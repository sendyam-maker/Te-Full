Attribute VB_Name = "acc_sav"
'Memo By Morgan 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'*************************************************
'  儲存資料表記錄
'
'*************************************************

Public Sub Frmacc2150_Save()
Dim rsTmp As New ADODB.Recordset
Dim strCP09 As String 'Add By Sindy 2022/6/28

    'Added by Lydia 2021/12/07 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        strControlButton = MsgText(602) 'Added by Lydia 2021/12/29 觀察玫音的操作是用滑鼠點選確定存檔，也有可能切換畫面有影響; CFP-028113-0-13收文BB0045703做帳單U11010920，帳單沒有ACC150
        Exit Sub
    End If
    'Added by Lydia 2021/12/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        If PUB_ChkUniText(Frmacc2150, , True, "TextBox") = False Then
            strControlButton = MsgText(602)
            'Memo by Lydia 2021/12/07 因為MsgBox選"否"回傳Enter值,會造成類似直接存檔,但不是真寫入DB
            Exit Sub
        End If
    End If
    'end 2021/12/07
   
   On Error GoTo Checking
   With Frmacc2150
      If .Text2 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
'         .Text2.SetFocus
         Exit Sub
      'Added by Lydia 2022/03/22 帳單輸入的代理人D/N NO.應該是必填,原本只寫在Text4_LostFocus
      ElseIf .Text4 = MsgText(601) Then
         MsgBox MsgText(10) & .Label3, , MsgText(5)
         strControlButton = MsgText(602)
         Exit Sub
      'end 2022/03/22
      Else
         If .Text1 <> MsgText(601) Then
            'Add by Morgan 2005/1/14 補足9碼
            Select Case Len(.Text1)
               Case 6
                 .Text1 = AfterZero(.Text1)
               Case 8
                 .Text1 = .Text1 & "0"
            End Select
            '2005/1/14 end
            If ExistCheck("fagent", "fa01 || fa02", IIf(Len(.Text1) = 6, AfterZero(.Text1), .Text1), .Label2) = False Then
               strControlButton = MsgText(602)
               .Text1.SetFocus
               Exit Sub
            End If
         '2005/12/7 ADD BY SONIA
         Else
            MsgBox MsgText(149), , MsgText(5)
            strControlButton = MsgText(602)
            .Text1.SetFocus
            Exit Sub
         '2005/12/7 END
         End If
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label4 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox .Label4 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
         End If
         'Add by Morgan 2006/4/25 檢查不可大於系統日
         If Val(ChangeTDateStringToTString(.MaskEdBox1.Text)) > Val(strSrvDate(2)) Then
            MsgBox "帳單日期不可大於系統日！", vbExclamation
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         End If
         '2006/4/25 end
   
         If Val(.Text6) <> Val(.Text14) Then
            MsgBox MsgText(59), , MsgText(5)
            strControlButton = MsgText(602)
            .Text6.SetFocus
            Exit Sub
         End If
      End If
      
      'Added by Morgan 2023/4/7
      If .Check3.Value = vbChecked Then
         If .Combo2.ListIndex = -1 Then
            MsgBox "急件需點選付款日期！", vbExclamation
            strControlButton = MsgText(602)
            .Combo2.SetFocus
            Exit Sub
         End If
         If .Text13 <> "" And .Text15 = "" Then
            MsgBox .Label14 & "輸入錯誤！", vbCritical
            strControlButton = MsgText(602)
            .Text13.SetFocus
            Exit Sub
         End If
      End If
      
      'end 2023/4/7
      
'      'Add by Morgan 2006/4/26
'      '檢查帳單資料是否重覆
'      strControlButton = MsgText(601)
'      If PUB_ChkDNDup(.MaskEdBox1.Text, .Text1.Text, .Text4.Text, .Text2.Text) = True Then
'         strControlButton = MsgText(602)
'         .Text4.SetFocus
'         Exit Sub
'      End If
      'Add By Sindy 2009/07/01
      strControlButton = MsgText(601)
      .DataGrid1.row = 0
      .DataGrid1.col = 0
      '檢查抵帳單資料是否重覆
      '若為專利處只須以代理人+代理人D/N No.做重覆檢核
      If Left(.DataGrid1.Text, 1) = "P" And Left(.DataGrid1.Text, 2) <> "PS" And _
         Left(Trim(GetStaffDepartment(strUserNum)), 2) = "P1" Then
         If PUB_ChkDNDup("", .Text1.Text, .Text4.Text, .Text2.Text, , 0) = True Then
            strControlButton = MsgText(602)
            .Text4.SetFocus
            Exit Sub
         End If
      Else
         If PUB_ChkDNDup(.MaskEdBox1.Text, .Text1.Text, .Text4.Text, .Text2.Text, , 0) = True Then
            strControlButton = MsgText(602)
            .Text4.SetFocus
            Exit Sub
         End If
      End If
      '2009/07/01 End
            
      '2009/6/15 add by sonia V09700097以修改功能刪除資料
      If .Combo1 = MsgText(601) Then
         MsgBox "無帳單幣別, 請檢核...", , MsgText(5)
         strControlButton = MsgText(602)
         .Combo1.SetFocus
         Exit Sub
      '2012/11/12 ADD BY SONIA 檢查幣別與之前帳單不符則提醒
      Else
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open " Select MAX((A1502+19110000)||A1505) From ACC150 Where A1503='" & .Text1.Text & "' And A1501<>'" & .Text2.Text & "' AND A1507 IS NULL ", cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If .Combo1 <> Mid(rsTmp.Fields(0), 9) Then
               If MsgBox("您輸入的幣別與之前的幣別 " & Mid(rsTmp.Fields(0), 9) & " 不符, 是否修改幣別 ? ", vbYesNo + vbDefaultButton1, .Caption) = vbYes Then
                  strControlButton = MsgText(602)
                  .Combo1.SetFocus
                  If rsTmp.State <> adStateClosed Then rsTmp.Close
                  Set rsTmp = Nothing
                  Exit Sub
               End If
            End If
         End If
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
      '2012/11/12 END
      End If
      If Val(.Text6) = 0 Then
         MsgBox MsgText(58), , MsgText(5)
         strControlButton = MsgText(602)
         .Text6.SetFocus
         Exit Sub
      End If
      '2009/6/15 end
      
      'Added by Morgan 2024/3/21
      'Y55766德國專利局帳單防呆檢查 1.不可有電子檔 2.金額非整數提醒可繼續
      If PUB_Y55766BillCheck(.Text2.Text, .Text1.Text, .Text6.Text) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If
      'end 2024/3/21
      
      'Add By Sindy 2018/2/22
      If .m_strIR01 <> "" Then
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open "select axf01,axf02 from acc151 where axf01 = '" & .Text2 & "' and axf03='" & .m_CP01 & .m_CP02 & .m_CP03 & .m_CP04 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            MsgBox "信件輸入必須與信件本所案號(" & .m_CP01 & "-" & .m_CP02 & "-" & .m_CP03 & "-" & .m_CP04 & ")一致！", , MsgText(5)
            strControlButton = MsgText(602)
            rsTmp.Close
            Exit Sub
         'Add By Sindy 2022/6/28
         Else
            If Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  rsTmp.Close
                  Exit Sub
               End If
            End If
            strCP09 = "" & rsTmp.Fields("axf02")
            '2022/6/28 END
         End If
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
      End If
      '2018/2/22 END
      
      '2013/3/18 add by sonia 婧瑄說000 南京捷恩凱信息技術有限公司,Y52268000江蘇舜禹信息技術有限公司都是翻譯費帳單,不必再審核
      'modify by sonia 2017/10/12 再+Y54868迅達翻譯社
      If Left(.Text1, 6) = "Y53541" Or Left(.Text1, 6) = "Y52268" Or Left(.Text1, 6) = "Y54868" Then
         .strYes = MsgText(601)
      End If
      '2013/3/18 end
      
      If strSaveConfirm = MsgText(3) Then
      
'Modify by Morgan 2006/4/26 修改也要檢查,往上移改Call共用函數
'        '檢查帳單資料是否重覆
'        If .ChkDataRepaet(.MaskEdBox1.Text, .Text1.Text, .Text4.Text) = True Then
'            .Text4.SetFocus
'            Exit Sub
'        End If
'2006/4/26 end
        
         If .adoacc150.RecordCount <> 0 Then
            .adoacc150.Find "a1501 = '" & .Text2 & "'", 0, adSearchForward, 1
            If .adoacc150.EOF = False Then
                'Modify By Cheng 2003/01/30
                '若帳單編號欄作用中
'               .Text2.SetFocus
               If .Text2.Enabled Then .Text2.SetFocus
               Exit Sub
            End If
         End If
         .adoacc150.AddNew
      End If
      .adoacc150.Fields("a1501").Value = .Text2
      If .Text1 <> MsgText(601) Then
         'Modify by Morgan 2006/4/26 原由Trigger控制
         If .Text1 = "Y51469000" Then
            .Text1 = "Y51566000"
         End If
         .adoacc150.Fields("a1503").Value = .Text1
      Else
         .adoacc150.Fields("a1503").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc150.Fields("a1504").Value = .Text4
      Else
         .adoacc150.Fields("a1504").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc150.Fields("a1502").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc150.Fields("a1502").Value = Null
      End If
      If .Combo1 <> MsgText(601) Then
         .adoacc150.Fields("a1505").Value = .Combo1
      Else
         .adoacc150.Fields("a1505").Value = Null
      End If
      If .Text6 <> MsgText(601) Then
         .adoacc150.Fields("a1506").Value = Val(.Text6)
      Else
         .adoacc150.Fields("a1506").Value = 0
      End If
      If .MaskEdBox2.Text <> MsgText(601) And .MaskEdBox2.Text <> MsgText(29) Then
         .adoacc150.Fields("a1507").Value = Val(FCDate(.MaskEdBox2.Text))
      Else
         .adoacc150.Fields("a1507").Value = Null
      End If
      If .Text8 <> MsgText(601) Then
         .adoacc150.Fields("a1509").Value = .Text8
      Else
         .adoacc150.Fields("a1509").Value = Null
      End If
      If .strYes = MsgText(603) Then
         .adoacc150.Fields("a1521").Value = .strYes
      Else
         .adoacc150.Fields("a1521").Value = Null
      End If
      
      'Added by Morgan 2019/3/12
      If .Check1.Value = vbChecked Then
         .adoacc150.Fields("a1525").Value = "Y"
      Else
         .adoacc150.Fields("a1525").Value = Null
      End If
      'end 2019/3/12
      
      'Added by Morgan 2019/3/15
      If .Check2.Value = vbChecked Then
         .adoacc150.Fields("a1526").Value = "Y"
      Else
         .adoacc150.Fields("a1526").Value = Null
      End If
      'end 2019/3/15
      
      'Added by Morgan 2023/4/7
      If .Check3.Value = vbChecked Then
         .adoacc150.Fields("a1527").Value = FCDate(.Combo2)
         If .Text13 = MsgText(601) Then
            .adoacc150.Fields("a1528").Value = Null '空字串要存Null否則若沒有重讀資料再存檔就會有錯誤
         Else
            .adoacc150.Fields("a1528").Value = .Text13
         End If
      'Added by Morgan 2023/4/20
      Else
         .adoacc150.Fields("a1527").Value = Null
         .adoacc150.Fields("a1528").Value = Null
      End If
      'end 2023/4/7
      
      If strSaveConfirm = MsgText(3) Then
         .adoacc150.Fields("a1514").Value = Val(strSrvDate(2))
         .adoacc150.Fields("a1515").Value = ServerTime
         .adoacc150.Fields("a1516").Value = strUserNum
      Else
         .adoacc150.Fields("a1517").Value = Val(strSrvDate(2))
         .adoacc150.Fields("a1518").Value = ServerTime
         .adoacc150.Fields("a1519").Value = strUserNum
      End If
      .adoacc150.UpdateBatch
      
      'Add by Sindy 2018/2/22
      If .m_strIR01 <> "" Then
         'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", strCP09, "")
         'Modify By Sindy 2023/7/12 IIf(Pub_StrUserSt03 = "F22", strCP09, "") => IIf(Left(Pub_StrUserSt03, 1) = "F", strCP09, "")
         PUB_UpdateEMailRec .m_strIR01, .m_strIR02, .m_strIR03, .m_strIR04, "Frmacc2150", IIf(Left(Pub_StrUserSt03, 1) = "F", strCP09, "")
      End If
      '2018/2/22 END
      
      cnnConnection.Execute "update caseprogress set cp61=cp61 where '" & .Text2 & "' in (cp61,cp62,cp63,cp87,cp88)", intI 'Added by Morgan 2017/8/11 為了觸發 Trigger "CASEPROGRESS_BEFORE"(acc150的檢查要存檔後才會有資料)
      
      'Added by Morgan 2025/7/2
      'FMP案於程序輸入帳單後能發信告知承辦人。Ex:【主旨：P129461年費帳單已輸入，請自行下載】--品薇
      If Left(Pub_StrUserSt03, 2) = "P1" Then
         'Modified by Morgan 2025/7/31 增加香港案都要通知--品薇
         'Modified by Morgan 2025/10/16 +分割案所有程序--品薇
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc13,mc14)" & _
            " select '" & strUserNum & "' mc01,cp13 mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||cpm04||'帳單已輸入，請自行下載' mc07" & _
            ",'如旨' mc08,cp09 mc13,'Y' mc14 from acc151,caseprogress a,patent,casepropertymap where axf01='" & .Text2 & "' and cp09(+)=axf02" & _
            " and cp01='P' and cp12 like 'F%' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and (pa09='013' or cp10 in (601,605,416,401,701)" & _
            " or exists(select * from caseprogress b where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='307'))" & _
            " and rownum<2"
         cnnConnection.Execute strSql, intI
      End If
      'end 2025/7/2
      
      .RecordShow
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   strControlButton = MsgText(602)
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc2160_Save()
Dim rsTmp As New ADODB.Recordset
Dim strCP09 As String 'Add By Sindy 2022/6/28
   
    'Added by Lydia 2021/12/07 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        strControlButton = MsgText(602) 'Added by Lydia 2021/12/29 參考Frmacc2150_Save
        Exit Sub
    End If
    'Added by Lydia 2021/12/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        If PUB_ChkUniText(Frmacc2160, , True, "TextBox") = False Then
            strControlButton = MsgText(602)
            Exit Sub
        End If
    End If
    'end 2021/12/07
    
   On Error GoTo Checking
   With Frmacc2160
      If .Text2 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
'         .Text2.SetFocus
         Exit Sub
      Else
         'If ExistCheck("acc150", "a1501", .Text4, .Label2) = False Then
         '   strControlButton = MsgText(602)
         '   .Text4.SetFocus
         '   Exit Sub
         'End If
         
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label5 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox .Label5 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
         End If
         '2006/5/15 復原 BY SONIA, 原已被REMARK
         If Val(.Text7) <> Val(.Text8) Then
            MsgBox MsgText(59), , MsgText(5)
            strControlButton = MsgText(602)
            .Text7.SetFocus
            Exit Sub
         End If
         '2006/5/15 END
'2012/8/9 cancel by sonia 移到下面去
'         '2009/6/15 add by sonia V09700097以修改功能刪除資料
'         If .Text1 = MsgText(601) Then
'            MsgBox MsgText(149), , MsgText(5)
'            strControlButton = MsgText(602)
'            .Text1.SetFocus
'            Exit Sub
'         End If
'         If .Combo1 = MsgText(601) Then
'            MsgBox "無抵帳幣別, 請檢核...", , MsgText(5)
'            strControlButton = MsgText(602)
'            .Combo1.SetFocus
'            Exit Sub
'         End If
'         If Val(.Text7) = 0 Then
'            MsgBox MsgText(58), , MsgText(5)
'            strControlButton = MsgText(602)
'            .Text7.SetFocus
'            Exit Sub
'         End If
'         '2009/6/15 end
''2012/8/9 end
         'Add By Sindy 2009/07/01
         strControlButton = MsgText(601)
         .DataGrid1.row = 0
         .DataGrid1.col = 1
         '檢查抵帳單資料是否重覆
         '若為專利處只須以代理人+代理人D/N No.做重覆檢核
         If Left(.DataGrid1.Text, 1) = "P" And Left(.DataGrid1.Text, 2) <> "PS" And _
            Left(Trim(GetStaffDepartment(strUserNum)), 2) = "P1" Then
            If PUB_ChkDNDup("", .Text1.Text, .Text5.Text, .Text2.Text, , 1) = True Then
               strControlButton = MsgText(602)
               .Text5.SetFocus
               Exit Sub
            End If
         Else
            If PUB_ChkDNDup(.MaskEdBox1.Text, .Text1.Text, .Text5.Text, .Text2.Text, , 1) = True Then
               strControlButton = MsgText(602)
               .Text5.SetFocus
               Exit Sub
            End If
         End If
         '2009/07/01 End
      End If
      
      'Add By Sindy 2018/2/22
      If .m_strIR01 <> "" Then
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open "select axg01,axg02 from acc161 where axg01 = '" & .Text2 & "' and axg03='" & .m_strCP01 & .m_strCP02 & .m_strCP03 & .m_strCP04 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            MsgBox "信件輸入必須與信件本所案號(" & .m_strCP01 & "-" & .m_strCP02 & "-" & .m_strCP03 & "-" & .m_strCP04 & ")一致！", , MsgText(5)
            strControlButton = MsgText(602)
            rsTmp.Close
            Exit Sub
         'Add By Sindy 2022/6/28
         Else
            If Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  rsTmp.Close
                  Exit Sub
               End If
            End If
            strCP09 = "" & rsTmp.Fields("axg02")
            '2022/6/28 END
         End If
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
      End If
      '2018/2/22 END

      If strSaveConfirm = MsgText(3) Then
         If .adoacc160.RecordCount <> 0 Then
            .adoacc160.Find "a1601 = '" & .Text2 & "'", 0, adSearchForward, 1
            If .adoacc160.EOF = False Then
               Exit Sub
            End If
         End If
         .adoacc160.AddNew
      End If
      .adoacc160.Fields("a1601").Value = .Text2
      If .Text1 <> MsgText(601) Then
         .adoacc160.Fields("a1603").Value = .Text1
      Else
         .adoacc160.Fields("a1603").Value = Null
      End If
      If .Text5 <> MsgText(601) Then
         .adoacc160.Fields("a1604").Value = .Text5
      Else
         .adoacc160.Fields("a1604").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc160.Fields("a1602").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc160.Fields("a1602").Value = Null
      End If
      If .Combo1 <> MsgText(601) Then
         .adoacc160.Fields("a1605").Value = .Combo1
      Else
         .adoacc160.Fields("a1605").Value = Null
      End If
      If .Text7 <> MsgText(601) Then
         .adoacc160.Fields("a1606").Value = Val(.Text7)
      Else
         .adoacc160.Fields("a1606").Value = 0
      End If
      If .Text9 <> MsgText(601) Then
         .adoacc160.Fields("a1608").Value = .Text9
      Else
         .adoacc160.Fields("a1608").Value = Null
      End If
      If strSaveConfirm = MsgText(3) Then
         .adoacc160.Fields("a1612").Value = Val(strSrvDate(2))
         .adoacc160.Fields("a1613").Value = ServerTime
         .adoacc160.Fields("a1614").Value = strUserNum
      Else
         .adoacc160.Fields("a1615").Value = Val(strSrvDate(2))
         .adoacc160.Fields("a1616").Value = ServerTime
         .adoacc160.Fields("a1617").Value = strUserNum
      End If
      .adoacc160.UpdateBatch
      .RecordShow
      
      '2012/8/9 add by sonia 從上面移下來,否則不會新增
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
      '2012/11/12 ADD BY SONIA 檢查幣別與之前帳單不符則提醒
      Else
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open " Select MAX((A1502+19110000)||A1505) From ACC150 Where A1503='" & .Text1.Text & "' And A1501<>'" & .Text2.Text & "' AND A1507 IS NULL ", cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If .Combo1 <> Mid(rsTmp.Fields(0), 9) Then
               If MsgBox("您輸入的幣別與之前的幣別 " & Mid(rsTmp.Fields(0), 9) & " 不符, 是否修改幣別 ? ", vbYesNo + vbDefaultButton1, .Caption) = vbYes Then
                  strControlButton = MsgText(602)
                  .Combo1.SetFocus
                  If rsTmp.State <> adStateClosed Then rsTmp.Close
                  Set rsTmp = Nothing
                  Exit Sub
               End If
            End If
         End If
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
      '2012/11/12 END
      End If
      If Val(.Text7) = 0 Then
         MsgBox MsgText(58), , MsgText(5)
         strControlButton = MsgText(602)
         .Text7.SetFocus
         Exit Sub
      End If
      '2012/8/9 end
      
      'Add by Sindy 2018/2/22
      If .m_strIR01 <> "" Then
         'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", strCP09, "")
         'Modify By Sindy 2023/7/12 IIf(Pub_StrUserSt03 = "F22", strCP09, "") => IIf(Left(Pub_StrUserSt03, 1) = "F", strCP09, "")
         PUB_UpdateEMailRec .m_strIR01, .m_strIR02, .m_strIR03, .m_strIR04, "Frmacc2160", IIf(Left(Pub_StrUserSt03, 1) = "F", strCP09, "")
      End If
      '2018/2/22 END
      
Checking:
   If Err.NUMBER = 0 Or Err.NUMBER = -2147217864 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc21g0_Save(ByRef oForm As Form)
Dim strYes As String
'Add By Cheng 2002/02/04
Dim rsTmp As New ADODB.Recordset

   On Error GoTo Checking
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21g0
   With oForm
      'Added by Morgan 2022/1/12 檢查畫面輸入欄位是否含有Unicode文字
      If PUB_ChkUniText(oForm, , True, "TextBox") = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If
      'end 2022/1/12
   
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If .Text2 = MsgText(601) Then
            MsgBox MsgText(10) & .Label2, , MsgText(5)
            strControlButton = MsgText(602)
            .Text2.SetFocus
            Exit Sub
         End If
         If .Text9 <> MsgText(601) Then
            If ExistCheck("acc010", "a0101", .Text9, .Label10) = False Then
               strControlButton = MsgText(602)
               .Text9.SetFocus
               Exit Sub
            End If
         End If
         If CheckLen(.Label3, .Text3, 100) = MsgText(603) Then
            strControlButton = MsgText(602)
            .Text3.SetFocus
            Exit Sub
         End If
      End If
      'Add by Morgan 2007/4/20
      If .Text1 = "FCP" Then
         If .Text2 = "601" Then
            If InStr(.Text11, "第 1 年") = 0 Then
               strControlButton = MsgText(602)
               MsgBox "為配合請款單列印，FCP 的 601 項目日文欄位必須有【第 1 年】5 字(含兩個半形空白)！"
               .Text11.SetFocus
               Exit Sub
            End If
         ElseIf .Text2 = "605" Then
            If InStr(.Text11, "第  年") = 0 Then
               strControlButton = MsgText(602)
               MsgBox "為配合請款單列印，FCP 的 605 項目日文欄位必須有【第  年】4 字(含兩個半形空白)！"
               .Text11.SetFocus
               Exit Sub
            End If
         End If
      End If
      'Add By Cheng 2002/02/04
      If rsTmp.State <> adStateClosed Then rsTmp.Close
      Set rsTmp = Nothing
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open " Select * From Staff,Staff_Group Where ST11=SG01(+) And ST01='" & strUserNum & "' And SG02='" & .Text1.Text & "'", cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         MsgBox "您無權維護此系統類別的相關請款資料", vbExclamation
         Exit Sub
      Else
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
      End If
      
      If strSaveConfirm = MsgText(3) Then
         .adoquery.CursorLocation = adUseClient
         .adoquery.Open "select a1j01 from acc1j0 where a1j01 = '" & .Text1 & "' and a1j02 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            .adoquery.Close
            .Text1.SetFocus
            Exit Sub
         End If
         .adoquery.Close
         .adoacc1j0.AddNew
      End If
      .adoacc1j0.Fields("a1j01").Value = .Text1
      .adoacc1j0.Fields("a1j02").Value = .Text2
      If .Text3 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j03").Value = .Text3
      Else
         .adoacc1j0.Fields("a1j03").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j04").Value = .Text4
      Else
         .adoacc1j0.Fields("a1j04").Value = Null
      End If
      If .Text5 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j05").Value = .Text5
      Else
         .adoacc1j0.Fields("a1j05").Value = Null
      End If
      If .Text6 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j06").Value = .Text6
      Else
         .adoacc1j0.Fields("a1j06").Value = Null
      End If
      If .Text11 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j16").Value = .Text11
      Else
         .adoacc1j0.Fields("a1j16").Value = Null
      End If
      If .Text7 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j07").Value = Val(.Text7)
      Else
         .adoacc1j0.Fields("a1j07").Value = 0
      End If
      If .Text8 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j08").Value = Val(.Text8)
      Else
         .adoacc1j0.Fields("a1j08").Value = 0
      End If
      If .Text9 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j09").Value = .Text9
      Else
         .adoacc1j0.Fields("a1j09").Value = Null
      End If
      If .Text10 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j17").Value = Val(.Text10)
      Else
         .adoacc1j0.Fields("a1j17").Value = 0
      End If
      'Add by Morgan 2010/11/4
      If .Text13 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j18").Value = .Text13
      Else
         .adoacc1j0.Fields("a1j18").Value = Null
      End If
      If .Text14 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j19").Value = .Text14
      Else
         .adoacc1j0.Fields("a1j19").Value = Null
      End If
      If .Text15 <> MsgText(601) Then
         .adoacc1j0.Fields("a1j20").Value = .Text15
      Else
         .adoacc1j0.Fields("a1j20").Value = Null
      End If
      'end 2010/11/14
      If strSaveConfirm = MsgText(3) Then
         .adoacc1j0.Fields("a1j10").Value = Val(strSrvDate(2))
         .adoacc1j0.Fields("a1j11").Value = ServerTime
         .adoacc1j0.Fields("a1j12").Value = strUserNum
      Else
         .adoacc1j0.Fields("a1j13").Value = Val(strSrvDate(2))
         .adoacc1j0.Fields("a1j14").Value = ServerTime
         .adoacc1j0.Fields("a1j15").Value = strUserNum
      End If
      .adoacc1j0.UpdateBatch
      .RecordShow
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc21h0_Save()
Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
'Dim strA1K18 As String 'Add By Sindy 2013/1/18
'Add By Sindy 2016/8/19
Dim bolCheck As Boolean, strText As String
Dim i As Integer, strCP09 As String
'2016/8/19 END
   
On Error GoTo Checking
   
   With Frmacc21h0
      If .Text5 = MsgText(601) Then
         MsgBox MsgText(10) & .Label6, , MsgText(5)
         strControlButton = MsgText(602)
         .Text5.SetFocus
         Exit Sub
      Else
         .adoquery.CursorLocation = adUseClient
         .adoquery.Open "select cp09 from caseprogress where cp01 = '" & .Text1 & "' and cp02 = '" & .Text6 & "' and cp03 = '" & .Text7 & "' and cp04 = '" & .Text8 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount = 0 Then
            strControlButton = MsgText(602)
            .Text1.SetFocus
            .adoquery.Close
            Exit Sub
         End If
         .adoquery.Close
      End If
      
      'Add By Sindy 2016/8/19 存檔時, 若上方之未開請款單GRID中仍有有發文規費(cp84>0)但無收文規費(cp17=0)的資料,
      '(不限制 FCP案)則提示"....有發文規費但尚未請款, 是否要一併點選至此請款單？"
      '讓使用者可選擇是(不存檔)或否(存檔)
      bolCheck = False
      strText = ""
      If .Adodc1.Recordset.RecordCount > 0 Then
         .Adodc1.Recordset.MoveFirst
         For i = 1 To .Adodc1.Recordset.RecordCount
            strCP09 = .Adodc1.Recordset.Fields(0)
            stSQL = "Select cp84,cp17 From caseprogress Where cp09='" & strCP09 & "' and nvl(cp84,0)>0 and nvl(cp17,0)=0"
            intR = 1
            Set adoRst = ClsLawReadRstMsg(intR, stSQL)
            If intR = 1 Then
               bolCheck = True
               strText = strText & "、" & .Adodc1.Recordset.Fields(2)
            End If
            .Adodc1.Recordset.MoveNext
         Next i
         If bolCheck = True And strText <> "" Then
            strText = Mid(strText, 2)
            If MsgBox(strText & "有發文規費但尚未請款，是否要一併點選至此請款單？", vbYesNo) = vbYes Then
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
      End If
      '2016/8/19 END
      
      'add by sonia 2017/10/17 存檔時, 若上方之未開請款單GRID中仍有有帳單(cp61)但無收文費用(cp16=0)的資料,
      '讓使用者可選擇是(不存檔)或否(存檔)
      bolCheck = False
      strText = ""
      If .Adodc1.Recordset.RecordCount > 0 Then
         .Adodc1.Recordset.MoveFirst
         For i = 1 To .Adodc1.Recordset.RecordCount
            strCP09 = .Adodc1.Recordset.Fields(0)
            stSQL = "Select cp61,cp16 From caseprogress Where cp09='" & strCP09 & "' and cp61 is not null and nvl(cp16,0)=0"
            intR = 1
            Set adoRst = ClsLawReadRstMsg(intR, stSQL)
            If intR = 1 Then
               bolCheck = True
               strText = strText & "、" & .Adodc1.Recordset.Fields(2)
            End If
            .Adodc1.Recordset.MoveNext
         Next i
         If bolCheck = True And strText <> "" Then
            strText = Mid(strText, 2)
            If MsgBox(strText & "有帳單但尚未請款，是否要一併點選至此請款單？", vbYesNo) = vbYes Then
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
      End If
      'end 2017/10/17
      
      If .Adodc2.Recordset.RecordCount = 0 Then
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      End If
      If strSaveConfirm = MsgText(3) Then
         If .adoacc1k0.RecordCount <> 0 Then
            .adoacc1k0.Find "a1k01 = '" & .Text5 & "'", 0, adSearchForward, 1
            If .adoacc1k0.EOF = False Then
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               .Text5.SetFocus
               Exit Sub
            End If
         End If
         .adoacc1k0.AddNew
         'Add by Morgan 2010/5/18
         '新增時FMP案預設特殊請款
         stSQL = "select cp01,cp12 from caseprogress where cp60='" & .Text5 & "' order by cp05,cp09"
         intR = 1
         Set adoRst = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
            '2010/5/26 MODIFY BY SONIA FMT也要預設
            'If adoRst(0) = "P" And Left(adoRst(1), 1) = "F" Then
'            If (adoRst(0) = "P" Or adoRst(0) = "T") And Left(adoRst(1), 1) = "F" Then
'               .adoacc1k0.Fields("a1k32").Value = "Y"
'            End If
         End If
         'end 2010/5/18
         '2010/5/26 add by sonia FCL,CFL,LIN預設特殊請款
         If .Text1 = "FCL" Or .Text1 = "CFL" Or .Text1 = "LIN" Then
            .adoacc1k0.Fields("a1k32").Value = "Y"
         End If
         '2010/5/26 end
      Else
         .adoacc1k0.Find "a1k01 = '" & .Text5 & "'", 0, adSearchForward, 1
      End If
      .adoacc1k0.Fields("a1k01").Value = .Text5
      .adoquery.CursorLocation = adUseClient
        'FC代理人-->申請人
'      .adoquery.Open "select pa75 as fag from patent where pa01 = '" & .Adodc2.Recordset.Fields("cp01").Value & "' and pa02 = '" & .Adodc2.Recordset.Fields("cp02").Value & "' and pa03 = '" & .Adodc2.Recordset.Fields("cp03").Value & "' and pa04 = '" & .Adodc2.Recordset.Fields("cp04").Value & "' " & _
'                     "union all select tm44 as fag from trademark where tm01 = '" & .Adodc2.Recordset.Fields("cp01").Value & "' and tm02 = '" & .Adodc2.Recordset.Fields("cp02").Value & "' and tm03 = '" & .Adodc2.Recordset.Fields("cp03").Value & "' and tm04 = '" & .Adodc2.Recordset.Fields("cp04").Value & "' " & _
'                     "union all select lc22 as fag from lawcase where lc01 = '" & .Adodc2.Recordset.Fields("cp01").Value & "' and lc02 = '" & .Adodc2.Recordset.Fields("cp02").Value & "' and lc03 = '" & .Adodc2.Recordset.Fields("cp03").Value & "' and lc04 = '" & .Adodc2.Recordset.Fields("cp04").Value & "' " & _
'                     "union all select sp26 as fag from servicepractice where sp01 = '" & .Adodc2.Recordset.Fields("cp01").Value & "' and sp02 = '" & .Adodc2.Recordset.Fields("cp02").Value & "' and sp03 = '" & .Adodc2.Recordset.Fields("cp03").Value & "' and sp04 = '" & .Adodc2.Recordset.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      .adoquery.Open "select Nvl(pa75, PA26) as fag from patent where pa01 = '" & .Adodc2.Recordset.Fields("cp01").Value & "' and pa02 = '" & .Adodc2.Recordset.Fields("cp02").Value & "' and pa03 = '" & .Adodc2.Recordset.Fields("cp03").Value & "' and pa04 = '" & .Adodc2.Recordset.Fields("cp04").Value & "' " & _
                     "union all select Nvl(tm44, TM23) as fag from trademark where tm01 = '" & .Adodc2.Recordset.Fields("cp01").Value & "' and tm02 = '" & .Adodc2.Recordset.Fields("cp02").Value & "' and tm03 = '" & .Adodc2.Recordset.Fields("cp03").Value & "' and tm04 = '" & .Adodc2.Recordset.Fields("cp04").Value & "' " & _
                     "union all select Nvl(lc22, LC11) as fag from lawcase where lc01 = '" & .Adodc2.Recordset.Fields("cp01").Value & "' and lc02 = '" & .Adodc2.Recordset.Fields("cp02").Value & "' and lc03 = '" & .Adodc2.Recordset.Fields("cp03").Value & "' and lc04 = '" & .Adodc2.Recordset.Fields("cp04").Value & "' " & _
                     "union all select Nvl(sp26, SP08) as fag from servicepractice where sp01 = '" & .Adodc2.Recordset.Fields("cp01").Value & "' and sp02 = '" & .Adodc2.Recordset.Fields("cp02").Value & "' and sp03 = '" & .Adodc2.Recordset.Fields("cp03").Value & "' and sp04 = '" & .Adodc2.Recordset.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If .adoquery.RecordCount <> 0 Then
        If IsNull(.adoquery.Fields(0).Value) Then
           .adoacc1k0.Fields("a1k03").Value = Null
        Else
          .adoacc1k0.Fields("a1k03").Value = .adoquery.Fields(0).Value
        End If
      End If
      .adoquery.Close
      If strSaveConfirm = MsgText(3) Then
         .adoacc1k0.Fields("a1k02").Value = Val(strSrvDate(2))
      End If
      'Remove by Morgan 2005/5/17 移到21h1作
'      If IsNull(.Adodc2.Recordset.Fields("CP17").Value) Then
'         .adoacc1k0.Fields("A1K09").Value = 0
'      Else
'         .adoacc1k0.Fields("A1K09").Value = .Adodc2.Recordset.Fields("CP17").Value
'      End If
'      'Modify By Sindy 2013/1/18
'      '.adoacc1k0.Fields("a1k18").Value = "USD"   '2009/6/24 由下面移上來
'      .adoacc1k0.Fields("a1k33").Value = PUB_GetDefaultCurrPrintType(.Text1, .adoacc1k0.Fields("a1k03").Value, "", strA1K18)
'      .adoacc1k0.Fields("a1k18").Value = strA1K18
'      '2013/1/18 End
'      .adorate.CursorLocation = adUseClient
'      '2009/6/24 modify by sonia 先依a1k03抓預設之請款幣別及匯率
'      '.adorate.Open "select usxr02 from usxrate where usxr01 <= " & Val(.adoacc1k0.Fields("a1k02").Value) & " order by usxr01 desc", adoTaie, adOpenStatic, adLockReadOnly
'      'If .adorate.RecordCount <> 0 Then
'      '   If IsNull(.adorate.Fields(0).Value) Then
'      '      .adoacc1k0.Fields("a1k10").Value = 0
'      '   Else
'      '      .adoacc1k0.Fields("a1k10").Value = .adorate.Fields(0).Value
'      '   End If
'      'Else
'      '   .adoacc1k0.Fields("a1k10").Value = 0
'      'End If
'      'Modify By Sindy 2011/3/3
'      'Select Case pub_GetCurrency(.adoacc1k0.Fields("a1k03").Value)
'      Select Case pub_GetCurrency(.adoacc1k0.Fields("a1k03").Value, .Text1)
'      '2011/3/3 End
'         'Modify By Sindy 2013/1/18
'         Case "USD", "NTD", ""
'            .adorate.Open "select usxr02 from usxrate where usxr01 <= " & Val(.adoacc1k0.Fields("a1k02").Value) & " order by usxr01 desc", adoTaie, adOpenStatic, adLockReadOnly
'            If .adorate.RecordCount <> 0 Then
'               If IsNull(.adorate.Fields(0).Value) Then
'                  .adoacc1k0.Fields("a1k10").Value = 0
'               Else
'                  .adoacc1k0.Fields("a1k10").Value = .adorate.Fields(0).Value
'               End If
'            Else
'               .adoacc1k0.Fields("a1k10").Value = 0
'            End If
'            .adorate.Close
'         Case "RMB"
'            .adoacc1k0.Fields("a1k18").Value = "RMB"
'            .adoacc1k0.Fields("a1k10").Value = PUB_GetUSXRate_1(Val(.adoacc1k0.Fields("a1k02").Value), .adoacc1k0.Fields("a1k18").Value) * (1 / PUB_GetDNRate(Val(.adoacc1k0.Fields("a1k02").Value), .adoacc1k0.Fields("a1k18").Value))
'      End Select
'      '2009/6/24 end
'      'Modify By Sindy 2013/1/18 抓請款匯率
'      .adoacc1k0.Fields("a1k10").Value = PUB_GetUSXRate_1(Val(.adoacc1k0.Fields("a1k02").Value), .adoacc1k0.Fields("a1k18").Value)
'      '2013/1/18 End
      If .Text1 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k13").Value = .Text1
      Else
         .adoacc1k0.Fields("a1k13").Value = Null
      End If
      If .Text6 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k14").Value = .Text6
      Else
         .adoacc1k0.Fields("a1k14").Value = Null
      End If
      If .Text7 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k15").Value = .Text7
      Else
         .adoacc1k0.Fields("a1k15").Value = Null
      End If
      If .Text8 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k16").Value = .Text8
      Else
         .adoacc1k0.Fields("a1k16").Value = Null
      End If
      If strSaveConfirm = MsgText(3) Then
         .adoacc1k0.Fields("a1k19").Value = Val(strSrvDate(2))
         .adoacc1k0.Fields("a1k20").Value = ServerTime
         .adoacc1k0.Fields("a1k21").Value = strUserNum
      Else
         .adoacc1k0.Fields("a1k22").Value = Val(strSrvDate(2))
         .adoacc1k0.Fields("a1k23").Value = ServerTime
         .adoacc1k0.Fields("a1k24").Value = strUserNum
      End If
      .adoacc1k0.UpdateBatch
      strSaveConfirm = MsgText(601)
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
   Set adoRst = Nothing
End Sub

Public Sub Frmacc21h1_Save()
    Dim strMsg As String, strMsg2 As String '2013/10/30 顯示訊息
    
    'Added by Lydia 2021/12/08 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        strControlButton = MsgText(602) 'Added by Lydia 2021/12/29 參考Frmacc2150_Save
        Exit Sub
    End If
    'Added by Lydia 2021/12/08 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        If PUB_ChkUniText(Frmacc21h1, , True, "TextBox") = False Then
            Exit Sub
        End If
    End If
    'end 2021/12/08
    
   On Error GoTo Checking
   
   With Frmacc21h1
        'Add by Amy 2013/10/30 +帳款處理訊息
         strMsg = GetDizhang(strCustNo, , False) '舊編號設定
         If strCustNo <> Left(.Text8.Text, 8) Then
            strMsg2 = GetDizhang(.Text8.Text, , False) '新編號設定
         Else
            strMsg = ""
         End If
         If strMsg <> "" And strMsg2 <> "" Then
            strMsg = "原 請款對象編號 " & strMsg & "新 請款對象編號 " & strMsg2
         ElseIf strMsg <> "" Then
            strMsg = "原 請款對象編號 " & strMsg
         ElseIf strMsg2 <> "" Then
            strMsg = "新 請款對象編號 " & strMsg2
         End If
         If strMsg <> "" Then
            MsgBox strMsg & "詳細情形請與財務處聯繫!!", vbInformation
         End If
         'end 2013/10/30
      
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc1k0.Fields("a1k02").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc1k0.Fields("a1k02").Value = Null
      End If
      '2009/4/24 add by sonia
      If .Combo3.Text <> MsgText(601) Then
         .adoacc1k0.Fields("a1k18").Value = .Combo3.Text
      Else
         .adoacc1k0.Fields("a1k18").Value = Null
      End If
      '2009/4/24 end
      '2009/4/24 modify by sonia
      '.adoquery.CursorLocation = adUseClient
      '.adoquery.Open "select usxr02 from usxrate where usxr01 <= " & Val(.adoacc1k0.Fields("a1k02").Value) & " order by usxr01 desc", adoTaie, adOpenStatic, adLockReadOnly
      'If .adoquery.RecordCount <> 0 Then
      '   If IsNull(.adoquery.Fields(0).Value) Then
      '      .adoacc1k0.Fields("a1k10").Value = 0
      '   Else
      '      .adoacc1k0.Fields("a1k10").Value = .adoquery.Fields(0).Value
      '   End If
      'Else
      '   .adoacc1k0.Fields("a1k10").Value = 0
      'End If
      '.adoquery.Close
      
      'Modified by Morgan 2012/9/18 HP用報價匯率
      If .Text19.Enabled = True And .Combo3 = "USD" Then
         .adoacc1k0.Fields("a1k10").Value = Val(.Text19)
      Else
         '依請款幣別對台幣匯率及對美金匯率換算成美金對台幣匯率
         '.adoacc1k0.Fields("a1k10").Value = PUB_GetUSXRate_1(Replace(.MaskEdBox1.Text, "/", ""), .Combo3.Text) * (1 / PUB_GetDNRate(Replace(.MaskEdBox1.Text, "/", ""), .Combo3.Text))
         'Modify By Sindy 2012/12/27 存請款幣別對台幣匯率
         .adoacc1k0.Fields("a1k10").Value = Val(.Text19)
         '2012/12/27 End
      End If
      '2009/4/24 end
      If .Text4 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k04").Value = .Text4
      Else
         .adoacc1k0.Fields("a1k04").Value = Null
      End If
      
      'Add by Morgan 2010/5/13
      If .Text5 <> "" Then
         .adoacc1k0.Fields("a1k32").Value = .Text5
      Else
         .adoacc1k0.Fields("a1k32").Value = Null
      End If
      
      'Added by Morgan 2012/12/6 列印幣別格式
      .adoacc1k0.Fields("a1k33").Value = .Combo4.ListIndex + 1
      
'      If .Text13 <> MsgText(601) Then
''Modified by Morgan 2012/11/2 改都不捨去--David,Frances
''            'Modify By Cheng 2004/04/23
''            '美金欄位取至整數位(無條件捨去)
'''         .adoacc1k0.Fields("a1k08").Value = Val(.Text13)
''
''         'Modify by Morgan 2004/6/28
''         '.adoacc1k0.Fields("a1k08").Value = Fix(Val(.Text13))
''         'Y48673000,Y49575000 存小數兩位
''         If .m_strCP10 = 605 And (.Text8 = "Y48673000" Or .Text8 = "Y49575000") Then
''            .adoacc1k0.Fields("a1k08").Value = Format(Val(.Text13), FAmount)
''
''         'Added by Morgan 2012/1/31 Y52218 PanKorea Patent & Law Firm 美金加總保留小數點--David
''         'Modified by Morgan 2012/7/6 +Y34126 L'AIR LIQUIDE SA DIRECTION DE LA PROPRIETE INTELLECTUELLE--David
''         'Modified by Morgan 2012/9/18 +Y48292000 HP
''         'Modified by Morgan 2012/10/11 +Y45149000,Y45149010
''         'Modified by Morgan 2012/11/2 +Y23045000 --陳芊穎
''         ElseIf .Text8 = "Y52218000" Or .Text8 = "Y34126000" Or .Text8 = "Y48292000" Or .Text8 = "Y45149000" Or .Text8 = "Y45149010" Or .Text8 = "Y23045000" Then
''            .adoacc1k0.Fields("a1k08").Value = Val(.Text13)
''
''         Else
''            .adoacc1k0.Fields("a1k08").Value = Fix(Val(.Text13))
''         End If
''
''            'End
'         .adoacc1k0.Fields("a1k08").Value = Val(.Text13)
''end 2012/11/2
'      Else
'         .adoacc1k0.Fields("a1k08").Value = 0
'      End If
      'Modify By Sindy 2012/12/27
      If .Text12 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k08").Value = Val(.Text12)
      Else
         .adoacc1k0.Fields("a1k08").Value = 0
      End If
      '2012/12/27 End
      If .Text14 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k11").Value = Val(.Text14)
      Else
         .adoacc1k0.Fields("a1k11").Value = 0
      End If
      If .Text6 <> MsgText(601) Then
            'Modify By Cheng 2003/06/25
'         .adoacc1k0.Fields("a1k27").Value = .Text6
         .adoacc1k0.Fields("a1k27").Value = Left(.Text6 & "000000000", 9)
      Else
         .adoacc1k0.Fields("a1k27").Value = Null
      End If
      If .Text8 <> MsgText(601) Then
            'Modify By Cheng 2003/06/25
'         .adoacc1k0.Fields("a1k28").Value = .Text8
         .adoacc1k0.Fields("a1k28").Value = Left(.Text8 & "000000000", 9)
         If .Text2 <> MsgText(601) Then
                'Modify By Cheng 2003/06/25
'            .adoacc1k0.Fields("a1k03").Value = .Text2
            .adoacc1k0.Fields("a1k03").Value = Left(.Text2 & "000000000", 9)
         Else
                'Modify By Cheng 2003/06/28
'            .adoacc1k0.Fields("a1k03").Value = .Text8
            .adoacc1k0.Fields("a1k03").Value = Left(.Text8 & "000000000", 9)
         End If
      Else
         .adoacc1k0.Fields("a1k28").Value = Null
         .adoacc1k0.Fields("a1k03").Value = Null
      End If
      If .Text11 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k05").Value = .Text11
      Else
         .adoacc1k0.Fields("a1k05").Value = Null
      End If
      
      'Added by Morgan 2013/10/18
      If .Text24 <> MsgText(601) Then
         .adoacc1k0.Fields("a1k34").Value = .Text24
      Else
         .adoacc1k0.Fields("a1k34").Value = Null
      End If
      'end 2013/10/18
      
      .adoquery.CursorLocation = adUseClient
      'Modify by Morgan 2005/5/17 加判斷 a1l03='T' 抓 al04='03',a1l03<>'T' 抓 al04末二碼='99'
      '.adoquery.Open "select sum(a1l05) from acc1l0 where a1l01 = '" & .Text1 & "' and substr(a1l04, length(a1l04) - 1, 2) = '99'", adoTaie, adOpenStatic, adLockReadOnly
      'Modify by Morgan 2006/4/12 改判斷 a1l03='T'開頭 抓 al04='03',a1l03<>'T'開頭 抓 al04末二碼='99'
      '2009/7/27 MODIFY BY SONIA a1l03='T'除抓 al04='03'外,al04末二碼='9也要抓 X09807783
      '.adoquery.Open "select nvl(sum(a1l05),0) from acc1l0 where a1l01 = '" & .Text1 & "' and ( (substr(a1l03,1,1)<>'T' and substr(a1l04, length(a1l04) - 1, 2) = '99') or (substr(a1l03,1,1)='T' and a1l04='03') )", adoTaie, adOpenStatic, adLockReadOnly
      'Modify by Morgan 2010/4/22 規費也可能有折扣
      '.adoquery.Open "select nvl(sum(a1l05),0) from acc1l0 where a1l01 = '" & .Text1 & "' and ( (substr(a1l03,1,1)<>'T' and substr(a1l04, length(a1l04) - 1, 2) = '99') or (substr(a1l03,1,1)='T' and (a1l04='03' OR substr(a1l04, length(a1l04) - 1, 2) = '99') ) )", adoTaie, adOpenStatic, adLockReadOnly
      'Modify By Sindy 2013/1/28
      If strSrvDate(1) >= AccFMPImputCurrStarDate Then
         'Modify By Sindy 2012/12/27 +項目後2碼為98時,代表是代收代付加入規費
         'Modify By Sindy 2013/1/24 規費=98+99且均不減折扣
         .adoquery.Open "select nvl(sum(a1l05),0) from acc1l0 where a1l01='" & .Text1 & "' and (substr(a1l04,-2)='98' or substr(a1l04,-2)='99' or (substr(a1l03,1,1)='T' and a1l04='03'))", adoTaie, adOpenStatic, adLockReadOnly
      Else
      '2013/1/28 End
         .adoquery.Open "select nvl(sum(a1l05-nvl(a1l07,0)),0) from acc1l0 where a1l01 = '" & .Text1 & "' and ( substr(a1l04,-2) = '99' or (substr(a1l03,1,1)='T' and a1l04='03') )", adoTaie, adOpenStatic, adLockReadOnly
      End If
      If .adoquery.RecordCount <> 0 Then
         If IsNull(.adoquery.Fields(0).Value) = False Then
            'Modify by Morgan 2010/5/13
            'adoTaie.Execute "update acc1k0 set a1k09 = " & .adoquery.Fields(0).Value & " where a1k01 = '" & .Text1 & "'"
            .adoacc1k0.Fields("a1k09").Value = .adoquery.Fields(0).Value
            'Added by Morgan 2013/2/20
            'FMP新案請款+安全基金2000
            If .m_bolFMPnewcase = True Then
               .adoacc1k0.Fields("a1k09").Value = .adoacc1k0.Fields("a1k09").Value + 2000
            End If
            'end 2013/2/20
            'add by sonia 2017/10/17 FCP及FMP案B類收文927其他翻譯且承辦人為外翻編號且相關總收文號為C類之結匯金額,要再扣除(扣點數),例FCP-050279(U10608558)
            .adoselect.CursorLocation = adUseClient
            .adoselect.Open "select a1p07,a1w01,a1w02,cp60,cp61 from acc1w0,caseprogress,acc1p0 where a1w01='" & .adoacc1k0.Fields("a1k01").Value & "' and substr(a1w02,1,1)='B' and a1w02=cp09(+) " & _
                            "and cp01 in ('P','FCP') and cp10='927' and substr(cp14,1,1)='F' and substr(cp43,1,1)='C' and cp61||a1w02=a1p23 and a1p07>0", adoTaie, adOpenStatic, adLockReadOnly
            If .adoselect.RecordCount <> 0 Then
               .adoacc1k0.Fields("a1k09").Value = .adoacc1k0.Fields("a1k09").Value + Val(.adoselect.Fields("a1p07"))
            End If
            .adoselect.Close
            'end 2017/10/17
         End If
      End If
      .adoquery.Close
      .adoacc1k0.Fields("a1k22").Value = Val(strSrvDate(2))
      .adoacc1k0.Fields("a1k23").Value = ServerTime
      .adoacc1k0.Fields("a1k24").Value = strUserNum
      .adoacc1k0.UpdateBatch
      If .m_bolIsBatch Then .Frmacc21p1_Save 'Added by Morgan 2014/8/18

      .adoacc1k0.ReQuery
      
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc21i0_Save()
   On Error GoTo Checking
   With Frmacc21i0
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If .MaskEdBox2.Text = MsgText(601) Or .MaskEdBox2.Text = MsgText(29) Then
            MsgBox .Label10 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox2.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox2.Text) = MsgText(603) Then
               MsgBox .Label10 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox2.SetFocus
               Exit Sub
            End If
         End If
      End If
      
      'Added by Morgan 2018/3/21
      '折讓後台幣不應小於(規費-規費折讓金額)
      If Val(.Text23) > Val("" & .adoacc1k0.Fields("a1k09").Value) Then
         MsgBox "規費折讓金額不可高於規費金額！", vbExclamation
         strControlButton = MsgText(602)
         .Text23.SetFocus
         Exit Sub
      ElseIf Val(.Text19) < (Val("" & .adoacc1k0.Fields("a1k09").Value) - Val(.Text23)) Then
         MsgBox "折讓後金額已低於規費金額，請輸入正確的規費折讓金額！", vbExclamation
         strControlButton = MsgText(602)
         .Text23.SetFocus
         Exit Sub
      'Added by Morgan 2018/10/5
      ElseIf Val(.Text23) > Val(.Text17) Then
         MsgBox "規費折讓金額不可大於折讓金額！", vbExclamation
         strControlButton = MsgText(602)
         .Text23.SetFocus
         Exit Sub
      End If
      'end 2018/3/21
      
      'Added by Morgan 2020/8/5 --婉莘
      '台幣折讓與外幣折讓必須同時有值或為0(取消折讓)
      If (Val(.Text17) > 0 And Val(.Text9) = 0) Or (Val(.Text17) = 0 And Val(.Text9) > 0) Then
         MsgBox "折讓金額(外幣)與折讓金額(台幣)必須同時有值或為0(取消折讓)！", vbExclamation
         strControlButton = MsgText(602)
         Exit Sub
      End If
      'end 2020/8/5
      
      '有外幣折讓金額時
      If .Text9 <> MsgText(601) Then
         '2009/4/24 MODIFY BY SONIA
         '.adoacc1k0.Fields("a1k06").Value = Val(.Text9)
         
         'Modify By Sindy 2012/11/29
         '.adoacc1k0.Fields("a1k06").Value = Val(.Text22)
         .adoacc1k0.Fields("a1k06").Value = Val(.Text17) '台幣折讓金額
         '2012/11/29 End
         
         .adoacc1k0.Fields("a1k31").Value = Val(.Text9) '外幣折讓金額
         '2009/4/24 END
      Else
         .adoacc1k0.Fields("a1k06").Value = 0
         .adoacc1k0.Fields("a1k31").Value = 0   '2009/4/24 ADD BY SONIA
      End If
      If Val(.Text9) <> 0 Then
         If .MaskEdBox2.Text <> MsgText(601) And .MaskEdBox2.Text <> MsgText(29) Then
            .adoacc1k0.Fields("a1k07").Value = Val(FCDate(.MaskEdBox2.Text))
         Else
            .adoacc1k0.Fields("a1k07").Value = Null
         End If
      Else
         .adoacc1k0.Fields("a1k07").Value = Null
      End If
      .adoacc1k0.Fields("a1k22").Value = Val(strSrvDate(2))
      .adoacc1k0.Fields("a1k23").Value = ServerTime
      .adoacc1k0.Fields("a1k24").Value = strUserNum
      .adoacc1k0.Fields("a1k36").Value = Val(.Text23) '台幣折讓金額 Added by Morgan 2018/3/20
      .adoacc1k0.UpdateBatch
      .RecordShow
      
      'Add by Morgan 2010/5/21
      '檢查是否須進入點數分配畫面
      PUB_ReAssignPoint .Text1
      If PUB_ChkPointOk(.Text1) = False Then
         MsgBox "請款點數與分配點數不符，按確定後進入點數分配輸入作業！"
         strItemNo = .Text1
         Frmacc21h3.Show vbModal
         strFormName = "Frmacc21i0"
      End If
      'end 2010/5/21
      
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc21j0_Save()
Dim adoquery As New ADODB.Recordset
Dim strCP09 As String 'Add By Sindy 2022/6/28
   
   On Error GoTo Checking
   
   With Frmacc21j0
      If .Text2 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text2.SetFocus
         Exit Sub
      Else
         If .MaskEdBox2.Text = MsgText(601) Or .MaskEdBox2.Text = MsgText(29) Then
            MsgBox .Label7 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox2.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox2.Text) = MsgText(603) Then
               MsgBox .Label7 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox2.SetFocus
               Exit Sub
            End If
         End If
         
         'Added by Morgan 2025/7/14
         If "" & .adoacc150.Fields("a1512") <> "" Then
            MsgBox "已抵帳不可作廢！", vbCritical
            strControlButton = MsgText(602)
            Exit Sub
         End If
         'end 2025/7/14
         
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select a1702 from acc170 where a1702 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(218), , MsgText(5)
            strControlButton = MsgText(602)
            adoquery.Close
            .Text2.SetFocus
            Exit Sub
         End If
         adoquery.Close
      End If
      
      'Add By Sindy 2018/2/22
      If .m_strIR01 <> "" Then
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select axf01,axf02 from acc151 where axf01 = '" & .Text2 & "' and axf03='" & .m_strCP01 & .m_strCP02 & .m_strCP03 & .m_strCP04 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <= 0 Then
            MsgBox "信件輸入必須與信件本所案號(" & .m_strCP01 & "-" & .m_strCP02 & "-" & .m_strCP03 & "-" & .m_strCP04 & ")一致！", , MsgText(5)
            strControlButton = MsgText(602)
            adoquery.Close
            .Text2.SetFocus
            Exit Sub
         'Add By Sindy 2022/6/28
         Else
            If Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  adoquery.Close
                  Exit Sub
               End If
            End If
            strCP09 = "" & adoquery.Fields("axf02")
            '2022/6/28 END
         End If
         adoquery.Close
      End If
      '2018/2/22 END
      
      adoTaie.BeginTrans 'Added by Lydia 2016/12/27 包在Transaction
      .adoacc150q.CursorLocation = adUseClient
      .adoacc150q.Open "select * from acc150 where a1501 = '" & .Text2 & "' order by a1501 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If .adoacc150q.RecordCount = 0 Then
         MsgBox MsgText(33), , MsgText(5)
         Exit Sub
      End If
      If .MaskEdBox2.Text <> MsgText(601) And .MaskEdBox2.Text <> MsgText(29) Then
         .adoacc150q.Fields("a1507").Value = Val(FCDate(.MaskEdBox2.Text))
      Else
         .adoacc150q.Fields("a1507").Value = Null
      End If
      .adoacc150q.UpdateBatch
      .adoacc150q.Close
      '2010/3/29 MODIFY BY SONIA
      'adoTaie.Execute "update caseprogress set cp61 = decode(cp61, '" & .Text2 & "', null, cp61), cp62 = decode(cp62, '" & .Text2 & "', null, cp62), cp63 = decode(cp63, '" & .Text2 & "', null, cp63) where cp61 = '" & .Text2 & "' or cp62 = '" & .Text2 & "' or cp63 = '" & .Text2 & "'"
      adoTaie.Execute "update caseprogress set cp61 = decode(cp61, '" & .Text2 & "', null, cp61), cp62 = decode(cp62, '" & .Text2 & "', null, cp62), cp63 = decode(cp63, '" & .Text2 & "', null, cp63), cp87 = decode(cp87, '" & .Text2 & "', null, cp87), cp88 = decode(cp88, '" & .Text2 & "', null, cp88) where cp61 = '" & .Text2 & "' or cp62 = '" & .Text2 & "' or cp63 = '" & .Text2 & "' or cp87 = '" & .Text2 & "' or cp88 = '" & .Text2 & "'"
      '2010/3/29 END
      
      .adoacc150.Close
      .adoacc150.CursorLocation = adUseClient
      .adoacc150.Open "select * from acc150 where a1507 is not null order by a1501 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If .adoacc150.RecordCount <> 0 Then
         .adoacc150.Find "a1501 = '" & .Text2 & "'", 0, adSearchForward, 1
         If .adoacc150.EOF Then
            .adoacc150.MoveFirst
            Frmacc21j0_Clear
         End If
      Else
         Frmacc21j0_Clear
      End If
      
      'Add by Sindy 2018/2/22
      If .m_strIR01 <> "" Then
         'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", strCP09, "")
         'Modify By Sindy 2023/7/12 IIf(Pub_StrUserSt03 = "F22", strCP09, "") => IIf(Left(Pub_StrUserSt03, 1) = "F", strCP09, "")
         PUB_UpdateEMailRec .m_strIR01, .m_strIR02, .m_strIR03, .m_strIR04, "Frmacc21j0", IIf(Left(Pub_StrUserSt03, 1) = "F", strCP09, "")
      End If
      '2018/2/22 END
      
      adoTaie.CommitTrans 'Added by Lydia 2016/12/27 包在Transaction
      .AdodcRefresh
      .RecordShow
      
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   
   adoTaie.RollbackTrans 'Added by Lydia 2016/12/27
   MsgBox Err.Description, , MsgText(5)
      
   End With
End Sub

Public Sub Frmacc21k0_Save()
   On Error GoTo Checking
      
    'Added by Lydia 2021/12/08 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        strControlButton = MsgText(602) 'Added by Lydia 2021/12/29 參考Frmacc2150_Save
        Exit Sub
    End If
    'Added by Lydia 2021/12/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        If PUB_ChkUniText(Frmacc21k0, , True, "TextBox") = False Then
            strControlButton = MsgText(602)
            Exit Sub
        End If
    End If
    'end 2021/12/01
    
   With Frmacc21k0
      If .Text5 = MsgText(601) Then
         MsgBox MsgText(10) & .Label6, , MsgText(5)
         strControlButton = MsgText(602)
         .Text5.SetFocus
         Exit Sub
      Else
         If .MaskEdBox2.Text = MsgText(601) Or .MaskEdBox2.Text = MsgText(29) Then
            MsgBox .Label7 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox2.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox2.Text) = MsgText(603) Then
               MsgBox .Label7 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox2.SetFocus
               Exit Sub
            End If
         End If
      End If
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select * from acc1k0 where a1k01 = '" & .Text5 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If .adoquery.RecordCount = 0 Then
         .adoquery.Close
         Exit Sub
      End If
      
      'add by sonia 2018/2/22
      If Val("" & .adoquery.Fields("a1k12").Value) > 0 Or Val("" & .adoquery.Fields("a1k30").Value) > 0 Or "" & .adoquery.Fields("a1k25").Value <> "" Then
         MsgBox "此筆資料已銷帳或已收款或已作廢！", , MsgText(5)
         strControlButton = MsgText(602)
         .Text5.SetFocus
         .adoquery.Close
         Exit Sub
      End If
      'end 2018/2/22

      adoTaie.BeginTrans 'Added by Lydia 2016/12/27 包在Transaction
      
      If .MaskEdBox2.Text <> MsgText(601) And .MaskEdBox2.Text <> MsgText(29) Then
         .adoquery.Fields("a1k12").Value = Val(FCDate(.MaskEdBox2.Text))
      Else
         .adoquery.Fields("a1k12").Value = Null
      End If
      '2013/10/22 modify by sonia a1k05改為a1k34
      If .Text10 <> MsgText(601) Then
         .adoquery.Fields("a1k34").Value = .Text10
      Else
         .adoquery.Fields("a1k34").Value = Null
      End If
      .adoquery.UpdateBatch
      .adoquery.Close
      adoTaie.Execute "update caseprogress set cp60 = null where cp60 = '" & .Text5 & "'"
      adoTaie.Execute "delete acc1t0 where a1t01='" & .Text5 & "'"  'add by sonia 2018/2/27
      
      adoTaie.CommitTrans 'Added by Lydia 2016/12/27 包在Transaction
      .AdodcRefresh
      .RecordShow
       
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   
   adoTaie.RollbackTrans 'Added by Lydia 2016/12/27
   MsgBox Err.Description, , MsgText(5)
      
   End With
End Sub

Public Sub Frmacc21m0_Save(ByRef oForm As Form)
   On Error GoTo Checking
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21m0
   With oForm
      If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .MaskEdBox1.SetFocus
         Exit Sub
      Else
         If .Text5 = MsgText(601) Then
            MsgBox MsgText(10) & .Label2, , MsgText(5)
            strControlButton = MsgText(602)
            .Text5.SetFocus
            Exit Sub
         End If
      End If
      .adousxrate.CursorLocation = adUseClient
      .adousxrate.Open "select * from usxrate where usxr01 = " & Val(FCDate(.MaskEdBox1.Text)) & "", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If strSaveConfirm = MsgText(3) Then
         If .adousxrate.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            .adousxrate.Close
            .MaskEdBox1.SetFocus
            Exit Sub
         End If
         .adousxrate.AddNew
      Else
         If strSaveConfirm = MsgText(4) Then
            If .adousxrate.RecordCount = 0 Then
               MsgBox MsgText(28), , MsgText(5)
               strControlButton = MsgText(602)
               .adousxrate.Close
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
         End If
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adousxrate.Fields("usxr01").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adousxrate.Fields("usxr01").Value = Null
      End If
      If .Text5 <> MsgText(601) Then
         .adousxrate.Fields("usxr02").Value = Val(.Text5)
      Else
         .adousxrate.Fields("usxr02").Value = 0
      End If
      .adousxrate.UpdateBatch
      .adousxrate.Close
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'Add By Sindy 2009/06/06
Public Sub Frmacc21s0_Save(ByRef oForm As Form)
   On Error GoTo Checking
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21s0
   With oForm
      If .Combo1.Text = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Combo1.SetFocus
         Exit Sub
      ElseIf .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
         MsgBox MsgText(10) & .Label4, , MsgText(5)
         strControlButton = MsgText(602)
         .MaskEdBox1.SetFocus
         Exit Sub
      ElseIf .textDNR03 = MsgText(601) Then
         MsgBox MsgText(10) & .Label2, , MsgText(5)
         strControlButton = MsgText(602)
         .textDNR03.SetFocus
         Exit Sub
      Else
         '2011/3/31 modify by sonia 人民幣才一定要輸
         'If .textDNR04 = MsgText(601) Then
         If .textDNR04 = MsgText(601) And .Combo1.Text = "RMB" Then
            MsgBox MsgText(10) & .Label3, , MsgText(5)
            strControlButton = MsgText(602)
            .textDNR04.SetFocus
            Exit Sub
         End If
      End If
      .adousxrate.CursorLocation = adUseClient
      .adousxrate.Open "select * from debitnoterate where dnr01 = '" & .Combo1.Text & "' and dnr02 = " & Val(FCDate(.MaskEdBox1.Text)) & "", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If strSaveConfirm = MsgText(3) Then
         If .adousxrate.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            .adousxrate.Close
            .MaskEdBox1.SetFocus
            Exit Sub
         End If
         .adousxrate.AddNew
      Else
         If strSaveConfirm = MsgText(4) Then
            If .adousxrate.RecordCount = 0 Then
               MsgBox MsgText(28), , MsgText(5)
               strControlButton = MsgText(602)
               .adousxrate.Close
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
         End If
      End If
      If .Combo1.Text <> MsgText(601) Then
         .adousxrate.Fields("dnr01").Value = .Combo1.Text
      Else
         .adousxrate.Fields("dnr01").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adousxrate.Fields("dnr02").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adousxrate.Fields("dnr02").Value = Null
      End If
      If .textDNR03 <> MsgText(601) Then
         .adousxrate.Fields("dnr03").Value = Val(.textDNR03)
      Else
         .adousxrate.Fields("dnr03").Value = 0
      End If
      If .textDNR04 <> MsgText(601) Then
         .adousxrate.Fields("dnr04").Value = Val(.textDNR04)
      Else
         .adousxrate.Fields("dnr04").Value = 0
      End If
      .adousxrate.UpdateBatch
      .adousxrate.Close
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc21o0_Save(ByRef oForm As Form)

   Dim bolUpdate As Boolean, bolAddNew As Boolean
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21o0
   With oForm
      'Add by Morgan 2005/6/14
      If .FormCheck(bolUpdate, bolAddNew) = False Then Exit Sub

   adoTaie.BeginTrans

On Error GoTo Checking

   
      If bolAddNew = True Then
         .adoacc210.AddNew
      End If
      
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc210.Fields("a2101").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc210.Fields("a2101").Value = Null
      End If
      If .Combo1.Text <> MsgText(601) Then
         .adoacc210.Fields("a2102").Value = .Combo1
      Else
         .adoacc210.Fields("a2102").Value = Null
      End If
      If .Text5 <> MsgText(601) Then
         .adoacc210.Fields("a2103").Value = Val(.Text5)
      Else
         .adoacc210.Fields("a2103").Value = 0
      End If
      'Add by Morgan 2005/6/14
      .adoacc210.Fields("a2110").Value = Val(.txtBase)
      .adoacc210.UpdateBatch
      adoTaie.Execute "update acc210 set a2110=" & Val(.txtBase) & " where a2101=" & .adoacc210.Fields("a2101").Value
      .adoacc210.Close
      adoTaie.CommitTrans
      
Checking:

      If Err.NUMBER <> 0 Then
         adoTaie.RollbackTrans
         MsgBox Err.Description, , MsgText(5)
      Else
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub
