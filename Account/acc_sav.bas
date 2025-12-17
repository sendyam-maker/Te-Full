Attribute VB_Name = "aacc_sav"
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

'*************************************************
'  儲存資料表記錄
'
'*************************************************
'Modify by Morgan 2011/8/12 清除a0k12,a0k14,a0k15相關程式以便保留再使用
Public Sub Frmacc1110_Save()

   Dim intCounter As Integer
   Dim strAutoNo, strYes As String
      
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   
On Error GoTo Checking

   adoTaie.BeginTrans
   With Frmacc1110
      '專利
      If (.Text4 <> MsgText(601) And .Text4 <> "E") And (.Text5 <> MsgText(601) And .Text5 <> "E") Then
         For intCounter = Val(Mid(.Text4, 5, 5)) To Val(Mid(.Text5, 5, 5))
            strAutoNo = Mid(.Text5, 1, 4) & ZeroBeforeNo(intCounter - 1, 5)
            strSql = "insert into acc0k0 (a0k01, a0k02, a0k06, a0k07, a0k09, a0k11, a0k17, a0k18, a0k19, a0k20, a0k21, a0k24, a0k25, a0k26) " & _
               " select '" & strAutoNo & "', 0, 0, 0, 0, '2', 0, 0, 0, '" & .Text2 & "', '1', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "'" & _
               " from dual where not exists(select * from acc0k0 where a0k01='" & strAutoNo & "')"
            adoTaie.Execute strSql
         Next intCounter
      End If
      '商標
      If (.Text6 <> MsgText(601) And .Text6 <> "E") And (.Text7 <> MsgText(601) And .Text7 <> "E") Then
         For intCounter = Val(Mid(.Text6, 5, 5)) To Val(Mid(.Text7, 5, 5))
            strAutoNo = Mid(.Text7, 1, 4) & ZeroBeforeNo(intCounter - 1, 5)
            adoTaie.Execute "insert into acc0k0 (a0k01, a0k02, a0k06, a0k07, a0k09, a0k11, a0k17, a0k18, a0k19, a0k20, a0k21, a0k24, a0k25, a0k26) " & _
               " values ('" & strAutoNo & "', 0, 0, 0, 0, '1', 0, 0, 0, '" & .Text2 & "', '2', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "')"
         Next intCounter
      End If
      '法律
      If (.Text10 <> MsgText(601) And .Text10 <> "E") And (.Text11 <> MsgText(601) And .Text11 <> "E") Then
         For intCounter = Val(Mid(.Text10, 5, 5)) To Val(Mid(.Text11, 5, 5))
            strAutoNo = Mid(.Text11, 1, 4) & ZeroBeforeNo(intCounter - 1, 5)
            adoTaie.Execute "insert into acc0k0 (a0k01, a0k02, a0k06, a0k07, a0k09, a0k11, a0k17, a0k18, a0k19, a0k20, a0k21, a0k24, a0k25, a0k26) " & _
               " values ('" & strAutoNo & "', 0, 0, 0, 0, '" & .Text8 & "', 0, 0, 0, '" & .Text2 & "', '3', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "')"
         Next intCounter
      End If
      .AutoNoQuery
   End With
   adoTaie.CommitTrans
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   
Checking:
   If Err.Number <> 0 Then
      adoTaie.RollbackTrans
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub

Public Sub Frmacc1130_Save()
   On Error GoTo Checking
   With Frmacc1130
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label2 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
                MsgBox .Label2 & MsgText(63), , MsgText(5)
                strControlButton = MsgText(602)
                .MaskEdBox1.SetFocus
                Exit Sub
            End If
         End If
         If .Text1 <> MsgText(601) Then
            'Add By Sindy 2013/12/25 J公司請款單若已開發票則不可作廢
            If .Text9 = "J" Then
               If .adoquery.State = adStateOpen Then
                  .adoquery.Close
               End If
               .adoquery.CursorLocation = adUseClient
               If .strRelateNoList <> .Text1 Then
                  .adoquery.Open "select * from acc431 where axc02 in ( '" & Replace(.strRelateNoList, ",", "','") & "' )", adoTaie, adOpenStatic, adLockReadOnly
               Else
                  .adoquery.Open "select * from acc431 where axc02 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
               End If
               If .adoquery.RecordCount <> 0 Then
                   MsgBox "已開發票不可作廢...", , MsgText(5)
                   strControlButton = MsgText(602)
                   .Text1.SetFocus
                   .adoquery.Close
                   Exit Sub
               End If
               .adoquery.Close
            End If
            '2013/12/25 END
            
            'Add by Morgan 2011/9/21
            .strRelateNoList = PUB_GetRelateNo(.Text1)
            If .strRelateNoList <> .Text1 Then
               .Command1.Enabled = True
               If .ShowRelNo(True, True) = False Then
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
            End If
            'end 2011/9/21
            
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            'Modify by Morgan 2011/9/21 考慮拆收據情形
            If .strRelateNoList <> .Text1 Then
               .adoquery.Open "select * from acc0m0 where a0m02 in ( '" & Replace(.strRelateNoList, ",", "','") & "' )", adoTaie, adOpenStatic, adLockReadOnly
            Else
               .adoquery.Open "select * from acc0m0 where a0m02 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
            End If
            If .adoquery.RecordCount <> 0 Then
                MsgBox MsgText(203), , MsgText(5)
                strControlButton = MsgText(602)
                .Text1.SetFocus
                .adoquery.Close
                Exit Sub
            End If
            .adoquery.Close
            
            '2011/8/19 add by sonia 已銷帳不可作廢
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            'Modify by Morgan 2011/9/21 考慮拆收據情形
            If .strRelateNoList <> .Text1 Then
               .adoquery.Open "select * from acc0k0 where a0k01 in ( '" & Replace(.strRelateNoList, ",", "','") & "' ) and a0k10 is not null ", adoTaie, adOpenStatic, adLockReadOnly
            Else
               .adoquery.Open "select * from acc0k0 where a0k01 = '" & .Text1 & "' and a0k10 is not null ", adoTaie, adOpenStatic, adLockReadOnly
            End If
            If .adoquery.RecordCount <> 0 Then
                MsgBox "已銷帳不可作廢...", , MsgText(5)
                strControlButton = MsgText(602)
                .Text1.SetFocus
                .adoquery.Close
                Exit Sub
            End If
            .adoquery.Close
            '2011/8/19 end
            
            'Added by Morgan 2023/10/26
            '已繳款不可作廢
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            If .strRelateNoList <> .Text1 Then
               .adoquery.Open "select * from acc441 where axd04 in ( '" & Replace(.strRelateNoList, ",", "','") & "' )", adoTaie, adOpenStatic, adLockReadOnly
            Else
               .adoquery.Open "select * from acc441 where axd04 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
            End If
            If .adoquery.RecordCount <> 0 Then
                MsgBox "已繳款不可作廢...", , MsgText(5)
                strControlButton = MsgText(602)
                .Text1.SetFocus
                .adoquery.Close
                Exit Sub
            End If
            .adoquery.Close
            
            'end 2023/10/26
         End If
      End If
      
       'Add by Amy 2013/11/15 +Transation
       adoTaie.BeginTrans
       'Modify by Morgan 2011/9/21 程式相同,保留一段就好
       'If strSaveConfirm = MsgText(3) Then
         .adoacc0k0a.Close
         .adoacc0k0a.CursorLocation = adUseClient
         'Modify by Morgan 2011/9/21 考慮拆收據情形
         If .strRelateNoList <> .Text1 Then
            .adoacc0k0a.Open "select * from acc0k0 where a0k01 in ( '" & Replace(.strRelateNoList, ",", "','") & "' )", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Else
            .adoacc0k0a.Open "select * from acc0k0 where a0k01 = '" & .Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         End If
         If .adoacc0k0a.RecordCount <> 0 Then
            Do While Not .adoacc0k0a.EOF
               If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
                  .adoacc0k0a.Fields("a0k09").Value = Val(FCDate(.MaskEdBox1.Text))
                  .adoacc0k0a.Fields("a0k32").Value = Null   '2010/5/20 add by sonia
               Else
                  .adoacc0k0a.Fields("a0k09").Value = 0
               End If
               If .Text12 <> MsgText(601) Then
                  .adoacc0k0a.Fields("a0k08").Value = .Text12
               Else
                  .adoacc0k0a.Fields("a0k08").Value = Null
               End If
               .adoacc0k0a.Fields("a0k27").Value = Val(strSrvDate(2))
               .adoacc0k0a.Fields("a0k28").Value = ServerTime
               .adoacc0k0a.Fields("a0k29").Value = strUserNum
               .adoacc0k0a.MoveNext
            Loop
            .adoacc0k0a.UpdateBatch
         End If
'      Else
'         .adoacc0k0.Close
'         .adoacc0k0.CursorLocation = adUseClient
'         .adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & .Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         If .adoacc0k0.RecordCount <> 0 Then
'            If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
'               .adoacc0k0.Fields("a0k09").Value = Val(FCDate(.MaskEdBox1.Text))
'               .adoacc0k0a.Fields("a0k32").Value = Null   '2010/5/20 add by sonia
'            Else
'               .adoacc0k0.Fields("a0k09").Value = 0
'            End If
'            If .Text12 <> MsgText(601) Then
'               .adoacc0k0.Fields("a0k08").Value = .Text12
'            Else
'               .adoacc0k0.Fields("a0k08").Value = Null
'            End If
'            .adoacc0k0.Fields("a0k27").Value = Val(ACDate(ServerDate))
'            .adoacc0k0.Fields("a0k28").Value = ServerTime
'            .adoacc0k0.Fields("a0k29").Value = strUserNum
'            .adoacc0k0.UpdateBatch
'         End If
'      End If
      
      'Modify by Morgan 2011/9/21 考慮拆收據情形
      If .strRelateNoList <> .Text1 Then
         adoTaie.Execute "delete from acc0j0 where a0j13 in ( '" & Replace(.strRelateNoList, ",", "','") & "' )"
         'Modified by Morgan 2016/6/2 收據作廢時不可更新cp79=0,會影響自動收文之P案發文
         'adoTaie.Execute "update caseprogress set cp60 = '',CP73=0,CP74=0,CP75=0,CP77=0,CP78=0,CP79=0 where cp60 in ( '" & Replace(.strRelateNoList, ",", "','") & "' )"
         adoTaie.Execute "update caseprogress set cp60 = '',CP73=0,CP74=0,CP75=0,CP77=0,CP78=0,CP79=cp16 where cp60 in ( '" & Replace(.strRelateNoList, ",", "','") & "' )"
      Else
         adoTaie.Execute "delete from acc0j0 where a0j13 = '" & .Text1 & "'"
         'Modify by Morgan 2005/10/31 cp73,cp74,cp75,cp79,cp77,cp78 清成0
         'Modified by Morgan 2016/6/2 收據作廢時不可更新cp79=0,會影響自動收文之P案發文
         'adoTaie.Execute "update caseprogress set cp60 = '',CP73=0,CP74=0,CP75=0,CP77=0,CP78=0,CP79=0 where cp60 = '" & .Text1 & "'"
         adoTaie.Execute "update caseprogress set cp60 = '',CP73=0,CP74=0,CP75=0,CP77=0,CP78=0,CP79=cp16 where cp60 = '" & .Text1 & "'"
      End If
      adoTaie.CommitTrans 'Add by Amy 2013/11/15
      .Acc0k0Refresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   adoTaie.RollbackTrans 'Add by amy 2013/11/15
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc1160_Save()
   On Error GoTo Checking
   With Frmacc1160
      'Mark by Amy 2023/05/15 程式搬至frmacc1160
'      If .Text1 = MsgText(601) Then
'         MsgBox MsgText(10), , MsgText(5)
'         strControlButton = MsgText(602)
'         .Text1.SetFocus
'         Exit Sub
'      Else
'         If CheckLen(.Label2, .Text2, 40) = MsgText(603) Then
'            strControlButton = MsgText(602)
'            .Text2.SetFocus
'            Exit Sub
'         End If
'         'Add by Amy 2015/09/10 +有值才檢查
'         If Trim(.Text3) <> MsgText(601) Then
'            If CheckLen(.Label3, .Text3, 70) = MsgText(603) Then
'               strControlButton = MsgText(602)
'               .Text3.SetFocus
'               Exit Sub
'            '2011/10/18 add by sonia 檢查地址
'            ElseIf CheckTaiwanAddr(.Text3, "000", "地址") = False Then
'               strControlButton = MsgText(602)
'               .Text3.SetFocus
'               Exit Sub
'            '2011/10/18 end
'            End If
'         End If
         'end 2015/09/10
'         If CheckLen(.Label6, .Text4, 10) = MsgText(603) Then
'            strControlButton = MsgText(602)
'            .Text4.SetFocus
'            Exit Sub
'         End If
         
'         'Add by Morgan 2011/6/20
'         If .Text5 = "1" Then
'            If .Text15 = "" Then
'               MsgBox "電匯廠商的【" & .Label16 & "】欄位不可空白！"
'               strControlButton = MsgText(602)
'               .Text15.SetFocus
'               Exit Sub
'
'            ElseIf Len(.Text15) <> 7 Then
'               MsgBox "【" & .Label16 & "】欄位必須為 7 碼數字！"
'               strControlButton = MsgText(602)
'               .Text15.SetFocus
'               Exit Sub
'            End If
'
'            If .Text7 = "" Then
'               MsgBox "電匯廠商的【" & .Label9 & "】欄位不可空白！"
'               strControlButton = MsgText(602)
'               .Text7.SetFocus
'               Exit Sub
'
'            ElseIf Len(.Text7) <> 14 Then
'               MsgBox "【" & .Label9 & "】欄位必須為 14 碼數字！"
'               strControlButton = MsgText(602)
'               .Text7.SetFocus
'               Exit Sub
'            End If
'
'         End If
'         'end 2011/6/20
'         'Add by Amy 2014/01/07
'         If .TxtValidate = False Then
'             strControlButton = MsgText(602)
'             Exit Sub
'         End If
'         'end 2014/01/07
         'Add by Morgan 2007/8/20
'         If Left(.Text1, 1) = "F" Then
'            If .Text3 = "" Then
'               If MsgBox("此編號為翻譯人員但未輸入" & .Label3 & "，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
'                  strControlButton = MsgText(602)
'                  .Text3.SetFocus
'                  Exit Sub
'               End If
'            End If
'
''Remove by Morgan 2012/3/26 改以員工檔為主廠商檔不可修改(Trigger同步)
''            'Add by Morgan 2010/4/23
''            MsgBox "若扣單地址也需修改時請通知人事做相同修改！"
'
''Remove by Morgan 2010/1/11 證明單不用印故不必再控制--瑞婷
''            If .Text11 = "" Then
''               If MsgBox("此編號為翻譯人員但未輸入" & .Label13 & "，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
''                  strControlButton = MsgText(602)
''                  .Text11.SetFocus
''                  Exit Sub
''               End If
''            End If
'         End If
'         'End 2007/8/20
'      End If
      If strSaveConfirm = MsgText(3) Then
'         If .Adodc1.Recordset.RecordCount <> 0 Then
'            .Adodc1.Recordset.Find "a0i01 = '" & .Text1 & "'", 0, adSearchForward, 1
'            If .Adodc1.Recordset.EOF = False Then
'               MsgBox "廠商資料已存在！", vbExclamation
'               Exit Sub
'            End If
'         End If
'end 2023/05/15
         .Adodc1.Recordset.AddNew
      End If
      .Adodc1.Recordset.Fields("a0i01").Value = .Text1
      If .Text2 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i02").Value = .Text2
      Else
         .Adodc1.Recordset.Fields("a0i02").Value = Null
      End If
      If .Text3 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i03").Value = .Text3
      Else
         .Adodc1.Recordset.Fields("a0i03").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i05").Value = .MaskEdBox1.Text
      Else
         .Adodc1.Recordset.Fields("a0i05").Value = Null
      End If
      If .MaskEdBox2.Text <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i07").Value = .MaskEdBox2.Text
      Else
         .Adodc1.Recordset.Fields("a0i07").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i08").Value = .Text4
      Else
         .Adodc1.Recordset.Fields("a0i08").Value = Null
      End If
      'Add by Morgan 2006/6/12
      If .Text5 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i12").Value = .Text5
      Else
         .Adodc1.Recordset.Fields("a0i12").Value = Null
      End If
      If .Text6 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i13").Value = .Text6
      Else
         .Adodc1.Recordset.Fields("a0i13").Value = Null
      End If
      If .Text7 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i14").Value = .Text7
      Else
         .Adodc1.Recordset.Fields("a0i14").Value = Null
      End If
      If .Text8 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i15").Value = .Text8
      Else
         .Adodc1.Recordset.Fields("a0i15").Value = Null
      End If
      If .Text9 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i16").Value = .Text9
      Else
         .Adodc1.Recordset.Fields("a0i16").Value = Null
      End If
      'end 2006/6/12
      'Add by Morgan 2007/2/5
      If .Text10 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i17").Value = .Text10
      Else
         .Adodc1.Recordset.Fields("a0i17").Value = Null
      End If
      'end 2007/2/5
      'Add by Morgan 2007/6/6
      If .Text11 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i18").Value = .Text11
      Else
         .Adodc1.Recordset.Fields("a0i18").Value = Null
      End If
      'end 2007/6/6
      
      'Add by Morgan 2007/12/24
      If .Text12 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i19").Value = .Text12
      Else
         .Adodc1.Recordset.Fields("a0i19").Value = Null
      End If
      
      'Add by Morgan 2009/1/21
      If .Text13 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i06").Value = .Text13
      Else
         .Adodc1.Recordset.Fields("a0i06").Value = Null
      End If
      If .Text14 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i04").Value = .Text14
      Else
         .Adodc1.Recordset.Fields("a0i04").Value = Null
      End If
      
      'Add by Morgan 2011/6/20
      If .Text15 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0i20").Value = .Text15
      Else
         .Adodc1.Recordset.Fields("a0i20").Value = Null
      End If
      
      .Adodc1.Recordset.UpdateBatch
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc11f0_Save()
   On Error GoTo Checking
   With Frmacc11f0
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      End If
      adoTaie.BeginTrans 'Added by Lydia 2016/12/27 包在Transaction
      adoTaie.Execute "update acc1v0 set a1v14 = null, a1v15 = null where a1v15 = '" & .Text1 & "'"
      adoTaie.Execute "update acc0w0 set a0w15 = " & Val(FCDate(.MaskEdBox1.Text)) & " where a0w02 = '" & .Text1 & "'"
      adoTaie.Execute "delete from acc1p0 where substr(a1p04, 1, 9) = '" & .Text1 & "'"
      adoTaie.CommitTrans 'Added by Lydia 2016/12/27 包在Transaction
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   adoTaie.RollbackTrans 'Added by Lydia 2016/12/27
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc2110_Save()
   'Added by Morgan 2021/5/27
   Dim arrA1p01() As String '傳票公司別
   Dim arrA1p22() As String '傳票號
   Dim intPos As Integer
   'end 2021/5/27
   Dim strMsg As String 'Added by Morgan 2022/6/29

    'Added by Lydia 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        If PUB_ChkUniText(Frmacc2110, , True, "TextBox") = False Then
            strControlButton = MsgText(602)
            Exit Sub
        End If
    End If
    'end 2021/12/03
    
On Error GoTo Checking
   With Frmacc2110
      If .Text2 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text2.SetFocus
         Exit Sub
      Else
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label2 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox .Label2 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
            
            'Added by Morgan 2022/6/29
            If .MaskEdBox1.Enabled = True Then
               If ChkWorkData("1", DBDATE(.MaskEdBox1), strMsg) = False Then
                   MsgBox .Label2 & strMsg, , MsgText(5)
                   strControlButton = MsgText(602)
                  .MaskEdBox1.SetFocus
                   Exit Sub
               End If
            End If
            'end 2022/6/29
         End If
         'Add by Amy 2021/01/29  匯率未輸會錯
         If Trim(.Text3) = MsgText(601) Then
            MsgBox "匯率不可為空！", , MsgText(5)
            strControlButton = MsgText(602)
            .Text3.SetFocus
            Exit Sub
         ElseIf IsNumeric(.Text3) = False Then
            MsgBox MsgText(130), , MsgText(5)
            strControlButton = MsgText(602)
            .Text3.SetFocus
            Exit Sub
         End If
      End If
      If strSaveConfirm = MsgText(4) Then
         '.CreDebCheck 'Removed by Morgan 2021/5/31
         If .CreDebCheck <> MsgText(602) Then
            MsgBox MsgText(11), , MsgText(5)
            strControlButton = MsgText(602)
            Exit Sub
         End If
      End If
      
      If strSaveConfirm = MsgText(3) Then
         If .adoacc0y0.RecordCount <> 0 Then
            .adoacc0y0.Find "a0y01 = '" & .Text2 & "'", 0, adSearchForward, 1
            If .adoacc0y0.EOF = False Then
               Exit Sub
            End If
         End If
         .adoacc0y0.AddNew
      End If
      
      .adoacc0y0.Fields("a0y01").Value = .Text2
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc0y0.Fields("a0y02").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc0y0.Fields("a0y02").Value = Null
      End If
      If .Combo4 <> MsgText(601) Then
         .adoacc0y0.Fields("a0y03").Value = .Combo4
      Else
         .adoacc0y0.Fields("a0y03").Value = Null
      End If
      'Add By Sindy 2013/1/31
      '因使用者會在第1畫面修改幣別後,直接離開畫面沒繼續進入明細,所以在此重新再更新a0z03幣別欄位值
      adoTaie.Execute "update acc0z0 set a0z03=" & IIf(.Combo4 <> MsgText(601), CNULL(.Combo4), CNULL("")) & " where a0z01='" & .Text2 & "'"
      '2013/1/31 End
      If Val(.Text3) <> .adoacc0y0.Fields("a0y04").Value Then
         adoTaie.Execute "update acc1p0 set a1p20 = " & Val(.Text3) & ", a1p07 = decode(a1p07, 0, 0, round(a1p21 * " & Val(.Text3) & ", 2)), a1p08 = decode(a1p08, 0, 0, round(a1p21 * " & Val(.Text3) & ", 2)) where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & .Text2 & "' and substr(a1p05, 1, 2) <> '22'"
      End If
      If .Text3 <> MsgText(601) Then
         .adoacc0y0.Fields("a0y04").Value = Val(.Text3)
      Else
         .adoacc0y0.Fields("a0y04").Value = 0
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc0y0.Fields("a0y11").Value = .Text4
      Else
         .adoacc0y0.Fields("a0y11").Value = Null
      End If
      If strSaveConfirm = MsgText(3) Then
         .adoacc0y0.Fields("a0y12").Value = Val(strSrvDate(2))
         .adoacc0y0.Fields("a0y13").Value = ServerTime
         .adoacc0y0.Fields("a0y14").Value = strUserNum
      Else
         .adoacc0y0.Fields("a0y15").Value = Val(strSrvDate(2))
         .adoacc0y0.Fields("a0y16").Value = ServerTime
         .adoacc0y0.Fields("a0y17").Value = strUserNum
      End If
      .adoacc0y0.UpdateBatch
      If strSaveConfirm <> MsgText(3) Then
         'Modified by Morgan 2021/5/27 目前會有1或L兩家公司
         'If strCon10 <> "" Then
         '   adoTaie.Execute "update acc1p0 set a1p22 = '" & strCon10 & "', a1p27 = 'Y', a1p18 = " & Val(FCDate(.MaskEdBox1.Text)) & " where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & .Text2 & "'"
         'End If
         If .strA1P01s <> "" Then
            arrA1p01 = Split(.strA1P01s, ";")
            arrA1p22 = Split(.strA1P22s, ";")
            For intPos = LBound(arrA1p22) To UBound(arrA1p22)
               If arrA1p22(intPos) <> "" Then
                  adoTaie.Execute "update acc1p0 set a1p22 = '" & arrA1p22(intPos) & "', a1p27 = 'Y' where a1p01 = '" & arrA1p01(intPos) & "' and a1p02 = 'F' and a1p04 = '" & .Text2 & "'", intI
               End If
            Next
         End If
         'end 2021/5/27
      End If
        'Add By Cheng 2004/04/27
        '若未產生傳票則更新A1P18
        'Modified by Morgan 2021/5/26 取消 a1p01 = '1'條件,目前會有1或L兩家公司
        adoTaie.Execute "update acc1p0 set a1p18 = " & Val(FCDate(.MaskEdBox1.Text)) & " where a1p02 = 'F' and a1p04 = '" & .Text2 & "' And A1P22 is Null "
        'End
      
      .RecordShow
'      .AdodcRefresh
'      .SumShow
      If strSaveConfirm <> MsgText(3) Then
         .adoaccsum.CursorLocation = adUseClient
         .adoaccsum.Open "select sum(a1p08) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & .Text2 & "' and a1p03 not in (select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & .Text2 & "' and a1p08 <> 0)", adoTaie, adOpenStatic, adLockReadOnly
         If .adoaccsum.RecordCount <> 0 Then
         '   adoTaie.Execute "update acc1p0 set a1p08 = " & Val(.Text5) - Val(IIf(IsNull(.adoaccsum.Fields(0).Value), 0, .adoaccsum.Fields(0).Value)) & " where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & .Text2 & "' and a1p03 in (select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & .Text2 & "' and a1p08 <> 0  and substr(a1p05, 1, 2) <> '22')"
         End If
         .adoaccsum.Close
         .SumShow
      End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If

   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc2111_Save()
    
   On Error GoTo Checking
   With Frmacc2111
      If .Text3 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .Text3.SetFocus
         Exit Sub
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc0y0.Fields("a0y07").Value = .Text4
      Else
         .adoacc0y0.Fields("a0y07").Value = Null
      End If
      If .Text6 <> MsgText(601) Then
         .adoacc0y0.Fields("a0y08").Value = .Text6
      Else
         .adoacc0y0.Fields("a0y08").Value = Null
      End If
      If .Text8 <> MsgText(601) Then
         .adoacc0y0.Fields("a0y09").Value = .Text8
      Else
         .adoacc0y0.Fields("a0y09").Value = Null
      End If
      If .Text10 <> MsgText(601) Then
         .adoacc0y0.Fields("a0y10").Value = .Text10
      Else
         .adoacc0y0.Fields("a0y10").Value = Null
      End If
      If .Text2 <> MsgText(601) Then
         .adoacc0y0.Fields("a0y06").Value = Val(Replace(.Text2, ",", ""))
      Else
         .adoacc0y0.Fields("a0y06").Value = 0
      End If
      If .Option1.Value = True Then
         .adoacc0y0.Fields("a0y18").Value = 1
      Else
         If .Option2.Value = True Then
            .adoacc0y0.Fields("a0y18").Value = 2
         Else
            .adoacc0y0.Fields("a0y18").Value = 3
         End If
      End If
      .adoacc0y0.Fields("a0y15").Value = Val(strSrvDate(2))
      .adoacc0y0.Fields("a0y16").Value = ServerTime
      .adoacc0y0.Fields("a0y17").Value = strUserNum
      'Add By Sindy 2013/4/29
      .adoacc0y0.Fields("a0y19").Value = .Text14
      '2013/4/29 End
      .adoacc0y0.UpdateBatch

Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc2130_Save()

    'Added by Lydia 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        If PUB_ChkUniText(Frmacc2130, , True, "TextBox") = False Then
            strControlButton = MsgText(602)
            Exit Sub
        End If
    End If
    'end 2021/12/03

   On Error GoTo Checking
   With Frmacc2130
      .Acc120Query
      If .Text2 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
'         If ExistCheck("acc010", "a0101", .Text5, .Label5) = False Then
'            strControlButton = MsgText(602)
'            .Text1.SetFocus
'            Exit Sub
'         End If
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label3 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox .Label3 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
         End If
      End If
      If strSaveConfirm = MsgText(3) Then
         If .Acc130Query Then
            strControlButton = MsgText(602)
            .Text1.SetFocus
            Exit Sub
         End If
         If .Adodc1.Recordset.RecordCount <> 0 Then
            .Adodc1.Recordset.Find "a1301 = '" & .Text2 & "'", 0, adSearchForward, 1
            If .Adodc1.Recordset.EOF = False Then
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               .Text2.SetFocus
               Exit Sub
            End If
         End If
         .Adodc1.Recordset.AddNew
      End If
      .Adodc1.Recordset.Fields("a1301").Value = .Text2
      If .Text1 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1303").Value = .Text1
      Else
         .Adodc1.Recordset.Fields("a1303").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .Adodc1.Recordset.Fields("a1302").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .Adodc1.Recordset.Fields("a1302").Value = Null
      End If
      If .Text3 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1304").Value = .Text3
      Else
         .Adodc1.Recordset.Fields("a1304").Value = Null
      End If
      If .Text5 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1305").Value = .Text5
      Else
         .Adodc1.Recordset.Fields("a1305").Value = Null
      End If
      If .Combo1 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1306").Value = .Combo1
      Else
         .Adodc1.Recordset.Fields("a1306").Value = Null
      End If
      If .Text10 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1309").Value = Val(.Text10)
      Else
         .Adodc1.Recordset.Fields("a1309").Value = 0
      End If
      If .Text7 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1307").Value = Val(.Text7)
      Else
         .Adodc1.Recordset.Fields("a1307").Value = 0
      End If
      If .Text8 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1308").Value = .Text8
      Else
         .Adodc1.Recordset.Fields("a1308").Value = Null
      End If
      If strSaveConfirm = MsgText(3) Then
         .Adodc1.Recordset.Fields("a1311").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a1312").Value = ServerTime
         .Adodc1.Recordset.Fields("a1313").Value = strUserNum
      Else
         .Adodc1.Recordset.Fields("a1314").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a1315").Value = ServerTime
         .Adodc1.Recordset.Fields("a1316").Value = strUserNum
      End If
'      If .Text5 <> MsgText(601) Then
   '貸方------------------------------------------------
'         .adoacc1p0.CursorLocation = adUseClient
'         .adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'H' and a1p04 = '" & .Text2 & "' and a1p05 = '" & .Text5 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         If .adoacc1p0.RecordCount = 0 Then
'            .adoacc1p0.AddNew
'         End If
'         .adoacc1p0.Fields("a1p01").Value = "1"
'         .adoacc1p0.Fields("a1p02").Value = "H"
'         .adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'H' and a1p04 = '" & .Text2 & "'", 3)
'         .adoacc1p0.Fields("a1p04").Value = .Text2
'         .adoacc1p0.Fields("a1p05").Value = .Text5
'         .adoacc1p0.Fields("a1p06").Value = MsgText(55)
'         If .Text1 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p23").Value = .Text1
'         Else
'            .adoacc1p0.Fields("a1p23").Value = Null
'         End If
'         If .Combo1 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p19").Value = .Combo1
'         Else
'            .adoacc1p0.Fields("a1p19").Value = Null
'         End If
'         If .Text10 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p20").Value = Val(.Text10)
'         Else
'            .adoacc1p0.Fields("a1p20").Value = 0
'         End If
'         If .Text7 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p21").Value = Val(.Text7)
'            .adoacc1p0.Fields("a1p08").Value = Val(Format(Val(.Text7) * Val(.Text10), FAmount))
'         Else
'            .adoacc1p0.Fields("a1p21").Value = 0
'            .adoacc1p0.Fields("a1p08").Value = 0
'         End If
'         .adoacc1p0.Fields("a1p07").Value = 0
'         If .Text8 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p14").Value = .Text8
'         Else
'            .adoacc1p0.Fields("a1p14").Value = .Combo1 & " " & .Text7
'         End If
'         If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
'            .adoacc1p0.Fields("a1p18").Value = Val(FCDate(.MaskEdBox1.Text))
'         Else
'            .adoacc1p0.Fields("a1p18").Value = Null
'         End If
'         If IsNull(.adoacc1p0.Fields("a1p22").Value) = False Then
'            .adoacc1p0.Fields("a1p27").Value = MsgText(602)
'         End If
'         If strSaveConfirm <> MsgText(3) And IsNull(.adoacc1p0.Fields("a1p27").Value) = False Then
'            .adoacc1p0.Fields("a1p27").Value = MsgText(602)
'         End If
'         .adoacc1p0.UpdateBatch
'         .adoacc1p0.Close
   '借方------------------------------------------------
'         .adoacc1p0.CursorLocation = adUseClient
'         .adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'H' and a1p04 = '" & .Text2 & "' and a1p05 = '2401'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         If .adoacc1p0.RecordCount = 0 Then
'            .adoacc1p0.AddNew
'         End If
'         .adoacc1p0.Fields("a1p01").Value = "1"
'         .adoacc1p0.Fields("a1p02").Value = "H"
'         .adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'H' and a1p04 = '" & .Text2 & "'", 3)
'         .adoacc1p0.Fields("a1p04").Value = .Text2
'         .adoacc1p0.Fields("a1p05").Value = "2401"
'         .adoacc1p0.Fields("a1p06").Value = MsgText(55)
'         If .Text1 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p23").Value = .Text1
'         Else
'            .adoacc1p0.Fields("a1p23").Value = Null
'         End If
'         If .Combo1 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p19").Value = .Combo1
'         Else
'            .adoacc1p0.Fields("a1p19").Value = Null
'         End If
'         If .Text10 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p20").Value = Val(.Text10)
'         Else
'            .adoacc1p0.Fields("a1p20").Value = 0
'         End If
'         If .Text7 <> MsgText(601) Then
'            .adoacc1p0.Fields("a1p21").Value = Val(.Text7)
'            .adoacc1p0.Fields("a1p07").Value = Val(Format(.douAmount, FAmount))
'         Else
'            .adoacc1p0.Fields("a1p21").Value = 0
'            .adoacc1p0.Fields("a1p07").Value = 0
'         End If
'         .adoacc1p0.Fields("a1p08").Value = 0
'         .adoacc1p0.Fields("a1p14").Value = .strCurrency & " " & .Text7 & "/" & Val(FCDate(.MaskEdBox1.Text)) & "/" & .Text4
'         If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
'            .adoacc1p0.Fields("a1p18").Value = Val(FCDate(.MaskEdBox1.Text))
'         Else
'            .adoacc1p0.Fields("a1p18").Value = Null
'         End If
'         If IsNull(.adoacc1p0.Fields("a1p22").Value) = False Then
'            .adoacc1p0.Fields("a1p27").Value = MsgText(602)
'         End If
'         If strSaveConfirm <> MsgText(3) And IsNull(.adoacc1p0.Fields("a1p27").Value) = False Then
'            .adoacc1p0.Fields("a1p27").Value = MsgText(602)
'         End If
'         .adoacc1p0.UpdateBatch
'         .adoacc1p0.Close
   '匯兌損益------------------------------------------------
'         If .douAmount <> Val(.Text11) Then
'            .adoacc1p0.CursorLocation = adUseClient
'            .adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'H' and a1p04 = '" & .Text2 & "' and a1p05 = '7128'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            If .adoacc1p0.RecordCount = 0 Then
'               .adoacc1p0.AddNew
'            End If
'            .adoacc1p0.Fields("a1p01").Value = "1"
'            .adoacc1p0.Fields("a1p02").Value = "H"
'            .adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'H' and a1p04 = '" & .Text2 & "'", 3)
'            .adoacc1p0.Fields("a1p04").Value = .Text2
'            .adoacc1p0.Fields("a1p05").Value = "7128"
'            .adoacc1p0.Fields("a1p06").Value = MsgText(55)
'            If .Text1 <> MsgText(601) Then
'               .adoacc1p0.Fields("a1p23").Value = .Text1
'            Else
'               .adoacc1p0.Fields("a1p23").Value = Null
'            End If
'            If .Combo1 <> MsgText(601) Then
'               .adoacc1p0.Fields("a1p19").Value = .Combo1
'            Else
'               .adoacc1p0.Fields("a1p19").Value = Null
'            End If
'            If .Text10 <> MsgText(601) Then
'               .adoacc1p0.Fields("a1p20").Value = Val(.Text10)
'            Else
'               .adoacc1p0.Fields("a1p20").Value = 0
'            End If
'            If Val(Format(.douAmount, FAmount)) > Val(.Text11) Then
'               .adoacc1p0.Fields("a1p08").Value = Val(Format(Val(Format(.douAmount, FAmount)) - Val(Format(Val(.Text7) * Val(.Text10), FAmount)), FAmount))
'               .adoacc1p0.Fields("a1p07").Value = 0
'            Else
'               .adoacc1p0.Fields("a1p07").Value = Val(Format(Val(Format(Val(.Text7) * Val(.Text10), FAmount)) - Val(Format(.douAmount, FAmount)), FAmount))
'               .adoacc1p0.Fields("a1p08").Value = 0
'            End If
'            .adoacc1p0.Fields("a1p14").Value = MsgText(90)
'            If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
'               .adoacc1p0.Fields("a1p18").Value = Val(FCDate(.MaskEdBox1.Text))
'            Else
'               .adoacc1p0.Fields("a1p18").Value = Null
'            End If
'            If IsNull(.adoacc1p0.Fields("a1p22").Value) = False Then
'               .adoacc1p0.Fields("a1p27").Value = MsgText(602)
'            End If
'            If strSaveConfirm <> MsgText(3) And IsNull(.adoacc1p0.Fields("a1p27").Value) = False Then
'               .adoacc1p0.Fields("a1p27").Value = MsgText(602)
'            End If
'            .adoacc1p0.UpdateBatch
'            .adoacc1p0.Close
'         Else
'            adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'H' and a1p04 = '" & .Text2 & "' and a1p05 = '7128'"
'         End If
'      End If
      .Adodc1.Recordset.UpdateBatch
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc2140_Save()

    'Added by Lydia 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        If PUB_ChkUniText(Frmacc2140, , True, "TextBox") = False Then
            strControlButton = MsgText(602)
            Exit Sub
        End If
    End If
    'end 2021/12/03
    
   On Error GoTo Checking
   With Frmacc2140
      If .Text2 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .Text2.SetFocus
         Exit Sub
      Else
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label2 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox .Label2 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
         End If
      End If
      'Add By Sindy 2011/8/9
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      End If
      '2011/8/9 End
      If strSaveConfirm = MsgText(3) Then
         '2006/7/18 ADD BY SONIA
         If .Text3 = MsgText(601) Then
            MsgBox .Label3 & MsgText(28), , MsgText(5)
            strControlButton = MsgText(602)
            .Text1.SetFocus
            Exit Sub
         End If
         '2006/7/18 END
         If .Adodc1.Recordset.RecordCount <> 0 Then
            .Adodc1.Recordset.Find "a1401 = '" & .Text2 & "'", 0, adSearchForward, 1
            If .Adodc1.Recordset.EOF = False Then
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               .Text2.SetFocus
               Exit Sub
            End If
         End If
         '.Adodc1.Recordset.AddNew 'Remove by Lydia 2016/12/27
      End If
      
      adoTaie.BeginTrans 'Added by Lydia 2016/12/27 包在Transaction
      
      If strSaveConfirm = MsgText(3) Then .Adodc1.Recordset.AddNew 'Added by Lydia 2016/12/27 從上面移下來
      
      .Adodc1.Recordset.Fields("a1401").Value = .Text2
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .Adodc1.Recordset.Fields("a1402").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .Adodc1.Recordset.Fields("a1402").Value = Null
      End If
      If .Text1 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1403").Value = .Text1
      Else
         .Adodc1.Recordset.Fields("a1403").Value = Null
      End If
      If .Text6 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1404").Value = .Text6
      Else
         .Adodc1.Recordset.Fields("a1404").Value = Null
      End If
      If .Text3 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("A1411").Value = .Text3
      End If
      If .Text4 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("A1412").Value = Val(.Text4)
      Else
         .Adodc1.Recordset.Fields("A1412").Value = 0
      End If
      If .Text5 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("A1413").Value = .Text5
      End If
      If strSaveConfirm = MsgText(3) Then
         .Adodc1.Recordset.Fields("a1405").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a1406").Value = ServerTime
         .Adodc1.Recordset.Fields("a1407").Value = strUserNum
      Else
         .Adodc1.Recordset.Fields("a1408").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a1409").Value = ServerTime
         .Adodc1.Recordset.Fields("a1410").Value = strUserNum
      End If
      adoTaie.Execute "update acc1k0 set a1k25 = '" & .Text2 & "' where a1k01 = '" & .Text1 & "'"
      'Add By Sindy 2010/12/7 修改請款編號時, 原請款編號之A1K25一併清除
      If Trim(.Text1) <> "" And Trim(.m_Old_a1k01) <> "" Then 'Add By Sindy 2011/8/10
         If Trim(.Text1) <> Trim(.m_Old_a1k01) Then
            adoTaie.Execute "update acc1k0 set a1k25 = null where a1k01 = '" & .m_Old_a1k01 & "'"
         End If
      End If
      '2010/12/7 End
      .Adodc1.Recordset.UpdateBatch
      
      adoTaie.CommitTrans 'Added by Lydia 2016/12/27 包在Transaction
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   
   adoTaie.RollbackTrans 'Added by Lydia 2016/12/27
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc21n0_Save()
   On Error GoTo Checking
   With Frmacc21n0
      If .Combo1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Combo1.SetFocus
         Exit Sub
      Else
         If .Text5 = MsgText(601) Then
            MsgBox MsgText(10) & .Label2, , MsgText(5)
            strControlButton = MsgText(602)
            .Text5.SetFocus
            Exit Sub
         End If
      End If
      .adoacc1x0.CursorLocation = adUseClient
      .adoacc1x0.Open "select * from acc1x0 where a1x01 = '" & .Combo1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If strSaveConfirm = MsgText(3) Then
         If .adoacc1x0.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            .adoacc1x0.Close
            .Combo1.SetFocus
            Exit Sub
         End If
         .adoacc1x0.AddNew
      Else
         If strSaveConfirm = MsgText(4) Then
            If .adoacc1x0.RecordCount = 0 Then
               MsgBox MsgText(28), , MsgText(5)
               strControlButton = MsgText(602)
               .adoacc1x0.Close
               .Combo1.SetFocus
               Exit Sub
            End If
         End If
      End If
      If .Combo1.Text <> MsgText(601) Then
         .adoacc1x0.Fields("a1x01").Value = .Combo1
      Else
         .adoacc1x0.Fields("a1x01").Value = Null
      End If
      If .Text5 <> MsgText(601) Then
         .adoacc1x0.Fields("a1x02").Value = Val(.Text5)
      Else
         .adoacc1x0.Fields("a1x02").Value = 0
      End If
      .adoacc1x0.UpdateBatch
      .adoacc1x0.Close
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub
'Remove by Morgan 2005/8/3 財務不用
'Public Sub Frmacc21o0_Save()
'   On Error GoTo Checking
'   With Frmacc21o0
'      If .Combo1 = MsgText(601) Then
'         MsgBox MsgText(10) & .Label1, , MsgText(5)
'         strControlButton = MsgText(602)
'         .Combo1.SetFocus
'         Exit Sub
'      Else
'         If .Text5 = MsgText(601) Then
'            MsgBox MsgText(10) & .Label2, , MsgText(5)
'            strControlButton = MsgText(602)
'            .Text5.SetFocus
'            Exit Sub
'         End If
'      End If
'      .adoacc210.CursorLocation = adUseClient
'      .adoacc210.Open "select * from acc210 where a2101 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a2102 = '" & .Combo1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'      If strSaveConfirm = MsgText(3) Then
'         If .adoacc210.RecordCount <> 0 Then
'            MsgBox MsgText(9), , MsgText(5)
'            strControlButton = MsgText(602)
'            .adoacc210.Close
'            .Combo1.SetFocus
'            Exit Sub
'         End If
'         .adoacc210.AddNew
'      Else
'         If strSaveConfirm = MsgText(4) Then
'            If .adoacc210.RecordCount = 0 Then
'               MsgBox MsgText(28), , MsgText(5)
'               strControlButton = MsgText(602)
'               .adoacc210.Close
'               .Combo1.SetFocus
'               Exit Sub
'            End If
'         End If
'      End If
'      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
'         .adoacc210.Fields("a2101").Value = Val(FCDate(.MaskEdBox1.Text))
'      Else
'         .adoacc210.Fields("a2101").Value = Null
'      End If
'      If .Combo1.Text <> MsgText(601) Then
'         .adoacc210.Fields("a2102").Value = .Combo1
'      Else
'         .adoacc210.Fields("a2102").Value = Null
'      End If
'      If .Text5 <> MsgText(601) Then
'         .adoacc210.Fields("a2103").Value = Val(.Text5)
'      Else
'         .adoacc210.Fields("a2103").Value = 0
'      End If
'      .adoacc210.UpdateBatch
'      .adoacc210.Close
'      .AdodcRefresh
'      .RecordShow
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'   End With
'End Sub

'Add By Cheng 2003/07/22
'Added by Lydia 2015/03/30 改funciton判斷是否寫入
Public Function Frmacc21q0_Save() As Boolean
   Dim oControl As Control
   Dim strA As String
   Dim strT As String 'Added by Lydia 2016/04/13
   'Added by Lydia 2016/06/29
   Dim Rs21q0 As New ADODB.Recordset
   Dim inX As Integer
   Dim bolUpd As Boolean
   'Added by Lydia 2017/09/12
   Dim tmpBol As Boolean
   
    'Added by Lydia 2021/12/02 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        Exit Function
    End If
    'Added by Lydia 2021/12/02 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        If PUB_ChkUniText(Frmacc21q0, , True, "TextBox") = False Then
            strControlButton = MsgText(602)
            Exit Function
        End If
    End If
    'end 2021/12/02
    
   On Error GoTo Checking
   With Frmacc21q0
      'Added by Morgan 2013/3/14
      For Each oControl In .Controls
         If TypeName(oControl) = "TextBox" Then
            If oControl.MaxLength > 0 And oControl.Enabled = True Then
               If StrLength(oControl) > oControl.MaxLength Then
                  MsgBox "超過欄位長度 " & oControl.MaxLength & " 個半形字，請修正！", vbCritical
                  strControlButton = MsgText(602)
                  oControl.SetFocus
                  Exit Function
               End If
            End If
         End If
      Next
      'end 2013/3/14
      'Added By Lydia 2015/03/24 剔除跳行符號+不過濾的文字框.name
      PUB_FilterFormText Frmacc21q0, "Text22"
      
    'Modified by Lydia 2015/03/27 針對台銀媒體檔格式,受款人名稱+地址最長140字
     strA = Trim(.Text3.Text) & IIf(Len(.Text9) > 0, " " & Trim(.Text9), "") & IIf(Len(.Text2) > 0, " " & Trim(.Text2), "")
     'Added by Lydia 2017/08/07 提醒名稱+原地址
     If Trim(.Text21) = "" And GetTextLength_1(Trim(strA & " " & Trim(.Text17))) > 140 Then
        MsgBox "請注意受款人名稱+原地址超過140字(中文字=2個字)!!", vbInformation
     End If
     'end 2017/08/07
     
     'Modified by Lydia 2017/09/12 改用模組
     'strA = strA & " " & Trim(.Text21)
     'If GetTextLength_1(Trim(strA)) > 140 Then
     '   'Modified by Lydia 2017/08/07　(中文字=2個字)
     '   MsgBox "受款人名稱+短地址最長140字(中文字=2個字),請修改!!", vbExclamation
     '   Exit Function
     'End If
     tmpBol = False
     .Text21_Validate tmpBol
     If tmpBol = True Then
        Exit Function
     End If
     'end 2017/09/12
     
     'Added by Lydia 2020/05/28 媒體備註：檢查是否為英數字
     tmpBol = False
     .Text15_Validate tmpBol
     If tmpBol = True Then
        Exit Function
     End If
     'end 2020/05/28
     'Added by Lydia 2025/07/21
     tmpBol = False
     If Trim(.txtFAddr.Text) <> "" Then
        If Trim(.txtA2224) = "" Then
            MsgBox "受款人地址城市不可空白!!", vbCritical
            Exit Function
        Else
            .txtA2224_Validate tmpBol
            If tmpBol = True Then
               Exit Function
            End If
            If Len(.txtA2224.Text) > .txtA2224.MaxLength Then
               MsgBox "受款人地址城市長度不可超過" & .txtA2224.MaxLength & "字元!!", vbCritical
               Exit Function
            End If
        End If
        If Trim(.txtA2225) = "" Then
            MsgBox "受款人地址國家代號不可空白!!", vbCritical
            Exit Function
        Else
            .txtA2225_Validate tmpBol
            If tmpBol = True Then
               Exit Function
            End If
        End If
     End If
     'end 2025/07/21
     'Added by Lydia 2025/10/17
     If Trim(.txtA2226) = "" Then
         MsgBox "是否為IBAN不可定為空白!!", vbCritical
         Exit Function
     Else
         If .txtA2226 <> "Y" And .txtA2226 <> "N" Then
            MsgBox "是否為IBAN，請輸入Y或N!!", vbCritical
            Exit Function
         End If
     End If
     'end 2025/10/17
     
      If .Text1.Text = MsgText(601) Then
         MsgBox MsgText(10) & .Label9, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Function
      End If
      If .Combo2.Text = MsgText(601) Then
         MsgBox MsgText(10) & .Label7, , MsgText(5)
         strControlButton = MsgText(602)
         .Combo2.SetFocus
         Exit Function
      End If
        'Add By Cheng 2004/03/25
        '檢查受款人名稱欄位
        If .Text3.Text <> "" Then
            If CheckLengthIsOK_1(.Text3.Text, 35) = False Then
                strControlButton = MsgText(602)
                .Text3.SetFocus
                Exit Function
            End If
        End If
        If .Text9.Text <> "" Then
            If CheckLengthIsOK_1(.Text9.Text, 35) = False Then
                strControlButton = MsgText(602)
                .Text9.SetFocus
                Exit Function
            End If
        End If
        If .Text2.Text <> "" Then
            If CheckLengthIsOK_1(.Text2.Text, 35) = False Then
                strControlButton = MsgText(602)
                .Text2.SetFocus
                Exit Function
            End If
        End If
         'Modified by Lydia 2015/03/27 名稱只要3欄
'        If .Text15.Text <> "" Then
'            If CheckLengthIsOK_1(.Text15.Text, 35) = False Then
'                strControlButton = MsgText(602)
'                .Text15.SetFocus
'                Exit Sub
'            End If
'        End If

        'End
      
      If .adoacc220.State <> adStateClosed Then .adoacc220.Close
      Set .adoacc220 = Nothing
      .adoacc220.CursorLocation = adUseClient
      'Modified by Lydia 2016/12/27 統一用cnnConnection
      '.adoacc220.Open "select * from acc220 where a2201 = '" & .Text1.Text & "' And a2202='" & .Combo2.Text & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
      .adoacc220.Open "select * from acc220 where a2201 = '" & .Text1.Text & "' And a2202='" & .Combo2.Text & "' ", cnnConnection, adOpenDynamic, adLockBatchOptimistic
      If strSaveConfirm = MsgText(3) Then
         If .adoacc220.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            .adoacc220.Close
            .Acc220Refresh
            .RecordShow
            .Text1.SetFocus
            Exit Function
         End If
         '.adoacc220.AddNew 'Remove by Lydia 2016/12/27
      Else
         If strSaveConfirm = MsgText(4) Then
            If .adoacc220.RecordCount = 0 Then
               MsgBox MsgText(28), , MsgText(5)
               strControlButton = MsgText(602)
               .adoacc220.Close
               .Acc220Refresh
               .RecordShow
               .Text1.SetFocus
               Exit Function
            End If
            'Added by Lydia 2016/04/13 保留歷史記錄(依據台銀水單的需求保留)
            'Modified by Lydia 2025/07/21 +A2224受款人地址城市, A2225受款人地址國家代號
            'Mddified by Lyida 22025/10/16 +A2226 是否為IBAN
            strT = "update acc220 set a2202='" & .Combo2.Text & "',a2203=" & CNULL(.Text3.Text) & ",a2204=" & CNULL(.Text9.Text) & _
                 ",a2205=" & CNULL(.Text2.Text) & ",a2207=" & CNULL(.Text8.Text) & ",a2208=" & CNULL(.Text10.Text) & _
                 ",a2209=" & CNULL(.Text11.Text) & ",a2210=" & CNULL(.Combo1.Text) & ",a2211=" & CNULL(.Text12.Text) & _
                 ",a2212=" & CNULL(.Text14.Text) & ",a2213=" & CNULL(.Text13.Text) & ",a2214=" & CNULL(.Text18.Text) & _
                 ",a2215=" & CNULL(.Text19.Text) & ",a2216=" & CNULL(.Text20.Text) & ",a2217=" & CNULL(.Text7.Text) & _
                 ",a2218=" & CNULL(.Text21.Text) & ",a2219=" & CNULL(.Combo3.Text) & ",a2220=" & CNULL(.Text6.Text) & _
                 ",a2221=" & CNULL(ChgSQL(.Text22.Text)) & ",a2222=" & CNULL(ChgSQL(.Text15.Text)) & ",a2223=" & CNULL(ChgSQL(.Text23.Text)) & _
                 ",a2224=" & CNULL(ChgSQL(.txtA2224.Text)) & ",a2225=" & CNULL(ChgSQL(.txtA2225.Text)) & _
                 ",a2226=" & CNULL(ChgSQL(.txtA2226.Text)) & " where a2201 = '" & .Text1.Text & "' And a2202='" & .Combo2.Text & "' and updatedate='" & strSrvDate(1) & "' and updatetime='" & ServerTime & "' "
             Pub_SeekTbLog strT
             'end 2016/04/13
         End If
      End If
      
      cnnConnection.BeginTrans 'Added by Lydia 2016/12/27 包在Transaction
      'Added by Lydia 2016/12/27 從上面移下來
      If strSaveConfirm = MsgText(3) Then
        .adoacc220.AddNew
      End If
      
        .adoacc220.Fields("a2201").Value = .Text1.Text
        .adoacc220.Fields("a2202").Value = .Combo2.Text
        .adoacc220.Fields("a2203").Value = .Text3.Text
        .adoacc220.Fields("a2204").Value = .Text9.Text
        .adoacc220.Fields("a2205").Value = .Text2.Text
        'Modified by Lydia 2015/03/27 名稱只要3欄
       ' .adoacc220.Fields("a2206").Value = .Text15.Text
        .adoacc220.Fields("a2207").Value = .Text8.Text
        .adoacc220.Fields("a2208").Value = .Text10.Text
        .adoacc220.Fields("a2209").Value = .Text11.Text
        .adoacc220.Fields("a2210").Value = .Combo1.Text
        .adoacc220.Fields("a2211").Value = .Text12.Text
        .adoacc220.Fields("a2212").Value = .Text14.Text
        .adoacc220.Fields("a2213").Value = .Text13.Text
        .adoacc220.Fields("a2214").Value = .Text18.Text
        .adoacc220.Fields("a2215").Value = .Text19.Text
        .adoacc220.Fields("a2216").Value = .Text20.Text
        'Modified by Lydia 2015/3/24 針對台銀水單媒體化增加的欄位
        .adoacc220.Fields("a2217").Value = .Text7.Text
        .adoacc220.Fields("a2218").Value = .Text21.Text
        .adoacc220.Fields("a2219").Value = .Combo3.Text
        .adoacc220.Fields("a2220").Value = .Text6.Text
        .adoacc220.Fields("a2221").Value = .Text22.Text
        'end 2015/03/24
        'Modified by Lydia 2015/03/27 媒體,台一備註
        .adoacc220.Fields("a2222").Value = .Text15.Text
        .adoacc220.Fields("a2223").Value = .Text23.Text
        'Added by Lydia 2025/07/21 A2224受款人地址城市, A2225受款人地址國家代號
        .adoacc220.Fields("a2224").Value = .txtA2224.Text
        .adoacc220.Fields("a2225").Value = .txtA2225.Text
        'end 2025/07/21
        .adoacc220.Fields("a2226").Value = .txtA2226.Text 'Added by Lydia 2025/10/17 是否為IBAN
        .adoacc220.UpdateBatch
        .adoacc220.Close
        .Acc220Refresh
        .RecordShow
        .Command1.Enabled = True
        'Added by Lydia 2016/06/29 依代理人資料,變更未輸入匯票號碼(a1908)的國外付款
        strA = "select a1801, a1803,a1903,a1917 a0k11,a1811 ptype,a1718,a1702 from acc180,acc190,acc170 " & _
               "where a1801=a1901 and a1908 is null and a1801=a1709 and a1803=a1705 and a1803='" & Trim(.Text1.Text) & "' and a1903='" & Trim(.Combo2.Text) & "' " & _
               "group by a1801,a1803,a1903,a1917,a1811,a1718,a1702 "
        inX = 1
        Set Rs21q0 = ClsLawReadRstMsg(inX, strA)
        If inX = 1 Then
           Rs21q0.MoveFirst
           'cnnConnection.BeginTrans 'Remove by Lydia 2016/12/27
           Do While Not Rs21q0.EOF
              strA = GetTermOfPayment(Rs21q0.Fields("a1702"), Rs21q0.Fields("a1903"), "" & Rs21q0.Fields("a0k11"), "" & Rs21q0.Fields("a1718"))
              If strA <> "" And strA <> "" & Rs21q0.Fields("ptype") Then
                 bolUpd = True
                 strT = "update acc180 set a1811='" & strA & "',a1807=" & Val(strSrvDate(2)) & ",a1808=" & ServerTime & ",a1809='" & strUserNum & "' where a1801='" & Rs21q0.Fields("a1801") & "' and a1803='" & Rs21q0.Fields("a1803") & "' "
                 cnnConnection.Execute strT, inX
              End If
              Rs21q0.MoveNext
           Loop
           'cnnConnection.CommitTrans 'Remove by Lydia 2016/12/27
        End If
        'end 2016/06/29
        
        cnnConnection.CommitTrans 'Added by Lydia 2016/12/27
        'Added by Lydia 2015/03/30
        Frmacc21q0_Save = True
        
Checking:
   If Err.Number = 0 Then
      Exit Function
   'Added by Lydia 2016/06/29
   'Remove by Lydia 2016/12/27
   'ElseIf bolUpd Then
   '   cnnConnection.RollbackTrans
   'end 2016/06/29
   End If
   
   cnnConnection.RollbackTrans 'Added by Lydia 2016/12/27
   MsgBox Err.Description, , MsgText(5)
   
   End With
End Function

Public Sub Frmacc3130_Save()
   On Error GoTo Checking
   With Frmacc3130
      If .Text5 = MsgText(601) Then
         MsgBox MsgText(10) & .Label9, , MsgText(5)
         strControlButton = MsgText(602)
         .Text5.SetFocus
         Exit Sub
      Else
         If .Text11 = MsgText(601) Then
            MsgBox MsgText(10) & .Label5, , MsgText(5)
            strControlButton = MsgText(602)
            .Text11.SetFocus
            Exit Sub
         End If
      End If
      If .Text9 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e19").Value = .Text9
      Else
         .adoacc0e0.Fields("a0e19").Value = Null
      End If
      If .Combo1 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e20").Value = .Combo1
      Else
         .adoacc0e0.Fields("a0e20").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc0e0.Fields("a0e14").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc0e0.Fields("a0e14").Value = 0
      End If
      If strSaveConfirm = MsgText(3) Then
         .adoacc0e0.Fields("a0e26").Value = Val(strSrvDate(2))
         .adoacc0e0.Fields("a0e27").Value = ServerTime
         .adoacc0e0.Fields("a0e28").Value = strUserNum
      Else
         .adoacc0e0.Fields("a0e29").Value = Val(strSrvDate(2))
         .adoacc0e0.Fields("a0e30").Value = ServerTime
         .adoacc0e0.Fields("a0e31").Value = strUserNum
      End If
      .adoacc0e0.UpdateBatch
      .AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Public Sub Frmacc4140_Save()
   On Error GoTo Checking
   With Frmacc4140
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label3, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If .Text2 <> MsgText(602) And .Text2 <> MsgText(603) Then
            MsgBox .Label2 & MsgText(54), , MsgText(5)
            strControlButton = MsgText(602)
            .Text2.SetFocus
            Exit Sub
         End If
         If CheckLen(.Label1, .Text3, 20) = MsgText(603) Then
            strControlButton = MsgText(602)
            .Text3.SetFocus
            Exit Sub
         End If
      End If
      
      'Add by Morgan 2010/4/26
      If .TxtValidate = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If
      
      If strSaveConfirm = MsgText(3) Then
         If .Adodc1.Recordset.RecordCount <> 0 Then
            .Adodc1.Recordset.Find "a0901 = '" & .Text1 & "'", 0, adSearchForward, 1
            If .Adodc1.Recordset.EOF = False Then
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               .Text1.SetFocus
               Exit Sub
            End If
         End If
         .Adodc1.Recordset.AddNew
      End If
      .Adodc1.Recordset.Fields("a0901").Value = .Text1
      If .Text2 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0904").Value = .Text2
      Else
         .Adodc1.Recordset.Fields("a0904").Value = Null
      End If
      If .Text3 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0902").Value = .Text3
      Else
         .Adodc1.Recordset.Fields("a0902").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0903").Value = .Text4
      Else
         .Adodc1.Recordset.Fields("a0903").Value = Null
      End If
      'Add by Morgan 2010/4/26
      .Adodc1.Recordset.Fields("a0908") = .Text5
      .Adodc1.Recordset.Fields("a0909") = .Text6
      .Adodc1.Recordset.Fields("a0910") = .Text7
      
      .Adodc1.Recordset.UpdateBatch
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'Public Sub Frmacc4150_Save()
'   On Error GoTo Checking
'   With Frmacc4150
'      If .Text1 = MsgText(601) Or .Text2 = MsgText(601) Or .Text15 = MsgText(601) Or .Text17 = MsgText(601) Then
'         MsgBox MsgText(10), , MsgText(5)
'         Exit Sub
'      End If
'      If strSaveConfirm = MsgText(3) Then
'         If .adoacc0a0.RecordCount <> 0 Then
'            .adoacc0a0.Find "a0a01 = " & Val(.Text1) & "", 0, adSearchForward, 1
'            If .adoacc0a0.EOF = False Then
'               .adoacc0a0.Find "a0a02 = " & Val(.Text2) & "", 0, adSearchForward, .adoacc0a0.Bookmark
'               If .adoacc0a0.EOF = False Then
'                  .adoacc0a0.Find "a0a03 = '" & .Text15 & "'", 0, adSearchForward, .adoacc0a0.Bookmark
'                  If .adoacc0a0.EOF = False Then
'                     .adoacc0a0.Find "a0a04 = '" & .Text17 & "'", 0, adSearchForward, .adoacc0a0.Bookmark
'                     If .adoacc0a0.EOF = False Then
'                        MsgBox MsgText(9), , MsgText(5)
'                        Exit Sub
'                     End If
'                  End If
'               End If
'            End If
'         End If
'         .adoacc0a0.AddNew
'      End If
'      .adoacc0a0.Fields("a0a01").Value = Val(.Text1)
'      .adoacc0a0.Fields("a0a02").Value = Val(.Text2)
'      .adoacc0a0.Fields("a0a03").Value = .Text15
'      .adoacc0a0.Fields("a0a04").Value = .Text17
'      If .Text3 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a13").Value = Val(.Text3)
'      Else
'         .adoacc0a0.Fields("a0a13").Value = 0
'      End If
'      If .Text9 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a14").Value = Val(.Text9)
'      Else
'         .adoacc0a0.Fields("a0a14").Value = 0
'      End If
'      If .Text4 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a16").Value = Val(.Text4)
'      Else
'         .adoacc0a0.Fields("a0a16").Value = 0
'      End If
'      If .Text10 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a15").Value = Val(.Text10)
'      Else
'         .adoacc0a0.Fields("a0a15").Value = 0
'      End If
'      If .Text5 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a05").Value = Val(.Text5)
'      Else
'         .adoacc0a0.Fields("a0a05").Value = 0
'      End If
'      If .Text11 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a08").Value = Val(.Text11)
'      Else
'         .adoacc0a0.Fields("a0a08").Value = 0
'      End If
'      If .Text6 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a06").Value = Val(.Text6)
'      Else
'         .adoacc0a0.Fields("a0a06").Value = 0
'      End If
'      If .Text12 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a07").Value = Val(.Text12)
'      Else
'         .adoacc0a0.Fields("a0a07").Value = 0
'      End If
'      If .Text7 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a09").Value = Val(.Text7)
'      Else
'         .adoacc0a0.Fields("a0a09").Value = 0
'      End If
'      If .Text13 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a10").Value = Val(.Text13)
'      Else
'         .adoacc0a0.Fields("a0a10").Value = 0
'      End If
'      If .Text8 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a11").Value = Val(.Text8)
'      Else
'         .adoacc0a0.Fields("a0a11").Value = 0
'      End If
'      If .Text14 <> MsgText(601) Then
'         .adoacc0a0.Fields("a0a12").Value = Val(.Text14)
'      Else
'         .adoacc0a0.Fields("a0a12").Value = 0
'      End If
'      If strSaveConfirm = MsgText(3) Then
'         .adoacc0a0.Fields("a0a17").Value = Val(ACDate(ServerDate))
'         .adoacc0a0.Fields("a0a18").Value = ServerTime
'         .adoacc0a0.Fields("a0a19").Value = strUserNum
'      Else
'         .adoacc0a0.Fields("a0a20").Value = Val(ACDate(ServerDate))
'         .adoacc0a0.Fields("a0a21").Value = ServerTime
'         .adoacc0a0.Fields("a0a22").Value = strUserNum
'      End If
'      .adoacc0a0.UpdateBatch
'      .RecordShow
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'   End With
'End Sub

Public Sub Frmacc4180_Save()
   On Error GoTo Checking
   With Frmacc4180
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If CheckLen(.Label2, .Text5, 20) = MsgText(603) Then
            strControlButton = MsgText(602)
            .Text5.SetFocus
            Exit Sub
         End If
      End If
      If strSaveConfirm = MsgText(3) Then
         If .Adodc1.Recordset.RecordCount <> 0 Then
            .Adodc1.Recordset.Find "a0701 = '" & .Text1 & "'", 0, adSearchForward, 1
            If .Adodc1.Recordset.EOF = False Then
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               .Text1.SetFocus
               Exit Sub
            End If
         End If
         .Adodc1.Recordset.AddNew
      End If
      .Adodc1.Recordset.Fields("a0701").Value = .Text1
      If .Text5 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0702").Value = .Text5
      Else
         .Adodc1.Recordset.Fields("a0702").Value = Null
      End If
      .Adodc1.Recordset.UpdateBatch
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub
