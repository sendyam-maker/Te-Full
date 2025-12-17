Attribute VB_Name = "acc_sav"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

Public Sub Frmacc7100_Save()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   On Error GoTo Checking
   With Frmacc7100
      If .Text1.Text = MsgText(601) Then
         MsgBox MsgText(10) & .Label3, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      End If
      If .Text2.Text = MsgText(601) Then
         MsgBox MsgText(10) & .Label12, , MsgText(5)
         strControlButton = MsgText(602)
         .Text2.SetFocus
         Exit Sub
      End If
      'add by nick 2004/10/20
      If .oState = "1" Or .oState = "2" Then
            If .MaskEdBox1.Text = "" Or .MaskEdBox1.Text = "___/__/__" Then
                MsgBox "收款日期不可空白", , MsgText(5)
                strControlButton = MsgText(602)
                .MaskEdBox1.SetFocus
                Exit Sub
            End If
            'add by nick 2004/12/28
            If .Text2.Text = MsgText(4) Then
                MsgBox "人工收據不能為 E ", , MsgText(5)
                strControlButton = MsgText(602)
                .MaskEdBox1.SetFocus
                Exit Sub
            End If
      End If
      If Val(.Text4) = "0" Then
        If .MaskEdBox2.Text <> "" And .MaskEdBox2.Text <> "___/__/__" Then
            MsgBox "不可輸入到期日", , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox2.SetFocus
            Exit Sub
        End If
        If .Text5.Text <> "" Then
            MsgBox "不可輸入帳號", , MsgText(5)
            strControlButton = MsgText(602)
            .Text5.SetFocus
            Exit Sub
        End If
        If .Text6.Text <> "" Then
            MsgBox "不可輸入票號", , MsgText(5)
            strControlButton = MsgText(602)
            .Text6.SetFocus
            Exit Sub
        End If
'        If .Text7.Text <> "" Then
'            MsgBox "不可輸入付款地", , MsgText(5)
'            strControlButton = MsgText(602)
'            .Text7.SetFocus
'            Exit Sub
'        End If
      Else
'edit by nick 2004/08/20
'        If .MaskEdBox2.Text = "" Or .MaskEdBox2.Text = "___/__/__" Then
'            MsgBox "請輸入到期日", , MsgText(5)
'            strControlButton = MsgText(602)
'            .MaskEdBox2.SetFocus
'            Exit Sub
'        End If
'        If .Text5.Text = "" Then
'            MsgBox "請輸入帳號", , MsgText(5)
'            strControlButton = MsgText(602)
'            .Text5.SetFocus
'            Exit Sub
'        End If
'        If .Text6.Text = "" Then
'            MsgBox "請輸入票號", , MsgText(5)
'            strControlButton = MsgText(602)
'            .Text6.SetFocus
'            Exit Sub
'        End If
        If .Text7.Text = "" Then
            MsgBox "請輸入付款地", , MsgText(5)
            strControlButton = MsgText(602)
            .Text7.SetFocus
            Exit Sub
        End If
      End If
      If Val(.Text9.Text) > Val(.Text3.Text) + Val(.Text4.Text) Then
         MsgBox "留分所金額不可大於現金金額+支票金額", , MsgText(5)
         strControlButton = MsgText(602)
         .Text9.SetFocus
         Exit Sub
      End If
'        strSQLA = "Select A0K20||' '||ST02, A0K04, A0J20, Round(Nvl(A0J09,0)/1000,1), A0J02, Nvl(A0J09,0) + Nvl(A0J10,0) From ACC0K0, ACC0J0, Staff Where A0K01=A0J13 And A0K20=ST01 And A0K01='" & ChgSQL(.Text1.Text) & "' And ST06='" & pub_strUserOffice & "' Order By A0J03 "
'        rsA.CursorLocation = adUseClient
'        rsA.Open strSQLA, adoTaie, adOpenStatic, adLockReadOnly
'        If rsA.RecordCount <= 0 Then
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            MsgBox "查無此電腦收據資料!!!", vbExclamation + vbOKOnly
'            strControlButton = MsgText(602)
'            .Text1.SetFocus
'            Exit Sub
'        End If
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
      If .adoacc310.State <> adStateClosed Then .adoacc310.Close
      Set .adoacc310 = Nothing
      .adoacc310.CursorLocation = adUseClient
      .adoacc310.Open "Select * From ACC310 Where A3103 = '" & ChgSQL(.Text1.Text) & "' And A3104='" & ChgSQL(.Text2.Text) & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If strSaveConfirm = MsgText(3) Then
         If .adoacc310.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            .adoacc310.Close
            .Acc310Refresh
            .RecordShow
            .Text1.SetFocus
            Exit Sub
         End If
         .adoacc310.AddNew
         .adoacc310.Fields("a3114").Value = strUserNum
         .adoacc310.Fields("a3115").Value = strSrvDate(2)
         .adoacc310.Fields("a3116").Value = ServerTime
      Else
         If strSaveConfirm = MsgText(4) Then
            If .adoacc310.RecordCount = 0 Then
               MsgBox MsgText(28), , MsgText(5)
               strControlButton = MsgText(602)
               .adoacc310.Close
               .Acc310Refresh
               .RecordShow
               .Text1.SetFocus
               Exit Sub
            End If
            .adoacc310.Fields("a3117").Value = strUserNum
            .adoacc310.Fields("a3118").Value = strSrvDate(2)
            .adoacc310.Fields("a3119").Value = ServerTime
         End If
      End If
        'add by nick 2004/08/20 因為電腦中心可以用，所以要判斷若是電腦中心或不是新增，則不存
        If UCase(strUserDept) <> "M51" Or strSaveConfirm = MsgText(3) Then
            .adoacc310.Fields("a3101").Value = pub_strUserOffice
        End If
        If .MaskEdBox1.Text <> "" And .MaskEdBox1.Text <> "___/__/__" Then
            .adoacc310.Fields("a3102").Value = Val(FCDate(.MaskEdBox1.Text))
        Else
            .adoacc310.Fields("a3102").Value = Null
        End If
        .adoacc310.Fields("a3103").Value = .Text1.Text
        .adoacc310.Fields("a3104").Value = .Text2.Text
        .adoacc310.Fields("a3105").Value = Val(.Text3.Text)
        .adoacc310.Fields("a3106").Value = Val(.Text4.Text)
        If .MaskEdBox2.Text <> "" And .MaskEdBox2.Text <> "___/__/__" Then
            .adoacc310.Fields("a3107").Value = Val(FCDate(.MaskEdBox2.Text))
        Else
            .adoacc310.Fields("a3107").Value = Null
        End If
        .adoacc310.Fields("a3108").Value = .Text5.Text
        .adoacc310.Fields("a3109").Value = .Text6.Text
        .adoacc310.Fields("a3110").Value = .Text7.Text
        If .MaskEdBox3.Text <> "" And .MaskEdBox3.Text <> "___/__/__" Then
            .adoacc310.Fields("a3111").Value = Val(FCDate(.MaskEdBox3.Text))
        Else
            .adoacc310.Fields("a3111").Value = Null
        End If
        .adoacc310.Fields("a3112").Value = Val(.Text8.Text)
        .adoacc310.Fields("a3113").Value = Val(.Text9.Text)
        'add by nick 2004/08/20
        .adoacc310.Fields("a3121").Value = .Text13.Text
        .adoacc310.Fields("a3122").Value = .Text14.Text
        .adoacc310.Fields("a3123").Value = Val(.Text16.Text)
        .adoacc310.Fields("a3124").Value = .Text11.Text
        .adoacc310.UpdateBatch
        .adoacc310.Close
        .Acc310Refresh
        .RecordShow
        .Command1.Enabled = True
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'Added by Lydia 2020/03/26 從account.aacc_sav複製
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
