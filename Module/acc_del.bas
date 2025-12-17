Attribute VB_Name = "acc_del"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'*************************************************
'  刪除資料表記錄
'
'*************************************************

Public Sub Frmacc2150_Delete()
On Error GoTo Checking
   With Frmacc2150
      If DeleteCheck("select a1501 from acc150 where a1501 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      'Add by Amy 2016/02/01 +Transation 因1/30婉莘刪某資料卡住後資料刪除不完整
      cnnConnection.BeginTrans
      adoTaie.Execute "delete from acc151 where axf01 = '" & .Text2 & "'"
'      adoTaie.Execute "update caseprogress set cp61 = null where cp61 = '" & .Text2 & "'"
'      adoTaie.Execute "update caseprogress set cp62 = null where cp62 = '" & .Text2 & "'"
'      adoTaie.Execute "update caseprogress set cp63 = null where cp63 = '" & .Text2 & "'"
      If .adoquery.State = adStateOpen Then
         .adoquery.Close
      End If
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select * from caseprogress where cp61 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While .adoquery.EOF = False
         adoTaie.Execute "update caseprogress set cp61 = cp62 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp62 = cp63 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         'Modify by Morgan 2007/9/27
         'adoTaie.Execute "update caseprogress set cp63 = null where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp63 = cp87 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp87 = cp88 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp88 = null where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         'end 2007/9/27
         .adoquery.MoveNext
      Loop
      .adoquery.Close
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select * from caseprogress where cp62 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While .adoquery.EOF = False
         adoTaie.Execute "update caseprogress set cp62 = cp63 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         'Modify by Morgan 2007/9/27
         'adoTaie.Execute "update caseprogress set cp63 = null where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp63 = cp87 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp87 = cp88 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp88 = null where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         'end 2007/9/27
         .adoquery.MoveNext
      Loop
      .adoquery.Close
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select * from caseprogress where cp63 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While .adoquery.EOF = False
         'Modify by Morgan 2007/9/27
         'adoTaie.Execute "update caseprogress set cp63 = null where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp63 = cp87 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp87 = cp88 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp88 = null where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         'end 2007/9/27
         .adoquery.MoveNext
      Loop
      .adoquery.Close
      'Add by Morgan 2007/9/27
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select * from caseprogress where cp87 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While .adoquery.EOF = False
         adoTaie.Execute "update caseprogress set cp87 = cp88 where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         adoTaie.Execute "update caseprogress set cp88 = null where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         .adoquery.MoveNext
      Loop
      .adoquery.Close
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select * from caseprogress where cp88 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While .adoquery.EOF = False
         adoTaie.Execute "update caseprogress set cp88 = null where cp09 = '" & .adoquery.Fields("cp09").Value & "'"
         .adoquery.MoveNext
      Loop
      .adoquery.Close
      'end 2007/9/27
      .AdodcRefresh
      .AdodcClear
      adoTaie.Execute "delete from acc150 where a1501 = '" & .Text2 & "'"
      
      'Added by Morgan 2018/5/21
      '帳單電子檔也要刪除
      strSql = "update acc152 set ayf01=ayf01 where ayf01='" & .Text2 & "'"
      cnnConnection.Execute strSql, intI
      If intI > 0 Then
         If PUB_DelFtpFile2(.Text2, , "ACC152") = False Then
            Err.Raise 999, , " 刪除Ftp檔案失敗!!"
         Else
            strSql = "delete acc152 where ayf01='" & .Text2 & "'"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2018/5/21
      
      cnnConnection.CommitTrans
      .adoacc150.Requery
      If .adoacc150.RecordCount <> 0 Then
         .adoacc150.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   cnnConnection.RollbackTrans
   'end 2016/02/01
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc2160_Delete()
On Error GoTo Checking
   With Frmacc2160
      If DeleteCheck("select a1601 from acc160 where a1601 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc161 where axg01 = '" & .Text2 & "'"
      .AdodcRefresh
      .AdodcClear
      adoTaie.Execute "delete from acc160 where a1601 = '" & .Text2 & "'"
      .adoacc160.Requery
      If .adoacc160.RecordCount <> 0 Then
         .adoacc160.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc21g0_Delete(ByRef oForm As Form)
On Error GoTo Checking
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21g0
   With oForm
      If DeleteCheck("select a1j01 from acc1j0 where a1j01 = '" & .Text1 & "' and a1j02 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc1j0 where a1j01 = '" & .Text1 & "' and a1j02 = '" & .Text2 & "'"
      .adoacc1j0.Requery
      If .adoacc1j0.RecordCount <> 0 Then
         .adoacc1j0.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc21h0_Delete()
   With Frmacc21h0
      If DeleteCheck("select a1k01 from acc1k0 where a1k01 = '" & .Text5 & "'") = MsgText(603) Then
         Exit Sub
      End If
      
'Modify by Morgan 2004/10/5 加Transaction 及錯誤控制

      adoTaie.BeginTrans '2004/10/5
            
On Error GoTo ErrTrans '2004/10/5

      adoTaie.Execute "update caseprogress set cp60 = null where cp60 = '" & .Text5 & "'"
      adoTaie.Execute "delete from acc1k0 where a1k01 = '" & .Text5 & "'"
      adoTaie.Execute "delete from acc1l0 where a1l01 = '" & .Text5 & "'"
      adoTaie.Execute "delete from acc1w0 where a1w01 = '" & .Text5 & "'" '2004/10/5
      adoTaie.Execute "delete from acc1n0 where a1n01 = '" & .Text5 & "'" 'Added by Morgan 2018/2/26
      adoTaie.CommitTrans
      
On Error GoTo ErrHnd '2004/10/5


      .adoacc1k0.Requery
      If .adoacc1k0.RecordCount <> 0 Then
         .adoacc1k0.MoveFirst
         .AdodcRefresh
         .RecordShow
      Else
         StatusClear
      End If
   
   
   End With
   
'Add by Morgan 2004/10/5

   Exit Sub
   
ErrTrans:
   adoTaie.RollbackTrans
ErrHnd:
   MsgBox Err.Description

'2004/10/5 end
End Sub

Public Sub Frmacc21j0_Delete()
On Error GoTo Checking
   With Frmacc21j0
      If DeleteCheck("select a1501 from acc150 where a1501 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "update acc150 set a1507 = '' where a1501 = '" & .Text2 & "'"
      .adoacc150.Requery
      If .adoacc150.RecordCount <> 0 Then
         .adoacc150.MoveFirst
         .AdodcRefresh
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc21k0_Delete()
On Error GoTo Checking
   With Frmacc21k0
      If DeleteCheck("select a1k01 from acc1k0 where a1k01 = '" & .Text5 & "'") = MsgText(603) Then
         Exit Sub
      End If
      MsgBox "作廢不可刪除！": Exit Sub 'Add by Morgan 2010/8/5
      adoTaie.Execute "update acc1k0 set a1k12 = null where a1k01 = '" & .Text5 & "'"
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select a1w02 from acc1w0 where a1w01 = '" & .Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While .adoquery.EOF = False
         adoTaie.Execute "update caseprogress set cp60 = '" & .Text5 & "' where cp09 = '" & .adoquery.Fields("a1w02").Value & "'"
         .adoquery.MoveNext
      Loop
      .adoquery.Close
      .adoacc1k0.Requery
      If .adoacc1k0.RecordCount <> 0 Then
         Frmacc21k0_Clear
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc21m0_Delete(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21m0
   With oForm
      If DeleteCheck("select usxr01 from usxrate where usxr01 = " & Val(FCDate(.MaskEdBox1.Text)) & "") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from usxrate where usxr01 = " & Val(FCDate(.MaskEdBox1.Text)) & ""
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
End Sub

Public Sub Frmacc21o0_Delete(ByRef oForm As Form)
On Error GoTo Checking
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21o0
   With oForm
      If DeleteCheck("select a2101 from acc210 where a2101 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a2102 = '" & .Combo1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc210 where a2101 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a2102 = '" & .Combo1 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'Add By Sindy 2009/06/06
Public Sub Frmacc21s0_Delete(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21s0
   With oForm
      If DeleteCheck("select dnr01,dnr02 from debitnoterate where dnr01 = '" & .Combo1.Text & "' and dnr02 = " & Val(FCDate(.MaskEdBox1.Text)) & "") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from debitnoterate where dnr01 = '" & .Combo1.Text & "' and dnr02 = " & Val(FCDate(.MaskEdBox1.Text)) & ""
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
End Sub


