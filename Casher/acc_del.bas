Attribute VB_Name = "acc_del"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

Public Sub Frmacc7100_Delete()
On Error GoTo Checking
   With Frmacc7100
      If DeleteCheck("Select A3101 From ACC310 Where A3103='" & ChgSQL(.Text1.Text) & "' And A3104='" & ChgSQL(.Text2.Text) & "' ") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "Delete From ACC310 Where A3103 = '" & ChgSQL(.Text1.Text) & "' And A3104='" & ChgSQL(.Text2.Text) & "' "
      .Acc310Refresh
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Added by Lydia 2020/03/26 從account.aacc_del複製
'沒有使用(收據抬頭修改不可刪除) Memoed by Morgan 2011/12/23
Public Sub Frmacc1140_Delete()
Dim StrSQLa As String

On Error GoTo Checking

   With Frmacc1140
      If .adoacc0k0.RecordCount = 0 Then
         Exit Sub
      End If
'      'Add By Sindy 2012/11/12 收據自動列印時間點=NULL
'      StrSQLa = "update caseprogress " & _
'                "set cp151=null " & _
'                "where cp09 in(select a0j01 from acc0j0 where a0j13='" & .Text1 & "')"
'      cnnConnection.Execute StrSQLa
'      '2012/11/12 End
      If .adoacc0j0.State = adStateOpen Then
         .adoacc0j0.Close
      End If
      .adoacc0j0.CursorLocation = adUseClient
      .adoacc0j0.Open "select * from acc0j0 where a0j13 = '" & .Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Do While .adoacc0j0.EOF = False
         .adoacc0j0.Fields("a0j06").Value = Null
         .adoacc0j0.Fields("a0j08").Value = Null
         .adoacc0j0.Fields("a0j13").Value = Null
         .adoacc0j0.UpdateBatch
         .adoacc0j0.MoveNext
      Loop
      If .adocaseprogress.State = adStateOpen Then
         .adocaseprogress.Close
      End If
      .adocaseprogress.CursorLocation = adUseClient
      .adocaseprogress.Open "select * from caseprogress where cp60 = '" & .Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Do While .adocaseprogress.EOF = False
         .adocaseprogress.Fields("cp60").Value = Null
         .adocaseprogress.UpdateBatch
         .adocaseprogress.MoveNext
      Loop
      adoTaie.Execute "delete from acc0k0 where a0k01 = '" & .Text1 & "'"
      .Acc0k0Refresh
      If .adoacc0k0.RecordCount <> 0 Then
         .adoacc0k0.MoveFirst
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
