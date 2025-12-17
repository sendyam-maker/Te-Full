Attribute VB_Name = "acc_lst"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit
'*************************************************
'  移動至最後一筆記錄
'
'*************************************************
Public Sub Frmacc7100_Last()
   With Frmacc7100
      If .adoacc310.RecordCount <> 0 Then
         .adoacc310.MoveLast
         .FormShow
         'edit by nick 2004/10/07
         '.Acc310Refresh
         .RecordShow
      End If
   End With
End Sub

'Added by Lydia 2020/03/26 從account.aacc_lst複製
Public Sub Frmacc1130_Last()
   With Frmacc1130
      If .adoacc0k0.RecordCount <> 0 Then
         .adoacc0k0.MoveLast
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc1140_Last()
   With Frmacc1140
      If .adoacc0k0.RecordCount <> 0 Then
         .adoacc0k0.MoveLast
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub
'end 2020/03/26
