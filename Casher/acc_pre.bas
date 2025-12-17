Attribute VB_Name = "acc_pre"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'*************************************************
'  移動至上一筆記錄
'
'*************************************************
Public Sub Frmacc7100_Previous()
   With Frmacc7100
      If .adoacc310.BOF = False Then
         .adoacc310.MovePrevious
         If .adoacc310.BOF Then
            .adoacc310.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         'edit by nick 2004/10/07
         '.Acc310Refresh
         .RecordShow
      End If
   End With
End Sub

'Added by Lydia 2020/03/26 從account.aacc_pre複製
Public Sub Frmacc1130_Previous()
   With Frmacc1130
      If .adoacc0k0.BOF = False Then
         .adoacc0k0.MovePrevious
         If .adoacc0k0.BOF Then
            .adoacc0k0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc1140_Previous()
   With Frmacc1140
      If .adoacc0k0.BOF = False Then
         .adoacc0k0.MovePrevious
         If .adoacc0k0.BOF Then
            .adoacc0k0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub
'end 2020/03/26
