Attribute VB_Name = "acc_nxt"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'*************************************************
'  移動至下一筆記錄
'
'*************************************************
Public Sub Frmacc7100_Next()
   With Frmacc7100
      If .adoacc310.EOF = False Then
         .adoacc310.MoveNext
         If .adoacc310.EOF Then
            .adoacc310.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         'edit by nick 2004/10/07
         '.Acc310Refresh
         .RecordShow
      End If
   End With
End Sub

'Added by Lydia 2020/03/26 從account.aacc_nxt複製
Public Sub Frmacc1130_Next()
   With Frmacc1130
      If .adoacc0k0.EOF = False Then
         .adoacc0k0.MoveNext
         If .adoacc0k0.EOF Then
            .adoacc0k0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc1140_Next()
   With Frmacc1140
      If .adoacc0k0.EOF = False Then
         .adoacc0k0.MoveNext
         If .adoacc0k0.EOF Then
            .adoacc0k0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub
'end 2020/03/26
