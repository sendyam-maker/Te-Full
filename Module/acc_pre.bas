Attribute VB_Name = "acc_pre"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'*************************************************
'  移動至上一筆記錄
'
'*************************************************

Public Sub Frmacc2150_Previous()
   With Frmacc2150
      If .adoacc150.BOF = False Then
         .adoacc150.MovePrevious
         If .adoacc150.BOF Then
            .adoacc150.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2160_Previous()
   With Frmacc2160
      If .adoacc160.BOF = False Then
         .adoacc160.MovePrevious
         If .adoacc160.BOF Then
            .adoacc160.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21g0_Previous(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21g0
   With oForm
      If .adoacc1j0.BOF = False Then
         .adoacc1j0.MovePrevious
         If .adoacc1j0.BOF Then
            .adoacc1j0.MoveFirst
            MsgBox MsgText(9008), , MsgText(9001)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21h0_Previous()
   With Frmacc21h0
      If .adoacc1k0.BOF = False Then
         .adoacc1k0.MovePrevious
         If .adoacc1k0.BOF Then
            .adoacc1k0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21i0_Previous()
   With Frmacc21i0
      If .adoacc1k0.BOF = False Then
         .adoacc1k0.MovePrevious
         If .adoacc1k0.BOF Then
            .adoacc1k0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21j0_Previous()
   With Frmacc21j0
      If .adoacc150.BOF = False Then
         .adoacc150.MovePrevious
         If .adoacc150.BOF Then
            .adoacc150.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21k0_Previous()
   With Frmacc21k0
      If .adoacc1k0.BOF = False Then
         .adoacc1k0.MovePrevious
         If .adoacc1k0.BOF Then
            .adoacc1k0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21m0_Previous(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21m0
   With oForm
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21o0_Previous(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21o0
   With oForm
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub
'Add By Sindy 2009/06/06
Public Sub Frmacc21s0_Previous(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21s0
   With oForm
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

