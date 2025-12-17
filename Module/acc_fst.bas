Attribute VB_Name = "acc_fst"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'*************************************************
'  移動至第一筆記錄
'
'*************************************************

Public Sub Frmacc2150_First()
   With Frmacc2150
      If .adoacc150.RecordCount <> 0 Then
         .adoacc150.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2160_First()
   With Frmacc2160
      If .adoacc160.RecordCount <> 0 Then
         .adoacc160.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21g0_First(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21g0
   With oForm
      If .adoacc1j0.RecordCount <> 0 Then
         .adoacc1j0.MoveFirst
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21h0_First()
   With Frmacc21h0
      If .adoacc1k0.RecordCount <> 0 Then
         .adoacc1k0.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21i0_First()
   With Frmacc21i0
      If .adoacc1k0.RecordCount <> 0 Then
         .adoacc1k0.MoveFirst
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21j0_First()
   With Frmacc21j0
      If .adoacc150.RecordCount <> 0 Then
         .adoacc150.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21k0_First()
   With Frmacc21k0
      If .adoacc1k0.RecordCount <> 0 Then
         .adoacc1k0.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21m0_First(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21m0
   With oForm
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21o0_First(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21o0
   With oForm
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub
'Add By Sindy 2009/06/06
Public Sub Frmacc21s0_First(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21s0
   With oForm
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

