Attribute VB_Name = "acc_nxt"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'*************************************************
'  移動至下一筆記錄
'
'*************************************************

Public Sub Frmacc2150_Next()
   With Frmacc2150
      If .adoacc150.EOF = False Then
         .adoacc150.MoveNext
         If .adoacc150.EOF Then
            .adoacc150.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2160_Next()
   With Frmacc2160
      If .adoacc160.EOF = False Then
         .adoacc160.MoveNext
         If .adoacc160.EOF Then
            .adoacc160.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21g0_Next(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21g0
   With oForm
      If .adoacc1j0.EOF = False Then
         .adoacc1j0.MoveNext
         If .adoacc1j0.EOF Then
            .adoacc1j0.MoveLast
            MsgBox MsgText(9009), , MsgText(9001)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21h0_Next()
   With Frmacc21h0
      If .adoacc1k0.EOF = False Then
         .adoacc1k0.MoveNext
         If .adoacc1k0.EOF Then
            .adoacc1k0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21i0_Next()
   With Frmacc21i0
      If .adoacc1k0.EOF = False Then
         .adoacc1k0.MoveNext
         If .adoacc1k0.EOF Then
            .adoacc1k0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21j0_Next()
   With Frmacc21j0
      If .adoacc150.EOF = False Then
         .adoacc150.MoveNext
         If .adoacc150.EOF Then
            .adoacc150.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21k0_Next()
   With Frmacc21k0
      If .adoacc1k0.EOF = False Then
         .adoacc1k0.MoveNext
         If .adoacc1k0.EOF Then
            .adoacc1k0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21m0_Next(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21m0
   With oForm
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21o0_Next(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21o0
   With oForm
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub
'Add By Sindy 2009/06/06
Public Sub Frmacc21s0_Next(ByRef oForm As Form)
   'Modify by Morgan 2010/8/5 改用傳的專案才能不用加
   'With Frmacc21s0
   With oForm
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

