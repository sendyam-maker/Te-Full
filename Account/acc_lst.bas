Attribute VB_Name = "aacc_lst"
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
'*************************************************
'  移動至最後一筆記錄
'
'*************************************************
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

Public Sub Frmacc1160_Last()
   With Frmacc1160
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc1190_Last()
   With Frmacc1190
      If .m_IsOpen = False Then .OpenTable 'Add by Morgan 2011/10/14
      If .Option1.Value Then
         If .adoacc0s0.RecordCount <> 0 Then
            .adoacc0s0.MoveLast
            .FormShowE
            Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0s0.Bookmark & MsgText(35) & .adoacc0s0.RecordCount
         End If
      Else
         If .adoacc0t0.RecordCount <> 0 Then
            .adoacc0t0.MoveLast
            .FormShowJ
            Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0t0.Bookmark & MsgText(35) & .adoacc0t0.RecordCount
         End If
      End If
   End With
End Sub

Public Sub Frmacc11a0_Last()
   With Frmacc11a0
      If .adoacc0t0.RecordCount <> 0 Then
         .adoacc0t0.MoveLast
         .FormShow
         .AdodcRefresh
         .AdodcClear
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11d0_Last()
   With Frmacc11d0
      If .adocaseprogress.RecordCount <> 0 Then
         .adocaseprogress.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11f0_Last()
   With Frmacc11f0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2110_Last()
   With Frmacc2110
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc0y0.RecordCount <> 0 Then
'         .adoacc0y0.MoveLast
         .adoaccsum.CursorLocation = adUseClient
         .adoaccsum.Open "select max(a0y01) from acc0y0", adoTaie, adOpenStatic, adLockReadOnly
         If .adoaccsum.EOF = False Then
            If IsNull(.adoaccsum.Fields(0).Value) = False Then
              .Text2 = .adoaccsum.Fields(0).Value
            End If
         End If
         .adoaccsum.Close
         .Acc0y0Refresh
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2120_Last()
   With Frmacc2120
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2130_Last()
   With Frmacc2130
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2140_Last()
   With Frmacc2140
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21d0_Last()
   With Frmacc21d0
      If .adoacc1b0.RecordCount <> 0 Then
         .adoacc1b0.MoveLast
         .FormShow
         .AdodcRefresh
         .SumShow
         .Adodc3Clear
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21e0_Last()
   With Frmacc21e0
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc1e0.RecordCount <> 0 Then
         .adoacc1e0.MoveLast
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21f0_Last()
   With Frmacc21f0
      If .adoacc1g0.RecordCount <> 0 Then
         .adoacc1g0.MoveLast
         .FormShow
         .AdodcRefresh1
         .AdodcRefresh2
         .SumShow1
         .SumShow2
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21n0_Last()
   With Frmacc21n0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

'Add By Cheng 2003/07/23
Public Sub Frmacc21q0_Last()
    With Frmacc21q0
        If .adoacc220.RecordCount <> 0 Then
            .adoacc220.MoveLast
            .FormShow
            .RecordShow
        End If
    End With
End Sub

Public Sub Frmacc3110_Last()
   With Frmacc3110
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .Adodc2Refresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3120_Last()
   With Frmacc3120
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .Adodc2Refresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3140_Last()
   With Frmacc3140
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3150_Last()
   With Frmacc3150
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3160_Last()
   With Frmacc3160
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3170_Last()
   With Frmacc3170
      If .adoacc0f0.RecordCount <> 0 Then
         .adoacc0f0.MoveLast
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3180_Last()
   With Frmacc3180
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3190_Last()
   With Frmacc3190
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc31a0_Last()
   With Frmacc31a0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc31c0_Last()
   With Frmacc31c0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4130_Last()
   With Frmacc4130
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4140_Last()
   With Frmacc4140
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4150_Last()
'   With Frmacc4150
'      If .adoacc0a0.RecordCount <> 0 Then
'         .adoacc0a0.MoveLast
'         .FormShow
'         .RecordShow
'      End If
'   End With
End Sub

Public Sub Frmacc4160_Last()
   Call Frmacc4160.MoveData("Lst")  'Modify by Amy 2024/08/23 原程式搬回表單中
End Sub

Public Sub Frmacc4170_Last()
   With Frmacc4170
      If .adoacc0d1.RecordCount <> 0 Then
         .adoacc0d1.MoveLast
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4180_Last()
   With Frmacc4180
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4190_Last()
   With Frmacc4190
      If .adoacc061.RecordCount <> 0 Then
         .adoacc061.MoveLast
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

'Add by Morgan 2004/10/27
Public Sub Frmacc41c0_Last()
   With Frmacc41c0
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc020.RecordCount <> 0 Then
         .adoacc020.MoveLast
         .FormShow
         .AdodcRefresh
         .SumShow
      End If
      .AdodcClear
      .RecordShow
   End With
End Sub

Public Sub Frmacc5200_Last()
   With Frmacc5200
      If .adoacc040T.RecordCount <> 0 Then
         .adoacc040T.MoveLast
         .FormShow
         .QueryTable
         .RecordShow
      End If
   End With
End Sub

