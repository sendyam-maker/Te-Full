Attribute VB_Name = "aacc_pre"
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
'*************************************************
'  移動至上一筆記錄
'
'*************************************************
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

Public Sub Frmacc1160_Previous()
   With Frmacc1160
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc1190_Previous()
   With Frmacc1190
      If .m_IsOpen = False Then .OpenTable 'Add by Morgan 2011/10/14
      If .Option1.Value Then
         If .adoacc0s0.BOF = False Then
            .adoacc0s0.MovePrevious
            If .adoacc0s0.BOF Then
               .adoacc0s0.MoveFirst
               MsgBox MsgText(7), , MsgText(5)
            End If
            .FormShowE
            Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0s0.Bookmark & MsgText(35) & .adoacc0s0.RecordCount
         End If
      Else
         If .adoacc0t0.BOF = False Then
            .adoacc0t0.MovePrevious
            If .adoacc0t0.BOF Then
               .adoacc0t0.MoveFirst
               MsgBox MsgText(7), , MsgText(5)
            End If
            .FormShowJ
            Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0t0.Bookmark & MsgText(35) & .adoacc0t0.RecordCount
         End If
      End If
   End With
End Sub

Public Sub Frmacc11a0_Previous()
   With Frmacc11a0
      If .adoacc0t0.BOF = False Then
         .adoacc0t0.MovePrevious
         If .adoacc0t0.BOF Then
            .adoacc0t0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .AdodcClear
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11d0_Previous()
   With Frmacc11d0
      If .adocaseprogress.BOF = False Then
         .adocaseprogress.MovePrevious
         If .adocaseprogress.BOF Then
            .adocaseprogress.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11f0_Previous()
   With Frmacc11f0
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2110_Previous()
   With Frmacc2110
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc0y0.BOF = False Then
         .adoacc0y0.MovePrevious
         If .adoacc0y0.BOF Then
'            .adoacc0y0.MoveFirst
             .adoaccsum.CursorLocation = adUseClient
             .adoaccsum.Open "select max(a0y01) from acc0y0 where a0y01 <  '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
            If .adoaccsum.EOF = False Then
                If IsNull(.adoaccsum.Fields(0).Value) = False Then
                  .Text2 = .adoaccsum.Fields(0).Value
               End If
            Else
               MsgBox MsgText(7), , MsgText(5)
            End If
             .adoaccsum.Close
             .Acc0y0Refresh
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2120_Previous()
   With Frmacc2120
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2130_Previous()
   With Frmacc2130
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2140_Previous()
   With Frmacc2140
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21d0_Previous()
   With Frmacc21d0
      If .adoacc1b0.BOF = False Then
         .adoacc1b0.MovePrevious
         If .adoacc1b0.BOF Then
            .adoacc1b0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .Adodc3Clear
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21e0_Previous()
   With Frmacc21e0
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc1e0.BOF = False Then
         .adoacc1e0.MovePrevious
         If .adoacc1e0.BOF Then
            .adoacc1e0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21f0_Previous()
   With Frmacc21f0
      If .adoacc1g0.BOF = False Then
         .adoacc1g0.MovePrevious
         If .adoacc1g0.BOF Then
            .adoacc1g0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh1
         .AdodcRefresh2
         .SumShow1
         .SumShow2
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21n0_Previous()
   With Frmacc21n0
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

'Add By Cheng 2003/07/23
Public Sub Frmacc21q0_Previous()
   With Frmacc21q0
      If .adoacc220.BOF = False Then
         .adoacc220.MovePrevious
         If .adoacc220.BOF Then
            .adoacc220.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3110_Previous()
   With Frmacc3110
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .Adodc2Refresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3120_Previous()
   With Frmacc3120
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .Adodc2Refresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3140_Previous()
   With Frmacc3140
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3150_Previous()
   With Frmacc3150
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3160_Previous()
   With Frmacc3160
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3170_Previous()
   With Frmacc3170
      If .adoacc0f0.BOF = False Then
         .adoacc0f0.MovePrevious
         If .adoacc0f0.BOF Then
            .adoacc0f0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3180_Previous()
   With Frmacc3180
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3190_Previous()
   With Frmacc3190
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc31a0_Previous()
   With Frmacc31a0
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc31c0_Previous()
   With Frmacc31c0
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4130_Previous()
   With Frmacc4130
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4140_Previous()
   With Frmacc4140
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4150_Previous()
'   With Frmacc4150
'      If .adoacc0a0.BOF = False Then
'         .adoacc0a0.MovePrevious
'         If .adoacc0a0.BOF Then
'            .adoacc0a0.MoveFirst
'            MsgBox MsgText(7), , MsgText(5)
'         End If
'         .FormShow
'         .RecordShow
'      End If
'   End With
End Sub

Public Sub Frmacc4160_Previous()
   Call Frmacc4160.MoveData("Pre")  'Modify by Amy 2024/08/23 原程式搬回表單中
End Sub

Public Sub Frmacc4170_Previous()
   With Frmacc4170
      If .adoacc0d1.BOF = False Then
         .adoacc0d1.MovePrevious
         If .adoacc0d1.BOF Then
            .adoacc0d1.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4180_Previous()
   With Frmacc4180
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4190_Previous()
   With Frmacc4190
      If .adoacc061.BOF = False Then
         .adoacc061.MovePrevious
         If .adoacc061.BOF Then
            .adoacc061.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

'Add by Morgan 2004/10/27
Public Sub Frmacc41c0_Previous()
   With Frmacc41c0
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc020.BOF = False Then
         .adoacc020.MovePrevious
         If .adoacc020.BOF Then
            .adoacc020.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
      .AdodcClear
   End With
End Sub

Public Sub Frmacc5200_Previous()
   With Frmacc5200
      If .adoacc040T.BOF = False Then
         .adoacc040T.MovePrevious
         If .adoacc040T.BOF Then
            .adoacc040T.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .QueryTable
         .RecordShow
      End If
   End With
End Sub

