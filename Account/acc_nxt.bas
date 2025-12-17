Attribute VB_Name = "aacc_nxt"
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
'*************************************************
'  移動至下一筆記錄
'
'*************************************************
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

Public Sub Frmacc1160_Next()
   With Frmacc1160
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc1190_Next()
   With Frmacc1190
      If .m_IsOpen = False Then .OpenTable 'Add by Morgan 2011/10/14
      If .Option1.Value Then
         If .adoacc0s0.EOF = False Then
            .adoacc0s0.MoveNext
            If .adoacc0s0.EOF Then
               .adoacc0s0.MoveLast
               MsgBox MsgText(8), , MsgText(5)
            End If
            .FormShowE
            Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0s0.Bookmark & MsgText(35) & .adoacc0s0.RecordCount
         End If
      Else
         If .adoacc0t0.EOF = False Then
            .adoacc0t0.MoveNext
            If .adoacc0t0.EOF Then
               .adoacc0t0.MoveLast
               MsgBox MsgText(8), , MsgText(5)
            End If
            .FormShowJ
            Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0t0.Bookmark & MsgText(35) & .adoacc0t0.RecordCount
         End If
      End If
   End With
End Sub

Public Sub Frmacc11a0_Next()
   With Frmacc11a0
      If .adoacc0t0.EOF = False Then
         .adoacc0t0.MoveNext
         If .adoacc0t0.EOF Then
            .adoacc0t0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .AdodcClear
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11d0_Next()
   With Frmacc11d0
      If .adocaseprogress.EOF = False Then
         .adocaseprogress.MoveNext
         If .adocaseprogress.EOF Then
            .adocaseprogress.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11f0_Next()
   With Frmacc11f0
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2110_Next()
   With Frmacc2110
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc0y0.EOF = False Then
         .adoacc0y0.MoveNext
         If .adoacc0y0.EOF Then
'            .adoacc0y0.MoveLast
'            MsgBox MsgText(8), , MsgText(5)
            .Acc0y0Refresh
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2120_Next()
   With Frmacc2120
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2130_Next()
   With Frmacc2130
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc2140_Next()
   With Frmacc2140
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21d0_Next()
   With Frmacc21d0
      If .adoacc1b0.EOF = False Then
         .adoacc1b0.MoveNext
         If .adoacc1b0.EOF Then
            .adoacc1b0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .Adodc3Clear
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21e0_Next()
   With Frmacc21e0
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc1e0.EOF = False Then
         .adoacc1e0.MoveNext
         If .adoacc1e0.EOF Then
            .adoacc1e0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21f0_Next()
   With Frmacc21f0
      If .adoacc1g0.EOF = False Then
         .adoacc1g0.MoveNext
         If .adoacc1g0.EOF Then
            .adoacc1g0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
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

Public Sub Frmacc21n0_Next()
   With Frmacc21n0
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

'Add By Cheng 2003/07/23
Public Sub Frmacc21q0_Next()
    With Frmacc21q0
        If .adoacc220.EOF = False Then
            .adoacc220.MoveNext
            If .adoacc220.EOF Then
                .adoacc220.MoveLast
            MsgBox MsgText(8), , MsgText(5)
            End If
            .FormShow
            .RecordShow
        End If
    End With
End Sub

Public Sub Frmacc3110_Next()
   With Frmacc3110
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .Adodc2Refresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3120_Next()
   With Frmacc3120
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .Adodc2Refresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3140_Next()
   With Frmacc3140
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3150_Next()
   With Frmacc3150
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3160_Next()
   With Frmacc3160
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3170_Next()
   With Frmacc3170
      If .adoacc0f0.EOF = False Then
         .adoacc0f0.MoveNext
         If .adoacc0f0.EOF Then
            .adoacc0f0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3180_Next()
   With Frmacc3180
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc3190_Next()
   With Frmacc3190
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc31a0_Next()
   With Frmacc31a0
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc31c0_Next()
   With Frmacc31c0
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4130_Next()
   With Frmacc4130
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4140_Next()
   With Frmacc4140
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4150_Next()
'   With Frmacc4150
'      If .adoacc0a0.EOF = False Then
'         .adoacc0a0.MoveNext
'         If .adoacc0a0.EOF Then
'            .adoacc0a0.MoveLast
'            MsgBox MsgText(8), , MsgText(5)
'         End If
'         .FormShow
'         .RecordShow
'      End If
'   End With
End Sub

Public Sub Frmacc4160_Next()
   Call Frmacc4160.MoveData("Nxt")  'Modify by Amy 2024/08/23 原程式搬回表單中
End Sub

Public Sub Frmacc4170_Next()
   With Frmacc4170
      If .adoacc0d1.EOF = False Then
         .adoacc0d1.MoveNext
         If .adoacc0d1.EOF Then
            .adoacc0d1.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4180_Next()
   With Frmacc4180
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc4190_Next()
   With Frmacc4190
      If .adoacc061.EOF = False Then
         .adoacc061.MoveNext
         If .adoacc061.EOF Then
            .adoacc061.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
         .RecordShow
      End If
   End With
End Sub

'Add by Morgan 2004/10/27
Public Sub Frmacc41c0_Next()
   With Frmacc41c0
      .CreDebCheck
      If .CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If .adoacc020.EOF = False Then
         .adoacc020.MoveNext
         'Add by Morgan 2004/10/27
         If .adoacc020.EOF Then
            .Acc020Refresh 1
            If .adoacc020.RecordCount > 0 Then
               .adoacc020.MoveNext
            Else
               Exit Sub
            End If
         End If
         '2004/10/27 end
         If .adoacc020.EOF Then
            .adoacc020.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .SumShow
      End If
      .AdodcClear
      .RecordShow
   End With
End Sub

Public Sub Frmacc5200_Next()
   With Frmacc5200
      If .adoacc040T.EOF = False Then
         .adoacc040T.MoveNext
         If .adoacc040T.EOF Then
            .adoacc040T.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .QueryTable
         .RecordShow
      End If
   End With
End Sub

