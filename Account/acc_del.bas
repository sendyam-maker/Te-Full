Attribute VB_Name = "aacc_del"
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
'*************************************************
'  刪除資料表記錄
'
'*************************************************
Public Sub Frmacc1110_Delete()
On Error GoTo Checking
   With Frmacc1110
      If .Text4 <> MsgText(601) And .Text5 <> MsgText(601) Then
         adoTaie.Execute "delete from acc0k0 where a0k01 >= '" & .Text4 & "' and a0k01 <= '" & .Text5 & "'"
      End If
      If .Text6 <> MsgText(601) And .Text7 <> MsgText(601) Then
         adoTaie.Execute "delete from acc0k0 where a0k01 >= '" & .Text6 & "' and a0k01 <= '" & .Text7 & "'"
      End If
      If .Text10 <> MsgText(601) And .Text11 <> MsgText(601) Then
         adoTaie.Execute "delete from acc0k0 where a0k01 >= '" & .Text10 & "' and a0k01 <= '" & .Text11 & "'"
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Public Sub Frmacc1130_Delete()
'On Error GoTo Checking
'   With Frmacc1130
'      If DeleteCheck("select a0k01 from acc0k0 where a0k01 = '" & .Text1 & "'") = MsgText(603) Then
'         Exit Sub
'      End If
''      adoTaie.Execute "update acc0k0 set a0k09 = 0 where a0k01 = '" & .Text1 & "'"
'      adoTaie.Execute "delete from acc0k0 where a0k01 = '" & .Text1 & "'"
'      .Acc0k0Refresh
'      If .adoacc0k0.RecordCount <> 0 Then
'         .adoacc0k0.MoveFirst
'         .RecordShow
'      Else
'         StatusClear
'      End If
'   End With
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

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

Public Sub Frmacc1150_Delete()
   Dim bInTrans As Boolean
   
On Error GoTo Checking
   With Frmacc1150
      If DeleteCheck("select a0l01 from acc0l0 where a0l01 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      'Added by Morgan 2014/1/17
      adoTaie.BeginTrans
      bInTrans = True
      adoTaie.Execute "Update acc440 set a4416=null where a4416='" & .Text2 & "'", intI
      'end 2014/1/17
      
      adoTaie.Execute "delete from acc0e0 where a0e02 in (select a1p09 from acc1p0 where a1p04 = '" & .Text2 & "' and a1p05 = '113001' and a1p09 is not null)"
      'Modified by Morgan 2014/1/20 收款會有J公司,取消 a1p01='1' 條件
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & .Text2 & "'"
      .adoacc1p0.Requery
      'Modified by Morgan 2018/6/20 +a0k37=null
      adoTaie.Execute "update acc0k0 set a0k17 = (select sum(nvl(a1u04, 0)) from acc1u0 where a1u02 = a0k01 and a1u01 <> '" & .Text2 & "'), a0k18 = (select sum(nvl(a1u05, 0)) from acc1u0 where a1u02 = a0k01 and a1u01 <> '" & .Text2 & "'),a0k37=null where a0k01 in (select a0m02 from acc0m0 where a0m01 = '" & .Text2 & "')"
      adoTaie.Execute "delete from acc1u0 where a1u01 = '" & .Text2 & "'"
      '更新caseprogress
      'Modify by Morgan 2006/3/17
      'adoTaie.Execute "update caseprogress set cp73 = (select sum(nvl(a1u04, 0)) from acc1u0 where a1u03 = cp09 and a1u01 <> '" & .Text2 & "'), cp74 = (select sum(nvl(a1u05, 0)) from acc1u0 where a1u03 = cp09 and a1u01 <> '" & .Text2 & "') where cp60 in (select a0m02 from acc0m0 where a0m01 = '" & .Text2 & "')"
      'adoTaie.Execute "update caseprogress set cp75 = nvl(cp73, 0) - nvl(cp74, 0), cp79 = nvl(cp16, 0) - cp75 where cp60 in (select a0m02 from acc0m0 where a0m01 = '" & .Text2 & "')"
      'Modify by Morgan 2011/10/13 考慮拆收據情形
      'strSql = "update caseprogress set (cp73,cp74,cp75,cp76,cp77,cp78,cp79 )" & _
         " = (select nvl(sum(a1u04),0) cp73, nvl(sum(a1u05),0) cp74" & _
         ", nvl(sum(a1u04),0)+nvl(sum(a1u05),0) cp75, nvl(sum(a1u06),0) cp76" & _
         ", nvl(sum(a1u07),0)+nvl(sum(a1u09),0) cp77, nvl(sum(a1u08),0)+nvl(sum(a1u10),0) cp78" & _
         ", cp16-nvl(sum(a1u04),0)-nvl(sum(a1u05),0)-nvl(sum(a1u07),0)-nvl(sum(a1u09),0) cp79" & _
         " from acc1u0 where a1u03=cp09)" & _
         " where cp60 in (select a0m02 from acc0m0 where a0m01 = '" & .Text2 & "')"
      strSql = "update caseprogress set (cp73,cp74,cp75,cp76,cp77,cp78,cp79 )" & _
         " = (select nvl(sum(a1u04),0) cp73, nvl(sum(a1u05),0) cp74" & _
         ", nvl(sum(a1u04),0)+nvl(sum(a1u05),0) cp75, nvl(sum(a1u06),0) cp76" & _
         ", nvl(sum(a1u07),0)+nvl(sum(a1u09),0) cp77, nvl(sum(a1u08),0)+nvl(sum(a1u10),0) cp78" & _
         ", cp16-nvl(sum(a1u04),0)-nvl(sum(a1u05),0)-nvl(sum(a1u07),0)-nvl(sum(a1u09),0) cp79" & _
         " from acc1u0 where a1u03=cp09)" & _
         " where cp09 in (select a0j01 from acc0m0,acc0j0 where a0m01='" & .Text2 & "' and a0j13(+)=a0m02)"
      '2011/10/13
      '2006/3/17 end
      adoTaie.Execute strSql
      
      'Added by Morgan 2025/3/4
      '刪除扣繳資料(該次收款的收據且沒有其他收款記錄者)
      strSql = "delete from acc1v0 where a1v02 in (select a0m02 from acc0m0 where a0m01= '" & .Text2 & "') and not exists(select * from acc1u0 where a1u02=a1v02 and a1u03=a1v01 and substr(a1u01,1,1)='F')"
      adoTaie.Execute strSql, intI
      '更新部份收款資料的扣繳資料
      strSql = "update acc1v0 set (a1v05,a1v06,a1v07)=(select decode(max(nvl(a0j09,0)+nvl(a0j10,0))-sum(nvl(a1u07,0)+nvl(a1u09,0))-sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)),0,'N','Y')" & _
         ",sum(a1u06),a1v04-sum(a1u06) from acc0j0,acc1u0 where a0j01=a1v01 and a0j13=a1v02 and a1u02(+)=a0j13 and a1u03(+)=a0j01)" & _
         " where a1v02 in (select a0m02 from acc0m0 where a0m01= '" & .Text2 & "')"
      adoTaie.Execute strSql, intI
      'end 2025/3/4
      
      adoTaie.Execute "delete from acc0m0 where a0m01 = '" & .Text2 & "'"
      adoTaie.Execute "delete from acc0l0 where a0l01 = '" & .Text2 & "'"
      .adoacc0l0.Requery
      '93.11.9 ADD BY SONIA
      adoTaie.Execute "delete from acc0T0 where a0T07 = '" & .Text2 & "'"
      '93.11.9 END
      
      adoTaie.CommitTrans 'Added by Morgan 2014/1/17
      
      '.adoquery.Close  '93.11.9 CANCEL BY SONIA
      .AdodcRefresh
      .SumShow
      If .adoacc0l0.RecordCount <> 0 Then
         .adoacc0l0.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
   Exit Sub
   
Checking:
   If bInTrans = True Then adoTaie.RollbackTrans 'Added by Morgan 2014/1/17
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc1160_Delete()
On Error GoTo Checking
   With Frmacc1160
      If DeleteCheck("select a0i01 from acc0i0 where a0i01 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc0i0 where a0i01 = '" & .Text1 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc11f0_Delete()
On Error GoTo Checking
   With Frmacc11f0
      If DeleteCheck("select a0w01 from acc0w0 where a0w02 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "update acc1v0 set a1v14 = '" & MsgText(602) & "' where a1v15 = '" & .Text1 & "'"
      adoTaie.Execute "update acc0w0 set a0w15 = null where a0w02 = '" & .Text1 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc2110_Delete()
On Error GoTo Checking
   With Frmacc2110
      If DeleteCheck("select a0y01 from acc0y0 where a0y01 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      .adoaccsum.CursorLocation = adUseClient
      .adoaccsum.Open "select a0z02 from acc0z0 where a0z01 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While .adoaccsum.EOF = False
         adoTaie.Execute "update acc1k0 set a1k29 = null, a1k30 = null where a1k01 = '" & .adoaccsum.Fields("a0z02").Value & "'"
         .adoaccsum.MoveNext
      Loop
      .adoaccsum.Close
      'Add By Sindy 2015/12/31
      adoTaie.Execute "delete from acc1v0 where a1v02 in(select a0z02 from acc0z0 where a0z01 = '" & .Text2 & "')"
      '2015/12/31 END
      adoTaie.Execute "delete from acc0z0 where a0z01 = '" & .Text2 & "'"
      adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & .Text2 & "'"
      adoTaie.Execute "delete from acc0y0 where a0y01 = '" & .Text2 & "'"
      adoTaie.Execute "delete from acc120 where a1210 = '" & .Text2 & "'"
      .adoacc0y0.Requery
      .AdodcRefresh
      .SumShow
      If .adoacc0y0.RecordCount <> 0 Then
         .adoacc0y0.MoveFirst
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

Public Sub Frmacc2120_Delete()
On Error GoTo Checking
   With Frmacc2120
      If DeleteCheck("select a1201 from acc120 where a1201 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p04 = '" & .Adodc1.Recordset.Fields("a1201").Value & "'"
      adoTaie.Execute "delete from acc120 where a1201 = '" & .Text2 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc2130_Delete()
On Error GoTo Checking
   With Frmacc2130
      If DeleteCheck("select a1301 from acc130 where a1301 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'H' and a1p04 = '" & .Adodc1.Recordset.Fields("a1301").Value & "'"
      adoTaie.Execute "delete from acc130 where a1301 = '" & .Text2 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc2140_Delete()
On Error GoTo Checking
   With Frmacc2140
      If DeleteCheck("select a1401 from acc140 where a1401 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "update acc1k0 set a1k25 = null where a1k01 = '" & .Text1 & "'"
'      .adoacc1k0.Requery
      adoTaie.Execute "delete from acc140 where a1401 = '" & .Text2 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc21d0_Delete()
On Error GoTo Checking
   With Frmacc21d0
      If DeleteCheck("select a1b01 from acc1b0 where a1b01 = '" & .Text3 & "' and a1b02 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      '2012/10/17 add by sonia
      adoTaie.Execute "update acc150 set a1520 = 0 where a1501 in (select a1902 from acc190 where a1908 = '" & .Text3 & "')"
      adoTaie.Execute "update acc160 set a1607 = null where a1601 in (select a1902 from acc190 where a1908 = '" & .Text3 & "')"
      '2012/10/17 end
      adoTaie.Execute "delete from acc1c0 where a1c01 = '" & .Text3 & "' and a1c02 = '" & .Text1 & "'"
      'Modified by Morgan 2015/9/11
      'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & .Text3 & .Text1 & "'"
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'I' and a1p04 = '" & .Text3 & .Text1 & "'"
      'end 2015/9/11
      .AdodcRefresh
      adoTaie.Execute "delete from acc1b0 where a1b01 = '" & .Text3 & "' and a1b02 = '" & .Text1 & "'"
      adoTaie.Execute "update acc190 set a1908 = null where a1908 = '" & .Text3 & "'"
      .adoacc1b0.Requery
      .SumShow
      .Adodc3Clear
      If .adoacc1b0.RecordCount <> 0 Then
         .adoacc1b0.MoveFirst
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

Public Sub Frmacc21e0_Delete()
On Error GoTo Checking
   With Frmacc21e0
      If DeleteCheck("select a1e01 from acc1e0 where a1e01 = '" & .Text1 & "' and a1e02 = " & Val(FCDate(.MaskEdBox1.Text)) & "") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'J' and a1p04 = '" & .Text1 & Val(FCDate(.MaskEdBox1.Text)) & "'"
      .AdodcRefresh
      .SumShow
      adoTaie.Execute "delete from acc1e0 where a1e01 = '" & .Text1 & "' and a1e02 = " & Val(FCDate(.MaskEdBox1.Text)) & ""
      .adoacc1e0.Requery
      .AdodcClear
      If .adoacc1e0.RecordCount <> 0 Then
         .adoacc1e0.MoveFirst
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

Public Sub Frmacc21f0_Delete()
On Error GoTo Checking
   With Frmacc21f0
      If DeleteCheck("select a1g01 from acc1g0 where a1g01 = '" & .Text9 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "update acc1k0 set a1k17 = null, a1k29 = null, a1k30 = 0 where a1k17 = '" & .Text9 & "'"
      adoTaie.Execute "update acc150 set a1512 = null, a1520 = 0 where a1512 = '" & .Text9 & "'"
      .AdodcRefresh1
      .AdodcRefresh2
      adoTaie.Execute "delete from acc1i0 where a1i01 = '" & .Text9 & "'"
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'K' and a1p04 = '" & .Text9 & "'"
      adoTaie.Execute "delete from acc1g0 where a1g01 = '" & .Text9 & "'"
      .adoacc1g0.Requery
      If .adoacc1g0.RecordCount <> 0 Then
         .adoacc1g0.MoveFirst
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

Public Sub Frmacc21n0_Delete()
On Error GoTo Checking
   With Frmacc21n0
      If DeleteCheck("select a1x01 from acc1x0 where a1x01 = '" & .Combo1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc1x0 where a1x01 = '" & .Combo1 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

'Add By Cheng 2003/07/23
Public Sub Frmacc21q0_Delete()
On Error GoTo Checking
   With Frmacc21q0
      If DeleteCheck("select a2201 from acc220 where a2201 = '" & .Text1.Text & "' And a2202='" & .Combo2.Text & "' ") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc220 where a2201 = '" & .Text1.Text & "' And a2202='" & .Combo2.Text & "' "
      .Acc220Refresh
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc3180_Delete()
On Error GoTo Checking
   With Frmacc3180
      If DeleteCheck("select a0g01 from acc0g0 where a0g01 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc0g0 where a0g01 = '" & .Text1 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc3190_Delete()
On Error GoTo Checking
   With Frmacc3190
      If DeleteCheck("select a0h01 from acc0h0 where a0h01 = '" & .Text5 & "' and a0h02 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc0h0 where a0h01 = '" & .Text5 & "' and a0h02 = '" & .Text1 & "'"
      .adoacc0h0.Requery
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc4130_Delete()
On Error GoTo Checking
   With Frmacc4130
      If DeleteCheck("select a0801 from acc080 where a0801 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc080 where a0801 = '" & .Text1 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc4140_Delete()
On Error GoTo Checking
   With Frmacc4140
      If DeleteCheck("select a0901 from acc090 where a0901 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc090 where a0901 = '" & .Text1 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc4170_Delete()
On Error GoTo Checking
   With Frmacc4170
      If DeleteCheck("select axd01 from acc0d1 where axd01 = '" & .Text1 & "' and axd02 = '" & .Text3 & "'") = MsgText(603) Then
         Exit Sub
      End If
      '2008/3/27 modify by sonia 不應有a0d03的條件
      'adoTaie.Execute "delete from acc0d0 where a0d01 = '" & .Text1 & "' and a0d02 = " & Val(.Text3) & " and a0d03 = '" & .Text6 & "'"
      adoTaie.Execute "delete from acc0d0 where a0d01 = '" & .Text1 & "' and a0d02 = " & Val(.Text3) & ""
      .AdodcRefresh
      .AdodcClear
      adoTaie.Execute "delete from acc0d1 where axd01 = '" & .Text1 & "' and axd02 = '" & .Text3 & "'"
      .adoacc0d1.Requery
      If .adoacc0d1.RecordCount <> 0 Then
         .adoacc0d1.MoveFirst
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

Public Sub Frmacc4180_Delete()
On Error GoTo Checking
   With Frmacc4180
      If DeleteCheck("select a0701 from acc070 where a0701 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc070 where a0701 = '" & .Text1 & "'"
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

Public Sub Frmacc4190_Delete()
On Error GoTo Checking
   With Frmacc4190
      If DeleteCheck("select ax601 from acc061 where ax601 = " & Val(.Text6) & " and ax602 = '" & .Text5 & "' and ax603 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc060 where a0603 = '" & .Text1 & "' and a0601 = " & Val(.Text6) & " and a0602 = '" & .Text5 & "'"
      adoTaie.Execute "delete from acc061 where ax601 = " & Val(.Text6) & " and ax602 = '" & .Text5 & "' and ax603 = '" & .Text1 & "'"
      .adoacc061.Requery
      .AdodcRefresh
      If .adoacc061.RecordCount <> 0 Then
         .adoacc061.MoveFirst
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

Public Sub Frmacc41a0_Delete()
   Dim intR As Integer
On Error GoTo Checking
   With Frmacc41a0
   'Added by Morgan 2022/9/30
   adoTaie.BeginTrans
   adoTaie.Execute "update ACC1P0 set a1p22=a1p22 WHERE A1P04='" & .Text10 & "' and a1p22 is not null", intR
   If intR > 0 Then
      Err.Raise 999, , "已有傳票號，不可刪除！"
   Else
      adoTaie.Execute "DELETE ACC1P0 WHERE A1P04='" & .Text10 & "' and a1p22 is null", intR
      adoTaie.Execute "UPDATE ACC240 SET A240015=NULL,A240016=NULL WHERE A240002='" & .Text10 & "'", intR
      adoTaie.CommitTrans
   End If
   'end 2022/9/30
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   'Modified by Morgan 2022/9/30
   'MsgBox Err.Description, , MsgText(5)
   adoTaie.RollbackTrans
   Err.Raise Err.Number, , Err.Description
   'end 2022/9/30
End Sub

Public Sub Frmacc41b0_Delete()
On Error GoTo Checking
   With Frmacc41b0
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc41d0_Delete()
On Error GoTo Checking
   With Frmacc41d0
      If DeleteCheck("select a1p01 from acc1p0 where a1p01 = '" & .Text1 & "' and a1p04 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc1p0 where a1p01 = '" & .Text1 & "' and a1p04 = '" & .Text2 & "'"
      .AdodcRefresh
      .AdodcClear
      .adoacc020.Requery
      If .adoacc020.RecordCount <> 0 Then
         .adoacc020.MoveFirst
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
'Add by Morgan 2006/10/17
Public Sub Frmacc41e0_Delete()
On Error GoTo Checking
   With Frmacc41e0
      If DeleteCheck("select a2301 from acc230 where a2301='" & .txtA2301 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc230 where a2301 = '" & .txtA2301 & "'"
      FormMoveLast
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
Public Sub Frmacc5200_Delete()
Dim strDepartment As String

On Error GoTo Checking
   With Frmacc5200
      If .Text1 = MsgText(601) Then
         strDepartment = MsgText(55)
      Else
         strDepartment = .Text1
      End If
      If DeleteCheck("select a0401 from acc040 where a0401 = " & Val(.Text6) & " and a0403 = '" & .Text4 & "' and a0404 = '" & strDepartment & "' and a0405 = '" & .Text3 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc040 where a0401 = " & Val(.Text6) & " and a0403 = '" & .Text4 & "' and a0404 = '" & strDepartment & "' and a0405 = '" & .Text3 & "'"
      .adoacc040T.Requery
      Frmacc5200_Clear
      .QueryTable
      If .adoacc040T.RecordCount <> 0 Then
         .adoacc040T.MoveFirst
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

