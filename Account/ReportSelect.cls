VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Acc11e0(strTitle As String, strMem As String, strPrintDate As String)
Dim arrPrintDate1
    arrPrintDate1 = Split(strPrintDate & ",2", ",")
   If AccEnvironment.rsRpt11e0.State = adStateOpen Then
      AccEnvironment.rsRpt11e0.Close
   End If
   AccEnvironment.rsRpt11e0.Open "select * from accrpt111 order by r11101 asc, r11104 asc"
   With Rptacc11e0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = arrPrintDate1(0)
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc1411(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt1411.State = adStateOpen Then
      AccEnvironment.rsRpt1411.Requery
   End If
   With Rptacc1411
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .Sections(2).Controls("Label2").Caption = strCon1
      .Sections(2).Controls("Label14").Caption = strCon2
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc1430(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
Dim arrcon1
Dim arrCon2
Dim arrCon3 'Add by Amy 2014/01/16
   
    arrcon1 = Split(strCon1, ",")
    arrCon2 = Split(strCon2, ",")
    arrCon3 = Split(strTitle, ",") 'Add 2014/01/16 +公司名稱
   If AccEnvironment.rsRpt1430_Grouping.State = adStateOpen Then
      AccEnvironment.rsRpt1430_Grouping.Requery
   End If
   With Rptacc1430
      'Modify by Amy 2014/01/16
      .Sections(2).Controls("Label14").Caption = arrCon3(0)
      .Sections(2).Controls("Label3").Caption = arrCon3(1) '原:strTitle
      'end 2014/01/16
      .Sections(2).Controls("Label2").Caption = arrcon1(0)
      .Sections(2).Controls("Label8").Caption = arrCon2(0)
      .Sections(2).Controls("Label23").Caption = arrcon1(1)
      .Sections(2).Controls("Label24").Caption = arrCon2(1)
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc1470(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt1470.State = adStateOpen Then
      AccEnvironment.rsRpt1470.Requery
   End If
   With Rptacc1470
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label2").Caption = strCon1
      .Sections(2).Controls("Label8").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc1480(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt1480.State = adStateOpen Then
      AccEnvironment.rsRpt1480.Close
   End If
   AccEnvironment.rsRpt1480.CursorLocation = adUseClient
   If strCon4 = "" Then
      AccEnvironment.rsRpt1480.Open "SELECT * FROM ACCRPT108 WHERE R10801 = '" & strCon6 & "' AND R10809 is null ORDER BY R10802 ASC", AccEnvironment.Connections.Item(1), adOpenStatic, adLockReadOnly
   Else
      AccEnvironment.rsRpt1480.Open "SELECT * FROM ACCRPT108 WHERE R10801 = '" & strCon6 & "' AND R10809 = " & strCon4 & " ORDER BY R10802 ASC", AccEnvironment.Connections.Item(1), adOpenStatic, adLockReadOnly
   End If
'   If AccEnvironment.rsRpt1480.State = adStateOpen Then
'      AccEnvironment.rsRpt1480.Requery
'   End If
   With Rptacc1480
      If strTitle = "***  智權人員帳齡分析表  ***" Then
         .Sections(2).Controls("Label23").Caption = ""
         .Sections(2).Controls("Label6").Caption = "智權人員"
      Else
         .Sections(2).Controls("Label6").Caption = "客戶編號"
      End If
      If strTitle = "***  客戶帳齡分析表  ***" Then
         .Sections(2).Controls("Label23").Caption = ""
      End If
      If strTitle = "***  智權人員客戶帳齡分析表  ***" Then
         .Sections(2).Controls("Label23").Caption = "智權人員:"
      End If
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label2").Caption = strCon1
      .Sections(2).Controls("Label8").Caption = strCon2
      .Sections(2).Controls("Label5").Caption = strCon3
      .Sections(2).Controls("Label24").Caption = strCon4
      .Sections(2).Controls("Label25").Caption = strCon5
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc1490(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   Dim arrCon
   arrCon = Split(strCon2, ",")
   If AccEnvironment.rsRpt1490.State = adStateOpen Then
      AccEnvironment.rsRpt1490.Requery
   End If
   With Rptacc1490
      .Sections(2).Controls("Label3").Caption = strTitle
      If strCon1 <> "" Then
         .Sections(2).Controls("Label2").Caption = "銷退日期：" & strCon1
      Else
         .Sections(2).Controls("Label2").Caption = ""
      End If
      If arrCon(0) <> "" Then
         .Sections(2).Controls("Label24").Caption = "類　　別：" & arrCon(0)
      Else
         .Sections(2).Controls("Label24").Caption = ""
      End If
      If arrCon(1) <> "" Then
         .Sections(2).Controls("Label25").Caption = "退費方式：" & arrCon(1)
      Else
         .Sections(2).Controls("Label25").Caption = ""
      End If
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc14a0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt14a0.State = adStateOpen Then
      AccEnvironment.rsRpt14a0.Requery
   End If
   With Rptacc14a0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label2").Caption = strCon1
      .Sections(2).Controls("Label8").Caption = strCon2
      .Sections(2).Controls("Label15").Caption = strCon3
      .Sections(2).Controls("Label16").Caption = strCon4
      .Sections(2).Controls("Label23").Caption = strCon5
      .Sections(2).Controls("Label24").Caption = strCon6
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc14b0(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
'edit by nick 2004/07/07 因為沒切好，所以修正
'Dim arrPrintDate
Dim arrPrintDate1
'Add by Amy 2014/01/28
Dim arrTitle
    arrTitle = Split(strTitle, ",")
'end 2014/01/28
    'edit by nick 2004/07/07
    'arrPrintDate = Split(strPrintDate, ",")
    arrPrintDate1 = Split(strPrintDate & ",2", ",")
    If AccEnvironment.rsRpt14b0.State = adStateOpen Then
        AccEnvironment.rsRpt14b0.Close
    End If
    '不含寄分所
    'edit by nick 2004/07/07
    'If arrPrintDate(1) = "1" Then
    If arrPrintDate1(1) = "1" Then
        AccEnvironment.rsRpt14b0.Open "select * from accrpt111 Where R11108 Is Null Or substr(R11108, 1, 1) <> '3' " ', AccEnvironment.Commands.Item(1), adOpenStatic, adLockReadOnly
    '寄分所
    Else
        'add by nick 2004/07/12 加入修正，因為 繳款書印出明細，也有用到，所以修正
        If arrPrintDate1(1) = "2" Then
            AccEnvironment.rsRpt14b0.Open "select * from accrpt111 order by r11101 asc, r11110 asc "
        Else
            AccEnvironment.rsRpt14b0.Open "select * from accrpt111 Where substr(R11108, 1, 1) = '3' " ', AccEnvironment.Commands.Item(1), adOpenStatic, adLockReadOnly
        End If
    End If
    'edit by nick 2004/07/07
'   If AccEnvironment.rsRpt14b0.State = adStateOpen Then
'      AccEnvironment.rsRpt14b0.Requery
'   End If
   With Rptacc14b0
      'Modify by Amy 2014/01/28 +公司別
      .Sections(2).Controls("Label1").Caption = arrTitle(0)
      .Sections(2).Controls("Label3").Caption = arrTitle(1) '原:strTitle
      'end 2014/01/28
      .Sections(2).Controls("Label2").Caption = strCon1
      .Sections(2).Controls("Label8").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strMem
      'edit by nick 2004/07/07
      '.Sections(2).Controls("Label19").Caption = strPrintDate
      .Sections(2).Controls("Label19").Caption = arrPrintDate1(0)
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc14c0(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   'Add by Amy 2014/02/06
   Dim arrTitle
   arrTitle = Split(strTitle, ",")
   'end 2014/02/06
   If AccEnvironment.rsRpt14c0.State = adStateOpen Then
      AccEnvironment.rsRpt14c0.Requery
   End If
   With Rptacc14c0
      'Modify by Amy 2014/02/06
      .Sections(2).Controls("Label13").Caption = arrTitle(0)
      .Sections(2).Controls("Label3").Caption = arrTitle(1) '原:strTitle
      'end 2014/02/06
      .Sections(2).Controls("Label2").Caption = strCon1
      .Sections(2).Controls("Label8").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc14e0(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt14e0.State = adStateOpen Then
      AccEnvironment.rsRpt14e0.Requery
   End If
   With Rptacc14e0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label2").Caption = strCon1
      .Sections(2).Controls("Label8").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc2160(strTitle As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt2160.State = adStateOpen Then
      AccEnvironment.rsRpt2160.Requery
   End If
   With Rptacc2160
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc2170(strTitle As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt2170.State = adStateOpen Then
      AccEnvironment.rsRpt2170.Requery
   End If
   With Rptacc2170
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc2180(strTitle As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt2180_Grouping.State = adStateOpen Then
      AccEnvironment.rsRpt2180_Grouping.Requery
   End If
   With Rptacc2180
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label12").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc24a0(strTitle As String, strCon1 As String, intSelect As Integer, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt24a0_Grouping.State = adStateOpen Then
      AccEnvironment.rsRpt24a0_Grouping.Requery
   End If
   With Rptacc24a0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label12").Caption = strPrintDate
      .Sections(2).Controls("Label4").Caption = strCon1
      .Sections(2).Controls("Label5").Caption = Val(strCon1) - 1
      .Sections(2).Controls("Label6").Caption = Val(strCon1) - 2
      .Sections(2).Controls("Label7").Caption = Val(strCon1) - 3
      .Sections(2).Controls("Label8").Caption = Val(strCon1) - 4
      .Sections(2).Controls("Label13").Caption = (Val(strCon1) - 5) & " ~ " & (Val(strCon1) - 9)
      .Sections(2).Controls("Label14").Caption = Val(strCon1) - 10 & " ~ "
      Select Case intSelect
         Case 2
            .Sections(3).Visible = False
            .Sections(4).Visible = False
            .Sections(5).Visible = False
         Case Else
            .Sections(3).Visible = True
            .Sections(4).Visible = True
            .Sections(5).Visible = True
      End Select
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc24b0(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt24b0.State = adStateOpen Then
      AccEnvironment.rsRpt24b0.Requery
   End If
   With Rptacc24b0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label4").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label12").Caption = strPrintDate
'      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc3410(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strCon7 As String, strMem As String, strPrintDate As String)
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
   
   If AccEnvironment.rsRpt3410.State = adStateOpen Then
      AccEnvironment.rsRpt3410.Requery
   End If
   With Rptacc3410
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label32").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label15").Caption = strCon5
      .Sections(2).Controls("Label26").Caption = strCon6
      .Sections(2).Controls("Label30").Caption = strCon7
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc3420(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strCon7 As String, strCon8 As String, strMem As String, strPrintDate As String)
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
   
   If AccEnvironment.rsRpt3420.State = adStateOpen Then
      AccEnvironment.rsRpt3420.Requery
   End If
   With Rptacc3420
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label35").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label12").Caption = strCon5
      .Sections(2).Controls("Label27").Caption = strCon6
      .Sections(2).Controls("Label31").Caption = strCon7
      .Sections(2).Controls("Label33").Caption = strCon8
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc3430(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strMem As String, strPrintDate As String)
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
   
   If AccEnvironment.rsRpt3430.State = adStateOpen Then
      AccEnvironment.rsRpt3430.Requery
   End If
   With Rptacc3430
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label22").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label33").Caption = strCon5
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc3440(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt3440_Grouping.State = adStateOpen Then
      AccEnvironment.rsRpt3440_Grouping.Requery
   End If
   With Rptacc3440
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc3450(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strCon7 As String, strCon8 As String, strCon9 As String, strCon10 As String, strCon11 As String, strCon12 As String, strMem As String, strPrintDate As String)
'   If AccEnvironment.Connections.Item(1).State = adStateOpen Then
'      AccEnvironment.Connections.Item(1).Close
'   End If
'   AccEnvironment.Connections.Item(1).ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON;Persist Security Info=True"
'   AccEnvironment.Connections.Item(1).Open
   If AccEnvironment.rsRpt3450.State = adStateOpen Then
      AccEnvironment.rsRpt3450.Close
   End If
   AccEnvironment.rsRpt3450.CursorLocation = adUseClient
   AccEnvironment.rsRpt3450.Open "SELECT * FROM ACCRPT305 WHERE R30501 = '" & strCon12 & "' AND R30510 = '" & strCon3 & "' AND R30511 = '" & strCon4 & "' ORDER BY R30507 ASC, R30502 ASC", AccEnvironment.Connections.Item(1), adOpenStatic, adLockReadOnly
   With Rptacc3450
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("Label10").Caption = strCon4
      .Sections(2).Controls("Label17").Caption = strCon5
      .Sections(2).Controls("Label19").Caption = strCon6
      .Sections(2).Controls("Label21").Caption = strCon7
      .Sections(2).Controls("Label23").Caption = strCon8
      .Sections(2).Controls("Label35").Caption = strCon9
      .Sections(2).Controls("Label37").Caption = strCon10
      .Sections(2).Controls("Label39").Caption = strCon11
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
'   AccEnvironment.rsRpt3450.Close
'   If AccEnvironment.Connections.Item(1).State = adStateOpen Then
'      AccEnvironment.Connections.Item(1).Close
'   End If
End Sub

Public Sub Acc3460(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strCon7 As String, strCon8 As String, strCon9 As String, strCon10 As String, strCon11 As String, strCon12 As String, strCon13 As String, strMem As String, strPrintDate As String)
'   If AccEnvironment.Connections.Item(1).State = adStateOpen Then
'      AccEnvironment.Connections.Item(1).Close
'   End If
'   AccEnvironment.Connections.Item(1).ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON;Persist Security Info=True"
'   AccEnvironment.Connections.Item(1).Open
   If AccEnvironment.rsRpt3460.State = adStateOpen Then
      AccEnvironment.rsRpt3460.Close
   End If
   AccEnvironment.rsRpt3460.CursorLocation = adUseClient
   'Modify By Cheng 2002/01/21
'   AccEnvironment.rsRpt3460.Open "select * from accrpt307 where r30701 = '" & strCon13 & "' and r30710 = '" & strCon5 & "' and r30711 = '" & strCon6 & "' order by r30707 asc, r30702 asc"
   'Modify By Cheng 2002/01.31
'   AccEnvironment.rsRpt3460.Open "select * from accrpt307 where r30701 = '" & strCon13 & "' and r30710 = '" & strCon5 & "' and r30711 = '" & strCon6 & "' order by r30709 asc, r30706 asc, r30707 asc, r30702 asc, r30704 asc, r30705 asc, r30711 asc"
   AccEnvironment.rsRpt3460.Open "select * from accrpt307 where r30701 = '" & strCon13 & "' and r30710 = '" & strCon5 & "' and r30712 = '" & strCon6 & "' order by r30709 asc, r30706 asc, r30707 asc, r30702 asc, r30704 asc, r30705 asc, r30711 asc"
   With Rptacc3460
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label36").Caption = strCon3
      .Sections(2).Controls("Label38").Caption = strCon4
      .Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      'Modify By Cheng 2002/01/30
'      .Sections(2).Controls("Label6").Caption = strCon5
'      .Sections(2).Controls("Label10").Caption = strCon6
'      .Sections(2).Controls("Label17").Caption = strCon7
'      .Sections(2).Controls("Label19").Caption = strCon8
'      .Sections(2).Controls("Label21").Caption = strCon9
'      .Sections(2).Controls("Label23").Caption = strCon10
'      .Sections(2).Controls("Label34").Caption = strCon11
      .Sections(2).Controls("Label40").Caption = strCon12
      'Add By Cheng 2002/01/31
      .Sections(2).Controls("Label27").Visible = True
      .Sections("Rpt3460_Detail").Controls("text4").Visible = True
      .Sections("ReportFooter").Controls("funSumP").Visible = True
      .Sections("ReportFooter").Controls("linP1").Visible = True
      .Sections("ReportFooter").Controls("linP2").Visible = True
      .Sections(2).Controls("Label26").Visible = True
      .Sections("Rpt3460_Detail").Controls("text3").Visible = True
      .Sections("ReportFooter").Controls("funSumR").Visible = True
      .Sections("ReportFooter").Controls("linR1").Visible = True
      .Sections("ReportFooter").Controls("linR2").Visible = True
      'Add By Cheng 2002/01/30
      '若列印應收票據
      If Left(.Sections(2).Controls("Label40").Caption, 1) = "R" Then
         .Sections(2).Controls("Label27").Visible = False
         .Sections("Rpt3460_Detail").Controls("text4").Visible = False
         .Sections("ReportFooter").Controls("funSumP").Visible = False
         .Sections("ReportFooter").Controls("linP1").Visible = False
         .Sections("ReportFooter").Controls("linP2").Visible = False
      '若列印應付票據
      ElseIf Left(.Sections(2).Controls("Label40").Caption, 1) = "P" Then
         .Sections(2).Controls("Label26").Visible = False
         .Sections("Rpt3460_Detail").Controls("text3").Visible = False
         .Sections("ReportFooter").Controls("funSumR").Visible = False
         .Sections("ReportFooter").Controls("linR1").Visible = False
         .Sections("ReportFooter").Controls("linR2").Visible = False
      End If
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
'   AccEnvironment.rsRpt3460.Close
'   If AccEnvironment.Connections.Item(1).State = adStateOpen Then
'      AccEnvironment.Connections.Item(1).Close
'   End If
End Sub

Public Sub Acc3470(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strMem As String, strPrintDate As String)
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
   
   If AccEnvironment.rsRpt3470.State = adStateOpen Then
      AccEnvironment.rsRpt3470.Requery
   End If
   With Rptacc3470
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label23").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label6").Caption = strCon1
      .Sections(2).Controls("Label1").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strCon3
      .Sections(2).Controls("Label19").Caption = strCon4
      .Sections(2).Controls("Label21").Caption = strCon5
      .Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub


Public Sub Acc3480(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strCon7 As String, strCon8 As String, strCon9 As String, strMem As String, strPrintDate As String)
'   If AccEnvironment.Connections.Item(1).State = adStateOpen Then
'      AccEnvironment.Connections.Item(1).Close
'   End If
'   AccEnvironment.Connections.Item(1).ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON;Persist Security Info=True"
'   AccEnvironment.Connections.Item(1).Open
   
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
   
   If AccEnvironment.rsRpt3480.State = adStateOpen Then
      AccEnvironment.rsRpt3480.Close
   End If
   AccEnvironment.rsRpt3480.CursorLocation = adUseClient
   AccEnvironment.rsRpt3480.Open "select * from accrpt309 where r30901 = '" & strCon9 & "' and r30912 = '" & strCon7 & "' order by r30909 asc, r30905 asc"
   With Rptacc3480
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label37").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label9").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strCon3
      .Sections(2).Controls("Label5").Caption = strCon4
      .Sections(2).Controls("Label7").Caption = strCon5
      .Sections(2).Controls("Label13").Caption = strCon6
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .Sections(2).Controls("Label31").Caption = strCon7
      .Sections(2).Controls("Label34").Caption = strCon8
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
'   AccEnvironment.rsRpt3480.Close
'   If AccEnvironment.Connections.Item(1).State = adStateOpen Then
'      AccEnvironment.Connections.Item(1).Close
'   End If
End Sub

'Modify By Cheng 2003/05/27
'加傳入參數--後續處理
'Public Sub Acc3490(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
Public Sub Acc3490(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String, strFollow As String)
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
      
   If AccEnvironment.rsRpt3490.State = adStateOpen Then
      AccEnvironment.rsRpt3490.Requery
   End If
   With Rptacc3490
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label15").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
        'Add By Cheng 2003/05/27
        '後續處理
        .Sections(2).Controls("Label10").Caption = strFollow
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc34a0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
   
   If AccEnvironment.rsRpt34a0.State = adStateOpen Then
      AccEnvironment.rsRpt34a0.Requery
   End If
   With Rptacc34a0
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label24").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc34b0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt34b0.State = adStateOpen Then
      AccEnvironment.rsRpt34b0.Requery
   End If
   With Rptacc34b0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc34c0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt34c0.State = adStateOpen Then
      AccEnvironment.rsRpt34c0.Requery
   End If
   With Rptacc34c0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label15").Caption = strCon3
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc34d0(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
   
   If AccEnvironment.rsRpt34d0.State = adStateOpen Then
      AccEnvironment.rsRpt34d0.Requery
   End If
   With Rptacc34d0
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label7").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc34e0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt34e0.State = adStateOpen Then
      AccEnvironment.rsRpt34e0.Requery
   End If
   With Rptacc34e0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("label8").Caption = strCon4
      .Sections(2).Controls("Label28").Caption = strCon5
      .Sections(2).Controls("Label30").Caption = strCon6
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc34g0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   '20140120START 增加公司別  Add By eric
   Dim arrcon1
   arrcon1 = Split(strTitle, "-")
   '20140120END
   
   If AccEnvironment.rsRpt34g0.State = adStateOpen Then
      AccEnvironment.rsRpt34g0.Requery
   End If
   With Rptacc34g0
      '20140120START 增加公司別 Modify By eric
      .Sections(2).Controls("Label23").Caption = arrcon1(1)
      .Sections(2).Controls("Label3").Caption = arrcon1(0)
      '.Sections(2).Controls("Label3").Caption = strTitle
      '20140120END
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label6").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub
Public Sub Acc34h0(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt34h0.State = adStateOpen Then
      AccEnvironment.rsRpt34h0.Requery
   End If
   With Rptacc34h0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label1").Caption = strCon1
      .Sections(2).Controls("Label2").Caption = strCon2
      .Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc34i0(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt34i0.State = adStateOpen Then
      AccEnvironment.rsRpt34i0.Requery
   End If
   With Rptacc34i0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label6").Caption = strCon1
      .Sections(2).Controls("Label8").Caption = strCon2
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc4270(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strTmp As String, strPrintDate As String)
   Dim strTp() As String
   
   'Modify by Amy 2013/11/05 因與frmacc41f0共用accrpt423 Table 只抓strUserNum 資料
   strTp = Split(strMem, "-")
'   If AccEnvironment.rsRpt4270.State = adStateOpen Then
'      AccEnvironment.rsRpt4270.Requery
'   End If
   If AccEnvironment.rsRpt4270.State = adStateOpen Then
      AccEnvironment.rsRpt4270.Close
   End If
   AccEnvironment.rsRpt4270.CursorLocation = adUseClient
   AccEnvironment.rsRpt4270.Open "SELECT * FROM ACCRPT423 WHERE R42301 = '" & strTp(0) & "' ORDER BY R42302,R42310 ", AccEnvironment.Connections.Item(1), adOpenStatic, adLockReadOnly
   
   With Rptacc4270
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strCon3
      .Sections(2).Controls("Label11").Caption = strCon4
      .Sections(2).Controls("Label23").Caption = strTmp
      'Modify by Amy 2013/11/05
      '.Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label8").Caption = strTp(1)
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc4410(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt4410_Grouping.State = adStateOpen Then
      AccEnvironment.rsRpt4410_Grouping.Requery
   End If
   With Rptacc4410
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strCon3
      .Sections(2).Controls("Label11").Caption = strCon4
      .Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc4420(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strCon7 As String, strCon8 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt4420.State = adStateOpen Then
      AccEnvironment.rsRpt4420.Requery
   End If
   With Rptacc4420
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strCon3
      .Sections(2).Controls("Label7").Caption = strCon4
      .Sections(2).Controls("Label9").Caption = strCon5
      .Sections(2).Controls("Label10").Caption = strCon6
      .Sections(2).Controls("Label13").Caption = strCon7
      .Sections(2).Controls("Label15").Caption = strCon8
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc4440(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt4440.State = adStateOpen Then
      AccEnvironment.rsRpt4440.Requery
   End If
   With Rptacc4440
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strCon3
      .Sections(2).Controls("Label7").Caption = strCon4
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc4450(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt4450.State = adStateOpen Then
      AccEnvironment.rsRpt4450.Requery
   End If
   strCondition1 = strCon5
   With Rptacc4450
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label13").Caption = strCon3
      .Sections(2).Controls("Label15").Caption = strCon4
      .Sections(2).Controls("Label9").Caption = strCon5
      .Sections(2).Controls("Label1").Caption = strCon6
      .Sections(2).Controls("Label17").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
   Unload Rptacc4450
End Sub

Public Sub Acc4460(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt4460.State = adStateOpen Then
      AccEnvironment.rsRpt4460.Requery
   End If
   With Rptacc4460
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strCon3
      .Sections(2).Controls("Label11").Caption = strCon4
      .Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc4470(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt4470.State = adStateOpen Then
      AccEnvironment.rsRpt4470.Requery
   End If
   With Rptacc4470
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label13").Caption = strCon3
      .Sections(2).Controls("Label15").Caption = strCon4
      .Sections(2).Controls("Label2").Caption = strCon5
      .Sections(2).Controls("Label7").Caption = strCon6
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc4480(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   'Modify by Amy 2015/04/09 財務處可能同時兩個執行此報表,造成資料錯誤
   Dim strTp() As String
'   If AccEnvironment.rsRpt4480.State = adStateOpen Then
'      AccEnvironment.rsRpt4480.Requery
'   End If

   strTp = Split(strMem, "-")
   If AccEnvironment.rsRpt4480.State = adStateOpen Then
      AccEnvironment.rsRpt4480.Close
   End If
   AccEnvironment.rsRpt4480.CursorLocation = adUseClient
   AccEnvironment.rsRpt4480.Open "Select * From accrpt409 Where r40901 = '" & strTp(0) & "' order by r40901 asc, r40908 asc ", AccEnvironment.Connections.Item(1), adOpenStatic, adLockReadOnly
   'end 2015/04/09
   
   With Rptacc4480
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label10").Caption = strTp(1) 'Modify by Amy 2015/04/09 原:strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc4490(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt4490.State = adStateOpen Then
      AccEnvironment.rsRpt4490.Requery
   End If
   With Rptacc4490
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strCon3
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc44a0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt44a0.State = adStateOpen Then
      AccEnvironment.rsRpt44a0.Requery
   End If
   With Rptacc44a0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc44b0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   'Modify by Amy 2015/04/15 財務處可能同時兩個人執行此報表,造成資料錯誤
    Dim strTp() As String
'   If AccEnvironment.rsRpt44b0.State = adStateOpen Then
'      AccEnvironment.rsRpt44b0.Requery
'   End If
  strTp = Split(strMem, "-")
   If AccEnvironment.rsRpt44b0.State = adStateOpen Then
      AccEnvironment.rsRpt44b0.Close
   End If
   AccEnvironment.rsRpt44b0.CursorLocation = adUseClient
   AccEnvironment.rsRpt44b0.Open "Select * From accrpt412 Where r41201 ='" & strTp(0) & "' order by r41201 asc, r41202 asc", AccEnvironment.Connections.Item(1), adOpenStatic, adLockReadOnly
   'end 2014/04/15
   With Rptacc44b0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label10").Caption = strTp(1) 'Modify by Amy 2015/04/15 原:strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc44c0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt44c0.State = adStateOpen Then
      AccEnvironment.rsRpt44c0.Requery
   End If
   With Rptacc44c0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strCon3
      .Sections(2).Controls("Label8").Caption = strCon4
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc44d0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt44c0.State = adStateOpen Then
      AccEnvironment.rsRpt44c0.Requery
   End If
   With Rptacc44d0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label23").Caption = strCon3
      .Sections(2).Controls("Label24").Caption = strCon4
      .Sections(2).Controls("Label2").Caption = strCon5
      .Sections(2).Controls("Label8").Caption = strCon6
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc44e0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt44e0.State = adStateOpen Then
      AccEnvironment.rsRpt44e0.Requery
   End If
   With Rptacc44e0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label10").Caption = strCon3
      .Sections(2).Controls("Label11").Caption = strCon4
      .Sections(2).Controls("Label8").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc44k0(strTitle As String, strCon1 As String, strCon2 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt44k0.State = adStateOpen Then
      AccEnvironment.rsRpt44k0.Requery
   End If
   With Rptacc44k0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label8").Caption = strCon1
      .Sections(2).Controls("Label1").Caption = strCon2
      .Sections(2).Controls("Label2").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc44l0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strCon5 As String, strCon6 As String, strCon7 As String, strCon8 As String, strCon9 As String, strMem As String, strPrintDate As String)
'   If AccEnvironment.Connections.Item(1).State = adStateOpen Then
'      AccEnvironment.Connections.Item(1).Close
'   End If
'   AccEnvironment.Connections.Item(1).ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON;Persist Security Info=True"
'   AccEnvironment.Connections.Item(1).Open
   If AccEnvironment.rsRpt44l0.State = adStateOpen Then
      AccEnvironment.rsRpt44l0.Close
   End If
   AccEnvironment.rsRpt44l0.CursorLocation = adUseClient
   '2014/3/3 modify by sonia
   'AccEnvironment.rsRpt44l0.Open "select * from accrpt418 where r41801 = '" & strCon9 & "' and  r41802 = '" & strCon7 & "' and r41803 = '" & strCon8 & "' order by r41804 asc"
   If strCon7 <> "" Then
      AccEnvironment.rsRpt44l0.Open "select * from accrpt418 where r41801 = '" & strCon9 & "' and  r41802 = '" & IIf(strCon7 = "2", "J", strCon7) & "' and r41803 = '" & strCon8 & "' order by r41804 asc"
   Else
      AccEnvironment.rsRpt44l0.Open "select * from accrpt418 where r41801 = '" & strCon9 & "' and r41803 = '" & strCon8 & "' order by r41804 asc"
   End If
   '2014/3/3 end
   With Rptacc44l0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label5").Caption = strCon1
      .Sections(2).Controls("Label6").Caption = strCon2
      .Sections(2).Controls("Label23").Caption = strCon3
      .Sections(2).Controls("Label24").Caption = strCon4
      .Sections(2).Controls("Label2").Caption = strCon5
      .Sections(2).Controls("Label8").Caption = strCon6
      .Sections(2).Controls("Label10").Caption = strMem
      .Sections(2).Controls("Label19").Caption = strPrintDate
      .Sections(2).Controls("Label34").Caption = strCon7
      .Sections(2).Controls("Label36").Caption = strCon8
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
'   AccEnvironment.rsRpt44l0.Close
'   If AccEnvironment.Connections.Item(1).State = adStateOpen Then
'      AccEnvironment.Connections.Item(1).Close
'   End If
End Sub

Public Sub Acc44m0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
   If AccEnvironment.rsRpt44m0_Grouping.State = adStateOpen Then
      AccEnvironment.rsRpt44m0_Grouping.Requery
   End If
   With Rptacc44m0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label20").Caption = strCon1
      .Sections(2).Controls("Label21").Caption = strCon2
      .Sections(2).Controls("Label8").Caption = strCon3
      .Sections(2).Controls("Label1").Caption = strCon4
      .Sections(2).Controls("Label2").Caption = strMem
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
End Sub

Public Sub Acc44o0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String, strMem As String, strPrintDate As String)
'   AccEnvironment.Connections.Item(1).ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON;Persist Security Info=True"
'   AccEnvironment.Connections.Item(1).Open
   If AccEnvironment.rsRpt44o0.State = adStateOpen Then
      AccEnvironment.rsRpt44o0.Close
   End If
   AccEnvironment.rsRpt44o0.CursorLocation = adUseClient
   AccEnvironment.rsRpt44o0.Open "SELECT * FROM ACCRPT420 WHERE R42002 = '" & strCon2 & "' ORDER BY R42002 ASC", AccEnvironment.Connections.Item(1), adOpenStatic, adLockReadOnly
   With Rptacc44o0
      .Sections(2).Controls("Label3").Caption = strTitle
      .Sections(2).Controls("Label20").Caption = strCon1
      .Sections(2).Controls("Label16").Caption = strCon2
      .Sections(2).Controls("Label39").Caption = strCon3
      .Sections(2).Controls("Label7").Caption = strCon4
      .Sections(2).Controls("Label14").Caption = strPrintDate
      .LeftMargin = intLeftMargin
      .WindowState = 2
      .PrintReport
   End With
'   AccEnvironment.rsRpt44o0.Close
'   AccEnvironment.Connections.Item(1).Close
End Sub
'Modify by Morgan 2004/421
Public Sub Acc44p0(strTitle As String, strCon1 As String, strCon2 As String, strCon3 As String, strMem As String, strPrintDate As String)

   Dim arrcon1() As String
   arrcon1 = Split(strCon1, Chr(23))
   
   If AccEnvironment.rsRpt44p0.State = adStateOpen Then
      AccEnvironment.rsRpt44p0.Close
   End If
   AccEnvironment.rsRpt44p0.CursorLocation = adUseClient
   AccEnvironment.rsRpt44p0.Open "SELECT * FROM ACCRPT421 WHERE R42103 = '" & strCon2 & "' ORDER BY R42102 ASC"
   With Rptacc44p0
      .LeftMargin = intLeftMargin
      
      If arrcon1(1) = "Y" Then
         .Sections(2).Controls("Label3").Caption = strTitle
         .Sections(2).Controls("Label20").Caption = arrcon1(0)
         .Sections(2).Controls("Label4").Caption = strCon2
         .Sections(2).Controls("Label6").Caption = strCon3
         If arrcon1(2) = "Y" Then
            .Sections(2).Controls("Label11").Visible = True
            .Sections(3).Controls("Text4").Visible = True
         End If
         .Sections(2).Controls("Line1").Visible = False
         .Sections(2).Controls("Label2").Visible = False
         .Sections(2).Controls("Label7").Visible = False
         .Sections(2).Controls("Label13").Visible = False
         .Sections(2).Controls("Label14").Visible = False
         .Sections(2).Controls("Label15").Visible = False
         .Sections(2).Controls("Label16").Visible = False
         'Modify By Sindy 2018/2/27
         '.ExportReport .ExportFormats(3), "C:\TAIE.TMP", True, False
         .ExportReport .ExportFormats(3), "c:\xls\TAIE.TMP", True, False
         '2018/2/27 END
      Else
         .Sections(2).Controls("Label3").Caption = strTitle
         .Sections(2).Controls("Label20").Caption = arrcon1(0)
         If arrcon1(2) = "Y" Then
            .Sections(2).Controls("Label11").Visible = True
            .Sections(2).Controls("Line1").Width = .Sections(2).Controls("Line1").Width + 2200
            .Sections(3).Controls("Text4").Visible = True
         End If
         .Sections(2).Controls("Label2").Caption = strMem
         .Sections(2).Controls("Label14").Caption = strPrintDate
         .Sections(2).Controls("Label4").Caption = strCon2
         .Sections(2).Controls("Label6").Caption = strCon3
         .PrintReport
      End If
      
   End With
End Sub

Private Sub Class_Terminate()
   Set AccEnvironment = Nothing
   End
End Sub
