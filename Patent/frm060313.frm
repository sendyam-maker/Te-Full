VERSION 5.00
Begin VB.Form frm060313 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人准駁明細表"
   ClientHeight    =   2415
   ClientLeft      =   1980
   ClientTop       =   3780
   ClientWidth     =   3465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3465
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2448
      TabIndex        =   8
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1656
      TabIndex        =   7
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1260
      TabIndex        =   6
      Top             =   2076
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   984
      TabIndex        =   5
      Top             =   1752
      Width           =   210
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   984
      TabIndex        =   4
      Top             =   1428
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2100
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1104
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   984
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1104
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   984
      TabIndex        =   1
      Top             =   792
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   984
      TabIndex        =   0
      Top             =   468
      Width           =   2424
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   2016
      TabIndex        =   18
      Top             =   1476
      Width           =   1452
   End
   Begin VB.Line Line1 
      X1              =   1980
      X2              =   2070
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:印)"
      Height          =   180
      Index           =   8
      Left            =   1548
      TabIndex        =   17
      Top             =   2124
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "(1.承辦人 2.准駁日)"
      Height          =   180
      Index           =   7
      Left            =   1260
      TabIndex        =   16
      Top             =   1800
      Width           =   1728
   End
   Begin VB.Label Label1 
      Caption         =   "(1.准 2.駁)"
      Height          =   180
      Index           =   6
      Left            =   1320
      TabIndex        =   15
      Top             =   852
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "是否列印明細："
      Height          =   180
      Index           =   5
      Left            =   12
      TabIndex        =   14
      Top             =   2136
      Width           =   1308
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   4
      Left            =   12
      TabIndex        =   13
      Top             =   1812
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   3
      Left            =   12
      TabIndex        =   12
      Top             =   1488
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "准駁日期："
      Height          =   180
      Index           =   2
      Left            =   12
      TabIndex        =   11
      Top             =   1164
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "准駁代碼："
      Height          =   180
      Index           =   1
      Left            =   12
      TabIndex        =   10
      Top             =   840
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   12
      TabIndex        =   9
      Top             =   516
      Width           =   996
   End
End
Attribute VB_Name = "frm060313"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String
Dim iPrint As Integer, Page As Integer, strTemp(0 To 10) As String
Dim PLeft(0 To 10) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrTemp4(0 To 4) As String, StrTemp5(0 To 4) As String

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Len(txt1(1)) = 0 Then
             s = MsgBox("准駁代碼不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             Exit Sub
         Else
            'Add By Cheng 2002/03/20
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
               Me.txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
             
             If Len(txt1(3)) = 0 Then
                 s = MsgBox("准駁日期區間不可空白!!", , "USER 輸入錯誤")
                 If Len(txt1(2)) = 0 Then txt1(2).SetFocus
                 Exit Sub
             Else
                 If Len(txt1(5)) = 0 Then
                     s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                     txt1(5).SetFocus
                     Exit Sub
                 Else
                     Me.Enabled = False
                     Screen.MousePointer = vbHourglass
                     Process
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                 End If
             End If
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
cnnConnection.Execute "DELETE FROM R060313 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
'edit by nickc 2007/02/08
'StrSQL6 = ""

'系統類別
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/13
End If
'准駁通知日
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA20>=" & Val(ChangeTStringToWString(txt1(2)))
End If
If Len(txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA20<=" & Val(ChangeTStringToWString(txt1(3)))
End If
If Len(txt1(2)) <> 0 Or Len(txt1(3)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/13
End If
If txt1(1) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.准" 'Add By Sindy 2010/12/13
ElseIf txt1(1) = "2" Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.駁" 'Add By Sindy 2010/12/13
End If
If Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & lbl1 'Add By Sindy 2010/12/13
End If
If Trim(txt1(4)) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & "1.承辦人" 'Add By Sindy 2010/12/13
ElseIf Trim(txt1(4)) = "2" Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & "2.准駁日" 'Add By Sindy 2010/12/13
End If
If Trim(txt1(6)) = "Y" Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(6)  'Add By Sindy 2010/12/13
End If
'If Len(txt1(4)) <> 0 Then
'    strSQL1 = strSQL1 + " AND CP14='" & txt1(4) & "' "
'End If

'92.04.02 nick add left join
'strSQL = "SELECT distinct cp14," & SQLDate("PA20") & ",PA01||'-'||PA02||'-'||PA03||'-'||PA04," & _
   SQLDate("CP27") & ",NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),NVL(NVL(N2.NA03,N2.NA04)," & _
   "NVL(N1.NA03,N1.NA04)),DECODE(PA16,'1','准','2','駁',NULL,' '),pa01,pa02,pa03,pa04," & _
   "dEcode(pa09,'000',CPM03,CPM04),CP10 FROM PATENT,CASEPROGRESS,NATION N1,NATION N2," & _
   "PATENTTRADEMARKMAP,CUSTOMER,FAGENT,CASEPROPERTYMAP WHERE " & _
   "PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND " & _
   "PA16=CP24 AND PA20=CP25 AND " & _
   "SUBSTR(PA26,1,8)=CU01(+) AND " & _
   "decode(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND " & _
   "decode(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND CU10=N1.NA01(+) AND " & _
   "FA10=N2.NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND " & _
   "CP01=cpm01(+) AND cp10=CPM02(+) " & strSQL1 & "  " & _
   ""
strSql = "SELECT distinct cp14," & SQLDate("PA20") & ",PA01||'-'||PA02||'-'||PA03||'-'||PA04," & _
   SQLDate("CP27") & ",NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),NVL(NVL(N2.NA03,N2.NA04)," & _
   "NVL(N1.NA03,N1.NA04)),DECODE(PA16,'1','准','2','駁',NULL,' '),pa01,pa02,pa03,pa04," & _
   "dEcode(pa09,'000',CPM03,CPM04),CP10 FROM PATENT,CASEPROGRESS,NATION N1,NATION N2," & _
   "PATENTTRADEMARKMAP,CUSTOMER,FAGENT,CASEPROPERTYMAP WHERE " & _
   "PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND " & _
   "PA16=CP24(+) AND PA20=CP25(+) AND " & _
   "SUBSTR(PA26,1,8)=CU01(+) AND " & _
   "decode(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND " & _
   "decode(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND CU10=N1.NA01(+) AND " & _
   "FA10=N2.NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND " & _
   "CP01=cpm01(+) AND cp10=CPM02(+) " & strSQL1 & "  " & _
   ""

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    'edit by nickc 2007/02/08
    'k = 0
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/13
    With adoRecordset
      .MoveFirst
      Do While .EOF = False
          
      For i = 0 To 7
          strTemp(i) = CheckStr(.Fields(i))
      Next i
      '910621 Sieg
      'strSQL = "select st02 from caseprogress,staff where cp01='" & CheckStr(.Fields(8)) & "' and cp02='" & CheckStr(.Fields(9)) & "' and cp03='" & CheckStr(.Fields(10)) & "' and cp04='" & CheckStr(.Fields(11)) & "' and cp14=st01(+) and (cp10='203' or cp10='204') "
      
      Select Case .Fields("CP10")
      Case "101", "102", "103", "104", "105"
         strSql = "select count(*) from caseprogress,staff where cp01='" & CheckStr(.Fields(8)) & "' and cp02='" & CheckStr(.Fields(9)) & "' and cp03='" & CheckStr(.Fields(10)) & "' and cp04='" & CheckStr(.Fields(11)) & "' and cp14=st01 and ST15='F22'"
         CheckOC2
         adoRecordset1.CursorLocation = adUseClient
         adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset1.Fields(0) > 0 Then
            'Modify By Cheng 2002/10/07
            '取得案件性質為"203"至"206"且發文日最大的資料
'             strSQL = "select cp14 from caseprogress where cp01='" & CheckStr(.Fields(8)) & "' and cp02='" & CheckStr(.Fields(9)) & "' and cp03='" & CheckStr(.Fields(10)) & "' and cp04='" & CheckStr(.Fields(11)) & "' and (cp10='203' or cp10='204') "
             strSql = "select cp14 from caseprogress where cp01='" & CheckStr(.Fields(8)) & "' and cp02='" & CheckStr(.Fields(9)) & "' and cp03='" & CheckStr(.Fields(10)) & "' and cp04='" & CheckStr(.Fields(11)) & "' and (cp10>='203' and cp10<='206') AND CP27 IS NOT NULL ORDER BY CP27 DESC "
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount = 0 Then
               'Modify By Cheng 2002/10/07
               '取得案件性質為"201"核稿人的資料
'               strSQL = "select cp14 from caseprogress where cp01='" & CheckStr(.Fields(8)) & "' and cp02='" & CheckStr(.Fields(9)) & "' and cp03='" & CheckStr(.Fields(10)) & "' and cp04='" & CheckStr(.Fields(11)) & "' and cp10='205'"
               strSql = "select Ep04 from caseprogress,ENGINEERPROGRESS where CP09=EP02(+) AND cp01='" & CheckStr(.Fields(8)) & "' and cp02='" & CheckStr(.Fields(9)) & "' and cp03='" & CheckStr(.Fields(10)) & "' and cp04='" & CheckStr(.Fields(11)) & "' and cp10='201' AND EP04 IS NOT NULL "
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount = 0 Then
                  strSql = "select cp14 from caseprogress where cp01='" & CheckStr(.Fields(8)) & "' and cp02='" & CheckStr(.Fields(9)) & "' and cp03='" & CheckStr(.Fields(10)) & "' and cp04='" & CheckStr(.Fields(11)) & "' and cp10='201'"
                  CheckOC2
                  adoRecordset1.CursorLocation = adUseClient
                  adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If adoRecordset1.RecordCount = 0 Then
                     strTemp(0) = "null"
                  Else
                     strTemp(0) = CheckStr(adoRecordset1.Fields(0))
                  End If
               Else
                  strTemp(0) = CheckStr(adoRecordset1.Fields(0))
               End If
             Else
               strTemp(0) = CheckStr(adoRecordset1.Fields(0))
             End If
         End If
         CheckOC2
      End Select
      '若有輸入承辦人
      If txt1(4).Text <> "" Then
          If strTemp(0) = txt1(4).Text Then
            'Modify By Cheng 2002/10/08
            '加存案件性質代號,系統類別
'             strSQL = "insert into r060313 values('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & CheckStr(.Fields(12)) & "','" & strUserNum & "') "
             strSql = "insert into r060313 values('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & CheckStr(.Fields(12)) & "','" & strUserNum & "','" & .Fields("CP10") & "','" & .Fields("PA01") & "') "
             cnnConnection.Execute strSql
          End If
      '若未輸入承辦人
      Else
            'Modify By Cheng 2002/10/08
            '加存案件性質代號,系統類別
'          strSQL = "insert into r060313 values('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & CheckStr(.Fields(12)) & "','" & strUserNum & "') "
          strSql = "insert into r060313 values('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & CheckStr(.Fields(12)) & "','" & strUserNum & "','" & .Fields("CP10") & "','" & .Fields("PA01") & "') "
          cnnConnection.Execute strSql
      End If
       .MoveNext
      Loop
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/13
    ShowNoData
    Exit Sub
End If
CheckOC
PrintData
End Sub

Sub PrintData()
strSQL1 = ""
Select Case Val(txt1(1))
Case 1
      strSQL1 = strSQL1 + " AND r046008='准' "
Case 2
      strSQL1 = strSQL1 + " AND r046008='駁' "
Case Else
End Select
If txt1(5) = "1" Then
   'Modify By Cheng 2002/10/08
'    strSQL = "select st02,r046002,r046003,r046004,r046005,r046006,r046007,r046008,r046001 from r060313,staff where r046001=st01(+) and id='" & strUserNum & "' " & strSQL1 & " order by decode(r046001,'','0',r046001),r046003 "
    strSql = "select st02,r046002,r046003,r046004,r046005,r046009,r046006,r046007,r046008,r046001,r046010 from r060313,staff where r046001=st01(+) and id='" & strUserNum & "' " & strSQL1 & " order by decode(r046001,'','0',r046001),r046003,r046010 "
Else
   'Modify By Cheng 2002/10/08
'    strSQL = "select st02,r046002,r046003,r046004,r046005,r046006,r046007,r046008,r046001 from r060313,staff where r046001=st01(+) and id='" & strUserNum & "' " & strSQL1 & " order by decode(r046002,'','0',r046002),r046003 "
    strSql = "select st02,r046002,r046003,r046004,r046005,r046009,r046006,r046007,r046008,r046001 from r060313,staff where r046001=st01(+) and id='" & strUserNum & "' " & strSQL1 & " order by decode(r046002,'','0',r046002),r046003 "
End If
CheckOC
SavDay1 = " "
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
         'Modify By Cheng 2002/10/08
'        SavDay1 = CheckStr(.Fields(8))
        SavDay1 = CheckStr(.Fields(9))
        SavDay2 = CheckStr(.Fields(0))
        PrintTitle
        PrintTitle1
        Do While .EOF = False
'            For i = 0 To 7
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Cheng 2002/10/08
'            If SavDay1 <> CheckStr(.Fields(8)) Then
            If SavDay1 <> CheckStr(.Fields(9)) Then
                If txt1(5) = "1" Then
                    Printer.CurrentX = 500
                    Printer.CurrentY = iPrint
                    If txt1(6) <> "Y" Then
                        Printer.Print SavDay2
                    Else
                        Printer.Print String(250, "-")
                    End If
                    iPrint = iPrint + 300
                    If iPrint >= 10000 Then
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                        PrintTitle1
                    End If
                    PrintEnd
                    PrintTitle1
'                    SavDay1 = CheckStr(.Fields(8))
                    SavDay1 = CheckStr(.Fields(9))
                    SavDay2 = strTemp(0)
                End If
            End If
            strTemp(0) = StrToStr(strTemp(0), 3)
'            strTemp(4) = StrToStr(strTemp(4), 20)
            strTemp(4) = StrToStr(strTemp(4), 16)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(7) = StrToStr(strTemp(7), 4)
            If txt1(6) = "Y" Then
                PrintDatil
            End If
            If iPrint >= 10000 Then
                Printer.NewPage
                Page = Page + 1
                PrintTitle
                PrintTitle1
            End If
            .MoveNext
        Loop
        'Add by Morgan 2004/2/27
        '不印明細時最後一筆承辦人名稱小計
        If txt1(5).Text = "1" Then
            If txt1(6) <> "Y" Then
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print SavDay2
                iPrint = iPrint + 300
                PrintEnd
                iPrint = iPrint - 300
            End If
        End If
    End With
Else
    ShowNoData
    Exit Sub
End If
   
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
'Modify by Morgan 2004/2/27
'If txt1(5) = "1" Then
'   PrintEnd
'   'Add By Cheng 2002/10/08
'   PrintEnd2
'
'Else
'   PrintEnd2
'End If
PrintEnd2
CheckOC
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd()
'Add By Cheng 2002/10/08
Dim ii As Integer
Dim jj As Integer
Dim kk As Integer
Dim strSubTotal As String

strSQL1 = ""
Select Case Val(txt1(1))
Case 1
      strSQL1 = strSQL1 + " AND r046008='准' "
Case 2
      strSQL1 = strSQL1 + " AND r046008='駁' "
Case Else
End Select
If Len(SavDay1) <> 0 Then
   strSql = "select count(*) from r060313 where id='" & strUserNum & "' " & strSQL1 & " and r046001='" & SavDay1 & "' group by r046001 "
Else
   strSql = "select count(*) from r060313 where id='" & strUserNum & "' " & strSQL1 & " and r046001 is null  group by r046001 "
End If
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'   Printer.CurrentX = 500
'   Printer.CurrentY = iPrint
'   Printer.Print "共 " & CheckStr(adoRecordset1.Fields(0)) & " 件 "
   strSubTotal = "共 " & CheckStr(adoRecordset1.Fields(0)) & " 件 "
End If
CheckOC2
'Add By Cheng 2002/10/08
strSQL1 = ""
Select Case Val(txt1(1))
Case 1
      strSQL1 = strSQL1 + " AND r046008='准' "
Case 2
      strSQL1 = strSQL1 + " AND r046008='駁' "
Case Else
End Select
If Len(SavDay1) <> 0 Then
   strSql = "select r046001,CPM03,r046010,count(*) from r060313,casepropertymap where R046011=CPM01(+) AND R046010=CPM02(+) AND id='" & strUserNum & "' " & strSQL1 & " and r046001='" & SavDay1 & "' group by r046001,cpm03,r046010 order by r046010"
Else
   strSql = "select r046001,CPM03,r046010,count(*) from r060313,casepropertymap where R046011=CPM01(+) AND R046010=CPM02(+) AND id='" & strUserNum & "' " & strSQL1 & " and r046001 is null  group by r046001,cpm03,r046010 order by r046010"
End If
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   ii = adoRecordset1.RecordCount Mod 5
   ii = IIf(ii = 0, (adoRecordset1.RecordCount / 5), (adoRecordset1.RecordCount / 5) + 1)
   For jj = 1 To ii
      iPrint = iPrint + 300 * (jj - 1)
      For kk = 0 To 4
         Printer.CurrentX = 500 + 2500 * kk
         Printer.CurrentY = iPrint
         Printer.Print "" & Left(adoRecordset1.Fields(1).Value, 4) & " : 共 " & Val(adoRecordset1.Fields(3).Value) & " 件"
         If jj = 1 And kk = 0 Then
            Printer.CurrentX = 500 + 2500 * 5
            Printer.CurrentY = iPrint
            Printer.Print strSubTotal
         End If
         adoRecordset1.MoveNext
         If adoRecordset1.EOF Then GoTo ExitFor
      Next kk
   Next jj
End If
ExitFor:
CheckOC2
iPrint = iPrint + 300
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
End If
End Sub

Sub PrintEnd2()
'Add By Cheng 2002/10/08
Dim ii As Integer
Dim jj As Integer
Dim kk As Integer
Dim strTotal As String

'Add By Cheng 2002/10/08
If iPrint >= 10000 Then
    Printer.NewPage
    Page = Page + 1
    PrintTitle
    PrintTitle1
End If
strSQL1 = ""
Select Case Val(txt1(1))
Case 1
      strSQL1 = strSQL1 + " AND r046008='准' "
Case 2
      strSQL1 = strSQL1 + " AND r046008='駁' "
Case Else
End Select

strSql = "select count(*) from r060313 where id='" & strUserNum & "' " & strSQL1
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
'   iPrint = iPrint + 300
'   Printer.CurrentX = 500
'   Printer.CurrentY = iPrint
'   Printer.Print "總合計共 " & CheckStr(adoRecordset1.Fields(0)) & " 件"
'   iPrint = iPrint + 300
   strTotal = "總合計共 " & CheckStr(adoRecordset1.Fields(0)) & " 件"
End If
CheckOC2
'Add By Cheng 2002/10/08
If iPrint >= 10000 Then
    Printer.NewPage
    Page = Page + 1
    PrintTitle
    PrintTitle1
End If
strSQL1 = ""
Select Case Val(txt1(1))
Case 1
      strSQL1 = strSQL1 + " AND r046008='准' "
Case 2
      strSQL1 = strSQL1 + " AND r046008='駁' "
Case Else
End Select

strSql = "select CPM03,r046010,count(*) from r060313,CASEPROPERTYMAP where R046011=CPM01(+) AND R046010=CPM02(+) AND id='" & strUserNum & "' " & strSQL1 & " group by CPM03,r046010 order by r046010"
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
   ii = adoRecordset1.RecordCount Mod 5
   ii = IIf(ii = 0, (adoRecordset1.RecordCount / 5), (adoRecordset1.RecordCount / 5) + 1)
   For jj = 1 To ii
      iPrint = iPrint + 300 * (jj - 1)
      For kk = 0 To 4
         Printer.CurrentX = 500 + 2500 * kk
         Printer.CurrentY = iPrint
         Printer.Print "" & Left(adoRecordset1.Fields(0).Value, 4) & " : 共 " & Val(adoRecordset1.Fields(2).Value) & " 件"
         If jj = 1 And kk = 0 Then
            Printer.CurrentX = 500 + 2500 * 5
            Printer.CurrentY = iPrint
            Printer.Print strTotal
         End If
         adoRecordset1.MoveNext
         If adoRecordset1.EOF Then GoTo ExitFor
         If iPrint >= 10000 Then
             Printer.NewPage
             Page = Page + 1
             PrintTitle
             PrintTitle1
         End If
      Next kk
   Next jj
End If
ExitFor:
CheckOC2
End Sub

Sub PrintDatil()
'For i = 0 To 7
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "承辦人准駁明細表") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "承辦人准駁明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 7500 - (Printer.TextWidth("准駁日期：" & Format(ChangeTStringToTDateString(txt1(2)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))) / 2)
Printer.CurrentY = iPrint
Printer.Print "准駁日期：" & Format(ChangeTStringToTDateString(txt1(2)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
'Add by Morgan 2004/2/25
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Select Case txt1(1)
    Case "1"
        Printer.Print "准/駁：准"
    Case "2"
        Printer.Print "准/駁：駁"
    Case Else
        Printer.Print "准/駁："
End Select

Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300

End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1500
PLeft(2) = 3000
PLeft(3) = 5000
PLeft(4) = 6500
PLeft(5) = 10000 + 500
PLeft(6) = 11500 + 500
PLeft(7) = 13000 + 500
PLeft(8) = 14500 + 500
End Sub


Sub PrintTitle1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "准駁日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "專利種類"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "國籍"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "准/駁"
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm060313 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CmdOK(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
    Next i
Case 3
     If RunNick(txt1(2), txt1(3)) Then
       txt1(2).SetFocus
     End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
If txt1(Index) = "" Then Exit Sub
Cancel = False
Select Case Index
Case 1
     Select Case Val(txt1(1))
     Case 1, 2
     Case Else
          s = MsgBox("准駁代碼只能 1 或 2 !!", , "USER 輸入錯誤")
          Cancel = True
     End Select
Case 2, 3 '准駁日期
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
   End If
Case 4
      'edit by nickc 2007/02/08 不用 dll 了
      'If objPublicData.GetStaff(txt1(Index), strExc(0)) Then
      If ClsPDGetStaff(txt1(Index), strExc(0)) Then
         lbl1 = strExc(0)
      Else
         lbl1 = ""
         Cancel = True
      End If
Case 5
     Select Case Val(txt1(5))
     Case 1, 2
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Cancel = True
     End Select
Case 6
     Select Case txt1(6)
     Case "Y", ""
     Case Else
          s = MsgBox("是否列印明細只能輸入 Y 或空白!!", , "USER 輸入錯誤")
          Cancel = True
     End Select
End Select
If Cancel Then TextInverse txt1(Index)
End Sub
