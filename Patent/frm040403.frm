VERSION 5.00
Begin VB.Form frm040403 
   BorderStyle     =   1  '單線固定
   Caption         =   "准駁預估統計表"
   ClientHeight    =   2520
   ClientLeft      =   -48
   ClientTop       =   1740
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5040
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   4104
      TabIndex        =   8
      Top             =   36
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3180
      TabIndex        =   7
      Top             =   36
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1005
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1905
      Width           =   285
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2052
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1005
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2052
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1230
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1005
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1230
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1005
      MaxLength       =   1
      TabIndex        =   1
      Top             =   915
      Width           =   285
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1005
      TabIndex        =   0
      Top             =   585
      Width           =   3570
   End
   Begin VB.Line Line2 
      X1              =   1590
      X2              =   2346
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line Line1 
      X1              =   1620
      X2              =   2304
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "(1.發明 2.新型 3.設計 4.再審 5.訴願再訴行訴       6.異議舉發答辯 7.總件數)"
      Height          =   480
      Index           =   6
      Left            =   1335
      TabIndex        =   15
      Top             =   1920
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "(1.准 2.駁 3.全部)"
      Height          =   180
      Index           =   5
      Left            =   1380
      TabIndex        =   14
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   13
      Top             =   1935
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   12
      Top             =   1590
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "准駁日期："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   11
      Top             =   1260
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "准/駁："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   930
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   630
      Width           =   1005
   End
End
Attribute VB_Name = "frm040403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL6 As String, i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 21) As String, strTemp1 As Variant, strTemp2 As Variant, k As Integer, StrTemp7(0 To 4) As String
Dim PLeft(0 To 21) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 2) As String, BolOk As Boolean
'Add By Cheng 2002/09/12
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
      'Add By Cheng 2002/09/12
      blnClkSure = False
      
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
            'Add By Cheng 2002/03/19
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
            'Add By Cheng 2002/09/12
            If Me.txt1(2).Text <> "" And Me.txt1(3).Text <> "" Then
               If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
                  MsgBox "准駁日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(2).SetFocus
                  txt1_GotFocus 2
                  Exit Sub
               End If
            End If
             
            If Len(txt1(3)) = 0 Then
               s = MsgBox("准駁日期區間不可空白!!", , "USER 輸入錯誤")
               txt1(2).SetFocus
               txt1_GotFocus (2)
               Exit Sub
            Else
               If Len(txt1(6)) = 0 Then
                   s = MsgBox("列印順序不可空白!!", , "USER輸入錯誤")
                   txt1(6).SetFocus
                   Exit Sub
               Else
                  'Add By Cheng 2002/09/12
                  If Me.txt1(4).Text <> "" And Me.txt1(5).Text <> "" Then
                     If Me.txt1(4).Text > Me.txt1(5).Text Then
                        MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(4).SetFocus
                        txt1_GotFocus 4
                        Exit Sub
                     End If
                  End If
                   
                   Screen.MousePointer = vbHourglass
                   Me.Enabled = False
                   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
                   Process
                   Me.Enabled = True
                   Screen.MousePointer = vbDefault
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
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R040403 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/2
End If
StrSQL6 = ""
If Len(Trim(txt1(2))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP25>=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(3))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP25<=" & Val(ChangeTStringToWString(txt1(3))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/2
End If
Select Case Val(txt1(1))
Case 1
     StrSQL6 = StrSQL6 + " AND CP24='1' "
     pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.准" 'Add By Sindy 2010/12/2
Case 2
     StrSQL6 = StrSQL6 + " AND CP24='2' "
     pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.駁" 'Add By Sindy 2010/12/2
Case Else
     pub_QL05 = pub_QL05 & ";" & Label1(1) & "3.全部" 'Add By Sindy 2010/12/2
End Select
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(4) & "' "
End If
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(5) & "' "
End If
If Len(txt1(4)) <> 0 Or Len(txt1(5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/12/2
End If
strSql = "SELECT ST02,CP10,PA08,CP23,CP24 FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6
strSql = strSql + " UNION ALL SELECT ST02,CP10,' ',CP23,CP24 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & strSQL2 & StrSQL6
CheckOC
k = 0
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        DoEvents
        Do While .EOF = False
            BolOk = True
            For i = 0 To 4
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            For i = 0 To 21
                strTemp(i) = ""
            Next i
            strTemp(0) = StrTemp7(0)
            Select Case Val(StrTemp7(1))
            Case 101
                 If StrTemp7(3) = StrTemp7(4) Then
                    strTemp(1) = "1"
                 Else
                    strTemp(2) = "1"
                 End If
            Case 102
                 If StrTemp7(3) = StrTemp7(4) Then
                    strTemp(4) = "1"
                 Else
                    strTemp(5) = "1"
                 End If
            Case 103, 105
                 If StrTemp7(3) = StrTemp7(4) Then
                    strTemp(7) = "1"
                 Else
                    strTemp(8) = "1"
                 End If
            Case 104
                 Select Case Val(StrTemp7(2))
                 Case 1
                      If StrTemp7(3) = StrTemp7(4) Then
                          strTemp(1) = "1"
                      Else
                          strTemp(2) = "1"
                      End If
                 Case 2
                      If StrTemp7(3) = StrTemp7(4) Then
                          strTemp(4) = "1"
                      Else
                          strTemp(5) = "1"
                      End If
                 Case Else
                 End Select
            Case 107
                 'Modified by Morgan 2018/5/14
                 'If StrTemp7(3) = strTemp(4) Then
                 If StrTemp7(3) = StrTemp7(4) Then
                    strTemp(10) = "1"
                 Else
                    strTemp(11) = "1"
                 End If
            Case 501, 502, 503, 504
                 If StrTemp7(3) = StrTemp7(4) Then
                    strTemp(13) = "1"
                 Else
                    strTemp(14) = "1"
                 End If
            Case 801, 802, 803, 804
                 If StrTemp7(3) = StrTemp7(4) Then
                    strTemp(16) = "1"
                 Else
                    strTemp(17) = "1"
                 End If
            Case Else
                 BolOk = False
            End Select
            If BolOk = True Then
                strSql = "INSERT INTO R040403 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & ",'" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If
            .MoveNext
            k = k + 1
            DoEvents
        Loop
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/2
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
CheckOC
strSql = "SELECT R032001,SUM(R032002),SUM(R032003),SUM(R032004),SUM(R032005),SUM(R032006),SUM(R032007),SUM(R032008),SUM(R032009),SUM(R032010),SUM(R032011),SUM(R032012),SUM(R032013),SUM(R032014),SUM(R032015),SUM(R032016),SUM(R032017),SUM(R032018),SUM(R032019),SUM(R032002+R032005+R032008+R032011+R032014+R032017),SUM(R032003+R032006+R032009+R032012+R032015+R032018),SUM(R032022) FROM R040403 WHERE ID='" & strUserNum & "' GROUP BY R032001 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    cnnConnection.Execute "DELETE FROM R040403 WHERE ID='" & strUserNum & "'"
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/2
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 21
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Val(strTemp(1)) = 0 And Val(strTemp(2)) = 0 Then
                strTemp(3) = "0"
            Else
                strTemp(3) = Trim(str(Val(strTemp(1)) / (Val(strTemp(1)) + Val(strTemp(2))) * 100))
            End If
            If Val(strTemp(4)) = 0 And Val(strTemp(5)) = 0 Then
                strTemp(6) = "0"
            Else
                strTemp(6) = Trim(str(Val(strTemp(4)) / (Val(strTemp(4)) + Val(strTemp(5))) * 100))
            End If
            If Val(strTemp(7)) = 0 And Val(strTemp(8)) = 0 Then
                strTemp(9) = "0"
            Else
                strTemp(9) = Trim(str(Val(strTemp(7)) / (Val(strTemp(7)) + Val(strTemp(8))) * 100))
            End If
            If Val(strTemp(10)) = 0 And Val(strTemp(11)) = 0 Then
                strTemp(12) = "0"
            Else
                strTemp(12) = Trim(str(Val(strTemp(10)) / (Val(strTemp(10)) + Val(strTemp(11))) * 100))
            End If
            If Val(strTemp(13)) = 0 And Val(strTemp(14)) = 0 Then
                strTemp(15) = "0"
            Else
                strTemp(15) = Trim(str(Val(strTemp(13)) / (Val(strTemp(13)) + Val(strTemp(14))) * 100))
            End If
            If Val(strTemp(16)) = 0 And Val(strTemp(17)) = 0 Then
                strTemp(18) = "0"
            Else
                strTemp(18) = Trim(str(Val(strTemp(16)) / (Val(strTemp(16)) + Val(strTemp(17))) * 100))
            End If
            If Val(strTemp(19)) = 0 And Val(strTemp(20)) = 0 Then
                strTemp(21) = "0"
            Else
                strTemp(21) = Trim(str(Val(strTemp(19)) / (Val(strTemp(19)) + Val(strTemp(20))) * 100))
            End If
            strSql = "INSERT INTO R040403 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & ",'" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End With
End If
CheckOC
PrintData
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
strSql = "SELECT * FROM R040403 WHERE ID='" & strUserNum & "' "
Select Case Val(txt1(6))
Case 1
      strSql = strSql + " ORDER BY R032004 desc,R032001 "
Case 2
      strSql = strSql + " ORDER BY R032007 desc,R032001 "
Case 3
      strSql = strSql + " ORDER BY R032010 desc,R032001 "
Case 4
      strSql = strSql + " ORDER BY R032013 desc,R032001 "
Case 5
      strSql = strSql + " ORDER BY R032016 desc,R032001 "
Case 6
      strSql = strSql + " ORDER BY R032019 desc,R032001 "
Case 7
      strSql = strSql + " ORDER BY R032021 desc,R032001 "
Case Else
End Select
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 21
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(0) = StrToStr(strTemp(0), 4)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End With
Else
   Exit Sub
End If
PrintEnd
Printer.EndDoc
ShowPrintOk
CheckOC
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6600
Printer.CurrentY = iPrint
'Modify By Cheng 2003/02/04
'若選擇准, 為核駁預估統計表; 若選擇駁, 為核駁預估統計; 若選擇全部, 為准/駁預估統計表
'Printer.Print "准/駁預估統計表"
Printer.Print IIf(Me.txt1(1).Text = "1", "核准預估統計表", IIf(Me.txt1(1).Text = "2", "核駁預估統計表", "准/駁預估統計表"))
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
'Modify By Cheng 2003/02/04
'若選擇准, 為核准日; 若選擇駁, 為核駁日; 若選擇全部, 為准駁日
'Printer.Print "准駁日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
Printer.Print IIf(Me.txt1(1).Text = "1", "核准日", IIf(Me.txt1(1).Text = "2", "核駁日", "准駁日")) & "：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
'Modify By Cheng 2003/02/04
'以全形空白取代半形空白
'Printer.Print "頁    次：" & str(Page)
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "發明"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "新型"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "設計"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "再審"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
'2013/5/7 MODIFY BY SONIA
'Printer.Print "訴願再訴行訴"
Printer.Print "訴願 及 行訴"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
'2013/5/7 MODIFY BY SONIA
'Printer.Print "異議舉發答辯"
Printer.Print "舉發 及 答辯"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "總件數"
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "對"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "錯"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "準確率"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "對"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "錯"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "準確率"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "對"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "錯"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "準確率"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "對"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "錯"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "準確率"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "對"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "錯"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "準確率"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "對"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "錯"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "準確率"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "對"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "錯"
Printer.CurrentX = PLeft(21)
Printer.CurrentY = iPrint
Printer.Print "準確率"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil()
Printer.Font.Size = 10
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 0 To 6
    Printer.CurrentX = PLeft((i * 3) + 1) + 100 - Printer.TextWidth(strTemp((i * 3) + 1))
    Printer.CurrentY = iPrint
    Printer.Print strTemp((i * 3) + 1)
    Printer.CurrentX = PLeft((i * 3) + 2) + 100 - Printer.TextWidth(strTemp((i * 3) + 2))
    Printer.CurrentY = iPrint
    Printer.Print strTemp((i * 3) + 2)
    Printer.CurrentX = PLeft((i * 3) + 3) + 600 - Printer.TextWidth(Format(strTemp((i * 3) + 3), "###.00") + "%")
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp((i * 3) + 3), "###.00") + "%"
Next i
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1800
PLeft(2) = 2250
PLeft(3) = 2700
PLeft(4) = 3600
PLeft(5) = 4050
PLeft(6) = 4500
PLeft(7) = 5400
PLeft(8) = 5850
PLeft(9) = 6300
PLeft(10) = 7200
PLeft(11) = 7650
PLeft(12) = 8100
PLeft(13) = 9000
PLeft(14) = 9450
PLeft(15) = 9900
PLeft(16) = 10800
PLeft(17) = 11250
PLeft(18) = 11700
PLeft(19) = 12600
PLeft(20) = 13050
PLeft(21) = 13500
End Sub

Sub PrintEnd()
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
strSql = "SELECT ' ',SUM(R032002),SUM(R032003),SUM(R032004),SUM(R032005),SUM(R032006),SUM(R032007),SUM(R032008),SUM(R032009),SUM(R032010),SUM(R032011),SUM(R032012),SUM(R032013),SUM(R032014),SUM(R032015),SUM(R032016),SUM(R032017),SUM(R032018),SUM(R032019),SUM(R032020),SUM(R032021),SUM(R032022) FROM R040403 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 21
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Val(strTemp(1)) = 0 And Val(strTemp(2)) = 0 Then
                strTemp3(0) = "0"
            Else
                strTemp(3) = Trim(str(Val(strTemp(1)) / (Val(strTemp(1)) + Val(strTemp(2))) * 100))
            End If
            If Val(strTemp(4)) = 0 And Val(strTemp(5)) = 0 Then
                strTemp(6) = "0"
            Else
                strTemp(6) = Trim(str(Val(strTemp(4)) / (Val(strTemp(4)) + Val(strTemp(5))) * 100))
            End If
            If Val(strTemp(7)) = 0 And Val(strTemp(8)) = 0 Then
                strTemp(9) = "0"
            Else
                strTemp(9) = Trim(str(Val(strTemp(7)) / (Val(strTemp(7)) + Val(strTemp(8))) * 100))
            End If
            If Val(strTemp(10)) = 0 And Val(strTemp(11)) = 0 Then
                strTemp(12) = "0"
            Else
                strTemp(12) = Trim(str(Val(strTemp(10)) / (Val(strTemp(10)) + Val(strTemp(11))) * 100))
            End If
            If Val(strTemp(13)) = 0 And Val(strTemp(14)) = 0 Then
                strTemp(15) = "0"
            Else
                strTemp(15) = Trim(str(Val(strTemp(13)) / (Val(strTemp(13)) + Val(strTemp(14))) * 100))
            End If
            If Val(strTemp(16)) = 0 And Val(strTemp(17)) = 0 Then
                strTemp(18) = "0"
            Else
                strTemp(18) = Trim(str(Val(strTemp(16)) / (Val(strTemp(16)) + Val(strTemp(17))) * 100))
            End If
            If Val(strTemp(19)) = 0 And Val(strTemp(20)) = 0 Then
                strTemp(21) = "0"
            Else
                strTemp(21) = Trim(str(Val(strTemp(19)) / (Val(strTemp(19)) + Val(strTemp(20))) * 100))
            End If
            strTemp(0) = "合計："
            PrintDatil
            .MoveNext
        Loop
    End With
End If
CheckOC
strTemp(0) = ""
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040403 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/12
   Select Case Index
   Case 1 '准駁
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 6 '列印順序
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End Select
   
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
        End If
     Next i
Case 3, 5
   'Modify By Cheng 2002/09/12
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
      End If
   Else
      blnClkSure = False
   End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 2, 3 '准駁日期
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
   End If
Case 1
     Select Case Val(txt1(1))
     Case 1, 2, 3
     Case Else
         s = MsgBox("准駁代碼只能 1 或 2 或 3 !!", , "USER 輸入錯誤")
         'Add By Cheng 2002/09/26
         Cancel = True
     End Select
Case 6
     Select Case Val(txt1(6))
     Case 1, 2, 3, 4, 5, 6, 7
     Case Else
         s = MsgBox("列印順序只能 1~7 !!", , "USER 輸入錯誤")
         'Add By Cheng 2002/09/26
         Cancel = True
     End Select
End Select
If Cancel Then TextInverse txt1(Index)
End Sub
