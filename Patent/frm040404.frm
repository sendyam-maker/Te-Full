VERSION 5.00
Begin VB.Form frm040404 
   BorderStyle     =   1  '單線固定
   Caption         =   "准駁統計總表"
   ClientHeight    =   2730
   ClientLeft      =   480
   ClientTop       =   2880
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   4104
      TabIndex        =   10
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3312
      TabIndex        =   9
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2265
      Width           =   225
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1875
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2310
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1530
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1035
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1530
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2310
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1215
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1035
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1215
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   2
      Top             =   900
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1035
      MaxLength       =   7
      TabIndex        =   1
      Top             =   900
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1035
      TabIndex        =   0
      Top             =   585
      Width           =   3720
   End
   Begin VB.Line Line3 
      X1              =   2175
      X2              =   2340
      Y1              =   1635
      Y2              =   1665
   End
   Begin VB.Line Line2 
      X1              =   2235
      X2              =   2355
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line1 
      X1              =   2190
      X2              =   2310
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "(1.承辦人 2.智權人員)"
      Height          =   180
      Index           =   7
      Left            =   1365
      TabIndex        =   18
      Top             =   2340
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "(1.發明 2.新型 3.設計  4.再審 5.訴願再訴行訴     6.異議舉發答辯 7.總件數)"
      Height          =   360
      Index           =   6
      Left            =   1350
      TabIndex        =   17
      Top             =   1920
      Width           =   3630
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   5
      Left            =   135
      TabIndex        =   16
      Top             =   2325
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      Height          =   180
      Index           =   4
      Left            =   135
      TabIndex        =   15
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   3
      Left            =   135
      TabIndex        =   14
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   13
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "准駁日期："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   12
      Top             =   945
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   630
      Width           =   975
   End
End
Attribute VB_Name = "frm040404"
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

Private Sub cmdOK_Click(index As Integer)
Select Case index
Case 0
      'Add By Cheng 2002/09/12
      blnClkSure = False
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
         'Add By Cheng 2002/03/19
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            txt1_GotFocus 2
            Exit Sub
         End If
         'Add By Cheng 2002/09/12
         If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
            If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
               MsgBox "准駁日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
         End If
        
        If Len(txt1(2)) = 0 Then
            s = MsgBox("准駁日期區間不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1_GotFocus (1)
            Exit Sub
        Else
            'Add By Cheng 2002/09/12
            If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
               If Me.txt1(3).Text > Me.txt1(4).Text Then
                  MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
            If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
               If Me.txt1(5).Text > Me.txt1(6).Text Then
                  MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(5).SetFocus
                  txt1_GotFocus 5
                  Exit Sub
               End If
            End If
            
            If Len(txt1(7)) = 0 Then
                s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
                txt1(7).SetFocus
                Exit Sub
            Else
                If Len(txt1(8)) = 0 Then
                    s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                    txt1(8).SetFocus
                    Exit Sub
                Else
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
cnnConnection.Execute "DELETE FROM R040404 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/2
End If
StrSQL6 = ""
If Len(Trim(txt1(1))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP25>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP25<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/2
End If
If Len(txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/2
End If
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/2
End If
'Add By Cheng 2003/05/13
'收文日不為111111, 計件案件 ( 但案件性質在101~105之間者不論是否計件都要算入 )
StrSQL6 = StrSQL6 & " And CP05<>19221111 And (CP26 Is Null Or (To_Number(CP10)>=101 and To_Number(CP10)<=105 )) "
strSql = "SELECT S1.ST02,CP10,PA08,CP24,S2.ST02 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
strSql = strSql + " UNION ALL SELECT S1.ST02,CP10,' ',CP24,S2.ST02 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
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
            If Val(txt1(8)) = "1" Then
                strTemp(0) = StrTemp7(0)
            Else
                strTemp(0) = StrTemp7(4)
            End If
            Select Case Val(StrTemp7(1))
            Case 101 '發明申請
                 Select Case Val(StrTemp7(3))
                 Case 1
                      strTemp(1) = "1"
                 Case 2
                      strTemp(2) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 102 '新型申請
                 Select Case Val(StrTemp7(3))
                 Case 1
                      strTemp(4) = "1"
                 Case 2
                      strTemp(5) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 103, 105 '設計申請,聯合申請
                 Select Case Val(StrTemp7(3))
                 Case 1
                      strTemp(7) = "1"
                 Case 2
                      strTemp(8) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 104 '追加申請
                 Select Case Val(StrTemp7(2))
                 Case 1
                      Select Case Val(StrTemp7(3))
                      Case 1
                          strTemp(1) = "1"
                      Case 2
                          strTemp(2) = "1"
                      Case Else
                           BolOk = False
                      End Select
                 Case 2
                      Select Case Val(StrTemp7(3))
                      Case 1
                          strTemp(4) = "1"
                      Case 2
                          strTemp(5) = "1"
                      Case Else
                           BolOk = False
                      End Select
                 Case Else
                      BolOk = False
                 End Select
            Case 107 '再審申請
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(10) = "1"
                 Case 2
                    strTemp(11) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 501, 502, 503, 504
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(13) = "1"
                 Case 2
                    strTemp(14) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 801, 802, 803, 804
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(16) = "1"
                 Case 2
                    strTemp(17) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case Else
                 BolOk = False
            End Select
            If BolOk = True Then
                strSql = "INSERT INTO R040404 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & ",'" & strUserNum & "') "
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
strSql = "SELECT R033001,SUM(R033002),SUM(R033003),SUM(R033004),SUM(R033005),SUM(R033006),SUM(R033007),SUM(R033008),SUM(R033009),SUM(R033010),SUM(R033011),SUM(R033012),SUM(R033013),SUM(R033014),SUM(R033015),SUM(R033016),SUM(R033017),SUM(R033018),SUM(R033019),SUM(R033002+R033005+R033008+R033011+R033014+R033017),SUM(R033003+R033006+R033009+R033012+R033015+R033018),SUM(R033022) FROM R040404 WHERE ID='" & strUserNum & "' GROUP BY R033001 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    cnnConnection.Execute "DELETE FROM R040404 WHERE ID='" & strUserNum & "'"
    With adoRecordset
        .MoveFirst
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/2
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
            strSql = "INSERT INTO R040404 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & ",'" & strUserNum & "') "
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
strSql = "SELECT * FROM R040404 WHERE ID='" & strUserNum & "' "
'列印順序
Select Case Val(txt1(7))
Case 1 '發明(核准率)
'      strSQL = strSQL + " ORDER BY R033004 desc,R033001 "
      strSql = strSql + " ORDER BY R033004 desc, R033002 Desc, R033001 "
Case 2 '新型(核准率)
'      strSQL = strSQL + " ORDER BY R033007 desc,R033001 "
      strSql = strSql + " ORDER BY R033007 desc, R033005 Desc, R033001 "
Case 3 '設計(核准率)
'      strSQL = strSQL + " ORDER BY R033010 desc,R033001 "
      strSql = strSql + " ORDER BY R033010 desc, R033008 Desc,R033001 "
Case 4 '再審(核准率)
'      strSQL = strSQL + " ORDER BY R033013 desc,R033001 "
      strSql = strSql + " ORDER BY R033013 desc, R033011 Desc, R033001 "
Case 5 '訴願再訴行訴(核准率)
'      strSQL = strSQL + " ORDER BY R033016 desc,R033001 "
      strSql = strSql + " ORDER BY R033016 desc, R033014 Desc ,R033001 "
Case 6 '異議舉發答辯(核准率)
'      strSQL = strSQL + " ORDER BY R033019 desc,R033001 "
      strSql = strSql + " ORDER BY R033019 desc, R033017 Desc, R033001 "
Case 7 '總件數(核准率)
        'Modify By Cheng 2003/05/13
'      strSQL = strSQL + " ORDER BY R033021 desc,R033001 "
      strSql = strSql + " ORDER BY R033022 desc, R033020 Desc, R033001 "
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
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
Printer.Print "准/駁統計總表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
Printer.Print "准駁日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
'Add By Cheng 2003/10/20
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Me.txt1(3).Text & "－" & Me.txt1(4).Text
'End
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
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
If Val(txt1(8)) = 1 Then
    Printer.Print "承辦人"
Else
    Printer.Print "智權人員"
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(21)
Printer.CurrentY = iPrint
Printer.Print "核准率"
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
strSql = "SELECT ' ',SUM(R033002),SUM(R033003),SUM(R033004),SUM(R033005),SUM(R033006),SUM(R033007),SUM(R033008),SUM(R033009),SUM(R033010),SUM(R033011),SUM(R033012),SUM(R033013),SUM(R033014),SUM(R033015),SUM(R033016),SUM(R033017),SUM(R033018),SUM(R033019),SUM(R033020),SUM(R033021),SUM(R033022) FROM R040404 WHERE ID='" & strUserNum & "' "
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
Set frm040404 = Nothing
End Sub

Private Sub txt1_GotFocus(index As Integer)
txt1(index).SelStart = 0
txt1(index).SelLength = Len(txt1(index))
End Sub

Private Sub txt1_KeyPress(index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/12
   Select Case index
   Case 7 '列印順序
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 8 '列印別
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End Select
   
End Sub

Private Sub txt1_LostFocus(index As Integer)
Select Case index
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
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 2, 4, 6
   'Modify By Cheng 2002/09/11
   If blnClkSure = False Then
      If RunNick(txt1(index - 1), txt1(index)) Then
         txt1(index - 1).SetFocus
      End If
   Else
      blnClkSure = False
   End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(index As Integer, Cancel As Boolean)
Select Case index
Case 1, 2 '准駁日期起, 迄
   If PUB_CheckKeyInDate(Me.txt1(index)) = -1 Then
      Cancel = True
   End If
Case 7
     Select Case Val(txt1(7))
     Case 1, 2, 3, 4, 5, 6, 7
     Case Else
          s = MsgBox("列印順序只能 1 ~ 7 !!", , "USER 輸入錯誤")
          Cancel = True
     End Select
Case 8
     Select Case Val(txt1(8))
     Case 1, 2
     Case Else
          s = MsgBox("列印別只能 1 或 2 !!", , "USER 輸入錯誤")
          Cancel = True
     End Select
End Select
If Cancel Then TextInverse txt1(index)
End Sub
