VERSION 5.00
Begin VB.Form frm040401 
   BorderStyle     =   1  '單線固定
   Caption         =   "收文統計表"
   ClientHeight    =   2280
   ClientLeft      =   1236
   ClientTop       =   2832
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3900
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3012
      TabIndex        =   9
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2220
      TabIndex        =   8
      Top             =   20
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   804
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1830
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1530
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1000
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1530
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2430
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1215
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1000
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1215
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2430
      MaxLength       =   7
      TabIndex        =   2
      Top             =   915
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1000
      MaxLength       =   7
      TabIndex        =   1
      Top             =   900
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1000
      TabIndex        =   0
      Top             =   600
      Width           =   2760
   End
   Begin VB.Line Line3 
      X1              =   2205
      X2              =   2405
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line2 
      X1              =   2205
      X2              =   2405
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   2205
      X2              =   2405
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "(1.承辦人 2.智權人員)"
      Height          =   180
      Index           =   5
      Left            =   1170
      TabIndex        =   15
      Top             =   1875
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   14
      Top             =   1875
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Top             =   1575
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   2
      Left            =   90
      TabIndex        =   12
      Top             =   1260
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   11
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   645
      Width           =   915
   End
End
Attribute VB_Name = "frm040401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL6 As String, i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 20) As String, strTemp1 As Variant, strTemp2 As Variant, k As Integer
Dim PLeft(0 To 16) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 2) As String
'Add By Cheng 2002/09/12
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
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
               MsgBox "收文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
         End If

         If Len(txt1(2)) = 0 Then
            s = MsgBox("收文日區間不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1_GotFocus (1)
            Exit Sub
         Else
            'Add By Cheng 2002/09/12
            If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
               If Me.txt1(3).Text > Me.txt1(4).Text Then
                  MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
            If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
               If Me.txt1(5).Text > Me.txt1(6).Text Then
                  MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(5).SetFocus
                  txt1_GotFocus 5
                  Exit Sub
               End If
            End If
         
            If Len(txt1(7)) = 0 Then
                s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                txt1(7).SetFocus
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
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
strSql = "DELETE FROM R040401 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute strSql
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
   StrSQL6 = StrSQL6 & " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/2
End If
StrSQL6 = StrSQL6 & " AND CP26 IS NULL AND CP57 IS NULL "
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/2
End If
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(6) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/2
End If
'Modify By Cheng 2003/02/04
'原抓姓名改抓代號
''**************** 將業務區改成抓案件進度檔   91.08.15  nick
'strSQL = "SELECT S1.ST02,CP10,PA08,NVL(A0902,A0903),S2.ST02 FROM CASEPROGRESS,PATENT,ACC090,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP12=A0901(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " UNION ALL SELECT S1.ST02,CP10,' ',NVL(A0902,A0903),S2.ST02 FROM CASEPROGRESS,SERVICEPRACTICE,ACC090,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP12=A0901(+) " & strSQL2 & StrSQL6
strSql = "SELECT S1.ST01,CP10,PA08,NVL(A0902,A0903),S2.ST01 FROM CASEPROGRESS,PATENT,ACC090,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP12=A0901(+) AND CP09<'C' " & strSQL1 & StrSQL6
strSql = strSql + " UNION ALL SELECT S1.ST01,CP10,' ',NVL(A0902,A0903),S2.ST01 FROM CASEPROGRESS,SERVICEPRACTICE,ACC090,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP12=A0901(+) AND CP09<'C' " & strSQL2 & StrSQL6
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/2
        .MoveFirst
        k = 0
        DoEvents
        Do While .EOF = False
            For i = 0 To 17
                strTemp(i) = ""
            Next i
            strTemp(0) = CheckStr(.Fields(0))
            strTemp(19) = CheckStr(.Fields(1))
            strTemp(20) = CheckStr(.Fields(2))
            strTemp(16) = CheckStr(.Fields(3))
            strTemp(17) = CheckStr(.Fields(4))
            
            'Modified by Morgan 2011/11/28 -502(再訴願),802(異答),801(異議);+205(申復),422(加速審查),941(分析)
            Select Case Val(strTemp(19))
            Case 101
                 strTemp(1) = "1"
            Case 104
                 Select Case Val(strTemp(20))
                 Case 1
                      strTemp(1) = "1"
                 Case 2
                      strTemp(2) = "1"
                 End Select
            Case 102
                 strTemp(2) = "1"
            Case 103, 105
                 strTemp(3) = "1"
            Case 107
                 strTemp(4) = "1"
            Case 205
                 strTemp(5) = "1"
            'Modified by Morgan 2024/11/18 +447再審查加速審查
            Case 422, 447
                 strTemp(6) = "1"
            Case 941
                 strTemp(7) = "1"
            Case 501
                 strTemp(8) = "1"
            Case 503, 504
                 strTemp(9) = "1"
            Case 804
                 strTemp(10) = "1"
            Case 803
                 strTemp(11) = "1"
            Case 203, 204
                 strTemp(12) = "1"
            Case 903
                 strTemp(13) = "1"
            Case 906
                 strTemp(14) = "1"
                 
            'Add by Morgan 2003/12/03
            Case 113, 307
               Select Case Val(strTemp(20))
                 Case 1
                      strTemp(1) = "1"
                 Case 2
                      strTemp(2) = "1"
                 Case 3
                      strTemp(3) = "1"
                 End Select
            'End 2003/12/03
                 
            Case Else
                 strTemp(15) = "1"
            End Select
            strSql = "INSERT INTO R040401 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            k = k + 1
            DoEvents
            .MoveNext
        Loop
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/2
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
CheckOC
If Val(txt1(7)) = 1 Then '依承辦人列印
    PrintData1
Else '依智權人員列印
    PrintData2
End If
Screen.MousePointer = vbDefault
End Sub

Sub PrintData1()
'Modify By Cheng 2003/02/04
'以員工代號排序
'strSQL = "SELECT R030001,SUM(R030002),SUM(R030003),SUM(R030004),SUM(R030005),SUM(R030006),SUM(R030007),SUM(R030008),SUM(R030009),SUM(R030010),SUM(R030011),SUM(R030012),SUM(R030013),SUM(R030014),SUM(R030015),SUM(R030016),SUM(R030002+R030003+R030004+R030005+R030006+R030007+R030008+R030009+R030010+R030011+R030012+R030013+R030014+R030015+R030016) FROM R040401 WHERE ID='" & strUserNum & "' GROUP BY R030001 "
strSql = "SELECT R030001,SUM(R030002),SUM(R030003),SUM(R030004),SUM(R030005),SUM(R030006),SUM(R030007),SUM(R030008),SUM(R030009),SUM(R030010),SUM(R030011),SUM(R030012),SUM(R030013),SUM(R030014),SUM(R030015),SUM(R030016),SUM(R030002+R030003+R030004+R030005+R030006+R030007+R030008+R030009+R030010+R030011+R030012+R030013+R030014+R030015+R030016) FROM R040401 WHERE ID='" & strUserNum & "' GROUP BY R030001 ORDER BY R030001 "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 16
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            '取得員工姓名
'            strTemp(0) = StrToStr(strTemp(0), 4)
            strTemp(0) = GetStaffName(strTemp(0), True)
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle1
            End If
            PrintDatil
            .MoveNext
        Loop
    End With
Else
End If
PrintEnd1
Printer.EndDoc
ShowPrintOk
CheckOC
End Sub

Sub PrintEnd1()
Printer.Font.Name = "細明體"
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
End If
strSql = "SELECT SUM(R030002),SUM(R030003),SUM(R030004),SUM(R030005),SUM(R030006),SUM(R030007),SUM(R030008),SUM(R030009),SUM(R030010),SUM(R030011),SUM(R030012),SUM(R030013),SUM(R030014),SUM(R030015),SUM(R030016),SUM(R030002+R030003+R030004+R030005+R030006+R030007+R030008+R030009+R030010+R030011+R030012+R030013+R030014+R030015+R030016) FROM R040401 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        For i = 1 To 16
            strTemp(i) = CheckStr(.Fields(i - 1))
        Next i
        strTemp(0) = "合計："
        PrintDatil
    End With
End If
CheckOC
End Sub

Sub PrintEnd2()
Printer.Font.Name = "細明體"
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
End If
If Trim(strTemp3(0)) = "" Then
    strSql = "SELECT SUM(R030002),SUM(R030003),SUM(R030004),SUM(R030005),SUM(R030006),SUM(R030007),SUM(R030008),SUM(R030009),SUM(R030010),SUM(R030011),SUM(R030012),SUM(R030013),SUM(R030014),SUM(R030015),SUM(R030016),SUM(R030002+R030003+R030004+R030005+R030006+R030007+R030008+R030009+R030010+R030011+R030012+R030013+R030014+R030015+R030016) FROM R040401 WHERE ID='" & strUserNum & "' AND (R030017='' OR R030017 IS NULL) "
    CheckOC2
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        With adoRecordset1
            .MoveFirst
            Printer.Font.Name = "細明體"
            Printer.Font.Size = 10
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print "小計："
            For i = 1 To 16
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(CheckStr(.Fields(i - 1)))
                Printer.CurrentY = iPrint
                Printer.Print CheckStr(.Fields(i - 1))
            Next i
            iPrint = iPrint + 300
            Printer.Font.Size = 10
        End With
    End If
    CheckOC2
Else
    strSql = "SELECT SUM(R030002),SUM(R030003),SUM(R030004),SUM(R030005),SUM(R030006),SUM(R030007),SUM(R030008),SUM(R030009),SUM(R030010),SUM(R030011),SUM(R030012),SUM(R030013),SUM(R030014),SUM(R030015),SUM(R030016),SUM(R030002+R030003+R030004+R030005+R030006+R030007+R030008+R030009+R030010+R030011+R030012+R030013+R030014+R030015+R030016) FROM R040401 WHERE ID='" & strUserNum & "' AND R030017='" & strTemp3(0) & "' "
    CheckOC2
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        With adoRecordset1
            .MoveFirst
            Printer.Font.Name = "細明體"
            Printer.Font.Size = 10
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print "小計："
            For i = 1 To 16
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(CheckStr(.Fields(i - 1)))
                Printer.CurrentY = iPrint
                Printer.Print CheckStr(.Fields(i - 1))
            Next i
            iPrint = iPrint + 300
            Printer.Font.Size = 10
        End With
    End If
    CheckOC2
End If
End Sub

Sub PrintEnd3()
Printer.Font.Name = "細明體"
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
End If

strSql = "SELECT SUM(R030002),SUM(R030003),SUM(R030004),SUM(R030005),SUM(R030006),SUM(R030007),SUM(R030008),SUM(R030009),SUM(R030010),SUM(R030011),SUM(R030012),SUM(R030013),SUM(R030014),SUM(R030015),SUM(R030016),SUM(R030002+R030003+R030004+R030005+R030006+R030007+R030008+R030009+R030010+R030011+R030012+R030013+R030014+R030015+R030016) FROM R040401 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        For i = 1 To 16
            strTemp(i) = CheckStr(.Fields(i - 1))
        Next i
        strTemp(0) = "合計："
        PrintDatil
    End With
End If
CheckOC
End Sub


Sub PrintData2()
'Modify By Cheng 2003/02/04
'以業務區及員工號排序
'strSQL = "SELECT R030018,SUM(R030002),SUM(R030003),SUM(R030004),SUM(R030005),SUM(R030006),SUM(R030007),SUM(R030008),SUM(R030009),SUM(R030010),SUM(R030011),SUM(R030012),SUM(R030013),SUM(R030014),SUM(R030015),SUM(R030016),SUM(R030002+R030003+R030004+R030005+R030006+R030007+R030008+R030009+R030010+R030011+R030012+R030013+R030014+R030015+R030016),R030017 FROM R040401 WHERE ID='" & strUserNum & "' GROUP BY R030017,R030018 "
strSql = "SELECT R030018,SUM(R030002),SUM(R030003),SUM(R030004),SUM(R030005),SUM(R030006),SUM(R030007),SUM(R030008),SUM(R030009),SUM(R030010),SUM(R030011),SUM(R030012),SUM(R030013),SUM(R030014),SUM(R030015),SUM(R030016),SUM(R030002+R030003+R030004+R030005+R030006+R030007+R030008+R030009+R030010+R030011+R030012+R030013+R030014+R030015+R030016),R030017 FROM R040401 WHERE ID='" & strUserNum & "' GROUP BY R030017,R030018 ORDER BY R030017,R030018 "
CheckOC
Page = 1
strTemp3(0) = " "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        strTemp(17) = CheckStr(.Fields(17))
        strTemp3(0) = strTemp(17)
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 17
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp(17) <> strTemp3(0) Then
                PrintEnd2
                strTemp3(0) = CheckStr(.Fields(17))
                strTemp(17) = CheckStr(.Fields(17))
                'Modify By Cheng 2003/02/04
                '取得員工姓名
'                strTemp(0) = CheckStr(.Fields(0))
                strTemp(0) = GetStaffName("" & .Fields(0), True)
                iPrint = iPrint + 300
                If iPrint >= 10000 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitle
                End If
                PrintTitle2
            End If
            'Modify By Cheng 2003/02/04
            '取得員工姓名
'            strTemp(0) = CheckStr(.Fields(0))
            strTemp(0) = GetStaffName("" & .Fields(0), True)
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle2
            End If
            PrintDatil
            .MoveNext
        Loop
    End With
End If
PrintEnd2
PrintEnd3
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
Printer.Print "收文統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Format(txt1(5) & " ", "@@@@") & "－" & txt1(6)
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
End Sub

Sub PrintTitle1()
Printer.Font.Name = "細明體"
Printer.Font.Size = 10
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "發明"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "新型"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "設計"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "再審"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "申復"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "加速審查"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "分析"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "訴願"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "行政訴訟"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "舉答"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "舉發"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "修正"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "專利調查"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "鑑定報告"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "總計"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintTitle2()
Printer.Font.Name = "細明體"
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print strTemp3(0)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.Font.Size = 10
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "發明"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "新型"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "設計"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "再審"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "訴願"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "再訴願"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "行政訴訟"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "異答"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "舉答"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "異議"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "舉發"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "修正"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "專利調查"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "鑑定報告"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "總計"
iPrint = iPrint + 300
Printer.Font.Size = 12
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.Font.Size = 10
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1700
PLeft(2) = 2500
PLeft(3) = 3300
PLeft(4) = 4100
PLeft(5) = 5100
PLeft(6) = 5900
PLeft(7) = 6900
PLeft(8) = 7700
PLeft(9) = 8700
PLeft(10) = 9500
PLeft(11) = 10300
PLeft(12) = 11100
PLeft(13) = 11900
PLeft(14) = 12900
PLeft(15) = 13900
PLeft(16) = 14700
End Sub

Sub PrintDatil()
Printer.Font.Size = 10
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 1 To 16
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040401 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/12
   Select Case Index
   Case 7 '列印別
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
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
Case 2, 4, 6
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
Case 1, 2 '收文日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
   End If
Case 7
     Select Case Val(txt1(7))
     Case 1, 2
     Case Else
          s = MsgBox("列印別只能 1 或 2 !!", , "USER 輸入錯誤")
          Cancel = True
     End Select
End Select
If Cancel Then TextInverse txt1(Index)
End Sub
