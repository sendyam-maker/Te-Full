VERSION 5.00
Begin VB.Form frm040402 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文統計表"
   ClientHeight    =   1980
   ClientLeft      =   4320
   ClientTop       =   1992
   ClientWidth     =   3312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3312
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1440
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2190
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1440
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2436
      TabIndex        =   8
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1644
      TabIndex        =   7
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2190
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1155
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1155
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2190
      MaxLength       =   7
      TabIndex        =   2
      Top             =   870
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   1
      Top             =   870
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   570
      Width           =   2190
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   3
      Left            =   165
      TabIndex        =   12
      Top             =   1485
      Width           =   915
   End
   Begin VB.Line Line3 
      X1              =   2055
      X2              =   2220
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Line Line2 
      X1              =   2055
      X2              =   2220
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Line Line1 
      X1              =   2055
      X2              =   2220
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   11
      Top             =   1200
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   10
      Top             =   930
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   9
      Top             =   615
      Width           =   1050
   End
End
Attribute VB_Name = "frm040402"
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
Dim PLeft(0 To 17) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 2) As String
'Add By Cheng 2002/09/12
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
'Add By Cheng 2003/02/04
Dim m_blnSysKindMoreThanOne As Boolean '判斷是否系統類別多於一種

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
               MsgBox "發文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
         End If
         If Len(txt1(2)) = 0 Then
            s = MsgBox("發文日區間不可空白!!", , "USER 輸入錯誤")
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
            '2012/1/2 add by sonia
            If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
               If Me.txt1(5).Text > Me.txt1(6).Text Then
                  MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(5).SetFocus
                  txt1_GotFocus 5
                  Exit Sub
               End If
            End If
            '2012/1/2 end
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
            Process
            Me.Enabled = True
            Screen.MousePointer = vbDefault
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
strSql = "DELETE FROM R040402 WHERE ID='" & strUserNum & "' "
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
   StrSQL6 = StrSQL6 & " AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/2
End If
StrSQL6 = StrSQL6 & " AND CP26 IS NULL "
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
'2012/1/2 add by sonia
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(5) & "-" & txt1(6)
End If
'2012/1/2 end
'Modify By Cheng 2003/02/04
'原抓員工姓名改抓員工代號
'strSQL = "SELECT ST02,CP10,PA08,PA01,CP15 FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " UNION ALL SELECT ST02,CP10,' ',SP01,CP15 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & strSQL2 & StrSQL6
strSql = "SELECT ST01,CP10,PA08,PA01,CP15 FROM CASEPROGRESS,PATENT,STAFF WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND CP09<'C' " & strSQL1 & StrSQL6
strSql = strSql + " UNION ALL SELECT ST01,CP10,' ',SP01,CP15 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP09<'C' " & strSQL2 & StrSQL6
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
            For i = 0 To 18
                strTemp(i) = ""
            Next i
            strTemp(0) = CheckStr(.Fields(0))
            strTemp(19) = CheckStr(.Fields(1))
            strTemp(20) = CheckStr(.Fields(2))
            strTemp(18) = CheckStr(.Fields(3))
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
            strSql = "INSERT INTO R040402 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(17)) & ",'" & ChgSQL(strTemp(18)) & "','" & strUserNum & "') "
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
PrintData
Screen.MousePointer = vbDefault
End Sub

'Modified by Morgan 2011/11/28 -再訴願,異答,異議;+申復,加速審查,分析
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
Printer.Print "發文統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "系統別：" & strTemp3(0)
'Modify By Cheng 2003/02/04
'與發文日對齊
'Printer.CurrentX = 3000
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Format(txt1(3) & " ", "@@@@") & "－" & txt1(4)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
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
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "總支援時數"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintData()
'Modify By Cheng 2003/02/04
'依系統類別及員工代號排序
'strSQL = "SELECT R031001,SUM(R031002),SUM(R031003),SUM(R031004),SUM(R031005),SUM(R031006),SUM(R031007),SUM(R031008),SUM(R031009),SUM(R031010),SUM(R031011),SUM(R031012),SUM(R031013),SUM(R031014),SUM(R031015),SUM(R031016),SUM(R031002+R031003+R031004+R031005+R031006+R031007+R031008+R031009+R031010+R031011+R031012+R031013+R031014+R031015+R031016),SUM(R031017),R031018 FROM R040402 WHERE ID='" & strUserNum & "' GROUP BY R031018,R031001 "
strSql = "SELECT R031001,SUM(R031002),SUM(R031003),SUM(R031004),SUM(R031005),SUM(R031006),SUM(R031007),SUM(R031008),SUM(R031009),SUM(R031010),SUM(R031011),SUM(R031012),SUM(R031013),SUM(R031014),SUM(R031015),SUM(R031016),SUM(R031002+R031003+R031004+R031005+R031006+R031007+R031008+R031009+R031010+R031011+R031012+R031013+R031014+R031015+R031016),SUM(R031017),R031018 FROM R040402 WHERE ID='" & strUserNum & "' GROUP BY R031018,R031001 ORDER BY R031018,R031001 "
CheckOC
'Add By Cheng 2003/02/04
'預設系統類別只有一種
m_blnSysKindMoreThanOne = False
Page = 1
strTemp3(0) = " "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        strTemp3(0) = CheckStr(.Fields(18))
        PrintTitle
        Do While .EOF = False
            For i = 0 To 18
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Cheng 2003/02/04
            '取得員工姓名
'            strTemp(0) = StrToStr(strTemp(0), 4)
            strTemp(0) = GetStaffName(strTemp(0), True)
            If strTemp3(0) <> CheckStr(.Fields(18)) Then
                'Add By Cheng 2003/02/04
                '系統類別超過一種
                m_blnSysKindMoreThanOne = True
                PrintEnd
                strTemp(0) = StrToStr(CheckStr(.Fields(0)), 4)
                For i = 1 To 18
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                strTemp3(0) = CheckStr(.Fields(18))
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End With
End If
PrintEnd
'Modify By Cheng 2003/02/04
'停止對印表機丟資料
'Printer.NewPage
Printer.EndDoc
'若系統類別不只一個才要印"ALL"的資料
If m_blnSysKindMoreThanOne Then
    PrintData1
    PrintEnd1
    Printer.EndDoc
End If
ShowPrintOk
CheckOC
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
Printer.CurrentX = PLeft(17) + 800 - Printer.TextWidth(strTemp(17))
Printer.CurrentY = iPrint
Printer.Print strTemp(17)
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintData1()
'Modify By Cheng 2003/02/04
'依員工代號排序
'strSQL = "SELECT R031001,SUM(R031002),SUM(R031003),SUM(R031004),SUM(R031005),SUM(R031006),SUM(R031007),SUM(R031008),SUM(R031009),SUM(R031010),SUM(R031011),SUM(R031012),SUM(R031013),SUM(R031014),SUM(R031015),SUM(R031016),SUM(R031002+R031003+R031004+R031005+R031006+R031007+R031008+R031009+R031010+R031011+R031012+R031013+R031014+R031015+R031016),SUM(R031017) FROM R040402 WHERE ID='" & strUserNum & "' GROUP BY R031001 "
strSql = "SELECT R031001,SUM(R031002),SUM(R031003),SUM(R031004),SUM(R031005),SUM(R031006),SUM(R031007),SUM(R031008),SUM(R031009),SUM(R031010),SUM(R031011),SUM(R031012),SUM(R031013),SUM(R031014),SUM(R031015),SUM(R031016),SUM(R031002+R031003+R031004+R031005+R031006+R031007+R031008+R031009+R031010+R031011+R031012+R031013+R031014+R031015+R031016),SUM(R031017) FROM R040402 WHERE ID='" & strUserNum & "' GROUP BY R031001 ORDER BY R031001 "
CheckOC
strTemp3(0) = " "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        strTemp3(0) = "ALL"
        'Add By Cheng 2003/02/04
        '頁次加一
        Page = Page + 1
        PrintTitle
        Do While .EOF = False
            For i = 0 To 16
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Cheng 2003/02/04
'            strTemp(0) = StrToStr(strTemp(0), 4)
            strTemp(0) = GetStaffName(strTemp(0), True)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End With
End If
CheckOC
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
If Trim(strTemp3(0)) = "" Then
    strSql = "SELECT SUM(R031002),SUM(R031003),SUM(R031004),SUM(R031005),SUM(R031006),SUM(R031007),SUM(R031008),SUM(R031009),SUM(R031010),SUM(R031011),SUM(R031012),SUM(R031013),SUM(R031014),SUM(R031015),SUM(R031016),SUM(R031002+R031003+R031004+R031005+R031006+R031007+R031008+R031009+R031010+R031011+R031012+R031013+R031014+R031015+R031016),SUM(R031017) FROM R040402 WHERE ID='" & strUserNum & "' AND (R031018='" & strTemp3(0) & "' OR R031018 IS NULL) "
    CheckOC2
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        With adoRecordset1
            .MoveFirst
            For i = 1 To 17
                strTemp(i) = CheckStr(.Fields(i - 1))
            Next i
            strTemp(0) = "小計："
            PrintDatil
        End With
    End If
    CheckOC2
Else
    strSql = "SELECT SUM(R031002),SUM(R031003),SUM(R031004),SUM(R031005),SUM(R031006),SUM(R031007),SUM(R031008),SUM(R031009),SUM(R031010),SUM(R031011),SUM(R031012),SUM(R031013),SUM(R031014),SUM(R031015),SUM(R031016),SUM(R031002+R031003+R031004+R031005+R031006+R031007+R031008+R031009+R031010+R031011+R031012+R031013+R031014+R031015+R031016),SUM(R031017) FROM R040402 WHERE ID='" & strUserNum & "' AND R031018='" & strTemp3(0) & "' "
    CheckOC2
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        With adoRecordset1
            .MoveFirst
            For i = 1 To 17
                strTemp(i) = CheckStr(.Fields(i - 1))
            Next i
            strTemp(0) = "小計："
            PrintDatil
        End With
    End If
    CheckOC2
End If
strTemp(0) = ""
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1700
PLeft(2) = 2450
PLeft(3) = 3200
PLeft(4) = 3950
PLeft(5) = 4700
PLeft(6) = 5450
PLeft(7) = 6400
PLeft(8) = 7150
PLeft(9) = 7900
PLeft(10) = 8850
PLeft(11) = 9600
PLeft(12) = 10350
PLeft(13) = 11100
PLeft(14) = 12050
PLeft(15) = 12950
PLeft(16) = 13800
PLeft(17) = 14550
End Sub

Sub PrintEnd1()
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
strSql = "SELECT SUM(R031002),SUM(R031003),SUM(R031004),SUM(R031005),SUM(R031006),SUM(R031007),SUM(R031008),SUM(R031009),SUM(R031010),SUM(R031011),SUM(R031012),SUM(R031013),SUM(R031014),SUM(R031015),SUM(R031016),SUM(R031002+R031003+R031004+R031005+R031006+R031007+R031008+R031009+R031010+R031011+R031012+R031013+R031014+R031015+R031016),SUM(R031017) FROM R040402 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        For i = 1 To 17
            strTemp(i) = CheckStr(.Fields(i - 1))
        Next i
        strTemp(0) = "合計："
        PrintDatil
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
Set frm040402 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 2, 4
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
Case 1, 2 '發文日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub
