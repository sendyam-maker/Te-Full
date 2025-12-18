VERSION 5.00
Begin VB.Form frm050403 
   BorderStyle     =   1  '單線固定
   Caption         =   "准駁統計表"
   ClientHeight    =   1995
   ClientLeft      =   2925
   ClientTop       =   2805
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form22"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3225
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2268
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1560
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   948
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1560
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   960
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1188
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   960
      MaxLength       =   7
      TabIndex        =   1
      Top             =   852
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   948
      TabIndex        =   0
      Top             =   516
      Width           =   2145
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   1992
      TabIndex        =   7
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1200
      TabIndex        =   6
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2268
      MaxLength       =   7
      TabIndex        =   2
      Top             =   852
      Width           =   800
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   2160
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "申請國家："
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   1560
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1872
      X2              =   2112
      Y1              =   972
      Y2              =   972
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "統計對象："
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "准駁日："
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   885
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
      Height          =   180
      Left            =   90
      TabIndex        =   9
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.承辦人 2.智權人員)"
      Height          =   180
      Left            =   1290
      TabIndex        =   8
      Top             =   1230
      Width           =   1695
   End
End
Attribute VB_Name = "frm050403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSql As String, strTemp(0 To 18) As String, PLeft(0 To 18) As Integer, iPrint As Integer, Page As Integer
Dim i As Integer, j As Integer, s As Integer, strTemp1 As Variant, strTemp2 As Variant
Dim strSQL1 As String
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
   'Add By Cheng 2002/09/16
   blnClkSure = False
   
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Len(txt1(2)) = 0 Then
             s = MsgBox("公告日區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
            'Add By Cheng 2002/03/20
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
            'Add By Cheng 2002/09/16
            If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
               If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                  MsgBox "准駁日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
                         
             If Len(txt1(3)) = 0 Then
                 s = MsgBox("統計對象不可空白!!", , "USER 輸入錯誤")
                 txt1(3).SetFocus
                 Exit Sub
             Else
                 Screen.MousePointer = vbHourglass
                 Me.Enabled = False
                 ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
                 Process
                 Me.Enabled = True
                 Screen.MousePointer = vbDefault
             End If
            '92.10.22 ADD BY SONIA
            If Me.txt1(4).Text <> "" And Me.txt1(4).Text <> "" Then
               If Val(Me.txt1(4).Text) > Val(Me.txt1(5).Text) Then
                  MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
            '92.10.22 END
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from R050403 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & GetAddStr(txt1(0)) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/12/7
End If
If Len(Trim(txt1(1))) <> 0 Then
   strSQL1 = strSQL1 + " AND CP25>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP25<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/7
End If
'92.10.22 ADD BY SONIA
If Len(Trim(txt1(4))) <> 0 Then
   strSQL1 = strSQL1 + " AND PA09>='" & txt1(4) & "' "
End If
If Len(Trim(txt1(5))) <> 0 Then
   strSQL1 = strSQL1 & " AND PA09<='" & txt1(5) & "' "
End If
If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label5 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/12/7
End If
'92.10.22 END
If Val(txt1(3)) = 1 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & "1.承辦人" 'Add By Sindy 2010/12/7
   'Modify By Cheng 2002/04/12
'    strSQL = "SELECT ST02,NVL(CPM03,CPM04),CP24,CP10 FROM CASEPROPERTYMAP,CASEPROGRESS,STAFF WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP14=ST01(+)  AND (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B') " & strSQL1
    '92.10.22 MODIFY BY SONIA 加申請國家條件
    'strSQL = "SELECT ST02,NVL(CPM03,CPM04),CP24,CP10 FROM CASEPROPERTYMAP,CASEPROGRESS,STAFF WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP14=ST01(+)  AND ( CP09<'C' ) " & strSQL1
    strSql = "SELECT ST02,NVL(CPM03,CPM04),CP24,CP10 FROM CASEPROPERTYMAP,CASEPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND CP14=ST01(+)  AND ( CP09<'C' ) " & strSQL1
    '92.10.22 END
Else
   pub_QL05 = pub_QL05 & ";" & Label3 & "2.智權人員" 'Add By Sindy 2010/12/7
   'Modify By Cheng 2002/04/12
'    strSQL = "SELECT ST02,NVL(CPM03,CPM04),CP24,CP10 FROM CASEPROPERTYMAP,CASEPROGRESS,STAFF WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP13=ST01(+)  AND (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B') " & strSQL1
    '92.10.22 MODIFY BY SONIA 加申請國家條件
    'strSQL = "SELECT ST02,NVL(CPM03,CPM04),CP24,CP10 FROM CASEPROPERTYMAP,CASEPROGRESS,STAFF WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP13=ST01(+)  AND ( CP09<'C' ) " & strSQL1
    strSql = "SELECT ST02,NVL(CPM03,CPM04),CP24,CP10 FROM CASEPROPERTYMAP,CASEPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND CP13=ST01(+)  AND ( CP09<'C' ) " & strSQL1
    '92.10.22 END
End If
CheckOC

adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 18
                strTemp(i) = ""
            Next i
            For i = 0 To 3
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            Select Case Val(strTemp(3))
            Case 101
                 If Val(strTemp(2)) = 1 Then
                    strTemp(1) = "1"
                    strTemp(16) = "1"
                    strTemp(2) = ""
                 Else
                     If Val(strTemp(2)) = 2 Then
                        strTemp(2) = "1"
                        strTemp(17) = "1"
                        strTemp(1) = ""
                     End If
                 End If
                 strTemp(3) = ""
                 strTemp(4) = ""
                 
            Case 102
                 If Val(strTemp(2)) = 1 Then
                    strTemp(4) = "1"
                    strTemp(16) = "1"
                 Else
                     If Val(strTemp(2)) = 2 Then
                        strTemp(5) = "1"
                        strTemp(17) = "1"
                     End If
                 End If
                 strTemp(1) = ""
                 strTemp(2) = ""
                 strTemp(3) = ""
            Case 103, 104, 105
                 If Val(strTemp(2)) = 1 Then
                    strTemp(7) = "1"
                    strTemp(16) = "1"
                 Else
                     If Val(strTemp(2)) = 2 Then
                        strTemp(8) = "1"
                        strTemp(17) = "1"
                     End If
                 End If
                 strTemp(1) = ""
                 strTemp(2) = ""
                 strTemp(3) = ""
            Case 501, 502, 503
                 If Val(strTemp(2)) = 1 Then
                    strTemp(10) = "1"
                    strTemp(16) = "1"
                 Else
                     If Val(strTemp(2)) = 2 Then
                        strTemp(11) = "1"
                        strTemp(17) = "1"
                     End If
                 End If
                 strTemp(2) = ""
                 strTemp(1) = ""
                 strTemp(3) = ""
            Case 801, 802, 803
                 If Val(strTemp(2)) = 1 Then
                    strTemp(13) = "1"
                    strTemp(16) = "1"
                 Else
                     If Val(strTemp(2)) = 2 Then
                        strTemp(14) = "1"
                        strTemp(17) = "1"
                     End If
                 End If
                 strTemp(1) = ""
                 strTemp(2) = ""
                 strTemp(3) = ""
            Case Else
                 strTemp(0) = ""
                 strTemp(1) = ""
                 strTemp(2) = ""
                 strTemp(3) = ""
                 strTemp(4) = ""
                 strTemp(5) = ""
            End Select
            If strTemp(0) <> "" Then
                strSql = "INSERT INTO R050403 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & ",'" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If
            .MoveNext

        Loop
    End With
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/7
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Private Sub PrintData()
strSql = "SELECT R017001,SUM(R017002),SUM(R017003),SUM(R017004),SUM(R017005),SUM(R017006),SUM(R017007),SUM(R017008),SUM(R017009),SUM(R017010),SUM(R017011),SUM(R017012),SUM(R017013),SUM(R017014),SUM(R017015),SUM(R017016),SUM(R017017),SUM(R017018),SUM(R017019) FROM R050403 WHERE ID='" & strUserNum & "' GROUP BY R017001 "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        'cnnConnection.Execute "DELETE FROM R050403"
        PrintTitle
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 18
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Val(strTemp(1)) + Val(strTemp(2)) = 0 Then
                strTemp(3) = "0"
            Else
                strTemp(3) = str(Val(strTemp(1)) / (Val(strTemp(1)) + Val(strTemp(2))) * 100)
            End If
            If Val(strTemp(4)) + Val(strTemp(5)) = 0 Then
                strTemp(6) = "0"
            Else
                strTemp(6) = str(Val(strTemp(4)) / (Val(strTemp(4)) + Val(strTemp(5))) * 100)
            End If
            If Val(strTemp(7)) + Val(strTemp(8)) = 0 Then
                strTemp(9) = "0"
            Else
                strTemp(9) = str(Val(strTemp(7)) / (Val(strTemp(7)) + Val(strTemp(8))) * 100)
            End If
            If Val(strTemp(10)) + Val(strTemp(11)) = 0 Then
                strTemp(12) = "0"
            Else
                strTemp(12) = str(Val(strTemp(10)) / (Val(strTemp(10)) + Val(strTemp(11))) * 100)
            End If
            If Val(strTemp(13)) + Val(strTemp(14)) = 0 Then
                strTemp(15) = "0"
            Else
                strTemp(15) = str(Val(strTemp(13)) / (Val(strTemp(13)) + Val(strTemp(14))) * 100)
            End If
            If Val(strTemp(16)) + Val(strTemp(17)) = 0 Then
                strTemp(18) = "0"
            Else
                strTemp(18) = str(Val(strTemp(16)) / (Val(strTemp(16)) + Val(strTemp(17))) * 100)
            End If
            strTemp(3) = Format(strTemp(3), "0.00") & " %"
            strTemp(6) = Format(strTemp(6), "0.00") & " %"
            strTemp(9) = Format(strTemp(9), "0.00") & " %"
            strTemp(12) = Format(strTemp(12), "0.00") & " %"
            strTemp(15) = Format(strTemp(15), "0.00") & " %"
            strTemp(18) = Format(strTemp(18), "0.00") & " %"
            'StrSQL = "INSERT INTO R050403 VALUES('" & chgsql(strTemp(0)) & "'," & Val(StrTemp(1)) & "," & Val(StrTemp(2)) & "," & Val(StrTemp(3)) & "," & Val(StrTemp(4)) & "," & Val(StrTemp(5)) & "," & Val(StrTemp(6)) & "," & Val(StrTemp(7)) & "," & Val(StrTemp(8)) & "," & Val(StrTemp(9)) & "," & Val(StrTemp(10)) & "," & Val(StrTemp(11)) & "," & Val(StrTemp(12)) & "," & Val(StrTemp(13)) & "," & Val(StrTemp(14)) & "," & Val(StrTemp(15)) & "," & Val(StrTemp(16)) & "," & Val(StrTemp(17)) & "," & Val(StrTemp(18)) & ") "
            If iPrint > 10000 Then
                Page = Page + 1
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                Printer.NewPage
                PrintTitle
            End If
            PrintDatil
            'cnnConnection.Execute StrSQL
            .MoveNext
        Loop
    End With
End If
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
strSql = "SELECT '合    計',SUM(R017002),SUM(R017003),SUM(R017004),SUM(R017005),SUM(R017006),SUM(R017007),SUM(R017008),SUM(R017009),SUM(R017010),SUM(R017011),SUM(R017012),SUM(R017013),SUM(R017014),SUM(R017015),SUM(R017016),SUM(R017017),SUM(R017018),SUM(R017019) FROM R050403 WHERE ID='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    For i = 0 To 18
        strTemp(i) = CheckStr(adoRecordset.Fields(i))
    Next i
    If Val(strTemp(1)) + Val(strTemp(2)) = 0 Then
        strTemp(3) = "0"
    Else
        strTemp(3) = str(Val(strTemp(1)) / (Val(strTemp(1)) + Val(strTemp(2))) * 100)
    End If
    If Val(strTemp(4)) + Val(strTemp(5)) = 0 Then
        strTemp(6) = "0"
    Else
        strTemp(6) = str(Val(strTemp(4)) / (Val(strTemp(4)) + Val(strTemp(5))) * 100)
    End If
    If Val(strTemp(7)) + Val(strTemp(8)) = 0 Then
        strTemp(9) = "0"
    Else
        strTemp(9) = str(Val(strTemp(7)) / (Val(strTemp(7)) + Val(strTemp(8))) * 100)
    End If
    If Val(strTemp(10)) + Val(strTemp(11)) = 0 Then
        strTemp(12) = "0"
    Else
        strTemp(12) = str(Val(strTemp(10)) / (Val(strTemp(10)) + Val(strTemp(11))) * 100)
    End If
    If Val(strTemp(13)) + Val(strTemp(14)) = 0 Then
        strTemp(15) = "0"
    Else
        strTemp(15) = str(Val(strTemp(13)) / (Val(strTemp(13)) + Val(strTemp(14))) * 100)
    End If
    If Val(strTemp(16)) + Val(strTemp(17)) = 0 Then
        strTemp(18) = "0"
    Else
        strTemp(18) = str(Val(strTemp(16)) / (Val(strTemp(16)) + Val(strTemp(17))) * 100)
    End If
    strTemp(3) = Format(strTemp(3), "0.00") & " %"
    strTemp(6) = Format(strTemp(6), "0.00") & " %"
    strTemp(9) = Format(strTemp(9), "0.00") & " %"
    strTemp(12) = Format(strTemp(12), "0.00") & " %"
    strTemp(15) = Format(strTemp(15), "0.00") & " %"
    strTemp(18) = Format(strTemp(18), "0.00") & " %"
    PrintDatil
End If
Printer.EndDoc
End Sub

Private Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 1 To 18
   Select Case i
   Case 3, 6, 9, 12, 15, 18
      Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Case Else
      Printer.CurrentX = PLeft(i) + 150 - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   End Select
Next i
iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 3000
PLeft(2) = 3600
PLeft(3) = 4200
PLeft(4) = 5000
PLeft(5) = 5600
PLeft(6) = 6200
PLeft(7) = 7000
PLeft(8) = 7600
PLeft(9) = 8200
PLeft(10) = 9000
PLeft(11) = 9600
PLeft(12) = 10200
PLeft(13) = 11000
PLeft(14) = 11600
PLeft(15) = 12200
PLeft(16) = 13000
PLeft(17) = 13600
PLeft(18) = 14200
End Sub

Private Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "准 駁 統 計 表"
iPrint = iPrint + 500
Printer.Font.Size = 8
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "准駁日期：" & ChangeTStringToTDateString(txt1(1)) & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
'92.10.22 ADD BY SONIA
If txt1(4) <> "" Or txt1(5) <> "" Then
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   Printer.Print "申請國家：" & txt1(4) & "－" & txt1(5)
End If
'92.10.22 END
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
'Printer.Font.Underline = True
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "發         明"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "新         型"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "新    式    樣"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "訴願  再訴   行訴"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "異議  舉發   答辯"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "總    件    數"
'Printer.Font.Underline = False
Printer.Line (PLeft(1), iPrint + 200)-(PLeft(3) + 500, iPrint + 200)
Printer.Line (PLeft(4), iPrint + 200)-(PLeft(6) + 500, iPrint + 200)
Printer.Line (PLeft(7), iPrint + 200)-(PLeft(9) + 500, iPrint + 200)
Printer.Line (PLeft(10), iPrint + 200)-(PLeft(12) + 500, iPrint + 200)
Printer.Line (PLeft(13), iPrint + 200)-(PLeft(15) + 500, iPrint + 200)
Printer.Line (PLeft(16), iPrint + 200)-(PLeft(18) + 500, iPrint + 200)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
If Val(txt1(3)) = 1 Then
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
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300

End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050403 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/16
   Select Case Index
   Case 3
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
Case 2
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If

Case 3
     Select Case Trim(txt1(3))
     Case "1", "2"
     Case Else
          s = MsgBox("統計對象只能 1 或 2 !!", , "USER 輸入錯誤")
          txt1(3).SetFocus
          txt1(3).SelStart = 0
          txt1(3).SelLength = Len(txt1(3))
          Exit Sub
     End Select
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '准駁日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub
