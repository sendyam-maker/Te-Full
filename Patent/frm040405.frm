VERSION 5.00
Begin VB.Form frm040405 
   BorderStyle     =   1  '單線固定
   Caption         =   "准駁統計明細表"
   ClientHeight    =   2805
   ClientLeft      =   1785
   ClientTop       =   1890
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5280
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   615
      Width           =   3915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   1
      Top             =   930
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2370
      MaxLength       =   7
      TabIndex        =   2
      Top             =   930
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1245
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2370
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1245
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1545
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2370
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1545
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1875
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2220
      Width           =   225
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3156
      TabIndex        =   9
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3948
      TabIndex        =   10
      Top             =   60
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "准駁日期："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   17
      Top             =   975
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   16
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   15
      Top             =   1605
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   14
      Top             =   1935
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "(1.發明 2.新型 3.設計 4.再審 5.訴願再訴行訴      6.異議舉發答辯 7.總件數)"
      Height          =   360
      Index           =   6
      Left            =   1410
      TabIndex        =   12
      Top             =   1875
      Width           =   3630
   End
   Begin VB.Label Label1 
      Caption         =   "(1.承辦員 2.智權人員)"
      Height          =   180
      Index           =   7
      Left            =   1425
      TabIndex        =   11
      Top             =   2310
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   2250
      X2              =   2370
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   2400
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line3 
      X1              =   2220
      X2              =   2385
      Y1              =   1650
      Y2              =   1680
   End
End
Attribute VB_Name = "frm040405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL6 As String, i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 39) As String, strTemp1 As Variant, strTemp2 As Variant, k As Integer, StrTemp7(0 To 5) As String
Dim PLeft(0 To 39) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 2) As String, BolOk As Boolean
'Add By Cheng 2002/09/12
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
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
Case 1 '結束
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R040405 WHERE ID='" & strUserNum & "' "
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
'strSQL = "SELECT S1.ST02,CP10,PA08,CP24,S2.ST02,CP43 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " UNION ALL SELECT S1.ST02,CP10,' ',CP24,S2.ST02,CP43 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
strSql = "SELECT S1.ST02,CP10,PA08,CP24,S2.ST02,CP43, CP09 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
strSql = strSql + " UNION ALL SELECT S1.ST02,CP10,' ',CP24,S2.ST02,CP43, CP09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
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
            For i = 0 To 5
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            For i = 0 To 39
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
            Case 103, 105 '設計申請, 聯合申請
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
                'Modify By Cheng 2003/05/13
                '改抓基本檔專利種類
'                 strSQL = "SELECT CP10,PA08,CP24 FROM CASEPROGRESS,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP09='" & StrTemp7(5) & "' "
'                 strSQL = strSQL + " UNION ALL SELECT CP10,' ',CP24 FROM CASEPROGRESS,SERVICEPRACTICE WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP09='" & StrTemp7(5) & "' "
                 strSql = "SELECT CP10,PA08,CP24 FROM CASEPROGRESS,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) And CP09='" & .Fields("CP09").Value & "' "
                 CheckOC2
                 adoRecordset1.CursorLocation = adUseClient
                 adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                 If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                        'Modify By Cheng 2003/05/13
'                        Select Case Val(CheckStr(adoRecordset1.Fields(0)))
'                        Case 101 '發明申請
'                             Select Case Val(CheckStr(adoRecordset1.Fields(2)))
'                             Case 1
'                                  strTemp(10) = "1"
'                             Case 2
'                                  strTemp(11) = "1"
'                             Case Else
'                                  BolOk = False
'                             End Select
'                        Case 102 '新型申請
'                             Select Case Val(CheckStr(adoRecordset1.Fields(2)))
'                             Case 1
'                                  strTemp(13) = "1"
'                             Case 2
'                                  strTemp(14) = "1"
'                             Case Else
'                                  BolOk = False
'                             End Select
'                        Case 103, 105 '設計申請, 聯合申請
'                             Select Case Val(CheckStr(adoRecordset1.Fields(2)))
'                             Case 1
'                                  strTemp(16) = "1"
'                             Case 2
'                                  strTemp(17) = "1"
'                             Case Else
'                                  BolOk = False
'                             End Select
'                        Case 104 '追加申請
'                             Select Case Val(CheckStr(adoRecordset1.Fields(1)))
'                             Case 1
'                                  Select Case Val(CheckStr(adoRecordset1.Fields(2)))
'                                  Case 1
'                                       strTemp(10) = "1"
'                                  Case 2
'                                       strTemp(11) = "1"
'                                  Case Else
'                                       BolOk = False
'                                  End Select
'                             Case 2
'                                  Select Case Val(CheckStr(adoRecordset1.Fields(2)))
'                                  Case 1
'                                       strTemp(13) = "1"
'                                  Case 2
'                                       strTemp(14) = "1"
'                                  Case Else
'                                       BolOk = False
'                                  End Select
'                             Case Else
'                                  BolOk = False
'                             End Select
'                        Case Else
'                        End Select
                        '判斷專利種類
                        Select Case Val(CheckStr(adoRecordset1.Fields(1)))
                        Case 1 '發明申請
                             Select Case Val(CheckStr(adoRecordset1.Fields(2)))
                             Case 1
                                  strTemp(10) = "1"
                             Case 2
                                  strTemp(11) = "1"
                             Case Else
                                  BolOk = False
                             End Select
                        Case 2 '新型申請
                             Select Case Val(CheckStr(adoRecordset1.Fields(2)))
                             Case 1
                                  strTemp(13) = "1"
                             Case 2
                                  strTemp(14) = "1"
                             Case Else
                                  BolOk = False
                             End Select
                        Case 3 '設計申請
                             Select Case Val(CheckStr(adoRecordset1.Fields(2)))
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
                 Else
                     BolOk = False
                 End If
                 CheckOC2
            Case 501
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(19) = "1"
                 Case 2
                    strTemp(20) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 502
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(22) = "1"
                 Case 2
                    strTemp(23) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 503, 504
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(25) = "1"
                 Case 2
                    strTemp(26) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 801
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(28) = "1"
                 Case 2
                    strTemp(29) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 803
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(31) = "1"
                 Case 2
                    strTemp(32) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case 802, 804
                 Select Case Val(StrTemp7(3))
                 Case 1
                    strTemp(34) = "1"
                 Case 2
                    strTemp(35) = "1"
                 Case Else
                      BolOk = False
                 End Select
            Case Else
                 BolOk = False
            End Select
            If BolOk = True Then
                strSql = "INSERT INTO R040405 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & "," & Val(strTemp(27)) & "," & Val(strTemp(28)) & "," & Val(strTemp(29)) & "," & Val(strTemp(30)) & "," & Val(strTemp(31)) & "," & Val(strTemp(32)) & "," & Val(strTemp(33)) & "," & Val(strTemp(34)) & "," & Val(strTemp(35)) & "," & Val(strTemp(36)) & "," & Val(strTemp(37)) & "," & _
                          Val(strTemp(38)) & "," & Val(strTemp(39)) & ",'" & strUserNum & "') "
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
strSql = "SELECT R034001,SUM(R034002),SUM(R034003),SUM(R034004),SUM(R034005),SUM(R034006),SUM(R034007),SUM(R034008),SUM(R034009),SUM(R034010),SUM(R034011),SUM(R034012),SUM(R034013),SUM(R034014),SUM(R034015),SUM(R034016),SUM(R034017),SUM(R034018),SUM(R034019),sum(R034020),SUM(R034021),SUM(R034022),SUM(R034023),SUM(R034024),SUM(R034025),SUM(R034026),SUM(R034027),SUM(R034028),SUM(R034029),SUM(R034030),SUM(R034031),SUM(R034032),SUM(R034033),SUM(R034034),SUM(R034035),SUM(R034036),SUM(R034037),SUM(R034002+R034005+R034008+R034011+R034014+R034017+R034020+R034023+R034026+R034029+R034032+R034035+R034038),SUM(R034003+R034006+R034009+R034012+R034015+R034018+R034021+R034024+R034027+R034030+R034033+R034036+R034039),SUM(R034040) FROM R040405 WHERE ID='" & strUserNum & "' GROUP BY R034001 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    cnnConnection.Execute "DELETE FROM R040405 WHERE ID='" & strUserNum & "'"
    With adoRecordset
        .MoveFirst
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/2
        Do While .EOF = False
            For i = 0 To 39
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
            If Val(strTemp(22)) = 0 And Val(strTemp(23)) = 0 Then
                strTemp(24) = "0"
            Else
                strTemp(24) = Trim(str(Val(strTemp(22)) / (Val(strTemp(22)) + Val(strTemp(23))) * 100))
            End If
            If Val(strTemp(25)) = 0 And Val(strTemp(26)) = 0 Then
                strTemp(27) = "0"
            Else
                strTemp(27) = Trim(str(Val(strTemp(25)) / (Val(strTemp(25)) + Val(strTemp(26))) * 100))
            End If
            If Val(strTemp(28)) = 0 And Val(strTemp(29)) = 0 Then
                strTemp(30) = "0"
            Else
                strTemp(30) = Trim(str(Val(strTemp(28)) / (Val(strTemp(28)) + Val(strTemp(29))) * 100))
            End If
            If Val(strTemp(31)) = 0 And Val(strTemp(32)) = 0 Then
                strTemp(33) = "0"
            Else
                strTemp(33) = Trim(str(Val(strTemp(31)) / (Val(strTemp(31)) + Val(strTemp(32))) * 100))
            End If
            If Val(strTemp(34)) = 0 And Val(strTemp(35)) = 0 Then
                strTemp(36) = "0"
            Else
                strTemp(36) = Trim(str(Val(strTemp(34)) / (Val(strTemp(34)) + Val(strTemp(35))) * 100))
            End If
            If Val(strTemp(37)) = 0 And Val(strTemp(38)) = 0 Then
                strTemp(39) = "0"
            Else
                strTemp(39) = Trim(str(Val(strTemp(37)) / (Val(strTemp(37)) + Val(strTemp(38))) * 100))
            End If
            strSql = "INSERT INTO R040405 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & "," & Val(strTemp(27)) & "," & Val(strTemp(28)) & "," & Val(strTemp(29)) & "," & Val(strTemp(30)) & "," & Val(strTemp(31)) & "," & Val(strTemp(32)) & "," & Val(strTemp(33)) & "," & Val(strTemp(34)) & "," & Val(strTemp(35)) & "," & Val(strTemp(36)) & "," & Val(strTemp(37)) & "," & _
                          Val(strTemp(38)) & "," & Val(strTemp(39)) & ",'" & strUserNum & "') "
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
strSql = "SELECT * FROM R040405 WHERE ID='" & strUserNum & "' "
'列印順序
Select Case Val(txt1(7))
Case 1 '發明
'      strSQL = strSQL + " ORDER BY R034004 desc,R034001 "
      strSql = strSql + " ORDER BY R034004 desc ,R034002 Desc ,R034001 "
Case 2 '新型
'      strSQL = strSQL + " ORDER BY R034007 desc,R034001 "
      strSql = strSql + " ORDER BY R034007 desc, R034005 Desc, R034001 "
Case 3 '設計
'      strSQL = strSQL + " ORDER BY R034010 desc,R034001 "
      strSql = strSql + " ORDER BY R034010 desc, R034008 Desc, R034001 "
Case 4 '再審
'      strSQL = strSQL + " ORDER BY R034013 desc,R034016 DESC,R034019 DESC,R034001 "
      strSql = strSql + " ORDER BY R034013 desc,R034016 DESC,R034019 DESC, R034011 Desc, R034001 "
Case 5 '訴願再訴行訴
'      strSQL = strSQL + " ORDER BY R034022 desc,R034025 DESC,R034028 DESC,R034001 "
      strSql = strSql + " ORDER BY R034022 desc,R034025 DESC,R034028 DESC,R034020 Desc ,R034001 "
Case 6 '異議舉發答辯
'      strSQL = strSQL + " ORDER BY R034031 desc,R034034 DESC,R034037 DESC,R034001 "
      strSql = strSql + " ORDER BY R034031 desc,R034034 DESC,R034037 DESC, R034029 Desc, R034001 "
Case 7 '總件數
'      strSQL = strSQL + " ORDER BY R034040 desc,R034001 "
      strSql = strSql + " ORDER BY R034040 desc, R034038 Desc, R034001 "
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
            For i = 0 To 39
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
Printer.Print "准/駁統計明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6800
Printer.CurrentY = iPrint
Printer.Print "准駁日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(210, "-")
iPrint = iPrint + 300
Printer.Font.Size = 8
Printer.Font.Underline = True
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "發　　　明"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "新　　　型"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "新　式　樣"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "再　　　　　　　　　　　　　　　　審"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "訴　　　願"
Printer.CurrentX = PLeft(22)
Printer.CurrentY = iPrint
Printer.Print "再　訴　願"
Printer.CurrentX = PLeft(25)
Printer.CurrentY = iPrint
Printer.Print "行 政 訴 訟"
Printer.CurrentX = PLeft(28)
Printer.CurrentY = iPrint
Printer.Print "異　　　議"
Printer.CurrentX = PLeft(31)
Printer.CurrentY = iPrint
Printer.Print "舉　　　發"
Printer.CurrentX = PLeft(34)
Printer.CurrentY = iPrint
Printer.Print "異 答 舉 答"
Printer.CurrentX = PLeft(37)
Printer.CurrentY = iPrint
Printer.Print "總　件　數"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "發　　　明"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "新　　　型"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "新　式　樣"
iPrint = iPrint + 300
Printer.Font.Size = 8
Printer.Font.Underline = False
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
Printer.CurrentX = PLeft(22)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(23)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(24)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(25)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(26)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(27)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(28)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(29)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(30)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(31)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(32)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(33)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(34)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(35)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(36)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(37)
Printer.CurrentY = iPrint
Printer.Print "准"
Printer.CurrentX = PLeft(38)
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = PLeft(39)
Printer.CurrentY = iPrint
Printer.Print "核准率"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(210, "-")
iPrint = iPrint + 300
Printer.Font.Size = 8
End Sub

Sub PrintDatil()
Printer.Font.Size = 8
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 0 To 12
    Printer.CurrentX = PLeft((i * 3) + 1) + 200 - Printer.TextWidth(strTemp((i * 3) + 1))
    Printer.CurrentY = iPrint
    Printer.Print strTemp((i * 3) + 1)
    Printer.CurrentX = PLeft((i * 3) + 2) + 200 - Printer.TextWidth(strTemp((i * 3) + 2))
    Printer.CurrentY = iPrint
    Printer.Print strTemp((i * 3) + 2)
    Printer.CurrentX = PLeft((i * 3) + 3) + 550 - Printer.TextWidth(Format(strTemp((i * 3) + 3), "###.00") + "%")
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp((i * 3) + 3), "###.0") + "%"
Next i
iPrint = iPrint + 300
Printer.Font.Size = 8
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1500
PLeft(2) = 1750
PLeft(3) = 2000
PLeft(4) = 2600
PLeft(5) = 2850
PLeft(6) = 3100
PLeft(7) = 3700
PLeft(8) = 3950
PLeft(9) = 4200
PLeft(10) = 4800
PLeft(11) = 5050
PLeft(12) = 5300
PLeft(13) = 5900
PLeft(14) = 6150
PLeft(15) = 6400
PLeft(16) = 7000
PLeft(17) = 7250
PLeft(18) = 7500
PLeft(19) = 8100
PLeft(20) = 8350
PLeft(21) = 8600
PLeft(22) = 9200
PLeft(23) = 9450
PLeft(24) = 9700
PLeft(25) = 10300
PLeft(26) = 10550
PLeft(27) = 10800
PLeft(28) = 11400
PLeft(29) = 11650
PLeft(30) = 11900
PLeft(31) = 12500
PLeft(32) = 12750
PLeft(33) = 13000
PLeft(34) = 13600
PLeft(35) = 13850
PLeft(36) = 14100
PLeft(37) = 14700
PLeft(38) = 14950
PLeft(39) = 15200
End Sub

Sub PrintEnd()
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(210, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
'Modify By Cheng 2003/05/13
'strSQL = "SELECT ' ',SUM(R033002),SUM(R033003),SUM(R033004),SUM(R033005),SUM(R033006),SUM(R033007),SUM(R033008),SUM(R033009),SUM(R033010),SUM(R033011),SUM(R033012),SUM(R033013),SUM(R033014),SUM(R033015),SUM(R033016),SUM(R033017),SUM(R033018),SUM(R033019),SUM(R033020),SUM(R033021),SUM(R033022) FROM R040404 WHERE ID='" & strUserNum & "' "
strSql = "SELECT ' ',SUM(R034002),SUM(R034003),SUM(R034004),SUM(R034005),SUM(R034006),SUM(R034007),SUM(R034008),SUM(R034009),SUM(R034010),SUM(R034011),SUM(R034012),SUM(R034013),SUM(R034014),SUM(R034015),SUM(R034016),SUM(R034017),SUM(R034018),SUM(R034019),SUM(R034020),SUM(R034021),SUM(R034022) " & _
            ",SUM(R034023),SUM(R034024),SUM(R034025),SUM(R034026),SUM(R034027),SUM(R034028),SUM(R034029),SUM(R034030),SUM(R034031),SUM(R034032),SUM(R034033),SUM(R034034),SUM(R034035),SUM(R034036),SUM(R034037),SUM(R034038),SUM(R034039),SUM(R034040) FROM R040405 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Do While .EOF = False
'            For i = 0 To 21
            For i = 0 To 39
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
            'Add By Cheng 2003/05/13
            If Val(strTemp(22)) = 0 And Val(strTemp(23)) = 0 Then
                strTemp(24) = "0"
            Else
                strTemp(24) = Trim(str(Val(strTemp(22)) / (Val(strTemp(22)) + Val(strTemp(23))) * 100))
            End If
            If Val(strTemp(25)) = 0 And Val(strTemp(26)) = 0 Then
                strTemp(27) = "0"
            Else
                strTemp(27) = Trim(str(Val(strTemp(25)) / (Val(strTemp(25)) + Val(strTemp(26))) * 100))
            End If
            If Val(strTemp(28)) = 0 And Val(strTemp(29)) = 0 Then
                strTemp(30) = "0"
            Else
                strTemp(30) = Trim(str(Val(strTemp(28)) / (Val(strTemp(28)) + Val(strTemp(29))) * 100))
            End If
            If Val(strTemp(31)) = 0 And Val(strTemp(32)) = 0 Then
                strTemp(33) = "0"
            Else
                strTemp(33) = Trim(str(Val(strTemp(31)) / (Val(strTemp(31)) + Val(strTemp(32))) * 100))
            End If
            If Val(strTemp(34)) = 0 And Val(strTemp(35)) = 0 Then
                strTemp(36) = "0"
            Else
                strTemp(36) = Trim(str(Val(strTemp(34)) / (Val(strTemp(34)) + Val(strTemp(35))) * 100))
            End If
            If Val(strTemp(37)) = 0 And Val(strTemp(38)) = 0 Then
                strTemp(39) = "0"
            Else
                strTemp(39) = Trim(str(Val(strTemp(37)) / (Val(strTemp(37)) + Val(strTemp(38))) * 100))
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

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/12
   Select Case Index
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
Case 1, 2 '准駁日期起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
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
If Cancel Then TextInverse txt1(Index)
End Sub
