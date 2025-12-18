VERSION 5.00
Begin VB.Form frm060403 
   BorderStyle     =   1  '單線固定
   Caption         =   "准駁統計總表"
   ClientHeight    =   2430
   ClientLeft      =   1980
   ClientTop       =   2220
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3816
      TabIndex        =   6
      Top             =   24
      Width           =   972
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3024
      TabIndex        =   5
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1155
      TabIndex        =   4
      Top             =   1680
      Width           =   195
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1155
      TabIndex        =   3
      Top             =   1350
      Width           =   210
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1020
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1155
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1020
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1155
      TabIndex        =   0
      Top             =   690
      Width           =   3300
   End
   Begin VB.Line Line1 
      X1              =   2055
      X2              =   2175
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "(1.發明 2.新型 3.設計 4.再審 5.救濟程序       6.異議舉發答辯 7.總件數)"
      Height          =   360
      Index           =   5
      Left            =   1425
      TabIndex        =   12
      Top             =   1695
      Width           =   3315
   End
   Begin VB.Label Label1 
      Caption         =   "(1.承辦人 2.智權人員)"
      Height          =   180
      Index           =   4
      Left            =   1455
      TabIndex        =   11
      Top             =   1380
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      Height          =   180
      Index           =   3
      Left            =   195
      TabIndex        =   10
      Top             =   1710
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   2
      Left            =   195
      TabIndex        =   9
      Top             =   1395
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "准駁日期："
      Height          =   180
      Index           =   1
      Left            =   195
      TabIndex        =   8
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   195
      TabIndex        =   7
      Top             =   750
      Width           =   915
   End
End
Attribute VB_Name = "frm060403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL6 As String, i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 21) As String, strTemp1 As Variant, strTemp2 As Variant, k As Integer, StrTemp7(0 To 4) As String, SeekPrint As Integer, SeekPrintL As Integer
Dim PLeft(0 To 21) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 2) As String, BolOk As Boolean
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         'Add By Cheng 2002/09/16
         blnClkSure = False
         
           Printer.Orientation = 2
           DoEvents
           If Len(txt1(0)) = 0 Then
               s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
               txt1(0).SetFocus
               Exit Sub
           Else
               If Len(txt1(2)) = 0 Then
                   s = MsgBox("准駁日期區間不可空白!!", , "USER 輸入錯誤")
                   txt1(1).SetFocus
                   txt1_GotFocus (1)
                   Exit Sub
               Else
                  'Add By Cheng 2002/03/21
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
                       s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                       txt1(3).SetFocus
                       Exit Sub
                   Else
                       If Len(txt1(4)) = 0 Then
                           s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
                           txt1(4).SetFocus
                           Exit Sub
                       Else
                           Screen.MousePointer = vbHourglass
                           Me.Enabled = False
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
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
   cnnConnection.Execute "DELETE FROM R060403 WHERE ID='" & strUserNum & "' "
   strSQL1 = ""
   'Modify By Cheng 2002/10/07
   'strSQL2 = ""
   StrSQL6 = ""
   '系統類別
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
      'Modify By Cheng 2002/10/07
   '   strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/13
   End If
   '准駁日期
   If Len(Trim(txt1(1))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP25>=" & Val(ChangeTStringToWString(txt1(1))) & " "
   End If
   If Len(Trim(txt1(2))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP25<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/13
   End If
   If Val(txt1(3)) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label1(2) & "1.承辦人" 'Add By Sindy 2010/12/13
   Else
      pub_QL05 = pub_QL05 & ";" & Label1(2) & "2.智權人員" 'Add By Sindy 2010/12/13
   End If
   Select Case Val(txt1(4))
   Case 1
         pub_QL05 = pub_QL05 & ";" & Label1(3) & "1.發明" 'Add By Sindy 2010/12/13
   Case 2
         pub_QL05 = pub_QL05 & ";" & Label1(3) & "2.新型" 'Add By Sindy 2010/12/13
   Case 3
         pub_QL05 = pub_QL05 & ";" & Label1(3) & "3.設計" 'Add By Sindy 2010/12/13
   Case 4
         pub_QL05 = pub_QL05 & ";" & Label1(3) & "4.再審" 'Add By Sindy 2010/12/13
   Case 5
         pub_QL05 = pub_QL05 & ";" & Label1(3) & "5.救濟程序" 'Add By Sindy 2010/12/13
   Case 6
         pub_QL05 = pub_QL05 & ";" & Label1(3) & "6.異議舉發答辯" 'Add By Sindy 2010/12/13
   Case 7
         pub_QL05 = pub_QL05 & ";" & Label1(3) & "7.總件數" 'Add By Sindy 2010/12/13
   Case Else
   End Select
   'Modify By Cheng 2002/10/07
   'strSQL = "SELECT S1.ST02,CP10,PA08,CP24,S2.ST02 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
   'strSQL = strSQL + " UNION ALL SELECT S1.ST02,CP10,' ',CP24,S2.ST02 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
   '92.04.02 nick add left join
   'strSQL = "SELECT S1.ST02,CP10,PA08,CP24,S2.ST02,CP01,CP02,CP03,CP04 FROM PATENT,CASEPROGRESS,STAFF S1,STAFF S2 WHERE PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA16=CP24 AND PA20=CP25 AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
   strSql = "SELECT S1.ST02,CP10,PA08,CP24,S2.ST02,CP01,CP02,CP03,CP04 FROM PATENT,CASEPROGRESS,STAFF S1,STAFF S2 WHERE cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp24=pa16(+) and cp25=pa20(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
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
               If Val(txt1(3)) = "1" Then
                   strTemp(0) = StrTemp7(0)
               Else
                   strTemp(0) = StrTemp7(4)
               End If
               Select Case Val(StrTemp7(1))
               Case 101
                    Select Case Val(StrTemp7(3))
                    Case 1
                         strTemp(1) = "1"
                    Case 2
                         strTemp(2) = "1"
                    Case Else
                         BolOk = False
                    End Select
               Case 102
                    Select Case Val(StrTemp7(3))
                    Case 1
                         strTemp(4) = "1"
                    Case 2
                         strTemp(5) = "1"
                    Case Else
                         BolOk = False
                    End Select
               Case 103, 105
                    Select Case Val(StrTemp7(3))
                    Case 1
                         strTemp(7) = "1"
                    Case 2
                         strTemp(8) = "1"
                    Case Else
                         BolOk = False
                    End Select
               Case 104
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
               Case 107
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
                  '若為新申請案
                  If Val(StrTemp7(1)) >= 101 And Val(StrTemp7(1)) <= 105 Then
                     strSql = "select count(*) from caseprogress,staff where cp01='" & CheckStr(.Fields(5)) & "' and cp02='" & CheckStr(.Fields(6)) & "' and cp03='" & CheckStr(.Fields(7)) & "' and cp04='" & CheckStr(.Fields(8)) & "' and cp14=st01(+) and ST15='F22'"
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.Fields(0) > 0 Then
                        'Modify By Cheng 2002/10/07
                        '取得案件性質為"203"至"206"且發文日最大的資料
                         strSql = "select ST02 from caseprogress,STAFF where CP14=ST01(+) AND cp01='" & CheckStr(.Fields(5)) & "' and cp02='" & CheckStr(.Fields(6)) & "' and cp03='" & CheckStr(.Fields(7)) & "' and cp04='" & CheckStr(.Fields(8)) & "' and (cp10>='203' and cp10<='206') AND CP27 IS NOT NULL ORDER BY CP27 DESC "
                         CheckOC2
                         adoRecordset1.CursorLocation = adUseClient
                         adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                         If adoRecordset1.RecordCount = 0 Then
                           'Modify By Cheng 2002/10/07
                           '取得案件性質為"201"核稿人的資料
                           strSql = "select ST02 from caseprogress,ENGINEERPROGRESS,STAFF where EP04=ST01(+) AND CP09=EP02(+) AND cp01='" & CheckStr(.Fields(5)) & "' and cp02='" & CheckStr(.Fields(6)) & "' and cp03='" & CheckStr(.Fields(7)) & "' and cp04='" & CheckStr(.Fields(8)) & "' and cp10='201' AND EP04 IS NOT NULL "
                           CheckOC2
                           adoRecordset1.CursorLocation = adUseClient
                           adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                           If adoRecordset1.RecordCount = 0 Then
                              strSql = "select ST02 from caseprogress,STAFF where CP14=ST01(+) AND cp01='" & CheckStr(.Fields(5)) & "' and cp02='" & CheckStr(.Fields(6)) & "' and cp03='" & CheckStr(.Fields(7)) & "' and cp04='" & CheckStr(.Fields(8)) & "' and cp10='201'"
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
                  End If
                  strSql = "INSERT INTO R060403 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & ",'" & strUserNum & "') "
                  cnnConnection.Execute strSql
               End If
               .MoveNext
               k = k + 1
               DoEvents
           Loop
       End With
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/12/13
       ShowNoData
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   CheckOC
   strSql = "SELECT R051001,SUM(R051002),SUM(R051003),SUM(R051004),SUM(R051005),SUM(R051006),SUM(R051007),SUM(R051008),SUM(R051009),SUM(R051010),SUM(R051011),SUM(R051012),SUM(R051013),SUM(R051014),SUM(R051015),SUM(R051016),SUM(R051017),SUM(R051018),SUM(R051019),SUM(R051002+R051005+R051008+R051011+R051014+R051017),SUM(R051003+R051006+R051009+R051012+R051015+R051018),SUM(R051022) FROM R060403 WHERE ID='" & strUserNum & "' GROUP BY R051001 "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       cnnConnection.Execute "DELETE FROM R060403 WHERE ID='" & strUserNum & "'"
       With adoRecordset
           InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
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
               strSql = "INSERT INTO R060403 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & ",'" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
           Loop
       End With
   End If
   CheckOC
   PrintData
   ShowPrintOk
   Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
   strSql = "SELECT * FROM R060403 WHERE ID='" & strUserNum & "' "
   Select Case Val(txt1(4))
   Case 1
         strSql = strSql + " ORDER BY R051004 desc,R051001 "
   Case 2
         strSql = strSql + " ORDER BY R051007 desc,R051001 "
   Case 3
         strSql = strSql + " ORDER BY R051010 desc,R051001 "
   Case 4
         strSql = strSql + " ORDER BY R051013 desc,R051001 "
   Case 5
         strSql = strSql + " ORDER BY R051016 desc,R051001 "
   Case 6
         strSql = strSql + " ORDER BY R051019 desc,R051001 "
   Case 7
         strSql = strSql + " ORDER BY R051021 desc,R051001 "
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
   End If
   PrintEnd
   Printer.EndDoc
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
   Printer.Print GetTitleNick & "准/駁統計總表"
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   iPrint = iPrint + 500
   Printer.CurrentX = 6700
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
   Printer.Print "救濟程序"
   Printer.CurrentX = PLeft(16)
   Printer.CurrentY = iPrint
   Printer.Print "異議舉發答辯"
   Printer.CurrentX = PLeft(19)
   Printer.CurrentY = iPrint
   Printer.Print "總件數"
   iPrint = iPrint + 300
   Printer.Font.Size = 10
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
   strSql = "SELECT ' ',SUM(R051002),SUM(R051003),SUM(R051004),SUM(R051005),SUM(R051006),SUM(R051007),SUM(R051008),SUM(R051009),SUM(R051010),SUM(R051011),SUM(R051012),SUM(R051013),SUM(R051014),SUM(R051015),SUM(R051016),SUM(R051017),SUM(R051018),SUM(R051019),SUM(R051020),SUM(R051021),SUM(R051022) FROM R060403 WHERE ID='" & strUserNum & "' "
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

'Private Sub Combo1_Change()
'If Combo1.ListIndex >= SeekPrint Then
'   j = Combo1.ListIndex + 1
'Else
'   j = Combo1.ListIndex
'End If
'Set Printer = Printers(j)

'End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060403 = Nothing
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
   Case 4
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
         'Add By Cheng 2002/09/16
         If blnClkSure = False Then
            If RunNick(txt1(1), txt1(2)) Then
               txt1(1).SetFocus
               txt1_GotFocus (1)
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
      Case 3
         'Modify By Cheng 2002/09/17
         'If Me.txt1(3).Text <> "" Then  MODIFY BY SONIA 91.9.25
           Select Case Val(txt1(3))
           Case 1, 2
           Case Else
                s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
                txt1(3).SetFocus
                txt1(3).SelStart = 0
                txt1(3).SelLength = Len(txt1(3))
                Exit Sub
           End Select
         'End If  91.9.25 END
      Case 4
         'Modify By Cheng 2002/09/17
         'If Me.txt1(4).Text <> "" Then MODIFY BY SONIA 91.9.25
           Select Case Val(txt1(4))
           Case 1, 2, 3, 4, 5, 6, 7
           Case Else
                s = MsgBox("列印順序只能輸入 1 到 7 !!", , "USER 輸入錯誤")
                txt1(4).SetFocus
                txt1(4).SelStart = 0
                txt1(4).SelLength = Len(txt1(4))
                Exit Sub
           End Select
         'End If  91.9.25 END
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1, 2 '准駁日期起, 迄
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Cancel = True
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
         End If
   End Select
End Sub
