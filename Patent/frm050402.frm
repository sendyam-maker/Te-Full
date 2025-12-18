VERSION 5.00
Begin VB.Form frm050402 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人發文統計表"
   ClientHeight    =   2055
   ClientLeft      =   2370
   ClientTop       =   1230
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form21"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4110
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1125
      TabIndex        =   0
      Top             =   735
      Width           =   2130
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1035
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2445
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1035
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1125
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1335
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2445
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1335
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1635
      Width           =   255
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2190
      TabIndex        =   6
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2985
      TabIndex        =   7
      Top             =   36
      Width           =   800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
      Height          =   180
      Left            =   165
      TabIndex        =   12
      Top             =   735
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "發文日期："
      Height          =   180
      Left            =   165
      TabIndex        =   11
      Top             =   1035
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "申請國家："
      Height          =   180
      Left            =   165
      TabIndex        =   10
      Top             =   1335
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否含不計件案件："
      Height          =   180
      Left            =   165
      TabIndex        =   9
      Top             =   1635
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(Y.是)"
      Height          =   180
      Left            =   2325
      TabIndex        =   8
      Top             =   1635
      Width           =   465
   End
   Begin VB.Line Line1 
      X1              =   2085
      X2              =   2325
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line2 
      X1              =   2085
      X2              =   2325
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "frm050402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 14) As String
Dim PLeft(0 To 14) As Integer, strTemp1 As Variant, strTemp2 As Variant, BolNo As Boolean
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
   'Add By Cheng 2002/09/16
   blnClkSure = True
   
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
        If Len(txt1(2)) = 0 Then
            s = MsgBox("發文日期區間不可空白!!", , "USER 輸入錯誤")
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
                  MsgBox "發文日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
            If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
               If Me.txt1(3).Text > Me.txt1(4).Text Then
                  MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
                        
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
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

Private Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050402 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R050402_1 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/12/7
End If
If Len(Trim(txt1(1))) <> 0 Then
   strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " "
   strSQL2 = strSQL2 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL2 = strSQL2 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/7
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
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/7
End If
'Modify By Cheng 2003/04/09
'If Len(txt1(5)) = 0 Then
'    strSQL1 = strSQL1 + " AND CP21 IS NULL "
'    strSQL2 = strSQL2 + " AND CP21 IS NULL "
'End If
'若不含不計件案件
'If Len(txt1(5)) <> 0 Then
If Len(txt1(5)) = 0 Then
   strSQL1 = strSQL1 + " AND CP26 IS NULL "
   strSQL2 = strSQL2 + " AND CP26 IS NULL "
Else
   pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) 'Add By Sindy 2010/12/7
End If
'91.7.29 MODIFY BY SONIA 不管是否算案件數  92.3.26 張瓊玉72006不印
'strSQL = "SELECT ST02,decode(pa09,'000',cpm03,cpm04),CP10,PA08,CP14,ST03,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE ST01=CP14 AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND CP26 IS NULL AND CP57 IS NULL " & strSQL1
'strSQL = strSQL + " union all select ST02,decode(sp09,'000',cpm03,cpm04),CP10,'',CP14,ST03,CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE ST01=CP14 AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND CP26 IS NULL AND CP57 IS NULL " & strSQL2
'modify by sonia 2016/3/3 +CP14<>'87025'
strSql = "SELECT ST02,decode(pa09,'000',cpm03,cpm04),CP10,PA08,CP14,ST03,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE CP14=st01(+) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14<>'72006' AND CP14<>'87025' AND cp01=cpm01(+) AND cp10=cpm02(+) " & strSQL1
strSql = strSql + " union all select ST02,decode(sp09,'000',cpm03,cpm04),CP10,'',CP14,ST03,CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE CP14=st01(+) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP14<>'72006' AND CP14<>'87025' AND cp01=cpm01(+) AND cp10=cpm02(+) " & strSQL2
'91.7.29 END
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 14
                strTemp(i) = ""
            Next i
            For i = 0 To 6
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            BolNo = True
            '91.7.29 MODIFY BY SONIA
            'If Val(Mid(strTemp(5), 3, 1)) <> 9 Then
            If strTemp(5) = "" Or strTemp(5) <> "P12" Then
            '91.7.29 END
                Select Case Val(strTemp(2))
                Case 101
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                        strTemp(1) = "1"
                        strTemp(2) = ""
                        strTemp(3) = ""
                        strTemp(4) = ""
                        strTemp(5) = ""
                        strTemp(6) = ""
                     Else
                        BolNo = False
                     End If
                Case 102
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                        strTemp(2) = "1"
                        strTemp(1) = ""
                        strTemp(3) = ""
                        strTemp(4) = ""
                        strTemp(5) = ""
                        strTemp(6) = ""
                     Else
                        BolNo = False
                     End If
                Case 103
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                        strTemp(3) = "1"
                        strTemp(1) = ""
                        strTemp(2) = ""
                        strTemp(4) = ""
                        strTemp(5) = ""
                        strTemp(6) = ""
                     Else
                        BolNo = False
                     End If
                Case 104
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                         If Val(strTemp(3)) = 1 Then
                            strTemp(1) = "1"
                            strTemp(2) = ""
                            strTemp(3) = ""
                            strTemp(4) = ""
                            strTemp(5) = ""
                            strTemp(6) = ""
                         Else
                            strTemp(2) = "1"
                            strTemp(1) = ""
                            strTemp(3) = ""
                            strTemp(4) = ""
                            strTemp(5) = ""
                            strTemp(6) = ""
                         End If
                     Else
                         BolNo = False
                     End If
                Case 201
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                         strTemp(4) = "1"
                         strTemp(3) = ""
                         strTemp(5) = ""
                         strTemp(2) = ""
                         strTemp(1) = ""
                         strTemp(6) = ""
                     Else
                         BolNo = False
                     End If
                Case 1002
                     If UCase(Mid(strTemp(6), 1, 1)) = "C" Then
                         strTemp(5) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(6) = ""
                     Else
                         BolNo = False
                     End If
                Case 107
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                         strTemp(6) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(5) = ""
                     Else
                         BolNo = False
                     End If
                Case 203, 204
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                         strTemp(7) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(5) = ""
                         strTemp(6) = ""
                     Else
                         BolNo = False
                     End If
                Case 1209
                     If UCase(Mid(strTemp(6), 1, 1)) = "C" Then
                         strTemp(8) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(5) = ""
                         strTemp(6) = ""
                     Else
                         BolNo = False
                     End If
                Case 207
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                         strTemp(9) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(5) = ""
                         strTemp(6) = ""
                     Else
                         BolNo = False
                     End If
                Case 1205
                     If UCase(Mid(strTemp(6), 1, 1)) = "C" Then
                         strTemp(10) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(5) = ""
                         strTemp(6) = ""
                     Else
                        BolNo = False
                     End If
                Case 208
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                         strTemp(11) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(5) = ""
                         strTemp(6) = ""
                     Else
                         BolNo = False
                     End If
                Case 1206
                     If UCase(Mid(strTemp(6), 1, 1)) = "C" Then
                         strTemp(12) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(5) = ""
                         strTemp(6) = ""
                     Else
                         BolNo = False
                     End If
                     
               'Add by Morgan 2003/12/03
               '集體設計申請 -->設計申請
                Case 105
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                        strTemp(3) = "1"
                        strTemp(1) = ""
                        strTemp(2) = ""
                        strTemp(4) = ""
                        strTemp(5) = ""
                        strTemp(6) = ""
                     Else
                        BolNo = False
                     End If
                'CIP申請,分割-->依專利種類併入申請案
                Case 113, 307
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                        strTemp(1) = ""
                        strTemp(2) = ""
                        strTemp(4) = ""
                        strTemp(5) = ""
                        strTemp(6) = ""
                        Select Case Val(strTemp(3))
                           
                           Case 1
                              strTemp(1) = "1"
                              strTemp(3) = ""
                           Case 2
                              strTemp(2) = "1"
                              strTemp(3) = ""
                           Case 3
                              strTemp(3) = "1"
                              
                        End Select
                     Else
                        BolNo = False
                     End If
                     
                'End 2003/12/03
                
                Case Else
                     If UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                         strTemp(13) = "1"
                         strTemp(1) = ""
                         strTemp(2) = ""
                         strTemp(3) = ""
                         strTemp(4) = ""
                         strTemp(5) = ""
                         strTemp(6) = ""
                     Else
                         BolNo = False
                     End If
                End Select
                strTemp(14) = "1"
                If BolNo = True Then
                     'Modify By Cheng 2002/08/06
'                    strsql = "INSERT INTO R050402 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & ",'" & strUserNum & "') "
                    strSql = "INSERT INTO R050402 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & ",'" & strUserNum & "','" & .Fields(5) & "','" & .Fields(4) & "') "
                End If
                For i = 0 To 14
                    strTemp(i) = ""
                Next i
                If BolNo = True Then
                    cnnConnection.Execute strSql
                End If
            Else
               If strTemp(5) = "P12" And UCase(Mid(strTemp(6), 1, 1)) <> "C" Then
                  Select Case Val(strTemp(2))
                  Case 605
                       strTemp(1) = "1"
                       strTemp(2) = ""
                       strTemp(3) = ""
                       strTemp(4) = ""
                       strTemp(5) = ""
                       strTemp(6) = ""
                  Case 607
                       strTemp(2) = "1"
                       strTemp(1) = ""
                       strTemp(3) = ""
                       strTemp(4) = ""
                       strTemp(5) = ""
                       strTemp(6) = ""
                  Case 606
                       strTemp(3) = "1"
                       strTemp(2) = ""
                       strTemp(1) = ""
                       strTemp(4) = ""
                       strTemp(5) = ""
                       strTemp(6) = ""
                  Case 601
                       strTemp(4) = "1"
                       strTemp(2) = ""
                       strTemp(3) = ""
                       strTemp(1) = ""
                       strTemp(5) = ""
                       strTemp(6) = ""
                  Case 701
                       strTemp(5) = "1"
                       strTemp(2) = ""
                       strTemp(3) = ""
                       strTemp(4) = ""
                       strTemp(1) = ""
                       strTemp(6) = ""
                  Case 416
                       strTemp(6) = "1"
                       strTemp(2) = ""
                       strTemp(3) = ""
                       strTemp(4) = ""
                       strTemp(5) = ""
                       strTemp(1) = ""
                  Case 106
                       strTemp(7) = "1"
                       strTemp(2) = ""
                       strTemp(3) = ""
                       strTemp(4) = ""
                       strTemp(5) = ""
                       strTemp(1) = ""
                       strTemp(6) = ""
                  Case 405
                       strTemp(8) = "1"
                       strTemp(2) = ""
                       strTemp(3) = ""
                       strTemp(4) = ""
                       strTemp(5) = ""
                       strTemp(1) = ""
                       strTemp(6) = ""
                  Case Else
                       strTemp(9) = "1"
                       strTemp(2) = ""
                       strTemp(3) = ""
                       strTemp(4) = ""
                       strTemp(5) = ""
                       strTemp(1) = ""
                       strTemp(6) = ""
                  End Select
                  strTemp(10) = "1"
                  If BolNo = True Then
                        'Modify By Cheng 2002/08/06
'                      strsql = "INSERT INTO R050402_1 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & ",'" & strUserNum & "') "
                      strSql = "INSERT INTO R050402_1 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & ",'" & strUserNum & "','" & .Fields(5) & "','" & .Fields(4) & "') "
                  End If
                  For i = 0 To 14
                      strTemp(i) = ""
                  Next i
                  If BolNo = True Then
                      cnnConnection.Execute strSql
                  End If
               End If
            End If
            .MoveNext

            DoEvents
        Loop
    End With
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/7
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
'Modify By Cheng 2002/08/06
'strsql = "SELECT R015001,SUM(R015002),SUM(R015003),SUM(R015004),SUM(R015005),SUM(R015006),SUM(R015007),SUM(R015008),SUM(R015009),SUM(R015010),SUM(R015011),SUM(R015012),SUM(R015013),SUM(R015014),SUM(R015015) FROM R050402 WHERE ID='" & strUserNum & "' GROUP BY R015001 "
strSql = "SELECT R015001,SUM(R015002),SUM(R015003),SUM(R015004),SUM(R015005),SUM(R015006),SUM(R015007),SUM(R015008),SUM(R015009),SUM(R015010),SUM(R015011),SUM(R015012),SUM(R015013),SUM(R015014),SUM(R015015),R015016,R015017 FROM R050402 WHERE ID='" & strUserNum & "' GROUP BY R015001,R015016,R015017 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    cnnConnection.Execute "DELETE FROM R050402 WHERE ID='" & strUserNum & "' "
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        For i = 0 To 14
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
         'Modify By Cheng 2002/08/06
'        strsql = "INSERT INTO R050402 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & ",'" & strUserNum & "') "
        strSql = "INSERT INTO R050402 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & ",'" & strUserNum & "','" & adoRecordset.Fields(15) & "','" & adoRecordset.Fields(16) & "') "
        cnnConnection.Execute strSql
        adoRecordset.MoveNext
    Loop
End If
CheckOC
'Modify By Cheng 2002/08/06
'strsql = "SELECT R016001,SUM(R016002),SUM(R016003),SUM(R016004),SUM(R016005),SUM(R016006),SUM(R016007),SUM(R016008),SUM(R016009),SUM(R016010),SUM(R016011) FROM R050402_1 WHERE ID='" & strUserNum & "' GROUP BY R016001 "
strSql = "SELECT R016001,SUM(R016002),SUM(R016003),SUM(R016004),SUM(R016005),SUM(R016006),SUM(R016007),SUM(R016008),SUM(R016009),SUM(R016010),SUM(R016011),R016012,R016013 FROM R050402_1 WHERE ID='" & strUserNum & "' GROUP BY R016001,R016012,R016013 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    cnnConnection.Execute "DELETE FROM R050402_1 WHERE ID='" & strUserNum & "' "
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        For i = 0 To 10
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
         'Modify By Cheng 2002/08/06
'        strsql = "INSERT INTO R050402_1 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & ",'" & strUserNum & "') "
        strSql = "INSERT INTO R050402_1 VALUES('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & ",'" & strUserNum & "','" & adoRecordset.Fields(11) & "','" & adoRecordset.Fields(12) & "') "
        cnnConnection.Execute strSql
        adoRecordset.MoveNext
    Loop
End If
CheckOC
PrintData
PrintData1
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Private Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "承辦人發文統計表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "日期：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
'Add By Cheng 2003/10/20
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Me.txt1(3) & "－" & Me.txt1(4).Text
'End
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
If Len(txt1(5)) = 0 Then
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
'   Printer.Print "不含多國案"
   Printer.Print "不含不計件案件"
Else
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
'   Printer.Print "含多國案"
   Printer.Print "含不計件案件"
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "工程師"
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
Printer.Print "翻譯"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "核駁分析"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "答辯"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "修正"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "檢索報告"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "提供前案"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "通知提供前案"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "選取"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "通知要求選取"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "合計"
iPrint = iPrint + 300
Printer.Font.Underline = False
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Private Sub PrintTitle1()
GetPleft1
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "承辦人發文統計表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "日期：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
'Add By Cheng 2003/10/20
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Me.txt1(3) & "－" & Me.txt1(4).Text
'End
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
If Len(txt1(5)) = 0 Then
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
'   Printer.Print "不含多國案"
   Printer.Print "不含不計件案件"
Else
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
'   Printer.Print "含多國案"
   Printer.Print "含不計件案件"
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "程    序"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "繳 年 費"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "延    展"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "維 持 費"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "領    證"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "讓    與"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "實體審查"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "主張優先權"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "優先權證明書"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "合    計"
iPrint = iPrint + 300
Printer.Font.Underline = False
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Private Sub PrintData()
'Modify By Cheng 2002/08/06
'strsql = "SELECT * FROM R050402 WHERE ID='" & strUserNum & "'"
strSql = "SELECT * FROM R050402 WHERE ID='" & strUserNum & "' ORDER BY R015016,R015017 "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 14
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If iPrint > 10000 Then
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            PrintDatil
            .MoveNext
        Loop
    End With
Else
   Exit Sub
End If
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint > 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
strSql = "SELECT '合計',SUM(R015002),SUM(R015003),SUM(R015004),SUM(R015005),SUM(R015006),SUM(R015007),SUM(R015008),SUM(R015009),SUM(R015010),SUM(R015011),SUM(R015012),SUM(R015013),SUM(R015014),SUM(R015015) FROM R050402 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        For i = 0 To 14
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        PrintDatil
        adoRecordset.MoveNext
    Loop
End If
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
Printer.EndDoc
End Sub

Private Sub PrintDatil()
For i = 0 To 14
   If i = 0 Then
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Else
      Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   End If
Next i
iPrint = iPrint + 300
End Sub

Private Sub PrintDatil1()
For i = 0 To 10
   If i = 0 Then
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Else
      Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   End If
Next i
iPrint = iPrint + 300
End Sub

Private Sub PrintData1()
'Modify By Cheng 2002/08/06
'strsql = "SELECT * FROM R050402_1 WHERE ID='" & strUserNum & "'"
strSql = "SELECT * FROM R050402_1 WHERE ID='" & strUserNum & "' ORDER BY R016012,R016013 "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If iPrint > 10000 Then
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                Page = Page + 1
                Printer.NewPage
                PrintTitle1
            End If
            PrintDatil1
            .MoveNext
        Loop
    End With
Else
   Exit Sub
End If
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint > 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
strSql = "SELECT '合計',SUM(R016002),SUM(R016003),SUM(R016004),SUM(R016005),SUM(R016006),SUM(R016007),SUM(R016008),SUM(R016009),SUM(R016010),SUM(R016011) FROM R050402_1 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        For i = 0 To 10
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        PrintDatil1
        adoRecordset.MoveNext
    Loop
End If
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
Printer.EndDoc
End Sub

Private Sub GetPleft1()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 2200
PLeft(2) = 3500
PLeft(3) = 4700
PLeft(4) = 6200
PLeft(5) = 7500
PLeft(6) = 8700
PLeft(7) = 10000
PLeft(8) = 11400
PLeft(9) = 12900
PLeft(10) = 13800

End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1350
PLeft(2) = 2000
PLeft(3) = 2600
PLeft(4) = 3400
PLeft(5) = 4000
PLeft(6) = 5100
PLeft(7) = 5700
PLeft(8) = 6300
PLeft(9) = 7500
PLeft(10) = 8700
PLeft(11) = 10300
PLeft(12) = 11000
PLeft(13) = 12700
PLeft(14) = 13500
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050402 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/16
   Select Case Index
   Case 5 '小計是否含不計件案件
        'Modify By Cheng 2003/04/09
      If KeyAscii <> 89 And KeyAscii <> 8 Then
'        If KeyAscii <> 78 And KeyAscii <> 8 Then
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
Case 2, 4
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
Case 5 '小計是否含不計件案件
     Select Case Trim(txt1(5))
     Case "Y", "y", "", " "
'     Case "N", "n", "", " "
     Case Else
'          s = MsgBox("是否計算多國案件只能 Y 或 空白 !!", , "USER 輸入錯誤")
'          s = MsgBox("是否含不計件案件只能為 N 或 空白 !!", , "USER 輸入錯誤")
          s = MsgBox("小計是否含不計件案件只能為 Y 或 空白 !!", , "USER 輸入錯誤")
          txt1(5).SetFocus
          txt1(5).SelStart = 0
          txt1(5).SelLength = Len(txt1(5))
          Exit Sub
     End Select
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '發文日期起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select

End Sub
