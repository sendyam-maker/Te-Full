VERSION 5.00
Begin VB.Form frm040308 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收文明細表"
   ClientHeight    =   2700
   ClientLeft      =   5730
   ClientTop       =   2085
   ClientWidth     =   3915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3915
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3024
      TabIndex        =   13
      Top             =   12
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2208
      TabIndex        =   12
      Top             =   12
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   2220
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2304
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1020
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2304
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   2220
      MaxLength       =   9
      TabIndex        =   9
      Top             =   1992
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1020
      MaxLength       =   9
      TabIndex        =   8
      Top             =   1992
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1680
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1020
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1680
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2220
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1368
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1368
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1020
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1056
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2220
      MaxLength       =   3
      TabIndex        =   2
      Top             =   744
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   1
      Top             =   744
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   432
      Width           =   2325
   End
   Begin VB.Line Line5 
      X1              =   2124
      X2              =   2214
      Y1              =   2412
      Y2              =   2412
   End
   Begin VB.Line Line4 
      X1              =   2124
      X2              =   2229
      Y1              =   2184
      Y2              =   2184
   End
   Begin VB.Line Line3 
      X1              =   2076
      X2              =   2226
      Y1              =   1848
      Y2              =   1848
   End
   Begin VB.Line Line2 
      X1              =   2136
      X2              =   2226
      Y1              =   1512
      Y2              =   1512
   End
   Begin VB.Line Line1 
      X1              =   2124
      X2              =   2199
      Y1              =   888
      Y2              =   888
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   2244
      TabIndex        =   21
      Top             =   1068
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   2316
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   2016
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1716
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "收文日期："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1404
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1104
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   804
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   492
      Width           =   996
   End
End
Attribute VB_Name = "frm040308"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 8) As String, strTemp1 As Variant, strTemp2 As Variant
Dim PLeft(0 To 8) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 4) As String, StrTemp5(0 To 4) As String, StrTemp6(0 To 4) As String
'Add By Cheng 2002/09/11
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
      'Add By Cheng 2002/09/11
      blnClkSure = False
      
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
         'Add By Cheng 2002/03/19
         If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
            Me.txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(5)) = -1 Then
            Me.txt1(5).SetFocus
            txt1_GotFocus 5
            Exit Sub
         End If
         'Add By Cheng 2002/09/11
         If Me.txt1(4).Text <> "" And Me.txt1(5).Text <> "" Then
            If Val(Me.txt1(4).Text) > Val(Me.txt1(5).Text) Then
               MsgBox "收文日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(4).SetFocus
               txt1_GotFocus 4
               Exit Sub
            End If
         End If
         
         If Len(txt1(5)) = 0 Then
            s = MsgBox("收文日期區間不可空白!!", , "USER 輸入錯誤")
             txt1(4).SetFocus
             txt1_GotFocus (4)
            Exit Sub
         Else
            If Len(Trim(txt1(8))) <> 0 And Len(Trim(txt1(9))) <> 0 Then
                If Left(txt1(8), 6) <> Left(txt1(9), 6) Then
                    s = MsgBox("申請人前六碼必須相同!!", , "USER 輸入錯誤")
                     blnClkSure = True
                    txt1(8).SetFocus
                    txt1(8).SelStart = 0
                    txt1(8).SelLength = Len(txt1(8))
                    Exit Sub
                End If
            End If
            If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) <> 0 Then
                If Left(txt1(10), 6) <> Left(txt1(11), 6) Then
                    s = MsgBox("代理人前六碼必須相同!!", , "USER 輸入錯誤")
                     blnClkSure = True
                    txt1(10).SetFocus
                    txt1(10).SelStart = 0
                    txt1(10).SelLength = Len(txt1(10))
                    Exit Sub
                End If
            End If
            'Add By Cheng 2002/09/11
            If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
               If Me.txt1(1).Text > Me.txt1(2).Text Then
                  MsgBox "業務區別範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
            If Me.txt1(3).Text <> "" Then
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetStaff(txt1(3), strExc(0)) Then
               If ClsPDGetStaff(txt1(3), strExc(0)) Then
                  lbl1 = strExc(0)
               Else
                  lbl1 = ""
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
            If Me.txt1(4).Text <> "" And Me.txt1(5).Text <> "" Then
               If Val(Me.txt1(4).Text) > Val(Me.txt1(5).Text) Then
                  MsgBox "收文日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 4
                  Exit Sub
               End If
            End If
            If Me.txt1(6).Text <> "" And Me.txt1(7).Text <> "" Then
               If Me.txt1(6).Text > Me.txt1(7).Text Then
                  MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(6).SetFocus
                  txt1_GotFocus 6
                  Exit Sub
               End If
            End If
            If Me.txt1(8).Text <> "" And Me.txt1(9).Text <> "" Then
               If Me.txt1(8).Text > Me.txt1(9).Text Then
                  MsgBox "申請人代號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(8).SetFocus
                  txt1_GotFocus 8
                  Exit Sub
               End If
            End If
            If Me.txt1(10).Text <> "" And Me.txt1(11).Text <> "" Then
               If Me.txt1(10).Text > Me.txt1(11).Text Then
                  MsgBox "代理人代號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(10).SetFocus
                  txt1_GotFocus 10
                  Exit Sub
               End If
            End If
            
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
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
strSql = "DELETE FROM R040308 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute strSql
strSQL1 = ""
strSQL2 = ""
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " and cp12>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND cp12<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/2
End If
If Len(txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " and CP13='" & txt1(3) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & lbl1 'Add By Sindy 2010/12/2
End If
If Len(Trim(txt1(4))) <> 0 Then
   strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(4))) & " "
End If
If Len(Trim(txt1(5))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(5))) & " "
End If
If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/12/2
End If
If Len(txt1(6)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & txt1(6) & "' "
End If
If Len(txt1(7)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & txt1(7) & "' "
End If
If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/2
End If
strSQL2 = strSQL1
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 & " and CP01 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/2
End If
If Len(Trim(txt1(8))) <> 0 And Len(Trim(txt1(9))) <> 0 Then
    strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(8)) & "' AND PA26<='" & GetNewFagent(txt1(9)) & "') OR (PA27>='" & GetNewFagent(txt1(8)) & "' AND PA27<='" & GetNewFagent(txt1(9)) & "') OR (PA28>='" & GetNewFagent(txt1(8)) & "' AND PA28<='" & GetNewFagent(txt1(9)) & "') OR (PA29>='" & GetNewFagent(txt1(8)) & "' AND PA29<='" & GetNewFagent(txt1(9)) & "') OR (PA30>='" & GetNewFagent(txt1(8)) & "' AND PA30<='" & GetNewFagent(txt1(9)) & "')) "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(8)) & "' AND SP08<='" & GetNewFagent(txt1(9)) & "') OR (SP58<='" & GetNewFagent(txt1(8)) & "' AND SP58<='" & GetNewFagent(txt1(9)) & "') OR (SP59>='" & GetNewFagent(txt1(8)) & "' AND SP59<='" & GetNewFagent(txt1(9)) & "')) "
Else
    If Len(Trim(txt1(8))) <> 0 And Len(Trim(txt1(9))) = 0 Then
        strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(8)) & "' OR PA27>='" & GetNewFagent(txt1(8)) & "' OR PA28>='" & GetNewFagent(txt1(8)) & "' OR PA29>='" & GetNewFagent(txt1(8)) & "' OR PA30>='" & GetNewFagent(txt1(8)) & "') "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(8)) & "' OR SP58>='" & GetNewFagent(txt1(8)) & "' OR SP59>='" & GetNewFagent(txt1(8)) & "') "
    Else
        If Len(Trim(txt1(8))) = 0 And Len(Trim(txt1(9))) <> 0 Then
            strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(9)) & "' OR PA27<='" & GetNewFagent(txt1(9)) & "' OR PA28<='" & GetNewFagent(txt1(9)) & "' OR PA29<='" & GetNewFagent(txt1(9)) & "' OR PA30<='" & GetNewFagent(txt1(9)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(9)) & "' OR SP58<='" & GetNewFagent(txt1(9)) & "' OR SP59<='" & GetNewFagent(txt1(9)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(8))) <> 0 Or Len(Trim(txt1(9))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(8) & "-" & txt1(9) 'Add By Sindy 2010/12/2
End If
'代理人
If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) <> 0 Then
    strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(10)) & "' AND PA75<='" & GetNewFagent(txt1(11)) & "' "
    strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(10)) & "' AND SP26<='" & GetNewFagent(txt1(11)) & "' "
Else
    If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) = 0 Then
        strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(10)) & "' "
        strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(10)) & "' "
    Else
        If Len(Trim(txt1(10))) = 0 And Len(Trim(txt1(11))) <> 0 Then
            strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(11)) & "' "
            strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(11)) & "' "
        End If
    End If
End If
If Len(Trim(txt1(10))) <> 0 Or Len(Trim(txt1(11))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/12/2
End If
strSql = "SELECT NVL(A0902,A0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",'" & strUserNum & "' FROM CASEPROGRESS,ACC090,STAFF,PATENT,CASEPROPERTYMAP,NATION,CUSTOMER WHERE CP12=a0901(+) AND cp13=ST01(+) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND cu10=NA01(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) " & strSQL1
strSql = strSql + " UNION ALL SELECT NVL(A0902,A0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",'" & strUserNum & "' FROM CASEPROGRESS,ACC090,STAFF,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,CUSTOMER WHERE cp12=A0901(+) AND cp13=ST01(+) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND cu10=NA01(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) " & strSQL2
cnnConnection.Execute "insert into r040308 " & strSql
CheckOC
strSql = "select * from r040308 where id='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/2
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/2
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
strSql = "SELECT * FROM R040308 WHERE ID='" & strUserNum & "' ORDER BY R025001,R025002,R025003,R025004 "
CheckOC
Page = 1
strTemp3(0) = " "
strTemp3(1) = " "
strTemp3(2) = " "
strTemp3(3) = " "
strTemp3(4) = " "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        strTemp(0) = CheckStr(.Fields(0))
        strTemp(1) = CheckStr(.Fields(1))
        strTemp3(0) = strTemp(0)
        strTemp3(1) = strTemp(1)
        PrintTitle
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp3(0) <> strTemp(0) Then
                PrintEnd
                PrintTitle2
                strTemp3(0) = strTemp(0)
                strTemp3(1) = strTemp(1)
                strTemp3(2) = strTemp(2)
                strTemp3(3) = strTemp(3)
                strTemp3(4) = strTemp(4)
            Else
                If strTemp3(1) <> strTemp(1) Then
                    PrintEnd
                    PrintTitle2
                    strTemp3(1) = strTemp(1)
                    strTemp3(2) = strTemp(2)
                    strTemp3(3) = strTemp(3)
                    strTemp3(4) = strTemp(4)
                Else
                    If strTemp3(2) <> strTemp(2) Then
                        strTemp3(2) = strTemp(2)
                        strTemp3(3) = strTemp(3)
                        strTemp3(4) = strTemp(4)
                    Else
                        strTemp(2) = ""
                        If strTemp3(3) <> strTemp(3) Then
                            strTemp3(3) = strTemp(3)
                            strTemp3(4) = strTemp(4)
                        Else
                            strTemp(3) = ""
                            If strTemp3(4) <> strTemp(4) Then
                                strTemp3(4) = strTemp(4)
                            Else
                                strTemp(4) = ""
                            End If
                        End If
                    End If
                End If
            End If
            strTemp(4) = StrToStr(strTemp(4), 24)
            strTemp(5) = StrToStr(strTemp(5), 6)
            strTemp(6) = StrToStr(strTemp(6), 4)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                .MoveNext
                If .EOF = False Then
                    Printer.NewPage
                    PrintTitle
                    If strTemp3(0) = CheckStr(.Fields(0)) Or strTemp3(1) = CheckStr(.Fields(1)) Then
                        PrintTitle2
                    End If
                End If
                .MovePrevious
            End If
            .MoveNext
        Loop
    End With
End If
PrintEnd
Printer.EndDoc
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "智權人員收文明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(4)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))
iPrint = iPrint + 300
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
End Sub

Sub PrintTitle2()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區：" & strTemp(0)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員：" & strTemp(1)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
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
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請人國籍"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "發文日"
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
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
End Sub

Sub PrintDatil()
For i = 2 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintEnd()
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
strSql = "SELECT R025006,COUNT(R025006) FROM R040308 WHERE R025001='" & strTemp3(0) & "' AND R025002='" & strTemp3(1) & "' AND ID='" & strUserNum & "' GROUP BY R025006 "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    With adoRecordset1
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 4
                StrTemp5(i) = CheckStr(.Fields(0))
                StrTemp6(i) = CheckStr(.Fields(1))
                Printer.CurrentX = 500 + (i * 2500)
                Printer.CurrentY = iPrint
                Printer.Print StrToStr(StrTemp5(i), 6)
                Printer.CurrentX = 2500 + (i * 2500) - Printer.TextWidth(StrTemp6(i))
                Printer.CurrentY = iPrint
                Printer.Print StrTemp6(i)
                .MoveNext
                If .EOF = True Then
                    Exit For
                End If
            Next i
            iPrint = iPrint + 300
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If .EOF = False Then
                    PrintTitle2
                End If
            End If
            Loop
    End With
End If
CheckOC2
iPrint = iPrint + 300
    
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 3000
PLeft(2) = 500
PLeft(3) = 1500
PLeft(4) = 3500
PLeft(5) = 9500
PLeft(6) = 11000
PLeft(7) = 12500
PLeft(8) = 14000
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040308 = Nothing
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
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
        End If
     Next i
   Case 2, 5, 7
      'Modify By Cheng 2002/09/11
      If blnClkSure = False Then
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
         End If
      Else
         blnClkSure = False
      End If
   Case 9
      'Modify By Cheng 2002/09/11
      If blnClkSure = False Then
         If Len(txt1(Index - 1)) <> 0 Then
            If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
                s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
                txt1(Index - 1).SetFocus
                Exit Sub
            End If
         End If
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
         End If
      Else
         blnClkSure = False
      End If
   Case 11
      'Modify By Cheng 2002/09/11
      If blnClkSure = False Then
         If Len(txt1(Index - 1)) <> 0 Then
            If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
                s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
                txt1(Index - 1).SetFocus
                Exit Sub
            End If
         End If
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
         End If
      Else
         blnClkSure = False
      End If
End Select

End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 3
   If txt1(Index) <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaff(txt1(Index), strExc(0)) Then
      If ClsPDGetStaff(txt1(Index), strExc(0)) Then
         lbl1 = strExc(0)
      Else
         lbl1 = ""
         Cancel = True
      End If
   End If
Case 4, 5 '收文日期起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
   End If
End Select
If Cancel Then TextInverse txt1(Index)
End Sub
