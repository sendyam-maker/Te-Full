VERSION 5.00
Begin VB.Form frm050308 
   BorderStyle     =   1  '單線固定
   Caption         =   "收文明細表"
   ClientHeight    =   3840
   ClientLeft      =   2850
   ClientTop       =   1770
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4185
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   2160
      MaxLength       =   9
      TabIndex        =   12
      Top             =   3168
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   960
      MaxLength       =   9
      TabIndex        =   11
      Top             =   3168
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2868
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   2160
      MaxLength       =   9
      TabIndex        =   14
      Top             =   3504
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   960
      MaxLength       =   9
      TabIndex        =   13
      Top             =   3504
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2505
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   960
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2196
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   960
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1860
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1524
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   960
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1524
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1188
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   960
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1188
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   2
      Top             =   888
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   960
      MaxLength       =   7
      TabIndex        =   1
      Top             =   888
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   552
      Width           =   1995
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2910
      TabIndex        =   16
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2115
      TabIndex        =   15
      Top             =   24
      Width           =   756
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "智權人員所別："
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   2535
      Width           =   1260
   End
   Begin VB.Line Line5 
      X1              =   1800
      X2              =   2040
      Y1              =   3624
      Y2              =   3624
   End
   Begin VB.Line Line4 
      X1              =   1800
      X2              =   2040
      Y1              =   3288
      Y2              =   3288
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   2040
      Y1              =   1644
      Y2              =   1644
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   2040
      Y1              =   1308
      Y2              =   1308
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   2040
      Y1              =   1008
      Y2              =   1008
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   3165
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(Y 計算)"
      Height          =   180
      Left            =   2160
      TabIndex        =   29
      Top             =   2868
      Width           =   648
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "是否計算多國案件："
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   2865
      Width           =   1620
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   120
      TabIndex        =   27
      Top             =   3510
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(1. 北 2. 中 3. 南 4. 高 5. ALL)"
      Height          =   180
      Left            =   1755
      TabIndex        =   26
      Top             =   2535
      Width           =   2250
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1908
      TabIndex        =   24
      Top             =   2232
      Width           =   1512
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "(1. 承辦人2. 所別)"
      Height          =   180
      Left            =   1440
      TabIndex        =   22
      Top             =   1860
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "列印順序："
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   1530
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   1185
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   885
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   555
      Width           =   900
   End
End
Attribute VB_Name = "frm050308"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSql As String, i As Integer, j As Integer, s As Integer, k As Integer
Dim strTemp1 As Variant, strTemp2 As Variant, StrTest1 As String, StrTest2 As String
Dim StrSQL3 As String, strSQL2 As String, strSQL1 As String, strTemp(0 To 9) As String
Dim PLeft1(0 To 9) As Integer, PLeft(0 To 9) As Integer, iPrint As Integer, Page As Integer, iLine As Integer, St As String, TmpArea As String
Dim StrTemp4(0 To 5) As String, StrTemp5(0 To 5) As String, StrTemp6 As String, StrTemp7 As String, StrTemp8 As String
Dim strSystem As String
'Add By Cheng 2002/07/26
Dim strRecDate As String '用來判斷當收文日相同時, 除第一筆外, 其餘不列印出來
Dim strSaleZone As String '業務區
Dim strSales As String '智權人員
'Add By Cheng 2002/09/12
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
      'Add By Cheng 2002/09/12
      blnClkSure = False
      
     If Len(Trim(txt1(0))) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1(0).SelStart = 0
        txt1(0).SelLength = Len(txt1(0))
        Exit Sub
     Else
        If Len(Trim(txt1(7))) = 0 Then
            s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
            txt1(7).SetFocus
            txt1(7).SelStart = 0
            txt1(7).SelLength = Len(txt1(7))
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
            
            If Len(Trim(txt1(1))) = 0 Then
                s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
                txt1(1).SetFocus
                txt1_GotFocus (1)
                Exit Sub
            Else
                If Len(Trim(txt1(9))) = 0 Then
                    s = MsgBox("所別不可空白!!", , "USER 輸入錯誤")
                    txt1(9).SetFocus
                    txt1(9).SelStart = 0
                    txt1(9).SelLength = Len(txt1(9))
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
                        txt1_GotFocus 6
                        Exit Sub
                     End If
                  End If
                  If Me.txt1(8).Text <> "" Then
                     lbl1 = GetPrjSales(txt1(8))
                     If Me.lbl1.Caption = Me.txt1(8).Text Then
                        Me.lbl1.Caption = ""
                        Me.txt1(8).SetFocus
                        txt1_GotFocus 8
                        Exit Sub
                     End If
                  End If
                    If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) <> 0 Then
                        If Left(txt1(11), 6) <> Left(txt1(12), 6) Then
                            s = MsgBox("申請人前六碼必須相同!!", , "USER 輸入錯誤")
                            blnClkSure = True
                            txt1(11).SetFocus
                            txt1(11).SelStart = 0
                            txt1(11).SelLength = Len(txt1(11))
                            Exit Sub
                        End If
                    End If
                  If Me.txt1(11).Text <> "" And Me.txt1(12).Text <> "" Then
                     If Me.txt1(11).Text > Me.txt1(12).Text Then
                        MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(11).SetFocus
                        txt1_GotFocus 11
                        Exit Sub
                     End If
                  End If
                    
                    If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) <> 0 Then
                        If Left(txt1(13), 6) <> Left(txt1(14), 6) Then
                            s = MsgBox("代理人前六碼必須相同!!", , "USER 輸入錯誤")
                            blnClkSure = True
                            txt1(13).SetFocus
                            txt1(13).SelStart = 0
                            txt1(13).SelLength = Len(txt1(13))
                            Exit Sub
                        End If
                    End If
                  If Me.txt1(13).Text <> "" And Me.txt1(14).Text <> "" Then
                     If Me.txt1(13).Text > Me.txt1(14).Text Then
                        MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(13).SetFocus
                        txt1_GotFocus 13
                        Exit Sub
                     End If
                  End If
                    
                    strTemp1 = Split(txt1(0), ",")
                    strSystem = strTemp1(0)
                    ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位
                    Select Case Val(txt1(7))
                    Case 1
                         Process             '承辦人
                    Case 2
                         Process1            '所別
                    Case Else
                    End Select
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
'910318 nick
cnnConnection.Execute "DELETE FROM R050308_1 WHERE ID='" & strUserNum & "' "
pub_QL05 = pub_QL05 & ";" & Label5 & "1. 承辦人" 'Add By Sindy 2010/01/22
'系統類別
strSQL1 = ""
strSQL2 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/01/22
End If
'收文日
If Len(Trim(txt1(1))) <> 0 Then
   strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
   strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL2 = strSQL2 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/01/22
'申請國家
If Len(Trim(txt1(3))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
End If
If Len(Trim(txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
End If
If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/01/22
End If
'案件性質
If Len(Trim(txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND CP10>='" & txt1(5) & "' "
End If
If Len(Trim(txt1(6))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & txt1(6) & "' "
    strSQL2 = strSQL2 + " AND CP10<='" & txt1(6) & "' "
End If
If Len(Trim(txt1(5))) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/01/22
End If
'承辦人
If Len(txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP14='" & txt1(8) & "' "
    strSQL2 = strSQL2 + " AND CP14='" & txt1(8) & "' "
    pub_QL05 = pub_QL05 & ";" & Label7 & txt1(8) 'Add By Sindy 2010/01/22
End If
'所別
If Len(txt1(9)) <> 0 And Val(txt1(9)) <> 5 Then
    strSQL1 = strSQL1 + " AND ST06='" & Trim(txt1(9)) & "' "
    strSQL2 = strSQL2 + " AND ST06='" & Trim(txt1(9)) & "' "
    pub_QL05 = pub_QL05 & ";" & Label9 & txt1(9) & Label10 'Add By Sindy 2010/01/22
End If
'是否計算多國案件
If UCase(txt1(10)) = "Y" Then
    '2005/6/2 CANCEL BY SONIA
    'strSQL1 = strSQL1 + " AND CP21='Y' "
    'strSQL2 = strSQL2 + " AND CP21='Y' "
    '2005/6/2 END
    pub_QL05 = pub_QL05 & ";" & Label11 & txt1(10) 'Add By Sindy 2010/01/22
Else
    strSQL1 = strSQL1 + " AND CP21 is null "
    strSQL2 = strSQL2 + " AND CP21 is null "
End If
'申請人
If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) <> 0 Then
    strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(11)) & "' AND PA26<='" & GetNewFagent(txt1(12)) & "') OR (PA27>='" & GetNewFagent(txt1(11)) & "' AND PA27<='" & GetNewFagent(txt1(12)) & "') OR (PA28>='" & GetNewFagent(txt1(11)) & "' AND PA28<='" & GetNewFagent(txt1(12)) & "') OR (PA29>='" & GetNewFagent(txt1(11)) & "' AND PA29<='" & GetNewFagent(txt1(12)) & "') OR (PA30>='" & GetNewFagent(txt1(11)) & "' AND PA30<='" & GetNewFagent(txt1(12)) & "')) "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(11)) & "' AND SP08<='" & GetNewFagent(txt1(12)) & "') OR (SP58<='" & GetNewFagent(txt1(11)) & "' AND SP58<='" & GetNewFagent(txt1(12)) & "') OR (SP59>='" & GetNewFagent(txt1(11)) & "' AND SP59<='" & GetNewFagent(txt1(12)) & "')) "
Else
    If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) = 0 Then
        strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(11)) & "' OR PA27>='" & GetNewFagent(txt1(11)) & "' OR PA28>='" & GetNewFagent(txt1(11)) & "' OR PA29>='" & GetNewFagent(txt1(11)) & "' OR PA30>='" & GetNewFagent(txt1(11)) & "') "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(11)) & "' OR SP58>='" & GetNewFagent(txt1(11)) & "' OR SP59>='" & GetNewFagent(txt1(11)) & "') "
    Else
        If Len(Trim(txt1(11))) = 0 And Len(Trim(txt1(12))) <> 0 Then
            strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(12)) & "' OR PA27<='" & GetNewFagent(txt1(12)) & "' OR PA28<='" & GetNewFagent(txt1(12)) & "' OR PA29<='" & GetNewFagent(txt1(12)) & "' OR PA30<='" & GetNewFagent(txt1(12)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(12)) & "' OR SP58<='" & GetNewFagent(txt1(12)) & "' OR SP59<='" & GetNewFagent(txt1(12)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label13 & txt1(11) & "-" & txt1(12) 'Add By Sindy 2010/01/22
End If
'代理人
If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) <> 0 Then
    strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(13)) & "' AND PA75<='" & GetNewFagent(txt1(14)) & "' "
    strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(13)) & "' AND SP26<='" & GetNewFagent(txt1(14)) & "' "
Else
    If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) = 0 Then
        strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(13)) & "' "
        strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(13)) & "' "
    Else
        If Len(Trim(txt1(13))) = 0 And Len(Trim(txt1(14))) <> 0 Then
            strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(14)) & "' "
            strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(14)) & "' "
        End If
    End If
End If
If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/01/22
End If
'組合
'Modify By Cheng 2002/08/22
'910318  nick
'91.5.1 MODIFY BY SONIA, PTM03 ONLY
'strSQL = "select cp14," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),PTM03,cp10,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "',cp01,pa09 FROM CASEPROGRESS,PATENT,STAFF,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND pa09=na01(+) AND CP14=ST01(+)  " & strsql1
strSql = "select cp14," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),PTM03,cp10,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "',cp01,pa09 FROM CASEPROGRESS,PATENT,STAFF,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND pa09=na01(+) AND CP13=ST01(+)  " & strSQL1
'strSQL = "select cp14," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',PTM03,PTM04),cp10,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "',cp01,pa09 FROM CASEPROGRESS,PATENT,STAFF,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND pa09=na01(+) AND CP14=ST01(+)  " & strSQL1
'91.5.1 END
'strSQL = strSQL + " union all select cp14," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),'',cp10,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "',cp01,sp09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,NATION,CASEPROPERTYMAP WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) AND CP14=ST01(+)  " & strsql2
strSql = strSql + " union all select cp14," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),'',cp10,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "',cp01,sp09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,NATION,CASEPROPERTYMAP WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) AND CP13=ST01(+)  " & strSQL2
'strSQL = strSQL + " ORDER BY ST01,CP05,CP01,CP02,CP03,CP04 "
'910318 nick
cnnConnection.Execute "insert into r050308_1 " & strSql
strSql = "select * from r050308_1 where id='" & strUserNum & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
Else
   InsertQueryLog (0) 'Add By Sindy 2010/01/22
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()             '印表主程式
'搜尋承辦人及筆數
'StrSQL = "SELECT * FROM R050308_1 ORDER BY R007001,R007002,R007003 "
strSql = "SELECT R007001,COUNT(R007006) FROM R050308_1 WHERE ID='" & strUserNum & "' GROUP BY R007001 "
CheckOC
Page = 0
'Add By Cheng 2002/07/26
strRecDate = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    
    Do While adoRecordset.EOF = False
        DoEvents
        'Add By Cheng 2002/07/26
        strRecDate = ""
        
        StrTemp6 = CheckStr(adoRecordset.Fields(0))
        Page = Page + 1
        PrintTitle
        '搜尋個別承辦人收文明細資料
        strSql = "SELECT R007001,R007002,R007003, R007004, R007005, decode(r007012,'000',cpm03,decode(cpm04,null,R007006,cpm04)) , R007007, R007008 , R007009 , R007010,r007011,r007012  FROM R050308_1,casepropertymap WHERE r007011 = cpm01(+) and r007006=cpm02(+) and R007001='" & StrTemp6 & "' AND ID='" & strUserNum & "' ORDER BY R007002,R007003 "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      'Modify By Cheng 2002/07/26
'        If adoRecordset1.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            With adoRecordset1
                .MoveFirst
                Do While .EOF = False
                    For i = 0 To 9
                        strTemp(i) = CheckStr(.Fields(i))
                    Next i
                    'Add By Cheng 2002/07/26
                    '若收文日同前一筆, 則收文日不印出來
                    If strRecDate = strTemp(1) Then
                       strTemp(1) = ""
                    Else
                       strRecDate = "" & strTemp(1)
                    End If
                    strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 32), vbUnicode)
                    strTemp(5) = StrToStr(strTemp(5), 4)
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                        'Add By Cheng 2002/07/26
                        strTemp(1) = "" & strRecDate
                    
                    End If
                   PrintDatil
                    .MoveNext
                Loop
            End With
        End If
        CheckOC2
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
      'Modify By Cheng 2002/08/06
'        strsql = "select decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),COUNT(*) FROM R050308_1,casepropertymap WHERE '" & strSystem & "' = cpm01(+) and r007006=cpm02(+) and R007001='" & StrTemp6 & "' AND ID='" & strUserNum & "' GROUP BY decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) "
        strSql = "select decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),COUNT(*),r007006 FROM R050308_1,casepropertymap WHERE '" & strSystem & "' = cpm01(+) and r007006=cpm02(+) and R007001='" & StrTemp6 & "' AND ID='" & strUserNum & "' GROUP BY decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),r007006 ORDER BY r007006 "
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 = 0 Then
                adoRecordset1.MoveFirst
                For j = 1 To (adoRecordset1.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'If adoRecordset1.EOF = False Then
                        PrintTotil
                    'End If
                Next j
            Else
                If adoRecordset1.RecordCount < 5 Then
                    adoRecordset1.MoveFirst
                    For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    For k = adoRecordset1.RecordCount Mod 5 To 4
                        StrTemp4(k) = ""
                        StrTemp5(k) = ""
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'If adoRecordset1.EOF = False Then
                     PrintTotil
                    'End If
                Else
                    If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 <> 0 Then
                        adoRecordset1.MoveFirst
                        For j = 1 To (adoRecordset1.RecordCount \ 5)
                            For k = 0 To 4
                                StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                                StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                                If adoRecordset1.EOF = False Then
                                    adoRecordset1.MoveNext
                                End If
                            Next k
                            If iPrint > 10000 Then
                                PrintEnd
                                Printer.NewPage
                                Page = Page + 1
                                PrintTitle
                            End If
                            'If adoRecordset1.EOF = False Then
                                 PrintTotil
                            'End If
                        Next j
                        For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                            StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                            StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                            If adoRecordset1.EOF = False Then
                                adoRecordset1.MoveNext
                            End If
                            Next k
                        For k = adoRecordset1.RecordCount Mod 5 To 4
                            StrTemp4(k) = ""
                            StrTemp5(k) = ""
                        Next k
                        If iPrint > 10000 Then
                            PrintEnd
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle
                        End If
                        'If adoRecordset1.EOF = False Then
                           PrintTotil
                        'End If
                    End If
                End If
            End If
        End If
        CheckOC2
        PrintEnd
        Printer.NewPage
        adoRecordset.MoveNext
    Loop
    CheckOC
    StrTemp6 = "ALL"
    Page = Page + 1
    PrintTitle
   'Modify By Cheng 2002/08/06
'    strsql = "select decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),COUNT(*) FROM R050308_1,casepropertymap WHERE '" & strSystem & "'= cpm01(+) and r007006=cpm02(+) AND ID='" & strUserNum & "' GROUP BY decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) "
    strSql = "select decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),COUNT(*),r007006 FROM R050308_1,casepropertymap WHERE '" & strSystem & "'= cpm01(+) and r007006=cpm02(+) AND ID='" & strUserNum & "' GROUP BY decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),r007006 ORDER BY r007006 "
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 = 0 Then
            adoRecordset.MoveFirst
            For j = 1 To (adoRecordset.RecordCount \ 5)
                For k = 0 To 4
                    StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                    StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                    If adoRecordset.EOF = False Then
                        adoRecordset.MoveNext
                    End If
                Next k
                If iPrint > 10000 Then
                    PrintEnd
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle
                End If
                'If adoRecordset.EOF = False Then
                  PrintTotilEnd
                'End If
            Next j
        Else
            If adoRecordset.RecordCount < 5 Then
                adoRecordset.MoveFirst
                For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                    StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                    StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                    If adoRecordset.EOF = False Then
                        adoRecordset.MoveNext
                    End If
                Next k
                For k = adoRecordset.RecordCount Mod 5 To 4
                    StrTemp4(k) = ""
                    StrTemp5(k) = ""
                Next k
                If iPrint > 10000 Then
                    PrintEnd
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle
                End If
                'If adoRecordset.EOF = False Then
                  PrintTotilEnd
                'End If
            Else
                If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 <> 0 Then
                    adoRecordset.MoveFirst
                    'Modify By Cheng 2002/07/26
'                    For j = 1 To (adoRecordset.RecordCount \ 5)
                    For j = 1 To (adoRecordset.RecordCount \ 5)
                        For k = 0 To 4
                            StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                            StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                            If adoRecordset.EOF = False Then
                                adoRecordset.MoveNext
                            End If
                        Next k
                        If iPrint > 10000 Then
                            PrintEnd
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle
                        End If
                        'If adoRecordset.EOF = False Then
                           PrintTotilEnd
                        'End If
                    Next j
                    For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                        StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                        If adoRecordset.EOF = False Then
                            adoRecordset.MoveNext
                        End If
                    Next k
                    For k = adoRecordset.RecordCount Mod 5 To 4
                        StrTemp4(k) = ""
                        StrTemp5(k) = ""
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'If adoRecordset.EOF = False Then
                        PrintTotil
                    'End If
                End If
            End If
        End If
    End If
    CheckOC
End If
PrintEnd
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle()          '印抬頭
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "收文明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & IIf(StrTemp6 = "ALL", StrTemp6, GetPrjSalesNM(StrTemp6))
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "專利種類"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "准/駁"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintEnd()            '印結尾
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintDatil()           '印內容
For j = 1 To 9
    Printer.CurrentX = PLeft(j)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(j)
Next j
iPrint = iPrint + 300
End Sub

Sub PrintTotil()            '印小計
'Add By Cheng 2002/07/28
Dim jj As Integer

For jj = 0 To 4
    Printer.CurrentX = 500 + (jj * 2000)
    Printer.CurrentY = iPrint
    Printer.Print StrConv(MidB(StrConv(StrTemp4(jj), vbFromUnicode), 1, 10), vbUnicode)
    Printer.CurrentX = 2300 + (jj * 2000) - Printer.TextWidth(StrTemp5(jj))
    Printer.CurrentY = iPrint
    Printer.Print StrTemp5(jj)
Next jj
iPrint = iPrint + 300

End Sub

Sub PrintTotilEnd()          '印合計
'Add By Cheng 2002/07/28
Dim jj As Integer

For jj = 0 To 4
    Printer.CurrentX = 500 + (jj * 2000)
    Printer.CurrentY = iPrint
    Printer.Print StrConv(MidB(StrConv(StrTemp4(jj), vbFromUnicode), 1, 10), vbUnicode)
    Printer.CurrentX = 2300 + (jj * 2000) - Printer.TextWidth(StrTemp5(jj))
    Printer.CurrentY = iPrint
    Printer.Print StrTemp5(jj)
Next jj
iPrint = iPrint + 300
End Sub
 
Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 3700
PLeft(4) = 7700
PLeft(5) = 9100
PLeft(6) = 10300
PLeft(7) = 11300
PLeft(8) = 12800
PLeft(9) = 13800
End Sub

Sub PrintData1()

'Add By Cheng 2002/07/26
strSaleZone = ""
strSales = ""
strRecDate = ""

Page = 1
'搜尋所別
strSql = "SELECT distinct R008001,decode(R008001,'北所','1','中所','2','南所','3','其他','4') as a FROM R050308_2 where id='" & strUserNum & "' group BY R008001,decode(R008001,'北所','1','中所','2','南所','3','其他','4') order by a "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        StrTemp6 = CheckStr(adoRecordset.Fields(0))
        CheckOC2
        '搜尋業務區及智權人員
        If Len(StrTemp6) = 0 Then
            strSql = "SELECT distinct R008002,R008003 FROM R050308_2 WHERE R008001 is null and id='" & strUserNum & "' GROUP BY R008002,R008003 "
        Else
            strSql = "SELECT distinct R008002,R008003 FROM R050308_2 WHERE R008001='" & StrTemp6 & "' and id='" & strUserNum & "' GROUP BY R008002,R008003 "
        End If
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            adoRecordset1.MoveFirst
            Do While adoRecordset1.EOF = False
                'Add By Cheng 2002/07/26
                strSaleZone = ""
                strSales = ""
                strRecDate = ""
                
                PrintTitle1
                StrTemp7 = CheckStr(adoRecordset1.Fields(0))
                StrTemp8 = CheckStr(adoRecordset1.Fields(1))
                strSql = "SELECT decode(R008001,'北所','1','中所','2','南所','3','其他','4'),nvl(a0902,r008002),st02,r008004,r008005,r008006,r008007,decode(r008012,'000',cpm03,decode(cpm04,null,r008008,cpm04)),r008009,r008010,r008003 FROM R050308_2,acc090,staff,casepropertymap WHERE r008011=cpm01(+) and r008008=cpm02(+) and r008003 = st01(+) and r008002=a0901(+) and id='" & strUserNum & "' " & IIf(Len(StrTemp6) = 0, " and r008001 is null ", " and R008001='" & StrTemp6 & "' ") & IIf(Len(StrTemp7) = 0, " and r008002 is null ", " and R008002='" & StrTemp7 & "' ") & IIf(Len(StrTemp7) = 0, " and r008003 is null ", " and R008003='" & StrTemp8 & "' ") & " ORDER BY R008004 "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                    AdoRecordSet3.MoveFirst
                    Do While AdoRecordSet3.EOF = False
                        For i = 0 To 9
                            strTemp(i) = CheckStr(AdoRecordSet3.Fields(i))
                        Next i
                        strTemp(5) = StrConv(MidB(StrConv(strTemp(5), vbFromUnicode), 1, 28), vbUnicode)
                        'Add By Cheng 2002/07/26
                        If strSaleZone = "" & strTemp(1) Then
                           strTemp(1) = ""
                           If strSales = "" & strTemp(2) Then
                              strTemp(2) = ""
                              If strRecDate = "" & strTemp(3) Then
                                 strTemp(3) = ""
                              Else
                                 strRecDate = "" & strTemp(3)
                              End If
                           Else
                              strSales = "" & strTemp(2)
                              strRecDate = "" & strTemp(3)
                           End If
                        Else
                           strSaleZone = "" & strTemp(1)
                           strSales = "" & strTemp(2)
                           strRecDate = "" & strTemp(3)
                        End If
                                          
                        PrintDatil1
                        If iPrint > 10000 Then
                            PrintEnd1
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle1
                            'Add By Cheng 2002/07/26
                            strTemp(1) = "" & strSaleZone
                            strTemp(2) = "" & strSales
                            strTemp(3) = "" & strRecDate
                            
                        End If
                        AdoRecordSet3.MoveNext
                    Loop
                End If
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
                If iPrint > 10000 Then
                    PrintEnd1
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle1
                End If
                CheckOC3
               'Modify By Cheng 2002/08/06
'                strsql = "SELECT decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),COUNT(*) FROM R050308_2,casepropertymap  WHERE id='" & strUserNum & "' and '" & strSystem & "'=cpm01(+) and r008008=cpm02(+) " & IIf(Len(StrTemp6) = 0, " and r008001 is null ", "and R008001='" & StrTemp6 & "' ") & IIf(Len(StrTemp7) = 0, " and r008002 is null ", "and R008002='" & StrTemp7 & "' ") & IIf(Len(StrTemp7) = 0, " and r008003 is null ", "and R008003='" & StrTemp8 & "' ") & " GROUP BY decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) "
                strSql = "SELECT decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),COUNT(*),r008008 FROM R050308_2,casepropertymap  WHERE id='" & strUserNum & "' and '" & strSystem & "'=cpm01(+) and r008008=cpm02(+) " & IIf(Len(StrTemp6) = 0, " and r008001 is null ", "and R008001='" & StrTemp6 & "' ") & IIf(Len(StrTemp7) = 0, " and r008002 is null ", "and R008002='" & StrTemp7 & "' ") & IIf(Len(StrTemp7) = 0, " and r008003 is null ", "and R008003='" & StrTemp8 & "' ") & " GROUP BY decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),r008008 ORDER BY r008008 "
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                    If AdoRecordSet3.RecordCount \ 5 <> 0 And AdoRecordSet3.RecordCount Mod 5 = 0 Then
                        AdoRecordSet3.MoveFirst
                        For j = 1 To (AdoRecordSet3.RecordCount \ 5)
                            For k = 0 To 4
                                StrTemp4(k) = CheckStr(AdoRecordSet3.Fields(0))
                                StrTemp5(k) = CheckStr(AdoRecordSet3.Fields(1))
                                If AdoRecordSet3.EOF = False Then
                                    AdoRecordSet3.MoveNext
                                End If
                            Next k
                            If iPrint > 10000 Then
                                PrintEnd1
                                Printer.NewPage
                                Page = Page + 1
                                PrintTitle1
                            End If
                            'If AdoRecordSet3.EOF = False Then
                              PrintTotil1
                            'End If
                            
                        Next j
                    Else
                        If AdoRecordSet3.RecordCount < 5 Then
                            AdoRecordSet3.MoveFirst
                            For k = 0 To ((AdoRecordSet3.RecordCount Mod 5) - 1)
                                StrTemp4(k) = CheckStr(AdoRecordSet3.Fields(0))
                                StrTemp5(k) = CheckStr(AdoRecordSet3.Fields(1))
                                If AdoRecordSet3.EOF = False Then
                                    AdoRecordSet3.MoveNext
                                End If
                            Next k
                            For k = AdoRecordSet3.RecordCount Mod 5 To 4
                                StrTemp4(k) = ""
                                StrTemp5(k) = ""
                            Next k
                            If iPrint > 10000 Then
                                PrintEnd1
                                Printer.NewPage
                                Page = Page + 1
                                PrintTitle1
                            End If
                            'If AdoRecordSet3.EOF = False Then
                              PrintTotil1
                            'End If
                            
                        Else
                            If AdoRecordSet3.RecordCount \ 5 <> 0 And AdoRecordSet3.RecordCount Mod 5 <> 0 Then
                                AdoRecordSet3.MoveFirst
                                For j = 1 To (AdoRecordSet3.RecordCount \ 5)
                                    For k = 0 To 4
                                        StrTemp4(k) = CheckStr(AdoRecordSet3.Fields(0))
                                        StrTemp5(k) = CheckStr(AdoRecordSet3.Fields(1))
                                        If AdoRecordSet3.EOF = False Then
                                            AdoRecordSet3.MoveNext
                                        End If
                                    Next k
                                    PrintTotil1
                                    If iPrint > 10000 Then
                                        PrintEnd1
                                        Printer.NewPage
                                        Page = Page + 1
                                        PrintTitle1
                                    End If
                                Next j
                                For k = 0 To ((AdoRecordSet3.RecordCount Mod 5) - 1)
                                    StrTemp4(k) = CheckStr(AdoRecordSet3.Fields(0))
                                    StrTemp5(k) = CheckStr(AdoRecordSet3.Fields(1))
                                    If AdoRecordSet3.EOF = False Then
                                        AdoRecordSet3.MoveNext
                                    End If
                                Next k
                                For k = AdoRecordSet3.RecordCount Mod 5 To 4
                                    StrTemp4(k) = ""
                                    StrTemp5(k) = ""
                                Next k
                                If iPrint > 10000 Then
                                    PrintEnd1
                                    Printer.NewPage
                                    Page = Page + 1
                                    PrintTitle1
                                End If
                                'If AdoRecordSet3.EOF = False Then
                                 PrintTotil1
                                'End If
                                
                            End If
                        End If
                    End If
                End If
                CheckOC3
                adoRecordset1.MoveNext
                If adoRecordset1.EOF = False Then
                    Printer.NewPage
                    Page = Page + 1
                End If
            Loop
        End If
        adoRecordset.MoveNext
        If adoRecordset.EOF = False Then
            PrintEnd1
            Printer.NewPage
            Page = Page + 1
        End If
    Loop
    CheckOC
    PrintEnd1
    Printer.NewPage
    Page = Page + 1
    StrTemp6 = "ALL"
    PrintTitle1
   'Modify By Cheng 2002/08/06
'    '910318  nick
'    strsql = "SELECT decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),COUNT(*) FROM R050308_2,casepropertymap  WHERE id='" & strUserNum & "' and '" & strSystem & "'=cpm01(+) and r008008=cpm02(+) GROUP BY decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) "
    strSql = "SELECT decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),COUNT(*),r008008 FROM R050308_2,casepropertymap  WHERE id='" & strUserNum & "' and '" & strSystem & "'=cpm01(+) and r008008=cpm02(+) GROUP BY decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),r008008 ORDER BY r008008 "
    'strSQL = "select R008008,COUNT(R008008) FROM R050308_2 where id='" & strUserNum & "' GROUP BY R008008 "
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 = 0 Then
            adoRecordset.MoveFirst
            For j = 1 To (adoRecordset.RecordCount \ 5)
                For k = 0 To 4
                    StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                    StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                    If adoRecordset.EOF = False Then
                        adoRecordset.MoveNext
                    End If
                Next k
                If iPrint > 10000 Then
                    PrintEnd1
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle1
                End If
                'If adoRecordset.EOF = False Then
                     PrintTotilEnd1
                'End If
            Next j
        Else
            If adoRecordset.RecordCount < 5 Then
                adoRecordset.MoveFirst
                For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                    StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                    StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                    If adoRecordset.EOF = False Then
                        adoRecordset.MoveNext
                    End If
                Next k
                For k = adoRecordset.RecordCount Mod 5 To 4
                    StrTemp4(k) = ""
                    StrTemp5(k) = ""
                Next k
                If iPrint > 10000 Then
                    PrintEnd1
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle1
                End If
                'If adoRecordset.EOF = False Then
                        PrintTotilEnd1
                'End If
                
            Else
                If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 <> 0 Then
                    adoRecordset.MoveFirst
                    For j = 1 To (adoRecordset.RecordCount \ 5)
                        For k = 0 To 4
                            StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                            StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                            If adoRecordset.EOF = False Then
                                adoRecordset.MoveNext
                            End If
                        Next k
                        If iPrint > 10000 Then
                            PrintEnd1
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle1
                        End If
                        'If adoRecordset.EOF = False Then
                              PrintTotilEnd1
                        'End If
                    Next j
                    For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                        StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                        If adoRecordset.EOF = False Then
                            adoRecordset.MoveNext
                        End If
                    Next k
                    For k = adoRecordset.RecordCount Mod 5 To 4
                        StrTemp4(k) = ""
                        StrTemp5(k) = ""
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd1
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle1
                    End If
                    'If adoRecordset.EOF = False Then
                        PrintTotilEnd1
                    'End If
                End If
            End If
        End If
    End If
    CheckOC
End If
CheckOC
PrintEnd1
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle1()
GetPleft1
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "收文明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "所　別：" & IIf(StrTemp6 = "ALL", StrTemp6, IIf(StrTemp6 = "1", "北所", IIf(StrTemp6 = "2", "中所", IIf(StrTemp6 = "3", "南所", IIf(StrTemp6 = "4", "其他", "")))))
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft1(1)
Printer.CurrentY = iPrint
Printer.Print "業務區"
Printer.CurrentX = PLeft1(2)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft1(3)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft1(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft1(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft1(6)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft1(7)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft1(8)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft1(9)
Printer.CurrentY = iPrint
Printer.Print "發文日"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintEnd1()
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(200, "-")
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint + 300
'Printer.Print "註：1.依所別小計且跳頁，再合計。"
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint + 600
'Printer.Print "        2.列印順序：所別＋業務區＋智權人員"
'iPrint = iPrint + 600
End Sub

Sub PrintDatil1()
For j = 1 To 9
    Printer.CurrentX = PLeft1(j)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(j)
Next j
iPrint = iPrint + 300

End Sub

Sub PrintTotil1()
'Add By Cheng 2002/07/28
Dim jj As Integer

For jj = 0 To 4
    Printer.CurrentX = 500 + (jj * 2000)
    Printer.CurrentY = iPrint
    Printer.Print StrConv(MidB(StrConv(StrTemp4(jj), vbFromUnicode), 1, 10), vbUnicode)
    Printer.CurrentX = 2300 + (jj * 2000) - Printer.TextWidth(StrTemp5(jj))
    Printer.CurrentY = iPrint
    Printer.Print StrTemp5(jj)
Next jj
iPrint = iPrint + 300

End Sub

Sub PrintTotilEnd1()
'Add By Cheng 2002/07/28
Dim jj As Integer

For jj = 0 To 4
    Printer.CurrentX = 500 + (jj * 2000)
    Printer.CurrentY = iPrint
    Printer.Print StrConv(MidB(StrConv(StrTemp4(jj), vbFromUnicode), 1, 10), vbUnicode)
    Printer.CurrentX = 2300 + (jj * 2000) - Printer.TextWidth(StrTemp5(jj))
    Printer.CurrentY = iPrint
    Printer.Print StrTemp5(jj)
Next jj
iPrint = iPrint + 300
End Sub

Sub GetPleft1()
Erase PLeft1
PLeft1(0) = 500
PLeft1(1) = 500
PLeft1(2) = 1500
PLeft1(3) = 2700
PLeft1(4) = 3800
PLeft1(5) = 6000
PLeft1(6) = 9500
PLeft1(7) = 11000
PLeft1(8) = 13000
PLeft1(9) = 14500
End Sub

Sub Process1()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050308_2 WHERE ID='" & strUserNum & "' "
pub_QL05 = pub_QL05 & ";" & Label5 & "2. 所別" 'Add By Sindy 2010/01/22
'系統類別
strSQL1 = ""
strSQL2 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/01/22
End If
'收文日
If Len(Trim(txt1(1))) <> 0 Then
   strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
   strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL2 = strSQL2 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/01/22
'申請國家
If Len(Trim(txt1(3))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
End If
If Len(Trim(txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
End If
If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/01/22
End If
'案件性質
If Len(Trim(txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND CP10>='" & txt1(5) & "' "
End If
If Len(Trim(txt1(6))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & txt1(6) & "' "
    strSQL2 = strSQL2 + " AND CP10<='" & txt1(6) & "' "
End If
If Len(Trim(txt1(5))) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/01/22
End If
'承辦人
If Len(txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP14='" & txt1(8) & "' "
    strSQL2 = strSQL2 + " AND CP14='" & txt1(8) & "' "
    pub_QL05 = pub_QL05 & ";" & Label7 & txt1(8) 'Add By Sindy 2010/01/22
End If
'所別
If Len(txt1(9)) <> 0 And Val(txt1(9)) <> 5 Then
    strSQL1 = strSQL1 + " AND S1.ST06='" & Trim(txt1(9)) & "' "
    strSQL2 = strSQL2 + " AND S1.ST06='" & Trim(txt1(9)) & "' "
    pub_QL05 = pub_QL05 & ";" & Label9 & txt1(9) & Label10 'Add By Sindy 2010/01/22
End If
'是否計算多國案件
If UCase(txt1(10)) = "Y" Then
    '2005/6/2 CANCEL BY SONIA
    'strSQL1 = strSQL1 + " AND CP21='Y' "
    'strSQL2 = strSQL2 + " AND CP21='Y' "
    '2005/6/2 END
Else
    strSQL1 = strSQL1 + " AND CP21 is null "
    strSQL2 = strSQL2 + " AND CP21 is null "
End If
'申請人
If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) <> 0 Then
    strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(11)) & "' AND PA26<='" & GetNewFagent(txt1(12)) & "') OR (PA27>='" & GetNewFagent(txt1(11)) & "' AND PA27<='" & GetNewFagent(txt1(12)) & "') OR (PA28>='" & GetNewFagent(txt1(11)) & "' AND PA28<='" & GetNewFagent(txt1(12)) & "') OR (PA29>='" & GetNewFagent(txt1(11)) & "' AND PA29<='" & GetNewFagent(txt1(12)) & "') OR (PA30>='" & GetNewFagent(txt1(11)) & "' AND PA30<='" & GetNewFagent(txt1(12)) & "')) "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(11)) & "' AND SP08<='" & GetNewFagent(txt1(12)) & "') OR (SP58<='" & GetNewFagent(txt1(11)) & "' AND SP58<='" & GetNewFagent(txt1(12)) & "') OR (SP59>='" & GetNewFagent(txt1(11)) & "' AND SP59<='" & GetNewFagent(txt1(12)) & "')) "
Else
    If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) = 0 Then
        strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(11)) & "' OR PA27>='" & GetNewFagent(txt1(11)) & "' OR PA28>='" & GetNewFagent(txt1(11)) & "' OR PA29>='" & GetNewFagent(txt1(11)) & "' OR PA30>='" & GetNewFagent(txt1(11)) & "') "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(11)) & "' OR SP58>='" & GetNewFagent(txt1(11)) & "' OR SP59>='" & GetNewFagent(txt1(11)) & "') "
    Else
        If Len(Trim(txt1(11))) = 0 And Len(Trim(txt1(12))) <> 0 Then
            strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(12)) & "' OR PA27<='" & GetNewFagent(txt1(12)) & "' OR PA28<='" & GetNewFagent(txt1(12)) & "' OR PA29<='" & GetNewFagent(txt1(12)) & "' OR PA30<='" & GetNewFagent(txt1(12)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(12)) & "' OR SP58<='" & GetNewFagent(txt1(12)) & "' OR SP59<='" & GetNewFagent(txt1(12)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label13 & txt1(11) & "-" & txt1(12) 'Add By Sindy 2010/01/22
End If
'代理人
If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) <> 0 Then
    strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(13)) & "' AND PA75<='" & GetNewFagent(txt1(14)) & "' "
    strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(13)) & "' AND SP26<='" & GetNewFagent(txt1(14)) & "' "
Else
    If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) = 0 Then
        strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(13)) & "' "
        strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(13)) & "' "
    Else
        If Len(Trim(txt1(13))) = 0 And Len(Trim(txt1(14))) <> 0 Then
            strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(14)) & "' "
            strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(14)) & "' "
        End If
    End If
End If
If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/01/22
End If
'組合
'910318  nick
'strSQL = "select decode(S1.ST06,'1','北所','2','中所','3','南所','4','其他') AS A,S1.ST03 AS B,S1.ST02 AS C," & SQLDate("CP05") & " AS D,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),nvl(decode(pa09,'000',CPM03,CPM04),cp10),S2.ST02," & SQLDate("CP27") & ",'" & strUserNum & "' FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,PATENT WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND pa09=na01(+) AND cp13=S1.ST01(+) AND CP14=s2.ST01(+)  " & strSQL1
'strSQL = strSQL + " union all select decode(S1.ST06,'1','北所','2','中所','3','南所','4','其他') AS A,S1.ST03 AS B,S1.ST02 AS C," & SQLDate("CP05") & " AS D,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(NA03,NA04),nvl(decode(sp09,'000',CPM03,CPM04),cp10),S2.ST02," & SQLDate("CP27") & ",'" & strUserNum & "' FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) AND cp13=S1.ST01(+) AND CP14=s2.ST01(+)  " & strSQL2
strSql = "select S1.ST06 AS A,cp12 AS B,cp13 AS C," & SQLDate("CP05") & " AS D,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),cp10,S2.ST02," & SQLDate("CP27") & ",'" & strUserNum & "',cp01,pa09 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,PATENT WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND pa09=na01(+) AND cp13=S1.ST01(+) AND CP14=s2.ST01(+)  " & strSQL1
strSql = strSql + " union all select S1.ST06 AS A,cp12 AS B,cp13 AS C," & SQLDate("CP05") & " AS D,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(NA03,NA04),cp10,S2.ST02," & SQLDate("CP27") & ",'" & strUserNum & "',cp01,sp09 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) AND cp13=S1.ST01(+) AND CP14=s2.ST01(+)  " & strSQL2
'strSQL = strSQL + " ORDER BY A,B,C,D "
cnnConnection.Execute "insert into r050308_2 " & strSql
strSql = "select * from r050308_2 where id='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
Else
   InsertQueryLog (0) 'Add By Sindy 2010/01/22
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
PrintData1
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050308 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/12
   Select Case Index
   Case 7 '列印順序
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 9 '智權人員所別
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 10 '是否計算多國案件
      If KeyAscii <> 89 And KeyAscii <> 8 Then
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
Case 2, 4, 6
   'Modify By Cheng 2002/09/12
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
        txt1(Index - 1).SetFocus
        txt1_GotFocus (Index - 1)
        Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case 7
   If Me.txt1(7).Text <> "" Then
     Select Case txt1(7)
     Case "1", "2"
     Case Else
         s = MsgBox("列印順序只能輸入 1 或 2 !!", , "USER 輸入錯誤")
         txt1(7).SetFocus
         txt1(7).SelStart = 0
         txt1(7).SelLength = Len(txt1(7))
         Exit Sub
     End Select
   End If
Case 8
     lbl1 = GetPrjSales(txt1(8))
     'Add By Cheng 2002/09/26
     If Me.txt1(8).Text <> "" Then
         If Me.txt1(8).Text = Me.lbl1.Caption Then
            Me.lbl1.Caption = ""
            Me.txt1(8).SetFocus
            txt1_GotFocus 8
            Exit Sub
         End If
     End If
Case 9
   'Modify By Cheng 2002/09/26
   If Me.txt1(9).Text <> "" Then
     Select Case txt1(9)
     Case "1", "2", "3", "4", "5"
     Case Else
         s = MsgBox("所別只能輸入 1 或 2 或 3 或 4 或 5 !!", , "USER 輸入錯誤")
         txt1(9).SetFocus
         txt1(9).SelStart = 0
         txt1(9).SelLength = Len(txt1(9))
         Exit Sub
     End Select
   End If
Case 10
   'Modify By Cheng 2002/09/26
   If Me.txt1(10).Text <> "" Then
     Select Case txt1(10)
     Case "Y", ""
     Case Else
         s = MsgBox("是否計算多國案件只能輸入 Y 或 空白 !!", , "USER 輸入錯誤")
         txt1(10).SetFocus
         txt1(10).SelStart = 0
         txt1(10).SelLength = Len(txt1(10))
         Exit Sub
     End Select
   End If
Case 12
   'Modify By Cheng 2002/09/12
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
Case 14
   'Modify By Cheng 2002/09/26
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
Case Else

End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '收文日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub
