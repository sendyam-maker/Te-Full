VERSION 5.00
Begin VB.Form frm040304 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限管制表"
   ClientHeight    =   4152
   ClientLeft      =   996
   ClientTop       =   2148
   ClientWidth     =   3636
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4152
   ScaleWidth      =   3636
   Begin VB.CheckBox Check1 
      Caption         =   "含FMP外專管制期限"
      Height          =   285
      Left            =   180
      TabIndex        =   30
      Top             =   3750
      Width           =   3345
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   13
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1950
      Width           =   360
   End
   Begin VB.TextBox txtPA46 
      Height          =   264
      Left            =   1710
      TabIndex        =   14
      Top             =   3435
      Width           =   360
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   350
      Index           =   1
      Left            =   2796
      TabIndex        =   16
      Top             =   50
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   1992
      TabIndex        =   15
      Top             =   50
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   2256
      MaxLength       =   9
      TabIndex        =   13
      Top             =   3105
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1056
      MaxLength       =   9
      TabIndex        =   12
      Top             =   3105
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   2256
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2820
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1056
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2820
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2508
      Width           =   525
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2235
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1332
      Width           =   360
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1644
      Width           =   360
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1032
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1032
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   2
      Top             =   744
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   1
      Top             =   744
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1035
      TabIndex        =   0
      Top             =   432
      Width           =   2460
   End
   Begin VB.OptionButton Option1 
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   168
      TabIndex        =   25
      Top             =   1044
      Width           =   1305
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   168
      TabIndex        =   24
      Top             =   756
      Value           =   -1  'True
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列印對象：　   （1.非智權部同仁 2.全部）"
      Height          =   180
      Index           =   10
      Left            =   168
      TabIndex        =   29
      Top             =   1992
      Width           =   3330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PCT進入國家階段：　   （Y：國家階段）"
      Height          =   180
      Index           =   9
      Left            =   165
      TabIndex        =   28
      Top             =   3480
      Width           =   3270
   End
   Begin VB.Line Line2 
      X1              =   2250
      X2              =   2355
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line4 
      X1              =   2085
      X2              =   2220
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   2100
      X2              =   2235
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line Line1 
      X1              =   2268
      X2              =   2373
      Y1              =   864
      Y2              =   864
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   1620
      TabIndex        =   27
      Top             =   2550
      Width           =   1920
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   2235
      TabIndex        =   26
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   6
      Left            =   165
      TabIndex        =   23
      Top             =   3150
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   165
      TabIndex        =   22
      Top             =   2865
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   4
      Left            =   165
      TabIndex        =   21
      Top             =   2550
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員/承辦人："
      Height          =   180
      Index           =   3
      Left            =   165
      TabIndex        =   20
      Top             =   2277
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列印別：　　     (1.未收文 2.已收文 3.全部)"
      Height          =   180
      Index           =   2
      Left            =   168
      TabIndex        =   19
      Top             =   1374
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "管制對象：　     (1.承辦人 2.智權人員)"
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   18
      Top             =   1680
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   168
      TabIndex        =   17
      Top             =   468
      Width           =   936
   End
End
Attribute VB_Name = "frm040304"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim s As Integer, strSql As String, iPrint As Integer, Page As Integer, strSQL1 As String, strSQL2 As String
Dim strTemp2 As Variant, StrTest2 As String, i As Integer, j As Integer, strTemp1 As Variant, StrTest As String
Dim strTemp(0 To 20) As String, StrYear1 As String, StrYear2 As String, PLeft(0 To 12) As Integer
Dim Area As String
'Add By Cheng 2002/09/11
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim Iman As String        '20140217ADD By eric 非智權同仁之期限管制表,改為每個人一張分開跑

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     'Add By Cheng 2002/09/11
      blnClkSure = False
     
     If (Len(Txt1(9)) = 0 And Len(Txt1(10)) <> 0) Or (Len(Txt1(9)) <> 0 And Len(Txt1(10)) = 0) Then
        s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
        Txt1(9).SetFocus
        txt1_GotFocus (9)
        Exit Sub
     End If
     If (Len(Txt1(11)) = 0 And Len(Txt1(12)) <> 0) Or (Len(Txt1(11)) <> 0 And Len(Txt1(12)) = 0) Then
        s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
        Txt1(11).SetFocus
        txt1_GotFocus (11)
        Exit Sub
     End If
     If Len(Txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白", , "USER 輸入錯誤")
        Txt1(0).SetFocus
        Exit Sub
     Else
         'Add By Cheng 2002/03/19
         '本所期限
         If Me.Option1(0).Value Then
            If PUB_CheckKeyInDate(Me.Txt1(1)) = -1 Then
               Me.Txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Txt1(2)) = -1 Then
               Me.Txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            'Add By Cheng 2002/09/11
            If Me.Txt1(1).Text <> "" And Me.Txt1(2).Text <> "" Then
               If Val(Me.Txt1(1).Text) > Val(Me.Txt1(2).Text) Then
                  MsgBox "本所期限範圍輸入錯誤", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.Txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
            
         '法定期限
         Else
            If PUB_CheckKeyInDate(Me.Txt1(3)) = -1 Then
               Me.Txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Txt1(4)) = -1 Then
               Me.Txt1(4).SetFocus
               txt1_GotFocus 4
               Exit Sub
            End If
            'Add By Cheng 2002/09/11
            If Me.Txt1(3).Text <> "" And Me.Txt1(4).Text <> "" Then
               If Val(Me.Txt1(3).Text) > Val(Me.Txt1(4).Text) Then
                  MsgBox "法定期限範圍輸入錯誤", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.Txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
         
         End If
        
        If Len(Txt1(5)) = 0 Then
            s = MsgBox("管制對象不可空白", , "USER 輸入錯誤")
            Txt1(5).SetFocus
            Exit Sub
        Else
            StrTest = StrTest2
            strTemp1 = Split(UCase(StrTest), ",")
            strTemp2 = Split(UCase(Txt1(0)), ",")
            For i = 0 To UBound(strTemp1)
                s = 0
                For j = 0 To UBound(strTemp2)
                    If strTemp2(j) = strTemp1(i) Then
                        s = 1
                        Exit For
                    End If
                Next j
                If s = 0 Then
                    StrTest = Replace(StrTest, strTemp1(i), "")
                End If
            Next i
            If Len(Txt1(6)) = 0 Then
                s = MsgBox("列印別不可空白", , "USER 輸入錯誤")
                Txt1(6).SetFocus
                Exit Sub
            Else
               If Txt1(8) <> "" Then
                  'Add By Cheng 2002/09/11
                  lbl1(1) = GetPrjState6HM("P", Txt1(8))
                  If lbl1(1) = "" Then
                     MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
                     Me.Txt1(8).SetFocus
                     txt1_GotFocus 8
                     Exit Sub
                  End If
                End If
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                Select Case Val(Txt1(6))
                Case 1
                      StrMenu              '未收文
                Case 2
                      StrMenu1             '已收文
                Case 3
                      StrMenu              '全部
                      StrMenu1
                End Select
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

Sub StrMenu()                       '未收文智權人員
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R040304_1 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""

ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/30 清除查詢印表記錄檔欄位
If Txt1(6) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(2), 4) & "1.未收文" 'Add By Sindy 2010/11/30
ElseIf Txt1(6) = "2" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(2), 4) & "2.已收文" 'Add By Sindy 2010/11/30
ElseIf Txt1(6) = "3" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(2), 4) & "3.全部" 'Add By Sindy 2010/11/30
End If
pub_QL05 = pub_QL05 & ";內專智權人員期限管制表" 'Add By Sindy 2010/11/30

'本所期限
If Option1(0).Value = True Then
   '91.11.4 modify by sonia
    'strSQL1 = strSQL1 + " AND NP06 IS NULL "
    'Modify by Morgan 2009/7/13 +995,996
    'Modify by Morgan 2010/2/1 +994
    '2010/10/14 MODIFY BY SONIA 改用常數strNpSqlOfNoSalesDuty
    'strSQL1 = strSQL1 + " AND NP06 IS NULL AND NP07 NOT IN ('997','998','994','995','996','411','1204') "
    strSQL1 = strSQL1 + " AND NP06 IS NULL " & strNpSqlOfNoSalesDuty
   '91.11.4 end
   
   pub_QL05 = pub_QL05 & ";" & Option1(0).Caption 'Add By Sindy 2010/11/30
   If Len(Trim(Txt1(1))) <> 0 Then
      strSQL1 = strSQL1 & " AND NP08>=" & ChangeTStringToWString(Txt1(1)) & " "
      pub_QL05 = pub_QL05 & Txt1(1) 'Add By Sindy 2010/11/30
   End If
   
   If Len(Trim(Txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND NP08<=" & ChangeTStringToWString(Txt1(2)) & " "
      pub_QL05 = pub_QL05 & "-" & Txt1(2) 'Add By Sindy 2010/11/30
   'Add By Cheng 2002/03/19
   Else
      If Len(Trim(Txt1(1))) <> 0 Then
         strSQL1 = strSQL1 & " AND NP08<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
         pub_QL05 = pub_QL05 & "-" & (ServerDate - 19110000) 'Add By Sindy 2010/11/30
      End If
   End If
'法定期限
Else
   '91.11.4 modify by sonia
    'strSQL1 = strSQL1 + " AND NP06 IS NULL "
    'Modify by Morgan 2009/7/13 +995,996
    'Modify by Morgan 2010/2/1 +994
    'strSQL1 = strSQL1 + " AND NP06 IS NULL AND NP07 NOT IN ('997','998','994','995','996','411','1204') "
    '2010/10/14 MODIFY BY SONIA 改用常數strNpSqlOfNoSalesDuty
    strSQL1 = strSQL1 + " AND NP06 IS NULL " & strNpSqlOfNoSalesDuty
   '91.11.4 end
   'Modify By Cheng 2002/03/19
'    If Len(Trim(txt1(1))) <> 0 Then
'      strSQL1 = strSQL1 & " AND NP09>=" & ChangeTStringToWString(txt1(1)) & " "
    pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/11/30
    If Len(Trim(Txt1(3))) <> 0 Then
      strSQL1 = strSQL1 & " AND NP09>=" & ChangeTStringToWString(Txt1(3)) & " "
      pub_QL05 = pub_QL05 & Txt1(3) 'Add By Sindy 2010/11/30
    End If
   'Modify By Cheng
'    If Len(Trim(txt1(2))) <> 0 Then
'      strSQL1 = strSQL1 & " AND NP09<=" & ChangeTStringToWString(txt1(2)) & " "
    If Len(Trim(Txt1(4))) <> 0 Then
      strSQL1 = strSQL1 & " AND NP09<=" & ChangeTStringToWString(Txt1(4)) & " "
      pub_QL05 = pub_QL05 & "-" & Txt1(4) 'Add By Sindy 2010/11/30
   'Add By Cheng 2002/03/19
   Else
      If Len(Trim(Txt1(3))) <> 0 Then
         strSQL1 = strSQL1 & " AND NP09<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
         pub_QL05 = pub_QL05 & "-" & (ServerDate - 19110000) 'Add By Sindy 2010/11/30
      End If
    End If
End If

If Txt1(5) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 5) & "1.承辦人" 'Add By Sindy 2010/11/30
ElseIf Txt1(5) = "2" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 5) & "2.智權人員" 'Add By Sindy 2010/11/30
End If
If Len(Txt1(5)) <> 0 And Len(Txt1(7)) <> 0 Then
'Modify by Morgan 2006/10/17
'    If Val(txt1(6)) = 1 Then
'        strSQL1 = strSQL1 + " AND NP10='" & txt1(7) & "' "
'    Else
'        strSQL1 = strSQL1 + " AND CP14='" & txt1(7) & "' "
'    End If
    strSQL1 = strSQL1 + " AND NP10='" & Txt1(7) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(3) & Txt1(7) & lbl1(0) 'Add By Sindy 2010/11/30
End If

'92.1.10 CANCEL BY SONIA 未收文資料無此限制
'Add By Cheng 2002/12/12
'除非有指定承辦人否則, 承辦人為陳玲玲與莊敏惠的資料不印
'If Val(txt1(6)) = 1 Then
'    If Me.txt1(7).Text = "" Then
'        stCon = stCon + " AND CP14<>'81002' And CP14<>'73017' "
'    End If
'End If
'92.1.10 END
If Len(Txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " AND np07='" & Txt1(8) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(8) & lbl1(1) 'Add By Sindy 2010/11/30
End If
strSQL2 = strSQL1
If Len(StrTest) <> 0 Then
   strSQL1 = strSQL1 + " AND np02 in (" & SQLGrpStr(StrTest, 1) & ") "
   strSQL2 = strSQL2 + " AND np02 in (" & SQLGrpStr(StrTest, 5) & ") "
End If
If Len(Txt1(9)) <> 0 And Len(Txt1(10)) <> 0 Then
    strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(Txt1(9)) & "' AND PA26<='" & GetNewFagent(Txt1(10)) & "') OR (PA27>='" & GetNewFagent(Txt1(9)) & "' AND PA27<='" & GetNewFagent(Txt1(10)) & "') OR (PA28>='" & GetNewFagent(Txt1(9)) & "' AND PA28<='" & GetNewFagent(Txt1(10)) & "') OR (PA29>='" & GetNewFagent(Txt1(9)) & "' AND PA29<='" & GetNewFagent(Txt1(10)) & "') OR (PA30>='" & GetNewFagent(Txt1(9)) & "' AND PA30<='" & GetNewFagent(Txt1(10)) & "'))"
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(Txt1(9)) & "' AND SP05<='" & GetNewFagent(Txt1(10)) & "') OR (SP58>='" & GetNewFagent(Txt1(9)) & "' AND SP58<='" & GetNewFagent(Txt1(10)) & "') OR (SP59>='" & GetNewFagent(Txt1(9)) & "' AND SP59<='" & GetNewFagent(Txt1(10)) & "')) "
Else
    If Len(Txt1(9)) <> 0 Then
        strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(Txt1(9)) & "') OR (PA27>='" & GetNewFagent(Txt1(9)) & "') OR (PA28>='" & GetNewFagent(Txt1(9)) & "') OR (PA29>='" & GetNewFagent(Txt1(9)) & "') OR (PA30>='" & GetNewFagent(Txt1(9)) & "'))"
        strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(Txt1(9)) & "') OR (SP58>='" & GetNewFagent(Txt1(9)) & "') OR (SP59>='" & GetNewFagent(Txt1(9)) & "')) "
    Else
        If Len(Txt1(10)) <> 0 Then
            strSQL1 = strSQL1 + " AND ((PA26<='" & GetNewFagent(Txt1(10)) & "') OR (PA27<='" & GetNewFagent(Txt1(10)) & "') OR (PA28<='" & GetNewFagent(Txt1(10)) & "') OR (PA29<='" & GetNewFagent(Txt1(10)) & "') OR (PA30<='" & GetNewFagent(Txt1(10)) & "'))"
            strSQL2 = strSQL2 + " AND ((SP05<='" & GetNewFagent(Txt1(10)) & "') OR (SP58<='" & GetNewFagent(Txt1(10)) & "') OR (SP59<='" & GetNewFagent(Txt1(10)) & "')) "
        End If
    End If
End If
If Len(Txt1(9)) <> 0 Or Len(Txt1(10)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & Txt1(9) & "-" & Txt1(10) 'Add By Sindy 2010/11/30
End If
If Len(Txt1(11)) <> 0 And Len(Txt1(12)) <> 0 Then
    strSQL1 = strSQL1 + " AND (PA75>='" & GetNewFagent(Txt1(11)) & "' AND PA75<='" & GetNewFagent(Txt1(12)) & "')"
    strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(Txt1(11)) & "' AND SP26<='" & GetNewFagent(Txt1(12)) & "')"
Else
    If Len(Txt1(11)) <> 0 Then
        strSQL1 = strSQL1 + " AND (PA75>='" & GetNewFagent(Txt1(11)) & "')"
        strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(Txt1(11)) & "') "
    Else
        If Len(Txt1(12)) <> 0 Then
            strSQL1 = strSQL1 + " AND (PA75<='" & GetNewFagent(Txt1(12)) & "')"
            strSQL2 = strSQL2 + " AND (SP26<='" & GetNewFagent(Txt1(12)) & "') "
        End If
    End If
End If
If Len(Txt1(11)) <> 0 Or Len(Txt1(12)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(11) & "-" & Txt1(12) 'Add By Sindy 2010/11/30
End If

'Add by Morgan 2005/2/14
If txtPA46 = "Y" Then
   strSQL1 = strSQL1 & " And PA09<>'056' AND PA46='Y' "
   pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 10) & txtPA46 'Add By Sindy 2010/11/30
End If

'Add by Morgan 2006/5/30
If Txt1(5) = "2" Then
   If Txt1(13) = "1" Then
      strSQL1 = strSQL1 & " And SUBSTR(S1.ST15,1,1)<>'S'"
      pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 5) & "1.非智權部同仁" 'Add By Sindy 2010/11/30
   ElseIf Txt1(13) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 5) & "2.全部" 'Add By Sindy 2010/11/30
   End If
End If

CheckOC
'Modify By Cheng 2002/07/24
'已閉卷的資料不印
'strSQL = "select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),decode(pa09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),PA72,PA08,NP07,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST03=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND pa09=na01(+) " & strSQL1
'strSQL = strSQL + " union all select A0902,S1.ST02 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),decode(sp09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),' ',' ',CP10,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST03=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND sp09=na01(+) " & strSQL2
'91.11.4 MODIFY BY SONIA NP10->ST15 NOT ST03
'strSQL = "select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),decode(pa09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),PA72,PA08,NP07,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST03=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND pa09=na01(+) AND PA57 IS NULL " & strSQL1
'strSQL = strSQL + " union all select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),decode(sp09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),' ',' ',CP10,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST03=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND sp09=na01(+) AND SP15 IS NULL " & strSQL2
'2010/10/14 modify by sonia FMP案件依外專規則抓智權人員
'strSql = "select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),decode(pa09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),PA72,PA08,NP07,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND pa09=na01(+) AND PA57 IS NULL " & strSQL1
'strSql = strSql + " union all select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),decode(sp09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),' ',' ',CP10,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND sp09=na01(+) AND SP15 IS NULL " & strSQL2
strSql = "select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(pa09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),PA72,PA08,NP07,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND pa09=na01(+) AND PA57 IS NULL AND SUBSTR(S1.ST15,1,1)<>'F' " & strSQL1
strSql = strSql + " union all select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(sp09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),' ',' ',CP10,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND sp09=na01(+) AND SP15 IS NULL AND SUBSTR(S1.ST15,1,1)<>'F' " & strSQL2

'Addec by Morgan 2012/5/22 FMP案改可選擇
If Check1.Value = 1 Then
   pub_QL05 = pub_QL05 & ";" & Check1.Caption
'end 2012/5/22
'   strSql = strSql + " union all select A0902,n2.na51 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(pa09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(n1.NA03,n1.NA04),PA72,PA08,NP07,n2.Na51,n1.NA21,n1.NA23,n1.NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION n1,fagent,nation n2 WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND decode(SUBSTR(Pa75,9,1),null,'0',substr(pa75,9,1))=FA02(+) and fa10=n2.na01(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND pa09=N1.na01(+) AND PA57 IS NULL AND SUBSTR(S1.ST15,1,1)='F' " & strSQL1
'   strSql = strSql + " union all select A0902,n2.na51 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(sp09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(n1.NA03,n1.NA04),' ',' ',CP10,n2.Na51,n1.NA21,n1.NA23,n1.NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION n1,fagent,nation n2 WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND SUBSTR(sP26,1,8)=FA01(+) AND decode(SUBSTR(sP26,9,1),null,'0',substr(sp26,9,1))=FA02(+) and fa10=n2.na01(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND sp09=N1.na01(+) AND SP15 IS NULL AND SUBSTR(S1.ST15,1,1)='F' " & strSQL2
'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
 '代理人Y51333010=Pub_GetSpecMan("北京銀龍FCP案承辦業務")
   Dim midStr As String, midStr2 As String
   'Modified by Lydia 2024/05/27 改成用Y編號+日文承辦業務
   'midStr = Pub_GetSpecMan("北京銀龍FCP案承辦業務")
   'strSql = strSql + " union all select A0902,decode(pa75,'Y51333010','" & midStr & "',n2.na51) AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(pa09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(n1.NA03,n1.NA04),PA72,PA08,NP07,decode(pa75,'Y51333010','" & midStr & "',n2.na51) Na51,n1.NA21,n1.NA23,n1.NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION n1,fagent,nation n2 WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND decode(SUBSTR(Pa75,9,1),null,'0',substr(pa75,9,1))=FA02(+) and fa10=n2.na01(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND pa09=N1.na01(+) AND PA57 IS NULL AND SUBSTR(S1.ST15,1,1)='F' " & strSQL1
   'strExc(0) = strSql + " union all select A0902,decode(sp26,'Y51333010','" & midStr & "',n2.na51) AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(sp09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(n1.NA03,n1.NA04),' ',' ',CP10,decode(sp26,'Y51333010','" & midStr & "',n2.na51) Na51 "
   midStr = Pub_GetSpecFCP
   strSql = strSql + " union all select A0902,decode(pa75," & midStr & ",n2.na51) AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(pa09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(n1.NA03,n1.NA04),PA72,PA08,NP07,decode(pa75," & midStr & ",n2.na51) Na51,n1.NA21,n1.NA23,n1.NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION n1,fagent,nation n2 WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND decode(SUBSTR(Pa75,9,1),null,'0',substr(pa75,9,1))=FA02(+) and fa10=n2.na01(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND pa09=N1.na01(+) AND PA57 IS NULL AND SUBSTR(S1.ST15,1,1)='F' " & strSQL1
   strExc(0) = strSql + " union all select A0902,decode(sp26," & midStr & ",n2.na51) AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(sp09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(n1.NA03,n1.NA04),' ',' ',CP10,decode(sp26," & midStr & ",n2.na51) Na51 "
   'end 2024/05/27
   strSql = strExc(0) + ",n1.NA21,n1.NA23,n1.NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION n1,fagent,nation n2 WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND SUBSTR(sP26,1,8)=FA01(+) AND decode(SUBSTR(sP26,9,1),null,'0',substr(sp26,9,1))=FA02(+) and fa10=n2.na01(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND sp09=N1.na01(+) AND SP15 IS NULL AND SUBSTR(S1.ST15,1,1)='F' " & strSQL2
End If 'Added by Morgan 2012/5/22

'2010/10/14 END
'91.11.4 END
strSql = strSql + " ORDER BY A0902,B,CP05,A "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/30
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 20
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            StrYear2 = ""
            If Val(strTemp(14)) = 605 Or Val(strTemp(14)) = 606 Or Val(strTemp(14)) = 607 Then
            If Len(Trim(strTemp(12))) <> 0 Then
                strTemp1 = Split(strTemp(12), ",")
                'Modify By Cheng 2002/07/24
                '取得起始繳費年度
                StrYear1 = strTemp1(LBound(strTemp1))
                For i = LBound(strTemp1) To UBound(strTemp1)
                  If Val(strTemp1(i)) > Val(StrYear1) Then
                     StrYear1 = strTemp1(i)
                  End If
                Next i
'                StrYear1 = strTemp1(UBound(strTemp1))
                Select Case Val(strTemp(13))
                Case 1
'                     strTemp1 = Split(strTemp(16))
                     strTemp1 = Split(strTemp(16), ",")
                     For i = 0 To UBound(strTemp1)
                        'Modify By Cheng 2002/07/24
'                        If Trim(strTemp1(i)) = Trim(StrYear1) Then
                        If Val(Trim(strTemp1(i))) > Val(Trim(StrYear1)) Then
'                            If i <> UBound(strTemp1) Then
'                                StrYear2 = strTemp1(i + 1)
'                            Else
'                                StrYear2 = ""
'                            End If
                             StrYear2 = strTemp1(i)
                             Exit For
                        Else
                           StrYear2 = ""
                        End If
                     Next i
                Case 2
                     'Modify By Cheng 2002/07/24
'                     strTemp1 = Split(strTemp(17))
                     strTemp1 = Split(strTemp(17), ",")
                     For i = 0 To UBound(strTemp1)
'                        If Trim(strTemp1(i)) = Trim(StrYear1) Then
                        If Val(Trim(strTemp1(i))) > Val(Trim(StrYear1)) Then
'                            If i <> UBound(strTemp1) Then
'                                StrYear2 = strTemp1(i + 1)
'                            Else
'                                StrYear2 = ""
'                            End If
                           StrYear2 = strTemp1(i)
                           Exit For
                        Else
                           StrYear2 = ""
                        End If
                     Next i
                Case 3
                     'Modify By Cheng 2002/07/24
'                     strTemp1 = Split(strTemp(18))
                     strTemp1 = Split(strTemp(18), ",")
                     For i = 0 To UBound(strTemp1)
'                        If Trim(strTemp1(i)) = Trim(StrYear1) Then
                        If Val(Trim(strTemp1(i))) > Val(Trim(StrYear1)) Then
'                            If i <> UBound(strTemp1) Then
'                                StrYear2 = strTemp1(i + 1)
'                            Else
'                                StrYear2 = ""
'                            End If
                           StrYear2 = strTemp1(i)
                           Exit For
                        Else
                           StrYear2 = ""
                        End If
                     Next i
                End Select
            End If
            Else
                strTemp(12) = ""
            End If
            
            If strTemp(3) < Format(Now, "YYYYMMDD") Then
                strTemp(3) = "*" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
            Else
                If strTemp(3) = Format(Now, "YYYYMMDD") Then
                    strTemp(3) = "V" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                Else
                    If UCase(Mid(strTemp(19), 1, 1)) = "C" And Trim(strTemp(20)) = "" Then
                        strTemp(3) = "#" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                    Else
                        'add by nick 年費駐記 2004/07/13
                        If Val(strTemp(14)) = 605 And (SystemNumber(strTemp(4), 1) = "FCP" Or SystemNumber(strTemp(4), 1) = "P") Then
                            strSql = "select pa14 from patent where pa01='" & SystemNumber(strTemp(4), 1) & "' and pa02='" & _
                                         SystemNumber(strTemp(4), 2) & "' and pa03='" & SystemNumber(strTemp(4), 3) & "' and pa04='" & _
                                         SystemNumber(strTemp(4), 4) & "' and pa09='000' "
                            CheckOC3
                            AdoRecordSet3.CursorLocation = adUseClient
                            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                            If AdoRecordSet3.RecordCount <> 0 Then
                                If CheckStr(AdoRecordSet3.Fields(0).Value) = "" Then
                                       strTemp(3) = "!" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                                Else
                                       strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                                End If
                            Else
                                strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                            End If
                        Else
                            strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                        End If
                        'add end 2004/07/13
                    End If
                End If
            End If
            strTemp(2) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
            strTemp(10) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(10)))
            strTemp(12) = StrYear2
            strSql = "INSERT INTO R040304_1 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            DoEvents
            .MoveNext
        Loop
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/30
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
CheckOC
PrintData2

'Add by Morgan 2005/8/26
'若條件為本所期限,未收文時 加印1.申復or修正有修改的期限表2.15日內收文,期限為本次期限前推10日的期限表
If Option1(0).Value = True And Txt1(6) = "1" And Len(Trim(Txt1(1))) <> 0 Then
   StrMenu3
End If

'91.7.22 modify by sonia
'ShowPrintOk
   s = MsgBox("智權人員管制表列印完成!!", , "列印成功")
'91.7.22 end
Screen.MousePointer = vbDefault
End Sub

'Add by Morgan 2005/8/26
'選本所期限且未收文才出
'1.15日內收文,期限為本次期限前推10日的期限表
'2.申復or修正有修改的期限表
Sub StrMenu3()
Dim stCon As String
   
   cnnConnection.Execute "DELETE FROM R040304_1 WHERE ID='" & strUserNum & "' "
   strSQL1 = ""
   strSQL2 = ""
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/30 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";內專智權人員期限管制表-(前期限補通知)" 'Add By Sindy 2010/11/30
   
   'Modify by Morgan 2009/7/13 +995,996
   'Modify by Morgan 2010/2/1 +994
   '2010/10/14 MODIFY BY SONIA 改用常數strNpSqlOfNoSalesDuty
   'strSQL1 = strSQL1 + " AND NP06 IS NULL AND NP07 NOT IN ('997','998','994','995','996','411','1204') "
   strSQL1 = strSQL1 + " AND NP06 IS NULL " & strNpSqlOfNoSalesDuty
   '2010/10/14 END
   strSQL1 = strSQL1 & " AND NP08<" & TransDate(Txt1(1), 2)
   strSQL1 = strSQL1 & " AND CP05>=" & CompDate(2, -15, strSrvDate(2))
   strSQL1 = strSQL1 & " AND NP08>=" & CompDate(2, -10, TransDate(Txt1(1), 2))
   pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Txt1(1) & ";15日內收文,期限為本次期限前推10日的期限表" 'Add By Sindy 2010/11/30
   
   If Txt1(5) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 5) & "1.承辦人" 'Add By Sindy 2010/11/30
   ElseIf Txt1(5) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 5) & "2.智權人員" 'Add By Sindy 2010/11/30
   End If
   If Len(Txt1(5)) <> 0 And Len(Txt1(7)) <> 0 Then
      If Val(Txt1(6)) = 1 Then
         strSQL1 = strSQL1 + " AND NP10='" & Txt1(7) & "' "
         stCon = stCon & " AND NP10='" & Txt1(7) & "' "
      Else
         strSQL1 = strSQL1 + " AND CP14='" & Txt1(7) & "' "
         stCon = stCon & " AND CP14='" & Txt1(7) & "' "
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(3) & Txt1(7) & lbl1(0) 'Add By Sindy 2010/11/30
   End If
   If Len(Txt1(8)) <> 0 Then
       strSQL1 = strSQL1 + " AND np07='" & Txt1(8) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(8) & lbl1(1) 'Add By Sindy 2010/11/30
   End If
   
   strSQL2 = strSQL1
   
   If Len(StrTest) <> 0 Then
      strSQL1 = strSQL1 + " AND np02 in (" & SQLGrpStr(StrTest, 1) & ") "
      strSQL2 = strSQL2 + " AND np02 in (" & SQLGrpStr(StrTest, 5) & ") "
      stCon = stCon + " AND np02 in (" & SQLGrpStr(StrTest, 1) & ") "
   End If
   If Len(Txt1(9)) <> 0 And Len(Txt1(10)) <> 0 Then
       strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(Txt1(9)) & "' AND PA26<='" & GetNewFagent(Txt1(10)) & "') OR (PA27>='" & GetNewFagent(Txt1(9)) & "' AND PA27<='" & GetNewFagent(Txt1(10)) & "') OR (PA28>='" & GetNewFagent(Txt1(9)) & "' AND PA28<='" & GetNewFagent(Txt1(10)) & "') OR (PA29>='" & GetNewFagent(Txt1(9)) & "' AND PA29<='" & GetNewFagent(Txt1(10)) & "') OR (PA30>='" & GetNewFagent(Txt1(9)) & "' AND PA30<='" & GetNewFagent(Txt1(10)) & "'))"
       strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(Txt1(9)) & "' AND SP05<='" & GetNewFagent(Txt1(10)) & "') OR (SP58>='" & GetNewFagent(Txt1(9)) & "' AND SP58<='" & GetNewFagent(Txt1(10)) & "') OR (SP59>='" & GetNewFagent(Txt1(9)) & "' AND SP59<='" & GetNewFagent(Txt1(10)) & "')) "
   Else
       If Len(Txt1(9)) <> 0 Then
           strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(Txt1(9)) & "') OR (PA27>='" & GetNewFagent(Txt1(9)) & "') OR (PA28>='" & GetNewFagent(Txt1(9)) & "') OR (PA29>='" & GetNewFagent(Txt1(9)) & "') OR (PA30>='" & GetNewFagent(Txt1(9)) & "'))"
           strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(Txt1(9)) & "') OR (SP58>='" & GetNewFagent(Txt1(9)) & "') OR (SP59>='" & GetNewFagent(Txt1(9)) & "')) "
       Else
           If Len(Txt1(10)) <> 0 Then
               strSQL1 = strSQL1 + " AND ((PA26<='" & GetNewFagent(Txt1(10)) & "') OR (PA27<='" & GetNewFagent(Txt1(10)) & "') OR (PA28<='" & GetNewFagent(Txt1(10)) & "') OR (PA29<='" & GetNewFagent(Txt1(10)) & "') OR (PA30<='" & GetNewFagent(Txt1(10)) & "'))"
               strSQL2 = strSQL2 + " AND ((SP05<='" & GetNewFagent(Txt1(10)) & "') OR (SP58<='" & GetNewFagent(Txt1(10)) & "') OR (SP59<='" & GetNewFagent(Txt1(10)) & "')) "
           End If
       End If
   End If
   If Len(Txt1(9)) <> 0 Or Len(Txt1(10)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & Txt1(9) & "-" & Txt1(10) 'Add By Sindy 2010/11/30
   End If
   If Len(Txt1(11)) <> 0 And Len(Txt1(12)) <> 0 Then
       strSQL1 = strSQL1 + " AND (PA75>='" & GetNewFagent(Txt1(11)) & "' AND PA75<='" & GetNewFagent(Txt1(12)) & "')"
       strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(Txt1(11)) & "' AND SP26<='" & GetNewFagent(Txt1(12)) & "')"
   Else
       If Len(Txt1(11)) <> 0 Then
           strSQL1 = strSQL1 + " AND (PA75>='" & GetNewFagent(Txt1(11)) & "')"
           strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(Txt1(11)) & "') "
       Else
           If Len(Txt1(12)) <> 0 Then
               strSQL1 = strSQL1 + " AND (PA75<='" & GetNewFagent(Txt1(12)) & "')"
               strSQL2 = strSQL2 + " AND (SP26<='" & GetNewFagent(Txt1(12)) & "') "
           End If
       End If
   End If
   If Len(Txt1(11)) <> 0 Or Len(Txt1(12)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(11) & "-" & Txt1(12) 'Add By Sindy 2010/11/30
   End If
   
   If txtPA46 = "Y" Then
      strSQL1 = strSQL1 & " And PA09<>'056' AND PA46='Y' "
      pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 10) & txtPA46 'Add By Sindy 2010/11/30
   End If
   
   'Add by Morgan 2006/5/30
   If Txt1(5) = "2" Then
      If Txt1(13) = "1" Then
         strSQL1 = strSQL1 & " And SUBSTR(S1.ST15,1,1)<>'S'"
         stCon = stCon & " And SUBSTR(S1.ST15,1,1)<>'S'"
         pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 5) & "1.非智權部同仁" 'Add By Sindy 2010/11/30
      ElseIf Txt1(13) = "2" Then
         pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 5) & "2.全部" 'Add By Sindy 2010/11/30
      End If
   End If
   
   CheckOC
   strSql = "select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(pa09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),PA72,PA08,NP07,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND pa09=na01(+) AND PA57 IS NULL " & strSQL1
   strSql = strSql + " union all select A0902,np10 AS B,CP05,NP08,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(sp09,'000',cpm03,cpm04),S2.ST02,NP09,NVL(NA03,NA04),' ',' ',CP10,NP10,NA21,NA23,NA25,CP09,CP27,np22,np01 FROM ACC090,STAFF S1,STAFF S2,NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION WHERE S1.ST15=a0901(+) AND np10=S1.ST01(+) AND cp14=S2.ST01(+) AND NP01=cp09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND np02=cpm01(+) AND to_char(np07)=cpm02(+) AND sp09=na01(+) AND SP15 IS NULL " & strSQL2
   strSql = strSql + " ORDER BY A0902,B,CP05,A "
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/30
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 20
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               StrYear2 = ""
               If Val(strTemp(14)) = 605 Or Val(strTemp(14)) = 606 Or Val(strTemp(14)) = 607 Then
               If Len(Trim(strTemp(12))) <> 0 Then
                   strTemp1 = Split(strTemp(12), ",")
                   StrYear1 = strTemp1(LBound(strTemp1))
                   For i = LBound(strTemp1) To UBound(strTemp1)
                     If Val(strTemp1(i)) > Val(StrYear1) Then
                        StrYear1 = strTemp1(i)
                     End If
                   Next i
                   Select Case Val(strTemp(13))
                   Case 1
                        strTemp1 = Split(strTemp(16), ",")
                        For i = 0 To UBound(strTemp1)
                           If Val(Trim(strTemp1(i))) > Val(Trim(StrYear1)) Then
                                StrYear2 = strTemp1(i)
                                Exit For
                           Else
                              StrYear2 = ""
                           End If
                        Next i
                   Case 2
                        strTemp1 = Split(strTemp(17), ",")
                        For i = 0 To UBound(strTemp1)
                           If Val(Trim(strTemp1(i))) > Val(Trim(StrYear1)) Then
                              StrYear2 = strTemp1(i)
                              Exit For
                           Else
                              StrYear2 = ""
                           End If
                        Next i
                   Case 3
                        strTemp1 = Split(strTemp(18), ",")
                        For i = 0 To UBound(strTemp1)
                           If Val(Trim(strTemp1(i))) > Val(Trim(StrYear1)) Then
                              StrYear2 = strTemp1(i)
                              Exit For
                           Else
                              StrYear2 = ""
                           End If
                        Next i
                   End Select
               End If
               Else
                   strTemp(12) = ""
               End If
               
               If strTemp(3) < Format(Now, "YYYYMMDD") Then
                   strTemp(3) = "*" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
               Else
                   If strTemp(3) = Format(Now, "YYYYMMDD") Then
                       strTemp(3) = "V" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                   Else
                       If UCase(Mid(strTemp(19), 1, 1)) = "C" And Trim(strTemp(20)) = "" Then
                           strTemp(3) = "#" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                       Else
                           If Val(strTemp(14)) = 605 And (SystemNumber(strTemp(4), 1) = "FCP" Or SystemNumber(strTemp(4), 1) = "P") Then
                               strSql = "select pa14 from patent where pa01='" & SystemNumber(strTemp(4), 1) & "' and pa02='" & _
                                            SystemNumber(strTemp(4), 2) & "' and pa03='" & SystemNumber(strTemp(4), 3) & "' and pa04='" & _
                                            SystemNumber(strTemp(4), 4) & "' and pa09='000' "
                               CheckOC3
                               AdoRecordSet3.CursorLocation = adUseClient
                               AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                               If AdoRecordSet3.RecordCount <> 0 Then
                                   If CheckStr(AdoRecordSet3.Fields(0).Value) = "" Then
                                          strTemp(3) = "!" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                                   Else
                                          strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                                   End If
                               Else
                                   strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                               End If
                           Else
                               strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                           End If
                       End If
                   End If
               End If
               strTemp(2) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
               strTemp(10) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(10)))
               strTemp(12) = StrYear2
               strSql = "INSERT INTO R040304_1 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               DoEvents
               .MoveNext
           Loop
       End With
       PrintData2 "(前期限補通知)"
   End If
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/30 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";內專智權人員期限管制表-(恢復期限)" 'Add By Sindy 2010/11/30
   pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Txt1(1) 'Add By Sindy 2010/11/30
   stCon = stCon & " and NP08<" & TransDate(Txt1(1), 2)
   
   '2.申復or修正有修改的期限表
   strSql = "select A0902 C01,S1.ST02 C02,CP05 C03,NP08 C04,NP02||'-'||NP03||'-'||NP04||'-'||NP05 C05" & _
      ",NVL(PA05,NVL(PA06,PA07)) C06,PA11 C07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) C08" & _
      ",decode(pa09,'000',cpm03,cpm04) C09,S2.ST02 C10,NP09 C11,NVL(NA03,NA04) C12,NULL C13" & _
      ",DECODE( SIGN(TO_CHAR(SYSDATE,'YYYYMMDD')-NP08),1,'*',0,'V',DECODE(CP27,NULL,'#')) C14,NP10 C15" & _
      " FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION,STAFF S1,STAFF S2,ACC090,CUSTOMER,CASEPROPERTYMAP" & _
      " where NP06 IS NULL and np07 in ('204','205') and np20 is not null AND CP09(+)=NP01" & _
      " AND NP01=cp09(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+)" & _
      " and pa08='2' and pa09='000' AND PA57 IS NULL AND na01(+)=pa09" & stCon & _
      " AND S1.ST01(+)=np10 AND S2.ST01(+)=cp14 AND a0901(+)=S1.ST15" & _
      " AND cu01(+)=SUBSTR(PA26,1,8) AND cu02(+)=SUBSTR(Pa26,9,1)" & _
      " AND cpm01(+)=np02 AND cpm02(+)=to_char(np07)" & _
      " ORDER BY 1,15,3,5"
   
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/30
         PrintData3 AdoRecordSet3
      End If
   End With
End Sub

'Add by Morgan 2005/8/26
Private Sub PrintData3(ByRef p_adoTmp As ADODB.Recordset)
   Dim stTitlePlus As String
   stTitlePlus = "(恢復期限)"
   Page = 1
   With p_adoTmp
      .MoveFirst
      Area = CheckStr(.Fields(0))
      PrintTitle1 stTitlePlus
      Do While .EOF = False
          For i = 0 To 12
            Select Case i
               Case 2, 3, 10 '日期欄位
                  strTemp(i) = Format(CheckStr(.Fields(i)) - 19110000, "###/##/##")
                  If i = 3 Then
                     strTemp(i) = .Fields(13) & strTemp(i)
                  End If
               Case Else
                  strTemp(i) = CheckStr(.Fields(i))
            End Select
          Next i
          If Area <> strTemp(0) Then
              Area = strTemp(0)
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print String(200, "-")
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函, ! 表示尚未公告，年費為預估值 "
              Printer.NewPage
              Page = Page + 1
              PrintTitle1 stTitlePlus
          End If
          strTemp(5) = StrConv(MidB(StrConv(strTemp(5), vbFromUnicode), 1, 12), vbUnicode)
          strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 12), vbUnicode)
          strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 14), vbUnicode)
          strTemp(8) = StrConv(MidB(StrConv(strTemp(8), vbFromUnicode), 1, 8), vbUnicode)
          strTemp(9) = StrConv(MidB(StrConv(strTemp(9), vbFromUnicode), 1, 8), vbUnicode)
          strTemp(11) = StrConv(MidB(StrConv(strTemp(11), vbFromUnicode), 1, 8), vbUnicode)
          PrintDatil1
          If iPrint > 10000 Then
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print String(200, "-")
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函 "
              Printer.NewPage
              Page = Page + 1
              PrintTitle1 stTitlePlus
          End If
          .MoveNext
      Loop
   End With
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函 "
   Printer.EndDoc
End Sub

Sub StrMenu1()                      '已收文承辦人
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R040304_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""

ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/30 清除查詢印表記錄檔欄位
If Txt1(6) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(2), 4) & "1.未收文" 'Add By Sindy 2010/11/30
ElseIf Txt1(6) = "2" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(2), 4) & "2.已收文" 'Add By Sindy 2010/11/30
ElseIf Txt1(6) = "3" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(2), 4) & "3.全部" 'Add By Sindy 2010/11/30
End If
pub_QL05 = pub_QL05 & ";內專承辦人期限管制表" 'Add By Sindy 2010/11/30

'本所期限
If Option1(0).Value = True Then
    strSQL1 = strSQL1 + " AND CP27 IS NULL AND CP57 IS NULL "
    pub_QL05 = pub_QL05 & ";" & Option1(0).Caption 'Add By Sindy 2010/11/30
    If Len(Trim(Txt1(1))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP06>=" & ChangeTStringToWString(Txt1(1)) & " "
      pub_QL05 = pub_QL05 & Txt1(1) 'Add By Sindy 2010/11/30
    End If
    If Len(Trim(Txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP06<=" & ChangeTStringToWString(Txt1(2)) & " "
      pub_QL05 = pub_QL05 & "-" & Txt1(2) 'Add By Sindy 2010/11/30
   'Add By Cheng 2002/03/19
    Else
      If Len(Trim(Txt1(1))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP06<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
         pub_QL05 = pub_QL05 & "-" & (ServerDate - 19110000) 'Add By Sindy 2010/11/30
      End If
    End If
'法定期限
Else
    strSQL1 = strSQL1 + " AND CP27 IS NULL AND CP57 IS NULL "
   'Modify By Cheng 2002/03/19
'    If Len(Trim(txt1(1))) <> 0 Then
'      strSQL1 = strSQL1 & " AND CP07>=" & ChangeTStringToWString(txt1(1)) & " "
    pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/11/30
    If Len(Trim(Txt1(3))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP07>=" & ChangeTStringToWString(Txt1(3)) & " "
      pub_QL05 = pub_QL05 & Txt1(3) 'Add By Sindy 2010/11/30
    End If
   'Modify By Cheng 2002/03/19
'    If Len(Trim(txt1(2))) <> 0 Then
'      strSQL1 = strSQL1 & " AND CP07<=" & ChangeTStringToWString(txt1(2)) & " "
    If Len(Trim(Txt1(4))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP07<=" & ChangeTStringToWString(Txt1(4)) & " "
      pub_QL05 = pub_QL05 & "-" & Txt1(4) 'Add By Sindy 2010/11/30
    Else
      If Len(Trim(Txt1(3))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP07<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
         pub_QL05 = pub_QL05 & "-" & (ServerDate - 19110000) 'Add By Sindy 2010/11/30
      End If
    End If
End If
If Txt1(5) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 5) & "1.承辦人" 'Add By Sindy 2010/11/30
ElseIf Txt1(5) = "2" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 5) & "2.智權人員" 'Add By Sindy 2010/11/30
End If
If Len(Txt1(5)) <> 0 And Len(Txt1(7)) <> 0 Then
   'Modify by Morgan 2006/10/17
   'If Val(txt1(6)) = 1 Then
   If Val(Txt1(5)) = 2 Then
        strSQL1 = strSQL1 + " AND CP14='" & Txt1(7) & "' "
    Else
        strSQL1 = strSQL1 + " AND CP13='" & Txt1(7) & "' "
    End If
    pub_QL05 = pub_QL05 & ";" & Label1(3) & Txt1(7) & lbl1(0) 'Add By Sindy 2010/11/30
End If

'Removed by Morgan 2013/10/23 本段為已收文而條件卻判斷未收文???
''Add By Cheng 2002/12/12
''除非有指定承辦人否則, 承辦人為陳玲玲與莊敏惠的資料不印
'If Val(txt1(6)) = 1 Then
'    If Me.txt1(7).Text = "" Then
'        strSQL1 = strSQL1 + " AND CP14<>'81002' And CP14<>'73017' "
'    End If
'End If

If Len(Txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10='" & Txt1(8) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(8) & lbl1(1) 'Add By Sindy 2010/11/30
End If
strSQL2 = strSQL1
If Len(StrTest) <> 0 Then
   strSQL1 = strSQL1 + " AND cp01 in (" & SQLGrpStr(StrTest, 1) & ") "
   strSQL2 = strSQL2 + " AND cp01 in (" & SQLGrpStr(StrTest, 5) & ") "
End If
If Len(Txt1(9)) <> 0 And Len(Txt1(10)) <> 0 Then
    strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(Txt1(9)) & "' AND PA26<='" & GetNewFagent(Txt1(10)) & "') OR (PA27>='" & GetNewFagent(Txt1(9)) & "' AND PA27<='" & GetNewFagent(Txt1(10)) & "') OR (PA28>='" & GetNewFagent(Txt1(9)) & "' AND PA28<='" & GetNewFagent(Txt1(10)) & "') OR (PA29>='" & GetNewFagent(Txt1(9)) & "' AND PA29<='" & GetNewFagent(Txt1(10)) & "') OR (PA30>='" & GetNewFagent(Txt1(9)) & "' AND PA30<='" & GetNewFagent(Txt1(10)) & "'))"
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(Txt1(9)) & "' AND SP05<='" & GetNewFagent(Txt1(10)) & "') OR (SP58>='" & GetNewFagent(Txt1(9)) & "' AND SP58<='" & GetNewFagent(Txt1(10)) & "') OR (SP59>='" & GetNewFagent(Txt1(9)) & "' AND SP59<='" & GetNewFagent(Txt1(10)) & "')) "
Else
    If Len(Txt1(9)) <> 0 Then
        strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(Txt1(9)) & "') OR (PA27>='" & GetNewFagent(Txt1(9)) & "') OR (PA28>='" & GetNewFagent(Txt1(9)) & "') OR (PA29>='" & GetNewFagent(Txt1(9)) & "') OR (PA30>='" & GetNewFagent(Txt1(9)) & "'))"
        strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(Txt1(9)) & "') OR (SP58>='" & GetNewFagent(Txt1(9)) & "') OR (SP59>='" & GetNewFagent(Txt1(9)) & "')) "
    Else
        If Len(Txt1(10)) <> 0 Then
            strSQL1 = strSQL1 + " AND ((PA26<='" & GetNewFagent(Txt1(10)) & "') OR (PA27<='" & GetNewFagent(Txt1(10)) & "') OR (PA28<='" & GetNewFagent(Txt1(10)) & "') OR (PA29<='" & GetNewFagent(Txt1(10)) & "') OR (PA30<='" & GetNewFagent(Txt1(10)) & "'))"
            strSQL2 = strSQL2 + " AND ((SP05<='" & GetNewFagent(Txt1(10)) & "') OR (SP58<='" & GetNewFagent(Txt1(10)) & "') OR (SP59<='" & GetNewFagent(Txt1(10)) & "')) "
        End If
    End If
End If
If Len(Txt1(9)) <> 0 Or Len(Txt1(10)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & Txt1(9) & "-" & Txt1(10) 'Add By Sindy 2010/11/30
End If
If Len(Txt1(11)) <> 0 And Len(Txt1(12)) <> 0 Then
    strSQL1 = strSQL1 + " AND (PA75>='" & GetNewFagent(Txt1(11)) & "' AND PA75<='" & GetNewFagent(Txt1(12)) & "')"
    strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(Txt1(11)) & "' AND SP26<='" & GetNewFagent(Txt1(12)) & "')"
Else
    If Len(Txt1(11)) <> 0 Then
        strSQL1 = strSQL1 + " AND (PA75>='" & GetNewFagent(Txt1(11)) & "')"
        strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(Txt1(11)) & "') "
    Else
        If Len(Txt1(12)) <> 0 Then
            strSQL1 = strSQL1 + " AND (PA75<='" & GetNewFagent(Txt1(12)) & "')"
            strSQL2 = strSQL2 + " AND (SP26<='" & GetNewFagent(Txt1(12)) & "') "
        End If
    End If
End If
If Len(Txt1(11)) <> 0 Or Len(Txt1(12)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(11) & "-" & Txt1(12) 'Add By Sindy 2010/11/30
End If

'Add by Morgan 2005/2/14
If txtPA46 = "Y" Then
   strSQL1 = strSQL1 & " And PA09<>'056' AND PA46='Y' "
   pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 10) & txtPA46 'Add By Sindy 2010/11/30
End If

'Add by Morgan 2006/5/30
If Txt1(5) = "2" Then
   If Txt1(13) = "1" Then
      strSQL1 = strSQL1 & " And SUBSTR(S2.ST15,1,1)<>'S'"
      pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 5) & "1.非智權部同仁" 'Add By Sindy 2010/11/30
   ElseIf Txt1(13) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 5) & "2.全部" 'Add By Sindy 2010/11/30
   End If
End If

'Addec by Morgan 2012/5/22 FMP案改可選擇
If Check1.Value = 1 Then
   pub_QL05 = pub_QL05 & ";" & Check1.Caption
Else
    strSQL1 = strSQL1 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and ep33>0)) "
    strSQL2 = strSQL2 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and ep33>0)) "
End If
'end 2012/5/22

CheckOC
'Modify By Cheng 2002/07/24
'已閉卷的資料不印
'strSQL = "SELECT S1.ST02,CP06,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),decode(pa09,'000',cpm03,cpm04),S2.ST02,CP07,CP27,CP64 FROM STAFF S1,STAFF S2,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION WHERE cp14=S1.ST01(+) AND cp13=S2.ST01(+) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND pa09=na01(+) " & strSQL1
'strSQL = strSQL + " union all select S1.ST02,CP06,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),decode(sp09,'000',cpm03,cpm04),S2.ST02,CP07,CP27,CP64 FROM STAFF S1,STAFF S2,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION WHERE cp14=S1.ST01(+) AND cp13=S2.ST01(+) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) " & strSQL2
'Modified by Morgan 2012/5/22 +Engineerprogress
strSql = "SELECT S1.ST02,CP06,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),PA11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(pa09,'000',cpm03,cpm04),S2.ST02,CP07,CP27,CP64 FROM STAFF S1,STAFF S2,CASEPROGRESS,PATENT,CUSTOMER,CASEPROPERTYMAP,NATION,Engineerprogress WHERE cp14=S1.ST01(+) AND cp13=S2.ST01(+) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND pa09=na01(+) AND PA57 IS NULL and ep02(+)=cp09 " & strSQL1
strSql = strSql + " union all select S1.ST02,CP06,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),SP11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(sp09,'000',cpm03,cpm04),S2.ST02,CP07,CP27,CP64 FROM STAFF S1,STAFF S2,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,CASEPROPERTYMAP,NATION,Engineerprogress WHERE cp14=S1.ST01(+) AND cp13=S2.ST01(+) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) AND SP15 IS NULL  and ep02(+)=cp09 " & strSQL2
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/30
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp(1) < Format(Now, "YYYYMMDD") Then
                strTemp(1) = "*" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
            Else
                If strTemp(1) = Format(Now, "YYYYMMDD") Then
                    strTemp(1) = "V" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                Else
                    strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                End If
            End If
            strTemp(8) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(8)))
            strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
            strSql = "INSERT INTO R040304_2 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            DoEvents
            .MoveNext
        Loop
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/30
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
CheckOC
PrintData1
'91.7.22 modify by sonia
'ShowPrintOk
   s = MsgBox("承辦人管制表列印完成!!", , "列印成功")
'91.7.22 end
Screen.MousePointer = vbDefault
End Sub

Sub PrintData1()                     '已收文承辦人
'91/3/12  nick 修改
'因為邱小姐說日期加符號，排序時，不能用符號排序
'strSQL = "SELECT * FROM R040304_2 WHERE ID='" & strUserNum & "' ORDER BY R023001,R023002,R023003 "
strSql = "SELECT * FROM R040304_2 WHERE ID='" & strUserNum & "' ORDER BY R023001,R023002,decode(substr(R023003,1,1),'#',substr(r023003,2,10),'V',substr(r023003,2,10),'*',substr(r023003,2,10),r023003) "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Area = CheckStr(.Fields(0))
        PrintTitle
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Area <> strTemp(0) Then
     
                Area = strTemp(0)
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函 "
                Printer.NewPage
                Page = Page + 1
                PrintTitle
     
            End If
            strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 18), vbUnicode)
            strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 12), vbUnicode)
            strTemp(5) = StrConv(MidB(StrConv(strTemp(5), vbFromUnicode), 1, 20), vbUnicode)
            strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(10) = StrConv(MidB(StrConv(strTemp(10), vbFromUnicode), 1, 8), vbUnicode)
            PrintDatil
            If iPrint > 10000 Then
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函 "
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            .MoveNext
        Loop
    End With
End If
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函 "
Printer.EndDoc
End Sub

'Add by Morgan 2005/8/26 加參數
Sub PrintData2(Optional ByVal p_TitlePlus As String)                     '未收文智權人員
strSql = "SELECT nvl(a0902,st03),st02,R022003,R022004,R022005,R022006,R022007,R022008,R022009,R022010,R022011,r022012,r022013,R022001,R022002 FROM R040304_1,staff,acc090 WHERE ID='" & strUserNum & "' and r022002=st01(+) and st03=a0901(+) ORDER BY st03,R022002,R022004,R022005 "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Area = CheckStr(.Fields(0))
        Iman = CheckStr(.Fields(1))                '20140227ADD By eric
        PrintTitle1 p_TitlePlus
        Do While .EOF = False
            For i = 0 To 12
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            '20140227START ADD By eric
            'If Area <> strTemp(0) Then
            '
            '    Area = strTemp(0)
            '    Printer.CurrentX = 500
            '    Printer.CurrentY = iPrint
            '    Printer.Print String(200, "-")
            '    iPrint = iPrint + 300
            '    Printer.CurrentX = 500
            '    Printer.CurrentY = iPrint
            '    Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函, ! 表示尚未公告，年費為預估值 "
            '    Printer.NewPage
            '    Page = Page + 1
            '    PrintTitle1 p_TitlePlus
            '
            'End If
            If Area <> strTemp(0) Then
               Area = strTemp(0)
               Iman = strTemp(1)
                
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               iPrint = iPrint + 300
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函, ! 表示尚未公告，年費為預估值 "
               Printer.NewPage
               Page = Page + 1
               PrintTitle1 p_TitlePlus
            Else
               If Iman <> strTemp(1) Then
                  Iman = strTemp(1)
             
                  If Txt1(13) = "1" Then                        '非智權部同仁才換頁
                     Printer.CurrentX = 500
                     Printer.CurrentY = iPrint
                     Printer.Print String(200, "-")
                     iPrint = iPrint + 300
                     Printer.CurrentX = 500
                     Printer.CurrentY = iPrint
                     Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函, ! 表示尚未公告，年費為預估值 "
                     Printer.NewPage
                     Page = Page + 1
                     PrintTitle1 p_TitlePlus
                  End If
               End If
            End If
            '20140227END
            strTemp(5) = StrConv(MidB(StrConv(strTemp(5), vbFromUnicode), 1, 12), vbUnicode)
            strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 12), vbUnicode)
            strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 14), vbUnicode)
            strTemp(8) = StrConv(MidB(StrConv(strTemp(8), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(9) = StrConv(MidB(StrConv(strTemp(9), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(11) = StrConv(MidB(StrConv(strTemp(11), vbFromUnicode), 1, 8), vbUnicode)
            PrintDatil1
            If iPrint > 10000 Then
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函 "
                Printer.NewPage
                Page = Page + 1
                PrintTitle1 p_TitlePlus
            End If
            .MoveNext
        Loop
    End With
End If
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "* 表示逾本所期限, V 表示當日本所期限, # 表示承辦人尚未通知主管機關來函 "
Printer.EndDoc
End Sub

'Modify by Morgan 2005/8/26 加參數
Sub PrintTitle1(Optional ByVal p_TitlePlus As String)
GetPleft1
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
'Modify by Morgan 2005/8/26
Printer.CurrentX = 6000 - Printer.TextWidth(p_TitlePlus) / 2
Printer.CurrentY = iPrint

'Modify by Morgan 2005/8/26
'Printer.Print "內專智權人員期限管制表"
Printer.Print "內專智權人員期限管制表" & p_TitlePlus
'2005/8/26 end

iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
   'Modify by Morgan 2005/8/29 附加報表不印
   If p_TitlePlus = "" Then
      Printer.Print "本所期限：" & ChangeTStringToTDateString(Txt1(1)) & "－" & ChangeTStringToTDateString(Txt1(2))
   End If
Else
    Printer.Print "法定期限：" & ChangeTStringToTDateString(Txt1(3)) & "－" & ChangeTStringToTDateString(Txt1(4))
End If
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
Printer.Print "業務區：" & Area
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
Printer.Print "智權人員"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請案號"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "法定期限"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "繳費年度"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
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
Printer.Print "內專承辦人期限管制表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(Txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(2))
Else
    Printer.Print "法定期限：" & Format(ChangeTStringToTDateString(Txt1(3)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(4))
End If
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
Printer.Print "承辦人：" & Area
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
Printer.Print "本所期限"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "申請案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "法定期限"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "進度備註"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintDatil1()
For i = 1 To 12
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintDatil()
For i = 1 To 10
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft1()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 1600
PLeft(3) = 2600 + 100 + 50
PLeft(4) = 3900 + 100 + 50
PLeft(5) = 5400 + 100 + 100 + 50
PLeft(6) = 6900 + 100 + 100 + 50
PLeft(7) = 8400 + 100 + 100 + 50
PLeft(8) = 10200 + 100 + 100 + 50
PLeft(9) = 11300 + 100 + 100 + 50
PLeft(10) = 12500 + 100 + 100 + 50
PLeft(11) = 13600 + 100 + 100 + 50
PLeft(12) = 14700 + 100 + 100 + 50
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 1750
PLeft(3) = 3400
PLeft(4) = 5700
PLeft(5) = 7700
PLeft(6) = 10200
PLeft(7) = 11300
PLeft(8) = 12500
PLeft(9) = 13600
PLeft(10) = 14500
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
StrTest2 = "P,PS"
StrTest = StrTest2
strTemp1 = Split(UCase(StrTest), ",")
strTemp2 = Split(UCase(GetSystemKindByNick), ",")
For i = 0 To UBound(strTemp1)
    s = 0
    For j = 0 To UBound(strTemp2)
        If strTemp2(j) = strTemp1(i) Then
            s = 1
            Exit For
        End If
    Next j
    If s = 0 Then
        StrTest = Replace(StrTest, strTemp1(i), "")
    End If
Next i
Txt1(0) = StrTest
Txt1(6) = "3"
Txt1(5) = "2"
Txt1(13).Enabled = True: Txt1(13) = "1" 'Add by Morgan 2006/5/30
Option1_Click 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040304 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
If Option1(0).Value = True Then
   Txt1(1).Enabled = True
   Txt1(2).Enabled = True
   Txt1(3).Enabled = False
   Txt1(4).Enabled = False
   Txt1(1).SetFocus
   txt1_GotFocus (1)
Else
   Txt1(1).Enabled = False
   Txt1(2).Enabled = False
   Txt1(3).Enabled = True
   Txt1(4).Enabled = True
   Txt1(3).SetFocus
   txt1_GotFocus (3)
End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add by Morgan 2006/5/30
   Select Case Index
      Case 5
         If KeyAscii = Asc("2") Then
            Txt1(13).Enabled = True: Txt1(13) = "1"
         Else
            Txt1(13) = "": Txt1(13).Enabled = False
         End If
   End Select
   'end 2006/5/30
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 2, 4 '本所期限, 法定期限
         'Modify By Cheng 2002/09/11
         If blnClkSure = False Then
            If Txt1(Index - 1) <> "" Then
               If RunNick(Txt1(Index - 1), Txt1(Index)) Then
                  Txt1(Index - 1).SetFocus
               End If
            End If
         Else
            blnClkSure = False
         End If
      Case 10
         If Len(Txt1(Index - 1)) <> 0 Then
            If Left(Txt1(Index - 1), 6) <> Left(Txt1(Index), 6) Then
                s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
                Txt1(Index - 1).SetFocus
                Exit Sub
            End If
         End If
         If RunNick(Txt1(Index - 1), Txt1(Index)) Then
            Txt1(Index - 1).SetFocus
         End If
      Case 12
         If Len(Txt1(Index - 1)) <> 0 Then
            If Left(Txt1(Index - 1), 6) <> Left(Txt1(Index), 6) Then
                s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
                Txt1(Index - 1).SetFocus
                Exit Sub
            End If
         End If
         If RunNick(Txt1(Index - 1), Txt1(Index)) Then
            Txt1(Index - 1).SetFocus
         End If
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Cancel = False
Select Case Index
Case 0
      strTemp1 = Split(GetSystemKindByNick, ",")
      strTemp2 = Split(Txt1(0), ",")
      For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox("USER : " & strUserNum & " 沒有 " & strTemp2(i) & " 的使用權限 ", , "USER 輸入錯誤")
            Cancel = True
        End If
      Next i
Case 1, 2, 3, 4
   If Txt1(Index) <> "" Then Cancel = Not ChkDate(Txt1(Index).Text)
Case 5
      Select Case Val(Txt1(Index))
      Case 1, 2
      Case Else
         s = MsgBox("管制對象只能輸入 1,2 !!", , "USER 輸入錯誤")
         Cancel = True
      End Select
Case 6
     Select Case Val(Txt1(Index))
     Case 1, 2, 3
     Case Else
          s = MsgBox("列印別只能輸入 1,2,3 !!", , "USER 輸入錯誤")
          Cancel = True
     End Select
Case 7
     If Txt1(7) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(txt1(7), strExc(0)) Then
         If ClsPDGetStaff(Txt1(7), strExc(0)) Then
            lbl1(0) = strExc(0)
         Else
            lbl1(0) = ""
            Cancel = True
         End If
     Else
         lbl1(0) = ""
     End If
Case 8 '案件性質
     If Txt1(8) <> "" Then
         lbl1(1) = GetPrjState6HM("P", Txt1(8))
         If lbl1(1) = "" Then
            MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
     Else
         lbl1(1) = ""
     End If
'Add by Morgan 2006/5/30
Case 13
      Select Case Val(Txt1(Index))
      Case 1, 2
      Case Else
         s = MsgBox("列印對象只能輸入 1,2 !!", , "USER 輸入錯誤")
         Cancel = True
      End Select
End Select
If Cancel Then TextInverse Txt1(Index)
End Sub
'Add by Morgan 2005/2/14 加PCT進入國家階段條件
Private Sub txtPA46_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtPA46.IMEMode = 2
   CloseIme
   TextInverse txtPA46
End Sub

Private Sub txtPA46_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub
