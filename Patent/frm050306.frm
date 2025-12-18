VERSION 5.00
Begin VB.Form frm050306 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員案件明細表"
   ClientHeight    =   3195
   ClientLeft      =   1530
   ClientTop       =   2085
   ClientWidth     =   3120
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3120
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2256
      TabIndex        =   14
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1416
      TabIndex        =   13
      Top             =   36
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   2196
      MaxLength       =   9
      TabIndex        =   12
      Top             =   2868
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   996
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2868
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   2196
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2532
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   996
      MaxLength       =   9
      TabIndex        =   9
      Top             =   2532
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2196
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2196
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   996
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2196
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1008
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1848
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2196
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1524
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   996
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1524
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2196
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1188
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   996
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1188
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   996
      MaxLength       =   1
      TabIndex        =   1
      Top             =   852
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   996
      TabIndex        =   0
      Top             =   516
      Width           =   2085
   End
   Begin VB.Line Line1 
      X1              =   1836
      X2              =   2076
      Y1              =   1308
      Y2              =   1308
   End
   Begin VB.Line Line5 
      X1              =   1836
      X2              =   2076
      Y1              =   2988
      Y2              =   2988
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   150
      TabIndex        =   24
      Top             =   2865
      Width           =   720
   End
   Begin VB.Line Line4 
      X1              =   1836
      X2              =   2076
      Y1              =   2652
      Y2              =   2652
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   150
      TabIndex        =   23
      Top             =   2535
      Width           =   720
   End
   Begin VB.Line Line3 
      X1              =   1836
      X2              =   2076
      Y1              =   2316
      Y2              =   2316
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   150
      TabIndex        =   22
      Top             =   2190
      Width           =   900
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1872
      TabIndex        =   21
      Top             =   1920
      Width           =   1128
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   150
      TabIndex        =   20
      Top             =   1860
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1836
      X2              =   2076
      Y1              =   1644
      Y2              =   1644
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   150
      TabIndex        =   19
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   150
      TabIndex        =   18
      Top             =   1185
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(1. 收文 2.發文)"
      Height          =   180
      Left            =   1836
      TabIndex        =   17
      Top             =   888
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "報表方式："
      Height          =   180
      Left            =   150
      TabIndex        =   16
      Top             =   885
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   150
      TabIndex        =   15
      Top             =   510
      Width           =   900
   End
End
Attribute VB_Name = "frm050306"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strTemp1 As Variant, strTemp2 As Variant, s As Integer, strSQL1 As String, strSQL2 As String
Dim strSql As String, j As Integer, PLeft(0 To 7) As Integer
Dim i As Integer, StrTest1 As String, StrTest2 As String
Dim strTemp(0 To 9) As String, Page As Integer, iPrint As Integer
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
        If Len(Trim(txt1(1))) = 0 Then
            s = MsgBox("報表方式不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1(1).SelStart = 0
            txt1(1).SelLength = Len(txt1(1))
            Exit Sub
        Else
            'Add By Cheng 2002/03/20
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
               Me.txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
            'Add By Cheng 2002/09/12
            If Me.txt1(2).Text <> "" And Me.txt1(3).Text <> "" Then
               If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
                  MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(2).SetFocus
                  txt1_GotFocus 2
                  Exit Sub
               End If
            End If
                     
            If Len(Trim(txt1(3))) = 0 Then
                s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
                txt1(2).SetFocus
                txt1_GotFocus (2)
                Exit Sub
            Else
               'Add By Cheng 2002/09/12
               If Me.txt1(4).Text <> "" And Me.txt1(5).Text <> "" Then
                  If Me.txt1(4).Text > Me.txt1(5).Text Then
                     MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(4).SetFocus
                     txt1_GotFocus 4
                     Exit Sub
                  End If
               End If
               If txt1(6) <> "" Then
                  'edit by nickc 2007/02/02 不用 dll 了
                  'If objPublicData.GetStaff(txt1(6), strExc(0)) Then
                  If ClsPDGetStaff(txt1(6), strExc(0)) Then
                     lbl1 = strExc(0)
                  Else
                     lbl1 = ""
                     Me.txt1(6).SetFocus
                     txt1_GotFocus 6
                     Exit Sub
                  End If
               End If
               If Me.txt1(7).Text <> "" And Me.txt1(8).Text <> "" Then
                  If Me.txt1(7).Text > Me.txt1(8).Text Then
                     MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(7).SetFocus
                     txt1_GotFocus 7
                     Exit Sub
                  End If
               End If
                
                If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
                    If Left(txt1(9), 6) <> Left(txt1(10), 6) Then
                        s = MsgBox("申請人前六碼必須相同!!", , "USER 輸入錯誤")
                        blnClkSure = True
                        txt1(9).SetFocus
                        txt1(9).SelStart = 0
                        txt1(9).SelLength = Len(txt1(9))
                        Exit Sub
                    End If
                End If
               'Add By Cheng 2002/09/12
               If Me.txt1(9).Text <> "" And Me.txt1(10).Text <> "" Then
                  If Me.txt1(9).Text > Me.txt1(10).Text Then
                     MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(9).SetFocus
                     txt1_GotFocus 9
                     Exit Sub
                  End If
               End If
                
                If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) <> 0 Then
                    If Left(txt1(11), 6) <> Left(txt1(12), 6) Then
                        s = MsgBox("代理人前六碼必須相同!!", , "USER 輸入錯誤")
                        blnClkSure = True
                        txt1(11).SetFocus
                        txt1(11).SelStart = 0
                        txt1(11).SelLength = Len(txt1(11))
                        Exit Sub
                    End If
                End If
               'Add By Cheng 2002/09/12
               If Me.txt1(11).Text <> "" And Me.txt1(12).Text <> "" Then
                  If Me.txt1(11).Text > Me.txt1(12).Text Then
                     MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(11).SetFocus
                     txt1_GotFocus 11
                     Exit Sub
                  End If
               End If
                
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                Process
                Me.Enabled = True
                Screen.MousePointer = vbDefault
                'PrintData
            End If
        End If
    End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Process()
ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位

Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from  R050306 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "delete from  R050306_1 WHERE ID='" & strUserNum & "' "
'系統類別
strSQL1 = " "
strSQL2 = " "
'Add By Cheng 2002/05/09
strSQL1 = strSQL1 + " AND CP09<'B' "
strSQL2 = strSQL2 + " AND CP09<'B' "

If Len(Trim(txt1(0))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
    strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
    pub_QL05 = pub_QL05 & ";" & Label1 & Trim(txt1(0)) 'Add By Sindy 2010/01/22
End If
'報表方式,日期
If Trim(txt1(1)) = "1" Then
   If Len(Trim(txt1(2))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & " "
      strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   If Len(Trim(txt1(3))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(3))) & " "
      strSQL2 = strSQL2 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(3))) & " "
   End If
   pub_QL05 = pub_QL05 & ";" & "收文" & Label4 & Trim(txt1(2)) & "-" & Trim(txt1(3)) 'Add By Sindy 2010/01/22
Else
   If Len(Trim(txt1(2))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(2))) & " "
      strSQL2 = strSQL2 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   If Len(Trim(txt1(3))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(3))) & " "
      strSQL2 = strSQL2 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(3))) & " "
   End If
   pub_QL05 = pub_QL05 & ";" & "發文" & Label4 & Trim(txt1(2)) & "-" & Trim(txt1(3)) 'Add By Sindy 2010/01/22
End If
'業務區
If Len(Trim(txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND cp12>='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND cp12>='" & txt1(4) & "' "
End If
If Len(Trim(txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND cp12<='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND cp12<='" & txt1(5) & "' "
End If
If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label5 & Trim(txt1(4)) & "-" & Trim(txt1(5)) 'Add By Sindy 2010/01/22
End If
'智權人員
If Len(Trim(txt1(6))) <> 0 Then
    strSQL1 = strSQL1 + " AND cp13='" & txt1(6) & "' "
    strSQL2 = strSQL2 + " AND cp13='" & txt1(6) & "' "
End If
pub_QL05 = pub_QL05 & ";" & Label6 & Trim(txt1(6)) 'Add By Sindy 2010/01/22
'案件性質
If Len(Trim(txt1(7))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & txt1(7) & "' "
    strSQL2 = strSQL2 + " AND CP10>='" & txt1(7) & "' "
End If
If Len(Trim(txt1(8))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & txt1(8) & "' "
    strSQL2 = strSQL2 + " AND CP10<='" & txt1(8) & "' "
End If
If Len(Trim(txt1(7))) <> 0 Or Len(Trim(txt1(8))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label8 & Trim(txt1(7)) & "-" & Trim(txt1(8)) 'Add By Sindy 2010/01/22
End If
'申請人
If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
    strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(9)) & "' AND PA26<='" & GetNewFagent(txt1(10)) & "') OR (PA27>='" & GetNewFagent(txt1(9)) & "' AND PA27<='" & GetNewFagent(txt1(10)) & "') OR (PA28>='" & GetNewFagent(txt1(9)) & "' AND PA28<='" & GetNewFagent(txt1(10)) & "') OR (PA29>='" & GetNewFagent(txt1(9)) & "' AND PA29<='" & GetNewFagent(txt1(10)) & "') OR (PA30>='" & GetNewFagent(txt1(9)) & "' AND PA30<='" & GetNewFagent(txt1(10)) & "')) "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(9)) & "' AND SP08<='" & GetNewFagent(txt1(10)) & "') OR (SP58<='" & GetNewFagent(txt1(9)) & "' AND SP58<='" & GetNewFagent(txt1(10)) & "') OR (SP59>='" & GetNewFagent(txt1(9)) & "' AND SP59<='" & GetNewFagent(txt1(10)) & "')) "
Else
    If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) = 0 Then
        strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(9)) & "' OR PA27>='" & GetNewFagent(txt1(9)) & "' OR PA28>='" & GetNewFagent(txt1(9)) & "' OR PA29>='" & GetNewFagent(txt1(9)) & "' OR PA30>='" & GetNewFagent(txt1(9)) & "') "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(9)) & "' OR SP58>='" & GetNewFagent(txt1(9)) & "' OR SP59>='" & GetNewFagent(txt1(9)) & "') "
    Else
        If Len(Trim(txt1(9))) = 0 And Len(Trim(txt1(10))) <> 0 Then
            strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(10)) & "' OR PA27<='" & GetNewFagent(txt1(10)) & "' OR PA28<='" & GetNewFagent(txt1(10)) & "' OR PA29<='" & GetNewFagent(txt1(10)) & "' OR PA30<='" & GetNewFagent(txt1(10)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(10)) & "' OR SP58<='" & GetNewFagent(txt1(10)) & "' OR SP59<='" & GetNewFagent(txt1(10)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(9))) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label9 & Trim(txt1(9)) & "-" & Trim(txt1(10)) 'Add By Sindy 2010/01/22
End If
'代理人
If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) <> 0 Then
    strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(11)) & "' AND PA75<='" & GetNewFagent(txt1(12)) & "' "
    strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(11)) & "' AND SP26<='" & GetNewFagent(txt1(12)) & "' "
Else
    If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) = 0 Then
        strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(11)) & "' "
        strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(11)) & "' "
    Else
        If Len(Trim(txt1(11))) = 0 And Len(Trim(txt1(12))) <> 0 Then
            strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(12)) & "' "
            strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(12)) & "' "
        End If
    End If
End If
If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label10 & Trim(txt1(11)) & "-" & Trim(txt1(12)) 'Add By Sindy 2010/01/22
End If
'Modify By Cheng 2002/05/09
'若已閉卷仍然要列印出來, 且在其本所案號右加印"*"號
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'strSQL = "SELECT cp12,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "' FROM PATENT,CASEPROGRESS,NATION,CASEPROPERTYMAP,STAFF,ACC090 WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND pa09=na01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND CP57 IS NULL AND CP13=ST01(+) and st03=a0901(+)  " & strSQL1
'strSQL = strSQL + " union all select cp12,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "' FROM SERVICEPRACTICE,CASEPROGRESS,NATION,CASEPROPERTYMAP,STAFF,ACC090 WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND (SP15<>'Y' OR SP15 IS NULL)  AND sp09=na01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND CP57 IS NULL AND CP13=ST01(+) and st03=a0901(+)   " & strSQL2
strSql = "SELECT cp12,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "' FROM PATENT,CASEPROGRESS,NATION,CASEPROPERTYMAP,STAFF,ACC090 WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND CP57 IS NULL AND CP13=ST01(+) and cp12=a0901(+) " & strSQL1
strSql = strSql + " union all select cp12,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",decode(CP24,'1','准','2','駁',''),'" & strUserNum & "' FROM SERVICEPRACTICE,CASEPROGRESS,NATION,CASEPROPERTYMAP,STAFF,ACC090 WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND CP57 IS NULL AND CP13=ST01(+) and cp12=a0901(+) " & strSQL2
cnnConnection.Execute "insert into r050306 " & strSql
CheckOC
strSql = "select * from r050306 where id='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
   cnnConnection.Execute "insert into r050306_1 select r004001,r004002,r004006,count(*),'" & strUserNum & "' from r050306 where id='" & strUserNum & "' group by r004001,r004002,r004006,'" & strUserNum & "' "
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

Private Sub PrintData()
strSql = "select nvl(a0902,r004001),nvl(st02,r004002),r004003,r004004,r004005,r004006,r004007,r004008,r004009,r004010,r004001,r004002 from r050306,acc090,staff where id='" & strUserNum & "' and r004001=a0901(+) and r004002=st01(+) "
If Trim(txt1(1)) = "1" Then
    strSql = strSql & " order by r004001,r004002,r004003,r004004 "
Else
    strSql = strSql & " order by r004001,r004002,r004009,r004004 "
End If
CheckOC
StrTest1 = "null"   '業務區
StrTest2 = "null"   '智權人員
Page = 1
PrintTitle
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    .MoveFirst
    If .RecordCount <> 0 Then
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If StrTest1 = "null" And StrTest2 = "null" Then
                StrTest1 = strTemp(0)
                StrTest2 = strTemp(1)
                PrintSubTitle
            End If
            strTemp(4) = StrToStr(strTemp(4), 24)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 4)
            If strTemp(0) = StrTest1 And strTemp(1) = StrTest2 Then
            Else
                .MovePrevious
                Printer.CurrentX = PLeft(0)
                Printer.CurrentY = iPrint
                Printer.Line (PLeft(0), iPrint)-(16000, iPrint)
                iPrint = iPrint + 300
                PrintEnd
                StrTest1 = strTemp(0)
                StrTest2 = strTemp(1)
                .MoveNext
                iPrint = iPrint + 300
                PrintSubTitle
            End If
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintSubTitle
            End If
            PrintDatil
            .MoveNext
        Loop
    End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Line (PLeft(0), iPrint)-(16000, iPrint)
iPrint = iPrint + 300
.MovePrevious
PrintEnd
End With
Printer.EndDoc
CheckOC

'Printer.Orientation = 2
'If Trim(txt1(1)) = "1" Then
'    DataEnvironment1.Commands(1).CommandText = "SELECT * FROM R050306 WHERE ID='" & strUserNum & "' ORDER BY R004001,R004002,R004003,R004004 "
'Else
'    DataEnvironment1.Commands(1).CommandText = "SELECT * FROM R050306 WHERE ID='" & strUserNum & "' ORDER BY R004001,R004002,R004009,R004004 "
'End If
'If DataEnvironment1.rsRPT050306_Grouping.State <> 0 Then
'    DataEnvironment1.rsRPT050306_Grouping.Close
'End If

'DataEnvironment1.RPT050306_Grouping
'DR050306.Orientation = rptOrientLandscape
'DR050306.Sections(2).Controls("lblusername").Caption = strUserName
'If txt1(1) = "1" Then
'    DR050306.Sections(2).Controls("Label5").Caption = "收文日期："
'Else
'    DR050306.Sections(2).Controls("Label5").Caption = "發文日期："
'End If
'DR050306.Sections(2).Controls("lblday1").Caption = ChangeTStringToTDateString(txt1(2))
'DR050306.Sections(2).Controls("lblday2").Caption = ChangeTStringToTDateString(txt1(3))
'DR050306.Sections(2).Controls("lbltoday").Caption = ChangeTStringToTDateString(GetTaiwanTodayDate)
'If Len(txt1(9)) <> 0 And Len(txt1(10)) <> 0 Then
'    DR050306.Sections(2).Controls("LABEL18").Caption = "申請人  " & txt1(9) & "-" & txt1(10)
'Else
'    DR050306.Sections(2).Controls("LABEL18").Caption = ""
'End If
'If Len(txt1(11)) <> 0 And Len(txt1(12)) <> 0 Then
'    DR050306.Sections(2).Controls("LABEL19").Caption = "代理人  " & txt1(11) & "-" & txt1(12)
'Else
'    DR050306.Sections(2).Controls("LABEL19").Caption = ""
'End If

'DR050306.Show
'DR050306.PrintReport
'DataEnvironment1.rsRPT050306_Grouping.Close
'DoEvents
'Sleep (5)
'strSQL = "delete from R050306_1"
'cnnConnection.Execute strSQL

'strSQL = "INSERT INTO R050306_1 (SELECT R004001,R004002,R004006,COUNT(R004006),ID FROM R050306 WHERE ID='" & strUserNum & "' GROUP BY R004001,R004002,R004006,ID)"
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'DoEvents
'DataEnvironment1.Commands(1).CommandText = "SELECT * FROM R050306_1 WHERE ID='" & strUserNum & "' ORDER BY R005001,R005002,R005003  "
'If DataEnvironment1.rsR050306_1.State <> 0 Then
'    DataEnvironment1.rsR050306_1.Close
'End If

'DataEnvironment1.R050306_1
'DR050306_1.Orientation = rptOrientLandscape
'If txt1(1) = "1" Then
'    DR050306_1.Sections(2).Controls("Label5").Caption = "收文日期："
'Else
'    DR050306_1.Sections(2).Controls("Label5").Caption = "發文日期："
'End If
'DR050306_1.Sections(2).Controls("lblusername").Caption = strUserName
'DR050306_1.Sections(2).Controls("lblday1").Caption = ChangeTStringToTDateString(txt1(2))
'DR050306_1.Sections(2).Controls("lblday2").Caption = ChangeTStringToTDateString(txt1(3))
'DR050306_1.Sections(2).Controls("lbltoday").Caption = ChangeTStringToTDateString(GetTaiwanTodayDate)
'DR050306_1.Show
'DR050306_1.PrintReport
'DataEnvironment1.rsR050306_1.Close
'DoEvents
ShowPrintOk
End Sub
Private Sub PrintTitle()
GetPleft
Printer.Orientation = 2
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "智權人員案件明細表") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "智權人員案件明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
If Trim(txt1(1)) = "1" Then
    Printer.Print "收文日期：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
Else
    Printer.Print "發文日期：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
End If
iPrint = iPrint + 300
Printer.CurrentX = 0
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
'If Len(txt1(9)) <> 0 And Len(txt1(10)) <> 0 Then
'    Printer.CurrentX = 10000
'    Printer.CurrentX = iPrint
'    Printer.Print "申請人  " & txt1(9) & "-" & txt1(10)
'    iPrint = iPrint + 300
'End If
'If Len(txt1(11)) <> 0 And Len(txt1(12)) <> 0 Then
'    Printer.CurrentX = 10000
'    Printer.CurrentY = iPrint
'    Printer.Print "代理人  " & txt1(11) & "-" & txt1(12)
'    iPrint = iPrint + 300
'End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Line (PLeft(0), iPrint)-(16000, iPrint)
iPrint = iPrint + 300
End Sub

Sub PrintEnd()
strSql = "select * from r050306_1 where id='" & strUserNum & "' and r005001='" & CheckStr(adoRecordset.Fields(10)) & "' and r005002='" & CheckStr(adoRecordset.Fields(11)) & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 4
                If .EOF = False Then
                    Printer.CurrentX = 0 + (i * 2600)
                    Printer.CurrentY = iPrint
                    Printer.Print CheckStr(.Fields(2))
                    Printer.CurrentX = 2600 + (i * 2600) - 600 + 500 - (Printer.TextWidth(CheckStr(.Fields(3))))
                    Printer.CurrentY = iPrint
                    Printer.Print CheckStr(.Fields(3))
                End If
                If .EOF = True Then Exit For
                
                .MoveNext
            Next i
            iPrint = iPrint + 300
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If .EOF = False Then
                  PrintSubTitle
                Else
                  iPrint = iPrint - 300
                End If
            End If
        Loop
    End If
End With
End Sub

Sub PrintSubTitle()
If iPrint >= 8600 Then
   Page = Page + 1
   Printer.NewPage
   PrintTitle
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區：" & StrTest1
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "智權人員：" & StrTest2
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintSubTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Line (PLeft(0), iPrint)-(16000, iPrint)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintSubTitle
    Exit Sub
End If
'Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "准駁"
'Printer.Font.Underline = False
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintSubTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Line (PLeft(0), iPrint)-(16000, iPrint)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintSubTitle
    Exit Sub
End If
End Sub

Private Sub PrintDatil()
For i = 2 To 9
    Printer.CurrentX = PLeft(i - 2)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1100
PLeft(2) = 3100
PLeft(3) = 10100
PLeft(4) = 11200
PLeft(5) = 12300
PLeft(6) = 13400
PLeft(7) = 14500
'Pleft(8) = 12300
'Pleft(9) = 13400
'Pleft(10) = 14500
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050306 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 202/09/12
   Select Case Index
   Case 1 '報表方式
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
Case 3, 5, 8
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
      End If
Case 10
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
Case 12
   'Add By Cheng 2002/09/12
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
Case 1
     Select Case txt1(1)
     Case "1", "2"
     Case Else
         s = MsgBox("報表方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
         Cancel = True
     End Select

Case 2, 3 '日期起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
Case 6
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
End Select
If Cancel Then TextInverse txt1(Index)
End Sub
