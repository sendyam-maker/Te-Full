VERSION 5.00
Begin VB.Form frm090610 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人每日分案情形查詢"
   ClientHeight    =   2865
   ClientLeft      =   705
   ClientTop       =   2850
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4320
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2280
      TabIndex        =   9
      Top             =   50
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3060
      TabIndex        =   10
      Top             =   50
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1164
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2484
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1992
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2484
      Width           =   465
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1152
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1536
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1944
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1536
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1164
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2040
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1632
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2040
      Width           =   300
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1188
      TabIndex        =   0
      Top             =   504
      Width           =   1995
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1176
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1020
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1956
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1020
      Width           =   705
   End
   Begin VB.Label Label2 
      Caption         =   "月"
      Height          =   180
      Index           =   1
      Left            =   2544
      TabIndex        =   18
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label2 
      Caption         =   "年"
      Height          =   180
      Index           =   0
      Left            =   1608
      TabIndex        =   17
      Top             =   2532
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "收文年月："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   16
      Top             =   2505
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "部門代號："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   150
      TabIndex        =   14
      Top             =   2070
      Width           =   660
   End
   Begin VB.Line Line2 
      X1              =   1596
      X2              =   2271
      Y1              =   1692
      Y2              =   1692
   End
   Begin VB.Line Line3 
      X1              =   1332
      X2              =   1797
      Y1              =   2184
      Y2              =   2184
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   7
      Left            =   2016
      TabIndex        =   13
      Top             =   2088
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   12
      Top             =   570
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   1584
      X2              =   2214
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frm090610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/12 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 32) As String, strTemp3 As String, TestOk As Boolean, StrSQL3 As String, k As Integer
Dim PLeft(0 To 32) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, SeekPrint As Integer, SeekPrintL As Integer
Dim bol911001checkRange  As Boolean

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Len(txt1(7)) = 0 Or Len(txt1(8)) = 0 Then
             s = MsgBox("收文年月不可空白!!", , "USER 輸入錯誤")
             'If Len(txt1(8)) = 0 Then txt1(8).SetFocus
             If Len(txt1(7)) = 0 Then txt1(7).SetFocus
             Exit Sub
         Else
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/17 清除查詢印表記錄檔欄位
             Process
             Me.Hide
             frm090610_1.Show
             'Me.Enabled = True
             Screen.MousePointer = vbDefault
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
cnnConnection.Execute "DELETE FROM R090610 WHERE ID='" & strUserNum & "'  "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(txt1(0), 2) & ") "
   StrSQL3 = StrSQL3 + " and CP01 in (" & SQLGrpStr(txt1(0), 3) & ") "
   StrSQL4 = StrSQL4 + " and CP01 in (" & SQLGrpStr(txt1(0), 4) & ") "
   strSQL5 = strSQL5 + " and CP01 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/17
End If
StrSQL6 = ""
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(1) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & txt1(1) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(2) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & txt1(2) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/17
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST03>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST03<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/17
End If
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & "-" & txt1(6) & Label1(7) 'Add By Sindy 2010/12/17
End If
'If Len(txt1(7)) <> 0 Then
'    StrSQL6 = StrSQL6 + " AND SUBSTR(CP05,1,4)=" & Val(txt1(7)) + 1911
'End If
'If Len(txt1(8)) <> 0 Then
'    StrSQL6 = StrSQL6 + " AND SUBSTR(CP05,5,2)=" & Val(txt1(8))
'End If
StrSQL6 = StrSQL6 & " and cp05>=" & Trim(Val(txt1(7)) + 1911) & Trim(Right(ChgNumByNick(txt1(8)), 2)) & "01 and cp05<=" & Trim(Val(txt1(7)) + 1911) & Trim(Right(ChgNumByNick(txt1(8)), 2)) & "31 "
pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(7) & txt1(8) 'Add By Sindy 2010/12/17
'strSQL = "SELECT ST02,DECODE(SUBSTR(CP05,7,2),01,1,0)," & _
         "DECODE(SUBSTR(CP05,7,2),02,1,0),DECODE(SUBSTR(CP05,7,2),03,1,0),DECODE(SUBSTR(CP05,7,2),04,1,0),DECODE(SUBSTR(CP05,7,2),05,1,0),DECODE(SUBSTR(CP05,7,2),06,1,0),DECODE(SUBSTR(CP05,7,2),07,1,0),DECODE(SUBSTR(CP05,7,2),08,1,0),DECODE(SUBSTR(CP05,7,2),09,1,0),DECODE(SUBSTR(CP05,7,2),10,1,0)," & _
         "DECODE(SUBSTR(CP05,7,2),11,1,0),DECODE(SUBSTR(CP05,7,2),12,1,0),DECODE(SUBSTR(CP05,7,2),13,1,0),DECODE(SUBSTR(CP05,7,2),14,1,0),DECODE(SUBSTR(CP05,7,2),15,1,0),DECODE(SUBSTR(CP05,7,2),16,1,0),DECODE(SUBSTR(CP05,7,2),17,1,0),DECODE(SUBSTR(CP05,7,2),18,1,0),DECODE(SUBSTR(CP05,7,2),19,1,0)," & _
         "DECODE(SUBSTR(CP05,7,2),20,1,0),DECODE(SUBSTR(CP05,7,2),21,1,0),DECODE(SUBSTR(CP05,7,2),22,1,0),DECODE(SUBSTR(CP05,7,2),23,1,0),DECODE(SUBSTR(CP05,7,2),24,1,0),DECODE(SUBSTR(CP05,7,2),25,1,0),DECODE(SUBSTR(CP05,7,2),26,1,0),DECODE(SUBSTR(CP05,7,2),27,1,0),DECODE(SUBSTR(CP05,7,2),28,1,0)," & _
         "DECODE(SUBSTR(CP05,7,2),29,1,0),DECODE(SUBSTR(CP05,7,2),30,1,0),DECODE(SUBSTR(CP05,7,2),31,1,0),1,'" & strUserNum & "' FROM CASEPROGRESS,STAFF,PATENT WHERE CP14=ST01(+) AND CP10 IN ('103','105') AND PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & StrSQL6 + StrSQL1
'strSQL = strSQL + " UNION all  SELECT ST02,DECODE(SUBSTR(CP05,7,2),01,1,0)," & _
         "DECODE(SUBSTR(CP05,7,2),02,1,0),DECODE(SUBSTR(CP05,7,2),03,1,0),DECODE(SUBSTR(CP05,7,2),04,1,0),DECODE(SUBSTR(CP05,7,2),05,1,0),DECODE(SUBSTR(CP05,7,2),06,1,0),DECODE(SUBSTR(CP05,7,2),07,1,0),DECODE(SUBSTR(CP05,7,2),08,1,0),DECODE(SUBSTR(CP05,7,2),09,1,0),DECODE(SUBSTR(CP05,7,2),10,1,0)," & _
         "DECODE(SUBSTR(CP05,7,2),11,1,0),DECODE(SUBSTR(CP05,7,2),12,1,0),DECODE(SUBSTR(CP05,7,2),13,1,0),DECODE(SUBSTR(CP05,7,2),14,1,0),DECODE(SUBSTR(CP05,7,2),15,1,0),DECODE(SUBSTR(CP05,7,2),16,1,0),DECODE(SUBSTR(CP05,7,2),17,1,0),DECODE(SUBSTR(CP05,7,2),18,1,0),DECODE(SUBSTR(CP05,7,2),19,1,0)," & _
         "DECODE(SUBSTR(CP05,7,2),20,1,0),DECODE(SUBSTR(CP05,7,2),21,1,0),DECODE(SUBSTR(CP05,7,2),22,1,0),DECODE(SUBSTR(CP05,7,2),23,1,0),DECODE(SUBSTR(CP05,7,2),24,1,0),DECODE(SUBSTR(CP05,7,2),25,1,0),DECODE(SUBSTR(CP05,7,2),26,1,0),DECODE(SUBSTR(CP05,7,2),27,1,0),DECODE(SUBSTR(CP05,7,2),28,1,0)," & _
         "DECODE(SUBSTR(CP05,7,2),29,1,0),DECODE(SUBSTR(CP05,7,2),30,1,0),DECODE(SUBSTR(CP05,7,2),31,1,0),2,'" & strUserNum & "' FROM CASEPROGRESS,STAFF,PATENT WHERE CP14=ST01(+) AND CP10 NOT IN ('103','105') AND PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & StrSQL6 + StrSQL1
StrSQL6 = StrSQL6 + " and CP26 IS NULL  "
'strSQL = "INSERT INTO R090610 " & strSQL
'Modify By Cheng 2003/10/01
'Begin
'strSQL = "select NVL(st02,CP14),cp05,cp10,cp09 from caseprogress,patent,staff where CP01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & strSQL1
''911107 nick 因為設計只有專利有
''strSQL = strSQL & " UNION all  SELECT NVL(st02,CP14),cp05,cp10,cp09 from caseprogress,trademark,staff where CP01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & strSQL2
''strSQL = strSQL & " UNION all  SELECT NVL(st02,CP14),cp05,cp10,cp09 from caseprogress,lawcase,staff where CP01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & StrSQL3
''strSQL = strSQL & " UNION all  SELECT NVL(st02,CP14),cp05,cp10,cp09 from caseprogress,hirecase,staff where CP01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & StrSQL4
''strSQL = strSQL & " UNION all  SELECT NVL(st02,CP14),cp05,cp10,cp09 from caseprogress,servicepractice,staff where CP01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & StrSQL5
'strSQL = strSQL & " UNION all  SELECT NVL(st02,CP14),cp05,'999',cp09 from caseprogress,trademark,staff where CP01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & strSQL2
'strSQL = strSQL & " UNION all  SELECT NVL(st02,CP14),cp05,'999',cp09 from caseprogress,lawcase,staff where CP01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & StrSQL3
'strSQL = strSQL & " UNION all  SELECT NVL(st02,CP14),cp05,'999',cp09 from caseprogress,hirecase,staff where CP01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & StrSQL4
'strSQL = strSQL & " UNION all  SELECT NVL(st02,CP14),cp05,'999',cp09 from caseprogress,servicepractice,staff where CP01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & StrSQL5
strSql = "select CP14,cp05,cp10,cp09 from caseprogress,patent,staff where CP01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & strSQL1
strSql = strSql & " UNION all  SELECT CP14,cp05,'999',cp09 from caseprogress,trademark,staff where CP01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & strSQL2
strSql = strSql & " UNION all  SELECT CP14,cp05,'999',cp09 from caseprogress,lawcase,staff where CP01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & StrSQL3
strSql = strSql & " UNION all  SELECT CP14,cp05,'999',cp09 from caseprogress,hirecase,staff where CP01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & StrSQL4
strSql = strSql & " UNION all  SELECT CP14,cp05,'999',cp09 from caseprogress,servicepractice,staff where CP01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp14=st01(+) and cp05 is not null " & StrSQL6 & strSQL5
'End
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            Select Case Val(CheckStr(.Fields(2)))
            Case 103, 105
                 cnnConnection.Execute "insert into r090610 (r105001,r1050" & Format(Val(Right(CheckStr(.Fields(1)), 2)) + 1, "00") & ",r105033,id) values ('" & CheckStr(.Fields(0)) & "',1,1,'" & strUserNum & "') "
            Case Else
                 cnnConnection.Execute "insert into r090610 (r105001,r1050" & Format(Val(Right(CheckStr(.Fields(1)), 2)) + 1, "00") & ",r105033,id) values ('" & CheckStr(.Fields(0)) & "',1,2,'" & strUserNum & "') "
            End Select
            .MoveNext
        Loop
    End If
End With
CheckOC
'cnnConnection.Execute strSQL
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = Systemkind_g
bol911001checkRange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090610 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '系統類別
     'Add By Cheng 2002/01/07
     Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
     strTemp1 = Split(UCase(Systemkind_g), ",")
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
   If RunNick(txt1(Index - 1), txt1(Index)) Then
       txt1(Index - 1).SetFocus
       txt1_GotFocus (Index - 1)
       Exit Sub
   End If
Case 5
     bol911001checkRange = True
     Select Case Trim(txt1(5))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(5).SetFocus
          txt1(5).SelStart = 0
          txt1(5).SelLength = Len(txt1(5))
          bol911001checkRange = False
          Exit Sub
     End Select
Case 6
     If bol911001checkRange = True Then
          Select Case Trim(txt1(6))
          Case "1", "2", "3", "4", "5", ""
          Case Else
               s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
               txt1(6).SetFocus
               txt1(6).SelStart = 0
               txt1(6).SelLength = Len(txt1(6))
               Exit Sub
          End Select
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
        End If
     End If
     bol911001checkRange = True
Case 8
     If Len(txt1(8)) <> 0 Then
        If Val(txt1(8)) > 12 Or Val(txt1(8)) < 1 Then
            s = MsgBox("月份輸入錯誤!!", , "USER 輸入錯誤")
            txt1(8).SetFocus
            txt1(8).SelStart = 0
            txt1(8).SelLength = Len(txt1(8))
            Exit Sub
        End If
     End If
Case Else
End Select
End Sub
