VERSION 5.00
Begin VB.Form frm090703 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員每日分案情形查詢"
   ClientHeight    =   2160
   ClientLeft      =   105
   ClientTop       =   1305
   ClientWidth     =   4425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4425
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1935
      MaxLength       =   4
      TabIndex        =   2
      Top             =   860
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   1
      Top             =   860
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1170
      TabIndex        =   0
      Top             =   480
      Width           =   2685
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1240
      Width           =   300
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1240
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1950
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1620
      Width           =   465
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1620
      Width           =   405
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3180
      TabIndex        =   8
      Top             =   24
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2400
      TabIndex        =   7
      Top             =   24
      Width           =   756
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   2190
      Y1              =   1008
      Y2              =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   920
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   540
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   7
      Left            =   1944
      TabIndex        =   13
      Top             =   1300
      Width           =   2412
   End
   Begin VB.Line Line3 
      X1              =   1296
      X2              =   1761
      Y1              =   1383
      Y2              =   1383
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   12
      Top             =   1300
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "文齊年月："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "年"
      Height          =   180
      Index           =   0
      Left            =   1635
      TabIndex        =   10
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label Label2 
      Caption         =   "月"
      Height          =   180
      Index           =   1
      Left            =   2505
      TabIndex        =   9
      Top             =   1680
      Width           =   360
   End
End
Attribute VB_Name = "frm090703"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/07 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 32) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 32) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, SeekPrint As Integer, SeekPrintL As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(Txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         Txt1(0).SetFocus
         Exit Sub
     Else
         If Len(Txt1(5)) = 0 Or Len(Txt1(6)) = 0 Then
             s = MsgBox("收文年月不可空白!!", , "USER 輸入錯誤")
             If Len(Txt1(6)) = 0 Then Txt1(6).SetFocus
             If Len(Txt1(5)) = 0 Then Txt1(5).SetFocus
             Exit Sub
         Else
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
             Process
             Me.Hide
             frm090703_1.Show
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
cnnConnection.Execute "DELETE FROM R090703 WHERE ID='" & strUserNum & "'  "
strSQL1 = ""
If Len(Txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and cp01 in (" & SQLGrpStr(Txt1(0), 1) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/12/20
End If
StrSQL6 = ""
If Len(Txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & Txt1(1) & "' "
End If
If Len(Txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & Txt1(2) & "' "
End If
If Len(Txt1(1)) <> 0 Or Len(Txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(1) & "-" & Txt1(2) 'Add By Sindy 2010/12/20
End If
If Len(Txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06>='" & Txt1(3) & "' "
End If
If Len(Txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06<='" & Txt1(4) & "' "
End If
If Len(Txt1(3)) <> 0 Or Len(Txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(3) & "-" & Txt1(4) & Label1(7) 'Add By Sindy 2010/12/20
End If
If Len(Txt1(5)) <> 0 Then
    'Modify By Cheng 2004/03/23
'    StrSQL6 = StrSQL6 + " AND SUBSTR(CP05,1,4)=" & Val(txt1(5)) + 1911
    StrSQL6 = StrSQL6 + " AND SUBSTR(EP06,1,4)=" & Val(Txt1(5)) + 1911
    'End
End If
If Len(Txt1(6)) <> 0 Then
    'Modify By Cheng 2004/03/23
'    StrSQL6 = StrSQL6 + " AND SUBSTR(CP05,5,2)=" & Val(txt1(6))
    StrSQL6 = StrSQL6 + " AND SUBSTR(EP06,5,2)=" & Val(Txt1(6))
    'End
End If
If Len(Txt1(5)) <> 0 Or Len(Txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & Txt1(5) & Label2(0) & Txt1(6) & Label2(1) 'Add By Sindy 2010/12/20
End If
'Modify By Cheng 2003/07/17
'StrSQL6 = StrSQL6 + " and ep20 is null  "
StrSQL6 = StrSQL6 + " and ep29 is null  "
'92.04.03 nick add left join
'strSQL = "select st02,cp05 from caseprogress,patent,staff,engineerprogress where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep13=st01(+) and cp05 is not null and cp05>0 " & StrSQL6 & strSQL1
'Modify By Cheng 2003/07/17
'strSQL = "select st02,cp05 from caseprogress,patent,staff,engineerprogress where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep13=st01(+) AND ST05 IN ('79','81','82','AC') and cp05 is not null and cp05>0 " & StrSQL6 & strSQL1
'Modify By Cheng 2004/04/07
'收文日改為文齊日
'strSQL = "select st01,cp05 from caseprogress,patent,staff,engineerprogress where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep13=st01(+) AND ST05 IN ('79','81','82','AC') and cp05 is not null and cp05>0 " & StrSQL6 & strSQL1
strSql = "select st01,EP06 from caseprogress,patent,staff,engineerprogress where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep13=st01(+) AND ST05 IN ('79','81','82','AC') and EP06 is not null and EP06>0 " & StrSQL6 & strSQL1
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            cnnConnection.Execute "insert into r090703 (r105001,r1050" & Format(Val(Right(CheckStr(.Fields(1)), 2)) + 1, "00") & ",id) values ('" & CheckStr(.Fields(0)) & "',1,'" & strUserNum & "') "
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'Modify By Cheng 2004/03/08
'txt1(0) = Systemkind_g_P
Txt1(0) = "P,PS,CFP,CPS,FCP,FG,CFC"
'End
'Add By Cheng 2004/02/18
'收文年月預設系統年月
Me.Txt1(5).Text = Left(strSrvDate(1), 4) - 1911
Me.Txt1(6).Text = Val(Mid(strSrvDate(1), 5, 2))
'End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090703 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
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
        'Marked By Cheng 2004/03/08
'      'Add By Cheng 2002/01/07
'      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
'        'Modify By Cheng 2004/03/08
''     strTemp1 = Split(UCase(Systemkind_g_P), ",")
'     strTemp1 = Split(UCase(Systemkind_g), ",")
'        'End
'     strTemp2 = Split(UCase(txt1(0)), ",")
'     For i = 0 To UBound(strTemp2)
'        s = 0
'        For j = 0 To UBound(strTemp1)
'            If strTemp2(i) = strTemp1(j) Then
'                s = 1
'                Exit For
'            End If
'        Next j
'        If s = 0 Then
'            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
'            txt1(0).SetFocus
'            txt1(0).SelStart = 0
'            txt1(0).SelLength = Len(txt1(0))
'            Exit Sub
'        End If
'     Next i
        'End
Case 3
     Select Case Trim(Txt1(3))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(3).SetFocus
          Txt1(3).SelStart = 0
          Txt1(3).SelLength = Len(Txt1(3))
          Exit Sub
     End Select
Case 4
     Select Case Trim(Txt1(4))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(4).SetFocus
          Txt1(4).SelStart = 0
          Txt1(4).SelLength = Len(Txt1(4))
          Exit Sub
     End Select
Case 6
     If Len(Txt1(6)) <> 0 Then
        If Val(Txt1(6)) > 12 Or Val(Txt1(6)) < 1 Then
            s = MsgBox("月份輸入錯誤!!", , "USER 輸入錯誤")
            Txt1(6).SetFocus
            Txt1(6).SelStart = 0
            Txt1(6).SelLength = Len(Txt1(6))
            Exit Sub
        End If
     End If
Case Else
End Select
End Sub


