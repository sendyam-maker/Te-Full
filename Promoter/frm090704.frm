VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090704 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員作業天數統計查詢"
   ClientHeight    =   5325
   ClientLeft      =   2340
   ClientTop       =   990
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4530
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   18
      Left            =   135
      TabIndex        =   20
      Text            =   "5"
      Top             =   4395
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   19
      Left            =   1590
      TabIndex        =   21
      Text            =   "6"
      Top             =   4395
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   20
      Left            =   135
      TabIndex        =   22
      Text            =   "6"
      Top             =   4695
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   21
      Left            =   1590
      TabIndex        =   23
      Text            =   "7"
      Top             =   4695
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   22
      Left            =   135
      TabIndex        =   24
      Text            =   "8"
      Top             =   5010
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   23
      Left            =   1590
      TabIndex        =   25
      Text            =   "999"
      Top             =   5010
      Width           =   645
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2448
      TabIndex        =   26
      Top             =   72
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3228
      TabIndex        =   27
      Top             =   72
      Width           =   1200
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   504
      Width           =   1650
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   1020
      MaxLength       =   4
      TabIndex        =   1
      Top             =   792
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   2
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   2
      Top             =   792
      Width           =   675
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   3
      Left            =   1020
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1092
      Width           =   660
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   4
      Left            =   1788
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1104
      Width           =   720
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   5
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1380
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   6
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1380
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   7
      Left            =   1020
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1680
      Width           =   900
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   8
      Left            =   135
      TabIndex        =   10
      Text            =   "0"
      Top             =   2850
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   9
      Left            =   1590
      TabIndex        =   11
      Text            =   "1"
      Top             =   2850
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   10
      Left            =   135
      TabIndex        =   12
      Text            =   "1"
      Top             =   3165
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   11
      Left            =   1590
      TabIndex        =   13
      Text            =   "2"
      Top             =   3165
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   12
      Left            =   135
      TabIndex        =   14
      Text            =   "2"
      Top             =   3465
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   13
      Left            =   1590
      TabIndex        =   15
      Text            =   "3"
      Top             =   3465
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   14
      Left            =   135
      TabIndex        =   16
      Text            =   "3"
      Top             =   3780
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   15
      Left            =   1590
      TabIndex        =   17
      Text            =   "4"
      Top             =   3780
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   16
      Left            =   135
      TabIndex        =   18
      Text            =   "4"
      Top             =   4080
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   17
      Left            =   1590
      TabIndex        =   19
      Text            =   "5"
      Top             =   4080
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   24
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1965
      Width           =   315
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   25
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2280
      Width           =   300
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Left            =   1980
      TabIndex        =   55
      Top             =   1710
      Width           =   1920
      VariousPropertyBits=   27
      Size            =   "3387;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "< X < ="
      Height          =   180
      Index           =   26
      Left            =   930
      TabIndex        =   54
      Top             =   5055
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "< X < ="
      Height          =   180
      Index           =   25
      Left            =   930
      TabIndex        =   53
      Top             =   4740
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "< X < ="
      Height          =   180
      Index           =   14
      Left            =   930
      TabIndex        =   52
      Top             =   4440
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(5-6)"
      Height          =   180
      Index           =   9
      Left            =   2265
      TabIndex        =   51
      Top             =   4425
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(6-7)"
      Height          =   180
      Index           =   6
      Left            =   2265
      TabIndex        =   50
      Top             =   4740
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(8-999)"
      Height          =   180
      Index           =   3
      Left            =   2265
      TabIndex        =   49
      Top             =   5040
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "完稿年月："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   48
      Top             =   1152
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   47
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   46
      Top             =   576
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   45
      Top             =   852
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "< X < ="
      Height          =   180
      Index           =   5
      Left            =   930
      TabIndex        =   44
      Top             =   2895
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "作業天數統計範圍："
      Height          =   180
      Index           =   7
      Left            =   60
      TabIndex        =   43
      Top             =   2610
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員："
      Height          =   180
      Index           =   8
      Left            =   60
      TabIndex        =   42
      Top             =   1728
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "< X < ="
      Height          =   180
      Index           =   10
      Left            =   930
      TabIndex        =   41
      Top             =   4125
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "< X < ="
      Height          =   180
      Index           =   11
      Left            =   930
      TabIndex        =   40
      Top             =   3825
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "< X < ="
      Height          =   180
      Index           =   12
      Left            =   930
      TabIndex        =   39
      Top             =   3510
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "< X < ="
      Height          =   180
      Index           =   13
      Left            =   930
      TabIndex        =   38
      Top             =   3210
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(0-1)"
      Height          =   180
      Index           =   15
      Left            =   2265
      TabIndex        =   37
      Top             =   2880
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(1-2)"
      Height          =   180
      Index           =   16
      Left            =   2265
      TabIndex        =   36
      Top             =   3195
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(2-3)"
      Height          =   180
      Index           =   17
      Left            =   2265
      TabIndex        =   35
      Top             =   3495
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(3-4)"
      Height          =   180
      Index           =   18
      Left            =   2265
      TabIndex        =   34
      Top             =   3810
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(4-5)"
      Height          =   180
      Index           =   19
      Left            =   2265
      TabIndex        =   33
      Top             =   4125
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(1.明細 2.統計)"
      Height          =   180
      Index           =   20
      Left            =   1365
      TabIndex        =   32
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "(1.螢幕 2.報表)"
      Height          =   180
      Index           =   21
      Left            =   1395
      TabIndex        =   31
      Top             =   2010
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "顯示方式："
      Height          =   180
      Index           =   22
      Left            =   45
      TabIndex        =   30
      Top             =   2025
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "顯示內容："
      Height          =   180
      Index           =   23
      Left            =   45
      TabIndex        =   29
      Top             =   2325
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   24
      Left            =   1860
      TabIndex        =   28
      Top             =   1452
      Width           =   2412
   End
   Begin VB.Line Line1 
      X1              =   1332
      X2              =   2217
      Y1              =   936
      Y2              =   936
   End
   Begin VB.Line Line2 
      X1              =   1380
      X2              =   2145
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line3 
      X1              =   1128
      X2              =   1578
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frm090704"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; lbl1 ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 11) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 23) As String, StrTemp7(0 To 11) As String
Dim PLeft(0 To 11) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, Seekok As Integer, k As Integer
Dim iStr(0 To 1) As Integer, iInt(0 To 1) As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(Txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!""USER 輸入錯誤")
         Txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Cheng 2002/03/21
         If PUB_CheckKeyInYYMM(Me.Txt1(3)) = -1 Then
            Me.Txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Sub
         End If
         If PUB_CheckKeyInYYMM(Me.Txt1(4)) = -1 Then
            Me.Txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Sub
         End If
         
         If Len(Txt1(3)) = 0 Or Len(Txt1(4)) = 0 Then
             s = MsgBox("完稿年月區間不可空白!!", , "USER 輸入錯誤")
             If Len(Txt1(4)) = 0 Then Txt1(4).SetFocus
             If Len(Txt1(3)) = 0 Then Txt1(3).SetFocus
             Exit Sub
         Else
             If Len(Txt1(8)) = 0 And Len(Txt1(9)) = 0 And Len(Txt1(10)) = 0 And Len(Txt1(11)) = 0 And Len(Txt1(12)) = 0 And Len(Txt1(13)) = 0 And Len(Txt1(14)) = 0 And Len(Txt1(15)) = 0 And Len(Txt1(16)) = 0 And Len(Txt1(17)) = 0 And Len(Txt1(18)) = 0 And Len(Txt1(19)) = 0 And Len(Txt1(20)) = 0 And Len(Txt1(21)) = 0 And Len(Txt1(22)) = 0 And Len(Txt1(23)) = 0 Then
                 s = MsgBox("承辦天數範圍不可空白!!", , "USER 輸入錯誤")
                 Txt1(8).SetFocus
                 Txt1(8).SelStart = 0
                 Txt1(8).SelLength = Len(Txt1(8))
                 Exit Sub
             Else
                 If Len(Txt1(24)) = 0 Then
                     s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                     Txt1(24).SetFocus
                     Exit Sub
                 Else
                     If Len(Txt1(25)) = 0 Then
                         s = MsgBox("顯示內容不可空白!!", , "USER 輸入錯誤")
                         Txt1(25).SetFocus
                         Exit Sub
                     Else
                         Screen.MousePointer = vbHourglass
                         Me.Enabled = False
                         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
                         For i = 0 To 23
                            If StrTemp99(i) <> Txt1(i) Then
                                Process
                                Exit For
                            End If
                         Next i
                         For i = 0 To 23
                            StrTemp99(i) = Txt1(i)
                         Next i
                         Process1
                         Me.Enabled = True
                         Screen.MousePointer = vbDefault
                     End If
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
cnnConnection.Execute "DELETE FROM R090704_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090704_2 WHERE ID='" & strUserNum & "' "
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
StrSQL6 = StrSQL6 + " AND SUBSTR(EP09,1,6)>=" & Val(Txt1(3)) + 191100 & " AND SUBSTR(EP09,1,6)<=" & Val(Txt1(4)) + 191100 & " "
pub_QL05 = pub_QL05 & ";" & Label1(2) & Txt1(3) & "-" & Txt1(4) 'Add By Sindy 2010/12/20
If Len(Txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " and s1.st06>='" & Txt1(5) & "' "
End If
If Len(Txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " and s1.st06<='" & Txt1(6) & "' "
End If
If Len(Txt1(5)) <> 0 Or Len(Txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(5) & "-" & Txt1(6) & Label1(24) 'Add By Sindy 2010/12/20
End If
If Len(Txt1(7)) <> 0 Then
    StrSQL6 = StrSQL6 + " and ep13='" & Txt1(7) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & Txt1(7) & lbl1 'Add By Sindy 2010/12/20
End If
'Modify By Cheng 2003/07/17
'StrSQL6 = StrSQL6 + " and ep20 is null  "
StrSQL6 = StrSQL6 + " and ep29 is null  "
CheckOC
'Modify By Cheng 2002/04/29
'若已閉卷, 則在本所案號後加"*"號
'strSQL = "select s1.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s2.st02," & SQLDate("ep14") & "," & SQLDate("ep15") & "," & SQLDate("ep18") & "," & SQLDate("cp27") & ",EP18,CP48 from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep13=s1.st01(+) and ep05=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
'92.04.03 nick add left join
'strSQL = "select s1.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s2.st02," & SQLDate("ep14") & "," & SQLDate("ep15") & "," & SQLDate("ep18") & "," & SQLDate("cp27") & ",EP18,CP48 from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep13=s1.st01(+) and ep05=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
'Modify By Cheng 2003/07/17
'strSQL = "select s1.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s2.st02," & SQLDate("ep14") & "," & SQLDate("ep15") & "," & SQLDate("ep18") & "," & SQLDate("cp27") & ",EP18,CP48 from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep13=s1.st01(+) AND S1.ST05 IN ('79','81','82','AC') and ep05=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
strSql = "select s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s2.st02," & SQLDate("ep14") & "," & SQLDate("ep15") & "," & SQLDate("ep18") & "," & SQLDate("cp27") & ",EP18,CP48 from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep13=s1.st01(+) AND S1.ST05 IN ('79','81','82','AC') and ep05=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        k = 0
        'FRM100.Show
        'FRM100.Tag = Trim(str(.RecordCount)) & "=0"
        DoEvents
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
'            If Len(strTemp(9)) <> 0 And Val(strTemp(9)) <> 0 Then
            If (Len(strTemp(9)) <> 0 And Val(strTemp(9)) <> 0) And (Len(strTemp(7)) <> 0 And Val(strTemp(7)) <> 0) Then
                strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(9))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(7)))))
            Else
                strTemp(1) = 0
            End If
            strSql = "insert into r090704_1 values ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            If Len(strTemp(9)) <> 0 Then
                If Val(Format(strTemp(9), "m")) = Val(Right(Txt1(3), 2)) Or Val(Format(strTemp(9), "m")) = Val(Right(Txt1(4), 2)) Then
                    Seekok = 1
                Else
                   Seekok = 0
                End If
            Else
                Seekok = 0
            End If
            '超過件數
            If Len(CheckStr(.Fields(11))) <> 0 And Len(CheckStr(.Fields(12))) <> 0 Then
                If Val(CheckStr(.Fields(11))) > Val(CheckStr(.Fields(12))) Then
                    iStr(0) = 1
                Else
                    iStr(0) = 0
                End If
            Else
                iStr(0) = 0
            End If
            '填入資料
            If Val(strTemp(1)) >= Val(Txt1(8)) And Val(strTemp(1)) <= Val(Txt1(9)) Then
                pub_QL05 = pub_QL05 & ";" & Label1(7) & Txt1(8) & Label1(5) & Txt1(9) 'Add By Sindy 2010/12/20
                If Len(strTemp(8)) <> 0 Then
                    strSql = "insert into r090704_2 (r108001,R108002,r108003,R108011,R108012,id) values ('" & strTemp(0) & "',1,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                    cnnConnection.Execute strSql
                End If
                If Len(strTemp(9)) <> 0 Then
                    cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108003,R108011,R108012,id) values ('" & strTemp(0) & "',2,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                End If
            Else
                If Val(strTemp(1)) >= Val(Txt1(10)) And Val(strTemp(1)) <= Val(Txt1(11)) Then
                    pub_QL05 = pub_QL05 & ";" & Label1(7) & Txt1(10) & Label1(13) & Txt1(11) 'Add By Sindy 2010/12/20
                    If Len(strTemp(8)) <> 0 Then
                        cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108004,R108011,R108012,id) values ('" & strTemp(0) & "',1,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                    End If
                    If Len(strTemp(9)) <> 0 Then
                        cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108004,R108011,R108012,id) values ('" & strTemp(0) & "',2,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                    End If
                Else
                    If Val(strTemp(1)) >= Val(Txt1(12)) And Val(strTemp(1)) <= Val(Txt1(13)) Then
                        pub_QL05 = pub_QL05 & ";" & Label1(7) & Txt1(12) & Label1(12) & Txt1(13) 'Add By Sindy 2010/12/20
                        If Len(strTemp(8)) <> 0 Then
                            cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108005,R108011,R108012,id) values ('" & strTemp(0) & "',1,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                        End If
                        If Len(strTemp(9)) <> 0 Then
                            cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108005,R108011,R108012,id) values ('" & strTemp(0) & "',2,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                        End If
                    Else
                        If Val(strTemp(1)) >= Val(Txt1(14)) And Val(strTemp(1)) <= Val(Txt1(15)) Then
                            pub_QL05 = pub_QL05 & ";" & Label1(7) & Txt1(14) & Label1(11) & Txt1(15) 'Add By Sindy 2010/12/20
                            If Len(strTemp(8)) <> 0 Then
                                cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108006,R108011,R108012,id) values ('" & strTemp(0) & "',1,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                            End If
                            If Len(strTemp(9)) <> 0 Then
                                cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108006,R108011,R108012,id) values ('" & strTemp(0) & "',2,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                            End If
                        Else
                            If Val(strTemp(1)) >= Val(Txt1(16)) And Val(strTemp(1)) <= Val(Txt1(17)) Then
                                pub_QL05 = pub_QL05 & ";" & Label1(7) & Txt1(16) & Label1(10) & Txt1(17) 'Add By Sindy 2010/12/20
                                If Len(strTemp(8)) <> 0 Then
                                    cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108007,R108011,R108012,id) values ('" & strTemp(0) & "',1,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                                End If
                                If Len(strTemp(9)) <> 0 Then
                                    cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108007,R108011,R108012,id) values ('" & strTemp(0) & "',2,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                                End If
                            Else
                                If Val(strTemp(1)) >= Val(Txt1(18)) And Val(strTemp(1)) <= Val(Txt1(19)) Then
                                    pub_QL05 = pub_QL05 & ";" & Label1(7) & Txt1(18) & Label1(14) & Txt1(19) 'Add By Sindy 2010/12/20
                                    If Len(strTemp(8)) <> 0 Then
                                        cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108008,R108011,R108012,id) values ('" & strTemp(0) & "',1,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                                    End If
                                    If Len(strTemp(9)) <> 0 Then
                                        cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108008,R108011,R108012,id) values ('" & strTemp(0) & "',2,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                                    End If
                                Else
                                    If Val(strTemp(1)) >= Val(Txt1(20)) And Val(strTemp(1)) <= Val(Txt1(21)) Then
                                        pub_QL05 = pub_QL05 & ";" & Label1(7) & Txt1(20) & Label1(25) & Txt1(21) 'Add By Sindy 2010/12/20
                                        If Len(strTemp(8)) <> 0 Then
                                            cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108009,R108011,R108012,id) values ('" & strTemp(0) & "',1,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                                        End If
                                        If Len(strTemp(9)) <> 0 Then
                                            cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108009,R108011,R108012,id) values ('" & strTemp(0) & "',2,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                                        End If
                                    Else
                                        If Val(strTemp(1)) >= Val(Txt1(22)) And Val(strTemp(1)) <= Val(Txt1(23)) Then
                                            pub_QL05 = pub_QL05 & ";" & Label1(7) & Txt1(22) & Label1(26) & Txt1(23) 'Add By Sindy 2010/12/20
                                            If Len(strTemp(8)) <> 0 Then
                                                cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108010,R108011,R108012,id) values ('" & strTemp(0) & "',1,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                                            End If
                                            If Len(strTemp(9)) <> 0 Then
                                                cnnConnection.Execute "insert into r090704_2 (r108001,R108002,r108010,R108011,R108012,id) values ('" & strTemp(0) & "',2,1," & Val(strTemp(1)) & "," & iStr(0) & ",'" & strUserNum & "') "
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            .MoveNext
            k = k + 1
            'FRM100.Tag = Trim(str(.RecordCount)) & "=" & Trim(str(K))
            'FRM100.StrMenu
            DoEvents
        Loop
    End If
End With
CheckOC
''UNLOAD FRM100
End Sub

Sub Process1()
If Val(Txt1(24)) = 1 Then
   pub_QL05 = pub_QL05 & ";" & Label1(22) & "1.螢幕" 'Add By Sindy 2010/12/20
   If Val(Txt1(25)) = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(23) & "1.明細" 'Add By Sindy 2010/12/20
      Me.Hide
      frm090704_1.Show
   Else
      pub_QL05 = pub_QL05 & ";" & Label1(23) & "2.統計" 'Add By Sindy 2010/12/20
      Me.Hide
      frm090704_2.Show
   End If
Else
   pub_QL05 = pub_QL05 & ";" & Label1(22) & "2.報表" 'Add By Sindy 2010/12/20
   If Val(Txt1(25)) = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(23) & "1.明細" 'Add By Sindy 2010/12/20
      PrintData1
   Else
      pub_QL05 = pub_QL05 & ";" & Label1(23) & "2.統計" 'Add By Sindy 2010/12/20
      PrintData2
   End If
End If
End Sub

Sub ProGetOther()
iInt(0) = 0
iInt(1) = 0
CheckOC2
strSql = "select count(*) from r090704_1 where id='" & strUserNum & "' and r106001='" & strTemp3 & "' and r106009 is not null "
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        iInt(0) = Val(CheckStr(.Fields(0)))
    End If
    strSql = "select count(*) from r090704_1 where id='" & strUserNum & "' and r106001='" & strTemp3 & "' and r106009 is not null "
    CheckOC2
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenDynamic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        iInt(1) = Val(CheckStr(.Fields(0)))
    End If
End With
CheckOC2
End Sub

Sub PrintData1()
'列印資料
'Modify By Cheng 2003/07/17
'strSQL = " SELECT r106001,r106002,r106003,r106004,r106005,r106006,r106007,r106008,r106009,r106010,r106011 FROM R090704_1 WHERE ID='" & strUserNum & "' order by r106001,r106003  "
strSql = " SELECT r106001, r106002,r106003,r106004,r106005,r106006,r106007,r106008,r106009,r106010,r106011 FROM R090704_1, Staff WHERE r106001=ST01(+) And ID='" & strUserNum & "' order by ST06, r106001,r106003  "
CheckOC
Page = 1
strTemp3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/20
        .MoveFirst
        strTemp3 = CheckStr(.Fields(0))
        ProGetOther
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp3 <> strTemp(0) Then
                ProGetOther
                ShowLine3
                PrintEnd1
                ShowLine1
                PrintEnd3
                Page = Page + 1
                strTemp3 = strTemp(0)
                Printer.NewPage
                PrintTitle1
            End If
            strTemp(0) = GetStaffName(strTemp(0), True)
            strTemp(3) = StrToStr(strTemp(3), 18)
            strTemp(4) = StrToStr(strTemp(4), 4)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 4)
            PrintDatil1
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle1
            End If
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/20
    End If
End With
ShowLine3
PrintEnd1
ShowLine1
Printer.EndDoc
CheckOC
End Sub

Sub PrintEnd3()
'列印結尾
'Modify By Cheng 2003/07/17
'strSQL = "select " & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("r108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & ",R108002 from r090704_2 where id='" & strUserNum & "' and r108001='" & strTemp3 & "' "
strSql = "select " & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("r108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & ",R108002 from r090704_2 where id='" & strUserNum & "' and r108001='" & strTemp3 & "' Group By R108002 "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        ShowLine1
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        If Val(CheckStr(.Fields(8))) = 1 Then
            Printer.Print "草圖"
        Else
            Printer.Print "墨圖"
        End If
        Printer.CurrentX = 2500
        Printer.CurrentY = iPrint
        Printer.Print Txt1(8) & "-" & Txt1(9)
        Printer.CurrentX = 3700
        Printer.CurrentY = iPrint
        Printer.Print Txt1(10) & "-" & Txt1(11)
        Printer.CurrentX = 4900
        Printer.CurrentY = iPrint
        Printer.Print Txt1(12) & "-" & Txt1(13)
        Printer.CurrentX = 6100
        Printer.CurrentY = iPrint
        Printer.Print Txt1(14) & "-" & Txt1(15)
        Printer.CurrentX = 7300
        Printer.CurrentY = iPrint
        Printer.Print Txt1(16) & "-" & Txt1(17)
        Printer.CurrentX = 8500
        Printer.CurrentY = iPrint
        Printer.Print Txt1(18) & "-" & Txt1(19)
        Printer.CurrentX = 9700
        Printer.CurrentY = iPrint
        Printer.Print Txt1(20) & "-" & Txt1(21)
        Printer.CurrentX = 10900
        Printer.CurrentY = iPrint
        Printer.Print Txt1(22) & "-" & Txt1(23)
        iPrint = iPrint + 300
        ShowLine3
        Printer.CurrentX = 3000 - Printer.TextWidth(Format(CheckStr(.Fields(0)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(0)), "####0")
        Printer.CurrentX = 4200 - Printer.TextWidth(Format(CheckStr(.Fields(1)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(1)), "####0")
        Printer.CurrentX = 5400 - Printer.TextWidth(Format(CheckStr(.Fields(2)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(2)), "####0")
        Printer.CurrentX = 6600 - Printer.TextWidth(Format(CheckStr(.Fields(3)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(3)), "####0")
        Printer.CurrentX = 7800 - Printer.TextWidth(Format(CheckStr(.Fields(4)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(4)), "####0")
        Printer.CurrentX = 9000 - Printer.TextWidth(Format(CheckStr(.Fields(5)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(5)), "####0")
        Printer.CurrentX = 10200 - Printer.TextWidth(Format(CheckStr(.Fields(6)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(6)), "####0")
        Printer.CurrentX = 11400 - Printer.TextWidth(Format(CheckStr(.Fields(7)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(7)), "####0")
        iPrint = iPrint + 300
        ShowLine1
    End If
End With
CheckOC2
End Sub

Sub PrintData2()
'列印資料2
'strSQL = "select r108001,decode(r108001,1,'草圖','墨圖')," & sqlsum("r108003") & "," & sqlsum("r108004") & "," & sqlsum("r108005") & "," & sqlsum("r108006") & "," & sqlsum("r108007") & "," & sqlsum("r108008") & "," & sqlsum("r108009") & "," & sqlsum("r108010") &  from r090704_2 where id='" & strUserNum & "' group by r108001 order by r108001 "
'strSQL = "select r108001,decode(r108002,1,'草圖','墨圖')," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & "," & SQLSum("R108011") & "/" & SQLSum("r108003") & "+" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & "," & SQLSum("R108012") & " from r090704_2 where id='" & strUserNum & "' group by r108001,decode(r108002,1,'草圖','墨圖') order by r108001,decode(r108002,1,'草圖','墨圖') "
'strSQL = "select r108001,decode(r108002,1,'草圖','墨圖')," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & "," & SQLSum("R108011") & "/" & SQLSum("r108003") & "+" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & "," & SQLSum("R108012") & ", ST06 from r090704_2, Staff where r108001=ST01(+) And id='" & strUserNum & "' group by ST06, r108001,decode(r108002,1,'草圖','墨圖') order by ST06, r108001,decode(r108002,1,'草圖','墨圖') "

'Modify by Morgan 2004/5/4
'strSQL = "select r108001,decode(r108002,1,'草圖','墨圖')," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & "," & SQLSum("R108011") & "/Nvl(Sum(r108003),1) +" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & "," & SQLSum("R108012") & ", ST06 from r090704_2, Staff where r108001=ST01(+) And id='" & strUserNum & "' group by ST06, r108001,decode(r108002,1,'草圖','墨圖') order by ST06, r108001,decode(r108002,1,'草圖','墨圖') "
strSql = "select r108001,decode(r108002,1,'草圖','墨圖')," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & "," & SQLSum("R108011") & "/(" & SQLSum("r108003") & "+" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & ")," & SQLSum("R108012") & ", ST06 from r090704_2, Staff where r108001=ST01(+) And id='" & strUserNum & "' group by ST06, r108001,decode(r108002,1,'草圖','墨圖') order by ST06, r108001,decode(r108002,1,'草圖','墨圖') "
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/20
        .MoveFirst
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 11
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(0) = GetStaffName(strTemp(0), True)
            strTemp(0) = StrToStr(strTemp(0), 10)
            PrintDatil2
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle2
            End If
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/20
    End If
End With
ShowLine3
PrintEnd2
ShowLine2
Printer.EndDoc
CheckOC
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Txt1(0) = Systemkind_g_P
For i = 0 To 23
    StrTemp99(i) = ""
Next i
Txt1(24) = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090704 = Nothing
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
      'Add By Cheng 2002/01/07
      Me.Txt1(Index).Text = GetAllSysKind(Me.Txt1(Index))
     strTemp1 = Split(UCase(Systemkind_g_P), ",")
     strTemp2 = Split(UCase(Txt1(0)), ",")
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
            Txt1(0).SetFocus
            Txt1(0).SelStart = 0
            Txt1(0).SelLength = Len(Txt1(0))
            Exit Sub
        End If
     Next i
Case 5
     Select Case Trim(Txt1(5))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(5).SetFocus
          Txt1(5).SelStart = 0
          Txt1(5).SelLength = Len(Txt1(5))
          Exit Sub
     End Select
Case 6
     Select Case Trim(Txt1(6))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(6).SetFocus
          Txt1(6).SelStart = 0
          Txt1(6).SelLength = Len(Txt1(6))
          Exit Sub
     End Select
Case 7
     lbl1 = GetPrjSales(Txt1(7))
Case 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23
     For i = 1 To Len(Txt1(Index))
          strSql = Mid(Txt1(Index), i, 1)
          If InStr(1, "0123456789 ", strSql) = 0 Then
                s = MsgBox("承辦天數只能輸入數字!!", , "USER 輸入錯誤")
                Txt1(Index).SetFocus
                Txt1(Index).SelStart = 0
                Txt1(Index).SelLength = Len(Txt1(Index))
                Exit Sub
          End If
     Next i
Case 24
     Select Case Trim(Txt1(24))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(24).SetFocus
          Txt1(24).SelStart = 0
          Txt1(24).SelLength = Len(Txt1(24))
          Exit Sub
     End Select
Case 25
     Select Case Trim(Txt1(25))
     Case "1", "2", ""
     Case Else
          s = MsgBox("查詢對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(25).SetFocus
          Txt1(25).SelStart = 0
          Txt1(25).SelLength = Len(Txt1(25))
          Exit Sub
     End Select
Case Else
End Select
End Sub

Sub PrintTitle1()
'列印抬頭
GetPleft1
iPrint = 0
'Printer.Orientation = 1
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "繪圖人員作業天數明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "完稿年月：" & Mid(Txt1(3), 1, 2) & "/" & Mid(Txt1(3), 3, 2) & "－" & Mid(Txt1(4), 1, 2) & "/" & Mid(Txt1(4), 3, 2)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "繪圖人員：" & strTemp3
Printer.CurrentX = 3500
Printer.CurrentY = iPrint
Printer.Print "完稿草圖: " & Trim(str(iInt(0))) & " 件,墨圖: " & Trim(str(iInt(1))) & " 件 "
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "作業天數"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "草圖齊備日"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "草圖完稿日"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "墨圖完稿日"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "發文日"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
    Exit Sub
End If
End Sub

Sub PrintDatil1() '列印資料
Printer.CurrentX = PLeft(1) + 500 - Printer.TextWidth(Format(strTemp(1), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(1), "####0")
For i = 2 To 10
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft1()
'定陣列位置
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = 1000
PLeft(3) = 3300
PLeft(4) = 7800
PLeft(5) = 8800
PLeft(6) = 10000
PLeft(7) = 11000
PLeft(8) = 12000
PLeft(9) = 13000
PLeft(10) = 14000
End Sub

Sub ShowLine1()
'畫線
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
End If
End Sub

Sub ShowLine2()
'畫線
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
End If
End Sub

Sub ShowLine3()
'畫2線
Printer.CurrentX = 0
Printer.CurrentY = iPrint
For i = 120 To 125
    Printer.Line (0, iPrint + i)-(15000, iPrint + i)
    Printer.Line (0, iPrint + 50 + i)-(15000, iPrint + 50 + i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintEnd1()
'列印結尾
strSql = "SELECT '合  計','',SUM(R106002) from r090704_1 WHERE ID='" & strUserNum & "' AND R106001='" & strTemp3 & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        Printer.Print CheckStr(.Fields(0))
        Printer.CurrentX = PLeft(2) + 500 - Printer.TextWidth(Format(CheckStr(.Fields(1)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(1)), "####0")
        iPrint = iPrint + 300
    End If
End With
CheckOC2
End Sub

Sub PrintTitle2()
'列抬頭  地2張
GetPleft2
iPrint = 0
'Printer.Orientation = 1
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "繪圖人員作業天數統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "完稿年月：" & Mid(Txt1(3), 1, 2) & "/" & Mid(Txt1(3), 3, 2) & "－" & Mid(Txt1(4), 1, 2) & "/" & Mid(Txt1(4), 3, 2)
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
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "繪圖人員"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "當月承辦"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print Txt1(8) & "-" & Txt1(9)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print Txt1(10) & "-" & Txt1(11)
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print Txt1(12) & "-" & Txt1(13)
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print Txt1(14) & "-" & Txt1(15)
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print Txt1(16) & "-" & Txt1(17)
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print Txt1(18) & "-" & Txt1(19)
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print Txt1(20) & "-" & Txt1(21)
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print Txt1(22) & "-" & Txt1(23)
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "平均天數"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "超過天數"
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
End Sub

Sub PrintDatil2()
'列印資料
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 11
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0")
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft2()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1500
PLeft(2) = 2700
PLeft(3) = 3900
PLeft(4) = 5100
PLeft(5) = 6300
PLeft(6) = 7500
PLeft(7) = 8700
PLeft(8) = 9900
PLeft(9) = 11100
PLeft(10) = 12300
PLeft(11) = 14500
End Sub

Sub PrintEnd2()
'列印結尾
'strSQL = "select '合  計',decode(sum(r108003),null,0,sum(r108003))+decode(sum(r108004),null,0,sum(r108004))+decode(sum(r108005),null,0,sum(r108005))+decode(sum(r108006),null,0,sum(r108006))+decode(sum(r108007),null,0,sum(r108007)),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007) from r090704_2 where id='" & strUserNum & "' "

'Modify by Morgan 2004/5/4
'strSQL = "select '合  計',''," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & "," & SQLSum("R108011") & "/" & SQLSum("r108003") & "+" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & "," & SQLSum("R108012") & " from r090704_2 where id='" & strUserNum & "' "
strSql = "select '合  計',''," & SQLSum("r108003") & "," & SQLSum("r108004") & "," & SQLSum("r108005") & "," & SQLSum("r108006") & "," & SQLSum("r108007") & "," & SQLSum("R108008") & "," & SQLSum("R108009") & "," & SQLSum("R108010") & "," & SQLSum("R108011") & "/(" & SQLSum("r108003") & "+" & SQLSum("r108004") & "+" & SQLSum("r108005") & "+" & SQLSum("r108006") & "+" & SQLSum("r108007") & "+" & SQLSum("R108008") & "+" & SQLSum("R108009") & "+" & SQLSum("R108010") & ")," & SQLSum("R108012") & " from r090704_2 where id='" & strUserNum & "' "

CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 11
                StrTemp7(i) = CheckStr(.Fields(i))
                If Len(StrTemp7(i)) = 0 And i > 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 11
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "####0")
            Next i
            iPrint = iPrint + 300
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle2
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 3, 4 '完稿年月
   If PUB_CheckKeyInYYMM(Me.Txt1(Index)) = -1 Then
      Cancel = True
      Me.Txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub
