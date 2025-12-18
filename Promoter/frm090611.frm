VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090611 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦天數統計查詢"
   ClientHeight    =   5940
   ClientLeft      =   1950
   ClientTop       =   1710
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   4380
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   23
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   17
      Top             =   3645
      Width           =   300
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   22
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   16
      Top             =   3315
      Width           =   315
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   11
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2985
      Width           =   900
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   10
      Left            =   1752
      MaxLength       =   3
      TabIndex        =   14
      Top             =   2670
      Width           =   525
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   9
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   13
      Top             =   2670
      Width           =   480
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   8
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   12
      Top             =   2355
      Width           =   900
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   7
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2070
      Width           =   315
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   6
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1785
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   5
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1785
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   2
      Left            =   1812
      MaxLength       =   4
      TabIndex        =   2
      Top             =   780
      Width           =   675
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   1
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   1
      Top             =   780
      Width           =   660
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   468
      Width           =   1650
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3084
      TabIndex        =   29
      Top             =   24
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2304
      TabIndex        =   28
      Top             =   24
      Width           =   756
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   675
      Left            =   48
      TabIndex        =   54
      Top             =   1110
      Width           =   4212
      Begin VB.TextBox Txt1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   25
         Left            =   2328
         MaxLength       =   7
         TabIndex        =   8
         Top             =   345
         Width           =   810
      End
      Begin VB.TextBox Txt1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   24
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   7
         Top             =   345
         Width           =   810
      End
      Begin VB.OptionButton Option1 
         Caption         =   "完稿日期："
         Height          =   204
         Index           =   1
         Left            =   48
         TabIndex        =   6
         Top             =   360
         Width           =   1212
      End
      Begin VB.OptionButton Option1 
         Caption         =   "會稿日期："
         Height          =   204
         Index           =   0
         Left            =   48
         TabIndex        =   3
         Top             =   48
         Value           =   -1  'True
         Width           =   1212
      End
      Begin VB.TextBox Txt1 
         Height          =   300
         Index           =   3
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   4
         Top             =   48
         Width           =   810
      End
      Begin VB.TextBox Txt1 
         Height          =   300
         Index           =   4
         Left            =   2328
         MaxLength       =   7
         TabIndex        =   5
         Top             =   48
         Width           =   810
      End
      Begin VB.Line Line5 
         X1              =   1770
         X2              =   2535
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line2 
         X1              =   1764
         X2              =   2529
         Y1              =   192
         Y2              =   192
      End
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   12
      Left            =   45
      TabIndex        =   18
      Text            =   "0"
      Top             =   4200
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   13
      Left            =   1700
      TabIndex        =   19
      Text            =   "7"
      Top             =   4185
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   14
      Left            =   45
      TabIndex        =   20
      Text            =   "8"
      Top             =   4530
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   15
      Left            =   1700
      TabIndex        =   21
      Text            =   "14"
      Top             =   4515
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   16
      Left            =   45
      TabIndex        =   22
      Text            =   "15"
      Top             =   4845
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   17
      Left            =   1700
      TabIndex        =   23
      Text            =   "21"
      Top             =   4830
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   18
      Left            =   45
      TabIndex        =   24
      Text            =   "22"
      Top             =   5190
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   19
      Left            =   1700
      TabIndex        =   25
      Text            =   "28"
      Top             =   5175
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   20
      Left            =   45
      TabIndex        =   26
      Text            =   "29"
      Top             =   5505
      Width           =   645
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   21
      Left            =   1700
      TabIndex        =   27
      Text            =   "999"
      Top             =   5490
      Width           =   645
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   56
      Top             =   3000
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   2040
      TabIndex        =   55
      Top             =   2340
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line4 
      X1              =   1395
      X2              =   2115
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line Line3 
      X1              =   1185
      X2              =   1635
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line1 
      X1              =   1350
      X2              =   2235
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   24
      Left            =   1860
      TabIndex        =   53
      Top             =   1830
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "顯示內容："
      Height          =   180
      Index           =   23
      Left            =   15
      TabIndex        =   52
      Top             =   3690
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "顯示方式："
      Height          =   180
      Index           =   22
      Left            =   15
      TabIndex        =   51
      Top             =   3375
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "(1.螢幕 2.報表)"
      Height          =   180
      Index           =   21
      Left            =   1440
      TabIndex        =   50
      Top             =   3345
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "(1.明細 2.統計)"
      Height          =   180
      Index           =   20
      Left            =   1410
      TabIndex        =   49
      Top             =   3675
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.承辦人 2.智權人員)"
      Height          =   180
      Index           =   14
      Left            =   1530
      TabIndex        =   43
      Top             =   2130
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   9
      Left            =   15
      TabIndex        =   38
      Top             =   2715
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   8
      Left            =   15
      TabIndex        =   37
      Top             =   2415
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   6
      Left            =   15
      TabIndex        =   35
      Top             =   3030
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   15
      TabIndex        =   33
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   15
      TabIndex        =   32
      Top             =   504
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   15
      TabIndex        =   31
      Top             =   1845
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "查詢對象："
      Height          =   180
      Index           =   3
      Left            =   15
      TabIndex        =   30
      Top             =   2115
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "(29-999)"
      Height          =   180
      Index           =   19
      Left            =   2415
      TabIndex        =   48
      Top             =   5550
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(22-28)"
      Height          =   180
      Index           =   18
      Left            =   2415
      TabIndex        =   47
      Top             =   5235
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(15-21)"
      Height          =   180
      Index           =   17
      Left            =   2415
      TabIndex        =   46
      Top             =   4905
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(8-14)"
      Height          =   180
      Index           =   16
      Left            =   2415
      TabIndex        =   45
      Top             =   4575
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "(0-7)"
      Height          =   180
      Index           =   15
      Left            =   2415
      TabIndex        =   44
      Top             =   4230
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "< = X < ="
      Height          =   180
      Index           =   13
      Left            =   825
      TabIndex        =   42
      Top             =   4575
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "< = X < ="
      Height          =   180
      Index           =   12
      Left            =   825
      TabIndex        =   41
      Top             =   4890
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "< = X < ="
      Height          =   180
      Index           =   11
      Left            =   825
      TabIndex        =   40
      Top             =   5235
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "< = X < ="
      Height          =   180
      Index           =   10
      Left            =   825
      TabIndex        =   39
      Top             =   5550
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "承辦天數統計範圍："
      Height          =   180
      Index           =   7
      Left            =   15
      TabIndex        =   36
      Top             =   3975
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "< = X < ="
      Height          =   180
      Index           =   5
      Left            =   825
      TabIndex        =   34
      Top             =   4230
      Width           =   765
   End
End
Attribute VB_Name = "frm090611"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ; lbl1(index) ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 10) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String, StrTemp7(0 To 10) As String, StrSQL3 As String
Dim PLeft(0 To 10) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, Seekok As Integer, k As Integer
Dim bolData As Boolean
Dim bol911001checkRange As Boolean

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0 '確定
     If Len(Txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!""USER 輸入錯誤")
         Txt1(0).SetFocus
         Exit Sub
     Else
        '若選擇會稿日期
        If Me.Option1(0).Value Then
             'Add By Cheng 2002/03/21
    '         If PUB_CheckKeyInYYMM(Me.txt1(3)) = -1 Then
             If PUB_CheckKeyInDate(Me.Txt1(3)) = -1 Then
                Me.Txt1(3).SetFocus
                txt1_GotFocus 3
                Exit Sub
             End If
    '         If PUB_CheckKeyInYYMM(Me.txt1(4)) = -1 Then
             If PUB_CheckKeyInDate(Me.Txt1(4)) = -1 Then
                Me.Txt1(4).SetFocus
                txt1_GotFocus 4
                Exit Sub
             End If
        '若選擇完稿日期
        Else
             If PUB_CheckKeyInDate(Me.Txt1(24)) = -1 Then
                Me.Txt1(24).SetFocus
                txt1_GotFocus 24
                Exit Sub
             End If
             If PUB_CheckKeyInDate(Me.Txt1(25)) = -1 Then
                Me.Txt1(25).SetFocus
                txt1_GotFocus 25
                Exit Sub
             End If
        End If
'         If Len(Txt1(3)) = 0 Or Len(Txt1(4)) = 0 Then
         If (Me.Option1(0).Value And (Len(Txt1(3)) = 0 Or Len(Txt1(4)) = 0)) Or (Me.Option1(1).Value And (Len(Txt1(24)) = 0 Or Len(Txt1(25)) = 0)) Then
            If Me.Option1(0).Value Then
                    'Modify By Cheng 2003/06/05
    '             s = MsgBox("會稿年月區間不可空白!!", , "USER 輸入錯誤")
                 s = MsgBox("會稿日期區間不可空白!!", , "USER 輸入錯誤")
                 If Len(Txt1(4)) = 0 Then Txt1(4).SetFocus: txt1_GotFocus 4
                 If Len(Txt1(3)) = 0 Then Txt1(3).SetFocus: txt1_GotFocus 3
                 Exit Sub
            Else
                 s = MsgBox("完稿日期區間不可空白!!", , "USER 輸入錯誤")
                 If Len(Txt1(24)) = 0 Then Txt1(24).SetFocus: txt1_GotFocus 24
                 If Len(Txt1(25)) = 0 Then Txt1(25).SetFocus: txt1_GotFocus 25
                 Exit Sub
            End If
         Else
             If Len(Txt1(7)) = 0 Then
                 s = MsgBox("查詢對象不可空白!!", , "USER 輸入錯誤")
                 Txt1(7).SetFocus
                 Exit Sub
             Else
                 If Len(Txt1(12)) = 0 And Len(Txt1(13)) = 0 And Len(Txt1(14)) = 0 And Len(Txt1(15)) = 0 And Len(Txt1(16)) = 0 And Len(Txt1(17)) = 0 And Len(Txt1(18)) = 0 And Len(Txt1(19)) = 0 And Len(Txt1(20)) = 0 And Len(Txt1(21)) = 0 Then
                     s = MsgBox("承辦天數範圍不可空白!!", , "USER 輸入錯誤")
                     Txt1(12).SetFocus
                     Txt1(12).SelStart = 0
                     Txt1(12).SelLength = Len(Txt1(12))
                     Exit Sub
                 Else
                     If Len(Txt1(22)) = 0 Then
                         s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                         Txt1(22).SetFocus
                         Exit Sub
                     Else
                         If Len(Txt1(23)) = 0 Then
                             s = MsgBox("顯示內容不可空白!!", , "USER 輸入錯誤")
                             Txt1(23).SetFocus
                             Exit Sub
                         Else
                             Screen.MousePointer = vbHourglass
                             Me.Enabled = False
                             ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/17 清除查詢印表記錄檔欄位
                             'For i = 0 To 21
                             '   If StrTemp99(i) <> txt1(i) Then
                                    Process
                             '       Exit For
                             '   End If
                             'Next i
                             For i = 0 To 21
                                StrTemp99(i) = Txt1(i)
                             Next i
                             If bolData = True Then
                                 Process1
                             End If
                             Me.Enabled = True
                             Screen.MousePointer = vbDefault
                         End If
                     End If
                 End If
             End If
         End If
     End If
Case 1 '回前畫面
     Unload Me
Case Else
End Select
End Sub

Sub Process()
bolData = False
cnnConnection.Execute "DELETE FROM R090611_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090611_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""

If Len(Txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(Txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(Txt1(0), 2) & ") "
   StrSQL3 = StrSQL3 + " and CP01 in (" & SQLGrpStr(Txt1(0), 3) & ") "
   StrSQL4 = StrSQL4 + " and CP01 in (" & SQLGrpStr(Txt1(0), 4) & ") "
   strSQL5 = strSQL5 + " and CP01 in (" & SQLGrpStr(Txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/12/17
End If
StrSQL6 = ""
If Len(Txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & Txt1(1) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & Txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & Txt1(1) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & Txt1(1) & "' "
End If
If Len(Txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & Txt1(2) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & Txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & Txt1(2) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & Txt1(2) & "' "
End If
If Len(Txt1(1)) <> 0 Or Len(Txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(1) & "-" & Txt1(2) 'Add By Sindy 2010/12/17
End If
'若選擇會稿日期
If Me.Option1(0).Value Then
    'Modify By Cheng 2003/06/05
    'StrSQL6 = StrSQL6 + " AND EP07>=" & Val(Txt1(3)) + 191100 & "01 AND EP07<=" & Val(Txt1(4)) + 191100 & "31 "
    StrSQL6 = StrSQL6 + " AND EP07>=" & Val(Txt1(3)) + 19110000 & " AND EP07<=" & Val(Txt1(4)) + 19110000 & " "
    pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Txt1(3) & "-" & Txt1(4) 'Add By Sindy 2010/12/17
'若選擇完稿日期
Else
    StrSQL6 = StrSQL6 + " AND EP09>=" & Val(Txt1(24)) + 19110000 & " AND EP09<=" & Val(Txt1(25)) + 19110000 & " "
    pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Txt1(24) & "-" & Txt1(25) 'Add By Sindy 2010/12/17
End If
If Val(Txt1(7)) = 1 Then
    If Len(Txt1(5)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06>='" & Txt1(5) & "' "
    End If
    If Len(Txt1(6)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06<='" & Txt1(6) & "' "
    End If
Else
    If Len(Txt1(5)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s2.st06>='" & Txt1(5) & "' "
    End If
    If Len(Txt1(6)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s2.st06<='" & Txt1(6) & "' "
    End If
End If
If Len(Txt1(5)) <> 0 Or Len(Txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(5) & "-" & Txt1(6) & Label1(24) 'Add By Sindy 2010/12/17
End If
If Len(Txt1(8)) <> 0 Then
    StrSQL6 = StrSQL6 + " and ep05='" & Txt1(8) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & Txt1(8) & lbl1(0) 'Add By Sindy 2010/12/17
End If
If Len(Txt1(9)) <> 0 Then
    StrSQL6 = StrSQL6 + " and cp12>='" & Txt1(9) & "' "
End If
If Len(Txt1(10)) <> 0 Then
    StrSQL6 = StrSQL6 + " and cp12<='" & Txt1(10) & "' "
End If
If Len(Txt1(9)) <> 0 Or Len(Txt1(10)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(9) & Txt1(9) & "-" & Txt1(10) 'Add By Sindy 2010/12/17
End If
If Len(Txt1(11)) <> 0 Then
    StrSQL6 = StrSQL6 + " and cp13='" & Txt1(11) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(11) & lbl1(1) 'Add By Sindy 2010/12/17
End If
StrSQL6 = StrSQL6 + " and CP26 IS NULL "
CheckOC

'Modify By Sindy 2012/5/31
pub_QL05 = pub_QL05 & ";" & Label1(7)
If Txt1(12) <> "" Or Txt1(13) <> "" Then
   pub_QL05 = pub_QL05 & Txt1(12) & "-" & Txt1(13) & ","
End If
If Txt1(14) <> "" Or Txt1(15) <> "" Then
   pub_QL05 = pub_QL05 & Txt1(14) & "-" & Txt1(15) & ","
End If
If Txt1(16) <> "" Or Txt1(17) <> "" Then
   pub_QL05 = pub_QL05 & Txt1(16) & "-" & Txt1(17) & ","
End If
If Txt1(18) <> "" Or Txt1(19) <> "" Then
   pub_QL05 = pub_QL05 & Txt1(18) & "-" & Txt1(19) & ","
End If
If Txt1(20) <> "" Or Txt1(21) <> "" Then
   pub_QL05 = pub_QL05 & Txt1(20) & "-" & Txt1(21) & ","
End If
If Right(pub_QL05, 1) = "," Then
   pub_QL05 = Left(pub_QL05, Len(pub_QL05) - 1)
End If

'查詢承辦人
If Val(Txt1(7)) = 1 Then
   'Modify By Cheng 2002/04/26
   '若已閉卷, 則在本所案號後加"*"號
    'Modify By Cheng 2003/06/10
'    strSQL = "select s1.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s2.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s1.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),s2.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,trademARK,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s1.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),s2.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,lawcase where EP02=CP09(+) AND  cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s1.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,s2.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,hirecase where EP02=CP09(+) AND  cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s1.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),s2.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,servicepractice,staff s1,staff s2,casepropertymap where EP02=CP09(+) AND  cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL5 & StrSQL6
    'edit by nick 2004/08/02 fcp 因為沒有文件齊備日，所以用收文日去算工作天數
'    strSQL = "select s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,trademARK,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,lawcase where EP02=CP09(+) AND  cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,hirecase where EP02=CP09(+) AND  cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,servicepractice,staff s1,staff s2,casepropertymap where EP02=CP09(+) AND  cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL5 & StrSQL6
    'modify by sonia 2017/11/13 +s1.st03承辦人的部門
    strSql = "select s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
    strSql = strSql + " UNION all  SELECT s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,trademARK,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) " & strSQL2 & StrSQL6
    strSql = strSql + " UNION all  SELECT s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,lawcase where EP02=CP09(+) AND  cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL6
    strSql = strSql + " UNION all  SELECT s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,hirecase where EP02=CP09(+) AND  cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL6
    strSql = strSql + " UNION all  SELECT s1.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),s2.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,servicepractice,staff s1,staff s2,casepropertymap where EP02=CP09(+) AND  cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL5 & StrSQL6
'查詢智權人員
Else
   'Modify By Cheng 2002/04/26
   '若已閉卷, 則在本所案號後加"*"號
    'Modify By Cheng 2003/06/10
'    strSQL = "select s2.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s2.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),s1.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,trademARK,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s2.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),s1.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,lawcase where EP02=CP09(+) AND  cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s2.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,s1.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,hirecase where EP02=CP09(+) AND  cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s2.st02,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),s1.st02," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,servicepractice,staff s1,staff s2,casepropertymap where EP02=CP09(+) AND  cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL5 & StrSQL6
    'edit by nick 2004/08/02 fcp 因為沒有文件齊備日，所以用收文日去算工作天數
'    strSQL = "select s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,trademARK,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,lawcase where EP02=CP09(+) AND  cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,hirecase where EP02=CP09(+) AND  cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & " from engineerprogress,caseprogress,servicepractice,staff s1,staff s2,casepropertymap where EP02=CP09(+) AND  cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL5 & StrSQL6
    'modify by sonia 2017/11/13 +s1.st03承辦人的部門
    strSql = "select s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,patent,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) " & strSQL1 & StrSQL6
    strSql = strSql + " UNION all  SELECT s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,trademARK,staff s1,staff s2,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) " & strSQL2 & StrSQL6
    strSql = strSql + " UNION all  SELECT s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,lawcase where EP02=CP09(+) AND  cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & StrSQL6
    strSql = strSql + " UNION all  SELECT s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,staff s1,staff s2,casepropertymap,hirecase where EP02=CP09(+) AND  cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL4 & StrSQL6
    strSql = strSql + " UNION all  SELECT s2.st01,0,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),s1.st01," & SQLDate("ep06") & "," & SQLDate("ep09") & "," & SQLDate("ep07") & "," & SQLDate("cp27") & "," & SQLDate("cp05") & ",s1.st03 ep05DP from engineerprogress,caseprogress,servicepractice,staff s1,staff s2,casepropertymap where EP02=CP09(+) AND  cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL5 & StrSQL6
End If
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            '若選擇會稿日期
            If Me.Option1(0).Value Then
                'edit by nick 2004/08/02  判斷，fcp 沒有文件齊備日，以收文日和完稿日算
                'modify by sonia 2017/11/13 FCP案或承辦人的部門(ep05DP)為國外部者
                If UCase(Mid(strTemp(2), 1, 3)) = "FCP" Or Left(.Fields("ep05DP").Value, 1) = "F" Then
                     If Len(strTemp(9)) <> 0 And Val(strTemp(9)) <> 0 Then
                         strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(9))), ChangeTStringToWString(ChangeTDateStringToTString(CheckStr(.Fields(11).Value)))))
                     Else
                         strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(8))), ChangeTStringToWString(ChangeTDateStringToTString(CheckStr(.Fields(11).Value)))))
                     End If
                Else
                     If Len(strTemp(9)) <> 0 And Val(strTemp(9)) <> 0 Then
                         'modify by sonia 2017/11/13 P-118402 摘要英譯,沒有文件齊備日，以收文日算
                         'strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(9))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(7)))))
                         If Len(strTemp(7)) <> 0 And Val(strTemp(7)) <> 0 Then
                            strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(9))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(7)))))
                         Else
                            strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(9))), ChangeTStringToWString(ChangeTDateStringToTString(CheckStr(.Fields(11).Value)))))
                         End If
                         'end 2017/11/13
                     Else
                         strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(8))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(7)))))
                     End If
                End If
            '若選擇完稿日期
            Else
                'edit by nick 2004/08/02  判斷，fcp 沒有文件齊備日，以收文日和完稿日算
                'modify by sonia 2017/11/13 FCP案或承辦人的部門(ep05DP)為國外部者
                If UCase(Mid(strTemp(2), 1, 3)) = "FCP" Or Left(.Fields("ep05DP").Value, 1) = "F" Then
                    strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(8))), ChangeTStringToWString(ChangeTDateStringToTString(CheckStr(.Fields(11).Value)))))
                Else
                    'modify by sonia 2017/11/13 P-118402 摘要英譯,沒有文件齊備日，以收文日算
                    'strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(8))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(7)))))
                    If Len(strTemp(7)) <> 0 And Val(strTemp(7)) <> 0 Then
                       strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(8))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(7)))))
                    Else
                       strTemp(1) = str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(8))), ChangeTStringToWString(ChangeTDateStringToTString(CheckStr(.Fields(11).Value)))))
                    End If
                    'end 2017/11/13
                End If
            End If
            'Modify By Cheng 2003/06/05
'            strSQL = "insert into r090611_1 values ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strUserNum & "') "
            strSql = "insert into r090611_1 values ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & ChgSQL(strTemp(3)) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            If Len(strTemp(9)) <> 0 Then
                'Modify By Cheng 2003/06/05
'                If Val(Format(strTemp(9), "m")) = Val(Right(Txt1(3), 2)) Or Val(Format(strTemp(9), "m")) = Val(Right(Txt1(4), 2)) Then
                If Val(strTemp(9)) >= (Val(Me.Txt1(3).Text) + 19110000) And Val(strTemp(9)) <= (Val(Me.Txt1(4).Text) + 19110000) Then
                    Seekok = 1
                Else
                   Seekok = 0
                End If
            Else
                Seekok = 0
            End If
            If Val(strTemp(1)) >= Val(Txt1(12)) And Val(strTemp(1)) <= Val(Txt1(13)) Then
                cnnConnection.Execute "insert into r090611_2 (r108001,r108003,id) values ('" & strTemp(0) & "',1,'" & strUserNum & "') "
            Else
                If Val(strTemp(1)) >= Val(Txt1(14)) And Val(strTemp(1)) <= Val(Txt1(15)) Then
                    cnnConnection.Execute "insert into r090611_2 (r108001,r108004,id) values ('" & strTemp(0) & "',1,'" & strUserNum & "') "
                Else
                    If Val(strTemp(1)) >= Val(Txt1(16)) And Val(strTemp(1)) <= Val(Txt1(17)) Then
                        cnnConnection.Execute "insert into r090611_2 (r108001,r108005,id) values ('" & strTemp(0) & "',1,'" & strUserNum & "') "
                    Else
                        If Val(strTemp(1)) >= Val(Txt1(18)) And Val(strTemp(1)) <= Val(Txt1(19)) Then
                            cnnConnection.Execute "insert into r090611_2 (r108001,r108006,id) values ('" & strTemp(0) & "',1,'" & strUserNum & "') "
                        Else
                            If Val(strTemp(1)) >= Val(Txt1(20)) And Val(strTemp(1)) <= Val(Txt1(21)) Then
                                cnnConnection.Execute "insert into r090611_2 (r108001,r108007,id) values ('" & strTemp(0) & "',1,'" & strUserNum & "') "
                            End If
                        End If
                    End If
                End If
            End If
            .MoveNext
            DoEvents
        Loop
    Else
          ShowNoData
          Exit Sub
    End If
End With
CheckOC
bolData = True
End Sub

Sub Process1()
'查詢
If Val(Txt1(22)) = 1 Then
    pub_QL05 = pub_QL05 & ";" & Label1(22) & "1.螢幕" 'Add By Sindy 2010/12/17
    '明細
    If Val(Txt1(23)) = 1 Then
        pub_QL05 = pub_QL05 & ";" & Label1(23) & "1.明細" 'Add By Sindy 2010/12/17
        Me.Hide
        frm090611_1.Show
    '統計
    Else
        pub_QL05 = pub_QL05 & ";" & Label1(23) & "2.統計" 'Add By Sindy 2010/12/17
        Me.Hide
        frm090611_2.Show
    End If
'列印
Else
    pub_QL05 = pub_QL05 & ";" & Label1(22) & "2.報表" 'Add By Sindy 2010/12/17
    '明細
    If Val(Txt1(23)) = 1 Then
        pub_QL05 = pub_QL05 & ";" & Label1(23) & "1.明細" 'Add By Sindy 2010/12/17
        PrintData1
    '統計
    Else
        pub_QL05 = pub_QL05 & ";" & Label1(23) & "2.統計" 'Add By Sindy 2010/12/17
        PrintData2
    End If
End If
End Sub

Sub PrintData1()
'Modify By Cheng 2003806/05
'strSQL = " SELECT r106001,r106002,r106003,r106004,r106005,r106006,r106007,r106008,r106009,r106010,r106011 FROM R090611_1 WHERE ID='" & strUserNum & "' order by r106001,r106003  "
'Modify By Cheng 2003/06/10
'列印對象為承辦人
If Me.Txt1(7).Text = "1" Then
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "1.承辦人" 'Add By Sindy 2010/12/17
    strSql = "SELECT r106001,r106002,r106003,r106004,r106005,r106006,r106007,r106008,r106009,r106010,r106011 FROM R090611_1, Staff WHERE R106001=ST01(+) And ID='" & strUserNum & "' Order By ST06, ST03, r106001,r106002, r106003 "
'列印對象為智權人員
Else
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "2.智權人員" 'Add By Sindy 2010/12/17
    strSql = "SELECT r106001,r106002,r106003,r106004,r106005,r106006,r106007,r106008,r106009,r106010,r106011 FROM R090611_1, Staff WHERE R106001=ST01(+) And ID='" & strUserNum & "' Order By ST06, ST15, r106001,r106002, r106003 "
End If
CheckOC
Page = 1
strTemp3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        .MoveFirst
        strTemp3 = CheckStr(.Fields(0))
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp3 <> strTemp(0) Then
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
'            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(6) = GetStaffName(strTemp(6), True)
            PrintDatil1
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle1
            End If
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
ShowLine3
PrintEnd1
ShowLine1
Printer.EndDoc
CheckOC
ShowPrintOk
End Sub

Sub PrintEnd3()
'列印結尾
strSql = "select sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007) from r090611_2 where id='" & strUserNum & "' and r108001='" & strTemp3 & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        ShowLine1
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print Txt1(12) & "-" & Txt1(13)
        Printer.CurrentX = 2500
        Printer.CurrentY = iPrint
        Printer.Print Txt1(14) & "-" & Txt1(15)
        Printer.CurrentX = 5000
        Printer.CurrentY = iPrint
        Printer.Print Txt1(16) & "-" & Txt1(17)
        Printer.CurrentX = 7500
        Printer.CurrentY = iPrint
        Printer.Print Txt1(18) & "-" & Txt1(19)
        Printer.CurrentX = 10000
        Printer.CurrentY = iPrint
        Printer.Print Txt1(20) & "-" & Txt1(21)
        iPrint = iPrint + 300
        ShowLine3
        Printer.CurrentX = 500 - Printer.TextWidth(Format(CheckStr(.Fields(0)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(0)), "####0")
        Printer.CurrentX = 3000 - Printer.TextWidth(Format(CheckStr(.Fields(1)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(1)), "####0")
        Printer.CurrentX = 5500 - Printer.TextWidth(Format(CheckStr(.Fields(2)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(2)), "####0")
        Printer.CurrentX = 8000 - Printer.TextWidth(Format(CheckStr(.Fields(3)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(3)), "####0")
        Printer.CurrentX = 10500 - Printer.TextWidth(Format(CheckStr(.Fields(4)), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(CheckStr(.Fields(4)), "####0")
        iPrint = iPrint + 300
        ShowLine1
    End If
End With
CheckOC2
End Sub

Sub PrintData2()
'Modify By Cheng 2003/06/10
'strSQL = "select r108001,decode(sum(r108003),null,0,sum(r108003))+decode(sum(r108004),null,0,sum(r108004))+decode(sum(r108005),null,0,sum(r108005))+decode(sum(r108006),null,0,sum(r108006))+decode(sum(r108007),null,0,sum(r108007)),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007) from r090611_2 where id='" & strUserNum & "' group by r108001 order by r108001 "
'列印對象為承辦人
If Me.Txt1(7).Text = "1" Then
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "1.承辦人" 'Add By Sindy 2010/12/17
    strSql = "select r108001,decode(sum(r108003),null,0,sum(r108003))+decode(sum(r108004),null,0,sum(r108004))+decode(sum(r108005),null,0,sum(r108005))+decode(sum(r108006),null,0,sum(r108006))+decode(sum(r108007),null,0,sum(r108007)),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007), ST06, ST03 from r090611_2, Staff where R108001=ST01(+) And id='" & strUserNum & "' group by r108001, ST06, ST03 order by ST06, ST03, r108001 "
'列印對象為智權人員
Else
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "2.智權人員" 'Add By Sindy 2010/12/17
    strSql = "select r108001,decode(sum(r108003),null,0,sum(r108003))+decode(sum(r108004),null,0,sum(r108004))+decode(sum(r108005),null,0,sum(r108005))+decode(sum(r108006),null,0,sum(r108006))+decode(sum(r108007),null,0,sum(r108007)),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007), ST06, ST15 from r090611_2, Staff where R108001=ST01(+) And id='" & strUserNum & "' group by r108001, ST06, ST15 order by ST06, ST15, r108001 "
End If
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        .MoveFirst
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 6
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Cheng 2003/06/10
'            strTemp(0) = StrToStr(strTemp(0), 6)
            strTemp(0) = GetStaffName(strTemp(0), True)
            PrintDatil2
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle2
            End If
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
ShowLine3
PrintEnd2
ShowLine2
Printer.EndDoc
CheckOC
ShowPrintOk
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Txt1(0) = Systemkind_g
For i = 0 To 21
    StrTemp99(i) = ""
Next i
Txt1(22) = "1"
bol911001checkRange = True

'Added by Lydia 2022/01/12
lbl1(0).Caption = ""
lbl1(1).Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090611 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
    Case 0 '會稿日期
        Me.Option1(1).Value = False
        Me.Txt1(3).Enabled = True: Me.Txt1(4).Enabled = True
        Me.Txt1(24).Enabled = False: Me.Txt1(25).Enabled = False
        Me.Txt1(3).SetFocus
    Case 1 '完稿日期
        Me.Option1(0).Value = False
        Me.Txt1(3).Enabled = False: Me.Txt1(4).Enabled = False
        Me.Txt1(24).Enabled = True: Me.Txt1(25).Enabled = True
        Me.Txt1(24).SetFocus
    End Select
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
     strTemp1 = Split(UCase(Systemkind_g), ",")
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
Case 2, 10
   If RunNick(Txt1(Index - 1), Txt1(Index)) Then
       Txt1(Index - 1).SetFocus
       txt1_GotFocus (Index - 1)
       Exit Sub
   End If
Case 3, 4 '會稿日期
    'Modify By Cheng 2003/06/05
'   If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
    '若選擇會稿日期
    If Me.Option1(0).Value Then
        If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
           Me.Txt1(Index).SetFocus
           txt1_GotFocus Index
           Exit Sub
        End If
        If Index = 4 Then
            If RunNick(Txt1(Index - 1), Txt1(Index)) Then
                Txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
        End If
    End If
Case 5
     bol911001checkRange = True
     Select Case Trim(Txt1(5))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(5).SetFocus
          Txt1(5).SelStart = 0
          Txt1(5).SelLength = Len(Txt1(5))
          bol911001checkRange = False
          Exit Sub
     End Select
Case 6
     If bol911001checkRange = True Then
            Select Case Trim(Txt1(6))
            Case "1", "2", "3", "4", "5", ""
            Case Else
                 s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
                 Txt1(6).SetFocus
                 Txt1(6).SelStart = 0
                 Txt1(6).SelLength = Len(Txt1(6))
                 Exit Sub
            End Select
            If RunNick(Txt1(Index - 1), Txt1(Index)) Then
                Txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
     End If
     bol911001checkRange = True
Case 7
     Select Case Trim(Txt1(7))
     Case "1", "2", ""
     Case Else
          s = MsgBox("查詢對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(7).SetFocus
          Txt1(7).SelStart = 0
          Txt1(7).SelLength = Len(Txt1(7))
          Exit Sub
     End Select
Case 8
     lbl1(0) = GetPrjSalesNM(Txt1(8))
     If Trim(Txt1(Index)) <> "" Then
        If Trim(lbl1(0).Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 11
     lbl1(1) = GetPrjSalesNM(Txt1(11))
     If Trim(Txt1(Index)) <> "" Then
        If Trim(lbl1(1).Caption) = "" Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 12, 13, 14, 15, 16, 17, 18, 19, 20, 21
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
        If Index = 13 Or Index = 15 Or Index = 17 Or Index = 19 Or Index = 21 Then
        If RunNick(Txt1(Index - 1), Txt1(Index)) Then
            Txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
        End If
        End If
Case 22
     Select Case Trim(Txt1(22))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(22).SetFocus
          Txt1(22).SelStart = 0
          Txt1(22).SelLength = Len(Txt1(22))
          Exit Sub
     End Select
Case 23
     Select Case Trim(Txt1(23))
     Case "1", "2", ""
     Case Else
          s = MsgBox("查詢對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(23).SetFocus
          Txt1(23).SelStart = 0
          Txt1(23).SelLength = Len(Txt1(23))
          Exit Sub
     End Select
'Add By Cheng 2003/06/05
Case 24, 25 '完稿日期
    '若選擇完稿日期
    If Me.Option1(1).Value Then
        If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
           Me.Txt1(Index).SetFocus
           txt1_GotFocus Index
           Exit Sub
        End If
        If Index = 25 Then
            If RunNick(Txt1(Index - 1), Txt1(Index)) Then
                Txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
        End If
    End If
Case Else
End Select
End Sub

Sub PrintTitle1() '列印抬頭

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
Printer.Print "承辦人天數明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
'Modify By Cheng 2003/06/05
'Printer.Print "年月：" & Mid(Txt1(3), 1, 2) & "/" & Mid(Txt1(3), 3, 2) & "－" & Mid(Txt1(4), 1, 2) & "/" & Mid(Txt1(4), 3, 2)
If Me.Option1(0).Value Then
    Printer.Print "會稿日期：" & ChangeTStringToTDateString(Me.Txt1(3).Text) & "－" & ChangeTStringToTDateString(Me.Txt1(4).Text)
Else
    Printer.Print "完稿日期：" & ChangeTStringToTDateString(Me.Txt1(24).Text) & "－" & ChangeTStringToTDateString(Me.Txt1(25).Text)
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000 + 500
Printer.CurrentY = iPrint
'Modify By Cheng 2003/06/05
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
If Val(Txt1(7)) = 1 Then
    Printer.Print "承辦人：" & GetStaffName(strTemp3, True)
Else
    Printer.Print "智權人員：" & GetStaffName(strTemp3, True)
End If
Printer.CurrentX = 13000 + 500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000 + 750, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "天數"
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
If Val(Txt1(7)) = 1 Then
    Printer.Print "智權人員"
Else
    Printer.Print "承辦人"
End If
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "會稿日"
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
Printer.Line (0, iPrint + 150)-(15000 + 750, iPrint + 150)
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
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = 1000
PLeft(3) = 3300
PLeft(4) = 7800
PLeft(5) = 8800
PLeft(6) = 10000
PLeft(7) = 11000
PLeft(8) = 12000 + 250
PLeft(9) = 13000 + 500
PLeft(10) = 14000 + 750
End Sub

Sub ShowLine1()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
Printer.Line (0, iPrint + 150)-(15000 + 750, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
End If
End Sub

Sub ShowLine2()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
Printer.Line (0, iPrint + 150)-(15000 + 750, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
End If
End Sub

Sub ShowLine3()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
For i = 120 To 125
'    Printer.Line (0, iPrint + i)-(15000, iPrint + i)
'    Printer.Line (0, iPrint + 50 + i)-(15000, iPrint + 50 + i)
    Printer.Line (0, iPrint + i)-(15000 + 750, iPrint + i)
    Printer.Line (0, iPrint + 50 + i)-(15000 + 750, iPrint + 50 + i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintEnd1()
'列印結尾
strSql = "SELECT '合　計','',SUM(R106002) from r090611_1 WHERE ID='" & strUserNum & "' AND R106001='" & strTemp3 & "' "
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


Sub PrintTitle2() '列印抬頭

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
Printer.Print "承辦人承辦天數統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
'Modify By Cheng 2003/06/05
'Printer.Print "年月：" & Mid(Txt1(3), 1, 2) & "/" & Mid(Txt1(3), 3, 2) & "－" & Mid(Txt1(4), 1, 2) & "/" & Mid(Txt1(4), 3, 2)
If Me.Option1(0).Value Then
    Printer.Print "會稿日期：" & ChangeTStringToTDateString(Me.Txt1(3).Text) & "－" & ChangeTStringToTDateString(Me.Txt1(4).Text)
Else
    Printer.Print "完稿日期：" & ChangeTStringToTDateString(Me.Txt1(24).Text) & "－" & ChangeTStringToTDateString(Me.Txt1(25).Text)
End If
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
'Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
Printer.Line (0, iPrint + 150)-(15000 + 750, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'Modify By Cheng 2003/06/10
If Me.Option1(0).Value Then
    Printer.Print "當月會稿"
Else
    Printer.Print "當月完稿"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print Txt1(12) & "-" & Txt1(13)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print Txt1(14) & "-" & Txt1(15)
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print Txt1(16) & "-" & Txt1(17)
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print Txt1(18) & "-" & Txt1(19)
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print Txt1(20) & "-" & Txt1(21)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000 + 750, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
End Sub

Sub PrintDatil2() '列印資料

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 1 To 6
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
PLeft(1) = 3000
PLeft(2) = 5000
PLeft(3) = 7000
PLeft(4) = 9000
PLeft(5) = 11000
PLeft(6) = 13000
End Sub

Sub PrintEnd2()
'列印結尾
strSql = "select '合　計',decode(sum(r108003),null,0,sum(r108003))+decode(sum(r108004),null,0,sum(r108004))+decode(sum(r108005),null,0,sum(r108005))+decode(sum(r108006),null,0,sum(r108006))+decode(sum(r108007),null,0,sum(r108007)),sum(r108003),sum(r108004),sum(r108005),sum(r108006),sum(r108007) from r090611_2 where id='" & strUserNum & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 6
                StrTemp7(i) = CheckStr(.Fields(i))
                If Len(StrTemp7(i)) = 0 And i > 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            For i = 1 To 6
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
'Add By Cheng 2003/06/10
'列印百分比
strSql = " select '百分比',' ',sum(r108003)/(Nvl(sum(r108003),0)+Nvl(sum(r108004),0)+Nvl(sum(r108005),0)+Nvl(sum(r108006),0)+Nvl(sum(r108007),0)) * 100,sum(r108004)/(Nvl(sum(r108003),0)+Nvl(sum(r108004),0)+Nvl(sum(r108005),0)+Nvl(sum(r108006),0)+Nvl(sum(r108007),0)) * 100,sum(r108005)/(Nvl(sum(r108003),0)+Nvl(sum(r108004),0)+Nvl(sum(r108005),0)+Nvl(sum(r108006),0)+Nvl(sum(r108007),0)) * 100,sum(r108006)/(Nvl(sum(r108003),0)+Nvl(sum(r108004),0)+Nvl(sum(r108005),0)+Nvl(sum(r108006),0)+Nvl(sum(r108007),0)) * 100,sum(r108007)/(Nvl(sum(r108003),0)+Nvl(sum(r108004),0)+Nvl(sum(r108005),0)+Nvl(sum(r108006),0)+Nvl(sum(r108007),0)) * 100 from r090611_2 where id='" & strUserNum & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 6
                StrTemp7(i) = CheckStr(.Fields(i))
                If Len(StrTemp7(i)) = 0 And i > 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            For i = 2 To 6
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(StrTemp7(i), "##0.00") & "%")
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "##0.00") & "%"
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
