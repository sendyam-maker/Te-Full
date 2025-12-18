VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090613 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件處理時間統計查詢"
   ClientHeight    =   6300
   ClientLeft      =   1416
   ClientTop       =   1416
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5220
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   30
      Left            =   2205
      MaxLength       =   1
      TabIndex        =   23
      Top             =   3525
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   29
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   54
      Top             =   1695
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "齊備日："
      Height          =   180
      Index           =   5
      Left            =   360
      TabIndex        =   55
      Top             =   1755
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   28
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   53
      Top             =   1695
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   27
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   16
      Top             =   2301
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "會稿日："
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   14
      Top             =   2361
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   26
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   15
      Top             =   2301
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   25
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   34
      Top             =   6000
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   23
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   18
      Top             =   2610
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   21
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   12
      Top             =   1998
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1392
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1089
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "會完日："
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   2655
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   24
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   19
      Top             =   2610
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "完稿日："
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   2058
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   22
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   13
      Top             =   1998
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   20
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   33
      Top             =   5670
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   19
      Left            =   1995
      MaxLength       =   1
      TabIndex        =   32
      Top             =   5355
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   17
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   25
      Top             =   4125
      Width           =   480
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   18
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   26
      Top             =   4125
      Width           =   525
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   16
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   22
      Top             =   3225
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   15
      Left            =   1350
      MaxLength       =   4
      TabIndex        =   21
      Top             =   3225
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "不管繪圖的日期"
      Height          =   210
      Left            =   3120
      TabIndex        =   31
      Top             =   5085
      Width           =   1770
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Index           =   1
      Left            =   3870
      TabIndex        =   36
      Top             =   45
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3060
      TabIndex        =   35
      Top             =   45
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1392
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1089
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "發文日期："
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1452
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "收文日期："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1149
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   14
      Left            =   1245
      MaxLength       =   1
      TabIndex        =   30
      Top             =   5055
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   13
      Left            =   1245
      MaxLength       =   6
      TabIndex        =   29
      Top             =   4755
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   1875
      MaxLength       =   3
      TabIndex        =   28
      Top             =   4455
      Width           =   525
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   11
      Left            =   1245
      MaxLength       =   3
      TabIndex        =   27
      Top             =   4455
      Width           =   480
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   1245
      MaxLength       =   6
      TabIndex        =   24
      Top             =   3825
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   20
      Top             =   2910
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   4
      Top             =   786
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1215
      MaxLength       =   1
      TabIndex        =   3
      Top             =   786
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1980
      MaxLength       =   4
      TabIndex        =   2
      Top             =   483
      Width           =   675
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1215
      MaxLength       =   4
      TabIndex        =   1
      Top             =   483
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1215
      TabIndex        =   0
      Top             =   180
      Width           =   1650
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   2220
      TabIndex        =   58
      Top             =   4770
      Width           =   1500
      VariousPropertyBits=   27
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   2220
      TabIndex        =   57
      Top             =   3840
      Width           =   1500
      VariousPropertyBits=   27
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "是否只查詢新申請案：              Y: 僅查詢新申請案(含改請) "
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   360
      TabIndex        =   56
      Top             =   3570
      Width           =   4605
   End
   Begin VB.Line Line11 
      X1              =   1770
      X2              =   2655
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line Line10 
      X1              =   1770
      X2              =   2655
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明/新型案件屬性：        ( 1.機械 2.電子電機 3.化學生醫)"
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   52
      Top             =   6045
      Width           =   4725
   End
   Begin VB.Line Line9 
      X1              =   1770
      X2              =   2655
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line8 
      X1              =   1770
      X2              =   2655
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Line Line7 
      X1              =   1620
      X2              =   2505
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line6 
      X1              =   1770
      X2              =   2655
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Line Line3 
      X1              =   1740
      X2              =   2625
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   1410
      X2              =   1770
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "複雜或特殊案件：         ( Y:是 N:否 空白:全部)"
      Height          =   180
      Index           =   13
      Left            =   360
      TabIndex        =   51
      Top             =   5700
      Width           =   3555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "會稿加乘適用規則：         ( 1:甲規則 2:乙規則 )"
      Height          =   180
      Index           =   12
      Left            =   360
      TabIndex        =   50
      Top             =   5400
      Width           =   3630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人部門："
      Height          =   180
      Index           =   11
      Left            =   360
      TabIndex        =   49
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Line Line5 
      X1              =   1770
      X2              =   2490
      Y1              =   4260
      Y2              =   4260
   End
   Begin VB.Label Label2 
      Caption         =   "案件性質："
      Height          =   210
      Left            =   360
      TabIndex        =   48
      Top             =   3240
      Width           =   930
   End
   Begin VB.Line Line4 
      X1              =   1500
      X2              =   2220
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      X1              =   1500
      X2              =   2385
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   24
      Left            =   1980
      TabIndex        =   47
      Top             =   780
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "顯示方式："
      Height          =   180
      Index           =   22
      Left            =   360
      TabIndex        =   46
      Top             =   5100
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "(1.螢幕 2.報表)"
      Height          =   180
      Index           =   21
      Left            =   1650
      TabIndex        =   45
      Top             =   5100
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.承辦人 2.智權人員)"
      Height          =   180
      Index           =   14
      Left            =   1740
      TabIndex        =   44
      Top             =   2955
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   9
      Left            =   360
      TabIndex        =   43
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   8
      Left            =   360
      TabIndex        =   42
      Top             =   3870
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   6
      Left            =   360
      TabIndex        =   41
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   40
      Top             =   495
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   39
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   38
      Top             =   810
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "查詢對象："
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   37
      Top             =   2955
      Width           =   1095
   End
End
Attribute VB_Name = "frm090613"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/08 改成Form2.0 ; lbl1(index) ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 17) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 13) As String, StrTemp98(0 To 10) As String, StrSQL3 As String
Dim PLeft(0 To 17) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, k As Integer
Dim bol911001checkRange As Boolean
'add by nickc 2006/01/18 判斷是否已經開始印錯誤資料
Dim IsStartPrintErrData As Boolean


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(Txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         Txt1(0).SetFocus
         Exit Sub
     Else
         If Len(Txt1(9)) = 0 Then
             s = MsgBox("查詢對象不可空白!!", , "USER 輸入錯誤")
             Txt1(9).SetFocus
             Exit Sub
         Else
             'Modify By Sindy 2015/8/25
             If Option1(0).Value = True Then
                 If Len(Txt1(5)) = 0 Then
                     s = MsgBox("收文起始日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(5).SetFocus
                     Exit Sub
                 End If
                 If Len(Txt1(6)) = 0 Then
                     s = MsgBox("收文迄止日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(6).SetFocus
                     Exit Sub
                 End If
             ElseIf Option1(1).Value = True Then
                 If Len(Txt1(7)) = 0 Then
                     s = MsgBox("發文起始日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(7).SetFocus
                     Exit Sub
                 End If
                 If Len(Txt1(8)) = 0 Then
                     s = MsgBox("發文迄止日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(8).SetFocus
                     Exit Sub
                 End If
                 
             'Added by Morgan 2017/8/17
             ElseIf Option1(5).Value = True Then
                 If Len(Txt1(28)) = 0 Then
                     s = MsgBox("齊備起始日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(28).SetFocus
                     Exit Sub
                 End If
                 If Len(Txt1(29)) = 0 Then
                     s = MsgBox("齊備迄止日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(29).SetFocus
                     Exit Sub
                 End If
             'end 2017/8/17
             ElseIf Option1(2).Value = True Then
                 If Len(Txt1(21)) = 0 Then
                     s = MsgBox("完稿起始日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(21).SetFocus
                     Exit Sub
                 End If
                 If Len(Txt1(22)) = 0 Then
                     s = MsgBox("完稿迄止日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(22).SetFocus
                     Exit Sub
                 End If
             ElseIf Option1(3).Value = True Then
                 If Len(Txt1(23)) = 0 Then
                     s = MsgBox("會完起始日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(23).SetFocus
                     Exit Sub
                 End If
                 If Len(Txt1(24)) = 0 Then
                     s = MsgBox("會完迄止日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(24).SetFocus
                     Exit Sub
                 End If
             'Add By Sindy 2016/3/23
             Else
                 If Len(Txt1(26)) = 0 Then
                     s = MsgBox("會稿起始日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(26).SetFocus
                     Exit Sub
                 End If
                 If Len(Txt1(27)) = 0 Then
                     s = MsgBox("會稿迄止日期不可空白!!", , "USER 輸入錯誤")
                     Txt1(27).SetFocus
                     Exit Sub
                 End If
             '2016/3/23 END
             End If
             '2015/8/25 END
             If Len(Txt1(14)) = 0 Then
                 s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                 Txt1(14).SetFocus
                 Exit Sub
             Else
                 Screen.MousePointer = vbHourglass
                 Me.Enabled = False
                 ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/17 清除查詢印表記錄檔欄位
                 'For i = 0 To 13
                 '   If (StrTemp99(i) <> txt1(i)) And (i <> 10) Then
                        Process
                 '       For j = 0 To 13
                 '           StrTemp99(j) = txt1(j)
                 '       Next j
                 '       Exit For
                 '    End If
                 ' Next i
                  Process1
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

'2007/03/08 加入控制碼，原先錯誤的資料一律以 # 表示區隔，
'現在以其他代號表示
'收文==>1       齊備==>2        會稿==>3        會稿完成==>4        發文==>5        草齊==>6        草完==>7        墨齊==>8        墨完==>9
Sub Process()
Dim strPKey As String 'Add By Sindy 2015/8/25

cnnConnection.Execute "DELETE FROM R090613 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
If Len(Txt1(0)) <> 0 Then
    strSQL1 = strSQL1 & " and CP01 in (" & SQLGrpStr(Txt1(0), 1) & ") " '專利
    strSQL2 = strSQL2 & " and CP01 in (" & SQLGrpStr(Txt1(0), 2) & ") " '商標
    StrSQL3 = StrSQL3 & " and CP01 in (" & SQLGrpStr(Txt1(0), 3) & ") " '法務
    StrSQL4 = StrSQL4 & " and CP01 in (" & SQLGrpStr(Txt1(0), 4) & ") " '顧問
    strSQL5 = strSQL5 & " and CP01 in (" & SQLGrpStr(Txt1(0), 5) & ") "  '服務
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
If Val(Txt1(9)) = 1 Then
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "1.承辦人" 'Add By Sindy 2010/12/17
    If Len(Txt1(3)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06>='" & Txt1(3) & "' "
    End If
    If Len(Txt1(4)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06<='" & Txt1(4) & "' "
    End If
Else
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "2.智權人員" 'Add By Sindy 2010/12/17
    If Len(Txt1(3)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s2.st06>='" & Txt1(3) & "' "
    End If
    If Len(Txt1(4)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s2.st06<='" & Txt1(4) & "' "
    End If
End If
If Len(Txt1(3)) <> 0 Or Len(Txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(3) & "-" & Txt1(4) & Label1(24) 'Add By Sindy 2010/12/17
End If
If Len(Txt1(10)) <> 0 Then
    StrSQL6 = StrSQL6 + " and CP14='" & Txt1(10) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & Txt1(10) & LBL1(0) 'Add By Sindy 2010/12/17
End If
If Len(Txt1(11)) <> 0 Then
    StrSQL6 = StrSQL6 + " and cp12>='" & Txt1(11) & "' "
End If
If Len(Txt1(12)) <> 0 Then
    StrSQL6 = StrSQL6 + " and cp12<='" & Txt1(12) & "' "
End If
If Len(Txt1(11)) <> 0 Or Len(Txt1(12)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(9) & Txt1(11) & "-" & Txt1(12) 'Add By Sindy 2010/12/17
End If
'add by nickc 2007/03/08
If Len(Txt1(17)) <> 0 Then
    StrSQL6 = StrSQL6 + " and s1.st03>='" & Txt1(17) & "' "
End If
If Len(Txt1(18)) <> 0 Then
    StrSQL6 = StrSQL6 + " and s1.st03<='" & Txt1(18) & "' "
End If
If Len(Txt1(17)) <> 0 Or Len(Txt1(18)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(11) & Txt1(17) & "-" & Txt1(18) 'Add By Sindy 2010/12/17
End If
If Len(Txt1(13)) <> 0 Then
    StrSQL6 = StrSQL6 + " and cp13='" & Txt1(13) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(13) & LBL1(1) 'Add By Sindy 2010/12/17
End If
'add by nickc 2006/01/18
If Len(Txt1(15)) <> 0 Then
    StrSQL6 = StrSQL6 & " and cp10>='" & Txt1(15) & "' "
End If
If Len(Txt1(16)) <> 0 Then
    StrSQL6 = StrSQL6 & " and cp10<='" & Txt1(16) & "' "
End If
If Len(Txt1(15)) <> 0 Or Len(Txt1(16)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & Txt1(15) & "-" & Txt1(16) 'Add By Sindy 2010/12/17
End If

'系統種類(1-->專  2-->商  3-->法  4-->顧  5-->服)
'Added by Lydia 2021/11/15  是否只統計新申請案 (服務業務之新申請案案件性質為801、802、805、806)
If Len(Trim(Txt1(30))) <> 0 Then
    strExc(1) = "": strExc(2) = "": strExc(3) = ""

    strExc(1) = SQLGrpStr(strExc(4), 1) '專利
    strExc(2) = SQLGrpStr(strExc(4), 2) '商標
    strExc(3) = SQLGrpStr(strExc(4), 5) '服務

    strExc(1) = Replace(strExc(1), ",' '", "")
    strExc(2) = Replace(strExc(2), ",' '", "")
    strExc(3) = Replace(strExc(3), ",' '", "")
    strExc(0) = ""
   If Len(strExc(1)) > 0 Then  '專利
      '常數 + 含改請(3開頭)
      'Memo by Lydia 2024/12/02 若齊備日的條件有異動，請一併變更frm090642-齊備'Memo by Lydia 2024/12/02 若會稿日的條件有異動，請一併變更frm090642-會稿
      strExc(0) = "(cp01 in (" & strExc(1) & ") and (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3')) "
   End If
   If Len(strExc(2)) > 0 Then  '商標
      If Len(strExc(0)) > 0 Then strExc(0) = strExc(0) & " or "
      strExc(0) = strExc(0) & "(cp01 in (" & strExc(2) & ") and CP10='101') "
   End If
   If Len(strExc(3)) > 0 Then  '服務
      If Len(strExc(0)) > 0 Then strExc(0) = strExc(0) & " or "
      strExc(0) = strExc(0) & "(cp01 in (" & strExc(3) & ") and instr('801,802,805,806',CP10)>0) "
   End If
   StrSQL6 = StrSQL6 & " and (" & strExc(0) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(Label3, 10) & Txt1(30)
End If
'end 2021/11/15

'Modify By Sindy 2015/8/25
'If Option1(0).Value = True Then
'    StrSQL6 = StrSQL6 + " AND CP05>=" & Trim(Val(Txt1(5)) + 1911) & Right(ChgNumByNick(Txt1(6)), 2) & "01 AND CP05<=" & Trim(Val(Txt1(5)) + 1911) & Right(ChgNumByNick(Txt1(6)), 2) & "31 "
'    pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Txt1(5) & Txt1(6) 'Add By Sindy 2010/12/17
'Else
'    StrSQL6 = StrSQL6 + " AND CP27>=" & Trim(Val(Txt1(7)) + 1911) & Right(ChgNumByNick(Txt1(8)), 2) & "01 AND CP27<=" & Trim(Val(Txt1(7)) + 1911) & Right(ChgNumByNick(Txt1(8)), 2) & "31 "
'    pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Txt1(7) & Txt1(8) 'Add By Sindy 2010/12/17
'End If
If Option1(0).Value = True Then
   strPKey = "CP09=EP02(+)"
   StrSQL6 = StrSQL6 + " AND CP05>=" & DBDATE(Txt1(5)) & " AND CP05<=" & DBDATE(Txt1(6))
   pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Txt1(5) & "~" & Txt1(6) 'Add By Sindy 2010/12/17
ElseIf Option1(1).Value = True Then
   strPKey = "CP09=EP02(+)"
   StrSQL6 = StrSQL6 + " AND CP27>=" & DBDATE(Txt1(7)) & " AND CP27<=" & DBDATE(Txt1(8))
   pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Txt1(7) & "~" & Txt1(8) 'Add By Sindy 2010/12/17

'Added by Morgan 2017/8/17
'齊備日
ElseIf Option1(5).Value = True Then
   'Memo by Lydia 2024/12/02 若齊備日的條件有異動，請一併變更frm090642-齊備
   strPKey = "CP09(+)=EP02"
   StrSQL6 = StrSQL6 + " AND EP06>=" & DBDATE(Txt1(28)) & " AND EP06<=" & DBDATE(Txt1(29))
   pub_QL05 = pub_QL05 & ";" & Option1(5).Caption & Txt1(28) & "~" & Txt1(29)
'end 2017/8/17
'完稿日
ElseIf Option1(2).Value = True Then
   strPKey = "CP09(+)=EP02"
   StrSQL6 = StrSQL6 + " AND EP09>=" & DBDATE(Txt1(21)) & " AND EP09<=" & DBDATE(Txt1(22))
   pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & Txt1(21) & "~" & Txt1(22)

'會完日
ElseIf Option1(3).Value = True Then
   strPKey = "CP09(+)=EP02"
   'Modified by Morgan 2017/12/18 會稿日或會完日都要剔除不會稿案件--王副總
   'StrSQL6 = StrSQL6 + " AND EP08>=" & DBDATE(txt1(23)) & " AND EP08<=" & DBDATE(txt1(24))
   StrSQL6 = StrSQL6 + " AND EP08>=" & DBDATE(Txt1(23)) & " AND EP08<=" & DBDATE(Txt1(24)) & " and NVL(ep34,'Y')='Y' "
   pub_QL05 = pub_QL05 & ";" & Option1(3).Caption & Txt1(23) & "~" & Txt1(24)
'會稿日
Else
   'Memo by Lydia 2024/12/02 若會稿日的條件有異動，請一併變更frm090642-會稿
   strPKey = "CP09(+)=EP02"
   'Modified by Morgan 2017/12/18 會稿日或會完日都要剔除不會稿案件--王副總
   'StrSQL6 = StrSQL6 + " AND EP07>=" & DBDATE(txt1(26)) & " AND EP07<=" & DBDATE(txt1(27))
   StrSQL6 = StrSQL6 + " AND EP07>=" & DBDATE(Txt1(26)) & " AND EP07<=" & DBDATE(Txt1(27)) & " and NVL(ep34,'Y')='Y' "
   pub_QL05 = pub_QL05 & ";" & Option1(4).Caption & Txt1(26) & "~" & Txt1(27)
End If
'2015/8/25 END
StrSQL6 = StrSQL6 + " and CP26 IS NULL  "

'Add By Sindy 2015/9/30
'發明/新型案件屬性
If Len(Txt1(25)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA08 in('1','2') AND PA158='" & Txt1(25) & "' " '專利
End If
'2015/9/30 END

'Added by Morgan 2012/10/19
If Txt1(20) = "Y" Then
   StrSQL6 = StrSQL6 & " and cp147='Y' "
ElseIf Txt1(20) = "N" Then
   StrSQL6 = StrSQL6 & " and cp147 is null "
End If
pub_QL05 = pub_QL05 & ";" & Replace(Label1(13), "：", "：" & Txt1(20))
'end 2012/10/19

If Val(Txt1(9)) = 1 Then
'edit by nickc 2006/01/18
'    strSQL = "SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,PATENT WHERE CP09=EP02(+)  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK WHERE CP09=EP02(+)  AND cp01=tm01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE WHERE CP09=EP02(+)  AND CP01=lc01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND Cp04=lC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE WHERE CP09=EP02(+)  AND CP01=hc01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND Cp04=hC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,SERVICEPRACTICE WHERE CP09=EP02(+)  AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL5 & StrSQL6
'edit by nickc 2006/04/27 原先判斷計件才算繪圖，現在改加有繪圖人員才算
'    strSQL = "SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,PATENT WHERE CP09=EP02(+)  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK WHERE CP09=EP02(+)  AND cp01=tm01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE WHERE CP09=EP02(+)  AND CP01=lc01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND Cp04=lC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE WHERE CP09=EP02(+)  AND CP01=hc01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND Cp04=hC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,SERVICEPRACTICE WHERE CP09=EP02(+)  AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL5 & StrSQL6
'edit by nickc 2007/04/04 加入是否會稿
'    strSQL = "SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,PATENT WHERE CP09=EP02(+)  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK WHERE CP09=EP02(+)  AND cp01=tm01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE WHERE CP09=EP02(+)  AND CP01=lc01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND Cp04=lC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE WHERE CP09=EP02(+)  AND CP01=hc01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND Cp04=hC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,SERVICEPRACTICE WHERE CP09=EP02(+)  AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL5 & StrSQL6
    'Memo by Lydia 2024/12/02 若齊備日的條件有異動，請一併變更frm090642-齊備, 'Memo by Lydia 2024/12/02 若會稿日的條件有異動，請一併變更frm090642-會稿
    strSql = "SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,PATENT WHERE " & strPKey & " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
    strSql = strSql + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK WHERE " & strPKey & " AND cp01=tm01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
    strSql = strSql + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE WHERE " & strPKey & " AND CP01=lc01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND Cp04=lC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3 & StrSQL6
    strSql = strSql + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE WHERE " & strPKey & " AND CP01=hc01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND Cp04=hC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4 & StrSQL6
    strSql = strSql + " UNION all  SELECT CP14,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,SERVICEPRACTICE WHERE " & strPKey & " AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL5 & StrSQL6

Else
'edit by nickc 2006/01/18
'    strSQL = "SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,PATENT WHERE CP09=EP02(+)  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK WHERE CP09=EP02(+)  AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE WHERE CP09=EP02(+)  AND CP01=lc01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND Cp04=lC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE WHERE CP09=EP02(+)  AND CP01=hc01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND Cp04=hC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,SERVICEPRACTICE WHERE CP09=EP02(+)  AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL5 & StrSQL6
'edit by nickc 2006/04/27 原先判斷計件才算繪圖，現在改加有繪圖人員才算
'    strSQL = "SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,PATENT WHERE CP09=EP02(+)  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK WHERE CP09=EP02(+)  AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE WHERE CP09=EP02(+)  AND CP01=lc01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND Cp04=lC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE WHERE CP09=EP02(+)  AND CP01=hc01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND Cp04=hC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,SERVICEPRACTICE WHERE CP09=EP02(+)  AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL5 & StrSQL6
'edit by nickc 2007/04/04 加入是否會稿
'    strSQL = "SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,PATENT WHERE CP09=EP02(+)  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK WHERE CP09=EP02(+)  AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE WHERE CP09=EP02(+)  AND CP01=lc01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND Cp04=lC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE WHERE CP09=EP02(+)  AND CP01=hc01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND Cp04=hC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4 & StrSQL6
'    strSQL = strSQL + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,SERVICEPRACTICE WHERE CP09=EP02(+)  AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL5 & StrSQL6
    strSql = "SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,PATENT WHERE " & strPKey & " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1 & StrSQL6
    strSql = strSql + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK WHERE " & strPKey & " AND cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2 & StrSQL6
    strSql = strSql + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE WHERE " & strPKey & " AND CP01=lc01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND Cp04=lC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3 & StrSQL6
    strSql = strSql + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE WHERE " & strPKey & " AND CP01=hc01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND Cp04=hC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4 & StrSQL6
    strSql = strSql + " UNION all  SELECT CP13,CP05,EP06,EP14,EP15,EP07,EP08,EP17,EP18,CP27,cp09,ep20,ep29,cp29,ep34,cp01,cp10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,SERVICEPRACTICE WHERE " & strPKey & " AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL5 & StrSQL6
End If

'Add by Morgan 2009/7/17
If Txt1(19) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1(12), 9) & Txt1(19) & "( 1:甲規則 2:乙規則 )" 'Add By Sindy 2010/12/17
   strSql = " select x.* from (" & strSql & ") x,casepropertymap where cpm01(+)=cp01 and cpm02(+)=cp10 and cpm05='" & IIf(Txt1(19) = "1", "A", "B") & "'"
End If

CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        SavDay3 = str(.RecordCount)
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 16
                strTemp(i) = ""
            Next i
            For i = 0 To 10
                StrTemp98(i) = CheckStr(.Fields(i))
            Next i
            strTemp(0) = StrTemp98(0)
            strTemp(16) = StrTemp98(10)
            If Len(StrTemp98(1)) <> 0 And Len(StrTemp98(2)) <> 0 And Val(StrTemp98(1)) <> 0 And Val(StrTemp98(2)) <> 0 Then
                If Val(StrTemp98(1)) > Val(StrTemp98(2)) Then
                    strTemp(1) = GetWorkDay(StrTemp98(1), StrTemp98(2))
                Else
                    strTemp(1) = GetWorkDay(StrTemp98(2), StrTemp98(1))
                End If
                strTemp(2) = "1"
            Else
                'add by nickc 2006/04/26
                'edit by nickc 2007/03/08
                'If Val(StrTemp98(2)) <> 0 Or Val(StrTemp98(5)) <> 0 Or Val(StrTemp98(6)) <> 0 Or Val(StrTemp98(9)) <> 0 Then
                '    strTemp(15) = "#"
                'End If
                If Val(StrTemp98(1)) <> 0 And Val(StrTemp98(2)) = 0 Then
                    strTemp(15) = "2"
                ElseIf Val(StrTemp98(1)) = 0 And Val(StrTemp98(2)) <> 0 Then
                    strTemp(15) = "1"
                End If
            End If
            'add by nickc 2006/01/18   不計件不算
            'edit by nickc 2006/04/27 原先判斷計件才算繪圖，現在改加有繪圖人員才算
            'If CheckStr(.Fields("ep20")) = "" Then
            'edit by nickc 2007/04/04 加入不繪圖不算
            'If CheckStr(.Fields("ep20")) = "" And CheckStr(.Fields("cp29")) <> "" Then
            If CheckStr(.Fields("ep20")) = "" And CheckStr(.Fields("cp29")) <> "" And CheckStr(.Fields("cp29")) <> "99999" Then
                If Len(StrTemp98(3)) <> 0 And Len(StrTemp98(4)) <> 0 And Val(StrTemp98(3)) <> 0 And Val(StrTemp98(4)) <> 0 Then
                    strTemp(3) = GetWorkDay(StrTemp98(4), StrTemp98(3))
                    strTemp(4) = "1"
                Else
                    'add by nickc 2006/01/18  要計算繪圖
                    If Check1.Value = vbUnchecked Then
                        'edit by nickc 2007/03/08
                        'If Val(StrTemp98(4)) <> 0 Or Val(StrTemp98(5)) <> 0 Or Val(StrTemp98(6)) <> 0 Or Val(StrTemp98(9)) <> 0 Then
                        '    strTemp(15) = "#"
                        'End If
                        If Val(StrTemp98(3)) <> 0 And Val(StrTemp98(4)) = 0 Then
                            strTemp(15) = "7"
                        ElseIf Val(StrTemp98(3)) = 0 And Val(StrTemp98(4)) <> 0 Then
                            strTemp(15) = "6"
                        End If
                    Else
                        pub_QL05 = pub_QL05 & ";" & Check1.Caption 'Add By Sindy 2010/12/17
                    End If
                End If
            End If
            'add by nickc 2007/04/04 加入會稿才算
            'Modify by Morgan 2011/3/8
            'If CheckStr(.Fields("ep34")) = "" Then
            If CheckStr(.Fields("ep34")) <> "N" Then
                If Len(StrTemp98(2)) <> 0 And Len(StrTemp98(5)) <> 0 And Val(StrTemp98(2)) <> 0 And Val(StrTemp98(5)) <> 0 Then
                    strTemp(5) = GetWorkDay(StrTemp98(5), StrTemp98(2))
                    strTemp(6) = "1"
                Else
                    'add by nickc 2006/04/26
                    'edit by nickc 2007/03/08
                    'If Val(StrTemp98(5)) <> 0 Or Val(StrTemp98(6)) <> 0 Or Val(StrTemp98(9)) <> 0 Or IIf(Check1.Value = vbUnchecked, Val(StrTemp98(8)) <> 0 Or Val(StrTemp98(7)) <> 0, False) Then
                    '    strTemp(15) = "#"
                    'End If
                    If Val(StrTemp98(2)) <> 0 And Val(StrTemp98(5)) = 0 Then
                        strTemp(15) = "3"
                    ElseIf Val(StrTemp98(2)) = 0 And Val(StrTemp98(5)) <> 0 Then
                        strTemp(15) = "2"
                    End If
                    
                End If
                If Len(StrTemp98(5)) <> 0 And Len(StrTemp98(6)) <> 0 And Val(StrTemp98(5)) <> 0 And Val(StrTemp98(6)) <> 0 Then
                    strTemp(7) = GetWorkDay(StrTemp98(6), StrTemp98(5))
                    strTemp(8) = "1"
                Else
                    'add by nickc 2006/04/26
                    'edit by nickc 2007/03/08
                    'If Val(StrTemp98(6)) <> 0 Or Val(StrTemp98(9)) <> 0 Or IIf(Check1.Value = vbUnchecked, Val(StrTemp98(8)) <> 0 Or Val(StrTemp98(7)) <> 0, False) Then
                    '    strTemp(15) = "#"
                    'End If
                    If Val(StrTemp98(5)) <> 0 And Val(StrTemp98(6)) = 0 Then
                        strTemp(15) = "4"
                    ElseIf Val(StrTemp98(5)) = 0 And Val(StrTemp98(6)) <> 0 Then
                        strTemp(15) = "3"
                    End If
                End If
            End If
            'add by nickc 2006/01/18 '不計件不算
            'edit by nickc 2006/04/27 原先判斷計件才算繪圖，現在改加有繪圖人員才算
            'If CheckStr(.Fields("ep29")) = "" Then
            'edit by nickc 2007/04/04 加入不繪圖不算
            'If CheckStr(.Fields("ep29")) = "" And CheckStr(.Fields("cp29")) <> "" Then
            If CheckStr(.Fields("ep29")) = "" And CheckStr(.Fields("cp29")) <> "" And CheckStr(.Fields("cp29")) <> "99999" Then
                If Len(StrTemp98(7)) <> 0 And Len(StrTemp98(8)) <> 0 And Val(StrTemp98(7)) <> 0 And Val(StrTemp98(8)) <> 0 Then
                    strTemp(9) = GetWorkDay(StrTemp98(8), StrTemp98(7))
                    strTemp(10) = "1"
                Else
                    'add by nickc 2006/01/18 要計算繪圖
                    If Check1.Value = vbUnchecked Then
                        'add by nickc 2006/04/26
                        'edit by nickc 2007/03/08
                        'If Val(StrTemp98(8)) <> 0 Or Val(StrTemp98(9)) <> 0 Then
                        '    strTemp(15) = "#"
                        'End If
                        If Val(StrTemp98(7)) <> 0 And Val(StrTemp98(8)) = 0 Then
                            strTemp(15) = "9"
                        ElseIf Val(StrTemp98(7)) = 0 And Val(StrTemp98(8)) <> 0 Then
                            strTemp(15) = "8"
                        End If
                    End If
                End If
            End If
            'add by nickc 2007/04/04 加入會稿才算
            'Modify by Morgan 2011/3/8
            'If CheckStr(.Fields("ep34")) = "" Then
            If CheckStr(.Fields("ep34")) <> "N" Then
                If Len(StrTemp98(6)) <> 0 And Len(StrTemp98(9)) <> 0 And Val(StrTemp98(6)) <> 0 And Val(StrTemp98(9)) <> 0 Then
                    strTemp(11) = GetWorkDay(StrTemp98(9), StrTemp98(6))
                    strTemp(12) = "1"
                Else
                    'add by nickc 2006/04/26
                    'edit by nickc 2007/03/08
                    'If Val(StrTemp98(9)) <> 0 Then
                    '    strTemp(15) = "#"
                    'End If
                    If Val(StrTemp98(6)) <> 0 And Val(StrTemp98(9)) = 0 Then
                        strTemp(15) = "5"
                    ElseIf Val(StrTemp98(6)) = 0 And Val(StrTemp98(9)) <> 0 Then
                        strTemp(15) = "4"
                    End If
                End If
            End If
            If Len(StrTemp98(1)) <> 0 And Len(StrTemp98(9)) <> 0 And Val(StrTemp98(1)) <> 0 And Val(StrTemp98(9)) <> 0 Then
                strTemp(13) = GetWorkDay(StrTemp98(9), StrTemp98(1))
                strTemp(14) = "1"
            Else
                'add by nickc 2006/04/26
                'edit by nickc 2007/03/08
                'If Val(StrTemp98(9)) <> 0 Then
                '    strTemp(15) = "#"
                'End If
                If Val(StrTemp98(1)) <> 0 And Val(StrTemp98(9)) = 0 Then
                    strTemp(15) = "5"
                ElseIf Val(StrTemp98(1)) = 0 And Val(StrTemp98(9)) <> 0 Then
                    strTemp(15) = "1"
                End If
            End If
            strSql = "INSERT INTO R090613 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & ",'" & strTemp(15) & "','" & strTemp(16) & "','" & strUserNum & "') "
            'Debug.Print strTemp(16), strTemp(0)
            'If strTemp(0) = "90006" Then
            '   Debug.Print
            'End If
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
    End If
End With
CheckOC
End Sub

Sub Process1()
If Val(Txt1(14)) = 1 Then
    pub_QL05 = pub_QL05 & ";" & Label1(22) & "1.螢幕" 'Add By Sindy 2010/12/17
    Me.Hide
    frm090613_1.Show
Else
    pub_QL05 = pub_QL05 & ";" & Label1(22) & "2.報表" 'Add By Sindy 2010/12/17
    PrintData
End If
End Sub

Sub PrintEnd1()
'edit by nickc 2006/05/02
'strSQL = "SELECT '合計',round(SUM(R109002)/SUM(R109003),2),SUM(R109003),round(SUM(R109004)/SUM(R109005),2),SUM(R109005),round(SUM(R109006)/SUM(R109007),2),SUM(R109007),round(SUM(R109008)/SUM(R109009),2),SUM(R109009),round(SUM(R109010)/SUM(R109011),2),SUM(R109011),round(SUM(R109012)/SUM(R109013),2),SUM(R109013),round(SUM(R109014)/SUM(R109015),2),SUM(R109015) FROM R090613 WHERE ID='" & strUserNum & "' and (r109016 is null or r109016='') "
strSql = "SELECT '合計',decode(SUM(R109003),0,0,null,0,round(SUM(R109002)/SUM(R109003),2)),SUM(R109003),decode(SUM(R109005),0,0,null,0,round(SUM(R109004)/SUM(R109005),2)),SUM(R109005),decode(SUM(R109007),0,0,null,0,round(SUM(R109006)/SUM(R109007),2)),SUM(R109007),decode(SUM(R109009),0,0,null,0,round(SUM(R109008)/SUM(R109009),2)),SUM(R109009),decode(SUM(R109011),0,0,null,0,round(SUM(R109010)/SUM(R109011),2)),SUM(R109011),decode(SUM(R109013),0,0,null,0,round(SUM(R109012)/SUM(R109013),2)),SUM(R109013),decode(SUM(R109015),0,0,null,0,round(SUM(R109014)/SUM(R109015),2)),SUM(R109015) FROM R090613 WHERE ID='" & strUserNum & "' and (r109016 is null or r109016='') "
'列印結尾
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        For i = 0 To 14
            strTemp(i) = CheckStr(.Fields(i))
        Next i
        strTemp(0) = StrToStr(strTemp(0), 4)
        PrintDatil
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle
        End If
    End If
End With
CheckOC
End Sub

Sub PrintEnd2()
'列印結尾
'edit by nickc 2007/03/08
'strSQL = "SELECT '合計',DECODE(SUM(R109003),0,0,round(SUM(R109002)/SUM(R109003),2)),SUM(R109003),DECODE(SUM(R109005),0,0,round(SUM(R109004)/SUM(R109005),2)),SUM(R109005),DECODE(SUM(R109007),0,0,round(SUM(R109006)/SUM(R109007),2)),SUM(R109007),DECODE(SUM(R109009),0,0,round(SUM(R109008)/SUM(R109009),2)),SUM(R109009),DECODE(SUM(R109011),0,0,round(SUM(R109010)/SUM(R109011),2)),SUM(R109011),DECODE(SUM(R109013),0,0,round(SUM(R109012)/SUM(R109013),2)),SUM(R109013),DECODE(SUM(R109015),0,0,round(SUM(R109014)/SUM(R109015),2)),SUM(R109015) FROM R090613 WHERE ID='" & strUserNum & "' and r109016='#' "
strSql = "SELECT '合計',DECODE(SUM(R109003),0,0,round(SUM(R109002)/SUM(R109003),2)),SUM(R109003),DECODE(SUM(R109005),0,0,round(SUM(R109004)/SUM(R109005),2)),SUM(R109005),DECODE(SUM(R109007),0,0,round(SUM(R109006)/SUM(R109007),2)),SUM(R109007),DECODE(SUM(R109009),0,0,round(SUM(R109008)/SUM(R109009),2)),SUM(R109009),DECODE(SUM(R109011),0,0,round(SUM(R109010)/SUM(R109011),2)),SUM(R109011),DECODE(SUM(R109013),0,0,round(SUM(R109012)/SUM(R109013),2)),SUM(R109013),DECODE(SUM(R109015),0,0,round(SUM(R109014)/SUM(R109015),2)),SUM(R109015) FROM R090613 WHERE ID='" & strUserNum & "' and r109016 is not null  "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        For i = 0 To 14
            strTemp(i) = CheckStr(.Fields(i))
        Next i
        strTemp(0) = StrToStr(strTemp(0), 4)
        PrintDatil
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle
        End If
    End If
End With
CheckOC
End Sub

Sub PrintData()
'add by nickc 2006/01/18
IsStartPrintErrData = False
'edit by nickc 2006/01/18
'strSQL = "SELECT nvl(ST02,R109001),round(SUM(R109002)/SUM(R109003),2),SUM(R109003),round(SUM(R109004)/SUM(R109005),2),SUM(R109005),round(SUM(R109006)/SUM(R109007),2),SUM(R109007),round(SUM(R109008)/SUM(R109009),2),SUM(R109009),round(SUM(R109010)/SUM(R109011),2),SUM(R109011),round(SUM(R109012)/SUM(R109013),2),SUM(R109013),round(SUM(R109014)/SUM(R109015),2),SUM(R109015),R109001 FROM R090613,STAFF  WHERE ID='" & strUserNum & "' AND R109001=ST01(+) and (r109016 is null or r109016='') GROUP BY R109001,nvl(ST02,R109001) ORDER BY R109001,nvl(ST02,R109001) "
strSql = "SELECT nvl(ST02,R109001),decode(SUM(R109003),0,0,round(SUM(R109002)/SUM(R109003),2)),SUM(R109003),decode(SUM(R109005),0,0,round(SUM(R109004)/SUM(R109005),2)),SUM(R109005),decode(SUM(R109007),0,0,round(SUM(R109006)/SUM(R109007),2)),SUM(R109007),decode(SUM(R109009),0,0,round(SUM(R109008)/SUM(R109009),2)),SUM(R109009),decode(SUM(R109011),0,0,round(SUM(R109010)/SUM(R109011),2)),SUM(R109011),decode(SUM(R109013),0,0,round(SUM(R109012)/SUM(R109013),2)),SUM(R109013),decode(SUM(R109015),0,0,round(SUM(R109014)/SUM(R109015),2)),SUM(R109015),R109001 FROM R090613,STAFF  WHERE ID='" & strUserNum & "' AND R109001=ST01(+) and (r109016 is null or r109016='') GROUP BY R109001,nvl(ST02,R109001) ORDER BY R109001,nvl(ST02,R109001) "
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SeekTemp = .RecordCount
        PrintTitle
        Do While .EOF = False
            For i = 0 To 14
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(0) = StrToStr(strTemp(0), 4)
            PrintDatil
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
PrintEnd1
Page = Page + 1
Printer.NewPage
'add by nickc  2006/01/18
IsStartPrintErrData = True
'edit by nickc 2007/03/08
'strSQL = "SELECT nvl(ST02,R109001),DECODE(SUM(R109003),0,0,round(SUM(R109002)/SUM(R109003),2)),SUM(R109003),DECODE(SUM(R109005),0,0,round(SUM(R109004)/SUM(R109005),2)),SUM(R109005),DECODE(SUM(R109007),0,0,round(SUM(R109006)/SUM(R109007),2)),SUM(R109007),DECODE(SUM(R109009),0,0,round(SUM(R109008)/SUM(R109009),2)),SUM(R109009),DECODE(SUM(R109011),0,0,round(SUM(R109010)/SUM(R109011),2)),SUM(R109011),DECODE(SUM(R109013),0,0,round(SUM(R109012)/SUM(R109013),2)),SUM(R109013),DECODE(SUM(R109015),0,0,round(SUM(R109014)/SUM(R109015),2)),SUM(R109015),R109001 FROM R090613,STAFF WHERE ID='" & strUserNum & "' AND R109001=ST01(+) and r109016='#' GROUP BY R109001,nvl(ST02,R109001) ORDER BY R109001,nvl(ST02,R109001) "
strSql = "SELECT nvl(ST02,R109001),DECODE(SUM(R109003),0,0,round(SUM(R109002)/SUM(R109003),2)),SUM(R109003),DECODE(SUM(R109005),0,0,round(SUM(R109004)/SUM(R109005),2)),SUM(R109005),DECODE(SUM(R109007),0,0,round(SUM(R109006)/SUM(R109007),2)),SUM(R109007),DECODE(SUM(R109009),0,0,round(SUM(R109008)/SUM(R109009),2)),SUM(R109009),DECODE(SUM(R109011),0,0,round(SUM(R109010)/SUM(R109011),2)),SUM(R109011),DECODE(SUM(R109013),0,0,round(SUM(R109012)/SUM(R109013),2)),SUM(R109013),DECODE(SUM(R109015),0,0,round(SUM(R109014)/SUM(R109015),2)),SUM(R109015),R109001 FROM R090613,STAFF WHERE ID='" & strUserNum & "' AND R109001=ST01(+) and r109016 is not null GROUP BY R109001,nvl(ST02,R109001) ORDER BY R109001,nvl(ST02,R109001) "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        .MoveFirst
        SeekTemp = .RecordCount
        PrintTitle
        Do While .EOF = False
            For i = 0 To 14
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            PrintDatil
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
        'add by nickc 2006/04/26
        ShowLine
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle
        End If
        PrintEnd2
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
CheckOC
'add by nickc 2007/03/08 加入錯誤明細表
Page = Page + 1
Printer.NewPage
strSql = "SELECT nvl(ST02,R109001),r109016,r109017,cp01,cp02,cp03,cp04,r109001 FROM R090613,STAFF,caseprogress WHERE ID='" & strUserNum & "' AND R109001=ST01(+) and r109016 is not null and r109017=cp09(+) ORDER BY R109001,nvl(ST02,R109001) "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SeekTemp = .RecordCount
        PrintTitleErr
        Do While .EOF = False
            For i = 0 To 7
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(8) = ""
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print strTemp(0)
            strTemp(8) = "收文號：" & strTemp(2) & "；本所案號：" & strTemp(3) & "-" & strTemp(4) & "-" & strTemp(5) & "-" & strTemp(6) & "；錯誤資訊：" & IIf(strTemp(1) = "1", "沒有收文日", IIf(strTemp(1) = "2", "沒有齊備日", IIf(strTemp(1) = "3", "沒有會稿日", IIf(strTemp(1) = "4", "沒有會稿完成日", IIf(strTemp(1) = "5", "沒有發文日", IIf(strTemp(1) = "6", "沒有草圖齊備日", IIf(strTemp(1) = "7", "沒有草圖完稿日", IIf(strTemp(1) = "8", "沒有墨圖齊備日", IIf(strTemp(1) = "9", "沒有墨圖完稿日", "其他未知錯誤")))))))))
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print strTemp(8)
            iPrint = iPrint + 300
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitleErr
            End If
            .MoveNext
        Loop
        'add by nickc 2006/04/26
        ShowLine
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitleErr
        End If
    End If
End With
CheckOC
'edit by nickc 2006/04/26
'ShowLine
'If iPrint >= 9000 Then
'    Page = Page + 1
'    Printer.NewPage
'    PrintTitle
'End If
'PrintEnd2
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle() '列印抬頭
Printer.Orientation = 2
iPrint = 0
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "案件處理時間統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
'Modify By Sindy 2015/8/25
'If Option1(0).Value = True Then
'    Printer.Print "年月：" & Txt1(5) & "/" & Txt1(6)
'Else
'    Printer.Print "年月：" & Txt1(7) & "/" & Txt1(8)
'End If
If Option1(0).Value = True Then
   Printer.Print "收文日期：" & ChangeTStringToTDateString(Txt1(5)) & "~" & ChangeTStringToTDateString(Txt1(6))
ElseIf Option1(1).Value = True Then
   Printer.Print "發文日期：" & ChangeTStringToTDateString(Txt1(7)) & "~" & ChangeTStringToTDateString(Txt1(8))
'Added by Morgan 2017/8/17
ElseIf Option1(5).Value = True Then
   Printer.Print "齊備日：" & ChangeTStringToTDateString(Txt1(28)) & "~" & ChangeTStringToTDateString(Txt1(29))
'end 2017/8/17
ElseIf Option1(2).Value = True Then
   Printer.Print "完稿日：" & ChangeTStringToTDateString(Txt1(21)) & "~" & ChangeTStringToTDateString(Txt1(22))
ElseIf Option1(3).Value = True Then
   Printer.Print "會完日：" & ChangeTStringToTDateString(Txt1(23)) & "~" & ChangeTStringToTDateString(Txt1(24))
Else
   Printer.Print "會稿日：" & ChangeTStringToTDateString(Txt1(26)) & "~" & ChangeTStringToTDateString(Txt1(27))
End If
'2015/8/25 END
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    'edit by nickc 2006/01/18
    'Printer.Print Txt1(6) & " 月收文共計" & SavDay3 & " 件 "
    'edit by nickc 2006/04/26
    'Printer.Print Txt1(6) & " 月收文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "，包含繪圖錯誤資料", "")
    'edit by nickc 2007/03/08 加入系統別
    'Printer.Print txt1(6) & " 月收文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
    'Modify By Sindy 2015/8/25
    'Printer.Print Txt1(6) & " 月" & IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "收文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
    Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "收文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
    '2015/8/25 END
ElseIf Option1(1).Value = True Then
    'edit by nickc 2006/01/18
    'Printer.Print Txt1(8) & " 月發文共計" & SavDay3 & " 件 "
    'edit by nickc 2006/04/26
    'Printer.Print Txt1(8) & " 月發文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "，包含繪圖錯誤資料", "")
    'edit by nickc 2007/03/08 加入系統別
    'Printer.Print txt1(8) & " 月發文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
    'Modify By Sindy 2015/8/25
    'Printer.Print Txt1(8) & " 月" & IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "發文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
    Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "發文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
    '2015/8/25 END
    
'Added by Morgan 2017/8/17
ElseIf Option1(5).Value = True Then
    Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "齊備共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
'end 2017/8/17
'Modify By Sindy 2015/8/25
ElseIf Option1(2).Value = True Then
    Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "完稿共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
ElseIf Option1(3).Value = True Then
    Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "會完共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
'2015/8/25 END
'Add By Sindy 2016/3/23
Else
    Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "會稿共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
End If
'add by nickc 2006/01/18
If IsStartPrintErrData = True Then
    'edit by nickc 2006/04/26
    'Printer.CurrentX = 7000 - (Printer.TextWidth("錯誤資料" & IIf(Check1.Value = vbChecked, "，繪圖錯誤資料不列入錯誤", "")) / 2)
    Printer.CurrentX = 7000 - (Printer.TextWidth("錯誤資料" & IIf(Check1.Value = vbChecked, "：繪圖錯誤資料不列入錯誤", "")) / 2)
    Printer.CurrentY = iPrint
    'edit by nickc 2006/04/26
    'Printer.Print "錯誤資料" & IIf(Check1.Value = vbChecked, "，繪圖錯誤資料不列入錯誤", "")
    Printer.Print "錯誤資料" & IIf(Check1.Value = vbChecked, "：繪圖錯誤資料不列入錯誤", "")
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
GetPleft
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "草圖"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "草圖"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "會稿"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "墨圖"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "墨圖"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "會稿"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "收文----"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "齊備"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "齊備----"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "完稿"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "齊備----"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "會稿"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "會稿----"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "完成"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "齊備----"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "完稿"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "完成----"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "發文"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "收文----"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "發文"
iPrint = iPrint + 300
For i = 1 To 13 Step 2
    Printer.Line (PLeft(i), iPrint + 150)-(PLeft(i + 1) - PLeft(i) + PLeft(i + 1) - 100, iPrint + 150)
Next i
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "平均"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "平均"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "平均"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "平均"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "平均"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "平均"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "平均"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
If Val(Txt1(9)) = 1 Then
    Printer.Print "承辦人"
Else
    Printer.Print "智權人員"
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "天數"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "天數"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "天數"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "天數"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "天數"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "天數"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "天數"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "件數"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub
'add by nickc 2007/03/08 加入錯誤資料明細表頭
Sub PrintTitleErr() '列印抬頭
iPrint = 0
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "案件處理時間統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
'Modify By Sindy 2015/8/25
'If Option1(0).Value = True Then
'    Printer.Print "年月：" & Txt1(5) & "/" & Txt1(6)
'Else
'    Printer.Print "年月：" & Txt1(7) & "/" & Txt1(8)
'End If
If Option1(0).Value = True Then
   Printer.Print "收文日期：" & ChangeTStringToTDateString(Txt1(5)) & "~" & ChangeTStringToTDateString(Txt1(6))
ElseIf Option1(1).Value = True Then
   Printer.Print "發文日期：" & ChangeTStringToTDateString(Txt1(7)) & "~" & ChangeTStringToTDateString(Txt1(8))
'Added by Morgan 2017/8/17
ElseIf Option1(5).Value = True Then
   Printer.Print "齊備日：" & ChangeTStringToTDateString(Txt1(28)) & "~" & ChangeTStringToTDateString(Txt1(29))
'end 2017/8/17
ElseIf Option1(2).Value = True Then
   Printer.Print "完稿日：" & ChangeTStringToTDateString(Txt1(21)) & "~" & ChangeTStringToTDateString(Txt1(22))
ElseIf Option1(3).Value = True Then
   Printer.Print "會完日：" & ChangeTStringToTDateString(Txt1(23)) & "~" & ChangeTStringToTDateString(Txt1(24))
'Add By Sindy 2016/3/23
Else
   Printer.Print "會稿日：" & ChangeTStringToTDateString(Txt1(26)) & "~" & ChangeTStringToTDateString(Txt1(27))
'2016/3/23 END
End If
'2015/8/25 END
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'Modify By Sindy 2015/8/25
'If Option1(0).Value = True Then
'    Printer.Print Txt1(6) & " 月" & IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "收文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
'Else
'    Printer.Print Txt1(8) & " 月" & IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "發文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
'End If
If Option1(0).Value = True Then
   Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "收文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
ElseIf Option1(1).Value = True Then
   Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "發文共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
'Added by Morgan 2017/8/17
ElseIf Option1(5).Value = True Then
   Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "完稿共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
'end 2017/8/17
ElseIf Option1(2).Value = True Then
   Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "完稿共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
ElseIf Option1(3).Value = True Then
   Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "會完共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
'Add By Sindy 2016/3/23
Else
   Printer.Print IIf(Trim(Txt1(0)) <> "", "(" & Txt1(0) & ")", "") & "會稿共計" & SavDay3 & " 件 " & IIf(Check1.Value = vbChecked, "（包含繪圖錯誤資料）", "")
'2016/3/23 END
End If
'2015/8/25 END
If IsStartPrintErrData = True Then
    Printer.CurrentX = 7000 - (Printer.TextWidth("錯誤資料明細" & IIf(Check1.Value = vbChecked, "：繪圖錯誤資料不列入錯誤", "")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "錯誤資料明細" & IIf(Check1.Value = vbChecked, "：繪圖錯誤資料不列入錯誤", "")
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
GetPleft
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
If Val(Txt1(9)) = 1 Then
    Printer.Print "承辦人"
Else
    Printer.Print "智權人員"
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "錯誤內容"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub

Sub PrintDatil() '列印資料
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 0 To 6
    Printer.CurrentX = PLeft((i * 2) + 1) + 500 - Printer.TextWidth(Format(strTemp((i * 2) + 1), "####0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp((i * 2) + 1), "####0.00")
    Printer.CurrentX = PLeft((i * 2) + 2) + 500 - Printer.TextWidth(Format(strTemp((i * 2) + 2), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp((i * 2) + 2), "####0")
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
'定陣列
Erase PLeft
PLeft(0) = 0
For i = 1 To 14
    PLeft(i) = 500 + (1000 * i)
Next i
End Sub

Sub ShowLine()
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Txt1(0) = Systemkind_g
Txt1(14) = "1"
bol911001checkRange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090613 = Nothing
End Sub

'Added by Morgan 2017/8/17
Private Sub Option1_Click(Index As Integer)
   Dim ii As Integer
   Select Case Index
      Case 0: ii = 5
      Case 1: ii = 7
      Case 5: ii = 28
      Case 2: ii = 21
      Case 4: ii = 26
      Case 3: ii = 23
   End Select
   Txt1(ii).SetFocus
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add by Morgan 2009/7/17
   Select Case Index
      Case 19
         If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Added by Morgan 2012/10/19
      Case 20
         If KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Add By Sindy 2015/9/30
      Case 25
         If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Added by Lydia 2021/11/15 是否只查詢新申請案
      Case 30
         If KeyAscii <> Asc("Y") And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
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
'edit by nickc 2006/01/18
'Case 2, 12
'edit by nickc 2007/03/08
'Case 2, 12, 16
Case 2, 12, 16, 18
     If RunNick(Txt1(Index - 1), Txt1(Index)) Then
        Txt1(Index - 1).SetFocus
        txt1_GotFocus (Index - 1)
        Exit Sub
     End If
Case 3
     bol911001checkRange = True
     Select Case Trim(Txt1(3))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(3).SetFocus
          Txt1(3).SelStart = 0
          Txt1(3).SelLength = Len(Txt1(3))
          bol911001checkRange = False
          Exit Sub
     End Select
Case 4
     If bol911001checkRange = True Then
            Select Case Trim(Txt1(4))
            Case "1", "2", "", "3", "4", "5"
            Case Else
                 s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
                 Txt1(4).SetFocus
                 Txt1(4).SelStart = 0
                 Txt1(4).SelLength = Len(Txt1(4))
                 Exit Sub
            End Select
            If RunNick(Txt1(Index - 1), Txt1(Index)) Then
               Txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
            End If
     End If
     bol911001checkRange = True
'Modify By Sindy 2015/8/25
Case 5, 6
      If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
         Me.Txt1(Index).SetFocus
         txt1_GotFocus Index
         Exit Sub
      End If
      If Index = 6 Then
        If Not nickChgRan(Txt1(5), Txt1(6), "日期") Then
            Txt1(5).SetFocus
            txt1_GotFocus (5)
            Exit Sub
        End If
      End If
Case 7, 8
      If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
         Me.Txt1(Index).SetFocus
         txt1_GotFocus Index
         Exit Sub
      End If
      If Index = 6 Then
        If Not nickChgRan(Txt1(7), Txt1(8), "日期") Then
            Txt1(7).SetFocus
            txt1_GotFocus (7)
            Exit Sub
        End If
      End If
Case 21, 22
      If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
         Me.Txt1(Index).SetFocus
         txt1_GotFocus Index
         Exit Sub
      End If
      If Index = 6 Then
        If Not nickChgRan(Txt1(21), Txt1(22), "日期") Then
            Txt1(21).SetFocus
            txt1_GotFocus (21)
            Exit Sub
        End If
      End If
Case 23, 24
      If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
         Me.Txt1(Index).SetFocus
         txt1_GotFocus Index
         Exit Sub
      End If
      If Index = 6 Then
        If Not nickChgRan(Txt1(23), Txt1(24), "日期") Then
            Txt1(23).SetFocus
            txt1_GotFocus (23)
            Exit Sub
        End If
      End If
'2015/8/25 END
'Add By Sindy 2016/3/24
Case 26, 27
      If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
         Me.Txt1(Index).SetFocus
         txt1_GotFocus Index
         Exit Sub
      End If
      If Index = 6 Then
        If Not nickChgRan(Txt1(26), Txt1(27), "日期") Then
            Txt1(26).SetFocus
            txt1_GotFocus (26)
            Exit Sub
        End If
      End If
'2016/3/24 END
'Case 5
'     If Option1(0).Value = True Then
'        If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
'            s = MsgBox("年輸入錯誤!!", , "USER 輸入錯誤")
'            txt1(Index).SetFocus
'            txt1(Index).SelStart = 0
'            txt1(Index).SelLength = Len(txt1(Index))
'            Exit Sub
'        End If
'     End If
'Case 7
'     If Option1(1).Value = True Then
'        If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
'            s = MsgBox("年輸入錯誤!!", , "USER 輸入錯誤")
'            txt1(Index).SetFocus
'            txt1(Index).SelStart = 0
'            txt1(Index).SelLength = Len(txt1(Index))
'            Exit Sub
'        End If
'     End If
'Case 6
'     If Option1(0).Value = True Then
'        If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
'            s = MsgBox("月輸入錯誤!!", , "USER 輸入錯誤")
'            txt1(Index).SetFocus
'            txt1(Index).SelStart = 0
'            txt1(Index).SelLength = Len(txt1(Index))
'            Exit Sub
'        Else
'            If Val(txt1(Index)) < 1 Or Val(txt1(Index)) > 12 Then
'                s = MsgBox("月輸入錯誤!!", , "USER 輸入錯誤")
'                txt1(Index).SetFocus
'                txt1(Index).SelStart = 0
'                txt1(Index).SelLength = Len(txt1(Index))
'                Exit Sub
'            End If
'         End If
'     End If
'Case 8
'     If Option1(1).Value = True Then
'        If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
'            s = MsgBox("月輸入錯誤!!", , "USER 輸入錯誤")
'            txt1(Index).SetFocus
'            txt1(Index).SelStart = 0
'            txt1(Index).SelLength = Len(txt1(Index))
'            Exit Sub
'        Else
'            If Val(txt1(Index)) < 1 Or Val(txt1(Index)) > 12 Then
'                s = MsgBox("月輸入錯誤!!", , "USER 輸入錯誤")
'                txt1(Index).SetFocus
'                txt1(Index).SelStart = 0
'                txt1(Index).SelLength = Len(txt1(Index))
'                Exit Sub
'            End If
'         End If
'     End If
Case 9
     Select Case Trim(Txt1(9))
     Case "1", "2", ""
     Case Else
          s = MsgBox("查詢對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(9).SetFocus
          Txt1(9).SelStart = 0
          Txt1(9).SelLength = Len(Txt1(9))
          Exit Sub
     End Select
Case 10
     LBL1(0).Caption = GetPrjSalesNM(Txt1(10))
     If Trim(Txt1(Index)) <> "" Then
        If Trim(LBL1(0).Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 13
     LBL1(1).Caption = GetPrjSalesNM(Txt1(13))
     If Trim(Txt1(Index)) <> "" Then
        If Trim(LBL1(1).Caption) = "" Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 14
     Select Case Trim(Txt1(14))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(14).SetFocus
          Txt1(14).SelStart = 0
          Txt1(14).SelLength = Len(Txt1(14))
          Exit Sub
     End Select
'Add By Sindy 2015/9/30
Case 25
     Select Case Trim(Txt1(Index))
     Case "1", "2", "3", ""
     Case Else
          s = MsgBox("發明/新型案件屬性只能輸入 1 或 2 或 3 !!", , "USER 輸入錯誤")
          Txt1(Index).SetFocus
          Txt1(Index).SelStart = 0
          Txt1(Index).SelLength = Len(Txt1(Index))
          Exit Sub
     End Select
Case Else
End Select
End Sub
