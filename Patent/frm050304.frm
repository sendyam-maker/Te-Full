VERSION 5.00
Begin VB.Form frm050304 
   BorderStyle     =   1  '單線固定
   Caption         =   "收文未發文明細表"
   ClientHeight    =   5610
   ClientLeft      =   3165
   ClientTop       =   1200
   ClientWidth     =   3360
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3360
   Begin VB.CheckBox Check1 
      Caption         =   "含FMP外專管制期限"
      Height          =   285
      Left            =   135
      TabIndex        =   20
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox txtPA46 
      Height          =   285
      Left            =   1665
      TabIndex        =   19
      Top             =   4950
      Width           =   285
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   18
      Left            =   1695
      MaxLength       =   1
      TabIndex        =   14
      Text            =   "N"
      Top             =   3900
      Width           =   345
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   17
      Left            =   990
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "1"
      Top             =   3570
      Width           =   345
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2268
      TabIndex        =   22
      Top             =   12
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1488
      TabIndex        =   21
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   16
      Left            =   2040
      MaxLength       =   9
      TabIndex        =   18
      Top             =   4620
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   15
      Left            =   840
      MaxLength       =   9
      TabIndex        =   17
      Top             =   4620
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   2040
      MaxLength       =   9
      TabIndex        =   16
      Top             =   4290
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   840
      MaxLength       =   9
      TabIndex        =   15
      Top             =   4290
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   1695
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3204
      Width           =   345
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2868
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2532
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   960
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2532
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2196
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   960
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2196
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1860
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   960
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1860
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   960
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1524
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   960
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1188
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   2
      Top             =   852
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   960
      MaxLength       =   3
      TabIndex        =   1
      Top             =   852
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   516
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "PCT進入國家階段：　（Y：國家階段）"
      Height          =   180
      Index           =   7
      Left            =   135
      TabIndex        =   42
      Top             =   5010
      Width           =   3180
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "是否含未到期資料："
      Height          =   180
      Left            =   120
      TabIndex        =   41
      Top             =   3930
      Width           =   1620
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "(N:不含)"
      Height          =   180
      Left            =   2295
      TabIndex        =   40
      Top             =   3915
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "列印對象："
      Height          =   180
      Left            =   120
      TabIndex        =   39
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1:承辦人 2:智權人員)"
      Height          =   180
      Left            =   1455
      TabIndex        =   38
      Top             =   3585
      Width           =   1695
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   1920
      TabIndex        =   37
      Top             =   1590
      Width           =   1155
   End
   Begin VB.Line Line6 
      X1              =   1680
      X2              =   1920
      Y1              =   4740
      Y2              =   4740
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   720
   End
   Begin VB.Line Line5 
      X1              =   1680
      X2              =   1920
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   120
      TabIndex        =   35
      Top             =   4320
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "(Y:計算)"
      Height          =   180
      Left            =   2292
      TabIndex        =   34
      Top             =   3216
      Width           =   648
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "是否計算多國案："
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   3240
      Width           =   1440
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "(Y:印)"
      Height          =   180
      Left            =   1896
      TabIndex        =   32
      Top             =   2868
      Width           =   468
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否列印明細："
      Height          =   180
      Left            =   120
      TabIndex        =   31
      Top             =   2880
      Width           =   1260
   End
   Begin VB.Line Line4 
      X1              =   1920
      X2              =   2160
      Y1              =   2652
      Y2              =   2652
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   2535
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   1920
      X2              =   2160
      Y1              =   2316
      Y2              =   2316
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   2190
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   2160
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   120
      TabIndex        =   27
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   1908
      TabIndex        =   26
      Top             =   1248
      Width           =   1152
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   1230
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   2160
      Y1              =   972
      Y2              =   972
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   900
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   570
      Width           =   900
   End
End
Attribute VB_Name = "frm050304"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strSql As String, strTemp1 As Variant, strTemp2 As Variant, StrTest1 As String, StrTest2 As String, i As Integer, j As Integer, s As Integer
Dim PLeft(0 To 11) As Integer, k As Integer, TmpArea As String, iLine As Integer, Page As Integer
Dim strTemp3(0 To 11) As String, iPrint As Integer
Dim StrTest3 As String, Day1 As String, Day2 As String, StrTemp4 As String
Dim St As String, iK As Integer, iTatle As Integer
'Add By Cheng 2003/02/26
Dim iKK As Integer '業務區合計

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
        If Len(txt1(5)) = 0 Then
            s = MsgBox("收文日期(起)不可空白!!", , "USER 輸入錯誤")
            txt1(5).SetFocus
            txt1_GotFocus (5)
            Exit Sub
        'Added by Morgan 2013/4/30 起迄都是必要條件
        ElseIf Len(txt1(6)) = 0 Then
            s = MsgBox("收文日期(迄)不可空白!!", , "USER 輸入錯誤")
            txt1(6).SetFocus
            txt1_GotFocus (6)
            Exit Sub
        'end 2013/4/30
        Else
            'Add By Cheng 2003/02/24
            '若未輸入列印對象
            If Me.txt1(17).Text = "" Then
                s = MsgBox("請輸入列印對象", , "USER 輸入錯誤")
                txt1(17).SetFocus
                Exit Sub
            End If
            If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
                If Left(txt1(13), 6) <> Left(txt1(14), 6) Then
                    s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
                    txt1(13).SetFocus
                    Exit Sub
                End If
            End If
            If Len(Trim(txt1(15))) <> 0 Or Len(Trim(txt1(16))) <> 0 Then
                If Left(txt1(15), 6) <> Left(txt1(16), 6) Then
                    s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
                    txt1(15).SetFocus
                    txt1_GotFocus (15)
                    Exit Sub
                End If
            End If
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/30 清除查詢印表記錄檔欄位
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            'Modify By Cheng 2003/02/24
            '區分列印對象
'            StrMenu
            '列印對象為承辦人
            If Me.txt1(17).Text = "1" Then
                pub_QL05 = pub_QL05 & ";" & Left(Label6, 5) & "承辦人" 'Add By Sindy 2010/9/30
                StrMenu
            '列印對象為智權人員
            Else
                pub_QL05 = pub_QL05 & ";" & Left(Label6, 5) & "智權人員" 'Add By Sindy 2010/9/30
                StrMenu1
            End If
            Me.Enabled = True
            Screen.MousePointer = vbDefault
        End If
    End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub StrMenu()
Screen.MousePointer = vbHourglass
StrTest1 = ""
StrTest2 = ""
StrTest3 = ""
If Len(txt1(0)) <> 0 Then
   StrTest1 = StrTest1 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   StrTest2 = StrTest2 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   StrTest3 = StrTest3 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/9/30
End If
'92.4.2 MODIFY BY SONIA 林錦山, 楊文照己調區, 以前收文案件轉至新業務區列印
'故CP12改為S2.ST15
'If Len(txt1(1)) <> 0 Then
'    StrTest1 = StrTest1 + " AND CP12>='" & txt1(1) & "' "
'    StrTest2 = StrTest2 + " AND CP12>='" & txt1(1) & "' "
'    StrTest3 = StrTest3 + " AND CP12>='" & txt1(1) & "' "
'End If
'If Len(txt1(2)) <> 0 Then
'    StrTest1 = StrTest1 + " AND CP12<='" & txt1(2) & "' "
'    StrTest2 = StrTest2 + " AND CP12<='" & txt1(2) & "' "
'    StrTest3 = StrTest3 + " AND CP12<='" & txt1(2) & "' "
'End If
If Len(txt1(1)) <> 0 Then
    StrTest1 = StrTest1 + " AND S2.ST15>='" & txt1(1) & "' "
    StrTest2 = StrTest2 + " AND S2.ST15>='" & txt1(1) & "' "
    StrTest3 = StrTest3 + " AND S2.ST15>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    StrTest1 = StrTest1 + " AND S2.ST15<='" & txt1(2) & "' "
    StrTest2 = StrTest2 + " AND S2.ST15<='" & txt1(2) & "' "
    StrTest3 = StrTest3 + " AND S2.ST15<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/9/30
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label7 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/9/30
End If
'92.4.2 end
If Len(txt1(3)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP13='" & txt1(3) & "' "
    StrTest2 = StrTest2 + " AND CP13='" & txt1(3) & "' "
    StrTest3 = StrTest3 + " AND CP13='" & txt1(3) & "' "
    pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & lbl1(0)  'Add By Sindy 2010/9/30
End If
If Len(txt1(4)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP14='" & txt1(4) & "' "
    StrTest2 = StrTest2 + " AND CP14='" & txt1(4) & "' "
    StrTest3 = StrTest3 + " AND CP14='" & txt1(4) & "' "
    pub_QL05 = pub_QL05 & ";" & Label5 & txt1(4) & lbl1(1)   'Add By Sindy 2010/9/30
End If

'Add By Cheng 2002/12/12
'除非有指定承辦人, 否則承辦人為陳玲玲及莊敏惠的資料不印
If Me.txt1(4).Text = "" Then
    'Modified by Morgan 2013/10/23 考慮程序新人
    'StrTest1 = StrTest1 + " AND CP14<>'81002' And CP14<>'73017' "
    'StrTest2 = StrTest2 + " AND CP14<>'81002' And CP14<>'73017' "
    'StrTest3 = StrTest3 + " AND CP14<>'81002' And CP14<>'73017' "
    StrTest1 = StrTest1 + " AND NVL(S1.ST05,' ')<>'75' "
    StrTest2 = StrTest2 + " AND NVL(S1.ST05,' ')<>'75' "
    StrTest3 = StrTest3 + " AND NVL(S1.ST05,' ')<>'75' "
    'end 2013/10/23
End If
If Len(txt1(7)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP10>='" & txt1(7) & "' "
    StrTest2 = StrTest2 + " AND CP10>='" & txt1(7) & "' "
    StrTest3 = StrTest3 + " AND CP10>='" & txt1(7) & "' "
End If
If Len(txt1(8)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP10<='" & txt1(8) & "' "
    StrTest2 = StrTest2 + " AND CP10<='" & txt1(8) & "' "
    StrTest3 = StrTest3 + " AND CP10<='" & txt1(8) & "' "
End If
If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label8 & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/9/30
End If
If Len(txt1(9)) <> 0 Then
    StrTest1 = StrTest1 + " AND PA09>='" & txt1(9) & "' "
    StrTest2 = StrTest2 + " AND TM10>='" & txt1(9) & "' "
    StrTest3 = StrTest3 + " AND SP09>='" & txt1(9) & "' "
End If
If Len(txt1(10)) <> 0 Then
    StrTest1 = StrTest1 + " AND PA09<='" & txt1(10) & "' "
    StrTest2 = StrTest2 + " AND TM10<='" & txt1(10) & "' "
    StrTest3 = StrTest3 + " AND SP09<='" & txt1(10) & "' "
End If
If Len(txt1(9)) <> 0 Or Len(txt1(10)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label9 & txt1(9) & "-" & txt1(10) 'Add By Sindy 2010/9/30
End If
If Len(txt1(12)) <> 0 Then
    pub_QL05 = pub_QL05 & ";" & Label12 & "計算" 'Add By Sindy 2010/9/30
Else
    StrTest1 = StrTest1 + " AND CP21 IS NULL "
    StrTest2 = StrTest2 + " AND CP21 IS NULL "
    StrTest3 = StrTest3 + " AND CP21 IS NULL "
End If
If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) <> 0 Then
    StrTest1 = StrTest1 + " AND ((PA26>='" & GetNewFagent(txt1(13)) & "' AND PA26<='" & GetNewFagent(txt1(14)) & "') OR (PA27>='" & GetNewFagent(txt1(13)) & "' AND PA27<='" & GetNewFagent(txt1(14)) & "') OR (PA28>='" & GetNewFagent(txt1(13)) & "' AND PA28<='" & GetNewFagent(txt1(14)) & "') OR (PA29>='" & GetNewFagent(txt1(13)) & "' AND PA29<='" & GetNewFagent(txt1(14)) & "') OR (PA30>='" & GetNewFagent(txt1(13)) & "' AND PA30<='" & GetNewFagent(txt1(14)) & "')) "
'edit by nickc 2007/01/10
'    StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(txt1(13)) & "' AND TM23<='" & GetNewFagent(txt1(14)) & "') "
'    StrTest3 = StrTest3 + " AND ((SP08>='" & GetNewFagent(txt1(13)) & "' AND SP08<='" & GetNewFagent(txt1(14)) & "') OR (SP58<='" & GetNewFagent(txt1(13)) & "' AND SP58<='" & GetNewFagent(txt1(14)) & "') OR (SP59>='" & GetNewFagent(txt1(13)) & "' AND SP59<='" & GetNewFagent(txt1(14)) & "')) "
    StrTest2 = StrTest2 & " AND ((TM23>='" & GetNewFagent(txt1(13)) & "' AND TM23<='" & GetNewFagent(txt1(14)) & "') or (TM78>='" & GetNewFagent(txt1(13)) & "' AND TM78<='" & GetNewFagent(txt1(14)) & "') or (TM79>='" & GetNewFagent(txt1(13)) & "' AND TM79<='" & GetNewFagent(txt1(14)) & "') or (TM80>='" & GetNewFagent(txt1(13)) & "' AND TM80<='" & GetNewFagent(txt1(14)) & "') or (TM81>='" & GetNewFagent(txt1(13)) & "' AND TM81<='" & GetNewFagent(txt1(14)) & "'))"
    StrTest3 = StrTest3 + " AND ((SP08>='" & GetNewFagent(txt1(13)) & "' AND SP08<='" & GetNewFagent(txt1(14)) & "') OR (SP58<='" & GetNewFagent(txt1(13)) & "' AND SP58<='" & GetNewFagent(txt1(14)) & "') OR (SP59>='" & GetNewFagent(txt1(13)) & "' AND SP59<='" & GetNewFagent(txt1(14)) & "') or (SP65>='" & GetNewFagent(txt1(13)) & "' AND SP65<='" & GetNewFagent(txt1(14)) & "') or (SP66>='" & GetNewFagent(txt1(13)) & "' AND SP66<='" & GetNewFagent(txt1(14)) & "')) "
    pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/9/30
'edit by nickc 2007/01/10
'Else
'    If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) = 0 Then
'        StrTest1 = StrTest1 + " AND (PA26>='" & GetNewFagent(txt1(13)) & "' OR PA27>='" & GetNewFagent(txt1(13)) & "' OR PA28>='" & GetNewFagent(txt1(13)) & "' OR PA29>='" & GetNewFagent(txt1(13)) & "' OR PA30>='" & GetNewFagent(txt1(13)) & "') "
'        StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(txt1(13)) & "' ) "
'        StrTest3 = StrTest3 + " AND (SP08>='" & GetNewFagent(txt1(13)) & "' OR SP58>='" & GetNewFagent(txt1(13)) & "' OR SP59>='" & GetNewFagent(txt1(13)) & "') "
'    Else
'        If Len(Trim(txt1(13))) = 0 And Len(Trim(txt1(14))) <> 0 Then
'            StrTest1 = StrTest1 + " AND (PA26<='" & GetNewFagent(txt1(14)) & "' OR PA27<='" & GetNewFagent(txt1(14)) & "' OR PA28<='" & GetNewFagent(txt1(14)) & "' OR PA29<='" & GetNewFagent(txt1(14)) & "' OR PA30<='" & GetNewFagent(txt1(14)) & "') "
'            StrTest2 = StrTest2 & " AND (TM23<='" & GetNewFagent(txt1(14)) & "') "
'            StrTest3 = StrTest3 + " AND (SP08<='" & GetNewFagent(txt1(14)) & "' OR SP58<='" & GetNewFagent(txt1(14)) & "' OR SP59<='" & GetNewFagent(txt1(14)) & "') "
'        End If
'    End If
End If

If Len(Trim(txt1(15))) <> 0 And Len(Trim(txt1(16))) <> 0 Then
    StrTest1 = StrTest1 + " AND ((PA75>='" & GetNewFagent(txt1(15)) & "' AND PA75<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
    StrTest2 = StrTest2 + " AND ((TM44>='" & GetNewFagent(txt1(15)) & "' AND TM44<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
    StrTest3 = StrTest3 + " AND ((SP26>='" & GetNewFagent(txt1(15)) & "' AND SP26<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
Else
    If Len(Trim(txt1(15))) <> 0 And Len(Trim(txt1(16))) = 0 Then
        StrTest1 = StrTest1 + " AND (PA75>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
        StrTest2 = StrTest2 + " AND (TM44>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
        StrTest3 = StrTest3 + " AND (SP26>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
    Else
        If Len(Trim(txt1(15))) = 0 And Len(Trim(txt1(16))) <> 0 Then
            StrTest1 = StrTest1 + " AND (PA75<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
            StrTest2 = StrTest2 + " AND (TM44<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
            StrTest3 = StrTest3 + " AND (SP26<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(15))) <> 0 Or Len(Trim(txt1(16))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label15 & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/9/30
End If
'93.12.7 add by sonia 若有本所期限者, 不能大於系統日
If txt1(18) = "N" Then
   StrTest1 = StrTest1 & " AND ( CP06 Is Null Or CP06<=" & ServerDate & " ) "
   StrTest2 = StrTest2 & " AND ( CP06 Is Null Or CP06<=" & ServerDate & " ) "
   StrTest3 = StrTest3 & " AND ( CP06 Is Null Or CP06<=" & ServerDate & " ) "
   pub_QL05 = pub_QL05 & ";" & Label17 & "不含" 'Add By Sindy 2010/9/30
End If
'93.12.7 end

'Add by Morgan 2005/2/14
If txtPA46 = "Y" Then
   StrTest1 = StrTest1 & " And PA09<>'056' AND PA46='Y' "
   pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 10) & "國家階段"  'Add By Sindy 2010/9/30
End If


'Addec by Morgan 2012/5/22 FMP案改可選擇
If Check1.Value = 1 Then
   pub_QL05 = pub_QL05 & ";" & Check1.Caption
Else
    StrTest1 = StrTest1 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and ep33>0)) "
    StrTest2 = StrTest2 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and ep33>0)) "
    StrTest3 = StrTest3 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and ep33>0)) "
End If
'end 2012/5/22

'Modify By Cheng 2003/09/02
'StrTemp4 = DateSerial(Year(Now), Month(Now), Day(Now) + (Val(txt1(5)) * -1))
'Day1 = ChangeWDateStringToWString(StrTemp4)
'StrTemp4 = DateSerial(Year(Now), Month(Now), Day(Now) + (Val(txt1(6)) * -1))
'Day2 = ChangeWDateStringToWString(StrTemp4)
Day1 = ChangeTStringToWString(Val(txt1(6)))
Day2 = ChangeTStringToWString(Val(txt1(5)))

'Added by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
   strExc(5) = ",(select cp01 v1c1,cp02 v1c2,cp03 v1c3,cp04 v1c4,cp06 v1c6,cp07 v1c7,cp12 v1c8 from casemap,caseprogress where cm10 in ('4','5') and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and ((cm10='4' and cp10='110') or (cm10='5' and cp10 in (" & CaseMapIn & "))) ) VT1 " & _
               ",(select cp01 v2c1,cp02 v2c2,cp03 v2c3,cp04 v2c4,cp06 v2c6,cp07 v2c7,cp12 v2c8 from divisioncase,caseprogress where dc01 in ('P','FCP') and dc01=cp01(+) and dc02=cp02(+) and dc03=cp03(+) and dc04=cp04(+) and cp10 = '307' ) VT2 "
   '判斷條件
   'strExc(6) = "and decode(v2c1,null,decode(v1c1,null,1,decode(v1c6,null,decode(substr(v1c8,1,1),'F',0,1))),decode(v2c6,null,decode(substr(v2c8,1,1),'F',0,1)))=1"
   strExc(6) = "and decode(v1c1||v2c1,null,1,decode(substr(v1c6||v1c8,1,1),'F',0,decode(substr(v2c6||v2c8,1,1),'F',0,1)))=1 "
   '專利案431(PPH)無期限不出現
   'Modified by Lydia 2015/10/02 debug
   'strExc(6) = strExc(6) & " and decode(cc1.cp10||cc1.cp06,'431',0,1)=1 "
   strExc(6) = strExc(6) & " and decode(cp10||cp06,'431',0,1)=1 "
'end 2015/09/09

'Modify By Cheng 2003/01/09
'910819 Sieg 307
'列印明細
If Me.txt1(11).Text = "Y" Then
    pub_QL05 = pub_QL05 & ";" & Label10 & "印"  'Add By Sindy 2010/9/30
    If intPCaseKind = 專利 And intPWhere = 國外_CF Then
        'Modify By Cheng 2003/02/24
        '承辦人名稱改抓代號
'       strSQL = "SELECT S1.ST02 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
'          "PTM03,CPM03,S2.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
'          "PATENTTRADEMARKMAP,NATION WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
'          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
'          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1
        'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
      ' strSql = "SELECT S1.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
          "PTM03,CPM03,S2.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44,S1.ST03 AS D FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
          "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1
       'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
       strSql = "SELECT S1.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
          "PTM03,CPM03,S2.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44,S1.ST03 AS D " & _
          "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS " & _
          strExc(5) & "WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND " & _
          "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & _
          "and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
          strExc(6) & StrTest1
    Else
        'Modify By Cheng 2003/02/24
        '承辦人名稱改抓代號
'       strSQL = "SELECT S1.ST02 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
'          "DECODE(PA09,'000',PTM03,PTM04),CPM03,S2.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
'          "PATENTTRADEMARKMAP,NATION WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
'          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
'          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1
       'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
'       strSql = "SELECT S1.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
          "DECODE(PA09,'000',PTM03,PTM04),CPM03,S2.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44,S1.ST03 AS D FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
          "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1
       'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
       strSql = "SELECT S1.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
          "DECODE(PA09,'000',PTM03,PTM04),CPM03,S2.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44,S1.ST03 AS D " & _
          "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS " & _
          strExc(5) & "WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND " & _
          "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & _
          "and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
          strExc(6) & StrTest1
          
    End If
    'Modify By Cheng 2003/02/24
    '承辦人名稱改抓代號
'    strSQL = strSQL + " union all select S1.ST02 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07))," & _
'       "'','',CP06,CP07,PTM03,CPM03,S2.ST02,NA03,TM23,'','','','',TM44,CP44 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
'       "PATENTTRADEMARKMAP,NATION WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
'       "CP57 IS NULL AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & _
'       "cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2
'    strSQL = strSQL + " union all select S1.ST02 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07))," & _
'       "'','',CP06,CP07,'',CPM03,S2.ST02,NA03,SP08,SP58,SP59,'','',SP26,CP44 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
'       "STAFF S2,NATION WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL AND " & _
'       "cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp01=cpm01(+) AND " & _
'       "CP10=CPM02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3
    '2009/3/11 modify by sonia 取消cp14 is not null條件
    'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
    strSql = strSql + " union all select S1.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07))," & _
       "'','',CP06,CP07,PTM03,CPM03,S2.ST02,NA03,TM23,'','','','',TM44,CP44,S1.ST03 AS D FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
       "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND " & _
       "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & _
       "cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2
    '2009/3/11 modify by sonia 取消cp14 is not null條件
    'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
    strSql = strSql + " union all select S1.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07))," & _
       "'','',CP06,CP07,'',CPM03,S2.ST02,NA03,SP08,SP58,SP59,'','',SP26,CP44,S1.ST03 AS D FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
       "STAFF S2,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND " & _
       "cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp01=cpm01(+) AND " & _
       "CP10=CPM02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3
    strSql = strSql + " ORDER BY D,A,B,C "
'不列印明細
Else
    If intPCaseKind = 專利 And intPWhere = 國外_CF Then
        'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
      ' strSql = "SELECT CP14,S1.ST02,COUNT(CP14)  FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
          "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1 & _
          " Group By CP14,S1.ST02 "
       'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
       strSql = "SELECT CP14,S1.ST02,sum(decode(v2c1,null,decode(v1c1,null,1,decode(v1c6,null,0,1)),decode(v2c6,null,0,1))) cnt " & _
          "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS " & _
          strExc(5) & "WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND " & _
          "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & _
          " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
          strExc(6) & StrTest1 & " Group By CP14,S1.ST02 "
    Else
        'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
     '  strSql = "SELECT CP14,S1.ST02,COUNT(CP14) FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
          "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1 & _
          " Group By CP14,S1.ST02 "
       'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
       strSql = "SELECT CP14,S1.ST02,sum(decode(v2c1,null,decode(v1c1,null,1,decode(v1c6,null,0,1)),decode(v2c6,null,0,1))) cnt " & _
          "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS " & _
          strExc(5) & "WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND " & _
          "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & _
          " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
          strExc(6) & StrTest1 & " Group By CP14,S1.ST02 "
    End If
    
    '2009/3/11 modify by sonia 取消cp14 is not null條件
    'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
    strSql = strSql + " union all select CP14,S1.ST02,COUNT(CP14) FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
        "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND " & _
        "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & _
        "cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2 & _
        " Group By CP14,S1.ST02 "

    '2009/3/11 modify by sonia 取消cp14 is not null條件
    'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
    strSql = strSql + " union all select CP14,S1.ST02,COUNT(CP14) FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
        "STAFF S2,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND " & _
        "cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp01=cpm01(+) AND " & _
        "CP10=CPM02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3 & _
        " Group By CP14,S1.ST02 "
    strSql = strSql + " ORDER BY 1 "
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/30
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
    adoRecordset.MoveNext
    Loop
    'Modify By Cheng 2003/01/09
'    StrPrintDoc       '列印主程式
    '若選擇列印明細
    If Me.txt1(11).Text = "Y" Then
        StrPrintDoc       '列印主程式
    '若選擇不列印明細
    Else
        StrPrintDocTotal
    End If
    CheckOC
Else
    InsertQueryLog (0) 'Add By Sindy 2010/9/30
    ShowNoData
    CheckOC
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Screen.MousePointer = vbDefault
End Sub

'承辦人明細
Sub StrPrintDoc()
Dim strTempName As String '代理人名稱

GetPrintLeft
iLine = 1
Page = 1
If txt1(12) = "Y" Or txt1(12) = "y" Then
    TmpArea = "(含多國案)"
Else
    TmpArea = ""
End If
StrPrintTital TmpArea, str(Page)
iPrint = 2700
iTatle = 0       ' 總筆數
iK = 0           ' 小計
With adoRecordset
    .MoveFirst
    Do While .EOF = False
        For j = 0 To 11
            If Not IsNull(.Fields(j)) Then
                strTemp3(j) = .Fields(j)
            Else
                strTemp3(j) = ""
            End If
        Next j
        If Not IsNull(adoRecordset.Fields(12)) Then
            strTemp3(4) = adoRecordset.Fields(12)
        Else
            If Not IsNull(adoRecordset.Fields(13)) Then
                strTemp3(4) = adoRecordset.Fields(13)
            Else
                If Not IsNull(adoRecordset.Fields(14)) Then
                    strTemp3(4) = adoRecordset.Fields(14)
                Else
                    If Not IsNull(adoRecordset.Fields(15)) Then
                        strTemp3(4) = adoRecordset.Fields(15)
                    Else
                        If Not IsNull(adoRecordset.Fields(16)) Then
                            strTemp3(4) = adoRecordset.Fields(16)
                        End If
                    End If
                End If
            End If
        End If
        strTemp3(4) = GetPrjPeople1(strTemp3(4))
        If Not IsNull(adoRecordset.Fields(17)) Then
            strTemp3(5) = adoRecordset.Fields(17)
        Else
            If Not IsNull(adoRecordset.Fields(18)) Then
                strTemp3(5) = adoRecordset.Fields(18)
            End If
        End If
      'Modify By Cheng 2002/07/05
      '若系統種類對照檔SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'        strTemp3(5) = GetPrjName1(strTemp3(5))
      If PUB_GetAgentName(SystemNumber(strTemp3(2), 1), strTemp3(5), strTempName) = True Then
           strTemp3(5) = strTempName
      Else
           strTemp3(5) = ""
      End If
        St = strTemp3(0)
        iK = iK + 1
        iTatle = iTatle + 1
        If Len(strTemp3(1)) > 7 Then
            strTemp3(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(1)))
        End If
        If Len(strTemp3(6)) > 7 Then
            strTemp3(6) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(6)))
        End If
        If Len(strTemp3(7)) > 7 Then
            strTemp3(7) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(7)))
        End If
        'Printer.Font.Name = "Arial"
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        If iK = 1 Then
            'Modify By Cheng 2003/02/24
            '取得承辦人姓名
'           Printer.Print StrToStr(strTemp3(0), 4)
           Printer.Print StrToStr(GetStaffName(strTemp3(0), True), 4)
        Else
           Printer.Print ""
        End If
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(1)
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(2)
        Printer.CurrentX = PLeft(3)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(3), 4)
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(4), 6)
        Printer.CurrentX = PLeft(5)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(5), 7)
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(6)
        Printer.CurrentX = PLeft(7)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(7)
        Printer.CurrentX = PLeft(8)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(8), 4)
        Printer.CurrentX = PLeft(9)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(9), 4)
        Printer.CurrentX = PLeft(10)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(10), 4)
        Printer.CurrentX = PLeft(11)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(11), 4)
        .MoveNext
        If .EOF = False Then
            If Not IsNull(.Fields(0)) Then
                StrTest1 = .Fields(0)
            Else
                StrTest1 = ""
            End If
        End If
        If .EOF = False Then
            If StrTest1 <> St Then
            
               'Add by Morgan 2003/12/17
               If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
                  Printer.NewPage
                  Page = Page + 1
                  StrPrintTital TmpArea, str(Page)
                  iPrint = 2400
                  iLine = 0
               End If
               iLine = iLine + 1
               'Add 2003/12/17
               
                iPrint = iPrint + 300
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                
               'Add by Morgan 2003/12/17
               If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
                  Printer.NewPage
                  Page = Page + 1
                  StrPrintTital TmpArea, str(Page)
                  iPrint = 2400
                  iLine = 0
               End If
               iLine = iLine + 1
               'Add 2003/12/17
               
                iPrint = iPrint + 300
                Printer.CurrentX = 1000
                Printer.CurrentY = iPrint
                Printer.Print "小計： " & Trim(str(iK)) & " 筆"
                iK = 0
                
               'Add by Morgan 2003/12/17
               If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
                  Printer.NewPage
                  Page = Page + 1
                  StrPrintTital TmpArea, str(Page)
                  iPrint = 2400
                  iLine = 0
               End If
               iLine = iLine + 1
               'Add 2003/12/17
               
                iPrint = iPrint + 300
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                
                'Remove by  Morgan 2003/12/17
                'iLine = iLine + 3
                
                St = StrTest1
            End If
            'Modify by Morgan 2003/12/17
            'If (iLine Mod 20 = 0) Or iPrint >= 10000 Then
            If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
            'Modify 2003/12/17
                iPrint = iPrint + 300
                'Printer.CurrentX = 500
                'Printer.CurrentY = iPrint
                'Printer.Print String(200, "-")
                'iPrint = iPrint + 300
                'Printer.CurrentX = 1000
                'Printer.CurrentY = iPrint
                'Printer.Print "小計： " & Trim(Str(iK)) & " 筆"
                'iK = 0
                'iPrint = iPrint + 300
                'Printer.CurrentX = 500
                'Printer.CurrentY = iPrint
                'Printer.Print String(200, "-")
                'StrPrintEnd
                Printer.NewPage
                Page = Page + 1
                StrPrintTital TmpArea, str(Page)
                iPrint = 2400
                iLine = 0
            End If
            iLine = iLine + 1
            iPrint = iPrint + 300
        End If
    Loop
End With
'合計

'Add by Morgan 2003/12/17
iLine = iLine + 1
If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
   Printer.NewPage
   Page = Page + 1
   StrPrintTital TmpArea, str(Page)
   iPrint = 2400
   iLine = 0
End If

iLine = iLine + 1
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
 
'Add by Morgan 2003/12/17
If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
   Printer.NewPage
   Page = Page + 1
   StrPrintTital TmpArea, str(Page)
   iPrint = 2400
   iLine = 0
End If

iLine = iLine + 1
iPrint = iPrint + 300
Printer.CurrentX = 1000
Printer.CurrentY = iPrint
Printer.Print "小計： " & Trim(str(iK)) & " 筆"

If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
   Printer.NewPage
   Page = Page + 1
   StrPrintTital TmpArea, str(Page)
   iPrint = 2400
   iLine = 0
End If
iLine = iLine + 1
'Add 2003/12/17
 
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")

'Add by Morgan 2003/12/17
If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
   Printer.NewPage
   Page = Page + 1
   StrPrintTital TmpArea, str(Page)
   iPrint = 2400
   iLine = 0
End If
'Add 2003/12/17
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "合計：共 " & Trim(str(iTatle)) & " 筆"
Printer.EndDoc
ShowPrintOk
CheckOC


End Sub

'Add By Cheng 2003/01/09
Sub StrPrintDocTotal()
Dim strTempName As String '代理人名稱

GetPrintLeft
iLine = 1
Page = 1
If txt1(12) = "Y" Or txt1(12) = "y" Then
    TmpArea = "(含多國案)"
Else
    TmpArea = ""
End If
StrPrintTital TmpArea, str(Page)
iPrint = 2700
iTatle = 0       ' 總數
With adoRecordset
    .MoveFirst
    Do While .EOF = False
        For j = 0 To 2
            If Not IsNull(.Fields(j)) Then
                strTemp3(j) = .Fields(j)
            Else
                strTemp3(j) = ""
            End If
        Next j
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(1)
        Printer.CurrentX = PLeft(2) + 750 - TextWidth(Format(strTemp3(2), "##0") & " 筆")
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp3(2), "##0") & " 筆"
        iTatle = iTatle + Val(strTemp3(2))
        .MoveNext
        If (iLine Mod 20 = 0) Or iPrint >= 10000 Then
            iPrint = iPrint + 300
            Printer.NewPage
            Page = Page + 1
            StrPrintTital TmpArea, str(Page)
            iPrint = 2400
            iLine = 0
        End If
        iLine = iLine + 1
        iPrint = iPrint + 300
    Loop
End With
'合計
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(2) + 750 - TextWidth("合計：共 " & Format(iTatle, "##0") & " 筆")
Printer.CurrentY = iPrint
Printer.Print "合計：共 " & Format(iTatle, "##0") & " 筆"
Printer.EndDoc
ShowPrintOk
CheckOC
End Sub

Sub StrPrintTital(ByRef Area As String, ByRef Page As String)
GetPrintLeft
k = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = i
Printer.Print "收文未發文明細表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6800
Printer.CurrentY = k + 500
Printer.Print "收文日期：" & Format(txt1(5) & " ", "@@@@@") & "-" & txt1(6)
Printer.Font.Bold = False
Printer.CurrentX = 0
Printer.CurrentY = k + 800
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = k + 800
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.CurrentX = 0
Printer.CurrentY = k + 1100
Printer.Print Area
Printer.CurrentX = 13000
Printer.CurrentY = k + 1100
Printer.Print "頁    次：" & Page
Printer.CurrentX = 0
Printer.CurrentY = k + 1400
Printer.Print String(200, "-")
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = k + 1700
'Modify By Cheng 2003/02/25
'標題依列印對象不同而不同
'Printer.Print "承辦人"
Printer.Print IIf(Me.txt1(17).Text = "1", "承辦人", "智權人員")
Printer.CurrentX = PLeft(1)
Printer.CurrentY = k + 1700
Printer.Print "收文日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = k + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = k + 1700
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = k + 1700
Printer.Print "申請人"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = k + 1700
Printer.Print "代理人"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = k + 1700
Printer.Print "本所期限"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = k + 1700
Printer.Print "法定期限"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = k + 1700
Printer.Print "種類"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = k + 1700
Printer.Print "案件性質"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = k + 1700
'Modify By Cheng 2003/02/25
'標題依列印對象不同而不同
'Printer.Print "智權人員"
Printer.Print IIf(Me.txt1(17).Text = "1", "智權人員", "承辦人")
Printer.CurrentX = PLeft(11)
Printer.CurrentY = k + 1700
Printer.Print "申請國家"
Printer.Font.Underline = False
Printer.CurrentX = 0
Printer.CurrentY = k + 2000
Printer.Print String(200, "-")
End Sub

Sub StrPrintEnd()
End Sub
Sub GetPrintLeft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 2100 '2400
PLeft(3) = 4200 '4500
PLeft(4) = 5400 '5700
PLeft(5) = 7000 '7300
PLeft(6) = 8900 '9200
PLeft(7) = 10100 '10200
PLeft(8) = 11300
PLeft(9) = 12400
PLeft(10) = 13500
PLeft(11) = 14500
End Sub
Private Sub Form_Load()
MoveFormToCenter Me
txt1(12) = "Y"
strTemp1 = Split(UCase(GetSystemKindByNick), ",")
For i = 0 To UBound(strTemp1)
    If strTemp1(i) <> "FCP" Then
        txt1(0) = txt1(0) + strTemp1(i) + ","
    End If
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050304 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'Add By Cheng 2003/02/24
Select Case Index
Case 17 '列印對象
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
        'If UCase(StrTemp2(I)) = "FCP" Then
        '    S = MsgBox("此功能無法查詢  FCP 的報表，請從 FCP 系統進入", , "報表格式不同")
        '    TXT1(0).SetFocus
        '    TXT1(0).SelStart = 0
        '    TXT1(0).SelLength = Len(TXT1(0))
        '    Exit Sub
        ' End If
    Next i
    
Case 14
     'If Trim(txt1(13)) <> "" And Trim(txt1(14)) <> "" Then
        If Left(txt1(13), 6) <> Left(txt1(14), 6) Then
            s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
            txt1(13).SetFocus
            Exit Sub
        End If
      'End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 16
     'If Trim(txt1(15)) <> "" And Trim(txt1(16)) <> "" Then
        If Left(txt1(15), 6) <> Left(txt1(16), 6) Then
            s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
            txt1(15).SetFocus
            Exit Sub
        End If
      'End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 11
     Select Case Trim(txt1(11))
     Case "y", "Y", ""
     Case Else
         s = MsgBox("是否列印明細只能輸入 Y 或空白 !!", , "USER 輸入錯誤")
         txt1(11).SetFocus
         txt1(11).SelStart = 0
         txt1(11).SelLength = Len(txt1(11))
         Exit Sub
     End Select
Case 12
     Select Case Trim(txt1(12))
     Case "y", "Y", ""
     Case Else
         s = MsgBox("是否計算多國家只能輸入 Y 或空白 !!", , "USER 輸入錯誤")
         txt1(12).SetFocus
         txt1(12).SelStart = 0
         txt1(12).SelLength = Len(txt1(12))
         Exit Sub
     End Select
Case 3
     lbl1(0) = GetPrjSalesNM(txt1(Index))
     If Len(txt1(Index)) <> 0 Then
        If Len(lbl1(0).Caption) = 0 Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 4
     lbl1(1) = GetPrjSalesNM(txt1(Index))
     If Len(txt1(Index)) <> 0 Then
        If Len(lbl1(1).Caption) = 0 Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 2, 8, 10
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 5, 6
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 6 Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
End Select

End Sub

'Add By Cheng 2003/02/24
'列印對象為智權人員
Sub StrMenu1()
Screen.MousePointer = vbHourglass
StrTest1 = ""
StrTest2 = ""
StrTest3 = ""
If Len(txt1(0)) <> 0 Then
   StrTest1 = StrTest1 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   StrTest2 = StrTest2 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   StrTest3 = StrTest3 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/9/30
End If
'92.4.2 MODIFY BY SONIA 林錦山, 楊文照己調區, 以前收文案件轉至新業務區列印
'故CP12改為S2.ST15
'If Len(txt1(1)) <> 0 Then
'    StrTest1 = StrTest1 + " AND CP12>='" & txt1(1) & "' "
'    StrTest2 = StrTest2 + " AND CP12>='" & txt1(1) & "' "
'    StrTest3 = StrTest3 + " AND CP12>='" & txt1(1) & "' "
'End If
'If Len(txt1(2)) <> 0 Then
'    StrTest1 = StrTest1 + " AND CP12<='" & txt1(2) & "' "
'    StrTest2 = StrTest2 + " AND CP12<='" & txt1(2) & "' "
'    StrTest3 = StrTest3 + " AND CP12<='" & txt1(2) & "' "
'End If
If Len(txt1(1)) <> 0 Then
    StrTest1 = StrTest1 + " AND S2.ST15>='" & txt1(1) & "' "
    StrTest2 = StrTest2 + " AND S2.ST15>='" & txt1(1) & "' "
    StrTest3 = StrTest3 + " AND S2.ST15>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    StrTest1 = StrTest1 + " AND S2.ST15<='" & txt1(2) & "' "
    StrTest2 = StrTest2 + " AND S2.ST15<='" & txt1(2) & "' "
    StrTest3 = StrTest3 + " AND S2.ST15<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/9/30
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label7 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/9/30
End If
'92.4.2 end
If Len(txt1(3)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP13='" & txt1(3) & "' "
    StrTest2 = StrTest2 + " AND CP13='" & txt1(3) & "' "
    StrTest3 = StrTest3 + " AND CP13='" & txt1(3) & "' "
    pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & lbl1(0)  'Add By Sindy 2010/9/30
End If
If Len(txt1(4)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP14='" & txt1(4) & "' "
    StrTest2 = StrTest2 + " AND CP14='" & txt1(4) & "' "
    StrTest3 = StrTest3 + " AND CP14='" & txt1(4) & "' "
    pub_QL05 = pub_QL05 & ";" & Label5 & txt1(4) & lbl1(1)   'Add By Sindy 2010/9/30
End If
If Len(txt1(7)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP10>='" & txt1(7) & "' "
    StrTest2 = StrTest2 + " AND CP10>='" & txt1(7) & "' "
    StrTest3 = StrTest3 + " AND CP10>='" & txt1(7) & "' "
End If
If Len(txt1(8)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP10<='" & txt1(8) & "' "
    StrTest2 = StrTest2 + " AND CP10<='" & txt1(8) & "' "
    StrTest3 = StrTest3 + " AND CP10<='" & txt1(8) & "' "
End If
If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label8 & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/9/30
End If
If Len(txt1(9)) <> 0 Then
    StrTest1 = StrTest1 + " AND PA09>='" & txt1(9) & "' "
    StrTest2 = StrTest2 + " AND TM10>='" & txt1(9) & "' "
    StrTest3 = StrTest3 + " AND SP09>='" & txt1(9) & "' "
End If
If Len(txt1(10)) <> 0 Then
    StrTest1 = StrTest1 + " AND PA09<='" & txt1(10) & "' "
    StrTest2 = StrTest2 + " AND TM10<='" & txt1(10) & "' "
    StrTest3 = StrTest3 + " AND SP09<='" & txt1(10) & "' "
End If
If Len(txt1(9)) <> 0 Or Len(txt1(10)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label9 & txt1(9) & "-" & txt1(10) 'Add By Sindy 2010/9/30
End If
If Len(txt1(12)) <> 0 Then
    pub_QL05 = pub_QL05 & ";" & Label12 & "計算" 'Add By Sindy 2010/9/30
Else
    StrTest1 = StrTest1 + " AND CP21 IS NULL "
    StrTest2 = StrTest2 + " AND CP21 IS NULL "
    StrTest3 = StrTest3 + " AND CP21 IS NULL "
End If
If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) <> 0 Then
    StrTest1 = StrTest1 + " AND ((PA26>='" & GetNewFagent(txt1(13)) & "' AND PA26<='" & GetNewFagent(txt1(14)) & "') OR (PA27>='" & GetNewFagent(txt1(13)) & "' AND PA27<='" & GetNewFagent(txt1(14)) & "') OR (PA28>='" & GetNewFagent(txt1(13)) & "' AND PA28<='" & GetNewFagent(txt1(14)) & "') OR (PA29>='" & GetNewFagent(txt1(13)) & "' AND PA29<='" & GetNewFagent(txt1(14)) & "') OR (PA30>='" & GetNewFagent(txt1(13)) & "' AND PA30<='" & GetNewFagent(txt1(14)) & "')) "
'edit by nickc 2007/01/10
'    StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(txt1(13)) & "' AND TM23<='" & GetNewFagent(txt1(14)) & "') "
'    StrTest3 = StrTest3 + " AND ((SP08>='" & GetNewFagent(txt1(13)) & "' AND SP08<='" & GetNewFagent(txt1(14)) & "') OR (SP58<='" & GetNewFagent(txt1(13)) & "' AND SP58<='" & GetNewFagent(txt1(14)) & "') OR (SP59>='" & GetNewFagent(txt1(13)) & "' AND SP59<='" & GetNewFagent(txt1(14)) & "')) "
    StrTest2 = StrTest2 & " AND ((TM23>='" & GetNewFagent(txt1(13)) & "' AND TM23<='" & GetNewFagent(txt1(14)) & "') or (TM78>='" & GetNewFagent(txt1(13)) & "' AND TM78<='" & GetNewFagent(txt1(14)) & "') or (TM79>='" & GetNewFagent(txt1(13)) & "' AND TM79<='" & GetNewFagent(txt1(14)) & "') or (TM80>='" & GetNewFagent(txt1(13)) & "' AND TM80<='" & GetNewFagent(txt1(14)) & "') or (TM81>='" & GetNewFagent(txt1(13)) & "' AND TM81<='" & GetNewFagent(txt1(14)) & "'))"
    StrTest3 = StrTest3 + " AND ((SP08>='" & GetNewFagent(txt1(13)) & "' AND SP08<='" & GetNewFagent(txt1(14)) & "') OR (SP58<='" & GetNewFagent(txt1(13)) & "' AND SP58<='" & GetNewFagent(txt1(14)) & "') OR (SP59>='" & GetNewFagent(txt1(13)) & "' AND SP59<='" & GetNewFagent(txt1(14)) & "') or (SP65>='" & GetNewFagent(txt1(13)) & "' AND SP65<='" & GetNewFagent(txt1(14)) & "') or (SP66>='" & GetNewFagent(txt1(13)) & "' AND SP66<='" & GetNewFagent(txt1(14)) & "')) "
    pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/9/30
'edit by nickc 2007/01/10
'Else
'    If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) = 0 Then
'        StrTest1 = StrTest1 + " AND (PA26>='" & GetNewFagent(txt1(13)) & "' OR PA27>='" & GetNewFagent(txt1(13)) & "' OR PA28>='" & GetNewFagent(txt1(13)) & "' OR PA29>='" & GetNewFagent(txt1(13)) & "' OR PA30>='" & GetNewFagent(txt1(13)) & "') "
'        StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(txt1(13)) & "' ) "
'        StrTest3 = StrTest3 + " AND (SP08>='" & GetNewFagent(txt1(13)) & "' OR SP58>='" & GetNewFagent(txt1(13)) & "' OR SP59>='" & GetNewFagent(txt1(13)) & "') "
'    Else
'        If Len(Trim(txt1(13))) = 0 And Len(Trim(txt1(14))) <> 0 Then
'            StrTest1 = StrTest1 + " AND (PA26<='" & GetNewFagent(txt1(14)) & "' OR PA27<='" & GetNewFagent(txt1(14)) & "' OR PA28<='" & GetNewFagent(txt1(14)) & "' OR PA29<='" & GetNewFagent(txt1(14)) & "' OR PA30<='" & GetNewFagent(txt1(14)) & "') "
'            StrTest2 = StrTest2 & " AND (TM23<='" & GetNewFagent(txt1(14)) & "') "
'            StrTest3 = StrTest3 + " AND (SP08<='" & GetNewFagent(txt1(14)) & "' OR SP58<='" & GetNewFagent(txt1(14)) & "' OR SP59<='" & GetNewFagent(txt1(14)) & "') "
'        End If
'    End If
End If
If Len(Trim(txt1(15))) <> 0 And Len(Trim(txt1(16))) <> 0 Then
    StrTest1 = StrTest1 + " AND ((PA75>='" & GetNewFagent(txt1(15)) & "' AND PA75<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
    StrTest2 = StrTest2 + " AND ((TM44>='" & GetNewFagent(txt1(15)) & "' AND TM44<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
    StrTest3 = StrTest3 + " AND ((SP26>='" & GetNewFagent(txt1(15)) & "' AND SP26<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
Else
    If Len(Trim(txt1(15))) <> 0 And Len(Trim(txt1(16))) = 0 Then
        StrTest1 = StrTest1 + " AND (PA75>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
        StrTest2 = StrTest2 + " AND (TM44>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
        StrTest3 = StrTest3 + " AND (SP26>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
    Else
        If Len(Trim(txt1(15))) = 0 And Len(Trim(txt1(16))) <> 0 Then
            StrTest1 = StrTest1 + " AND (PA75<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
            StrTest2 = StrTest2 + " AND (TM44<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
            StrTest3 = StrTest3 + " AND (SP26<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(15))) <> 0 Or Len(Trim(txt1(16))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label15 & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/9/30
End If
'93.12.7 add by sonia 若有本所期限者, 不能大於系統日
If txt1(18) = "N" Then
   StrTest1 = StrTest1 & " AND ( CP06 Is Null Or CP06<=" & ServerDate & " ) "
   StrTest2 = StrTest2 & " AND ( CP06 Is Null Or CP06<=" & ServerDate & " ) "
   StrTest3 = StrTest3 & " AND ( CP06 Is Null Or CP06<=" & ServerDate & " ) "
   pub_QL05 = pub_QL05 & ";" & Label17 & "不含" 'Add By Sindy 2010/9/30
End If
'93.12.7 end

'Add by Morgan 2005/2/4
If txtPA46 = "Y" Then
   StrTest1 = StrTest1 & " And PA09<>'056' AND PA46='Y' "
   pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 10) & "國家階段"  'Add By Sindy 2010/9/30
End If

'Addec by Morgan 2012/5/22 FMP案改可選擇
If Check1.Value = 1 Then
   pub_QL05 = pub_QL05 & ";" & Check1.Caption
Else
    StrTest1 = StrTest1 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and ep33>0)) "
    StrTest2 = StrTest2 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and ep33>0)) "
    StrTest3 = StrTest3 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and ep33>0)) "
End If
'end 2012/5/22

'Modify By Cheng 2003/09/02
'StrTemp4 = DateSerial(Year(Now), Month(Now), Day(Now) + (Val(txt1(5)) * -1))
'Day1 = ChangeWDateStringToWString(StrTemp4)
'StrTemp4 = DateSerial(Year(Now), Month(Now), Day(Now) + (Val(txt1(6)) * -1))
'Day2 = ChangeWDateStringToWString(StrTemp4)
Day1 = ChangeTStringToWString(Val(txt1(6)))
Day2 = ChangeTStringToWString(Val(txt1(5)))
'Added by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
   strExc(5) = ",(select cp01 v1c1,cp02 v1c2,cp03 v1c3,cp04 v1c4,cp06 v1c6,cp07 v1c7,cp12 v1c8 from casemap,caseprogress where cm10 in ('4','5') and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and ((cm10='4' and cp10='110') or (cm10='5' and cp10 in (" & CaseMapIn & "))) ) VT1 " & _
               ",(select cp01 v2c1,cp02 v2c2,cp03 v2c3,cp04 v2c4,cp06 v2c6,cp07 v2c7,cp12 v2c8 from divisioncase,caseprogress where dc01 in ('P','FCP') and dc01=cp01(+) and dc02=cp02(+) and dc03=cp03(+) and dc04=cp04(+) and cp10 = '307' ) VT2 "
    '判斷條件
   'strExc(6) = "and decode(v2c1,null,decode(v1c1,null,1,decode(v1c6,null,decode(substr(v1c8,1,1),'F',0,1))),decode(v2c6,null,decode(substr(v2c8,1,1),'F',0,1)))=1"
   strExc(6) = "and decode(v1c1||v2c1,null,1,decode(substr(v1c6||v1c8,1,1),'F',0,decode(substr(v2c6||v2c8,1,1),'F',0,1)))=1 "
   '專利案431(PPH)無期限不出現
      'Modified by Lydia 2015/10/02 debug
   'strExc(6) = strExc(6) & " and decode(cc1.cp10||cc1.cp06,'431',0,1)=1 "
   strExc(6) = strExc(6) & " and decode(cp10||cp06,'431',0,1)=1 "
'Modify By Cheng 2003/01/09
'910819 Sieg 307
'列印明細
If Me.txt1(11).Text = "Y" Then
    pub_QL05 = pub_QL05 & ";" & Label10 & "印"  'Add By Sindy 2010/9/30
    If intPCaseKind = 專利 And intPWhere = 國外_CF Then
        'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
      ' strSql = "SELECT S2.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
          "PTM03,CPM03,S1.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44,CP12,A0902,S2.ST15 AS D FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
          "PATENTTRADEMARKMAP,NATION,ACC090,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND CP12=A0901(+) " & StrTest1
        'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
        strSql = "SELECT S2.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
          "PTM03,CPM03,S1.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44,CP12,A0902,S2.ST15 AS D " & _
          "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ACC090,ENGINEERPROGRESS " & _
          strExc(5) & "WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND " & _
          "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND CP12=A0901(+) " & _
          "and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
          strExc(6) & StrTest1
    Else
        'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
    '   strSql = "SELECT S2.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
          "DECODE(PA09,'000',PTM03,PTM04),CPM03,S1.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44,CP12,A0902,S2.ST15 AS D FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
          "PATENTTRADEMARKMAP,NATION,ACC090,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND CP12=A0901(+) " & StrTest1
        'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
        strSql = "SELECT S2.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',CP06,CP07," & _
          "DECODE(PA09,'000',PTM03,PTM04),CPM03,S1.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,CP44,CP12,A0902,S2.ST15 AS D " & _
          "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ACC090,ENGINEERPROGRESS " & _
          strExc(5) & "WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND " & _
          "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND CP12=A0901(+) " & _
          "and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
          strExc(6) & StrTest1
    End If
    
    '2009/3/11 modify by sonia 取消cp14 is not null條件
    'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
    strSql = strSql + " union all select S2.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07))," & _
       "'','',CP06,CP07,PTM03,CPM03,S1.ST02,NA03,TM23,'','','','',TM44,CP44,CP12,A0902,S2.ST15 AS D FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
       "PATENTTRADEMARKMAP,NATION,ACC090,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND " & _
       "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & _
       "cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND (TM29<>'Y' OR TM29 IS NULL) AND CP12=A0901(+) " & StrTest2
    '2009/3/11 modify by sonia 取消cp14 is not null條件
    'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
    strSql = strSql + " union all select S2.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07))," & _
       "'','',CP06,CP07,'',CPM03,S1.ST02,NA03,SP08,SP58,SP59,'','',SP26,CP44,CP12,A0902,S2.ST15 AS D FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
       "STAFF S2,NATION,ACC090,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND " & _
       "cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp01=cpm01(+) AND " & _
       "CP10=CPM02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) AND CP12=A0901(+) " & StrTest3
    strSql = strSql + " ORDER BY D,A,B,C "
'不列印明細
Else
    If intPCaseKind = 專利 And intPWhere = 國外_CF Then
        'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
       'strSql = "SELECT CP13,S2.ST02,COUNT(CP13),CP12  FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
          "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1 & _
          " Group By CP13,S2.ST02,CP12 "
       'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
       strSql = "SELECT CP13,S2.ST02,sum(decode(v2c1,null,decode(v1c1,null,1,decode(v1c6,null,0,1)),decode(v2c6,null,0,1))) cnt,CP12 " & _
          "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS " & _
          strExc(5) & "WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND " & _
          "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & _
          " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
          strExc(6) & StrTest1 & " Group By CP13,S2.ST02,CP12 "
    Else
       'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
'       strSql = "SELECT CP13,S2.ST02,COUNT(CP13),CP12 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
          "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND " & _
          "CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1 & _
          " Group By CP13,S2.ST02,CP12 "
       'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
       strSql = "SELECT CP13,S2.ST02,sum(decode(v2c1,null,decode(v1c1,null,1,decode(v1c6,null,0,1)),decode(v2c6,null,0,1))) cnt,CP12 " & _
          "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS " & _
          strExc(5) & "WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND " & _
          "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
          "AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) " & _
          " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
          strExc(6) & StrTest1 & " Group By CP13,S2.ST02,CP12 "
    End If
    
    '2009/3/11 modify by sonia 取消cp14 is not null條件
    'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
    strSql = strSql + " union all select CP13,S2.ST02,COUNT(CP13),CP12 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
        "PATENTTRADEMARKMAP,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND " & _
        "CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & _
        "cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2 & _
        " Group By CP13,S2.ST02,CP12 "

    '2009/3/11 modify by sonia 取消cp14 is not null條件
    'Modified by Lydia 2016/12/21 +排除D類收文 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'
    strSql = strSql + " union all select CP13,S2.ST02,COUNT(CP13),CP12 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
        "STAFF S2,NATION,ENGINEERPROGRESS WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D' AND " & _
        "cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp01=cpm01(+) AND " & _
        "CP10=CPM02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3 & _
        " Group By CP13,S2.ST02,CP12 "
    strSql = strSql + " ORDER BY 4, 1 "
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/30
    adoRecordset.MoveFirst
    'Marked By Cheng 2004/05/05
'    Do While adoRecordset.EOF = False
'    adoRecordset.MoveNext
'    Loop
    'End
    '若選擇列印明細
    If Me.txt1(11).Text = "Y" Then
        StrPrintDoc1       '列印主程式
    '若選擇不列印明細
    Else
        StrPrintDocTotal1
    End If
    CheckOC
Else
    InsertQueryLog (0) 'Add By Sindy 2010/9/30
    ShowNoData
    CheckOC
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Screen.MousePointer = vbDefault
End Sub

'Add By Cheng 2003/02/25
'智權人員明細
Sub StrPrintDoc1()
Dim strTempName As String '代理人名稱
Dim strSalesGrp As String '業務區
Dim strSalesGrpName As String '業務區名稱

GetPrintLeft
iLine = 1
Page = 1
With adoRecordset
    .MoveFirst
    If txt1(12) = "Y" Or txt1(12) = "y" Then
        TmpArea = "（含多國案）" & "業務區：" & .Fields("A0902").Value
    Else
        TmpArea = "業務區：" & .Fields("A0902").Value
    End If
    'Add By Cheng 2004/05/05
    '記錄智權人員
    St = "" & .Fields(0)
    'End
    StrPrintTital TmpArea, str(Page)
    iPrint = 2700
    iTatle = 0       ' 總筆數
    iKK = 0           ' 合計
    iK = 0           ' 小計
    Do While .EOF = False
        For j = 0 To 11
            If Not IsNull(.Fields(j)) Then
                strTemp3(j) = .Fields(j)
            Else
                strTemp3(j) = ""
            End If
        Next j
        If Not IsNull(adoRecordset.Fields(12)) Then
            strTemp3(4) = adoRecordset.Fields(12)
        Else
            If Not IsNull(adoRecordset.Fields(13)) Then
                strTemp3(4) = adoRecordset.Fields(13)
            Else
                If Not IsNull(adoRecordset.Fields(14)) Then
                    strTemp3(4) = adoRecordset.Fields(14)
                Else
                    If Not IsNull(adoRecordset.Fields(15)) Then
                        strTemp3(4) = adoRecordset.Fields(15)
                    Else
                        If Not IsNull(adoRecordset.Fields(16)) Then
                            strTemp3(4) = adoRecordset.Fields(16)
                        End If
                    End If
                End If
            End If
        End If
        strTemp3(4) = GetPrjPeople1(strTemp3(4))
        If Not IsNull(adoRecordset.Fields(17)) Then
            strTemp3(5) = adoRecordset.Fields(17)
        Else
            If Not IsNull(adoRecordset.Fields(18)) Then
                strTemp3(5) = adoRecordset.Fields(18)
            End If
        End If
      'Modify By Cheng 2002/07/05
      '若系統種類對照檔SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'        strTemp3(5) = GetPrjName1(strTemp3(5))
      If PUB_GetAgentName(SystemNumber(strTemp3(2), 1), strTemp3(5), strTempName) = True Then
           strTemp3(5) = strTempName
      Else
           strTemp3(5) = ""
      End If
        'Add By Cheng 2003/02/26
        '記錄業務區
        strSalesGrp = "" & .Fields("CP12").Value
        strSalesGrpName = "" & .Fields("A0902").Value
        iK = iK + 1
        'Add By Cheng 2003/02/26
        '記錄業務區筆數
        iKK = iKK + 1
        iTatle = iTatle + 1
        If Len(strTemp3(1)) > 7 Then
            strTemp3(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(1)))
        End If
        If Len(strTemp3(6)) > 7 Then
            strTemp3(6) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(6)))
        End If
        If Len(strTemp3(7)) > 7 Then
            strTemp3(7) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(7)))
        End If
        'Printer.Font.Name = "Arial"
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        If iK = 1 Then
            'Modify By Cheng 2003/02/24
            '取得承辦人姓名
'           Printer.Print StrToStr(strTemp3(0), 4)
           Printer.Print StrToStr(GetStaffName(strTemp3(0), True), 4)
        Else
           Printer.Print ""
        End If
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(1)
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(2)
        Printer.CurrentX = PLeft(3)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(3), 4)
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(4), 6)
        Printer.CurrentX = PLeft(5)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(5), 7)
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(6)
        Printer.CurrentX = PLeft(7)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(7)
        Printer.CurrentX = PLeft(8)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(8), 4)
        Printer.CurrentX = PLeft(9)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(9), 4)
        Printer.CurrentX = PLeft(10)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(10), 4)
        Printer.CurrentX = PLeft(11)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(11), 4)
        .MoveNext
        If .EOF = False Then
            If Not IsNull(.Fields(0)) Then
                StrTest1 = .Fields(0)
            Else
                StrTest1 = ""
            End If
        End If
        If .EOF = False Then
            '若智權人員不同時
            If StrTest1 <> St Then
            
               'Add by Morgan 2003/12/17
               If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
                  Printer.NewPage
                  Page = Page + 1
                  StrPrintTital TmpArea, str(Page)
                  iPrint = 2400
                  iLine = 0
               End If
               iLine = iLine + 1
               'Add 2003/12/17
               
                iPrint = iPrint + 300
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                
               'Add by Morgan 2003/12/17
               If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
                  Printer.NewPage
                  Page = Page + 1
                  StrPrintTital TmpArea, str(Page)
                  iPrint = 2400
                  iLine = 0
               End If
               iLine = iLine + 1
               'Add 2003/12/17
               
                iPrint = iPrint + 300
                Printer.CurrentX = 1000
                Printer.CurrentY = iPrint
                Printer.Print "小計： " & Trim(str(iK)) & " 筆"
                iK = 0
                
               'Add by Morgan 2003/12/17
               If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
                  Printer.NewPage
                  Page = Page + 1
                  StrPrintTital TmpArea, str(Page)
                  iPrint = 2400
                  iLine = 0
               End If
               'Add 2003/12/17
               
                iPrint = iPrint + 300
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                
                'Remove by  Morgan 2003/12/17
                'iLine = iLine + 3
                
                St = StrTest1
            End If
            
            'Modify by Morgan 2003/12/17
            'If (iLine Mod 20 = 0) Or iPrint >= 10000 Then
            If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
            'Modify 2003/12/17
            
                iPrint = iPrint + 300
                Printer.NewPage
                Page = Page + 1
                StrPrintTital TmpArea, str(Page)
                iPrint = 2400
                iLine = 0
            End If
            
            'Remove by Morgan 2003/12/17
            'iLine = iLine + 1
            'iPrint = iPrint + 300
            'Remove 2003/12/17
            
            '若業務區不同時
            If strSalesGrp <> "" & .Fields("CP12").Value Then
               'Add by Morgan 2003/12/17
                  iLine = iLine + 1
                  If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
                     Printer.NewPage
                     Page = Page + 1
                     StrPrintTital TmpArea, str(Page)
                     iPrint = 2400
                     iLine = 0
                  End If
                  iPrint = iPrint + 300
               'Add 2003/12/17
                
                Printer.CurrentX = PLeft(0)
                Printer.CurrentY = iPrint
                Printer.Print strSalesGrpName & "合計： " & Trim(str(iKK)) & " 筆"
                iKK = 0
                
               'Add by Morgan 2003/12/17
                  iLine = iLine + 1
                  If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
                     Printer.NewPage
                     Page = Page + 1
                     StrPrintTital TmpArea, str(Page)
                     iPrint = 2400
                     iLine = 0
                  End If
               'Add 2003/12/17
               
                iPrint = iPrint + 300
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
                Printer.NewPage
                Page = Page + 1
                If txt1(12) = "Y" Or txt1(12) = "y" Then
                    TmpArea = "（含多國案）" & "業務區：" & .Fields("A0902").Value
                Else
                    TmpArea = "業務區：" & .Fields("A0902").Value
                End If
                StrPrintTital TmpArea, str(Page)
                iPrint = 2400
                iLine = 0
               'Remove by Morgan 2003/12/17
               'iLine = iLine + 1
               'iPrint = iPrint + 300
               'Remove 2003/12/17
            End If
            'Add by Morgan 2003/12/17
            iLine = iLine + 1
            iPrint = iPrint + 300
            'Add 2003/12/17
        End If
    Loop
End With
'Add By Cheng 2003/02/26
'智權人員小計

'Add by Morgan 2003/12/17
iLine = iLine + 1
If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
   Printer.NewPage
   Page = Page + 1
   StrPrintTital TmpArea, str(Page)
   iPrint = 2400
   iLine = 0
End If
'Add 2003/12/17
               
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")

'Add by Morgan 2003/12/17
iLine = iLine + 1
If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
   Printer.NewPage
   Page = Page + 1
   StrPrintTital TmpArea, str(Page)
   iPrint = 2400
   iLine = 0
End If
'Add 2003/12/17

iPrint = iPrint + 300
Printer.CurrentX = 1000
Printer.CurrentY = iPrint
Printer.Print "小計： " & Trim(str(iK)) & " 筆"
iK = 0

'Add by Morgan 2003/12/17
iLine = iLine + 1
If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
   Printer.NewPage
   Page = Page + 1
   StrPrintTital TmpArea, str(Page)
   iPrint = 2400
   iLine = 0
End If
'Add 2003/12/17

iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")

'Modify by Morgan 2003/12/17
'iLine = iLine + 3
iLine = iLine + 1


St = StrTest1
'Modify by Morgan 2003/12/17
'If (iLine Mod 20 = 0) Or iPrint >= 10000 Then
If (iLine Mod 28 = 0) Or iPrint >= 10000 Then
'Modify 2003/12/17

    iPrint = iPrint + 300
    Printer.NewPage
    Page = Page + 1
    StrPrintTital TmpArea, str(Page)
    iPrint = 2400
    iLine = 0
End If
iLine = iLine + 1
iPrint = iPrint + 300
'業務區合計
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strSalesGrpName & "合計： " & Trim(str(iKK)) & " 筆"
iKK = 0
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")

'iPrint = iPrint + 300
''合計
'iPrint = iPrint + 300
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print String(200, "-")
'iPrint = iPrint + 300
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print "合計：共 " & Trim(str(iTatle)) & " 筆"

Printer.EndDoc
ShowPrintOk
CheckOC

End Sub

'Add By Cheng 2003/01/09
Sub StrPrintDocTotal1()
Dim strTempName As String '代理人名稱

GetPrintLeft
iLine = 1
Page = 1
If txt1(12) = "Y" Or txt1(12) = "y" Then
    TmpArea = "(含多國案)"
Else
    TmpArea = ""
End If
StrPrintTital TmpArea, str(Page)
iPrint = 2700
iTatle = 0       ' 總數
With adoRecordset
    .MoveFirst
    Do While .EOF = False
        For j = 0 To 2
            If Not IsNull(.Fields(j)) Then
                strTemp3(j) = .Fields(j)
            Else
                strTemp3(j) = ""
            End If
        Next j
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(1)
        Printer.CurrentX = PLeft(2) + 750 - TextWidth(Format(strTemp3(2), "##0") & " 筆")
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp3(2), "##0") & " 筆"
        iTatle = iTatle + Val(strTemp3(2))
        .MoveNext
        If (iLine Mod 20 = 0) Or iPrint >= 10000 Then
            iPrint = iPrint + 300
            Printer.NewPage
            Page = Page + 1
            StrPrintTital TmpArea, str(Page)
            iPrint = 2400
            iLine = 0
        End If
        iLine = iLine + 1
        iPrint = iPrint + 300
    Loop
End With
'合計
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(2) + 750 - TextWidth("合計：共 " & Format(iTatle, "##0") & " 筆")
Printer.CurrentY = iPrint
Printer.Print "合計：共 " & Format(iTatle, "##0") & " 筆"
Printer.EndDoc
ShowPrintOk
CheckOC
End Sub
'Add by Morgan 2005/2/14 加PCT進入國家階段條件
Private Sub txtPA46_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
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
