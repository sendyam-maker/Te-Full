VERSION 5.00
Begin VB.Form frm030411 
   BorderStyle     =   1  '單線固定
   Caption         =   "延遲承辦案件明細查詢"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4005
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   5
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1770
      Width           =   300
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   1515
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1470
      Width           =   240
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1170
      Width           =   660
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   1965
      MaxLength       =   7
      TabIndex        =   2
      Top             =   855
      Width           =   705
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   1
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   1
      Top             =   855
      Width           =   705
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   510
      Width           =   1740
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   30
      Width           =   825
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   2220
      TabIndex        =   6
      Top             =   30
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "(1. 查詢  2.印表)"
      Height          =   180
      Index           =   5
      Left            =   1860
      TabIndex        =   15
      Top             =   1815
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   14
      Top             =   1815
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:是)"
      Height          =   180
      Index           =   11
      Left            =   1830
      TabIndex        =   13
      Top             =   1530
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "是否列印明細："
      Height          =   180
      Index           =   8
      Left            =   210
      TabIndex        =   12
      Top             =   1530
      Width           =   1260
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   1830
      TabIndex        =   11
      Top             =   1215
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   10
      Top             =   1230
      Width           =   990
   End
   Begin VB.Line Line1 
      X1              =   1245
      X2              =   2505
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   900
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   540
      Width           =   915
   End
End
Attribute VB_Name = "frm030411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
'create by nickc 2008/01/10 陳經理有請作單
Option Explicit
Dim i As Integer, s As Integer, j As Integer, strTemp1 As Variant, strTemp2 As Variant
'add by nickc 2008/04/03 陳經理加控制
Dim stST05 As String

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     If Len(TXT1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         TXT1(0).SetFocus
         Exit Sub
     Else
         If Len(TXT1(5)) = 0 Then
             s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
             TXT1(5).SetFocus
             Exit Sub
         Else
            If PUB_CheckKeyInDate(Me.TXT1(1)) = -1 Then
               Me.TXT1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.TXT1(2)) = -1 Then
               Me.TXT1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If Len(Trim(TXT1(2))) = 0 Then
                 s = MsgBox("發文日期區間不可空白!!", , "USER 輸入錯誤")
                 TXT1(2).SetFocus
                 txt1_GotFocus (2)
                 Exit Sub
             Else
                'add by nickc 2008/04/03 加入外商主管  可以輸入相同組別的
                If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") Then
                    If Trim(TXT1(3)) = "" Then
                        MsgBox "承辦人不可以空白！", vbExclamation, "操作錯誤！"
                        TXT1(3).SetFocus
                        txt1_GotFocus 3
                        Exit Sub
                    End If
                    If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(TXT1(3)) Then
                        MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
                        TXT1(3).SetFocus
                        txt1_GotFocus 3
                        Exit Sub
                    End If
                End If
                Screen.MousePointer = vbHourglass
                DoEvents
                Me.Enabled = False
                frm030411_1.Hide
                DoEvents
                ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
                If TXT1(5) = "1" Then
                     pub_QL05 = pub_QL05 & ";" & Label1(1) & "查詢" 'Add By Sindy 2010/10/22
                Else
                     pub_QL05 = pub_QL05 & ";" & Label1(1) & "印表" 'Add By Sindy 2010/10/22
                End If
                frm030411_1.Process
                If TXT1(5) = "1" Then
                    Me.Hide
                    frm030411_1.Show
                Else
                    frm030411_1.PrintData
                    Unload frm030411_1
                End If
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

Private Sub Form_Load()
'add by nickc 2008/04/03 陳經理加控制
stST05 = PUB_GetST05(strUserNum)
Select Case stST05
Case "11", "00", "01"
Case "21", "26", "28"
Case Else
    TXT1(3) = strUserNum
    TXT1(3).Enabled = False
End Select

MoveFormToCenter Me
TXT1(0) = GetSystemKindByNick

End Sub

Private Sub Form_Unload(cancel As Integer)
Set frm030411 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
TXT1(Index).SelStart = 0
TXT1(Index).SelLength = Len(TXT1(Index))
CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, cancel As Boolean)
Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(TXT1(0)), ",,", ""), ",")
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
            TXT1(0).SetFocus
            TXT1(0).SelStart = 0
            TXT1(0).SelLength = Len(TXT1(0))
            cancel = True
            Exit Sub
        End If
     Next i
Case 2, 1
   If PUB_CheckKeyInDate(Me.TXT1(Index)) = -1 Then
      Me.TXT1(Index).SetFocus
      txt1_GotFocus Index
      cancel = True
      Exit Sub
   End If
   If Index = 2 Then
     If RunNick(TXT1(Index - 1), TXT1(Index)) Then
         TXT1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
Case 3
     lbl1(1) = GetPrjSalesNM(TXT1(Index))
     If Trim(TXT1(Index)) <> "" Then
        If Trim(lbl1(1).Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            TXT1(Index).SetFocus
            txt1_GotFocus (Index)
            cancel = True
            Exit Sub
        End If
     End If
Case 4
     Select Case Trim(TXT1(Index))
     Case "Y", ""
     Case Else
        s = MsgBox("是否列印明細只能輸入Y 或空白！", , "錯誤！")
        TXT1(Index).SetFocus
        txt1_GotFocus (Index)
        cancel = True
        Exit Sub
     End Select
Case 5
     Select Case Trim(TXT1(Index))
     Case "1", "2", ""
     Case Else
        s = MsgBox("列印別只能輸入1或2！", , "錯誤！")
        TXT1(Index).SetFocus
        txt1_GotFocus (Index)
        cancel = True
        Exit Sub
     End Select
End Select
End Sub


