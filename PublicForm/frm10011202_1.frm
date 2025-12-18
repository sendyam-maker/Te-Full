VERSION 5.00
Begin VB.Form frm10011202_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "後金收回查詢"
   ClientHeight    =   2400
   ClientLeft      =   510
   ClientTop       =   2490
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5340
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1452
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1080
      Width           =   1572
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請人編號："
      Height          =   180
      Index           =   1
      Left            =   84
      TabIndex        =   5
      Top             =   1140
      Width           =   1380
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   84
      TabIndex        =   0
      Top             =   750
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1425
      TabIndex        =   10
      Top             =   1800
      Width           =   2300
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1452
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1440
      Width           =   1572
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   3372
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1440
      Width           =   1572
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   3924
      MaxLength       =   2
      TabIndex        =   4
      Top             =   720
      Width           =   345
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   3456
      MaxLength       =   1
      TabIndex        =   3
      Top             =   720
      Width           =   360
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   2124
      MaxLength       =   6
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3432
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4224
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   3372
      MaxLength       =   9
      TabIndex        =   7
      Top             =   1080
      Width           =   1572
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1452
      MaxLength       =   3
      TabIndex        =   1
      Top             =   720
      Width           =   570
   End
   Begin VB.Line Line2 
      X1              =   3150
      X2              =   3270
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3150
      X2              =   3270
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                      (ALL： 全部)"
      Height          =   180
      Index           =   2
      Left            =   465
      TabIndex        =   14
      Top             =   1800
      Width           =   4365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收回日期："
      Height          =   180
      Index           =   1
      Left            =   495
      TabIndex        =   13
      Top             =   1500
      Width           =   900
   End
End
Attribute VB_Name = "frm10011202_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/07 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer, strSql As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
           cmdState = -1
            If PUB_CheckKeyInDate(Me.txt1(6)) = -1 Then
               Me.txt1(6).SetFocus
               txt1_GotFocus 6
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(7)) = -1 Then
               Me.txt1(7).SetFocus
               txt1_GotFocus 7
               Exit Sub
            End If
            If Option1(0).Value = True Then
               If Len(Trim(txt1(0))) = 0 Or Len(Trim(txt1(1))) = 0 Or Len(Trim(txt1(6))) = 0 Or Len(Trim(txt1(7))) = 0 Then
                  s = MsgBox("本所案號與收回日期不可空白", , "USER 輸入資料錯誤")
                  If Len(Trim(txt1(7))) = 0 Then txt1(7).SetFocus
                  If Len(Trim(txt1(6))) = 0 Then txt1(6).SetFocus
                  If Len(Trim(txt1(3))) = 0 Then txt1(3).SetFocus
                  If Len(Trim(txt1(2))) = 0 Then txt1(2).SetFocus
                  If Len(Trim(txt1(1))) = 0 Then txt1(1).SetFocus
                  If Len(Trim(txt1(0))) = 0 Then txt1(0).SetFocus
                  Exit Sub
               End If
            Else
               If Option1(1).Value = True Then
                  If Len(Trim(txt1(4))) = 0 Or Len(Trim(txt1(5))) = 0 Or Len(Trim(txt1(6))) = 0 Or Len(Trim(txt1(7))) = 0 Then
                      s = MsgBox("申請人編號與收回日期不可空白", , "USRE 輸入資料錯誤")
                      If Len(Trim(txt1(7))) = 0 Then txt1(7).SetFocus
                      If Len(Trim(txt1(6))) = 0 Then txt1(6).SetFocus
                      If Len(Trim(txt1(5))) = 0 Then txt1(5).SetFocus
                      If Len(Trim(txt1(4))) = 0 Then txt1(4).SetFocus
                      Exit Sub
                  End If
               End If
            End If
            If Left(txt1(4), 6) <> Left(txt1(5), 6) Then
               s = MsgBox("申請人編號前 6 碼必須相同", , "USER 輸入錯誤")
               txt1(5).SetFocus
               txt1(5).SelStart = 0
               txt1(5).SelLength = Len(txt1(5))
               Exit Sub
            End If
            Me.Enabled = False
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/4 清除查詢印表記錄檔欄位
            Screen.MousePointer = vbHourglass
            frm10011202_2.Show
            frm10011202_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
      Case 1
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdGoInput_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(8).Text)) = 0 Then
       Me.txt1(8).Text = "ALL"
   End If
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
''92.04.16 nick 以下無效
'Select Case Index
'Case 0
'      'Modify By Cheng 2002/03/14
''      'Add By Cheng 2002/01/07
''      txt1_LostFocus 8
'      'Add By Cheng 2002/03/18
'      If PUB_CheckKeyInDate(Me.txt1(6)) = -1 Then
'         Me.txt1(6).SetFocus
'         txt1_GotFocus 6
'         Exit Sub
'      End If
'      If PUB_CheckKeyInDate(Me.txt1(7)) = -1 Then
'         Me.txt1(7).SetFocus
'         txt1_GotFocus 7
'         Exit Sub
'      End If
'
'      If Option1(0).Value = True Then
'         If Len(Trim(txt1(0))) = 0 Or Len(Trim(txt1(1))) = 0 Or Len(Trim(txt1(6))) = 0 Or Len(Trim(txt1(7))) = 0 Then
'            s = MsgBox("本所案號與收回日期不可空白", , "USER 輸入資料錯誤")
'            If Len(Trim(txt1(7))) = 0 Then txt1(7).SetFocus
'            If Len(Trim(txt1(6))) = 0 Then txt1(6).SetFocus
'            If Len(Trim(txt1(3))) = 0 Then txt1(3).SetFocus
'            If Len(Trim(txt1(2))) = 0 Then txt1(2).SetFocus
'            If Len(Trim(txt1(1))) = 0 Then txt1(1).SetFocus
'            If Len(Trim(txt1(0))) = 0 Then txt1(0).SetFocus
'            Exit Sub
'         End If
'      Else
'         If Option1(1).Value = True Then
'            If Len(Trim(txt1(4))) = 0 Or Len(Trim(txt1(5))) = 0 Or Len(Trim(txt1(6))) = 0 Or Len(Trim(txt1(7))) = 0 Then
'                s = MsgBox("申請人編號與收回日期不可空白", , "USRE 輸入資料錯誤")
'                If Len(Trim(txt1(7))) = 0 Then txt1(7).SetFocus
'                If Len(Trim(txt1(6))) = 0 Then txt1(6).SetFocus
'                If Len(Trim(txt1(5))) = 0 Then txt1(5).SetFocus
'                If Len(Trim(txt1(4))) = 0 Then txt1(4).SetFocus
'                Exit Sub
'            End If
'         End If
'      End If
'      If Left(txt1(4), 6) <> Left(txt1(5), 6) Then
'         s = MsgBox("申請人編號前 6 碼必須相同", , "USER 輸入錯誤")
'         txt1(5).SetFocus
'         txt1(5).SelStart = 0
'         txt1(5).SelLength = Len(txt1(5))
'         Exit Sub
'      End If
'      'If Len(Trim(txt1(4))) < 9 And Option1(1).Value = True Then
'      '   For i = 1 To 9 - Len(Trim(txt1(4)))
'      '      txt1(4) = txt1(4) + "0"
'      '   Next i
'      'End If
'      'If Len(Trim(txt1(5))) < 9 And Option1(1).Value = True Then
'      '   For i = 1 To 9 - Len(Trim(txt1(5)))
'      '       txt1(5) = txt1(5) + "0"
'      '   Next i
'      'End If
'
'      Me.Enabled = False
'      Screen.MousePointer = vbHourglass
'      frm10011202_2.Show
'     ' frm10011202_2.Hide
'
'      frm10011202_2.StrMenu
'      Screen.MousePointer = vbDefault
'      Me.Hide
'      'frm10011202_2.Show
'      Do
'      DoEvents
'      If bolToEndByNick = True Then Unload Me: Exit Sub
'      Loop Until Not frm10011202_2.Visible
'      Unload frm10011202_2
'      Me.Enabled = True
'      Me.Show
'Case 1
'      Unload Me
'Case Else
'End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   bolToEndByNick = False
   Option1(1).Value = False
   txt1(8) = Systemkind_g
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm10011202_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
            If Option1(0).Value = True Then
              txt1(0).SetFocus
              txt1_GotFocus (0)
              Option1(1).Value = False
           End If
      Case 1
           If Option1(1).Value = True Then
              Option1(0).Value = False
              txt1(4).SetFocus
              txt1_GotFocus (4)
           End If
      Case Else
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   'Add By Cheng 2002/01/07
   Select Case Index
      Case 6, 7
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 7 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
         End If
      Case 5
            If Mid(txt1(Index - 1), 1, 6) <> Mid(txt1(Index), 1, 6) Then
                s = MsgBox("前6碼必須相同！", , "錯誤！")
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
            If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
      Case 8 '系統類別
         'Modify By Cheng 2002/03/14
      '   Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
   End Select
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
      Case 0, 1, 2, 3
          Option1(0).Value = True
      Case 4, 5
          Option1(1).Value = True
      Case Else
   End Select
      
End Sub


