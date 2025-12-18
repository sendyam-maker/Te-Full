VERSION 5.00
Begin VB.Form frm10011201_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "後金案件及結果查詢"
   ClientHeight    =   2220
   ClientLeft      =   3210
   ClientTop       =   2160
   ClientWidth     =   5205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5205
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1230
      TabIndex        =   5
      Top             =   1680
      Width           =   2800
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4005
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3210
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   0
      Top             =   690
      Width           =   192
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1020
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   2430
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1020
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1350
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   2430
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1350
      Width           =   852
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                                 (ALL：全部)"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   1740
      Width           =   4815
   End
   Begin VB.Line Line2 
      X1              =   2190
      X2              =   2310
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   2190
      X2              =   2310
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      Height          =   180
      Left            =   390
      TabIndex        =   11
      Top             =   750
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(1.收文 2.無結果 3.有結果)"
      Height          =   180
      Left            =   1590
      TabIndex        =   10
      Top             =   750
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   180
      Left            =   405
      TabIndex        =   9
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "結果日："
      Height          =   180
      Index           =   0
      Left            =   390
      TabIndex        =   8
      Top             =   1410
      Width           =   720
   End
End
Attribute VB_Name = "frm10011201_1"
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
Dim s As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
           cmdState = -1
           If txt1(0) = "1" Or txt1(0) = "2" Then
               If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
               If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
                  Me.txt1(2).SetFocus
                  txt1_GotFocus 2
                  Exit Sub
               End If
                        
               If Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0 Then
               Else
                  s = MsgBox("收文日區間不可空白", , "USER 輸入錯誤")
                  If Len(Trim(txt1(2))) = 0 Then txt1(2).SetFocus
                  If Len(Trim(txt1(1))) = 0 Then txt1(1).SetFocus
                  
                  Exit Sub
               End If
           Else
              If txt1(0) = "3" Then
                  If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
                     Me.txt1(3).SetFocus
                     txt1_GotFocus 3
                     Exit Sub
                  End If
                  If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
                     Me.txt1(4).SetFocus
                     txt1_GotFocus 4
                     Exit Sub
                  End If
                  If Len(Trim(txt1(3))) <> 0 And Len(Trim(txt1(4))) <> 0 Then
                  Else
                      s = MsgBox("結果日區間不可空白", , "USER 輸入錯誤")
                      If Len(Trim(txt1(4))) = 0 Then txt1(4).SetFocus
                      If Len(Trim(txt1(3))) = 0 Then txt1(3).SetFocus
                      
                      Exit Sub
                  End If
              Else
                  s = MsgBox("查詢別必須輸入", , "USER 輸入錯誤")
                  txt1(0).SetFocus
                  Exit Sub
              End If
           End If
           Me.Enabled = False
          If fnSaveParentForm(Me) = False Then
              Me.Enabled = True
              Exit Sub
          End If
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/4 清除查詢印表記錄檔欄位
         Screen.MousePointer = vbHourglass
         frm10011201_2.Show
         frm10011201_2.StrMenu
         Screen.MousePointer = vbDefault
         Me.Enabled = True
      Case 1
         fnCloseAllFrm100
      Case Else
   End Select
End Sub
 
Private Sub cmdGoInput_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(5).Text)) = 0 Then
       Me.txt1(5).Text = "ALL"
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
''      txt1_LostFocus 5
'
'     If txt1(0) = "1" Or txt1(0) = "2" Then
'         'Add By Cheng 2002/03/18
'         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
'            Me.txt1(1).SetFocus
'            txt1_GotFocus 1
'            Exit Sub
'         End If
'         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
'            Me.txt1(2).SetFocus
'            txt1_GotFocus 2
'            Exit Sub
'         End If
'
'         If Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0 Then
'         Else
'            s = MsgBox("收文日區間不可空白", , "USER 輸入錯誤")
'            If Len(Trim(txt1(2))) = 0 Then txt1(2).SetFocus
'            If Len(Trim(txt1(1))) = 0 Then txt1(1).SetFocus
'
'            Exit Sub
'         End If
'     Else
'        If txt1(0) = "3" Then
'            'Add By Cheng 2002/03/18
'            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
'               Me.txt1(3).SetFocus
'               txt1_GotFocus 3
'               Exit Sub
'            End If
'            If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
'               Me.txt1(4).SetFocus
'               txt1_GotFocus 4
'               Exit Sub
'            End If
'
'            If Len(Trim(txt1(3))) <> 0 And Len(Trim(txt1(4))) <> 0 Then
'            Else
'                s = MsgBox("結果日區間不可空白", , "USER 輸入錯誤")
'                If Len(Trim(txt1(4))) = 0 Then txt1(4).SetFocus
'                If Len(Trim(txt1(3))) = 0 Then txt1(3).SetFocus
'
'                Exit Sub
'            End If
'        Else
'            s = MsgBox("查詢別必須輸入", , "USER 輸入錯誤")
'            txt1(0).SetFocus
'            Exit Sub
'        End If
'     End If
'     Me.Enabled = False
'     Screen.MousePointer = vbHourglass
'     frm10011201_2.Show
'     'frm10011201_2.Hide
'
'     frm10011201_2.StrMenu
'     Screen.MousePointer = vbDefault
'     Me.Hide
'     'frm10011201_2.Show
'     Do
'     DoEvents
'     If bolToEndByNick = True Then Unload Me: Exit Sub
'     Loop Until Not frm10011201_2.Visible
'     Unload frm10011201_2
'     Me.Enabled = True
'     Me.Show
'Case 1
'     Unload Me
'Case Else
'End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   bolToEndByNick = False
   txt1(5) = Systemkind_g
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm10011201_1 = Nothing
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
   Select Case Index
      Case 0
            If (InStr(1, "123 ", txt1(0)) = 0) Then
               s = MsgBox("查詢別只能輸入 1 或 2 或 3 !!", , "USER 輸入錯誤")
               txt1(0).SetFocus
               txt1(0).SelStart = 0
               txt1(0).SelLength = Len(txt1(0))
               Exit Sub
            End If
      Case 1, 2, 3, 4
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 2 Or Index = 4 Then
             If RunNick(txt1(Index - 1), txt1(Index)) Then
                  txt1(Index - 1).SetFocus
                  txt1_GotFocus (Index - 1)
                  Exit Sub
             End If
         End If
      Case 5 '系統類別
            'Modify By Cheng 2002/03/14
      '      'Add By Cheng 2002/01/07
      '      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      Case Else
   End Select
End Sub

