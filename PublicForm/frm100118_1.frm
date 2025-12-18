VERSION 5.00
Begin VB.Form frm100118_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "監視系統案件查詢"
   ClientHeight    =   2505
   ClientLeft      =   810
   ClientTop       =   4395
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4965
   Begin VB.OptionButton Option1 
      Caption         =   "CCC CODE："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   1635
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "BTTM："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   1290
      Width           =   1185
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所發文號："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   1005
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   705
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   3510
      MaxLength       =   2
      TabIndex        =   4
      Top             =   645
      Width           =   360
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   3150
      MaxLength       =   1
      TabIndex        =   3
      Top             =   645
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2190
      MaxLength       =   6
      TabIndex        =   2
      Top             =   645
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1605
      MaxLength       =   3
      TabIndex        =   1
      Top             =   645
      Width           =   495
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4170
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3375
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   30
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1590
      MaxLength       =   15
      TabIndex        =   6
      Top             =   960
      Width           =   1680
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1590
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1290
      Width           =   1305
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Index           =   6
      Left            =   1590
      MaxLength       =   209
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   10
      Top             =   1605
      Width           =   3045
   End
End
Attribute VB_Name = "frm100118_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/07 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer ', StrTemp As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
           cmdState = -1
            If Option1(0).Value = True Then
                If Len(Trim(txt1(1))) = 0 Then
                   s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
                   txt1(1).SetFocus
                   Exit Sub
                End If
            Else
                If Option1(1).Value = True Then
                   If Len(Trim(txt1(4))) = 0 Then
                      s = MsgBox("本所發文號不可空白", , "USER 輸入錯誤")
                      txt1(4).SetFocus
                      Exit Sub
                   End If
                Else
                   If Option1(2).Value = True Then
                      If Len(Trim(txt1(5))) = 0 Then
                          s = MsgBox("BTTM 不可空白", , "USER 輸入錯誤")
                          txt1(5).SetFocus
                          Exit Sub
                      End If
                   Else
                      If Option1(3).Value = True Then
                          If Len(Trim(txt1(6))) = 0 Then
                              s = MsgBox("CCC CODE 不可空白", , "USER 輸入錯誤")
                              txt1(6).SetFocus
                              Exit Sub
                          End If
                      End If
                   End If
               End If
           End If
           Me.Enabled = False
          If fnSaveParentForm(Me) = False Then
              Me.Enabled = True
              Exit Sub
          End If
           Screen.MousePointer = vbHourglass
           ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
           frm100118_2.Show
           frm100118_2.StrMenu
           Screen.MousePointer = vbDefault
           Me.Enabled = True
      Case 1
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdGoInput_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
''92.04.16 nick 以下無效
'Select Case Index
'Case 0
'      If Option1(0).Value = True Then
'          If Len(Trim(txt1(1))) = 0 Then
'             s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
'             txt1(1).SetFocus
'             Exit Sub
'          End If
'      Else
'          If Option1(1).Value = True Then
'             If Len(Trim(txt1(4))) = 0 Then
'                s = MsgBox("本所發文號不可空白", , "USER 輸入錯誤")
'                txt1(4).SetFocus
'                Exit Sub
'             End If
'          Else
'             If Option1(2).Value = True Then
'                If Len(Trim(txt1(5))) = 0 Then
'                    s = MsgBox("BTTM 不可空白", , "USER 輸入錯誤")
'                    txt1(5).SetFocus
'                    Exit Sub
'                End If
'             Else
'                If Option1(3).Value = True Then
'                    If Len(Trim(txt1(6))) = 0 Then
'                        s = MsgBox("CCC CODE 不可空白", , "USER 輸入錯誤")
'                        txt1(6).SetFocus
'                        Exit Sub
'                    End If
'                End If
'             End If
'         End If
'     End If
'     Me.Enabled = False
'     Screen.MousePointer = vbHourglass
'     frm100118_2.Show
'     'frm100118_2.Hide
'     frm100118_2.StrMenu
'     Screen.MousePointer = vbDefault
'     Me.Hide
'     'frm100118_2.Show
'     Do
'     DoEvents
'     If bolToEndByNick = True Then Unload Me: Exit Sub
'     Loop Until Not frm100118_2.Visible
'     Unload frm100118_2
'     Me.Enabled = True
'     Me.Show
'Case 1
'      Unload Me
'Case Else
'End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolToEndByNick = False
   Option1(0).Value = True
   Option1(1).Value = False
   Option1(2).Value = False
   Option1(3).Value = False
   'txt1(4).Enabled = False
   'txt1(5).Enabled = False
   'txt1(6).Enabled = False
   txt1(0) = "TM"
   txt1(0).Enabled = False
   'StrTemp = Systemkind_g
   'If bolFNation = False Then
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100118_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
            If Option1(0).Value = True Then
               txt1(1).SetFocus
               txt1_GotFocus (1)
            End If
      Case 1
            If Option1(1).Value = True Then
               txt1(4).SetFocus
               txt1_GotFocus (4)
            End If
      Case 2
            If Option1(2).Value = True Then
               txt1(5).SetFocus
               txt1_GotFocus (5)
            End If
      Case 3
            If Option1(3).Value = True Then
               txt1(6).SetFocus
               txt1_GotFocus (6)
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

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
      Case 0, 1, 2, 3
          Option1(0).Value = True
      Case 4
          Option1(1).Value = True
      Case 5
          Option1(2).Value = True
      Case 6
          Option1(3).Value = True
      Case Else
   End Select
End Sub
