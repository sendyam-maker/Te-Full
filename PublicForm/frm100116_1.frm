VERSION 5.00
Begin VB.Form frm100116_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "以國別查詢"
   ClientHeight    =   3330
   ClientLeft      =   7100
   ClientTop       =   4730
   ClientWidth     =   4150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4150
   Begin VB.CheckBox ChkPCT 
      Caption         =   "是否顯示PCT 案"
      Height          =   225
      Left            =   300
      TabIndex        =   10
      Top             =   2610
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   690
      Left            =   90
      TabIndex        =   18
      Top             =   528
      Width           =   1515
      Begin VB.OptionButton Option1 
         Caption         =   "申請人國籍："
         CausesValidation=   0   'False
         Height          =   204
         Index           =   0
         Left            =   72
         TabIndex        =   20
         Top             =   48
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "申請國家："
         CausesValidation=   0   'False
         Height          =   228
         Index           =   1
         Left            =   72
         TabIndex        =   19
         Top             =   390
         Width           =   1212
      End
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   2850
      MaxLength       =   4
      TabIndex        =   3
      Top             =   870
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1650
      MaxLength       =   4
      TabIndex        =   2
      Top             =   870
      Width           =   852
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3372
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2580
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1530
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1650
      MaxLength       =   4
      TabIndex        =   0
      Top             =   528
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   2850
      MaxLength       =   4
      TabIndex        =   1
      Top             =   528
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1215
      Width           =   285
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1530
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2190
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2190
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   1860
      Width           =   1572
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "注意!!   有輸入 ""往來日期"" 區間時會很久  "
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   21
      Top             =   2970
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Line Line4 
      X1              =   2640
      X2              =   2760
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line1 
      X1              =   2625
      X2              =   2745
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   2640
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   2400
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      Height          =   180
      Left            =   240
      TabIndex        =   17
      Top             =   1275
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(1.收文 2.發文)"
      Height          =   180
      Left            =   1740
      TabIndex        =   16
      Top             =   1275
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                        (ALL：全部)"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   1890
      Width           =   3690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   2220
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   1575
      Width           =   540
   End
End
Attribute VB_Name = "frm100116_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/29 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         cmdState = -1
         If Option1(0).Value = True Then
            If Len(Trim(txt1(1))) = 0 Then
               s = MsgBox("申請人國籍不可空白", , "USER 輸入錯誤")
               txt1(1).SetFocus
               Exit Sub
            End If
         Else
            If Len(Trim(txt1(3))) = 0 Then
               s = MsgBox("申請國家不可空白", , "USER 輸入錯誤")
               txt1(3).SetFocus
               Exit Sub
             End If
         End If
         If Len(Trim(txt1(4))) = 0 Then
            s = MsgBox("查詢別不可空白", , "USER 輸入錯誤")
            txt1(4).SetFocus
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(5)) = -1 Then
             Me.txt1(5).SetFocus
             txt1_GotFocus 5
             Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(6)) = -1 Then
            Me.txt1(6).SetFocus
            txt1_GotFocus 6
            Exit Sub
         End If
         Me.Enabled = False
         If fnSaveParentForm(Me) = False Then
             Me.Enabled = True
             Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/15 清除查詢印表記錄檔欄位
         frm100116_2.Show
         frm100116_2.StrMenu
         Screen.MousePointer = vbDefault
         Me.Enabled = True
      Case 1
         fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdGoInput_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(7).Text)) = 0 Then
       Me.txt1(7).Text = "ALL"
   End If
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolToEndByNick = False
   bolToEndByNick = False
   txt1(7) = Systemkind_g
   '92.04.16 nick
   cmdState = -1

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100116_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
           If Option1(0).Value = True Then
               txt1(0).SetFocus
               txt1_GotFocus (0)
           End If
      Case 1
           If Option1(1).Value = True Then
               txt1(2).SetFocus
               txt1_GotFocus (2)
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
   Select Case Index
      Case 1
           If RunNick(txt1(0), txt1(1)) Then
               txt1(0).SetFocus
               txt1_GotFocus (0)
            End If
      Case 3
           If RunNick(txt1(2), txt1(3)) Then
               txt1(2).SetFocus
               txt1_GotFocus (2)
            End If
      Case 4
            If InStr(1, "12", txt1(Index)) = 0 Then
              s = MsgBox("查詢別只可輸入 1 或 2 !!", , "USER  輸入錯誤")
              txt1(Index).SetFocus
              txt1(Index).SelStart = 0
              txt1(Index).SelLength = Len(txt1(Index))
              Exit Sub
            End If
      Case 5, 6 '日期起, 迄
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 6 Then
           If RunNick(txt1(5), txt1(6)) Then
               txt1(5).SetFocus
               txt1_GotFocus (5)
            End If
          End If
      Case 7 '系統類別
            'Modify By Cheng 2002/03/14
      '      'Add By Cheng 2002/01/07
      '      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      Case 9
           If RunNick(txt1(8), txt1(9)) Then
               txt1(8).SetFocus
               txt1_GotFocus (8)
            End If
      Case Else
   End Select
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
      Case 0, 1
          Option1(0).Value = True
      Case 2, 3
          Option1(1).Value = True
      Case Else
   End Select
End Sub
