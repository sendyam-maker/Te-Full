VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100107_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "收文未發文查詢"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   1370
   ClientWidth     =   4610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4610
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   15
      Top             =   3690
      Width           =   345
   End
   Begin VB.CheckBox Check1 
      Caption         =   "已收款未發文"
      Height          =   255
      Left            =   1128
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   14
      Top             =   3360
      Width           =   732
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1110
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3360
      Width           =   732
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   12
      Top             =   3060
      Width           =   732
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1110
      MaxLength       =   4
      TabIndex        =   11
      Top             =   3060
      Width           =   732
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1128
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1332
      Width           =   972
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3048
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3840
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1128
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1032
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1128
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1932
      Width           =   732
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2232
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1932
      Width           =   732
   End
   Begin VB.OptionButton Option1 
      Caption         =   " 收文日"
      Height          =   180
      Index           =   0
      Left            =   1128
      TabIndex        =   9
      Top             =   2595
      Value           =   -1  'True
      Width           =   1332
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號"
      Height          =   180
      Index           =   1
      Left            =   1128
      TabIndex        =   10
      Top             =   2805
      Width           =   1332
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1128
      MaxLength       =   5
      TabIndex        =   2
      Top             =   732
      Width           =   540
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   1128
      TabIndex        =   5
      Top             =   1632
      Width           =   2100
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   1
      Top             =   456
      Width           =   1212
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1128
      MaxLength       =   7
      TabIndex        =   0
      Top             =   456
      Width           =   1092
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   1
      Left            =   2160
      TabIndex        =   27
      Top             =   1368
      Width           =   2412
      Size            =   "4254;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   2160
      TabIndex        =   26
      Top             =   1080
      Width           =   2412
      Size            =   "4254;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否含已取消收文資料：           (Y:含)"
      Height          =   180
      Index           =   8
      Left            =   60
      TabIndex        =   31
      Top             =   3720
      Width           =   2940
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍："
      Height          =   180
      Index           =   7
      Left            =   60
      TabIndex        =   30
      Top             =   3405
      Width           =   1080
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1980
      X2              =   2100
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   29
      Top             =   3105
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1980
      X2              =   2100
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "紅 色 欄 位 必 須 輸 入 !!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   180
      TabIndex        =   28
      Top             =   4080
      Width           =   3435
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   2400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1992
      X2              =   2112
      Y1              =   2052
      Y2              =   2052
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "（A , B , C）可複選"
      Height          =   180
      Left            =   1725
      TabIndex        =   25
      Top             =   795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   24
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   75
      TabIndex        =   23
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                    (ALL：全部)"
      Height          =   180
      Index           =   3
      Left            =   75
      TabIndex        =   22
      Top             =   1695
      Width           =   4230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   4
      Left            =   75
      TabIndex        =   21
      Top             =   1365
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Index           =   5
      Left            =   75
      TabIndex        =   20
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文種類："
      Height          =   180
      Index           =   6
      Left            =   75
      TabIndex        =   19
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "資料順序："
      Height          =   180
      Left            =   75
      TabIndex        =   18
      Top             =   2595
      Width           =   900
   End
End
Attribute VB_Name = "frm100107_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/20 改成Form2.0(lbl1(0),lbl1(1))
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/10 日期欄已修改
Option Explicit
Dim s As Integer, strSql As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
            cmdState = -1
            If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
               Me.txt1(0).SetFocus
               txt1_GotFocus 0
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If Len(txt1(1)) = 0 Then
               s = MsgBox("收文日期不可空白", , "USER 輸入錯誤")
               txt1(0).SetFocus
               txt1_GotFocus (0)
               Exit Sub
            End If
            Me.Enabled = False
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位
            Screen.MousePointer = vbHourglass
              If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
              End If
            frm100107_2.Show
            frm100107_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
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
''      'Add By Cheng 2002/01/07
''      txt1_LostFocus 5
'      'Add By Cheng 2002/03/18
'      If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
'         Me.txt1(0).SetFocus
'         txt1_GotFocus 0
'         Exit Sub
'      End If
'      If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
'         Me.txt1(1).SetFocus
'         txt1_GotFocus 1
'         Exit Sub
'      End If
'      If Len(txt1(1)) = 0 Then
'         s = MsgBox("收文日期不可空白", , "USER 輸入錯誤")
'         txt1(0).SetFocus
'         txt1_GotFocus (0)
'         Exit Sub
'      End If
'      Me.Enabled = False
'      Screen.MousePointer = vbHourglass
'      frm100107_2.Show
'      'frm100107_2.Hide
'      frm100107_2.StrMenu
'      Screen.MousePointer = vbDefault
'      'frm100107_2.Show
'      Me.Hide
'      Do
'      DoEvents
'      If bolToEndByNick = True Then Unload Me: Exit Sub
'      Loop Until Not frm100107_2.Visible
'      Unload frm100107_2
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
   txt1(5) = Systemkind_g
   bolToEndByNick = False
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100107_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
           If Option1(0).Value = True Then
              Option1(1).Value = False
           End If
      Case 1
           If Option1(1).Value = True Then
              Option1(0).Value = False
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
   'Add By Cheng 2002/12/03
   Select Case Index
      Case 2 '收文種類
          Select Case KeyAscii
          Case 8, 44, 65, 66, 67
          Case Else
              KeyAscii = 0
          End Select
      'Add By Cheng 2003/06/02
      Case 12 '是否含取消收文資料
          If KeyAscii <> 8 And KeyAscii <> 89 Then
              KeyAscii = 0
          End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 0, 1
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 1 Then
              If RunNick(txt1(Index - 1), txt1(Index)) Then
                  txt1(Index - 1).SetFocus
                  txt1_GotFocus (Index - 1)
                  Exit Sub
              End If
         End If
      Case 2
          'Modify By Cheng 2002/12/03
      '     If InStr(1, "A,B,C, ", UCase(txt1(Index))) = 0 Then
      '         s = MsgBox("請輸入 A 或 B 或 C 或加入分隔符號 , !!", , "輸入錯誤")
      '         txt1(Index).SetFocus
      '         txt1(Index).SelStart = 0
      '         txt1(Index).SelLength = Len(txt1(Index))
      '      End If
      Case 3
           If Len(txt1(Index)) <> 0 Then
                 lbl1(0).Caption = GetPrjSalesNM(txt1(Index))
                 If Trim(lbl1(0).Caption) = "" Then
                     s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                     txt1(Index).SetFocus
                     txt1_GotFocus (Index)
                     Exit Sub
                 End If
            Else
              lbl1(0) = ""
            End If
      Case 4
            If Len(txt1(Index)) <> 0 Then
                 lbl1(1).Caption = GetPrjSalesNM(txt1(Index))
                 If Trim(lbl1(1).Caption) = "" Then
                     s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
                     txt1(Index).SetFocus
                     txt1_GotFocus (Index)
                     Exit Sub
                 End If
            Else
                 lbl1(1) = ""
            End If
      Case 5 '系統類別
         'Modify By Cheng 2002/03/14
      '   'Add By Cheng 2002/01/07
      '   Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      Case 7, 9, 11
              If RunNick(txt1(Index - 1), txt1(Index)) Then
                  txt1(Index - 1).SetFocus
                  txt1_GotFocus (Index - 1)
                  Exit Sub
              End If
      Case Else
   End Select
End Sub

