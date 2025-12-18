VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100111_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人收/發文量查詢"
   ClientHeight    =   2280
   ClientLeft      =   1944
   ClientTop       =   2448
   ClientWidth     =   3876
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3876
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2070
      TabIndex        =   7
      Top             =   1890
      Width           =   400
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   972
      TabIndex        =   6
      Top             =   1584
      Width           =   1700
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2328
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3120
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   972
      MaxLength       =   7
      TabIndex        =   2
      Top             =   996
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   972
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1284
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1284
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   3
      Top             =   996
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   972
      MaxLength       =   6
      TabIndex        =   0
      Top             =   420
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   972
      MaxLength       =   1
      TabIndex        =   1
      Top             =   708
      Width           =   192
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   2016
      TabIndex        =   15
      Top             =   456
      Width           =   1812
      Size            =   "3196;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "不計件之案件是否統計：            (N : 不統計)"
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   17
      Top             =   1920
      Width           =   3435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                        (ALL：全部)"
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   1620
      Width           =   3690
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2280
      Y1              =   1152
      Y2              =   1152
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   450
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1035
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   750
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "(1.收文 2.發文)"
      Height          =   180
      Left            =   1440
      TabIndex        =   10
      Top             =   750
      Width           =   1155
   End
End
Attribute VB_Name = "frm100111_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/20 改成Form2.0(lbl1)
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
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
               Me.txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
           If Len(Trim(txt1(0))) <> 0 And Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0 And Len(Trim(txt1(3))) <> 0 Then
           
           Else
              s = MsgBox("承辦人、 查詢別、日期區間不可空白!!", , "USER 輸入錯誤")
              txt1(0).SetFocus
              txt1(0).SelStart = 0
              txt1(0).SelLength = Len(txt1(0))
              Exit Sub
           End If
           txt1_LostFocus (0)
           Me.Enabled = False
          If fnSaveParentForm(Me) = False Then
              Me.Enabled = True
              Exit Sub
          End If
          ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/4 清除查詢印表記錄檔欄位
           Screen.MousePointer = vbHourglass
           frm100111_2.Show
           frm100111_2.StrMenu
           Screen.MousePointer = vbDefault
           Me.Enabled = True
      Case 1
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdGoInput_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(6).Text)) = 0 Then
       Me.txt1(6).Text = "ALL"
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
''      txt1_LostFocus 6
'      'Add By Cheng 2002/03/18
'      If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
'         Me.txt1(2).SetFocus
'         txt1_GotFocus 2
'         Exit Sub
'      End If
'      If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
'         Me.txt1(3).SetFocus
'         txt1_GotFocus 3
'         Exit Sub
'      End If
'
'     If Len(Trim(txt1(0))) <> 0 And Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0 And Len(Trim(txt1(3))) <> 0 Then
'
'     Else
'        s = MsgBox("承辦人、 查詢別、日期區間不可空白!!", , "USER 輸入錯誤")
'        txt1(0).SetFocus
'        txt1(0).SelStart = 0
'        txt1(0).SelLength = Len(txt1(0))
'        Exit Sub
'     End If
'     txt1_LostFocus (0)
'
'     Me.Enabled = False
'     Screen.MousePointer = vbHourglass
'     frm100111_2.Show
'     'frm100111_2.Hide
'
'     frm100111_2.StrMenu
'     Screen.MousePointer = vbDefault
'     Me.Hide
'
'     'frm100111_2.Show
'     Do
'     DoEvents
'     If bolToEndByNick = True Then Unload Me: Exit Sub
'     Loop Until Not frm100111_2.Visible
'     Unload frm100111_2
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
   txt1(6) = Systemkind_g
   '92.04.16 nick
   cmdState = -1
   
   'Added by Lydia 2024/01/15 有「以收/發文量查詢」(frm100105_1)權限的人才可以查詢所有人的資料；其他人僅可查個人承辦的案件量。
   If Pub_StrUserSt03 <> "M51" Then
     If CheckUse("frm100105_1", strExec, False) = False Then
        Me.txt1(0).Text = strUserNum
        Me.lbl1.Caption = strUserName
        Me.txt1(0).Enabled = False
     End If
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100111_1 = Nothing
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
            lbl1.Caption = GetPrjSalesNM(txt1(0))
            If Trim(txt1(Index)) <> "" Then
                 If Trim(lbl1.Caption) = "" Then
                     s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                     txt1(Index).SetFocus
                     txt1_GotFocus (Index)
                     Exit Sub
                 End If
            End If
      Case 1
            If InStr(1, "12", txt1(1)) = 0 Then
               s = MsgBox("查詢別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
               txt1(1).SetFocus
               txt1(1).SelStart = 0
               txt1(1).SelLength = Len(txt1(1))
               Exit Sub
            End If
      Case 2, 3
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 3 Then
             If RunNick(txt1(Index - 1), txt1(Index)) Then
                 txt1(Index - 1).SetFocus
                 txt1_GotFocus (Index - 1)
                 Exit Sub
             End If
         End If
      
      Case 5
             If RunNick(txt1(Index - 1), txt1(Index)) Then
                 txt1(Index - 1).SetFocus
                 txt1_GotFocus (Index - 1)
                 Exit Sub
             End If
      Case 6 '系統類別
            'Modify By Cheng 2002/03/14
      '      'Add By Cheng 2002/01/07
      '      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      Case 7
           If Trim(txt1(Index)) <> "" And UCase(Trim(txt1(Index))) <> "N" Then
               s = MsgBox("只能輸入 N 或 空白！", , "錯誤！")
               txt1(Index).SetFocus
               txt1_GotFocus (Index)
               Exit Sub
           End If
      Case Else
   End Select
End Sub

