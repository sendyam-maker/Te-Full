VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100117_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文日查詢代理人作業進度"
   ClientHeight    =   2270
   ClientLeft      =   500
   ClientTop       =   3090
   ClientWidth     =   6220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2270
   ScaleWidth      =   6220
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1770
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1275
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1770
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1275
      TabIndex        =   5
      Top             =   1440
      Width           =   3804
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4560
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5340
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   90
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1275
      MaxLength       =   7
      TabIndex        =   0
      Top             =   435
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2640
      MaxLength       =   7
      TabIndex        =   1
      Top             =   435
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1275
      MaxLength       =   4
      TabIndex        =   2
      Top             =   780
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   3
      Top             =   780
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1275
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1116
      Width           =   1080
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   2430
      TabIndex        =   13
      Top             =   1170
      Width           =   3765
      Size            =   "6641;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   1800
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   2400
      X2              =   2520
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                                                 (ALL：全部)"
      Height          =   180
      Index           =   1
      Left            =   230
      TabIndex        =   14
      Top             =   1500
      Width           =   5930
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2520
      Y1              =   915
      Y2              =   915
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2520
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "發文日期："
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   495
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   180
      TabIndex        =   11
      Top             =   825
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1155
      Width           =   720
   End
End
Attribute VB_Name = "frm100117_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/20 改成Form2.0(lbl1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim s As Integer, i As Integer, j As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim m_bolFinalCheck As Boolean '最後檢查控制

'92.04.16 nick
Public Sub PubShowNextData()
Dim oText As TextBox

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
         
         If Len(Trim(txt1(1))) = 0 Then
            s = MsgBox("發文日區間不可空白", , "USER 輸入錯誤")
            If Len(Trim(txt1(0))) = 0 Then txt1(0).SetFocus
            Exit Sub
         End If
         
         'Add By Sindy 2012/2/8
         m_bolFinalCheck = True
         For Each oText In txt1
            txt1_LostFocus oText.Index
            If m_bolFinalCheck = False Then
               Exit Sub
            End If
         Next
         
         Me.Enabled = False
         If fnSaveParentForm(Me) = False Then
            Me.Enabled = True
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/15 清除查詢印表記錄檔欄位
         frm100117_2.Show
         frm100117_2.StrMenu
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
'      'Modify By Cheng 2002/03/15
''     'Add By Cheng 2002/01/07
''     txt1_LostFocus 5
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
'
'     If Len(Trim(txt1(1))) = 0 Then
'        s = MsgBox("發文日區間不可空白", , "USER 輸入錯誤")
'        If Len(Trim(txt1(0))) = 0 Then txt1(0).SetFocus
'        Exit Sub
'     End If
'     Me.Enabled = False
'     Screen.MousePointer = vbHourglass
'     frm100117_2.Show
'     'frm100117_2.Hide
'     frm100117_2.StrMenu
'     Screen.MousePointer = vbDefault
'     Me.Hide
'     'frm100117_2.Show
'     Do
'     DoEvents
'     If bolToEndByNick = True Then Unload Me: Exit Sub
'     Loop Until Not frm100117_2.Visible
'     Unload frm100117_2
'     Me.Enabled = True
'     Me.Show
'Case 1
'     Unload Me
'Case Else
'End Select
End Sub

Private Sub Form_Activate()
   If bolFNation = False Then
       s = MsgBox("國內人員不可查詢代理人案件", , "違規.....")
       Unload Me
       Exit Sub
   End If
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
   Set frm100117_1 = Nothing
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
'Add By Cheng 2002/07/09
Dim strTmp As String
Dim strTemp1

   Select Case Index
      Case 0, 1
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 1 Then
           If RunNick(txt1(0), txt1(1)) Then
            txt1(0).SetFocus
            txt1_GotFocus (0)
           End If
         End If
      Case 3
           If RunNick(txt1(2), txt1(3)) Then
            txt1(2).SetFocus
            txt1_GotFocus (2)
           End If
      Case 4
      '     If Len(Trim(txt1(4))) > 0 Then
      '        lbl1 = Left(GetPrjName2(txt1(4)), 20)
      '     Else
      '        lbl1 = ""
      '     End If
            'Modify By Cheng 2002/07/08
            '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
            If Me.txt1(5).Text = "" Then
               strTemp1 = Split(" ", ",")
            Else
               strTemp1 = Split(IIf(Me.txt1(5).Text = "ALL", Systemkind_g, Me.txt1(5).Text), ",")
            End If
            If PUB_GetAgentName(IIf(Len(Trim(txt1(5))) <> 0, strTemp1(0), ""), Me.txt1(4).Text, strTmp) Then
               lbl1 = strTmp
            Else
               lbl1 = ""
               If Trim(txt1(Index)) <> "" Then
                  s = MsgBox("代理人錯誤！", , "錯誤！")
                  txt1(Index).SetFocus
                  txt1_GotFocus (Index)
                  Exit Sub
               End If
            End If
      Case 7 '案件性質
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               m_bolFinalCheck = False
               Exit Sub
            End If
      Case Else
   End Select
End Sub
