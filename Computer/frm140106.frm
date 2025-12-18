VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140106 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所內商延展、第二期註冊費銷卷作業"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3930
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   3
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1500
      Width           =   285
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1770
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   2775
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   0
      Left            =   1110
      MaxLength       =   8
      TabIndex        =   0
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   1
      Left            =   2490
      MaxLength       =   8
      TabIndex        =   1
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   2
      Left            =   1110
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "(1:延展/續展 2:第二期註冊費)"
      Height          =   180
      Left            =   1470
      TabIndex        =   13
      Top             =   1590
      Width           =   2280
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1620
      Width           =   900
   End
   Begin MSForms.Label lblst06 
      Height          =   255
      Left            =   1110
      TabIndex        =   11
      Top             =   480
      Width           =   315
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(2:中;3:南;4:高)"
      Height          =   180
      Left            =   1530
      TabIndex        =   10
      Top             =   510
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   90
      TabIndex        =   9
      Top             =   510
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Left            =   90
      TabIndex        =   8
      Top             =   840
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2850
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   90
      TabIndex        =   7
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lblST 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   1425
   End
End
Attribute VB_Name = "frm140106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/03 Form2.0 已修改 lblST
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Trim(txt1(0)) = "" Then
        MsgBox "法定期限區間起不可空白！", vbExclamation
        txt1(0).SetFocus
        Exit Sub
     End If
     If Trim(txt1(1)) = "" Then
        MsgBox "法定期限區間迄不可空白！", vbExclamation
        txt1(1).SetFocus
        Exit Sub
     End If
     Me.Hide
     frm140106_1.Show
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   lblst06 = pub_strUserOffice
   txt1(0) = ChangeWDateStringToTString(DateAdd("m", -7, ChangeWStringToWDateString(Mid(strSrvDate(1), 1, 6) & "01")))
   txt1(1) = ChangeWDateStringToTString(DateAdd("d", -1, DateAdd("m", -6, ChangeWStringToWDateString(Mid(strSrvDate(1), 1, 6) & "01"))))
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm140106 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         If KeyAscii < 48 And KeyAscii > 57 And KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      'Add By Sindy 2010/11/26
      Case 2
         KeyAscii = UpperCase(KeyAscii)
      Case 3
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim s
Cancel = True
Select Case Index
    Case 0, 1
        If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
           Me.txt1(Index).SetFocus
           txt1_GotFocus Index
           Exit Sub
        End If
        If Index = 1 Then
           If RunNick2(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
           End If
         End If
    Case 2
        lblST.Caption = GetPrjSalesNM(txt1(Index))
        If Trim(txt1(Index)) <> "" Then
             If Trim(lblST.Caption) = "" Then
                 s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
                 txt1(Index).SetFocus
                 txt1_GotFocus (Index)
                 Exit Sub
             End If
        End If
    Case Else
End Select
Cancel = False
End Sub
