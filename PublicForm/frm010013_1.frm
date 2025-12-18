VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010013_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所收文量查詢"
   ClientHeight    =   2070
   ClientLeft      =   285
   ClientTop       =   1755
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6645
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1290
      MaxLength       =   5
      TabIndex        =   0
      Top             =   90
      Width           =   1332
   End
   Begin VB.OptionButton Option1 
      Caption         =   "收文日期"
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   18
      Top             =   390
      Width           =   1125
   End
   Begin VB.OptionButton Option1 
      Caption         =   "收文月份"
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   120
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5070
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton CmdOk 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5856
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1650
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1650
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   1
      Top             =   360
      Width           =   1332
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   2
      Left            =   1140
      TabIndex        =   16
      Top             =   720
      Width           =   735
      ForeColor       =   0
      VariousPropertyBits=   27
      Size            =   "1296;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   2160
      X2              =   2280
      Y1              =   1800
      Y2              =   1800
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   735
      ForeColor       =   0
      VariousPropertyBits=   27
      Size            =   "1296;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   1800
      TabIndex        =   14
      Top             =   1020
      Width           =   2175
      VariousPropertyBits=   27
      Size            =   "3836;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "輸入民國年"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2730
      TabIndex        =   11
      Top             =   390
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "所　別："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "智權人員："
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   900
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   2160
      X2              =   2280
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frm010013_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 lbl1()
'Memo By Sonia 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

'2004/1/27
Public Sub PubShowNextData()
    'edit by nick 2004/10/08
    If Option1(0).Value = True Then
        If txt1(0) = "" Then
            MsgBox "收文月份不可空白 !", vbCritical
            txt1(1).SetFocus
            Exit Sub
        End If
    Else
        If txt1(1) = "" Then
            MsgBox "收文日不可空白 !", vbCritical
            txt1(1).SetFocus
            Exit Sub
        End If
    End If
    If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
        txt1_GotFocus 1
        Exit Sub
    End If
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    frm010013_2.Show
    frm010013_2.StrMenu
    Screen.MousePointer = vbDefault
    Me.Enabled = True
End Sub

'2004/1/27
Public Sub UnloadChild()
    Dim oForm As Form
    For Each oForm In Forms
        If oForm.Name = "frm010013_2" Then
            Unload frm010013_2
            Exit For
        End If
    Next
End Sub

'2004/1/27
Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
        Case 0
            PubShowNextData
        Case 1
            Unload Me
    End Select
End Sub

'2004/1/27
Private Sub Form_Load()
    MoveFormToCenter Me
    txt1(1) = strSrvDate(2)
    '依使用者設定所別
    CheckOC
    lbl1(0).Caption = "" & PUB_GetST06(strUserNum)
    Select Case lbl1(0)
        Case "1"
            lbl1(2).Caption = "北"
        Case "2"
            lbl1(2).Caption = "中"
        Case "3"
            lbl1(2).Caption = "南"
        Case "4"
            lbl1(2).Caption = "高"
        Case Else '5
            lbl1(2).Caption = "其他"
    End Select
End Sub

'2004/1/27
Private Sub Form_Unload(Cancel As Integer)
    Set frm010013_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(Index).Value = True Then
    If Index = 0 Then
        txt1(0).Enabled = True
        txt1(1).Enabled = False
        txt1(0).SetFocus
    Else
        txt1(0).Enabled = False
        txt1(1).Enabled = True
        txt1(1).SetFocus
    End If
End If
End Sub

'2004/1/27
Private Sub txt1_GotFocus(Index As Integer)
    txt1(Index).SelStart = 0
    txt1(Index).SelLength = Len(txt1(Index))
    'edit by nickc 2007/06/06 切換輸入法改用API
    'txt1(Index).IMEMode = 2
    CloseIme
End Sub

'2004/1/27
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'2004/1/28
Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        'add by nick 2004/10/08
        Case 0
            If txt1(Index).Text <> "" Then
                If DateCheck(DBDATE(txt1(Index).Text & "01")) = "N" Then
                    MsgBox "收文月份錯誤！", , "錯誤！"
                    Call txt1_GotFocus(Index)
                    Cancel = True
                End If
            End If
            
        Case 1  '收文日期
            If PUB_CheckKeyInDate(txt1(Index)) = -1 Then
                Call txt1_GotFocus(Index)
                Cancel = True
            End If
            
        Case 8  '智權人員
            If Len(txt1(Index)) <> 0 Then
               CheckOC
               lbl1(1).Caption = "" & GetStaffName(txt1(Index), True)
               If lbl1(1).Caption = "" Then
                    Call MsgBox("智權人員輸入錯誤！", , "錯誤！")
                    Call txt1_GotFocus(Index)
                    Cancel = True
               End If
             Else
                lbl1(1).Caption = ""
             End If
             
        Case 9, 13 '申請國家,案件性質
            If Trim(txt1(Index)) <> "" And Trim(txt1(Index + 1)) <> "" Then
                If RunNick(txt1(Index), txt1(Index + 1)) Then
                    Call txt1_GotFocus(Index)
                    Cancel = True
                End If
            End If
            
        Case 10, 14 '申請國家,案件性質
            If Trim(txt1(Index)) <> "" And Trim(txt1(Index - 1)) <> "" Then
                If RunNick(txt1(Index - 1), txt1(Index)) Then
                    Call txt1_GotFocus(Index)
                    Cancel = True
                End If
            End If
    End Select
End Sub
