VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090616_0 
   BorderStyle     =   1  '單線固定
   Caption         =   "月考核"
   ClientHeight    =   1680
   ClientLeft      =   2355
   ClientTop       =   1725
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3645
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   3
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "1"
      Top             =   1260
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   2
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   3
      Top             =   915
      Width           =   1020
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束"
      Height          =   375
      Index           =   1
      Left            =   2715
      TabIndex        =   6
      Top             =   30
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1755
      TabIndex        =   5
      Top             =   30
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   1
      Left            =   2385
      MaxLength       =   5
      TabIndex        =   2
      Top             =   510
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1155
      MaxLength       =   5
      TabIndex        =   1
      Top             =   525
      Width           =   1020
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   2220
      TabIndex        =   10
      Top             =   930
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1：查詢；2：報表)"
      Height          =   180
      Left            =   1530
      TabIndex        =   9
      Top             =   1305
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "查詢方式："
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   150
      TabIndex        =   7
      Top             =   960
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1650
      X2              =   2835
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "考核年月："
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   585
      Width           =   900
   End
End
Attribute VB_Name = "frm090616_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; lbl1 ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
Dim i As Integer
Select Case Index
Case 0
         For i = 0 To 4
            Cancel = False
            txt1_Validate i, Cancel
            If Cancel = True Then
                Exit Sub
            End If
         Next i
         If Len(Trim(txt1(0))) = 0 Then
            MsgBox "考核年月不可空白！", , " 錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If Len(Trim(txt1(1))) = 0 Then
            MsgBox "考核年月不可空白！", , " 錯誤！"
            txt1(1).SetFocus
            Exit Sub
         End If
         If Len(Trim(txt1(3))) = 0 Then
            MsgBox "查詢方式不可空白！", , " 錯誤！"
            txt1(3).SetFocus
            Exit Sub
         End If
         If txt1(3) = "1" Then
            Me.Hide
         End If
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         'Added by Morgan 2019/3/21 +108考核(工程師)
         strExc(1) = DBDATE(Trim(txt1(0)) & "01")
         If ProSysState = "1" And strExc(1) >= PUB_108RuleDate Then
            frm090616_2.Show
         Else
         'end 2019/3/21
            frm090616_1.Show
         End If 'Added by Morgan 2019/3/21
         Me.Enabled = True
         Screen.MousePointer = vbDefault
Case 1
         Unload Me
Case Else
End Select
End Sub

Private Sub Form_Activate()
Static bolActivated As Boolean

ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
If ProState <> "2" Then
   If Not bolActivated Then 'Added by Morgan 2019/3/25
      txt1(2) = strUserNum
      txt1_Validate 2, False
      txt1(2).Enabled = False
   End If
   txt1(3).Enabled = False
   
   SetUserNumEnabled 'Added by Morgan 2019/3/25
End If

bolActivated = True
End Sub

Private Sub Form_Load()
m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
MoveFormToCenter Me
'edit by nickc 2005/04/18 游經理說預設系統日前一個月
'edit by nickc 2005/08/30 7號以前抓前一個月
If Val(Mid(ServerDate, 7, 2)) <= 7 Then
'txt1(0).Text = Val(Mid(ServerDate, 1, 6)) - 191100
'txt1(1).Text = Val(Mid(ServerDate, 1, 6)) - 191100
   txt1(0).Text = Val(Mid(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Mid(ServerDate, 1, 6) & "01"))), 1, 6)) - 191100
   txt1(1).Text = Val(Mid(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Mid(ServerDate, 1, 6) & "01"))), 1, 6)) - 191100
Else
   txt1(0).Text = Val(Mid(ServerDate, 1, 6)) - 191100
   txt1(1).Text = Val(Mid(ServerDate, 1, 6)) - 191100
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090616_0 = Nothing
End Sub


Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 0
      If Trim(txt1(Index)) <> "" Then
         If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
      End If
      
      SetUserNumEnabled 'Added by Morgan 2019/3/25
      
Case 1
      If Trim(txt1(Index)) <> "" Then
         If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
      End If
      If RunNick(txt1(0).Text, txt1(1).Text) Then
         txt1(1).SetFocus
         Cancel = True
         Exit Sub
      End If
Case 2
      lbl1.Caption = "" 'Added by Morgan 2019/3/25
      If Trim(txt1(Index)) <> "" Then
          strSql = "select * from staff where st01='" & txt1(Index) & "' " & IIf(ProSysState = "1", " and st03>='P10' and st03<='P11' ", " and st03='P13' ")
          CheckOC3
          AdoRecordSet3.CursorLocation = adUseClient
          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If AdoRecordSet3.RecordCount <> 0 Then
               lbl1.Caption = AdoRecordSet3.Fields("st02").Value
          Else
              MsgBox "請輸入有效" & IIf(ProSysState = "1", " 承辦人 ", " 繪圖人員 ") & "員工編號！", , "錯誤！"
              Cancel = True
              Exit Sub
          End If
      End If
Case 3
        If Trim(txt1(Index)) <> "" Then
            Select Case txt1(Index)
            Case "1", "2"
            Case Else
                    MsgBox "請輸入 1 或 2 選擇查詢方式！", , "錯誤！"
                    txt1(Index).SetFocus
                    Cancel = True
                    Exit Sub
            End Select
        End If
Case Else
End Select
End Sub

'Added by Morgan 2019/3/25 108考核,開放工程師成員(P11)可以查詢所有人員最近一季的季考核成績(在王副總完成評分之後)
Private Sub SetUserNumEnabled()
   If ProState = "2" Then
      txt1(2).Enabled = True
   Else
      txt1(2).Enabled = False
      
      If Pub_StrUserSt03 = "P11" And ProSysState = "1" And strSrvDate(1) >= PUB_108RuleDate And Val(txt1(0)) > 0 Then
         If Trim(Val(txt1(0)) + 191100) >= Left(CompDate(1, -1, strSrvDate(1)), 6) Then
            txt1(2).Enabled = True
         End If
      End If
      
      If txt1(2).Enabled = False Then
         txt1(2) = strUserNum
         txt1_Validate 2, False
         txt1(2).Enabled = False
      End If
   End If
End Sub
