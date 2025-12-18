VERSION 5.00
Begin VB.Form frm210133_2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "本所案號"
   ClientHeight    =   972
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   3396
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   972
   ScaleWidth      =   3396
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      Top             =   480
      Width           =   525
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   720
      MaxLength       =   6
      TabIndex        =   1
      Top             =   480
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   2
      Top             =   480
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1956
      MaxLength       =   2
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   430
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   50
      Width           =   800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   456
      X2              =   2016
      Y1              =   612
      Y2              =   612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "欲操作之本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frm210133_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2025/07/21
Option Explicit

Dim intType As Integer, strClose As String
Dim stA1K04 As String, stA1K27 As String, stA1K28 As String, stA1K29 As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         intType = Empty: strClose = Empty
         If ChkForm() = False Then
            Exit Sub
         End If
         If setNextForm = False Then
            Exit Sub
         End If
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210133_2 = Nothing
End Sub

Private Function ChkForm() As Boolean
   Dim stMsg As String
   
   ChkForm = False
   
   If txt1(0) = "" Or txt1(1) = "" Then
      MsgBox "案號不可為空", vbCritical, "操作錯誤！"
      If txt1(0) = "" Then
         Call txt1_GotFocus(0)
      Else
         Call txt1_GotFocus(1)
      End If
      Exit Function
   End If
   If txt1(0) <> "FCT" And txt1(0) <> "T" And txt1(0) <> "S" And txt1(0) <> "CFT" And txt1(0) <> "CFC" Then
      MsgBox "系統別輸入錯誤", vbCritical, "操作錯誤！"
      Call txt1_GotFocus(0)
      Exit Function
   End If
    
   If Trim(txt1(2)) = "" Then txt1(2) = "0"
   If Trim(txt1(3)) = "" Then txt1(3) = "00"
   
   intType = Pub_ChkFCAg(Me.Name, txt1(0), txt1(1), txt1(2), txt1(3), strClose, stMsg)
   If intType = 9 Or strClose = "Y" Then
      Call txt1_GotFocus(0)
      MsgBox stMsg, vbCritical, "輸入錯誤！"
      Exit Function
   '不進入 國外 or 國內 結案單
   ElseIf intType = 98 Or intType = 99 Then
      Call txt1_GotFocus(0)
      Exit Function
   End If
   
   ChkForm = True
End Function

Private Function setNextForm() As Boolean
   Dim nxtFrm As Form, strMsg As String
   
'*** 此函數有修改需確認 Trademark1.mdiMain的 ChkAndShowClose 函數 是否也需要改  ***
   setNextForm = False
   '[有]FC代理人
   If intType = 1 Then
      Set nxtFrm = Forms(0).GetForm("frm210133_F")
   '[無FC代理人
   Else
      Set nxtFrm = Forms(0).GetForm("frm210133")
   End If
   Call nxtFrm.SetParent(Me)
   nxtFrm.txt1(0) = Me.txt1(0)
   nxtFrm.txt1(1) = Me.txt1(1)
   nxtFrm.txt1(2) = Me.txt1(2)
   nxtFrm.txt1(3) = Me.txt1(3)
   '國外結案單
   If UCase(nxtFrm.Name) = "FRM210133_F" Then
      If nxtFrm.doQuery(strMsg) = False Then
         If strMsg <> "" Then
            MsgBox strMsg, vbExclamation, "警告！"
         End If
      Else
         setNextForm = True
      End If
   '國內結案單
   Else
      '外商承辦之案件都會有FC代理人,故國內結案單程式不修改,[無]FC代理人再看到時情況該如何處理-秀玲
      '秀玲:印紙本 讓程序補上 FC代理人,且信件不沖銷代表異常,讓User操作變麻煩才有警覺
      setNextForm = True
      Call nxtFrm.cmdok_Click(2)
   End If
   
   If setNextForm = True Then
      nxtFrm.txt1(0).Locked = True
      nxtFrm.txt1(1).Locked = True
      nxtFrm.txt1(2).Locked = True
      nxtFrm.txt1(3).Locked = True
      nxtFrm.cmdOK(2).Enabled = False
      Call FormClear '避免下次使用資料殘留,先清
      Me.Hide
   Else
      txt1(1).SetFocus
   End If
End Function

Private Sub FormClear()
   Dim obj As Object
   
   For Each obj In txt1
      obj.Text = ""
   Next
   intType = Empty
   strClose = Empty
   stA1K04 = Empty
   stA1K27 = Empty
   stA1K28 = Empty
   stA1K29 = Empty
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
