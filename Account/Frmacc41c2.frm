VERSION 5.00
Begin VB.Form Frmacc41c2 
   AutoRedraw      =   -1  'True
   Caption         =   "Account"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2880
   Begin VB.ComboBox CboComp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Text            =   "CboComp"
      Top             =   480
      Width           =   1600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "欲操作之公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "Frmacc41c2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 (無需修改)
Option Explicit

Dim intBT As Integer 'Add by Amy 2018/01/04

'Add by Amy 2020/03/17
Private Sub CboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii = 13 Then Exit Sub
End Sub

Private Sub CboComp_Validate(Cancel As Boolean)
    Dim strBKCmp As String
    
    If cboComp = MsgText(601) Then Exit Sub
    
    strBKCmp = GetBookKeepCmp
    If InStr(strBKCmp, cboComp) = 0 Then
        MsgBox Label1 & "作帳公司別輸入有誤,請確認！", , MsgText(5)
        Cancel = True
        Exit Sub
    End If

End Sub
'end 2020/03/17

'Modify by Amy 2020/03/17 改用下拉試選單 原:Text1
Private Sub cmdok_Click(Index As Integer)
    intBT = Index
    If Index = 0 Then
        If CheckComp = False Then
           Exit Sub
        End If
        strCompanyNo = Me.cboComp
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    intBT = 2 '可能按x離開
    Call Pub_SetCboCmpNo(cboComp, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add by Amy 2018/01/04
    If intBT <> 0 Then
        Call PUB_GetLock("", "Frmacc4120")
    End If
    
   Set Frmacc41c2 = Nothing
End Sub

Private Function CheckComp() As Boolean
    Dim bCancel As Boolean
    
    bCancel = False
    If cboComp = MsgText(601) Then
        MsgBox Label1 & "不可為空白", , MsgText(5)
        CheckComp = False
        cboComp.SetFocus
        Exit Function
    End If
    'Call Text1_Validate(bCancel)
    Call CboComp_Validate(bCancel)
    If bCancel = True Then
        CheckComp = False
        cboComp.SetFocus
        Exit Function
    End If
    CheckComp = True
End Function

'Private Sub Text1_GotFocus()
'    TextInverse Text1
'    CloseIme
'End Sub
'
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    KeyAscii = UpperCase(KeyAscii)
'End Sub

'Private Sub Text1_Validate(Cancel As Boolean)
'
'    If Text1 = MsgText(601) Then
'        Exit Sub
'    End If
'    If Text1 <> "1" And Text1 <> "J" Then
'        MsgBox Label1 & "只能輸入1 或 J ！", , MsgText(5)
'        Cancel = True
'        Exit Sub
'     End If
'End Sub
'end 2020/03/17
