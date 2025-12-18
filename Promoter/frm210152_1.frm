VERSION 5.00
Begin VB.Form frm210152_1 
   AutoRedraw      =   -1  'True
   Caption         =   "選擇登入部門"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3270
   Begin VB.ComboBox CboDept 
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
      Width           =   2115
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
      Left            =   2370
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
      Left            =   2370
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "欲操作部門別"
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
Attribute VB_Name = "frm210152_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/03 Form2.0已檢查 (無需修改的物件)
'Create by Amy 2021/10/07
Option Explicit

Dim RsQ As New ADODB.Recordset
Dim strQ As String, intQ As Integer, i As Integer

Private Sub CboDept_KeyPress(KeyAscii As Integer)
    If Pub_StrUserSt15 <> "M51" Then KeyAscii = 0: Exit Sub
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii = 13 Then Exit Sub
End Sub

Private Sub CboDept_Validate(Cancel As Boolean)
    If Trim(CboDept) = MsgText(601) Then Exit Sub
End Sub

Private Sub cmdOK_Click(Index As Integer)
    If Index = 0 Then
        If CboDept = MsgText(601) Then
            Exit Sub
        End If
        strPublicTemp = Left(Me.CboDept, 3)
    Else
        strPublicTemp = ""
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    CboDept.Clear
    Call ShowCboDept
    strPublicTemp = ""
End Sub

Private Sub ShowCboDept()
    strQ = "Select a0901,a0902 From Acc090 Where a0901 in(" & strPublicTemp & ")"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        CboDept.AddItem ""
        Do While RsQ.EOF = False
            CboDept.AddItem "" & RsQ.Fields("a0901") & " " & RsQ.Fields("a0902")
            RsQ.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm210152_1 = Nothing
End Sub
