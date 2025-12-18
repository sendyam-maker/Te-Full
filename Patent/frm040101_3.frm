VERSION 5.00
Begin VB.Form frm040101_3 
   AutoRedraw      =   -1  'True
   Caption         =   "接洽單案件性質數量"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3675
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   6
      Top             =   720
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      Caption         =   "其他"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   2205
      TabIndex        =   5
      Top             =   720
      Width           =   1635
   End
   Begin VB.OptionButton Option1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1770
      TabIndex        =   4
      Top             =   720
      Width           =   500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1170
      TabIndex        =   3
      Top             =   720
      Width           =   500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   50
      TabIndex        =   1
      Top             =   720
      Width           =   500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   500
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
      Left            =   960
      TabIndex        =   7
      Top             =   1080
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
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   780
      TabIndex        =   10
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "接洽單案件性質數量"
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
      Left            =   50
      TabIndex        =   9
      Top             =   360
      Width           =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請選擇"
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
      Left            =   50
      TabIndex        =   8
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frm040101_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/3 改成Form2.0(無)
'Create by Amy 2014/11/18
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
    Dim bolChoose As Boolean
    Dim ii As Integer
    
    strPublicTemp = MsgText(601)
    If Index = 0 Then
        For ii = 0 To Option1.UBound
            If Option1(ii).Value = True Then
                If ii = 4 Then
                    If Trim(Text1) = MsgText(601) Then
                        MsgBox "選擇其他請輸入數量！", , MsgText(5)
                        Exit Sub
                    Else
                        strPublicTemp = Text1
                    End If
                Else
                    strPublicTemp = Option1(ii).Caption
                End If
                bolChoose = True
                Exit For
            End If
        Next ii
        If bolChoose = False Then
            MsgBox "請選擇接洽單案件性質數量！", vbCritical
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
   'MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040101_3 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 4 Then
        Text1.SetFocus
    End If
End Sub

Private Sub Text1_GotFocus()
    CloseIme
    Option1(4).Value = True
    TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Trim(Text1) = "" Then Exit Sub
    If IsNumeric(Text1) = False Then
        MsgBox "請輸入數字", , MsgText(5)
        Cancel = True
        Exit Sub
    End If
End Sub
