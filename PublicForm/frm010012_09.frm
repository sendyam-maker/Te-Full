VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010012_09 
   BorderStyle     =   1  '單線固定
   Caption         =   "轉案至他所結果輸入"
   ClientHeight    =   4650
   ClientLeft      =   5580
   ClientTop       =   1860
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8520
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   7650
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame fraWindow1 
      BorderStyle     =   0  '沒有框線
      Height          =   3972
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   8532
      Begin VB.TextBox textCP13 
         Height          =   315
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1710
         Width           =   945
      End
      Begin MSForms.TextBox TextNote 
         Height          =   1305
         Left            =   1020
         TabIndex        =   1
         Top             =   2550
         Width           =   7365
         VariousPropertyBits=   -1466939365
         ScrollBars      =   2
         Size            =   "12991;2302"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM05 
         Height          =   1155
         Left            =   1020
         TabIndex        =   5
         Top             =   120
         Width           =   7365
         VariousPropertyBits=   -1466939365
         ScrollBars      =   2
         Size            =   "12991;2037"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LabDept 
         Caption         =   "LabDept"
         Height          =   255
         Left            =   5640
         TabIndex        =   18
         Top             =   1770
         Width           =   2895
      End
      Begin VB.Label LabNation 
         Caption         =   "LabNation"
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   1380
         Width           =   2895
      End
      Begin MSForms.Label LabSubject 
         Height          =   255
         Left            =   1020
         TabIndex        =   16
         Top             =   2160
         Width           =   7155
         VariousPropertyBits=   27
         Caption         =   "LabSubject"
         Size            =   "12621;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabSales 
         Height          =   255
         Left            =   1980
         TabIndex        =   15
         Top             =   1770
         Width           =   2415
         VariousPropertyBits=   27
         Caption         =   "LabSales"
         Size            =   "4260;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabApp 
         Height          =   255
         Left            =   1020
         TabIndex        =   14
         Top             =   1380
         Width           =   3375
         VariousPropertyBits=   27
         Caption         =   "LabApp"
         Size            =   "5953;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         Caption         =   "申請國家："
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label LabelPA52 
         Alignment       =   1  '靠右對齊
         Caption         =   "申 請 人："
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label LabelPA51 
         Alignment       =   1  '靠右對齊
         Caption         =   "案件名稱："
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   150
         Width           =   945
      End
      Begin VB.Label LabelPA56 
         Alignment       =   1  '靠右對齊
         Caption         =   "受文者部門："
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   1770
         Width           =   1155
      End
      Begin VB.Label LabelPA55 
         Alignment       =   1  '靠右對齊
         Caption         =   "備      註："
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label LabelPA54 
         Alignment       =   1  '靠右對齊
         Caption         =   "主      旨："
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label LabelPA53 
         Alignment       =   1  '靠右對齊
         Caption         =   "受 文 者："
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   1770
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   5664
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6510
      TabIndex        =   3
      Top             =   70
      Width           =   1100
   End
End
Attribute VB_Name = "frm010012_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/14 改成Form2.0 (textLC05,grdList...)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create By Sindy 2012/2/21
Option Explicit

'回傳值
Public BolOk As Boolean     'True: 確定  False: 取消


Private Sub cmdOK_Click(Index As Integer)
   If Index = 0 Then '確定
      BolOk = True
      Me.Hide
      Call frm010001.CP24_2_T728progress
   ElseIf Index = 1 Then '回前畫面
      BolOk = False
      Me.Hide
      Call frm010001.CP24_2_T728progress
   ElseIf Index = 2 Then '結束
      Unload frm010001
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010012_09 = Nothing
End Sub

Private Sub TextNote_GotFocus()
   InverseTextBox TextNote
   OpenIme
End Sub

Private Sub textCP13_Change()
   LabSales.Caption = ""
   LabDept = ""
End Sub

Private Sub textCP13_GotFocus()
   textCP13.SelStart = 0
   textCP13.SelLength = Len(textCP13.Text)
   CloseIme
End Sub

Private Sub textCP13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP13_Validate(Cancel As Boolean)
   If textCP13 <> "" Then
      If textCP13.Text <> "" And textCP13.Text < "63001" Then
         MsgBox "智權人員編號不可小於 63001！", , "注意！"
         Cancel = True
         textCP13_GotFocus
         Exit Sub
      End If
      
      Dim strTemp As String, strTemp1 As String, strST15 As String
      If Not ClsPDGetStaff(textCP13.Text, strTemp, strTemp1) Then
         Cancel = True
         textCP13_GotFocus
         Exit Sub
      End If
      strST15 = GetST15(textCP13.Text, strTemp1)
      LabSales.Caption = strTemp
      LabDept = strST15 & " " & strTemp1
   End If
End Sub
