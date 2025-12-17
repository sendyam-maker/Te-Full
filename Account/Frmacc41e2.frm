VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41e2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "Email通知"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdReceiver 
      Caption         =   "副本..."
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
      Left            =   135
      TabIndex        =   3
      Top             =   1410
      Width           =   1095
   End
   Begin VB.CommandButton cmdReceiver 
      Caption         =   "收件者..."
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
      Left            =   135
      TabIndex        =   2
      Top             =   270
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   3420
      TabIndex        =   1
      Top             =   4170
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "傳送"
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
      Left            =   2295
      TabIndex        =   0
      Top             =   4170
      Width           =   1095
   End
   Begin MSForms.TextBox txtContent 
      Height          =   1845
      Left            =   180
      TabIndex        =   8
      Top             =   2220
      Width           =   4335
      VariousPropertyBits=   -1467989989
      BackColor       =   16777215
      ScrollBars      =   2
      Size            =   "7646;3254"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtReceiver 
      Height          =   1035
      Left            =   1260
      TabIndex        =   7
      Top             =   270
      Width           =   3255
      VariousPropertyBits=   -1467989985
      BackColor       =   14737632
      Size            =   "5741;1826"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtSubject 
      Height          =   345
      Left            =   1260
      TabIndex        =   5
      Top             =   1800
      Width           =   3255
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCopy 
      Height          =   345
      Left            =   1260
      TabIndex        =   4
      Top             =   1380
      Width           =   3255
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "主旨："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   6
      Top             =   1875
      Width           =   615
   End
End
Attribute VB_Name = "Frmacc41e2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 txtReceiver/txtCopy/txtSubject/txtContent
'Created by Morgan 2013/4/26
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdReceiver_Click(Index As Integer)
   Dim arrList1, arrList2
   Dim ii As Integer
   
   If Index = 0 Then
      Frmacc41e3.Caption = "收件者"
      arrList1 = Split(txtReceiver, ";")
      arrList2 = Split(txtReceiver.Tag, ";")
   Else
      Frmacc41e3.Caption = "副本"
      arrList1 = Split(txtCopy, ";")
      arrList2 = Split(txtCopy.Tag, ";")
   End If
   
   For ii = 0 To UBound(arrList1)
      If arrList1(ii) <> "" Then
         Frmacc41e3.List2.AddItem arrList1(ii) & vbTab & arrList2(ii)
      End If
   Next
   Frmacc41e3.Show vbModal
   
   If intI = 1 Then
      If Index = 0 Then
         txtReceiver = strExc(1)
         txtReceiver.Tag = strExc(2)
      Else
         txtCopy = strExc(1)
         txtCopy.Tag = strExc(2)
      End If
   End If
   
   strFormName = Me.Name
End Sub

Private Sub cmdSend_Click()
   'Added by Morgan 2022/9/16
   If txtReceiver = "" Then
      MsgBox "收件者不可空白！", vbCritical
      Exit Sub
   End If
   'end 2022/9/16
   
   PUB_SendMail strUserNum, txtReceiver.Tag, "", txtSubject, txtContent, , , , , , txtCopy.Tag, , , , True
   If bolMailSendOk = True Then
      strSql = "update acc230 set a2324=sysdate where a2301='" & Frmacc41e0.txtA2301 & "'"
      cnnConnection.Execute strSql, intI
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc41e2 = Nothing
End Sub

