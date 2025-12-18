VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmpic002 
   BackColor       =   &H80000018&
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4620
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1800
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      Height          =   465
      Left            =   270
      ScaleHeight     =   405
      ScaleWidth      =   420
      TabIndex        =   4
      Top             =   2250
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   1020
      Left            =   240
      ScaleHeight     =   960
      ScaleWidth      =   975
      TabIndex        =   2
      Top             =   990
      Width           =   1035
      Begin MSForms.Label lblUniText 
         Height          =   960
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   975
         VariousPropertyBits=   8388627
         Caption         =   "字"
         Size            =   "1720;1693"
         FontName        =   "新細明體"
         FontHeight      =   960
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   585
      Left            =   2700
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
      VariousPropertyBits=   746637339
      Size            =   "2355;1032"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   5
      Top             =   2460
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "圖片讀取中...請稍候..."
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   26.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6270
   End
End
Attribute VB_Name = "frmpic002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Public m_bFixIME As Boolean 'Added by Morgan 2022/7/15

Private Sub Form_Activate()
   'Added by Morgan 2022/7/15
   '修正輸入法失效問題(觸發後卸載)
   If m_bFixIME Then
      TextBox1.SetFocus
      'SendKeys "^{ }"
      'SendKeys "^{ }"
      Text1.SetFocus
'WIN10 SendKeys 無效
'      If Val(PUB_GetVersionNo) >= 6.2 Then
'         keybd_event vbKeyShift, 0, 0, 0
'         keybd_event vbKeyShift, 0, &H2, 0
'      End If
      Unload Me
   End If
   'end 2022/7/15
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmpic002 = Nothing
End Sub
