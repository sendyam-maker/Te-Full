VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010301_12 
   BorderStyle     =   1  '單線固定
   Caption         =   "Pdf 轉檔失敗!!"
   ClientHeight    =   4335
   ClientLeft      =   435
   ClientTop       =   2610
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5640
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   525
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   3780
      Width           =   5430
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   3645
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   5415
      VariousPropertyBits=   -1400879073
      ScrollBars      =   2
      Size            =   "9551;6429"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm06010301_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/14 Form2.0已檢查 (無需修改的物件)
Option Explicit

Private Sub Command1_Click(Index As Integer)
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010301_12 = Nothing
End Sub

