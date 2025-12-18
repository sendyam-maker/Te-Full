VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100137_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "frm100137_1"
   ClientHeight    =   3708
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5748
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   5748
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   372
      Left            =   4750
      TabIndex        =   0
      Top             =   30
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢條件："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   216
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   996
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "同時查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   216
      Left            =   120
      TabIndex        =   3
      Top             =   396
      Width           =   912
   End
   Begin MSForms.TextBox txtChg 
      Height          =   3000
      Left            =   120
      TabIndex        =   2
      Top             =   624
      Width           =   5496
      VariousPropertyBits=   -1476378597
      BackColor       =   16777215
      Size            =   "9701;5292"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtOrg 
      Height          =   336
      Left            =   1200
      TabIndex        =   1
      Top             =   30
      Width           =   3504
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      Size            =   "6174;593"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100137_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2023/08/17
Option Explicit

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100137_1 = Nothing
End Sub
