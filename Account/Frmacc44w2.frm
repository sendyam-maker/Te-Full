VERSION 5.00
Begin VB.Form Frmacc44w2 
   AutoRedraw      =   -1  'True
   Caption         =   "代填繳款書客戶明細說明"
   ClientHeight    =   2256
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6108
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2256
   ScaleWidth      =   6108
   Begin VB.Image Image1 
      Height          =   132
      Left            =   3960
      Top             =   120
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "1. 若畫面全部不勾代表全部名單(含單筆收款扣繳合計未達2000元)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   5700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "說明："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   50
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "Frmacc44w2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Amy 2025/04/01 說明
Option Explicit

Private Sub Form_Load()
   Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single
   
   Me.Width = 6000
   Me.Height = 2112
   PUB_InitForm Me, Me.Width, Me.Height
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   Label2.Caption = "1. 若畫面全部不勾代表全部名單" & vbCrLf & _
                                    "     (含單筆收款扣繳合計未達2000元)" & vbCrLf & _
                                    "2. 勾選「單筆代填」或全部不勾,「單筆代填」都會出現" & vbCrLf & _
                                    "3. 剔除未達稅額2000者:請全部不勾選," & vbCrLf & _
                                    "    只勾選「不含單筆收款扣繳合計未達2000元 (Excel條件)」"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   Set Frmacc44w2 = Nothing
End Sub
