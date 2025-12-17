VERSION 5.00
Begin VB.Form Frmacc6200 
   AutoRedraw      =   -1  'True
   Caption         =   "關於"
   ClientHeight    =   5030
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "Frmacc6200.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5030
   ScaleWidth      =   7380
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "台一國際智慧財產事務所"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "授與使用權給"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '單線固定
      Height          =   735
      Left            =   840
      Stretch         =   -1  'True
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   $"Frmacc6200.frx":000C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   600
      X2              =   6705
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   600
      X2              =   6704
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "All rights reserved."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "寶研科技股份有限公司     "
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   16.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   4392
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Copyright (C)1998-2002 INFOBASE International, Inc."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   5556
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "財務管理系統  V1.0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   16.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   3180
   End
End
Attribute VB_Name = "Frmacc6200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/2/9 Form2.0不用改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Dim Counter As String

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 7500
   Me.Height = 5400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image2 = LoadPicture(strBackPicPath6)
   sglWidth = Image2.Width
   sglHeight = Image2.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image2, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Image1 = LoadPicture(strPicPath & "1.bmp")
   Timer1.Interval = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc6200 = Nothing
End Sub

Private Sub Timer1_Timer()
   Counter = Val(Counter) + 1
   Image1 = LoadPicture(strPicPath & Counter & ".bmp")
   If Val(Counter) = 20 Then
      Counter = "1"
   End If
End Sub
