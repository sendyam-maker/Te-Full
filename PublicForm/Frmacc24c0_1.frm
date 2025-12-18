VERSION 5.00
Begin VB.Form Frmacc24c0_1 
   AutoRedraw      =   -1  'True
   Caption         =   "FC業務請款／收款明細表(預覽)"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5055
   Begin VB.HScrollBar HScroll1 
      Height          =   228
      Left            =   0
      Max             =   20
      TabIndex        =   2
      Top             =   2988
      Width           =   4800
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3192
      Left            =   4824
      TabIndex        =   1
      Top             =   0
      Width           =   228
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2940
      Left            =   0
      ScaleHeight     =   2880
      ScaleWidth      =   4710
      TabIndex        =   0
      Top             =   0
      Width           =   4776
   End
End
Attribute VB_Name = "Frmacc24c0_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Sindy 2010/9/2 日期欄已修改
Option Explicit

Public m_ImageH As Long, m_ImageW As Long, m_iPages As Integer

Dim m_iPageNow As Integer, m_iPageLast As Integer
Dim m_PicX As Long, m_PicY As Long
Dim iDotH As Integer, iDotV As Integer
Dim m_Pictures(1 To 2) As StdPicture

Private Sub Form_Load()

   iDotH = 40: iDotV = 30
   
   With Frmacc24c0_1
      .Top = 0
      .Left = 0
      .Width = 11800 '12000 'Screen.Width - 100
      .Height = 6800 '8500 'Screen.Height - 1500
      .Picture1.Width = .ScaleWidth - .VScroll1.Width
      .Picture1.Height = .ScaleHeight - .HScroll1.Height
      .VScroll1.Left = .Picture1.Width
      .VScroll1.Height = .Picture1.Height
      .HScroll1.Top = .Picture1.Height
      .HScroll1.Width = .Picture1.Width
   End With
  
   '垂直捲軸
   If Picture1.Height >= m_ImageH Then
      VScroll1.Visible = False
   Else
      VScroll1.max = iDotV * m_iPages - Int(iDotV * Picture1.Height / m_ImageH)
   End If
   
   '水平捲軸
   If Picture1.Width >= m_ImageW Then
      HScroll1.Visible = False
   Else
      HScroll1.max = iDotH - Int(iDotH * Picture1.Width / m_ImageW)
   End If
      
   '載入第一張報表
   m_iPageLast = 0
   m_iPageNow = 1
   m_PicX = 0: m_PicY = 0
   PaintPic
End Sub

Private Sub GetPic(idx As Integer, idx1 As Integer)
   Dim strPicFileName As String
   strPicFileName = App.Path & "\$tmp_" & idx & ".tmp"
   Set m_Pictures(idx1) = LoadPicture(strPicFileName)
End Sub

Private Sub PaintPic()
   Picture1.Line (0, 0)-(Picture1.Width, Picture1.Height), QBColor(15), BF
   If m_iPageLast <> m_iPageNow Then
      GetPic m_iPageNow, 1
      If m_iPages > m_iPageNow Then
         GetPic m_iPageNow + 1, 2
      End If
   End If
   
   Picture1.PaintPicture m_Pictures(1), m_PicX, m_PicY
   If m_iPages > m_iPageNow Then
      Picture1.PaintPicture m_Pictures(2), m_PicX, m_PicY + m_ImageH
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DelPic
   Set Frmacc24c0_1 = Nothing
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.Path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
End Sub

Private Sub HScroll1_Change()
   m_PicX = -1 * m_ImageW * HScroll1.Value / iDotH
   PaintPic
End Sub

Private Sub HScroll1_Scroll()
   m_PicX = -1 * m_ImageW * HScroll1.Value / iDotH
   PaintPic
End Sub

Private Sub VScroll1_Change()
   m_iPageLast = m_iPageNow
   m_iPageNow = 1 + VScroll1.Value \ iDotV
   m_PicY = -1 * m_ImageH * (VScroll1.Value Mod iDotV) / iDotV
   PaintPic
End Sub

Private Sub VScroll1_Scroll()
   m_iPageLast = m_iPageNow
   m_iPageNow = 1 + VScroll1.Value \ iDotV
   m_PicY = -1 * m_ImageH * (VScroll1.Value Mod iDotV) / iDotV
   PaintPic
End Sub
