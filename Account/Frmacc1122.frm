VERSION 5.00
Begin VB.Form Frmacc1122 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收據編號輸入"
   ClientHeight    =   780
   ClientLeft      =   40
   ClientTop       =   320
   ClientWidth     =   3510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3510
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   492
      Left            =   120
      Top             =   120
      Width           =   3252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "手開收據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "Frmacc1122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoacc1m0 As New ADODB.Recordset


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
   Me.Width = 3600
   Me.Height = 1150
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Checking
   If Text1 = MsgText(601) Then
      tool3_enabled
      Frmacc1120.Enabled = True
      Exit Sub
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select a0k01 from acc0k0 where a0k01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      tool3_enabled
      Cancel = 1
      adoacc0k0.Close
      Exit Sub
   End If
   adoacc0k0.Close
   adoacc1m0.CursorLocation = adUseClient
   adoacc1m0.Open "select a1m01 from acc1m0 where a1m01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1m0.RecordCount <> 0 Then
      MsgBox MsgText(42), , MsgText(5)
      tool3_enabled
      Cancel = 1
      adoacc1m0.Close
      Exit Sub
   End If
   strReceiptConfirm = MsgText(602)
   If Text1 <> MsgText(601) Then
      strItemNo = Text1
   Else
      strItemNo = MsgText(601)
   End If
   tool3_enabled
   Frmacc1121.Show
   Set Frmacc1122 = Nothing
Checking:
   tool3_enabled
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
