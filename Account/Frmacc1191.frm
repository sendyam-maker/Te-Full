VERSION 5.00
Begin VB.Form Frmacc1191 
   AutoRedraw      =   -1  'True
   Caption         =   "付款方式輸入"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   6255
   Begin VB.OptionButton optDeliver 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   4635
      TabIndex        =   15
      Top             =   1590
      Width           =   240
   End
   Begin VB.OptionButton optDeliver 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2880
      TabIndex        =   13
      Top             =   1590
      Width           =   240
   End
   Begin VB.OptionButton optDeliver 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1575
      TabIndex        =   11
      Top             =   1590
      Width           =   240
   End
   Begin VB.OptionButton optDeliver 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   9
      Top             =   1590
      Value           =   -1  'True
      Width           =   240
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   1
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text3 
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
      Left            =   4320
      MaxLength       =   8
      TabIndex        =   2
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   3
      Top             =   960
      Width           =   1572
   End
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
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "寄出特別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   4995
      TabIndex        =   16
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "交智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3240
      TabIndex        =   14
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "寄分所"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1935
      TabIndex        =   12
      Top             =   1590
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "寄出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   630
      TabIndex        =   10
      Top             =   1590
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   972
      Left            =   3240
      Top             =   480
      Width           =   2772
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   972
      Left            =   240
      Top             =   480
      Width           =   2772
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收票銀行"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "票期"
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
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "(1.開票 2.退原票)"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   2052
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "付款方式"
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc1191"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset


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
   Me.Width = 6350
   Me.Height = 2340
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   If adoquery.RecordCount <> 0 Then
      FormShow
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Add by Morgan 2006/10/31
   If optDeliver(1).Value = True Or optDeliver(2).Value = True Then
       If Frmacc1190.CheckSendOpt(IIf(optDeliver(1).Value, 1, 2)) = False Then
         Cancel = 1
         Exit Sub
       End If
   End If
   'end 2006/10/31
   
   If Text1 = MsgText(601) Then
      strCon1 = MsgText(601)
   Else
      strCon1 = Text1
   End If
   If Text4 = MsgText(601) Then
      strCon2 = MsgText(601)
   Else
      strCon2 = Text4
   End If
   If Text3 = MsgText(601) Then
      strCon3 = MsgText(601)
   Else
      strCon3 = Text3
   End If
   If Text2 = MsgText(601) Then
      strCon4 = MsgText(601)
   Else
      strCon4 = Text2
   End If
   strCon5 = MsgText(601)
   tool2_enabled
   Frmacc1190.Enabled = True
   Frmacc1190.Text17.SetFocus
   'Add by Morgan 2004/12/28 回傳銷退方式選項 0=寄出 1=寄分所 2=交智權人員 3=寄出特別
   Frmacc1190.m_stDeliver = Abs(optDeliver(1).Value) * 1 + Abs(optDeliver(2).Value) * 2 + Abs(optDeliver(3).Value) * 3
   Set Frmacc1191 = Nothing
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   Select Case Text1
      Case "1"
         Text4.Enabled = True
         Text2.Enabled = False
         Text3.Enabled = False
      Case "2"
         Text4.Enabled = False
         Text2.Enabled = True
         Text3.Enabled = True
   End Select
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a0e01 from acc0e0 where a0e01 = '" & Text2 & "' and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
      MsgBox MsgText(111), , MsgText(5)
      Cancel = True
      Text2.SetFocus
   End If
   adocheck.Close
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc0s0 where a0s01 = '" & strCon5 & "'", adoTaie, adOpenStatic, adLockReadOnly
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國內銷帳退費資料)
'
'*************************************************
Public Sub FormShow()
   If IsNull(adoquery.Fields("a0s19").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoquery.Fields("a0s19").Value
   End If
   If IsNull(adoquery.Fields("a0s22").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoquery.Fields("a0s22").Value
   End If
   If IsNull(adoquery.Fields("a0s20").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoquery.Fields("a0s20").Value
   End If
   If IsNull(adoquery.Fields("a0s21").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoquery.Fields("a0s21").Value
   End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = MsgText(601) Then
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a0e01 from acc0e0 where a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
      MsgBox MsgText(111), , MsgText(5)
      Cancel = True
      Text3.SetFocus
   End If
   adocheck.Close
End Sub
