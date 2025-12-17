VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1110 
   AutoRedraw      =   -1  'True
   Caption         =   "手開收據開立"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   5430
   Begin VB.TextBox txtLaw 
      Height          =   330
      Left            =   1665
      TabIndex        =   9
      Top             =   3870
      Width           =   555
   End
   Begin VB.TextBox txtTradeMark 
      Height          =   330
      Left            =   1665
      TabIndex        =   5
      Top             =   2400
      Width           =   555
   End
   Begin VB.TextBox txtPatent 
      Height          =   330
      Left            =   1665
      TabIndex        =   2
      Top             =   1230
      Width           =   555
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      TabIndex        =   24
      Top             =   180
      Width           =   1572
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   11
      Top             =   4230
      Width           =   1572
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   10
      Top             =   4230
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      TabIndex        =   20
      Top             =   3510
      Width           =   2652
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3510
      Width           =   852
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2760
      Width           =   1572
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1665
      MaxLength       =   15
      TabIndex        =   6
      Top             =   2745
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1575
      Width           =   1572
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
      Height          =   330
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1590
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
      Height          =   330
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   1
      Top             =   540
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
      Height          =   330
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Top             =   180
      Width           =   612
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   3240
      TabIndex        =   15
      Top             =   540
      Width           =   1935
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "3413;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "張數"
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
      Left            =   1125
      TabIndex        =   30
      Top             =   3900
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "律師"
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
      Left            =   270
      TabIndex        =   29
      Top             =   3330
      Width           =   450
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      Height          =   1395
      Left            =   135
      Top             =   3240
      Width           =   5190
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   1035
      Left            =   135
      Top             =   2100
      Width           =   5190
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1035
      Left            =   135
      Top             =   960
      Width           =   5190
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "張數"
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
      Left            =   1125
      TabIndex        =   28
      Top             =   2430
      Width           =   450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
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
      Left            =   675
      TabIndex        =   27
      Top             =   2790
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "張數"
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
      Left            =   1125
      TabIndex        =   26
      Top             =   1260
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
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
      Left            =   675
      TabIndex        =   25
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "目前編號"
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
      Left            =   2640
      TabIndex        =   23
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   22
      Top             =   4230
      Width           =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
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
      Left            =   675
      TabIndex        =   21
      Top             =   4230
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   900
      TabIndex        =   19
      Top             =   3540
      Width           =   675
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "商標"
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
      Left            =   270
      TabIndex        =   17
      Top             =   2190
      Width           =   450
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   1590
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "專利"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1050
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      TabIndex        =   13
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      TabIndex        =   12
      Top             =   240
      Width           =   612
   End
End
Attribute VB_Name = "Frmacc1110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Dim strJump As String
Dim strNo As String
Dim lngAmount1 As Long
Dim lngAmount2 As Long
Dim strAmount1 As String
Dim strAmount2 As String
Dim intLength As Integer
Dim intCounter As Integer


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 5670, 5200
   Text1 = Mid(ACDate(ServerDate), 1, 3)
   OpenTable
   AutoNoQuery
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1110 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select * from acc0k0 where rownum<1", adoTaie, adOpenStatic, adLockReadOnly
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_LostFocus()
   strJump = MsgText(601)
End Sub

Private Sub Text10_GotFocus()
   Dim stLastNum As String
   If Val(txtLaw) > 0 Then
      If Text7 <> "" And Text7 <> "E" Then
         stLastNum = Text7
      ElseIf Text5 <> "" And Text5 <> "E" Then
         stLastNum = Text5
      ElseIf Text12 <> "" Then
         stLastNum = Text12
      End If
      If stLastNum <> "" Then
         Text10 = Left(stLastNum, 5) & Format(Val(Mid(stLastNum, 6)) + 1, "0000")
         Text11 = Left(stLastNum, 5) & Format(Val(Mid(stLastNum, 6)) + Val(txtLaw), "0000")
      End If
   End If
   If Len(Text10) > 0 And Left(Text10, 1) = "E" Then
      Text10.SelStart = 1
      Text10.SelLength = Len(Text10) - 1
   End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_LostFocus()
   Text10Jump
End Sub

Private Sub Text11_GotFocus()
   If Len(Text11) > 0 And Left(Text11, 1) = "E" Then
      Text11.SelStart = 1
      Text11.SelLength = Len(Text11) - 1
   End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_LostFocus()
   If strJump = MsgText(602) Then
      strJump = MsgText(601)
      Exit Sub
   End If
   strJump = MsgText(601)
   If Text10 = MsgText(601) Then
      If Text11 = MsgText(601) Then
         Exit Sub
      End If
   Else
      If Val(Mid(Text11, 5, 5)) >= Val(Mid(Text10, 5, 5)) Then
         Exit Sub
      End If
   End If
   MsgBox MsgText(32), , MsgText(5)
   strJump = MsgText(602)
   Text11.SetFocus
End Sub

Private Sub Text2_Change()
   Text3 = StaffQuery(Text2)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_LostFocus()
   strJump = MsgText(601)
End Sub

Private Sub Text4_GotFocus()
   If Val(txtPatent) > 0 Then
      If Text12 <> "" Then
         Text4 = Left(Text12, 5) & Format(Val(Mid(Text12, 6)) + 1, "0000")
         Text5 = Left(Text12, 5) & Format(Val(Mid(Text12, 6)) + Val(txtPatent), "0000")
      End If
   End If
   If Len(Text4) > 0 And Left(Text4, 1) = "E" Then
      Text4.SelStart = 1
      Text4.SelLength = Len(Text4) - 1
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_LostFocus()
   If Me.Text4.Text = "" Or Me.Text4.Text = "E" Then Exit Sub
   Text4Jump
End Sub

Private Sub Text5_GotFocus()
   If Text5 = "" Then
      Text5 = Text4
   End If
   If Len(Text5) > 0 And Left(Text5, 1) = "E" Then
      Text5.SelStart = 1
      Text5.SelLength = Len(Text5) - 1
   End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
   If Me.Text5.Text = "" Or Me.Text5.Text = "E" Then Exit Sub
   Text5Jump
End Sub

Private Sub Text6_GotFocus()
   Dim stLastNum As String
   If Val(txtTradeMark) > 0 Then
      If Text5 <> "" And Text5 <> "E" Then
         stLastNum = Text5
      ElseIf Text12 <> "" Then
         stLastNum = Text12
      End If
      If stLastNum <> "" Then
         Text6 = Left(stLastNum, 5) & Format(Val(Mid(stLastNum, 6)) + 1, "0000")
         Text7 = Left(stLastNum, 5) & Format(Val(Mid(stLastNum, 6)) + Val(txtTradeMark), "0000")
      End If
   End If
   If Len(Text6) > 0 And Left(Text6, 1) = "E" Then
      Text6.SelStart = 1
      Text6.SelLength = Len(Text6) - 1
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_LostFocus()
   If Me.Text6.Text = "" Or Me.Text6.Text = "E" Then Exit Sub
   Text6Jump
End Sub

Private Sub Text7_GotFocus()
   If Text7 = "" Then
      Text7 = Text6
   End If
   If Len(Text7) > 0 And Left(Text7, 1) = "E" Then
      Text7.SelStart = 1
      Text7.SelLength = Len(Text7) - 1
   End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_LostFocus()
    If Me.Text7.Text = "" Or Me.Text7.Text = "E" Then Exit Sub
   strJump = MsgText(601)
   If strJump = MsgText(602) Then
      strJump = MsgText(601)
      Exit Sub
   End If
   If Text6 = MsgText(601) Then
      If Text7 = MsgText(601) Then
         Exit Sub
      End If
   Else
      If Val(Mid(Text7, 5, 5)) >= Val(Mid(Text6, 5, 5)) Then
         Exit Sub
      End If
   End If
   MsgBox MsgText(32), , MsgText(5)
   strJump = MsgText(602)
   Text7.SetFocus
End Sub

Private Sub Text8_Change()
   Text9 = A0802Query(Text8)
End Sub
'*************************************************
'  查詢目前收據編號
'
'*************************************************
Public Sub AutoNoQuery()

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open "select max(a0k01) from acc0k0 where a0k01 <= '" & "E" & Mid(CFDate(ACDate(ServerDate)), 1, 3) & "02000" & "'", adoTaie, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         If IsNull(.Fields(0).Value) Then
            Text12 = "E" & Text1 & "00000"
         Else
            If Mid(.Fields(0).Value, 2, 3) <> Text1 Then  '非系統年
               CheckOC3
               With AdoRecordSet3
                  .CursorLocation = adUseClient
                  .Open "select max(a0k01) from acc0k0 where a0k01 <= '" & "E" & Text1 & "02000" & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If .RecordCount <> 0 Then
                     If IsNull(.Fields(0).Value) Then
                        Text12 = "E" & Text1 & "00000"
                     Else
                        If Mid(.Fields(0).Value, 2, 3) <> Text1 Then
                           Text12 = "E" & Text1 & "00000"
                        Else
                           Text12 = .Fields(0).Value
                        End If
                     End If
                  Else
                     Text12 = "E" & Text1 & "00000"
                  End If
               End With
            Else
               Text12 = .Fields(0).Value                  '系統年
            End If
         End If
      Else
         Text12 = "E" & Text1 & "00000"
      End If
   End With
End Sub
'*************************************************
'  Text4 跳位控制
'
'*************************************************
Private Sub Text4Jump()
   strJump = MsgText(601)
   If Text4 = MsgText(601) Or Text4 = MsgText(802) Then
      Exit Sub
   End If
   If Text12 = MsgText(601) Then
      If Val(Mid(Text4, 5, 5)) = 1 Then
         Exit Sub
      End If
   Else
      If Mid(Text4, 1, 4) = Mid(Text12, 1, 4) Then
         If Val(Mid(Text4, 5, 5)) = (Val(Mid(Text12, 5, 5)) + 1) Then
            Exit Sub
         End If
      Else
         Exit Sub
      End If
   End If
   MsgBox MsgText(32), , MsgText(5)
   strJump = MsgText(602)
   Text4.SetFocus
End Sub
'*************************************************
'  Text5 跳位控制
'
'*************************************************
Private Sub Text5Jump()
   strJump = MsgText(601)
   If Text4 = MsgText(601) Or Text4 = MsgText(802) Then
      If Text5 = MsgText(601) Or Text5 = MsgText(802) Then
         Exit Sub
      End If
   Else
      If Val(Mid(Text5, 5, 5)) >= Val(Mid(Text4, 5, 5)) Then
         Exit Sub
      End If
   End If
   MsgBox MsgText(32), , MsgText(5)
   strJump = MsgText(602)
   Text5.SetFocus
End Sub

'*************************************************
'  Text6 跳位控制
'
'*************************************************
Private Sub Text6Jump()
   strJump = MsgText(601)
   If Text6 = MsgText(601) Or Text6 = MsgText(802) Then
      Exit Sub
   End If
   If Text5 <> MsgText(601) And Text5 <> "E" Then
      If Val(Mid(Text6, 5, 5)) = (Val(Mid(Text5, 5, 5)) + 1) Then
         Exit Sub
      End If
   Else
      If Text12 = MsgText(601) Then
         If Val(Mid(Text6, 5, 5)) = 1 Then
            Exit Sub
         End If
      Else
         If Val(Mid(Text6, 5, 5)) = (Val(Mid(Text12, 5, 5)) + 1) Then
            Exit Sub
         End If
      End If
   End If
   MsgBox MsgText(32), , MsgText(5)
   strJump = MsgText(602)
   Text6.SetFocus
End Sub

'*************************************************
'  Text10 跳位控制
'
'*************************************************
Private Sub Text10Jump()
   strJump = MsgText(601)
   If Text10 = MsgText(601) Or Text10 = MsgText(802) Then
      Exit Sub
   End If
   If Text7 <> MsgText(601) Or Text7 = MsgText(802) Then
      If Val(Mid(Text10, 5, 5)) = (Val(Mid(Text7, 5, 5)) + 1) Or Text7 = MsgText(802) Then
         Exit Sub
      End If
   Else
      If Text5 <> MsgText(601) Or Text5 = MsgText(802) Then
         If Val(Mid(Text10, 5, 5)) = (Val(Mid(Text5, 5, 5)) + 1) Or Text5 = MsgText(802) Then
            Exit Sub
         End If
      Else
         If Text12 <> MsgText(601) Or Text12 = MsgText(802) Then
            If Val(Mid(Text10, 5, 5)) = (Val(Mid(Text12, 5, 5)) + 1) Or Text12 = MsgText(802) Then
               Exit Sub
            End If
         Else
            If Val(Mid(Text10, 5, 5)) = 1 Then
               Exit Sub
            End If
         End If
      End If
   End If
   MsgBox MsgText(32), , MsgText(5)
   strJump = MsgText(602)
   Text10.SetFocus
End Sub

Private Sub Text8_LostFocus()
   strJump = MsgText(601)
End Sub

'*************************************************
'  列印
'
'*************************************************
Public Sub PrintDoc(strStartNo As String, strEndNo As String, intI As Integer)
   Screen.MousePointer = vbHourglass
   strNo = ""
   intLength = 0
   strSql = MsgText(601)
   If strStartNo <> MsgText(601) Then
      strSql = strSql & " and a0k01 >= '" & strStartNo & "'"
   End If
   If strEndNo <> MsgText(601) Then
      strSql = strSql & " and a0k01 <= '" & strEndNo & "'"
   End If
   If strSql = MsgText(601) Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   'Modify by Morgan 2008/3/25 控制 9x 才自訂
   '9x
   If pub_OS = "1" Then
      Printer.Height = 8750
      Printer.Width = 13000
   Else
      Printer.PaperSize = PUB_GetPaperSize(1)
   End If
   'end 2008/3/25
   Printer.FontSize = 12
   CheckOC
   With adoacc0k0
      If .State = adStateOpen Then .Close
      .CursorLocation = adUseClient
      .Open "select * from acc0k0, customer where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+)" & strSql & " order by a0k01 asc", adoTaie, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         Select Case intI
            Case 1
               MsgBox MsgText(94), , MsgText(5)
            Case 2
               MsgBox MsgText(95), , MsgText(5)
            Case 3
               MsgBox MsgText(96), , MsgText(5)
         End Select
      End If
      Do While .EOF = False
         If strNo <> .Fields("a0k01").Value Then
            If strNo <> "" Then
               Printer.NewPage
            End If
            intCounter = 0
            If IsNull(.Fields("a0k19").Value) Then
               intCounter = 1
            Else
               intCounter = Val(.Fields("a0k19").Value) + 1
            End If
            adoTaie.Execute "update acc0k0 set a0k19 = " & intCounter & " where a0k01 = '" & .Fields("a0k01").Value & "'"
            PrintHead intI
            strNo = .Fields("a0k01").Value
         End If
         .MoveNext
      Loop
   End With
   Printer.EndDoc
   Screen.MousePointer = vbDefault
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead(intI As Integer)
   If intI = 3 Then
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         .Open "select * from acc080 where a0801 = '" & adoacc0k0.Fields("a0k11").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .RecordCount <> 0 Then
            Printer.FontSize = 16
            Printer.CurrentX = 4300
            Printer.CurrentY = 200
            If IsNull(.Fields("a0802").Value) Then
               Printer.Print ""
            Else
               Printer.Print .Fields("a0802").Value
            End If
            Printer.FontSize = 10
            Printer.CurrentX = 3400
            Printer.CurrentY = 600
            If IsNull(.Fields("a0804").Value) Then
               Printer.Print "地址: "
            Else
               Printer.Print "地址: " & .Fields("a0804").Value
            End If
            Printer.CurrentX = 3400
            Printer.CurrentY = 800
            If IsNull(.Fields("a0813").Value) Then
               Printer.Print "電話: "
            Else
               Printer.Print "電話: " & .Fields("a0813").Value
            End If
            Printer.CurrentX = 6300
            Printer.CurrentY = 800
            '2008/12/15 MODIFY BY SONIA 與婧瑄確認應為統一編號
            'If IsNull(.Fields("a0814").Value) Then
            '   Printer.Print "扣繳編號: "
            'Else
            '   Printer.Print "扣繳編號: " & .Fields("a0814").Value
            'End If
            If IsNull(.Fields("a0807").Value) Then
               Printer.Print "扣繳編號: "
            Else
               Printer.Print "扣繳編號: " & .Fields("a0807").Value
            End If
            '2008/12/15 END
         End If
         .Close
      End With
   End If
   Printer.FontSize = 12
   Printer.CurrentX = 1200
   Printer.CurrentY = 2200
   If IsNull(adoacc0k0.Fields("a0k03").Value) = False Then
      Printer.Print adoacc0k0.Fields("a0k03").Value
   Else
      Printer.Print ""
   End If
   Printer.CurrentX = 8200
   Printer.CurrentY = 2200
   Printer.Print adoacc0k0.Fields("a0k01").Value
   Printer.CurrentX = 1200
   Printer.CurrentY = 2450
   If IsNull(adoacc0k0.Fields("a0k04").Value) = False Then
      Printer.Print adoacc0k0.Fields("a0k04").Value
   Else
      Printer.Print ""
   End If
   Printer.CurrentX = 8200
   Printer.CurrentY = 2450
   If IsNull(adoacc0k0.Fields("a0k20").Value) = False Then
      Printer.Print StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
      Printer.CurrentX = 8650
      Printer.CurrentY = 2450
      Printer.Print adoacc0k0.Fields("a0k20").Value
   Else
      Printer.Print ""
   End If
   Printer.CurrentX = 1200
   Printer.CurrentY = 2700
   If IsNull(adoacc0k0.Fields("cu31").Value) = False Then
      Printer.Print adoacc0k0.Fields("cu31").Value
   Else
      Printer.Print ""
   End If
   Printer.CurrentX = 10200
   Printer.CurrentY = 6800
   Printer.Print intCounter
   
End Sub

Private Sub txtLaw_GotFocus()
   TextInverse txtLaw
End Sub

Private Sub txtLaw_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii = 8 Or IsNumeric(Chr(KeyAscii))) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtPatent_GotFocus()
   TextInverse txtPatent
End Sub

Private Sub txtPatent_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii = 8 Or IsNumeric(Chr(KeyAscii))) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtTradeMark_GotFocus()
   TextInverse txtTradeMark
End Sub

Private Sub txtTradeMark_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii = 8 Or IsNumeric(Chr(KeyAscii))) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Public Function SaveCheck() As Boolean
   Text4_LostFocus
   If strJump <> "" Then
      Text4_GotFocus
      SaveCheck = False
      Exit Function
   End If
   Text5_LostFocus
   If strJump <> "" Then
      Text5_GotFocus
      SaveCheck = False
      Exit Function
   End If
   Text6_LostFocus
   If strJump <> "" Then
      Text6_GotFocus
      SaveCheck = False
      Exit Function
   End If
   Text7_LostFocus
   If strJump <> "" Then
      Text7_GotFocus
      SaveCheck = False
      Exit Function
   End If
   Text10_LostFocus
   If strJump <> "" Then
      Text10_GotFocus
      SaveCheck = False
      Exit Function
   End If
   Text11_LostFocus
   If strJump <> "" Then
      Text11_GotFocus
      SaveCheck = False
      Exit Function
   End If
   SaveCheck = True
End Function
