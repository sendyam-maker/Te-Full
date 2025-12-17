VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc3420 
   AutoRedraw      =   -1  'True
   Caption         =   "應付票據資料表"
   ClientHeight    =   3984
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3984
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   90
      Width           =   3525
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Top             =   1930
      Width           =   612
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   1570
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   500
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   500
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   2280
      Width           =   4692
   End
   Begin VB.ComboBox Combo12 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      TabIndex        =   18
      Top             =   4880
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo11 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   17
      Top             =   4880
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo10 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      TabIndex        =   16
      Top             =   4520
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   15
      Top             =   4520
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      TabIndex        =   14
      Top             =   4160
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   13
      Top             =   4160
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      TabIndex        =   12
      Top             =   3800
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   11
      Top             =   3800
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      TabIndex        =   10
      Top             =   3440
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   9
      Top             =   3440
      Visible         =   0   'False
      Width           =   1812
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   850
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3240
      TabIndex        =   4
      Top             =   850
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   1210
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   300
      Left            =   3240
      TabIndex        =   6
      Top             =   1210
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   36
      Top             =   150
      Width           =   972
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "是否含作廢票據(Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   1930
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "(N不含)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   34
      Top             =   1930
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   1210
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "開票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   1210
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(N不含)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   31
      Top             =   1570
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "是否含已兌領票據(Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   1570
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   29
      Top             =   500
      Width           =   252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "開票帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   28
      Top             =   490
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   27
      Top             =   850
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   26
      Top             =   850
      Width           =   252
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2415
      Left            =   240
      Top             =   2960
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3420.frx":0000
      Stretch         =   -1  'True
      Top             =   4880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   4880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3420.frx":0442
      Stretch         =   -1  'True
      Top             =   4520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   4520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3420.frx":0884
      Stretch         =   -1  'True
      Top             =   4160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   4160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3420.frx":0CC6
      Stretch         =   -1  'True
      Top             =   3800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   22
      Top             =   3800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3420.frx":1108
      Stretch         =   -1  'True
      Top             =   3440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   3440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc3420"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc0e0 As New ADODB.Recordset
Public adoaccrpt302 As New ADODB.Recordset
Dim strSort1, strSort2, strSort3, strSort4, strSort5 As String
Dim dllaccrpt302 As Object
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17


'Add by Sindy 2020/04/17
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, False, True)
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label17 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17

Private Sub Combo10_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo11.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo11_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo12.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo12_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text1.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo4.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo5.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo6.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo6_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo7.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo7_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo8.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo8_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo9.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo9_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo10.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
Dim bCancel As Boolean
   
   'Add By Sindy 2020/4/23
   If CboCmp.Text = MsgText(601) Then
      MsgBox Label17 & MsgText(52), , MsgText(5)
      Exit Sub
   End If
   Call CboCmp_Validate(bCancel)
   If bCancel = True Then
      Exit Sub
   End If
   '2020/4/23 END
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   Call SetCompN 'Add by Sindy 2020/04/23
   
   Screen.MousePointer = vbHourglass
   Accrpt302Delete
   ProduceData
   If adoaccrpt302.State = adStateOpen Then
      adoaccrpt302.Close
   End If
   adoaccrpt302.CursorLocation = adUseClient
   adoaccrpt302.Open "select * from accrpt302", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt302.RecordCount <> 0 Then
      '20140120START Modify By eric
      dllaccrpt302.Acc3420 ReportTitle(302) & "-" & strCmpN, Text1, Text2, MaskEdBox1.Text, MaskEdBox2.Text, MaskEdBox3.Text, MaskEdBox4.Text, Text3, Text4, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      'dllaccrpt302.Acc3420 ReportTitle(302), Text1, Text2, MaskEdBox1.Text, MaskEdBox2.Text, MaskEdBox3.Text, MaskEdBox4.Text, Text3, Text4, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      '20140120END
   End If
   adoaccrpt302.Close
   Screen.MousePointer = vbDefault
   FormClear
   'Modify By Cheng 2002/01/17
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Frmacc0000.StatusBar1.Panels(1).Text = "請更換為 A4 紙張"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      'Modify By Cheng 2002/01/17
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
      Frmacc0000.StatusBar1.Panels(1).Text = "請更換為 A4 紙張"
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 3400 'Modify by Amy 2023/08/18 原:3100
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo6.AddItem MsgText(1)
   Combo6.AddItem MsgText(2)
   Combo8.AddItem MsgText(1)
   Combo8.AddItem MsgText(2)
   Combo10.AddItem MsgText(1)
   Combo10.AddItem MsgText(2)
   Combo12.AddItem MsgText(1)
   Combo12.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   Combo8 = MsgText(1)
   Combo10 = MsgText(1)
   Combo12 = MsgText(1)
   ComboAdd
   'Modify By Cheng 2002/01/17
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Frmacc0000.StatusBar1.Panels(1).Text = "請更換為 A4 紙張"
   Set dllaccrpt302 = CreateObject("AccReport.ReportSelect")
   
   'Add by Sindy 2020/04/17 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, False, False, False, , 1)
   'end 2020/04/17
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt302 = Nothing
   Set Frmacc3420 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "單據號碼"
   strSort2 = "收票日期"
   strSort3 = "客戶代號"
   strSort4 = "票據號碼"
   strSort5 = "收票銀行"
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
   Combo3.AddItem strSort3
   Combo3.AddItem strSort4
   Combo3.AddItem strSort5
   Combo5.AddItem strSort1
   Combo5.AddItem strSort2
   Combo5.AddItem strSort3
   Combo5.AddItem strSort4
   Combo5.AddItem strSort5
   Combo7.AddItem strSort1
   Combo7.AddItem strSort2
   Combo7.AddItem strSort3
   Combo7.AddItem strSort4
   Combo7.AddItem strSort5
   Combo9.AddItem strSort1
   Combo9.AddItem strSort2
   Combo9.AddItem strSort3
   Combo9.AddItem strSort4
   Combo9.AddItem strSort5
   Combo11.AddItem strSort1
   Combo11.AddItem strSort2
   Combo11.AddItem strSort3
   Combo11.AddItem strSort4
   Combo11.AddItem strSort5
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1, strOrder2, strOrder3, strOrder4, strOrder5 As String
Dim strSql As String

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e03 asc"
         Else
            strOrder1 = " order by a0e03 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e13 asc"
         Else
            strOrder1 = " order by a0e13 desc"
         End If
      Case strSort3
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e06 asc"
         Else
            strOrder1 = " order by a0e06 desc"
         End If
      Case strSort4
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e02 asc"
         Else
            strOrder1 = " order by a0e02 desc"
         End If
      Case strSort5
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e01 asc"
         Else
            strOrder1 = " order by a0e01 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e03 asc"
         Else
            strOrder2 = ", a0e03 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e13 asc"
         Else
            strOrder2 = ", a0e13 desc"
         End If
      Case strSort3
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e06 asc"
         Else
            strOrder2 = ", a0e06 desc"
         End If
      Case strSort4
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e02 asc"
         Else
            strOrder2 = ", a0e02 desc"
         End If
      Case strSort5
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e01 asc"
         Else
            strOrder2 = ", a0e01 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo7
      Case strSort1
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e03 asc"
         Else
            strOrder3 = ", a0e03 desc"
         End If
      Case strSort2
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e13 asc"
         Else
            strOrder3 = ", a0e13 desc"
         End If
      Case strSort3
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e06 asc"
         Else
            strOrder3 = ", a0e06 desc"
         End If
      Case strSort4
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e02 asc"
         Else
            strOrder3 = ", a0e02 desc"
         End If
      Case strSort5
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e01 asc"
         Else
            strOrder3 = ", a0e01 desc"
         End If
      Case Else
         strOrder3 = MsgText(601)
   End Select
   Select Case Combo9
      Case strSort1
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e03 asc"
         Else
            strOrder4 = ", a0e03 desc"
         End If
      Case strSort2
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e13 asc"
         Else
            strOrder4 = ", a0e13 desc"
         End If
      Case strSort3
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e06 asc"
         Else
            strOrder4 = ", a0e06 desc"
         End If
      Case strSort4
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e02 asc"
         Else
            strOrder4 = ", a0e02 desc"
         End If
      Case strSort5
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e01 asc"
         Else
            strOrder4 = ", a0e01 desc"
         End If
      Case Else
         strOrder4 = MsgText(601)
   End Select
   Select Case Combo11
      Case strSort1
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e03 asc"
         Else
            strOrder5 = ", a0e03 desc"
         End If
      Case strSort2
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e13 asc"
         Else
            strOrder5 = ", a0e13 desc"
         End If
      Case strSort3
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e06 asc"
         Else
            strOrder5 = ", a0e06 desc"
         End If
      Case strSort4
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e02 asc"
         Else
            strOrder5 = ", a0e02 desc"
         End If
      Case strSort5
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e01 asc"
         Else
            strOrder5 = ", a0e01 desc"
         End If
      Case Else
         strOrder5 = MsgText(601)
   End Select
   
   '20140120START Modify By eric
   'Modify By Sindy 2020/4/23
   'strSql = " and a0e23 = '" & IIf(Text5 = "2", "J", "1") & "' "
   strSql = " and a0e23 = '" & strCmp & "' "
   '2020/4/23 END
   
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0e07 >= '" & Text1 & "'"
   End If
   'If Text1 <> MsgText(601) Then
   '   strSql = " and a0e07 >= '" & Text1 & "'"
   'End If
   '20140120END
   
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0e07 <= '" & Text2 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
      strSql = strSql & " and a0e13 >= " & Val(FCDate(MaskEdBox3.Text)) & ""
   End If
   If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
      strSql = strSql & " and a0e13 <= " & Val(FCDate(MaskEdBox4.Text)) & ""
   End If
   If Text3 = MsgText(603) Then
      strSql = strSql & " and (a0e37 = 0 or a0e37 is null)"
   End If
   If Text4 = MsgText(603) Then
      strSql = strSql & " and (a0e25 is null or a0e25 = 0)"
   End If
  
   adoaccrpt302.CursorLocation = adUseClient
   adoaccrpt302.Open "select * from accrpt302", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "'" & strSql & strOrder1 & strOrder2 & strOrder3 & strOrder4 & strOrder5, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      adoacc0e0.Close
      adoaccrpt302.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0e0.EOF = False
      If adoaccrpt302.RecordCount = 0 Then
         adoaccrpt302.AddNew
         adoaccrpt302.Fields("r30201").Value = strUserNum
         adoaccrpt302.UpdateBatch
      End If
      adoaccrpt302.AddNew
      adoaccrpt302.Fields("r30201").Value = strUserNum
      If IsNull(adoacc0e0.Fields("a0e03").Value) Then
         adoaccrpt302.Fields("r30202").Value = Null
      Else
         adoaccrpt302.Fields("r30202").Value = adoacc0e0.Fields("a0e03").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e13").Value) Then
         adoaccrpt302.Fields("r30203").Value = Null
      Else
         adoaccrpt302.Fields("r30203").Value = adoacc0e0.Fields("a0e13").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e06").Value) Then
         adoaccrpt302.Fields("r30204").Value = Null
      Else
         Select Case adoacc0e0.Fields("a0e05").Value
            Case "1"
               adoaccrpt302.Fields("r30204").Value = MidB(CustomerQuery(adoacc0e0.Fields("a0e06").Value, 1), 1, 20)
            Case "2"
               adoaccrpt302.Fields("r30204").Value = MidB(A0i02Query(adoacc0e0.Fields("a0e06").Value), 1, 20)
            Case "3"
               adoaccrpt302.Fields("r30204").Value = MidB(StaffQuery(adoacc0e0.Fields("a0e06").Value), 1, 20)
         End Select
      End If
      If IsNull(adoacc0e0.Fields("a0e08").Value) Then
         adoaccrpt302.Fields("r30205").Value = Null
      Else
         Select Case adoacc0e0.Fields("a0e08").Value
            Case Mid(ComboItem(11), 1, 1)
               adoaccrpt302.Fields("r30205").Value = Mid(ComboItem(11), 4, 2)
            Case Mid(ComboItem(12), 1, 1)
               adoaccrpt302.Fields("r30205").Value = Mid(ComboItem(12), 4, 2)
            Case Mid(ComboItem(13), 1, 1)
               adoaccrpt302.Fields("r30205").Value = Mid(ComboItem(13), 4, 2)
         End Select
      End If
      If IsNull(adoacc0e0.Fields("a0e10").Value) Then
         adoaccrpt302.Fields("r30207").Value = Null
      Else
         adoaccrpt302.Fields("r30207").Value = adoacc0e0.Fields("a0e10").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e02").Value) Then
         adoaccrpt302.Fields("r30208").Value = Null
      Else
         adoaccrpt302.Fields("r30208").Value = adoacc0e0.Fields("a0e02").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e11").Value) Then
         adoaccrpt302.Fields("r30209").Value = 0
      Else
         adoaccrpt302.Fields("r30209").Value = Val(adoacc0e0.Fields("a0e11").Value)
      End If
      If IsNull(adoacc0e0.Fields("a0e25").Value) Then
         adoaccrpt302.Fields("r30210").Value = Null
      Else
         adoaccrpt302.Fields("r30210").Value = adoacc0e0.Fields("a0e25").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e37").Value) Then
         adoaccrpt302.Fields("r30211").Value = Null
      Else
         adoaccrpt302.Fields("r30211").Value = adoacc0e0.Fields("a0e37").Value
      End If
      adoaccrpt302.Fields("r30212").Value = adoacc0e0.Fields("a0e01").Value
      If IsNull(adoacc0e0.Fields("a0e07").Value) Then
         adoaccrpt302.Fields("r30213").Value = Null
      Else
         adoaccrpt302.Fields("r30213").Value = adoacc0e0.Fields("a0e07").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e12").Value) Then
         adoaccrpt302.Fields("r30214").Value = Null
      Else
         adoaccrpt302.Fields("r30214").Value = adoacc0e0.Fields("a0e12").Value
      End If
      adoaccrpt302.UpdateBatch
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   adoaccrpt302.Close
   adoTaie.Execute "delete from accrpt302 where r30208 is null"
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt302Delete()
   adoTaie.Execute "delete from accrpt302"
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = ""
   MaskEdBox4.Text = ""
   MaskEdBox4.Mask = DFormat
   Text3 = ""
   Combo3 = ""
   Combo5 = ""
   Combo7 = ""
   Combo9 = ""
   Combo11 = ""
   Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox3.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox4.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Mark by Sindy 2020/4/23 公司別改下拉式選單
'Private Sub Text5_Change()
'   Label19.Caption = A0802Query(IIf(Text5 = "2", "J", "1"))
'End Sub
'
''20140120START By eric
'Private Sub Text5_LostFocus()
'   If Me.Text5.Text = "" Then
'      MsgBox "公司別不可空白 !"
'      Me.Text5.SetFocus
'      Exit Sub
'   End If
'   If Me.Text5.Text <> "1" And Me.Text5.Text <> "2" Then
'      MsgBox "公司別僅能為 1 或 2 !"
'      Me.Text5.Text = ""
'      Me.Text5.SetFocus
'      Exit Sub
'   End If
'End Sub
'
''20140120START By eric
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'   CloseIme
'End Sub
