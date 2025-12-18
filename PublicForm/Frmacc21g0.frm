VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21g0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "請款項目資料維護"
   ClientHeight    =   3660
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3660
   ScaleWidth      =   8760
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7650
      MaxLength       =   5
      TabIndex        =   14
      Top             =   3135
      Width           =   855
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      MaxLength       =   5
      TabIndex        =   13
      Top             =   3135
      Width           =   855
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   12
      Top             =   3135
      Width           =   855
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
      Height          =   315
      Left            =   2520
      TabIndex        =   28
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   7080
      Picture         =   "Frmacc21g0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   240
      Width           =   350
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      MaxLength       =   14
      TabIndex        =   11
      Top             =   2760
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      MaxLength       =   14
      TabIndex        =   9
      Top             =   2400
      Width           =   1572
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   14
      TabIndex        =   8
      Top             =   2400
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      MaxLength       =   6
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSForms.TextBox Text11 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   2040
      Width           =   6825
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "12039;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1680
      Width           =   6825
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "12039;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   6825
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "12039;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   6825
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "12039;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   6825
      VariousPropertyBits=   671107099
      Size            =   "12039;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "EXPENSE_CODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2835
      TabIndex        =   31
      Top             =   3180
      Width           =   1800
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "ACTIVITY_CODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5715
      TabIndex        =   30
      Top             =   3180
      Width           =   1890
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "TASK_CODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   29
      Top             =   3180
      Width           =   1395
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "日文名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "固定請款金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   26
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "上限"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "下限"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "請款金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   3435
      Left            =   150
      Top             =   120
      Width           =   8475
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "英文名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "中文名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "項目代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "系統類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc21g0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (Text3,Text4,Text5,Text6,Text11)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
Public adoacc1j0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
'Add By Cheng 2002/02/04
Dim m_QuerySystem As String

Private Sub Command5_Click()
   If adoacc1j0.RecordCount = 0 Or Text1 = MsgText(601) Or Text2 = MsgText(601) Then
      Exit Sub
   End If
   adoacc1j0.Find "a1j01 = '" & Text1 & "'", 0, adSearchForward, 1
   If adoacc1j0.EOF = False Then
      adoacc1j0.Find "a1j02 = '" & Text2 & "'", 0, adSearchForward, adoacc1j0.Bookmark
      If adoacc1j0.EOF = False Then
         FormShow
         RecordShow
      Else
         MsgBox MsgText(33), , MsgText(5)
         adoacc1j0.MoveFirst
      End If
   Else
      MsgBox MsgText(33), , MsgText(5)
      adoacc1j0.MoveFirst
   End If
End Sub

Private Sub Command5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command5_Click
         Exit Sub
   End Select
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Form_Activate()
   '93.3.16 ADD BY SONIA
   If IsObject(mdiMain) Then
      ToolShow
   End If
   '93.3.16 END
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc1j0.RecordCount <> 0 Then
      adoacc1j0.MoveFirst
   End If
   adoacc1j0.Find "a1j01 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoacc1j0.EOF = False Then
      adoacc1j0.Find "a1j02 = '" & strCustNo & "'", 0, adSearchForward, adoacc1j0.Bookmark
      If adoacc1j0.EOF = False Then
         FormShow
         RecordShow
      End If
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   
'   Dim intX As Integer
'   Dim intY As Integer
'   Dim sglWidth As Single
'   Dim sglHeight As Single
'
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   'Me.Height = 3800
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   
   'Add By Cheng 2002/02/04
   FilterSystem
   
   OpenTable
   If adoacc1j0.RecordCount <> 0 Then
      adoacc1j0.MoveLast
      adoacc1j0.MoveFirst
      RecordShow
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc21g0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc1j0.CursorLocation = adUseClient
   'Modify By Cheng 2002/02/04
'   adoacc1j0.Open "select * from acc1j0 order by a1j01 asc, a1j02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1j0.Open "select * from acc1j0 Where A1J01 IN " & m_QuerySystem & " order by a1j01 asc, a1j02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國外請款項目資料)
'
'*************************************************
Public Sub FormShow()
   Text1 = adoacc1j0.Fields("a1j01").Value
   Text2 = adoacc1j0.Fields("a1j02").Value
   If IsNull(adoacc1j0.Fields("a1j03").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoacc1j0.Fields("a1j03").Value
   End If
   If IsNull(adoacc1j0.Fields("a1j04").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc1j0.Fields("a1j04").Value
   End If
   If IsNull(adoacc1j0.Fields("a1j05").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc1j0.Fields("a1j05").Value
   End If
   If IsNull(adoacc1j0.Fields("a1j06").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc1j0.Fields("a1j06").Value
   End If
   If IsNull(adoacc1j0.Fields("a1j16").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc1j0.Fields("a1j16").Value
   End If
   If IsNull(adoacc1j0.Fields("a1j07").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = adoacc1j0.Fields("a1j07").Value
   End If
   If IsNull(adoacc1j0.Fields("a1j08").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc1j0.Fields("a1j08").Value
   End If
   If IsNull(adoacc1j0.Fields("a1j09").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc1j0.Fields("a1j09").Value
   End If
   If IsNull(adoacc1j0.Fields("a1j17").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc1j0.Fields("a1j17").Value
   End If
   'Add by Morgan 2010/11/4
   Text13 = "" & adoacc1j0.Fields("a1j18").Value
   Text14 = "" & adoacc1j0.Fields("a1j19").Value
   Text15 = "" & adoacc1j0.Fields("a1j20").Value
   'end 2010/11/4
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
   CloseIme
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
   CloseIme
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
   CloseIme
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text3_GotFocus()
   StatusView MsgText(65) & "100"
   TextInverse Text3
End Sub

Private Sub Text3_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text3_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If CheckLen(Label3, Text3, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

Private Sub Text9_Change()
   If Text9 <> MsgText(601) Then
      Text12 = A0102Query(Text9)
   Else
      Text12 = MsgText(601)
   End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter CInt(KeyCode)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc1j0.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow adoacc1j0.Bookmark, adoacc1j0.RecordCount
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc010", "a0101", Text9, Label10) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub
'Add By Cheng 2002/02/04
'取得使用者系統類別使用權限
Private Sub FilterSystem()
   Dim nIndex As Integer
   Dim nCount As Integer
   Dim strSys As String
   Dim strTemp As String
   m_QuerySystem = Empty
   
   strSys = GetSystemKindByNick
   nCount = GetSubStringCount(strSys)
   For nIndex = 1 To nCount
      strTemp = GetSubString(strSys, nIndex)
      If IsEmptyText(m_QuerySystem) = False Then m_QuerySystem = m_QuerySystem & ","
      m_QuerySystem = m_QuerySystem & "'" & strTemp & "'"
NextRecord:
   Next nIndex
   
   m_QuerySystem = "(" & m_QuerySystem & ")"

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得字串以逗點分隔的Sub字串總數
' Input : strTemp ==> 所要計算子字串個數的母字串
' Output : 傳回以逗號分隔的子字串總數
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSubStringCount(ByVal strTemp As String) As Integer
   Dim nLength As Integer
   Dim nPos As Integer
   Dim nSection As Integer
   nSection = 0
   nLength = Len(strTemp)
   For nPos = 1 To nLength
      If Mid(strTemp, nPos, 1) = "," Then
         If nPos <> nLength Then
            nSection = nSection + 1
         End If
      End If
   Next nPos
   If nSection > 0 Then
      nSection = nSection + 1
   Else
      If IsEmptyText(strTemp) = False Then
         nSection = 1
      End If
   End If
   GetSubStringCount = nSection
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得字串中的子字串
' Input : strTemp ==> 原始字串
'         nSec ==> 子字串的索引值 (以1為基底)
' Output : 傳回第nSec個子字串
' 說明 : 若傳入的nSec值超過子字串的總數時, 回傳回空字串
'        此功能需配合函式(GetSubStringCount)來使用
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSubString(ByVal strTemp As String, ByVal nSec As Integer)
   Dim strSub As String
   Dim nLength As Integer
   Dim nPos As Integer
   Dim nX As Integer
   Dim nSection As Integer
   nSection = 1
   nLength = Len(strTemp)
   For nPos = 1 To nLength
      If nSection = nSec Then
         strSub = Empty
         For nX = nPos To nLength
            If Mid(strTemp, nX, 1) = "," Then
               Exit For
            Else
               strSub = strSub & Mid(strTemp, nX, 1)
            End If
         Next nX
         Exit For
      End If
      If Mid(strTemp, nPos, 1) = "," Then
         nSection = nSection + 1
      End If
   Next nPos
   strSub = Trim(strSub)
   GetSubString = strSub
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查字串是否為空白
' Input : strData ==> 所要檢查的字串
' Output : 若輸入的字串中沒有非空白的字元存在時, 則傳回True
'          否則傳回 False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEmptyText(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmptyText = False
   
   If Len(strData) <= 0 Then
      IsEmptyText = True
   Else
      IsEmptyText = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmptyText = False
            Exit For
         End If
      Next nIndex
   End If
End Function

