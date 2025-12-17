VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc34e0 
   AutoRedraw      =   -1  'True
   Caption         =   "銀行調節資料表"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   5160
   Begin VB.TextBox Text3 
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
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   1860
      Width           =   4692
   End
   Begin VB.ComboBox Combo12 
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
      TabIndex        =   13
      Top             =   4860
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo11 
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
      Left            =   960
      TabIndex        =   12
      Top             =   4860
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo10 
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
      TabIndex        =   11
      Top             =   4500
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo9 
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
      Left            =   960
      TabIndex        =   10
      Top             =   4500
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo8 
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
      TabIndex        =   9
      Top             =   4140
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo7 
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
      Left            =   960
      TabIndex        =   8
      Top             =   4140
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo6 
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
      TabIndex        =   7
      Top             =   3780
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   960
      TabIndex        =   6
      Top             =   3780
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
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
      TabIndex        =   5
      Top             =   3420
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   960
      TabIndex        =   4
      Top             =   3420
      Visible         =   0   'False
      Width           =   1812
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
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
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
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "(銀行對帳單餘額)"
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
      Left            =   3000
      TabIndex        =   26
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   25
      Top             =   630
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "上期餘額"
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
      Top             =   1380
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2415
      Left            =   240
      Top             =   2880
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc34e0.frx":0000
      Stretch         =   -1  'True
      Top             =   4860
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "5."
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
      Left            =   720
      TabIndex        =   23
      Top             =   4860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc34e0.frx":0442
      Stretch         =   -1  'True
      Top             =   4500
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "4."
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
      Left            =   720
      TabIndex        =   22
      Top             =   4500
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc34e0.frx":0884
      Stretch         =   -1  'True
      Top             =   4140
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label10 
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
      Left            =   720
      TabIndex        =   21
      Top             =   4140
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc34e0.frx":0CC6
      Stretch         =   -1  'True
      Top             =   3780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
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
      Left            =   720
      TabIndex        =   20
      Top             =   3780
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc34e0.frx":1108
      Stretch         =   -1  'True
      Top             =   3420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
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
      Left            =   720
      TabIndex        =   19
      Top             =   3420
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
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
      Top             =   3060
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   5340
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
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
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   990
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
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
      TabIndex        =   15
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc34e0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoaccrpt315 As New ADODB.Recordset
Public adoacc0h0 As New ADODB.Recordset
Public adoacc0b0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSort1, strSort2, strSort3, strSort4, strSort5 As String
Dim dllaccrpt315 As Object
Dim lngBalance As Long, lngBankBalance As Long
Dim strSql As String
Dim douAmount As Double

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
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt315Delete
   ProduceData
'   SumShow
   If adoaccrpt315.State = adStateOpen Then
      adoaccrpt315.Close
   End If
   adoaccrpt315.CursorLocation = adUseClient
   adoaccrpt315.Open "select * from accrpt315", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt315.RecordCount <> 0 Then
      'Modify By Cheng 2002/03/29
'      dllaccrpt315.Acc34e0 ReportTitle(315), Text1, Text2, MaskEdBox1.Text, MaskEdBox2.Text, Format(Val(Text3), DDollar), Format(Val(Text3) - douAmount, DDollar), StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      dllaccrpt315.Acc34e0 ReportTitle(315), Text1, Me.lbl(0).Caption, MaskEdBox1.Text, MaskEdBox2.Text, Format(Val(Text3), DDollar), Format(Val(Text3) - douAmount, DDollar), StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   End If
   adoaccrpt315.Close
   Screen.MousePointer = vbDefault
   FormClear
   'Modify By Cheng 2002/01/30
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      'Modify By Cheng 2002/01/30
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
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
'   Me.Height = 2400
   Me.Height = 2820
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
   'Modify By Cheng 2002/01/30
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Set dllaccrpt315 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt315 = Nothing
   Set Frmacc34e0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'Private Sub Text2_GotFocus()
'   TextInverse Text2
'End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "銀行代號"
   strSort2 = "票據號碼"
   strSort3 = "票別"
   strSort4 = "調節日期"
   strSort5 = "開票日期"
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
   
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e01 asc"
         Else
            strOrder1 = " order by a0e01 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e02 asc"
         Else
            strOrder1 = " order by a0e02 desc"
         End If
      Case strSort3
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e04 asc"
         Else
            strOrder1 = " order by a0e04 desc"
         End If
      Case strSort4
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e22 asc"
         Else
            strOrder1 = " order by a0e22 desc"
         End If
      Case strSort5
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e13 asc"
         Else
            strOrder1 = " order by a0e13 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e01 asc"
         Else
            strOrder2 = ", a0e01 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e02 asc"
         Else
            strOrder2 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e04 asc"
         Else
            strOrder2 = ", a0e04 desc"
         End If
      Case strSort4
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e22 asc"
         Else
            strOrder2 = ", a0e22 desc"
         End If
      Case strSort5
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e13 asc"
         Else
            strOrder2 = ", a0e13 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo7
      Case strSort1
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e01 asc"
         Else
            strOrder3 = ", a0e01 desc"
         End If
      Case strSort2
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e02 asc"
         Else
            strOrder3 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e04 asc"
         Else
            strOrder3 = ", a0e04 desc"
         End If
      Case strSort4
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e22 asc"
         Else
            strOrder3 = ", a0e22 desc"
         End If
      Case strSort5
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e13 asc"
         Else
            strOrder3 = ", a0e13 desc"
         End If
      Case Else
         strOrder3 = MsgText(601)
   End Select
   Select Case Combo9
      Case strSort1
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e01 asc"
         Else
            strOrder4 = ", a0e01 desc"
         End If
      Case strSort2
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e02 asc"
         Else
            strOrder4 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e04 asc"
         Else
            strOrder4 = ", a0e04 desc"
         End If
      Case strSort4
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e22 asc"
         Else
            strOrder4 = ", a0e22 desc"
         End If
      Case strSort5
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e13 asc"
         Else
            strOrder4 = ", a0e13 desc"
         End If
      Case Else
         strOrder4 = MsgText(601)
   End Select
   Select Case Combo11
      Case strSort1
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e01 asc"
         Else
            strOrder5 = ", a0e01 desc"
         End If
      Case strSort2
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e02 asc"
         Else
            strOrder5 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e04 asc"
         Else
            strOrder5 = ", a0e04 desc"
         End If
      Case strSort4
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e22 asc"
         Else
            strOrder5 = ", a0e22 desc"
         End If
      Case strSort5
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e13 asc"
         Else
            strOrder5 = ", a0e13 desc"
         End If
      Case Else
         strOrder5 = MsgText(601)
   End Select
   'Add By Cheng 2002/03/29
   strSql = ""
   
   If Text1 <> MsgText(601) Then
      strSql = " and a0e07 = '" & Text1 & "'"
   End If
   'Modify By Cheng 2002/03/29
   '銀行帳號改成單選
'   If Text2 <> MsgText(601) Then
'      strSQL = strSQL & " and a0e07 <= '" & Text2 & "'"
'   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Me.adoacc0e0.State <> adStateClosed Then Me.adoacc0e0.Close
   Set Me.adoacc0e0 = Nothing
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select sum(a0e11) from acc0e0 where a0e15 = 0 and a0e25 = 0 and (a0e37 is not null and a0e37 <> 0) and (a0e22 = 0 or a0e22 is null) AND a0e04 = '" & MsgText(19) & "'" & strSql & strOrder1 & strOrder2 & strOrder3 & strOrder4 & strOrder5, _
                  adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount <> 0 Then
      If IsNull(adoacc0e0.Fields(0).Value) Then
         douAmount = 0
      Else
         douAmount = adoacc0e0.Fields(0).Value
      End If
   Else
      douAmount = 0
   End If
   adoacc0e0.Close
   adoaccrpt315.CursorLocation = adUseClient
   adoaccrpt315.Open "select * from accrpt315", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acc0e0 where a0e15 = 0 and a0e25 = 0 and (a0e37 is not null and a0e37 <> 0) and (a0e22 = 0 or a0e22 is null) AND a0e04 = '" & MsgText(19) & "'" & strSql & strOrder1 & strOrder2 & strOrder3 & strOrder4 & strOrder5, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      adoacc0e0.Close
      adoaccrpt315.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0e0.EOF = False
      adoaccrpt315.AddNew
      adoaccrpt315.Fields("r31501").Value = strUserNum
      adoaccrpt315.Fields("r31502").Value = adoacc0e0.Fields("a0e01").Value
      If IsNull(adoacc0e0.Fields("a0e07").Value) Then
         adoaccrpt315.Fields("r31503").Value = Null
      Else
         adoaccrpt315.Fields("r31503").Value = adoacc0e0.Fields("a0e07").Value
         adoaccrpt315.Fields("r31504").Value = A0g02Query(adoacc0e0.Fields("a0e01").Value)
      End If
      adoaccrpt315.Fields("r31505").Value = adoacc0e0.Fields("a0e02").Value
      Select Case adoacc0e0.Fields("a0e08").Value
         Case Mid(ComboItem(11), 1, 1)
            adoaccrpt315.Fields("r31506").Value = Mid(ComboItem(11), 4, 2)
         Case Mid(ComboItem(12), 1, 1)
            adoaccrpt315.Fields("r31506").Value = Mid(ComboItem(12), 4, 2)
         Case Mid(ComboItem(13), 1, 1)
            adoaccrpt315.Fields("r31506").Value = Mid(ComboItem(13), 4, 2)
         Case Else
            adoaccrpt315.Fields("r31506").Value = Null
      End Select
      If IsNull(adoacc0e0.Fields("a0e22").Value) Then
         adoaccrpt315.Fields("r31507").Value = Null
      Else
         adoaccrpt315.Fields("r31507").Value = adoacc0e0.Fields("a0e22").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e11").Value) Then
         adoaccrpt315.Fields("r31508").Value = 0
      Else
         adoaccrpt315.Fields("r31508").Value = Val(adoacc0e0.Fields("a0e11").Value)
      End If
      If IsNull(adoacc0e0.Fields("a0e13").Value) Then
         adoaccrpt315.Fields("r31509").Value = Null
      Else
         adoaccrpt315.Fields("r31509").Value = adoacc0e0.Fields("a0e13").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e10").Value) Then
         adoaccrpt315.Fields("r31510").Value = Null
      Else
         adoaccrpt315.Fields("r31510").Value = adoacc0e0.Fields("a0e10").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e06").Value) Then
         adoaccrpt315.Fields("r31511").Value = Null
      Else
         Select Case adoacc0e0.Fields("A0E05").Value
            Case "1"
               'Modify By Cheng 2002/01/30
'               adoaccrpt315.Fields("r31511").Value = adoacc0e0.Fields("a0e06").Value & CustomerQuery(adoacc0e0.Fields("A0E06").Value, 1)
               adoaccrpt315.Fields("r31511").Value = CustomerQuery(adoacc0e0.Fields("A0E06").Value, 1)
            Case "2"
               'Modify By Cheng 2002/01/30
'               adoaccrpt315.Fields("r31511").Value = adoacc0e0.Fields("a0e06").Value & A0i02Query(adoacc0e0.Fields("A0E06").Value)
               adoaccrpt315.Fields("r31511").Value = A0i02Query(adoacc0e0.Fields("A0E06").Value)
            Case "3"
               'Modify By Cheng 2002/01/30
'               adoaccrpt315.Fields("r31511").Value = adoacc0e0.Fields("a0e06").Value & StaffQuery(adoacc0e0.Fields("A0E06").Value)
               adoaccrpt315.Fields("r31511").Value = StaffQuery(adoacc0e0.Fields("A0E06").Value)
         End Select
      End If
      
      'Add By Cheng 2002/01/30
      '備註欄
      If IsNull(adoacc0e0.Fields("a0e12").Value) Then
         adoaccrpt315.Fields("r31512").Value = Null
      Else
         adoaccrpt315.Fields("r31512").Value = adoacc0e0.Fields("a0e12").Value
      End If
      
      adoaccrpt315.UpdateBatch
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   adoaccrpt315.Close
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
Private Sub Accrpt315Delete()
   adoTaie.Execute "delete from accrpt315"
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Me.lbl(0).Caption = ""
'   Text2 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Combo3 = ""
   Combo5 = ""
   Combo7 = ""
   Combo9 = ""
   Combo11 = ""
   Text1.SetFocus
   'Add By Cheng 2002/01/30
   Me.Text3.Text = ""
End Sub

'*************************************************
'  合計
'
'*************************************************
Public Sub SumShow()
Dim strSQL1 As String
Dim strYear As String
Dim strMonth As String

   If Text1 <> MsgText(601) Then
      'Modify By Cheng 2002/03/29
'      strSQL1 = strSQL1 & " and a0h02 >= '" & Text1 & "'"
      strSQL1 = strSQL1 & " and a0h02 = '" & Text1 & "'"
   End If
   'Modify By Cheng 2002/03/29
'   If Text2 <> MsgText(601) Then
'      strSQL1 = strSQL1 & " and a0h02 <= '" & Text2 & "'"
'   End If
   adoacc0b0.CursorLocation = adUseClient
   adoacc0b0.Open "select a0b02 from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
   If IsNull(adoacc0b0.Fields(0).Value) Then
      GoTo Checking
   Else
      strYear = Mid(CFDate(adoacc0b0.Fields(0).Value), 1, 3)
      strMonth = Mid(CFDate(adoacc0b0.Fields(0).Value), 5, 2)
   End If
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select sum(a0408) from acc040, acc0h0 where a0405 = a0h08 and a0401 = " & Val(strYear) & " and a0402 = " & Val(strMonth) & " and a0403 = '1' and a0404 = '" & MsgText(55) & "'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0h0.RecordCount <> 0 Then
      If IsNull(adoacc0h0.Fields(0).Value) = False Then
         lngBalance = adoacc0h0.Fields(0).Value
      Else
         lngBalance = 0
      End If
   Else
      lngBalance = 0
   End If
   adoacc0h0.Close
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select SUM(A0E11) from acc0e0 where a0e15 = 0 and a0e25 = 0 AND A0E22 = 0 and a0e37 = 0 AND a0e04 = '" & MsgText(19) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount <> 0 Then
      If IsNull(adoacc0e0.Fields(0).Value) Then
         lngBankBalance = lngBalance
      Else
         lngBankBalance = lngBalance - adoacc0e0.Fields(0).Value
      End If
   Else
      lngBankBalance = lngBalance
   End If
   adoacc0e0.Close
   adoacc0b0.Close
   Exit Sub
Checking:
   adoacc0b0.Close
   lngBalance = 0
   lngBankBalance = 0
End Sub

Private Sub Text1_LostFocus()
'Add By Cheng 2002/03/29
GetBankName Me.Text1.Text
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub
'取得銀行名稱
Private Sub GetBankName(strBankNo As String)
Dim rs As New ADODB.Recordset
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
rs.CursorLocation = adUseClient
rs.Open "Select A0H03 From ACC0H0 WHERE A0H02='" & strBankNo & "'", adoTaie, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
   Me.lbl(0).Caption = "" & rs.Fields(0).Value
Else
   Me.lbl(0).Caption = ""
End If
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
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
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

