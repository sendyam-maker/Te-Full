VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc3450 
   AutoRedraw      =   -1  'True
   Caption         =   "銀行帳號別票據明細表"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   5160
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   1575
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
      Left            =   3240
      TabIndex        =   1
      Top             =   240
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
      TabIndex        =   0
      Top             =   240
      Width           =   1572
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
      TabIndex        =   6
      Top             =   2640
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
      TabIndex        =   7
      Top             =   2640
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
      TabIndex        =   8
      Top             =   3000
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
      TabIndex        =   9
      Top             =   3000
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
      TabIndex        =   10
      Top             =   3360
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
      TabIndex        =   11
      Top             =   3360
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
      TabIndex        =   12
      Top             =   3720
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
      TabIndex        =   13
      Top             =   3720
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
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   1812
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
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1212
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
      TabIndex        =   5
      Top             =   1440
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   600
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
      TabIndex        =   3
      Top             =   600
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
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "應收/付"
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
      TabIndex        =   26
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
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
      TabIndex        =   25
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
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
      TabIndex        =   24
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銀行代號"
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
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
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
      TabIndex        =   22
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
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
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3450.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
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
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3450.frx":0442
      Stretch         =   -1  'True
      Top             =   3000
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
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3450.frx":0884
      Stretch         =   -1  'True
      Top             =   3360
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
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3450.frx":0CC6
      Stretch         =   -1  'True
      Top             =   3720
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
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3450.frx":1108
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2415
      Left            =   240
      Top             =   2160
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "Frmacc3450"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoaccrpt305 As New ADODB.Recordset
Public adoacc0h0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSort1 As String
Dim strSort2 As String
Dim strSort3 As String
Dim strSort4 As String
Dim strSort5 As String
Dim strBankNo As String
Dim strBankId As String
Dim strIdName As String
Dim lngDeposit As Long
Dim lngRecAmount As Long
Dim lngPayAmount As Long
Dim intCounter As Integer
Dim strDateRange As String
Dim dllaccrpt305 As Object

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
Dim strQuery As String

   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   intCounter = 0
   Screen.MousePointer = vbHourglass
   Accrpt305Delete
   ProduceData
   If Text1 <> "" Or Text2 <> "" Then
      strQuery = " WHERE A0H01 >= '" & Text1 & "' AND A0H01 <= '" & Text2 & "'"
   End If
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select * from acc0h0" & strQuery & " order by a0h01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0h0.EOF = False
      adoaccrpt305.CursorLocation = adUseClient
      adoaccrpt305.Open "select * from accrpt305 where r30510 = '" & adoacc0h0.Fields("a0h01").Value & "' and r30511 = '" & adoacc0h0.Fields("a0h02").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoaccrpt305.RecordCount <> 0 Then
         Acc0h0Show
         RunReportDll
      End If
      adoaccrpt305.Close
      intCounter = intCounter + 1
      adoacc0h0.MoveNext
   Loop
   adoacc0h0.Close
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
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
   Me.Height = 2500
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
   Combo1.AddItem ComboItem(181)
   Combo1.AddItem ComboItem(182)
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
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt305 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt305 = Nothing
   Set Frmacc3450 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "票據號碼"
   strSort2 = "票別"
   strSort3 = "開票日期"
   strSort4 = "到期日期"
   strSort5 = "往來對象"
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
   strDateRange = ""
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e02 asc"
         Else
            strOrder1 = " order by a0e02 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e08 asc"
         Else
            strOrder1 = " order by a0e08 desc"
         End If
      Case strSort3
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e13 asc"
         Else
            strOrder1 = " order by a0e13 desc"
         End If
      Case strSort4
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e10 asc"
         Else
            strOrder1 = " order by a0e10 desc"
         End If
      Case strSort5
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e06 asc"
         Else
            strOrder1 = " order by a0e06 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e02 asc"
         Else
            strOrder2 = ", a0e02 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e08 asc"
         Else
            strOrder2 = ", a0e08 desc"
         End If
      Case strSort3
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e13 asc"
         Else
            strOrder2 = ", a0e13 desc"
         End If
      Case strSort4
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e10 asc"
         Else
            strOrder2 = ", a0e10 desc"
         End If
      Case strSort5
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e06 asc"
         Else
            strOrder2 = ", a0e06 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo7
      Case strSort1
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e02 asc"
         Else
            strOrder3 = ", a0e02 desc"
         End If
      Case strSort2
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e08 asc"
         Else
            strOrder3 = ", a0e08 desc"
         End If
      Case strSort3
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e13 asc"
         Else
            strOrder3 = ", a0e13 desc"
         End If
      Case strSort4
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e10 asc"
         Else
            strOrder3 = ", a0e10 desc"
         End If
      Case strSort5
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e06 asc"
         Else
            strOrder3 = ", a0e06 desc"
         End If
      Case Else
         strOrder3 = MsgText(601)
   End Select
   Select Case Combo9
      Case strSort1
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e02 asc"
         Else
            strOrder4 = ", a0e02 desc"
         End If
      Case strSort2
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e08 asc"
         Else
            strOrder4 = ", a0e08 desc"
         End If
      Case strSort3
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e13 asc"
         Else
            strOrder4 = ", a0e13 desc"
         End If
      Case strSort4
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e10 asc"
         Else
            strOrder4 = ", a0e10 desc"
         End If
      Case strSort5
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e06 asc"
         Else
            strOrder4 = ", a0e06 desc"
         End If
      Case Else
         strOrder4 = MsgText(601)
   End Select
   Select Case Combo11
      Case strSort1
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e02 asc"
         Else
            strOrder5 = ", a0e02 desc"
         End If
      Case strSort2
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e08 asc"
         Else
            strOrder5 = ", a0e08 desc"
         End If
      Case strSort3
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e13 asc"
         Else
            strOrder5 = ", a0e13 desc"
         End If
      Case strSort4
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e10 asc"
         Else
            strOrder5 = ", a0e10 desc"
         End If
      Case strSort5
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e06 asc"
         Else
            strOrder5 = ", a0e06 desc"
         End If
      Case Else
         strOrder5 = MsgText(601)
   End Select
   If Text1 <> MsgText(601) Then
      strSql = " and ((A0E04 = '" & MsgText(19) & "' AND a0e01 >= '" & Text1 & "' and a0e25 = 0 and a0e37 = 0) OR (A0E04 = '" & MsgText(18) & "' AND A0E19 >= '" & Text1 & "' and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e34 = 0))"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and ((A0E04 = '" & MsgText(19) & "' AND a0e01 <= '" & Text2 & "' and a0e25 = 0 and a0e37 = 0) OR (A0E04 = '" & MsgText(18) & "' AND A0E19 <= '" & Text2 & "' and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e34 = 0))"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strDateRange = strDateRange & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strDateRange = strDateRange & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Combo1 <> MsgText(601) Then
      strSql = strSql & " and a0e04 = '" & Mid(Combo1, 1, 1) & "'"
      strDateRange = strDateRange & " and a0e04 = '" & Mid(Combo1, 1, 1) & "'"
   End If
   adoaccrpt305.CursorLocation = adUseClient
   adoaccrpt305.Open "select * from accrpt305", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acc0e0 where a0e04 in ('" & MsgText(18) & "', '" & MsgText(19) & "')" & strSql & strOrder1 & strOrder2 & strOrder3 & strOrder4 & strOrder5, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      adoacc0e0.Close
      adoaccrpt305.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0e0.EOF = False
      adoaccrpt305.AddNew
      adoaccrpt305.Fields("r30501").Value = strUserNum
      adoaccrpt305.Fields("r30502").Value = adoacc0e0.Fields("a0e02").Value
      If IsNull(adoacc0e0.Fields("a0e08").Value) Then
         adoaccrpt305.Fields("r30503").Value = Null
      Else
         Select Case adoacc0e0.Fields("a0e08").Value
            Case Mid(ComboItem(11), 1, 1)
               adoaccrpt305.Fields("r30503").Value = Mid(ComboItem(11), 4, 2)
            Case Mid(ComboItem(12), 1, 1)
               adoaccrpt305.Fields("r30503").Value = Mid(ComboItem(12), 4, 2)
            Case Mid(ComboItem(13), 1, 1)
               adoaccrpt305.Fields("r30503").Value = Mid(ComboItem(13), 4, 2)
         End Select
      End If
      If IsNull(adoacc0e0.Fields("a0e11").Value) Then
         adoaccrpt305.Fields("r30504").Value = 0
         adoaccrpt305.Fields("r30505").Value = 0
      Else
         Select Case adoacc0e0.Fields("a0e04").Value
            Case MsgText(18)
               adoaccrpt305.Fields("r30504").Value = Val(adoacc0e0.Fields("a0e11").Value)
               adoaccrpt305.Fields("r30505").Value = 0
               If IsNull(adoacc0e0.Fields("a0e19").Value) Then
                  adoaccrpt305.Fields("r30510").Value = Null
               Else
                  adoaccrpt305.Fields("r30510").Value = adoacc0e0.Fields("a0e19").Value
               End If
               If IsNull(adoacc0e0.Fields("a0e20").Value) Then
                  adoaccrpt305.Fields("r30511").Value = Null
               Else
                  adoaccrpt305.Fields("r30511").Value = adoacc0e0.Fields("a0e20").Value
               End If
            Case MsgText(19)
               adoaccrpt305.Fields("r30504").Value = 0
               adoaccrpt305.Fields("r30505").Value = Val(adoacc0e0.Fields("a0e11").Value)
               adoaccrpt305.Fields("r30510").Value = adoacc0e0.Fields("a0e01").Value
               If IsNull(adoacc0e0.Fields("a0e07").Value) Then
                  adoaccrpt305.Fields("r30511").Value = Null
               Else
                  adoaccrpt305.Fields("r30511").Value = adoacc0e0.Fields("a0e07").Value
               End If
            Case Else
               adoaccrpt305.Fields("r30504").Value = 0
               adoaccrpt305.Fields("r30505").Value = 0
               adoaccrpt305.Fields("r30510").Value = Null
               adoaccrpt305.Fields("r30511").Value = Null
         End Select
      End If
      If IsNull(adoacc0e0.Fields("a0e13").Value) Then
         adoaccrpt305.Fields("r30506").Value = Null
      Else
         adoaccrpt305.Fields("r30506").Value = adoacc0e0.Fields("a0e13").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e10").Value) Then
         adoaccrpt305.Fields("r30507").Value = Null
      Else
         adoaccrpt305.Fields("r30507").Value = adoacc0e0.Fields("a0e10").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e06").Value) Then
         adoaccrpt305.Fields("r30509").Value = Null
      Else
         Select Case adoacc0e0.Fields("A0E05").Value
            Case "1"
               adoaccrpt305.Fields("r30509").Value = CustomerQuery(adoacc0e0.Fields("A0E06").Value, 1)
            Case "2"
               adoaccrpt305.Fields("r30509").Value = A0i02Query(adoacc0e0.Fields("A0E06").Value)
            Case "3"
               adoaccrpt305.Fields("r30509").Value = StaffQuery(adoacc0e0.Fields("A0E06").Value)
         End Select
      End If
      If IsNull(adoacc0e0.Fields("A0E05").Value) = False Then
         adoaccrpt305.Fields("R30512").Value = adoacc0e0.Fields("A0E05").Value
      End If
      adoaccrpt305.UpdateBatch
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   adoaccrpt305.Close
   adoTaie.Execute "delete from accrpt305 where r30502 is null"
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
Private Sub Accrpt305Delete()
   adoTaie.Execute "delete from accrpt305"
End Sub

'*************************************************
'  顯示銀行帳戶餘額資料
'
'*************************************************
Private Sub Acc0h0Show()
Dim adoaccsum As New ADODB.Recordset
Dim adoacc040 As New ADODB.Recordset
Dim adoacc0b0 As New ADODB.Recordset
Dim intYear, intMonth As Integer

   adoacc0b0.CursorLocation = adUseClient
   adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0b0.RecordCount = 0 Then
      If Mid(ServerDate, 5, 2) = 1 Then
         intMonth = 12
         intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
      Else
         intMonth = Val(Mid(ServerDate, 5, 2)) - 1
         intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
      End If
   Else
      If IsNull(adoacc0b0.Fields("a0b01").Value) Then
         If Mid(ServerDate, 5, 2) = 1 Then
            intMonth = 12
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
         Else
            intMonth = Val(Mid(ServerDate, 5, 2)) - 1
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
         End If
      Else
        intMonth = Val(Mid(CFDate(adoacc0b0.Fields("a0b01").Value), 5, 2))
        intYear = Val(Mid(CFDate(adoacc0b0.Fields("a0b01").Value), 1, 3))
      End If
   End If
   adoacc0b0.Close
   adoacc040.CursorLocation = adUseClient
   adoacc040.Open "select a0408 from acc040 where a0401 = " & intYear & " and a0403 = '1' and a0404 = '" & MsgText(55) & "' and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0402 = " & intMonth & "", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc040.RecordCount <> 0 Then
      If IsNull(adoacc040.Fields(0).Value) Then
         lngDeposit = 0
      Else
         lngDeposit = adoacc040.Fields(0).Value
      End If
   Else
      lngDeposit = 0
   End If
   adoacc040.Close
   strBankNo = adoacc0h0.Fields("a0h01").Value
   strBankId = adoacc0h0.Fields("a0h02").Value
   If IsNull(adoacc0h0.Fields("a0h03").Value) Then
      strIdName = MsgText(601)
   Else
      strIdName = adoacc0h0.Fields("a0h03").Value
   End If
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e19 = '" & adoacc0h0.Fields("a0h01").Value & "' AND A0E20 = '" & adoacc0h0.Fields("A0H02").Value & "' and a0e04 = '" & MsgText(18) & "' and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and (a0e34 = 0 or a0e34 is null)" & strDateRange, adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         lngRecAmount = 0
      Else
         lngRecAmount = Val(adoaccsum.Fields(0).Value)
      End If
   Else
      lngRecAmount = 0
   End If
   adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e01 = '" & adoacc0h0.Fields("a0h01").Value & "' AND A0E07 = '" & adoacc0h0.Fields("A0H02").Value & "' and a0e04 = '" & MsgText(19) & "' and a0e25 = 0 and (a0e37 = 0 or a0e37 is null)" & strDateRange, adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         lngPayAmount = 0
      Else
         lngPayAmount = Val(adoaccsum.Fields(0).Value)
      End If
   Else
      lngPayAmount = 0
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  執行報表之 Dll
'
'*************************************************
Private Sub RunReportDll()
   dllaccrpt305.Acc3450 ReportTitle(305), Text1, Text2, strBankNo, strBankId, strIdName, Format((lngDeposit), DDollar), Format(lngRecAmount, DDollar), Format(lngPayAmount, DDollar), Combo1.Text, MaskEdBox1.Text, MaskEdBox2.Text, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
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
   Combo1 = ""
   Combo3 = ""
   Combo5 = ""
   Combo7 = ""
   Combo9 = ""
   Combo11 = ""
   Text1.SetFocus
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
   FormCheck = False
End Function

