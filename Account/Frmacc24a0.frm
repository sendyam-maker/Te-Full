VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24a0 
   AutoRedraw      =   -1  'True
   Caption         =   "國外帳齡分析表"
   ClientHeight    =   3564
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3564
   ScaleWidth      =   5700
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel檔(&E)"
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
      Left            =   1740
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1530
      Style           =   2  '單純下拉式
      TabIndex        =   14
      Top             =   3090
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.CommandButton Command2 
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
      Left            =   180
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1530
      TabIndex        =   12
      Top             =   1770
      Width           =   612
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4530
      TabIndex        =   9
      Top             =   1050
      Width           =   612
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3930
      TabIndex        =   8
      Top             =   1050
      Width           =   612
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3330
      TabIndex        =   7
      Top             =   1050
      Width           =   612
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   6
      Top             =   1050
      Width           =   612
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2130
      TabIndex        =   5
      Top             =   1050
      Width           =   612
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1530
      TabIndex        =   4
      Top             =   1050
      Width           =   612
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3450
      MaxLength       =   9
      TabIndex        =   3
      Top             =   690
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1530
      MaxLength       =   9
      TabIndex        =   2
      Top             =   690
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   1
      Top             =   330
      Width           =   852
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1530
      TabIndex        =   0
      Top             =   330
      Width           =   852
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1530
      TabIndex        =   10
      Top             =   1410
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
      Left            =   3450
      TabIndex        =   11
      Top             =   1410
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
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "大陸一定要輸！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3660
      TabIndex        =   27
      Top             =   360
      Width           =   1785
   End
   Begin VB.Label Label11 
      Caption         =   "註：A4紙張"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   510
      TabIndex        =   25
      Top             =   2790
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   510
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   270
      Top             =   2730
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "(1.細目 2.總計)"
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
      Left            =   2250
      TabIndex        =   23
      Top             =   1770
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "報表格式"
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
      Left            =   570
      TabIndex        =   22
      Top             =   1770
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3210
      TabIndex        =   21
      Top             =   1410
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "請款日期"
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
      Left            =   570
      TabIndex        =   20
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "系統類別"
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
      Left            =   570
      TabIndex        =   19
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   3210
      TabIndex        =   18
      Top             =   690
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
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
      Left            =   570
      TabIndex        =   17
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Left            =   2490
      TabIndex        =   16
      Top             =   330
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "國籍"
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
      Left            =   570
      TabIndex        =   15
      Top             =   330
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc24a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/09 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt210 As New ADODB.Recordset
'Dim dllaccrpt210 As Object
'Add By Sindy 2012/12/14
Dim iLine As Integer
Dim strTemp(0 To 11) As String 'Modify by Amy 2025/09/17 +未收規費NTD,原strTemp(0 To 10)
Dim PLeft(0 To 10) As Integer
Dim strPrinter As String
Dim strNation As String
Dim StrFa As String
'2012/12/14 End
Dim iPrintLine As Integer
'Add By Sindy 2014/6/5
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
Dim intPage As Integer
Dim dblSkipPageRow As Double
'2014/6/5 END


'產生Excel檔
Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt210Delete
   ProduceData
   PrintExcel
   FormClear
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = "" 'MsgText(102)
End Sub

''列印
'Private Sub Command2_Click()
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'   Screen.MousePointer = vbHourglass
'   Accrpt210Delete
'   ProduceData
'   'Modify By Sindy 2012/12/14 Mark改報表寫法, 不使用dllaccrpt210
''   If adoaccrpt210.State = adStateOpen Then
''      adoaccrpt210.Close
''   End If
''   adoaccrpt210.CursorLocation = adUseClient
''   adoaccrpt210.Open "select * from accrpt210", adoTaie, adOpenStatic, adLockReadOnly
''   If adoaccrpt210.RecordCount <> 0 Then
''      dllaccrpt210.Acc24a0 ReportTitle(210), Mid(ServerDate, 1, 4), Val(Text11), StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
''   End If
''   adoaccrpt210.Close
'   PrintData
'   '2012/12/14 End
'   FormClear
'   Screen.MousePointer = vbDefault
'   Frmacc0000.StatusBar1.Panels(1).Text = "" 'MsgText(102)
'End Sub

'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'   KeyEnter KeyCode
'   If KeyCode <> vbKeyEscape Then
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
'   End If
'End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 5820 '5250
'   Me.Height = 3270 '3930
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath4)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 5820, 3270, strBackPicPath4
   'end 2021/12/09
         
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
'   Set dllaccrpt210 = CreateObject("AccReport.ReportSelect")
   PUB_SetPrinter Me.Name, Combo1, strPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
'   Set dllaccrpt210 = Nothing
   Set Frmacc24a0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Len(Text3) = 6 Then
      Text3 = AfterZero(Text3)
   End If
   '2009/6/2 ADD BY SONIA 預設尾碼999
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Text3.Text <> "" Then Text4.Text = Left(Me.Text3.Text, 6) & "999"
   If Text3.Text <> "" Then Text4.Text = Left(Me.Text3.Text, 6) & "ZZZ"
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Len(Text4) = 6 Then
      Text4 = AfterZero(Text4)
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strAnd, strFagent As String
Dim intYears As Integer
Dim strSql As String
Dim strSQL1 As String

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   strFagent = MsgText(601)
   If Text5 <> MsgText(601) Or Text6 <> MsgText(601) Or Text7 <> MsgText(601) Or Text8 <> MsgText(601) Or Text9 <> MsgText(601) Or Text10 <> MsgText(601) Then
      strAnd = " and ("
      If Text5 <> MsgText(601) Then
         strFagent = strFagent & "a1k13 = '" & Text5 & "' or "
      End If
      If Text6 <> MsgText(601) Then
         strFagent = strFagent & "a1k13 = '" & Text6 & "' or "
      End If
      If Text7 <> MsgText(601) Then
         strFagent = strFagent & "a1k13 = '" & Text7 & "' or "
      End If
      If Text8 <> MsgText(601) Then
         strFagent = strFagent & "a1k13 = '" & Text8 & "' or "
      End If
      If Text9 <> MsgText(601) Then
         strFagent = strFagent & "a1k13 = '" & Text9 & "' or "
      End If
      If Text10 <> MsgText(601) Then
         strFagent = strFagent & "a1k13 = '" & Text10 & "' or "
      End If
      strFagent = Mid(strFagent, 1, Len(strFagent) - 4) & ") "
   Else
      strAnd = MsgText(601)
   End If
   'Modify By Sindy 2014/8/5 不輸國籍時剔除大陸,大陸一定要輸國籍
   If Text1 = "020" Then
      strSql = " and fa10 = '020'"
      strSQL1 = " and cu10 = '020'"
   Else
   '2014/8/5 END
      If Text1 <> MsgText(601) Then
         strSql = " and fa10 >= '" & Text1 & "'"
         strSQL1 = " and cu10 >= '" & Text1 & "'"
      End If
      If Text2 <> MsgText(601) Then
         strSql = strSql & " and fa10 <= '" & Text2 & "z'"
         strSQL1 = strSQL1 & " and cu10 <= '" & Text2 & "z'"
      End If
      'Add By Sindy 2014/8/5
      strSql = strSql & " and fa10 <> '020'"
      strSQL1 = strSQL1 & " and cu10 <> '020'"
      '2014/8/5 END
   End If
   'Modify By Sindy 2014/8/5 原抓案件代理人A1K03改抓請款對象.A1K28
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and A1K28 >= '" & Text3 & "'"
      strSQL1 = strSQL1 & " and A1K28 >= '" & Text3 & "'"
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and A1K28 <= '" & Text4 & "'"
      strSQL1 = strSQL1 & " and A1K28 <= '" & Text4 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQL1 = strSQL1 & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQL1 = strSQL1 & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
    'Add By Cheng 2004/03/18
    '抓未銷帳的資料
    strSql = strSql & " And A1K25 Is Null "
    strSQL1 = strSQL1 & " And A1K25 Is Null "
    'End
   adoacc1k0.CursorLocation = adUseClient
    'Modify By Cheng 2004/03/25
    '修改外幣金額的計算
'   adoacc1k0.Open "select a1k03, fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)),  sum(a1k08 - (nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) - nvl(a1k06, 0)), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), na03 from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0)" & strAnd & strFagent & strSQL & " group by a1k03, fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)), na03 union " & _
'                  "select a1k03, cu10 as fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)),  sum(a1k08 - (nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) - nvl(a1k06, 0)), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), na03 from acc1k0, customer, nation where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0)" & strAnd & strFagent & strSQL1 & " group by a1k03, cu10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)), na03", adoTaie, adOpenStatic, adLockReadOnly
   '2009/5/13 modify by sonia 修改未收金額算法,不以a0z04計算改以台幣已收金額/請款匯率,參考frmacc2210
   'adoacc1k0.Open "select a1k03, fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)),  sum(a1k08 - nvl(a1k06, 0)) - Sum(Nvl(A0Z04,0)), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), na03 from acc1k0, fagent, nation, ACC0Z0 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) And A1K01=A0Z02(+) " & strAnd & strFagent & strSQL & " group by a1k03, fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)), na03 union " & _
                  "select a1k03, cu10 as fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)),  sum(a1k08 - nvl(a1k06, 0)) - Sum(Nvl(A0Z04,0)), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), na03 from acc1k0, customer, nation, ACC0Z0 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) And A1K01=A0Z02(+) " & strAnd & strFagent & strSQL1 & " group by a1k03, cu10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)), na03", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2012/12/7 sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) ==> sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k31, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0))/a1k10,1)))
   '                          sum(round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) ==> sum(round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),0))
   '                          +a1k18
   'Modify By Sindy 2014/8/5 原抓案件代理人A1K03改抓請款對象.A1K28; +fa76
   'Modify by Amy 2019/06/04 抓大陸資料跑很慢 substr(A1K28, 1, 8) = fa01加(+)
   'Modify by Amy 2025/09/17 +,sum(nvl(a1k09,0)) 未收規費NTD
   adoacc1k0.Open "select A1K28, fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)),  sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k31, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0))/a1k10,1))), sum(round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),0)), na03,a1k18,fa76,sum(nvl(a1k09,0)) as a1k09 from acc1k0,fagent,nation where substr(A1K28, 1, 8) = fa01(+) and substr(A1K28, 9, 1) = fa02(+) and fa10 = na01 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) " & strAnd & strFagent & strSql & " group by A1K28, fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)), na03,a1k18,fa76 union " & _
                  "select A1K28, cu10 as fa10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)),  sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k31, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0))/a1k10,1))), sum(round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),0)), na03,a1k18,'B' as fa76,sum(nvl(a1k09,0)) as a1k09 from acc1k0,customer,nation where substr(A1K28, 1, 8) = cu01(+) and substr(A1K28, 9, 1) = cu02(+) and cu10 = na01 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) " & strAnd & strFagent & strSQL1 & " group by A1K28, cu10, decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)), na03,a1k18", adoTaie, adOpenStatic, adLockReadOnly
    'End
   '***** accrpt210 *****
   'a1k03 代理人編號
   'fa10  國籍
   'decode(length(a1k02), 6, substr(a1k02, 1, 2), 7, substr(a1k02, 1, 3)) 請款年度
   'sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k31, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0))/a1k10,1))) 未收外幣
   'sum(round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),0)) 未收台幣
   'na03 國家名稱
   If adoacc1k0.RecordCount = 0 Then
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc1k0.EOF = False
      adoaccrpt210.CursorLocation = adUseClient
      'Modify By Sindy 2014/8/5 原抓案件代理人A1K03改抓請款對象.A1K28
      adoaccrpt210.Open "select * from accrpt210 where r21001 = '" & strUserNum & "' and r21002 = '" & Left(Trim(adoacc1k0.Fields("fa10").Value), 3) & adoacc1k0.Fields("na03").Value & "' and r21003 = '" & adoacc1k0.Fields("A1K28").Value & "' and r21019 = '" & adoacc1k0.Fields("a1k18").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoaccrpt210.RecordCount = 0 Then
         adoaccrpt210.AddNew
         FagentSave
         adoaccrpt210.UpdateBatch
      End If
      If IsNull(adoacc1k0.Fields(2).Value) = False Then
            'Add by Amy 2025/09/17 +未收規費NTD
            adoaccrpt210.Fields("r21021").Value = adoaccrpt210.Fields("r21021").Value + Val(adoacc1k0.Fields("a1K09").Value)
            If IsNull(adoacc1k0.Fields(2).Value) = False Then
               intYears = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - Val(adoacc1k0.Fields(2).Value)
            Else
               intYears = 0
            End If
            Select Case intYears
               Case 0
                  If IsNull(adoacc1k0.Fields(3).Value) = False Then
                     adoaccrpt210.Fields("r21005").Value = adoaccrpt210.Fields("r21005").Value + Val(adoacc1k0.Fields(3).Value)
                     If IsNull(adoacc1k0.Fields(4).Value) = False Then
                        adoaccrpt210.Fields("r21012").Value = adoaccrpt210.Fields("r21012").Value + Val(adoacc1k0.Fields(4).Value)
                     End If
                  End If
               Case 1
                  If IsNull(adoacc1k0.Fields(3).Value) = False Then
                     adoaccrpt210.Fields("r21006").Value = adoaccrpt210.Fields("r21006").Value + Val(adoacc1k0.Fields(3).Value)
                     If IsNull(adoacc1k0.Fields(4).Value) = False Then
                        adoaccrpt210.Fields("r21013").Value = adoaccrpt210.Fields("r21013").Value + Val(adoacc1k0.Fields(4).Value)
                     End If
                  End If
               Case 2
                  If IsNull(adoacc1k0.Fields(3).Value) = False Then
                     adoaccrpt210.Fields("r21007").Value = adoaccrpt210.Fields("r21007").Value + Val(adoacc1k0.Fields(3).Value)
                     If IsNull(adoacc1k0.Fields(4).Value) = False Then
                        adoaccrpt210.Fields("r21014").Value = adoaccrpt210.Fields("r21014").Value + Val(adoacc1k0.Fields(4).Value)
                     End If
                  End If
               Case 3
                  If IsNull(adoacc1k0.Fields(3).Value) = False Then
                     adoaccrpt210.Fields("r21008").Value = adoaccrpt210.Fields("r21008").Value + Val(adoacc1k0.Fields(3).Value)
                     If IsNull(adoacc1k0.Fields(4).Value) = False Then
                        adoaccrpt210.Fields("r21015").Value = adoaccrpt210.Fields("r21015").Value + Val(adoacc1k0.Fields(4).Value)
                     End If
                  End If
               Case 4
                  If IsNull(adoacc1k0.Fields(3).Value) = False Then
                     adoaccrpt210.Fields("r21009").Value = adoaccrpt210.Fields("r21009").Value + Val(adoacc1k0.Fields(3).Value)
                     If IsNull(adoacc1k0.Fields(4).Value) = False Then
                        adoaccrpt210.Fields("r21016").Value = adoaccrpt210.Fields("r21016").Value + Val(adoacc1k0.Fields(4).Value)
                     End If
                  End If
               Case 5
                  Five_NineSave
               Case 6
                  Five_NineSave
               Case 7
                  Five_NineSave
               Case 8
                  Five_NineSave
               Case 9
                  Five_NineSave
               Case Else
                  If IsNull(adoacc1k0.Fields(3).Value) = False Then
                     adoaccrpt210.Fields("r21011").Value = adoaccrpt210.Fields("r21011").Value + Val(adoacc1k0.Fields(3).Value)
                     If IsNull(adoacc1k0.Fields(4).Value) = False Then
                        adoaccrpt210.Fields("r21018").Value = adoaccrpt210.Fields("r21018").Value + Val(adoacc1k0.Fields(4).Value)
                     End If
                  End If
            End Select
      End If
      adoaccrpt210.UpdateBatch
      adoaccrpt210.Close
      adoacc1k0.MoveNext
   Loop
   adoacc1k0.Close
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
Private Sub Accrpt210Delete()
   adoTaie.Execute "delete from accrpt210"
End Sub

'*************************************************
'  代理人儲存
'
'*************************************************
Private Sub FagentSave()
   adoaccrpt210.Fields("r21001").Value = strUserNum
   If IsNull(adoacc1k0.Fields("fa10").Value) Then
      adoaccrpt210.Fields("r21002").Value = Null
   Else
      adoaccrpt210.Fields("r21002").Value = Left(Trim(adoacc1k0.Fields("fa10").Value), 3)
      adoaccrpt210.Fields("r21002").Value = adoaccrpt210.Fields("r21002").Value & Trim(NationQuery(adoacc1k0.Fields("fa10").Value, 1))
   End If
   'Modify By Sindy 2014/8/5 原抓案件代理人A1K03改抓請款對象.A1K28
   If IsNull(adoacc1k0.Fields("A1K28").Value) Then
      adoaccrpt210.Fields("r21003").Value = Null
   Else
      'Modify By Sindy 2014/8/5 原抓案件代理人A1K03改抓請款對象.A1K28
      adoaccrpt210.Fields("r21003").Value = adoacc1k0.Fields("A1K28").Value
        'Modify By Cheng 2004/03/25
'      adoaccrpt210.Fields("r21004").Value = FagentQuery(adoaccrpt210.Fields("r21003").Value, 2)
        If Left("" & adoaccrpt210.Fields("r21003").Value, 1) = "Y" Then
            adoaccrpt210.Fields("r21004").Value = FagentQuery(adoaccrpt210.Fields("r21003").Value, 2)
        ElseIf Left("" & adoaccrpt210.Fields("r21003").Value, 1) = "X" Then
            adoaccrpt210.Fields("r21004").Value = CustomerQuery(adoaccrpt210.Fields("r21003").Value, 2)
        Else
            adoaccrpt210.Fields("r21004").Value = ""
        End If
        'End
      If adoaccrpt210.Fields("r21004").Value = "" Then
        'Modify By Cheng 2004/03/25
'        adoaccrpt210.Fields("r21004").Value = FagentQuery(adoaccrpt210.Fields("r21003").Value, 3)
        If Left("" & adoaccrpt210.Fields("r21003").Value, 1) = "Y" Then
            adoaccrpt210.Fields("r21004").Value = FagentQuery(adoaccrpt210.Fields("r21003").Value, 1)
        ElseIf Left("" & adoaccrpt210.Fields("r21003").Value, 1) = "X" Then
            adoaccrpt210.Fields("r21004").Value = CustomerQuery(adoaccrpt210.Fields("r21003").Value, 1)
        Else
            adoaccrpt210.Fields("r21004").Value = ""
        End If
        'End
      End If
      If adoaccrpt210.Fields("r21004").Value = "" Then
        'Modify By Cheng 2004/03/25
'        adoaccrpt210.Fields("r21004").Value = FagentQuery(adoaccrpt210.Fields("r21003").Value, 3)
        If Left("" & adoaccrpt210.Fields("r21003").Value, 1) = "Y" Then
            adoaccrpt210.Fields("r21004").Value = FagentQuery(adoaccrpt210.Fields("r21003").Value, 3)
        ElseIf Left("" & adoaccrpt210.Fields("r21003").Value, 1) = "X" Then
            adoaccrpt210.Fields("r21004").Value = CustomerQuery(adoaccrpt210.Fields("r21003").Value, 3)
        Else
            adoaccrpt210.Fields("r21004").Value = ""
        End If
        'End
      End If
   End If
   adoaccrpt210.Fields("r21005").Value = 0
   adoaccrpt210.Fields("r21006").Value = 0
   adoaccrpt210.Fields("r21007").Value = 0
   adoaccrpt210.Fields("r21008").Value = 0
   adoaccrpt210.Fields("r21009").Value = 0
   adoaccrpt210.Fields("r21010").Value = 0
   adoaccrpt210.Fields("r21011").Value = 0
   adoaccrpt210.Fields("r21012").Value = 0
   adoaccrpt210.Fields("r21013").Value = 0
   adoaccrpt210.Fields("r21014").Value = 0
   adoaccrpt210.Fields("r21015").Value = 0
   adoaccrpt210.Fields("r21016").Value = 0
   adoaccrpt210.Fields("r21017").Value = 0
   adoaccrpt210.Fields("r21018").Value = 0
   adoaccrpt210.Fields("r21019").Value = "" & adoacc1k0.Fields("a1k18").Value 'Add By Sindy 2012/12/13
   adoaccrpt210.Fields("r21020").Value = "" & adoacc1k0.Fields("fa76").Value 'Add By Sindy 2014/8/6
   adoaccrpt210.Fields("r21021").Value = 0 'Add by Amy 2025/09/17
End Sub

'*************************************************
'  五至九年帳款儲存
'
'*************************************************
Private Sub Five_NineSave()
   If IsNull(adoacc1k0.Fields(3).Value) = False Then
      '93.11.8 MODIFY BY SONIA
      'adoaccrpt210.Fields("r21010").Value = Val(adoacc1k0.Fields(3).Value)
      If IsNull(adoaccrpt210.Fields("r21010").Value) = True Then
         adoaccrpt210.Fields("r21010").Value = Val(adoacc1k0.Fields(3).Value)
      Else
         adoaccrpt210.Fields("r21010").Value = adoaccrpt210.Fields("r21010").Value + Val(adoacc1k0.Fields(3).Value)
      End If
      '93.11.8 END
      If IsNull(adoacc1k0.Fields(4).Value) = False Then
         '93.11.8 MODIFY BY SONIA
         'adoaccrpt210.Fields("r21017").Value = Val(adoacc1k0.Fields(4).Value)
         If IsNull(adoaccrpt210.Fields("r21017").Value) = True Then
            adoaccrpt210.Fields("r21017").Value = Val(adoacc1k0.Fields(4).Value)
         Else
            adoaccrpt210.Fields("r21017").Value = adoaccrpt210.Fields("r21017").Value + Val(adoacc1k0.Fields(4).Value)
         End If
         '93.11.8 END
      End If
   End If
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text11 = ""
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
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text5 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text6 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text7 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text8 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text9 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text10 <> MsgText(601) Then
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

''Add By Sindy 2012/12/14
'Private Sub PrintDetail(intFormat As Integer)
'   Dim i As Integer
'
'   For i = 0 To 10
'      If i <= 3 Then
'         Printer.CurrentX = PLeft(i)
'      Else
'         If intFormat = 0 Then
'            Printer.CurrentX = PLeft(i) - Printer.TextWidth(Format(strTemp(i), FDollar)) '美金
'         Else
'            Printer.CurrentX = PLeft(i) - Printer.TextWidth(Format(strTemp(i), DDollar2))
'         End If
'      End If
'      Printer.CurrentY = iLine * 300
'      If i <= 3 Then
'         Printer.Print "" & strTemp(i)
'      Else
'         If intFormat = 0 Then
'            Printer.Print Format(strTemp(i), FDollar) '美金
'         Else
'            Printer.Print Format(strTemp(i), DDollar2)
'         End If
'      End If
'   Next i
'   iLine = iLine + 1
'End Sub
'
''Add By Sindy 2012/12/14
'Private Sub PrintData()
'   Dim rsReport As ADODB.Recordset
'   Dim i As Integer
'
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   PUB_RestorePrinter Combo1
'   Printer.Orientation = 2 '1.直印 2.橫印
'   iPrintLine = 35
'   Printer.Font.Name = "細明體"
'
'   strNation = "": StrFa = ""
'   strExc(0) = "select r21002,r21003,r21004,r21019,r21005,r21006,r21007,r21008,r21009,r21010,r21011 from accrpt210 where r21001='" & strUserNum & "' order by r21002,r21003,r21019"
'   intI = 1
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With rsReport
'         iLine = 1
'         PrintHead
'         Do While Not .EOF
'            If strNation <> "" And strNation <> .Fields(0) Then '總計國籍帳齡
'               'Call PrintSum(strNation) '小計
'            End If
'            For i = 0 To 10
'               If i = 0 Then
'                  strTemp(i) = Left(.Fields(i), 6)
'               ElseIf i = 2 Then
'                  strTemp(i) = Left(.Fields(i), 15)
'               Else
'                  strTemp(i) = .Fields(i)
'               End If
'            Next i
'            If iLine > iPrintLine Then
'               Printer.NewPage
'               iLine = 1
'               PrintHead
'            End If
'            If strNation <> "" And strNation = .Fields(0) Then
'               strTemp(0) = ""
'            End If
'            If StrFa <> "" And StrFa = .Fields(1) Then
'               strTemp(1) = ""
'               strTemp(2) = ""
'            End If
'            If Text11 = "1" Then '1.細目
'               Call PrintDetail(0)
'            End If
'            strNation = .Fields(0)
'            StrFa = .Fields(1)
'            .MoveNext
'         Loop
'      End With
'      'Call PrintSum(strNation) '小計
'      Call PrintSum("") '合計
'
'      Printer.EndDoc
'      PUB_RestorePrinter strPrinter
'      Call ShowPrintOk
'   End If
'   Set rsReport = Nothing
'End Sub
'
''Add By Sindy 2012/12/14
'Private Sub PrintSum(strKey1 As String)
'Dim rs As ADODB.Recordset
'Dim i As Integer
'Dim intRow As Integer
'Dim strText As String
'Dim dbl_TotAmt As Double 'Add By Sindy 2013/5/8
'
'   '外幣
'   If strKey1 = "" Then '總計
'      strExc(0) = "select r21019,sum(r21005),sum(r21006),sum(r21007),sum(r21008),sum(r21009),sum(r21010),sum(r21011) from accrpt210 where r21001='" & strUserNum & "' group by r21019 order by r21019"
'   Else
'      If Text11 = "1" Then
'         Printer.CurrentX = PLeft(0)
'         Printer.CurrentY = iLine * 300
'         Printer.Print String(132, "-")
'         iLine = iLine + 1
'      End If
'      strExc(0) = "select r21019,sum(r21005),sum(r21006),sum(r21007),sum(r21008),sum(r21009),sum(r21010),sum(r21011) from accrpt210 where r21001='" & strUserNum & "' and r21002='" & strKey1 & "' group by r21019 order by r21019"
'   End If
'   intI = 1
'   Set rs = ClsLawReadRstMsg(intI, strExc(0))
'   intRow = 0
'   If intI = 1 Then
'      With rs
'         .MoveFirst
'         Do While Not .EOF
'            intRow = intRow + 1
'            For i = 0 To 7
'               strTemp(i) = .Fields(i)
'            Next i
'            If iLine > iPrintLine Then
'               Printer.NewPage
'               iLine = 1
'               PrintHead
'            End If
'            '只列印總計時,第一列第一個欄位放國籍
'            If Text11 = "2" And intRow = 1 And strKey1 <> "" Then
'               Printer.CurrentX = PLeft(0)
'               Printer.CurrentY = iLine * 300
'               Printer.Print strNation
'            End If
'            For i = 0 To 7
'               If i = 0 Then
'                  If strKey1 = "" Then '總計
'                     strText = strTemp(i) & "合計："
'                  Else
'                     strText = strTemp(i) & "小計："
'                  End If
'                  Printer.CurrentX = PLeft(i + 3) - Printer.TextWidth(strText)
'               Else
'                  Printer.CurrentX = PLeft(i + 3) - Printer.TextWidth(Format(strTemp(i), FDollar))
'               End If
'               Printer.CurrentY = iLine * 300
'               If i = 0 Then
'                  Printer.Print strText
'               Else
'                  Printer.Print Format(strTemp(i), FDollar)
'               End If
'            Next i
'            iLine = iLine + 1
'            .MoveNext
'         Loop
'      End With
'   End If
'   '台幣
'   dbl_TotAmt = 0 'Add By Sindy 2013/5/8
'   If strKey1 = "" Then '總計
'      strExc(0) = "select sum(r21012),sum(r21013),sum(r21014),sum(r21015),sum(r21016),sum(r21017),sum(r21018) from accrpt210 where r21001='" & strUserNum & "'"
'   Else
'      strExc(0) = "select sum(r21012),sum(r21013),sum(r21014),sum(r21015),sum(r21016),sum(r21017),sum(r21018) from accrpt210 where r21001='" & strUserNum & "' and r21002='" & strKey1 & "'"
'   End If
'   intI = 1
'   Set rs = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With rs
'         Do While Not .EOF
'            For i = 0 To 6
'               strTemp(i) = "" & .Fields(i)
'            Next i
'            If iLine > iPrintLine Then
'               Printer.NewPage
'               iLine = 1
'               PrintHead
'            End If
'            If strKey1 = "" Then '總計
'               strText = "台幣合計："
'            Else
'               strText = "台幣小計："
'            End If
'            Printer.CurrentX = PLeft(3) - Printer.TextWidth(strText)
'            Printer.CurrentY = iLine * 300
'            Printer.Print strText
'            For i = 0 To 6
'               Printer.CurrentX = PLeft(i + 4) - Printer.TextWidth(Format(strTemp(i), DDollar2))
'               Printer.CurrentY = iLine * 300
'               Printer.Print Format(strTemp(i), DDollar2)
'               dbl_TotAmt = dbl_TotAmt + Format(strTemp(i), DDollar2) 'Add By Sindy 2013/5/8
'            Next i
'            iLine = iLine + 1
'            .MoveNext
'         Loop
'      End With
'   End If
'   If strKey1 <> "" Then
'      Printer.CurrentX = PLeft(0)
'      Printer.CurrentY = iLine * 300
'      Printer.Print String(132, "-")
'      iLine = iLine + 1
'   'Add By Sindy 2013/5/8
'   Else
'      Printer.CurrentX = PLeft(3) - Printer.TextWidth("台幣總計：")
'      Printer.CurrentY = iLine * 300
'      Printer.Print "台幣總計："
'      Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dbl_TotAmt, DDollar2))
'      Printer.CurrentY = iLine * 300
'      Printer.Print Format(dbl_TotAmt, DDollar2)
'   '2013/5/8 End
'   End If
'
'   Set rs = Nothing
'End Sub
'
''Add By Sindy 2012/12/14
'Private Sub GetPleft()
'   PLeft(0) = 500
'   PLeft(1) = 1500
'   PLeft(2) = 2800
'   PLeft(3) = 4800 '幣別
'   PLeft(4) = 6500
'   PLeft(5) = 8000
'   PLeft(6) = 9500
'   PLeft(7) = 11000
'   PLeft(8) = 12500
'   PLeft(9) = 14500
'   PLeft(10) = 16000
'End Sub
'
''Add By Sindy 2012/12/14
'Private Sub PrintHead()
'   GetPleft
'
'   Printer.Font.Size = 14
'   Printer.Font.Underline = False
'   Printer.FontBold = True
'
'   strExc(1) = ReportTitle(210)
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(1)) / 2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print strExc(1)
'
'   Printer.Font.Size = 12
'   Printer.FontBold = False
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = 900
'   Printer.Print "列印人員：" & strUserName
'   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
'   Printer.CurrentY = 900
'   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
'   Printer.CurrentY = 1200
'   Printer.Print "頁　　次：" & Printer.Page
'   iLine = 6
'
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "國籍"
'   If Text11 = "1" Then
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = iLine * 300
'      Printer.Print "請款對象" 'Modify By Sindy 2014/8/6 "代理人編號"
'      Printer.CurrentX = PLeft(2)
'      Printer.CurrentY = iLine * 300
'      Printer.Print "名稱" 'Modify By Sindy 2014/8/6 "代理人名稱"
'      Printer.CurrentX = PLeft(3)
'      Printer.CurrentY = iLine * 300
'      Printer.Print "幣別"
'   End If
'   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Val(Mid(ServerDate, 1, 4)))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Val(Mid(ServerDate, 1, 4))
'   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Val(Mid(ServerDate, 1, 4)) - 1)
'   Printer.CurrentY = iLine * 300
'   Printer.Print Val(Mid(ServerDate, 1, 4)) - 1
'   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Val(Mid(ServerDate, 1, 4)) - 2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print Val(Mid(ServerDate, 1, 4)) - 2
'   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Val(Mid(ServerDate, 1, 4)) - 3)
'   Printer.CurrentY = iLine * 300
'   Printer.Print Val(Mid(ServerDate, 1, 4)) - 3
'   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Val(Mid(ServerDate, 1, 4)) - 4)
'   Printer.CurrentY = iLine * 300
'   Printer.Print Val(Mid(ServerDate, 1, 4)) - 4
'   Printer.CurrentX = PLeft(9) - Printer.TextWidth((Val(Mid(ServerDate, 1, 4)) - 5) & " ~ " & (Val(Mid(ServerDate, 1, 4)) - 9))
'   Printer.CurrentY = iLine * 300
'   Printer.Print (Val(Mid(ServerDate, 1, 4)) - 5) & " ~ " & (Val(Mid(ServerDate, 1, 4)) - 9)
'   Printer.CurrentX = PLeft(10) - Printer.TextWidth(Val(Mid(ServerDate, 1, 4)) - 10 & " ~ ")
'   Printer.CurrentY = iLine * 300
'   Printer.Print Val(Mid(ServerDate, 1, 4)) - 10 & " ~ "
'
'   iLine = iLine + 1
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = iLine * 300
'   Printer.Print String(132, "-")
'   iLine = iLine + 1
'End Sub

'Add By Sindy 2014/6/5
'*************************************************
' 產生Excel資料
'
'*************************************************
Public Sub PrintExcel()
Dim strFilePath As String
Dim i As Integer
   
On Error GoTo ErrHnd
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   If adoaccrpt210.State = adStateOpen Then
      adoaccrpt210.Close
   End If
   adoaccrpt210.CursorLocation = adUseClient
   'Modify by Amy 2025/09/17 幣別前加 R21021 (台幣未收規費)
   adoaccrpt210.Open "select r21002,r21003,r21004,sum(r21021),r21019,sum(r21005),sum(r21006),sum(r21007),sum(r21008),sum(r21009),sum(r21010),sum(r21011) from accrpt210 where r21001='" & strUserNum & "' group by r21002,r21003,r21004,r21019 order by r21002,r21003,r21004,r21019", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt210.RecordCount = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      adoaccrpt210.Close
      Exit Sub
   End If
   
   intPage = 0
   intCounter = 0
   dblSkipPageRow = 0
   strNation = "": StrFa = ""
   
   'Excel檔案路徑
   strFilePath = strExcelPath & Me.Caption & ACDate(ServerDate) & ServerTime & MsgText(43)
   If Dir(strFilePath) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFilePath
   End If
   
   Set xlsAnnuity = New Excel.Application
   xlsAnnuity.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   With wksAnnuity
      .PageSetup.Orientation = xlLandscape '橫印xlLandscape,直印xlPortrait
      .PageSetup.LeftMargin = 28.34
      .PageSetup.RightMargin = 28.34
      .PageSetup.TopMargin = 42.51
      .PageSetup.BottomMargin = 42.51
      .PageSetup.HeaderMargin = 28.34
      .PageSetup.FooterMargin = 28.34
      '設定各欄位長度
      .Columns("A:A").ColumnWidth = 10 '國籍
      .Columns("B:B").ColumnWidth = 11 '請款對象
      'Modify by Amy 2025/09/17 原:17
      .Columns("C:C").ColumnWidth = 20 '名稱
      'Modify by Amy 2025/09/17 幣別 前插入一欄 未收規費NTD
      .Columns("D:D").ColumnWidth = 15  '未收規費NTD
      .Columns("E:E").ColumnWidth = 5  '幣別
      .Columns("F:F").ColumnWidth = 12 '統計欄位1
      .Columns("G:G").ColumnWidth = 12 '統計欄位2
      .Columns("H:H").ColumnWidth = 12 '統計欄位3
      .Columns("I:I").ColumnWidth = 12 '統計欄位4
      .Columns("J:J").ColumnWidth = 12 '統計欄位5
      .Columns("K:K").ColumnWidth = 12 '統計欄位6
      .Columns("L:L").ColumnWidth = 12 '統計欄位7
      'end 2025/09/17
      '逐筆填值
      adoaccrpt210.MoveFirst
      Do While adoaccrpt210.EOF = False
         If strNation <> "" And strNation <> adoaccrpt210.Fields(0) Then '總計國籍帳齡
            'Call PrintExcelSum(strNation) '小計
         End If
         '一開始或資料已填滿一頁時跳頁
         If intCounter = 0 Or dblSkipPageRow >= 28 Then
            intCounter = intCounter + 1
            If dblSkipPageRow >= 28 Then
               '換頁
               .Range("A" & intCounter).Select
               .HPageBreaks.add Before:=.Application.ActiveCell
            End If
            dblSkipPageRow = 0
            Call PrintExcelTitle(strTemp()) 'Modify by Amy 2025/09/17 +strTemp()
         Else
            If Text11 = "1" Then '1.細目
               intCounter = intCounter + 1
            End If
         End If
         '讀取資料
         'Modify by Amy 2025/09/17 原:10->改抓 Ubound(strTemp)
         For i = 0 To UBound(strTemp)
            'Modify by Amy 2023/08/18 資料都顯示-婉莘
'            If i = 0 Then
'               strTemp(i) = Left(adoaccrpt210.Fields(i), 6)
'            ElseIf i = 2 Then
'               strTemp(i) = Left(adoaccrpt210.Fields(i), 15)
'            Else
               strTemp(i) = adoaccrpt210.Fields(i)
'            End If
         Next i
'         If strNation <> "" And strNation = adoaccrpt210.Fields(0) Then
'            strTemp(0) = ""
'         End If
'         If StrFa <> "" And StrFa = adoaccrpt210.Fields(1) Then
'            strTemp(1) = ""
'            strTemp(2) = ""
'         End If
         'end 2023/08/18
         If Text11 = "1" Then '1.細目
            .Range("A" & intCounter).Value = strTemp(0)
            .Range("B" & intCounter).Value = strTemp(1)
            .Range("C" & intCounter).Value = strTemp(2)
            .Range("D" & intCounter).Value = strTemp(3)
            'Add by Amy 2025/09/17 幣別 前插入一欄 未收規費NTD
            .Range("D" & intCounter).NumberFormatLocal = "#,##0.00_ "
            .Range("E" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0.00_ "
            .Range("E" & intCounter).Value = CStr(strTemp(4))
            .Range("F" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0.00_ "
            .Range("F" & intCounter).Value = CStr(strTemp(5))
            .Range("G" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0.00_ "
            .Range("G" & intCounter).Value = CStr(strTemp(6))
            .Range("H" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0.00_ "
            .Range("H" & intCounter).Value = CStr(strTemp(7))
            .Range("I" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0.00_ "
            .Range("I" & intCounter).Value = CStr(strTemp(8))
            .Range("J" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0.00_ "
            .Range("J" & intCounter).Value = CStr(strTemp(9))
            .Range("K" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0.00_ "
            .Range("K" & intCounter).Value = CStr(strTemp(10))
            'Add by Amy 2025/09/17
            .Range("L" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0.00_ "
            .Range("L" & intCounter).Value = CStr(strTemp(11))
            'end 2025/09/17
            dblSkipPageRow = dblSkipPageRow + 1
         End If
         strNation = adoaccrpt210.Fields(0)
         StrFa = adoaccrpt210.Fields(1)
         adoaccrpt210.MoveNext
      Loop
      'Call PrintExcelSum(strNation) '小計
      Call PrintExcelSum("") '合計
      intCounter = intCounter + 1
      'Modify by Amy 2025/09/17 原E欄
      .Range("F" & intCounter).Value = "***結束***"
   End With
'   xlsAnnuity.Visible = True
'   xlsAnnuity.WindowState = wdWindowStateMaximize
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=-4143
   Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=56
   End If
   'end 2016/06/23
   xlsAnnuity.Workbooks.Close
   xlsAnnuity.Quit
   
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   adoaccrpt210.Close
   MsgBox "檔案已產生！" & vbCrLf & vbCrLf & "存放至 " & strFilePath
   Exit Sub
   
ErrHnd:
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   adoaccrpt210.Close
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2014/6/5
Public Sub PrintExcelTitle(arrCol() As String)
Dim i As Integer, strTemp As String
Dim strText As String
   
   intPage = intPage + 1
   With wksAnnuity
      .Range("F" & intCounter).Value = "國外帳齡分析表"
      'Modify by Amy 2025/09/17 原K欄->& Chr(UBound(arrcol) + 65)
      strTemp = "A" & intCounter & ":" & Chr(UBound(arrCol) + 65) & intCounter
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter
         .Font.Size = 18
      End With
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "列印人員：" & strUserName
      .Range("J" & intCounter).Value = "列印日期：" & ChangeWStringToTDateString(strSrvDate(1))
      intCounter = intCounter + 1
      .Range("J" & intCounter).Value = "頁　　次：" & intPage
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "國籍"
      If Text11 = "1" Then '1.細目
         .Range("B" & intCounter).Value = "請款對象" 'Modify By Sindy 2014/8/6 "代理人編號"
         .Range("C" & intCounter).Value = "名稱" 'Modify By Sindy 2014/8/6 "代理人名稱"
     'Modify by Amy 2025/09/17 幣別 前插入一欄 未收規費NTD
         .Range("D" & intCounter).Value = "未收規費NTD"
         .Range("E" & intCounter).Value = "幣別"
      End If
      .Range("F" & intCounter).Value = Val(Mid(ServerDate, 1, 4))
      .Range("G" & intCounter).Value = Val(Mid(ServerDate, 1, 4)) - 1
      .Range("H" & intCounter).Value = Val(Mid(ServerDate, 1, 4)) - 2
      .Range("I" & intCounter).Value = Val(Mid(ServerDate, 1, 4)) - 3
      .Range("J" & intCounter).Value = Val(Mid(ServerDate, 1, 4)) - 4
      .Range("K" & intCounter).Value = (Val(Mid(ServerDate, 1, 4)) - 5) & " ~ " & (Val(Mid(ServerDate, 1, 4)) - 9)
      .Range("K" & intCounter).Select
      'end 2025/09/17
      .Application.Selection.HorizontalAlignment = xlRight
      'Modify by Amy 2025/09/24原K欄->& Chr(UBound(arrcol) + 65)
      .Range(Chr(UBound(arrCol) + 65) & intCounter).Value = Val(Mid(ServerDate, 1, 4)) - 10 & " ~ "
      .Range(Chr(UBound(arrCol) + 65) & intCounter).Select
      'end 2025/09/24
      .Application.Selection.HorizontalAlignment = xlRight
      'Modify by Amy 2025/09/17 原K欄->& Chr(UBound(arrcol) + 65)
      strTemp = "A" & intCounter & ":" & Chr(UBound(arrCol) + 65) & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
      End With
      intCounter = intCounter + 1
   End With
End Sub

'Add By Sindy 2014/6/5
Private Sub PrintExcelSum(strKey1 As String)
Dim Rs As ADODB.Recordset
Dim i As Integer
Dim intRow As Integer
Dim strText As String, strTitle As String
Dim dbl_TotAmt As Double
   
   If Text11 = "1" Then
      'Modify by Amy 2025/09/17 原K欄->& Chr(UBound(strTemp) + 65)
      strText = "A" & intCounter & ":" & Chr(UBound(strTemp) + 65) & intCounter
      wksAnnuity.Range(strText).Select
      With wksAnnuity.Application.Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
      End With
   End If
   '外幣
   'Modify by Amy 2025/09/17 +未收規費NTD
   If strKey1 = "" Then '總計
      'strExc(0) = "select r21019,sum(r21005),sum(r21006),sum(r21007),sum(r21008),sum(r21009),sum(r21010),sum(r21011) from accrpt210 where r21001='" & strUserNum & "' group by r21019 order by r21019"
      strExc(0) = "select r21019,sum(r21005),sum(r21006),sum(r21007),sum(r21008),sum(r21009),sum(r21010),sum(r21011) from accrpt210 where r21001='" & strUserNum & "' group by r21019 "
   Else
      'strExc(0) = "select r21019,sum(r21005),sum(r21006),sum(r21007),sum(r21008),sum(r21009),sum(r21010),sum(r21011) from accrpt210 where r21001='" & strUserNum & "' and r21002='" & strKey1 & "' group by r21019 order by r21019"
      strExc(0) = "select r21019,sum(r21005),sum(r21006),sum(r21007),sum(r21008),sum(r21009),sum(r21010),sum(r21011) from accrpt210 where r21001='" & strUserNum & "' and r21002='" & strKey1 & "' group by r21019 "
   End If
   strExc(0) = strExc(0) & "Union Select 'ZZZ',sum(r21021),0,0,0,0,0,0 from accrpt210 where r21001='" & strUserNum & "' order by r21019"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   intRow = 0
   If intI = 1 Then
      With Rs
         .MoveFirst
         Do While Not .EOF
            intRow = intRow + 1
            intCounter = intCounter + 1
            For i = 0 To 7
               strTemp(i) = "" & .Fields(i) 'Add by Amy 2020/09/10 避免null 會錯
            Next i
            '一開始或資料已填滿一頁時跳頁
            If intCounter = 1 Or dblSkipPageRow >= 28 Then
               If dblSkipPageRow >= 28 Then
                  '換頁
                  wksAnnuity.Range("A" & intCounter).Select
                  wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
               End If
               dblSkipPageRow = 0
               Call PrintExcelTitle(strTemp()) 'Modify by Amy 2025/09/17 +strtemp()
            End If
            '只列印總計時,第一列第一個欄位放國籍
            If Text11 = "2" And intRow = 1 And strKey1 <> "" Then
               wksAnnuity.Range("A" & intCounter).Value = strNation
            End If
            
            'Modify by Amy 2025/09/17 +if
            If strTemp(0) = "ZZZ" Then
               strText = "未收規費NTD合計："
            ElseIf strKey1 = "" Then '總計
               strText = strTemp(0) & "合計："
            Else
               strText = strTemp(0) & "小計："
            End If
            wksAnnuity.Range("C" & intCounter).Value = strText
            'Modify by Amy 2025/09/17 +if 幣別 前插入一欄 未收規費NTD,所有欄位往後移
            If strTemp(0) = "ZZZ" Then
               wksAnnuity.Range("D" & intCounter).Value = strTemp(1)
               wksAnnuity.Range("D" & intCounter).NumberFormatLocal = "#,##0.00_ "
            Else
            wksAnnuity.Range("F" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("F" & intCounter).Value = strTemp(1)
            wksAnnuity.Range("G" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("G" & intCounter).Value = strTemp(2)
            wksAnnuity.Range("H" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("H" & intCounter).Value = strTemp(3)
            wksAnnuity.Range("I" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("I" & intCounter).Value = strTemp(4)
            wksAnnuity.Range("J" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("J" & intCounter).Value = strTemp(5)
            wksAnnuity.Range("K" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("K" & intCounter).Value = strTemp(6)
            wksAnnuity.Range("L" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("L" & intCounter).Value = strTemp(7)
            End If
            'end 2025/09/17
            dblSkipPageRow = dblSkipPageRow + 1
            .MoveNext
         Loop
      End With
   End If
   '台幣
   dbl_TotAmt = 0
   If strKey1 = "" Then '總計
      strExc(0) = "select sum(r21012),sum(r21013),sum(r21014),sum(r21015),sum(r21016),sum(r21017),sum(r21018) from accrpt210 where r21001='" & strUserNum & "'"
   Else
      strExc(0) = "select sum(r21012),sum(r21013),sum(r21014),sum(r21015),sum(r21016),sum(r21017),sum(r21018) from accrpt210 where r21001='" & strUserNum & "' and r21002='" & strKey1 & "'"
   End If
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With Rs
         Do While Not .EOF
            intCounter = intCounter + 1
            For i = 0 To 6
               strTemp(i) = "" & .Fields(i)
            Next i
            '一開始或資料已填滿一頁時跳頁
            If intCounter = 1 Or dblSkipPageRow >= 28 Then
               If dblSkipPageRow >= 28 Then
                  '換頁
                  wksAnnuity.Range("A" & intCounter).Select
                  wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
               End If
               dblSkipPageRow = 0
               Call PrintExcelTitle(strTemp()) 'Modify by Amy 2025/09/17 +strTemp()
            End If
            If strKey1 = "" Then '總計
               strText = "台幣合計："
            Else
               strText = "台幣小計："
            End If
            wksAnnuity.Range("C" & intCounter).Value = strText
            'Modify by Amy 2025/09/17 幣別 前插入一欄 未收規費NTD,所有欄位往後移
            wksAnnuity.Range("F" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0_ "
            wksAnnuity.Range("F" & intCounter).Value = strTemp(0)
            wksAnnuity.Range("G" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0_ "
            wksAnnuity.Range("G" & intCounter).Value = strTemp(1)
            wksAnnuity.Range("H" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0_ "
            wksAnnuity.Range("H" & intCounter).Value = strTemp(2)
            wksAnnuity.Range("I" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0_ "
            wksAnnuity.Range("I" & intCounter).Value = strTemp(3)
            wksAnnuity.Range("J" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0_ "
            wksAnnuity.Range("J" & intCounter).Value = strTemp(4)
            wksAnnuity.Range("K" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0_ "
            wksAnnuity.Range("K" & intCounter).Value = strTemp(5)
            wksAnnuity.Range("L" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0_ "
            wksAnnuity.Range("L" & intCounter).Value = strTemp(6)
            'end 2025/09/17
            For i = 0 To 6
               dbl_TotAmt = dbl_TotAmt + strTemp(i)
            Next i
            dblSkipPageRow = dblSkipPageRow + 1
            .MoveNext
         Loop
      End With
   End If
   If strKey1 <> "" Then
      'Modify by Amy 2025/09/17 原K欄->& Chr(UBound(strTemp) + 65)
      strText = "A" & intCounter & ":" & Chr(UBound(strTemp) + 65) & intCounter
      wksAnnuity.Range(strText).Select
      With wksAnnuity.Application.Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
      End With
   Else
      intCounter = intCounter + 1
      wksAnnuity.Range("C" & intCounter).Value = "台幣總計："
      'Modify by Amy 2025/09/17 幣別 前插入一欄 未收規費NTD,所有欄位往後移
      wksAnnuity.Range("F" & intCounter).Select
      wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0_ "
      wksAnnuity.Range("F" & intCounter).Value = dbl_TotAmt
      'end 2025/09/17
      dblSkipPageRow = dblSkipPageRow + 1
   End If
   
   'Add By Sindy 2014/8/6 增加來所性質統計
   strExc(0) = "select r21019,sum(r21005),sum(r21006),sum(r21007),sum(r21008),sum(r21009),sum(r21010),sum(r21011),r21020 from accrpt210 where r21001='" & strUserNum & "' group by r21020,r21019 union " & _
               "select 'ZZZ',sum(r21012),sum(r21013),sum(r21014),sum(r21015),sum(r21016),sum(r21017),sum(r21018),r21020 from accrpt210 where r21001='" & strUserNum & "' group by r21020 " & _
               "order by r21020,r21019"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With Rs
         .MoveFirst
         strTitle = ""
         Do While Not .EOF
            For i = 0 To 8
               strTemp(i) = "" & .Fields(i) 'Add by Amy 2020/09/10 避免null 會錯
            Next i
            If strTemp(8) = "A" Then
               strTemp(8) = "代理人"
            ElseIf strTemp(8) = "B" Then
               strTemp(8) = "申請人"
            Else
               strTemp(8) = "其他"
            End If
            If strTitle = "" Or strTitle <> strTemp(8) Then
               dblSkipPageRow = dblSkipPageRow + 1
               intCounter = intCounter + 2
            Else
               intCounter = intCounter + 1
            End If
            '一開始或資料已填滿一頁時跳頁
            If intCounter = 1 Or dblSkipPageRow >= 28 Then
               If dblSkipPageRow >= 28 Then
                  '換頁
                  wksAnnuity.Range("A" & intCounter).Select
                  wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
               End If
               dblSkipPageRow = 0
               Call PrintExcelTitle(strTemp()) 'Modify by Amy 2025/009/17 +strtemp()
            End If
            If strTitle = "" Or strTitle <> strTemp(8) Then
               wksAnnuity.Range("B" & intCounter).Value = strTemp(8)
            End If
            strTitle = strTemp(8)
            strText = IIf(strTemp(0) = "ZZZ", "台幣", strTemp(0)) & "小計："
            wksAnnuity.Range("C" & intCounter).Value = strText
            'Modify by Amy 2025/09/17 幣別 前插入一欄 未收規費NTD,所有欄位往後移
            wksAnnuity.Range("F" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("F" & intCounter).Value = strTemp(1)
            wksAnnuity.Range("G" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("G" & intCounter).Value = strTemp(2)
            wksAnnuity.Range("H" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("H" & intCounter).Value = strTemp(3)
            wksAnnuity.Range("I" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("I" & intCounter).Value = strTemp(4)
            wksAnnuity.Range("J" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("J" & intCounter).Value = strTemp(5)
            wksAnnuity.Range("K" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("K" & intCounter).Value = strTemp(6)
            wksAnnuity.Range("L" & intCounter).Select
            wksAnnuity.Application.Selection.NumberFormatLocal = "#,##0.00_ "
            wksAnnuity.Range("L" & intCounter).Value = strTemp(7)
            'end 2025/09/17
            dblSkipPageRow = dblSkipPageRow + 1
            .MoveNext
         Loop
      End With
   End If
   'Sindy 2014/8/6 END
   
   Set Rs = Nothing
End Sub
