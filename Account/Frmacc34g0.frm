VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc34g0 
   AutoRedraw      =   -1  'True
   Caption         =   "兌現日別資金流動彙總表"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   0
      Top             =   90
      Width           =   3525
   End
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
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
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   500
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
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   500
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
      Left            =   840
      TabIndex        =   7
      Top             =   3360
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
      Left            =   3480
      TabIndex        =   8
      Top             =   3360
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
      Left            =   840
      TabIndex        =   9
      Top             =   3720
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
      Left            =   3480
      TabIndex        =   10
      Top             =   3720
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
      Left            =   840
      TabIndex        =   11
      Top             =   4080
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
      Left            =   3480
      TabIndex        =   12
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
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   1230
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   855
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
      Left            =   3120
      TabIndex        =   4
      Top             =   855
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
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   2880
      TabIndex        =   20
      Top             =   495
      Width           =   255
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   495
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "兌現日期"
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
      Left            =   240
      TabIndex        =   18
      Top             =   855
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
      Left            =   2880
      TabIndex        =   17
      Top             =   855
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3720
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
      Left            =   240
      TabIndex        =   16
      Top             =   3000
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
      Left            =   600
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   2880
      Picture         =   "Frmacc34g0.frx":0000
      Stretch         =   -1  'True
      Top             =   3360
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
      Left            =   600
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   2880
      Picture         =   "Frmacc34g0.frx":0442
      Stretch         =   -1  'True
      Top             =   3720
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
      Left            =   600
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   2880
      Picture         =   "Frmacc34g0.frx":0884
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1695
      Left            =   120
      Top             =   2880
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "Frmacc34g0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc0e0 As New ADODB.Recordset
Public adoaccrpt306 As New ADODB.Recordset
Public adoacc0h0 As New ADODB.Recordset
Public adoacc0g0 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoacc0b0 As New ADODB.Recordset
Dim strSort1 As String
Dim strSort2 As String
Dim strSort3 As String
Dim dllaccrpt306 As Object
Dim strSql As String
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
        MsgBox Label6 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17

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
         Text1.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
Dim bCancel As Boolean
   
   'Add By Sindy 2020/4/23
   If CboCmp.Text = MsgText(601) Then
      MsgBox Label6 & MsgText(52), , MsgText(5)
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
   Accrpt306Delete
   ProduceData
   adoaccrpt306.CursorLocation = adUseClient
   adoaccrpt306.Open "select * from accrpt306", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt306.RecordCount <> 0 Then
      '20140120START By eric
      dllaccrpt306.Acc34g0 ReportTitle(306) & "-" & strCmpN, Text1, Text2, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      'dllaccrpt306.Acc34g0 ReportTitle(306), Text1, Text2, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      '20140120END
   End If
   adoaccrpt306.Close
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
   'Modify by Amy 2023/10/12 原W5250 H2100
   Me.Width = 5280
   Me.Height = 2460
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
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   Combo8 = MsgText(1)
   Combo1.AddItem ComboItem(181)
   Combo1.AddItem ComboItem(182)
   ComboAdd
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt306 = CreateObject("AccReport.ReportSelect")
   
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
   Set dllaccrpt306 = Nothing
   Set Frmacc34g0 = Nothing
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
   strSort1 = "銀行代號"
   strSort2 = "銀行帳號"
   strSort3 = "帳號名稱"
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
   Combo3.AddItem strSort3
   Combo5.AddItem strSort1
   Combo5.AddItem strSort2
   Combo5.AddItem strSort3
   Combo7.AddItem strSort1
   Combo7.AddItem strSort2
   Combo7.AddItem strSort3
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1 As String
Dim strOrder2 As String
Dim strOrder3 As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim intYear As Integer
Dim intMonth As Integer
   
On Error GoTo Checking
   adoacc0b0.CursorLocation = adUseClient
   '20140120START Modify By eric
   adoacc0b0.Open "select * from acc0b0 where a0b04 = '" & strCmp & "' ", adoTaie, adOpenStatic, adLockReadOnly
   'adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
   '20140120END
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
   strSql = ""
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0h01 asc"
         Else
            strOrder1 = " order by a0h01 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0h02 asc"
         Else
            strOrder1 = " order by a0h02 desc"
         End If
      Case strSort3
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0h03 asc"
         Else
            strOrder1 = " order by a0h03 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0h01 asc"
         Else
            strOrder2 = ", a0h01 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0h02 asc"
         Else
            strOrder2 = ", a0h02 desc"
         End If
      Case strSort3
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0h03 asc"
         Else
            strOrder2 = ", a0h03 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo7
      Case strSort1
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0h01 asc"
         Else
            strOrder3 = ", a0h01 desc"
         End If
      Case strSort2
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0h02 asc"
         Else
            strOrder3 = ", a0h02 desc"
         End If
      Case strSort3
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0h03 asc"
         Else
            strOrder3 = ", a0h03 desc"
         End If
      Case Else
         strOrder3 = MsgText(601)
   End Select
   
   If Text1 <> MsgText(601) Then
      strSQL1 = " and a0h02 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSQL1 = strSQL1 & " and a0h02 <= '" & Text2 & "'"
   End If
   If strSQL1 <> "" Then
      strSQL1 = " where " & Mid(strSQL1, 5, Len(strSQL1) - 4)
   End If
   adoaccrpt306.CursorLocation = adUseClient
   adoaccrpt306.Open "select * from accrpt306", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select * from acc0h0" & strSQL1 & strOrder1 & strOrder2 & strOrder3, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0h0.RecordCount = 0 Then
      adoacc0h0.Close
      adoaccrpt306.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0h0.EOF = False
      adoaccrpt306.AddNew
      adoaccrpt306.Fields("r30601").Value = strUserNum
      adoaccrpt306.Fields("r30602").Value = adoacc0h0.Fields("a0h01").Value
      adoaccrpt306.Fields("r30603").Value = adoacc0h0.Fields("a0h02").Value
      adoacc0g0.CursorLocation = adUseClient
      adoacc0g0.Open "SELECT A0G02 FROM ACC0G0 WHERE A0G01 = '" & adoacc0h0.Fields("A0H01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0g0.RecordCount <> 0 Then
         If IsNull(adoacc0g0.Fields(0).Value) = False Then
            adoaccrpt306.Fields("r30604").Value = adoacc0g0.Fields(0).Value
         End If
      End If
      adoacc0g0.Close
      adoacc040.CursorLocation = adUseClient
      '20140120START Modify By eric
      adoacc040.Open "SELECT A0408 FROM ACC040 WHERE A0401 = " & intYear & " AND A0402 = " & intMonth & " AND A0403 ='" & strCmp & "' AND A0404 = '" & MsgText(55) & "' AND A0405 = '" & adoacc0h0.Fields("A0H08").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      'adoacc040.Open "SELECT A0408 FROM ACC040 WHERE A0401 = " & intYear & " AND A0402 = " & intMonth & " AND A0403 = '1' AND A0404 = '" & MsgText(55) & "' AND A0405 = '" & adoacc0h0.Fields("A0H08").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      '20140120END
      If adoacc040.RecordCount <> 0 Then
         If IsNull(adoacc040.Fields(0).Value) Then
            adoaccrpt306.Fields("R30605").Value = 0
         Else
            adoaccrpt306.Fields("R30605").Value = adoacc040.Fields(0).Value
         End If
      Else
         adoaccrpt306.Fields("R30605").Value = 0
      End If
      adoacc040.Close
      strSql = ""
      
      '20140120START Modify By eric
      strSql = " and a0e23 = '" & strCmp & "' "
      If Text1 <> MsgText(601) Then
         strSql = strSql & " and a0e20 >= '" & Text1 & "'"
      End If
      'If Text1 <> MsgText(601) Then
      '   strSql = " and a0e20 >= '" & Text1 & "'"
      'End If
      '20140120END
            
      If Text2 <> MsgText(601) Then
         strSql = strSql & " and a0e20 <= '" & Text2 & "'"
      End If
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         strSql = strSql & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      End If
      adoacc0e0.CursorLocation = adUseClient
      adoacc0e0.Open "select sum(a0e11) from acc0e0 where a0e19 = '" & adoacc0h0.Fields("a0h01").Value & "' and a0e20 = '" & adoacc0h0.Fields("a0h02").Value & "' and a0e04 = '" & MsgText(18) & "' and a0e21 = 0 and a0e15 = 0 and a0e17 = 0 and a0e34 = 0 and (a0e14 <> 0 and a0e14 is not null)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) Then
            adoaccrpt306.Fields("r30606").Value = 0
         Else
            adoaccrpt306.Fields("r30606").Value = Val(adoacc0e0.Fields(0).Value)
         End If
      Else
         adoaccrpt306.Fields("r30606").Value = 0
      End If
      adoacc0e0.Close
      strSql = ""
      If Text1 <> MsgText(601) Then
         strSql = " and a0e07 >= '" & Text1 & "'"
      End If
      If Text2 <> MsgText(601) Then
         strSql = strSql & " and a0e07 <= '" & Text2 & "'"
      End If
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         strSql = strSql & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      End If
      adoacc0e0.CursorLocation = adUseClient
      adoacc0e0.Open "select sum(a0e11) from acc0e0 where a0e01 = '" & adoacc0h0.Fields("a0h01").Value & "' and a0e07 = '" & adoacc0h0.Fields("a0h02").Value & "' and a0e04 = '" & MsgText(19) & "' and a0e37 = 0 and a0e25 = 0" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) Then
            adoaccrpt306.Fields("r30607").Value = 0
         Else
            adoaccrpt306.Fields("r30607").Value = Val(adoacc0e0.Fields(0).Value)
         End If
      Else
         adoaccrpt306.Fields("r30607").Value = 0
      End If
      adoacc0e0.Close
      If IsNull(adoaccrpt306.Fields("r30605").Value) Then
         adoaccrpt306.Fields("r30608").Value = Val(adoaccrpt306.Fields("r30606").Value) - Val(adoaccrpt306.Fields("r30607").Value)
      Else
         adoaccrpt306.Fields("r30608").Value = Val(adoaccrpt306.Fields("r30605").Value) + Val(adoaccrpt306.Fields("r30606").Value) - Val(adoaccrpt306.Fields("r30607").Value)
      End If
      adoaccrpt306.UpdateBatch
      adoacc0h0.MoveNext
   Loop
   adoacc0h0.Close
   adoaccrpt306.Close
   adoTaie.Execute "delete from accrpt306 where r30602 is null"
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
Private Sub Accrpt306Delete()
   adoTaie.Execute "delete from accrpt306"
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
'   Label11 = ""     '20140120ADD By eric
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

'Mark by Sindy 2020/4/23 公司別改下拉式選單
''20140120START ADD By eric
'Private Sub Text3_LostFocus()
'   If Text3.Text = "" Then
'      MsgBox "公司別不可空白 !"
'      Text3.SetFocus
'      Exit Sub
'   End If
'   If Text3.Text <> "1" And Text3.Text <> "2" Then
'      MsgBox "公司別僅能為 1 或 2 !"
'      Text3.Text = ""
'      Text3.SetFocus
'      Exit Sub
'   End If
'End Sub
''20140120END
'
''20140120START ADD By eric
'Private Sub Text3_GotFocus()
'   TextInverse Text3
'   CloseIme
'End Sub
''20140120END
'
''20140120START ADD By eric
'Private Sub Text3_Change()
'       Label11.Caption = A0802Query(IIf(Text3 = "2", "J", "1"))
'End Sub
''20140120END
