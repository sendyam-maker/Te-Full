VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm050324 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "內商-國外FC帳款明細表"
   ClientHeight    =   3495
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5160
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5160
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "Y"
      Top             =   2010
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "2"
      Top             =   2340
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   1305
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   1
      Top             =   600
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Width           =   3555
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1572
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "結束(&X)"
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
      Left            =   2700
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   2730
      Width           =   2235
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
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
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      TabIndex        =   21
      Top             =   3150
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查詢/列印(&P)"
      Default         =   -1  'True
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
      TabIndex        =   10
      Top             =   2730
      Width           =   2235
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1572
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.查詢 2.報表)"
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
      Left            =   2490
      TabIndex        =   26
      Top             =   2370
      Width           =   1485
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "報表種類："
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
      TabIndex        =   25
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3015
      TabIndex        =   24
      Top             =   1290
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "請款對象："
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
      Left            =   270
      TabIndex        =   23
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   " 印表機"
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
      Left            =   420
      TabIndex        =   22
      Top             =   3180
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "(Y:是)"
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
      Left            =   2460
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "是否列印明細："
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
      Left            =   270
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label8 
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
      Left            =   2280
      TabIndex        =   18
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "國籍："
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
      Left            =   270
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   16
      Top             =   960
      Width           =   252
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "請款日期："
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
      Left            =   270
      TabIndex        =   15
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(1.請款 2.應收帳款)"
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
      TabIndex        =   14
      Top             =   600
      Width           =   2772
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "資料性質："
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
      Left            =   270
      TabIndex        =   13
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
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
      Left            =   270
      TabIndex        =   12
      Top             =   270
      Width           =   1215
   End
End
Attribute VB_Name = "frm050324"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

'copy by nickc 2006/06/14 from frmacc24i0
Public adoacc1k0 As New ADODB.Recordset
Public adoacc0y0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccrpt213 As New ADODB.Recordset
Dim strSql As String
Dim intCounter As Integer
Dim intPage As Integer
Dim intRecord As Integer
Dim intLength As Integer
Dim strAmount As String
Dim prnPrint As Printer
Dim strPrint As String


Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   If FormCheck = False Then
      MsgBox "條件不足或錯誤，請檢查！"
      Exit Sub
   End If
   If Me.Text1.Text = "" Then
      MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
      Exit Sub
   Else
      If Not CheckSysKind(Me.Text1.Text) Then
         Me.Text1.SetFocus
         Exit Sub
      End If
   End If
   'add by nickc 2006/07/07
   If Text7 = "" Then
      MsgBox "請輸入報表種類！", vbExclamation + vbOKOnly
      Text7.SetFocus
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   If Text7 = "2" Then
        For Each prnPrint In Printers
           If prnPrint.DeviceName = Combo1 Then
              Set Printer = prnPrint
           End If
        Next
        Printer.Orientation = 1
   End If
   Accrpt213Delete
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 清除查詢印表記錄檔欄位
   ProduceData
   If Text7 = "2" Then
        pub_QL05 = pub_QL05 & ";" & Label13 & "報表" 'Add By Sindy 2010/10/4
        PrintData
        'FormClear
   Else
        pub_QL05 = pub_QL05 & ";" & Label13 & "查詢" 'Add By Sindy 2010/10/4
        Me.Hide
        frm050324_1.Show
   End If
   If Text7 = "2" Then
        For Each prnPrint In Printers
           If prnPrint.DeviceName = strPrint Then
              Set Printer = prnPrint
           End If
        Next
   End If
   Screen.MousePointer = vbDefault
   StatusView MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      StatusView MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
'2011/10/13 modify by sonia
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 5250
'   Me.Height = 3900
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath4)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   MoveFormToCenter Me
'2011/10/13 end

   Text1 = Systemkind_g
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   strPrint = Printer.DeviceName
   For Each prnPrint In Printers
      If prnPrint.DeviceName <> Printer.DeviceName Then
         Combo1.AddItem prnPrint.DeviceName
      End If
      If Combo1 = "" Then
         Combo1 = prnPrint.DeviceName
      End If
   Next
   StatusView MsgText(102)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'edit by nickc 2007/02/08
   'Set Frmacc24d0 = Nothing
   Set frm050324 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Me.Text1.Text <> "" Then
      If Not CheckSysKind("" & Me.Text1.Text) Then
         Me.Text1.SetFocus
         Cancel = True
      End If
   End If
   If Cancel Then Text1_GotFocus
End Sub

'檢查輸入的系統類別是否超出使用者權限下所能使用的系統類別
Private Function CheckSysKind(strSysKind As String) As Boolean
Dim arr1
Dim arr2
Dim ii As Integer
Dim jj As Integer
Dim blnKind As Boolean '是否能使用此系統類別
   
   CheckSysKind = False
   If Systemkind_g <> "" Then
      arr1 = Split(Systemkind_g, ",")
      arr2 = Split(Me.Text1.Text, ",")
      For ii = LBound(arr2) To UBound(arr2)
         blnKind = False
         For jj = LBound(arr1) To UBound(arr1)
            If arr2(ii) = arr1(jj) Then
               blnKind = True
            End If
         Next jj
         If blnKind = False Then
            MsgBox "系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
            Exit Function
         End If
      Next ii
   End If
   CheckSysKind = True
End Function

Private Sub Text2_GotFocus(Index As Integer)
   If Index = 1 And Text2(1) = "" And Text2(0) <> "" Then
      Text2(1) = Text2(0)
   End If
   TextInverse Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
    Select Case Len(Text2(Index))
        Case 6
            Text2(Index) = AfterZero(Text2(Index))
        Case 8
            Text2(Index) = Text2(Index) & "0"
    End Select
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim arr1
Dim ii As Integer
Dim strSystemKind As String

On Error GoTo Checking
   strSql = ""
   If Text1 <> MsgText(601) Then
      strSystemKind = ""
      arr1 = Split(Me.Text1.Text, ",")
      For ii = LBound(arr1) To UBound(arr1)
         strSystemKind = strSystemKind & "'" & arr1(ii) & "',"
      Next ii
      strSystemKind = Left(strSystemKind, Len(strSystemKind) - 1)
      strSql = strSql & " AND A1K13 IN ( " & strSystemKind & " ) "
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1 'Add By Sindy 2010/10/4
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If (MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29)) Or _
      (MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29)) Then
      pub_QL05 = pub_QL05 & ";" & Label4 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/10/4
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and fa10 >= '" & Text4 & "'"
   End If
   If Text5 <> MsgText(601) Then
      strSql = strSql & " and fa10 <= '" & Text5 & "z'"
   End If
   If Text4 <> MsgText(601) Or Text5 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label6 & Text4 & "-" & Text5 'Add By Sindy 2010/10/4
   End If
   'Add by Morgan 2004/12/28 加請款對象
   If Text2(0) <> "" Then
      strSql = strSql & " and a1k28 >= '" & Text2(0) & "'"
   End If
   If Text2(1) <> "" Then
      strSql = strSql & " and a1k28 <= '" & Text2(1) & "'"
   End If
   If Text2(0) <> "" Or Text2(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label11 & Text2(0) & "-" & Text2(1) 'Add By Sindy 2010/10/4
   End If
   '2004/12/28
   
   StatusView MsgText(26)
   adoaccrpt213.CursorLocation = adUseClient
   adoaccrpt213.Open "select * from accrpt213", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Select Case Text3
      Case Mid(ComboItem(1), 1, 1)
         pub_QL05 = pub_QL05 & ";" & Label2 & "請款" 'Add By Sindy 2010/10/4
         Select1
      Case Mid(ComboItem(2), 1, 1)
         pub_QL05 = pub_QL05 & ";" & Label2 & "應收帳款" 'Add By Sindy 2010/10/4
         Select2
      Case Else
         pub_QL05 = pub_QL05 & ";" & Label2 & "請款" 'Add By Sindy 2010/10/4
         Select1
   End Select
   adoaccrpt213.Close
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
Private Sub Accrpt213Delete()
   adoTaie.Execute "delete from accrpt213 where r21301 = '" & strUserNum & "'"
End Sub

'*************************************************
'  選擇往來帳款統計
'
'*************************************************
Private Sub Select1()
Dim douExchange As Double

   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k12 is null or a1k12 = 0)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc1k0.EOF = False
      adoacc0y0.CursorLocation = adUseClient
      adoacc0y0.Open "select * from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0y0.RecordCount = 0 Then
         adoaccrpt213.AddNew
         ARSave
         adoaccrpt213.UpdateBatch
      Else
         Do While adoacc0y0.EOF = False
            adoaccrpt213.AddNew
            ARSave
            adoaccrpt213.Fields("r21317").Value = adoacc0y0.Fields("a0z01").Value
            If IsNull(adoacc0y0.Fields("a0y02").Value) Then
               adoaccrpt213.Fields("r21309").Value = Null
            Else
               adoaccrpt213.Fields("r21309").Value = adoacc0y0.Fields("a0y02").Value
            End If
            'Modify By Sindy 2012/10/22
'            If IsNull(adoacc0y0.Fields("a0z03").Value) Then
'               adoaccrpt213.Fields("r21310").Value = Null
'            Else
'               adoaccrpt213.Fields("r21310").Value = adoacc0y0.Fields("a0z03").Value
'            End If
            If IsNull(adoacc0y0.Fields("a0y03").Value) Then
               adoaccrpt213.Fields("r21310").Value = Null
            Else
               adoaccrpt213.Fields("r21310").Value = adoacc0y0.Fields("a0y03").Value
            End If
            '2012/10/22 End
            If IsNull(adoacc0y0.Fields("a0z04").Value) Then
               adoaccrpt213.Fields("r21311").Value = 0
            Else
               adoaccrpt213.Fields("r21311").Value = adoacc0y0.Fields("a0z04").Value
            End If
            If IsNull(adoacc0y0.Fields("a0y04").Value) Then
               douExchange = 0
            Else
               douExchange = adoacc0y0.Fields("a0y04").Value
            End If
            'adoaccrpt213.Fields("r21312").Value = Val(adoaccrpt213.Fields("r21311").Value) * douExchange
            adoaccrpt213.Fields("r21313").Value = adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
'因為之前抓一次申請人，並不用再抓一次彼所案號
'            adoquery.CursorLocation = adUseClient
'            adoquery.Open "select pa77 as Yno from patent where pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
'                          "select tm45 as Yno from trademark where tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
'                          "select lc23 as Yno from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
'                          "select sp27 as Yno from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'            If adoquery.RecordCount <> 0 Then
'               If IsNull(adoquery.Fields("Yno").Value) = False Then
'                  adoaccrpt213.Fields("r21314").Value = adoquery.Fields("Yno").Value
'               End If
'            End If
'            adoquery.Close
            'adoaccrpt213.Fields("r21315").Value = Val(adoaccrpt213.Fields("r21312").Value) / 1000
            adoaccrpt213.Fields("r21316").Value = Val(adoaccrpt213.Fields("r21306").Value) - Val(adoaccrpt213.Fields("r21312").Value)
            adoaccrpt213.UpdateBatch
            adoacc0y0.MoveNext
         Loop
      End If
      adoacc0y0.Close
      adoacc1k0.MoveNext
   Loop
   adoacc1k0.Close
End Sub

'*************************************************
'  選擇應收帳款統計
'
'*************************************************
Private Sub Select2()
Dim douExchange As Double

   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0)" & strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1k0.RecordCount = 0 Then
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc1k0.EOF = False
      adoaccrpt213.AddNew
      ARSave
      adoaccrpt213.UpdateBatch
      adoacc1k0.MoveNext
   Loop
   adoacc1k0.Close
End Sub

'*************************************************
'  請款資料儲存
'
'*************************************************
Private Sub ARSave()
   adoaccrpt213.Fields("r21301").Value = strUserNum
   adoaccrpt213.Fields("r21302").Value = adoacc1k0.Fields("a1k01").Value
   If IsNull(adoacc1k0.Fields("a1k02").Value) Then
      adoaccrpt213.Fields("r21303").Value = Null
   Else
      adoaccrpt213.Fields("r21303").Value = adoacc1k0.Fields("a1k02").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k18").Value) Then
      adoaccrpt213.Fields("r21304").Value = Null
   Else
      adoaccrpt213.Fields("r21304").Value = adoacc1k0.Fields("a1k18").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k08").Value) = False Then
      'Modify By Sindy 2013/1/15
      'adoaccrpt213.Fields("r21305").Value = Val(adoacc1k0.Fields("a1k08").Value)
      If IsNull(adoacc1k0.Fields("a1k31").Value) = True Then
         adoaccrpt213.Fields("r21305").Value = Val(adoacc1k0.Fields("a1k08").Value)
      Else
         adoaccrpt213.Fields("r21305").Value = Val(adoacc1k0.Fields("a1k08").Value) - Val(adoacc1k0.Fields("a1k31").Value)
      End If
      '2013/1/15 End
   Else
      adoaccrpt213.Fields("r21305").Value = 0
   End If
   If IsNull(adoacc1k0.Fields("a1k11").Value) Then
      adoaccrpt213.Fields("r21306").Value = 0
   Else
      adoaccrpt213.Fields("r21306").Value = adoacc1k0.Fields("a1k11").Value
   End If
   'If IsNull(adoacc1k0.Fields("a1k09").Value) = False Then
   '   adoaccrpt213.Fields("r21306").Value = Val(adoaccrpt213.Fields("r21306").Value) - Val(adoacc1k0.Fields("a1k09").Value)
   'End If
   If IsNull(adoacc1k0.Fields("a1k09").Value) Then
      adoaccrpt213.Fields("r21308").Value = 0
   Else
      adoaccrpt213.Fields("r21308").Value = adoacc1k0.Fields("a1k09").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k30").Value) Then
      adoaccrpt213.Fields("r21312").Value = 0
   Else
      adoaccrpt213.Fields("r21312").Value = adoacc1k0.Fields("a1k30").Value
   End If
   adoaccrpt213.Fields("r21315").Value = Val(Format((Val(adoaccrpt213.Fields("r21312").Value) - Val(adoaccrpt213.Fields("r21308").Value)) / 1000, FAmount))
   adoaccrpt213.Fields("r21307").Value = Val(Format((Val(adoaccrpt213.Fields("r21306").Value) - Val(adoaccrpt213.Fields("r21308").Value)) / 1000, FAmount))
   adoaccrpt213.Fields("r21313").Value = adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
   adoquery.CursorLocation = adUseClient
   '存放申請人
   adoquery.Open "select pa26 as Yno from patent where pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select tm23 as Yno from trademark where tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select lc11 as Yno from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select sp08 as Yno from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("Yno").Value) = False Then
         adoaccrpt213.Fields("r21314").Value = adoquery.Fields("Yno").Value
      End If
   End If
   adoquery.Close
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = Systemkind_g
   Text3 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text4 = ""
   Text5 = ""
   Text6 = ""
   'Add by Morgan 2004/12/28
   Text2(0) = "": Text2(1) = ""
   Text1.SetFocus
End Sub

'*************************************************
'  產生對帳資料
'
'*************************************************
Public Sub PrintData()
Dim strNo As String
Dim dblLin As Double '行數
Dim intRow As Integer 'Add By Sindy 2012/12/12
   
On Error GoTo Checking
   strSql = ""
   intCounter = 3
   intRecord = 1
   intPage = 0
   dblLin = 0
   '是否列印明細
   'If Text6 = MsgText(602) Then
      adoquery.CursorLocation = adUseClient
      strSql = "select R21302,R21313,R21304,R21305,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) as FAName,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CUName,R21303 from accrpt213, acc1k0, fagent,customer where r21302 = a1k01 and substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and substr(R21314,1,8)=cu01(+) and substr(R21314,9,1)=cu02(+) and r21301 = '" & strUserNum & "' order by r21301 asc, a1k01 asc"
      adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount > 0 Then
         InsertQueryLog (adoquery.RecordCount) 'Add By Sindy 2010/10/4
         intCounter = 3
         intRecord = 1
         intPage = intPage + 1
         dblLin = 0
         PrintHead
         Do While adoquery.EOF = False
            If dblLin >= 35 Then
               Printer.NewPage
               intCounter = 3
               intRecord = 1
               intPage = intPage + 1
               dblLin = 0
               PrintHead
            End If
            '請款編號
            Printer.CurrentX = 0
            Printer.CurrentY = 300 + intCounter * 300
            If IsNull(adoquery.Fields("R21302").Value) Then
               Printer.Print ""
            Else
               Printer.Print adoquery.Fields("R21302").Value
            End If
            '本所案號
            Printer.CurrentX = 1500
            Printer.CurrentY = 300 + intCounter * 300
            If IsNull(adoquery.Fields("R21313").Value) Then
               Printer.Print ""
            Else
               Printer.Print adoquery.Fields("R21313").Value
            End If
            '幣別
            Printer.CurrentX = 3300
            Printer.CurrentY = 300 + intCounter * 300
            If IsNull(adoquery.Fields("R21304").Value) Then
               Printer.Print ""
            Else
               Printer.Print adoquery.Fields("R21304").Value
            End If
            '請款金額
            If IsNull(adoquery.Fields("R21305").Value) = False Then
               strAmount = Format(Val(adoquery.Fields("R21305").Value), FDollar)
               intLength = Printer.TextWidth(strAmount)
               Printer.CurrentX = 4700 - intLength
               Printer.CurrentY = 300 + intCounter * 300
               Printer.Print strAmount
            End If
   
            '代理人
            Printer.CurrentX = 4800
            Printer.CurrentY = 300 + intCounter * 300
            If IsNull(adoquery.Fields("FAName").Value) Then
               Printer.Print ""
            Else
   
               Printer.Print Left(adoquery.Fields("FAName").Value, 10)
            End If
            '申請人
            Printer.CurrentX = 7300
            Printer.CurrentY = 300 + intCounter * 300
            If IsNull(adoquery.Fields("CUName").Value) Then
               Printer.Print ""
            Else
               Printer.Print Left(adoquery.Fields("CUName").Value, 10)
            End If
            '請款日期
            Printer.CurrentX = 9800
            Printer.CurrentY = 300 + intCounter * 300
            If IsNull(adoquery.Fields("R21303").Value) Then
               Printer.Print ""
            Else
               Printer.Print CFDate(adoquery.Fields("R21303").Value)
            End If
            intCounter = intCounter + 1
            intRecord = intRecord + 1
            dblLin = dblLin + 1
            adoquery.MoveNext
         Loop
         Printer.CurrentX = 6800
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print "共 " & Trim(adoquery.RecordCount) & " 筆"
         
         adoquery.Close
         adoquery.CursorLocation = adUseClient
         'Modify By Sindy 2012/12/12
         'adoquery.Open "select sum(r21305) from accrpt213 where r21301 = '" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select r21304,sum(r21305) from accrpt213 where r21301 = '" & strUserNum & "' group by r21304 order by r21304", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            adoquery.MoveFirst
            intRow = 0
            Do While adoquery.EOF = False
               intRow = intRow + 1
               Printer.CurrentX = 8800
               Printer.CurrentY = 300 + intCounter * 300
               'Printer.Print "TOTAL：USD " & Format(Val(CheckStr(adoquery.Fields(0).Value)), FDollar)
               If intRow = 1 Then
                  Printer.Print "TOTAL：" & adoquery.Fields(0).Value & " " & Format(Val(CheckStr(adoquery.Fields(1).Value)), FDollar)
               Else
                  Printer.Print "                " & adoquery.Fields(0).Value & " " & Format(Val(CheckStr(adoquery.Fields(1).Value)), FDollar)
               End If
               intCounter = intCounter + 1
               adoquery.MoveNext
            Loop
         '2012/12/12 End
         End If
         adoquery.Close
         Printer.EndDoc
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/10/4
      End If
   'End If
   
   'PrintSum
   'Printer.EndDoc
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
   Printer.FontSize = 14
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("***  內商-國外FC帳款明細表  ***") / 2)
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "***  內商-國外FC帳款明細表  ***"
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期: " & CFDate(ACDate(ServerDate))) - 500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "系統類別: " & Me.Text1.Text
   Printer.CurrentX = 16000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "頁次: " & intPage
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "帳款日期: " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text & " " & IIf(Text3 = "2", "(應收帳款)", "(請款)")
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "請款編號"
   Printer.CurrentX = 1500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "本所案號"
   Printer.CurrentX = 3300
   Printer.CurrentY = 300 + intCounter * 300
   'Modify By Sindy 2012/12/7
   'Printer.Print "請款美金"
   Printer.Print "請款外幣"
   '2012/12/7 End
   Printer.CurrentX = 4800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "代理人"
   Printer.CurrentX = 7300
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "申請人"
   Printer.CurrentX = 9800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "請款日期"
   Printer.Line (0, 300 + intCounter * 300 + 350)-(19500 - 1000, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Add by Morgan 2004/12/28 加請款對象
   If Left(Text2(0), 6) <> Left(Text2(1), 6) Then
      FormCheck = False
      MsgBox "請款對象前六碼必須相同！"
      Text2(0).SetFocus
      Exit Function
   End If
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
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text5 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'add by nickc 2006/07/07
Private Sub Text7_GotFocus()
    TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 And KeyAscii <> 23 Then
        KeyAscii = 0
    End If
End Sub


