VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24i0 
   AutoRedraw      =   -1  'True
   Caption         =   "國外FC帳款明細表"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   5835
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
      Height          =   315
      Left            =   1305
      TabIndex        =   6
      Top             =   1710
      Width           =   1215
   End
   Begin VB.TextBox Text8 
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
      Left            =   2880
      TabIndex        =   7
      Top             =   1710
      Width           =   1215
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
      Top             =   3360
      Width           =   3495
   End
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
      TabIndex        =   10
      Top             =   2490
      Width           =   612
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
      TabIndex        =   9
      Top             =   2100
      Width           =   855
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
      TabIndex        =   8
      Top             =   2100
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      TabIndex        =   11
      Top             =   2880
      Width           =   4692
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
      Text            =   "ALL"
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
   Begin VB.Label Label1 
      Caption         =   "業務區說明："
      Height          =   180
      Index           =   13
      Left            =   4200
      TabIndex        =   30
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "　外商：F10~F19"
      Height          =   180
      Index           =   14
      Left            =   4320
      TabIndex        =   29
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "　外專：F20~F29"
      Height          =   180
      Index           =   15
      Left            =   4320
      TabIndex        =   28
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "　外法：F30~F49"
      Height          =   180
      Index           =   16
      Left            =   4320
      TabIndex        =   27
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "業務區："
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
      Index           =   1
      Left            =   360
      TabIndex        =   26
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label13 
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
      Left            =   2640
      TabIndex        =   25
      Top             =   1680
      Width           =   255
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
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   " 印表機："
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
      Left            =   300
      TabIndex        =   22
      Top             =   3360
      Width           =   975
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
      Left            =   2520
      TabIndex        =   20
      Top             =   2490
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
      Left            =   360
      TabIndex        =   19
      Top             =   2490
      Width           =   1455
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
      Top             =   2100
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
      Left            =   360
      TabIndex        =   17
      Top             =   2100
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
      Height          =   252
      Left            =   360
      TabIndex        =   15
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(1.往來帳款 2.應收帳款)"
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
      Height          =   252
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   972
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
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc24i0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

'Add By Sindy 2018/7/3 取小數3位; 因要同Frmacc24c0報表數字
Const FDollar As String = "###,###,###,###.000"
Const FAmount As String = "0.000"

Public adoacc1k0 As New ADODB.Recordset
Public adoacc0y0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccrpt213 As New ADODB.Recordset
Dim strSql As String
Dim strSQLCP As String     '2007/12/6 ADD BY SONIA
Dim intCounter As Integer
Dim intPage As Integer
Dim intRecord As Integer
Dim intLength As Integer
Dim strAmount As String
Dim prnPrint As Printer
Dim strPrinter As String
Dim prnstrPos As Integer   '2008/1/4 ADD BY SONIA  報表起始位置
Dim bolPrint As Boolean 'Added by Lydia 2018/02/21 是否列印


Private Sub Command2_Click()

   If FormCheck = False Then
      'Modify by Morgan 2004/12/28
      'MsgBox MsgText(181), , MsgText(5)
      MsgBox "條件不足或錯誤，請檢查！"
      Exit Sub
   End If
   If Me.Text1.Text = "" Then
      MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
      Exit Sub
   Else
'2007/11/29 cancel by sonia
'      If Not CheckSysKind(Me.Text1.Text) Then
'         Me.Text1.SetFocus
'         Exit Sub
'      End If
'2007/11/29 end
   End If
   Screen.MousePointer = vbHourglass
   For Each prnPrint In Printers
      If prnPrint.DeviceName = Combo1 Then
         Set Printer = prnPrint
      End If
   Next
   bolPrint = False 'Added by Lydia 2018/02/21
   Accrpt213Delete
   ProduceData
   'Modified by Lydia 2018/02/21
   'PrintData
   If bolPrint = True Then PrintData
   FormClear
   For Each prnPrint In Printers
      If prnPrint.DeviceName = strPrinter Then
         Set Printer = prnPrint
      End If
   Next
   Screen.MousePointer = vbDefault
   StatusView MsgText(102)
End Sub

Private Sub Form_Activate()
   '93.3.16 ADD BY SONIA
   If IsObject(mdiMain) Then
      ToolShow
   End If
   '93.3.16 END

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
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5955
   Me.Height = 4230
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Text1 = Systemkind_g    '2007/11/29 cancel by sonia
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   
   'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   
   StatusView MsgText(102)
   prnstrPos = 500 '2008/1/5 ADD BY SONIA
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   'edit by nickc 2007/02/08
   'Set Frmacc24d0 = Nothing
   Set Frmacc24i0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
'2007/11/29 MODIFY by sonia
'   If Me.Text1.Text <> "" Then
'      If Not CheckSysKind("" & Me.Text1.Text) Then
'         Me.Text1.SetFocus
'         Cancel = True
'      End If
'   End If
   If Me.Text1.Text <> "ALL" Then
      If Not CheckSysKind1("" & Me.Text1.Text) Then
         Me.Text1.SetFocus
         Cancel = True
      End If
   End If
'2007/11/29 end
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

'2007/12/10 add by sonia
'檢查輸入的系統類別是否正確
Private Function CheckSysKind1(strSysKind As String) As Boolean
Dim arr1
Dim arr2
Dim ii As Integer
Dim jj As Integer
   
   CheckSysKind1 = False
   arr2 = Split(Me.Text1.Text, ",")
   For ii = LBound(arr2) To UBound(arr2)
      If CheckSys(arr2(ii)) = "" Then
         MsgBox "系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
         Exit Function
      End If
   Next ii
   CheckSysKind1 = True
End Function
'2007/12/10 end

Private Sub Text2_GotFocus(Index As Integer)
   If Index = 1 And Text2(1) = "" And Text2(0) <> "" Then
      Text2(1) = Text2(0)
   End If
   TextInverse Text2(Index)
   CloseIme
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

'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim arr1
Dim ii As Integer
Dim strSystemKind As String
   
On Error GoTo Checking
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/22 清除查詢印表記錄檔欄位
   strSql = ""
   strSQLCP = ""                '2007/12/6 ADD BY SONIA
   strSystemKind = ""
   '2007/11/29 modify by sonia
   'If Text1 <> MsgText(601) Then
   If Text1 <> "ALL" Then
      arr1 = Split(Me.Text1.Text, ",")
      For ii = LBound(arr1) To UBound(arr1)
         strSystemKind = strSystemKind & "'" & arr1(ii) & "',"
      Next ii
      strSystemKind = Left(strSystemKind, Len(strSystemKind) - 1)
      strSql = strSql & " AND A1K13 IN ( " & strSystemKind & " ) "
   End If
   If Trim(Text1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Text1 'Add By Sindy 2010/12/22
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If (MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29)) Or _
      (MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29)) Then
      pub_QL05 = pub_QL05 & ";" & Label4 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/22
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and fa10 >= '" & Text4 & "'"
   End If
   If Text5 <> MsgText(601) Then
      strSql = strSql & " and fa10 <= '" & Text5 & "z'"
   End If
   If Text4 <> MsgText(601) Or Text5 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label6 & Text4 & "-" & Text5 'Add By Sindy 2010/12/22
   End If
   'Add by Morgan 2004/12/28 加請款對象
   If Text2(0) <> "" Then
      strSql = strSql & " and a1k28 >= '" & Text2(0) & "'"
   End If
   If Text2(1) <> "" Then
      strSql = strSql & " and a1k28 <= '" & Text2(1) & "'"
   End If
   If Text2(0) <> "" Or Text2(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label11(0) & Text2(0) & "-" & Text2(1) 'Add By Sindy 2010/12/22
   End If
   '2004/12/28
   '2007/12/6 add by sonia 加業務區
   If Text7 <> "" Then
      strSQLCP = " and CP12 >= '" & Text7 & "'"
   End If
   If Text8 <> "" Then
      strSQLCP = strSQLCP & " and CP12 <= '" & Text8 & "'"
   End If
   If Text7 <> "" Or Text8 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label11(1) & Text7 & "-" & Text8 'Add By Sindy 2010/12/22
   End If
   '2007/12/6 end
   If Trim(Text3) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label2 & "1.往來帳款" 'Add By Sindy 2010/12/22
   ElseIf Trim(Text3) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Label2 & "2.應收帳款" 'Add By Sindy 2010/12/22
   End If
   If Trim(Text6) = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Label9 & Text6 'Add By Sindy 2010/12/22
   End If
   
   StatusView MsgText(26)
   adoaccrpt213.CursorLocation = adUseClient
   adoaccrpt213.Open "select * from accrpt213", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Select Case Text3
      Case Mid(ComboItem(1), 1, 1)
         Select1
      Case Mid(ComboItem(2), 1, 1)
         Select2
      Case Else
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
Dim strA1k01 As String 'Add By Sindy 2018/7/2

   adoacc1k0.CursorLocation = adUseClient
   '2007/12/7 modify by sonia 同一請款單計入最後收文之智權人員,作廢或銷帳都不抓,請款金額抓台幣-折讓金額*匯率
   'adoacc1k0.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k12 is null or a1k12 = 0)" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
   '2009/4/28 modify by sonia 外幣,台幣都要扣除折讓
   'adoacc1k0.Open "select * from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k02,a1k08,a1k09,a1k11,a1k13,a1k14,a1k15,a1k16,a1k18,a1k30 from acc1k0, fagent, caseprogress " & _
                  "where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01=cp60(+) " & strSQL & _
                  " group by a1k01,a1k02,a1k08,a1k09,a1k11,a1k13,a1k14,a1k15,a1k16,a1k18,a1k30) new where cp09 in substr(new.cp,9,9)" & strSQLCP, adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2012/10/11 計算外幣金額round((a1k08 - nvl(a1k06, 0)),2)=>round((a1k08 - nvl(a1k31, 0)),2)
   '                           round((a1k11 - nvl(a1k06, 0) * a1k10),2)=>round((a1k11 - nvl(a1k06, 0)),2)
'   adoacc1k0.Open "select * from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k02,round((a1k08 - nvl(a1k06, 0)),2) as a1k08,a1k09,decode(nvl(a1k30,0),0,round((a1k11 - nvl(a1k06, 0) * a1k10),2),round((a1k11 - nvl(a1k06, 0) * a1k10),0)) as a1k11,a1k13,a1k14,a1k15,a1k16,a1k18,a1k30 from acc1k0, fagent, caseprogress " & _
'                  "where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01=cp60(+) " & strSql & _
'                  " group by a1k01,a1k02,round((a1k08 - nvl(a1k06, 0)),2),a1k09,decode(nvl(a1k30,0),0,round((a1k11 - nvl(a1k06, 0) * a1k10),2),round((a1k11 - nvl(a1k06, 0) * a1k10),0)),a1k13,a1k14,a1k15,a1k16,a1k18,a1k30) new where cp09 in substr(new.cp,9,9)" & strSQLCP, adoTaie, adOpenStatic, adLockReadOnly
   adoacc1k0.Open "select * from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k02,round((a1k08 - nvl(a1k31, 0)),2) as a1k08,a1k09,decode(nvl(a1k30,0),0,round((a1k11 - nvl(a1k06, 0)),2),round((a1k11 - nvl(a1k06, 0)),0)) as a1k11,a1k13,a1k14,a1k15,a1k16,a1k18,a1k30 from acc1k0, fagent, caseprogress " & _
                  "where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01=cp60(+) " & strSql & _
                  " group by a1k01,a1k02,round((a1k08 - nvl(a1k31, 0)),2),a1k09,decode(nvl(a1k30,0),0,round((a1k11 - nvl(a1k06, 0)),2),round((a1k11 - nvl(a1k06, 0)),0)),a1k13,a1k14,a1k15,a1k16,a1k18,a1k30) new" & _
                  " where cp09 in substr(new.cp,9,9)" & strSQLCP, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      bolPrint = True 'Added by Lydia 2018/02/21
      InsertQueryLog (adoacc1k0.RecordCount) 'Add By Sindy 2010/12/22
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
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select pa77 as Yno from patent where pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                          "select tm45 as Yno from trademark where tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                          "select lc23 as Yno from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                          "select sp27 as Yno from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               If IsNull(adoquery.Fields("Yno").Value) = False Then
                  adoaccrpt213.Fields("r21314").Value = adoquery.Fields("Yno").Value
               End If
            End If
            adoquery.Close
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
   'Add By Sindy 2018/7/2
   '有分次收款狀況,R21305,R21306,R21308只記錄一次金額即可,不然會重覆計算 ex:X10700372
   adoacc1k0.Open "select * from accrpt213" & _
                  " where r21301 = '" & strUserNum & "'" & _
                  " and r21302 in(select r21302 from accrpt213" & _
                  " where r21301 = '" & strUserNum & "' and r21317 is not null" & _
                  " group by r21302" & _
                  " having count(*)>1)" & _
                  " order by r21302 asc,r21309 asc,r21317 asc", adoTaie, adOpenStatic, adLockReadOnly
   strA1k01 = ""
   If adoacc1k0.RecordCount > 0 Then
      adoacc1k0.MoveFirst
      Do While Not adoacc1k0.EOF
         If strA1k01 = adoacc1k0.Fields("r21302") Then
            strSql = "update accrpt213 set r21305=null,r21306=null,r21308=null,r21307=null" & _
                     " where r21301 = '" & strUserNum & "'" & _
                     " and r21302='" & strA1k01 & "' and r21317='" & adoacc1k0.Fields("r21317") & "'"
            cnnConnection.Execute strSql
         End If
         strA1k01 = adoacc1k0.Fields("r21302")
         adoacc1k0.MoveNext
      Loop
   End If
   adoacc1k0.Close
   '2018/7/2 END
   
   '分配點數資料
   '",R21214=(select max('*') from acc1n0 where a1n01=R21302 and a1n02='1' and a1n04<>r21318)"
   strSql = "update accrpt213 set r21307=(select sum(a1n05) from acc1n0 where a1n01=R21302 and a1n02='1' and a1n04=r21318)" & _
      " where R21301='" & strUserNum & "' and exists(select * from acc1n0 where a1n01=R21302 and a1n02='1')" & _
      " and nvl(R21306,0)>0"
   cnnConnection.Execute strSql, intI
End Sub

'*************************************************
'  選擇應收帳款統計
'
'*************************************************
Private Sub Select2()
Dim douExchange As Double

   adoacc1k0.CursorLocation = adUseClient
   '2007/12/7 modify by sonia 同一請款單計入最後收文之智權人員,作廢或銷帳都不抓,請款金額抓台幣-折讓金額*匯率
   'adoacc1k0.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0)" & strSQL, adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2009/4/28 modify by sonia 外幣,台幣都要扣除折讓
   'adoacc1k0.Open "select * from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k02,a1k08,a1k09,a1k11,a1k13,a1k14,a1k15,a1k16,a1k18,a1k30 from acc1k0, fagent, caseprogress " & _
                  "where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01=cp60(+) " & strSQL & _
                  " group by a1k01,a1k02,a1k08,a1k09,a1k11,a1k13,a1k14,a1k15,a1k16,a1k18,a1k30) new where cp09 in substr(new.cp,9,9)" & strSQLCP, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify By Sindy 2012/10/11 計算外幣金額round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)=>round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0))/a1k10,1)
   '                                       round((a1k11 - nvl(a1k06, 0) * a1k10),1)=>round((a1k11 - nvl(a1k06, 0)),1)
   '                                       round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0))=>round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),0))
'   adoacc1k0.Open "select * from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k02,round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1) as a1k08,a1k09,decode(nvl(a1k30,0),0,round((a1k11 - nvl(a1k06, 0) * a1k10),1),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as a1k11,a1k13,a1k14,a1k15,a1k16,a1k18,a1k30 from acc1k0, fagent, caseprogress " & _
'                  "where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01=cp60(+) " & strSql & _
'                  " group by a1k01,a1k02,round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1),a1k09,decode(nvl(a1k30,0),0,round((a1k11 - nvl(a1k06, 0) * a1k10),1),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)),a1k13,a1k14,a1k15,a1k16,a1k18,a1k30) new where cp09 in substr(new.cp,9,9)" & strSQLCP, adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1k0.Open "select * from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k02,round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0))/a1k10,1) as a1k08,a1k09,decode(nvl(a1k30,0),0,round((a1k11 - nvl(a1k06, 0)),1),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),0)) as a1k11,a1k13,a1k14,a1k15,a1k16,a1k18,a1k30 from acc1k0, fagent, caseprogress " & _
                  "where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01=cp60(+) " & strSql & _
                  " group by a1k01,a1k02,round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0))/a1k10,1),a1k09,decode(nvl(a1k30,0),0,round((a1k11 - nvl(a1k06, 0)),1),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),0)),a1k13,a1k14,a1k15,a1k16,a1k18,a1k30) new where cp09 in substr(new.cp,9,9)" & strSQLCP, adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1k0.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      bolPrint = True 'Added by Lydia 2018/02/21
      InsertQueryLog (adoacc1k0.RecordCount) 'Add By Sindy 2010/12/22
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
Dim dblRate As Double
   
   adoaccrpt213.Fields("r21301").Value = strUserNum
   adoaccrpt213.Fields("r21302").Value = adoacc1k0.Fields("a1k01").Value
   If IsNull(adoacc1k0.Fields("a1k02").Value) Then
      adoaccrpt213.Fields("r21303").Value = Null
   Else
      adoaccrpt213.Fields("r21303").Value = adoacc1k0.Fields("a1k02").Value
   End If
   '請款幣別
   If IsNull(adoacc1k0.Fields("a1k18").Value) Then
      adoaccrpt213.Fields("r21304").Value = Null
   Else
      adoaccrpt213.Fields("r21304").Value = adoacc1k0.Fields("a1k18").Value
   End If
   '外幣金額
   'Modify By Sindy 2012/10/11 A1K08改為存放各請款幣別的請款金額
'   'Modify By Sindy 2010/3/12
'   If Not IsNull(adoacc1k0.Fields("a1k18").Value) And adoacc1k0.Fields("a1k18").Value <> "USD" Then
'      dblRate = PUB_GetUSXRate_1(adoacc1k0.Fields("a1k02").Value, adoacc1k0.Fields("a1k18").Value)
'      If IsNull(adoacc1k0.Fields("a1k11").Value) = False Then
'         adoaccrpt213.Fields("r21305").Value = Format(((Val(adoacc1k0.Fields("a1k11").Value) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
'      Else
'         adoaccrpt213.Fields("r21305").Value = 0
'      End If
'   Else
'   '2010/3/12 End
      If IsNull(adoacc1k0.Fields("a1k08").Value) = False Then
         adoaccrpt213.Fields("r21305").Value = Val(adoacc1k0.Fields("a1k08").Value)
      Else
         adoaccrpt213.Fields("r21305").Value = 0
      End If
'   End If
   '台幣金額
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
   'Modify By Sindy 2018/7/5 同acc24c0計算方式
   'adoaccrpt213.Fields("r21307").Value = Val(Format((Val(adoaccrpt213.Fields("r21306").Value) - Val(adoaccrpt213.Fields("r21308").Value)) / 1000, FAmount))
   adoaccrpt213.Fields("r21307").Value = (Val(Format(adoaccrpt213.Fields("r21306").Value, DAmount)) - Val(Format(adoaccrpt213.Fields("r21308").Value, DAmount))) / 1000
   '2018/7/5 END
   adoaccrpt213.Fields("r21313").Value = adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select pa77 as Yno from patent where pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select tm45 as Yno from trademark where tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select lc23 as Yno from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select sp27 as Yno from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("Yno").Value) = False Then
         adoaccrpt213.Fields("r21314").Value = adoquery.Fields("Yno").Value
      End If
   End If
   adoquery.Close
   'Add By Sindy 2018/7/4
   If IsNull(adoacc1k0.Fields("cp13").Value) Then
      adoaccrpt213.Fields("r21318").Value = Null
      adoaccrpt213.Fields("r21319").Value = Null
   Else
      adoaccrpt213.Fields("r21318").Value = adoacc1k0.Fields("cp13").Value
      adoaccrpt213.Fields("r21319").Value = StaffQuery(adoacc1k0.Fields("cp13").Value)
   End If
   '2018/7/4 END
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Text1 = Systemkind_g   '2007/11/29 cancel by sonia
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

On Error GoTo Checking
   strSql = ""
   intCounter = 3
   intRecord = 1
   intPage = 0
   dblLin = 0
   
   'Added by Morgan 2019/9/3 設定美國標準紙張直印
   Printer.PaperSize = PUB_GetPaperSize(15)
   Printer.Orientation = 1 '直印
   'end 2019/9/3
   
   '是否列印明細
   If Text6 = MsgText(602) Then
      adoquery.CursorLocation = adUseClient
      strSql = "select * from accrpt213, acc1k0, fagent where r21302 = a1k01 and substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and r21301 = '" & strUserNum & "' order by r21301 asc, a1k01 asc"
      adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount > 0 Then
         intCounter = 3
         intRecord = 1
         intPage = intPage + 1
         dblLin = 0
         PrintHead
      End If
      Do While adoquery.EOF = False
         
         'Modified by Morgan 2019/9/3
         'If dblLin >= 35 Then
         If Printer.CurrentY + 2 * Printer.TextHeight("測試") > Printer.ScaleHeight Then
         'end 2019/9/3
            Printer.NewPage
            intCounter = 3
            intRecord = 1
            intPage = intPage + 1
            dblLin = 0
            PrintHead
         End If
         '請款編號
         Printer.CurrentX = prnstrPos + 0
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21302").Value) Then
            Printer.Print ""
         Else
            Printer.Print adoquery.Fields("R21302").Value
         End If
         '請款日期
         Printer.CurrentX = prnstrPos + 1500
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21303").Value) Then
            Printer.Print ""
         Else
            Printer.Print CFDate(adoquery.Fields("R21303").Value)
         End If
         '幣別
         Printer.CurrentX = prnstrPos + 2800
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
            Printer.CurrentX = prnstrPos + 5000 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         'Add By Sindy 2018/7/2
         Else
            strAmount = "-"
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 5000 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         '2018/7/2 END
         End If
         '台幣請款
         If IsNull(adoquery.Fields("R21306").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21306").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 6500 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         'Add By Sindy 2018/7/2
         Else
            strAmount = "-"
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 6500 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         '2018/7/2 END
         End If
         '規費
         If IsNull(adoquery.Fields("R21308").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21308").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 8000 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         'Add By Sindy 2018/7/2
         Else
            strAmount = "-"
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 8000 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         '2018/7/2 END
         End If
         '收款日期
         Printer.CurrentX = prnstrPos + 8100
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21309").Value) Then
            Printer.Print ""
         Else
            Printer.Print CFDate(adoquery.Fields("R21309").Value)
         End If
         '幣別
         Printer.CurrentX = prnstrPos + 9400
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21310").Value) Then
            Printer.Print ""
         Else
            Printer.Print adoquery.Fields("R21310").Value
         End If
         '收款金額
         If IsNull(adoquery.Fields("R21311").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21311").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 11600 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         '台幣收款
         If IsNull(adoquery.Fields("R21312").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21312").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 13100 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         '本所案號
         Printer.CurrentX = prnstrPos + 13200
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21313").Value) Then
            Printer.Print ""
         Else
            Printer.Print adoquery.Fields("R21313").Value
         End If
         '彼所案號
         Printer.CurrentX = prnstrPos + 15000
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21314").Value) Then
            Printer.Print ""
         Else
            'Modify By Cheng 2003/03/07
            '只取前10碼
'            Printer.Print adoquery.Fields("R21314").Value
            Printer.Print Left(adoquery.Fields("R21314").Value, 10)
         End If
         '點數
         If IsNull(adoquery.Fields("R21307").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21307").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 17900 - intLength - 500
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         '差額
'         If IsNull(adoquery.Fields("R21316").Value) = False Then
            strAmount = Format(Val("" & adoquery.Fields("R21306").Value) - Val("" & adoquery.Fields("r21312").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 19500 - intLength - 1000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
'         End If
         intCounter = intCounter + 1
         intRecord = intRecord + 1
         dblLin = dblLin + 1
         adoquery.MoveNext
      Loop
      adoquery.Close
      Printer.NewPage
   End If
   PrintSum
   Printer.EndDoc
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
   Printer.CurrentX = prnstrPos + 7000
   Printer.CurrentY = 300 + intCounter * 300
   '2007/12/7 modify by sonia 表頭依業務區之部門決定
   'Printer.Print ReportTitle(2131)
   Select Case Mid(Text7, 1, 2)
      Case "F1"
         Printer.Print ReportTitle(2133)
      Case "F2"
         Printer.Print ReportTitle(2132)
      Case "F3"
         Printer.Print ReportTitle(2134)
      Case Else
         Printer.Print ReportTitle(2131)
   End Select
   '2007/12/7 end
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = prnstrPos + 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = prnstrPos + 16000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   Printer.CurrentX = prnstrPos + 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "系統類別: " & Me.Text1.Text
   Printer.CurrentX = prnstrPos + 16000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "頁次: " & intPage
   intCounter = intCounter + 1
   Printer.CurrentX = prnstrPos + 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "帳款日期: " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text & " " & IIf(Text3 = "2", "(應收帳款)", "(往來帳款)")
   intCounter = intCounter + 1
   Printer.CurrentX = prnstrPos + 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "請款編號"
   Printer.CurrentX = prnstrPos + 1500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "請款日期"
   Printer.CurrentX = prnstrPos + 2800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = prnstrPos + 3600
   Printer.CurrentY = 300 + intCounter * 300
   'Modify By Sindy 2010/3/12
   'Printer.Print "請款金額(USD)"
   Printer.Print "外幣金額"
   Printer.CurrentX = prnstrPos + 5500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "台幣請款"
   Printer.CurrentX = prnstrPos + 7200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "規費"
   Printer.CurrentX = prnstrPos + 8100
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "收款日期"
   Printer.CurrentX = prnstrPos + 9400
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = prnstrPos + 10200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "收款金額"
   Printer.CurrentX = prnstrPos + 11700
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "台幣收款"
   Printer.CurrentX = prnstrPos + 13200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "本所案號"
   Printer.CurrentX = prnstrPos + 15000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "彼所案號"
   Printer.CurrentX = prnstrPos + 17000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "點數"
   Printer.CurrentX = prnstrPos + 18000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "差額"
   'Modified by Lydia 2018/02/21 改X位置
   'Printer.Line (0, 300 + intCounter * 300 + 350)-(19500 - 1000, 300 + intCounter * 300 + 350)
   Printer.Line (500, 300 + intCounter * 300 + 350)-(20000 - 1000, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'*************************************************
' 合計位置
'
'*************************************************
Private Sub PrintSum()
Dim dblTotPoint As Double 'Add By Sindy 2013/5/3
   
   intCounter = 2
   intPage = intPage + 1
   Printer.FontSize = 14
   Printer.CurrentX = prnstrPos + 7000
   Printer.CurrentY = 300 + intCounter * 300
    'Modify By Cheng 2003/03/07
'   Printer.Print ReportTitle(213)
   '2007/12/7 modify by sonia 表頭依業務區之部門決定
   'Printer.Print ReportTitle(2131)
   Select Case Mid(Text7, 1, 2)
      Case "F1"
         Printer.Print ReportTitle(2133)
      Case "F2"
         Printer.Print ReportTitle(2132)
      Case "F3"
         Printer.Print ReportTitle(2134)
      Case Else
         Printer.Print ReportTitle(2131)
   End Select
   '2007/12/7 end
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = prnstrPos + 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = prnstrPos + 16000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   Printer.CurrentX = prnstrPos + 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "系統類別: " & Me.Text1.Text
   intCounter = intCounter + 1
   Printer.CurrentX = prnstrPos + 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "帳款日期: " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text & " " & IIf(Text3 = "2", "(應收帳款)", "(往來帳款)")
   Printer.CurrentX = prnstrPos + 16000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "頁次: " & intPage
   intCounter = 0
   Printer.CurrentX = prnstrPos + 0
   Printer.CurrentY = 3000
   Printer.Print "PS:"
   dblTotPoint = 0 'Add By Sindy 2013/5/3
   intCounter = intCounter + 2
   adoquery.CursorLocation = adUseClient
   'Modify By Sindy 2010/3/12 增加請款幣別判斷-多筆
   'adoquery.Open "select sum(r21305), sum(r21306), sum(round((r21306 - r21308) / 1000, 2)), sum(r21311), sum(r21312), sum(decode(r21312, 0, 0, round((r21312 - r21308) / 1000, 2))) from accrpt213 where r21301 = '" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2018/7/2 + nvl(xx,0)
   'Modify By Sindy 2018/7/3 取小數3位; 因要同Frmacc24c0報表數字 sum(round((nvl(r21306,0) - nvl(r21308,0)) / 1000, 2)) ==> sum(round((nvl(r21306,0) - nvl(r21308,0)) / 1000, 3))
   'adoquery.Open "select r21304,sum(nvl(r21305,0)), sum(r21311), sum(round((nvl(r21306,0) - nvl(r21308,0)) / 1000, 3)) from accrpt213 where r21301 = '" & strUserNum & "' group by r21304 ", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select r21304, sum(nvl(r21305,0)), sum(r21311), round(sum(r21307), 3) from accrpt213 where r21301 = '" & strUserNum & "' group by r21304 ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      With adoquery
      .MoveFirst
      Do While Not .EOF
         Printer.CurrentX = prnstrPos + 0
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print adoquery.Fields(0).Value & "請款金額合計:"
         If IsNull(adoquery.Fields(1).Value) = False Then
            strAmount = Format(Val(adoquery.Fields(1).Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 3900 - intLength
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print strAmount
         End If
         Printer.CurrentX = prnstrPos + 4100
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print adoquery.Fields(0).Value & "已收款合計:"
         If IsNull(adoquery.Fields(2).Value) = False Then
            strAmount = Format(Val(adoquery.Fields(2).Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 7900 - intLength
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print strAmount
         End If
         Printer.CurrentX = prnstrPos + 8100
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print "請款點數:"
         If IsNull(adoquery.Fields(3).Value) = False Then
            strAmount = Format(Val(adoquery.Fields(3).Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = prnstrPos + 10700 - intLength
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print strAmount
            dblTotPoint = dblTotPoint + CDbl(strAmount) 'Add By Sindy 2013/5/3
         End If
         intCounter = intCounter + 2
         .MoveNext
      Loop
      End With
   End If
   adoquery.Close
   'Add By Sindy 2013/5/3
   If dblTotPoint > 0 Then
      Printer.CurrentX = prnstrPos + 7600 '8100
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print "請款點數合計:"
      intLength = Printer.TextWidth(Format(dblTotPoint, FDollar))
      Printer.CurrentX = prnstrPos + 10700 - intLength
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print Format(dblTotPoint, FDollar)
      intCounter = intCounter + 2
   End If
   '2013/5/3 End
   'Modify By Sindy 2018/7/3 取小數3位; 因要同Frmacc24c0報表數字 sum(decode(r21312, 0, 0, round((r21312 - r21308) / 1000, 2))) ==> sum(decode(r21312, 0, 0, round((r21312 - r21308) / 1000, 3)))
   adoquery.Open "select sum(r21306), sum(r21312), sum(decode(r21312, 0, 0, round((r21312 - r21308) / 1000, 3))) from accrpt213 where r21301 = '" & strUserNum & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      Printer.CurrentX = prnstrPos + 0
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print "台幣請款金額合計:"
      If IsNull(adoquery.Fields(0).Value) = False Then
         strAmount = Format(Val(adoquery.Fields(0).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = prnstrPos + 3900 - intLength
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print strAmount
      End If
      Printer.CurrentX = prnstrPos + 4100
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print "台幣已收款合計:"
      If IsNull(adoquery.Fields(1).Value) = False And adoquery.Fields(1).Value <> 0 Then
         strAmount = Format(Val(adoquery.Fields(1).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = prnstrPos + 7900 - intLength
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print strAmount
      End If
      Printer.CurrentX = prnstrPos + 8100
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print "收款點數:"
      If IsNull(adoquery.Fields(2).Value) = False And adoquery.Fields(2).Value <> 0 Then
         strAmount = Format(Val(adoquery.Fields(2).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = prnstrPos + 10700 - intLength
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print strAmount
      End If
      '2007/11/29 add by sonia
      intCounter = intCounter + 4
      Printer.CurrentX = prnstrPos + 0
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print "PS: 外幣/台幣已收款,係指當期之請款到列印報表日已收回者."
      '2007/11/29 end
   End If
   adoquery.Close
End Sub

'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

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
    
'2007/11/29 cancel by sonia
'   If Text1 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'2007/11/29 end
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

Private Sub Text6_GotFocus()
   CloseIme
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   CloseIme
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
