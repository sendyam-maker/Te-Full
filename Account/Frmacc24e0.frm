VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24e0 
   AutoRedraw      =   -1  'True
   Caption         =   "國外應收規費、服務費分析表"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   5565
   Begin VB.TextBox Text10 
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
      Left            =   2670
      TabIndex        =   19
      Top             =   1395
      Width           =   852
   End
   Begin VB.TextBox Text9 
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
      Left            =   1470
      TabIndex        =   18
      Top             =   1395
      Width           =   852
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   2400
      Width           =   3450
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
      Height          =   300
      Left            =   2670
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1050
      Width           =   852
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
      Left            =   1470
      MaxLength       =   3
      TabIndex        =   8
      Top             =   1050
      Width           =   852
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
      Left            =   450
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   1890
      Width           =   4692
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
      Height          =   300
      Left            =   1470
      TabIndex        =   2
      Top             =   690
      Width           =   612
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
      Left            =   2070
      TabIndex        =   3
      Top             =   690
      Width           =   612
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
      Left            =   2670
      TabIndex        =   4
      Top             =   690
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
      Left            =   3270
      TabIndex        =   5
      Top             =   690
      Width           =   612
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
      Left            =   3870
      TabIndex        =   6
      Top             =   690
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
      Left            =   4470
      TabIndex        =   7
      Top             =   690
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1470
      TabIndex        =   0
      Top             =   330
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
      Left            =   3390
      TabIndex        =   1
      Top             =   330
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
      Left            =   3630
      TabIndex        =   22
      Top             =   1395
      Width           =   1785
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
      Left            =   2430
      TabIndex        =   21
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "請款對象國籍"
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
      Left            =   60
      TabIndex        =   20
      Top             =   1395
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      TabIndex        =   17
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   510
      TabIndex        =   16
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Left            =   2430
      TabIndex        =   15
      Top             =   1050
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
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
      Left            =   510
      TabIndex        =   14
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "請款日期"
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
      Left            =   510
      TabIndex        =   13
      Top             =   330
      Width           =   975
   End
   Begin VB.Label Label6 
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
      Left            =   3150
      TabIndex        =   12
      Top             =   330
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc24e0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoaccrpt214 As New ADODB.Recordset
Dim dllaccrpt214 As Object
'Add By Sindy 2012/5/23
Dim iLine As Integer
Dim pLeft(0 To 5) As Integer
Dim strPrinter As String
'2012/5/23 End


Private Sub Command2_Click()
   If FormCheck = False Then
      'Modify By Sindy 2012/5/23 Mark
      'MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2012/5/23
'   Accrpt214Delete
'   ProduceData
'   dllaccrpt214.Acc24e0 ReportTitle(214), StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   Call PrintData
   '2012/5/23 End
   FormClear
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
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
   Me.Width = 5685
   Me.Height = 3345
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
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add by Sindy 2012/5/25
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Set dllaccrpt214 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   'Add By Sindy 2012/5/25
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   Set dllaccrpt214 = Nothing
   Set Frmacc24e0 = Nothing
End Sub

'Add By Sindy 2012/5/22
Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'2012/5/22 End

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

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

''*************************************************
''  產生報表資料
''
''*************************************************
'Private Sub ProduceData()
'Dim strAnd, strYear As String
'Dim intCounter As Integer
'Dim strSql As String
'
'On Error GoTo Checking
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   '94.2.3 ADD BY SONIA 抓未銷帳未作廢未結清的資料
'   strSql = strSql & " And A1K25 Is Null AND (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) "
'   '94.2.3 END
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   adoaccrpt214.CursorLocation = adUseClient
'   adoaccrpt214.Open "select * from accrpt214", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   strYear = MsgText(601)
'   'edit by nickc 2007/02/08
'   'If Text2 <> MsgText(601) Or Text3 <> MsgText(601) Or Text4 <> MsgText(601) Or Text5 <> MsgText(601) Or Text6 <> MsgText(601) Or Text10 <> MsgText(601) Then
'   If Text2 <> MsgText(601) Or Text3 <> MsgText(601) Or Text4 <> MsgText(601) Or Text5 <> MsgText(601) Or Text6 <> MsgText(601) Then
'      strAnd = " and ("
'      If Text2 <> MsgText(601) Then
'         strYear = strYear & "a1k13 = '" & Text2 & "' or "
'      End If
'      If Text3 <> MsgText(601) Then
'         strYear = strYear & "a1k13 = '" & Text3 & "' or "
'      End If
'      If Text4 <> MsgText(601) Then
'         strYear = strYear & "a1k13 = '" & Text4 & "' or "
'      End If
'      If Text5 <> MsgText(601) Then
'         strYear = strYear & "a1k13 = '" & Text5 & "' or "
'      End If
'      If Text6 <> MsgText(601) Then
'         strYear = strYear & "a1k13 = '" & Text6 & "' or "
'      End If
'      If Text7 <> MsgText(601) Then
'         strYear = strYear & "a1k13 = '" & Text7 & "' or "
'      End If
'      strYear = Mid(strYear, 1, Len(strYear) - 4) & ") "
'   Else
'      strAnd = MsgText(601)
'   End If
'   For intCounter = Val(Mid(MaskEdBox1.Text, 1, 3)) To Val(Mid(MaskEdBox2.Text, 1, 3))
'      adoacc1k0.CursorLocation = adUseClient
'      '94.2.3 MODIFY BY SONIA
'      'adoacc1k0.Open "select sum(a1k08), sum(a1k11), sum(a1k09), sum(a1k11 - a1k09) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & strSQL & strAnd & strYear, adoTaie, adOpenStatic, adLockReadOnly
'      adoacc1k0.Open "select sum(a1k08 - nvl(a1k06, 0)) - Sum(Nvl(A0Z04,0)), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), sum(a1k09), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09) from acc1k0,ACC0Z0 where A1K01=A0Z02(+) AND decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & strSql & strAnd & strYear, adoTaie, adOpenStatic, adLockReadOnly
'      '94.2.3 END
'      If adoacc1k0.RecordCount <> 0 Then
'         adoaccrpt214.AddNew
'         adoaccrpt214.Fields("r21401").Value = strUserNum
'         adoaccrpt214.Fields("r21402").Value = intCounter
'         If IsNull(adoacc1k0.Fields(0).Value) Then
'            adoaccrpt214.Fields("r21403").Value = 0
'         Else
'            adoaccrpt214.Fields("r21403").Value = adoacc1k0.Fields(0).Value
'         End If
'         If IsNull(adoacc1k0.Fields(1).Value) Then
'            adoaccrpt214.Fields("r21404").Value = 0
'         Else
'            adoaccrpt214.Fields("r21404").Value = adoacc1k0.Fields(1).Value
'         End If
'         If IsNull(adoacc1k0.Fields(2).Value) Then
'            adoaccrpt214.Fields("r21405").Value = 0
'         Else
'            adoaccrpt214.Fields("r21405").Value = adoacc1k0.Fields(2).Value
'         End If
'         If IsNull(adoacc1k0.Fields(3).Value) Then
'            adoaccrpt214.Fields("r21406").Value = 0
'         Else
'            adoaccrpt214.Fields("r21406").Value = adoacc1k0.Fields(3).Value
'         End If
'      Else
'         adoaccrpt214.Fields("r21403").Value = 0
'         adoaccrpt214.Fields("r21404").Value = 0
'         adoaccrpt214.Fields("r21405").Value = 0
'         adoaccrpt214.Fields("r21406").Value = 0
'      End If
'      adoacc1k0.Close
'      adoacc1k0.CursorLocation = adUseClient
'      '94.2.3 MODIFY BY SONIA
'      'adoacc1k0.Open "select sum(a1k11) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(173) & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'      adoacc1k0.Open "select sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), sum(a1k09), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(173) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      '94.2.3 END
'      If adoacc1k0.RecordCount <> 0 Then
'         If IsNull(adoacc1k0.Fields(0).Value) Then
'            adoaccrpt214.Fields("r21407").Value = 0
'         Else
'            adoaccrpt214.Fields("r21407").Value = adoacc1k0.Fields(0).Value
'         End If
'         '94.2.3 ADD BY SONIA
'         If IsNull(adoacc1k0.Fields(1).Value) Then
'            adoaccrpt214.Fields("r21413").Value = 0
'         Else
'            adoaccrpt214.Fields("r21413").Value = adoacc1k0.Fields(1).Value
'         End If
'         If IsNull(adoacc1k0.Fields(2).Value) Then
'            adoaccrpt214.Fields("r21414").Value = 0
'         Else
'            adoaccrpt214.Fields("r21414").Value = adoacc1k0.Fields(2).Value
'         End If
'         '94.2.3 END
'      Else
'         adoaccrpt214.Fields("r21407").Value = 0
'         '94.2.3 ADD BY SONIA
'         adoaccrpt214.Fields("r21413").Value = 0
'         adoaccrpt214.Fields("r21414").Value = 0
'         '94.2.3 END
'      End If
'      adoacc1k0.Close
'      adoacc1k0.CursorLocation = adUseClient
'      '94.2.3 MODIFY BY SONIA
'      'adoacc1k0.Open "select sum(a1k11) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(174) & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'      adoacc1k0.Open "select sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), sum(a1k09), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(174) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      '94.2.3 END
'      If adoacc1k0.RecordCount <> 0 Then
'         If IsNull(adoacc1k0.Fields(0).Value) Then
'            adoaccrpt214.Fields("r21408").Value = 0
'         Else
'            adoaccrpt214.Fields("r21408").Value = adoacc1k0.Fields(0).Value
'         End If
'         '94.2.3 ADD BY SONIA
'         If IsNull(adoacc1k0.Fields(1).Value) Then
'            adoaccrpt214.Fields("r21415").Value = 0
'         Else
'            adoaccrpt214.Fields("r21415").Value = adoacc1k0.Fields(1).Value
'         End If
'         If IsNull(adoacc1k0.Fields(2).Value) Then
'            adoaccrpt214.Fields("r21416").Value = 0
'         Else
'            adoaccrpt214.Fields("r21416").Value = adoacc1k0.Fields(2).Value
'         End If
'         '94.2.3 END
'      Else
'         adoaccrpt214.Fields("r21408").Value = 0
'         '94.2.3 ADD BY SONIA
'         adoaccrpt214.Fields("r21415").Value = 0
'         adoaccrpt214.Fields("r21416").Value = 0
'         '94.2.3 END
'      End If
'      adoacc1k0.Close
'      adoacc1k0.CursorLocation = adUseClient
'      '94.2.3 MODIFY BY SONIA
'      'adoacc1k0.Open "select sum(a1k11) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(175) & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'      adoacc1k0.Open "select sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), sum(a1k09), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(175) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      '94.2.3 END
'      If adoacc1k0.RecordCount <> 0 Then
'         If IsNull(adoacc1k0.Fields(0).Value) Then
'            adoaccrpt214.Fields("r21409").Value = 0
'         Else
'            adoaccrpt214.Fields("r21409").Value = adoacc1k0.Fields(0).Value
'         End If
'         '94.2.3 ADD BY SONIA
'         If IsNull(adoacc1k0.Fields(1).Value) Then
'            adoaccrpt214.Fields("r21417").Value = 0
'         Else
'            adoaccrpt214.Fields("r21417").Value = adoacc1k0.Fields(1).Value
'         End If
'         If IsNull(adoacc1k0.Fields(2).Value) Then
'            adoaccrpt214.Fields("r21418").Value = 0
'         Else
'            adoaccrpt214.Fields("r21418").Value = adoacc1k0.Fields(2).Value
'         End If
'         '94.2.3 END
'      Else
'         adoaccrpt214.Fields("r21409").Value = 0
'         '94.2.3 ADD BY SONIA
'         adoaccrpt214.Fields("r21417").Value = 0
'         adoaccrpt214.Fields("r21418").Value = 0
'         '94.2.3 END
'      End If
'      adoacc1k0.Close
'      adoacc1k0.CursorLocation = adUseClient
'      '94.2.3 MODIFY BY SONIA
'      'adoacc1k0.Open "select sum(a1k11) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(176) & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'      adoacc1k0.Open "select sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), sum(a1k09), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(176) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      '94.2.3 END
'      If adoacc1k0.RecordCount <> 0 Then
'         If IsNull(adoacc1k0.Fields(0).Value) Then
'            adoaccrpt214.Fields("r21410").Value = 0
'         Else
'            adoaccrpt214.Fields("r21410").Value = adoacc1k0.Fields(0).Value
'         End If
'         '94.2.3 ADD BY SONIA
'         If IsNull(adoacc1k0.Fields(1).Value) Then
'            adoaccrpt214.Fields("r21419").Value = 0
'         Else
'            adoaccrpt214.Fields("r21419").Value = adoacc1k0.Fields(1).Value
'         End If
'         If IsNull(adoacc1k0.Fields(2).Value) Then
'            adoaccrpt214.Fields("r21420").Value = 0
'         Else
'            adoaccrpt214.Fields("r21420").Value = adoacc1k0.Fields(2).Value
'         End If
'         '94.2.3 END
'      Else
'         adoaccrpt214.Fields("r21410").Value = 0
'         '94.2.3 ADD BY SONIA
'         adoaccrpt214.Fields("r21419").Value = 0
'         adoaccrpt214.Fields("r21420").Value = 0
'         '94.2.3 END
'      End If
'      adoacc1k0.Close
'      adoacc1k0.CursorLocation = adUseClient
'      '94.2.3 MODIFY BY SONIA
'      'adoacc1k0.Open "select sum(a1k11) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(177) & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'      adoacc1k0.Open "select sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), sum(a1k09), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 = '" & ComboItem(177) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      '94.2.3 END
'      If adoacc1k0.RecordCount <> 0 Then
'         If IsNull(adoacc1k0.Fields(0).Value) Then
'            adoaccrpt214.Fields("r21411").Value = 0
'         Else
'            adoaccrpt214.Fields("r21411").Value = adoacc1k0.Fields(0).Value
'         End If
'         '94.2.3 ADD BY SONIA
'         If IsNull(adoacc1k0.Fields(1).Value) Then
'            adoaccrpt214.Fields("r21421").Value = 0
'         Else
'            adoaccrpt214.Fields("r21421").Value = adoacc1k0.Fields(1).Value
'         End If
'         If IsNull(adoacc1k0.Fields(2).Value) Then
'            adoaccrpt214.Fields("r21422").Value = 0
'         Else
'            adoaccrpt214.Fields("r21422").Value = adoacc1k0.Fields(2).Value
'         End If
'         '94.2.3 END
'      Else
'         adoaccrpt214.Fields("r21411").Value = 0
'         '94.2.3 ADD BY SONIA
'         adoaccrpt214.Fields("r21421").Value = 0
'         adoaccrpt214.Fields("r21422").Value = 0
'         '94.2.3 END
'      End If
'      adoacc1k0.Close
'      '94.2.3 ADD BY SONIA
'      adoacc1k0.CursorLocation = adUseClient
'      adoacc1k0.Open "select sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))), sum(a1k09), sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09) from acc1k0 where decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2)) = " & intCounter & " and a1k13 NOT IN ('" & ComboItem(173) & "','" & ComboItem(174) & "','" & ComboItem(175) & "','" & ComboItem(176) & "','" & ComboItem(177) & "') " & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc1k0.RecordCount <> 0 Then
'         If IsNull(adoacc1k0.Fields(0).Value) Then
'            adoaccrpt214.Fields("r21412").Value = 0
'         Else
'            adoaccrpt214.Fields("r21412").Value = adoacc1k0.Fields(0).Value
'         End If
'         If IsNull(adoacc1k0.Fields(1).Value) Then
'            adoaccrpt214.Fields("r21423").Value = 0
'         Else
'            adoaccrpt214.Fields("r21423").Value = adoacc1k0.Fields(1).Value
'         End If
'         If IsNull(adoacc1k0.Fields(2).Value) Then
'            adoaccrpt214.Fields("r21424").Value = 0
'         Else
'            adoaccrpt214.Fields("r21424").Value = adoacc1k0.Fields(2).Value
'         End If
'      Else
'         adoaccrpt214.Fields("r21412").Value = 0
'         adoaccrpt214.Fields("r21423").Value = 0
'         adoaccrpt214.Fields("r21424").Value = 0
'      End If
'      adoacc1k0.Close
'      '94.2.3 END
'      adoaccrpt214.UpdateBatch
'   Next intCounter
'   adoaccrpt214.Close
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

''*************************************************
''  刪除報表資料
''
''*************************************************
'Private Sub Accrpt214Delete()
'   adoTaie.Execute "delete from accrpt214"
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   'Add By Sindy 2012/5/23
   Text1 = ""
   Text8 = ""
   '2012/5/23 End
   MaskEdBox1.SetFocus
   'Add by Amy 2017/10/24
   Text9 = ""
   Text10 = ""
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Modify By Sindy 2012/5/23
   FormCheck = True
   If MaskEdBox1.Text = MsgText(29) Then
      FormCheck = False
      MsgBox "請輸入請款日期(起)，不可空白！", , MsgText(5)
      MaskEdBox1.SetFocus
      Exit Function
   End If
   If MaskEdBox2.Text = MsgText(29) Then
      FormCheck = False
      MsgBox "請輸入請款日期(迄)，不可空白！", , MsgText(5)
      MaskEdBox2.SetFocus
      Exit Function
   End If
'   If Text2 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If Text3 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If Text4 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If Text5 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If Text6 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If Text7 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   FormCheck = False
End Function

'Add By Sindy 2012/5/23
'Mark by Amy 2017/10/24 因加國籍串代理人檔及客戶檔會很慢,故改寫至暫存檔
Private Sub PrintData_Old()
'   Dim strCon As String
'   Dim rsReport As ADODB.Recordset
'   Dim intCounter As Integer
'   Dim i As Integer
'   Dim dblSum As Double, dblSum1 As Double, dblSum2 As Double, dblSum3 As Double, dblSum4 As Double
'
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   strSql = ""
'   '請款日期
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   '抓未銷帳未作廢未結清的資料
'   strSql = strSql & " And A1K25 Is Null AND (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0)"
'   '系統別
'   If Text2 <> MsgText(601) Or Text3 <> MsgText(601) Or Text4 <> MsgText(601) Or Text5 <> MsgText(601) Or Text6 <> MsgText(601) Then
'      strSql = strSql & " and ("
'      If Text2 <> MsgText(601) Then
'         strCon = strCon & "a1k13 = '" & Text2 & "' or "
'      End If
'      If Text3 <> MsgText(601) Then
'         strCon = strCon & "a1k13 = '" & Text3 & "' or "
'      End If
'      If Text4 <> MsgText(601) Then
'         strCon = strCon & "a1k13 = '" & Text4 & "' or "
'      End If
'      If Text5 <> MsgText(601) Then
'         strCon = strCon & "a1k13 = '" & Text5 & "' or "
'      End If
'      If Text6 <> MsgText(601) Then
'         strCon = strCon & "a1k13 = '" & Text6 & "' or "
'      End If
'      If Text7 <> MsgText(601) Then
'         strCon = strCon & "a1k13 = '" & Text7 & "' or "
'      End If
'      strSql = strSql & Mid(strCon, 1, Len(strCon) - 4) & ")"
'   End If
'   '業務區
'   If Text1 <> "" Then
'      strSql = strSql & " and cp12>='" & Text1 & "'"
'   End If
'   If Text8 <> "" Then
'      strSql = strSql & " and cp12<='" & Text8 & "'"
'   End If
'
'   '檢查有無資料
'   'Modify By Sindy 2012/12/10 sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))) sum1 ==> sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) sum1
'   '                           sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09) ==> sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09)
'   strExc(0) = "select a1k13,sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) sum1, sum(a1k09), sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09)" & _
'               " from (" & _
'               "select distinct acc1k0.*" & _
'               " From acc1k0,caseprogress" & _
'               " where A1K01=CP60(+)" & strSql & _
'               ") a" & _
'               " group by a1k13" & _
'               " order by sum1 desc"
'   intI = 1
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'   If intI <> 1 Then
'      MsgBox "查無資料！"
'      Exit Sub
'   End If
'
'   PUB_RestorePrinter Combo1
'   Printer.Orientation = 1 '1.直印 2.橫印
'   Printer.Font.Name = "細明體"
'   iLine = 1
'   PrintHead
'   For intCounter = Val(Mid(MaskEdBox1.Text, 1, 3)) To Val(Mid(MaskEdBox2.Text, 1, 3))
'      'Modify By Sindy 2012/11/07
''      strExc(0) = "select " & intCounter & ", nvl(sum(a1k08 - nvl(a1k06, 0)) - Sum(Nvl(A0Z04,0)),0), nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))),0), nvl(sum(a1k09),0), nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09),0)" & _
''                  " from ACC0Z0,(" & _
''                  "select distinct acc1k0.*" & _
''                  " From acc1k0,caseprogress" & _
''                  " where A1K01=CP60(+)" & _
''                  " AND decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2))=" & intCounter & strSql & _
''                  ") a" & _
''                  " where A1K01=A0Z02(+)"
'      'Modify By Sindy 2012/12/10 sum(a1k08 - nvl(a1k06, 0)) A1 ==> sum(a1k08 - nvl(a1k31, 0)) A1
'      '                           nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))),0) A2 ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) A2
'      '                           nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09),0) A4 ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0) A4
'      '                           +,a1k18
'      strExc(0) = "select strYear,a1k18,nvl(sum(A1-Z1),0),sum(A2),sum(A3),sum(A4) from(" & _
'                  " select '" & intCounter & "' strYear,a1k18,sum(a1k08 - nvl(a1k31, 0)) A1,0 Z1, nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) A2, nvl(sum(a1k09),0) A3, nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0) A4 from (select distinct acc1k0.* From acc1k0,caseprogress where A1K01=CP60(+) AND decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2))=" & intCounter & strSql & ") a" & _
'                  " group by substr(a1k02,1,length(a1k02)-4),a1k18" & _
'                  " Union" & _
'                  " select '" & intCounter & "' strYear,a1k18,0 A1,Sum(Nvl(A0Z04,0)) Z1,0 A2,0 A3,0 A4 from ACC0Z0,ACC0Y0,(select distinct acc1k0.* From acc1k0,caseprogress where A1K01=CP60(+) AND decode(length(a1k02), 7, substr(a1k02, 1, 3), 6, substr(a1k02, 1, 2))=" & intCounter & strSql & ") a" & _
'                  " where A1K01=A0Z02(+) and A0Z01=A0Y01" & _
'                  " group by substr(a1k02,1,length(a1k02)-4),a1k18" & _
'                  " ) group by strYear,a1k18"
'      '2012/11/07 End
'      intI = 1
'      Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         With rsReport
'            Do While Not .EOF
'               If iLine > 50 Then
'                  Printer.NewPage
'                  iLine = 1
'                  PrintHead
'               End If
'               For i = 0 To 5 '4
'                  'If i = 0 Then
'                  If i <= 1 Then
'                     Printer.CurrentX = pLeft(i)
'                  'ElseIf i = 1 Then
'                  ElseIf i = 2 Then
'                     Printer.CurrentX = pLeft(i) - Printer.TextWidth(Format(.Fields(i), FDollar))
'                  Else
'                     Printer.CurrentX = pLeft(i) - Printer.TextWidth(Format(.Fields(i), DDollar2))
'                  End If
'                  Printer.CurrentY = iLine * 300
'                  'If i = 0 Then
'                  If i <= 1 Then
'                     Printer.Print "" & .Fields(i)
'                  'ElseIf i = 1 Then
'                  ElseIf i = 2 Then
'                     Printer.Print Format(.Fields(i), FDollar)
'                  Else
'                     Printer.Print Format(.Fields(i), DDollar2)
'                  End If
'               Next i
''               dblSum1 = dblSum1 + Val("" & .Fields(1))
''               dblSum2 = dblSum2 + Val("" & .Fields(2))
''               dblSum3 = dblSum3 + Val("" & .Fields(3))
''               dblSum4 = dblSum4 + Val("" & .Fields(4))
'               iLine = iLine + 1
'               .MoveNext
'            Loop
'         End With
'      End If
'   Next intCounter
'   'iLine = iLine + 1
'   Printer.CurrentX = pLeft(0)
'   Printer.CurrentY = iLine * 300
'   Printer.Print String(85, "-")
'   iLine = iLine + 1
'   '合計
''   For i = 0 To 4
''      If i = 1 Then dblSum = dblSum1
''      If i = 2 Then dblSum = dblSum2
''      If i = 3 Then dblSum = dblSum3
''      If i = 4 Then dblSum = dblSum4
''      If i <> 0 Then
''         Printer.CurrentX = PLeft(i) - Printer.TextWidth(Format(dblSum, DDollar2))
''      Else
''         Printer.CurrentX = PLeft(i)
''      End If
''      Printer.CurrentY = iLine * 300
''      If i <> 0 Then
''         Printer.Print Format(dblSum, DDollar2)
''      Else
''         Printer.Print "合計："
''      End If
''   Next i
'   'Modify By Sindy 2012/12/10
'   strExc(0) = "select '',a1k18,nvl(sum(A1-Z1),0),sum(A2),sum(A3),sum(A4) from(" & _
'               " select a1k18,sum(a1k08 - nvl(a1k31, 0)) A1,0 Z1, nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) A2, nvl(sum(a1k09),0) A3, nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0) A4 from (select distinct acc1k0.* From acc1k0,caseprogress where A1K01=CP60(+) " & strSql & ") a" & _
'               " group by a1k18" & _
'               " Union" & _
'               " select a1k18,0 A1,Sum(Nvl(A0Z04,0)) Z1,0 A2,0 A3,0 A4 from ACC0Z0,ACC0Y0,(select distinct acc1k0.* From acc1k0,caseprogress where A1K01=CP60(+) " & strSql & ") a" & _
'               " where A1K01=A0Z02(+) and A0Z01=A0Y01" & _
'               " group by a1k18" & _
'               " ) group by a1k18"
'   intI = 1
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With rsReport
'         Do While Not .EOF
'            If iLine > 50 Then
'               Printer.NewPage
'               iLine = 1
'               PrintHead
'            End If
'            For i = 0 To 5
'               If i <= 1 Then
'                  Printer.CurrentX = pLeft(i)
'               ElseIf i = 2 Then
'                  Printer.CurrentX = pLeft(i) - Printer.TextWidth(Format(.Fields(i), FDollar))
'               Else
'                  Printer.CurrentX = pLeft(i) - Printer.TextWidth(Format(.Fields(i), DDollar2))
'               End If
'               Printer.CurrentY = iLine * 300
'               If i = 0 Then
'                  If intI = 1 Then
'                     Printer.Print "合計："
'                     intI = 0
'                  End If
'               ElseIf i = 1 Then
'                  Printer.Print "" & .Fields(i)
'               ElseIf i = 2 Then
'                  Printer.Print Format(.Fields(i), FDollar)
'               Else
'                  Printer.Print Format(.Fields(i), DDollar2)
'               End If
'            Next i
'            iLine = iLine + 1
'            .MoveNext
'         Loop
'      End With
'   End If
'   'iLine = iLine + 1
'   '2012/12/10 End
'   Printer.CurrentX = pLeft(0)
'   Printer.CurrentY = iLine * 300
'   Printer.Print String(85, "-")
'   iLine = iLine + 1
'   '系統別合計
'   'Modify By Sindy 2012/12/10 nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))),0) sum1 ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) sum1
'   '                           nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09),0) ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0)
'   strExc(0) = "select a1k13,nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) sum1, nvl(sum(a1k09),0), nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0)" & _
'               " from (" & _
'               "select distinct acc1k0.*" & _
'               " From acc1k0,caseprogress" & _
'               " where A1K01=CP60(+)" & strSql & _
'               ") a" & _
'               " group by a1k13" & _
'               " order by sum1 desc"
'   intI = 1
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With rsReport
'         Do While Not .EOF
'            If iLine > 50 Then
'               Printer.NewPage
'               iLine = 1
'               PrintHead
'            End If
'            For i = 0 To 3
'               If i <> 0 Then
'                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth(Format(.Fields(i), DDollar2))
'               Else
'                  'Printer.CurrentX = 1500
'                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth("" & Left(.Fields(i) & "   ", 3) & "小計：")
'               End If
'               Printer.CurrentY = iLine * 300
'               If i <> 0 Then
'                  Printer.Print Format(.Fields(i), DDollar2)
'               Else
'                  Printer.Print "" & Left(.Fields(i) & "   ", 3) & "小計："
'               End If
'            Next i
'            iLine = iLine + 1
'            .MoveNext
'         Loop
'      End With
'   End If
'   Printer.CurrentX = pLeft(0)
'   Printer.CurrentY = iLine * 300
'   Printer.Print String(85, "-")
'   iLine = iLine + 1
'   '業務區合計
'   '外商F10-F19
'   'Modify By Sindy 2012/11/07 ,cp12=>,substr(cp12,1,2) cp12
'   'Modify By Sindy 2012/12/10 nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))),0) sum1 ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) sum1
'   '                           nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09),0) ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0)
'   strExc(0) = "select '外商',nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) sum1, nvl(sum(a1k09),0), nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0)" & _
'               " from (" & _
'               "select distinct acc1k0.*,substr(cp12,1,2) cp12" & _
'               " From acc1k0,caseprogress" & _
'               " where A1K01=CP60(+)" & strSql & _
'               ") a" & _
'               " where cp12='F1'"
'   intI = 1
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With rsReport
'         Do While Not .EOF
'            If iLine > 50 Then
'               Printer.NewPage
'               iLine = 1
'               PrintHead
'            End If
'            For i = 0 To 3
'               If i <> 0 Then
'                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth(Format(.Fields(i), DDollar2))
'               Else
'                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth("" & .Fields(i) & "小計：")
'               End If
'               Printer.CurrentY = iLine * 300
'               If i <> 0 Then
'                  Printer.Print Format(.Fields(i), DDollar2)
'               Else
'                  Printer.Print "" & .Fields(i) & "小計："
'               End If
'            Next i
'            iLine = iLine + 1
'            .MoveNext
'         Loop
'      End With
'   End If
'   '外專F20-F29
'   'Modify By Sindy 2012/11/07 ,cp12=>,substr(cp12,1,2) cp12
'   'Modify By Sindy 2012/12/10 nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))),0) sum1 ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) sum1
'   '                           nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09),0) ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0)
'   strExc(0) = "select '外專',nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) sum1, nvl(sum(a1k09),0), nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0)" & _
'               " from (" & _
'               "select distinct acc1k0.*,substr(cp12,1,2) cp12" & _
'               " From acc1k0,caseprogress" & _
'               " where A1K01=CP60(+)" & strSql & _
'               ") a" & _
'               " where cp12='F2'"
'   intI = 1
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With rsReport
'         Do While Not .EOF
'            If iLine > 50 Then
'               Printer.NewPage
'               iLine = 1
'               PrintHead
'            End If
'            For i = 0 To 3
'               If i <> 0 Then
'                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth(Format(.Fields(i), DDollar2))
'               Else
'                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth("" & .Fields(i) & "小計：")
'               End If
'               Printer.CurrentY = iLine * 300
'               If i <> 0 Then
'                  Printer.Print Format(.Fields(i), DDollar2)
'               Else
'                  Printer.Print "" & .Fields(i) & "小計："
'               End If
'            Next i
'            iLine = iLine + 1
'            .MoveNext
'         Loop
'      End With
'   End If
'   '外法F30-F49
'   'Modify By Sindy 2012/11/07 ,cp12=>,substr(cp12,1,2) cp12
'   'Modify By Sindy 2012/12/10 nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1))),0) sum1 ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) sum1
'   '                           nvl(sum(a1k11 - nvl(a1k30, 0) - (nvl(a1k06, 0) * nvl(a1k10, 1)) - a1k09),0) ==> nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0)
'   strExc(0) = "select '外法',nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)),0) sum1, nvl(sum(a1k09),0), nvl(sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09),0)" & _
'               " from (" & _
'               "select distinct acc1k0.*,substr(cp12,1,2) cp12" & _
'               " From acc1k0,caseprogress" & _
'               " where A1K01=CP60(+)" & strSql & _
'               ") a" & _
'               " where cp12>='F3' and cp12<='F4'"
'   intI = 1
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With rsReport
'         Do While Not .EOF
'            If iLine > 50 Then
'               Printer.NewPage
'               iLine = 1
'               PrintHead
'            End If
'            For i = 0 To 3
'               If i <> 0 Then
'                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth(Format(.Fields(i), DDollar2))
'               Else
'                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth("" & .Fields(i) & "小計：")
'               End If
'               Printer.CurrentY = iLine * 300
'               If i <> 0 Then
'                  Printer.Print Format(.Fields(i), DDollar2)
'               Else
'                  Printer.Print "" & .Fields(i) & "小計："
'               End If
'            Next i
'            iLine = iLine + 1
'            .MoveNext
'         Loop
'      End With
'   End If
'
'   Printer.EndDoc
'   PUB_RestorePrinter strPrinter
'   Call ShowPrintOk
'   Set rsReport = Nothing
End Sub

'Add By Sindy 2012/5/23
Private Sub GetPleft()
   pLeft(0) = 500
   pLeft(1) = 1500 'Add By Sindy 2012/12/10
   pLeft(2) = 3500
   pLeft(3) = 5500
   pLeft(4) = 7500
   pLeft(5) = 9500
End Sub

'Add By Sindy 2012/5/23
Private Sub PrintHead()
Dim strText As String
   
   GetPleft
   
   Printer.Font.Size = 14
   Printer.Font.Underline = False
   Printer.FontBold = True
   
   strExc(1) = "國外應收規費及服務費分析表"
   Printer.CurrentX = 3000 'Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(1)) / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print strExc(1)
   
   Printer.Font.Size = 12
   Printer.FontBold = False
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) And _
      MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      Printer.CurrentX = 3000 'Printer.ScaleWidth / 2 - (Printer.TextWidth("補充資料日期：" & ChangeTStringToTDateString(txtDate(0)) & " - " & ChangeTStringToTDateString(txtDate(1))) / 2)
      Printer.CurrentY = 900
      Printer.Print "請款日期：" & MaskEdBox1.Text & " - " & MaskEdBox2.Text
   End If
   Printer.CurrentX = 7000 'Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
   Printer.CurrentY = 900
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Printer.CurrentX = pLeft(0)
   Printer.CurrentY = 1200
   Printer.Print "列印人員：" & strUserName
   If Text1 <> "" And Text8 <> "" Then
      Printer.CurrentX = 3000 'Printer.ScaleWidth / 2 - (Printer.TextWidth("補充資料日期：" & ChangeTStringToTDateString(txtDate(0)) & " - " & ChangeTStringToTDateString(txtDate(1))) / 2)
      Printer.CurrentY = 1200
      Printer.Print "　業務區：" & Text1.Text & " - " & Text8.Text
   End If
   Printer.CurrentX = 7000 'Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
   Printer.CurrentY = 1200
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = 6
   
   Printer.CurrentX = pLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print "年度"
   'Add By Sindy 2012/12/10
   Printer.CurrentX = pLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "幣別"
   '2012/12/10 End
   'Modify By Sindy 2012/12/10
'   Printer.CurrentX = PLeft(1) - Printer.TextWidth("請款美金")
'   Printer.CurrentY = iLine * 300
'   Printer.Print "請款美金"
   Printer.CurrentX = pLeft(2) - Printer.TextWidth("請款外幣")
   Printer.CurrentY = iLine * 300
   Printer.Print "請款外幣"
   '2012/12/10 End
   Printer.CurrentX = pLeft(3) - Printer.TextWidth("台幣金額")
   Printer.CurrentY = iLine * 300
   Printer.Print "台幣金額"
   Printer.CurrentX = pLeft(4) - Printer.TextWidth("規費")
   Printer.CurrentY = iLine * 300
   Printer.Print "規費"
   Printer.CurrentX = pLeft(5) - Printer.TextWidth("服務費")
   Printer.CurrentY = iLine * 300
   Printer.Print "服務費"
   
   iLine = iLine + 1
   Printer.CurrentX = pLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print String(85, "-")
   iLine = iLine + 1
End Sub

'Add by Amy 2017/10/24 改寫至暫存檔
Private Sub PrintData()
   Dim strCon As String, strQ As String, strWhere(1) As String
   Dim rsReport As ADODB.Recordset
   Dim intCounter As Integer
   Dim i As Integer
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   strSql = ""
   '請款日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   '抓未銷帳未作廢未結清的資料
   strSql = strSql & " And A1K25 Is Null AND (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0)"
   '系統別
   If Text2 <> MsgText(601) Or Text3 <> MsgText(601) Or Text4 <> MsgText(601) Or Text5 <> MsgText(601) Or Text6 <> MsgText(601) Then
      strSql = strSql & " and ("
      If Text2 <> MsgText(601) Then
         strCon = strCon & "a1k13 = '" & Text2 & "' or "
      End If
      If Text3 <> MsgText(601) Then
         strCon = strCon & "a1k13 = '" & Text3 & "' or "
      End If
      If Text4 <> MsgText(601) Then
         strCon = strCon & "a1k13 = '" & Text4 & "' or "
      End If
      If Text5 <> MsgText(601) Then
         strCon = strCon & "a1k13 = '" & Text5 & "' or "
      End If
      If Text6 <> MsgText(601) Then
         strCon = strCon & "a1k13 = '" & Text6 & "' or "
      End If
      If Text7 <> MsgText(601) Then
         strCon = strCon & "a1k13 = '" & Text7 & "' or "
      End If
      strSql = strSql & Mid(strCon, 1, Len(strCon) - 4) & ")"
   End If
   '業務區
   If Text1 <> "" Then
      strSql = strSql & " and cp12>='" & Text1 & "'"
   End If
   If Text8 <> "" Then
      strSql = strSql & " and cp12<='" & Text8 & "'"
   End If
   
   '抓取資料寫入暫存檔,修改未收金額算法,不以a0z04計算改以台幣已收金額/請款匯率,與frmacc24a0一致
   strQ = "Delete From Accrpt24e0 Where ID='" & strUserNum & "'"
   adoTaie.Execute strQ
   strQ = "Insert Into Accrpt24e0 (ID,R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13,R14) " & _
            "Select '" & strUserNum & "', a1k01,a1k02,a1k06,a1k08,a1k09,a1k11,a1k30,a1k31,a1k18,a1k13,cp12," & _
            "Decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k31, 0)),Round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0))/a1k10,1)), " & _
            "Round(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) - a1k09,0)," & _
            "a1k28" & _
            " From (" & _
                "Select Distinct acc1k0.*,cp12" & _
                " From acc1k0,caseprogress" & _
                " where A1K01=CP60(+)" & strSql & _
            ")"
    adoTaie.Execute strQ
    
    '不輸國籍時剔除大陸,大陸一定要輸國
    If Text9 = "020" Then
            strWhere(0) = " And fa10 = '020'"
            strWhere(1) = " And cu10 = '020'"
    Else
        If Text9 <> MsgText(601) Then
            strQ = strWhere(0) & " And fa10 >= '" & Text9 & "'"
            strWhere(1) = strWhere(1) & " And cu10 >= '" & Text9 & "'"
        End If
        If Text10 <> MsgText(601) Then
            strWhere(0) = strWhere(0) & " And fa10 >= '" & Text10 & "z'"
            strWhere(1) = strWhere(1) & " And cu10 >= '" & Text10 & "z'"
        End If
        strWhere(0) = strWhere(0) & " And fa10 <> '020'"
        strWhere(1) = strWhere(1) & " And cu10 <> '020'"
    End If
    strQ = "Delete From Accrpt24e0 Where ID='" & strUserNum & "' And R01 Not in (" & _
                "Select R01 From Accrpt24e0,Fagent Where ID='" & strUserNum & "' And SubStr(R14,1,1)='Y' And SubStr(R14, 1, 8) = fa01 And SubStr(R14, 9, 1) = fa02 " & strWhere(0) & _
         " Union Select R01 From Accrpt24e0,Customer Where ID='" & strUserNum & "' And SubStr(R14,1,1)='X' And SubStr(R14, 1, 8) = cu01 And SubStr(R14, 9, 1) = cu02 " & strWhere(1) & _
                ")"
    adoTaie.Execute strQ
    
    '檢查有無資料
    strQ = "Select * From Accrpt24e0 Where ID='" & strUserNum & "' "
    
    intI = 1
    Set rsReport = ClsLawReadRstMsg(intI, strQ)
    If intI <> 1 Then
       MsgBox "查無資料！"
       Exit Sub
    End If
    
    PUB_RestorePrinter Combo1
    Printer.Orientation = 1 '1.直印 2.橫印
    Printer.Font.Name = "細明體"
    iLine = 1
    PrintHead
    
    '依年度、幣別列印資料
    For intCounter = Val(Mid(MaskEdBox1.Text, 1, 3)) To Val(Mid(MaskEdBox2.Text, 1, 3))
        strQ = "Select '" & intCounter & "' as strYear,R09 as a1k18,Sum(Nvl(R12,0)),Sum(R06 - Nvl(R07, 0) - Nvl(R03, 0)) as A2,Sum(Nvl(R05,0)) as A3,Sum(R06 - Nvl(R07, 0) - Nvl(R03, 0) - Nvl(R05,0)) as A4" & _
              " From Accrpt24e0" & _
              " Where ID='" & strUserNum & "' And Decode(length(R02), 7, SubStr(R02, 1, 3), 6, SubStr(R02, 1, 2))=" & intCounter & _
              " Group by R09"
        intI = 1
        Set rsReport = ClsLawReadRstMsg(intI, strQ)
        If intI = 1 Then
           With rsReport
              Do While Not .EOF
                 If iLine > 50 Then
                    Printer.NewPage
                    iLine = 1
                    PrintHead
                 End If
                 For i = 0 To 5
                    If i <= 1 Then
                       Printer.CurrentX = pLeft(i)
                    ElseIf i = 2 Then
                       Printer.CurrentX = pLeft(i) - Printer.TextWidth(Format(.Fields(i), FDollar))
                    Else
                       Printer.CurrentX = pLeft(i) - Printer.TextWidth(Format(.Fields(i), DDollar2))
                    End If
                    Printer.CurrentY = iLine * 300
                    If i <= 1 Then
                       Printer.Print "" & .Fields(i)
                    ElseIf i = 2 Then
                       Printer.Print Format(.Fields(i), FDollar)
                    Else
                       Printer.Print Format(.Fields(i), DDollar2)
                    End If
                 Next i
                 iLine = iLine + 1
                 .MoveNext
              Loop
           End With
        End If
    Next intCounter
    Printer.CurrentX = pLeft(0)
    Printer.CurrentY = iLine * 300
    Printer.Print String(85, "-")
    iLine = iLine + 1
    
    '合計
    strQ = "Select '',R09 as a1k18,Sum(Nvl(R12,0)),Sum(R06 - Nvl(R07, 0) - Nvl(R03, 0)) as A2,Sum(Nvl(R05,0)) as A3,Sum(R06 - Nvl(R07, 0) - Nvl(R03, 0) - Nvl(R05,0)) as A4" & _
          " From Accrpt24e0 Where ID='" & strUserNum & "'" & _
          " Group by R09"
               
    intI = 1
    Set rsReport = ClsLawReadRstMsg(intI, strQ)
    If intI = 1 Then
         With rsReport
         Do While Not .EOF
            If iLine > 50 Then
               Printer.NewPage
               iLine = 1
               PrintHead
            End If
            For i = 0 To 5
               If i <= 1 Then
                  Printer.CurrentX = pLeft(i)
               ElseIf i = 2 Then
                  Printer.CurrentX = pLeft(i) - Printer.TextWidth(Format(.Fields(i), FDollar))
               Else
                  Printer.CurrentX = pLeft(i) - Printer.TextWidth(Format(.Fields(i), DDollar2))
               End If
               Printer.CurrentY = iLine * 300
               If i = 0 Then
                  If intI = 1 Then
                     Printer.Print "合計："
                     intI = 0
                  End If
               ElseIf i = 1 Then
                  Printer.Print "" & .Fields(i)
               ElseIf i = 2 Then
                  Printer.Print Format(.Fields(i), FDollar)
               Else
                  Printer.Print Format(.Fields(i), DDollar2)
               End If
            Next i
            iLine = iLine + 1
            .MoveNext
         Loop
      End With
    End If
    Printer.CurrentX = pLeft(0)
    Printer.CurrentY = iLine * 300
    Printer.Print String(85, "-")
    iLine = iLine + 1
    
    '系統別合計
    strQ = "Select R10,nvl(sum(R06 - nvl(R07, 0) - nvl(R03, 0)),0) sum1, nvl(sum(R05),0), nvl(sum(R06 - nvl(R07, 0) - nvl(R03, 0) - R05),0)" & _
               " From Accrpt24e0 Where ID='" & strUserNum & "'" & _
               " Group by R10" & _
               " Order by sum1 Desc"
    intI = 1
    Set rsReport = ClsLawReadRstMsg(intI, strQ)
    If intI = 1 Then
        With rsReport
            Do While Not .EOF
                If iLine > 50 Then
                   Printer.NewPage
                   iLine = 1
                   PrintHead
                End If
                For i = 0 To 3
                   If i <> 0 Then
                      Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth(Format(.Fields(i), DDollar2))
                   Else
                      Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth("" & Left(.Fields(i) & "   ", 3) & "小計：")
                   End If
                   Printer.CurrentY = iLine * 300
                   If i <> 0 Then
                      Printer.Print Format(.Fields(i), DDollar2)
                   Else
                      Printer.Print "" & Left(.Fields(i) & "   ", 3) & "小計："
                   End If
                Next i
                iLine = iLine + 1
                .MoveNext
            Loop
        End With
    End If
    Printer.CurrentX = pLeft(0)
    Printer.CurrentY = iLine * 300
    Printer.Print String(85, "-")
    iLine = iLine + 1
    
    '業務區合計
    '外商F10-F19/外專F20-F29/外法F30-F49
    strQ = "Select * From (" & _
           "Select '外商',nvl(sum(R06 - nvl(R07, 0) - nvl(R03, 0)),0) sum1, nvl(sum(R05),0), nvl(sum(R06 - nvl(R07, 0) - nvl(R03, 0) - R05),0),1 as Sort " & _
           "From Accrpt24e0 Where ID='" & strUserNum & "' And SubStr(R11,1,2)='F1' " & _
     "Union Select '外專',nvl(sum(R06 - nvl(R07, 0) - nvl(R03, 0)),0) sum1, nvl(sum(R05),0), nvl(sum(R06 - nvl(R07, 0) - nvl(R03, 0) - R05),0),2 as Sort " & _
           "From Accrpt24e0 Where ID='" & strUserNum & "' And SubStr(R11,1,2)='F2' " & _
     "Union Select '外法',nvl(sum(R06 - nvl(R07, 0) - nvl(R03, 0)),0) sum1, nvl(sum(R05),0), nvl(sum(R06 - nvl(R07, 0) - nvl(R03, 0) - R05),0),3 as Sort " & _
           "From Accrpt24e0 Where ID='" & strUserNum & "' And SubStr(R11,1,2)>='F3' And SubStr(R11,1,2)<='F4' " & _
           ") Order by sort"
    intI = 1
    Set rsReport = ClsLawReadRstMsg(intI, strQ)
    If intI = 1 Then
      With rsReport
         Do While Not .EOF
            If iLine > 50 Then
               Printer.NewPage
               iLine = 1
               PrintHead
            End If
            For i = 0 To 3
               If i <> 0 Then
                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth(Format(.Fields(i), DDollar2))
               Else
                  Printer.CurrentX = pLeft(i + 2) - Printer.TextWidth("" & .Fields(i) & "小計：")
               End If
               Printer.CurrentY = iLine * 300
               If i <> 0 Then
                  Printer.Print Format(.Fields(i), DDollar2)
               Else
                  Printer.Print "" & .Fields(i) & "小計："
               End If
            Next i
            iLine = iLine + 1
            .MoveNext
         Loop
      End With
    End If
               
    Printer.EndDoc
    PUB_RestorePrinter strPrinter
    Call ShowPrintOk
    Set rsReport = Nothing
End Sub
