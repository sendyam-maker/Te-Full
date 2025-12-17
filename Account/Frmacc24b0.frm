VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24b0 
   AutoRedraw      =   -1  'True
   Caption         =   "代理人帳目排名"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   5160
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel檔(&E)"
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
      Left            =   1500
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   2610
      Width           =   2295
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
      Left            =   330
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   3060
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
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
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   1572
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
      TabIndex        =   2
      Top             =   600
      Width           =   1572
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
      Width           =   852
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
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   852
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
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
      TabIndex        =   5
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
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "列印名次："
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
      TabIndex        =   17
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "(1.FC往來 2.FC未收         3.CF往來 4.CF未付         5.往來    6.未收未付)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2040
      TabIndex        =   16
      Top             =   1320
      Width           =   2532
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "排名順序："
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
      Top             =   1320
      Width           =   972
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
      Height          =   252
      Left            =   3000
      TabIndex        =   14
      Top             =   960
      Width           =   252
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "往來日期："
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
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label3 
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
      Height          =   252
      Left            =   3000
      TabIndex        =   12
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "代理人："
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
      TabIndex        =   11
      Top             =   600
      Width           =   972
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
      Height          =   252
      Left            =   2280
      TabIndex        =   10
      Top             =   240
      Width           =   252
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc24b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/09 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit

Public adoacc150 As New ADODB.Recordset
Public adoacc1k0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt211 As New ADODB.Recordset
Dim dllaccrpt211 As Object
Dim strSql As String
Dim strWhere(5) As String
'Add By Sindy 2014/5/29
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
Dim intPage As Integer
'2014/5/29 END


'Add By Sindy 2014/5/29 產生Excel檔
Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt211Delete
   ProduceData
   PrintExcel
   If strCon10 <> MsgText(602) Then
      FormClear
   End If
   Screen.MousePointer = vbDefault
   StatusView "" 'MsgText(101)
End Sub

'列印
Private Sub Command2_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt211Delete
   ProduceData
   If adoaccrpt211.State = adStateOpen Then
      adoaccrpt211.Close
   End If
   adoaccrpt211.CursorLocation = adUseClient
   adoaccrpt211.Open "select * from accrpt211 where R21101='" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt211.RecordCount <> 0 Then
      dllaccrpt211.Acc24b0 ReportTitle(211), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   End If
   adoaccrpt211.Close
   If strCon10 <> MsgText(602) Then
      FormClear
   End If
   Screen.MousePointer = vbDefault
   StatusView "" 'MsgText(101)
End Sub

Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      StatusView "" 'MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 5250
'   Me.Height = 3650
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
   PUB_InitForm Me, 5250, 3650, strBackPicPath4
   'end 2021/12/09
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   StatusView "" 'MsgText(101)
   Set dllaccrpt211 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt211 = Nothing
   Set Frmacc24b0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
   CloseIme
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
   CloseIme
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

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim intCounter As Integer
   
On Error GoTo Checking
   
   StatusView MsgText(26)
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/22 清除查詢印表記錄檔欄位
   strSql = ""
   For intCounter = 0 To 5
      strWhere(intCounter) = ""
   Next intCounter
   If Text1 <> "" Then
      'Modify by Morgan 2010/11/11 請款對象有可能是客戶
      'strWhere(0) = strWhere(0) & " and fa10 >= '" & Text1 & "'"
      strWhere(5) = strWhere(5) & " and fa10 >= '" & Text1 & "'"
'      strWhere(1) = strWhere(1) & " and fa10 >= '" & Text1 & "'"
      strWhere(2) = strWhere(2) & " and fa10 >= '" & Text1 & "'"
'      strWhere(3) = strWhere(3) & " and fa10 >= '" & Text1 & "'"
      strWhere(4) = strWhere(4) & " and cu10 >= '" & Text1 & "'" 'Add by Morgan 2010/11/11
   End If
   If Text2 <> "" Then
      'Modify by Morgan 2010/11/11 請款對象有可能是客戶
      'strWhere(0) = strWhere(0) & " and fa10 <= '" & Text2 & "z'"
      strWhere(5) = strWhere(5) & " and fa10 <= '" & Text2 & "z'"
'      strWhere(1) = strWhere(1) & " and fa10 <= '" & Text2 & "z'"
      strWhere(2) = strWhere(2) & " and fa10 <= '" & Text2 & "z'"
'      strWhere(3) = strWhere(3) & " and fa10 <= '" & Text2 & "z'"
      strWhere(4) = strWhere(4) & " and cu10 <= '" & Text2 & "z'" 'Add by Morgan 2010/11/11
   End If
   If Text1 <> "" Or Text2 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1 & "-" & Text2 'Add By Sindy 2010/12/22
   End If
   If Text3 <> "" Then
      'Modify by Morgan 2010/11/11 要抓請款對象
      'strWhere(0) = strWhere(0) & " and a1k03 >= '" & Text3 & "'"
      strWhere(0) = strWhere(0) & " and a1k28 >= '" & Text3 & "'"
'      strWhere(1) = strWhere(1) & " and a0y07 >= '" & Text3 & "'"
      strWhere(2) = strWhere(2) & " and a1503 >= '" & Text3 & "'"
'      strWhere(3) = strWhere(3) & " and a1803 >= '" & Text3 & "'"
   End If
   If Text4 <> "" Then
      'Modify by Morgan 2010/11/11 要抓請款對象
      'strWhere(0) = strWhere(0) & " and a1k03 <= '" & Text4 & "'"
      strWhere(0) = strWhere(0) & " and a1k28 <= '" & Text4 & "'"
'      strWhere(1) = strWhere(1) & " and a0y07 <= '" & Text4 & "'"
      strWhere(2) = strWhere(2) & " and a1503 <= '" & Text4 & "'"
'      strWhere(3) = strWhere(3) & " and a1803 <= '" & Text4 & "'"
   End If
   If Text3 <> "" Or Text4 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label2 & Text3 & "-" & Text4 'Add By Sindy 2010/12/22
   End If
   If MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "" Then
      strWhere(0) = strWhere(0) & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'      strWhere(1) = strWhere(1) & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(2) = strWhere(2) & " and a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'      strWhere(3) = strWhere(3) & " and a1802 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "" Then
      strWhere(0) = strWhere(0) & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'      strWhere(1) = strWhere(1) & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(2) = strWhere(2) & " and a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'      strWhere(3) = strWhere(3) & " and a1802 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If (MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "") Or _
      (MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "") Then
      pub_QL05 = pub_QL05 & ";" & Label4 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/22
   End If
   
'   Select Case Text5
'      Case "1"
'         pub_QL05 = pub_QL05 & ";" & Label5 & "1.FC往來" 'Add By Sindy 2010/12/22
'         'Modify by Morgan 2010/11/11 要抓請款對象且要考慮X編號
'         'strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k06, 0) * a1k10) as Namount, '1' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
'         'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
'         strSql = "select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k06, 0) * a1k10) as Namount, '1' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
'         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
'         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k06, 0) * a1k10) as Namount, '1' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+)" & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
'         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
'      Case "2"
'         pub_QL05 = pub_QL05 & ";" & Label5 & "2.FC未收" 'Add By Sindy 2010/12/22
'         'Modify by Morgan 2010/11/11 要抓請款對象且要考慮X編號並排除已銷帳
'         'strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
'         strSql = "select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k25 is null " & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
'         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k25 is null " & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
'      Case "3"
'         pub_QL05 = pub_QL05 & ";" & Label5 & "3.CF往來" 'Add By Sindy 2010/12/22
'         strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506) as Namount, '3' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
'         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506 - nvl(a1520, 0)) as Namount, '4' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1506 > nvl(a1520, 0)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
'      Case "4"
'         pub_QL05 = pub_QL05 & ";" & Label5 & "4.CF未付" 'Add By Sindy 2010/12/22
'         strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506 - nvl(a1520, 0)) as Namount, '4' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1506 > nvl(a1520, 0)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
'      Case "6"
'         pub_QL05 = pub_QL05 & ";" & Label5 & "6.未收未付" 'Add By Sindy 2010/12/22
'         'Modify by Morgan 2010/11/11 要抓請款對象且要考慮X編並排除已銷帳
'         'strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
'         strSql = "select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k25 is null " & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
'         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k25 is null " & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
'         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506 - nvl(a1520, 0)) as Namount, '4' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1506 > nvl(a1520, 0)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
'      Case Else
'         pub_QL05 = pub_QL05 & ";" & Label5 & "5.往來" 'Add By Sindy 2010/12/22
'         'Modify by Morgan 2010/11/11 要抓請款對象且要考慮X編
'         'strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k06, 0) * a1k10) as Namount, '1' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
'         'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
'         strSql = "select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k06, 0) * a1k10) as Namount, '1' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
'         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
'         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k06, 0) * a1k10) as Namount, '1' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+)" & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
'         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
'         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506) as Namount, '3' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
'         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506 - nvl(a1520, 0)) as Namount, '4' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1506 > nvl(a1520, 0)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
'   End Select
   'Modify By Sindy 2013/1/15 台幣折讓不用換算了,直接引用a1k06即可
   Select Case Text5
      Case "1"
         pub_QL05 = pub_QL05 & ";" & Label5 & "1.FC往來" 'Add By Sindy 2010/12/22
         'Modify by Morgan 2010/11/11 要抓請款對象且要考慮X編號
         'strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k06, 0) * a1k10) as Namount, '1' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
         'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
         strSql = "select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k06, 0)) as Namount, '1' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k06, 0)) as Namount, '1' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+)" & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) as Namount, '2' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
      Case "2"
         pub_QL05 = pub_QL05 & ";" & Label5 & "2.FC未收" 'Add By Sindy 2010/12/22
         'Modify by Morgan 2010/11/11 要抓請款對象且要考慮X編號並排除已銷帳
         'strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
         strSql = "select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k25 is null " & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) as Namount, '2' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k25 is null " & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
      Case "3"
         pub_QL05 = pub_QL05 & ";" & Label5 & "3.CF往來" 'Add By Sindy 2010/12/22
         strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506) as Namount, '3' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506 - nvl(a1520, 0)) as Namount, '4' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1506 > nvl(a1520, 0)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
      Case "4"
         pub_QL05 = pub_QL05 & ";" & Label5 & "4.CF未付" 'Add By Sindy 2010/12/22
         strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506 - nvl(a1520, 0)) as Namount, '4' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1506 > nvl(a1520, 0)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
      Case "6"
         pub_QL05 = pub_QL05 & ";" & Label5 & "6.未收未付" 'Add By Sindy 2010/12/22
         'Modify by Morgan 2010/11/11 要抓請款對象且要考慮X編並排除已銷帳
         'strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
         strSql = "select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k25 is null " & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) as Namount, '2' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k25 is null " & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506 - nvl(a1520, 0)) as Namount, '4' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1506 > nvl(a1520, 0)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
      Case Else
         pub_QL05 = pub_QL05 & ";" & Label5 & "5.往來" 'Add By Sindy 2010/12/22
         'Modify by Morgan 2010/11/11 要抓請款對象且要考慮X編
         'strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k06, 0) * a1k10) as Namount, '1' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
         'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0) * a1k10) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & " group by a1k03, nvl(fa05, nvl(fa06, fa04)), na03"
         strSql = "select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k06, 0)) as Namount, '1' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(fa05, nvl(fa06, fa04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) as Namount, '2' as No from acc1k0, fagent, nation where substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & strWhere(5) & " and substr(a1k28,1,1)='Y' group by a1k28"
         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k06, 0)) as Namount, '1' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+)" & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
         strSql = strSql & " union select a1k28 as FagentNo, max(nvl(cu05, nvl(cu06, cu04))) as FagentName, max(na03) as Nation, sum(a1k11 - nvl(a1k30, 0) - nvl(a1k06, 0)) as Namount, '2' as No from acc1k0, customer, nation where substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cu10 = na01 (+) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '')" & strWhere(0) & strWhere(4) & " and substr(a1k28,1,1)='X' group by a1k28"
         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506) as Namount, '3' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, sum(a1506 - nvl(a1520, 0)) as Namount, '4' as No from acc150, fagent, nation where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1506 > nvl(a1520, 0)" & strWhere(2) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03"
   End Select
   If Text6 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label9 & Text6 'Add By Sindy 2010/12/22
   End If
   strSql = "select FagentNo, FagentName, nation, sum(decode(No, '1', Namount, 0)) as FCall, sum(decode(No, '2', Namount, 0)) as FCpart, sum(decode(No, '3', Namount, 0)) as CFall, sum(decode(No, '4', Namount, 0)) as CFpart from (" & strSql & ") Old group by FagentNo, FagentName, nation"
   Select Case Text5
      Case "1"
         strSql = "select * from (" & strSql & ") new order by FCall desc, CFall desc"
      Case "2"
         strSql = "select * from (" & strSql & ") new order by FCpart desc, CFpart desc"
      Case "3"
         strSql = "select * from (" & strSql & ") new order by CFall desc, FCall desc"
      Case "4"
         strSql = "select * from (" & strSql & ") new order by CFpart desc, FCpart desc"
      Case Else
         strSql = "select * from (" & strSql & ") new order by FCall desc, CFall desc"
   End Select
   intCounter = 1
   strCon10 = ""
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
      strCon10 = MsgText(602)
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      InsertQueryLog (adoacc1k0.RecordCount) 'Add By Sindy 2010/12/22
   End If
   Do While adoacc1k0.EOF = False
      adoTaie.Execute "insert into accrpt211 (r21101, r21102, r21103, r21104, r21105, r21106, r21107, r21108, r21109) values ('" & strUserNum & "', " & intCounter & ", '" & adoacc1k0.Fields("FagentNo").Value & "', '" & ChgSQL("" & adoacc1k0.Fields("FagentName").Value) & "', '" & adoacc1k0.Fields("nation").Value & "', " & adoacc1k0.Fields("FCall").Value & ", " & adoacc1k0.Fields("FCpart").Value & ", " & adoacc1k0.Fields("CFall").Value & ", " & adoacc1k0.Fields("CFpart").Value & ")"
      intCounter = intCounter + 1
      adoacc1k0.MoveNext
   Loop
   adoacc1k0.Close
   If Text6 <> "" Then
      adoTaie.Execute "delete from accrpt211 where R21101='" & strUserNum & "' and r21102 > " & Val(Text6)
   End If
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
Private Sub Accrpt211Delete()
   'Modify by Morgan 2010/11/11 不可只刪本使用者資料，除非改Accreprot的語法
   adoTaie.Execute "delete from accrpt211"
End Sub

''*************************************************
''  代理人FC儲存
''
''*************************************************
'Private Sub FagentFCSave()
'Dim strSql As String
'
'   adoaccrpt211.Fields("r21101").Value = strUserNum
'   If IsNull(adoacc1k0.Fields("a1k03").Value) Then
'      adoaccrpt211.Fields("r21103").Value = Null
'   Else
'      adoaccrpt211.Fields("r21103").Value = adoacc1k0.Fields("a1k03").Value
'      adoaccrpt211.Fields("r21104").Value = FagentQuery(adoacc1k0.Fields("a1k03").Value, 2)
'   End If
'   If IsNull(adoacc1k0.Fields("fa10").Value) Then
'      adoaccrpt211.Fields("r21105").Value = Null
'   Else
'      adoaccrpt211.Fields("r21105").Value = NationQuery(adoacc1k0.Fields("fa10").Value, 1)
'   End If
'   adoaccrpt211.Fields("r21106").Value = 0
'   adoaccrpt211.Fields("r21107").Value = 0
'   adoaccrpt211.Fields("r21108").Value = 0
'   adoaccrpt211.Fields("r21109").Value = 0
'   adoaccsum.CursorLocation = adUseClient
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   adoaccsum.Open "select sum(a0z04) from acc0z0, acc1k0 where a0z02 = a1k01 and a1k03 = '" & adoacc1k0.Fields("a1k03").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt211.Fields("r21110").Value = 0
'      Else
'         adoaccrpt211.Fields("r21110").Value = adoaccsum.Fields(0).Value
'      End If
'   Else
'      adoaccrpt211.Fields("r21110").Value = 0
'   End If
'   adoaccsum.Close
'   adoaccsum.CursorLocation = adUseClient
'   strSql = MsgText(601)
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and 1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and 1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   adoaccsum.Open "select sum(a1905) from acc190, acc150 where a1902 = a1501 and a1503 = '" & adoacc1k0.Fields("a1k03").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt211.Fields("r21111").Value = 0
'      Else
'         adoaccrpt211.Fields("r21111").Value = adoaccsum.Fields(0).Value
'      End If
'   Else
'      adoaccrpt211.Fields("r21111").Value = 0
'   End If
'   adoaccsum.Close
'End Sub

''*************************************************
''  代理人CF儲存
''
''*************************************************
'Private Sub FagentCFSave()
'   adoaccrpt211.Fields("r21101").Value = strUserNum
'   If IsNull(adoacc1k0.Fields("a1503").Value) Then
'      adoaccrpt211.Fields("r21103").Value = Null
'   Else
'      adoaccrpt211.Fields("r21103").Value = adoacc150.Fields("a1503").Value
'      adoaccrpt211.Fields("r21104").Value = FagentQuery(adoacc150.Fields("a1503").Value, 2)
'   End If
'   If IsNull(adoacc150.Fields("fa10").Value) Then
'      adoaccrpt211.Fields("r21105").Value = Null
'   Else
'      adoaccrpt211.Fields("r21105").Value = NationQuery(adoacc150.Fields("fa10").Value, 1)
'   End If
'   adoaccrpt211.Fields("r21106").Value = 0
'   adoaccrpt211.Fields("r21107").Value = 0
'   adoaccrpt211.Fields("r21108").Value = 0
'   adoaccrpt211.Fields("r21109").Value = 0
'   adoaccrpt211.Fields("r21110").Value = 0
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(a1905) from acc190, acc150 where a1902 = a1501 and a1503 = '" & adoacc150.Fields("a1503").Value & "' and a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt211.Fields("r21111").Value = 0
'      Else
'         adoaccrpt211.Fields("r21111").Value = adoaccsum.Fields(0).Value
'      End If
'   Else
'      adoaccrpt211.Fields("r21111").Value = 0
'   End If
'   adoaccsum.Close
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text5 = ""
   Text6 = ""
   Text1.SetFocus
End Sub

''*************************************************
''  產生報表資料
''
''*************************************************
'Private Sub ProduceData1()
'Dim strSql As String
'Dim intCounter As Integer
'
'On Error GoTo Checking
'   Select Case Val(Text5)
'      Case 1, 2, 5, 6
'         If Text1 <> MsgText(601) Then
'            strSql = " and fa10 >= '" & Text1 & "'"
'         End If
'         If Text2 <> MsgText(601) Then
'            strSql = strSql & " and fa10 <= '" & Text2 & "'"
'         End If
'         If Text3 <> MsgText(601) Then
'            strSql = strSql & " and a1k03 >= '" & Text3 & "'"
'         End If
'         If Text4 <> MsgText(601) Then
'            strSql = strSql & " and a1k03 <= '" & Text4 & "'"
'         End If
'         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'            strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'         End If
'         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'            strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'         End If
'         StatusView MsgText(26)
'         adoaccrpt211.CursorLocation = adUseClient
'         adoaccrpt211.Open "select * from accrpt211", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         adoacc1k0.CursorLocation = adUseClient
'         adoacc1k0.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'         If adoacc1k0.RecordCount = 0 Then
'            adoacc1k0.Close
'            MsgBox MsgText(28), , MsgText(5)
'            Exit Sub
'         End If
'         Do While adoacc1k0.EOF = False
'            adoaccrpt211.Find "r21101 = '" & strUserNum & "'", 0, adSearchForward, 1
'            If adoaccrpt211.EOF Then
'               adoaccrpt211.AddNew
'               FagentFCSave
'            Else
'               adoaccrpt211.Find "r21103 = '" & adoacc1k0.Fields("a1k03").Value & "'", 0, adSearchForward, adoaccrpt211.Bookmark
'               If adoaccrpt211.EOF Then
'                  adoaccrpt211.AddNew
'                  FagentFCSave
'               End If
'            End If
'            If IsNull(adoacc1k0.Fields("a1k11").Value) = False Then
'               adoaccrpt211.Fields("r21106").Value = Val(adoaccrpt211.Fields("r21106").Value) + Val(adoacc1k0.Fields("a1k11").Value)
'            End If
'            adoaccrpt211.UpdateBatch
'            adoacc1k0.MoveNext
'         Loop
'         adoacc1k0.Close
'      Case 3, 4, 5, 6
'         adoacc150.CursorLocation = adUseClient
'         strSql = MsgText(601)
'         If Text1 <> MsgText(601) Then
'            strSql = " and fa10 >= '" & Text1 & "'"
'         End If
'         If Text2 <> MsgText(601) Then
'            strSql = strSql & " and fa10 <= '" & Text2 & "'"
'         End If
'         If Text3 <> MsgText(601) Then
'            strSql = strSql & " and a1503 >= '" & Text3 & "'"
'         End If
'         If Text4 <> MsgText(601) Then
'            strSql = strSql & " and a1503 <= '" & Text4 & "'"
'         End If
'         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'            strSql = strSql & " and a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'         End If
'         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'            strSql = strSql & " and a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'         End If
'         adoacc150.Open "select * from acc150, fagent where substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'         If adoacc150.RecordCount = 0 Then
'            adoacc150.Close
'            MsgBox MsgText(28), , MsgText(5)
'            Exit Sub
'         End If
'         Do While adoacc150.EOF = False
'            adoaccrpt211.Find "r21101 = '" & strUserNum & "'", 0, adSearchForward, 1
'            If adoaccrpt211.EOF Then
'               adoaccrpt211.AddNew
'               FagentCFSave
'            Else
'               adoaccrpt211.Find "r21103 = '" & adoacc150.Fields("a1503").Value & "'", 0, adSearchForward, adoaccrpt211.Bookmark
'               If adoaccrpt211.EOF Then
'                  adoaccrpt211.AddNew
'                  FagentCFSave
'               End If
'            End If
'            If IsNull(adoacc150.Fields("a1510").Value) = False Then
'               adoaccrpt211.Fields("r21108").Value = Val(adoaccrpt211.Fields("r21108").Value) + Val(adoacc150.Fields("a1510").Value)
'            End If
'            adoaccrpt211.UpdateBatch
'            adoacc150.MoveNext
'         Loop
'         adoacc150.Close
'   End Select
'   adoaccrpt211.Close
'   adoTaie.Execute "update accrpt211 set r21107 = r21106 - r21110, r21109 = r21108 - r21111"
'
'   intCounter = 1
'   adoaccrpt211.CursorLocation = adUseClient
'   adoaccrpt211.Open "select * from accrpt211 order by r21106 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   Do While adoaccrpt211.EOF = False
'      adoaccrpt211.Fields("r21102").Value = intCounter
'      intCounter = intCounter + 1
'      adoaccrpt211.MoveNext
'   Loop
'   adoaccrpt211.UpdateBatch
'   adoaccrpt211.Close
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

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
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text5 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add By Sindy 2014/5/29
'*************************************************
' 產生Excel資料
'
'*************************************************
Public Sub PrintExcel()
Dim strFilePath As String
Dim dblSkipPageRow As Double
   
On Error GoTo ErrHnd
   
   If adoaccrpt211.State = adStateOpen Then
      adoaccrpt211.Close
   End If
   adoaccrpt211.CursorLocation = adUseClient
   adoaccrpt211.Open "select * from accrpt211 where R21101='" & strUserNum & "' order by R21102 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt211.RecordCount = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      adoaccrpt211.Close
      Exit Sub
   End If
   
   intPage = 0
   intCounter = 0
   dblSkipPageRow = 0
   
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
      .PageSetup.Orientation = xlPortrait '橫印xlLandscape,直印xlPortrait
      .PageSetup.LeftMargin = 28.34
      .PageSetup.RightMargin = 28.34
      .PageSetup.TopMargin = 42.51
      .PageSetup.BottomMargin = 42.51
      .PageSetup.HeaderMargin = 28.34
      .PageSetup.FooterMargin = 28.34
      '設定各欄位長度
      .Columns("A:A").ColumnWidth = 5 '名次
      .Columns("B:B").ColumnWidth = 10 '代理人編號
      .Columns("C:C").ColumnWidth = 15 '代理人名稱
      .Columns("D:D").ColumnWidth = 8  '國籍
      .Columns("E:E").ColumnWidth = 11 'FC往來
      .Columns("F:F").ColumnWidth = 11 'FC未收
      .Columns("G:G").ColumnWidth = 11 'CF往來
      .Columns("H:H").ColumnWidth = 11 'CF未付
      '逐筆填值
      adoaccrpt211.MoveFirst
      Do While adoaccrpt211.EOF = False
         intCounter = intCounter + 1
         '一開始或資料已填滿一頁時跳頁
         If intCounter = 1 Or dblSkipPageRow >= 42 Then
            If dblSkipPageRow >= 42 Then
               '換頁
               .Range("A" & intCounter).Select
               .HPageBreaks.add Before:=.Application.ActiveCell
            End If
            dblSkipPageRow = 0
            Call PrintExcelTitle
         End If
         If IsNull(adoaccrpt211.Fields("R21102").Value) = False Then
            .Range("A" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "0;[紅色]0"
            .Range("A" & intCounter).Value = CStr(adoaccrpt211.Fields("R21102").Value)
         End If
         If IsNull(adoaccrpt211.Fields("R21103").Value) = False Then
            .Range("B" & intCounter).Value = adoaccrpt211.Fields("R21103").Value
         End If
         If IsNull(adoaccrpt211.Fields("R21104").Value) = False Then
            .Range("C" & intCounter).Value = adoaccrpt211.Fields("R21104").Value
         End If
         If IsNull(adoaccrpt211.Fields("R21105").Value) = False Then
            .Range("D" & intCounter).Value = adoaccrpt211.Fields("R21105").Value
         End If
         If IsNull(adoaccrpt211.Fields("R21106").Value) = False Then
            .Range("E" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("E" & intCounter).Value = CStr(adoaccrpt211.Fields("R21106").Value)
         End If
         If IsNull(adoaccrpt211.Fields("R21107").Value) = False Then
            .Range("F" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("F" & intCounter).Value = CStr(adoaccrpt211.Fields("R21107").Value)
         End If
         If IsNull(adoaccrpt211.Fields("R21108").Value) = False Then
            .Range("G" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("G" & intCounter).Value = CStr(adoaccrpt211.Fields("R21108").Value)
         End If
         If IsNull(adoaccrpt211.Fields("R21109").Value) = False Then
            .Range("H" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("H" & intCounter).Value = CStr(adoaccrpt211.Fields("R21109").Value)
         End If
         dblSkipPageRow = dblSkipPageRow + 1
         adoaccrpt211.MoveNext
      Loop
      intCounter = intCounter + 1
      .Range("E" & intCounter).Value = "***結束***"
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
   adoaccrpt211.Close
   'Modify by Amy 2021/06/22 路徑改中文字顯示
   MsgBox "檔案已產生！" & vbCrLf & vbCrLf & "存放至 " & strExcelPathN & Replace(strFilePath, strExcelPath, "")
   Exit Sub
   
ErrHnd:
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   adoaccrpt211.Close
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2014/5/29
Public Sub PrintExcelTitle()
Dim i As Integer, strTemp As String
Dim strText As String
   
   intPage = intPage + 1
   With wksAnnuity
      For i = 1 To 2
         If i = 1 Then
            .Range("E" & intCounter).Value = "代理人帳目排名"
         ElseIf i = 2 Then
            intCounter = intCounter + 1
            If MaskEdBox1.Text <> MsgText(29) Then
               If strText <> "" Then strText = strText & "　"
               strText = strText & "往來日期：" & MaskEdBox1.Text & "~" & MaskEdBox2.Text
            End If
            .Range("E" & intCounter).Value = strText
         End If
         strTemp = "A" & intCounter & ":H" & intCounter
         .Range(strTemp).Select
         With .Application.Selection
            .HorizontalAlignment = xlCenter
         End With
         If i = 1 Then
            With .Application.Selection
               .Font.Size = 18
            End With
         End If
      Next i
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "列印人員：" & strUserName
      .Range("G" & intCounter).Value = "列印日期：" & ChangeWStringToTDateString(strSrvDate(1))
      intCounter = intCounter + 1
      .Range("G" & intCounter).Value = "頁　　次：" & intPage
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "名次"
      .Range("B" & intCounter).Value = "代理人編號"
      .Range("C" & intCounter).Value = "代理人名稱"
      .Range("D" & intCounter).Value = "國籍"
      .Range("E" & intCounter).Value = "FC往來"
      .Range("F" & intCounter).Value = "FC未收"
      .Range("G" & intCounter).Value = "CF往來"
      .Range("H" & intCounter).Value = "CF未付"
      strTemp = "A" & intCounter & ":H" & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       intCounter = intCounter + 1
   End With
End Sub
