VERSION 5.00
Begin VB.Form Frmacc24f0 
   AutoRedraw      =   -1  'True
   Caption         =   "代理人逾期帳款分析表"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   5130
   Begin VB.TextBox Text14 
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   23
      Top             =   2190
      Width           =   612
   End
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
      TabIndex        =   22
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text13 
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
      Left            =   4200
      TabIndex        =   12
      Top             =   1560
      Width           =   612
   End
   Begin VB.TextBox Text12 
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
      Left            =   3600
      TabIndex        =   11
      Top             =   1560
      Width           =   612
   End
   Begin VB.TextBox Text11 
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
      Left            =   3000
      TabIndex        =   10
      Top             =   1560
      Width           =   612
   End
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
      Left            =   2400
      TabIndex        =   9
      Top             =   1560
      Width           =   612
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
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   612
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
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
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
      Height          =   315
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
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
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
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
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   612
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
      TabIndex        =   13
      Top             =   2940
      Visible         =   0   'False
      Width           =   2145
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
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
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
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   840
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
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   480
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
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin VB.Label Label9 
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
      Left            =   360
      TabIndex        =   24
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "(空白表全部)"
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
      Left            =   1800
      TabIndex        =   21
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "代理人國籍："
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
      TabIndex        =   20
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(1:帳齡 2:金額)"
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
      TabIndex        =   19
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "列印順序："
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
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "元者(外幣)"
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
      Left            =   3480
      TabIndex        =   17
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "帳款金額超過："
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
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "個月者"
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
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "帳款帳齡超過："
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
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Frmacc24f0"
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

Public adoacc1k0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt215 As New ADODB.Recordset
Dim intPage As Integer
Dim intCounter As Integer
Dim strName As String
Dim strAmount As String
Dim PLeft(0 To 6) As Integer 'Add By Sindy 2012/12/11
'Add By Sindy 2014/6/6
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
'2014/6/6 END


'產生Excel檔
Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt215Delete
   ProduceData
   PrintExcel
   If strCon10 <> MsgText(602) Then
      FormClear
   End If
   Screen.MousePointer = vbDefault
   StatusView MsgText(156)
End Sub

'列印
Private Sub Command2_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt215Delete
   ProduceData
   PrintData
   If strCon10 <> MsgText(602) Then
      FormClear
   End If
   Screen.MousePointer = vbDefault
   StatusView MsgText(156)
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
      StatusView MsgText(156)
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
'   Me.Height = 3550
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
   PUB_InitForm Me, 5250, 3550, strBackPicPath4
   'end 2021/12/09
   
   StatusView MsgText(156)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc24f0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
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

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
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
Dim strAnd As String
Dim strFagent As String
Dim strCustomer As String
Dim datDate As Date
Dim lngDate As Long
Dim strMonth As String
Dim strDay As String
Dim i As Integer 'Add By Sindy 2012/12/11
   
On Error GoTo Checking
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/23 清除查詢印表記錄檔欄位
   StatusView MsgText(26)
   adoaccrpt215.CursorLocation = adUseClient
   adoaccrpt215.Open "select * from accrpt215", adoTaie, adOpenDynamic, adLockBatchOptimistic
   strAnd = " and ("
   If Text4 <> MsgText(601) Or _
      Text5 <> MsgText(601) Or _
      Text6 <> MsgText(601) Or _
      Text7 <> MsgText(601) Or _
      Text8 <> MsgText(601) Or _
      Text9 <> MsgText(601) Or _
      Text10 <> MsgText(601) Or _
      Text11 <> MsgText(601) Or _
      Text12 <> MsgText(601) Or _
      Text13 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label7 & Text4 'Add By Sindy 2010/12/23
   End If
   If Text4 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text4 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text4 & "' or "
   End If
   If Text5 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text5 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text5 & "' or "
      pub_QL05 = pub_QL05 & "," & Text5 'Add By Sindy 2010/12/23
   End If
   If Text6 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text6 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text6 & "' or "
      pub_QL05 = pub_QL05 & "," & Text6 'Add By Sindy 2010/12/23
   End If
   If Text7 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text7 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text7 & "' or "
      pub_QL05 = pub_QL05 & "," & Text7 'Add By Sindy 2010/12/23
   End If
   If Text8 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text8 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text8 & "' or "
      pub_QL05 = pub_QL05 & "," & Text8 'Add By Sindy 2010/12/23
   End If
   If Text9 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text9 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text9 & "' or "
      pub_QL05 = pub_QL05 & "," & Text9 'Add By Sindy 2010/12/23
   End If
   If Text10 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text10 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text10 & "' or "
      pub_QL05 = pub_QL05 & "," & Text10 'Add By Sindy 2010/12/23
   End If
   If Text11 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text11 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text11 & "' or "
      pub_QL05 = pub_QL05 & "," & Text11 'Add By Sindy 2010/12/23
   End If
   If Text12 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text12 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text12 & "' or "
      pub_QL05 = pub_QL05 & "," & Text12 'Add By Sindy 2010/12/23
   End If
   If Text13 <> MsgText(601) Then
      strFagent = strFagent & "substr(fa10,1,3) = '" & Text13 & "' or "
      strCustomer = strCustomer & "substr(cu10,1,3) = '" & Text13 & "' or "
      pub_QL05 = pub_QL05 & "," & Text13 'Add By Sindy 2010/12/23
   End If
   If strFagent <> MsgText(601) Then
      strFagent = Mid(strFagent, 1, Len(strFagent) - 4) & ") "
      strCustomer = Mid(strCustomer, 1, Len(strCustomer) - 4) & ")"
   End If
   If Text2 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label4 & Text2 & Label1 'Add By Sindy 2010/12/23
   End If
   datDate = CDate(Format(ServerDate, ADFormat)) - Val(Text2) * 30
   If Month(datDate) > 9 Then
      strMonth = Month(datDate)
   Else
      strMonth = "0" & Month(datDate)
   End If
   If Day(datDate) > 9 Then
      strDay = Day(datDate)
   Else
      strDay = "0" & Day(datDate)
   End If
   lngDate = (Year(datDate) - 1911) & strMonth & strDay
   strCon10 = ""
   adoacc1k0.CursorLocation = adUseClient
   If Text1 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label2 & Text1 & Label3 'Add By Sindy 2010/12/23
   End If
   'Add By Sindy 2014/6/9
   If Text14 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label9 & Text14
   End If
   '2014/6/9 END
   If Text4 = MsgText(601) Then
      '2010/6/28 MODIFY BY SONIA 剔除銷帳資料
      '2011/4/29 modify by sonia 加a1k01否則同日同金額之請款會被distinct
      'adoacc1k0.Open "select a1k08, a1k03, a1k02, fa10, fa04, fa05, fa06 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null) union " & _
                     "select a1k08, a1k03, a1k02, cu10 as fa10, cu04 as fa04, cu05 as fa05, cu06 as fa06 from acc1k0, customer where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null) order by a1k02 asc", adoTaie, adOpenStatic, adLockReadOnly
      'Modify By Sindy 2014/6/9 + IIf(Text14 = "", "", " and a1k13='" & Text14 & "')
      'Modified by Lydia 2017/02/23
      'adoacc1k0.Open "select a1k01, a1k08, a1k03, a1k02, fa10, fa04, fa05, fa06 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & IIf(Text14 = "", "", " and a1k13='" & Text14 & "'") & " union " & _
                     "select a1k01, a1k08, a1k03, a1k02, cu10 as fa10, cu04 as fa04, cu05 as fa05, cu06 as fa06 from acc1k0, customer where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & IIf(Text14 = "", "", " and a1k13='" & Text14 & "'") & " order by a1k02 asc", adoTaie, adOpenStatic, adLockReadOnly
      strSql = "select a1k01, a1k08, a1k03, a1k02, fa10, fa04, fa05, fa06 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & IIf(Text14 = "", "", " and a1k13='" & Text14 & "'") & " union " & _
               "select a1k01, a1k08, a1k03, a1k02, cu10 as fa10, cu04 as fa04, cu05 as fa05, cu06 as fa06 from acc1k0, customer where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & IIf(Text14 = "", "", " and a1k13='" & Text14 & "'") & " order by a1k02 asc"
   Else
   
      '2010/6/28 MODIFY BY SONIA 剔除銷帳資料
      '2011/4/29 modify by sonia 加a1k01否則同日同金額之請款會被distinct
      'adoacc1k0.Open "select a1k08, a1k03, a1k02, fa10, fa04, fa05, fa06 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & strAnd & strFagent & " union " & _
                     "select a1k08, a1k03, a1k02, cu10 as fa10, cu04 as fa04, cu05 as fa05, cu06 as fa06 from acc1k0, customer where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & strAnd & strCustomer & " order by a1k02 asc", adoTaie, adOpenStatic, adLockReadOnly
      'Modify By Sindy 2014/6/9 + IIf(Text14 = "", "", " and a1k13='" & Text14 & "')
      'Modified by Lydia 2017/02/23
      'adoacc1k0.Open "select a1k01, a1k08, a1k03, a1k02, fa10, fa04, fa05, fa06 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & strAnd & strFagent & IIf(Text14 = "", "", " and a1k13='" & Text14 & "'") & " union " & _
                     "select a1k01, a1k08, a1k03, a1k02, cu10 as fa10, cu04 as fa04, cu05 as fa05, cu06 as fa06 from acc1k0, customer where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & strAnd & strCustomer & IIf(Text14 = "", "", " and a1k13='" & Text14 & "'") & " order by a1k02 asc", adoTaie, adOpenStatic, adLockReadOnly
      strSql = "select a1k01, a1k08, a1k03, a1k02, fa10, fa04, fa05, fa06 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & strAnd & strFagent & IIf(Text14 = "", "", " and a1k13='" & Text14 & "'") & " union " & _
               "select a1k01, a1k08, a1k03, a1k02, cu10 as fa10, cu04 as fa04, cu05 as fa05, cu06 as fa06 from acc1k0, customer where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null)" & strAnd & strCustomer & IIf(Text14 = "", "", " and a1k13='" & Text14 & "'") & " order by a1k02 asc"
   End If
   
   'Added by Lydia 2017/02/23
   adoacc1k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   
   If adoacc1k0.RecordCount = 0 Then
      strCon10 = MsgText(602)
      adoacc1k0.Close
      adoaccrpt215.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc1k0.EOF = False
      If adoaccrpt215.State = adStateOpen Then
         adoaccrpt215.Close
      End If
      adoaccrpt215.CursorLocation = adUseClient
      'Modify By Sindy 2012/12/11 新增時才更新資料,不逐筆加總了
      adoaccrpt215.Open "select * from accrpt215 where r21501 = '" & strUserNum & "' and r21505 = '" & adoacc1k0.Fields("a1k03").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoaccrpt215.RecordCount = 0 Then
         adoaccrpt215.AddNew
         FagentSave
      'End If
''      adoaccrpt215.Find "r21501 = '" & strUserNum & "'", 0, adSearchForward, 1
''      If adoaccrpt215.EOF Then
''         adoaccrpt215.AddNew
''         FagentSave
''      Else
''         adoaccrpt215.Find "r21505 = '" & adoacc1k0.Fields("a1k03").Value & "'", 0, adSearchForward, adoaccrpt215.Bookmark
''         If adoaccrpt215.EOF Then
''            adoaccrpt215.AddNew
''            FagentSave
''         End If
''      End If
'      If IsNull(adoacc1k0.Fields("a1k08").Value) = False Then
'         adoaccrpt215.Fields("r21508").Value = Val(adoaccrpt215.Fields("r21508").Value) + Val(adoacc1k0.Fields("a1k08").Value)
'      End If
'      adoaccsum.CursorLocation = adUseClient
'      '2010/6/28 MODIFY BY SONIA
'      'adoaccsum.Open "select sum(a1506) from acc150 where a1503 = '" & adoacc1k0.Fields("a1k03").Value & "' and a1506 > nvl(a1520, 0)", adoTaie, adOpenStatic, adLockReadOnly
'      adoaccsum.Open "select sum(a1506) from acc150 where a1503 = '" & adoacc1k0.Fields("a1k03").Value & "' and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccsum.RecordCount <> 0 Then
'         If IsNull(adoaccsum.Fields(0).Value) = False Then
'            '2010/6/28 MODIFY BY SONIA 會重覆
'            'adoaccrpt215.Fields("r21509").Value = Val(adoaccrpt215.Fields("r21509").Value) + Val(adoaccsum.Fields(0).Value)
'            adoaccrpt215.Fields("r21509").Value = Val(adoaccsum.Fields(0).Value)
'         End If
'      End If
'      adoaccsum.Close
'      adoaccrpt215.UpdateBatch
         
         'Modify By Sindy 2012/12/11 改為一次讀取此代理人的應收資料
         adoaccsum.CursorLocation = adUseClient
         'Modified by Lydia 2017/02/23 +系統別
         'adoaccsum.Open "select sum(a1k08),a1k18 from acc1k0 where a1k03='" & adoacc1k0.Fields("a1k03").Value & "' and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null) group by a1k18", adoTaie, adOpenStatic, adLockReadOnly
         strSql = "select sum(a1k08),a1k18 from acc1k0 where a1k03='" & adoacc1k0.Fields("a1k03").Value & "'" & IIf(Text14 <> "", " and a1k13=" & CNULL(Text14), "") & " and a1k02 <= " & lngDate & " and a1k08 >= " & Val(Text1) & " and (a1k29 is null) and (a1k12 is null) and (a1k25 is null) group by a1k18"
         adoaccsum.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         'end 2017/02/23
         
         i = 0
         If adoaccsum.RecordCount <> 0 Then
            adoaccsum.MoveFirst
            Do While adoaccsum.EOF = False
               i = i + 1
               If IsNull(adoaccsum.Fields(0).Value) = False Then
                  If i > 1 Then
                     adoaccrpt215.AddNew
                     FagentSave
                  End If
                  adoaccrpt215.Fields("r21508").Value = Val(adoaccsum.Fields(0).Value)
                  adoaccrpt215.Fields("r21510").Value = adoaccsum.Fields(1).Value
                  adoaccrpt215.UpdateBatch
               End If
               adoaccsum.MoveNext
            Loop
         End If
         adoaccsum.Close
         'Modify By Sindy 2012/12/11 改為一次讀取此代理人的CF未付資料
         adoaccsum.CursorLocation = adUseClient
         'Modified by Lydia 2017/02/23 +系統別
         'adoaccsum.Open "select sum(a1506),a1505 from acc150 where a1503 = '" & adoacc1k0.Fields("a1k03").Value & "' and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1505", adoTaie, adOpenStatic, adLockReadOnly
         '" & IIf(Text14 <> "", " and a1k13=" & CNULL(Text14), "") & "
         strSql = "select sum(axf04) amt,a1505 from acc150,acc151 where a1501=axf01(+) and a1503 = '" & adoacc1k0.Fields("a1k03").Value & "' and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null " & _
                  IIf(Text14 <> "", " and substr(axf03,1,length(axf03)-9)=" & CNULL(Text14), "") & "group by a1505"
         adoaccsum.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         'end 2017/02/23
         If adoaccsum.RecordCount <> 0 Then
            adoaccsum.MoveFirst
            Do While adoaccsum.EOF = False
               If IsNull(adoaccsum.Fields(0).Value) = False Then
                  If adoaccrpt215.State = adStateOpen Then
                     adoaccrpt215.Close
                  End If
                  adoaccrpt215.CursorLocation = adUseClient
                  adoaccrpt215.Open "select * from accrpt215 where r21501 = '" & strUserNum & "' and r21505 = '" & adoacc1k0.Fields("a1k03").Value & "' and r21510 = '" & adoaccsum.Fields(1).Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
                  If adoaccrpt215.RecordCount > 0 Then
                     adoaccrpt215.Fields("r21509").Value = Val(adoaccsum.Fields(0).Value)
                     adoaccrpt215.Fields("r21511").Value = adoaccsum.Fields(1).Value
                     adoaccrpt215.UpdateBatch
                  Else
                     If adoaccrpt215.State = adStateOpen Then
                        adoaccrpt215.Close
                     End If
                     adoaccrpt215.CursorLocation = adUseClient
                     adoaccrpt215.Open "select * from accrpt215 where r21501 = '" & strUserNum & "' and r21505 = '" & adoacc1k0.Fields("a1k03").Value & "' and r21511 is null", adoTaie, adOpenDynamic, adLockBatchOptimistic
                     If adoaccrpt215.RecordCount > 0 Then
                        adoaccrpt215.Fields("r21509").Value = Val(adoaccsum.Fields(0).Value)
                        adoaccrpt215.Fields("r21511").Value = adoaccsum.Fields(1).Value
                        adoaccrpt215.UpdateBatch
                     Else
                        adoaccrpt215.AddNew
                        FagentSave
                        adoaccrpt215.Fields("r21509").Value = Val(adoaccsum.Fields(0).Value)
                        adoaccrpt215.Fields("r21511").Value = adoaccsum.Fields(1).Value
                        adoaccrpt215.UpdateBatch
                     End If
                  End If
               End If
               adoaccsum.MoveNext
            Loop
         End If
         adoaccsum.Close
      End If
      '2012/12/11 End
      adoacc1k0.MoveNext
   Loop
   adoacc1k0.Close
   If adoaccrpt215.State = adStateOpen Then
      adoaccrpt215.Close
   End If
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
Private Sub Accrpt215Delete()
   adoTaie.Execute "delete from accrpt215"
End Sub

'*************************************************
'  代理人資料儲存
'
'*************************************************
Private Sub FagentSave()
   adoaccrpt215.Fields("r21501").Value = strUserNum
   If Text2 <> MsgText(601) Then
      adoaccrpt215.Fields("r21502").Value = Val(Text2)
   Else
      adoaccrpt215.Fields("r21502").Value = Null
   End If
   If Text1 <> MsgText(601) Then
      adoaccrpt215.Fields("r21503").Value = Val(Text1)
   Else
      adoaccrpt215.Fields("r21503").Value = Null
   End If
   If IsNull(adoacc1k0.Fields("fa10").Value) Then
      adoaccrpt215.Fields("r21504").Value = Null
   Else
      adoaccrpt215.Fields("r21504").Value = Mid(adoacc1k0.Fields("fa10").Value, 1, 3)
   End If
   If IsNull(adoacc1k0.Fields("a1k03").Value) Then
      adoaccrpt215.Fields("r21505").Value = Null
   Else
      adoaccrpt215.Fields("r21505").Value = adoacc1k0.Fields("a1k03").Value
   End If
   If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
      adoaccrpt215.Fields("r21506").Value = adoacc1k0.Fields("fa05").Value
   Else
      If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
         adoaccrpt215.Fields("r21506").Value = adoacc1k0.Fields("fa04").Value
      Else
         If IsNull(adoacc1k0.Fields("fa06").Value) = False Then
            adoaccrpt215.Fields("r21506").Value = adoacc1k0.Fields("fa06").Value
         Else
            adoaccrpt215.Fields("r21506").Value = Null
         End If
      End If
   End If
   If IsNull(adoacc1k0.Fields("a1k02").Value) Then
      adoaccrpt215.Fields("r21507").Value = Null
   Else
      adoaccrpt215.Fields("r21507").Value = adoacc1k0.Fields("a1k02").Value
   End If
   adoaccrpt215.Fields("r21508").Value = 0
   adoaccrpt215.Fields("r21509").Value = 0
   'Add By Sindy 2012/12/11
   adoaccrpt215.Fields("r21510").Value = Null '應收幣別
   adoaccrpt215.Fields("r21511").Value = Null 'CF幣別
   '2012/12/11 End
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text2 = ""
   Text1 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text12 = ""
   Text13 = ""
   Text2.SetFocus
End Sub

'*************************************************
' 列印明細資料
'
'*************************************************
Private Sub PrintData()
Dim intRow As Integer
   
   'add by nickc  2007/02/08
   Dim strNo As String
   
   intPage = 0
   intCounter = 0
   strNo = ""
   'Modify By Sindy 2013/1/14
   adoaccrpt215.CursorLocation = adUseClient
   Select Case Text3
      Case "2"
         pub_QL05 = pub_QL05 & ";" & Label5 & "2:金額" 'Add By Sindy 2010/12/23
         adoaccrpt215.Open "select r21504,r21505,min(r21507),sum(r21508) from accrpt215  where r21501 = '" & strUserNum & "' group by r21504,r21505 order by sum(r21508)", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         pub_QL05 = pub_QL05 & ";" & Label5 & "1:帳齡" 'Add By Sindy 2010/12/23
         adoaccrpt215.Open "select r21504,r21505,min(r21507),sum(r21508) from accrpt215  where r21501 = '" & strUserNum & "' group by r21504,r21505 order by min(r21507)", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   If adoaccrpt215.RecordCount > 0 Then
      adoaccrpt215.MoveFirst
      intRow = 0
      Do While Not adoaccrpt215.EOF
         intRow = intRow + 1
         'update排序
         strSql = "update accrpt215 set r21512=" & intRow & _
                  " where r21501='" & strUserNum & "' " & _
                  "and r21504='" & adoaccrpt215.Fields(0) & "' " & _
                  "and r21505='" & adoaccrpt215.Fields(1) & "' "
         cnnConnection.Execute strSql
         adoaccrpt215.MoveNext
      Loop
   End If
   adoaccrpt215.Close
   adoaccrpt215.CursorLocation = adUseClient
   adoaccrpt215.Open "select * from accrpt215 where r21501 = '" & strUserNum & "' order by substr(r21504, 1, 3) asc, r21512 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2013/1/14 End
   InsertQueryLog (adoaccrpt215.RecordCount) 'Add By Sindy 2010/12/23
   Do While adoaccrpt215.EOF = False
      If strNo <> Mid(adoaccrpt215.Fields("r21504").Value, 1, 3) Then
         If strNo <> "" Then
            Printer.NewPage
         End If
         intCounter = 0
         intPage = intPage + 1
         PrintHead
         'intCounter = intCounter + 1
         strNo = Mid(adoaccrpt215.Fields("r21504").Value, 1, 3)
      End If
      If intCounter > 40 Then
         Printer.NewPage
         intCounter = 0
         intPage = intPage + 1
         PrintHead
         'intCounter = intCounter + 1
      End If
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = 2000 + intCounter * 300
      If IsNull(adoaccrpt215.Fields("r21505").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt215.Fields("r21505").Value
      End If
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = 2000 + intCounter * 300
      If IsNull(adoaccrpt215.Fields("r21506").Value) Then
         Printer.Print ""
      Else
         Printer.Print StrConv(MidB(StrConv(adoaccrpt215.Fields("r21506").Value, vbFromUnicode), 1, 20), vbUnicode)
      End If
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = 2000 + intCounter * 300
      If IsNull(adoaccrpt215.Fields("r21507").Value) Then
         Printer.Print ""
      Else
         Printer.Print CFDate(adoaccrpt215.Fields("r21507").Value)
      End If
      'Add By Sindy 2012/12/11 +應收幣別
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = 2000 + intCounter * 300
      If IsNull(adoaccrpt215.Fields("r21510").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt215.Fields("r21510").Value
      End If
      '2012/12/11 End
      If IsNull(adoaccrpt215.Fields("r21508").Value) = False Then
         strAmount = Format(adoaccrpt215.Fields("r21508").Value, DDollar)
         Printer.CurrentX = PLeft(4) - Printer.TextWidth(strAmount)
         Printer.CurrentY = 2000 + intCounter * 300
         Printer.Print strAmount
      End If
      'Add By Sindy 2012/12/11 +未付幣別
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = 2000 + intCounter * 300
      If IsNull(adoaccrpt215.Fields("r21511").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt215.Fields("r21511").Value
      End If
      '2012/12/11 End
      If IsNull(adoaccrpt215.Fields("r21509").Value) = False Then
         strAmount = Format(adoaccrpt215.Fields("r21509").Value, DDollar)
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(strAmount)
         Printer.CurrentY = 2000 + intCounter * 300
         Printer.Print strAmount
      End If
      intCounter = intCounter + 1
      adoaccrpt215.MoveNext
   Loop
   adoaccrpt215.Close
   Printer.EndDoc
End Sub

'Add By Sindy 2012/12/11
Private Sub GetPleft()
   PLeft(0) = 500
   PLeft(1) = 1750
   PLeft(2) = 4800
   PLeft(3) = 6000
   PLeft(4) = 8200
   PLeft(5) = 8500
   PLeft(6) = 10500
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
   GetPleft 'Add By Sindy 2012/12/11
   
   Printer.FontSize = 16
   Printer.CurrentX = 3000
   Printer.CurrentY = 1000
   Printer.Print ReportTitle(215)
   Printer.FontSize = 12
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "頁次: " & intPage
   intCounter = intCounter + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print Label4 & " " & Text2 & " " & Label1
   intCounter = intCounter + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print Label2 & " " & Text1 & " " & Label3
   Printer.CurrentX = 5000
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "代理人國籍: " & IIf(IsNull(adoaccrpt215.Fields("r21504").Value), "", Mid(adoaccrpt215.Fields("r21504").Value, 1, 3) & " " & NationQuery(Mid(adoaccrpt215.Fields("r21504").Value, 1, 3), 1))
   
   intCounter = intCounter + 1
   'Printer.Line (0, 2000 + intCounter * 300 - 10)-(11000, 2000 + intCounter * 300 - 10)
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print String(140, "-")
   intCounter = intCounter + 1
   
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "應收最早"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("  應　　收")
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "  應　　收"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth("　  CF 　")
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "　  CF 　"
   
   
   intCounter = intCounter + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "代理人"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "代理人名稱"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "請款日期"
   
   
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("外幣請款總額")
   Printer.CurrentY = 2000 + intCounter * 300
   'Modify By Sindy 2012/12/11
   'Printer.Print "美金請款總額"
   Printer.Print "外幣請款總額"
   '2012/12/11 End
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth("未付款金額")
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print "未付款金額"
   
   intCounter = intCounter + 1
   'Printer.Line (0, 2000 + intCounter * 300 - 10)-(11000, 2000 + intCounter * 300 - 10)
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = 2000 + intCounter * 300
   Printer.Print String(140, "-")
   intCounter = intCounter + 1
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
   If Text11 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text12 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text13 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add By Sindy 2014/6/6
'*************************************************
' 產生Excel資料
'
'*************************************************
Public Sub PrintExcel()
Dim strFilePath As String
Dim dblSkipPageRow As Double
Dim strNo As String, intRow As Integer
   
On Error GoTo ErrHnd
   
   'Modify By Sindy 2013/1/14
   adoaccrpt215.CursorLocation = adUseClient
   Select Case Text3
      Case "2"
         pub_QL05 = pub_QL05 & ";" & Label5 & "2:金額" 'Add By Sindy 2010/12/23
         adoaccrpt215.Open "select r21504,r21505,min(r21507),sum(r21508) from accrpt215  where r21501 = '" & strUserNum & "' group by r21504,r21505 order by sum(r21508)", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         pub_QL05 = pub_QL05 & ";" & Label5 & "1:帳齡" 'Add By Sindy 2010/12/23
         adoaccrpt215.Open "select r21504,r21505,min(r21507),sum(r21508) from accrpt215  where r21501 = '" & strUserNum & "' group by r21504,r21505 order by min(r21507)", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   If adoaccrpt215.RecordCount > 0 Then
      adoaccrpt215.MoveFirst
      intRow = 0
      Do While Not adoaccrpt215.EOF
         intRow = intRow + 1
         'update排序
         strSql = "update accrpt215 set r21512=" & intRow & _
                  " where r21501='" & strUserNum & "' " & _
                  "and r21504='" & adoaccrpt215.Fields(0) & "' " & _
                  "and r21505='" & adoaccrpt215.Fields(1) & "' "
         cnnConnection.Execute strSql
         adoaccrpt215.MoveNext
      Loop
   End If
   adoaccrpt215.Close
   
   If adoaccrpt215.State = adStateOpen Then
      adoaccrpt215.Close
   End If
   adoaccrpt215.CursorLocation = adUseClient
   adoaccrpt215.Open "select * from accrpt215 where r21501 = '" & strUserNum & "' order by substr(r21504, 1, 3) asc, r21512 asc", adoTaie, adOpenStatic, adLockReadOnly
   InsertQueryLog (adoaccrpt215.RecordCount) 'Add By Sindy 2010/12/23
   '2013/1/14 End
   If adoaccrpt215.RecordCount = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      adoaccrpt215.Close
      Exit Sub
   End If
   
   intPage = 0
   intCounter = 0
   dblSkipPageRow = 0
   strNo = ""
   
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
   xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
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
      .Columns("A:A").ColumnWidth = 10 '代理人
      .Columns("B:B").ColumnWidth = 20 '代理人名稱
      .Columns("C:C").ColumnWidth = 10 '應收最早請款日期
      .Columns("D:D").ColumnWidth = 8  '幣別
      .Columns("E:E").ColumnWidth = 14 '應收外幣請款總額
      .Columns("F:F").ColumnWidth = 8 '幣別
      .Columns("G:G").ColumnWidth = 14 'CF未付款金額
      '逐筆填值
      adoaccrpt215.MoveFirst
      Do While adoaccrpt215.EOF = False
         intCounter = intCounter + 1
         '一開始或資料已填滿一頁或代理人國籍不同時跳頁
         If intCounter = 1 Or dblSkipPageRow >= 40 Or _
            (strNo <> "" And strNo <> Mid(adoaccrpt215.Fields("r21504").Value, 1, 3)) Then
            If intCounter > 1 Then
               '換頁
               .Range("A" & intCounter).Select
               .HPageBreaks.add Before:=.Application.ActiveCell
            End If
            dblSkipPageRow = 0
            Call PrintExcelTitle
            strNo = Mid(adoaccrpt215.Fields("r21504").Value, 1, 3)
         End If
         If IsNull(adoaccrpt215.Fields("r21505").Value) = False Then
            .Range("A" & intCounter).Value = CStr(adoaccrpt215.Fields("r21505").Value)
         End If
         If IsNull(adoaccrpt215.Fields("r21506").Value) = False Then
            .Range("B" & intCounter).Value = StrConv(MidB(StrConv(adoaccrpt215.Fields("r21506").Value, vbFromUnicode), 1, 20), vbUnicode)
         End If
         If IsNull(adoaccrpt215.Fields("r21507").Value) = False Then
            .Range("C" & intCounter).Value = CFDate(adoaccrpt215.Fields("r21507").Value)
         End If
         If IsNull(adoaccrpt215.Fields("r21510").Value) = False Then
            .Range("D" & intCounter).Value = adoaccrpt215.Fields("r21510").Value
         End If
         If IsNull(adoaccrpt215.Fields("r21508").Value) = False Then
            .Range("E" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("E" & intCounter).Value = CStr(adoaccrpt215.Fields("r21508").Value)
         End If
         If IsNull(adoaccrpt215.Fields("r21511").Value) = False Then
            .Range("F" & intCounter).Value = adoaccrpt215.Fields("r21511").Value
         End If
         If IsNull(adoaccrpt215.Fields("r21509").Value) = False Then
            .Range("G" & intCounter).Select
            .Application.Selection.NumberFormatLocal = "#,##0_ "
            .Range("G" & intCounter).Value = CStr(adoaccrpt215.Fields("r21509").Value)
         End If
         dblSkipPageRow = dblSkipPageRow + 1
         adoaccrpt215.MoveNext
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
   adoaccrpt215.Close
   'Modify by Amy 2021/06/22 路徑改中文字顯示
   MsgBox "檔案已產生！" & vbCrLf & vbCrLf & "存放至 " & strExcelPathN & Replace(strFilePath, strExcelPath, "")
   Exit Sub
   
ErrHnd:
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   adoaccrpt215.Close
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2014/6/6
Public Sub PrintExcelTitle()
Dim strTemp As String
Dim strText As String
   
   intPage = intPage + 1
   With wksAnnuity
      .Range("D" & intCounter).Value = Me.Caption
      strTemp = "A" & intCounter & ":G" & intCounter
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter
         .Font.Size = 18
      End With
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "列印人員：" & strUserName
      .Range("F" & intCounter).Value = "列印日期：" & ChangeWStringToTDateString(strSrvDate(1))
      intCounter = intCounter + 1
      .Range("F" & intCounter).Value = "頁　　次：" & intPage
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "帳款帳齡超過：" & Text2
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "帳款金額超過：" & Text1 & "元者(外幣)"
      .Range("C" & intCounter).Value = "代理人國籍：" & IIf(IsNull(adoaccrpt215.Fields("r21504").Value), "", Mid(adoaccrpt215.Fields("r21504").Value, 1, 3) & " " & NationQuery(Mid(adoaccrpt215.Fields("r21504").Value, 1, 3), 1))
      .Range("F" & intCounter).Value = "系統類別：" & Text14
      intCounter = intCounter + 1
      .Range("C" & intCounter).Value = "應收最早"
      .Range("E" & intCounter).Value = "應　　收"
      .Range("G" & intCounter).Value = "CF"
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "代理人"
      .Range("B" & intCounter).Value = "代理人名稱"
      .Range("C" & intCounter).Value = "請款日期"
      .Range("D" & intCounter).Value = "幣別"
      .Range("E" & intCounter).Value = "外幣請款總額"
      .Range("F" & intCounter).Value = "幣別"
      .Range("G" & intCounter).Value = "未付款金額"
      strTemp = "A" & intCounter & ":G" & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       intCounter = intCounter + 1
   End With
End Sub
