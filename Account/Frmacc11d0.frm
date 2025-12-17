VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11d0 
   AutoRedraw      =   -1  'True
   Caption         =   "收據金額修改"
   ClientHeight    =   4212
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6048
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4212
   ScaleWidth      =   6048
   Begin VB.CommandButton Command1 
      Caption         =   "拆收據其他收據號資料"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2970
      TabIndex        =   26
      Top             =   2730
      Width           =   2500
   End
   Begin VB.TextBox Text12 
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
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      MaxLength       =   14
      TabIndex        =   4
      Top             =   2040
      Width           =   1605
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      MaxLength       =   14
      TabIndex        =   3
      Top             =   1680
      Width           =   1605
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   20
      Top             =   960
      Width           =   1605
   End
   Begin VB.TextBox Text2 
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
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   2
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   19
      Top             =   960
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   16
      Top             =   2760
      Width           =   1572
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   3000
      Picture         =   "Frmacc11d0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   14
      Top             =   2040
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   12
      Top             =   1680
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   10
      Top             =   1320
      Width           =   1572
   End
   Begin VB.TextBox Text1 
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
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSForms.TextBox Text14 
      Height          =   330
      Left            =   1320
      TabIndex        =   6
      Top             =   3480
      Width           =   4365
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "7699;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text13 
      Height          =   330
      Left            =   1320
      TabIndex        =   23
      Top             =   3120
      Width           =   4365
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   30
      Size            =   "7699;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   330
      Left            =   2880
      TabIndex        =   15
      Top             =   600
      Width           =   1605
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   30
      Size            =   "2831;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "收據備註"
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
      Left            =   360
      TabIndex        =   25
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Left            =   360
      TabIndex        =   24
      Top             =   3150
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "(Y：是)"
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
      Left            =   2070
      TabIndex        =   22
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "是否合併"
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
      Left            =   360
      TabIndex        =   21
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "業務區 "
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
      Left            =   360
      TabIndex        =   18
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
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
      Left            =   360
      TabIndex        =   17
      Top             =   2790
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   3900
      Left            =   120
      Top             =   120
      Width           =   5700
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "規費"
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
      Left            =   360
      TabIndex        =   13
      Top             =   2070
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "服務費"
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
      Left            =   360
      TabIndex        =   11
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
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
      Left            =   360
      TabIndex        =   9
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   8
      Top             =   630
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收文號"
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
      Left            =   360
      TabIndex        =   7
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc11d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adocaseprogress As New ADODB.Recordset
Public adoacc0j0 As New ADODB.Recordset
Public adoacc0k0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public strRelateNoList As String '相關收據號碼清單 Add by Morgan 2011/9/22
Public adoAssign As ADODB.Recordset '金額分配資料 Add by Morgan 2011/9/23
Public m_A0K11 As String, m_A0K03 As String 'Add By Sindy 2013/12/25


Private Sub Command1_Click()
   Dim ado0k0 As ADODB.Recordset
   'Modified by Morgan 2022/1/13 +a0k04
   strExc(0) = "select '',a0k01,sqldatet(a0k02),a0k04,a0k06||'',a0k07||'' from acc0k0 where a0k01 in ('" & Replace(strRelateNoList, ",", "','") & "') and a0k01<>'" & Text9 & "' order by 2"
   intI = 1
   Set ado0k0 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If ado0k0.RecordCount = 1 Then
         Frmacc1133.strNo = ado0k0.Fields("a0k01")
         Frmacc1133.Show vbModal
         strFormName = Me.Name
      Else
         Do
            Set Frmacc1132.grdDataList.Recordset = ado0k0.Clone
            Set Frmacc1132.fmParent = Me
            Frmacc1132.Show vbModal
            If Me.Tag <> "" Then
               Frmacc1133.strNo = Me.Tag
               Frmacc1133.Show vbModal
            End If
            strFormName = Me.Name
         Loop While Me.Tag <> ""
      End If
   End If
   Set ado0k0 = Nothing
End Sub

Private Sub Command3_Click()
'   If adocaseprogress.RecordCount = 0 Or Text1 = MsgText(601) Then
'      Exit Sub
'   End If
'   adocaseprogress.Find "cp09 = '" & Text1 & "'", 0, adSearchForward, 1
'   If adocaseprogress.EOF = False Then
'      FormShow
'      RecordShow
'   Else
'      Frmacc11d0_Clear
'      MsgBox MsgText(33), , MsgText(5)
'      adocaseprogress.MoveFirst
'   End If
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   CaseRefresh
'   If adocaseprogress.RecordCount <> 0 Then
'      FormShow
'      RecordShow
'   End If
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   'If adocaseprogress.RecordCount <> 0 Then
   '   adocaseprogress.MoveFirst
   'End If
   'adocaseprogress.Find "cp09 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adocaseprogress.EOF = False Then
   '   FormShow
   '   RecordShow
   'End If
   Text1 = strItemNo
   CaseRefresh
   If adocaseprogress.RecordCount <> 0 Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/05 原:W:6045/H:4560
   Me.Width = 6165
   Me.Height = 4680
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strItemNo = MsgText(601)
   OpenTable
   If adocaseprogress.RecordCount <> 0 Then
      RecordShow
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache '發信 Add By Sindy 2022/12/6
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11d0 = Nothing
End Sub

Private Sub Text1_Change()
   Command1.Enabled = False
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adocaseprogress.CursorLocation = adUseClient
   adocaseprogress.MaxRecords = intMax
   adocaseprogress.Open "select cp01, cp02, cp03, cp04, cp09, cp13, cp16, cp17, cp18, cp60 from caseprogress where substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and cp09 >= '" & Text1 & "' order by cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0j0.CursorLocation = adUseClient
   adoacc0j0.Open "select * from acc0j0 where a0j01 = 'Z' order by a0j01 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select * from acc0k0 where a0k01 = 'Z' order by a0k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(案件進度檔)
'
'*************************************************
Public Sub FormShow()
   m_A0K11 = adocaseprogress.Fields("A0K11").Value 'Add By Sindy 2013/12/25
   m_A0K03 = adocaseprogress.Fields("A0K03").Value 'Add By Sindy 2013/12/25
   Text1 = adocaseprogress.Fields("cp09").Value
   '2012/8/15 MODIFY BY SONIA 改抓收據智權人員, 且更新時不可改收文智權人員,否則收文智權人員離職改由他人催收帳款時會同時把收文智權人員改到,影響後續程序抓智權人員的原則 T-176273
'   If IsNull(adocaseprogress.Fields("cp13").Value) Then
'      Text2 = MsgText(601)
'   Else
'      Text2 = adocaseprogress.Fields("cp13").Value
'   End If
'   'Add By Cheng 2002/04/23
'   '顯示業務區
'   If IsNull(adocaseprogress.Fields("CP12").Value) Then
'      Me.Text10.Text = MsgText(601)
'   Else
'      Me.Text10.Text = adocaseprogress.Fields("CP12").Value
'   End If
   If IsNull(adocaseprogress.Fields("a0k20").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adocaseprogress.Fields("a0k20").Value
   End If
   '顯示業務區
   If IsNull(adocaseprogress.Fields("ST15").Value) Then
      Me.Text10.Text = MsgText(601)
   Else
      Me.Text10.Text = adocaseprogress.Fields("ST15").Value
   End If
   '2012/8/15 END
   '顯示業務區名稱
   If IsNull(adocaseprogress.Fields("A0902").Value) Then
      Me.Text11.Text = MsgText(601)
   Else
      Me.Text11.Text = adocaseprogress.Fields("A0902").Value
   End If
   
   If IsNull(adocaseprogress.Fields("cp01").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adocaseprogress.Fields("cp01").Value
      If IsNull(adocaseprogress.Fields("cp02").Value) = False Then
         Text3 = Text3 & adocaseprogress.Fields("cp02").Value
      End If
      If IsNull(adocaseprogress.Fields("cp03").Value) = False Then
         Text3 = Text3 & adocaseprogress.Fields("cp03").Value
      End If
      If IsNull(adocaseprogress.Fields("cp04").Value) = False Then
         Text3 = Text3 & adocaseprogress.Fields("cp04").Value
      End If
   End If
   If IsNull(adocaseprogress.Fields("cp16").Value) Then
      Text4 = "0"
      Text5 = "0"
   Else
      If IsNull(adocaseprogress.Fields("cp17").Value) Then
         Text4 = adocaseprogress.Fields("cp16").Value
         Text5 = adocaseprogress.Fields("cp16").Value
         Text6 = "0"
         Text7 = "0"
      Else
         Text4 = Val(adocaseprogress.Fields("cp16").Value) - Val(adocaseprogress.Fields("cp17").Value)
         Text5 = Val(adocaseprogress.Fields("cp16").Value) - Val(adocaseprogress.Fields("cp17").Value)
         Text6 = adocaseprogress.Fields("cp17").Value
         Text7 = adocaseprogress.Fields("cp17").Value
      End If
   End If
   If IsNull(adocaseprogress.Fields("cp60").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adocaseprogress.Fields("cp60").Value
   End If
   '顯示是否合併
   'Modified by Morgan 2011/11/24
   'If IsNull(adocaseprogress.Fields("A0K30").Value) Then
   If IsNull(adocaseprogress.Fields("A0j07").Value) Then
      Me.Text12.Text = MsgText(601)
   Else
      'Modified by Morgan 2011/11/24
      'Me.Text12.Text = adocaseprogress.Fields("A0K30").Value
      Me.Text12.Text = adocaseprogress.Fields("A0j07").Value
   End If
   '顯示收據抬頭
   If IsNull(adocaseprogress.Fields("A0K04").Value) Then
      Me.Text13.Text = MsgText(601)
   Else
      Me.Text13.Text = adocaseprogress.Fields("A0K04").Value
   End If
   '2007/8/9 ADD BY SONIA 收據備註
   If IsNull(adocaseprogress.Fields("a0k08").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = adocaseprogress.Fields("a0k08").Value
   End If
   '2007/8/9 END
   
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
'   If ExistCheck("caseprogress", "cp09", Text1, Label1) = False Then
'      Cancel = True
'      Exit Sub
'   End If
End Sub

Private Sub Text12_Change()
   Me.Text12.Text = UCase(Me.Text12.Text)
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
   CloseIme
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case 8, 89, 121
         '無動作
      Case Else
         KeyAscii = 0
   End Select
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
   OpenIme
End Sub

Private Sub Text14_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

Private Sub Text14_LostFocus()
   CloseIme
End Sub

Private Sub Text2_Change()
   Text8 = StaffQuery(Text2)
   'Add By Cheng 2002/04/23
   Me.Text10.Text = StaffDeptQuery(Me.Text2)
   Me.Text11.Text = A0902Query(Me.Text10)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adocaseprogress.Bookmark & MsgText(35) & adocaseprogress.RecordCount
End Sub

'*************************************************
'  重新整理國內收據資料
'
'*************************************************
Public Sub CaseRefresh()
On Error GoTo Checking
   If adocaseprogress.State = adStateOpen Then
      adocaseprogress.Close
   End If
   adocaseprogress.CursorLocation = adUseClient
   adocaseprogress.MaxRecords = intMax
   'Modify By Cheng 2002/04/23
'   adocaseprogress.Open "select cp01, cp02, cp03, cp04, cp09, cp13, cp16, cp17, cp18, cp60 from caseprogress where substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp09 >= '" & Text1 & "' order by cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '92.12.24 MODIFY BY SONIA
   'adocaseprogress.Open "select cp01, cp02, cp03, cp04, cp09, cp13, cp16, cp17, cp18, cp60, CP12,A0902,A0K30,A0K04 from caseprogress,ACC090,ACC0K0 where CP12=A0901(+) AND CP60=A0K01(+) AND substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp09 >= '" & Text1 & "' order by cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Morgan 2006/1/3 加cp75,cp77, 2007/8/9加A0K08
   '2010/2/11 MODIFY BY SONIA 不控制substr(cp01, 1, 2) <> 'FC'改為SUBSTR(CP60,1,1)='E'(FCP-037910總收文號A99005516)
   'adocaseprogress.Open "select cp01, cp02, cp03, cp04, cp09, cp13, cp16, cp17, cp18, cp60, CP12,A0902,A0K30,A0K04,A0K08,CP79,cp75,cp77 from caseprogress,ACC090,ACC0K0 where CP12=A0901(+) AND CP60=A0K01(+) AND substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp09 >= '" & Text1 & "' order by cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Morgan 2011/11/24 是否合併改抓 a0j07
   'adocaseprogress.Open "select cp01, cp02, cp03, cp04, cp09, cp13, cp16, cp17, cp18, cp60, CP12,A0902,A0K30,A0K04,A0K08,CP79,cp75,cp77 from caseprogress,ACC090,ACC0K0 where CP12=A0901(+) AND CP60=A0K01(+) AND SUBSTR(CP60,1,1)='E' and (cp16 <> 0 and cp16 is not null) and cp09 >= '" & Text1 & "' order by cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2012/8/15 MODIFY BY SONIA 改抓收據智權人員, 且更新時不可改收文智權人員,否則收文智權人員離職改由他人催收帳款時會同時把收文智權人員改到,影響後續程序抓智權人員的原則 T-176273
   'adocaseprogress.Open "select cp01, cp02, cp03, cp04, cp09, cp13, cp16, cp17, cp18, cp60, CP12,A0902,a0j07,A0K04,A0K08,CP79,cp75,cp77 from caseprogress,ACC090,ACC0K0,acc0j0 where CP12=A0901(+) AND CP60=A0K01(+) and a0j01(+)=cp09 and a0j13(+)=cp60 AND SUBSTR(CP60,1,1)='E' and (cp16 <> 0 and cp16 is not null) and cp09 >= '" & Text1 & "' order by cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2013/8/14 modify by sonia 加發文規費cp84,申請國家a0j04
   'adocaseprogress.Open "select cp01, cp02, cp03, cp04, cp09, A0K20, cp16, cp17, cp18, cp60, ST15,A0902,a0j07,A0K04,A0K08,CP79,cp75,cp77 from caseprogress,ACC090,ACC0K0,acc0j0,STAFF where CP60=A0K01(+) and A0K20=ST01(+) AND ST15=A0901(+) AND a0j01(+)=cp09 and a0j13(+)=cp60 AND SUBSTR(CP60,1,1)='E' and (cp16 <> 0 and cp16 is not null) and cp09 >= '" & Text1 & "' order by cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify By Sindy 2013/12/25 +A0K11,A0K03
   adocaseprogress.Open "select cp01, cp02, cp03, cp04, cp09, A0K20, cp16, cp17, cp18, cp60, ST15,A0902,a0j07,A0K04,A0K08,CP79,cp75,cp77,cp84,a0j04,A0K11,A0K03 from caseprogress,ACC090,ACC0K0,acc0j0,STAFF where CP60=A0K01(+) and A0K20=ST01(+) AND ST15=A0901(+) AND a0j01(+)=cp09 and a0j13(+)=cp60 AND SUBSTR(CP60,1,1)='E' and (cp16 <> 0 and cp16 is not null) and cp09 >= '" & Text1 & "' order by cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2012/8/5 END
   If adocaseprogress.RecordCount <> 0 Then
      If Text1 <> MsgText(601) Then
         adocaseprogress.Find "cp09 = '" & Text1 & "'", 0, adSearchForward, 1
         If adocaseprogress.EOF = False Then
            FormShow
            RecordShow
            'Add by Morgan 2011/9/22
            strRelateNoList = PUB_GetRelateNo2(Text1)
            If strRelateNoList <> Text9 Then
               Command1.Enabled = True
            Else
               Command1.Enabled = False
            End If
            Command3.Tag = Text1 'Add By Sindy 2024/5/8 記錄查詢資料的文號 ex:CFP-034309(AB3008307,AB3008337)
            'Add by Morgan 2011/9/27
            If adocaseprogress.Fields("cp77") > 0 Then
               MsgBox "本收文號有銷帳不可修改!!", vbExclamation
            End If
            
            Call PUB_Chk440(Text1.Text, "2") 'Added by Morgan 2025/5/22
            Exit Sub
         Else
            adocaseprogress.MoveFirst
         End If
      End If
   End If
   MsgBox MsgText(28), , MsgText(5)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'Add by Morgan 2011/9/22
'收據金額分配畫面
Public Function FeeReAssign() As Boolean
   
   Me.Tag = ""
   strExc(0) = "select a0j13,a0j09,a0j10,a0k04 from acc0j0,acc0k0 where a0j01='" & Text1 & "' and a0k01(+)=a0j13"
   Frmacc11d2.Text1 = Text1
   Frmacc11d2.Text5 = Text5
   Frmacc11d2.Text7 = Text7
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/07/01 +FormName 改暫存TB
   Set adoAssign = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   If intI = 1 Then
      Set Frmacc11d2.MSHFlexGrid2.Recordset = adoAssign
      Frmacc11d2.Show vbModal
      strFormName = Me.Name
   End If
   If Me.Tag = "Y" Then
      FeeReAssign = True
   End If
End Function

'Modify by Morgan 2011/8/12 清除a0k12相關程式(目前沒用,保留再使用)
Public Sub Frmacc11d0_Save()
Dim strYes As String
Dim m_strMailDesc As String, m_strMailSubject As String  '2011/8/17 ADD BY SONIA
Dim bolTrans As Boolean
'Add By Sindy 2013/12/26
Dim bolHaveAcc430 As Boolean, bolHaveAcc0M0 As Boolean
Dim strA4317 As String, strA4302 As String, strA4111 As String, strAX210 As String, strA4301 As String
Dim dblAmt As Double, dblTax As Double
'2013/12/26 END
Dim bolHasA4319 As Boolean 'Add by Amy 2022/11/10 已上傳
Dim strCaseNo As String, strNA01 As String, strCaseProperty As String
Dim GetMailSubject As String, GetMailContent As String
Dim strSqlLog As String
'Dim strReason As String
Dim strUpdDate As String, strUpdTime As String 'Add By Sindy 2023/5/18
Dim strModifyNote As String 'Add By Sindy 2023/5/18
   
On Error GoTo Checking
   strSqlLog = "" 'Add By Sindy 2022/12/6
   
   With Frmacc11d0
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If ExistCheck("caseprogress", "cp09", .Text1, .Label1) = False Then
            strControlButton = MsgText(602)
            .Text1.SetFocus
            Exit Sub
         End If
      End If
      
      'Add by Morgan 2011/9/27
      If .adocaseprogress.Fields("cp77") > 0 Then
         MsgBox "本收文號有銷帳不可修改!!", vbCritical
         strControlButton = MsgText(602)
         Exit Sub
      End If
      
      'Add By Sindy 2024/5/8 檢查要存檔的文號,是否有做過查詢動作
      If Text1 <> Command3.Tag Then
         MsgBox "此收文號(" & Text1 & ")尚未按過查詢按鈕(目前記錄的文號=" & Command3.Tag & ")!!!" & vbCrLf & vbCrLf & _
                "畫面資料與此收文號不符!", , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      End If
      '2024/5/8 END
      
      'Added by Morgan 2023/9/13
      '智權人員已繳款但財務尚未收款時
      'Modified by Morgan 2025/5/22 改呼叫公用模組
      'strSql = " select * from acc441 where axd05='" & .Text1 & "' and exists(select * from acc440 where a4401=axd01 and a4402=axd02 and a4403=axd03 and a4416 is null)"
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      'If intI = 1 Then
      '   MsgBox "智權已繳款但尚未收款，請先通知智權刪除繳款紀錄，待金額修改後請智權再次繳款!!", vbCritical
      '   strControlButton = MsgText(602)
      '   Exit Sub
      'End If
      If PUB_Chk440(Text1.Text, "2") Then
         strControlButton = MsgText(602)
         Exit Sub
      End If
      'end 2025/5/22
      'end 2023/9/13
      
      'Add By Sindy 2013/12/25
      bolHaveAcc430 = False
      bolHaveAcc0M0 = False
      bolHasA4319 = False 'Add by Amy 2022/11/10
      If .m_A0K11 = "J" Then
         strSql = "select * From acc431,acc430 where axc02='" & .Text9 & "' and axc01=a4301"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            bolHaveAcc430 = True '已開發票
            strA4317 = "" & RsTemp.Fields("a4317")
            strA4301 = "" & RsTemp.Fields("a4301") '發票號碼
            strA4302 = "" & RsTemp.Fields("a4302") '發票日期
            strA4111 = GetInvDataA4111(strA4302) '發票申報日期
            'Add by Amy 2022/11/10 +是否已上傳
            If "" & RsTemp.Fields("a4319") <> MsgText(601) Then
                bolHasA4319 = True
            End If
            strAX210 = ""
            '有未收款沖帳傳票
            If strA4317 <> "" Then
               strSql = "select * From acc021 where ax201='J' and ax202='" & strA4317 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strAX210 = "" & RsTemp.Fields("ax210") '傳票過帳日期
               End If
            End If
            If (strA4317 <> "" And Val(strAX210) > 0) Or Val(strA4111) > 0 Then
               '傳票已過帳或發票已申報
               If Val(.Text4) + Val(.Text6) <> Val(.Text5) + Val(.Text7) Then
                  MsgBox "J公司請款單沖轉傳票已過帳或發票已申報，只可改總額相同，即服務費規費互調!!", vbCritical
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
               MsgBox "已開發票，請自行控制發票收回!!", vbInformation
            End If
         End If
         '檢查是否已收款
         strSql = "select a0m01 From acc0m0 where a0m02='" & .Text9 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            bolHaveAcc0M0 = True '已收款
         End If
      End If
      '2013/12/25 END
      
      '2013/8/14 add by sonia 台灣案若有發文規費cp84則檢查並提醒但仍可選擇是否修改T-168087(AA2025092)
      If "" & .adocaseprogress.Fields("a0j04").Value = "000" And Not IsNull(.adocaseprogress.Fields("cp84").Value) Then
         If Val(.Text7) <> Val(.adocaseprogress.Fields("cp84").Value) Then
            If MsgBox("規費與發文規費 " & Val(.adocaseprogress.Fields("cp84").Value) & " 不符，請確認是否仍要更改??" & vbCrLf & vbCrLf, vbYesNo + vbDefaultButton1) = vbNo Then
               strControlButton = MsgText(602)
               .Text7.SetFocus
               Exit Sub
            End If
         End If
      End If
      '2013/8/14 end
      
      'Add by Morgan 2011/9/21 拆收據分配點數
      .strRelateNoList = PUB_GetRelateNo2(.Text1)
      If .strRelateNoList <> .Text9 Then
         .Command1.Enabled = True
         If .FeeReAssign = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
      End If
      'end 2011/9/21
     
   'Modify By Sindy 2022/12/6 往上移至此處
   adoTaie.BeginTrans
   bolTrans = True
   '2022/12/6 END
      
      'Add By Sindy 2023/1/12
      If .adoquery.State = adStateOpen Then
         .adoquery.Close
      End If
      '2023/1/12 END
      .adoquery.CursorLocation = adUseClient
      'Modify by Morgan 2011/9/26 考慮拆收據
      '.adoquery.Open "select * from acc0m0, caseprogress where a0m02 = cp60 and cp09 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      .adoquery.Open "select * from caseprogress,acc0j0,acc0m0 where cp09='" & .Text1 & "' and a0j01(+)=cp09 and a0m02(+) = a0j13 and a0m01 is not null", adoTaie, adOpenStatic, adLockReadOnly
      If .adoquery.RecordCount <> 0 Then
         MsgBox MsgText(177), , MsgText(5)
         '2011/8/17 ADD BY SONIA
         'Modify by Morgan 2011/9/26 考慮拆收據
         'm_strMailSubject = "收款後修改收據金額【" & .Text9 & ", " & .Text1 & "】"
         m_strMailSubject = "收款後修改收據金額【" & .strRelateNoList & ", " & .Text1 & "】"
         
         '2013/8/15 modify by sonia 原金額改用畫面值,否則AA2034464點數應為1.16但cp18存1.2會有誤差
         'm_strMailDesc = "原資料：服務費 " & (.adocaseprogress.Fields("cp18").Value * 1000) & ", 規費 " & .adocaseprogress.Fields("cp17").Value & _
                  vbCrLf & "修改後：服務費 " & Val(.Text5) & ", 規費 " & Val(.Text7) & _
                  vbCrLf & vbCrLf & "請調整收款資料：CASEPROGRESS,ACC0K0,ACC0L0,ACC0M0,ACC1U0,ACC1V0" & _
                  vbCrLf & vbCrLf & "同時確認傳票資料：ACC1P0,ACC021,ACC031(若傳票已過帳,則此三檔案不改)"
         'Modify By Sindy 2022/8/11 若沒有改金額就不要發EMAIL
         If Val(.Text4) <> Val(.Text5) Or Val(.Text6) <> Val(.Text7) Then
         '2022/8/11 END
            m_strMailDesc = "原資料：服務費 " & Val(.Text4) & ", 規費 " & Val(.Text6) & _
                     vbCrLf & "修改後：服務費 " & Val(.Text5) & ", 規費 " & Val(.Text7) & _
                     vbCrLf & vbCrLf & "請調整收款資料：CASEPROGRESS,ACC0K0,ACC0L0,ACC0M0,ACC1U0,ACC1V0" & _
                     vbCrLf & vbCrLf & "同時確認傳票資料：ACC1P0,ACC021,ACC031(若傳票已過帳,則此三檔案不改)"
            'Modify By Sindy 2022/12/6
            'PUB_SendMail strUserNum, "83002", "", m_strMailSubject, m_strMailDesc
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               " values( '" & strUserNum & "','" & Pub_GetSpecMan("程式管理人員") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & m_strMailSubject & "','" & m_strMailDesc & "',null)"
            cnnConnection.Execute strSql
            '2022/12/6 END
            '2011/8/17 END
         End If
      End If
      .adoquery.Close
      
'   adoTaie.BeginTrans
'   bolTrans = True
      
      'Add By Sindy 2022/12/6
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select * from caseprogress,acc0j0 where cp09='" & .Text1 & "' and a0j01(+)=cp09", adoTaie, adOpenStatic, adLockReadOnly
      If .adoquery.RecordCount <> 0 Then
      
         'Modify By Sindy 2023/12/12 統一使用函數
         If strSrvDate(1) >= 接洽單電子收文啟用日 And Len("" & .adoquery.Fields("cp140")) = 10 Then
            If Val(.Text4) <> Val(.Text5) Or Val(.Text6) <> Val(.Text7) Then
'               strReason = InputBox("請輸入修改費用的原因！（不可空白）")
'               If strReason = "" Then
               If Trim(Text14.Text) = "" Then
                  MsgBox "請輸入修改費用的原因；備註不可空白！", vbCritical + vbOKOnly
                  GoTo Checking
               End If
               strCaseNo = .adoquery.Fields("cp01") & "-" & .adoquery.Fields("cp02") & "-" & .adoquery.Fields("cp03") & "-" & .adoquery.Fields("cp04")
               strNA01 = GetPrjNation1(strCaseNo)
               If PUB_ModCrLCRCData(Text1, .adoquery.Fields("cp140"), "", .adoquery.Fields("cp10") _
                  , strNA01, , Val(Text5) + Val(Text7), Text7, Val(Text4) + (Text6), Text6, False, Text14.Text) = False Then
                  GoTo Checking
               End If
            End If
         End If
         'Add By Sindy 2022/12/5 修改費用,規費時; 發Mail通知正本財務處,副本智權人員
'         If strSrvDate(1) >= 接洽單電子收文啟用日 And Len("" & .adoquery.Fields("cp140")) = 10 Then
'            If Val(.Text4) <> Val(.Text5) Or Val(.Text6) <> Val(.Text7) Then
''               strReason = InputBox("請輸入修改費用的原因！（不可空白）")
''               If strReason = "" Then
'               If Trim(Text14.Text) = "" Then
'                  MsgBox "請輸入修改費用的原因；備註不可空白！", vbCritical + vbOKOnly
'                  GoTo Checking
'               End If
'
'               '案件性質名稱
'               strCaseNo = .adoquery.Fields("cp01") & "-" & .adoquery.Fields("cp02") & "-" & .adoquery.Fields("cp03") & "-" & .adoquery.Fields("cp04")
'               strNA01 = GetPrjNation1(strCaseNo)
'               If strNA01 = "000" Then
'                  Call ClsPDGetCasePropertyL(1, .adoquery.Fields("cp01"), .adoquery.Fields("cp10"), strCaseProperty)
'               Else
'                  Call ClsPDGetCasePropertyL(2, .adoquery.Fields("cp01"), .adoquery.Fields("cp10"), strCaseProperty)
'               End If
'
'               '接洽記錄單案件性質
'               strSql = "UPDATE ConsultRecCMP SET CRC04='有修改',CRC05='有修改',CRC06='有修改'" & _
'                        " WHERE CRC08='" & .Text1 & "'"
'               cnnConnection.Execute strSql
'
'               strUpdDate = strSrvDate(1)
'               strUpdTime = Right("000000" & ServerTime, 6)
'               strModifyNote = "原費用 " + CStr((Val(.Text4) + Val(.Text6))) + "元 修改為 " + CStr((Val(.Text5) + Val(.Text7))) + "元; " + _
'                  "原規費 " + CStr(Val(.Text6)) + "元 修改為 " + CStr(Val(.Text7)) + "元; 修改原因：" & Trim(Text14.Text) & "。"
'               '流程備註檔
'               strSql = GetInsertFLOW004Sql("" & .adoquery.Fields("cp140"), strUserNum, strUpdDate, strUpdTime, "", _
'                  strCaseProperty & "(" & .Text1 & ")" & strModifyNote)
'               cnnConnection.Execute strSql
'               'Add By Sindy 2023/5/18 記錄進度備註
'               strSql = "update caseprogress set cp64='" & Format(Val(strUpdDate) - 19110000, "###/##/##") & " " & _
'                        Format(strUpdTime, "##:##:##") & " " & strUserName & "：" & strModifyNote & "'||cp64" & _
'                        " where cp09='" & .Text1 & "'"
'               cnnConnection.Execute strSql
'               '2023/5/18 END
'
''               '郵件暫存記錄
''               GetMailSubject = strCaseNo & " 案(" & strCaseProperty & ")「" & GetPrjName(strCaseNo) & "」"
''               GetMailContent = "本所案號： " + strCaseNo + vbCrLf + _
''                "案件名稱： " + GetPrjName(strCaseNo) + vbCrLf + _
''                "收 文 日： " + ChangeTStringToTDateString(TransDate(.adoquery.Fields("cp05"), 1)) + vbCrLf + _
''                "案件性質： " + strCaseProperty + vbCrLf + vbCrLf
''               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
''                  " values( '" & strUserNum & "','" & Pub_GetSpecMan("財務處總帳人員") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
''                  ",'" & GetMailSubject & " 費用已異動" & "','" & GetMailContent + _
''                  "原費用： " + CStr((Val(.Text4) + Val(.Text6))) + " 元" + vbCrLf + _
''                  "修改費用： " + CStr((Val(.Text5) + Val(.Text7))) + " 元" + vbCrLf + _
''                  "原規費： " + CStr(Val(.Text6)) + " 元" + vbCrLf + _
''                  "修改規費： " + CStr(Val(.Text7)) + " 元" + vbCrLf & "','" & .adoquery.Fields("cp13") & "')"
''               cnnConnection.Execute strSql
'            End If
'            '2022/12/5 END
'         End If
         '2023/12/12 END
      End If
      .adoquery.Close
      '2022/12/6 END
      
'2012/8/15 CANCEL BY SONIA 改抓收據智權人員, 且更新時不可改收文智權人員,否則收文智權人員離職改由他人催收帳款時會同時把收文智權人員改到,影響後續程序抓智權人員的原則 T-176273
'      'Add By Cheng 2002/04/23
'      '更新智權人員
'      If .Text2.Text <> MsgText(601) Then
'         .adocaseprogress.Fields("CP13").Value = .Text2.Text
'      Else
'         .adocaseprogress.Fields("CP13").Value = Null
'      End If
'      '更新業務區
'      If .Text10.Text <> MsgText(601) Then
'         .adocaseprogress.Fields("CP12").Value = .Text10.Text
'      Else
'         .adocaseprogress.Fields("CP12").Value = Null
'      End If
'2012/8/5 END
      
      'Modify by Morgan 2006/1/2 只改規費也要更新
      'If .Text5 <> MsgText(601) Then
      If .Text5 <> MsgText(601) Or .Text7 <> MsgText(601) Then
         strSqlLog = "update caseprogress set " 'Add By Sindy 2022/12/6
         If .Text7 <> MsgText(601) Then
            .adocaseprogress.Fields("cp16").Value = Val(.Text5) + Val(.Text7)
            .adocaseprogress.Fields("cp17").Value = Val(.Text7)
            'Modify by Morgan 2010/1/5 統一進位到小數一位
            '.adocaseprogress.Fields("cp18").Value = Val(.Text5) / 1000
            .adocaseprogress.Fields("cp18").Value = Format(Val(.Text5) / 1000, "0.0")
            'Remove by Morgan 2006/1/2 移到下面
            '.adocaseprogress.Fields("cp79").Value = Val(.Text5) + Val(.Text7)
            '2006/1/2 end
         Else
            .adocaseprogress.Fields("cp16").Value = Val(.Text5)
            .adocaseprogress.Fields("cp17").Value = 0
            'Modify by Morgan 2010/1/5 統一進位到小數一位
            '.adocaseprogress.Fields("cp18").Value = Val(.Text5) / 1000
            .adocaseprogress.Fields("cp18").Value = Format(Val(.Text5) / 1000, "0.0")
            'Remove by Morgan 2006/1/2 移到下面
            '.adocaseprogress.Fields("cp79").Value = Val(.Text5)
            '2006/1/2 end
         End If
         strSqlLog = strSqlLog & "cp16=" & .adocaseprogress.Fields("cp16").Value 'Add By Sindy 2022/12/6
         strSqlLog = strSqlLog & ",cp17=" & .adocaseprogress.Fields("cp17").Value 'Add By Sindy 2022/12/6
         strSqlLog = strSqlLog & ",cp18=" & .adocaseprogress.Fields("cp18").Value 'Add By Sindy 2022/12/6
         'Add by Morgan 2006/1/2 需扣已收及銷帳金額
         '2006/2/3 MODIFY BY SONIA 解決NULL的問題
         '.adocaseprogress.Fields("cp79").Value = .adocaseprogress.Fields("cp16").Value - .adocaseprogress.Fields("cp75").Value - .adocaseprogress.Fields("cp77").Value
         .adocaseprogress.Fields("cp79").Value = .adocaseprogress.Fields("cp16").Value - Val("" & .adocaseprogress.Fields("cp75").Value) - Val("" & .adocaseprogress.Fields("cp77").Value)
         strSqlLog = strSqlLog & ",cp79=" & .adocaseprogress.Fields("cp79").Value 'Add By Sindy 2022/12/6
         '2006/2/3 END
         '2006/1/2 end
         
         'Modify by Morgan 2011/9/26 拆收據分配點數
         'adoTaie.Execute "update acc0j0 set a0j09 = " & Val(.Text5) & ", a0j10 = " & Val(.Text7) & " where a0j01 = '" & .Text1 & "'"
         If .strRelateNoList <> .Text9 Then
            .adoAssign.MoveFirst
            Do While Not .adoAssign.EOF
               If Val(.adoAssign.Fields(1)) = 0 And Val(.adoAssign.Fields(2)) = 0 Then
                  adoTaie.Execute "delete acc0j0 where a0j01 = '" & .Text1 & "' and a0j13='" & .adoAssign.Fields(0) & "'"
               Else
                  adoTaie.Execute "update acc0j0 set a0j09 = " & Val(.adoAssign.Fields(1)) & ", a0j10 = " & Val(.adoAssign.Fields(2)) & " where a0j01 = '" & .Text1 & "' and a0j13='" & .adoAssign.Fields(0) & "'"
               End If
               .adoAssign.MoveNext
            Loop
         Else
            adoTaie.Execute "update acc0j0 set a0j09 = " & Val(.Text5) & ", a0j10 = " & Val(.Text7) & " where a0j01 = '" & .Text1 & "'"
         End If
         'end 2011/9/26
      End If
      
      If .adoacc0k0.State = adStateOpen Then
         .adoacc0k0.Close
      End If
      .adoacc0k0.CursorLocation = adUseClient
      'Modify by Morgan 2011/9/26 考慮有拆收據情形改抓 0j0 並跑迴圈更新
      '.adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & .Text9 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      .adoacc0k0.Open "select * from acc0k0 where a0k01 in ('" & Replace(.strRelateNoList, ",", "','") & "')", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If .adoacc0k0.RecordCount <> 0 Then
         Do While Not .adoacc0k0.EOF
            .adoacc0k0.Fields("a0k06").Value = 0
            .adoacc0k0.Fields("a0k07").Value = 0
            
            If .adoacc0j0.State = adStateOpen Then
               .adoacc0j0.Close
            End If
            .adoacc0j0.CursorLocation = adUseClient
            'Modify by Morgan 2011/9/26
            '.adoacc0j0.Open "select * from acc0j0 where a0j13 = '" & .Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
            'Do While .adoacc0j0.EOF = False
            '   If IsNull(.adoacc0j0.Fields("a0j09").Value) = False Then
            '      .adoacc0k0.Fields("a0k06").Value = Val(.adoacc0k0.Fields("a0k06").Value) + Val(.adoacc0j0.Fields("a0j09").Value)
            '   End If
            '   If IsNull(.adoacc0j0.Fields("a0j10").Value) = False Then
            '      .adoacc0k0.Fields("a0k07").Value = Val(.adoacc0k0.Fields("a0k07").Value) + Val(.adoacc0j0.Fields("a0j10").Value)
            '   End If
            '   .adoacc0j0.MoveNext
            'Loop
            .adoacc0j0.Open "select nvl(sum(a0j09),0),nvl(sum(a0j10),0) from acc0j0 where a0j13 = '" & .adoacc0k0("a0k01") & "'", adoTaie, adOpenStatic, adLockReadOnly
            If .adoacc0j0.RecordCount > 0 Then
               .adoacc0k0.Fields("a0k06").Value = .adoacc0j0.Fields(0)
               .adoacc0k0.Fields("a0k07").Value = .adoacc0j0.Fields(1)
            End If
            'end 2011/9/26
            
'cancel by sonia 2018/5/22 取消A0K30,統一改用A0J07
'            'Add By Cheng 2002/04/23
'            If .Text12.Text <> MsgText(601) Then
'               .adoacc0k0.Fields("a0k30").Value = .Text12.Text
'            Else
'               .adoacc0k0.Fields("a0k30").Value = Null
'            End If
            If .Text2 <> MsgText(601) Then
               .adoacc0k0.Fields("a0k20").Value = .Text2
            Else
               .adoacc0k0.Fields("a0k20").Value = Null
            End If
            '2007/8/9 ADD BY SONIA
            If .Text14.Text <> MsgText(601) Then
               .adoacc0k0.Fields("a0k08").Value = .Text14.Text
            Else
               .adoacc0k0.Fields("a0k08").Value = Null
            End If
            '2007/8/9 END
            '2009/5/5 ADD BY SONIA
            If .Text10.Text <> MsgText(601) Then
               .adoacc0k0.Fields("a0k22").Value = .Text10.Text
            Else
               .adoacc0k0.Fields("a0k22").Value = Null
            End If
            '2009/5/5 END
            'Add by Morgan 2004/7/30
            '加更新人員日期時間
            .adoacc0k0.Fields("a0k29").Value = strUserNum
            .adoacc0k0.Fields("a0k27").Value = strSrvDate(2)
            .adoacc0k0.Fields("a0k28").Value = ServerTime
            
            .adoacc0k0.MoveNext
         Loop
         .adoacc0k0.UpdateBatch
      End If
      Pub_SeekTbLog strSqlLog & " where cp09='" & .Text1 & "'" 'Add By Sindy 2022/12/6
      .adocaseprogress.UpdateBatch
      
      'Modify by Morgan 2011/9/26 0k0上面已經更新不必重複
      'If .Text9 <> "" Then
      '   .adoquery.CursorLocation = adUseClient
      '   .adoquery.Open "select sum(nvl(cp16,0) - nvl(cp17,0)), sum(cp17) from caseprogress where cp60 = '" & .Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
      '   If .adoquery.RecordCount <> 0 Then
      '      adoTaie.Execute "update acc0k0 set a0k06 = " & .adoquery.Fields(0).Value & ", a0k07 = " & .adoquery.Fields(1).Value & ", a0k30 = " & CNULL(.Text12) & ", a0k20 = '" & .Text2 & "' where a0k01 = '" & .Text9 & "'"
      '      '2010/10/4 MODIFY BY SONIA 同時更新智權人員
      '      'adoTaie.Execute "update acc0j0 set a0j07 = " & CNULL(.Text12) & " where a0j13 = '" & .Text9 & "'"
      '      adoTaie.Execute "update acc0j0 set a0j05 = " & CNULL(.Text2) & ",a0j07 = " & CNULL(.Text12) & " where a0j13 = '" & .Text9 & "'"
      '   End If
      '   .adoquery.Close
      'End If
      'Modified by Morgan 2011/12/26 取消a0j05
      'Modified by Morgan 2014/2/25 一張收據會有可能包含合併及不合併的收文資料,故只需更新該收文號的資料
      'adoTaie.Execute "update acc0j0 a set a0j07 = " & CNULL(.Text12) & " where a0j13 in (select b.a0j13 from acc0j0 b where b.a0j01='" & .Text1 & "')"
      adoTaie.Execute "update acc0j0 a set a0j07 = " & CNULL(.Text12) & " where a0j01='" & .Text1 & "'"
      'end 2011/9/26
      
      'Add by Morgan 2006/1/2 1v0也要更新
      '合併
      
'Modify by Morgan 2011/10/21 考慮拆收據情形改寫(部分收款註記改與0k0一致,目前仍不考慮部分扣繳)
'      If .Text12.Text <> MsgText(601) Then
'         strExc(4) = 0.1 * (Val("" & .adocaseprogress("cp16")) - Val("" & .adocaseprogress("cp77")))
'      Else
'         strExc(4) = 0.1 * (Val("" & .adocaseprogress("cp16")) - Val("" & .adocaseprogress("cp17")) - Val("" & .adocaseprogress("cp77")))
'      End If
'      If Val("" & .adocaseprogress("cp79")) = 0 Then
'         strExc(5) = "N"
'      Else
'         strExc(5) = "Y"
'      End If
'      strSql = "update acc1v0 set a1v04 = " & strExc(4) & ", a1v05 = '" & strExc(5) & "', a1v07=" & strExc(4) & "-a1v06 where a1v01 = '" & .Text1 & "'"
'      adoTaie.Execute strSql
      strSql = "update acc1v0 set (a1v04,a1v05,a1v06,a1v07)=(" & _
            " select 0.1*(nvl(max(a0j09),0)-nvl(sum(a1u07),0)+decode(max(a0j07),'Y',nvl(max(a0j10),0)-nvl(sum(a1u09),0),0)) a1v04" & _
            ",nvl(max(a0k13),'N') a1v05,nvl(sum(a1u06),0) a1v06" & _
            ",0.1*(nvl(max(a0j09),0)-nvl(sum(a1u07),0)+decode(max(a0j07),'Y',nvl(max(a0j10),0)-nvl(sum(a1u09),0),0))-nvl(sum(a1u06),0) a1v07" & _
            " from acc0j0,acc0k0,acc1u0" & _
            " where a0j01=a1v01 and a0j13=a1v02 and a0k01(+)=a0j13 and a1u02(+)=a1v02 and a1u03(+)=a1v01)" & _
            " where a1v01='" & .Text1 & "'"
      adoTaie.Execute strSql, intI
'end 2011/10/21

      '2006/1/2 end
      
      'Add by Morgan 2011/9/30
      If .strRelateNoList <> .Text9 Then
         adoTaie.Execute "update caseprogress set cp60=(select min(a0j13) from acc0j0 where a0j01=cp09) where cp60 in ('" & Replace(.strRelateNoList, ",", "','") & "')", intI
      End If
      
      'Add By Sindy 2013/12/27 已開發票者,必須重新計算發票的銷售額及稅額
      If bolHaveAcc430 = True Then
         Dim strKey As String
         If .strRelateNoList <> .Text9 Then
            strKey = "'" & Trim(Replace(.strRelateNoList, ",", "','")) & "'"
         Else
            strKey = "'" & Trim(.Text9) & "'"
         End If
         dblAmt = 0
         strSql = "select axc01,sum(nvl(a0k06,0))+sum(nvl(a0k07,0))-sum(nvl(A1uAmt,0)) Amt" & _
                  " From acc0k0,(select a1u02,sum(nvl(a1u07,0))+sum(nvl(a1u09,0)) A1uAmt from acc1u0 where a1u02 in(" & strKey & ") group by a1u02),acc431" & _
                  " Where a0k01 in(" & strKey & ")" & _
                  " and a0k01=a1u02(+)" & _
                  " and a0k01=axc02(+)" & _
                  " group by axc01"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               dblAmt = RsTemp.Fields(1)
               dblTax = dblAmt - Round((dblAmt / 1.05), 0)
               '2014/12/2 modify by sonia 個人無稅額
               'strSql = "update acc430 set a4304=" & (dblAmt - dblTax) & ",a4305=" & dblTax & _
                     " where a4301='" & RsTemp.Fields(0) & "'"
               'Modify By Sindy 2017/1/9 a4305=decode(a4303,null,0," & dblTax & ") ==> a4305=decode(a4303,null,0,'00000000',0," & dblTax & ")
               strSql = "update acc430 set a4304=decode(a4303,null," & dblAmt & "," & (dblAmt - dblTax) & "),a4305=decode(a4303,null,0,'00000000',0," & dblTax & ") where a4301='" & RsTemp.Fields(0) & "'"
               '2014/12/2 end
               adoTaie.Execute strSql, intI
               RsTemp.MoveNext
            Loop
         End If
      End If
      '2013/12/27 END
      
      adoTaie.CommitTrans
      bolTrans = False
      
      'Add By Sindy 2024/10/15 收據金額修改時, 如因收據已列印過, 請提醒(本張收據金額已列印,請重新列印收據)
      If Val(.Text4) <> Val(.Text5) Or Val(.Text6) <> Val(.Text7) Then
         strSql = " select * from acc0k0 where a0k01='" & .Text9 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Val("" & RsTemp.Fields("a0k19")) > 0 Then
               MsgBox "本張" & IIf("" & RsTemp.Fields("a0k11") = "J", "請款單", "收據") & "( " & .Text9 & " )已列印過, 請重新列印收據!!", vbInformation, "提醒"
            End If
         End If
      End If
      '2024/10/15 END
      
      .RecordShow
      
      'Add By Sindy 2013/12/27
      '已開發票有未收款沖帳傳票且傳票未過帳 或 已收款
      If (strA4317 <> "" And Val(strAX210) = 0) Or bolHaveAcc0M0 = True Then
         If bolHaveAcc430 = True Then '已開發票
            '開啟frmacc1127畫面
            strItemNo = .Text9
            strCustNo = .m_A0K03
            strTitle = Me.Name
            Me.Enabled = False
            Screen.MousePointer = vbHourglass
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
            Frmacc1127.Text1.Enabled = False
            'Frmacc1127.cmdSave.Visible = False 'Modify By Sindy 2014/3/31 Mark
            'Add by Amy 2022/11/10 +bolHasA4319 已上傳不可按發票存檔鈕
            Frmacc1127.cmdSave.Enabled = True
            If bolHasA4319 = True Then
                Frmacc1127.cmdSave.Enabled = False
            End If
            Frmacc1127.Show
            Screen.MousePointer = vbDefault
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
            'MsgBox IIf((strA4317 <> "" And Val(strAX210) = 0), "已產生未收款沖帳傳票" & IIf(bolHaveAcc0M0 = True, "及", ""), "") & IIf(bolHaveAcc0M0 = True, "已收款", "") & "，請自行調整傳票內容!!", vbInformation
            If strA4317 <> "" And Val(strAX210) = 0 Then
               MsgBox "已產生未收款沖帳傳票，請自行調整傳票內容!!", vbInformation
            End If
         End If
      End If
      '2013/12/27 END
      
      PUB_SendMailCache '發信 Add By Sindy 2022/12/6
      
      MsgBox "修改完畢！" 'Add By Sindy 2023/1/12
      
      Exit Sub
Checking:
   
   If bolTrans = True Then
      adoTaie.RollbackTrans
   End If
   
   If Err.Number <> 0 Then MsgBox Err.Description, , MsgText(5)
   
   End With
End Sub
