VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11o5 
   AutoRedraw      =   -1  'True
   Caption         =   "特殊發票客戶資料維護"
   ClientHeight    =   2540
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2540
   ScaleWidth      =   9030
   Begin VB.TextBox Text5 
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
      Left            =   2130
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1830
      Width           =   1330
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
      Left            =   2130
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1440
      Width           =   560
   End
   Begin VB.CommandButton Command3 
      Default         =   -1  'True
      Height          =   300
      Left            =   3930
      Picture         =   "Frmacc11o5.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   180
      Width           =   350
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2130
      MaxLength       =   9
      TabIndex        =   0
      Top             =   180
      Width           =   1725
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "先開發票核准日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   90
      TabIndex        =   11
      Top             =   1860
      Width           =   2100
   End
   Begin MSForms.TextBox Text3 
      Height          =   320
      Left            =   2130
      TabIndex        =   4
      Top             =   600
      Width           =   6410
      VariousPropertyBits=   671105043
      MaxLength       =   30
      Size            =   "7223;529"
      Value           =   "Text"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox Text4 
      Height          =   320
      Left            =   2130
      TabIndex        =   5
      Top             =   1020
      Width           =   1730
      VariousPropertyBits=   671105043
      MaxLength       =   100
      Size            =   "7223;529"
      Value           =   "Text"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "（N:不開發票 A:開立發票請款 B:等通知後再開立發票請款）"
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
      Left            =   2640
      TabIndex        =   10
      Top             =   1470
      Width           =   5840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "特殊發票："
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
      Left            =   960
      TabIndex        =   9
      Top             =   1470
      Width           =   1130
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶名稱："
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
      Left            =   960
      TabIndex        =   8
      Top             =   620
      Width           =   1130
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
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
      Left            =   960
      TabIndex        =   7
      Top             =   1050
      Width           =   1130
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶編號："
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
      Left            =   960
      TabIndex        =   6
      Top             =   200
      Width           =   1130
   End
End
Attribute VB_Name = "Frmacc11o5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/08 Form2.0已修改
'Create by Sindy 2013/12/13
Option Explicit


'尋找
Private Sub Command3_Click()
Dim rs As New ADODB.Recordset

   If Text1 = MsgText(601) Then
      MsgBox "請輸入客戶編號!!!"
      Text1.SetFocus
      Exit Sub
   Else
      Text1 = Left(Text1 & "000000000", 9)
   End If
   
   '讀取客戶資料
   rs.CursorLocation = adUseClient
   'Modify By Sindy 2023/9/4 +,cu195
   rs.Open "select a0902,st02,cu01||cu02 as custno,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04,cu144,cu195" & _
           " from customer,staff,acc090" & _
           " where cu13=st01(+)" & _
           " and cu12=a0901(+) and cu01||cu02='" & Text1 & "'" & _
           " order by cu12,cu13,cu01", adoTaie, adOpenStatic, adLockReadOnly
   If rs.RecordCount > 0 Then
      '存在帶出資料
      Text2 = "" & rs.Fields("cu144")
      Text3 = "" & rs.Fields("cu04")
      Text4 = "" & rs.Fields("st02")
      Text5 = ChangeWStringToTString("" & rs.Fields("cu195")) 'Add By Sindy 2023/9/4
      Forms(0).Toolbar1.Buttons.Item(6).Enabled = True
   Else
      Text2 = ""
      Text3 = ""
      Text4 = ""
      Text5 = "" 'Add By Sindy 2023/9/4
      MsgBox "無此客戶!!!"
      Forms(0).Toolbar1.Buttons.Item(6).Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   Text1.Text = strCompanyNo
   Call Command3_Click
   strCompanyNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim intCounter As Integer
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9150
   Me.Height = 2940
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11o5 = Nothing
End Sub

Private Sub Text1_GotFocus()
   CloseIme
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Public Sub Text1_Validate(Cancel As Boolean)
   Cancel = False
   If Text1.Text <> "" And Left(Text1.Text, 1) <> "X" Then
      Text1 = Left(Text1 & "000000000", 9)
      MsgBox "客戶編號只可輸入X !!"
      Text1.SetFocus
      Cancel = True
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Sub Frmacc11o5_Clear()
   With Frmacc11o5
      .Text1.Text = ""
      .Text2.Text = ""
      .Text3.Text = ""
      .Text4.Text = ""
      .Text5.Text = "" 'Add By Sindy 2023/9/4
      .Text1.SetFocus
   End With
End Sub

'儲存
Public Sub Frmacc11o5_Save()
Dim Cancel As Boolean

On Error GoTo Checking
   
   Cancel = False
   Call Text2_Validate(Cancel)
   If Cancel = True Then
      Exit Sub
   End If
   
   'Add By Sindy 2023/9/4
   Cancel = False
   Call Text5_Validate(Cancel)
   If Cancel = True Then
      Exit Sub
   End If
   '2023/9/4 END
   
   With Frmacc11o5
      adoTaie.BeginTrans
      'Modify by Amy 2021/07/23 未觸發 Trigger,導致未更新修改人員/日期/時間
      'Modify By Sindy 2023/9/4 + & ",cu195=" & CNULL(DBDATE(Text5))
      strSql = "begin user_data.user_enabled:=1; Update Customer " & _
               "Set cu144=" & CNULL(Text2) & ",cu195=" & CNULL(DBDATE(Text5)) & _
               " Where cu01='" & Left(.Text1, 8) & "'; end;"
      'end 2021/07/23
      adoTaie.Execute strSql
      adoTaie.CommitTrans
      'Call Command3_Click
      Call Frmacc11o5_Clear
      Forms(0).Toolbar1.Buttons.Item(6).Enabled = False
Checking:
   If Err.Number = 0 Then
      Exit Sub
   Else
      adoTaie.RollbackTrans
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   Text2 = Trim(Text2)
   'Modify By Sindy 2023/9/4 + And Text2 <> "A" And Text2 <> "B"
   If Text2 <> "" And Text2 <> "N" And Text2 <> "A" And Text2 <> "B" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入 N 或 A 或 B 或 空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Text2_GotFocus
   End If
End Sub

'Add By Sindy 2023/9/4
Private Sub Text5_Validate(Cancel As Boolean)
    If Trim(Text5) <> MsgText(601) Then
        If CheckIsTaiwanDate(Text5) = False Then
            Cancel = True
        Else
            If ChkWorkDay(Val(Text5) + 19110000) = False Then
                MsgBox "先開發票核准日期需為工作日！"
                Text5.SetFocus
                Cancel = True
            End If
        End If
    Else
        If Text2 = "A" Or Text2 = "B" Then
            MsgBox "先開發票核准日期不可空白！"
            Text5.SetFocus
            Cancel = True
        End If
    End If
End Sub
