VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc12d1 
   AutoRedraw      =   -1  'True
   Caption         =   "寄發E-Mail"
   ClientHeight    =   5290
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   8930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5290
   ScaleWidth      =   8930
   Begin VB.OptionButton Option1 
      Caption         =   "本筆帳款註明（管控日期）預定收款，今已逾期，請留意速收款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   300
      TabIndex        =   30
      Top             =   3810
      Width           =   8595
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   4680
      MaxLength       =   15
      TabIndex        =   28
      Top             =   1500
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   26
      Top             =   1500
      Width           =   1572
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   4680
      MaxLength       =   15
      TabIndex        =   24
      Top             =   60
      Width           =   1572
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   20
      Top             =   60
      Width           =   1572
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "傳送"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7770
      TabIndex        =   15
      Top             =   210
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本發票已於（　年　月　日）開立，至今未收款，請留意速收款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   300
      TabIndex        =   14
      Top             =   3480
      Width           =   8595
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本筆帳款（管控日期）註明待銷帳，至今仍未銷，請速處理"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   300
      TabIndex        =   13
      Top             =   3150
      Width           =   8595
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本筆帳款至今尚未發文，請問本帳款是否仍會辦理或有考慮銷帳，若要銷帳請速填銷帳單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   300
      TabIndex        =   12
      Top             =   2820
      Width           =   8595
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本筆帳款已於（　年　月　日）發文，款項尚未收回，請問預估何時收回或如何處理"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   300
      TabIndex        =   11
      Top             =   2490
      Width           =   8595
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "自行輸入說明"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   10
      Top             =   2160
      Value           =   -1  'True
      Width           =   8595
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   8
      Top             =   420
      Width           =   645
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   4680
      MaxLength       =   15
      TabIndex        =   1
      Top             =   420
      Width           =   825
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1140
      Width           =   4125
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   780
      Width           =   1572
      _ExtentX        =   2769
      _ExtentY        =   564
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
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
      Height          =   315
      Left            =   4680
      TabIndex        =   22
      Top             =   780
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   564
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   315
      Left            =   1440
      TabIndex        =   23
      Top             =   1140
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   564
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
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
   Begin MSForms.TextBox Text5 
      Height          =   825
      Left            =   270
      TabIndex        =   18
      Top             =   4410
      Width           =   8565
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "15108;1455"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   315
      Left            =   5520
      TabIndex        =   17
      Top             =   420
      Width           =   1545
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   15
      Size            =   "2725;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "未收金額："
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
      TabIndex        =   29
      Top             =   1530
      Width           =   1125
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "案件性質："
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
      TabIndex        =   27
      Top             =   1530
      Width           =   1125
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "本所案號："
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
      TabIndex        =   25
      Top             =   90
      Width           =   1125
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "收據編號："
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
      TabIndex        =   21
      Top             =   90
      Width           =   1125
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "信件內容： (勾黃底選項才能輸入內容)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4170
      Width           =   6735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "制式Mail內容："
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
      TabIndex        =   16
      Top             =   1890
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據公司："
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
      TabIndex        =   9
      Top             =   450
      Width           =   1125
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   450
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "發文日期："
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
      TabIndex        =   6
      Top             =   810
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "控管日期："
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
      TabIndex        =   5
      Top             =   1170
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   270
      Top             =   3570
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label29 
      BackStyle       =   0  '透明
      Caption         =   "發票日期："
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
      TabIndex        =   4
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label23 
      BackStyle       =   0  '透明
      Caption         =   "控管類別："
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
      TabIndex        =   3
      Top             =   1170
      Width           =   1155
   End
End
Attribute VB_Name = "Frmacc12d1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Create by Sindy 2016/6/13
Option Explicit

Private adoquery As New ADODB.Recordset
Public m_Appl As String '申請人
Public m_A0K04 As String '收據抬頭


'寄E-Mail
Private Sub cmdSend_Click()
Dim strSubject As String
Dim strContext As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
   If Option1(0).Value = True Then
      If Text5 = "" Then
         MsgBox "信件內容不可空白", , MsgText(5)
         Exit Sub
      End If
   End If
   
   If MsgBox("確定要發E-Mail嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
      Exit Sub
   End If
   
   strCP01 = Left(Text7, Len(Text7) - 9)
   strCP02 = Mid(Text7, Len(strCP01) + 1, 6)
   If Right(Text7, 3) <> "000" Then
      strCP03 = Left(Right(Text7, 3), 1)
      strCP04 = Right(Right(Text7, 3), 2)
   End If
   '主旨
   strSubject = strCP01 & "-" & strCP02 & IIf(Right(Text7, 3) <> "000", "-" & strCP03 & "-" & strCP04, "") & "(收據號碼" & Text6 & ")應收帳款提醒"
   '內容
   strContext = "申 請 人：" & m_Appl & vbCrLf
   strContext = strContext & "收據抬頭：" & m_A0K04 & vbCrLf
   strContext = strContext & "案件性質：" & Text8 & vbCrLf
   strContext = strContext & "未收金額：" & Text9
   If Option1(1).Value = True Then
      strContext = strContext & vbCrLf & vbCrLf & Option1(1).Caption
   ElseIf Option1(2).Value = True Then
      strContext = strContext & vbCrLf & vbCrLf & Option1(2).Caption
   ElseIf Option1(3).Value = True Then
      strContext = strContext & vbCrLf & vbCrLf & Option1(3).Caption
   ElseIf Option1(4).Value = True Then
      strContext = strContext & vbCrLf & vbCrLf & Option1(4).Caption
   'Add By Sindy 2017/11/16
   ElseIf Option1(5).Value = True Then
      strContext = strContext & vbCrLf & vbCrLf & Option1(5).Caption
   '2017/11/16 END
   End If
   'Modify By Sindy 2024/12/27
   'strContext = strContext & vbCrLf & vbCrLf & IIf(Text5.Enabled = True, Text5, "")
   strContext = strContext & vbCrLf & vbCrLf & IIf(Text5.Locked = False, Text5, "")
   '2024/12/27 END
   
   PUB_SendMail strUserNum, Text2, "", strSubject, strContext
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9045
   Me.Height = 5700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
      For intY = 0 To Int(ScaleHeight / sglHeight)
         PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
      Next
   Next
   FormShowE
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strCon1 = ""
   StatusClear
   tool3_enabled
   Select Case strFormLink
      Case "Frmacc12d0"
         Frmacc12d0.Enabled = True
   End Select
   Set Frmacc12d1 = Nothing
End Sub

'************************************************
'  顯示資料表
'
'************************************************
Private Sub FormShowE()
   adoquery.CursorLocation = adUseClient
   'Modify By Sindy 2024/11/27 + ,CP162
   strSql = "select a0k11,a0k20,st02,a0k38,a0k39,cp27,CP162 from acc0k0,staff,acc0j0,caseprogress" & _
             " where a0k01='" & strItemNo & "' and a0k20=st01(+)" & _
             " and a0j13(+)=a0k01 and cp09(+)=a0j01"
   adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   Text6.Text = strItemNo
   If adoquery.RecordCount <> 0 Then
      Text1.Text = "" & adoquery.Fields("a0k11")
      'Modify By Sindy 2024/11/27 法律所案件若有介紹人則收件者預設為介紹人，無介紹人才預設智權人員。
      If "" & adoquery.Fields("CP162") <> "" Then
         strExc(0) = "select LawOfficeSource.*,st02" & _
                     " from LawOfficeSource,staff" & _
                     " where LOS15='" & adoquery.Fields("CP162") & "' and substr(los04,1,5)=ST01(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Label2.Caption = "介紹人："
            Text2.Text = "" & RsTemp.Fields("los04")
            Text4.Text = "" & RsTemp.Fields("st02")
         End If
      Else
      '2024/11/27 END
         Text2.Text = "" & adoquery.Fields("a0k20")
         Text4.Text = "" & adoquery.Fields("st02")
      End If
      '發文日期
      MaskEdBox1.Mask = MsgText(601)
      Option1(1).Caption = "本筆帳款已於（　年　月　日）發文，款項尚未收回，請問預估何時收回或如何處理"
      If IsNull(adoquery.Fields("cp27").Value) Then
         MaskEdBox1.Text = MsgText(601)
      Else
         MaskEdBox1.Text = CFDate(Val(adoquery.Fields("cp27").Value) - 19110000)
         Option1(1).Caption = "本筆帳款已於（" & MaskEdBox1.Text & "）發文，款項尚未收回，請問預估何時收回或如何處理"
      End If
      MaskEdBox1.Mask = DFormat
      Text3.Text = "" & adoquery.Fields("a0k39")
      '控管日期
      MaskEdBox3.Mask = MsgText(601)
      Option1(3).Caption = "本筆帳款（　年　月　日）註明待銷帳，至今仍未銷，請速處理"
      Option1(5).Caption = "本筆帳款註明（　年　月　日）預定收款，今已逾期，請留意速收款"
      If IsNull(adoquery.Fields("a0k38").Value) Then
         MaskEdBox3.Text = MsgText(601)
      Else
         MaskEdBox3.Text = CFDate(adoquery.Fields("a0k38").Value)
         'Add By Sindy 2017/6/20
         If Text3 = "待銷帳" Then
         '2017/6/20 END
            Option1(3).Caption = "本筆帳款（" & MaskEdBox3.Text & "）註明待銷帳，至今仍未銷，請速處理"
         'Add By Sindy 2017/6/20
         ElseIf Text3 = "預計收款" Then
            Option1(5).Caption = "本筆帳款註明（" & MaskEdBox3.Text & "）預定收款，今已逾期，請留意速收款"
         End If
         '2017/6/20 END
      End If
      MaskEdBox3.Mask = DFormat
   End If
   '發票日期
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   strSql = "select a4302 from acc431,acc430" & _
             " where axc02='" & strItemNo & "'" & _
             " and axc01=a4301(+)"
   adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   MaskEdBox2.Mask = MsgText(601)
   Option1(4).Caption = "本發票已於（　年　月　日）開立，至今未收款，請留意速收款"
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("a4302").Value) Then
         MaskEdBox2.Text = MsgText(601)
      Else
         MaskEdBox2.Text = CFDate(adoquery.Fields("a4302").Value)
         Option1(4).Caption = "本發票已於（" & MaskEdBox2.Text & "）開立，至今未收款，請留意速收款"
      End If
   End If
   MaskEdBox2.Mask = DFormat
   adoquery.Close
End Sub

'Add By Sindy 2017/6/20
Private Sub Option1_Click(Index As Integer)
   If Index <> 0 Then
      Text5.Text = Option1(Index).Caption
      'Modify By Sindy 2024/12/27
      'Text5.Enabled = False
      Text5.Locked = True
      '2024/12/27 END
   Else
      'Modify By Sindy 2024/12/27
      'Text5.Enabled = True
      Text5.Locked = False
      '2024/12/27 END
   End If
End Sub

