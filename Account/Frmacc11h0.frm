VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11h0 
   AutoRedraw      =   -1  'True
   Caption         =   "扣繳憑單修正作業"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3390
   ScaleWidth      =   4800
   Begin VB.TextBox txtPreChar 
      BackColor       =   &H00E0E0E0&
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
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "K"
      Top             =   240
      Width           =   270
   End
   Begin VB.TextBox txtA0w04 
      BackColor       =   &H00E0E0E0&
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
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtA0w01 
      BackColor       =   &H00E0E0E0&
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
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   585
   End
   Begin VB.TextBox txtA0w14 
      BackColor       =   &H00E0E0E0&
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
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2760
      Width           =   3195
   End
   Begin VB.TextBox txtA0w13 
      BackColor       =   &H00E0E0E0&
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
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1572
   End
   Begin VB.CommandButton cmdFind 
      Default         =   -1  'True
      Height          =   300
      Left            =   2850
      Picture         =   "Frmacc11h0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   247
      Width           =   350
   End
   Begin VB.TextBox txtA0w05 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1125
   End
   Begin VB.TextBox txtA0w02 
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
      Left            =   1575
      MaxLength       =   8
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin MSForms.TextBox txtA0w06 
      Height          =   330
      Left            =   1305
      TabIndex        =   7
      Top             =   2040
      Width           =   3195
      VariousPropertyBits=   671105051
      BackColor       =   14737632
      MaxLength       =   2000
      Size            =   "5636;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA0w03 
      Height          =   330
      Left            =   1305
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   3195
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "5636;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  '透明
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
      Left            =   1890
      TabIndex        =   18
      Top             =   1350
      Width           =   2550
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
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
      Top             =   2820
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "補扣日期"
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
      TabIndex        =   15
      Top             =   2460
      Width           =   930
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
      Height          =   3105
      Left            =   240
      Top             =   120
      Width           =   4380
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Top             =   2070
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "扣單金額"
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
      TabIndex        =   13
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   360
      TabIndex        =   12
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
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
      TabIndex        =   11
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "扣單編號"
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
      TabIndex        =   10
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc11h0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/30 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Create by Morgan 2005/6/8
Option Explicit


Private Sub cmdFind_Click()
   If txtA0w02 <> "" Then
      strSql = "SELECT * FROM ACC0W0 WHERE A0W02='" & txtPreChar & txtA0w02 & "'"
   Else
      Exit Sub
   End If
   
On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         If "" & .Fields("a0w15") <> "" Then
            MsgBox "扣單已作廢！", vbInformation
         Else
            txtA0w01 = "" & .Fields("a0w01")
            txtA0w03 = "" & .Fields("a0w03")
            txtA0w04 = "" & .Fields("a0w04")
            txtA0w05 = "" & .Fields("a0w05")
            txtA0w06 = "" & .Fields("a0w06")
            txtA0w13 = "" & .Fields("a0w13")
            txtA0w14 = "" & .Fields("a0w14")
         End If
      Else
         MsgBox "無此單號！", vbExclamation
         txtA0w02.SetFocus
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
      
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 4900
   Me.Height = 4125
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
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11h0 = Nothing
End Sub

Public Function EditCheck() As Boolean
   If txtA0w01 <> "" Then
      EditCheck = True
   End If
End Function

Public Sub FormLocked(Optional ByVal p_Choice As Boolean = False)
   txtA0w02.Locked = p_Choice
   txtA0w04.Locked = Not p_Choice
   txtA0w06.Locked = Not p_Choice
   If p_Choice = True Then
      txtA0w02.BackColor = &HE0E0E0
      txtA0w04.BackColor = &H80000005
      txtA0w06.BackColor = &H80000005
   Else
      txtA0w02.BackColor = &H80000005
      txtA0w04.BackColor = &HE0E0E0
      txtA0w06.BackColor = &HE0E0E0
   End If
End Sub

Private Sub FormClear()
   txtA0w01 = ""
   txtA0w03 = ""
   txtA0w04 = ""
   txtA0w05 = ""
   txtA0w06 = ""
   txtA0w13 = ""
   txtA0w14 = ""
End Sub
Private Sub txtA0w02_Change()
   If txtA0w01 <> "" Then
      FormClear
   End If
End Sub

Private Sub txtA0w02_GotFocus()
   If txtA0w02.Locked = False Then
      'edit by nickc 2007/06/11  切換輸入法改用API
      'txtA0w02.IMEMode = 2
      CloseIme
      TextInverse txtA0w02
   End If
End Sub

Private Sub txtA0w02_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtA0w04_Change()
   If txtA0w04 <> "" Then
      lblCompany = A0802Query(txtA0w04)
   Else
      lblCompany = ""
   End If
End Sub

Private Sub txtA0w04_GotFocus()
   If txtA0w04.Locked = False Then
      'edit by nickc 2007/06/11  切換輸入法改用API
      'txtA0w04.IMEMode = 2
      CloseIme
      TextInverse txtA0w04
   End If
End Sub

Private Sub txtA0w04_KeyPress(KeyAscii As Integer)
   'Modify By Sindy 2020/5/15 + And Not (KeyAscii = Asc("L"))
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("1") And KeyAscii <= Asc("8")) And Not (KeyAscii = Asc("L")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtA0w04_Validate(Cancel As Boolean)
   If txtA0w04 = "" And txtA0w02 <> "" Then
      MsgBox "公司別不可空白！", vbExclamation
      Cancel = True
   End If
End Sub

Public Function FormSave() As Boolean

On Error GoTo ErrHnd

   strSql = "Update acc0w0 set a0w04='" & ChgSQL(txtA0w04) & "'" & _
      ",a0w06='" & ChgSQL(txtA0w06) & "',a0w12='" & strUserNum & "'" & _
      ",a0w10=to_char(sysdate,'yyyymmdd')-19110000,a0w11=to_char(sysdate,'HH24MiSS')" & _
      " where a0w02='" & txtPreChar & txtA0w02 & "'"
   
   adoTaie.BeginTrans
   adoTaie.Execute strSql
   adoTaie.CommitTrans
   FormSave = True

ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub txtA0w06_GotFocus()
   If txtA0w06.Locked = False Then
      TextInverse txtA0w06
   End If
End Sub
