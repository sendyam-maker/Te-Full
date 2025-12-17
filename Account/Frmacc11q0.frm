VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11q0 
   AutoRedraw      =   -1  'True
   Caption         =   "發票作廢作業"
   ClientHeight    =   4356
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8928
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4356
   ScaleWidth      =   8928
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7440
      MaxLength       =   1
      TabIndex        =   3
      Top             =   300
      Width           =   345
   End
   Begin VB.TextBox txtA4301 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   300
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2820
      Picture         =   "Frmacc11q0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   300
      Width           =   350
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   270
      Top             =   3930
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   4290
      TabIndex        =   2
      Top             =   300
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "請依流程 發票作廢>發票上傳作業>重開發票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1290
      TabIndex        =   26
      Top             =   3840
      Width           =   6000
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "本發票已收款，重開日期如欲更動，請至發票維護更改日期！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Index           =   1
      Left            =   4575
      TabIndex        =   25
      Top             =   3210
      Width           =   4005
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   615
      Left            =   0
      Top             =   150
      Width           =   8895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "發票是否收回：   (Y:收回)"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5670
      TabIndex        =   24
      Top             =   360
      Width           =   3270
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   23
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "發票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1290
      TabIndex        =   22
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "客戶編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1290
      TabIndex        =   21
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4590
      TabIndex        =   20
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1290
      TabIndex        =   19
      Top             =   1830
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "統一編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1290
      TabIndex        =   18
      Top             =   2370
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "銷 售 額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1290
      TabIndex        =   17
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "稅　　額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4590
      TabIndex        =   16
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "請款單號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4590
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "作廢日期"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   14
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label labA4302 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4302"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2490
      TabIndex        =   13
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label labA0K03 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA0K03"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2490
      TabIndex        =   12
      Top             =   1500
      Width           =   990
   End
   Begin MSForms.Label labA0K20 
      Height          =   285
      Left            =   5790
      TabIndex        =   11
      Top             =   1500
      Width           =   1920
      VariousPropertyBits=   19
      Caption         =   "labA0K20"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label labA0K04 
      Height          =   285
      Left            =   2490
      TabIndex        =   10
      Top             =   1830
      Width           =   6135
      VariousPropertyBits=   19
      Caption         =   "labA0K04"
      Size            =   "10821;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label labA4303 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4303"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2490
      TabIndex        =   9
      Top             =   2370
      Width           =   945
   End
   Begin VB.Label labA4304 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4304"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2490
      TabIndex        =   8
      Top             =   2700
      Width           =   945
   End
   Begin VB.Label labA4305 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4305"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5790
      TabIndex        =   7
      Top             =   2700
      Width           =   1875
   End
   Begin VB.Label labAxc02 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labAxc02"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5790
      TabIndex        =   6
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label labA4317 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4317"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2490
      TabIndex        =   5
      Top             =   3180
      Width           =   945
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "未收款沖帳傳票："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   570
      TabIndex        =   4
      Top             =   3180
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc11q0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/08 Form2.0已修改
'Create by Sindy 2014/1/8
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoacc0k0a As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim adoadodc1 As New ADODB.Recordset


Private Sub Command2_Click()
   Acc0k0Refresh
   If adoacc0k0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      RecordShow
   End If
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command2_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  重新整理國內收據資料
'
'*************************************************
Public Sub Acc0k0Refresh()
On Error GoTo Checking
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.MaxRecords = intMax
   adoacc0k0.Open "select * from acc430,acc431,acc0k0,staff where a4301>='" & txtA4301 & "' and a4301=axc01(+) and axc02=a0k01(+) and a0k20=st01(+) and nvl(a4308,0)>0 order by a4301 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0k0a.State = adStateOpen Then
      adoacc0k0a.Close
   End If
   adoacc0k0a.CursorLocation = adUseClient
   adoacc0k0a.MaxRecords = intMax
   adoacc0k0a.Open "select * from acc430,acc431,acc0k0,staff where a4301>='" & txtA4301 & "' and a4301=axc01(+) and axc02=a0k01(+) and a0k20=st01(+) and nvl(a4308,0)=0 order by a4301 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0k0.RecordCount <> 0 Then
      If txtA4301 <> MsgText(601) Then
         adoacc0k0.Find "a0k01 = '" & txtA4301 & "'", 0, adSearchForward, 1
         If adoacc0k0.EOF = False Then
            FormShow
            AdodcRefresh
            RecordShow
         Else
            adoacc0k0.MoveFirst
         End If
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   txtA4301 = strItemNo
   Acc0k0Refresh
   If adoacc0k0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      RecordShow
   End If
   strItemNo = MsgText(601)
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
   Me.Height = 5700
   Me.Width = 9048
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2 + 900
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   Forms(0).Toolbar1.Buttons.Item(5).Enabled = False
   Call AdodcClear
   
   strItemNo = MsgText(601)
   OpenTable
   If adoacc0k0.RecordCount <> 0 Then
      adoacc0k0.MoveLast
      adoacc0k0.MoveFirst
      RecordShow
   End If
   'Call FormDisabled
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
   Set Frmacc11q0 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Then
      If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
         MsgBox Label2(1) & MsgText(52), , MsgText(5)
         Cancel = True
         MaskEdBox1.SetFocus
         Exit Sub
      End If
      If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
         MsgBox Label2(1) & MsgText(63), , MsgText(5)
         Cancel = True
         MaskEdBox1.SetFocus
         Exit Sub
      End If
   End If
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   '已作廢發票
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.MaxRecords = intMax
   adoacc0k0.Open "select * from acc430,acc431,acc0k0,staff where a4301>='" & txtA4301 & "' and a4301=axc01(+) and axc02=a0k01(+) and a0k20=st01(+) and nvl(a4308,0)<>0 order by a4301 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '未作廢發票
   adoacc0k0a.CursorLocation = adUseClient
   adoacc0k0a.MaxRecords = intMax
   adoacc0k0a.Open "select * from acc430,acc431,acc0k0,staff where a4301>='" & txtA4301 & "' and a4301=axc01(+) and axc02=a0k01(+) and a0k20=st01(+) and nvl(a4308,0)=0 order by a4301 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc430,acc431,acc0k0,staff where a4301='" & txtA4301 & "' and a4301=axc01(+) and axc02=a0k01(+) and a0k20=st01(+) and nvl(a4308,0)>0 order by a4301 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   txtA4301 = adoacc0k0.Fields("a4301").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0k0.Fields("a4308").Value) Or adoacc0k0.Fields("a4308").Value = 0 Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0k0.Fields("a4308").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0k0.Fields("a4302").Value) Then
      labA4302 = MsgText(601)
   Else
      labA4302 = ChangeTStringToTDateString(adoacc0k0.Fields("a4302").Value)
   End If
   labAXC02 = adoacc0k0.Fields("Axc02").Value
   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
      labA0K03 = MsgText(601)
   Else
      labA0K03 = adoacc0k0.Fields("A0K03").Value
   End If
   If IsNull(adoacc0k0.Fields("st02").Value) Then
      labA0K20 = MsgText(601)
   Else
      labA0K20 = adoacc0k0.Fields("st02").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k04").Value) Then
      labA0K04 = MsgText(601)
   Else
      labA0K04 = adoacc0k0.Fields("A0K04").Value
   End If
   If IsNull(adoacc0k0.Fields("a4303").Value) Then
      labA4303 = MsgText(601)
   Else
      labA4303 = adoacc0k0.Fields("a4303").Value
   End If
   If IsNull(adoacc0k0.Fields("a4304").Value) Then
      labA4304 = MsgText(601)
   Else
      labA4304 = adoacc0k0.Fields("a4304").Value
   End If
   If IsNull(adoacc0k0.Fields("a4305").Value) Then
      labA4305 = MsgText(601)
   Else
      labA4305 = adoacc0k0.Fields("a4305").Value
   End If
   If IsNull(adoacc0k0.Fields("a4317").Value) Then
      labA4317 = MsgText(601)
      Label9(3).Visible = False
      labA4317.Visible = False
   Else
      labA4317 = adoacc0k0.Fields("a4317").Value
      Label9(3).Visible = True
      labA4317.Visible = True
   End If
End Sub

'*************************************************
'  顯示查詢資料(國內收據資料)
'
'*************************************************
Private Sub Acc0k0Query()
   If IsNull(adoacc0k0a.Fields("a4302").Value) Then
      labA4302 = MsgText(601)
   Else
      labA4302 = ChangeTStringToTDateString(adoacc0k0a.Fields("a4302").Value)
   End If
   labAXC02 = adoacc0k0a.Fields("Axc02").Value
   If IsNull(adoacc0k0a.Fields("a0k03").Value) Then
      labA0K03 = MsgText(601)
   Else
      labA0K03 = adoacc0k0a.Fields("A0K03").Value
   End If
   If IsNull(adoacc0k0a.Fields("st02").Value) Then
      labA0K20 = MsgText(601)
   Else
      labA0K20 = adoacc0k0a.Fields("st02").Value
   End If
   If IsNull(adoacc0k0a.Fields("a0k04").Value) Then
      labA0K04 = MsgText(601)
   Else
      labA0K04 = adoacc0k0a.Fields("A0K04").Value
   End If
   If IsNull(adoacc0k0a.Fields("a4303").Value) Then
      labA4303 = MsgText(601)
   Else
      labA4303 = adoacc0k0a.Fields("a4303").Value
   End If
   If IsNull(adoacc0k0a.Fields("a4304").Value) Then
      labA4304 = MsgText(601)
   Else
      labA4304 = adoacc0k0a.Fields("a4304").Value
   End If
   If IsNull(adoacc0k0a.Fields("a4305").Value) Then
      labA4305 = MsgText(601)
   Else
      labA4305 = adoacc0k0a.Fields("a4305").Value
   End If
   If IsNull(adoacc0k0a.Fields("a4317").Value) Then
      labA4317 = MsgText(601)
      Label9(3).Visible = False
      labA4317.Visible = False
   Else
      labA4317 = adoacc0k0a.Fields("a4317").Value
      Label9(3).Visible = True
      labA4317.Visible = True
   End If
   'add by sonia 2017/9/29
   If Val("" & adoacc0k0a.Fields("a0k17").Value) + Val("" & adoacc0k0a.Fields("a0k18").Value) = 0 Then
      Label9(1).Visible = False
   Else
      Label9(1).Visible = True
   End If
   'end 2017/9/29
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   txtA4301.Enabled = False
   MaskEdBox1.Enabled = False
   Text4.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   If strSaveConfirm = MsgText(3) Then '新增狀態時
      txtA4301.Enabled = True
   Else
      If txtA4301 <> "" Then
         txtA4301.Enabled = False
      Else
         txtA4301.Enabled = True
      End If
   End If
   MaskEdBox1.Enabled = True
   Text4.Enabled = True
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc430,acc431,acc0k0,staff where a4301='" & txtA4301 & "' and a4301=axc01(+) and axc02=a0k01(+) and a0k20=st01(+) and nvl(a4308,0)>0 order by a4301 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0k0.Bookmark & MsgText(35) & adoacc0k0.RecordCount
End Sub

Public Sub Frmacc11q0_Clear()
   With Frmacc11q0
      .txtA4301 = ""
      TextInverse .txtA4301
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
      .MaskEdBox1.Mask = DFormat
      .Text4 = ""
      .labA4302 = ""
      .labAXC02 = ""
      .labA0K03 = ""
      .labA0K20 = ""
      .labA0K04 = ""
      .labA4303 = ""
      .labA4304 = ""
      .labA4305 = ""
      .labA4317 = ""
      .Label9(3).Visible = False
      .Label9(1).Visible = False  'add by sonia 2017/9/29
      .labA4317.Visible = False
      .AdodcRefresh
      .txtA4301.SetFocus
   End With
End Sub

Public Sub Frmacc11q0_First()
   With Frmacc11q0
      If .adoacc0k0.RecordCount <> 0 Then
         .adoacc0k0.MoveFirst
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11q0_Last()
   With Frmacc11q0
      If .adoacc0k0.RecordCount <> 0 Then
         .adoacc0k0.MoveLast
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11q0_Next()
   With Frmacc11q0
      If .adoacc0k0.EOF = False Then
         .adoacc0k0.MoveNext
         If .adoacc0k0.EOF Then
            .adoacc0k0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11q0_Previous()
   With Frmacc11q0
      If .adoacc0k0.BOF = False Then
         .adoacc0k0.MovePrevious
         If .adoacc0k0.BOF Then
            .adoacc0k0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .AdodcRefresh
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11q0_Save()
Dim rec_Msg  As String   'add by sonia 2016/11/29

   On Error GoTo Checking
   With Frmacc11q0
      If .txtA4301 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .txtA4301.SetFocus
         Exit Sub
      Else
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label2(1) & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
                MsgBox .Label2(1) & MsgText(63), , MsgText(5)
                strControlButton = MsgText(602)
                .MaskEdBox1.SetFocus
                Exit Sub
            End If
         End If
         If .txtA4301 <> MsgText(601) Then
            'Add by Amy 2023/06/13 判斷發票未曾上傳至盟立,則不可作廢(要先將新增的發票先上傳,才能作廢,否則一起上傳同一張發票和作廢,盟立會回傳錯誤)
            If adoquery.State = adStateOpen Then adoquery.Close
            adoquery.CursorLocation = adUseClient
            adoquery.Open "Select * from acc430 where a4301 = '" & .txtA4301 & "' and a4319 is null ", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               MsgBox "此張發票開立後未曾上傳盟立,不可作廢！"
               strControlButton = MsgText(602)
               txtA4301.SetFocus
               adoquery.Close
               Exit Sub
            End If
            .adoquery.Close
            'end 2023/06/13
            
            '已銷帳不可作廢
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            .adoquery.Open "select * from acc430 where a4301 = '" & .txtA4301 & "' and a4309 is not null ", adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount <> 0 Then
               MsgBox "已銷帳不可作廢...", , MsgText(5)
               strControlButton = MsgText(602)
               .txtA4301.SetFocus
               .adoquery.Close
               Exit Sub
            End If
            .adoquery.Close
         End If
         If .labA4317.Visible = True And Trim(.labA4317.Caption) <> "" Then
            '已有未收款沖帳傳票
            strSql = "select * From acc021 where ax201='J' and ax202='" & Trim(.labA4317.Caption) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               '傳票已過帳不可作廢
               If Val("" & RsTemp.Fields("ax210")) > 0 Then
                  MsgBox "未收款沖帳傳票已過帳不可作廢...", , MsgText(5)
                  strControlButton = MsgText(602)
                  .txtA4301.SetFocus
                  Exit Sub
               End If
            End If
            '傳票未過帳
            MsgBox "已產生未收款沖帳傳票，請自行調整傳票內容!!", , MsgText(5)
         End If
         'Add by Amy 2021/11/09 因SX90796297 1100923 作廢日1080703
        If Val(FCDate(.labA4302)) > Val(FCDate(.MaskEdBox1.Text)) Then
            MsgBox .Label2(1) & "不可小於發票日", , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
        End If
         '發票已申報,不可作廢
         If Val(GetInvDataA4111(Val(FCDate(.labA4302)))) > 0 Then
            MsgBox "發票已申報, 不可作廢...", , MsgText(5)
            strControlButton = MsgText(602)
            .txtA4301.SetFocus
            Exit Sub
         End If
         '發票是否收回,一定要輸入Y才可存檔
         If .Text4 <> "Y" Then
            MsgBox "發票一定要收回, 才可作廢...", , MsgText(5)
            strControlButton = MsgText(602)
            .Text4.SetFocus
            Exit Sub
         End If
         'add by sonia 2016/11/29
         rec_Msg = ""
         If .txtA4301 <> MsgText(601) Then
            '已收款提醒改傳票摘要 發票號碼EF26342333作廢重開
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            .adoquery.Open "select distinct a1p01,a1p22 from acc431,acc1p0 where axc01='" & .txtA4301 & "' and axc03 is not null and axc03=a1p04(+) and a1p02='A' ", adoTaie, adOpenStatic, adLockReadOnly
            Do While .adoquery.EOF = False
               If IsNull(.adoquery.Fields(0).Value) = False Then
                  rec_Msg = rec_Msg & "(" & .adoquery.Fields(0).Value & ")" & .adoquery.Fields(1).Value & "；"
               End If
               .adoquery.MoveNext
            Loop
            If rec_Msg <> "" Then
               MsgBox "此發票已收款, 若為修改發票抬頭, 請記得修改收款傳票之摘要：" & rec_Msg, , MsgText(21)
            End If
            .adoquery.Close
         End If
         'end 2016/11/29
      End If
      
      adoTaie.Execute "update acc430 set a4308=" & Val(FCDate(.MaskEdBox1.Text)) & " where a4301 = '" & .txtA4301 & "'"
      adoTaie.Execute "update acc431 set axc02='E' where axc01 = '" & .txtA4301 & "'"
      .Acc0k0Refresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtA4301_GotFocus()
   TextInverse txtA4301
End Sub

Private Sub txtA4301_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA4301_Validate(Cancel As Boolean)
Dim strKey As String, StrKey2 As String
   
   If strSaveConfirm = MsgText(3) Then
      adoacc0k0a.Close
      adoacc0k0a.CursorLocation = adUseClient
      adoacc0k0a.Open "select * from acc430,acc431,acc0k0,staff where a4301='" & txtA4301 & "' and a4301=axc01(+) and axc02=a0k01(+) and a0k20=st01(+) order by a4301 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      strKey = txtA4301
      StrKey2 = MaskEdBox1
      If adoacc0k0a.RecordCount <> 0 Then
         If Val("" & adoacc0k0a.Fields("a4308")) <> 0 Then
            MsgBox "此發票號碼已作廢", , MsgText(5)
            AdodcClear
            txtA4301 = strKey
            MaskEdBox1 = StrKey2
            AdodcRefresh
            Cancel = True
            Exit Sub
         End If
         Acc0k0Query
         AdodcRefresh
      Else
         MsgBox MsgText(28), , MsgText(5)
         AdodcClear
         txtA4301 = strKey
         MaskEdBox1 = StrKey2
         AdodcRefresh
         Cancel = True
      End If
   End If
End Sub

'*************************************************
'  清除查詢資料
'
'*************************************************
Private Sub AdodcClear()
   txtA4301 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Text4 = ""
   labA4302 = ""
   labAXC02 = ""
   labA0K03 = ""
   labA0K20 = ""
   labA0K04 = ""
   labA4303 = ""
   labA4304 = ""
   labA4305 = ""
   labA4317 = ""
   Label9(3).Visible = False
   Label9(1).Visible = False  'add by sonia 2017/9/29
   labA4317.Visible = False
End Sub
