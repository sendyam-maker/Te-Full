VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc7100 
   AutoRedraw      =   -1  'True
   Caption         =   "分所收款作業"
   ClientHeight    =   4725
   ClientLeft      =   1050
   ClientTop       =   4095
   ClientWidth     =   8595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   8595
   Begin VB.TextBox Text11 
      Height          =   315
      Left            =   1410
      MaxLength       =   2000
      TabIndex        =   17
      Top             =   4010
      Width           =   6915
   End
   Begin VB.TextBox Text10 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   280
      Left            =   2250
      TabIndex        =   34
      Top             =   790
      Width           =   5985
   End
   Begin VB.TextBox Text2 
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
      Height          =   324
      Left            =   5550
      MaxLength       =   15
      TabIndex        =   2
      Top             =   470
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   7200
      Picture         =   "Frmacc7100.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   480
      Width           =   350
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1410
      TabIndex        =   16
      Top             =   3680
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4650
      TabIndex        =   15
      Top             =   3350
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   13
      Top             =   3020
      Width           =   6855
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4650
      MaxLength       =   8
      TabIndex        =   12
      Top             =   2690
      Width           =   1575
   End
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
      Height          =   324
      Left            =   1410
      MaxLength       =   12
      TabIndex        =   11
      Top             =   2690
      Width           =   1965
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1410
      TabIndex        =   8
      Top             =   2030
      Width           =   1575
   End
   Begin VB.TextBox Text16 
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
      Height          =   324
      Left            =   1410
      TabIndex        =   7
      Top             =   1700
      Width           =   1155
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
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
      Height          =   280
      Left            =   1410
      TabIndex        =   6
      Top             =   1410
      Width           =   6825
   End
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
      Height          =   324
      Left            =   1410
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1070
      Width           =   6825
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
      Height          =   324
      Left            =   1410
      TabIndex        =   4
      Top             =   750
      Width           =   795
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1410
      TabIndex        =   9
      Top             =   2360
      Width           =   1575
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
      Height          =   324
      Left            =   5550
      MaxLength       =   15
      TabIndex        =   1
      Top             =   130
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   130
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
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
      Height          =   324
      Left            =   4650
      TabIndex        =   10
      Top             =   2360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
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
      Left            =   1410
      TabIndex        =   14
      Top             =   3350
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "PS：類別：A暫收款、B代收款、C其他收款 　所別：C中所、N南所、K高所"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   240
      TabIndex        =   38
      Top             =   4350
      Width           =   7812
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "PS：費用及點數為未收款且未銷帳之金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3720
      TabIndex        =   37
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "所別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   360
      TabIndex        =   36
      Top             =   520
      Width           =   4185
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "備　　註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   360
      TabIndex        =   35
      Top             =   4060
      Width           =   972
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "留分所金額"
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
      Left            =   200
      TabIndex        =   33
      Top             =   3730
      Width           =   1275
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "扣繳金額"
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
      Left            =   3720
      TabIndex        =   32
      Top             =   3400
      Width           =   1092
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "扣繳日期"
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
      TabIndex        =   31
      Top             =   3400
      Width           =   1092
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "付 款 地"
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
      TabIndex        =   30
      Top             =   3070
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   220
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "點　　數"
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
      TabIndex        =   28
      Top             =   1750
      Width           =   972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "到 期 日"
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
      Left            =   3720
      TabIndex        =   27
      Top             =   2410
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "票　　號"
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
      Left            =   3720
      TabIndex        =   26
      Top             =   2740
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "支　　票"
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
      TabIndex        =   25
      Top             =   2410
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   4632
      Left            =   120
      Top             =   12
      Width           =   8292
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "帳　　號"
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
      Top             =   2740
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "現　　金"
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
      TabIndex        =   23
      Top             =   2080
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "人工收據"
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
      Left            =   4560
      TabIndex        =   22
      Top             =   520
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "案件性質"
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
      TabIndex        =   21
      Top             =   1460
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      TabIndex        =   20
      Top             =   1120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收 款 人"
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
      Top             =   800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "電腦收據"
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
      Left            =   4560
      TabIndex        =   18
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "Frmacc7100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Public adoacc310 As New ADODB.Recordset
'add by nick 2004/08/20
'ostate = 1 insert;2 edit;3 delete;4 find
Public oState As String
Public DBOffice As String
'93.12.16 ADD BY SONIA 是否已收或已銷
Public M_REC As String
'add by nick 2004/12/21
Dim IsSeek As Boolean

Private Sub Command1_Click()
'edit by nick 2005/01/05 多筆查詢限當所，單筆查詢不限
'    Acc310Refresh
'    'edit by nick 2004/12/21
'    'If adoacc310.RecordCount <> 0 Then
'    If adoacc310.RecordCount <> 0 And IsSeek = True Then
'        FormShow
'        RecordShow
'    End If
Dim rsSearch As New ADODB.Recordset
Set rsSearch = New ADODB.Recordset
strSql = "select * from acc310 where A3103='" & Me.Text1.Text & "'And A3104 ='" & Me.Text2.Text & "' "
With rsSearch
    .CursorLocation = adUseClient
    .Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        DBOffice = "" & .Fields("A3101").Value
        Label18.Caption = "所別：" & DBOffice & " (1.北所 2.中所 3.南所 4.高所 5.其他)"
        Me.MaskEdBox1.Mask = ""
        If IsNull(.Fields("A3102").Value) Then
           Me.MaskEdBox1.Text = ""
        Else
           Me.MaskEdBox1.Text = CFDate(.Fields("A3102").Value)
        End If
        Me.MaskEdBox1.Mask = DFormat
        Me.Text1.Text = "" & .Fields("A3103").Value
        Me.Text2.Text = "" & .Fields("A3104").Value
        Me.Text3.Text = Val("" & .Fields("A3105").Value)
        Me.Text4.Text = Val("" & .Fields("A3106").Value)
        Me.MaskEdBox2.Mask = ""
        If IsNull(.Fields("A3107").Value) Then
           Me.MaskEdBox2.Text = ""
        Else
           Me.MaskEdBox2.Text = CFDate(.Fields("A3107").Value)
        End If
        Me.MaskEdBox2.Mask = DFormat
        Me.Text5.Text = "" & .Fields("A3108").Value
        Me.Text6.Text = "" & .Fields("A3109").Value
        Me.Text7.Text = "" & .Fields("A3110").Value
        Me.MaskEdBox3.Mask = ""
        If IsNull(.Fields("A3111").Value) Then
           Me.MaskEdBox3.Text = ""
        Else
           Me.MaskEdBox3.Text = CFDate(.Fields("A3111").Value)
        End If
        Me.MaskEdBox3.Mask = DFormat
        Me.Text8.Text = Val("" & .Fields("A3112").Value)
        Me.Text9.Text = Val("" & .Fields("A3113").Value)
        Me.Text13.Text = "" & .Fields("A3121").Value
        Me.Text14.Text = "" & .Fields("A3122").Value
        Me.Text16.Text = "" & .Fields("A3123").Value
        Me.Text11.Text = "" & .Fields("A3124").Value
        Text13_Validate False
    Else
         MsgBox "搜尋不到資料！", , "錯誤！"
    End If
    .Close
End With
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        Command1_Click
        Exit Sub
    End Select
    'KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
Dim arrItemNo
    
    strFormName = Name
    If strItemNo = MsgText(601) Then
        Exit Sub
    End If
    If adoacc310.RecordCount <> 0 Then
        adoacc310.MoveFirst
    End If
    arrItemNo = Split(strItemNo, ",")
    Text1 = arrItemNo(0)
    Text2 = arrItemNo(1)
    Acc310Refresh
    'edit by nick 2004/12/21
    'If adoacc310.RecordCount <> 0 Then
    If adoacc310.RecordCount <> 0 And IsSeek = True Then
        FormShow
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
   
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Width = 8688
    Me.Height = 5112
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath1)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    Me.MaskEdBox1.Mask = DFormat
    Me.MaskEdBox2.Mask = DFormat
    Me.MaskEdBox3.Mask = DFormat
    OpenTable
    If adoacc310.RecordCount <> 0 Then
        adoacc310.MoveLast
        adoacc310.MoveFirst
        RecordShow
    End If
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
    Set Frmacc7100 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If oState = "3" Or oState = "4" Then Exit Sub
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
        If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
           MsgBox Label9 & MsgText(63), , MsgText(5)
           Cancel = True
           MaskEdBox1.SetFocus
           Exit Sub
        End If
    End If
    If oState = "1" Or oState = "2" Then
        If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
           MsgBox Label9 & MsgText(52), , MsgText(5)
           Cancel = True
           MaskEdBox1.SetFocus
           Exit Sub
        End If
    End If
    
End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   'EDIT BY NICK 2004/11/09
   If Trim(Text4.Text) = "" Then Exit Sub
   If MaskEdBox2.Text <> MsgText(601) Or MaskEdBox2.Text <> MsgText(29) Then
        If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
           MsgBox Label7 & MsgText(63), , MsgText(5)
           Cancel = True
           MaskEdBox2.SetFocus
           Exit Sub
        End If
    End If
End Sub

Private Sub MaskEdBox3_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
        If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
           MsgBox Label11 & MsgText(63), , MsgText(5)
           Cancel = True
           MaskEdBox3.SetFocus
           Exit Sub
        End If
    End If
End Sub

Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'add by nick 2004/12/08
On Error GoTo ErrSub
    KeyAscii = UpperCase(KeyAscii)
    If oState = "1" Or oState = "2" Then
        If Len(Text1) = 0 Or Text1.SelLength <> 0 Then
            If KeyAscii <> 67 And KeyAscii <> 78 And KeyAscii <> 75 And KeyAscii <> 69 And KeyAscii <> 8 Then
                KeyAscii = 0
            Else
                If (pub_strUserOffice = "2" And KeyAscii = 67) Or (pub_strUserOffice = "3" And KeyAscii = 78) Or (pub_strUserOffice = "4" And KeyAscii = 75) Or KeyAscii = 8 Or KeyAscii = 69 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (pub_strUserOffice = "1" And (KeyAscii = 67 Or KeyAscii = 78 Or KeyAscii = 75)) Then
                
                Else
                    KeyAscii = 0
                End If
            End If
        ElseIf Len(Text1) = 1 And Text1.SelLength = 0 And Mid(Text1, 1, 1) <> "E" Then
            If KeyAscii >= 65 And KeyAscii <= 69 Then
                '取自動編號
                Dim AAA As String
                adoTaie.BeginTrans
                Text1.Text = Text1.Text & AccAutoNo(Text1.Text & Chr(KeyAscii), 3, Format(Mid(ServerDate, 1, 4) - 1911, "000"), Mid(ServerDate, 5, 2))
                AAA = AccSaveAutoNo(Mid(Text1.Text, 1, 2), Mid(Text1.Text, 8, 3), Format(Mid(ServerDate, 1, 4) - 1911, "000"), Mid(ServerDate, 5, 2))
                adoTaie.CommitTrans
                KeyAscii = 0
                Text1.Enabled = False
                Text2.Enabled = False
                '93.12.20 ADD BY SONIA
                Select Case Mid(Text1.Text, 1, 1)
                  Case "C"
                     DBOffice = "2"
                  Case "N"
                     DBOffice = "3"
                  Case "K"
                     DBOffice = "4"
                  Case Else
                     DBOffice = "1"
                  End Select
                Label18.Caption = "所別：" & DBOffice & " (1.北所 2.中所 3.南所 4.高所 5.其他)"
                '93.12.20 END
            ElseIf KeyAscii = 8 Then
            Else
                KeyAscii = 0
            End If
        ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
        ElseIf KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    End If
Exit Sub
ErrSub:
    adoTaie.RollbackTrans
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim ii As Integer
Dim strAmt As String
Dim strOurCaseNo As String
    
'edit by nick 2004/08/20 電腦收據第一碼不為E 時不檢查 ACC0K0
'If Me.Text1.Text = "" Or Me.Text1.Text = "E" Then Exit Sub
'edit by nick 2004/08/26
'If Me.Text1.Text = "" Or Me.Text1.Text = "E" Or (Len(Text1.Text) > 1 And UCase(Mid(Text1.Text, 1, 1)) <> "E") Then Exit Sub
If Me.Text1.Text = "" Or Me.Text1.Text = "E" Or (Len(Text1.Text) > 1 And UCase(Mid(Text1.Text, 1, 1)) <> "E") Then Text2.Text = Text1.Text: Text15.Text = "": Exit Sub
Screen.MousePointer = vbHourglass
'add by nick 2004/10/08 判斷所別
StrSQLa = "Select st06 From ACC0K0, Staff Where A0K20=ST01 And A0K01='" & ChgSQL(Me.Text1.Text) & "'   "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    DBOffice = "" & rsA.Fields(0).Value
    Label18.Caption = "所別：" & DBOffice & " (1.北所 2.中所 3.南所 4.高所 5.其他)"
Else
    DBOffice = ""
    Label18.Caption = "所別：" & DBOffice & " (1.北所 2.中所 3.南所 4.高所 5.其他)"
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'add by nick 2004/10/11
If oState = "1" Or oState = "2" Then
        If DBOffice <> pub_strUserOffice And UCase(strUserDept) <> "M51" Then
             MsgBox "不能修改它所資料", , MsgText(5)
             Screen.MousePointer = vbDefault
             Cancel = True
             Exit Sub
        End If
End If
'edit by nick 2004/08/19 及改分所也可以用
'strSQLA = "Select A0K20||' '||ST02, A0K04, A0J20, Round(Nvl(A0J09,0)/1000,1), A0J02, Nvl(A0J09,0) + Nvl(A0J10,0) From ACC0K0, ACC0J0, Staff Where A0K01=A0J13 And A0K20=ST01 And A0K01='" & ChgSQL(Me.Text1.Text) & "' And ST06='" & pub_strUserOffice & "' Order By A0J03 "
'Modified by Morgan 2011/12/26 取消a0j03 改抓 cp10
'Modified by Morgan 2011/12/27 取消 a0j20
StrSQLa = "Select A0K20, A0K04, getcp10desc(cp01,cp10,a0j04) cp10N, Round(Nvl(A0J09,0)/1000,1), A0J02, Nvl(A0J09,0) + Nvl(A0J10,0),st02 From ACC0K0, ACC0J0, Staff,caseprogress Where A0K01=A0J13(+) And A0K20=ST01(+) And A0K01='" & ChgSQL(Me.Text1.Text) & "' and cp09(+)=a0j01  Order By cp10 "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    ii = 0: strAmt = "": strOurCaseNo = ""
    While Not rsA.EOF
        strAmt = Val(strAmt) + Val("" & rsA.Fields(5).Value)
        ii = ii + 1
        If ii = 1 Then
            Me.Text13.Text = "" & rsA.Fields(0).Value
            'add by nick 2004/08/19
            Me.Text10.Text = "" & rsA.Fields(6).Value
            Me.Text14.Text = "" & rsA.Fields(1).Value
            strOurCaseNo = ReConBNOurCaseNO("" & rsA.Fields(4).Value)
            Me.Text15.Text = strOurCaseNo & rsA.Fields(2).Value
            Me.Text16.Text = "" & rsA.Fields(3).Value
        Else
            If strOurCaseNo = ReConBNOurCaseNO("" & rsA.Fields(4).Value) Then
                strOurCaseNo = ""
            Else
                strOurCaseNo = ReConBNOurCaseNO("" & rsA.Fields(4).Value)
            End If
            Me.Text15.Text = Me.Text15.Text & "及" & strOurCaseNo & rsA.Fields(2).Value
            Me.Text16.Text = Val(Me.Text16.Text) + Val("" & rsA.Fields(3).Value)
        End If
        rsA.MoveNext
    Wend
   '93.10.18 ADD BY SONIA 扣除已銷帳及已收款
   If rsB.State <> adStateClosed Then rsA.Close
   Set rsB = Nothing
   StrSqlB = "Select Round(Nvl(SUM(A1U04-A1U08+A1U07),0)/1000,1),SUM(A1U04+A1U05-A1U08-A1U10+A1U07+A1U09) From ACC1U0 Where A1U02='" & ChgSQL(Me.Text1.Text) & "'"
   rsB.CursorLocation = adUseClient
   rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      strAmt = Val(strAmt) - Val("" & rsB.Fields(1).Value)
      Me.Text16.Text = Val(Me.Text16.Text) - Val("" & rsB.Fields(0).Value)
   End If
   '93.10.18 END
    Me.Text15.Text = Me.Text15.Text & strAmt
   '93.10.18 ADD BY SONIA
   M_REC = ""  '93.12.16 ADD BY SONIA
   If Me.Text15.Text <> "" And Me.Text16.Text = 0 And strAmt = 0 Then
      If oState = "1" Or oState = "2" Then
         MsgBox "此筆電腦收據資料已收款或已銷帳!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
      M_REC = "Y" '93.12.16 ADD BY SONIA
   End If
   '93.10.18 END
Else
    MsgBox "查無此電腦收據資料!!!", vbExclamation + vbOKOnly
    Cancel = True
    'add by nick 2004/08/19
    Me.Text10.Text = ""
    Me.Text13.Text = ""
    Me.Text14.Text = ""
    Me.Text15.Text = ""
    Me.Text16.Text = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
Screen.MousePointer = vbDefault
If Cancel = True Then
    Text1_GotFocus
Else
    If Me.Text2.Text = "" Then Me.Text2.Text = Me.Text1.Text
End If
End Sub

'add by nick 2004/08/20
Private Sub Text11_GotFocus()
TextInverse Text11
'edit by nickc 2007/07/11 切換輸入法改用API
'Text11.IMEMode = 1
OpenIme
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(Text11, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "備註內容太長"
      Text11_GotFocus
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: Text11.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'add by nick 2004/08/20
Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'add by nick 2004/08/20
Private Sub Text13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   Text10 = Empty
   If IsEmptyText(Text13) = False Then
      Text10 = GetStaffBy7100(Text13)
      If IsEmptyText(Text10) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "收款人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text13_GotFocus
      End If
   End If
End Sub

'add by nick 2004/08/20
Private Sub Text14_GotFocus()
TextInverse Text14
End Sub

'add by nick 2004/08/20
Private Sub Text16_GotFocus()
TextInverse Text16
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
End Sub

'add by nick 2004/08/19
Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

On Error GoTo Checking
    strSql = "Select * From ACC310 Where A3101='" & pub_strUserOffice & "' Order By A3102, A3103, A3104 "
    adoacc310.CursorLocation = adUseClient
    adoacc310.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
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
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
On Error GoTo ErrorHandler
    'add by nick 2004/08/20 紀錄所別
    DBOffice = "" & adoacc310.Fields("A3101").Value
    Label18.Caption = "所別：" & DBOffice & " (1.北所 2.中所 3.南所 4.高所 5.其他)"
    Me.MaskEdBox1.Mask = ""
    If IsNull(adoacc310.Fields("A3102").Value) Then
       Me.MaskEdBox1.Text = ""
    Else
       Me.MaskEdBox1.Text = CFDate(adoacc310.Fields("A3102").Value)
    End If
    Me.MaskEdBox1.Mask = DFormat
    Me.Text1.Text = "" & adoacc310.Fields("A3103").Value
    Me.Text2.Text = "" & adoacc310.Fields("A3104").Value
    Me.Text3.Text = Val("" & adoacc310.Fields("A3105").Value)
    Me.Text4.Text = Val("" & adoacc310.Fields("A3106").Value)
    Me.MaskEdBox2.Mask = ""
    If IsNull(adoacc310.Fields("A3107").Value) Then
       Me.MaskEdBox2.Text = ""
    Else
       Me.MaskEdBox2.Text = CFDate(adoacc310.Fields("A3107").Value)
    End If
    Me.MaskEdBox2.Mask = DFormat
    Me.Text5.Text = "" & adoacc310.Fields("A3108").Value
    Me.Text6.Text = "" & adoacc310.Fields("A3109").Value
    Me.Text7.Text = "" & adoacc310.Fields("A3110").Value
    Me.MaskEdBox3.Mask = ""
    If IsNull(adoacc310.Fields("A3111").Value) Then
       Me.MaskEdBox3.Text = ""
    Else
       Me.MaskEdBox3.Text = CFDate(adoacc310.Fields("A3111").Value)
    End If
    Me.MaskEdBox3.Mask = DFormat
    Me.Text8.Text = Val("" & adoacc310.Fields("A3112").Value)
    Me.Text9.Text = Val("" & adoacc310.Fields("A3113").Value)
    'add by nick 2004/08/19
    Me.Text13.Text = "" & adoacc310.Fields("A3121").Value
    Me.Text14.Text = "" & adoacc310.Fields("A3122").Value
    Me.Text16.Text = "" & adoacc310.Fields("A3123").Value
    Me.Text11.Text = "" & adoacc310.Fields("A3124").Value
'    Text1_Validate False
    Text13_Validate False
    'Text1_Validate False

Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
    
On Error GoTo ErrorHandler
    If adoacc310.RecordCount = 0 Then
        Exit Sub
    End If
    CountShow adoacc310.Bookmark, adoacc310.RecordCount
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

'*************************************************
'  重新整理分所收款資料
'
'*************************************************
Public Sub Acc310Refresh()
On Error GoTo Checking
    If adoacc310.State = adStateOpen Then
        adoacc310.Close
    End If
    strSql = "Select * From ACC310 Where A3101='" & pub_strUserOffice & "' Order By A3102,A3103, A3104 "
    adoacc310.CursorLocation = adUseClient
    adoacc310.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
    'add by nick 2004/12/21
    adoacc310.MoveFirst
    IsSeek = False
    Do While True
        If adoacc310("A3103").Value = Me.Text1.Text And adoacc310("A3104").Value = Me.Text2.Text Then IsSeek = True: Exit Do
        'add by nick 2004/12/21
        If Not adoacc310.EOF = False Then
            Exit Do
        End If
        adoacc310.MoveNext
    Loop
    'add by nick 2004/12/21
    If IsSeek = False Then
        MsgBox "搜尋不到資料！", , "錯誤！"
    End If
Checking:
    If Err.Number = 0 Or Err.Number = 3021 Then
        Exit Sub
    End If
    MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text2_LostFocus()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    'edit by nick 2004/08/19
    Text2.Text = UCase(Text2.Text)
    
    If Me.Text2.Text = "" Then Me.Text2.Text = Me.Text1.Text
    If strSaveConfirm = "A" And Me.Text1.Text <> "" And Me.Text2.Text <> "" Then
        StrSQLa = "Select * From ACC310 Where A3103='" & Me.Text1.Text & "' And A3104='" & Me.Text2.Text & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            MsgBox "資料重覆, 請重新輸入!!!", vbExclamation + vbOKOnly
            Me.Text1.SetFocus
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If

End Sub

Private Sub Text3_GotFocus()
    TextInverse Me.Text3
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub Text4_GotFocus()
    TextInverse Me.Text4
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub Text5_GotFocus()
    TextInverse Me.Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub Text6_GotFocus()
    TextInverse Me.Text8
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub Text7_GotFocus()
    TextInverse Me.Text7
    OpenIme
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(Text7, 20) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "付款地內容太長"
      Text7_GotFocus
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: Text7.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub Text8_GotFocus()
    TextInverse Me.Text8
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Sub Text9_GotFocus()
    TextInverse Me.Text9
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
    'KeyEnter KeyCode
End Sub

Private Function ReConBNOurCaseNO(strCaseNo As String) As String

If strCaseNo <> "" Then
    ReConBNOurCaseNO = Replace(Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 3), 6) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 2), 1) & "-" & Right(strCaseNo, 2), "-0-00", "")
Else
    ReConBNOurCaseNO = ""
End If

End Function

Function GetStaffBy7100(ByVal strStuff As String, Optional ByVal bAll As Boolean = False) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetStaffBy7100 = Empty
   
   strSql = "SELECT * FROM Staff " & _
            "WHERE ST01 = '" & strStuff & "' " & IIf(oState = "1", " and st04='1' ", "")
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("ST02")) = False Then
         GetStaffBy7100 = rsTmp.Fields("ST02")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'判斷中英混雜字串之長度是否有超過最大長度
Public Function CheckLengthIsOK(ByRef strTemp As String, ByRef intTemp As Integer) As Boolean
If GetTextLength(strTemp) > intTemp Then
   Beep
   MsgBox "輸入之資料過長，超過" & Format(intTemp) & "個字（註：中文算兩個字)!", vbCritical + vbOKOnly, "警告!!"
   CheckLengthIsOK = False
Else
   CheckLengthIsOK = True
End If
End Function

'取得中英混雜字串之長度
Public Function GetTextLength(ByRef strTemp As String) As Integer
GetTextLength = LenB(StrConv(strTemp, vbFromUnicode))
End Function
