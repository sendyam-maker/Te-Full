VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2191 
   AutoRedraw      =   -1  'True
   Caption         =   "其他結匯作業"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   8880
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
      Height          =   330
      Left            =   4560
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1280
      Width           =   1275
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
      Height          =   330
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   2
      Top             =   900
      Width           =   4000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "結匯對象"
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
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   520
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "代理人"
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
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Value           =   -1  'True
      Width           =   1035
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
      Height          =   330
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "帳單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   264
      TabIndex        =   9
      Top             =   4875
      Width           =   1092
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
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
      Height          =   330
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1280
      Width           =   1572
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   6570
      TabIndex        =   3
      Top             =   900
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Caption         =   "抵帳單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1536
      TabIndex        =   11
      Top             =   4875
      Width           =   1092
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FC暫收款退費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2832
      TabIndex        =   13
      Top             =   4875
      Width           =   1692
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8040
      Picture         =   "Frmacc2191.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   8
      ToolTipText     =   "取消"
      Top             =   96
      Width           =   450
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2191.frx":066A
      Height          =   2100
      Left            =   120
      TabIndex        =   7
      Top             =   2610
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   3704
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "A1702"
         Caption         =   "單據編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "A1703"
         Caption         =   "幣別"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "A1704"
         Caption         =   "金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a1706"
         Caption         =   "D/N No."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "A1705"
         Caption         =   "其他對象代號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "FagentName"
         Caption         =   "其他對象名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "a1716"
         Caption         =   "備註"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "a1718"
         Caption         =   "申請人"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   5280
      TabIndex        =   14
      Top             =   75
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   7680
      Top             =   1800
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   4320
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   510
      Width           =   1500
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "2646;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   510
      Width           =   2760
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "4868;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label7 
      Height          =   300
      Left            =   5880
      TabIndex        =   24
      Top             =   1295
      Width           =   2775
      VariousPropertyBits=   19
      Size            =   "4895;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   795
      Left            =   1560
      TabIndex        =   6
      Top             =   1710
      Width           =   3855
      VariousPropertyBits=   -1467989989
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "6800;1402"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCCnt 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "lblCCnt"
      Height          =   180
      Index           =   1
      Left            =   5520
      TabIndex        =   22
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label lblCCnt 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "lblCCnt"
      Height          =   180
      Index           =   0
      Left            =   5520
      TabIndex        =   21
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "申請人:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   1310
      Width           =   975
   End
   Begin VB.Label Label4 
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
      TabIndex        =   19
      Top             =   1680
      Width           =   570
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "金額"
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
      Top             =   1310
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
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
      Left            =   6030
      TabIndex        =   17
      Top             =   900
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "代理人D/N No."
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
      Left            =   90
      TabIndex        =   16
      Top             =   900
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4752
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "作業日期"
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
      Left            =   4290
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSForms.Label Label6 
      Height          =   300
      Left            =   2880
      TabIndex        =   23
      Top             =   135
      Width           =   4935
      VariousPropertyBits=   19
      Size            =   "8705;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "Frmacc2191"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/06 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text1、Text2、Text4、Label6、Label7
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset
Public adoacc170 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim adoaccrpt216 As New ADODB.Recordset
Dim adoaccrpt217 As New ADODB.Recordset
Dim adoaccrpt218 As New ADODB.Recordset
Dim dllaccrpt As Object
Dim strAutoNo As String
Dim strYes As String

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo2, Label5) = False Then
      Cancel = True
      Combo2.SetFocus
   End If
End Sub

Private Sub Command2_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2180.Show
   Unload Me
End Sub

Private Sub Command3_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2170.Show
   Unload Me
End Sub

Private Sub Command4_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2190.Show
   Unload Me
End Sub

Private Sub Command5_Click()
   AdodcDelete
End Sub

Private Sub Form_Activate()
       Me.Text5.SetFocus 'Add by Amy 2013/08/19
End Sub

'Added by Lydia 2021/12/06
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/06 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   'Modify by Amy 2013/08/19
'   Me.Width = 9000 '原8850
'   Me.Height = 5700 '原5500
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 8880, 5700, strBackPicPath1
   'end 2021/12/07
   MaskEdBox1.Mask = MsgText(601)
   MaskEdBox1.Text = CFDate(ACDate(ServerDate))
   MaskEdBox1.Mask = DFormat
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
   Set dllaccrpt = CreateObject("AccReport.ReportSelect")
   'Added by Lydia 2016/01/26
   lblCCnt(0) = ""
   lblCCnt(1) = "(最長 " & Text4.MaxLength & "字元)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '2005/12/16 CANCEL BY SONIA 婧瑄說要取消
   'Screen.MousePointer = vbHourglass
   'Frmacc2170.ProcessData
   'Unload Frmacc2170
   '2005/12/16 END
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Screen.MousePointer = vbDefault
   If strExitControl = MsgText(602) Then
      strFormName = MsgText(601)
      strTrackMode = "" 'Added by Lydia 2021/12/06 Form2.0 記錄鍵盤傳入順序(清除)
      KeyEnter vbKeyEscape
      MenuEnabled
      Set dllaccrpt = Nothing
      Set Frmacc2191 = Nothing
   End If
   strExitControl = MsgText(602)
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
End Sub

Private Sub Text1_GotFocus()
   'Add by Amy 2013/08/19
   Option1(1).Value = True
   Text5 = ""
   Text5.Enabled = False
   Label6.Caption = ""
   OpenIme
   'end 2013/08/19
   TextInverse Text1
End Sub

'Modified by Lydia 2021/12/06 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Text1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   'Modified by Lydia 2015/10/06 + A1718
   'adoadodc1.Open "select a1702, a1703, a1704, a1705, a1706, a1707, a1717 as FagentName, a1716 from acc170 where a1701 = '4' and (a1709 is null or a1709 = '')", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select a1702, a1703, a1704, a1705, a1706, a1707, a1717 as FagentName, a1716, DECODE(SUBSTR(A1718,1,1),'X',A1718||'0','') as a1718 from acc170 where a1701 = '4' and (a1709 is null or a1709 = '')", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo2.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   Combo2 = "USD"
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
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
   'Modified by Lydia 2015/10/06 +A1718
   'adoadodc1.Open "select a1702, a1703, a1704, a1705, a1706, a1707, a1717 as FagentName, a1716 from acc170 where a1701 = '4' and (a1709 is null or a1709 = '')", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select a1702, a1703, a1704, a1705, a1706, a1707, a1717 as FagentName, a1716, DECODE(SUBSTR(A1718,1,1),'X',A1718||'0','') as a1718 from acc170 where a1701 = '4' and (a1709 is null or a1709 = '')", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      'Modify by Amy 2013/08/19 +代理人編號輸入
      If Option1(0).Value Then
        Adodc1.Recordset.Find "a1705 = '" & ChgSQL(Text5) & "'", 0, adSearchForward, 1
      Else
        Adodc1.Recordset.Find "FagentName = '" & ChgSQL(Text1) & "'", 0, adSearchForward, 1
      End If
      'end 2013/08/19
      If Adodc1.Recordset.EOF Then
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(國外結匯資料)
'
'*************************************************
Private Sub Acc170Save()
   strAutoNo = AccAutoNo("B", 5)
   strYes = AccSaveAutoNo("B", Mid(strAutoNo, 5, 5))
   'Modify by Amy 2013/08/19 新增可以代理人編號輸入及增加D/N No(存入a1706)
   'adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1716, a1710, a1711, a1712, a1717) values ('4', '" & strAutoNo & "', " & CNULL(Combo2) & ", " & CNULL(Text3) & ", " & CNULL(ChgSQL(PUB_StrToStr(Text2, 9))) & ", " & CNULL(ChgSQL(Text4)) & ", " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', " & CNULL(ChgSQL(Text1)) & ")"
   'Modified by Lydia 2015/10/06 +A1718
   If Option1(0).Value Then
      '2015/3/9 modify by sonia 編號輸入者不必存a1717,否則水單列印時若同時有帳單資料則不會印合計
      'adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1716, a1710, a1711, a1712, a1717,a1706) values ('4', '" & strAutoNo & "', " & CNULL(Combo2) & ", " & CNULL(Text3) & ", " & CNULL(ChgSQL(PUB_StrToStr(Text5, 9))) & ", " & CNULL(ChgSQL(Text4)) & ", " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', " & CNULL(ChgSQL(Label6.Caption)) & "," & CNULL(ChgSQL(Trim(Text6))) & ")"
      adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1716, a1710, a1711, a1712, a1706, a1718) values ('4', '" & strAutoNo & "', " & CNULL(Combo2) & ", " & CNULL(Text3) & ", " & CNULL(ChgSQL(PUB_StrToStr(Text5, 9))) & ", " & CNULL(ChgSQL(Text4)) & ", " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', " & CNULL(ChgSQL(Trim(Text6))) & ", " & CNULL(Left(ChgSQL(Trim(Text7)), 8)) & ")"
   Else
      adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1716, a1710, a1711, a1712, a1717,a1706, a1718) values ('4', '" & strAutoNo & "', " & CNULL(Combo2) & ", " & CNULL(Text3) & ", " & CNULL(ChgSQL(PUB_StrToStr(Text2, 9))) & ", " & CNULL(ChgSQL(Text4)) & ", " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', " & CNULL(ChgSQL(Text1)) & "," & CNULL(ChgSQL(Trim(Text6))) & "," & CNULL(Left(ChgSQL(Trim(Text7)), 8)) & ")"
   End If
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
Dim b_Cancel As Boolean
On Error GoTo Checking

   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/06 Form2.0 記錄鍵盤傳入順序
   Select Case KeyCode
      Case vbKeyInsert
         'Added by Lydia 2021/12/06 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'Added by Lydia 2021/12/06 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "TextBox") = False Then
             Exit Sub
         End If
         'end 2021/12/06
    
        'Modify by Amy 2013/08/19 +代理人及代理人D/N No判斷
'         If Text3 = MsgText(601) Then
'            MsgBox MsgText(58), , MsgText(5)
'            Text3.SetFocus
'            Exit Sub
'         End If
         b_Cancel = False
         Text3_Validate b_Cancel
         If b_Cancel Then
            Exit Sub
         End If
         
         b_Cancel = False
         Text5_Validate b_Cancel
         If b_Cancel Then
            Exit Sub
         End If
         
         If Len(Trim(Text6)) = 0 Then
            MsgBox MsgText(52), , MsgText(5)
            Text6.SetFocus
            TextInverse Text6
            Exit Sub
         End If
         b_Cancel = False
         Text6_Validate b_Cancel
         If b_Cancel Then
            Exit Sub
         End If
         'Added by Lydia 2015/10/06 +A1918
         b_Cancel = False
         Text7_Validate b_Cancel
         If b_Cancel Then
            Exit Sub
         End If
         
         'Add by Amy 2013/11/18 +帳款處理訊息
         strExc(0) = GetDizhang("" & Text5, , True) '代理人
         'end 2013/11/18
         If CheckDN(IIf(Option1(0).Value = True, Text5, Text2), Trim(Text6)) Then
            If MsgBox("D/N No重覆,是否仍要輸入?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbNo Then
                Exit Sub
            End If
         End If
         'end 2013/08/19
         Acc170Save
         FormClear
         'Add by Amy 2017/10/26
         If Option1(0).Value = True Then
            Text5.SetFocus
         Else
            Text1.SetFocus
         End If
      Case vbKeyF12
         AdodcRefresh
   End Select
   KeyEnter KeyCode
Checking:
   Exit Sub
End Sub

'*************************************************
'  刪除資料表
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adoTaie.Execute "delete from acc170 where a1701 = '4' and a1702 = '" & Adodc1.Recordset.Fields("a1702").Value & "'"
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  產生結匯前核對表
'
'*************************************************
'2005/9/9 CANCEL BY SONIA 改抓Frmacc2170.ProcessData
'Private Sub ProcessData()

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Text2 = Mid(Text1, 1, 5)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   'Modify by Amy 2013/08/19 修正輸入0仍可insert
   'If Text3 = MsgText(601) Then
   If Text3 = MsgText(601) Or Val(Text3) = 0 Then
      MsgBox MsgText(58), , MsgText(5)
      Cancel = True
      Text3.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

'Modified by Lydia 2021/12/06 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Text4_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Combo2 = "USD"
   Text3 = ""
   Text4 = ""
   'Text1.SetFocus
   'add by Amy 2013/08/19
   'Mark by Amy 2017/10/26 會造成無窮迴圈故mark
   'Option1(0).Value = True '一開始若選代理人輸入完不會觸發option1_click
   'Option1(1).Value = True
   Text5 = ""
   Label6.Caption = ""
   Text6 = ""
   'end 2013/08/19
   'Added by Lydia 2015/10/06
   Text7 = ""
   Label7.Caption = ""
   'end 2015/10/06
Checking:
   Exit Sub
End Sub

'Add by Amy 2013/08/19
'+1.代理人及代理人D/N No欄位 2.改datagrid 其他對象代號名稱位置對調
Private Sub Option1_Click(Index As Integer)
      Select Case Index
        Case 0
            Text5.Enabled = True
            Text5.SetFocus
            Text5_GotFocus
            
        Case 1
            Text1.Enabled = True
            Text1.SetFocus
            Text1_GotFocus
    End Select
End Sub

Private Sub Text5_GotFocus()
    Option1(0).Value = True
    Text1 = ""
    Text1.Enabled = False
    CloseIme
    TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
    If Len(Text5) = 9 Then
        If Right(Text5, 1) <> "0" Then
            Cancel = True
            MsgBox "此編號為舊名稱編號, 不可輸入!"
            Text5.SetFocus
            TextInverse Text5
            Exit Sub
        End If
    End If
    
    Select Case Len(Text5)
      Case 6
        Text5 = AfterZero(Text5)
      Case 8
        Text5 = Text5 & "0"
    End Select
    
    If Text5 <> MsgText(601) Then
      If ExistCheck("fagent", "fa01", Mid(Text5, 1, 8), Option1(0).Caption) = False Then
         Cancel = True
         Text5.SetFocus
         TextInverse Text5
         Exit Sub
      End If
   End If
   
   Label6.Caption = FagentName(Text5)
End Sub

Private Sub Text6_GotFocus()
    TextInverse Text6
    CloseIme
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
    If CheckLengthIsOK(Trim(Text6), 50) = False Then
        Cancel = True
        Text6.SetFocus
        TextInverse Text6
    End If
End Sub

'*************************************************
'  依代理人編號查詢Fagent且fa02='0',帶出名稱(英->中->日)
'
'*************************************************
Private Function FagentName(InputNo As String) As String
    Dim adofagent As New ADODB.Recordset
   
   FagentName = MsgText(601)
   adofagent.CursorLocation = adUseClient
   strExc(0) = "select Decode(NVL(fa05||fa63||fa64||fa65,'0'),'0',Decode(NVL(fa04,'0'),'0',fa06,fa04),fa05||fa63||fa64||fa65) FName from fagent " & _
                    "where fa01 = '" & ChgSQL(Mid(InputNo, 1, 8)) & "' and fa02 = '0' "
   adofagent.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   If adofagent.RecordCount <> 0 Then
      adofagent.MoveFirst
      If Not IsNull(adofagent.Fields("FName")) Then
        FagentName = adofagent.Fields("FName")
      End If
   End If
   adofagent.Close
End Function

'*************************************************
'  檢查同一對象已有相同D/N No，若已有回傳True
'
'*************************************************
Private Function CheckDN(ByVal strNo, ByVal strDnNo) As Boolean
    Dim rsQuery As New ADODB.Recordset
    
    CheckDN = False
    rsQuery.CursorLocation = adUseClient
    strExc(0) = "Select * From acc170 Where a1701 = '4' and a1705='" & strNo & "' " & _
                     "and a1706='" & strDnNo & "'  and (a1709 is null or a1709 = '') "
    intI = 1
    Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        CheckDN = True
    End If
End Function
'end 2013/08/19

'Added by Lydia 2015/10/06 +A1718代為
Private Sub Text7_GotFocus()
    CloseIme
    TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
    If Len(Text7) > 0 Then
        If Len(Text7) = 9 Then
            If Right(Text7, 1) <> "0" Then
                Cancel = True
                MsgBox "此編號為舊名稱編號, 不可輸入!"
                Text7.SetFocus
                TextInverse Text7
                Exit Sub
            End If
        End If
        
        Select Case Len(Text7)
          Case 6
            Text7 = AfterZero(Text7)
          Case 8
            Text7 = Text7 & "0"
        End Select
        
'        If Text7 <> MsgText(601) Then
'          If ExistCheck("customer", "cu01", Mid(Text7, 1, 8), Option1(0).Caption) = False Then
'             Cancel = True
'             Text7.SetFocus
'             TextInverse Text7
'             Exit Sub
'          End If
'        End If
        If ClsPDGetCustomer(Text7, strExc(1)) = False Then
             Text7.SetFocus
             TextInverse Text7
             Exit Sub
        End If
        Label7.Caption = strExc(1)
        If Text4 = MsgText(601) Then
            MsgBox "輸入申請人,必要輸入備註!"
            Text4.SetFocus
            TextInverse Text4
            Exit Sub
        End If
    End If
End Sub
'end 2015/10/06

'Added by Lydia 2016/01/26 備註可折行
Private Sub Text4_LostFocus()
    lblCCnt(0) = ""
    If Text4 <> "" Then
       lblCCnt(0) = "已輸入 " & GetTextLength(Text4) & "字元"
    End If
End Sub
