VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2130 
   AutoRedraw      =   -1  'True
   Caption         =   "暫收款退費作業"
   ClientHeight    =   5496
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8796
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5496
   ScaleWidth      =   8796
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   5
      Top             =   1347
      Width           =   1590
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2520
      Picture         =   "Frmacc2130.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   243
      Width           =   350
   End
   Begin VB.TextBox Text11 
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
      Height          =   330
      Left            =   6840
      TabIndex        =   21
      Top             =   1710
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox Text10 
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
      Height          =   330
      Left            =   4200
      MaxLength       =   13
      TabIndex        =   6
      Top             =   1347
      Visible         =   0   'False
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2130.frx":0102
      Height          =   2730
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   4805
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "a1301"
         Caption         =   "退費單號"
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
         DataField       =   "a1303"
         Caption         =   "暫收款單號"
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
         DataField       =   "a1302"
         Caption         =   "退費日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a1306"
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
      BeginProperty Column04 
         DataField       =   "a1307"
         Caption         =   "外幣金額"
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
      BeginProperty Column05 
         DataField       =   "a1304"
         Caption         =   "代理人"
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
            ColumnWidth     =   1404.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1284.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1344.189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text7 
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
      Height          =   330
      Left            =   6840
      TabIndex        =   16
      Top             =   1347
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   4
      Top             =   984
      Visible         =   0   'False
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
      Height          =   330
      Left            =   1320
      TabIndex        =   13
      Top             =   606
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      MaxLength       =   15
      TabIndex        =   2
      Top             =   228
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   228
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   6840
      TabIndex        =   3
      Top             =   228
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   593
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   150
      Top             =   2400
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
   Begin MSForms.TextBox Text8 
      Height          =   555
      Left            =   1320
      TabIndex        =   7
      Top             =   1710
      Width           =   4455
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "7858;979"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   330
      Left            =   2910
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   606
      Width           =   5535
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "9763;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   330
      Left            =   2910
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   984
      Visible         =   0   'False
      Width           =   5532
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "9758;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "台幣金額"
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
      Left            =   5880
      TabIndex        =   20
      Top             =   1748
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "匯率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   19
      Top             =   1386
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   18
      Top             =   1386
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2295
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Top             =   1748
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "外幣金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5880
      TabIndex        =   15
      Top             =   1386
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   14
      Top             =   1023
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   12
      Top             =   645
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "退費日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5880
      TabIndex        =   11
      Top             =   267
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "暫收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   10
      Top             =   267
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "暫退單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   9
      Top             =   267
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc2130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text4、Text6、Text8
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc120 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public douAmount As Double
Public strNDate As String
Public strCurrency As String
Public strDocNo As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo1, Label8) = False Then
      Cancel = True
      Combo1.SetFocus
   End If
End Sub

Private Sub Command3_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text2 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a1301 = '" & Text2 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      AdodcRefresh
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a1301 = '" & strItemNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

'Added by Lydia 2021/12/03
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
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
'   Me.Width = 8850
'   Me.Height = 5500
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
   'Modify by Amy 2023/08/18 W8850 H5800
   PUB_InitForm Me, 8890, 5940, strBackPicPath1
   'end 2021/12/07
   
   MaskEdBox1.Mask = DFormat
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveLast
      Adodc1.Recordset.MoveFirst
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
   strTrackMode = "" 'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2130 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text2 = UpdateNo("acc130", "a1301", 5, MaskEdBox1.Text, MsgText(810))
   Else
      'Text2 = AutoNo(MsgText(810), 5)
      Text2 = strDocNo
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  查詢資料表(國外暫收款資料)
'
'*************************************************
Public Sub Acc120Query()
Dim douAmount As Double

   adoacc120.CursorLocation = adUseClient
   adoacc120.Open "select * from acc120 where a1201 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc120.RecordCount = 0 Then
      MsgBox MsgText(33), , MsgText(5)
      adoacc120.Close
      QueryClear
      Exit Sub
   End If
   If IsNull(adoacc120.Fields("a1203").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoacc120.Fields("a1203").Value
   End If
   If IsNull(adoacc120.Fields("a1204").Value) Then
      strCurrency = ""
   Else
      strCurrency = adoacc120.Fields("a1204").Value
   End If
   If IsNull(adoacc120.Fields("a1202").Value) Then
      strNDate = ""
   Else
      strNDate = adoacc120.Fields("a1202").Value
   End If
   If IsNull(adoacc120.Fields("a1204").Value) Then
      Combo1 = ""
   Else
      Combo1 = adoacc120.Fields("a1204").Value
   End If
'   If IsNull(adoacc120.Fields("a1206").Value) Then
'      Text5 = MsgText(601)
'   Else
'      Text5 = adoacc120.Fields("a1206").Value
'   End If
'   If IsNull(adoacc120.Fields("a1204").Value) Then
'      Text9 = MsgText(601)
'   Else
'      Text9 = adoacc120.Fields("a1204").Value
'   End If
'   If IsNull(adoacc120.Fields("a1205").Value) Then
'      Text10 = MsgText(601)
'   Else
'      Text10 = adoacc120.Fields("a1205").Value
'   End If
'   If adoquery.State = adStateOpen Then
'      adoquery.Close
'   End If
'   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select sum(nvl(a1p21, 0)) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p30 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount <> 0 Then
'      If IsNull(adoquery.Fields(0).Value) Then
'         douAmount = 0
'      Else
'         douAmount = Val(adoquery.Fields(0).Value)
'      End If
'   Else
'      douAmount = 0
'   End If
'   adoquery.Close
   If IsNull(adoacc120.Fields("a1207").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Format(Val(adoacc120.Fields("a1207").Value), FAmount)
      douAmount = Val(Format(Val(adoacc120.Fields("a1205").Value) * Val(Text7), FAmount))
   End If
   adoacc120.Close
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc130 order by a1301 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo1.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   Combo1 = "USD"
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
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc130 order by a1301 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a1301 = '" & Text2 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
      End If
   End If
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
   Text2 = Adodc1.Recordset.Fields("a1301").Value
   If IsNull(Adodc1.Recordset.Fields("a1303").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a1303").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a1302").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a1302").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a1304").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a1304").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1305").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("a1305").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1306").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("a1306").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1309").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a1309").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1307").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Format(Adodc1.Recordset.Fields("a1307").Value, FAmount)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1308").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1308").Value
   End If
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyF12
         Acc120Query
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   If strSaveConfirm = MsgText(3) Then
      If Acc130Query Then
         Cancel = True
         Text1.SetFocus
         Exit Sub
      End If
   End If
   Acc120Query
End Sub

Private Sub Text10_Change()
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   Text11 = Format(Val(Text10) * Val(Text7), FAmount)
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Change()
   'Modify by Morgan 2006/7/20
   'Text4 = FagentQuery(Text3, 2)
   If Left(Text3, 1) = "X" Then
      Text4 = CustomerQuery(Text3, 2)
   Else
      Text4 = FagentQuery(Text3, 2)
   End If
End Sub

Private Sub Text5_Change()
   Text6 = A0102Query(Text5)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text7_Change()
   If Text7 = MsgText(601) Then
      Exit Sub
   End If
   Text11 = Format(Val(Text10) * Val(Text7), FAmount)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

'*************************************************
'  清除查詢資料
'
'*************************************************
Private Sub QueryClear()
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   'edit by nickc 2007/02/08
   'Text9 = ""
   Text10 = ""
   Text7 = ""
   Text11 = ""
End Sub

'*************************************************
'  查詢資料表(國外已暫收款退費資料)
'
'*************************************************
Public Function Acc130Query() As Boolean
   adoacc120.CursorLocation = adUseClient
   adoacc120.Open "select * from acc130 where a1303 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc120.RecordCount <> 0 Then
      MsgBox MsgText(166), , MsgText(5)
      adoacc120.Close
      Acc130Query = True
      Exit Function
   End If
   adoacc120.Close
   Acc130Query = False
End Function

Private Sub Text8_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub
