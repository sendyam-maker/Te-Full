VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2214 
   AutoRedraw      =   -1  'True
   Caption         =   "付款資料查詢"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   8730
   Begin VB.CommandButton Command2 
      Caption         =   "匯票內容"
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
      Left            =   7320
      TabIndex        =   7
      Top             =   4272
      Width           =   1212
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
      Height          =   330
      Left            =   1290
      TabIndex        =   3
      Top             =   112
      Width           =   1572
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
      Height          =   330
      Left            =   4050
      TabIndex        =   2
      Top             =   112
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
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
      Height          =   330
      Left            =   2520
      TabIndex        =   1
      Top             =   4200
      Width           =   1332
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2214.frx":0000
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a1902"
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
         DataField       =   "a1903"
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
         DataField       =   "a1904"
         Caption         =   "單據金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a1907"
         Caption         =   "國內客戶名稱(收據抬頭)"
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
         DataField       =   "a1916"
         Caption         =   "個人/公司"
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
            Locked          =   -1  'True
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4410.142
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   480
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
   Begin MSForms.TextBox Text4 
      Height          =   330
      Left            =   5640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   112
      Width           =   2775
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "付款單號"
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
      Left            =   330
      TabIndex        =   6
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人 "
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
      Left            =   3090
      TabIndex        =   5
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   480
      TabIndex        =   4
      Top             =   4238
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc2214"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text3
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc180 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset

Private Sub Command2_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   tool3_enabled
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select a1b01, a1b02 from acc190, acc1b0 where a1908 = a1b01 and a1901 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields("a1b01").Value) Then
         strItemNo = ""
      Else
         strItemNo = adoaccsum.Fields("a1b01").Value
      End If
      If IsNull(adoaccsum.Fields("a1b02").Value) Then
         strCompanyNo = ""
      Else
         strCompanyNo = adoaccsum.Fields("a1b02").Value
      End If
   Else
      strItemNo = ""
      strCompanyNo = ""
   End If
   Frmacc2215.Show
   Me.Enabled = False
   Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   Me.Height = 5100
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 8850, 5100, strBackPicPath1
   'end 2021/12/09
   
   OpenTable
   SumShow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strItemNo = ""
   tool3_enabled
   Select Case strFormLink
      Case "Frmacc2210"
         Frmacc2210.Enabled = True
      Case "Frmacc2220"
         Frmacc2220.Enabled = True
   End Select
   Set Frmacc2214 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc180.CursorLocation = adUseClient
   adoacc180.Open "select * from acc180 where a1801 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   FormShow
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc190 where a1901 = '" & strItemNo & "' order by a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
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
   Text2 = strItemNo
   If IsNull(adoacc180.Fields("a1803").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoacc180.Fields("a1803").Value
   End If
End Sub

'*************************************************
'  合計顯示
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1904) from acc190 where a1901 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text5 = MsgText(601)
      Else
        'Modify By Cheng 2004/04/28
'         Text5 = Format(adoaccsum.Fields(0).Value, DDollar)
         Text5 = Format(adoaccsum.Fields(0).Value, FDollar)
        'End
      End If
   Else
      Text5 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

Private Sub Text3_Change()
   Text4 = FagentQuery(Text3, 2)
End Sub

