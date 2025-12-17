VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2240 
   AutoRedraw      =   -1  'True
   Caption         =   "案件損益查詢"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8760
   Begin VB.TextBox Text13 
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
      Height          =   324
      Left            =   4600
      TabIndex        =   22
      Top             =   4536
      Width           =   1300
   End
   Begin VB.TextBox Text8 
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
      Height          =   300
      Left            =   1440
      TabIndex        =   19
      Top             =   4572
      Width           =   1596
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2240.frx":0000
      Height          =   3168
      Left            =   216
      TabIndex        =   18
      Top             =   1296
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   5583
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "案件損益查詢"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "cp05"
         Caption         =   "收文日"
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
      BeginProperty Column01 
         DataField       =   "PropertyName"
         Caption         =   "案件性質"
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
         DataField       =   "RecAmount"
         Caption         =   "應收金額"
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
         DataField       =   "cp18"
         Caption         =   "點數"
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
      BeginProperty Column04 
         DataField       =   "PayAmount"
         Caption         =   "代理人帳單金額"
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
      BeginProperty Column05 
         DataField       =   "st02"
         Caption         =   "智權人員"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "FagentName"
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
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4500.284
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00C0C0FF&
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
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   7395
      MaxLength       =   14
      TabIndex        =   16
      Top             =   885
      Width           =   1140
   End
   Begin VB.TextBox Text4 
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
      Left            =   5760
      TabIndex        =   15
      Top             =   108
      Width           =   576
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
      Height          =   330
      Left            =   3072
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   396
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
      Height          =   330
      Left            =   2724
      TabIndex        =   3
      Top             =   120
      Width           =   348
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
      Height          =   330
      Left            =   2484
      TabIndex        =   2
      Top             =   120
      Width           =   240
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
      Left            =   1704
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   780
   End
   Begin VB.TextBox Text11 
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
      Left            =   1212
      MaxLength       =   14
      TabIndex        =   8
      Top             =   864
      Width           =   1560
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
      Height          =   330
      Left            =   1212
      MaxLength       =   3
      TabIndex        =   0
      Top             =   120
      Width           =   492
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
      Height          =   330
      Left            =   6324
      TabIndex        =   7
      Top             =   108
      Width           =   2196
   End
   Begin VB.TextBox Text2 
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
      Height          =   300
      Left            =   6924
      TabIndex        =   6
      Top             =   4548
      Width           =   1596
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   336
      Left            =   240
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   2790
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   864
      Width           =   2715
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "4789;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   360
      Left            =   1215
      TabIndex        =   5
      Top             =   450
      Width           =   7305
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12885;635"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "財務其他支出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   3132
      TabIndex        =   23
      Top             =   4584
      Width           =   1368
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   $"Frmacc2240.frx":0015
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
      Height          =   450
      Left            =   240
      TabIndex        =   21
      Top             =   4920
      Width           =   7700
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "舊系統支出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   252
      TabIndex        =   20
      Top             =   4596
      Width           =   1140
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "浮動："
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
      Left            =   6735
      TabIndex        =   17
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "申請人："
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
      Left            =   75
      TabIndex        =   13
      Top             =   864
      Width           =   1155
   End
   Begin VB.Label Label11 
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
      Left            =   75
      TabIndex        =   12
      Top             =   135
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "案件名稱："
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
      Left            =   75
      TabIndex        =   11
      Top             =   510
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "申請國家："
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
      Left            =   4635
      TabIndex        =   10
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "目前盈虧"
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
      Left            =   5964
      TabIndex        =   9
      Top             =   4596
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -12
      Top             =   4752
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc2240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Combo1、Text3
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

'2006/1/6 整理
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSql As String
Dim strSQL1 As String
Dim strSQL2, StrSQL3, StrSQL4 As String 'add by sonia 2021/3/26
Dim strFloat As String
Dim m_PA08 As String

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo1_GotFocus()
OpenIme
End Sub

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo1_Validate(Cancel As Boolean)
   CloseIme
End Sub

Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5800
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   StatusView MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2240 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   Select Case Text10
      Case "TF"
         Text12.Visible = True
         Text5.MaxLength = 5 '2010/7/2 ADD BY SONIA
      Case Else
         Text12.Visible = False
         Text5.MaxLength = 6 '2010/7/2 ADD BY SONIA
   End Select
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'Modified by Lydia 2022/01/26 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Text3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from caseprogress where cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(傳票資料)
'
'*************************************************
Public Sub QueryTable()
   
On Error GoTo Checking
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/22 清除查詢印表記錄檔欄位
   strSql = ""
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   If Text10 <> MsgText(601) Then
      strSql = strSql & " and cp01 = '" & Text10 & "'"
      pub_QL05 = pub_QL05 & ";" & Label11 & Text10 'Add By Sindy 2010/12/22
   End If
   
   'modify by sonia 改用IF,因為CFP只有EPC(221)才要母子案一起計算,否則CFP案在抓ACC021都會很久
   'Select Case Text10
   '   Case "TF"
      If Text10 = "TF" Then
         If Text5 <> MsgText(601) And Text7 <> MsgText(601) Then
            '2006/1/5 MODIFY BY SONIA
            strSql = strSql & " and cp02 = '" & Text5 & Text7 & "'"
            pub_QL05 = pub_QL05 & "-" & Text5 & Text7 'Add By Sindy 2010/12/22
            '2010/3/26 cancel by sonia 領土延伸不可一起算,故改回
            'strSql = strSql & " and SUBSTR(cp02,1,5) = '" & Text5 & "'"
         End If
         '2006/1/5 CANCEL BY SONIA 母案子案一起計算
         'If Text9 <> MsgText(601) Then
         '   strSQL = strSQL & " and cp03 = '" & Text9 & "'"
         'End If
         'If Text12 <> MsgText(601) Then
         '   strSQL = strSQL & " and cp04 = '" & Text12 & "'"
         'End If
         '2006/1/5 END
         'add by nickc 2005/11/02
         Text6 = GetFloatPrepareCase(Text10.Text, Text5.Text & Text7.Text, Text9.Text, Text12.Text)
      '2006/1/5 ADD BY SONIA CFP母案子案一起計算
      'modify by sonia 改用IF,因為CFP只有EPC(221)才要母子案一起計算,否則CFP案在抓ACC021都會很久
      'Case "CFP"
      ElseIf Text10 = "CFP" And Text4 = "221" Then
         If Text5 <> MsgText(601) Then
            strSql = strSql & " and cp02 = '" & Text5 & "'"
            pub_QL05 = pub_QL05 & "-" & Text5 'Add By Sindy 2010/12/22
         End If
         '2010/3/26 ADD BY SONIA 接續案個別計算
         If Text9 <> MsgText(601) Then
            strSql = strSql & " and cp03 = '" & Text7 & "'"
            pub_QL05 = pub_QL05 & "-" & Text7 'Add By Sindy 2010/12/22
         End If
         '2010/3/26 END
         Text6 = GetFloatPrepareCase(Text10.Text, Text5.Text, Text7.Text, Text9.Text)
      '2006/1/5 END
      'modify by sonia 改用IF,因為CFP只有EPC(221)才要母子案一起計算,否則CFP案在抓ACC021都會很久
      'Case Else
      Else
         If Text5 <> MsgText(601) Then
            strSql = strSql & " and cp02 = '" & Text5 & "'"
            pub_QL05 = pub_QL05 & "-" & Text5 'Add By Sindy 2010/12/22
         End If
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and cp03 = '" & Text7 & "'"
            pub_QL05 = pub_QL05 & "-" & Text7 'Add By Sindy 2010/12/22
         End If
         If Text9 <> MsgText(601) Then
            strSql = strSql & " and cp04 = '" & Text9 & "'"
            pub_QL05 = pub_QL05 & "-" & Text9 'Add By Sindy 2010/12/22
         End If
         'add by nickc 2005/11/02
         Text6 = GetFloatPrepareCase(Text10.Text, Text5.Text, Text7.Text, Text9.Text)
   'modify by sonia 改用IF,因為CFP只有EPC(221)才要母子案一起計算,否則CFP案在抓ACC021都會很久
   'End Select
      End If
      
   'edit by nickc 2005/11/02 浮動改抓GetFloatPrepareCase
   
   '92.4.16 modify by sonia: cp75改成cp16-cp77
   '93.7.16 MODIFY BY SONIA cp16-cp77 改成 DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))
   '93.11.2 MODIFY BY SONIA 加入 CP09
   '2005/11/15 MODIFY BY SONIA 已付改抓AXF04*ACC190之A1906實際付款匯率,未付才抓AXF15
   '2006/1/6 MODIFY BY SONIA只顯示有收費或有帳單的進度資料
   'adoadodc1.Open "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, a1505, st02, cp61, sum(decode(a1507, null, nvl(NVL(AXF04*A1906,axf15), 0), 0)) as PayAmount, sum(nvl(axf04, 0)) as FpayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190 " & _
   '               "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null)" & strSQL & " group by cp05, decode(cpm03, '（無）', cpm04, cpm03), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), a1505, st02, cp61, CP09 union " & _
   '               "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, nvl(a1k30, 0) as RecAmount, nvl(cp18, 0) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, a1505, st02, cp61, sum(decode(a1507, null, nvl(NVL(AXF04*A1906,axf15), 0), 0)) as PayAmount, sum(decode(a1507, null, nvl(axf04, 0), 0)) as FpayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) and cp13 = st01 (+) and cp60 = a1k01 (+) and substr(cp60, 1, 1) = 'X'" & strSQL & " group by cp05, decode(cpm03, '（無）', cpm04, cpm03), nvl(a1k30, 0), cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), a1505, st02, cp61, CP09", adoTaie, adOpenStatic, adLockReadOnly
   '2006/7/17 MODIFY BY SONIA 抵帳改抓抵帳匯率
   'adoadodc1.Open "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, a1505, st02, cp61, sum(decode(a1507, null, nvl(NVL(AXF04*A1906,axf15), 0), 0)) as PayAmount, sum(nvl(axf04, 0)) as FpayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190 " & _
   '               "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null)" & strSQL & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))))>0 OR (decode(a1507, null, nvl(NVL(AXF04*A1906,axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode(cpm03, '（無）', cpm04, cpm03), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), a1505, st02, cp61, CP09 union " & _
   '               "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, NVL(nvl(a1k30, A1K11),0) as RecAmount, nvl(cp18, 0) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, a1505, st02, cp61, sum(decode(a1507, null, nvl(NVL(AXF04*A1906,axf15), 0), 0)) as PayAmount, sum(decode(a1507, null, nvl(axf04, 0), 0)) as FpayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190 " & _
   '               "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) and cp13 = st01 (+) and cp60 = a1k01 (+) and substr(cp60, 1, 1) = 'X'" & strSQL & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))))>0 OR (decode(a1507, null, nvl(NVL(AXF04*A1906,axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode(cpm03, '（無）', cpm04, cpm03), NVL(nvl(a1k30, A1K11),0), cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), a1505, st02, cp61, CP09", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Morgan 2007/3/13 當未結匯(a1906=0)時也抓axf15
   '2010/3/17 MODIFY BY SONIA CFT-012517-->DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))改成DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)
   '2010/3/18 MODIFY BY SONIA 取消帳單幣別A1505及付款外幣金額FpayAmount,否則一收文號不同幣別帳單會出現多筆CFT-012231
   'adoadodc1.Open "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, a1505, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0)) as PayAmount, sum(nvl(axf04, 0)) as FpayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190,ACC1G0 " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null)" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode(cpm03, '（無）', cpm04, cpm03), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), a1505, st02, cp61, CP09 union " & _
                  "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, NVL(nvl(a1k30, A1K11),0) as RecAmount, nvl(cp18, 0) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, a1505, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0)) as PayAmount, sum(decode(a1507, null, nvl(axf04, 0), 0)) as FpayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190,ACC1G0 " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) and cp13 = st01 (+) and cp60 = a1k01 (+) and substr(cp60, 1, 1) = 'X'" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode(cpm03, '（無）', cpm04, cpm03), NVL(nvl(a1k30, A1K11),0), cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), a1505, st02, cp61, CP09", adoTaie, adOpenStatic, adLockReadOnly
   '2010/3/18 MODIFY BY SONIA CFP-022665發明申請銷帳只銷服務費時點數CP18算錯,(CP16-CP77) / 1000再改為(CP16-A1U07合計-CP17) / 1000
   'adoadodc1.Open "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190,ACC1G0 " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null)" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode(cpm03, '（無）', cpm04, cpm03), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), st02, cp61, CP09 union " & _
                  "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, NVL(nvl(a1k30, A1K11),0) as RecAmount, nvl(cp18, 0) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190,ACC1G0 " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) and cp13 = st01 (+) and cp60 = a1k01 (+) and substr(cp60, 1, 1) = 'X'" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode(cpm03, '（無）', cpm04, cpm03), NVL(nvl(a1k30, A1K11),0), cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), st02, cp61, CP09", adoTaie, adOpenStatic, adLockReadOnly
   '2010/3/29 MODIFY BY SONIA 未付帳單改抓ACC210最新預估匯率計算
   'adoadodc1.Open "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-X.A1U07-CP17) / 1000)) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190,ACC1G0, " & _
                  "(SELECT A1U03,SUM(A1U07) A1U07 FROM ACC1U0,CASEPROGRESS WHERE CP09=A1U03 " & strSql & " GROUP BY A1U03) X " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) AND CP09=X.A1U03(+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null)" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-X.A1U07-CP17) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode(cpm03, '（無）', cpm04, cpm03), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-X.A1U07-CP17) / 1000)), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), st02, cp61, CP09 union " & _
                  "select nvl(cp05 - 19110000, 0) as cp05, decode(cpm03, '（無）', cpm04, cpm03) as PropertyName, NVL(nvl(a1k30, A1K11),0) as RecAmount, nvl(cp18, 0) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190,ACC1G0 " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) and cp13 = st01 (+) and cp60 = a1k01 (+) and substr(cp60, 1, 1) = 'X'" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode(cpm03, '（無）', cpm04, cpm03), NVL(nvl(a1k30, A1K11),0), cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), st02, cp61, CP09", adoTaie, adOpenStatic, adLockReadOnly
   '2010/5/11 MODIFY BY SONIA 案件性質改以申請國家判斷T-164001
   '2011/9/20 modify by sonia 抓acc1k0時原以cp60 = a1k01 (+),改以本所案號抓,否則會串不到CFT-013202的X10009761,另請款單的RecAmount及cp18改抓法
   'adoadodc1.Open "select nvl(cp05 - 19110000, 0) as cp05, decode('" & Text4 & "', '000', cpm03, cpm04) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-X.A1U07-CP17) / 1000)) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190,ACC1G0,ACC210, " & _
                  "(SELECT A1U03,SUM(A1U07) A1U07 FROM ACC1U0,CASEPROGRESS WHERE CP09=A1U03 " & strSql & " GROUP BY A1U03) X, " & _
                  "(SELECT A1505 CURR,MAX(A2101) PAYDATE FROM ACC150,ACC151,ACC210,CASEPROGRESS WHERE CP01||CP02||CP03||CP04=AXF03(+)" & strSql & " AND AXF01=A1501(+) AND A1505=A2102(+) GROUP BY A1505) Y " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) AND CP09=X.A1U03(+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) AND A1505=Y.CURR(+) AND Y.CURR=A2102(+) AND Y.PAYDATE=A2101(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null)" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-X.A1U07-CP17) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode('" & Text4 & "', '000', cpm03, cpm04), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-X.A1U07-CP17) / 1000)), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), st02, cp61, CP09 union " & _
                  "select nvl(cp05 - 19110000, 0) as cp05, decode('" & Text4 & "', '000', cpm03, cpm04) as PropertyName, NVL(nvl(a1k30, A1K11),0) as RecAmount, nvl(cp18, 0) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190,ACC1G0,ACC210, " & _
                  "(SELECT A1505 CURR,MAX(A2101) PAYDATE FROM ACC150,ACC151,ACC210,CASEPROGRESS WHERE CP01||CP02||CP03||CP04=AXF03(+)" & strSql & " AND AXF01=A1501(+) AND A1505=A2102(+) GROUP BY A1505) Y " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) AND A1505=Y.CURR(+) AND Y.CURR=A2102(+) AND Y.PAYDATE=A2101(+) and cp13 = st01 (+) and cp60 = a1k01 (+) and substr(cp60, 1, 1) = 'X'" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode('" & Text4 & "', '000', cpm03, cpm04), NVL(nvl(a1k30, A1K11),0), cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), st02, cp61, CP09", adoTaie, adOpenStatic, adLockReadOnly
   '2015/7/3 MODIFY BY SONIA CFP-026816 實審取消收文且未開立收據也不算故(substr(cp60, 1, 1) = 'E' or cp60 is null)改為(substr(cp60, 1, 1) = 'E' or (cp60 is null AND NVL(CP57,0)=0))
   'modify by sonia 2021/4/8 +order by cp05,cp09 CFP-030420同一請款單只顯示在有收文費用之最小收文號,P-097213之抵帳單ACC160要從帳單金額加回來
   'adoadodc1.Open "select nvl(cp05 - 19110000, 0) as cp05, decode('" & Text4 & "', '000', cpm03, cpm04) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190,ACC1G0,ACC210, " & _
                  "(SELECT A1U03,SUM(A1U07) A1U07 FROM ACC1U0,CASEPROGRESS WHERE CP09=A1U03 " & strSql & " GROUP BY A1U03) X, " & _
                  "(SELECT A1505 CURR,MAX(A2101) PAYDATE FROM ACC150,ACC151,ACC210,CASEPROGRESS WHERE CP01||CP02||CP03||CP04=AXF03(+)" & strSql & " AND AXF01=A1501(+) AND A1505=A2102(+) GROUP BY A1505) Y " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) AND CP09=X.A1U03(+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) AND A1505=Y.CURR(+) AND Y.CURR=A2102(+) AND Y.PAYDATE=A2101(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or (cp60 is null AND NVL(CP57,0)=0))" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode('" & Text4 & "', '000', cpm03, cpm04), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), st02, cp61, CP09 union " & _
                  "select nvl(cp05 - 19110000, 0) as cp05, decode('" & Text4 & "', '000', cpm03, cpm04) as PropertyName, DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)) as RecAmount, (nvl(a1k11, 0)-nvl(a1k09,0))/1000 as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, st02, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190,ACC1G0,ACC210, " & _
                  "(SELECT A1505 CURR,MAX(A2101) PAYDATE FROM ACC150,ACC151,ACC210,CASEPROGRESS WHERE CP01||CP02||CP03||CP04=AXF03(+)" & strSql & " AND AXF01=A1501(+) AND A1505=A2102(+) GROUP BY A1505) Y " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) AND A1505=Y.CURR(+) AND Y.CURR=A2102(+) AND Y.PAYDATE=A2101(+) and cp13 = st01 (+) and cp01=a1k13(+) and cp02=a1k14(+) and cp03=a1k15(+) and cp04=a1k16(+) and cp60=a1k01(+) and substr(cp60, 1, 1) = 'X'" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode('" & Text4 & "', '000', cpm03, cpm04), DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)), (nvl(a1k11, 0)-nvl(a1k09,0))/1000, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), st02, cp61, CP09", adoTaie, adOpenStatic, adLockReadOnly
   '2021/4/15 modify by sonia 解決FF案件請款單重覆計算問題CFP-030420
   'adoadodc1.Open "select nvl(cp05 - 19110000, 0) as cp05, decode('" & Text4 & "', '000', cpm03, cpm04) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(a.A1906,0,null,a.A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))-sum(nvl(AXg04*b.A1906,0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190 a,ACC1G0,ACC210,acc161,acc190 b, " & _
                  "(SELECT A1U03,SUM(A1U07) A1U07 FROM ACC1U0,CASEPROGRESS WHERE CP09=A1U03 " & strSql & " GROUP BY A1U03) X, " & _
                  "(SELECT A1505 CURR,MAX(A2101) PAYDATE FROM ACC150,ACC151,ACC210,CASEPROGRESS WHERE CP01||CP02||CP03||CP04=AXF03(+)" & strSql & " AND AXF01=A1501(+) AND A1505=A2102(+) GROUP BY A1505) Y " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) AND CP09=X.A1U03(+) and axf01 = a1501 (+) AND AXF01=a.A1902(+) AND A1512=A1G01(+) AND A1505=Y.CURR(+) AND Y.CURR=A2102(+) AND Y.PAYDATE=A2101(+) and cp09=axg02(+) and axg01=b.a1902(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or (cp60 is null AND NVL(CP57,0)=0))" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(a.A1906,0,null,a.A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode('" & Text4 & "', '000', cpm03, cpm04), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), cp61, CP09 union " & _
                  "select nvl(cp05 - 19110000, 0) as cp05, decode('" & Text4 & "', '000', cpm03, cpm04) as PropertyName, DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)) as RecAmount, (nvl(a1k11, 0)-nvl(a1k09,0))/1000 as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(a.A1906,0,null,a.A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))-sum(nvl(AXg04*b.A1906,0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190 a,ACC1G0,ACC210,acc161,acc190 b, " & _
                  "(SELECT A1505 CURR,MAX(A2101) PAYDATE FROM ACC150,ACC151,ACC210,CASEPROGRESS WHERE CP01||CP02||CP03||CP04=AXF03(+)" & strSql & " AND AXF01=A1501(+) AND A1505=A2102(+) GROUP BY A1505) Y " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=a.A1902(+) AND A1512=A1G01(+) AND A1505=Y.CURR(+) AND Y.CURR=A2102(+) AND Y.PAYDATE=A2101(+) and cp09=axg02(+) and axg01=b.a1902(+) and cp13 = st01 (+) and cp01=a1k13(+) and cp02=a1k14(+) and cp03=a1k15(+) and cp04=a1k16(+) and cp60=a1k01(+) and substr(cp60, 1, 1) = 'X'" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(a.A1906,0,null,a.A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode('" & Text4 & "', '000', cpm03, cpm04), DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)), (nvl(a1k11, 0)-nvl(a1k09,0))/1000, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), cp61, CP09 order by cp05,cp09", adoTaie, adOpenStatic, adLockReadOnly
   '2021/4/30 modify by sonia 1.解決抵帳單輸在無收入無支出的進度CFT-012084(V09900104輸在勝訴),2.CFP-025351帳單輸在閉卷未顯示故加CP61 is not null條件
   adoadodc1.Open "SELECT cp05, PropertyName,RecAmount,cp18, st02,FagentName, cp61,PayAmount-nvl(axg04,0) PayAmount, CP09 FROM " & _
                  "(select nvl(cp05 - 19110000, 0) as cp05, decode('" & Text4 & "', '000', cpm03, cpm04) as PropertyName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)) as cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(a.A1906,0,null,a.A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff,ACC190 a,ACC1G0,ACC210, " & _
                  " (SELECT A1U03,SUM(A1U07) A1U07 FROM ACC1U0,CASEPROGRESS WHERE CP09=A1U03 " & strSql & " GROUP BY A1U03) X, " & _
                  " (SELECT A1505 CURR,MAX(A2101) PAYDATE FROM ACC150,ACC151,ACC210,CASEPROGRESS WHERE CP01||CP02||CP03||CP04=AXF03(+)" & strSql & " AND AXF01=A1501(+) AND A1505=A2102(+) GROUP BY A1505) Y " & _
                  "  where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) AND CP09=X.A1U03(+) and axf01 = a1501 (+) AND AXF01=a.A1902(+) AND A1512=A1G01(+) AND A1505=Y.CURR(+) AND Y.CURR=A2102(+) AND Y.PAYDATE=A2101(+) and cp13 = st01 (+) and (substr(cp60, 1, 1) = 'E' or (cp60 is null AND NVL(CP57,0)=0) or (cp60 is null and CP61 IS NOT NULL))" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(a.A1906,0,null,a.A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, decode('" & Text4 & "', '000', cpm03, cpm04), nvl(cp16,0) - nvl(cp77,0), decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-nvl(X.A1U07,0)-nvl(CP17,0)) / 1000)), st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), cp61, CP09 " & _
                  " UNION select nvl(cp05 - 19110000, 0) AS cp05, decode('239', '000', cpm03, cpm04) AS PropertyName, 0 AS RecAmount, 0 AS cp18, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) AS FagentName, cp61, 0 AS PayAmount, CP09 FROM caseprogress, casepropertymap, fagent, acc151, acc161,staff " & _
                  "  WHERE cp01 = cpm01 AND cp10 = cpm02 AND substr(cp44, 1, 8) = fa01 (+) AND substr(cp44, 9, 1) = fa02 (+)" & strSql & " AND cp09 = axf02 (+) AND axf01 IS NULL AND cp09=axg02(+) AND axg02 IS NOT NULL AND cp13=st01(+) )," & _
                  "(SELECT axg02,sum(axg04*a1906) axg04 FROM acc161,acc190,caseprogress WHERE CP09=AXG02(+)" & strSql & " AND axg01=a1902(+) and axg02 is not null GROUP BY axg02) where cp09=axg02(+) " & _
                  "union " & _
                  "SELECT cp05, PropertyName,RecAmount,cp18, st02,FagentName, cp61,PayAmount-nvl(axg04,0) PayAmount, CP09 FROM " & _
                  "(select nvl(cp05 - 19110000, 0) as cp05, max(decode('" & Text4 & "', '000', cpm03, cpm04)) as PropertyName, max(decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0)) as RecAmount, max(decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0)) as cp18, st02, max(nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06))) as FagentName, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(a.A1906,0,null,a.A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0)) as PayAmount, CP09 from caseprogress, casepropertymap, fagent, acc151, acc150, staff, acc1k0,ACC190 a,ACC1G0,ACC210, " & _
                  " (SELECT A1505 CURR,MAX(A2101) PAYDATE FROM ACC150,ACC151,ACC210,CASEPROGRESS WHERE CP01||CP02||CP03||CP04=AXF03(+)" & strSql & " AND AXF01=A1501(+) AND A1505=A2102(+) GROUP BY A1505) Y,(select cp60 z1,min(cp09) z2 from caseprogress where 1=1 " & strSql & " group by cp60) Z  " & _
                  "where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=a.A1902(+) AND A1512=A1G01(+) AND A1505=Y.CURR(+) AND Y.CURR=A2102(+) AND Y.PAYDATE=A2101(+) and cp13 = st01 (+) and cp01=a1k13(+) and cp02=a1k14(+) and cp03=a1k15(+) and cp04=a1k16(+) and cp60=a1k01(+) and substr(cp60, 1, 1) = 'X'" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 or cp60>'X' OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(a.A1906,0,null,a.A1906),AXF04*A1G03),AXF04*NVL(A2103,1)), 0), 0))>0 OR (nvl(axf04, 0))>0) and z1(+)=cp60 group by cp05, decode('" & Text4 & "', '000', cpm03, cpm04), DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)), (nvl(a1k11, 0)-nvl(a1k09,0))/1000, st02, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)), cp61, CP09), " & _
                  "(SELECT axg02,sum(axg04*a1906) axg04 FROM acc161,acc190,caseprogress WHERE CP09=AXG02(+)" & strSql & " AND axg01=a1902(+) and axg02 is not null GROUP BY axg02) where cp09=axg02(+) order by cp05,cp09", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.ReQuery
   If Adodc1.Recordset.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
      Text2 = 0: Text6 = 0: Text8 = 0: Text13 = 0   'add by sonia 2021/5/5
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      InsertQueryLog (Adodc1.Recordset.RecordCount) 'Add By Sindy 2010/12/22
   End If
   
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
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            Select Case Text10
               Case "TF"
                  If Text9 = "" Then
                     Text9 = "0"
                  End If
                  If Text12 = "" Then
                     Text12 = "00"
                  End If
               Case Else
                  If Text7 = "" Then
                     Text7 = "0"
                  End If
                  If Text9 = "" Then
                     Text9 = "00"
                  End If
            End Select
            FormShow
            QueryTable
            Calculate
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   StatusView MsgText(98)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text10 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text5 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text7 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text9 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text10 = "TF" Then
      If Text12 <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   FormCheck = False
End Function

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Select Case Text10
      Case "TF"
         Text9 = "0"
         Text12 = "00"
      Case Else
         Text7 = "0"
         Text9 = "00"
   End Select
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  顯示畫面
'
'*************************************************
Public Sub FormShow()
   Combo1.Clear
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   '92.3.3 加PA08
   adoquery.Open "select pa05 as Name1, pa06 as Name2, pa07 as Name3, nvl(na03, na04) as NationName, na01, pa26 as CustomerNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as CustomerName,PA08 from patent, nation, customer where pa09 = na01 (+) and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and pa01 = '" & Text10 & "' and pa02 = '" & Text5 & "' and pa03 = '" & Text7 & "' and pa04 = '" & Text9 & "' union " & _
                 "select tm05 as Name1, tm06 as Name2, tm07 as Name3, nvl(na03, na04) as NationName, na01, tm23 as CustomerNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as CustomerName,'' from trademark, nation, customer where tm10 = na01 (+) and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and tm01 = '" & Text10 & "' and tm02 = '" & Text5 & "' and tm03 = '" & Text7 & "' and tm04 = '" & Text9 & "' and tm01 <> 'TF' union " & _
                 "select tm05 as Name1, tm06 as Name2, tm07 as Name3, nvl(na03, na04) as NationName, na01, tm23 as CustomerNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as CustomerName,'' from trademark, nation, customer where tm10 = na01 (+) and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and tm01 = '" & Text10 & "' and tm02 = '" & Text5 & Text7 & "' and tm03 = '" & Text9 & "' and tm04 = '" & Text12 & "' and tm01 = 'TF' union " & _
                 "select lc05 as Name1, lc06 as Name2, lc07 as Name3, nvl(na03, na04) as NationName, na01, lc11 as CustomerNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as CustomerName,'' from lawcase, nation, customer where lc15 = na01 (+) and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and lc01 = '" & Text10 & "' and lc02 = '" & Text5 & "' and lc03 = '" & Text7 & "' and lc04 = '" & Text9 & "' union " & _
                 "select hc06 as Name1, '' as Name2, '' as Name3, nvl(na03, na04) as NationName, '000' as na01, hc05 as CustomerNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as CustomerName,'' from hirecase, nation, customer where  '000'=na01(+) and substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and hc01 = '" & Text10 & "' and hc02 = '" & Text5 & "' and hc03 = '" & Text7 & "' and hc04 = '" & Text9 & "' union " & _
                 "select sp05 as Name1, sp06 as Name2, sp07 as Name3, nvl(na03, na04) as NationName, na01, sp08 as CustomerNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as CustomerName,'' from servicepractice, nation, customer where sp09 = na01 (+) and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and sp01 = '" & Text10 & "' and sp02 = '" & Text5 & "' and sp03 = '" & Text7 & "' and sp04 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Add By Cheng 2003/03/04
   m_PA08 = ""
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("Name1").Value) = False Then
         Combo1 = "中--" & adoquery.Fields("Name1").Value
         Combo1.AddItem "中--" & adoquery.Fields("Name1").Value
      End If
      If IsNull(adoquery.Fields("Name2").Value) = False Then
         Combo1.AddItem "英--" & adoquery.Fields("Name2").Value
      End If
      If IsNull(adoquery.Fields("Name3").Value) = False Then
         Combo1.AddItem "日--" & adoquery.Fields("Name3").Value
      End If
      Text4 = adoquery.Fields("na01").Value
      If IsNull(adoquery.Fields("NationName").Value) = False Then
         Text1 = adoquery.Fields("NationName").Value
      Else
         Text1 = MsgText(601)
      End If
      Text11 = "" & adoquery.Fields("CustomerNo").Value
      If IsNull(adoquery.Fields("CustomerName").Value) = False Then
         Text3 = adoquery.Fields("CustomerName").Value
      Else
         Text3 = MsgText(601)
      End If
        'Add By Cheng 2003/03/04
        '取得專利種類
        If "" & adoquery.Fields("PA08").Value <> "" Then m_PA08 = adoquery.Fields("PA08").Value
   Else
      Text4 = MsgText(601)
      Text1 = MsgText(601)
      Text11 = MsgText(601)
      Text3 = MsgText(601)
   End If
   adoquery.Close
End Sub

'*************************************************
'  計算並顯示盈虧
'
'*************************************************
Public Sub Calculate()
Dim intCounter As Integer

   intCounter = 0
   '2010/3/26 MODIFY BY SONIA 剔除作廢帳單
   'strSQL1 = ""
   strSQL1 = " and a1507 is null"
   '2010/3/26 END
   
   'modify by sonia 改用IF,因為CFP只有EPC(221)才要母子案一起計算,否則CFP案在抓ACC021都會很久
   'Select Case Text10
   '   Case "TF"
      If Text10 = "TF" Then
         '2009/6/25 MODIFY BY SONIA母案子案一起計算
         'strSQL1 = strSQL1 & " and axf03 = '" & Text10 & Text5 & Text7 & Text9 & Text12 & "'"
         '2010/3/26 modify by sonia 領土延伸不可一起算
         'strSQL1 = strSQL1 & " and axf03 = '" & Text10 & Text5 & "0000'"
         strSQL1 = strSQL1 & " and instr(axf03,'" & Text10 & Text5 & Text7 & "')=1 "
         'add by sonia 2021/3/26
         strSQL2 = " instr(ax214,'" & Text10 & Text5 & Text7 & "')=1 "
         StrSQL3 = " instr(a1p17,'" & Text10 & Text5 & Text7 & "')=1 "
         StrSQL4 = " and instr(axg03,'" & Text10 & Text5 & Text7 & "')=1 "
         'end 2021/3/26
      '2009/6/25 ADD BY SONIA EPC母案子案一起計算
      'modify by sonia 改用IF,因為CFP只有EPC(221)才要母子案一起計算,否則CFP案在抓ACC021都會很久
      'Case "CFP"
      ElseIf Text10 = "CFP" And Text4 = "221" Then
         strSQL1 = strSQL1 & " and instr(axf03,'" & Text10 & Text5 & Text7 & "')=1 "
         'add by sonia 2021/3/26
         strSQL2 = " instr(ax214,'" & Text10 & Text5 & Text7 & "')=1 "
         StrSQL3 = " instr(a1p17,'" & Text10 & Text5 & Text7 & "')=1 "
         StrSQL4 = " and instr(axg03,'" & Text10 & Text5 & Text7 & "')=1 "
         'end 2021/3/26
      '2009/6/25 END
      'modify by sonia 改用IF,因為CFP只有EPC(221)才要母子案一起計算,否則CFP案在抓ACC021都會很久
      'Case Else
      Else
         strSQL1 = strSQL1 & " and axf03 = '" & Text10 & Text5 & Text7 & Text9 & "'"
         'add by sonia 2021/3/26
         strSQL2 = " ax214 = '" & Text10 & Text5 & Text7 & Text9 & "' "
         StrSQL3 = " a1p17 = '" & Text10 & Text5 & Text7 & Text9 & "' "
         StrSQL4 = " and axg03 = '" & Text10 & Text5 & Text7 & Text9 & "'"
         'end 2021/3/26
   'modify by sonia 改用IF,因為CFP只有EPC(221)才要母子案一起計算,否則CFP案在抓ACC021都會很久
   'End Select
      End If
      
   adoquery.CursorLocation = adUseClient
   'modify by sonia 2021/5/5 原只抓舊系統帳單axf02 = '000000000',但舊系統之支出公簽證就抓不到CFT-007227,故改抓92/01/30以前傳票借方規費
   'adoquery.Open "select sum(nvl(a1520, 0)) from acc151, acc150 where axf01 = a1501 and axf02 = '000000000'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
   'If adoquery.RecordCount <> 0 Then
   '   If IsNull(adoquery.Fields(0).Value) Then
   '      Text8 = "0"
   '   Else
   '      Text8 = adoquery.Fields(0).Value
   '   End If
   'Else
   '   Text8 = "0"
   'End If
   adoquery.Open "SELECT ax201,ax202,ax203,ax206-ax207 AMT,ax212 FROM acc021,acc020,(SELECT * FROM acc1p0 WHERE " & StrSQL3 & " AND a1p02<>'L') WHERE " & strSQL2 & " AND ax201=a1p01(+) AND ax202=a1p22(+) and ax205=a1p05(+) and ax214=a1p17(+) and ax206=a1p07(+) and ax207=a1p08(+) AND a1p04 IS NULL AND ax205 LIKE '2201%' and ax206>0 and ax201=a0201(+) and ax202=a0202(+) and a0205<920130", adoTaie, adOpenStatic, adLockReadOnly
   Text8 = 0
      Do While adoquery.EOF = False
         If InStr("" & adoquery.Fields("ax212"), "結餘") = 0 Then
            Text8 = Text8 + adoquery.Fields("AMT").Value
         End If
         adoquery.MoveNext
      Loop
   adoquery.Close
   
   If Adodc1.Recordset.State = adStateOpen Then
      Text2 = ""
      Do While Adodc1.Recordset.EOF = False
            '收入
            Text2 = Val(Text2) + Val(Adodc1.Recordset.Fields("RecAmount").Value)
            '扣點數
            If IsNull(Adodc1.Recordset.Fields("cp18").Value) = False And Val(Adodc1.Recordset.Fields("RecAmount").Value) > 0 Then
               Text2 = Val(Text2) - (Val(Adodc1.Recordset.Fields("cp18").Value) * 1000)
            End If

'2010/3/18 CANCEL BY SONIA 因CFT-12231有不同幣別帳單,QueryTable取消帳單幣別A1505及付款外幣金額FpayAmount欄,故帳單移至下面另外做
'            If Adodc1.Recordset.Fields("PayAmount").Value <> 0 And IsNull(Adodc1.Recordset.Fields("PayAmount").Value) = False Then
'               Text2 = Val(Text2) - Val(Adodc1.Recordset.Fields("PayAmount").Value)
'            Else
'               If adoaccsum.State = adStateOpen Then
'                  adoaccsum.Close
'               End If
'               adoaccsum.CursorLocation = adUseClient
'               '2009/11/24 加註by sonia 所有未付帳單之損益都以最新的預估結匯匯率計算
'               adoaccsum.Open "select a2103 from acc210 where a2102 = '" & Adodc1.Recordset.Fields("a1505").Value & "' and a2101 = (select max(a2101) from acc210 where a2102 = '" & Adodc1.Recordset.Fields("a1505").Value & "' and a2101 <= " & Val(ACDate(ServerDate)) & ")", adoTaie, adOpenStatic, adLockReadOnly
'               If adoaccsum.RecordCount <> 0 Then
'                  If IsNull(Adodc1.Recordset.Fields("FpayAmount").Value) = False Then
'                     Text2 = Val(Text2) - (Val(Adodc1.Recordset.Fields("FpayAmount").Value) * Val(adoaccsum.Fields("a2103").Value))
'                  End If
'               Else
'                  If IsNull(Adodc1.Recordset.Fields("FpayAmount").Value) = False Then
'                     Text2 = Val(Text2) - Val(Adodc1.Recordset.Fields("FpayAmount").Value)
'                  End If
'               End If
'               adoaccsum.Close
'            End If
'2010/3/18 END
         Adodc1.Recordset.MoveNext
      Loop
   
      '2010/3/18 ADD BY SONIA
      '扣除帳單
      adoquery.CursorLocation = adUseClient
      '2010/3/19 MODIFY BY SONIA 直接抓ACC151帳單,不必分國內收據或國外請款單,且未付不抓AXF15
      'adoquery.Open "select nvl(cp05 - 19110000, 0) as cp05, a1505, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0)) as PayAmount, sum(nvl(axf04, 0)) as FpayAmount, CP09 from caseprogress, acc151, acc150,ACC190,ACC1G0 " & _
                  "where cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) and (substr(cp60, 1, 1) = 'E' or cp60 is null)" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, a1505, cp61, CP09 union " & _
                  "select nvl(cp05 - 19110000, 0) as cp05, a1505, cp61, sum(decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0)) as PayAmount, sum(decode(a1507, null, nvl(axf04, 0), 0)) as FpayAmount, CP09 from caseprogress, acc151, acc150,ACC190,ACC1G0 " & _
                  "where cp09 = axf02 (+) and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) and substr(cp60, 1, 1) = 'X'" & strSql & " AND ((nvl(cp16,0) - nvl(cp77,0))>0 OR (decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP16-CP77) / 1000)))>0 OR (decode(a1507, null, nvl(NVL(NVL(AXF04*decode(A1906,0,null,A1906),AXF04*A1G03),axf15), 0), 0))>0 OR (nvl(axf04, 0))>0) group by cp05, a1505, cp61, CP09", adoTaie, adOpenStatic, adLockReadOnly
      '2010/4/13 MODIFY BY SONIA 剔除舊系統帳單CFT-007227
      'modify by sonia 2021/4/8 P-097213之抵帳單ACC160要從帳單金額加回來
      'adoquery.Open "select a1505,AXF04,A1906,A1G03,NVL(AXF04*decode(A1906,0,null,A1906),NVL(AXF04*decode(A1G03,0,null,A1G03),0)) as PayAmount,nvl(axf04, 0) as FpayAmount, AXF02 from acc151, acc150,ACC190,ACC1G0 " & _
                    "where axf02<>'000000000' and axf01 = a1501 (+) AND AXF01=A1902(+) AND A1512=A1G01(+) " & strSQL1 & " AND nvl(axf04, 0)>0 ", adoTaie, adOpenStatic, adLockReadOnly
      'modify by sonia 2021/5/6 CFP-025854之V10600009於Z10600021抵帳抓抵帳單匯率計算
      'adoquery.Open "select a1505,AXF04,a.A1906,A1G03,NVL(AXF04*decode(a.A1906,0,null,a.A1906),NVL(AXF04*decode(A1G03,0,null,A1G03),0)) as PayAmount,nvl(axf04, 0) as FpayAmount, AXF02 from acc151, acc150,ACC190 a,ACC1G0 " & _
                    "where axf02<>'000000000' and axf01 = a1501 (+) AND AXF01=a.A1902(+) AND A1512=A1G01(+) " & strSQL1 & " AND nvl(axf04, 0)>0 " & _
                    "union select a1605,AXG04,b.A1906,0,nvl(AXg04*b.A1906,0)*-1 as PayAmount,nvl(axg04, 0)*-1 as FpayAmount, AXg02 from acc160,acc161,acc190 b where axg01 = a1601 (+) and axg01=b.a1902(+) " & StrSQL4 & " AND nvl(axg04, 0)>0", adoTaie, adOpenStatic, adLockReadOnly
      'modify by sonia 2022/7/26 T-124366之抵帳單V09000001要從帳單金額加回來但因and a1607=d.a1i03少寫(+)所以沒抓到
      adoquery.Open "select a1505,AXF04,a.A1906,A1G03,NVL(AXF04*decode(a.A1906,0,null,a.A1906),NVL(AXF04*decode(A1G03,0,null,A1G03),0)) as PayAmount,nvl(axf04, 0) as FpayAmount, AXF02 from acc151, acc150,ACC190 a,ACC1G0 " & _
                    "where axf02<>'000000000' and axf01 = a1501 (+) AND AXF01=a.A1902(+) AND A1512=A1G01(+) " & strSQL1 & " AND nvl(axf04, 0)>0 " & _
                    "union select a1605,AXG04,nvl(A1906,a1g03),0,nvl(AXg04*nvl(A1906,a1g03),0)*-1 as PayAmount,nvl(axg04, 0)*-1 as FpayAmount, AXg02 from acc160,acc161,acc190,acc1i0 c,acc1i0 d,acc1g0 " & _
                    "where axg01 = a1601 (+) and axg01=a1902(+) " & StrSQL4 & " AND nvl(axg04, 0)>0 and a1607=c.a1i03(+) and a1605=c.a1i05(+) and a1607=d.a1i03(+) and nvl(c.a1i01,d.a1i01)=a1g01(+)", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoquery.EOF = False
         '己付不再算
         If adoquery.Fields("PayAmount").Value <> 0 And IsNull(adoquery.Fields("PayAmount").Value) = False Then
            Text2 = Val(Text2) - Format(Val(adoquery.Fields("PayAmount").Value), FAmount)
         '未付帳單以最新的預估結匯匯率計算
         Else
            If adoaccsum.State = adStateOpen Then
               adoaccsum.Close
            End If
            adoaccsum.CursorLocation = adUseClient
            adoaccsum.Open "select a2103 from acc210 where a2102 = '" & adoquery.Fields("a1505").Value & "' and a2101 = (select max(a2101) from acc210 where a2102 = '" & adoquery.Fields("a1505").Value & "' and a2101 <= " & Val(ACDate(strSrvDate(1))) & ")", adoTaie, adOpenStatic, adLockReadOnly
            If adoaccsum.RecordCount <> 0 Then
               If IsNull(adoquery.Fields("FpayAmount").Value) = False Then
                  Text2 = Val(Text2) - Format((Val(adoquery.Fields("FpayAmount").Value) * Val(adoaccsum.Fields("a2103").Value)), FAmount)
               End If
            Else
               If IsNull(adoquery.Fields("FpayAmount").Value) = False Then
                  Text2 = Val(Text2) - Format(Val(adoquery.Fields("FpayAmount").Value), FAmount)
               End If
            End If
            adoaccsum.Close
         End If
         adoquery.MoveNext
      Loop
      '2010/3/18 END

      'add by sonia 2021/3/26 加財務其他支出Text13(直接由傳票輸入之規費),判斷AX212有無結餘二字若放在語法中當沒資料時會有點慢instr(ax212,'結餘')=0
      'modify by sonia 2021/5/5 +傳票日期>=920130條件,否則會抓到舊系統之帳單CFT-007227
      adoaccsum.CursorLocation = adUseClient
      adoaccsum.Open "SELECT ax201,ax202,ax203,ax206-ax207 AMT,ax212 FROM acc021,acc020,(SELECT * FROM acc1p0 WHERE " & StrSQL3 & " AND a1p02<>'L') WHERE " & strSQL2 & " AND ax201=a1p01(+) AND ax202=a1p22(+) and ax205=a1p05(+) and ax214=a1p17(+) and ax206=a1p07(+) and ax207=a1p08(+) AND a1p04 IS NULL AND ax205 LIKE '2201%' and ax201=a0201(+) and ax202=a0202(+) and a0205>=920130", adoTaie, adOpenStatic, adLockReadOnly
      Text13 = 0
      Do While adoaccsum.EOF = False
         If InStr("" & adoaccsum.Fields("ax212"), "結餘") = 0 Then
            Text13 = Text13 + adoaccsum.Fields("AMT").Value
         End If
         adoaccsum.MoveNext
      Loop
      adoaccsum.Close
      'end 2021/3/26
      
      'modify by sonia 2021/3/26 再減財務其他支出Text13
      'Text2 = Val(Text2) - Val(Text6) - Val(Text8)
      Text2 = Val(Text2) - Val(Text6) - Val(Text8) - Val(Text13)
      'end 2021/3/26
      Text2 = Format(Text2, FAmount)
      If Adodc1.Recordset.RecordCount <> 0 Then
         Adodc1.Recordset.MoveFirst
      End If
   End If
End Sub
