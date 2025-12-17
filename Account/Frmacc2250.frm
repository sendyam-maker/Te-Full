VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2250 
   AutoRedraw      =   -1  'True
   Caption         =   "其他結匯查詢"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   8760
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
      Height          =   330
      Left            =   6600
      MaxLength       =   50
      TabIndex        =   4
      Top             =   480
      Width           =   1920
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
      Height          =   330
      Left            =   1305
      TabIndex        =   0
      Top             =   90
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
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
      Left            =   1305
      TabIndex        =   5
      Top             =   870
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2250.frx":0000
      Height          =   3600
      Left            =   225
      TabIndex        =   6
      Top             =   1350
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6350
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "a1702"
         Caption         =   "單據號碼"
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
         DataField       =   "a1709"
         Caption         =   "付款單號"
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
         DataField       =   "a1b01"
         Caption         =   "匯票號碼"
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
      BeginProperty Column03 
         DataField       =   "a1b03"
         Caption         =   "結匯日期"
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
      BeginProperty Column04 
         DataField       =   "a1b06"
         Caption         =   "付款方式"
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
         DataField       =   "a1703"
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
      BeginProperty Column06 
         DataField       =   "a1704"
         Caption         =   "結匯金額"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "a1705"
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2505.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   210
      Top             =   1245
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   1305
      TabIndex        =   2
      Top             =   480
      Width           =   1500
      _ExtentX        =   2646
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   3105
      TabIndex        =   3
      Top             =   480
      Width           =   1500
      _ExtentX        =   2646
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
   Begin MSForms.TextBox Text1 
      Height          =   330
      Left            =   5040
      TabIndex        =   1
      Top             =   90
      Width           =   3495
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      Size            =   "6165;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
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
      Height          =   315
      Left            =   5040
      TabIndex        =   12
      Top             =   518
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
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
      Left            =   240
      TabIndex        =   11
      Top             =   128
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "往來日期"
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
      Left            =   240
      TabIndex        =   10
      Top             =   518
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2865
      TabIndex        =   9
      Top             =   518
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "其他對象名稱"
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
      Left            =   3585
      TabIndex        =   8
      Top             =   128
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -12
      Top             =   4776
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "外幣金額"
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
      Left            =   240
      TabIndex        =   7
      Top             =   908
      Width           =   1215
   End
End
Attribute VB_Name = "Frmacc2250"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset
Dim strSql As String

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
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
'   Me.Height = 5400
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath2)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 8850, 5500, strBackPicPath2
   'end 2021/12/09
   
   OpenTable
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   StatusView MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2250 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'Modified by Lydia 2021/12/09 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Text1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   Select Case Len(Text2)
      Case 6
         Text2 = Text2 & "000"
      Case 8
         Text2 = Text2 & "0"
   End Select
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   '2013/8/20 modify by sonia 不必抓資料故加入a1801='1'加快速度
   adoadodc1.Open "select a1b01, a1b03, decode(a1b06, '1', '票匯', '2', '電匯') as a1b06, sum(a1p07) as Amount from acc1b0, acc180, acc1p0 where a1801='1' and a1b02 = a1803 and a1b03 = a1802 and a1b01||a1b02 = a1p04 and a1p01 = '1' and a1p02 = 'I' and a1b01 = 'Z' group by a1b01, a1b03, a1b06", adoTaie, adOpenStatic, adLockReadOnly
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
Dim strSql As String
Dim strHaving As String

On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   
   If Text1 <> MsgText(601) Then
      '2013/8/20 modify by sonia 改以acc170為主table
      'strSql = " and a1810 like '" & Text1 & "%" & "'"
      strSql = " and a1717 like '" & Text1 & "%" & "'"
   End If
   If Text3 <> MsgText(601) Then
      '2013/8/20 modify by sonia 改以acc170為主table
      'strSql = strSql & " and a1904 = " & Val(Text3) & ""
      strSql = strSql & " and a1704 = " & Val(Text3) & ""
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      '2013/8/20 modify by sonia 改以acc170為主table
      'strSql = strSql & " and a1802 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSql = strSql & " and a1708 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      '2013/8/20 modify by sonia 改以acc170為主table
      'strSql = strSql & " and a1802 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSql = strSql & " and a1708 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text2 <> MsgText(601) Then
      '2013/8/20 modify by sonia 改以acc170為主table
      'strSql = strSql & " and a1803 = '" & Text2 & "'"
      strSql = strSql & " and a1705 = '" & Text2 & "'"
   End If
   'Add by Amy 2013/08/19 +代理人D/N No.查詢
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a1706='" & Text4 & "' "
   End If
   
   'adoadodc1.Open "select a1b01, a1b03, a1b06, sum(a1p07) as Amount from (select a1b01, a1b02, a1b03, decode(a1b06, '1', '票匯', '2', '電匯', '3', '旅行支票', '4', '現金') as a1b06, a1810 from acc1b0, acc180 where a1b02 = a1803 and a1b03 = a1802" & strSQL & " group by a1b01, a1b02, a1b03, a1b06, a1810) new, acc1p0 where a1b01||a1b02 = a1p04 and a1p01 = '1' and a1p02 = 'I'" & strSQL & " group by a1b01, a1b03, a1b06" & strHaving, adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2013/08/19 +代理人D/N No.查詢
   'adoadodc1.Open "select a1901, a1902, a1903, a1b01, a1b03, decode(a1b06, '1', '票匯', '2', '電匯', '3', '旅行支票', '4', '現金') as a1b06, a1904, a1803 from acc190, acc180, acc1b0 where a1901 = a1801 and a1908 = a1b01 (+)" & strSql & " order by a1902 asc, a1901 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2013/8/20 modify by sonia 改以acc170為主table且只抓其他結匯故加入substr(a1702,1,1)='B'
   'strExc(0) = "select a1901, a1902, a1903, a1b01, a1b03, decode(a1b06, '1', '票匯', '2', '電匯', '3', '旅行支票', '4', '現金') as a1b06, a1904, a1803,a1706 from acc190, acc180, acc1b0,acc170 where a1901 = a1801 and a1908 = a1b01 (+) and a1902=a1702 and a1803 = a1705(+) " & strSql & " order by a1902 asc, a1901 asc"
   strExc(0) = "select a1702, a1709, a1703, a1b01, a1b03, decode(a1b06, '1', '票匯', '2', '電匯', '3', '旅行支票', '4', '現金', '5', '商務卡', '6', '其他') as a1b06, a1704, a1703, a1706, a1705||' '||a1717 as a1705 from acc190, acc180, acc1b0,acc170 where a1709 = a1801(+) and a1709=a1901(+) and a1702=a1902(+) and a1908 = a1b01 (+) and substr(a1702,1,1)='B'" & strSql & " order by a1702 asc, a1709 asc"
   adoadodc1.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
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
            QueryTable
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
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   '2013/8/20 add by sonia
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   '2013/8/20 end
   FormCheck = False
End Function

'Add by Amy 2013/08/19 +代理人D/N No查詢
Private Sub Text4_GotFocus()
    TextInverse Text4
    CloseIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2013/08/19
