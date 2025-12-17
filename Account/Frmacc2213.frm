VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2213 
   AutoRedraw      =   -1  'True
   Caption         =   "帳單資料查詢"
   ClientHeight    =   4464
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8736
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4464
   ScaleWidth      =   8736
   Begin VB.CommandButton Command2 
      Caption         =   "電子檔"
      Enabled         =   0   'False
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
      Left            =   2520
      TabIndex        =   17
      Top             =   271
      Width           =   855
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
      Left            =   1140
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1365
   End
   Begin VB.TextBox Text1 
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
      Height          =   330
      Left            =   4440
      MaxLength       =   9
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
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
      Height          =   330
      Left            =   1740
      TabIndex        =   2
      Top             =   600
      Width           =   3915
   End
   Begin VB.TextBox Text5 
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
      Height          =   330
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   4
      Top             =   960
      Width           =   1440
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
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
      Height          =   330
      Left            =   4230
      TabIndex        =   5
      Top             =   960
      Width           =   1425
   End
   Begin VB.TextBox Text14 
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
      Left            =   3360
      TabIndex        =   6
      Top             =   4050
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   6960
      TabIndex        =   7
      Top             =   960
      Width           =   1455
      _ExtentX        =   2561
      _ExtentY        =   593
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   6960
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _ExtentX        =   2561
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
      Left            =   240
      Top             =   1680
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGrid1 
      Height          =   2055
      Left            =   240
      TabIndex        =   20
      Top             =   1890
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   3620
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSForms.TextBox Text8 
      Height          =   480
      Left            =   1140
      TabIndex        =   19
      Top             =   1320
      Width           =   7275
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "12832;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   5670
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   240
      Width           =   2760
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4868;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "帳單編號"
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
      Left            =   180
      TabIndex        =   16
      Top             =   270
      Width           =   1155
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   270
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人D/N No."
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
      Left            =   180
      TabIndex        =   14
      Top             =   630
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳單日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6000
      TabIndex        =   13
      Top             =   645
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   180
      TabIndex        =   12
      Top             =   1012
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳單金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3255
      TabIndex        =   11
      Top             =   1005
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "作廢日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6000
      TabIndex        =   10
      Top             =   1005
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   180
      TabIndex        =   9
      Top             =   1320
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   4050
      Width           =   855
   End
End
Attribute VB_Name = "Frmacc2213"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text3、Text8
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc150 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Dim str1stChar As String * 1 'Add by Morgan 2004/9/2 單號辨識字元
Dim m_AttachPath As String, m_FileName As String 'Added by Morgan 2019/5/30 帳單電子檔
'Add by Amy 2025/01/14 傳入案號/欄位名稱/大小
Public m_CaseNo As String
Dim i As Integer, strFieldN(), intWidth()

'Added by Morgan 2019/5/30
Private Sub Command2_Click()
   Dim stSaveFileName As String
   Dim hLocalFile As Long
   
   If PUB_GetAttachFile_Invoice(Text2, m_FileName, m_AttachPath, stSaveFileName) = True Then
      ShellExecute hLocalFile, "open", m_AttachPath & "\" & stSaveFileName, vbNullString, vbNullString, 1
   End If
End Sub
'Added by Morgan 2019/5/30
Private Sub SetFileButton()
   Dim stSQL As String, intQ As Integer
   
   stSQL = "select ayf02 from acc152 where ayf01='" & Text2 & "'"
   intQ = 1
   Set adoaccsum = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      m_FileName = adoaccsum(0)
      Command2.Enabled = True
      
      m_AttachPath = App.path & "\" & strUserNum
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      KillTemp
   End If
   adoaccsum.Close
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
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
'   Me.Height = 4850
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
   PUB_InitForm Me, 8850, 4900, strBackPicPath1
   'end 2021/12/09
   
   
   'Add by Morgan 2004/9/2
   str1stChar = Left(strItemNo, 1)
   If str1stChar = "V" Then
      Label1 = "抵" & Label1
      Label4 = "抵" & Label4
      Label6 = "抵" & Label6
      Me.Caption = "抵" & Me.Caption
      'Mark by Amy 2025/01/14 換Grid改至
      'DataGrid1.Columns(2).Caption = "抵" & DataGrid1.Columns(2).Caption
   End If
   'end 2025/01/14
   OpenTable
   SumShow
   SetFileButton 'Added by Morgan 2019/5/30
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_CaseNo = "" 'Add by Amy 2025/01/14
   strItemNo = ""
   tool3_enabled
   Select Case strFormLink
      Case "Frmacc2210"
         Frmacc2210.Enabled = True
      Case "Frmacc2220"
         Frmacc2220.Enabled = True
   End Select
   Set Frmacc2213 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim strSystemKind As String
Dim strSql As String 'Add by Amy 2025/01/14

On Error GoTo Checking
   adoacc150.CursorLocation = adUseClient
   'Modify by Morgan 2004/9/2 國外抵帳單編號抓160,161
   If str1stChar = "V" Then
      adoacc150.Open "select * from acc160 where a1601 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
      adoacc150.Open "select * from acc150 where a1501 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End If
   
   FormShow
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2004/9/2 國外抵帳單編號抓160,161
   'Modify by Amy 2025/01/14 改為MSHFlexGrid,故調整顯示欄位
   If str1stChar = "V" Then
      'adoadodc1.Open "select axg01 axf01, axg02 axf02, axg03 axf03, axg04 axf04, axg12 axf12, axg13 axf13 from acc161 where axg01 = '" & strItemNo & "' order by axg02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      strSql = "select '',axg02 axf02,axg03 axf03,axg04 axf04,axg12 axf12,axg13 axf13,axg01 axf01 from acc161 where axg01 = '" & strItemNo & "' order by axg02 asc"
   Else
      'adoadodc1.Open "select * from acc151 where axf01 = '" & strItemNo & "' order by axf02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      strSql = "select '',axf02,axf03,axf04,axf12,axf13,axf01,axf05,axf08,axf06,axf07,axf11,axf09,axf10,axf14,axf15,axf16 from acc151 where axf01 = '" & strItemNo & "' order by axf02 asc"
   End If
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   'Modify by Amy 2025/01/14 換Grid 設顏色
   Set Adodc1.Recordset = adoadodc1
   Set MSHFGrid1.Recordset = adoadodc1
   SetGridWidth
   Call SetGridColor(0)
   'end Add by Amy 2025/01/14
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
   'Modify by Morgan 2004/9/2 國外抵帳單編號抓160,161
   If str1stChar = "V" Then
      If IsNull(adoacc150.Fields("a1603").Value) Then
         Text1 = MsgText(601)
      Else
         Text1 = adoacc150.Fields("a1603").Value
      End If
      If Len(Text1) = 6 Then
         Text3 = FagentQuery(AfterZero(Text1), 2)
      Else
         Text3 = FagentQuery(Text1, 2)
         'add by sonia 2015/4/20
         If Text3 = "" Then
            Text3 = FagentQuery(Text1, 1)
         End If
         If Text3 = "" Then
            Text3 = FagentQuery(Text1, 3)
         End If
         '2015/4/20 end
      End If
      If IsNull(adoacc150.Fields("a1604").Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = adoacc150.Fields("a1604").Value
      End If
      MaskEdBox1.Mask = MsgText(601)
      If IsNull(adoacc150.Fields("a1602").Value) Then
         MaskEdBox1.Text = MsgText(601)
      Else
         MaskEdBox1.Text = CFDate(adoacc150.Fields("a1602").Value)
      End If
      MaskEdBox1.Mask = DFormat
      If IsNull(adoacc150.Fields("a1605").Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = adoacc150.Fields("a1605").Value
      End If
      If IsNull(adoacc150.Fields("a1606").Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = Format(adoacc150.Fields("a1606").Value, FAmount)
      End If
      MaskEdBox2.Mask = MsgText(601)
      If IsNull(adoacc150.Fields("a1608").Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = adoacc150.Fields("a1608").Value
      End If
      
   Else
      If IsNull(adoacc150.Fields("a1503").Value) Then
         Text1 = MsgText(601)
      Else
         Text1 = adoacc150.Fields("a1503").Value
      End If
      If Len(Text1) = 6 Then
         Text3 = FagentQuery(AfterZero(Text1), 2)
      Else
         Text3 = FagentQuery(Text1, 2)
         'add by sonia 2015/4/20
         If Text3 = "" Then
            Text3 = FagentQuery(Text1, 1)
         End If
         If Text3 = "" Then
            Text3 = FagentQuery(Text1, 3)
         End If
         '2015/4/20 end
      End If
      If IsNull(adoacc150.Fields("a1504").Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = adoacc150.Fields("a1504").Value
      End If
      MaskEdBox1.Mask = MsgText(601)
      If IsNull(adoacc150.Fields("a1502").Value) Then
         MaskEdBox1.Text = MsgText(601)
      Else
         MaskEdBox1.Text = CFDate(adoacc150.Fields("a1502").Value)
      End If
      MaskEdBox1.Mask = DFormat
      If IsNull(adoacc150.Fields("a1505").Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = adoacc150.Fields("a1505").Value
      End If
      If IsNull(adoacc150.Fields("a1506").Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = Format(adoacc150.Fields("a1506").Value, FAmount)
      End If
      MaskEdBox2.Mask = MsgText(601)
      If IsNull(adoacc150.Fields("a1507").Value) Then
         MaskEdBox2.Text = MsgText(601)
      Else
         MaskEdBox2.Text = CFDate(adoacc150.Fields("a1507").Value)
      End If
      MaskEdBox2.Mask = DFormat
      If IsNull(adoacc150.Fields("a1509").Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = adoacc150.Fields("a1509").Value
      End If
         
   End If
End Sub

'*************************************************
'  合計顯示
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'Modify by Morgan 2004/9/2 國外抵帳單編號抓160,161
   If str1stChar = "V" Then
      adoaccsum.Open "select sum(axg04) from acc161 where axg01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoaccsum.Open "select sum(axf04) from acc151 where axf01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   End If
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text14 = MsgText(601)
      Else
         Text14 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
   Else
      Text14 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'Add by Amy 2025/01/14 原DataGird
Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = LBound(strFieldN) To UBound(strFieldN)
        If UCase(strFieldN(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function

Private Sub SetGridWidth()
   Dim stTP As String
   ReDim strFieldN(4)
   ReDim intWidth(4)
   
   'Add by Amy 2025/01/14 改Grid
   If str1stChar = "V" Then
      stTP = "抵"
   End If
       
   strFieldN = Array(" ", "總收文號", "本所案號", stTP & "帳單金額", "案件名稱", "收據抬頭")
   intWidth = Array(200, 1320, 1400, 1390, 3400, 4356)
   MSHFGrid1.Cols = UBound(strFieldN) + 1

   MSHFGrid1.row = 0
   For i = LBound(strFieldN) To UBound(strFieldN)
      MSHFGrid1.col = i
      MSHFGrid1.ColWidth(i) = intWidth(i)
      MSHFGrid1.Text = strFieldN(i)
      
      MSHFGrid1.CellFontName = "新細明體-ExtB"
      MSHFGrid1.CellFontSize = 11
      MSHFGrid1.CellFontBold = True
      MSHFGrid1.CellBackColor = &HE0E0E0
      MSHFGrid1.CellAlignment = flexAlignLeftCenter
   Next i
   
End Sub

Private Sub SetGridColor(intChoose As Integer, Optional ByVal intRow As Integer)
   Dim j As Integer, intS As Integer, intE As Integer
   If intRow = 0 Then
      intS = 1: intE = MSHFGrid1.Rows - 1
   Else
      intS = intRow: intE = intRow
   End If
   
   If intChoose = 1 Then
      MSHFGrid1.row = intRow
      For i = 1 To MSHFGrid1.Cols - 1
         '第1欄 Or 前畫面由Frmacc2220 且有輸本所案號且與目前查到資料不同時,不變色
         If i = GetValue(" ") _
           Or (strFormLink = "Frmacc2220" And m_CaseNo <> MsgText(601) And m_CaseNo <> MSHFGrid1.TextMatrix(intRow, i) And i = GetValue("本所案號")) Then
            '不需改變,維持原設定
         Else
            MSHFGrid1.col = i
            MSHFGrid1.CellBackColor = &HFFC0C0 '整列底 藍色
         End If
      Next i
   Else
      For i = intS To intE
         MSHFGrid1.row = i
         For j = LBound(strFieldN) To UBound(strFieldN)
            MSHFGrid1.col = j
            If j = GetValue(" ") Then
               If intChoose = 2 And MSHFGrid1.TextMatrix(i, j) = "V" Then
                  MSHFGrid1.TextMatrix(i, j) = ""
               End If
               MSHFGrid1.CellBackColor = &HE0E0E0
            ElseIf strFormLink = "Frmacc2220" And m_CaseNo <> MsgText(601) And m_CaseNo <> MSHFGrid1.TextMatrix(i, j) And j = GetValue("本所案號") Then
               '由frmacc2210 輸案號查進入此畫面,案號[不同]時顯示黃色 ex:T-246337 之U11309756 會有1筆 T-246338資料
               MSHFGrid1.CellBackColor = vbYellow
            Else
               MSHFGrid1.CellBackColor = QBColor(15) '設回
            End If
            If j = GetValue("帳單金額") Then
               MSHFGrid1.CellAlignment = flexAlignRightCenter
            End If
         Next j
      Next i
   End If
End Sub

Private Sub MSHFGrid1_Click()
   Dim intR As Integer
   
   MSHFGrid1.row = MSHFGrid1.MouseRow
   MSHFGrid1.col = MSHFGrid1.MouseCol
   intR = MSHFGrid1.row
   If MSHFGrid1.row <> 0 Then
      Call SetGridColor(2) '先還原顏色
      Call SetGridColor(1, intR)
      MSHFGrid1.TextMatrix(intR, 0) = "V"
   End If
End Sub
'end 2025/01/14
