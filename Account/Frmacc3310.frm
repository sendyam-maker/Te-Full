VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc3310 
   AutoRedraw      =   -1  'True
   Caption         =   "票據兌現處理"
   ClientHeight    =   5240
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5240
   ScaleWidth      =   8740
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmacc3310.frx":0000
      Left            =   3405
      List            =   "Frmacc3310.frx":0002
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   480
      Width           =   2375
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
      Left            =   3945
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   612
   End
   Begin VB.TextBox Text7 
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
      Height          =   315
      Left            =   3405
      MaxLength       =   8
      TabIndex        =   19
      Top             =   830
      Width           =   1572
   End
   Begin VB.TextBox Text6 
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
      Height          =   315
      Left            =   2448
      TabIndex        =   18
      Top             =   4600
      Width           =   1092
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
      Height          =   315
      Left            =   5328
      TabIndex        =   16
      Top             =   4600
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Height          =   450
      Left            =   7680
      Picture         =   "Frmacc3310.frx":0004
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   960
      Width           =   450
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3310.frx":066E
      Height          =   3000
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   8055
      _ExtentX        =   14199
      _ExtentY        =   5292
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "兌現票據資料"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a0e02"
         Caption         =   "票據號碼"
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
         DataField       =   "a0e01"
         Caption         =   "收票銀行"
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
         DataField       =   "a0e07"
         Caption         =   "收票帳號"
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
         DataField       =   "a0e11"
         Caption         =   "票據金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "a0e35"
         Caption         =   "是否退補票"
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
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1429.795
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1700.221
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
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
      Height          =   315
      Left            =   6765
      MaxLength       =   10
      TabIndex        =   4
      Top             =   465
      Width           =   1572
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
      Height          =   315
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   5
      Top             =   830
      Width           =   1100
   End
   Begin VB.TextBox Text1 
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
      Left            =   1335
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   1170
      Width           =   1095
      _ExtentX        =   1940
      _ExtentY        =   529
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
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   564
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   1095
      _ExtentX        =   1940
      _ExtentY        =   529
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "批：到期日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   22
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "(1.批次作業 2.單筆作業)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4665
      TabIndex        =   21
      Top             =   105
      Width           =   3630
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "票據金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2460
      TabIndex        =   20
      Top             =   825
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      Left            =   1485
      TabIndex        =   17
      Top             =   4600
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "小計"
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
      Left            =   4365
      TabIndex        =   15
      Top             =   4600
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "託收帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2460
      TabIndex        =   14
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "託收銀行"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5835
      TabIndex        =   13
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "單：票據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   12
      Top             =   825
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "(1.收票 2.開票)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2010
      TabIndex        =   11
      Top             =   105
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "票據別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   540
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   48
      Top             =   4272
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "票據兌現日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   9
      Top             =   1170
      Width           =   1260
   End
End
Attribute VB_Name = "Frmacc3310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/30 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoBatch As New ADODB.Recordset
Dim strRemark As String
Public strA0E23 As String   '2014/1/24 add by sonia
Dim bolProcessOK As Boolean 'Add by Amy 2014/11/17

Private Sub Command1_Click()
   AdodcDelete
   SumShow
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/11 原W8850 H5500
   Me.Width = 8865
   Me.Height = 5700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath3)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Text = CFDate(strSrvDate(2))
   MaskEdBox1.Mask = DFormat
   'ADD BY SONIA 2013/5/31 加到期日期條件,改抓到期日期以前的資料做兌現,兌現日期做在票據兌現日期
   MaskEdBox2.Text = CFDate(ChangeWStringToTString(CompWorkDay(3, strSrvDate(1), 1)))   '預設系統日的前二個工作天,當天不算
   MaskEdBox2.Mask = DFormat
   '2013/5/31 END
   Text8 = "1"
   Text2.Enabled = False
   PUB_SetAccount Combo1 'Add by Morgan 2006/7/21
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc3310 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
'MODIFY BY SONIA 2013/5/31 改在MaskEdBox2做
'   Select Case Text1
'      Case "1"
'         Combo1.SetFocus
'      Case "2"
'         Combo1.SetFocus
'   End Select
'   AdodcRefresh
'   SumShow
   If ChkWorkDay(DBDATE(MaskEdBox1)) = False Then
      MsgBox "票據兌現日期必須為工作日!!"
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

'ADD BY SONIA 2013/5/31
Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   Select Case Text1
      Case "1"
         Combo1.SetFocus
      Case "2"
         Combo1.SetFocus
   End Select
   AdodcRefresh
   SumShow
End Sub
'2013/5/31 END

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = "1" Then
      Text8 = "2"
   Else
      Text8 = "1"
      'Combo1.Text = "0149950" 'Add by Morgan 2006/7/31 --瑞婷
      'modify by sonia 2020/6/19
      'Combo1.Text = "0149951" 'Mark by Amy 2020/07/23 改回1756650 2010/6/21 MODIFY BY SONIA --瑞婷
      'Modify by Amy 2023/05/24
      'Combo1.Text = "1756650"
      SetCombo1 "1756650"
   End If
   Select Case Text8
      Case "1"
         Text2.Enabled = False
      Case "2"
         Text2.Enabled = True
   End Select
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         adoquery.CursorLocation = adUseClient
         'modify by sonia 2024/12/5 +a0e04條件
         adoquery.Open "select a0e11 from acc0e0 where a0e02 = '" & Text2 & "' and a0e04 = '" & IIf(Text1 = "1", "R", "P") & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("a0e11").Value) Then
               Text7 = ""
            Else
               Text7 = Format(adoquery.Fields("a0e11").Value, DDollar)
            End If
         Else
            Text7 = ""
         End If
         adoquery.Close
   End Select
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   adoquery.CursorLocation = adUseClient
   'modify by sonia 2024/12/5 +a0e04條件
   adoquery.Open "select a0e11 from acc0e0 where a0e02 = '" & Text2 & "' and a0e04 = '" & IIf(Text1 = "1", "R", "P") & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("a0e11").Value) Then
         Text7 = ""
      Else
         Text7 = Format(adoquery.Fields("a0e11").Value, DDollar)
      End If
   Else
      Text7 = ""
   End If
   adoquery.Close
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   Select Case Text8
      Case "1"
         BatchProcess
   End Select
   AdodcRefresh
   SumShow
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   '2013/5/31 MODIFY BY SONIA 改為MaskEdBox2
   'adoadodc1.Open "select * from acc0e0 where a0e02 = '" & Text2 & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select * from acc0e0 where a0e02 = '" & Text2 & "' and a0e21 = " & Val(FCDate(MaskEdBox2.Text)) & " order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
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
Private Sub AdodcRefresh()

   'Add by Morgan 2005/9/28
   If Trim(Text1) = Empty Then Exit Sub
   
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2023/05/23 原:Combo1
   strExc(1) = ""
   If Combo1 <> MsgText(601) Then strExc(1) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   Select Case Text1
      Case "1"
         'Modify by Amy 2021/06/29 bug-原:MaskEdBox1.Text 改為MaskEdBox2.Text (與加總不一致)
         'Modify by Amy 2021/06/30 應抓MaskEdBox1-瑞婷
         adoadodc1.Open "select * from acc0e0 where a0e19 = '" & Text3 & "' and a0e20 = '" & strExc(1) & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e04 = 'R' and (a0e45 is null or a0e45 = '99') and (a0e14 is not null and a0e14 <> 0) order by a0e48 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case "2"
         '2010/6/21 MODIFY BY SONIA 因甲存多一個帳號故要判斷帳號
         'adoadodc1.Open "select * from acc0e0 where a0e37 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e04 = 'P' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         adoadodc1.Open "select * from acc0e0 where a0e37 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e04 = 'P' and a0e07 = '" & strExc(1) & "' and a0e01 = '" & Text3 & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   'end 2023/05/23
   Adodc1.Recordset.Requery
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub

'*************************************************
'  應付票據兌現處理
'
'*************************************************
Private Sub ProcessP(strValue As String)
Dim adocheck As New ADODB.Recordset
Dim strAutoNo As String
Dim strMsg As String 'Add by Amy 2014/11/17
Dim strCombo1 As String 'Add by Amy 2023/05/23
   
On Error GoTo Checking
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select * from acc0e0 where a0e02 = '" & strValue & "' and a0e04 = 'P' and (a0e37 = 0 or a0e37 is null) and a0e25 = 0", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
'      MsgBox MsgText(137), , MsgText(5)
      adocheck.Close
      bolProcessOK = True 'Add by Amy 2014/11/17
      Exit Sub
   End If
   'Add by Amy 2014/11/17
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        MsgBox Label3 & MsgText(52), , MsgText(5)
        adocheck.Close
        MaskEdBox1.SetFocus
       Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
        MsgBox Label3 & MsgText(63), , MsgText(5)
        adocheck.Close
        MaskEdBox1.SetFocus
      Exit Sub
   End If
   If ChkWorkDay(DBDATE(MaskEdBox1)) = False Then
      MsgBox "票據兌現日期必須為工作日!!"
      adocheck.Close
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If ChkWorkData(strA0E23, DBDATE(MaskEdBox1), strMsg) = False Then
        MsgBox Label3 & strMsg, , MsgText(5)
        adocheck.Close
        MaskEdBox1.SetFocus
       Exit Sub
   End If
   'end 2014/11/17
   adoquery.CursorLocation = adUseClient
   'adoquery.Open "select a1p14 from acc1p0 where a1p01 = '1' and a1p09 = '" & strValue & "' and a1p10 = '" & adocheck.Fields("a0e01").Value & "' and a1p05 = '2111' and a1p07 = 0", adoTaie, adOpenStatic, adLockReadOnly
   '93.11.26 MODIFY BY SONIA  A0E05='4'時同時要抓備註A0E12
   'adoquery.Open "select a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(cu04, 1, 12) from acc0e0, customer where substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e05 = '1' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "' union " & _
   '              "select a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(a0i02, 1, 12) from acc0e0, acc0i0 where a0e06 = a0i01 (+) and a0e05 = '2' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "' union " & _
   '              "select a0e10||'/'||a0e02||'/'||a0e07||'/'||st02 from acc0e0, staff where a0e06 = st01 (+) and a0e05 = '3' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "' union " & _
   '              "select a0e10||'/'||a0e02||'/'||a0e07 from acc0e0 where a0e05 = '4' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   '2005/11/25 MODIFY BY SONIA 無論哪種往來類別,若備註A0E12有值,傳票摘要都要抓備註
   'adoquery.Open "select a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(cu04, 1, 12) from acc0e0, customer where substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e05 = '1' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "' union " & _
   '              "select a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(a0i02, 1, 12) from acc0e0, acc0i0 where a0e06 = a0i01 (+) and a0e05 = '2' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "' union " & _
   '              "select a0e10||'/'||a0e02||'/'||a0e07||'/'||st02 from acc0e0, staff where a0e06 = st01 (+) and a0e05 = '3' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "' union " & _
   '              "select a0e10||'/'||a0e02||'/'||a0e07||'/'||A0E12 from acc0e0 where a0e05 = '4' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/07/23 +a0e07 因改為key
   'Modify by Amy 2021/08/18 a0e05 = '4'那句加 a0e07,因2020/07/23改時,未改到
   'modify by sonia 2024/12/5 +a0e04條件
   adoquery.Open "select NVL((a0e10||'/'||a0e02||'/'||a0e07||'/'||A0E12),(a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(cu04, 1, 12))) from acc0e0, customer where substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e04 = 'P' and a0e05 = '1' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "' And a0e07='" & adocheck.Fields("a0e07").Value & "' union " & _
                 "select NVL((a0e10||'/'||a0e02||'/'||a0e07||'/'||A0E12),(a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(a0i02, 1, 12))) from acc0e0, acc0i0 where a0e06 = a0i01 (+) and a0e04 = 'P' and a0e05 = '2' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "'  And a0e07='" & adocheck.Fields("a0e07").Value & "' union " & _
                 "select NVL((a0e10||'/'||a0e02||'/'||a0e07||'/'||st02),(a0e10||'/'||a0e02||'/'||a0e07||'/'||st02)) from acc0e0, staff where a0e06 = st01 (+) and a0e04 = 'P' and a0e05 = '3' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "'  And a0e07='" & adocheck.Fields("a0e07").Value & "' union " & _
                 "select a0e10||'/'||a0e02||'/'||a0e07||'/'||A0E12 from acc0e0 where a0e05 = '4' and a0e04 = 'P' and a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "'  And a0e07='" & adocheck.Fields("a0e07").Value & "' ", adoTaie, adOpenStatic, adLockReadOnly
   '2005/11/25 END
   '93.11.26 END
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         strRemark = ""
      Else
         strRemark = adoquery.Fields(0).Value
      End If
   End If
   adoquery.Close
   adoTaie.BeginTrans
   'Modify by Amy 2020/07/23 +a0e07 因改為key
   'modify by sonia 2024/12/5 +a0e04條件
   adoTaie.Execute "update acc0e0 set a0e37 = " & Val(FCDate(MaskEdBox1.Text)) & " where a0e02 = '" & strValue & "' and a0e01 = '" & adocheck.Fields("a0e01").Value & "' and a0e04 = 'P' And a0e07='" & adocheck.Fields("a0e07").Value & "' and (a0e25 = 0 or a0e25 is null)"
   strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & strValue & adocheck.Fields("a0e01").Value & "6" & "'", 3)
   adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('1', 'L', '" & strAutoNo & "', '" & strValue & adocheck.Fields("a0e01").Value & adocheck.Fields("a0e07").Value & "6" & "', '2111', '" & MsgText(55) & "', " & adocheck.Fields("a0e11").Value & ", 0, '" & strValue & "', '" & adocheck.Fields("a0e01").Value & "', '" & adocheck.Fields("a0e07").Value & "', " & _
                   "" & Val(adocheck.Fields("a0e10").Value) & ", '" & adocheck.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                   "'" & adocheck.Fields("a0e03").Value & "', null, null, '2', null)"
   strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & strValue & adocheck.Fields("a0e01").Value & adocheck.Fields("a0e07").Value & "6" & "'", 3)
   adoquery.CursorLocation = adUseClient
   'Modify by Amy 2023/05/23 原:Combo1
   strCombo1 = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   adoquery.Open "select a0h08 from acc0h0 where a0h01 = '" & Text3 & "' and a0h02 = '" & strCombo1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('1', 'L', '" & strAutoNo & "', '" & strValue & adocheck.Fields("a0e01").Value & adocheck.Fields("a0e07").Value & "6" & "', '" & adoquery.Fields(0).Value & "', '" & MsgText(55) & "', 0, " & adocheck.Fields("a0e11").Value & ", '" & strValue & "', '" & adocheck.Fields("a0e01").Value & "', '" & adocheck.Fields("a0e07").Value & "', " & _
                      "" & Val(adocheck.Fields("a0e10").Value) & ", '" & adocheck.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                      "'" & adocheck.Fields("a0e03").Value & "', null, null, '2', null)"
   Else
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('1', 'L', '" & strAutoNo & "', '" & strValue & adocheck.Fields("a0e01").Value & adocheck.Fields("a0e07").Value & "6" & "', '110201', '" & MsgText(55) & "', 0, " & adocheck.Fields("a0e11").Value & ", '" & strValue & "', '" & adocheck.Fields("a0e01").Value & "', '" & adocheck.Fields("a0e07").Value & "', " & _
                      "" & Val(adocheck.Fields("a0e10").Value) & ", '" & adocheck.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                      "'" & adocheck.Fields("a0e03").Value & "', null, null, '2', null)"
   End If
   adoquery.Close
   adoTaie.CommitTrans
   adocheck.Close
   bolProcessOK = True 'Add by Amy 2014/11/17
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  應收票據兌現處理
'
'*************************************************
Private Sub ProcessR(strValue As String)
'edit by nickc 2007/02/08
Dim strAutoNo As String
Dim strMsg As String 'Add byAmy 2014/11/17
Dim strCombo1 As String 'Add by Amy 2023/05/23

On Error GoTo Checking
   If strValue = MsgText(601) Or Text3 = MsgText(601) Or Combo1 = MsgText(601) Then
      AdodcRefresh
      bolProcessOK = True 'Add by Amy 2014/11/17
      Exit Sub
   End If
   'Modify by Amy 2014/11/17 +票據兌現日必填
   'adoTaie.BeginTrans
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Morgan 2004/11/2 加 and a0e04 = 'R' 控制應收資料
   'Modify by Amy 2023/05/23 原:Combo1
   strCombo1 = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   adoacc0e0.Open "select * from acc0e0 where a0e02 = '" & strValue & "' and a0e19 = '" & Text3 & "' and a0e20 = '" & strCombo1 & "' and (a0e21 = 0 or a0e21 is null) and a0e15 = 0 and a0e04 = 'R' order by a0e10 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0e0.RecordCount = 0 Then
      'adoTaie.RollbackTrans
      MsgBox MsgText(33) & " " & MsgText(39), , MsgText(5)
      adoacc0e0.Close
      Exit Sub
   End If
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        MsgBox Label3 & MsgText(52), , MsgText(5)
        adoacc0e0.Close
        MaskEdBox1.SetFocus
       Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
        MsgBox Label3 & MsgText(63), , MsgText(5)
        adoacc0e0.Close
        MaskEdBox1.SetFocus
      Exit Sub
   End If
   If ChkWorkDay(DBDATE(MaskEdBox1)) = False Then
      MsgBox "票據兌現日期必須為工作日!!"
      adoacc0e0.Close
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If ChkWorkData(strA0E23, DBDATE(MaskEdBox1), strMsg) = False Then
        MsgBox Label3 & strMsg, , MsgText(5)
        adoacc0e0.Close
        MaskEdBox1.SetFocus
       Exit Sub
   End If
   
   adoTaie.BeginTrans
   'end
   adoacc0e0.Fields("a0e21").Value = Val(FCDate(MaskEdBox1.Text))
   adoacc0e0.Fields("a0e29").Value = Val(strSrvDate(2))
   adoacc0e0.Fields("a0e30").Value = ServerTime
   adoacc0e0.Fields("a0e31").Value = strUserNum
   adoacc0e0.Fields("a0e48").Value = ServerTime
   adoquery.CursorLocation = adUseClient
   'adoquery.Open "select a1p14 from acc1p0 where a1p01 = '1' and a1p09 = '" & strValue & "' and a1p10 = '" & adoacc0e0.Fields("a0e01").Value & "' and a1p05 = '113001' and a1p08 = 0", adoTaie, adOpenStatic, adLockReadOnly
   '2010/3/3 MODIFY BY SONIA 兌現傳票摘要最後的票號改抓銀行名稱
   'adoquery.Open "select a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(a0e02, 1, 12) from acc0e0 where a0e02 = '" & strValue & "' and a0e01 = '" & adoacc0e0.Fields("a0e01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/07/23 +a0e07 因改為key
   'modify by sonia 2024/12/5 +a0e04條件
   adoquery.Open "select a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(a0g02, 1, 12) from acc0e0, acc0g0 where a0e01 = a0g01 (+) and a0e02 = '" & strValue & "' and a0e01 = '" & adoacc0e0.Fields("a0e01").Value & "' and a0e04='R' And a0e07='" & adoacc0e0.Fields("a0e07") & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         strRemark = ""
      Else
         strRemark = adoquery.Fields(0).Value
      End If
   End If
   adoquery.Close
   '2014/1/24 modify by sonia a1p01='1' 改為a1p01='" & strA0E23 & "
   'Modify by Amy 2020/07/23 +a0e07 因改為key
   strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & strA0E23 & "' and a1p02 = 'L' and a1p04 = '" & strValue & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07") & "5" & "'", 3)
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a0h08 from acc0h0 where a0h01 = '" & Text3 & "' and a0h02 = '" & strCombo1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('" & strA0E23 & "', 'L', '" & strAutoNo & "', '" & strValue & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "5" & "', '" & adoquery.Fields(0).Value & "', '" & MsgText(55) & "', " & adoacc0e0.Fields("a0e11").Value & ", 0, '" & strValue & "', '" & Text3 & "', '" & strCombo1 & "', " & _
                      "" & Val(adoacc0e0.Fields("a0e10").Value) & ", '" & adoacc0e0.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                      "'" & adoacc0e0.Fields("a0e03").Value & "', null, null, '1', null)"
   Else
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('" & strA0E23 & "', 'L', '" & strAutoNo & "', '" & strValue & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "5" & "', '110201', '" & MsgText(55) & "', " & adoacc0e0.Fields("a0e11").Value & ", 0, '" & strValue & "', '" & Text3 & "', '" & strCombo1 & "', " & _
                      "" & Val(adoacc0e0.Fields("a0e10").Value) & ", '" & adoacc0e0.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                      "'" & adoacc0e0.Fields("a0e03").Value & "', null, null, '1', null)"
   End If
   adoquery.Close
   strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & strA0E23 & "' and a1p02 = 'L' and a1p04 = '" & strValue & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "5" & "'", 3)
   adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('" & strA0E23 & "', 'L', '" & strAutoNo & "', '" & strValue & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "5" & "', '113001', '" & MsgText(55) & "', 0, " & adoacc0e0.Fields("a0e11").Value & ", '" & strValue & "', '" & Text3 & "', '" & strCombo1 & "', " & _
                   "" & Val(adoacc0e0.Fields("a0e10").Value) & ", '" & adoacc0e0.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                   "'" & adoacc0e0.Fields("a0e03").Value & "', null, null, '1', null)"
   'end 2020/07/23
   'end 2023/05/23
   adoacc0e0.UpdateBatch
   adoTaie.CommitTrans
   adoacc0e0.Close
   bolProcessOK = True 'Add by Amy 2014/11/17
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
Dim Cancel As Boolean   'add by sonia 2022/5/27
   Select Case KeyCode
      Case vbKeyInsert
         bolProcessOK = False 'Add by Amy 2014/11/17
         Select Case Text8
            Case "2"
               Select Case Text1
                  Case "1"
                     Combo1_Validate (Cancel)    'add by sonia 2022/5/27 未跳離直接按INSERT造成公司別錯誤6247061之1110526兌現
                     ProcessR Text2
                     'Add by Amy 2014/11/17
                     If bolProcessOK = False Then Exit Sub
                     Text2 = MsgText(601)
                     Text2.SetFocus
                  Case "2"
                     ProcessP Text2
                     If bolProcessOK = False Then Exit Sub
                     Text2 = MsgText(601)
                     Text2.SetFocus
               End Select
                AdodcRefresh
                SumShow
         End Select
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Select Case Text1
      Case "1"
         adoquery.CursorLocation = adUseClient
         '2014/1/24 modify by sonia a1p01='1' 改為a1p01='" & strA0E23 & "
         'Modify by Amy 2020/07/23 +a0e07 因改為key
         adoquery.Open "select ax210 from acc1p0, acc021 where a1p22 = ax202 and a1p01 = '" & strA0E23 & "' and a1p02 = 'L' and a1p04 = '" & Adodc1.Recordset.Fields("a0e02").Value & Adodc1.Recordset.Fields("a0e01").Value & Adodc1.Recordset.Fields("a0e07").Value & "5" & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(180), , MsgText(5)
            adoquery.Close
            Exit Sub
         End If
      Case "2"
         adoquery.CursorLocation = adUseClient
         'Modify by Amy 2020/07/23 +a0e07 因改為key
         adoquery.Open "select ax210 from acc1p0, acc021 where a1p22 = ax202 and a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Adodc1.Recordset.Fields("a0e02").Value & Adodc1.Recordset.Fields("a0e01").Value & Adodc1.Recordset.Fields("a0e07").Value & "6" & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(180), , MsgText(5)
            adoquery.Close
            Exit Sub
         End If
   End Select
   adoquery.Close
   adoTaie.BeginTrans
   'Modify by Amy 2020/07/23 +a0e07 因改為key
   Select Case Text1
      Case "1"
         '2014/1/24 modify by sonia a1p01='1' 改為a1p01='" & strA0E23 & "
         adoTaie.Execute "delete from acc1p0 where a1p01 = '" & strA0E23 & "' and a1p02 = 'L' and a1p04 = '" & Adodc1.Recordset.Fields("a0e02").Value & Adodc1.Recordset.Fields("a0e01").Value & Adodc1.Recordset.Fields("a0e07") & "5" & "'"
         'modify by sonia 2024/12/5 +a0e04條件
         adoTaie.Execute "update acc0e0 set a0e21 = 0 where a0e02 = '" & Adodc1.Recordset.Fields("a0e02").Value & "' and a0e01 = '" & Adodc1.Recordset.Fields("a0e01").Value & "' and a0e04='R' And a0e07='" & Adodc1.Recordset.Fields("a0e07") & "' "
      Case "2"
         adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Adodc1.Recordset.Fields("a0e02").Value & Adodc1.Recordset.Fields("a0e01").Value & Adodc1.Recordset.Fields("a0e07") & "6" & "'"
         'modify by sonia 2024/12/5 +a0e04條件
         adoTaie.Execute "update acc0e0 set a0e37 = 0 where a0e02 = '" & Adodc1.Recordset.Fields("a0e02").Value & "' and a0e01 = '" & Adodc1.Recordset.Fields("a0e01").Value & "' and a0e04='P' And a0e07='" & Adodc1.Recordset.Fields("a0e07") & "' "
   End Select
   'end 2020/07/23
'   Adodc1.Recordset.Fields("a0e21").Value = 0
'   Adodc1.Recordset.UpdateBatch
   adoTaie.CommitTrans
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計算並顯示筆數與小計
'
'*************************************************
Private Sub SumShow()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Amy 2023/05/23 原:Combo1
   strExc(1) = ""
   If Combo1 <> MsgText(601) Then strExc(1) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   Select Case Text1
      Case "1"
         '2013/5/31 MODIFY BY SONIA 改為MaskEdBox2
         'adoacc0e0.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e19 = '" & Text3 & "' and a0e20 = '" & Combo1 & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e04 = 'R' and (a0e45 is null or a0e45 = '99') and (a0e14 is not null and a0e14 <> 0)", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Amy 2021/06/30 改回MaskEdBox1-瑞婷,秀玲說忘了  2013/5/31 為何改為MaskEdBox2
         adoacc0e0.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e19 = '" & Text3 & "' and a0e20 = '" & strExc(1) & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e04 = 'R' and (a0e45 is null or a0e45 = '99') and (a0e14 is not null and a0e14 <> 0)", adoTaie, adOpenStatic, adLockReadOnly
      Case "2"
         '2010/6/21 MODIFY BY SONIA 因甲存多一個帳號故要判斷帳號
         'adoacc0e0.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e37 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e04 = 'P'", adoTaie, adOpenStatic, adLockReadOnly
         adoacc0e0.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e37 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e04 = 'P' and a0e07 = '" & strExc(1) & "' and a0e01 = '" & Text3 & "' ", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   'end 2023/05/23
   If adoacc0e0.RecordCount <> 0 Then
      If IsNull(adoacc0e0.Fields(0).Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = Format(adoacc0e0.Fields(0).Value, DDollar)
      End If
      If IsNull(adoacc0e0.Fields(1).Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = Format(adoacc0e0.Fields(1).Value, DDollar)
      End If
   Else
      Text6 = MsgText(601)
      Text6 = MsgText(601)
   End If
   adoacc0e0.Close
   AdodcRefresh
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
Dim adoacc0h0 As New ADODB.Recordset
Dim strCombo1 As String 'Add by Amy 2023/05/23

   adoacc0h0.CursorLocation = adUseClient
   'Modify by Amy 2023/05/23 原:Combo1
   If Combo1 <> MsgText(601) Then strCombo1 = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   adoacc0h0.Open "select a0h01 from acc0h0 where a0h02 = '" & strCombo1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0h0.RecordCount <> 0 Then
      Text3 = adoacc0h0.Fields(0).Value
      '2014/1/24 add by sonia
      'modify by sonia 2015/5/12 加智權 華銀長安0236819
      If strCombo1 = "1607750" Or strCombo1 = "0236819" Then
         strA0E23 = "J"
      'add by sonia 2020/4/7 加法律所
      ElseIf strCombo1 = "1756890" Then
      'end 2023/05/23
         strA0E23 = "L"
      'end 2020/4/7
      Else
         strA0E23 = "1"
      End If
      '2014/1/24 end
   Else
      Text3 = MsgText(601)
   End If
   adoacc0h0.Close
   AdodcRefresh
   SumShow
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   Select Case Text8
      Case "1"
         Text2.Enabled = False
      Case "2"
         Text2.Enabled = True
   End Select
End Sub

'*************************************************
'  整批作業
'
'*************************************************
Private Sub BatchProcess()
   
   'Add by Morgan 2005/9/28
   If Trim(Text1) = Empty Then Exit Sub
   
On Error GoTo Checking
   If adoBatch.State = adStateOpen Then
      adoBatch.Close
   End If
   adoBatch.CursorLocation = adUseClient
   'Modify by Amy 2023/05/23 原:Combo1
   strExc(1) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   Select Case Text1
      Case "1"
         '2013/5/31 MODIFY BY SONIA 改為MaskEdBox2
         'adoBatch.Open "select a0e02 from acc0e0 where a0e04 = 'R' and (a0e21 is null or a0e21 = 0) and (a0e25 is null or a0e25 = 0) and a0e10 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e20 = '" & Combo1 & "' and a0e19 = '" & Text3 & "' order by a0e10 asc, a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         adoBatch.Open "select a0e02 from acc0e0 where a0e04 = 'R' and (a0e21 is null or a0e21 = 0) and (a0e25 is null or a0e25 = 0) and a0e10 = " & Val(FCDate(MaskEdBox2.Text)) & " and a0e20 = '" & strExc(1) & "' and a0e19 = '" & Text3 & "' order by a0e10 asc, a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Do While adoBatch.EOF = False
            If IsNull(adoBatch.Fields("a0e02").Value) = False Then
               ProcessR adoBatch.Fields("a0e02").Value
            End If
            adoBatch.MoveNext
         Loop
      Case "2"
         '2013/5/31 MODIFY BY SONIA 改為MaskEdBox2
         'adoBatch.Open "select a0e02 from acc0e0 where a0e04 = 'P' and (a0e37 is null or a0e37 = 0) and (a0e25 is null or a0e25 = 0) and a0e10 <= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e07 = '" & Combo1 & "' and a0e01 = '" & Text3 & "' order by a0e10 asc, a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         adoBatch.Open "select a0e02 from acc0e0 where a0e04 = 'P' and (a0e37 is null or a0e37 = 0) and (a0e25 is null or a0e25 = 0) and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0e07 = '" & strExc(1) & "' and a0e01 = '" & Text3 & "' order by a0e10 asc, a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Do While adoBatch.EOF = False
            If IsNull(adoBatch.Fields("a0e02").Value) = False Then
               ProcessP adoBatch.Fields("a0e02").Value
            End If
            adoBatch.MoveNext
         Loop
   End Select
   'end 2023/05/23
   adoBatch.Close
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub

'Add by Amy 2023/05/23
Private Sub SetCombo1(pCode As String)
   Dim idx As Integer
   For idx = 0 To Combo1.ListCount - 1
      If InStr(Combo1.List(idx), pCode) = 1 Then
         Combo1.ListIndex = idx
         Exit For
      End If
   Next
   If idx = Combo1.ListCount Then
      Combo1.AddItem pCode
      Combo1 = pCode
   End If
End Sub

