VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc7160 
   AutoRedraw      =   -1  'True
   Caption         =   "分所收款資料查詢-智權人員繳款"
   ClientHeight    =   5004
   ClientLeft      =   1536
   ClientTop       =   2820
   ClientWidth     =   8724
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5200.274
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8730
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1710
      TabIndex        =   2
      Top             =   570
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "統計資料"
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
      Left            =   7140
      TabIndex        =   5
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "繳款內容"
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
      Left            =   7140
      TabIndex        =   4
      Top             =   270
      Width           =   1092
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc7160.frx":0000
      Height          =   3554
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   8445
      _ExtentX        =   14880
      _ExtentY        =   6287
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   26
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "R43105"
         Caption         =   "出納確認日"
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
         DataField       =   "R43104"
         Caption         =   "繳款人"
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
         DataField       =   "R43102"
         Caption         =   "收據號碼"
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
         DataField       =   "R43103"
         Caption         =   "收據抬頭"
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
         DataField       =   "R43109"
         Caption         =   "票據金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "R43110"
         Caption         =   "北所電匯"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "R43111"
         Caption         =   "分所電匯"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "R43112"
         Caption         =   "現金"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "R43113"
         Caption         =   "抵暫收款"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "R43120"
         Caption         =   "其他"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "R43114"
         Caption         =   "溢收款"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "R43115"
         Caption         =   "手續費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "R43116"
         Caption         =   "補扣繳"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "R43117"
         Caption         =   "外幣"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "R43118"
         Caption         =   "總計"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "R43119"
         Caption         =   "點數"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   984.866
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   924.73
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1345.114
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1477.299
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1045.002
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1045.002
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Object.Visible         =   0   'False
            ColumnWidth     =   1261.15
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            ColumnWidth     =   1068.829
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   3630
      TabIndex        =   1
      Top             =   210
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
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
      Left            =   8400
      Top             =   30
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1710
      TabIndex        =   0
      Top             =   210
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
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
   Begin MSForms.TextBox Text2 
      Height          =   314
      Left            =   1690
      TabIndex        =   3
      Top             =   930
      Width           =   5500
      VariousPropertyBits=   683687963
      MaxLength       =   100
      Size            =   "9701;554"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   288
      Left            =   2880
      TabIndex        =   10
      Top             =   596
      Width           =   1725
      VariousPropertyBits=   19
      Size            =   "3043;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "可模糊比對"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   7320
      TabIndex        =   12
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收款抬頭"
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
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   240
      TabIndex        =   8
      Top             =   570
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "出納確認日期"
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
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3390
      TabIndex        =   6
      Top             =   210
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc7160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/1/2 Form2.0已修改(lblSalesName,Text2,DataGrid1改Fonts)
'Create by Lydia 2014/10/3 分所收款資料查詢-智權人員繳款
'Memo by Lydia 顯示資料:除了明細資料,另外要有智權人員小計和全部人員總計
'並將處理好的資料丟到ACCRPT431供統計資料(Frmacc7161)直接使用
Option Explicit
Public adoadodc1 As New ADODB.Recordset
Dim mCheck01 As String  'check a4401
Dim mCheck02 As Long 'check a4402
Dim mCheck03 As Long 'check a4403
Dim mCheck04 As String 'check axd04
Dim mCheckDate As Long  'check a4413
'Added by Lydia 2015/07/15 設定欄位數
Private Const idx As Integer = 11
Dim mTT(0 To idx) As Double '各欄位,加總計
Dim Stt(0 To idx) As Double '智權人員小計
Dim Tsum(0 To idx) As Double '總計
''Modified by Lydia 2015/02/04 +點數(9=>10)
'Dim mTT(0 To 10) As Double '各欄位,加總計
'Dim Stt(0 To 10) As Double '智權人員小計
'Dim Tsum(0 To 10) As Double '總計
''end 2015/02/04
'end 2015/07/15
Dim intSno As Integer
Dim strSql As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim adoquery As ADODB.Recordset '繳款明細
Dim mAdodc1chk As Boolean 'Add by Lydia 2014/10/15 無資料時,不可按統計
Private Sub Command1_Click()
   'Add by Lydia 2014/10/15 無資料時,不可按統計
   If mAdodc1chk = False Then
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If

   tool3_enabled

Call cmdDetail_Click
End Sub

Private Sub Command2_Click()
   'Add by Lydia 2014/10/15 無資料時,不可按統計
   If mAdodc1chk = False Then
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strItemNo = ""
   strCon4 = MaskEdBox2.Text
   strCon5 = MaskEdBox3.Text
    strCon6 = Me.Text1.Text
    strCon7 = Me.lblSalesName.Caption
   strExitControl = ""
   tool3_enabled
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Frmacc7161.Show
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Me.Hide
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
   Me.Height = 5400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox2.Text = MsgText(601)
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = MsgText(601)
   MaskEdBox3.Text = MsgText(601)
   MaskEdBox3.Mask = DFormat
   OpenTable
   strExitControl = MsgText(602)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strExitControl = MsgText(602) Then
      StatusClear
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
      Set Frmacc7160 = Nothing
      Exit Sub
   End If
   strExitControl = MsgText(602)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
'on error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   'edit by nick 2004/08/20 讓分所可以查其他所  cancel
   adoadodc1.Open "Select * From ACC310 Where A3101='" & pub_strUserOffice & "' And A3103='" & Text1 & "' Order By A3103 ", adoTaie, adOpenStatic, adLockReadOnly
   'adoadodc1.Open "Select * From ACC310 Where A3103='" & Text1 & "' Order By A3103 ", adoTaie, adOpenStatic, adLockReadOnly
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
Public Sub AdodcRefresh()

'on error GoTo Checking
   
    adoTaie.Execute "Delete From ACCRPT431 Where R43100='" & strUserNum & "' "
    strSql = ""
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
        strSql = " And A4413 >= " & Val(FCDate(MaskEdBox2.Text)) + 19110000 & ""
    End If
    If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
        strSql = strSql & " And A4413 <= " & Val(FCDate(MaskEdBox3.Text)) + 19110000 & ""
    End If
    If Me.Text1.Text <> "" Then
        strSql = strSql & " And A4401='" & Me.Text1.Text & "' "
    End If

    If Me.Text2.Text <> "" Then
        strSql = strSql & " and instr(A0K04, '" & Me.Text2.Text & "') > 0 "
    End If

       
    intSno = 0
    
    Erase mTT
    Erase Stt
    Erase Tsum
 'Modified by Lydia 2015/02/04 + 點數=(服務費/1000)
 'Modified by Lydia 2015/07/15 +其他A4430
    StrSQLa = " SELECT AXD04,a0k04,ST02,A4413,A4401,A4402,A4403,A4405,A4406,A4407,A4408,A4409,A4410,A4411,A4422,A4426,nvl(a0k09, 0) as a0k09,(AXD06/1000)  AXD06 " & _
          " ,A4430 From ACC440, ACC441, ACC0K0 , STAFF " & _
          " WHERE A4401=AXD01(+) AND A4402=AXD02(+) AND A4403=AXD03(+) AND AXD04=a0k01(+) AND A4401=st01(+) and st06 ='" & pub_strUserOffice & "' "
    StrSQLa = StrSQLa & strSql & " group by  AXD04,a0k04,ST02,A4413,A4401,A4402,A4403,A4405,A4406,A4407,A4408,A4409,A4410,A4411,A4422,A4426,a0k09,(AXD06/1000),A4430 "
    StrSQLa = StrSQLa & " Order by A4413,A4401,A4402,A4403 " '依出納確認日,智權人員

    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly

    If rsA.RecordCount > 0 Then
        rsA.MoveFirst

        Do While rsA.EOF = False

          If Trim(mCheck01) = "" Then '首筆
            Call Pub7160_1
          Else
            If mCheckDate + 19110000 = rsA!a4413 And Trim(mCheck01) = Trim(rsA!A4401) And mCheck02 = rsA!A4402 And mCheck03 = rsA!A4403 Then
               '同一張繳款單,有多筆收據,只印第一筆
            Else
             If rsA.AbsolutePosition >= 19 Then
               intSno = intSno
             End If
              If mCheckDate + 19110000 <> rsA!a4413 Or Trim(mCheck01) <> Trim(rsA!A4401) Then
                Call Pub7160_2 '智權人員合計
                Call Pub7160_1
              Else
                If mCheck02 <> rsA!A4402 Or mCheck03 <> rsA!A4403 Then
                  Call Pub7160_1
                End If
              End If
            End If
          End If
           rsA.MoveNext
        Loop
        If rsA.EOF = True Then
           Call Pub7160_2 '智權人員合計
           Call Pub7160_3 '總計
        End If
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    If adoadodc1.State = adStateOpen Then
        adoadodc1.Close
    End If
    adoadodc1.CursorLocation = adUseClient
    StrSQLa = "Select * From ACCRPT431 Where R43100='" & strUserNum & "' Order By 2 "
    adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    Adodc1.Recordset.ReQuery
    
    mAdodc1chk = False 'Add by Lydia 2014/10/15 無資料時,不可按統計
    If Adodc1.Recordset.RecordCount = 0 Then
        Adodc1.Recordset.Close
        MsgBox MsgText(28), , MsgText(5)
        Exit Sub
    Else
        mAdodc1chk = True
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
            AdodcRefresh
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub



Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = MaskEdBox2.Text
   MaskEdBox3.Mask = DFormat
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'add by sonia 2022/1/2 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2022/1/2
   
   If Text1 <> MsgText(601) And Text1 <> MsgText(802) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox3.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   
   FormCheck = False
End Function


Private Sub Text1_Validate(Cancel As Boolean)
    If Me.Text1.Text = "" Then Me.lblSalesName.Caption = "": Exit Sub
    Me.lblSalesName.Caption = GetStaffName(Me.Text1.Text)
    If Me.lblSalesName.Caption = "" Then
        MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
        Cancel = True
    End If
    If Cancel = True Then Text1_GotFocus
End Sub

Private Sub Text2_GotFocus()
TextInverse Text2
End Sub

Private Sub Pub7160_1()
Dim mChkDel As String
    intSno = intSno + 1
    mCheck01 = rsA!A4401
    mCheck02 = rsA!A4402
    mCheck03 = rsA!A4403
    mCheck04 = Trim(rsA!axd04)
    If rsA!a0k09 > 0 Then
       mChkDel = mCheck04 & "(廢)"
    Else
       mChkDel = mCheck04
    End If
    mCheckDate = rsA!a4413 - 19110000 '西元轉民國
    mTT(0) = IIf(IsNull(rsA!A4405) = True, 0, rsA!A4405) '票據金額
    mTT(1) = IIf(IsNull(rsA!A4406) = True, 0, rsA!A4406) '北所電匯
    mTT(2) = IIf(IsNull(rsA!A4407) = True, 0, rsA!A4407) '分所電匯
    mTT(3) = IIf(IsNull(rsA!A4408) = True, 0, rsA!A4408) '現金
    mTT(4) = IIf(IsNull(rsA!A4409) = True, 0, rsA!A4409) '抵暫收款
    mTT(5) = IIf(IsNull(rsA!A4410) = True, 0, rsA!A4410) '溢收款
    mTT(6) = IIf(IsNull(rsA!A4411) = True, 0, rsA!A4411) '手續費
    mTT(7) = IIf(IsNull(rsA!A4422) = True, 0, rsA!A4422) '補扣繳
    mTT(8) = IIf(IsNull(rsA!A4426) = True, 0, rsA!A4426) '外幣
    'Added by Lydia 2015/07/15 + A4430
    mTT(11) = IIf(IsNull(rsA!A4430) = True, 0, rsA!A4430) '其他
    'Modified by Lydia 2015/07/15 + mTT(11)
    mTT(9) = mTT(0) + mTT(1) + mTT(2) + mTT(3) + mTT(4) - mTT(5) + mTT(6) + mTT(7) + mTT(8) + mTT(11)  '減溢收款
    
    'Modified by Lydia 2015/02/04 + 點數(R43119)
    mTT(10) = IIf(IsNull(rsA!AXD06) = True, 0, rsA!AXD06)
    'Added by Lydia 2015/07/15 + mTT(11)=R43120
    strSql = " Insert Into ACCRPT431(R43100,R43101, R43102, R43103, R43104, R43105, R43106, R43107, R43108, R43109, R43110, R43111, R43112,R43113,R43114, R43115, R43116,R43117,R43118,R43119,R43120)  "
    strSql = strSql & " values('" & strUserNum & "', " & intSno & ",'" & mChkDel & "','" & rsA!A0K04 & "','" & rsA!st02 & "' "
    strSql = strSql & " ," & mCheckDate & ",'" & mCheck01 & "'," & mCheck02 & " ," & mCheck03 & "," & mTT(0) & "," & mTT(1) & "," & mTT(2) & "," & mTT(3) & "," & mTT(4) & "," & mTT(5) & "," & mTT(6) & "," & mTT(7) & "," & mTT(8) & "," & mTT(9) & "," & mTT(10) & "," & mTT(11) & ") "
    adoTaie.Execute strSql
    
    'Modified by Lydia 2015/02/04  9=>10
    'Modified by Lydia 2015/07/15
    'For intI = 0 To 10
    For intI = 0 To idx
       Stt(intI) = Stt(intI) + mTT(intI)
    Next intI
    
End Sub

Private Sub Pub7160_2()
    intSno = intSno + 1
    'Modified by Lydia 2015/02/04 + 點數(R43119)
    'Added by Lydia 2015/07/15 + 其他(R43120)
    strSql = " Insert Into ACCRPT431(R43100,R43101, R43103, R43109, R43110, R43111, R43112,R43113,R43114, R43115, R43116,R43117,R43118,R43119,R43120)  "
    strSql = strSql & " values('" & strUserNum & "', " & intSno & ",'小計：'," & Stt(0) & "," & Stt(1) & "," & Stt(2) & "," & Stt(3) & "," & Stt(4) & "," & Stt(5) & "," & Stt(6) & "," & Stt(7) & "," & Stt(8) & "," & Stt(9) & "," & Stt(10) & "," & Stt(11) & ") "
    adoTaie.Execute strSql
    'Modified by Lydia 2015/02/04  9=>10
    'Modified by Lydia 2015/07/15
    'For intI = 0 To 10
    For intI = 0 To idx
       Tsum(intI) = Tsum(intI) + Stt(intI)
       Stt(intI) = 0
    Next intI
    mCheck01 = ""
    mCheck02 = 0
    mCheck03 = 0
    mCheck04 = ""
End Sub

Private Sub Pub7160_3()
    intSno = intSno + 1
    'Modified by Lydia 2015/02/04 + 點數(R43119)
    'Added by Lydia 2015/07/15 + 其他(R43120)
    strSql = " Insert Into ACCRPT431(R43100,R43101, R43103, R43109, R43110, R43111, R43112,R43113,R43114, R43115, R43116,R43117,R43118,R43119,R43120)  "
    strSql = strSql & " values('" & strUserNum & "', " & intSno & ",'總計：'," & Tsum(0) & "," & Tsum(1) & "," & Tsum(2) & "," & Tsum(3) & "," & Tsum(4) & "," & Tsum(5) & "," & Tsum(6) & "," & Tsum(7) & "," & Tsum(8) & "," & Tsum(9) & "," & Tsum(10) & "," & Tsum(11) & ") "
    adoTaie.Execute strSql
End Sub

Private Sub cmdDetail_Click()
'copy from Promoter.frm210142
Dim stVTB11 As String, stVTB22 As String
Dim iCol    As Integer
Dim rtNo    As String
Dim strCon  As String
Dim Role    As String
Dim PayDate As Long
Dim PayTime As Long
'Dim PayNo As String 'Modified by Lydia 2015/02/04 收據號碼
Dim dblVal(3) As Double
     
   Role = Adodc1.Recordset("R43106")                '智權人員
   PayDate = Adodc1.Recordset("R43107")             '繳款日期
   PayTime = Adodc1.Recordset("R43108")             '繳款時間
   'PayNo = Adodc1.Recordset("R43102")             '收據號碼
   'Modified by Lydia 2015/02/04 +收據號碼
   'Modified by Morgan 2015/10/2 不需收據號碼條件
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據;decode(nvl(a0k19,0),0,'◎')=> decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))
   strExc(0) = "select sqldatet(a0k02) 單據日期" & _
       ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
       ",decode(a0j04,'000',cpm03,cpm04) 案件性質" & _
       ",na03 國別,axd06 服務費,axd07 規費,axd08 扣繳金額,a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','＊') 收據編號" & _
       ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱,a0k03,a0k04" & _
       " from ACC441,ACC0J0,acc0k0,acc431,caseprogress,casepropertymap,nation" & _
       ",trademark,patent,lawcase,servicepractice,hirecase" & _
       " where A0J01(+)=AXD05 AND A0J13(+)=AXD04" & _
       " and axd01='" & Role & "' and axd02='" & PayDate & "' and axd03='" & PayTime & "'" & _
       " and a0k01(+)=a0j13 and axc02(+)=a0j13 and cp09(+)=a0j01" & _
       " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04" & _
       " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
       " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
       " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
       " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
       " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
       " order by a0k02,a0j13,a0j01"
                                   
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
    
   If adoquery.RecordCount > 0 Then
      With frm210141_3
      'Modified by Lydia 2019/07/03 更名
      'frm210141_3.Caption = "智權人員繳款資料查詢-繳款資料明細"
      frm210141_3.Caption = "繳款資料查詢-繳款資料明細"

      Set .Adodc1.Recordset = PUB_CreateRecordset(adoquery, , , , .Name)
         With adoquery
           .MoveFirst
           Do While Not .EOF
              dblVal(1) = dblVal(1) + Val("" & .Fields("服務費"))
              dblVal(2) = dblVal(2) + Val("" & .Fields("規費"))
              dblVal(3) = dblVal(3) + Val("" & .Fields("扣繳金額"))
              .MoveNext
           Loop
         End With
         
         .txtTot(2) = Format(dblVal(1), "#,##0")
         .txtTot(3) = Format(dblVal(2), "#,##0")
         .txtTot(4) = Format(dblVal(3), "#,##0")
         .txtTot(5) = Format(dblVal(1) + dblVal(2) - dblVal(3), "#,##0")
         .Show vbModal
      End With
   End If
   
End Sub

