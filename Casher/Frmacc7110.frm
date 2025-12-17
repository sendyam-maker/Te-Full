VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc7110 
   AutoRedraw      =   -1  'True
   Caption         =   "分所收款資料查詢"
   ClientHeight    =   5025
   ClientLeft      =   1530
   ClientTop       =   2820
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8760
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1230
      MaxLength       =   100
      TabIndex        =   11
      Top             =   930
      Width           =   6045
   End
   Begin VB.CommandButton Command2 
      Caption         =   "統計資料"
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
      Left            =   7140
      TabIndex        =   9
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "收款內容"
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
      Left            =   7140
      TabIndex        =   8
      Top             =   270
      Width           =   1092
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc7110.frx":0000
      Height          =   3705
      Left            =   150
      TabIndex        =   6
      Top             =   1290
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   6535
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "R43003"
         Caption         =   "收款日"
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
         DataField       =   "R43004"
         Caption         =   "收款人"
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
         DataField       =   "R43005"
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
      BeginProperty Column03 
         DataField       =   "R43006"
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
      BeginProperty Column04 
         DataField       =   "R43007"
         Caption         =   "人工號"
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
         DataField       =   "R43008"
         Caption         =   "電腦號"
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
         DataField       =   "R43009"
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
      BeginProperty Column07 
         DataField       =   "R43010"
         Caption         =   "支票"
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
         DataField       =   "R43011"
         Caption         =   "到期日"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "R43012"
         Caption         =   "扣繳額"
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
         DataField       =   "R43013"
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
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2234.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   884.976
         EndProperty
      EndProperty
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
      Left            =   1230
      TabIndex        =   2
      Top             =   570
      Width           =   975
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1230
      TabIndex        =   0
      Top             =   210
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   3150
      TabIndex        =   1
      Top             =   210
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   8400
      Top             =   30
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
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblSalesName 
      BackStyle       =   0  '透明
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
      Left            =   2400
      TabIndex        =   7
      Top             =   570
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   570
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
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
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label10 
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
      Left            =   2910
      TabIndex        =   3
      Top             =   210
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc7110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset
Dim mAdodc1chk As Boolean 'Add by Lydia 2014/10/15 無資料時,不可按統計

Private Sub Command1_Click()

   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
    If Adodc1.Recordset.Fields("R43008").Value & "," & Adodc1.Recordset.Fields("R43007").Value = "," Then
        Exit Sub
    End If
   strItemNo = Adodc1.Recordset.Fields("R43008").Value & "," & Adodc1.Recordset.Fields("R43007").Value
   strCon4 = MaskEdBox2.Text
   strCon5 = MaskEdBox3.Text
   strExitControl = MsgText(601)
   tool3_enabled
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Frmacc7111.Show
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Me.Hide
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
   strCon8 = Me.Text2.Text   'add by sonia 2018/4/19
   strExitControl = ""
   tool3_enabled
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Frmacc7112.Show
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Me.Hide
End Sub

Private Sub Form_Activate()
'   MaskEdBox2.Mask = ""
'   MaskEdBox2.Text = strCon4
'   MaskEdBox2.Mask = DFormat
'   MaskEdBox3.Mask = ""
'   MaskEdBox3.Text = strCon5
'   MaskEdBox3.Mask = DFormat
'   If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(601) Then
'      AdodcRefresh
'   End If
'   strExitControl = MsgText(602)
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
      Set Frmacc7110 = Nothing
      Exit Sub
   End If
   strExitControl = MsgText(602)
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
Dim strSql As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strNo As String
Dim strOurCaseNo As String
Dim strCaseData  As String
Dim strAmt  As String
Dim strPoint  As String
Dim strSalesNo As String
'add by nick 2004/08/20
Dim strA3124 As String
'add by nick 2004/08/26
Dim dblPoint As Double
Dim dblTPoint As Double

Dim ii As Integer
Dim dblCash As Double
Dim dblCheck As Double
Dim dblTot As Double
Dim dblTCash As Double
Dim dblTCheck As Double
Dim dblTTOT As Double

On Error GoTo Checking
    adoTaie.Execute "Delete From ACCRPT430 Where R43001='" & strUserNum & "' "
    strSql = ""
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
        strSql = " And A3102 >= " & Val(FCDate(MaskEdBox2.Text)) & ""
    End If
    If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
        strSql = strSql & " And A3102 <= " & Val(FCDate(MaskEdBox3.Text)) & ""
    End If
    If Me.Text1.Text <> "" Then
        'edit by nick 2004/12/30
        'StrSql = StrSql & " And A0K20='" & Me.Text1.Text & "' "
        strSql = strSql & " And A3121='" & Me.Text1.Text & "' "
    End If
    'add by nick 2005/01/05 增加收據抬頭用模糊
    If Me.Text2.Text <> "" Then
        strSql = strSql & " and instr(A3122, '" & Me.Text2.Text & "') > 0 "
    End If
    
    strNo = ""
    strOurCaseNo = ""
    strCaseData = ""
    strAmt = ""
    strPoint = ""
    strSalesNo = ""
    'add by nick 2004/08/20
    strA3124 = ""
    
    ii = 0
    'edit by nick 2004/08/20 可以查他所和新增欄位和以acc310 為主 cancel
    'strSQLA = "Select A3102 As 收款日, ST02 As 收款人, A0K04 As 收據抬頭, A0J02 As 本所案號, A0J20 As 案件性質名稱, Nvl(A0J09,0)+Nvl(A0J10,0) As 費用, A3104 As 人工號, A3103 As 電腦號, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳額, A3113 As 留分所金額, Round(Nvl(A0J09,0)/1000,1) As 點數, A0J09, A0J10, A0K20 From ACC310, ACC0k0, ACC0J0, Staff Where A3103=A0K01(+) And A0K01=A0J13(+) And A0K20=ST01(+) And A3101='" & pub_strUserOffice & "' " & strSQL
    'strSQLA = strSQLA & " Order By ST03, A0K20, A3101, A3102, A3103, A3104 "
    'Modified by Morgan 2011/12/27 取消 a0j20
    StrSQLa = "Select A3102 As 收款日, ST02 As 收款人, A3122 As 收據抬頭, A0J02 As 本所案號, getcp10desc(cp01,cp10,a0j04) As 案件性質名稱, Nvl(A0J09,0)+Nvl(A0J10,0) As 費用, A3104 As 人工號, A3103 As 電腦號, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳額, A3113 As 留分所金額, A3123 As 點數, A0J09, A0J10,A3124 as 備註,A3121 From ACC310,  ACC0J0, Staff,caseprogress Where A3103=A0J13(+) And A3121=ST01(+)  And A3101='" & pub_strUserOffice & "' " & strSql & " and cp09(+)=a0j01 "
    StrSQLa = StrSQLa & " Order By A3102, A3121, A3115, A3116 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        rsA.MoveFirst
        'edit by nick 2004/08/20
        'strSalesNo = "" & rsA("A0K20").Value
        strSalesNo = "" & rsA("A3121").Value
        Do While rsA.EOF = False
            If strNo = "" Then
                strNo = "" & rsA("人工號").Value & rsA("電腦號").Value
                strCaseData = IIf(strOurCaseNo <> "" & rsA("本所案號").Value, ReConBNOurCaseNO("" & rsA("本所案號").Value), "") & rsA("案件性質名稱").Value
                If strOurCaseNo <> "" & rsA("本所案號").Value Then
                    strOurCaseNo = "" & rsA("本所案號").Value
                End If
                strAmt = Val("" & rsA("A0J09").Value) + Val("" & rsA("A0J10").Value)
                strPoint = Val("" & rsA("點數").Value)
                GoTo NextRec
            Else
                'edit by  nick 2004/08/20
                'If strNo <> "" & rsA("人工號").Value & rsA("電腦號").Value Then
                If Trim(strNo) <> Trim("" & rsA("人工號").Value & rsA("電腦號").Value) Then
                    rsA.MovePrevious
                    GoTo PrintRec
                Else
                    strCaseData = strCaseData & "及" & IIf(strOurCaseNo <> "" & rsA("本所案號").Value, ReConBNOurCaseNO("" & rsA("本所案號").Value), "") & rsA("案件性質名稱").Value
                    If strOurCaseNo <> "" & rsA("本所案號").Value Then
                        strOurCaseNo = "" & rsA("本所案號").Value
                    End If
                    strAmt = Val(strAmt) + Val("" & rsA("A0J09").Value) + Val("" & rsA("A0J10").Value)
                    strPoint = Val(strPoint) + Val("" & rsA("點數").Value)
                    GoTo NextRec
                End If
            End If
            'add by nick 2004/08/20
            strA3124 = "" & rsA("備註").Value
PrintRec:
            'edit by nick 2004/08/20
            'If strSalesNo <> "" & rsA.Fields("A0K20").Value Then
            If strSalesNo <> "" & rsA.Fields("A3121").Value Then
                ii = ii + 1
                'edit by nick 2004/08/20
                'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ) "
                'edit by nick 2004/08/26
                'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ,'" & strA3124 & "') "
                'edit by nick 2004/10/07
                'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011,r43013,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ," & dblPoint & ",'" & strA3124 & "') "
                adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43008, R43009, R43010, R43011,r43013,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & ",'" & Format(dblTot, "###,###,###,##0") & "'," & dblPoint & ",'" & strA3124 & "') "
                dblCash = 0: dblCheck = 0: dblTot = 0
                'add by nick 2004/08/26
                dblPoint = 0
                'edit by nick 2004/08/20
                strSalesNo = "" & rsA.Fields("A3121").Value
            End If
            dblCash = dblCash + Val("" & rsA("現金").Value)
            dblCheck = dblCheck + Val("" & rsA("支票").Value)
            dblTot = dblTot + Val("" & rsA("現金").Value) + Val("" & rsA("支票").Value)
            dblTCash = dblTCash + Val("" & rsA("現金").Value)
            dblTCheck = dblTCheck + Val("" & rsA("支票").Value)
            dblTTOT = dblTTOT + Val("" & rsA("現金").Value) + Val("" & rsA("支票").Value)
            'add by nick 2004/08/26
            dblPoint = dblPoint + Val("" & rsA("點數").Value)
            dblTPoint = dblTPoint + Val("" & rsA("點數").Value)
            
            
            ii = ii + 1
            'adoTaie.Execute "Insert Into ACCRPT430 Values('" & strUserNum & "', " & ii & "," & IIf(IsNull(rsA("收款日").Value) = True, "Null", rsA("收款日").Value) & ",'" & rsA("收款人").Value & "','" & rsA("收據抬頭").Value & "','" & strCaseData & IIf(strAmt = "0", "", strAmt) & "','" & _
                                rsA("人工號").Value & "','" & rsA("電腦號").Value & "'," & Val("" & rsA("現金").Value) & "," & Val("" & rsA("支票").Value) & "," & IIf(IsNull(rsA("到期日").Value) = True, "Null", rsA("到期日").Value) & "," & Val("" & rsA("扣繳額").Value) & "," & Val(strPoint) & " ) "
            'edit by nick 2004/10/07
            'adoTaie.Execute "Insert Into ACCRPT430 Values('" & strUserNum & "', " & ii & "," & IIf(IsNull(rsA("收款日").Value) = True, "Null", rsA("收款日").Value) & ",'" & rsA("收款人").Value & "','" & rsA("收據抬頭").Value & "','" & strCaseData & IIf(strAmt = "0", "", strAmt) & "','" & _
                                rsA("人工號").Value & "','" & rsA("電腦號").Value & "'," & Val("" & rsA("現金").Value) & "," & Val("" & rsA("支票").Value) & "," & IIf(IsNull(rsA("到期日").Value) = True, "Null", rsA("到期日").Value) & "," & Val("" & rsA("扣繳額").Value) & "," & Val(strPoint) & " ,'" & strA3124 & "') "
            '2013/8/30 modify by sonia 102/8/2收款之E10207686有12筆收文號,strCaseData寫入工作檔太大
            'adoTaie.Execute "Insert Into ACCRPT430 Values('" & strUserNum & "', " & ii & "," & IIf(IsNull(rsA("收款日").Value) = True, "Null", rsA("收款日").Value) & ",'" & rsA("收款人").Value & "','" & rsA("收據抬頭").Value & "','" & convForm(strCaseData, 100) & IIf(strAmt = "0", "", strAmt) & "','" & _
                                rsA("人工號").Value & "','" & rsA("電腦號").Value & "'," & Val("" & rsA("現金").Value) & "," & Val("" & rsA("支票").Value) & "," & IIf(IsNull(rsA("到期日").Value) = True, "Null", "'" & ChangeTStringToTDateString("" & rsA("到期日").Value) & "'") & "," & Val("" & rsA("扣繳額").Value) & "," & Val("" & rsA("點數").Value) & " ,'" & strA3124 & "') "
            adoTaie.Execute "Insert Into ACCRPT430 Values('" & strUserNum & "', " & ii & "," & IIf(IsNull(rsA("收款日").Value) = True, "Null", rsA("收款日").Value) & ",'" & rsA("收款人").Value & "','" & rsA("收據抬頭").Value & "','" & Trim(convForm(strCaseData, 90)) & IIf(strAmt = "0", "", strAmt) & "','" & _
                                rsA("人工號").Value & "','" & rsA("電腦號").Value & "'," & Val("" & rsA("現金").Value) & "," & Val("" & rsA("支票").Value) & "," & IIf(IsNull(rsA("到期日").Value) = True, "Null", "'" & ChangeTStringToTDateString("" & rsA("到期日").Value) & "'") & "," & Val("" & rsA("扣繳額").Value) & "," & Val("" & rsA("點數").Value) & " ,'" & strA3124 & "') "
            rsA.MoveNext
            If rsA.EOF = False Then
                strNo = "" & rsA("人工號").Value & rsA("電腦號").Value
                strCaseData = IIf(strOurCaseNo <> "" & rsA("本所案號").Value, ReConBNOurCaseNO("" & rsA("本所案號").Value), "") & rsA("案件性質名稱").Value
                If strOurCaseNo <> "" & rsA("本所案號").Value Then
                    strOurCaseNo = "" & rsA("本所案號").Value
                End If
                strAmt = Val("" & rsA("A0J09").Value) + Val("" & rsA("A0J10").Value)
                strPoint = Val("" & rsA("點數").Value)
                rsA.MoveNext
            End If
            GoTo NextRec1
NextRec:
            rsA.MoveNext
NextRec1:
        Loop
        rsA.MoveLast
        'edit by nick 2004/08/20
        'If strSalesNo <> "" & rsA.Fields("A0K20").Value Then
        If strSalesNo <> "" & rsA.Fields("A3121").Value Then
            ii = ii + 1
            'edit by nick 2004/08/20 加欄位
            'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ) "
            'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ,'" & strA3124 & "') "
            'edit by nick 2004/10/07
            'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011,r43013,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ," & dblPoint & " ,'" & strA3124 & "') "
            adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43008, R43009, R43010, R43011,r43013,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & ",'" & Format(dblTot, "###,###,###,##0") & "'," & dblPoint & " ,'" & strA3124 & "') "
            dblCash = 0: dblCheck = 0: dblTot = 0
            'add by nick 2004/08/26
            dblPoint = 0
            'edit by nick 2004/08/20
            'strSalesNo = "" & rsA.Fields("A0K20").Value
            strSalesNo = "" & rsA.Fields("A3121").Value
        End If
        dblCash = dblCash + Val("" & rsA("現金").Value)
        dblCheck = dblCheck + Val("" & rsA("支票").Value)
        dblTot = dblTot + Val("" & rsA("現金").Value) + Val("" & rsA("支票").Value)
        dblTCash = dblTCash + Val("" & rsA("現金").Value)
        dblTCheck = dblTCheck + Val("" & rsA("支票").Value)
        dblTTOT = dblTTOT + Val("" & rsA("現金").Value) + Val("" & rsA("支票").Value)
        'add by nick 2004/08/26
        dblPoint = dblPoint + Val("" & rsA("點數").Value)
        dblTPoint = dblTPoint + Val("" & rsA("點數").Value)
        ii = ii + 1
        'edit by nick 2004/08/20 加欄位
        'adoTaie.Execute "Insert Into ACCRPT430 Values('" & strUserNum & "', " & ii & "," & IIf(IsNull(rsA("收款日").Value) = True, "Null", rsA("收款日").Value) & ",'" & rsA("收款人").Value & "','" & rsA("收據抬頭").Value & "','" & strCaseData & IIf(strAmt = "0", "", strAmt) & "','" & _
                            rsA("人工號").Value & "','" & rsA("電腦號").Value & "'," & Val("" & rsA("現金").Value) & "," & Val("" & rsA("支票").Value) & "," & IIf(IsNull(rsA("到期日").Value) = True, "Null", rsA("到期日").Value) & "," & Val("" & rsA("扣繳額").Value) & "," & Val(strPoint) & " ) "
        'edit by nick 2004/10/07
        'adoTaie.Execute "Insert Into ACCRPT430 Values('" & strUserNum & "', " & ii & "," & IIf(IsNull(rsA("收款日").Value) = True, "Null", rsA("收款日").Value) & ",'" & rsA("收款人").Value & "','" & rsA("收據抬頭").Value & "','" & strCaseData & IIf(strAmt = "0", "", strAmt) & "','" & _
                            rsA("人工號").Value & "','" & rsA("電腦號").Value & "'," & Val("" & rsA("現金").Value) & "," & Val("" & rsA("支票").Value) & "," & IIf(IsNull(rsA("到期日").Value) = True, "Null", rsA("到期日").Value) & "," & Val("" & rsA("扣繳額").Value) & "," & Val(strPoint) & " ,'" & strA3124 & "') "
        adoTaie.Execute "Insert Into ACCRPT430 Values('" & strUserNum & "', " & ii & "," & IIf(IsNull(rsA("收款日").Value) = True, "Null", rsA("收款日").Value) & ",'" & rsA("收款人").Value & "','" & rsA("收據抬頭").Value & "','" & strCaseData & IIf(strAmt = "0", "", strAmt) & "','" & _
                            rsA("人工號").Value & "','" & rsA("電腦號").Value & "'," & Val("" & rsA("現金").Value) & "," & Val("" & rsA("支票").Value) & "," & IIf(IsNull(rsA("到期日").Value) = True, "Null", "'" & ChangeTStringToTDateString("" & rsA("到期日").Value) & "'") & "," & Val("" & rsA("扣繳額").Value) & "," & Val("" & rsA("點數").Value) & " ,'" & strA3124 & "') "
        ii = ii + 1
        'edit by nick 2004/08/20 加欄位
        'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ) "
        'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ,'" & strA3124 & "') "
        'edit by nick 2004/10/07
        'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011,r43013,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & "," & dblTOT & " ," & dblPoint & " ,'" & strA3124 & "') "
        adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43008, R43009, R43010, R43011,r43013,R43014) Values('" & strUserNum & "', " & ii & ",'小計'," & dblCash & "," & dblCheck & ",'" & Format(dblTot, "###,###,###,##0") & "'," & dblPoint & " ,'" & strA3124 & "') "
        dblCash = 0: dblCheck = 0: dblTot = 0
        ii = ii + 1
        'edit by nick 2004/08/20 加欄位
        'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011) Values('" & strUserNum & "', " & ii & ",'總計'," & dblTCash & "," & dblTCheck & "," & dblTTOT & " ) "
        'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011,R43014) Values('" & strUserNum & "', " & ii & ",'總計'," & dblTCash & "," & dblTCheck & "," & dblTTOT & " ,'" & strA3124 & "') "
        'edit by nick 2004/10/07
        'adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43005, R43009, R43010, R43011,r43013,R43014) Values('" & strUserNum & "', " & ii & ",'總計'," & dblTCash & "," & dblTCheck & "," & dblTTOT & " ," & dblTPoint & " ,'" & strA3124 & "') "
        adoTaie.Execute "Insert Into ACCRPT430(R43001, R43002, R43008, R43009, R43010, R43011,r43013,R43014) Values('" & strUserNum & "', " & ii & ",'總計'," & dblTCash & "," & dblTCheck & ",'" & Format(dblTTOT, "###,###,###,##0") & "'," & dblTPoint & " ,'" & strA3124 & "') "
        'add by nick 2004/08/26
        dblTPoint = 0
        dblTCash = 0: dblTCheck = 0: dblTTOT = 0
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    If adoadodc1.State = adStateOpen Then
        adoadodc1.Close
    End If
    adoadodc1.CursorLocation = adUseClient
    StrSQLa = "Select * From ACCRPT430 Where R43001='" & strUserNum & "' Order By 2 "
    adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    Adodc1.Recordset.Requery
    
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

Private Function ReConBNOurCaseNO(strCaseNo As String) As String
If strCaseNo <> "" Then
    ReConBNOurCaseNO = Replace(Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 3), 6) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 2), 1) & "-" & Right(strCaseNo, 2), "-0-00", "")
Else
    ReConBNOurCaseNO = ""
End If
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
