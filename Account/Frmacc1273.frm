VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1273 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "進項發票資料"
   ClientHeight    =   4755
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8801.03
   Begin VB.TextBox Text5 
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
      Left            =   3824
      TabIndex        =   16
      Top             =   960
      Width           =   1692
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
      Height          =   315
      Left            =   1320
      TabIndex        =   15
      Top             =   960
      Width           =   1692
   End
   Begin VB.TextBox Text7 
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
      Left            =   1320
      TabIndex        =   14
      Top             =   600
      Width           =   2588
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
      Height          =   315
      Left            =   5160
      TabIndex        =   13
      Top             =   600
      Width           =   1792
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
      Height          =   315
      Left            =   6960
      TabIndex        =   12
      Top             =   255
      Width           =   1493
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
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   11
      Top             =   255
      Width           =   492
   End
   Begin VB.TextBox Text6 
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
      Left            =   6120
      TabIndex        =   9
      Top             =   960
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
      Height          =   315
      Index           =   1
      Left            =   1845
      TabIndex        =   7
      Top             =   255
      Width           =   1792
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1273.frx":0000
      Height          =   3000
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5292
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "a4509"
         Caption         =   "項次"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "@"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "a4502"
         Caption         =   "格式代號"
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
         DataField       =   "a4503"
         Caption         =   "發票日期"
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
         DataField       =   "a4504"
         Caption         =   "發票號碼"
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
         DataField       =   "a4505"
         Caption         =   "銷售人統編"
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
         DataField       =   "a4506"
         Caption         =   "扣抵代號"
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
      BeginProperty Column06 
         DataField       =   "a4507"
         Caption         =   "銷售額"
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
         DataField       =   "a4508"
         Caption         =   "營業稅"
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
         DataField       =   "TotalAmt"
         Caption         =   "發票總額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   254
         BeginProperty Column00 
            ColumnWidth     =   557.623
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   512.626
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   994.495
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1250.808
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1295.805
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   512.626
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   904.5
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   904.5
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1099.868
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   1440
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4680
      TabIndex        =   17
      Top             =   255
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
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
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "項次"
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
      Left            =   5640
      TabIndex        =   8
      Top             =   960
      Width           =   498
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "發票總額"
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
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "扣抵代號"
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
      Left            =   345
      TabIndex        =   5
      Top             =   615
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "格式代號"
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
      Left            =   345
      TabIndex        =   4
      Top             =   255
      Width           =   985
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "發票日期"
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
      Left            =   3705
      TabIndex        =   3
      Top             =   255
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
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
      Left            =   5985
      TabIndex        =   2
      Top             =   255
      Width           =   985
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1300
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "營業稅"
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
      Left            =   3105
      TabIndex        =   1
      Top             =   960
      Width           =   737
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "銷售人統編"
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
      Left            =   3981
      TabIndex        =   0
      Top             =   615
      Width           =   1200
   End
End
Attribute VB_Name = "Frmacc1273"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/02/22 Form2.0已修改 (無需修改)
Option Explicit
Public adoAcc450 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset

Dim strSql As String
'Dim strTp() As String
Dim strSeqNo As String
Dim ii As Integer
Dim strA4501 As String  '前畫面公司別

Private Sub Form_Load()
    Dim intX As Integer
    Dim intY As Integer
    Dim sglWidth As Single
    Dim sglHeight As Single
   
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Width = 8850
    Me.Height = 5130
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath1)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    strA4501 = Frmacc1271.Text1
    OpenTable
    FormShow
    If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
         MaskEdBox1.Mask = ""
         MaskEdBox1.Text = ""
         MaskEdBox1.Mask = DFormat
    End If
    tool3_enabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   tool3_enabled
   Frmacc1271.Enabled = True
   Set Frmacc1273 = Nothing
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
    If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSeqNo = Adodc1.Recordset.Fields("a4509").Value
   FormShow
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoAcc450.CursorLocation = adUseClient
 
   strSql = "Select * From Acc450 Where a4501='" & strA4501 & "' Order by a4509"
   adoAcc450.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   adoadodc2.CursorLocation = adUseClient
   strSql = "Select a4502,a4503,a4504,a4505,a4506,a4507,a4508,a4507+a4508 as TotalAmt,to_char(a4509,'009') as a4509 From Acc450 " & _
                 "Where a4501='" & strA4501 & "' Order by a4509"
   adoadodc2.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc2
 
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub AdodcRefresh()
On Error GoTo Checking
     If adoadodc2.State = adStateOpen Then
        adoadodc2.Close
     End If
     adoadodc2.CursorLocation = adUseClient
   
     strSql = "Select a4502,a4503,a4504,a4505,a4506,a4507,a4508,a4507+a4508 as TotalAmt,to_char(a4509,'009') as a4509 From Acc450 " & _
                 "Where a4501='" & strA4501 & "' Order by a4509"
     adoadodc2.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
     Adodc1.Recordset.Requery
     If Adodc1.Recordset.RecordCount <> 0 Then
        Adodc1.Recordset.Find "a4509=" & Val(strSeqNo), 0, adSearchForward, 1
        If Adodc1.Recordset.EOF Then
            Exit Sub
        Else
            DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
        End If
     End If
     strSeqNo = MsgText(601)
     
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub FormShow()
    Text1(0) = Adodc1.Recordset.Fields("a4502")
    Text1(1) = GetFormatName(Text1(0))
    MaskEdBox1.Mask = MsgText(601)
    If IsNull(Adodc1.Recordset.Fields("a4503").Value) Then
       MaskEdBox1.Text = MsgText(601)
    Else
       MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a4503").Value)
    End If
    Text2 = Adodc1.Recordset.Fields("a4504")
     If IsNull(Adodc1.Recordset.Fields("a4506")) Then
        Text7 = MsgText(601)
     Else
        Text7 = GetFormatName1(Adodc1.Recordset.Fields("a4506").Value)
     End If
     Text3 = "" & Adodc1.Recordset.Fields("a4505")
     Text5 = Adodc1.Recordset.Fields("a4508")
     Text4 = Val(Adodc1.Recordset.Fields("a4507")) + Val(Text5) '發票總額
     Text6 = Adodc1.Recordset.Fields("a4509") '序號
End Sub

Private Function GetFormatName(ByVal strNo As String) As String
    Select Case strNo
        Case "21"
            GetFormatName = "三聯式"
        Case "22"
            GetFormatName = "二聯式"
        Case "23"
            GetFormatName = "三聯式進項退出"
        Case "24"
            GetFormatName = "二聯式進項退出"
        Case "25"
            GetFormatName = "收銀機發票"
    End Select
End Function

Private Function GetFormatName1(ByVal strNo As String) As String
    Select Case strNo
        Case "1"
            GetFormatName1 = "可扣抵進項費用"
        Case "2"
            GetFormatName1 = "可扣抵進項固定資產"
        Case "3"
            GetFormatName1 = "不可扣抵進項費用"
        Case "4"
            GetFormatName1 = "不可扣抵進項固定資"
        Case Else
            GetFormatName1 = ""
    End Select
    If strNo <> MsgText(601) Then
        GetFormatName1 = strNo & "-" & GetFormatName1
    End If
End Function





