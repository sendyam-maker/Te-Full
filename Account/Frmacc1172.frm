VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1172 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "進項發票明細維護"
   ClientHeight    =   4764
   ClientLeft      =   36
   ClientTop       =   312
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
   Begin VB.TextBox Text6 
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
      Height          =   300
      Left            =   6120
      TabIndex        =   19
      Top             =   960
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7200
      Picture         =   "Frmacc1172.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   17
      ToolTipText     =   "清除畫面"
      Top             =   720
      Width           =   550
   End
   Begin VB.CommandButton CmdDel 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7863
      Picture         =   "Frmacc1172.frx":08CA
      Style           =   1  '圖片外觀
      TabIndex        =   7
      ToolTipText     =   "取消"
      Top             =   720
      Width           =   549
   End
   Begin VB.TextBox Text4 
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
      Height          =   315
      Left            =   1320
      MaxLength       =   14
      TabIndex        =   5
      Top             =   960
      Width           =   1692
   End
   Begin VB.TextBox Text1 
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
      Height          =   315
      Index           =   1
      Left            =   1845
      TabIndex        =   16
      Top             =   255
      Width           =   1792
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      MaxLength       =   8
      TabIndex        =   4
      Top             =   600
      Width           =   1792
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1305
      TabIndex        =   3
      Top             =   585
      Width           =   2588
   End
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
      Height          =   315
      Index           =   0
      Left            =   1305
      MaxLength       =   2
      TabIndex        =   0
      Top             =   255
      Width           =   547
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6945
      MaxLength       =   10
      TabIndex        =   2
      Top             =   255
      Width           =   1493
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
      Height          =   315
      Left            =   3825
      MaxLength       =   12
      TabIndex        =   6
      Top             =   945
      Width           =   1692
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1172.frx":0F34
      Height          =   3000
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14626
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
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
            ColumnWidth     =   554.775
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   518.322
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   988.799
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1253.656
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1290.109
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   518.322
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
            ColumnWidth     =   1097.02
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Top             =   255
      Width           =   1204
      _ExtentX        =   2117
      _ExtentY        =   572
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
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "項次"
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
      Left            =   5640
      TabIndex        =   18
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
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "扣抵代號"
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
      Left            =   345
      TabIndex        =   14
      Top             =   615
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "格式代號"
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
      Left            =   345
      TabIndex        =   12
      Top             =   255
      Width           =   985
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "發票日期"
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
      Left            =   3705
      TabIndex        =   11
      Top             =   255
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
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
      Left            =   5985
      TabIndex        =   10
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
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3105
      TabIndex        =   9
      Top             =   960
      Width           =   737
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "銷售人統編"
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
      Left            =   3981
      TabIndex        =   8
      Top             =   615
      Width           =   1200
   End
End
Attribute VB_Name = "Frmacc1172"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 (無需修改)
'Create by Amy 2013/12/27 與frmacc1170及frmacc4120及frm41c0(103/07/09增加)共用
Option Explicit
Public adoAcc450 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Const SetCombo1 As String = ",可扣抵進項費用,可扣抵進項固定資產,不可扣抵進項費用,不可扣抵進項固定資產"

Dim strSql As String
Dim strTp() As String
'Dim bolIsFirst As Boolean   'Mark by Amy 2015/06/15 '是否為第一次進入(前畫面進入Insert後再回到此頁不算第一次)
Dim strSeqNo As String '2014/01/10
Dim ii As Integer
Dim strA4501 As String  '前畫面公司別
Public strBackForm As String 'Modify by Amy 2014/01/15

Private Sub cmdDel_Click()
    Dim strDel As String
On Error GoTo Checking

    If Adodc1.Recordset.RecordCount <> 0 Then
        strDel = "Delete From Acc450 Where A4501 = '" & strA4501 & "' And a4509 = " & Val(Text6)
        adoTaie.Execute strDel
    End If
    AdodcRefresh
    FormClear
    'Text2.Locked = False
    strSeqNo = MsgText(601) 'Add byAmy 2014/01/10
    Text1(0).SetFocus 'Add by Amy 2014/01/08
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.Text = "" Then
        Exit Sub
    Else
        If Val(Mid(Combo1, 1, 1)) = 0 Or Val(Mid(Combo1, 1, 1)) > 4 Then
                MsgBox Label4 & "輸入錯誤,請確認！", , MsgText(5)
                Cancel = True
                Exit Sub
        End If
    End If
End Sub

'Add by Amy 2014/01/08
Private Sub Command2_Click()
    FormClear
    'Text2.Locked = False
    strSeqNo = MsgText(601) 'Add byAmy 2014/01/10
    Text1(0).SetFocus
End Sub
'end 2014/01/08

Private Sub DataGrid1_SelChange(Cancel As Integer)
    If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSeqNo = Adodc1.Recordset.Fields("a4509").Value
   FormShow
   'Text2.Locked = True
End Sub

Private Sub Form_Activate()
    strFormName = Name
    If strSaveConfirm <> MsgText(601) Then
        Text1(0).SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Add  by Amy 2014/07/09
    If strBackForm = "Frmacc41c0" Then
        KeyDefine KeyCode
        Exit Sub
    End If
    'end 2014/07/09
    
    '若未判斷strSaveConfirm = MsgText(601) 跳離，則run KeyDefine KeyCode會錯
    If strSaveConfirm = MsgText(601) Then
        Exit Sub
    End If
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
    'Add by Amy 2014/01/15
    Select Case strBackForm
        Case "Frmacc1170"
           strA4501 = Frmacc1170.Text1
        Case "Frmacc4120"
            strA4501 = Frmacc4120.Text2
        'Add by Amy 2014/07/09
        Case "Frmacc41c0"
            strA4501 = Frmacc41c0.Text2
        Case Else
    End Select
    'end 2014/01/15
    OpenTable
    '設定扣抵代號
    strTp = Split(SetCombo1, ",")
    For ii = 0 To UBound(strTp)
        If ii = 0 Then
            Combo1.AddItem strTp(ii)
        Else
            Combo1.AddItem ii & "-" & strTp(ii)
        End If
    Next
    Combo1 = MsgText(601)
    If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
         MaskEdBox1.Mask = ""
         MaskEdBox1.Text = ""
         MaskEdBox1.Mask = DFormat
    End If
    'Modify by Amy 2014/01/15 +廠商編號不為F且統編是8碼才預帶
    If strBackForm = "Frmacc1170" Then
        If strSaveConfirm <> MsgText(601) And Left(Frmacc1170.Text2, 1) <> "F" And Len(Frmacc1170.Text21) = 8 Then
            Text3 = Frmacc1170.Text21
        End If
    End If
    'end 2014/01/15
    FormEnabled
    tool3_enabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case strBackForm
        Case "Frmacc1170"
            'Add by Amy 2014/01/10
            'Modify by Amy 2015/06/15 改應付款進入不論新增或修改都存acc10,故改判斷
            'If bolIsFirst = True Then
                '只有應付款資料第一次新增且第一次進入表單才需存 acc1p0
            If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
                Acc1p0Save
            End If
            'end 2015/06/15
            'end 2014/01/10
            'Add by Amy 2018/09/13 帶回第一筆明細(因G10700329 只frm1170畫面之發票號未更新明細,造成資料不一致)
            If (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) And Adodc1.Recordset.RecordCount > 0 Then
                Adodc1.Recordset.MoveFirst
                DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
                Frmacc1170.bolBack = True
                Frmacc1170.Text4 = DataGrid1.Columns(3)
            End If
            'end 2018/09/13
            Frmacc1170.Show
        Case "Frmacc4120"
            Frmacc4120.Show
        'Add by Amy 2014/07/09
        Case "Frmacc41c0"
            Frmacc41c0.Show
        Case Else
    End Select
    'Modify by Amy 2014/07/09 +strBackForm判斷
    If strBackForm = "Frmacc41c0" Then
        tool6_enabled
    ElseIf strSaveConfirm = MsgText(601) Then
        tool1_enabled
    Else
         tool2_enabled
    End If
    
    StatusClear
    Set Frmacc1172 = Nothing
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
  
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoAcc450.CursorLocation = adUseClient
   adoAcc450.MaxRecords = intMax
   'Modify by Amy2014/01/10 改order by 原:a4504
   strSql = "Select * From Acc450 Where a4501='" & strA4501 & "' Order by a4509"
   adoAcc450.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   adoadodc2.CursorLocation = adUseClient
   strSql = "Select a4502,a4503,a4504,a4505,a4506,a4507,a4508,a4507+a4508 as TotalAmt,to_char(a4509,'009') as a4509 From Acc450 " & _
                 "Where a4501='" & strA4501 & "' Order by a4509"
   adoadodc2.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc2
   
   'Mark by Amy 2015/06/15 應付款不論新增、修改都重產生acc1p0
'   If strBackForm = "Frmacc4120" Then
'        bolIsFirst = False '傳票輸入不需產生acc1p0
'   Else
'        If adoadodc2.RecordCount = 0 And strSaveConfirm = MsgText(3) Then
'            bolIsFirst = True  '應付款資料第一次進入需產生acc1p0
'        Else
'            bolIsFirst = False
'        End If
'   End If
    ' end 2015/06/15

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
On Error GoTo Checking

    If KeyCode = vbKeyInsert Then
        If TxtValidate = False Then Exit Sub
        FormSave
        'Modify by Amy 2014/01/10 Acc1p0Save 改至Form Unload做
        AdodcRefresh
        FormClear
        'Text2.Locked = False
        Text1(0).SetFocus
        Text1_GotFocus (0)
        Frmacc0000.StatusBar1.Panels(1).Text = MsgText(17)
    Else
        Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
    End If
   KeyEnter KeyCode
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgBox(5)
End Sub

Private Function TxtValidate() As Boolean
    Dim bCancel As Boolean
    bCancel = False
    TxtValidate = False
    If Trim(Text1(0)) = MsgText(601) Then
        MsgBox MsgText(10) & Label1, , MsgText(5)
        Text1(0).SetFocus
        Exit Function
    Else
        Call Text1_Validate(0, bCancel)
        If bCancel = True Then
            Text1(1) = ""
            MsgBox Label1 & "輸入錯誤,請確認！", , MsgText(5)
            Text1(0).SetFocus
            Exit Function
        End If
    End If
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        MsgBox MsgText(10) & Label2, , MsgText(5)
        MaskEdBox1.SetFocus
        Exit Function
    Else
        Call MaskEdBox1_Validate(bCancel)
        If bCancel = True Then
            Exit Function
        End If
    End If
    If Trim(Text2) = MsgText(601) Then
        MsgBox MsgText(10) & Label3, , MsgText(5)
        Text2.SetFocus
        Exit Function
    Else
        If Len(Trim(Text2)) <> 10 Then
            MsgBox Label3 & "位數不足10碼 ！", , MsgText(5)
            Text2.SetFocus
            Exit Function
        End If
        Call Text2_Validate(bCancel)
        If bCancel = True Then
            Exit Function
        End If
        'Add by Amy 2014/03/10 +同一個付款單號的發票號碼不可重覆
        'Modify by Amy 2019/09/16 改一年內發票號碼不可重覆,新增時不可存檔,修改時提示 ex:G10800400/409 客戶提供二次同一張發票,付了二次款
        'If ChkA4504 = True And Text6 = "" Then
        If ChkA4504(Text2) = True Then
            If Text6 = "" Then
                MsgBox Label3 & "不可重覆！", , MsgText(5)
                Text2.SetFocus
                Exit Function
            ElseIf MsgBox(Label3 & "重覆要繼續操作？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                Exit Function
            End If
        End If
        'end 2019/09/16
        'end 2014/03/10
    End If
    If Trim(Text3) = MsgText(601) Then
        MsgBox Label5 & "不可為空！", , MsgText(5)
        Text3.SetFocus
        Exit Function
    End If
    Call Text3_Validate(bCancel)
    If bCancel = True Then
        Exit Function
    End If
    If Combo1 = MsgText(601) Then
        MsgBox MsgText(10) & Label4, , MsgText(5)
        Combo1.SetFocus
        Exit Function
    Else
        Call Combo1_Validate(bCancel)
        If bCancel = True Then
            Combo1.SetFocus
            Exit Function
        End If
    End If
    If Text4 = MsgText(601) Or Val(Text4) = 0 Then
         MsgBox MsgText(58) & Label6, , MsgText(5)
         Text4.SetFocus
         Exit Function
    End If
    If Text5 = MsgText(601) Or Val(Text5) = 0 Then
         MsgBox MsgText(58) & Label8, , MsgText(5)
         Text5.SetFocus
         Exit Function
    End If
    TxtValidate = True
End Function

Private Sub FormSave()
    If adoAcc450.RecordCount = 0 Then
        If adoAcc450.EOF Then
            adoAcc450.AddNew
            Text6 = GetSerialNo("Select MAX(A4509) From Acc450 Where A4501 = '" & strA4501 & "' ", 3)
        End If
    Else
        adoAcc450.MoveFirst
        adoAcc450.Find "a4509=" & Val(Text6), 0, adSearchForward, 1
        If adoAcc450.EOF Then
           adoAcc450.AddNew
           Text6 = GetSerialNo("Select MAX(A4509) From Acc450 Where A4501 = '" & strA4501 & "' ", 3)
        End If
    End If
    
    adoAcc450.Fields("a4501").Value = strA4501
    adoAcc450.Fields("a4509").Value = Val(Text6)
    adoAcc450.Fields("a4504").Value = Text2
    adoAcc450.Fields("a4502").Value = Text1(0)
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         adoAcc450.Fields("a4503").Value = Val(FCDate(MaskEdBox1.Text))
    Else
         adoAcc450.Fields("a4503").Value = Null
    End If
    If Text3 <> MsgText(601) Then
         adoAcc450.Fields("a4505").Value = Text3
    Else
         adoAcc450.Fields("a4505").Value = Null
    End If
    If Combo1 <> MsgText(601) Then
        adoAcc450.Fields("a4506").Value = Mid(Combo1, 1, 1)
    Else
        adoAcc450.Fields("a4506").Value = Null
    End If
    adoAcc450.Fields("a4507").Value = Val(Text4) - Val(Text5)
    adoAcc450.Fields("a4508").Value = Val(Text5)
   adoAcc450.UpdateBatch

End Sub

Private Sub Acc1p0Save()
    '只有frmacc1170進入才需寫入Acc1p0
    Dim strSave As String
    Dim SeqNo As String    '2014/01/10 改Acc450 序號用
    Dim strMemo As String 'Add by Amy 2014/01/07
    'Add by Amy 2014/01/10
    Dim StrSQLa As String
    Dim intR As Integer
    Dim TotalAmt As Double
    Dim intA1p18 As Long 'Add by Amy 2014/01/14
    'Add by Amy 2015/06/15
    Dim RsQ As New ADODB.Recordset, StrSqlB As String, nowSeq As Integer
    Dim intQ As Integer, jj As Integer, intNowRec As Integer
    
    'Modify by Amy 2014/01/10 +抓acc450跑迴圈 依Acc450的序號由小至大新增至Acc1p0
    'SeqNo = GetSerialNo("Select MAX(a1p03) From acc1p0 Where a1p01 = '" & Frmacc1170.Text19 & "' And a1p02 = 'B' And a1p04 = '" & Frmacc1170.Text1 & "'", 3)
    'strMemo = Frmacc1170.Text3 & "/" & Text2 & " " & Text4
    'strSave = "Insert Into Acc1P0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p14,a1p15) " & _
                    "Values ('" & ChgSQL(Frmacc1170.Text19) & "','B','" & SeqNo & "','" & ChgSQL(Frmacc1170.Text1) & "','1211','TOT'," & Val(Text4) & ",'" & ChgSQL(strMemo) & "','" & ChgSQL(Frmacc1170.Text2) & "' )"
    'adoTaie.Execute strSave
    
    'Add by Amy 2015/06/15 應付款(acc1170)修改進入->新增進項->再回應付款,acc1p0因為序號控制故不會自動新增1121,需user手動增加,財務欲改自動增加
    '改成先刪除acc1p0中的1121再抓acc1p0空號填回
    StrSqlB = "Delete acc1p0 Where a1p01='" & ChgSQL(Frmacc1170.Text19) & "' And a1p02='B' And a1p04='" & ChgSQL(Frmacc1170.Text1) & "' And a1p05='1211' "
    cnnConnection.Execute StrSqlB, intQ
    
    StrSQLa = "Select * From Acc450 Where a4501='" & Frmacc1170.Text1 & "' Order by a4509"
    intR = 1
    Set RsTemp = ClsLawReadRstMsg(intR, StrSQLa)
    If intR = 1 Then
        'Add by Amy 2015/06/15 再抓acc1p0空序號填回
        intQ = 1
        StrSqlB = "Select * From Acc1p0 Where a1p01='" & ChgSQL(Frmacc1170.Text19) & "' And a1p02='B' And a1p04='" & ChgSQL(Frmacc1170.Text1) & "' Order by a1p02"
        Set RsQ = ClsLawReadRstMsg(intQ, StrSqlB)
        If intQ = 1 Then RsQ.MoveFirst
        'end 2015/06/15
        
        RsTemp.MoveFirst
        ii = 0:  intNowRec = 0: nowSeq = 1
        Do While Not RsTemp.EOF
            'Modify by Amy 2015/06/15 判斷序號,acc1p0空序號填回
            If intQ = 1 Then
                For jj = intNowRec To RsQ.RecordCount - 1
                    If Val(RsQ.Fields("a1p03")) = nowSeq Then
                        nowSeq = ZeroBeforeNo("" & nowSeq, 3)
                        RsQ.MoveNext
                    ElseIf Val(RsQ.Fields("a1p03")) > nowSeq Then
                        SeqNo = String(3 - Len("" & nowSeq), "0") & nowSeq
                        nowSeq = ZeroBeforeNo("" & nowSeq, 3)
                        intNowRec = jj
                        Exit For
                    Else
                        RsQ.MoveNext
                    End If
                Next jj
                If jj > RsQ.RecordCount - 1 Then SeqNo = String(3 - Len("" & nowSeq), "0") & nowSeq
            Else
                SeqNo = ZeroBeforeNo("" & ii, 3)
            End If
            'end 2015/06/15
            TotalAmt = Val(RsTemp.Fields("a4507")) + Val(RsTemp.Fields("a4508"))
            strMemo = Frmacc1170.Text3 & "/" & RsTemp.Fields("a4504") & " " & TotalAmt '廠商名稱/發票號碼 總額
            'Add by Amy 2014/01/14 +a1p18
            If Frmacc1170.MaskEdBox1.Text <> MsgText(601) And Frmacc1170.MaskEdBox1.Text <> MsgText(29) Then
                intA1p18 = Val(FCDate(Frmacc1170.MaskEdBox1.Text))
            Else
                intA1p18 = Null
            End If
            'end 2014/01/14
            strSave = "Insert Into Acc1P0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p18) " & _
                          "Values ('" & ChgSQL(Frmacc1170.Text19) & "','B','" & SeqNo & "','" & ChgSQL(Frmacc1170.Text1) & "','1211','TOT'," & Val(RsTemp.Fields("a4508")) & ",0,'" & ChgSQL(strMemo) & "','" & ChgSQL(Frmacc1170.Text2) & "'," & intA1p18 & " )"
            adoTaie.Execute strSave
            RsTemp.MoveNext
            ii = ii + 1
        Loop
        If RsQ.RecordCount > 0 Then MsgBox "進項稅額分錄資料會重新產生,請確認資料是否正確！", vbExclamation 'Add by Amy 2015/06/15
    End If
    'end 2014/01/10
End Sub

Private Sub AdodcRefresh()
On Error GoTo Checking
     If adoadodc2.State = adStateOpen Then
        adoadodc2.Close
     End If
     adoadodc2.CursorLocation = adUseClient
     'Modify by Amy2014/01/10 改order by及find 原:a4504
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
        Combo1 = MsgText(601)
     Else
        Combo1 = Combo1.List(Val(Adodc1.Recordset.Fields("a4506").Value))
     End If
     Text3 = "" & Adodc1.Recordset.Fields("a4505")
     Text5 = Adodc1.Recordset.Fields("a4508")
     Text4 = Val(Adodc1.Recordset.Fields("a4507")) + Val(Text5) '發票總額
     Text6 = Adodc1.Recordset.Fields("a4509") 'Add by Amy 2014/01/10 +序號
End Sub

Private Sub FormEnabled()
    'Modify by Amy 2014/07/09 +strBackForm判斷
    If strSaveConfirm <> MsgText(601) Or strBackForm = "Frmacc41c0" Then
        Command2.Enabled = True 'Add by Amy 2014/01/08
        CmdDel.Enabled = True
        Text1(0).Enabled = True
        Text2.Enabled = True 'Add by Amy 2013/01/10
        MaskEdBox1.Enabled = True
        Text2.Enabled = True
        Combo1.Enabled = True
        Text3.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
         Exit Sub
    End If
    'end 2014/07/09
    
    If strSaveConfirm = MsgText(601) Then
        Command2.Enabled = False 'Add by Amy 2014/01/08
        CmdDel.Enabled = False
        Text1(0).Enabled = False
        MaskEdBox1.Enabled = False
        Text2.Enabled = False
        Combo1.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
    End If
    
End Sub

Private Sub FormClear()
    Text1(0) = ""
    Text1(1) = ""
    If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
         MaskEdBox1.Mask = ""
         MaskEdBox1.Text = ""
         MaskEdBox1.Mask = DFormat
      End If
      Text2 = ""
      Combo1 = ""
      Text4 = ""
      Text5 = ""
      Text6 = "" 'Add by Amy 2014/01/10
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

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If Val(FCDate(MaskEdBox1.Text)) > Val(strSrvDate(2)) Then
        MsgBox Label2 & MsgText(9022), , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
'Mark by Amy 2014/12/16 因進項稅額發票日可能為非工作日
'   Else
'        If ChkWorkDay(Val(FCDate(MaskEdBox1.Text)) + 19110000) = False Then
'            MsgBox Label2 & "需為工作日", , MsgText(5)
'            Cancel = True
'            MaskEdBox1.SetFocus
'            Exit Sub
'        End If
   End If
End Sub

Private Sub Text1_Change(Index As Integer)
    If Index = 0 Then
        If Text1(0) = MsgText(601) Then
            Exit Sub
        End If
        
        If Len(Trim(Text1(0))) = 2 And InStr("21,22,23,24,25", Text1(0)) > 0 Then
            'Add by Amy 2014/01/07 +由frmacc1170進入不可輸入23 or 24
            If strBackForm = "Frmacc1170" And InStr("23,24", Text1(0)) > 0 Then
                Exit Sub
            Else
                Text1(1) = GetFormatName(Text1(0))
            End If
            'end 2014/01/07
        End If
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    CloseIme
    TextInverse Text1(0)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If Text1(0) <> MsgText(601) Then
            If Len(Trim(Text1(0))) <> 2 Or Not InStr("21,22,23,24,25", Text1(0)) > 0 Then
                Text1(1) = ""
                MsgBox Label1 & "輸入錯誤,請確認！", , MsgText(5)
                Cancel = True
                Exit Sub
            'Add by Amy 2014/01/07 +elseif
            ElseIf strBackForm = "Frmacc1170" And InStr("23,24", Text1(0)) > 0 Then
                Text1(1) = ""
                MsgBox Label1 & "輸入錯誤,請確認！", , MsgText(5)
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub Text2_GotFocus()
    CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Len(Trim(Text2)) = 10 Then
        For ii = 1 To 10
            If ii > 2 Then
                If Not IsNumeric(Mid(Text2, ii, 1)) Then
                    MsgBox Label3 & "後8碼需為數字！", , MsgText(5)
                    Cancel = True
                    Exit Sub
                End If
            Else
                If IsNumeric(Mid(Text2, ii, 1)) Then
                    MsgBox Label3 & "前2碼需為文字！", , MsgText(5)
                    Cancel = True
                    Exit Sub
                End If
            End If
        Next ii
    End If
End Sub

Private Sub Text3_GotFocus()
    TextInverse Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    If Trim(Text3.Text) = "" Then Exit Sub
    If GetTextLength(Text3) <> 8 Then
         If MsgBox("統編必須是8碼 ! 請確定 ?", vbYesNo + vbCritical) = vbNo Then
            Cancel = True
            Exit Sub
         End If
    End If
    If CheckID(1, Text3) = False Then
        If MsgBox("統一編號錯誤，是否確定 ?", vbYesNo + vbCritical) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If

End Sub

Private Sub Text4_Change()
    If Len(Trim(Text4)) > 0 Then
        If Not IsNumeric(Text4) Then
            MsgBox Label6 & "輸入錯誤,請確認！", , MsgText(5)
            Exit Sub
        End If
        Text5 = Val(Text4) - (Round(Val(Text4) / 1.05, 0))
    End If
End Sub

Private Sub Text5_Change()
    If Len(Trim(Text5)) > 0 Then
        If Not IsNumeric(Text5) Then
            MsgBox Label8 & "輸入錯誤,請確認！", , MsgText(5)
            Exit Sub
        End If
    End If
End Sub

