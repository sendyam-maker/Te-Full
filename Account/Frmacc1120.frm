VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1120 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收據開立作業"
   ClientHeight    =   5120
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5120
   ScaleWidth      =   9120
   Begin VB.CommandButton Command4 
      Caption         =   "檢視接洽單"
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
      Left            =   4500
      TabIndex        =   11
      Top             =   570
      Width           =   1410
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含2年前收文資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7170
      TabIndex        =   10
      Top             =   4800
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "特殊收據"
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
      Left            =   6255
      TabIndex        =   9
      Top             =   570
      Width           =   1230
   End
   Begin VB.CommandButton cmdPtAssign 
      Caption         =   "收文點數分配"
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
      Left            =   7590
      TabIndex        =   8
      Top             =   570
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "例外處理"
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
      Left            =   6270
      TabIndex        =   1
      Top             =   270
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1120.frx":0000
      Height          =   3765
      Left            =   120
      TabIndex        =   2
      Top             =   1000
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   6632
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "a0j06"
         Caption         =   "選取"
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
         DataField       =   "a0j07"
         Caption         =   "合併"
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
         DataField       =   "cp13"
         Caption         =   "智權人員"
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
         DataField       =   "a0j02"
         Caption         =   "本所案號"
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
         DataField       =   "cp10N"
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
      BeginProperty Column05 
         DataField       =   "na03"
         Caption         =   "申請國家"
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
         DataField       =   "PA161"
         Caption         =   "出"
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
      BeginProperty Column07 
         DataField       =   "crl119"
         Caption         =   "特"
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
         DataField       =   "a0j08"
         Caption         =   "手開收據"
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
      BeginProperty Column09 
         DataField       =   "a0j09"
         Caption         =   "服務費"
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
      BeginProperty Column10 
         DataField       =   "a0j10"
         Caption         =   "規費"
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
      BeginProperty Column11 
         DataField       =   "cp5354"
         Caption         =   "年費年度"
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
      BeginProperty Column12 
         DataField       =   "Rdate"
         Caption         =   "收文日期"
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
      BeginProperty Column13 
         DataField       =   "cp140"
         Caption         =   "接洽紀錄單號"
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
      BeginProperty Column14 
         DataField       =   "crl49"
         Caption         =   "接洽單收據公司"
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
         Size            =   275
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1239.874
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   789.732
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   260.22
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   260.22
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1039.748
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   120
      Top             =   900
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
   Begin VB.CommandButton Command1 
      Caption         =   "收據抬頭輸入"
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
      Left            =   7590
      TabIndex        =   3
      Top             =   270
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   0
      Top             =   210
      Width           =   1572
   End
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   2760
      TabIndex        =   5
      Top             =   210
      Width           =   3375
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "5953;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "輸完客戶編號請按Tab鍵以帶出資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   180
      Left            =   300
      TabIndex        =   7
      Top             =   600
      Width           =   3045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "若同時收文  移轉/讓與 及其他案件性質時 移轉/讓與 請最後開立！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   150
      TabIndex        =   6
      Top             =   4800
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FF0000&
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   9030
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
End
Attribute VB_Name = "Frmacc1120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0j0 As New ADODB.Recordset
Public adoacc0k0 As New ADODB.Recordset
Public adopatent As New ADODB.Recordset
Public adotrademark As New ADODB.Recordset
Public adolawcase As New ADODB.Recordset
Public adohirecase As New ADODB.Recordset
Public adoservice As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim stra0j01 As String
Dim stra0j02 As String
Dim stra0j04 As String
'Dim stra0j05 As String 'Removed by Morgan 2011/12/26 取消 a0j05
Dim stra0j07 As String
Dim stra0j11 As String
'Dim stra0j12 As String 'Removed by Morgan 2011/12/26 取消 a0j12
'Dim stra0j20 As String 'Removed by Morgan 2011/12/27 取消 a0j20
'Dim stra0j21 As String 'Removed by Morgan 2011/12/29 取消 a0j21
'Dim lnga0j03 As String 'Removed by Morgan 2011/12/26 取消 a0j03
Dim lnga0j09 As Long
Dim lnga0j10 As Long
Dim strA0K01 As String
Dim stra0k03 As String
Dim strA0K04 As String
Dim stra0k05 As String
Dim stra0k08 As String
Dim strA0K11 As String
Dim stra0k20 As String
Dim lnga0k02 As Long
Dim lnga0k06 As Long
Dim lnga0k07 As Long
Dim intY As Integer
Dim strY As String
Dim strNoNation As String
Public strMsgShow As String
'Add by Morgan 2010/12/6
Public m_AutoProcess As Boolean
Public m_CustNo As String
Public m_CP09 As String
Dim m_Assigning As Boolean 'Add by Morgan 2011/4/13
Dim m_PA161 As String, m_CRL119 As String, m_CRL49 As String, m_CP140 As String, ii As Integer 'Add By Sindy 2013/12/20
Dim m_CRL02 As String 'Add By Sindy 2020/3/31
Public m_cp10N As Boolean   'add by sonia 2014/3/18
Dim tmpfrm As Form 'Add By Sindy 2023/1/4


Private Sub cmdPtAssign_Click()
   CheckOC3
   Set AdoRecordSet3 = Adodc1.Recordset.Clone
   With AdoRecordSet3
      .Filter = "a0j06='Y'"
      If .RecordCount > 0 Then
         Frmacc11l0.m_sAssignNo = .Fields("a0j01")
         Frmacc11l0.m_sCallType = "A"
         Set Frmacc11l0.m_fCallForm = Me
         Frmacc11l0.Show
         m_Assigning = True
         Me.Visible = False
      Else
         MsgBox "請選擇要分配的收文..."
      End If
      .Filter = "" 'Added by Morgan 2023/5/10
   End With
End Sub

'Add by Morgan 2011/9/29 開收據檢查從Command1_Click抽出以便共用
Private Function CheckData() As Boolean
   'Add by Morgan 2005/9/15 顧問案的顧問聘任(0)時
   CheckOC3
   Set AdoRecordSet3 = Adodc1.Recordset.Clone
   With AdoRecordSet3
      'Modified by Morgan 2011/12/26 取消a0j03
      '.Filter = "a0j06='Y' and a0j03='0'"
      .Filter = "a0j06='Y' and cp10='0'"
      If .RecordCount > 0 Then
         .MoveFirst
         If Left("" & .Fields("A0J02"), 2) = "LA" Then
            strSql = "SELECT CP53,CP54 FROM CASEPROGRESS WHERE CP09=" & CNULL(.Fields("A0J01"))
            CheckOC3
            .Filter = ""
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
               If MsgBox("顧問期間：" & Format(TransDate("" & .Fields("CP53"), 1), "###/##/##") & " - " & Format(TransDate("" & .Fields("CP54"), 1), "###/##/##") & "，是否正確？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                  Exit Function
               End If
            End If
         End If
      End If
      .Filter = ""
   End With

   '2011/4/25 add by sonia 要智權同仁確認後或已列印定稿才可開收據
   CheckOC3
   Set AdoRecordSet3 = Adodc1.Recordset.Clone
   With AdoRecordSet3
      .Filter = "a0j06='Y'"
      If .RecordCount > 0 Then
         .MoveFirst
         If adocheck.State = adStateOpen Then
            adocheck.Close
         End If
         adocheck.CursorLocation = adUseClient
         '2011/8/25 modify by sonia
         'adocheck.Open "SELECT nvl(lc08,lc13) FROM LetterCache WHERE lc01=" & CNULL(.Fields("A0J01")), adoTaie, adOpenStatic, adLockReadOnly
'modify by sonia 2014/10/28 改為專業部列印定稿後才可開收據,否則可能收據先開但定稿還沒印
'         adocheck.Open "SELECT lc08,lc13 FROM LetterCache WHERE lc01=" & CNULL(.Fields("A0J01")), adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount >= 1 Then
'            '2011/8/25 modify by sonia
'            'If IsNull(adocheck.Fields(0)) Then
'            If IsNull(adocheck.Fields(0)) And IsNull(adocheck.Fields(1)) Then
'               MsgBox ("此收文號智權同仁尚未確認, 不可先開立收據!")
'               adocheck.Close
'               Exit Function
'            '2011/8/26 add by sonia 即使智權同仁已確認若有改金額也不可開立收據
'            Else
'               If adocheck.State = adStateOpen Then
'                  adocheck.Close
'               End If
'               adocheck.CursorLocation = adUseClient
'               adocheck.Open "SELECT * FROM LetterCachevar,caseprogress WHERE lcv06 is not null and lcv06<>cp16 and lcv05='Y' and lcv01=cp09 and lcv01=" & CNULL(.Fields("A0J01")), adoTaie, adOpenStatic, adLockReadOnly
'               If adocheck.RecordCount >= 1 Then
'                  MsgBox ("智權同仁已確認但專業部尚未修改金額, 不可先開立收據!")
'                  adocheck.Close
'                  Exit Function
'               End If
'            '2011/8/26 end
'            End If
'         End If
         'modify by sonia 2017/8/1 +lc02='0',否則會抓到年費報價資料CFP-021321
         adocheck.Open "SELECT lc13 FROM LetterCache WHERE lc01=" & CNULL(.Fields("A0J01")) & " and lc02='0'", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount >= 1 Then
            If IsNull(adocheck.Fields(0)) Then
               MsgBox ("此收文號專業部尚未列印報價定稿, 不可先開立收據!")
               adocheck.Close
               Exit Function
            '2014/11/6 add by sonia 即使已列印定稿,但專業部未改金額也不可開立收據
            Else
               If adocheck.State = adStateOpen Then
                  adocheck.Close
               End If
               adocheck.CursorLocation = adUseClient
               'modify by sonia 2017/8/1 +lc02='0',否則會抓到年費報價資料CFP-021321
               adocheck.Open "SELECT * FROM LetterCachevar,caseprogress WHERE lcv06 is not null and lcv06<>cp16 and lcv05='Y' and lcv01=cp09 and lcv01=" & CNULL(.Fields("A0J01")) & " and lcv02='0'", adoTaie, adOpenStatic, adLockReadOnly
               If adocheck.RecordCount >= 1 Then
                  MsgBox ("專業部尚未修改金額, 不可先開立收據!")
                  adocheck.Close
                  Exit Function
               End If
            'end 2014/11/6
            End If
         End If
'end 2014/10/28
         adocheck.Close
      End If
      .Filter = ""
   End With
   '2011/4/25 end
   
   'Add By Sindy 2023/9/6 ACS代收代付不可與其他案件性質一起開收據
   strExc(0) = "select cp01,cp10 from acc0j0,caseprogress" & _
      " where a0j06 = '" & MsgText(602) & "' and a0j11 = '" & Text1 & "' and a0j13=a0j01" & _
      " and cp09(+)=a0j01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If RsTemp.Fields("cp01") = "ACS" And RsTemp.Fields("cp10") = "706" Then
            If RsTemp.RecordCount > 1 Then
               MsgBox "ACS代收代付不可與其他案件性質一起開收據!!", , MsgText(5)
               Exit Function
            End If
         End If
         RsTemp.MoveNext
      Loop
   End If
   '2023/9/6 END
   
   'Added by Morgan 2019/12/19 勾選案件有未列印收據時提醒--瑞婷
   'Modified by Morgan 2020/1/7 +未結清判斷(舊資料有已結清但無列印次數 Ex:E08414378)
   'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據=> + and nvl(a0k32,'Y') <> 'Z'
   strExc(0) = "select distinct a0j02 from acc0j0,caseprogress a" & _
      " where a0j06 = '" & MsgText(602) & "' and a0j11 = '" & Text1 & "' and a0j13=a0j01" & _
      " and cp09(+)=a0j01 and exists(select * from caseprogress b,acc0j0 c,acc0k0" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04" & _
      " and c.a0j01(+)=b.cp09 and a0k01(+)=c.a0j13 and to_number(substr(a0k01, 5, 5)) > 2000 " & _
      " and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and a0k37 is null and nvl(a0k32,'Y') <> 'Z') "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         MsgBox RsTemp(0) & "案前有收據尚未列印, 請考慮新收文是否要合併！" & vbCrLf & "若要請作廢重開！", vbInformation
      Else
         strExc(1) = RsTemp.GetString()
         MsgBox "下列案號前有收據尚未列印, 請考慮新收文是否要合併！" & vbCrLf & "若要請作廢重開！" & vbCrLf & vbCrLf & strExc(1), vbInformation
      End If
   End If
   'end 2019/12/19
   
   CheckData = True
End Function

Private Sub Command1_Click()
Dim intOther As Integer 'Add By Sindy 2013/12/24
Dim bolFirst As Boolean, strCP09 As String 'Add by Amy 2016/08/18 是否為第一筆/記錄選取的第一筆總收文號
   
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   
   'Add By Sindy 2013/12/20
   m_PA161 = "": intOther = 0: m_CRL119 = "": m_CP140 = "": m_CRL49 = "": m_CRL02 = ""
   m_cp10N = False 'add by sonia 2014/3/18
   bolFirst = True 'Add by Amy 2016/08/18
   
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MoveFirst
      For ii = 1 To adoadodc1.RecordCount
         'Add by Amy 2016/08/18
         If bolFirst = True And adoadodc1.Fields("a0j06") = "Y" Then
            strCP09 = "" & adoadodc1.Fields("a0j01")
            bolFirst = False
         End If
         'end 2016/08/18
         If "" & adoadodc1.Fields("a0j06") = "Y" And "" & adoadodc1.Fields("pa161") = "J" Then
            m_PA161 = "" & adoadodc1.Fields("pa161")
         End If
         If "" & adoadodc1.Fields("a0j06") = "Y" And "" & adoadodc1.Fields("pa161") <> "J" Then
            intOther = intOther + 1
         End If
         If m_PA161 = "J" And intOther > 0 Then
            MsgBox "智權公司不可與其他公司別一起開收據!!", , MsgText(5)
            DataGrid1.SetFocus
            Exit Sub
         End If
         'Add By Sindy 2014/2/11
         If "" & adoadodc1.Fields("CP140") <> "" And "" & adoadodc1.Fields("a0j06") = "Y" Then
            m_CRL119 = "" & adoadodc1.Fields("CRL119")
            m_CP140 = "" & adoadodc1.Fields("CP140")
            m_CRL49 = "" & adoadodc1.Fields("CRL49")
            m_CRL02 = "" & adoadodc1.Fields("CRL02") 'Add By Sindy 2020/3/31
            'Add By Sindy 2023/5/19
            Call Command4_Click '檢視接洽單
            '2023/5/19 END
         End If
         '2014/2/11 END
         'add by sonia 2014/3/18
         If "" & adoadodc1.Fields("a0j06") = "Y" And "" & adoadodc1.Fields("cp10N") = "代辦退費" Then
            m_cp10N = True
         End If
         '2014/3/18 end
         adoadodc1.MoveNext
      Next ii
      adoadodc1.MoveFirst
   End If
   '2013/12/20 END
   
   If adocheck.State = adStateOpen Then
      adocheck.Close
   End If
   adocheck.CursorLocation = adUseClient
   'adocheck.Open "select count(*) from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j11 = '" & Text1 & "' and a0j13 is null", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Morgan 2011/9/19 a0j13改先放收文號
   'adocheck.Open "select * from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j11 = '" & Text1 & "' and a0j13 is null", adoTaie, adOpenStatic, adLockReadOnly
   adocheck.Open "select * from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j11 = '" & Text1 & "' and a0j13=a0j01", adoTaie, adOpenStatic, adLockReadOnly
   'Add By Sindy 2013/12/20
   If m_PA161 = "J" Then
      If adocheck.RecordCount > 5 Then
         MsgBox "一張收據不可選取超過五筆以上的收文...", , MsgText(5)
         DataGrid1.SetFocus
         adocheck.Close
         Exit Sub
      End If
   '2013/12/20 END
   ElseIf adocheck.RecordCount > 2 Then
      MsgBox MsgText(171), , MsgText(5)
      DataGrid1.SetFocus
      adocheck.Close
      Exit Sub
   End If
   adocheck.Close
   
   If CheckData = False Then Exit Sub
 
   strItemNo = ""
   strCustNo = Text1
   Me.Enabled = False
   adoacc0j0.Close
   adoacc0j0.CursorLocation = adUseClient
   'Modify by Morgan 2011/9/19 a0j13改先放收文號
   'adoacc0j0.Open "select * from acc0j0 where a0j06 = '" & MsgText(602) & "' and (a0j13 is null or a0j13 = '')", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0j0.Open "select * from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j13=a0j01", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0j0.RecordCount <> 0 Then
      If adoacc0j0.Fields("a0j08").Value = MsgText(602) Then
         Frmacc1122.Show
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
      Screen.MousePointer = vbDefault
      'Add By Sindy 2014/2/11
      Frmacc1121.strB_CP09 = strCP09 'Add by Amy 2016/08/18
      Frmacc1121.m_CRL119 = m_CRL119 '特殊收據
      Frmacc1121.m_CRL49 = m_CRL49 '收據公司
      Frmacc1121.m_CRL02 = m_CRL02 '填表日期 Add By Sindy 2020/3/31
      Frmacc1121.m_CP140 = m_CP140 '電子表單單號
      '2014/2/11 END
      Frmacc1121.Show
      'Add By Sindy 2024/3/26 將作業置前端
      Call SetFormZOrder("frm090801_Q")
      Call SetFormZOrder("frm090801_7")
      '2024/3/26 END
      
'      'Add By Sindy 2014/2/17
'      If UCase(Frmacc1121.m_CallForm) = UCase("frm090801_7") Then
'         frm090801_7.Show 'vbModal
'      End If
'      '2014/2/17 END
   Else
      MsgBox MsgText(97), , MsgText(5)
      Me.Enabled = True
      Exit Sub
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
End Sub

Private Sub Command2_Click()
   If Text1 = "" Then
      Exit Sub
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Acc0j0Delete
   Process1
End Sub

Private Sub Command3_Click()
   If Text1 = MsgText(601) Then Exit Sub
   If CheckData = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   'Modified by Morgan 2011/12/26 取消 a0j05 改抓 cp13
   'strExc(0) = "select count(distinct a0j05) from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j13=a0j01"
   strExc(0) = "select count(distinct cp13) from acc0j0,caseprogress where a0j06 = '" & MsgText(602) & "' and a0j13=a0j01 and cp09(+)=a0j01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields(0) = 0 Then
         MsgBox "尚未勾選收文資料！"
      ElseIf RsTemp.Fields(0) > 1 Then
         MsgBox "勾選收文資料必須為相同智權人員！"
      Else
         'Add By Sindy 2014/2/11
         m_CRL119 = "": m_CP140 = "": m_CRL49 = "": m_CRL02 = ""
         If adoadodc1.RecordCount > 0 Then
            adoadodc1.MoveFirst
            For ii = 1 To adoadodc1.RecordCount
               If "" & adoadodc1.Fields("CP140") <> "" And "" & adoadodc1.Fields("a0j06") = "Y" Then
                  m_CRL119 = "" & adoadodc1.Fields("CRL119")
                  m_CP140 = "" & adoadodc1.Fields("CP140")
                  m_CRL49 = "" & adoadodc1.Fields("CRL49")
                  m_CRL02 = "" & adoadodc1.Fields("CRL02") 'Add By Sindy 2020/3/31
                  'Add By Sindy 2023/5/19
                  Call Command4_Click '檢視接洽單
                  '2023/5/19 END
               End If
               adoadodc1.MoveNext
            Next ii
         End If
         adoadodc1.MoveFirst
         '2014/2/11 END
         
         Me.Enabled = False
         Set Frmacc1125.frmCallForm = Me
         'Add By Sindy 2014/2/11
         Frmacc1125.m_CRL119 = m_CRL119
         Frmacc1125.m_CRL49 = m_CRL49
         Frmacc1125.m_CRL02 = m_CRL02 'Add By Sindy 2020/3/31
         Frmacc1125.m_CP140 = m_CP140
         Screen.MousePointer = vbDefault
         '2014/2/11 END
         Frmacc1125.Show
         'Add By Sindy 2024/3/26 將作業置前端
         Call SetFormZOrder("frm090801_Q")
         Call SetFormZOrder("frm090801_7")
         '2024/3/26 END
         
'         'Add By Sindy 2014/2/17
'         If UCase(Frmacc1125.m_CallForm) = UCase("frm090801_7") Then
'            frm090801_7.Show 'vbModal
'         End If
'         '2014/2/17 END
      End If
   End If
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
End Sub

'Add By Sindy 2022/11/24 檢視檢洽單
Private Sub Command4_Click()
   Call PUB_Queryfrm090801("" & Adodc1.Recordset.Fields("cp140"), "", Me)
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   Select Case DataGrid1.col
      Case 0
'         adocheck.CursorLocation = adUseClient
'         adocheck.Open "select count(*) from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j11 = '" & Text1 & "' and a0j13 is null", adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount <> 0 Then
'            If IsNull(adocheck.Fields(0).Value) = False Then
'               If Val(adocheck.Fields(0).Value) > 2 Then
'                  adocheck.Close
'                  SendKeys "{BACKSPACE}"
'                  SendKeys "{DEL}"
'                  SendKeys "{ENTER}"
'                  MsgBox MsgText(171), , MsgText(5)
'                  DataGrid1.SetFocus
'                  Exit Sub
'               End If
'            End If
'         End If
'         adocheck.Close
   End Select
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
'   Adodc1.Recordset.UpdateBatch
End Sub

Private Sub DataGrid1_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Select Case DataGrid1.col
      Case 0
         If adocheck.State = adStateOpen Then
            adocheck.Close
         End If
         If DataGrid1.Columns(0).Text = MsgText(601) Then
            SendKeys "{Y}"
         Else
            SendKeys "{BACKSPACE}"
            SendKeys "{DEL}"
         End If
'         adocheck.CursorLocation = adUseClient
'         adocheck.Open "select count(*) from acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j11 = '" & Text1 & "' and a0j13 is null", adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount <> 0 Then
'            If IsNull(adocheck.Fields(0).Value) = False Then
'               If Val(adocheck.Fields(0).Value) > 2 Then
'                  adocheck.Close
'                  MsgBox MsgText(171), , MsgText(5)
'                  DataGrid1.SetFocus
'                  Exit Sub
'               End If
'            End If
'         End If
'         adocheck.Close
'         SendKeys "{ENTER}"
      Case 1
         If DataGrid1.Columns(1).Text = MsgText(601) Then
            SendKeys "{Y}"
         Else
            SendKeys "{BACKSPACE}"
            SendKeys "{DEL}"
         End If
      Case 6
         If DataGrid1.Columns(6).Text = MsgText(601) Then
            SendKeys "{Y}"
         Else
            SendKeys "{BACKSPACE}"
            SendKeys "{DEL}"
         End If
   End Select
'   If Adodc1.Recordset.RecordCount <> 0 Then
'      Adodc1.Recordset.UpdateBatch
'   End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case DataGrid1.col
      Case 6
      If KeyAscii = 89 Then
         intY = KeyAscii
      End If
   End Select
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim intCounter As Integer

   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid1.col
            Case 0
               SendKeys "{RIGHT}"
            Case 1
               For intCounter = 1 To 5
                  SendKeys "{RIGHT}"
               Next intCounter
            Case 6
               SendKeys "{DOWN}"
               For intCounter = 1 To 6
                  SendKeys "{LEFT}"
               Next intCounter
         End Select
   End Select
End Sub

Private Sub DataGrid1_LostFocus()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   SendKeys "{ENTER}"
   Adodc1.Recordset.UpdateBatch
Checking:
   Exit Sub
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.Fields("a0j08").Value = MsgText(602) Then
      If Text1 <> MsgText(601) Then
         strCustNo = Text1
      Else
         strCustNo = MsgText(601)
      End If
      Frmacc1120.Enabled = False
      Frmacc1122.Show
   End If
End Sub

'Add by Morgan 2010/12/6
Private Sub AutoProcess()
   Command2.Visible = False
   Text1 = m_CustNo
   Text1_Validate False
   With Adodc1.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         If .Fields("a0j01") = m_CP09 Then
            .Fields("a0j06") = "Y"
            .UpdateBatch
            'Add By Sindy 2022/11/28
            Call Command4_Click '檢視接洽單
            '2022/11/28 END
            Exit Do
         End If
         .MoveNext
      Loop
   End If
   End With
   Command1.SetFocus
End Sub

Private Sub Form_Activate()
   'Modify By Sindy 2015/2/11 瑞婷反應開收據時,若出現有人員休假訊息後,回到該畫面不會重Load資料
   Call CallFormActivate
'   Dim strYes As String
'   strFormName = Name
'
'   'Add by Morgan 2011/4/13
'   If m_Assigning Then
'      tool3_enabled
'      MenuDisabled
'      m_Assigning = False
'      Exit Sub
'   End If
'
'   If Text1 = MsgText(601) Or Text1 = "X" Then
'      If m_AutoProcess Then AutoProcess 'Add by Morgan 2010/12/6
'      Exit Sub
'   End If
'   Text1.SetFocus
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
''   ProduceData
'   AdodcRefresh
'   strReceiptConfirm = MsgText(601)
'   strItemNo = MsgText(601)
'   StatusClear
'   tool3_enabled
End Sub

'Modify By Sindy 2015/2/11 從Form_Activate移出來變共用函數
Public Sub CallFormActivate()
   Dim strYes As String
   strFormName = Name
   
   'Add by Morgan 2011/4/13
   If m_Assigning Then
      tool3_enabled
      MenuDisabled
      m_Assigning = False
      Exit Sub
   End If
   
   If Text1 = MsgText(601) Or Text1 = "X" Then
      If m_AutoProcess Then AutoProcess 'Add by Morgan 2010/12/6
      Exit Sub
   End If
   Text1.SetFocus
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   ProduceData
   AdodcRefresh
   strReceiptConfirm = MsgText(601)
   strItemNo = MsgText(601)
   StatusClear
   tool3_enabled
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
   Me.Width = 9500
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text1 = "X"
   strItemNo = MsgText(601)
   adoTaie.Execute "update acc0j0 set a0j06 = null where a0j06 = '" & MsgText(602) & "'"
   OpenTable
   'Mark by Amy 2014/06/30
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
   
   'Add By Sindy 2022/11/30
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      Command4.Visible = True
   Else
      Command4.Visible = False
   End If
   '2022/11/30 END
End Sub

Private Sub Form_Resize()
   strFormName = Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Sindy 2022/12/17 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q", tmpfrm) = True Then
      Unload tmpfrm 'frm090801_Q
   End If
   '2022/12/17 END
   
   adoTaie.Execute "update acc0j0 set a0j06 = null where a0j06 = '" & MsgText(602) & "'"
   StatusClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   
   If Not m_AutoProcess Then 'Added by Morgan 2024/5/24 整批進來的改在母畫面解鎖
      Call PUB_GetLock("", Me.Name) 'Add by Morgan 2006/6/29 解除鎖定
   End If
   
   Set Frmacc1120 = Nothing
   
   'Add by Morgan 2010/12/7
   If m_AutoProcess Then
      Frmacc1123.m_Continue = True
      Frmacc1123.Show
   End If
   
End Sub

Private Sub Text1_GotFocus()
Dim intPos As Integer  '2005/11/23 ADD BY SONIA
   
   TextInverse Text1
   '2005/11/23 ADD BY SONIA
   If Len("" & Text1) > 0 Then
      intPos = InStr("" & Text1, "X")
      If intPos - 1 = 0 Then
         If Len("" & Text1) > 1 Then
            Text1.SelStart = 1
         Else
            Text1.SelStart = 2
         End If
      End If
      Text1.SelLength = Len("" & Text1) - 1
   End If
   '2005/11/23 END
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         If Text1 = "" Then
            Exit Sub
         End If
         Select Case Len(Text1)
            Case 6
               Text1 = Text1 & "000"
            Case 8
               Text1 = Text1 & "0"
         End Select
         If ExistCheck("customer", "cu01", Mid(Text1, 1, 8), "", False) = False Then
            MsgBox MsgText(45) & Label1, , MsgText(5)
            Text1 = "X"
            TextInverse Text1
            'edit by nickc 2007/02/08
            'Cancel = True
            Text1.SetFocus
            Exit Sub
         End If
         Text2 = CustomerQuery(Text1, 1)
   End Select
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0j0.CursorLocation = adUseClient
   adoacc0j0.Open "select * from acc0j0 where a0j01 = 'Z'", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select * from acc0k0 where a0k01 = 'Z'", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.CursorLocation = adUseClient
   '2005/6/24 MODIFY BY SONIA
   'adoadodc1.Open "select * from acc0j0 where a0j13 is null and a0j11 like '" & Text1 & "' order by a0j05 asc, a0j02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Morgan 2011/9/19 a0j13改先放收文號
   'adoadodc1.Open "select * from acc0j0 where a0j13 is null and a0j11 = '" & Text1 & "' order by a0j05 asc, a0j02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Morgan 2011/12/26 取消 a0j05 改抓 cp13
   'Modified by Morgan 2012/1/2 取消 a0j20,a0j21
   adoadodc1.Open "select a.*,b.*,getcp10desc(cp01,cp10,a0j04) cp10N,na03 from acc0j0 a,caseprogress b,nation where a0j13=a0j01 and a0j11 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 order by cp13 asc, a0j02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2005/6/24 END
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'***************************************************
'  將未開收據之案件儲存於國內未開收據案件資料表內
'
'***************************************************
Private Sub Process()
   Dim strComputer As String
   Dim stCon As String, stCon1 As String
   Dim stVTB As String, iCnt As Integer 'Add by Morgan 2010/12/20
   
   iCnt = 100 'Add by Morgan 2010/12/20 控制案件數超過的用CP當主要資料表
   
   strNoNation = ""
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j11 = '" & Text1 & "'"
   adopatent.CursorLocation = adUseClient
'   adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                  "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and (pa26 like '" & Mid(Text1, 1, 8) & "%" & "' or pa27 like '" & Text1 & "%" & "' or pa28 like '" & Text1 & "%" & "' or pa29 like '" & Text1 & "%" & "' or pa30 like '" & Text1 & "%" & "') and (cp60 is null or cp60 = '') and cp10 not in ('907') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   '92.2.25 MODIFY BY SONIA CP05>=20011112 改成 20030101
   'adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
   '               "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and (cp56 = '" & Text1 & "' or pa26 = '" & Text1 & "' or pa27 = '" & Text1 & "' or pa28 = '" & Text1 & "' or pa29 = '" & Text1 & "' or pa30 = '" & Text1 & "') and (cp60 is null or cp60 = '') and cp10 not in ('907') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   '92.9.2 MODIFY BY SONIA  取消and cp10 not in ('907') 的限制
   'adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
   '               "and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and (cp56 = '" & Text1 & "' or pa26 = '" & Text1 & "' or pa27 = '" & Text1 & "' or pa28 = '" & Text1 & "' or pa29 = '" & Text1 & "' or pa30 = '" & Text1 & "') and (cp60 is null or cp60 = '') and cp10 not in ('907') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modify by Morgan 2006/7/25 加可用受讓人查該案件的其他程序
   'adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
                  "and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and (cp56 = '" & Text1 & "' or pa26 = '" & Text1 & "' or pa27 = '" & Text1 & "' or pa28 = '" & Text1 & "' or pa29 = '" & Text1 & "' or pa30 = '" & Text1 & "') and (cp60 is null or cp60 = '') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modify by Morgan 2010/12/20 改寫法調整效能(抓兩年內收文資料就好--辜)
   'stCon = " and (cp01,cp02,cp03,cp04) in ( select cp01,cp02,cp03,cp04 from caseprogress,patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and (cp56 = '" & Text1 & "' or pa26 = '" & Text1 & "' or pa27 = '" & Text1 & "' or pa28 = '" & Text1 & "' or pa29 = '" & Text1 & "' or pa30 = '" & Text1 & "') and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and (cp60 is null or cp60 = '') and (cp20 is null or cp20 = ''))"
   'stCon = " and pa75 is null" & stCon 'Add by Morgan 2010/12/10
   'adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
                  " and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and (cp60 is null or cp60 = '') and (cp20 is null or cp20 = '')" & stCon, adoTaie, adOpenStatic, adLockReadOnly
   
   'Modifed by Morgan 2013/5/9 排除待確認的電子送件程序
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0"
   'Modified by Morgan 2014/1/2
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0 and nvl(cp118,'N')<>'W'"
   If Check1.Value = vbUnchecked Then
      stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')"
   Else
      stCon = " and cp05 >= 20030101"
   End If
   stCon = stCon & " and cp20||cp57||cp60 is null and cp16>0 and nvl(cp118,'N')<>'W'"
   'end 2014/1/2
   
   'Added by Morgan 2023/1/7  --辜,秀玲
   '1. P及CFP的605、606、607，北所未分案時不出現。
    stCon = stCon & " and not(cp01 in ('P','CFP') and cp10 in ('605','606','607') and nvl(cp157,0)=0)"
    'end 2023/1/7
   
   'Modify by Morgan 2011/3/24 改智權部收文都要其他的才抓沒有FC代理人的
   'stCon1 = " and pa75 is null"
   'modify by sonia 2021/1/22 +pa26='X67934000' 此客戶改由唐韻如
   stCon1 = " and (pa75 is null or pa26='X67934000' or substr(cp12,1,1)='S')"
   
   strExc(0) = "select count(*) from patent where pa26='" & Text1 & "' having count(*)<=" & iCnt
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify By Sindy 2014/7/9 只抓取申請人1開立收據
'      stVTB = "select cp09 from patent,caseprogress" & _
'         " where pa26 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
'         " union all select cp09 from caseprogress,patent" & _
'         " where pa27 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
'         " union all select cp09 from caseprogress,patent" & _
'         " where pa28 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
'         " union all select cp09 from caseprogress,patent" & _
'         " where pa29 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
'         " union all select cp09 from caseprogress,patent" & _
'         " where pa30 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
'         " union all select cp09 from caseprogress,patent" & _
'         " where (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress" & _
'         " where cp56 = '" & Text1 & "'" & stCon & ")" & stCon & stCon1 & _
'         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null"
      stVTB = "select cp09 from patent,caseprogress" & _
         " where pa26 = '" & Text1 & "'" & stCon & stCon1 & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
         " union all select cp09 from caseprogress,patent" & _
         " where (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress" & _
         " where cp56 = '" & Text1 & "'" & stCon & ")" & stCon & stCon1 & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null"
   Else
      'Modify By Sindy 2014/7/9 只抓取申請人1開立收據
'      stVTB = "select cp09 from  caseprogress,patent" & _
'         " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
'         " and pa01 is not null " & stCon & stCon1 & _
'         " and ( pa26 = '" & Text1 & "' or pa27 = '" & Text1 & "' or pa28 = '" & Text1 & "'" & _
'         " or pa29 = '" & Text1 & "' or pa30 = '" & Text1 & "')" & _
'         " union all select cp09 from caseprogress,patent" & _
'         " where (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress" & _
'         " where cp56 = '" & Text1 & "'" & stCon & ")" & stCon & stCon1 & _
'         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null"
      stVTB = "select cp09 from  caseprogress,patent" & _
         " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and pa01 is not null " & stCon & stCon1 & _
         " and pa26 = '" & Text1 & "'" & _
         " union all select cp09 from caseprogress,patent" & _
         " where (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress" & _
         " where cp56 = '" & Text1 & "'" & stCon & ")" & stCon & stCon1 & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null"
   End If
   strSql = "select * from caseprogress a, patent, casepropertymap" & _
      " where cp09 in (" & stVTB & ")" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 "
      
   'Added by Morgan 2013/7/2 相同案號有待確認電子送件程序的都不能開
   strSql = strSql & " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')"
   'end 2013/7/2

   adopatent.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   'end 2010/12/17
   'end 2006/7/25
   '92.9.2 END
   '92.2.25 END
   Do While adopatent.EOF = False
      'Modify by Morgan 2011/9/19 a0j13改先放收文號
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adopatent.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adopatent.Fields("cp09").Value & "'"
      If adopatent.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveP
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select cu10, na03 from patent, customer, nation where substr(pa26, 1, 8) = cu01 and cu02 = '0' and pa09 = na01 and pa01 = '" & adopatent.Fields("pa01").Value & "' and pa02 = '" & adopatent.Fields("pa02").Value & "' and pa03 = '" & adopatent.Fields("pa03").Value & "' and pa04 = '" & adopatent.Fields("pa04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         
         'Modified by Morgan 2011/12/29 取消 a0j21
         'If IsNull(adocheck.Fields(1).Value) Then
         '   stra0j21 = ""
         'Else
         '   stra0j21 = adocheck.Fields(1).Value
         'End If
         
         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2010/12/6 不必再限制客戶國籍
            'If Val(adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  If adopatent.Fields("pa09").Value = "020" Then
                     stra0j07 = MsgText(602)
                  End If
                  '2005/6/8 ADD BY SONIA, 2005/7/5 加 201, 2005/7/26 加 507,508
                  'Modify by Morgan 2007/9/11 211,212,503,507,506不再合併
                  'If adopatent.Fields("CP10").Value = "211" Or adopatent.Fields("CP10").Value = "212" Or adopatent.Fields("CP10").Value = "503" Or adopatent.Fields("CP10").Value = "506" Or adopatent.Fields("CP10").Value = "906" Or adopatent.Fields("CP10").Value = "201" Or adopatent.Fields("CP10").Value = "507" Or adopatent.Fields("CP10").Value = "508" Then
                  'Modify by Morgan 2010/12/10 +927 也要合併
                  If adopatent.Fields("CP10").Value = "906" Or adopatent.Fields("CP10").Value = "201" Or adopatent.Fields("CP10").Value = "927" Or adopatent.Fields("CP10").Value = "508" Then
                     stra0j07 = MsgText(602)
                  End If
                  '2005/6/8 END
                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11, a0j14, a0j15, a0j16, a0j08, a0j07,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & stra0j04 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & strComputer & "', '" & stra0j07 & "','" & stra0j01 & "')"
               End If
            'End If
         Else
            Screen.MousePointer = vbDefault
            strNoNation = MsgText(602)
            adocheck.Close
            adopatent.Close
            AdodcRefresh
            Exit Sub
         End If
      End If
      adocheck.Close
      adopatent.MoveNext
   Loop
   adopatent.Close
   adotrademark.CursorLocation = adUseClient
'   adotrademark.Open "select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                     "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and tm23 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('703') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   '92.9.2 MODIFY BY SONIA 取消 and cp10 not in ('703') 的限制
   'adotrademark.Open "select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
   '                  "and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and (cp56 = '" & Text1 & "' or tm23 = '" & Text1 & "') and (cp60 is null or cp60 = '') and cp10 not in ('703') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modify by Morgan 2006/7/25 加可用受讓人查該案件的其他程序
   'adotrademark.Open "select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
                     "and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and (cp56 = '" & Text1 & "' or tm23 = '" & Text1 & "') and (cp60 is null or cp60 = '') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modify by Morgan 2010/12/20 改寫法調整效能(抓兩年內收文資料就好--辜)
   'stCon = " and (cp01,cp02,cp03,cp04) in ( select cp01,cp02,cp03,cp04 from caseprogress,trademark where tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and (cp56 = '" & Text1 & "' or tm23 = '" & Text1 & "') and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and (cp60 is null or cp60 = '') and (cp20 is null or cp20 = ''))"
   'stCon = " and tm44 is null" & stCon 'Add by Morgan 2010/12/10
   'adotrademark.Open "select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
                     "and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and (cp60 is null or cp60 = '') and (cp20 is null or cp20 = '')" & stCon, adoTaie, adOpenStatic, adLockReadOnly
   
   'Modifed by Morgan 2013/5/9 排除待確認的電子送件程序
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0"
   'Modified by Morgan 2014/1/2
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0 and nvl(cp118,'N')<>'W'"
      If Check1.Value = vbUnchecked Then
      stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')"
   Else
      stCon = " and cp05 >= 20030101"
   End If
   stCon = stCon & " and cp20||cp57||cp60 is null and cp16>0 and nvl(cp118,'N')<>'W'"
   'end 2014/1/2
   
   'Modify by Morgan 2011/3/24 改智權部收文都要其他的才抓沒有FC代理人的
   'stCon1 = " and tm44 is null"
   stCon1 = " and (tm44 is null or substr(cp12,1,1)='S')"
   
   strExc(0) = "select count(*) from trademark where tm23='" & Text1 & "' having count(*)<=" & iCnt
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify By Sindy 2014/7/9 只抓取申請人1開立收據
'      stVTB = "select cp09 from trademark,caseprogress" & _
'         " where tm23 = '" & Text1 & "'" & stCon & _
'         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04" & _
'         " union all select cp09 from trademark,caseprogress" & _
'         " where tm78 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04" & _
'         " union all select cp09 from trademark,caseprogress" & _
'         " where tm79 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04" & _
'         " union all select cp09 from trademark,caseprogress" & _
'         " where tm80 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04" & _
'         " union all select cp09 from trademark,caseprogress" & _
'         " where tm81 = '" & Text1 & "'" & stCon & stCon1 & _
'         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04" & _
'         " union all select cp09 from caseprogress,trademark" & _
'         " where (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress" & _
'         " where cp56 = '" & Text1 & "'" & stCon & ")" & stCon & stCon1 & _
'         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null"
      stVTB = "select cp09 from trademark,caseprogress" & _
         " where tm23 = '" & Text1 & "'" & stCon & _
         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04" & _
         " union all select cp09 from caseprogress,trademark" & _
         " where (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress" & _
         " where cp56 = '" & Text1 & "'" & stCon & ")" & stCon & stCon1 & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null"
   Else
      'Modify By Sindy 2014/7/9 只抓取申請人1開立收據
'      stVTB = "select cp09 from caseprogress,trademark" & _
'         " where tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
'         " and tm01 is not null " & stCon & _
'         " and (tm23='" & Text1 & "' or tm78='" & Text1 & "' or tm79='" & Text1 & "'" & _
'         " or tm80='" & Text1 & "' or tm81='" & Text1 & "')" & _
'         " union all select cp09 from caseprogress,trademark" & _
'         " where (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress" & _
'         " where cp56 = '" & Text1 & "'" & stCon & ")" & stCon & stCon1 & _
'         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null"
      stVTB = "select cp09 from caseprogress,trademark" & _
         " where tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
         " and tm01 is not null " & stCon & _
         " and tm23='" & Text1 & "'" & _
         " union all select cp09 from caseprogress,trademark" & _
         " where (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress" & _
         " where cp56 = '" & Text1 & "'" & stCon & ")" & stCon & stCon1 & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null"
   End If
      
   strSql = "select * from caseprogress a, trademark, casepropertymap" & _
      " where cp09 in (" & stVTB & ")" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 "
      
   'Added by Morgan 2013/7/2 相同案號有待確認電子送件程序的都不能開
   strSql = strSql & " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')"
   'end 2013/7/2

   adotrademark.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   'end 2010/12/20
   'end 2006/7/25
   '92.9.2 END
   Do While adotrademark.EOF = False
      'Modify by Morgan 2011/9/19 a0j13改先放收文號
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adotrademark.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adotrademark.Fields("cp09").Value & "'"
      '2014/2/26 remark by sonia 瑞婷說不顯示手開收據欄,故僅將欄位寬度改為0
      If adotrademark.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveT
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select cu10, na03 from trademark, customer, nation where substr(tm23, 1, 8) = cu01 and cu02 = '0' and tm10 = na01 and tm01 = '" & adotrademark.Fields("tm01").Value & "' and tm02 = '" & adotrademark.Fields("tm02").Value & "' and tm03 = '" & adotrademark.Fields("tm03").Value & "' and tm04 = '" & adotrademark.Fields("tm04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
      
         'Modified by Morgan 2011/12/29 取消 a0j21
         'If IsNull(adocheck.Fields(1).Value) Then
         '   stra0j21 = ""
         'Else
         '   stra0j21 = adocheck.Fields(1).Value
         'End If
         
         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2010/12/7 不必再限制客戶國籍
            'If Val(adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  If adotrademark.Fields("tm10").Value = "020" Then
                     stra0j07 = MsgText(602)
                  End If
                  '2005/6/8 ADD BY SONIA, 2004/7/26 加 408,410
                  'Modify by Morgan 2007/9/11 204,205,403,408,407不再合併
                  'If adotrademark.Fields("CP10").Value = "204" Or adotrademark.Fields("CP10").Value = "205" Or adotrademark.Fields("CP10").Value = "403" Or adotrademark.Fields("CP10").Value = "407" Or adotrademark.Fields("CP10").Value = "408" Or adotrademark.Fields("CP10").Value = "410" Then
                  If adotrademark.Fields("CP10").Value = "410" Then
                     stra0j07 = MsgText(602)
                  End If
                  '2005/6/8 END
                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11, a0j14, a0j15, a0j16, a0j08, a0j07,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & stra0j04 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & strComputer & "', '" & stra0j07 & "','" & stra0j01 & "')"
               End If
            'End If
         Else
            Screen.MousePointer = vbDefault
            strNoNation = MsgText(602)
            adocheck.Close
            adotrademark.Close
            AdodcRefresh
            Exit Sub
         End If
      End If
      adocheck.Close
      adotrademark.MoveNext
   Loop
   adotrademark.Close
   adolawcase.CursorLocation = adUseClient
'   adolawcase.Open "select * from caseprogress, lawcase, casepropertymap where (caseprogress.cp01 = lawcase.lc01  and caseprogress.cp02 = lawcase.lc02 and caseprogress.cp03 = lawcase.lc03 and caseprogress.cp04 = lawcase.lc04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                   "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and lc11 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('999') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modify by Morgan 2010/12/20 改寫法調整效能(抓兩年內收文資料就好--辜)
   'stCon = ""
   'stCon = " and lc22 is null" & stCon 'Add by Morgan 2010/12/10
   'adolawcase.Open "select * from caseprogress, lawcase, casepropertymap where (caseprogress.cp01 = lawcase.lc01  and caseprogress.cp02 = lawcase.lc02 and caseprogress.cp03 = lawcase.lc03 and caseprogress.cp04 = lawcase.lc04) and cp01 = cpm01 and cp10 = cpm02 " & _
                   "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and lc11 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('999') and (cp20 is null or cp20 = '')" & stCon, adoTaie, adOpenStatic, adLockReadOnly
   
   'Modifed by Morgan 2013/5/9 排除待確認的電子送件程序
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0 and cp10<>'999'"
   'Modified by Morgan 2014/1/2
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0 and cp10<>'999' and nvl(cp118,'N')<>'W'"
   If Check1.Value = vbUnchecked Then
      stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')"
   Else
      stCon = " and cp05 >= 20030101"
   End If
   stCon = stCon & " and cp20||cp57||cp60 is null and cp16>0 and cp10<>'999' and nvl(cp118,'N')<>'W'"
   'end 2014/1/2
   

   'Modify by Morgan 2011/3/24 改智權部收文都要其他的才抓沒有FC代理人的
   'stCon = stCon & " and lc22 is null"
   stCon = stCon & " and (lc22 is null or substr(cp12,1,1)='S')"
   
   'Modify By Sindy 2011/2/21 增加LC43,LC44,LC45,LC46
   'Modify By Sindy 2014/7/9 只抓取申請人1開立收據
'   stVTB = "select cp09 from lawcase,caseprogress" & _
'      " where lc11 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=lc01 and cp02(+)=lc02 and cp03(+)=lc03 and cp04(+)=lc04" & _
'      " union all select cp09 from lawcase,caseprogress" & _
'      " where lc43 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=lc01 and cp02(+)=lc02 and cp03(+)=lc03 and cp04(+)=lc04" & _
'      " union all select cp09 from lawcase,caseprogress" & _
'      " where lc44 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=lc01 and cp02(+)=lc02 and cp03(+)=lc03 and cp04(+)=lc04" & _
'      " union all select cp09 from lawcase,caseprogress" & _
'      " where lc45 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=lc01 and cp02(+)=lc02 and cp03(+)=lc03 and cp04(+)=lc04" & _
'      " union all select cp09 from lawcase,caseprogress" & _
'      " where lc46 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=lc01 and cp02(+)=lc02 and cp03(+)=lc03 and cp04(+)=lc04"
   stVTB = "select cp09 from lawcase,caseprogress" & _
      " where lc11 = '" & Text1 & "'" & stCon & _
      " and cp01(+)=lc01 and cp02(+)=lc02 and cp03(+)=lc03 and cp04(+)=lc04"
      
   strSql = "select * from caseprogress a, lawcase, casepropertymap" & _
      " where cp09 in (" & stVTB & ")" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 "
      
   'Added by Morgan 2013/7/2 相同案號有待確認電子送件程序的都不能開
   strSql = strSql & " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')"
   'end 2013/7/2

   adolawcase.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   'end 2010/12/20
   
   Do While adolawcase.EOF = False
      'Modify by Morgan 2011/9/19 a0j13改先放收文號
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adolawcase.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adolawcase.Fields("cp09").Value & "'"
      If adolawcase.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveL
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select cu10, na03 from lawcase, customer, nation where substr(lc11, 1, 8) = cu01 and cu02 = '0' and lc15 = na01 and lc01 = '" & adolawcase.Fields("lc01").Value & "' and lc02 = '" & adolawcase.Fields("lc02").Value & "' and lc03 = '" & adolawcase.Fields("lc03").Value & "' and lc04 = '" & adolawcase.Fields("lc04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         
         'Modified by Morgan 2011/12/29 取消 a0j21
         'If IsNull(adocheck.Fields(1).Value) Then
         '   stra0j21 = ""
         'Else
         '   stra0j21 = adocheck.Fields(1).Value
         'End If
         
         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2010/12/7 不必再限制客戶國籍
            'If Val(adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  If adolawcase.Fields("lc15").Value = "020" Then
                     stra0j07 = MsgText(602)
                  End If
                  'add by sonia 2021/3/30 ACS代收代付706要合併
                  If adolawcase.Fields("lc01").Value = "ACS" And adolawcase.Fields("CP10").Value = "706" Then
                     stra0j07 = MsgText(602)
                  End If
                  'end 2021/3/30

                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11, a0j14, a0j15, a0j16, a0j08, a0j07,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & stra0j04 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & strComputer & "', '" & stra0j07 & "','" & stra0j01 & "')"
               End If
            'End If
         Else
            Screen.MousePointer = vbDefault
            strNoNation = MsgText(602)
            adocheck.Close
            adolawcase.Close
            AdodcRefresh
            Exit Sub
         End If
      End If
      adocheck.Close
      adolawcase.MoveNext
   Loop
   adolawcase.Close
   adohirecase.CursorLocation = adUseClient
'   adohirecase.Open "select * from caseprogress, hirecase, casepropertymap where (caseprogress.cp01 = hirecase.hc01  and caseprogress.cp02 = hirecase.hc02 and caseprogress.cp03 = hirecase.hc03 and caseprogress.cp04 = hirecase.hc04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                     "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and hc05 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('999') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modify by Morgan 2010/12/20 改寫法調整效能(抓兩年內收文資料就好--辜)
   'adohirecase.Open "select * from caseprogress, hirecase, casepropertymap where (caseprogress.cp01 = hirecase.hc01  and caseprogress.cp02 = hirecase.hc02 and caseprogress.cp03 = hirecase.hc03 and caseprogress.cp04 = hirecase.hc04) and cp01 = cpm01 and cp10 = cpm02 " & _
                     "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and hc05 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('999') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modifed by Morgan 2013/5/9 排除待確認的電子送件程序
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0 and cp10<>'999'"
   'Modified by Morgan 2014/1/2
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0 and cp10<>'999' and nvl(cp118,'N')<>'W'"
   If Check1.Value = vbUnchecked Then
      stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')"
   Else
      stCon = " and cp05 >= 20030101"
   End If
   stCon = stCon & " and cp20||cp57||cp60 is null and cp16>0 and cp10<>'999' and nvl(cp118,'N')<>'W'"
   'end 2014/1/2
   
   'Modify By Sindy 2011/2/21 增加HC24,HC25,HC26,HC27
   'Modify By Sindy 2014/7/9 只抓取申請人1開立收據
'   stVTB = "select cp09 from hirecase,caseprogress" & _
'      " where hc05 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=hc01 and cp02(+)=hc02 and cp03(+)=hc03 and cp04(+)=hc04" & _
'      " union all select cp09 from hirecase,caseprogress" & _
'      " where hc24 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=hc01 and cp02(+)=hc02 and cp03(+)=hc03 and cp04(+)=hc04" & _
'      " union all select cp09 from hirecase,caseprogress" & _
'      " where hc25 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=hc01 and cp02(+)=hc02 and cp03(+)=hc03 and cp04(+)=hc04" & _
'      " union all select cp09 from hirecase,caseprogress" & _
'      " where hc26 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=hc01 and cp02(+)=hc02 and cp03(+)=hc03 and cp04(+)=hc04" & _
'      " union all select cp09 from hirecase,caseprogress" & _
'      " where hc27 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=hc01 and cp02(+)=hc02 and cp03(+)=hc03 and cp04(+)=hc04"
   stVTB = "select cp09 from hirecase,caseprogress" & _
      " where hc05 = '" & Text1 & "'" & stCon & _
      " and cp01(+)=hc01 and cp02(+)=hc02 and cp03(+)=hc03 and cp04(+)=hc04"
      
   strSql = "select * from caseprogress a, hirecase, casepropertymap" & _
      " where cp09 in (" & stVTB & ")" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 "
      
   'Added by Morgan 2013/7/2 相同案號有待確認電子送件程序的都不能開
   strSql = strSql & " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')"
   'end 2013/7/2

   adohirecase.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   'end 2010/12/20
   
   Do While adohirecase.EOF = False
      'Modify by Morgan 2011/9/19 a0j13改先放收文號
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adohirecase.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adohirecase.Fields("cp09").Value & "'"
      If adohirecase.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveLA
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select cu10 from hirecase, customer where substr(hc05, 1, 8) = cu01 and cu02 = '0' and hc01 = '" & adohirecase.Fields("hc01").Value & "' and hc02 = '" & adohirecase.Fields("hc02").Value & "' and hc03 = '" & adohirecase.Fields("hc03").Value & "' and hc04 = '" & adohirecase.Fields("hc04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         'stra0j21 = "" 'Removed by Morgan 2011/12/29 取消 a0j21
         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2010/12/7 不必再限制客戶國籍
            'If Val(adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11, a0j14, a0j15, a0j16, a0j08,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '000', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & strComputer & "','" & stra0j01 & "')"
               End If
            'End If
         Else
            Screen.MousePointer = vbDefault
            strNoNation = MsgText(602)
            adocheck.Close
            adohirecase.Close
            AdodcRefresh
            Exit Sub
         End If
      End If
      adocheck.Close
      adohirecase.MoveNext
   Loop
   adohirecase.Close
   adoservice.CursorLocation = adUseClient
'   adoservice.Open "select * from caseprogress, servicepractice, casepropertymap where (caseprogress.cp01 = sp01  and caseprogress.cp02 = sp02 and caseprogress.cp03 = sp03 and caseprogress.cp04 = sp04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                   "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and sp08 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and (cp20 is null or cp20 = '')", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modify by Morgan 2010/12/20 改寫法調整效能(抓兩年內收文資料就好--辜)
   'stCon = ""
   'stCon = " and sp26 is null" & stCon 'Add by Morgan 2010/12/10
   'adoservice.Open "select * from caseprogress, servicepractice, casepropertymap where (caseprogress.cp01 = sp01  and caseprogress.cp02 = sp02 and caseprogress.cp03 = sp03 and caseprogress.cp04 = sp04) and cp01 = cpm01 and cp10 = cpm02 " & _
                   "and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and sp08 = '" & Text1 & "' and (cp60 is null or cp60 = '') and (cp20 is null or cp20 = '')" & stCon, adoTaie, adOpenStatic, adLockReadOnly
   
   'Modifed by Morgan 2013/5/9 排除待確認的電子送件程序
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0"
   'Modified by Morgan 2014/1/2
   'stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')  and cp20||cp57||cp60 is null and cp16>0 and nvl(cp118,'N')<>'W'"
   If Check1.Value = vbUnchecked Then
      stCon = " and cp05 >=to_char(add_months(sysdate,-24),'yyyymmdd')"
   Else
      stCon = " and cp05 >= 20030101"
   End If
   stCon = stCon & " and cp20||cp57||cp60 is null and cp16>0 and nvl(cp118,'N')<>'W'"
   'end 2014/1/2
   
   'Modify by Morgan 2011/3/24 改智權部收文都要其他的才抓沒有FC代理人的
   'stCon = stCon & " and sp26 is null"
   stCon = stCon & " and (sp26 is null or substr(cp12,1,1)='S')"
   
   'Modify By Sindy 2014/7/9 只抓取申請人1開立收據
'   stVTB = "select cp09 from servicepractice,caseprogress" & _
'      " where sp08 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=sp01 and cp02(+)=sp02 and cp03(+)=sp03 and cp04(+)=sp04" & _
'      " union all select cp09 from servicepractice,caseprogress" & _
'      " where sp58 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=sp01 and cp02(+)=sp02 and cp03(+)=sp03 and cp04(+)=sp04" & _
'      " union all select cp09 from servicepractice,caseprogress" & _
'      " where sp59 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=sp01 and cp02(+)=sp02 and cp03(+)=sp03 and cp04(+)=sp04" & _
'      " union all select cp09 from servicepractice,caseprogress" & _
'      " where sp65 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=sp01 and cp02(+)=sp02 and cp03(+)=sp03 and cp04(+)=sp04" & _
'      " union all select cp09 from servicepractice,caseprogress" & _
'      " where sp66 = '" & Text1 & "'" & stCon & _
'      " and cp01(+)=sp01 and cp02(+)=sp02 and cp03(+)=sp03 and cp04(+)=sp04"
   stVTB = "select cp09 from servicepractice,caseprogress" & _
      " where sp08 = '" & Text1 & "'" & stCon & _
      " and cp01(+)=sp01 and cp02(+)=sp02 and cp03(+)=sp03 and cp04(+)=sp04"
      
   strSql = "select * from caseprogress a, servicepractice, casepropertymap" & _
      " where cp09 in (" & stVTB & ")" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 "
      
   'Added by Morgan 2013/7/2 相同案號有待確認電子送件程序的都不能開
   strSql = strSql & " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')"
   'end 2013/7/2

   adoservice.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   'end 2010/12/20
   
   Do While adoservice.EOF = False
      'Modify by Morgan 2011/9/19
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adoservice.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adoservice.Fields("cp09").Value & "'"
      If adoservice.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveS
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select cu10, na03 from servicepractice, customer, nation where substr(sp08, 1, 8) = cu01 and cu02 = '0' and sp09 = na01 and sp01 = '" & adoservice.Fields("sp01").Value & "' and sp02 = '" & adoservice.Fields("sp02").Value & "' and sp03 = '" & adoservice.Fields("sp03").Value & "' and sp04 = '" & adoservice.Fields("sp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         'Modified by Morgan 2011/12/29 取消 a0j21
         'If IsNull(adocheck.Fields(1).Value) Then
         '   stra0j21 = ""
         'Else
         '   stra0j21 = adocheck.Fields(1).Value
         'End If
         
         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2010/12/7 不必再限制客戶國籍
            'If Val(adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  'add by sonia 2020/6/23 TT-999999預設合併
                  If stra0j02 = "TT999999000" Then
                     stra0j07 = MsgText(602)
                  End If
                  'end 2020/6/23
                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11, a0j14, a0j15, a0j16, a0j08, a0j07,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & stra0j04 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & strComputer & "', '" & stra0j07 & "','" & stra0j01 & "')"
               End If
            'End If
         Else
            Screen.MousePointer = vbDefault
            strNoNation = MsgText(602)
            adocheck.Close
            adoservice.Close
            AdodcRefresh
            Exit Sub
         End If
      End If
      adocheck.Close
      adoservice.MoveNext
   Loop
   adoservice.Close
   AdodcRefresh
   
   Screen.MousePointer = vbDefault
   'Mark by Amy 2014/06/30
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
End Sub

'*************************************************
'  將專利資料放置系統變數中
'
'*************************************************
Private Sub Acc0j0SaveP()
   stra0j01 = adopatent.Fields("cp09").Value
   If IsNull(adopatent.Fields("cp01").Value) Then
      stra0j02 = MsgText(601)
   Else
      stra0j02 = adopatent.Fields("cp01").Value
      If IsNull(adopatent.Fields("cp02").Value) = False Then
         stra0j02 = stra0j02 & adopatent.Fields("cp02").Value
      End If
      If IsNull(adopatent.Fields("cp03").Value) = False Then
         stra0j02 = stra0j02 & adopatent.Fields("cp03").Value
      End If
      If IsNull(adopatent.Fields("cp04").Value) = False Then
         stra0j02 = stra0j02 & adopatent.Fields("cp04").Value
      End If
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j03
   'If IsNull(adopatent.Fields("cp10").Value) Then
   '   lnga0j03 = MsgText(601)
   'Else
   '   lnga0j03 = adopatent.Fields("cp10").Value
   'End If
   
   If IsNull(adopatent.Fields("pa09").Value) Then
      stra0j04 = MsgText(601)
   Else
      stra0j04 = adopatent.Fields("pa09").Value
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j05
   'If IsNull(adopatent.Fields("cp13").Value) Then
   '   stra0j05 = MsgText(601)
   'Else
   '   stra0j05 = adopatent.Fields("cp13").Value
   'End If
   
'   If Mid(adopatent.Fields("cp01").Value, 1, 2) = "CF" Then
'      stra0j07 = MsgText(602)
'   Else
      If adopatent.Fields("pa09").Value <> "000" Then
         stra0j07 = MsgText(602)
      Else
         stra0j07 = MsgText(601)
      End If
'   End If

   If IsNull(adopatent.Fields("cp16").Value) Then
      lnga0j09 = 0
   Else
      If IsNull(adopatent.Fields("cp17").Value) Then
         lnga0j09 = adopatent.Fields("cp16").Value
      Else
         lnga0j09 = Val(adopatent.Fields("cp16").Value) - Val(adopatent.Fields("cp17").Value)
      End If
   End If
   If IsNull(adopatent.Fields("cp17").Value) Then
      lnga0j10 = 0
   Else
      lnga0j10 = adopatent.Fields("cp17").Value
   End If
   If Text1 = MsgText(601) Then
      stra0j11 = MsgText(601)
   Else
      stra0j11 = Text1
   End If
   'If IsNull(adopatent.Fields("pa26").Value) Then
   '   If IsNull(adopatent.Fields("pa27").Value) Then
   '      If IsNull(adopatent.Fields("pa28").Value) Then
   '         If IsNull(adopatent.Fields("pa29").Value) Then
   '            If IsNull(adopatent.Fields("pa30").Value) Then
   '               stra0j11 = MsgText(601)
   '            Else
   '               stra0j11 = Mid(adopatent.Fields("pa30").Value, 1, 8) & "0"
   '            End If
   '         Else
   '            stra0j11 = Mid(adopatent.Fields("pa29").Value, 1, 8) & "0"
   '         End If
   '      Else
   '         stra0j11 = Mid(adopatent.Fields("pa28").Value, 1, 8) & "0"
   '      End If
   '   Else
   '      stra0j11 = Mid(adopatent.Fields("pa27").Value, 1, 8) & "0"
   '   End If
   'Else
   '   stra0j11 = Mid(adopatent.Fields("pa26").Value, 1, 8) & "0"
   'End If
   
   'Removed by Morgan 2011/12/26 取消 a0j12
   'If IsNull(adopatent.Fields("cp05").Value) = False Then
   '   If adopatent.Fields("cp05").Value <> 0 Then
   '      stra0j12 = ACDate(adopatent.Fields("cp05").Value)
   '   Else
   '      stra0j12 = ""
   '   End If
   'Else
   '   stra0j12 = ""
   'End If
   
   'Removed by Morgan 2011/12/27 取消 a0j20
   'If stra0j04 = "000" Then
   '   If InStr(adopatent.Fields("cpm03").Value, MsgText(173)) > 0 Then
   '      If IsNull(adopatent.Fields("cpm04").Value) = False Then
   '         stra0j20 = adopatent.Fields("cpm04").Value
   '      Else
   '         stra0j20 = ""
   '      End If
   '   Else
   '      stra0j20 = adopatent.Fields("cpm03").Value
   '   End If
   'Else
   '   If IsNull(adopatent.Fields("cpm04").Value) = False Then
   '      stra0j20 = adopatent.Fields("cpm04").Value
   '   Else
   '      stra0j20 = ""
   '   End If
   'End If
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
  'Mark by Amy 2014/06/30 目前使用Tab刪acc0j0再重抓資料,下列insert及F12程式用不到
'   Select Case KeyCode
'      Case vbKeyInsert
'         If Text1 = "" Then
'            Exit Sub
'         End If
'         Acc0j0Delete
'         Process
'      Case vbKeyF12
'         AdodcRefresh
'   End Select
   KeyEnter KeyCode
   '避免用F12可能會抓到之前暫存的資料造成混淆,故mark
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
   'end 2014/06/30
End Sub

'*************************************************
'  將商標資料放置系統變數中
'
'*************************************************
Private Sub Acc0j0SaveT()
   stra0j01 = adotrademark.Fields("cp09").Value
   If IsNull(adotrademark.Fields("cp01").Value) Then
      stra0j02 = MsgText(601)
   Else
      stra0j02 = adotrademark.Fields("cp01").Value
      If IsNull(adotrademark.Fields("cp02").Value) = False Then
         stra0j02 = stra0j02 & adotrademark.Fields("cp02").Value
      End If
      If IsNull(adotrademark.Fields("cp03").Value) = False Then
         stra0j02 = stra0j02 & adotrademark.Fields("cp03").Value
      End If
      If IsNull(adotrademark.Fields("cp04").Value) = False Then
         stra0j02 = stra0j02 & adotrademark.Fields("cp04").Value
      End If
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j03
   'If IsNull(adotrademark.Fields("cp10").Value) Then
   '   lnga0j03 = MsgText(601)
   'Else
   '   lnga0j03 = adotrademark.Fields("cp10").Value
   'End If
   
   If IsNull(adotrademark.Fields("tm10").Value) Then
      stra0j04 = MsgText(601)
   Else
      stra0j04 = adotrademark.Fields("tm10").Value
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j05
   'If IsNull(adotrademark.Fields("cp13").Value) Then
   '   stra0j05 = MsgText(601)
   'Else
   '   stra0j05 = adotrademark.Fields("cp13").Value
   'End If
   
'   If Mid(adotrademark.Fields("cp01").Value, 1, 2) = "CF" Then
'      stra0j07 = MsgText(602)
'   Else
      If adotrademark.Fields("tm10").Value <> "000" Then
         stra0j07 = MsgText(602)
      Else
         stra0j07 = MsgText(601)
      End If
'   End If
   If IsNull(adotrademark.Fields("cp16").Value) Then
      lnga0j09 = 0
   Else
      If IsNull(adotrademark.Fields("cp17").Value) Then
         lnga0j09 = adotrademark.Fields("cp16").Value
      Else
         lnga0j09 = Val(adotrademark.Fields("cp16").Value) - Val(adotrademark.Fields("cp17").Value)
      End If
   End If
   If IsNull(adotrademark.Fields("cp17").Value) Then
      lnga0j10 = 0
   Else
      lnga0j10 = adotrademark.Fields("cp17").Value
   End If
   'If IsNull(adotrademark.Fields("tm23").Value) Then
   If Text1 = MsgText(601) Then
      stra0j11 = MsgText(601)
   Else
   '   stra0j11 = adotrademark.Fields("tm23").Value
      stra0j11 = Text1
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j12
   'If IsNull(adotrademark.Fields("cp05").Value) = False Then
   '   If adotrademark.Fields("cp05").Value <> 0 Then
   '      stra0j12 = ACDate(adotrademark.Fields("cp05").Value)
   '   Else
   '      stra0j12 = ""
   '   End If
   'Else
   '   stra0j12 = ""
   'End If
   
   'Removed by Morgan 2011/12/27 取消 a0j20
   'If stra0j04 = "000" Then
   '   If InStr(adotrademark.Fields("cpm03").Value, MsgText(173)) > 0 Then
   '      If IsNull(adotrademark.Fields("cpm04").Value) = False Then
   '         stra0j20 = adotrademark.Fields("cpm04").Value
   '      Else
   '         stra0j20 = ""
   '      End If
   '   Else
   '      stra0j20 = adotrademark.Fields("cpm03").Value
   '   End If
   'Else
   '   If IsNull(adotrademark.Fields("cpm04").Value) = False Then
   '      stra0j20 = adotrademark.Fields("cpm04").Value
   '   Else
   '      stra0j20 = ""
   '   End If
   'End If
   
End Sub

'*************************************************
'  重新顯示 Adodc 之內容
'
'*************************************************
Private Sub AdodcRefresh()
Dim strVal As String 'Add By Sindy 2020/3/24

On Error GoTo Checking
   
   'Add By Sindy 2020/3/24
   If strSrvDate(1) >= 智慧所更名日 Then
      'Modify By Sindy 2020/4/10 ACS不屬於L
      'strVal = "decode(sk02,'3','L','4','L','7','L','8','L',pa161||tm130||sp85||lc48) pa161"
      strVal = "decode(instr(cp01,'L'),'0',pa161||tm130||sp85||lc48,'L') pa161"
   Else
      strVal = "pa161||tm130||sp85||lc48 pa161"
   End If
   '2020/3/24 END
   
   strY = ""
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   '2010/2/12 MODIFY BY SONIA 加P及CFP的601,605,606,607顯示繳年費年度
   'adoadodc1.Open "select * from acc0j0 where a0j13 is null and a0j11 like '" & Text1 & "%" & "' order by a0j02 asc, a0j05 asc, a0j01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Morgan 2011/9/19 a0j13 改先放收文號
   'adoadodc1.Open "select acc0j0.*,DECODE(CP53,'',NULL,DECODE(CP01||CP10,'P601',CP53||'-'||CP54,'P605',CP53||'-'||CP54,'P606',CP53||'-'||CP54,'P607',CP53||'-'||CP54,'CFP601',CP53||'-'||CP54,'CFP605',CP53||'-'||CP54,'CFP606',CP53||'-'||CP54,'CFP607',CP53||'-'||CP54,NULL)) CP5354 from acc0j0,CASEPROGRESS where a0j13 is null and a0j11 like '" & Text1 & "%" & "' AND A0J01=CP09(+) order by a0j02 asc, a0j05 asc, a0j01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Morgan 2011/12/26 +CP10,CP13;a0j05 改抓 cp13
   'Modified by Morgan 2012/1/2 取消 a0j20,a0j21
   'Modify By Sindy 2013/12/20 增加顯示特殊出名公司
   'adoadodc1.Open "select acc0j0.*,DECODE(CP53,'',NULL,DECODE(CP01||CP10,'P601',CP53||'-'||CP54,'P605',CP53||'-'||CP54,'P606',CP53||'-'||CP54,'P607',CP53||'-'||CP54,'CFP601',CP53||'-'||CP54,'CFP605',CP53||'-'||CP54,'CFP606',CP53||'-'||CP54,'CFP607',CP53||'-'||CP54,NULL)) CP5354,CP10,CP13,cp05-19110000 RDate,getcp10desc(cp01,cp10,a0j04) cp10N,na03  from acc0j0,CASEPROGRESS,nation where a0j13=a0j01 and a0j11 like '" & Text1 & "%" & "' AND A0J01=CP09(+) and na01(+)=a0j04 order by a0j02 asc, cp13 asc, a0j01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify By Sindy 2014/2/10 +CRL119,cp140,CRL49
   'Modify By Sindy 2020/3/24 + systemkind
   '   pa161||tm130||sp85||lc48 pa161=>strVal
   'Modify By Sindy 2020/3/31 + ,CRL02
   'Modify By Sindy 2023/6/15 and a0j11 like '" & Text1 & "%" & "' => and a0j11 like '" & Left(Text1, 8) & "%" & "'
   strSql = "select acc0j0.*," & strVal & ",DECODE(CP53,'',NULL,DECODE(CP01||CP10,'P601',CP53||'-'||CP54,'P605',CP53||'-'||CP54,'P606',CP53||'-'||CP54,'P607',CP53||'-'||CP54,'CFP601',CP53||'-'||CP54,'CFP605',CP53||'-'||CP54,'CFP606',CP53||'-'||CP54,'CFP607',CP53||'-'||CP54,NULL)) CP5354,CP10,CP13,cp05-19110000 RDate,getcp10desc(cp01,cp10,a0j04) cp10N,na03,CRL119,cp140,CRL49,CRL02" & _
            " from acc0j0,CASEPROGRESS,nation,patent,trademark,servicepractice,lawcase,hirecase,consultrecordlist,systemkind" & _
            " where a0j13=a0j01 and a0j11 like '" & Left(Text1, 8) & "%" & "' AND A0J01=CP09(+) and na01(+)=a0j04 and cp140=crl01(+)" & _
            " and substr(A0J02,1,length(A0J02)-9)=pa01(+) and substr(A0J02,(length(A0J02)-9)+1,6)=pa02(+) and substr(A0J02,(length(A0J02)-3)+1,1)=pa03(+) and substr(A0J02,(length(A0J02)-2)+1,2)=pa04(+)" & _
            " and substr(A0J02,1,length(A0J02)-9)=tm01(+) and substr(A0J02,(length(A0J02)-9)+1,6)=tm02(+) and substr(A0J02,(length(A0J02)-3)+1,1)=tm03(+) and substr(A0J02,(length(A0J02)-2)+1,2)=tm04(+)" & _
            " and substr(A0J02,1,length(A0J02)-9)=sp01(+) and substr(A0J02,(length(A0J02)-9)+1,6)=sp02(+) and substr(A0J02,(length(A0J02)-3)+1,1)=sp03(+) and substr(A0J02,(length(A0J02)-2)+1,2)=sp04(+)" & _
            " and substr(A0J02,1,length(A0J02)-9)=lc01(+) and substr(A0J02,(length(A0J02)-9)+1,6)=lc02(+) and substr(A0J02,(length(A0J02)-3)+1,1)=lc03(+) and substr(A0J02,(length(A0J02)-2)+1,2)=lc04(+)" & _
            " and substr(A0J02,1,length(A0J02)-9)=hc01(+) and substr(A0J02,(length(A0J02)-9)+1,6)=hc02(+) and substr(A0J02,(length(A0J02)-3)+1,1)=hc03(+) and substr(A0J02,(length(A0J02)-2)+1,2)=hc04(+)" & _
            " and substr(A0J02,1,length(A0J02)-9)=sk01(+)" & _
            " order by a0j02 asc, cp13 asc, a0j01 asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      If strNoNation = MsgText(602) Then
         MsgBox MsgText(174), , MsgText(5)
         strNoNation = ""
      Else
         If strMsgShow <> MsgText(602) Then
            MsgBox MsgText(165), , MsgText(5)
         End If
      End If
'      Text1 = ""
      strY = MsgText(602)
      Text1.SetFocus
   End If
   strMsgShow = ""
   'Mark by Amy 2014/06/30
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim salesArea As String, salesNo As String 'add by sonia 2017/7/11
   
   If Text1 = "" Then
      Exit Sub
   End If
   Select Case Len(Text1)
      Case 6
         Text1 = Text1 & "000"
      Case 8
         Text1 = Text1 & "0"
   End Select
   If ExistCheck("customer", "cu01", Mid(Text1, 1, 8), "", False) = False Then
      MsgBox MsgText(45) & Label1, , MsgText(5)
      Text1 = "X"
      TextInverse Text1
      Cancel = True
      Text1.SetFocus
      Exit Sub
   'add by sonia 2017/7/11 X54363010(E10614798)會造成收款點數掛收文智權人員
   Else
      salesArea = GetCuSales(Text1, salesNo)
      If Left(salesNo, 4) = "MCTF" Then
         'modify by sonia 2018/8/17 改開放但提醒 T-216118
         'MsgBox "此客戶為商標處大至台客戶, 不可開立國內收據, 請改開國外請款單 !", , MsgText(5)
         'TextInverse Text1
         'Cancel = True
         'Text1.SetFocus
         'Exit Sub
         MsgBox "此客戶為商標處MCT客戶, 國內收據請另外註明點數做 M0100 !", , MsgText(5)
      End If
   'end 2017/7/11
   End If
   Text2 = CustomerQuery(Text1, 1)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoTaie.Execute "update acc0j0 set a0j06 = null where a0j06 = '" & MsgText(602) & "'"
   Acc0j0Delete
   Process
   If strY <> MsgText(601) Then
'      Cancel = True
      Text1.SetFocus
      TextInverse Text1
      Exit Sub
   End If
End Sub

'*************************************************
'  產生並儲存國內收據資料
'
'*************************************************
Private Sub ProduceData()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If adoacc0j0.Fields("a0j08").Value = MsgText(602) Then
      If strReceiptConfirm = MsgText(601) Then
         Exit Sub
      End If
      adoacc0j0.MoveFirst
      Do While adoacc0j0.EOF = False
         'Modify by Morgan 2011/9/19 a0j13改先放收文號
         'adoTaie.Execute "update acc0j0 set a0j13 = '" & strItemNo & "' where a0j01 = '" & adoacc0j0.Fields("a0j01").Value & "' and a0j08 = '" & MsgText(602) & "' and a0j13 is null"
         adoTaie.Execute "update acc0j0 set a0j13 = '" & strItemNo & "' where a0j01 = '" & adoacc0j0.Fields("a0j01").Value & "' and a0j08 = '" & MsgText(602) & "' and a0j13=a0j01"
         adoTaie.Execute "update caseprogress set cp60 = '" & strItemNo & "' where cp09 = '" & adoacc0j0.Fields("a0j01").Value & "'"
         adoTaie.Execute "insert into acc1m0 values ('" & strItemNo & "', '" & adoacc0j0.Fields("a0j01").Value & "')"
         Acc0k0Save
         adoacc0k0.Close
         adoacc0k0.CursorLocation = adUseClient
         adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If adoacc0k0.RecordCount <> 0 Then
            adoacc0k0.Fields("a0k02").Value = lnga0k02
            adoacc0k0.Fields("a0k03").Value = stra0k03
            adoacc0k0.Fields("a0k04").Value = strA0K04
            adoacc0k0.Fields("a0k05").Value = stra0k05
            adoacc0k0.Fields("a0k06").Value = Val(adoacc0k0.Fields("a0k06").Value) + Val(adoacc0j0.Fields("a0j09").Value)
            adoacc0k0.Fields("a0k07").Value = Val(adoacc0k0.Fields("a0k07").Value) + Val(adoacc0j0.Fields("a0j10").Value)
            adoacc0k0.Fields("a0k08").Value = stra0k08
            adoacc0k0.Fields("a0k11").Value = strA0K11
            adoacc0k0.Fields("a0k27").Value = Val(strSrvDate(2))
            adoacc0k0.Fields("a0k28").Value = ServerTime
            adoacc0k0.Fields("a0k29").Value = strUserNum
            adoacc0k0.UpdateBatch
         End If
         adoacc0j0.MoveNext
      Loop
   End If
   strItemNo = ""
Checking:
   strItemNo = ""
   Exit Sub
End Sub

'*************************************************
'  將國內收據所需之資料放置系統變數中
'
'*************************************************
Private Sub Acc0k0Save()
   lnga0k02 = Val(strSrvDate(2))
   stra0k03 = strCustNo
   strA0K04 = strTitle
   stra0k05 = strComPer
   stra0k08 = strRemark
   strA0K11 = strCompanyNo
End Sub

'*************************************************
'  將法務資料放置系統變數中
'
'*************************************************
Private Sub Acc0j0SaveL()
   stra0j01 = adolawcase.Fields("cp09").Value
   If IsNull(adolawcase.Fields("cp01").Value) Then
      stra0j02 = MsgText(601)
   Else
      stra0j02 = adolawcase.Fields("cp01").Value
      If IsNull(adolawcase.Fields("cp02").Value) = False Then
         stra0j02 = stra0j02 & adolawcase.Fields("cp02").Value
      End If
      If IsNull(adolawcase.Fields("cp03").Value) = False Then
         stra0j02 = stra0j02 & adolawcase.Fields("cp03").Value
      End If
      If IsNull(adolawcase.Fields("cp04").Value) = False Then
         stra0j02 = stra0j02 & adolawcase.Fields("cp04").Value
      End If
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j03
   'If IsNull(adolawcase.Fields("cp10").Value) Then
   '   lnga0j03 = MsgText(601)
   'Else
   '   lnga0j03 = adolawcase.Fields("cp10").Value
   'End If
   
   If IsNull(adolawcase.Fields("lc15").Value) Then
      stra0j04 = MsgText(601)
   Else
      stra0j04 = adolawcase.Fields("lc15").Value
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j05
   'If IsNull(adolawcase.Fields("cp13").Value) Then
   '   stra0j05 = MsgText(601)
   'Else
   '   stra0j05 = adolawcase.Fields("cp13").Value
   'End If
   
'   If Mid(adolawcase.Fields("cp01").Value, 1, 2) = "CF" Then
'      stra0j07 = MsgText(602)
'   Else
   If adolawcase.Fields("lc15").Value <> "000" Then
      stra0j07 = MsgText(602)
   Else
      stra0j07 = MsgText(601)
   End If
'   End If
   'add by sonia 2016/4/12 L所有案件性質預設合併-婷
   'cancel by sonia 2024/8/13 瑞婷
   'If adolawcase.Fields("cp01").Value = "L" Then
   '   stra0j07 = MsgText(602)
   'End If
   'end 2024/8/13
   'end 2016/4/12
   
   'add by sonia 2020/5/26 法律所案源為B,C類者,不合併
   Set AdoRecordSet3 = Adodc1.Recordset.Clone
   With AdoRecordSet3
      strSql = "SELECT LOS02 FROM CASEPROGRESS,LAWOFFICESOURCE WHERE CP09='" & adolawcase.Fields("cp09").Value & "' AND CP09=LOS06(+)"
      CheckOC3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 And "" & .Fields("LOS02") > "B" Then
         stra0j07 = ""
      End If
   End With
   'end 2020/5/26

   If IsNull(adolawcase.Fields("cp16").Value) Then
      lnga0j09 = 0
   Else
      If IsNull(adolawcase.Fields("cp17").Value) Then
         lnga0j09 = adolawcase.Fields("cp16").Value
      Else
         lnga0j09 = Val(adolawcase.Fields("cp16").Value) - Val(adolawcase.Fields("cp17").Value)
      End If
   End If
   If IsNull(adolawcase.Fields("cp17").Value) Then
      lnga0j10 = 0
   Else
      lnga0j10 = adolawcase.Fields("cp17").Value
   End If
   'If IsNull(adolawcase.Fields("lc11").Value) Then
   If Text1 = MsgText(601) Then
      stra0j11 = MsgText(601)
   Else
   '   stra0j11 = Mid(adolawcase.Fields("lc11").Value, 1, 8) & "0"
      stra0j11 = Text1
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j12
   'If IsNull(adolawcase.Fields("cp05").Value) = False Then
   '   If adolawcase.Fields("cp05").Value <> 0 Then
   '      stra0j12 = ACDate(adolawcase.Fields("cp05").Value)
   '   Else
   '      stra0j12 = ""
   '   End If
   'Else
   '   stra0j12 = ""
   'End If
   
   'Removed by Morgan 2011/12/27 取消 a0j20
   'If stra0j04 = "000" Then
   '   If InStr(adolawcase.Fields("cpm03").Value, MsgText(173)) > 0 Then
   '      If IsNull(adolawcase.Fields("cpm04").Value) = False Then
   '         stra0j20 = adolawcase.Fields("cpm04").Value
   '      Else
   '         stra0j20 = ""
   '      End If
   '   Else
   '      stra0j20 = adolawcase.Fields("cpm03").Value
   '   End If
   'Else
   '   If IsNull(adolawcase.Fields("cpm04").Value) = False Then
   '      stra0j20 = adolawcase.Fields("cpm04").Value
   '   Else
   '      stra0j20 = ""
   '   End If
   'End If
   
End Sub

'*************************************************
'  將顧問資料放置系統變數中
'
'*************************************************
Private Sub Acc0j0SaveLA()
   stra0j01 = adohirecase.Fields("cp09").Value
   If IsNull(adohirecase.Fields("cp01").Value) Then
      stra0j02 = MsgText(601)
   Else
      stra0j02 = adohirecase.Fields("cp01").Value
      If IsNull(adohirecase.Fields("cp02").Value) = False Then
         stra0j02 = stra0j02 & adohirecase.Fields("cp02").Value
      End If
      If IsNull(adohirecase.Fields("cp03").Value) = False Then
         stra0j02 = stra0j02 & adohirecase.Fields("cp03").Value
      End If
      If IsNull(adohirecase.Fields("cp04").Value) = False Then
         stra0j02 = stra0j02 & adohirecase.Fields("cp04").Value
      End If
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j03
   'If IsNull(adohirecase.Fields("cp10").Value) Then
   '   lnga0j03 = MsgText(601)
   'Else
   '   lnga0j03 = adohirecase.Fields("cp10").Value
   'End If
   
   stra0j04 = "000"
   
   'Removed by Morgan 2011/12/26 取消 a0j05
   'If IsNull(adohirecase.Fields("cp13").Value) Then
   '   stra0j05 = MsgText(601)
   'Else
   '   stra0j05 = adohirecase.Fields("cp13").Value
   'End If
   
   If IsNull(adohirecase.Fields("cp16").Value) Then
      lnga0j09 = 0
   Else
      If IsNull(adohirecase.Fields("cp17").Value) Then
         lnga0j09 = adohirecase.Fields("cp16").Value
      Else
         lnga0j09 = Val(adohirecase.Fields("cp16").Value) - Val(adohirecase.Fields("cp17").Value)
      End If
   End If
   If IsNull(adohirecase.Fields("cp17").Value) Then
      lnga0j10 = 0
   Else
      lnga0j10 = adohirecase.Fields("cp17").Value
   End If
   'If IsNull(adohirecase.Fields("hc05").Value) Then
   If Text1 = MsgText(601) Then
      stra0j11 = MsgText(601)
   Else
   '   stra0j11 = Mid(adohirecase.Fields("hc05").Value, 1, 8) & "0"
      stra0j11 = Text1
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j12
   'If IsNull(adohirecase.Fields("cp05").Value) = False Then
   '   If adohirecase.Fields("cp05").Value <> 0 Then
   '      stra0j12 = ACDate(adohirecase.Fields("cp05").Value)
   '   Else
   '      stra0j12 = ""
   '   End If
   'Else
   '   stra0j12 = ""
   'End If
   
   'Removed by Morgan 2011/12/27 取消 a0j20
   'If stra0j04 = "000" Then
   '   If InStr(adohirecase.Fields("cpm03").Value, MsgText(173)) > 0 Then
   '      If IsNull(adohirecase.Fields("cpm04").Value) = False Then
   '         stra0j20 = adohirecase.Fields("cpm04").Value
   '      Else
   '         stra0j20 = ""
   '      End If
   '   Else
   '      stra0j20 = adohirecase.Fields("cpm03").Value
   '   End If
   'Else
   '   If IsNull(adohirecase.Fields("cpm04").Value) = False Then
   '      stra0j20 = adohirecase.Fields("cpm04").Value
   '   Else
   '      stra0j20 = ""
   '   End If
   'End If
   
End Sub

'*************************************************
'  將服務資料放置系統變數中
'
'*************************************************
Private Sub Acc0j0SaveS()
   stra0j01 = adoservice.Fields("cp09").Value
   If IsNull(adoservice.Fields("cp01").Value) Then
      stra0j02 = MsgText(601)
   Else
      stra0j02 = adoservice.Fields("cp01").Value
      If IsNull(adoservice.Fields("cp02").Value) = False Then
         stra0j02 = stra0j02 & adoservice.Fields("cp02").Value
      End If
      If IsNull(adoservice.Fields("cp03").Value) = False Then
         stra0j02 = stra0j02 & adoservice.Fields("cp03").Value
      End If
      If IsNull(adoservice.Fields("cp04").Value) = False Then
         stra0j02 = stra0j02 & adoservice.Fields("cp04").Value
      End If
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j03
   'If IsNull(adoservice.Fields("cp10").Value) Then
   '   lnga0j03 = MsgText(601)
   'Else
   '   lnga0j03 = adoservice.Fields("cp10").Value
   'End If
   
   If IsNull(adoservice.Fields("sp09").Value) Then
      stra0j04 = MsgText(601)
   Else
      stra0j04 = adoservice.Fields("sp09").Value
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j05
   'If IsNull(adoservice.Fields("cp13").Value) Then
   '   stra0j05 = MsgText(601)
   'Else
   '   stra0j05 = adoservice.Fields("cp13").Value
   'End If
   
'   If Mid(adoservice.Fields("cp01").Value, 1, 2) = "CF" Then
'      stra0j07 = MsgText(602)
'   Else
   'MODIFY BY SONIA 2015/6/4 TD大陸名稱也是向台灣網路資訊中心申請,故收據不必合併TD-000153
   If adoservice.Fields("sp09").Value <> "000" And adoservice.Fields("cp01").Value <> "TD" Then
      stra0j07 = MsgText(602)
   Else
      stra0j07 = MsgText(601)
   End If
'   End If
   'Add by Morgan 2004/11/10 TT 的文件簽證711預設合併
   If adoservice.Fields("cp01").Value = "TT" And adoservice.Fields("cp10").Value = "711" Then
      stra0j07 = MsgText(602)
   End If
   '2004/11/2 END
   
   If IsNull(adoservice.Fields("cp16").Value) Then
      lnga0j09 = 0
   Else
      If IsNull(adoservice.Fields("cp17").Value) Then
         lnga0j09 = adoservice.Fields("cp16").Value
      Else
         lnga0j09 = Val(adoservice.Fields("cp16").Value) - Val(adoservice.Fields("cp17").Value)
      End If
   End If
   If IsNull(adoservice.Fields("cp17").Value) Then
      lnga0j10 = 0
   Else
      lnga0j10 = adoservice.Fields("cp17").Value
   End If
   'If IsNull(adoservice.Fields("sp08").Value) Then
   If Text1 = MsgText(601) Then
      stra0j11 = MsgText(601)
   Else
   '   stra0j11 = Mid(adoservice.Fields("sp08").Value, 1, 8) & "0"
      stra0j11 = Text1
   End If
   
   'Removed by Morgan 2011/12/26 取消 a0j12
   'If IsNull(adoservice.Fields("cp05").Value) = False Then
   '   If adoservice.Fields("cp05").Value <> 0 Then
   '      stra0j12 = ACDate(adoservice.Fields("cp05").Value)
   '   Else
   '      stra0j12 = ""
   '   End If
   'Else
   '   stra0j12 = ""
   'End If
   
   'Removed by Morgan 2011/12/27 取消 a0j20
   'If stra0j04 = "000" Then
   '   If InStr(adoservice.Fields("cpm03").Value, MsgText(173)) > 0 Then
   '      If IsNull(adoservice.Fields("cpm04").Value) = False Then
   '         stra0j20 = adoservice.Fields("cpm04").Value
   '      Else
   '         stra0j20 = ""
   '      End If
   '   Else
   '      stra0j20 = adoservice.Fields("cpm03").Value
   '   End If
   'Else
   '   If IsNull(adoservice.Fields("cpm04").Value) = False Then
   '      stra0j20 = adoservice.Fields("cpm04").Value
   '   Else
   '      stra0j20 = ""
   '   End If
   'End If
   
End Sub

'*************************************************
' 刪除資料表 (國內未開收據資料表)
'
'*************************************************
Private Sub Acc0j0Delete()
'   adoTaie.Execute "delete from acc0j0 where (substr(a0j01, 1, 1) = 'C' or substr(a0j02, 1, 2) = 'FC')"
   'Modify by Morgan 2011/9/19 a0j13改先放收文號
   'adoTaie.Execute "delete from acc0j0 where a0j11 = '" & Text1 & "' and (a0j13 is null or a0j13 = '')"
   adoTaie.Execute "delete from acc0j0 where a0j11 = '" & Text1 & "' and a0j13=a0j01"
End Sub

'***************************************************
'  將未開收據之案件儲存於國內未開收據案件資料表內(例外處理)
'
'***************************************************
Private Sub Process1()
Dim strComputer As String

   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   'Modify by Morgan 2011/9/19 a0j13改先放收文號
   'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j11 = '" & Text1 & "'"
   adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j11 = '" & Text1 & "'"
   adopatent.CursorLocation = adUseClient
'   adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                  "and cp57 is null and cp05 >= 20011112 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and (pa26 like '" & Mid(Text1, 1, 8) & "%" & "' or pa27 like '" & Text1 & "%" & "' or pa28 like '" & Text1 & "%" & "' or pa29 like '" & Text1 & "%" & "' or pa30 like '" & Text1 & "%" & "') and (cp60 is null or cp60 = '') and cp10 not in ('907') ", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2014/7/9 只抓取申請人1開立收據
'   adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                  "and cp57 is null and cp05 >= 20030101 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and (pa26 = '" & Text1 & "' or pa27 = '" & Text1 & "' or pa28 = '" & Text1 & "' or pa29 = '" & Text1 & "' or pa30 = '" & Text1 & "') and (cp60 is null or cp60 = '') and cp10 not in ('907') ", adoTaie, adOpenStatic, adLockReadOnly
   adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
                  "and cp57 is null and cp05 >= 20030101 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and pa26 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('907') ", adoTaie, adOpenStatic, adLockReadOnly
   '2014/7/9 END
                  '"union select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
                  '"and cp57 is null and cp05 >= 20011112 and (cp16 is null or cp16 = 0) and (pa26 like '" & Mid(Text1, 1, 8) & "%" & "' or pa27 like '" & Text1 & "%" & "' or pa28 like '" & Text1 & "%" & "' or pa29 like '" & Text1 & "%" & "' or pa30 like '" & Text1 & "%" & "') and (cp60 is null or cp60 = '') and cp10 not in ('907')", adoTaie, adOpenStatic, adLockReadOnly
   Do While adopatent.EOF = False
      'Modify by Morgan 2011/9/19 a0j13改先放收文號
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adopatent.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adopatent.Fields("cp09").Value & "'"
      If adopatent.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveP
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select cu10, na03, pa75 from patent, customer, nation where substr(pa26, 1, 8) = cu01 and cu02 = '0' and pa09 = na01 and pa01 = '" & adopatent.Fields("pa01").Value & "' and pa02 = '" & adopatent.Fields("pa02").Value & "' and pa03 = '" & adopatent.Fields("pa03").Value & "' and pa04 = '" & adopatent.Fields("pa04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         
         'Modified by Morgan 2011/12/29 取消 a0j21
         'If IsNull(adocheck.Fields(1).Value) Then
         '   stra0j21 = ""
         'Else
         '   stra0j21 = adocheck.Fields(1).Value
         'End If
         
'         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2011/3/24 改判斷客戶國籍為台灣的FC案件才對
            'If adocheck.Fields(0).Value > 10 Or IsNull(adocheck.Fields(0).Value) Then
            If Val("" & adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  If adopatent.Fields("pa09").Value = "020" Then
                     stra0j07 = MsgText(602)
                  End If
                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11,  a0j14, a0j15, a0j16, a0j08, a0j07,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & stra0j04 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & strComputer & "', '" & stra0j07 & "','" & stra0j01 & "')"
               End If
            End If
'         End If
      End If
      adocheck.Close
      adopatent.MoveNext
   Loop
   adopatent.Close
   adotrademark.CursorLocation = adUseClient
'   adotrademark.Open "select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                     "and cp57 is null and cp05 >= 20011112 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and tm23 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('703') ", adoTaie, adOpenStatic, adLockReadOnly
   adotrademark.Open "select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
                     "and cp57 is null and cp05 >= 20030101 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and tm23 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('703') ", adoTaie, adOpenStatic, adLockReadOnly
                     '"union select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
                     '"and cp57 is null and cp05 >= 20011112 and (cp16 is null or cp16 = 0) and tm23 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('703')", adoTaie, adOpenStatic, adLockReadOnly
   Do While adotrademark.EOF = False
      'Modify by Morgan 2011/9/19 a0j13改先放收文號
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adotrademark.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adotrademark.Fields("cp09").Value & "'"
      If adotrademark.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveT
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select cu10, na03,tm44 from trademark, customer, nation where substr(tm23, 1, 8) = cu01 and cu02 = '0' and tm10 = na01 and tm01 = '" & adotrademark.Fields("tm01").Value & "' and tm02 = '" & adotrademark.Fields("tm02").Value & "' and tm03 = '" & adotrademark.Fields("tm03").Value & "' and tm04 = '" & adotrademark.Fields("tm04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         'Modified by Morgan 2011/12/29 取消 a0j21
         'If IsNull(adocheck.Fields(1).Value) Then
         '   stra0j21 = ""
         'Else
         '   stra0j21 = adocheck.Fields(1).Value
         'End If
         
'         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2011/3/24 改判斷客戶國籍為台灣的FC案件才對
            'If Val(adocheck.Fields(0).Value) > 10 Or IsNull(adocheck.Fields(0).Value) Then
            If Val("" & adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  If adotrademark.Fields("tm10").Value = "020" Then
                     stra0j07 = MsgText(602)
                  End If
                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11, a0j14, a0j15, a0j16, a0j08, a0j07,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & stra0j04 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & strComputer & "', '" & stra0j07 & "','" & stra0j01 & "')"
               End If
            End If
'         End If
      End If
      adocheck.Close
      adotrademark.MoveNext
   Loop
   adotrademark.Close
   adolawcase.CursorLocation = adUseClient
'   adolawcase.Open "select * from caseprogress, lawcase, casepropertymap where (caseprogress.cp01 = lawcase.lc01  and caseprogress.cp02 = lawcase.lc02 and caseprogress.cp03 = lawcase.lc03 and caseprogress.cp04 = lawcase.lc04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                   "and cp57 is null and cp05 >= 20011112 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and lc11 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('999') ", adoTaie, adOpenStatic, adLockReadOnly
   adolawcase.Open "select * from caseprogress, lawcase, casepropertymap where (caseprogress.cp01 = lawcase.lc01  and caseprogress.cp02 = lawcase.lc02 and caseprogress.cp03 = lawcase.lc03 and caseprogress.cp04 = lawcase.lc04) and cp01 = cpm01 and cp10 = cpm02 " & _
                   "and cp57 is null and cp05 >= 20030101 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and lc11 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('999') ", adoTaie, adOpenStatic, adLockReadOnly
                   '"union select * from caseprogress, lawcase, casepropertymap where (caseprogress.cp01 = lawcase.lc01  and caseprogress.cp02 = lawcase.lc02 and caseprogress.cp03 = lawcase.lc03 and caseprogress.cp04 = lawcase.lc04) and cp01 = cpm01 and cp10 = cpm02 " & _
                   '"and cp57 is null and cp05 >= 20011112 and (cp16 is null or cp16 = 0) and lc11 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('999')", adoTaie, adOpenStatic, adLockReadOnly
   Do While adolawcase.EOF = False
      'Modify by Morgan 2011/9/19 a0j13改先放收文號
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adolawcase.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adolawcase.Fields("cp09").Value & "'"
      If adolawcase.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveL
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select cu10, na03,lc22 from lawcase, customer, nation where substr(lc11, 1, 8) = cu01 and cu02 = '0' and lc15 = na01 and lc01 = '" & adolawcase.Fields("lc01").Value & "' and lc02 = '" & adolawcase.Fields("lc02").Value & "' and lc03 = '" & adolawcase.Fields("lc03").Value & "' and lc04 = '" & adolawcase.Fields("lc04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
      
         'Modified by Morgan 2011/12/29 取消 a0j21
         'If IsNull(adocheck.Fields(1).Value) Then
         '   stra0j21 = ""
         'Else
         '   stra0j21 = adocheck.Fields(1).Value
         'End If
         
'         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2011/3/24 改判斷客戶國籍為台灣的FC案件才對
            'If Val(adocheck.Fields(0).Value) > 10 Or IsNull(adocheck.Fields(0).Value) Then
            If Val("" & adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  If adolawcase.Fields("lc15").Value = "020" Then
                     stra0j07 = MsgText(602)
                  End If
                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11, a0j14, a0j15, a0j16, a0j08, a0j07,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & stra0j04 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & strComputer & "', '" & stra0j07 & "','" & stra0j01 & "')"
               End If
            End If
'         End If
      End If
      adocheck.Close
      adolawcase.MoveNext
   Loop
   adolawcase.Close
   
'Remove by Morgan 2011/3/24 顧問案件調整條件後不會出現在例外
'   adohirecase.CursorLocation = adUseClient
''   adohirecase.Open "select * from caseprogress, hirecase, casepropertymap where (caseprogress.cp01 = hirecase.hc01  and caseprogress.cp02 = hirecase.hc02 and caseprogress.cp03 = hirecase.hc03 and caseprogress.cp04 = hirecase.hc04) and cp01 = cpm01 and cp10 = cpm02 " & _
''                    "and cp57 is null and cp05 >= 20011112 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and hc05 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('999') ", adoTaie, adOpenStatic, adLockReadOnly
'   adohirecase.Open "select * from caseprogress, hirecase, casepropertymap where (caseprogress.cp01 = hirecase.hc01  and caseprogress.cp02 = hirecase.hc02 and caseprogress.cp03 = hirecase.hc03 and caseprogress.cp04 = hirecase.hc04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                    "and cp57 is null and cp05 >= 20030101 and (cp16 is not null and cp16 <> 0) and substr(cp01, 1, 2) = 'FC' and hc05 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('999') ", adoTaie, adOpenStatic, adLockReadOnly
'                    '"union select * from caseprogress, hirecase, casepropertymap where (caseprogress.cp01 = hirecase.hc01  and caseprogress.cp02 = hirecase.hc02 and caseprogress.cp03 = hirecase.hc03 and caseprogress.cp04 = hirecase.hc04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                    '"and cp57 is null and cp05 >= 20011112 and (cp16 is null or cp16 = 0) and hc05 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('999')", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adohirecase.EOF = False
'      adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adohirecase.Fields("cp09").Value & "'"
'      If adohirecase.Fields("cp32").Value = MsgText(603) Then
'         strComputer = MsgText(602)
'      Else
'         strComputer = MsgText(601)
'      End If
'      Acc0j0SaveLA
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select cu10 from hirecase, customer where substr(hc05, 1, 8) = cu01 and cu02 = '0' and hc01 = '" & adohirecase.Fields("hc01").Value & "' and hc02 = '" & adohirecase.Fields("hc02").Value & "' and hc03 = '" & adohirecase.Fields("hc03").Value & "' and hc04 = '" & adohirecase.Fields("hc04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount <> 0 Then
'         stra0j21 = "台灣"
''         If IsNull(adocheck.Fields(0).Value) = False Then
'            If Val(adocheck.Fields(0).Value) > 10 Or IsNull(adocheck.Fields(0).Value) Then
'               adoacc0j0.Close
'               adoacc0j0.CursorLocation = adUseClient
'               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
'               If adoacc0j0.RecordCount = 0 Then
'                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j03, a0j04, a0j05, a0j09, a0j10, a0j11, a0j12, a0j14, a0j15, a0j16, a0j20, a0j21, a0j08) " & _
'                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & lnga0j03 & "', '000', '" & stra0j05 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & Val(stra0j12) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & stra0j20 & "', '台灣', '" & strComputer & "')"
'               End If
'            End If
''         End If
'      End If
'      adocheck.Close
'      adohirecase.MoveNext
'   Loop
'   adohirecase.Close
'end 2011/3/24

   adoservice.CursorLocation = adUseClient
'   adoservice.Open "select * from caseprogress, servicepractice, casepropertymap where (caseprogress.cp01 = sp01  and caseprogress.cp02 = sp02 and caseprogress.cp03 = sp03 and caseprogress.cp04 = sp04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                   "and cp57 is null and cp05 >= 20011112 and (cp16 is not null and cp16 <> 0) and sp08 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '')", adoTaie, adOpenStatic, adLockReadOnly
   adoservice.Open "select * from caseprogress, servicepractice, casepropertymap where (caseprogress.cp01 = sp01  and caseprogress.cp02 = sp02 and caseprogress.cp03 = sp03 and caseprogress.cp04 = sp04) and cp01 = cpm01 and cp10 = cpm02 " & _
                   "and cp57 is null and cp05 >= 20030101 and (cp16 is not null and cp16 <> 0) and sp08 = '" & Text1 & "' and (cp60 is null or cp60 = '')", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoservice.EOF = False
      'Modify by Morgan 2011/9/19 a0j13改先放收文號
      'adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adoservice.Fields("cp09").Value & "'"
      adoTaie.Execute "delete from acc0j0 where a0j13=a0j01 and a0j01 = '" & adoservice.Fields("cp09").Value & "'"
      
      If adoservice.Fields("cp32").Value = MsgText(603) Then
         strComputer = MsgText(602)
      Else
         strComputer = MsgText(601)
      End If
      Acc0j0SaveS
      adocheck.CursorLocation = adUseClient
        'Modify By Cheng 2004/05/05
        '加國名(NA03)
'      adocheck.Open "select cu10 from servicepractice, customer where substr(sp08, 1, 8) = cu01 and cu02 = '0' and sp01 = '" & adoservice.Fields("sp01").Value & "' and sp02 = '" & adoservice.Fields("sp02").Value & "' and sp03 = '" & adoservice.Fields("sp03").Value & "' and sp04 = '" & adoservice.Fields("sp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adocheck.Open "select cu10, NA03,sp26 from servicepractice, customer, Nation where substr(sp08, 1, 8) = cu01 and cu02 = '0' And SP09=NA01 and sp01 = '" & adoservice.Fields("sp01").Value & "' and sp02 = '" & adoservice.Fields("sp02").Value & "' and sp03 = '" & adoservice.Fields("sp03").Value & "' and sp04 = '" & adoservice.Fields("sp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
        'End
      If adocheck.RecordCount <> 0 Then
         'stra0j21 = "" & adocheck.Fields(1).Value 'Removed by Morgan 2011/12/29 取消 a0j21

'         If IsNull(adocheck.Fields(0).Value) = False Then
            'Modify by Morgan 2011/3/24 改判斷客戶國籍為台灣的FC案件才對
            'If Val(adocheck.Fields(0).Value) > 10 Or IsNull(adocheck.Fields(0).Value) Then
            If Val("" & adocheck.Fields(0).Value) < 11 Then
               adoacc0j0.Close
               adoacc0j0.CursorLocation = adUseClient
               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0j0.RecordCount = 0 Then
                  '2005/5/23 MODIFY BY SONIA 加 ACC0J0
                  'adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j03, a0j04, a0j05, a0j09, a0j10, a0j11, a0j12, a0j14, a0j15, a0j16, a0j20, a0j21, a0j08) " & _
                  '                "values ('" & stra0j01 & "', '" & stra0j02 & "', " & lnga0j03 & ", '" & stra0j04 & "', '" & stra0j05 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & Val(stra0j12) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & stra0j20 & "', '" & stra0j21 & "', '" & strComputer & "')"
                  'Modify by Morgan 2011/9/19 +a0j13(改key,先放收文號)
                  'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j04, a0j09, a0j10, a0j11, a0j14, a0j15, a0j16, a0j08, a0j07,a0j13) " & _
                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & stra0j04 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "','" & strComputer & "', '" & stra0j07 & "','" & stra0j01 & "')"
                  '2005/5/23 END
               End If
            End If
'         End If
      End If
      adocheck.Close
      adoservice.MoveNext
   Loop
   adoservice.Close
   
'Remove by Morgan 2011/3/24 非FC案應皆為智權人員收文不再為例外
'   adopatent.CursorLocation = adUseClient
''   adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
''                  "and cp57 is null and cp05 >= 20011112 and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and (pa26 like '" & Mid(Text1, 1, 8) & "%" & "' or pa27 like '" & Text1 & "%" & "' or pa28 like '" & Text1 & "%" & "' or pa29 like '" & Text1 & "%" & "' or pa30 like '" & Text1 & "%" & "') and (cp60 is null or cp60 = '') and cp10 not in ('907')", adoTaie, adOpenStatic, adLockReadOnly
'   adopatent.Open "select * from caseprogress, patent, casepropertymap where (caseprogress.cp01 = patent.pa01 and caseprogress.cp02 = patent.pa02 and caseprogress.cp03 = patent.pa03 and caseprogress.cp04 = patent.pa04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                  "and cp57 is null and cp05 >= 20030101 and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and (pa26 = '" & Text1 & "' or pa27 = '" & Text1 & "' or pa28 = '" & Text1 & "' or pa29 = '" & Text1 & "' or pa30 = '" & Text1 & "') and (cp60 is null or cp60 = '') and cp10 not in ('907')", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adopatent.EOF = False
'      adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adopatent.Fields("cp09").Value & "'"
'      If adopatent.Fields("cp32").Value = MsgText(603) Then
'         strComputer = MsgText(602)
'      Else
'         strComputer = MsgText(601)
'      End If
'      Acc0j0SaveP
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select cu10, na03 from patent, customer, nation where substr(pa26, 1, 8) = cu01 and cu02 = '0' and pa09 = na01 and pa01 = '" & adopatent.Fields("pa01").Value & "' and pa02 = '" & adopatent.Fields("pa02").Value & "' and pa03 = '" & adopatent.Fields("pa03").Value & "' and pa04 = '" & adopatent.Fields("pa04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount <> 0 Then
'         If IsNull(adocheck.Fields(1).Value) Then
'            stra0j21 = ""
'         Else
'            stra0j21 = adocheck.Fields(1).Value
'         End If
'         If IsNull(adocheck.Fields(0).Value) = False Then
'            If Val(adocheck.Fields(0).Value) > 10 Then
'               adoacc0j0.Close
'               adoacc0j0.CursorLocation = adUseClient
'               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
'               If adoacc0j0.RecordCount = 0 Then
'                  If adopatent.Fields("pa09").Value = "020" Then
'                     stra0j07 = MsgText(602)
'                  End If
'                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j03, a0j04, a0j05, a0j09, a0j10, a0j11, a0j12, a0j14, a0j15, a0j16, a0j20, a0j21, a0j08, a0j07) " & _
'                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & lnga0j03 & "', '" & stra0j04 & "', '" & stra0j05 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & Val(stra0j12) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & stra0j20 & "', '" & stra0j21 & "', '" & strComputer & "', '" & stra0j07 & "')"
'               End If
'            End If
'         End If
'      End If
'      adocheck.Close
'      adopatent.MoveNext
'   Loop
'   adopatent.Close
'   adotrademark.CursorLocation = adUseClient
''   adotrademark.Open "select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
''                     "and cp57 is null and cp05 >= 20011112 and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and tm23 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('703')", adoTaie, adOpenStatic, adLockReadOnly
'   adotrademark.Open "select * from caseprogress, trademark, casepropertymap where (caseprogress.cp01 = trademark.tm01  and caseprogress.cp02 = trademark.tm02 and caseprogress.cp03 = trademark.tm03 and caseprogress.cp04 = trademark.tm04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                     "and cp57 is null and cp05 >= 20030101 and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and tm23 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('703')", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adotrademark.EOF = False
'      adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adotrademark.Fields("cp09").Value & "'"
'      If adotrademark.Fields("cp32").Value = MsgText(603) Then
'         strComputer = MsgText(602)
'      Else
'         strComputer = MsgText(601)
'      End If
'      Acc0j0SaveT
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select cu10, na03 from trademark, customer, nation where substr(tm23, 1, 8) = cu01 and cu02 = '0' and tm10 = na01 and tm01 = '" & adotrademark.Fields("tm01").Value & "' and tm02 = '" & adotrademark.Fields("tm02").Value & "' and tm03 = '" & adotrademark.Fields("tm03").Value & "' and tm04 = '" & adotrademark.Fields("tm04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount <> 0 Then
'         If IsNull(adocheck.Fields(1).Value) Then
'            stra0j21 = ""
'         Else
'            stra0j21 = adocheck.Fields(1).Value
'         End If
'         If IsNull(adocheck.Fields(0).Value) = False Then
'            If Val(adocheck.Fields(0).Value) > 10 Then
'               adoacc0j0.Close
'               adoacc0j0.CursorLocation = adUseClient
'               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
'               If adoacc0j0.RecordCount = 0 Then
'                  If adotrademark.Fields("tm10").Value = "020" Then
'                     stra0j07 = MsgText(602)
'                  End If
'                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j03, a0j04, a0j05, a0j09, a0j10, a0j11, a0j12, a0j14, a0j15, a0j16, a0j20, a0j21, a0j08, a0j07) " & _
'                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & lnga0j03 & "', '" & stra0j04 & "', '" & stra0j05 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & Val(stra0j12) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & stra0j20 & "', '" & stra0j21 & "', '" & strComputer & "', '" & stra0j07 & "')"
'               End If
'            End If
'         End If
'      End If
'      adocheck.Close
'      adotrademark.MoveNext
'   Loop
'   adotrademark.Close
'   adolawcase.CursorLocation = adUseClient
''   adolawcase.Open "select * from caseprogress, lawcase, casepropertymap where (caseprogress.cp01 = lawcase.lc01  and caseprogress.cp02 = lawcase.lc02 and caseprogress.cp03 = lawcase.lc03 and caseprogress.cp04 = lawcase.lc04) and cp01 = cpm01 and cp10 = cpm02 " & _
''                   "and cp57 is null and cp05 >= 20011112 and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and lc11 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('999')", adoTaie, adOpenStatic, adLockReadOnly
'   adolawcase.Open "select * from caseprogress, lawcase, casepropertymap where (caseprogress.cp01 = lawcase.lc01  and caseprogress.cp02 = lawcase.lc02 and caseprogress.cp03 = lawcase.lc03 and caseprogress.cp04 = lawcase.lc04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                   "and cp57 is null and cp05 >= 20030101 and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and lc11 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('999')", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adolawcase.EOF = False
'      adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adolawcase.Fields("cp09").Value & "'"
'      If adolawcase.Fields("cp32").Value = MsgText(603) Then
'         strComputer = MsgText(602)
'      Else
'         strComputer = MsgText(601)
'      End If
'      Acc0j0SaveL
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select cu10, na03 from lawcase, customer, nation where substr(lc11, 1, 8) = cu01 and cu02 = '0' and lc15 = na01 and lc01 = '" & adolawcase.Fields("lc01").Value & "' and lc02 = '" & adolawcase.Fields("lc02").Value & "' and lc03 = '" & adolawcase.Fields("lc03").Value & "' and lc04 = '" & adolawcase.Fields("lc04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount <> 0 Then
'         If IsNull(adocheck.Fields(1).Value) Then
'            stra0j21 = ""
'         Else
'            stra0j21 = adocheck.Fields(1).Value
'         End If
'         If IsNull(adocheck.Fields(0).Value) = False Then
'            If Val(adocheck.Fields(0).Value) > 10 Then
'               adoacc0j0.Close
'               adoacc0j0.CursorLocation = adUseClient
'               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
'               If adoacc0j0.RecordCount = 0 Then
'                  If adolawcase.Fields("lc15").Value = "020" Then
'                     stra0j07 = MsgText(602)
'                  End If
'                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j03, a0j04, a0j05, a0j09, a0j10, a0j11, a0j12, a0j14, a0j15, a0j16, a0j20, a0j21, a0j08, a0j07) " & _
'                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & lnga0j03 & "', '" & stra0j04 & "', '" & stra0j05 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & Val(stra0j12) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & stra0j20 & "', '" & stra0j21 & "', '" & strComputer & "', '" & stra0j07 & "')"
'               End If
'            End If
'         End If
'      End If
'      adocheck.Close
'      adolawcase.MoveNext
'   Loop
'   adolawcase.Close
'   adohirecase.CursorLocation = adUseClient
''   adohirecase.Open "select * from caseprogress, hirecase, casepropertymap where (caseprogress.cp01 = hirecase.hc01  and caseprogress.cp02 = hirecase.hc02 and caseprogress.cp03 = hirecase.hc03 and caseprogress.cp04 = hirecase.hc04) and cp01 = cpm01 and cp10 = cpm02 " & _
''                     "and cp57 is null and cp05 >= 20011112 and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and hc05 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '') and cp10 not in ('999')", adoTaie, adOpenStatic, adLockReadOnly
'   adohirecase.Open "select * from caseprogress, hirecase, casepropertymap where (caseprogress.cp01 = hirecase.hc01  and caseprogress.cp02 = hirecase.hc02 and caseprogress.cp03 = hirecase.hc03 and caseprogress.cp04 = hirecase.hc04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                     "and cp57 is null and cp05 >= 20030101 and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and hc05 = '" & Text1 & "' and (cp60 is null or cp60 = '') and cp10 not in ('999')", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adohirecase.EOF = False
'      adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adohirecase.Fields("cp09").Value & "'"
'      If adohirecase.Fields("cp32").Value = MsgText(603) Then
'         strComputer = MsgText(602)
'      Else
'         strComputer = MsgText(601)
'      End If
'      Acc0j0SaveLA
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select cu10 from hirecase, customer where substr(hc05, 1, 8) = cu01 and cu02 = '0' and hc01 = '" & adohirecase.Fields("hc01").Value & "' and hc02 = '" & adohirecase.Fields("hc02").Value & "' and hc03 = '" & adohirecase.Fields("hc03").Value & "' and hc04 = '" & adohirecase.Fields("hc04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount <> 0 Then
'         stra0j21 = "台灣"
'         If IsNull(adocheck.Fields(0).Value) = False Then
'            If Val(adocheck.Fields(0).Value) > 10 Then
'               adoacc0j0.Close
'               adoacc0j0.CursorLocation = adUseClient
'               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
'               If adoacc0j0.RecordCount = 0 Then
'                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j03, a0j04, a0j05, a0j09, a0j10, a0j11, a0j12, a0j14, a0j15, a0j16, a0j20, a0j21, a0j08) " & _
'                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & lnga0j03 & "', '" & stra0j04 & "', '" & stra0j05 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & Val(stra0j12) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & stra0j20 & "', '" & stra0j21 & "', '" & strComputer & "')"
'               End If
'            End If
'         End If
'      End If
'      adocheck.Close
'      adohirecase.MoveNext
'   Loop
'   adohirecase.Close
'   adoservice.CursorLocation = adUseClient
''   adoservice.Open "select * from caseprogress, servicepractice, casepropertymap where (caseprogress.cp01 = sp01  and caseprogress.cp02 = sp02 and caseprogress.cp03 = sp03 and caseprogress.cp04 = sp04) and cp01 = cpm01 and cp10 = cpm02 " & _
''                   "and cp57 is null and cp05 >= 20011112 and (cp16 <> 0 and cp16 is not null) and sp08 like '" & Mid(Text1, 1, 8) & "%" & "' and (cp60 is null or cp60 = '')", adoTaie, adOpenStatic, adLockReadOnly
'   adoservice.Open "select * from caseprogress, servicepractice, casepropertymap where (caseprogress.cp01 = sp01  and caseprogress.cp02 = sp02 and caseprogress.cp03 = sp03 and caseprogress.cp04 = sp04) and cp01 = cpm01 and cp10 = cpm02 " & _
'                   "and cp57 is null and cp05 >= 20030101 and (cp16 <> 0 and cp16 is not null) and sp08 = '" & Text1 & "' and (cp60 is null or cp60 = '')", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoservice.EOF = False
'      adoTaie.Execute "delete from acc0j0 where (a0j13 is null or a0j13 = '') and a0j01 = '" & adoservice.Fields("cp09").Value & "'"
'      If adoservice.Fields("cp32").Value = MsgText(603) Then
'         strComputer = MsgText(602)
'      Else
'         strComputer = MsgText(601)
'      End If
'      Acc0j0SaveS
'      adocheck.CursorLocation = adUseClient
'        'Modify By Cheng 2004/05/05
'        '加國名(NA03)
''      adocheck.Open "select cu10 from servicepractice, customer where substr(sp08, 1, 8) = cu01 and cu02 = '0' and sp01 = '" & adoservice.Fields("sp01").Value & "' and sp02 = '" & adoservice.Fields("sp02").Value & "' and sp03 = '" & adoservice.Fields("sp03").Value & "' and sp04 = '" & adoservice.Fields("sp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      adocheck.Open "select cu10, NA03 from servicepractice, customer, Nation where substr(sp08, 1, 8) = cu01 and cu02 = '0' And SP09=NA01 and sp01 = '" & adoservice.Fields("sp01").Value & "' and sp02 = '" & adoservice.Fields("sp02").Value & "' and sp03 = '" & adoservice.Fields("sp03").Value & "' and sp04 = '" & adoservice.Fields("sp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'        'End
'      If adocheck.RecordCount <> 0 Then
''         stra0j21 = ""
'         stra0j21 = "" & adocheck.Fields(1).Value
'         If IsNull(adocheck.Fields(0).Value) = False Then
'            If Val(adocheck.Fields(0).Value) > 10 Then
'               adoacc0j0.Close
'               adoacc0j0.CursorLocation = adUseClient
'               adoacc0j0.Open "select * from acc0j0 where a0j01 = '" & stra0j01 & "'", adoTaie, adOpenStatic, adLockReadOnly
'               If adoacc0j0.RecordCount = 0 Then
'                  '2005/5/23 MODIFY BY SONIA 加 ACC0J0
'                  'adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j03, a0j04, a0j05, a0j09, a0j10, a0j11, a0j12, a0j14, a0j15, a0j16, a0j20, a0j21, a0j08) " & _
'                  '                "values ('" & stra0j01 & "', '" & stra0j02 & "', " & lnga0j03 & ", '" & stra0j04 & "', '" & stra0j05 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & Val(stra0j12) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & stra0j20 & "', '" & stra0j21 & "', '" & strComputer & "')"
'                  adoTaie.Execute "insert into acc0j0 (a0j01, a0j02, a0j03, a0j04, a0j05, a0j09, a0j10, a0j11, a0j12, a0j14, a0j15, a0j16, a0j20, a0j21, a0j08, a0j07) " & _
'                                  "values ('" & stra0j01 & "', '" & stra0j02 & "', '" & lnga0j03 & "', '" & stra0j04 & "', '" & stra0j05 & "', " & lnga0j09 & ", " & lnga0j10 & ", '" & stra0j11 & "', " & Val(stra0j12) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & stra0j20 & "', '" & stra0j21 & "', '" & strComputer & "', '" & stra0j07 & "')"
'                  '2005/5/23 END
'               End If
'            End If
'         End If
'      End If
'      adocheck.Close
'      adoservice.MoveNext
'   Loop
'   adoservice.Close
'end 2011/3/24

   AdodcRefresh
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
End Sub
