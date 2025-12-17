VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11j0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "員工翻譯費率維護"
   ClientHeight    =   5124
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5124
   ScaleWidth      =   8760
   Begin VB.TextBox txtSPR 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   13
      Left            =   3780
      MaxLength       =   4
      TabIndex        =   4
      Top             =   629
      Width           =   900
   End
   Begin VB.TextBox txtSPR 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   14
      Left            =   6000
      MaxLength       =   4
      TabIndex        =   5
      Top             =   629
      Width           =   900
   End
   Begin VB.TextBox txtSPR 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   16
      Left            =   6000
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1008
      Width           =   900
   End
   Begin VB.TextBox txtSPR 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   15
      Left            =   3768
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1008
      Width           =   900
   End
   Begin VB.TextBox txtSPR 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   12
      Left            =   1572
      MaxLength       =   4
      TabIndex        =   3
      Top             =   629
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   2640
      Picture         =   "Frmacc11j0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11j0.frx":0102
      Height          =   2988
      Left            =   96
      TabIndex        =   8
      Top             =   1812
      Width           =   8484
      _ExtentX        =   14965
      _ExtentY        =   5271
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "SPR01"
         Caption         =   "員工號"
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
         DataField       =   "ST02"
         Caption         =   "姓名"
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
         DataField       =   "SPR12"
         Caption         =   "英翻費率"
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
      BeginProperty Column03 
         DataField       =   "SPR13"
         Caption         =   "日翻費率"
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
      BeginProperty Column04 
         DataField       =   "SPR14"
         Caption         =   "德翻費率"
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
         DataField       =   "SPR15"
         Caption         =   "中翻日費率"
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
         DataField       =   "SPR16"
         Caption         =   "中翻德費率"
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
         DataField       =   "ST03"
         Caption         =   "部門"
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
         DataField       =   "ST04"
         Caption         =   "狀態"
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
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   864
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   756.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   720
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   3210
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
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   1548
      MaxLength       =   6
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(中文字數計算)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   7104
      TabIndex        =   19
      Top             =   1075
      Width           =   1296
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "中翻日費率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   2712
      TabIndex        =   18
      Top             =   1069
      Width           =   1020
   End
   Begin MSForms.TextBox Text2 
      Height          =   315
      Left            =   4740
      TabIndex        =   2
      Top             =   233
      Width           =   3555
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "日文翻譯費率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   2520
      TabIndex        =   17
      Top             =   690
      Width           =   1260
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS: 108.8.15 以後完稿案件改以原文字數計算翻譯費並取消中文打字費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   4860
      Width           =   6420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(原文字數計算)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   7104
      TabIndex        =   15
      Top             =   696
      Width           =   1296
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "員工代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   285
      TabIndex        =   14
      Top             =   300
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "英文翻譯費率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   285
      TabIndex        =   13
      Top             =   689
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "中翻德費率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   4944
      TabIndex        =   12
      Top             =   1069
      Width           =   1020
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "德文翻譯費率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   4752
      TabIndex        =   11
      Top             =   690
      Width           =   1224
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1368
      Left            =   72
      Top             =   72
      Width           =   8508
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "費率計算單位 : NT$/千字"
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
      Left            =   6000
      TabIndex        =   10
      Top             =   1524
      Width           =   2508
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "員工姓名"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3405
      TabIndex        =   9
      Top             =   300
      Width           =   840
   End
End
Attribute VB_Name = "Frmacc11j0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改  Text2/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Modified by Morgan 2017/2/10 +英文翻譯費率2

'Modified by Morgan 2019/8/14
'108.8.15以後完稿案件翻譯費改以原文字數計算並取消中文打字費
'考慮適用舊費率的未付款案件,保留舊費率及欄位,另新增欄位紀錄新費率
'end 2019/8/14

'Add by Morgan 2007/5/16
Option Explicit
Public adoadodc1 As New ADODB.Recordset

Public Sub MoveFirst()
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MoveFirst
      FormShow
      RecordShow
   End If
End Sub

Public Sub MoveLast()
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MoveLast
      FormShow
      RecordShow
   End If
End Sub

Public Sub MoveNext()
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MoveNext
      If Not adoadodc1.EOF Then
         FormShow
         RecordShow
      Else
         adoadodc1.MoveLast
      End If
   End If
End Sub

Public Sub MovePrevious()
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MovePrevious
      If Not adoadodc1.BOF Then
         FormShow
         RecordShow
      Else
         adoadodc1.MoveFirst
      End If
   End If
End Sub

Public Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "SPR01 = '" & Text1 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(3) Then
      FormShow
   End If
   RecordShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "SPR01 = '" & strItemNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
   '表單初始化
   'Modified byMorgan 2019/7/29
   'PUB_InitForm Me, 8850, 5500
   PUB_InitForm Me, Me.Width, Me.Height
   'end 2019/7/29
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
      FormShow
      RecordShow
   End If
   FormEnable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11j0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2019/8/14 +,SPR12,SPR13
   'Modified by Morgan 2025/10/21
   'adoadodc1.Open "select SPR01,ST02,SPR02,SPR03,SPR04,SPR11,ST03,DECODE(ST04,'1','在職','離職') ST04,SPR12,SPR13 from Staff_PayRate,STAFF WHERE ST01(+)=SPR01 order by ST03 DESC,ST01", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select SPR01,ST02,ST03,DECODE(ST04,'1','在職','離職') ST04,SPR12,SPR13,SPR14,SPR15,SPR16 from Staff_PayRate,STAFF WHERE ST01(+)=SPR01 order by ST03 DESC,ST01", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub AdodcRefresh()
   Adodc1.Recordset.Requery
End Sub
'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   If Not (Adodc1.Recordset.EOF And Adodc1.Recordset.BOF) Then
      Text1 = "" & Adodc1.Recordset.Fields("SPR01").Value
      Text2 = "" & Adodc1.Recordset.Fields("ST02").Value
      'Removed by Morgan 2025/10/21
      'Text3 = Format("" & Adodc1.Recordset.Fields("SPR02").Value)
      'Text4 = Format("" & Adodc1.Recordset.Fields("SPR03").Value)
      'Text5 = Format("" & Adodc1.Recordset.Fields("SPR04").Value)
      'Text6 = Format("" & Adodc1.Recordset.Fields("SPR11").Value)
      'end 2025/10/21
      'Added by Morgan 2019/8/14
      txtSPR(12) = "" & Adodc1.Recordset.Fields("SPR12").Value
      txtSPR(13) = "" & Adodc1.Recordset.Fields("SPR13").Value
      'end 2019/8/14
      'Added by Morgan 2025/10/21
      txtSPR(14) = "" & Adodc1.Recordset.Fields("SPR14").Value
      txtSPR(15) = "" & Adodc1.Recordset.Fields("SPR15").Value
      txtSPR(16) = "" & Adodc1.Recordset.Fields("SPR16").Value
      'end 2019/8/14
      Text1.Tag = Text1.Text
   Else
      Text1.Tag = ""
   End If
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

Public Sub FormEnable()
   Dim bolLock As Boolean
   '新增
   If strSaveConfirm = MsgText(3) Then
      Text1.Locked = False
      Command1.Enabled = False
      bolLock = False
      DataGrid1.Enabled = False
   '修改
   ElseIf strSaveConfirm = MsgText(4) Then
      Text1.Locked = True
      Command1.Enabled = False
      bolLock = False
      DataGrid1.Enabled = False
   Else
      Text1.Locked = False
      Command1.Enabled = True
      bolLock = True
      DataGrid1.Enabled = True
   End If
   
   'Modified by Morgan 2025/10/21
   'Text3.Locked = bolLock
   'Text4.Locked = bolLock
   'Text5.Locked = bolLock
   'Text6.Locked = bolLock
   txtSPR(12).Locked = bolLock
   txtSPR(13).Locked = bolLock
   txtSPR(14).Locked = bolLock
   txtSPR(15).Locked = bolLock
   txtSPR(16).Locked = bolLock
   'end 2025/10/21
End Sub

Public Sub FormClear()
   Dim oObj As Object
   For Each oObj In Me.Controls
      If TypeName(oObj) = "TextBox" Then
         oObj.Text = Empty
      End If
   Next
   '新增時中打費預設120
   'Removed by Morgan 2025/10/21
   'If strSaveConfirm = MsgText(3) Then
   '   Text5.Text = 120
   'End If
   'end 2025/10/21
End Sub

Private Sub Text1_GotFocus()
   If Text1 = "F5" Then
      Text1.SelStart = 2
   Else
      TextInverse Text1
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   Text2.Text = GetStaffName(Text1, True)
End Sub

'Removed by Morgan 2025/10/21
'Private Sub Text3_GotFocus()
'   TextInverse Text3
'End Sub
'
'Private Sub Text4_GotFocus()
'   TextInverse Text4
'End Sub
'
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'End Sub
'
'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub
'end 2025/10/21

Private Function SaveCheck() As Boolean
   If Left(Text1, 2) <> "F5" Then
      MsgBox "請輸入F5開頭的外譯編號！"
      If Text1.Enabled Then Text1.SetFocus
      Exit Function
   End If
   
   If Text2.Text = "" Then
      MsgBox "員工代碼不存在，請重新輸入！"
      If Text1.Enabled Then Text1.SetFocus
      Exit Function
   End If
   
'Removed by Morgan 2025/10/21
'   If Text3 <> "" And Not IsNumeric(Text3) Then
'      MsgBox "英文翻譯費率輸入錯誤！"
'      Text3.SetFocus
'      Exit Function
'   End If
'
'   If Text4 <> "" And Not IsNumeric(Text4) Then
'      MsgBox "日文翻譯費率輸入錯誤！"
'      Text4.SetFocus
'      Exit Function
'   End If
'
'   If Text5 <> "" And Not IsNumeric(Text5) Then
'      MsgBox "中文打字費率輸入錯誤！"
'      Text5.SetFocus
'      Exit Function
'   End If

   SaveCheck = True
End Function

Public Function FormSave() As Boolean
   
   If SaveCheck = False Then
      Exit Function
   End If
   
On Error GoTo ErrHnd

   'Modified by Morgan 2019/8/14 +,SPR12,SPR13
   If strSaveConfirm = MsgText(3) Then
      'Modified by Morgan 2025/10/21
      'strSql = "INSERT INTO STAFF_PAYRATE(SPR01,SPR02,SPR03,SPR04,SPR11,SPR12,SPR13)" & _
         " VALUES('" & Text1.Text & "'," & CNULL(Format(Text3.Text)) & "," & CNULL(Format(Text4.Text)) & "," & CNULL(Format(Text5)) & "," & CNULL(Format(Text6)) & "," & CNULL(Format(txtSPR(12))) & "," & CNULL(Format(txtSPR(13))) & ")"
      strSql = "INSERT INTO STAFF_PAYRATE(SPR01,SPR12,SPR13,SPR14,SPR15,SPR16)" & _
         " VALUES('" & Text1.Text & "'," & CNULL(Format(txtSPR(12))) & "," & CNULL(Format(txtSPR(13))) & "," & CNULL(Format(txtSPR(14))) & "," & CNULL(Format(txtSPR(15))) & "," & CNULL(Format(txtSPR(16))) & ")"
   Else
      'Modified by Morgan 2025/10/21
      'strSql = "UPDATE STAFF_PAYRATE SET SPR02=" & CNULL(Format(Text3.Text)) & ",SPR03=" & CNULL(Format(Text4.Text)) & ",SPR04=" & CNULL(Format(Text5)) & ",SPR11=" & CNULL(Format(Text6)) & ",SPR12=" & CNULL(Format(txtSPR(12))) & ",SPR13=" & CNULL(Format(txtSPR(13))) & _
         " WHERE SPR01='" & Text1.Text & "'"
      strSql = "UPDATE STAFF_PAYRATE SET SPR12=" & CNULL(Format(txtSPR(12))) & ",SPR13=" & CNULL(Format(txtSPR(13))) & ",SPR14=" & CNULL(Format(txtSPR(14))) & ",SPR15=" & CNULL(Format(txtSPR(15))) & ",SPR16=" & CNULL(Format(txtSPR(16))) & _
         " WHERE SPR01='" & Text1.Text & "'"
   End If
   
   adoTaie.Execute strSql, intI
   FormSave = True
   AdodcRefresh
   Command1_Click
   Exit Function
   
ErrHnd:
   If Err.Number = -2147217873 Then
      MsgBox "資料已存在，請改為修改模式作業！"
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function
Public Function FormCheck() As Boolean
   If Text1.Text = "" Then
      MsgBox "員工代碼不可空白！"
      Exit Function
   ElseIf Text1.Tag <> Text1.Text Then
      MsgBox "員工代碼有改，請重新查詢！"
      Exit Function
   End If
   FormCheck = True
End Function
Public Function FormDelete() As Boolean
   If FormCheck = False Then
      Exit Function
   End If
On Error GoTo ErrHnd
   strSql = "delete from Staff_PayRate where SPR01='" & Text1.Text & "'"
   adoTaie.Execute strSql
   FormDelete = True
   FormClear
   AdodcRefresh
   FormShow
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub txtSPR_GotFocus(Index As Integer)
   TextInverse txtSPR(Index)
End Sub
