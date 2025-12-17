VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1130 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收據／請款單作廢作業"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8760
   Begin VB.CommandButton Command1 
      Caption         =   "拆收據其他收據號資料"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      TabIndex        =   27
      Top             =   4680
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2580
      Picture         =   "Frmacc1130.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   210
      Width           =   350
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      TabIndex        =   25
      Top             =   690
      Width           =   855
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   23
      Top             =   690
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1130.frx":0102
      Height          =   2385
      Left            =   120
      TabIndex        =   21
      Top             =   2190
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4207
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1154.835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   2040
      Visible         =   0   'False
      Width           =   972
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
   Begin VB.TextBox Text8 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      TabIndex        =   20
      Top             =   1770
      Width           =   1572
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   18
      Top             =   1770
      Width           =   1572
   End
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
      Height          =   330
      Left            =   6000
      TabIndex        =   16
      Top             =   1410
      Width           =   372
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4320
      TabIndex        =   14
      Top             =   1770
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   7
      Top             =   1050
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   4080
      TabIndex        =   2
      Top             =   210
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   1560
      TabIndex        =   11
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
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
   Begin MSForms.TextBox Text13 
      Height          =   330
      Left            =   6720
      TabIndex        =   28
      Top             =   690
      Width           =   1695
      VariousPropertyBits=   671105051
      Size            =   "2990;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text12 
      Height          =   330
      Left            =   6360
      TabIndex        =   3
      Top             =   210
      Width           =   2055
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "3625;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   330
      Left            =   1560
      TabIndex        =   15
      Top             =   1410
      Width           =   4335
      VariousPropertyBits=   671105051
      Size            =   "7646;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   5880
      TabIndex        =   8
      Top             =   1050
      Width           =   2535
      VariousPropertyBits=   671105051
      Size            =   "4471;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   5880
      TabIndex        =   26
      Top             =   240
      Width           =   612
   End
   Begin VB.Label Label12 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
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
      Left            =   4850
      TabIndex        =   24
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   360
      TabIndex        =   22
      Top             =   720
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   492
      Left            =   240
      Top             =   120
      Width           =   8292
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "總計"
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
      Left            =   6120
      TabIndex        =   19
      Top             =   1800
      Width           =   612
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "服務費合計"
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
      Left            =   360
      TabIndex        =   17
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "規費合計"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1.不可扣繳 2.可扣繳"
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
      Left            =   6480
      TabIndex        =   12
      Top             =   1480
      Width           =   1755
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
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
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label3 
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
      Height          =   252
      Left            =   3360
      TabIndex        =   6
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "作廢日期"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
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
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc1130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoacc0k0a As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public strRelateNoList As String '相關收據號碼清單 Add by Morgan 2011/9/21


'Add by Morgan 2011/9/21
Private Sub Command1_Click()
   Call ShowRelNo
End Sub

'Add by Morgan 2011/9/21
Public Function ShowRelNo(Optional pAll As Boolean, Optional pIsCancelConfirm As Boolean) As Boolean
   Dim strCon As String
   Dim stSQL As String, intR As Integer
   Dim ado0k0 As ADODB.Recordset
   
   If pAll = False Then
      strCon = " and a0k01<>'" & Text1 & "'"
   End If
   
   stSQL = "select '',a0k01,sqldatet(a0k02),a0k04,a0k06||'',a0k07||'' from acc0k0 where a0k01 in ('" & Replace(strRelateNoList, ",", "','") & "')" & strCon & " order by 2"
   intR = 1
   Set ado0k0 = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      '作廢確認
      If pIsCancelConfirm = True Then
         Frmacc1132.lblAlert.Visible = True
         Set Frmacc1132.grdDataList.Recordset = ado0k0.Clone
         Set Frmacc1132.fmParent = Me
         Frmacc1132.Show vbModal
         strFormName = Me.Name
         If Me.Tag = "Y" Then
            ShowRelNo = True
         End If
      '查詢
      Else
         If ado0k0.RecordCount = 1 Then
            Frmacc1133.strNo = ado0k0.Fields("a0k01")
            Frmacc1133.Show vbModal
            strFormName = Me.Name
         Else
            Do
               Set Frmacc1132.grdDataList.Recordset = ado0k0.Clone
               Set Frmacc1132.fmParent = Me
               Frmacc1132.Show vbModal
               If Me.Tag <> "" Then
                  Frmacc1133.strNo = Me.Tag
                  Frmacc1133.Show vbModal
               End If
               strFormName = Me.Name
            Loop While Me.Tag <> ""
         End If
      End If
   End If
   Set ado0k0 = Nothing
End Function

Private Sub Command2_Click()
   If strSaveConfirm <> "" Then Exit Sub 'Added by Morgan 2017/11/30 維護狀態下不可查詢
   
   'If adoacc0k0.RecordCount = 0 Or Text1 = MsgText(601) Then
   '   Exit Sub
   'End If
   'adoacc0k0.Find "a0k01 = '" & Text1 & "'", 0, adSearchForward, 1
   'If adoacc0k0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   RecordShow
   'Else
   '   MsgBox MsgText(33), , MsgText(5)
   '   adoacc0k0.MoveFirst
   'End If
   Acc0k0Refresh
   If adoacc0k0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      RecordShow
   End If
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command2_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   'If adoacc0k0.RecordCount <> 0 Then
   '   adoacc0k0.MoveFirst
   '   AdodcRefresh
   'End If
   'adoacc0k0.Find "a0k01 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adoacc0k0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   RecordShow
   'End If
   Text1 = strItemNo
   Acc0k0Refresh
   If adoacc0k0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5600 'Modify by Amy 2023/10/05 原:5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strItemNo = MsgText(601)
   MaskEdBox1.Mask = DFormat
   OpenTable
   If adoacc0k0.RecordCount <> 0 Then
      adoacc0k0.MoveLast
      adoacc0k0.MoveFirst
      RecordShow
   End If
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
   Set Frmacc1130 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_Change()
   Command1.Enabled = False
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.MaxRecords = intMax
   adoacc0k0.Open "select * from acc0k0 where a0k09 <> 0 and a0k01 >= '" & Text1 & "' order by a0k01", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0k0a.CursorLocation = adUseClient
   adoacc0k0a.MaxRecords = intMax
   adoacc0k0a.Open "select * from acc0k0 where a0k09 = 0 and a0k01 >= '" & Text1 & "' order by a0k01", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   adoadodc1.Open "select a.*,getcp10desc(cp01,cp10,a0j04) cp10N,na03 from acc0j0 a,caseprogress,nation where a0j13 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 order by a0j01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國內收據作廢資料)
'
'*************************************************
Public Sub FormShow()
   Text1 = adoacc0k0.Fields("a0k01").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0k0.Fields("a0k09").Value) Or adoacc0k0.Fields("a0k09").Value = 0 Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0k0.Fields("a0k09").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0k0.Fields("a0k08").Value) Then
      Text12 = MsgText(601)
   Else
      Text12 = adoacc0k0.Fields("a0k08").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k11").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc0k0.Fields("a0k11").Value
   End If
   
   'Removed by Morgan 2011/11/24 一張收據可同時包含合併及不合併的收文資料
   'If IsNull(adoacc0k0.Fields("a0k30").Value) Then
   '   Text10 = MsgText(601)
   'Else
   '   Text10 = adoacc0k0.Fields("a0k30").Value
   'End If
   
   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc0k0.Fields("a0k20").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0k0.Fields("a0k02").Value) Or adoacc0k0.Fields("a0k02").Value = 0 Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0k0.Fields("a0k02").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0k0.Fields("a0k03").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k04").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc0k0.Fields("a0k04").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k05").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc0k0.Fields("a0k05").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k07").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0k0.Fields("a0k07").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k06").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = adoacc0k0.Fields("a0k06").Value
   End If
   Text8 = Val(Text4) + Val(Text7)
End Sub

'*************************************************
'  顯示查詢資料(國內收據資料)
'
'*************************************************
Private Sub Acc0k0Query()
'   MaskEdBox1.Mask = MsgText(601)
'   If IsNull(adoacc0k0a.Fields("a0k09").Value) Then
'      MaskEdBox1.Text = MsgText(601)
'   Else
'      MaskEdBox1.Text = CFDate(adoacc0k0a.Fields("a0k09").Value)
'   End If
'   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0k0a.Fields("a0k08").Value) Then
      Text12 = MsgText(601)
   Else
      Text12 = adoacc0k0a.Fields("a0k08").Value
   End If
   If IsNull(adoacc0k0a.Fields("a0k11").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc0k0a.Fields("a0k11").Value
   End If
   
   'Removed by Morgan 2011/11/24 一張收據可同時包含合併及不合併的收文資料
   'If IsNull(adoacc0k0a.Fields("a0k30").Value) Then
   '   Text10 = MsgText(601)
   'Else
   '   Text10 = adoacc0k0a.Fields("a0k30").Value
   'End If
   
   If IsNull(adoacc0k0a.Fields("a0k20").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc0k0a.Fields("a0k20").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0k0a.Fields("a0k02").Value) Or adoacc0k0a.Fields("a0k02").Value = 0 Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0k0a.Fields("a0k02").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0k0a.Fields("a0k03").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0k0a.Fields("a0k03").Value
   End If
   If IsNull(adoacc0k0a.Fields("a0k04").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc0k0a.Fields("a0k04").Value
   End If
   If IsNull(adoacc0k0a.Fields("a0k05").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc0k0a.Fields("a0k05").Value
   End If
   If IsNull(adoacc0k0a.Fields("a0k07").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0k0a.Fields("a0k07").Value
   End If
   If IsNull(adoacc0k0a.Fields("a0k06").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = adoacc0k0a.Fields("a0k06").Value
   End If
   Text8 = Val(Text4) + Val(Text7)
End Sub

'*************************************************
'  重新整理 Adcdc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   adoadodc1.Open "select a.*,getcp10desc(cp01,cp10,a0j04) cp10N,na03 from acc0j0 a,caseprogress,nation where a0j13 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 order by a0j02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  清除查詢資料
'
'*************************************************
Private Sub AdodcClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Text12 = ""
   Text9 = ""
   'Text10 = "" 'Removed by Morgan 2011/11/24 一張收據可同時包含合併及不合併的收文資料
   Text11 = ""
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text2 = ""
   Text3 = ""
   Text6 = ""
   Text5 = ""
   Text4 = ""
   Text7 = ""
   Text8 = ""
   Text13 = ""  '2012/6/11 add by sonia
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Then
      adoacc0k0a.Close
      adoacc0k0a.CursorLocation = adUseClient
      adoacc0k0a.Open "select * from acc0k0 where a0k01 = '" & Text1 & "' order by a0k01", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc0k0a.RecordCount <> 0 Then
         Acc0k0Query
         AdodcRefresh
         'Add by Morgan 2011/9/21
         strRelateNoList = PUB_GetRelateNo(Text1)
         If strRelateNoList <> Text1 Then
            Command1.Enabled = True
         End If
      Else
         MsgBox MsgText(28), , MsgText(5)
         AdodcClear
         AdodcRefresh
         Cancel = True
      End If
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Text12_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text2_Change()
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Text3 = CustomerQuery(Text2, 1)
End Sub

Private Sub Text11_Change()
   If Text11 = MsgText(601) Then
      Exit Sub
   End If
   Text13 = GetPrjSalesNM(Text11)
End Sub

'*************************************************
'  重新整理國內收據資料
'
'*************************************************
Public Sub Acc0k0Refresh()
On Error GoTo Checking
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.MaxRecords = intMax
   adoacc0k0.Open "select * from acc0k0 where a0k09 <> 0 and a0k01 >= '" & Text1 & "' order by a0k01", adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   'Removed by Morgan 2017/11/30 沒用
   'If adoacc0k0a.State = adStateOpen Then
   '   adoacc0k0a.Close
   'End If
   'adoacc0k0a.CursorLocation = adUseClient
   'adoacc0k0a.MaxRecords = intMax
   'adoacc0k0a.Open "select * from acc0k0 where a0k09 = 0 and a0k01 >= '" & Text1 & "' order by a0k01", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2017/11/30
   
   If adoacc0k0.RecordCount <> 0 Then
      If Text1 <> MsgText(601) Then
         adoacc0k0.Find "a0k01 = '" & Text1 & "'", 0, adSearchForward, 1
         If adoacc0k0.EOF = False Then
            FormShow
            AdodcRefresh
            RecordShow
         Else
            adoacc0k0.MoveFirst
         End If
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0k0.Bookmark & MsgText(35) & adoacc0k0.RecordCount
End Sub

