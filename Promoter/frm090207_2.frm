VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm090207_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案例資料查詢"
   ClientHeight    =   5715
   ClientLeft      =   90
   ClientTop       =   1290
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9300
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   624
      Left            =   4176
      Top             =   1296
      Visible         =   0   'False
      Width           =   1236
      _ExtentX        =   2170
      _ExtentY        =   1111
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
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6864
      TabIndex        =   0
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "詳細資料(&F)"
      Height          =   400
      Left            =   8088
      TabIndex        =   1
      Top             =   10
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm090207_2.frx":0000
      Height          =   5208
      Left            =   0
      TabIndex        =   2
      Top             =   468
      Width           =   9276
      _ExtentX        =   16351
      _ExtentY        =   9181
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "N1"
         Caption         =   "類別"
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
         DataField       =   "N3"
         Caption         =   "文書日期"
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
         DataField       =   "N4"
         Caption         =   "文書種類"
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
         DataField       =   "PC09"
         Caption         =   "主旨"
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
         DataField       =   "PC11"
         Caption         =   "案情摘要"
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
         DataField       =   "N2"
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
      BeginProperty Column06 
         DataField       =   "PC10"
         Caption         =   "案例字號"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm090207_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset
Public a As String

Private Sub Command1_Click()
'A = pemain.Fields(0).Value
frm090207_3.Show
frm090207_3.lbl1(0).Caption = SystemNumber(CheckStr(Adodc1.Recordset.Fields(0)), 1)
frm090207_3.lbl1(1).Caption = SystemNumber(CheckStr(Adodc1.Recordset.Fields(0)), 2)
frm090207_3.lbl1(2).Caption = SystemNumber(CheckStr(Adodc1.Recordset.Fields(0)), 3)
frm090207_3.lbl1(3).Caption = SystemNumber(CheckStr(Adodc1.Recordset.Fields(0)), 4)
frm090207_3.lbl1(4).Caption = CheckStr(Adodc1.Recordset.Fields(3))
frm090207_3.lbl1(5).Caption = CheckStr(Adodc1.Recordset.Fields(1))
frm090207_3.lbl1(6).Caption = CheckStr(Adodc1.Recordset.Fields(4))
frm090207_3.lbl1(7).Caption = CheckStr(Adodc1.Recordset.Fields(2))
frm090207_3.lbl1(8).Caption = CheckStr(Adodc1.Recordset.Fields(6))
frm090207_3.lbl1(9).Caption = ChangeTStringToTDateString("" & Adodc1.Recordset.Fields(5))
frm090207_2.Hide
End Sub

Private Sub Command2_Click()
frm090207_1.Show
Unload Me
End Sub

Private Sub Form_Activate()
'If pemain.State = adStateOpen Then pemain.Close
'If frm090207_1.ADDR = 1 Then
'pemain.Open frm090207_1.SQLSTRING1, cnnConnection, adOpenStatic, adLockReadOnly
'ElseIf frm090207_1.ADDR = 2 Then
'pemain.Open frm090207_1.SQLSTRING, cnnConnection, adOpenStatic, adLockReadOnly
'End If
'If pemain.BOF And pemain.EOF Then MsgBox "資料庫內無資料,請回到上一頁重新輸入", vbInformation: frm090207_1.Show: Unload Me: Exit Sub
'Set Adodc1.Recordset = pemain
'Adodc1.Recordset.Requery

End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'If pemain.State = adStateOpen Then pemain.Close
'pemain.CursorLocation = adUseClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090207_2 = Nothing
End Sub
