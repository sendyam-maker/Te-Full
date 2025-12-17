VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2142 
   BorderStyle     =   1  '單線固定
   Caption         =   "各幣別最新請款匯率查詢"
   ClientHeight    =   5715
   ClientLeft      =   2565
   ClientTop       =   3105
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6435
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   5520
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   60
      Width           =   750
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmacc2142.frx":0000
      Height          =   5100
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   8996
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "dnr01"
         Caption         =   "幣別"
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
         DataField       =   "dnr02"
         Caption         =   "日期"
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
      BeginProperty Column02 
         DataField       =   "dnr03"
         Caption         =   "對台幣匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.0#####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "dnr04"
         Caption         =   "對美金匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.0#####"
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
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1409.953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
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
End
Attribute VB_Name = "Frmacc2142"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/03 Form2.0已檢查 (無需修改的物件)
'Add by Amy 20130912
Option Explicit

Dim rs As New ADODB.Recordset

Private Sub cmdOK_Click(Index As Integer)
    Unload Frmacc2142
End Sub

'OpenTable 放在Form_Load 會導致第一筆的第一欄變成空值
Private Sub Form_Activate()
    OpenTable
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()

On Error GoTo Checking
    'Modified by Morgan 2022/11/3 修正日期條件為轉西元問題 +19110000
    strExc(0) = "SELECT DNR01 ,DNR02,DNR03,DECODE(DNR04,0,NULL,DNR04) DNR04 " & _
                    "FROM (SELECT 'USD' AS DNR01,USXR01 AS DNR02,USXR02 AS DNR03,0 AS DNR04 FROM USXRATE WHERE USXR01 IN (SELECT MAX(USXR01) FROM USXRATE WHERE USXR01+19110000<=TO_CHAR(sysdate,'YYYYMMDD')) " & _
                    "UNION SELECT DNR01,DNR02,DNR03,DNR04 FROM DEBITNOTERATE WHERE (DNR01,DNR02) IN (SELECT DNR01,MAX(DNR02) FROM DEBITNOTERATE WHERE DNR01<>'NTD' AND DNR02+19110000<=TO_CHAR(sysdate,'YYYYMMDD') GROUP BY DNR01) " & _
                    ") ORDER BY DNR01 Desc"

   If rs.State <> adStateClosed Then rs.Close
   rs.CursorLocation = adUseClient
   rs.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   If rs.RecordCount <> 0 Then
       Set Adodc1.Recordset = rs
   ElseIf rs.RecordCount = 0 Then
        MsgBox MsgText(28), , MsgText(5)
   End If
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   rs.Close
   Set rs = Nothing
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    Set Frmacc2142 = Nothing
End Sub
