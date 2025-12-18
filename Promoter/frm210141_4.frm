VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm210141_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳款輸入-選擇應收客戶"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9390
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdOK 
      Caption         =   "選擇客戶代號(&H)"
      Height          =   400
      Index           =   2
      Left            =   6030
      TabIndex        =   3
      Top             =   120
      Width           =   1590
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "取消(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7680
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "選擇收據抬頭(&H)"
      Height          =   400
      Index           =   1
      Left            =   4410
      TabIndex        =   1
      Top             =   120
      Width           =   1590
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm210141_4.frx":0000
      Height          =   4875
      Left            =   225
      TabIndex        =   0
      Top             =   600
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   8599
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "A0K04"
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
      BeginProperty Column01 
         DataField       =   "A0K03"
         Caption         =   "客戶編號"
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
         DataField       =   "CU04"
         Caption         =   "客戶名稱"
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
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   4215.118
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8190
      Top             =   5310
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
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "說明：選擇客戶代號將包含所有關係企業收據"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   4
      Top             =   360
      Width           =   3600
   End
End
Attribute VB_Name = "frm210141_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo by Lydia 2019/07/01 表單名稱:智權人員繳款資料輸入=>繳款輸入
'Created by Morgan 2013/12/18
Option Explicit

Dim m_blnColOrderAsc As Boolean
Public m_Frm146_4 As Boolean 'Add by Lydia 2014/9/25 客戶請款明細
Private Sub cmdOK_Click(Index As Integer)
  If m_Frm146_4 = True Then
   Select Case Index
      Case 0
         frm210146.Tag = "0"
      Case 1
         frm210146.txtTitle = Adodc1.Recordset.Fields("A0K04").Value
         frm210146.txtCustNo(0) = "X"
         frm210146.txtCustNo(1) = "X"
         frm210146.Tag = "1"
      Case 2
         frm210146.txtTitle = ""
         frm210146.txtCustNo(0) = Left(Adodc1.Recordset.Fields("A0K03").Value, 6) & "000"
         frm210146.txtCustNo(1) = Left(Adodc1.Recordset.Fields("A0K03").Value, 6) & "999"
         frm210146.Tag = "1"
   End Select

  Else
   Select Case Index
      Case 0
         frm210141.Tag = "0"
      Case 1
         frm210141.txtTitle = Adodc1.Recordset.Fields("A0K04").Value
         frm210141.txtCustNo(0) = "X"
         frm210141.txtCustNo(1) = "X"
         frm210141.Tag = "1"
      Case 2
         frm210141.txtTitle = ""
         frm210141.txtCustNo(0) = Left(Adodc1.Recordset.Fields("A0K03").Value, 6) & "000"
         frm210141.txtCustNo(1) = Left(Adodc1.Recordset.Fields("A0K03").Value, 6) & "999"
         frm210141.Tag = "1"
   End Select

  End If
   m_Frm146_4 = False
   Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   If m_blnColOrderAsc Then
      Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " desc"
      m_blnColOrderAsc = False
   Else
      Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " asc"
      m_blnColOrderAsc = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
If m_Frm146_4 = True Then
   Me.Caption = "客戶應收帳款明細列印-選擇應收客戶"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frm210141_4 = Nothing
End Sub
