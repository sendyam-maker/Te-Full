VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210141_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳款輸入-帶入簽收金額"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7830
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txtSub 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3915
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2993
      Width           =   915
   End
   Begin VB.TextBox txtSub 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2993
      Width           =   915
   End
   Begin VB.TextBox txtSub 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   585
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2993
      Width           =   915
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6390
      Top             =   1770
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
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2993
      Width           =   1095
   End
   Begin VB.TextBox txtSales 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   2
      Top             =   240
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   6120
      TabIndex        =   1
      Top             =   90
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm210141_1.frx":0000
      Height          =   2385
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   4207
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "選取"
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
         DataField       =   "A2302"
         Caption         =   "繳款日期"
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
         Caption         =   "客戶"
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
         DataField       =   "A2318"
         Caption         =   "金額"
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
         DataField       =   "Type"
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
      BeginProperty Column05 
         DataField       =   "Memo"
         Caption         =   "備註"
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
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1950.236
         EndProperty
      EndProperty
   End
   Begin MSForms.Label lblSalesName 
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   270
      Width           =   1710
      VariousPropertyBits=   27
      Size            =   "3016;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "現金"
      Height          =   180
      Index           =   3
      Left            =   3510
      TabIndex        =   11
      Top             =   3045
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "電匯"
      Height          =   180
      Index           =   2
      Left            =   1845
      TabIndex        =   9
      Top             =   3045
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "票據"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   3045
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "合計"
      Height          =   180
      Index           =   10
      Left            =   5220
      TabIndex        =   5
      Top             =   3045
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   3
      Top             =   285
      Width           =   900
   End
End
Attribute VB_Name = "frm210141_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、lblSalesName
'Memo by Lydia 2019/07/01 表單名稱:智權人員繳款資料輸入=>繳款輸入
'Created by Morgan 2013/12/3
Option Explicit

Dim m_mouseRow As Integer, m_MouseCol As Integer
Dim m_CustNo As String

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub DataGrid1_Click()
   SelectDataGrid1 m_MouseCol, m_mouseRow
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   m_MouseCol = DataGrid1.ColContaining(x)
   m_mouseRow = DataGrid1.RowContaining(y)
End Sub

Private Sub Form_Activate()
   SetTot
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210141_1 = Nothing
End Sub

Private Sub SelectDataGrid1(pCol As Integer, pRow As Integer)
   Dim bUpdate As Boolean, stKey1 As String, stKey2 As String
   
   If pRow >= 0 Then
      Set RsTemp = Adodc1.Recordset.Clone
      With RsTemp
      intI = DataGrid1.FirstRow + pRow - 1
      .Move intI, adBookmarkFirst
      If .Fields("選取") = "Y" Then
         .Fields("選取") = ""
      Else
         If m_CustNo = "" Then
            m_CustNo = Left(.Fields("A2304"), 6)
         ElseIf Left(.Fields("A2304"), 6) <> m_CustNo Then
            MsgBox "多選時只能點選關係企業的電匯資料！", vbExclamation
            Exit Sub
         End If
         .Fields("選取") = "Y"
      End If
      'Modify by Amy 2014/06/16 改暫存TB 解決「找不到要更新的資料列」錯誤
      '.UpdateBatch
      .UPDATE
      
      'Added by Morgan 2015/7/16
      strExc(0) = .Fields("A2301")
      strExc(1) = .Fields("選取")
      .MoveFirst
      Do While Not .EOF
         If .Fields("A2301") = strExc(0) Then
            If "" & .Fields("選取") <> strExc(1) Then
               .Fields("選取") = strExc(1)
               .UPDATE
            End If
         End If
         .MoveNext
      Loop
      'end 2015/7/16
      End With
      SetTot
   End If
End Sub


Private Sub SetTot()
   Dim bYes As Boolean
   Set RsTemp = Adodc1.Recordset.Clone
   With RsTemp
   .MoveFirst
   txtTot = ""
   'Added by Morgan 2015/7/16
   txtSub(0) = ""
   txtSub(1) = ""
   txtSub(2) = ""
   'end 2015/7/16
   bYes = False
   Do While Not .EOF
      If .Fields("選取") = "Y" Then
         If m_CustNo = "" Then m_CustNo = Left(.Fields("A2304"), 6) 'Added by Morgan 2015/7/14
         txtTot = Val(txtTot) + Val("" & .Fields("A2318"))
         bYes = True
         'Added by Morgan 2015/7/16
         If .Fields("Src") = "1" Then
            txtSub(0) = Val(txtSub(0)) + Val("" & .Fields("A2318"))
         ElseIf .Fields("Src") = "2" Then
            txtSub(1) = Val(txtSub(1)) + Val("" & .Fields("A2318"))
         ElseIf .Fields("Src") = "3" Then
            txtSub(2) = Val(txtSub(2)) + Val("" & .Fields("A2318"))
         End If
         'end 2015/7/16
      End If
      .MoveNext
   Loop
   If bYes = False Then m_CustNo = ""
   txtTot = Format(txtTot, "#,##0")
   'Added by Morgan 2015/7/16
   txtSub(0) = Format(txtSub(0), "#,##0")
   txtSub(1) = Format(txtSub(1), "#,##0")
   txtSub(2) = Format(txtSub(2), "#,##0")
   'end 2015/7/16
   End With
End Sub
