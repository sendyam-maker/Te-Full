VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc5200 
   AutoRedraw      =   -1  'True
   Caption         =   "會計科目餘額維護"
   ClientHeight    =   5180
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5180
   ScaleWidth      =   8760
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
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
      Height          =   324
      Left            =   6576
      TabIndex        =   22
      Top             =   4800
      Width           =   1752
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
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
      Height          =   324
      Left            =   4800
      TabIndex        =   21
      Top             =   4800
      Width           =   1752
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '靠右對齊
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
      Height          =   324
      Left            =   2976
      TabIndex        =   20
      Top             =   4800
      Width           =   1800
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  '靠右對齊
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
      Height          =   324
      Left            =   1056
      TabIndex        =   18
      Top             =   4800
      Width           =   1896
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   300
      Left            =   1920
      TabIndex        =   4
      Top             =   864
      Width           =   350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "產生月份資料"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   5
      Top             =   504
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '靠右對齊
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
      Left            =   6840
      TabIndex        =   17
      Top             =   864
      Width           =   1575
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
      Height          =   300
      Left            =   4200
      TabIndex        =   15
      Top             =   864
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1320
      TabIndex        =   3
      Top             =   864
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc5200.frx":0000
      Height          =   3504
      Left            =   240
      TabIndex        =   6
      Top             =   1272
      Width           =   8292
      _ExtentX        =   14623
      _ExtentY        =   6191
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
      Caption         =   "會計科目餘額資料"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a0402"
         Caption         =   "月份"
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
         DataField       =   "a0406"
         Caption         =   "借方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "a0407"
         Caption         =   "貸方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a0409"
         Caption         =   "預算金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "a0408"
         Caption         =   "本月餘額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
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
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1899.78
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1810.205
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1152
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
      Height          =   300
      Left            =   2880
      TabIndex        =   12
      Top             =   504
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1320
      TabIndex        =   2
      Top             =   504
      Width           =   1572
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
      Height          =   300
      Left            =   6000
      TabIndex        =   11
      Top             =   144
      Width           =   2412
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   5400
      TabIndex        =   1
      Top             =   144
      Width           =   612
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1320
      TabIndex        =   0
      Top             =   144
      Width           =   612
   End
   Begin VB.TextBox Text13 
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
      Left            =   1920
      TabIndex        =   9
      Top             =   144
      Width           =   2412
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   564
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "年初餘額"
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
      TabIndex        =   16
      Top             =   864
      Width           =   1092
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "借/貸方"
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
      TabIndex        =   14
      Top             =   864
      Width           =   852
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      TabIndex        =   13
      Top             =   864
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   144
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4704
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1212
      Left            =   240
      Top             =   24
      Width           =   8292
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   8
      Top             =   504
      Width           =   972
   End
   Begin VB.Label Label3 
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
      TabIndex        =   7
      Top             =   144
      Width           =   732
   End
End
Attribute VB_Name = "Frmacc5200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/2/9 Form2.0不用改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Public adoacc040T As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset

Private Sub Command1_Click()
Dim intCounter As Integer

   If Text4 = MsgText(601) Then
      MsgBox MsgText(10) & Label3, , MsgText(5)
      strControlButton = MsgText(602)
      Text4.SetFocus
      Exit Sub
   Else
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0801 from acc080 where a0801 = '" & Text4 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label3
         adocheck.Close
         strControlButton = MsgText(602)
         Text4.SetFocus
         Exit Sub
      End If
      adocheck.Close
   End If
   If Text1 = MsgText(601) Then
      MsgBox MsgText(10) & Label1, , MsgText(5)
      strControlButton = MsgText(602)
      Text1.SetFocus
      Exit Sub
   Else
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0901 from acc090 where a0901 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label1
         adocheck.Close
         strControlButton = MsgText(602)
         Text1.SetFocus
         Exit Sub
      End If
      adocheck.Close
   End If
   If Text3 = MsgText(601) Then
      MsgBox MsgText(10) & Label2, , MsgText(5)
      strControlButton = MsgText(602)
      Text3.SetFocus
      Exit Sub
   Else
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0101 from acc010 where a0101 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label2
         adocheck.Close
         strControlButton = MsgText(602)
         Text3.SetFocus
         Exit Sub
      End If
      adocheck.Close
   End If
   If Text6 = MsgText(601) Then
      MsgBox MsgText(10) & Label4, , MsgText(5)
      strControlButton = MsgText(602)
      Text6.SetFocus
      Exit Sub
   End If
   adoacc040.CursorLocation = adUseClient
   adoacc040.Open "select * from acc040 where a0401 = " & Val(Text6) & " and a0403 = '" & Text4 & "' and a0404 = '" & Text1 & "' and a0405 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc040.RecordCount <> 0 Then
      MsgBox MsgText(9), , MsgText(5)
      strSaveConfirm = MsgText(4)
      QueryTable
   Else
      For intCounter = 1 To 12
         If Text1 = MsgText(601) Then
            adoTaie.Execute "insert into acc040 (a0401, a0402, a0403, a0404, a0405, a0406, a0407, a0408, a0409, a0411, a0412, a0413) values (" & Val(Text6) & ", " & intCounter & ", '" & Text4 & "', '" & MsgText(55) & "', '" & Text3 & "', 0, 0, 0, 0, " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "')"
         Else
            adoTaie.Execute "insert into acc040 (a0401, a0402, a0403, a0404, a0405, a0406, a0407, a0408, a0409, a0411, a0412, a0413) values (" & Val(Text6) & ", " & intCounter & ", '" & Text4 & "', '" & Text1 & "', '" & Text3 & "', 0, 0, 0, 0, " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "')"
         End If
      Next intCounter
      QueryTable
      adoacc040T.Requery
   End If
   adoacc040.Close
End Sub

Private Sub Command2_Click()
Dim strDepart As String

   If adoacc040T.RecordCount = 0 Then
      MsgBox MsgText(33), , MsgText(5)
      Exit Sub
   End If
   adoacc040T.MoveFirst
   adoacc040T.Find "a0401 = " & Val(Text6) & "", 0, adSearchForward, 1
   If adoacc040T.EOF = False Then
      adoacc040T.Find "a0403 = '" & Text4 & "'", 0, adSearchForward, adoacc040T.Bookmark
      If adoacc040T.EOF = False Then
         If Text1 <> MsgText(601) Then
            strDepart = Text1
         Else
            strDepart = MsgText(55)
         End If
         adoacc040T.Find "a0404 = '" & strDepart & "'", 0, adSearchForward, adoacc040T.Bookmark
         If adoacc040T.EOF = False Then
            adoacc040T.Find "a0405 = '" & Text3 & "'", 0, adSearchForward, adoacc040T.Bookmark
            If adoacc040T.EOF = False Then
               FormShow
               QueryTable
               RecordShow
               Exit Sub
            End If
         End If
      End If
   End If
   adoacc040T.MoveFirst
   QueryTable
   MsgBox MsgText(33), , MsgText(5)
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command2_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Select Case ColIndex
      Case 1
         If DataGrid1.Columns(1).Text = MsgText(601) Then
            DataGrid1.Columns(1).Value = 0
         End If
      Case 2
         If DataGrid1.Columns(2).Text = MsgText(601) Then
            DataGrid1.Columns(2).Value = 0
         End If
      Case 3
         If DataGrid1.Columns(3).Text = MsgText(601) Then
            DataGrid1.Columns(3).Value = 0
         End If
      Case 4
         If DataGrid1.Columns(4).Text = MsgText(601) Then
            DataGrid1.Columns(4).Value = 0
         End If
   End Select
   'Add by Morgan 2005/10/19 修改時要更新時間人員
   If strSaveConfirm = MsgText(4) Then
      Adodc1.Recordset.Fields("a0414") = strSrvDate(2)
      Adodc1.Recordset.Fields("a0415") = ServerTime
      Adodc1.Recordset.Fields("a0416") = strUserNum
   End If
   
   Adodc1.Recordset.UpdateBatch
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim intCounter As Integer

   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid1.col
            Case 1, 2, 3
               SendKeys "{RIGHT}"
            Case 4
               SendKeys "{DOWN}"
               For intCounter = 1 To 3
                  SendKeys "{LEFT}"
               Next intCounter
         End Select
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCon1 = MsgText(601) Then
      Exit Sub
   End If
   adoacc040T.Find "a0401 = " & Val(strCon1) & "", 0, adSearchForward, 1
   If adoacc040T.EOF = False Then
      adoacc040T.Find "a0403 = '" & strCon2 & "'", 0, adSearchForward, adoacc040T.Bookmark
      If adoacc040T.EOF = False Then
         adoacc040T.Find "a0404 = '" & strCon3 & "'", 0, adSearchForward, adoacc040T.Bookmark
         If adoacc040T.EOF = False Then
            adoacc040T.Find "a0405 = '" & strCon4 & "'", 0, adSearchForward, adoacc040T.Bookmark
            If adoacc040T.EOF = False Then
               FormShow
               QueryTable
               RecordShow
            End If
         End If
      End If
   End If
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
   Me.Height = 5550
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath5)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   If adoacc040T.RecordCount <> 0 Then
      adoacc040T.MoveLast
      adoacc040T.MoveFirst
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
   Set Frmacc5200 = Nothing
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Text2 = MsgText(601)
      Exit Sub
   End If
   Text2 = A0902Query(Text1)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text1 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0901 from acc090 where a0901 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label1
         adocheck.Close
         Cancel = True
         Exit Sub
      End If
      adocheck.Close
   End If
End Sub

Private Sub Text3_Change()
   If Text3 = MsgText(601) Then
      Exit Sub
   End If
   Text5 = A0102Query(Text3)
   adoacc010.CursorLocation = adUseClient
   adoacc010.Open "select a0103 from acc010 where a0101 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc010.RecordCount <> 0 Then
      If IsNull(adoacc010.Fields(0).Value) Then
         Text7 = MsgText(601)
      Else
         Select Case adoacc010.Fields(0).Value
            Case Mid(ComboItem(1), 1, 1)
               Text7 = ComboItem(1)
            Case Mid(ComboItem(2), 1, 1)
               Text7 = ComboItem(2)
         End Select
      End If
   Else
      Text7 = MsgText(601)
   End If
   adoacc010.Close
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a0101 from acc010 where a0101 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
      MessageShow Label2
      adocheck.Close
      Cancel = True
      Exit Sub
   End If
   adocheck.Close
End Sub

Private Sub Text4_Change()
   If Text4 = MsgText(601) Then
      Exit Sub
   End If
   Text13 = A0802Query(Text4)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc040T.CursorLocation = adUseClient
   adoacc040T.Open "select a0401, a0403, a0404, a0405 from acc040 group by a0401, a0403, a0404, a0405", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text6) & " and a0403 = '" & Text4 & "' and a0404 = '" & Text1 & "' and a0405 = '" & Text3 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(會計科目餘額資料)
'
'*************************************************
Public Sub QueryTable()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   If Text1 = MsgText(601) Then
      adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text6) & " and a0403 = '" & Text4 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text3 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
      adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text6) & " and a0403 = '" & Text4 & "' and a0404 = '" & Text1 & "' and a0405 = '" & Text3 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End If
   Adodc1.Recordset.Requery
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a0801 from acc080 where a0801 = '" & Text4 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adocheck.RecordCount = 0 Then
      MessageShow Label3
      adocheck.Close
      Cancel = True
      Exit Sub
   End If
   adocheck.Close
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text6 = MsgText(601) Then
      MsgBox Label4 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc040T.Bookmark & MsgText(35) & adoacc040T.RecordCount
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text4 = adoacc040T.Fields("a0403").Value
   If adoacc040T.Fields("a0404").Value = MsgText(55) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc040T.Fields("a0404").Value
   End If
   Text3 = adoacc040T.Fields("a0405").Value
   Text6 = adoacc040T.Fields("a0401").Value
   adoquery.CursorLocation = adUseClient
   If Text1 <> MsgText(601) Then
      adoquery.Open "select * from acc040 where a0401 = " & Val(Text6) - 1 & " and a0402 = 12 and a0403 = '" & Text4 & "' and a0404 = '" & Text1 & "' and a0405 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoquery.Open "select * from acc040 where a0401 = " & Val(Text6) - 1 & " and a0402 = 12 and a0403 = '" & Text4 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
   End If
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("a0408").Value) Then
         Text8 = MsgText(601)
      Else
         'Modify by Morgan 2005/10/19 改與資料庫位數相同
         'Text8 = Format(adoquery.Fields("a0408").Value, DDollar)
         Text8 = Format(adoquery.Fields("a0408").Value, FDollar)
      End If
   Else
      Text8 = MsgText(601)
   End If
   Select Case Mid(Text3, 1, 1)
      Case "4", "6", "7"
         Text8 = "0"
   End Select
   adoquery.Close
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
On Error GoTo Checking
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   If Text1 = MsgText(601) Then
      adoaccsum.Open "select sum(a0406) as Debit, sum(a0407) as Credit, sum(a0409) as Budget, sum(a0408) as Balance from acc040 where a0401 = " & Val(Text6) & " and a0403 = '" & Text4 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text3 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoaccsum.Open "select sum(a0406) as Debit, sum(a0407) as Credit, sum(a0409) as Budget, sum(a0408) as Balance from acc040 where a0401 = " & Val(Text6) & " and a0403 = '" & Text4 & "' and a0404 = '" & Text1 & "' and a0405 = '" & Text3 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields("Debit").Value) Then
         Text9 = MsgText(601)
      Else
         Text9 = Format(adoaccsum.Fields("Debit").Value, FDollar)
      End If
      If IsNull(adoaccsum.Fields("Credit").Value) Then
         Text10 = MsgText(601)
      Else
         Text10 = Format(adoaccsum.Fields("Credit").Value, FDollar)
      End If
      If IsNull(adoaccsum.Fields("Budget").Value) Then
         Text11 = MsgText(601)
      Else
         Text11 = Format(adoaccsum.Fields("Budget").Value, FDollar)
      End If
      If IsNull(adoaccsum.Fields("Balance").Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = Format(adoaccsum.Fields("Balance").Value, FDollar)
      End If
   Else
      Text9 = MsgText(601)
      Text10 = MsgText(601)
      Text11 = MsgText(601)
      Text12 = MsgText(601)
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

