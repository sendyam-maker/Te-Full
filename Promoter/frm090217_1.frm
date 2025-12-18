VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm090217_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案例資料彙整作業"
   ClientHeight    =   5715
   ClientLeft      =   90
   ClientTop       =   990
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton CMDOK 
      Caption         =   "搜尋(&S)"
      Default         =   -1  'True
      Height          =   400
      Index           =   3
      Left            =   4920
      TabIndex        =   7
      Top             =   70
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   510
      Top             =   5520
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
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8472
      TabIndex        =   6
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "修改刪除(M)"
      Height          =   400
      Index           =   1
      Left            =   7290
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "全部彙整(&A)"
      Height          =   400
      Index           =   0
      Left            =   6090
      TabIndex        =   4
      Top             =   70
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm090217_1.frx":0000
      Height          =   4884
      Left            =   36
      TabIndex        =   3
      Top             =   792
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8625
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   16
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "PC20"
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
      BeginProperty Column01 
         DataField       =   "PC01020304"
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
      BeginProperty Column02 
         DataField       =   "PC19C"
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
         DataField       =   "PC05060708"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         AllowFocus      =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   989.858
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   2892
      TabIndex        =   2
      Top             =   492
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1212
      TabIndex        =   1
      Top             =   492
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "文書日期：        　　　　　 －"
      Height          =   180
      Left            =   252
      TabIndex        =   0
      Top             =   540
      Width           =   2772
   End
End
Attribute VB_Name = "frm090217_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (DataGrid1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset

Private Sub cmdOK_Click(Index As Integer)

   Dim s As Integer, stKeys As String

On Error GoTo ErrFlg

   InitGrid
   Select Case Index
      Case 0 '全部彙整
         If Val(Text1) > Val(Text2) Then
             MsgBox "輸入資料範圍錯誤,請重新輸入", vbInformation
             Text1.Text = "": Text2.Text = ""
             Text1.SetFocus
             Text1.SelStart = 0
             Text1.SelLength = Len(Text1)
             Exit Sub
         End If
         
         If Adodc1.Recordset Is Nothing Then Exit Sub 'Added by Morgan 2022/1/14
         
         If Adodc1.Recordset.RecordCount <> 0 Then
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            pemain.MoveFirst
            stKeys = "''"
            Do While Not pemain.EOF
               stKeys = stKeys & ",'" & pemain.Fields("PC01").Value & pemain.Fields("PC02").Value & pemain.Fields("PC03").Value & pemain.Fields("PC04").Value & "'"
               pemain.MoveNext
            Loop
            
            cnnConnection.Execute "UPDATE PATENTCASE SET PC18='1' WHERE PC01||PC02||PC03||PC04 IN (" & stKeys & ")"
           
            MsgBox "彙整完畢", vbInformation
            Adodc1.Recordset.ReQuery
            DataGrid1.Refresh
            Me.Enabled = True
            Screen.MousePointer = vbDefault
         End If
         
      Case 1 '修改刪除
         If Adodc1.Recordset.RecordCount <> 0 Then
            frm090217_2.Show
            If frm090217_2.Process("" & pemain.Fields("PC01").Value & pemain.Fields("PC02").Value & pemain.Fields("PC03").Value & pemain.Fields("PC04").Value) Then
               frm090217_1.Hide
            Else
               Unload frm090217_2
            End If
         End If
         
      Case 2 '結束
         Unload Me
        
      Case 3 '搜尋
         If Val(Text1) > Val(Text2) Then
             MsgBox "輸入資料範圍錯誤,請重新輸入", vbInformation
             Text1.SetFocus
             Text1.SelStart = 0
             Text1.SelLength = Len(Text1)
             Exit Sub
         Else
            If Len(Trim(Text2)) <> 0 Then
               Screen.MousePointer = vbHourglass
               Me.Enabled = False
               Process
               Me.Enabled = True
               If pemain.RecordCount <= 0 Then
                   Me.Text1.SetFocus
                   Text1_GotFocus
               End If
               Screen.MousePointer = vbDefault
            End If
         End If
      End Select
      
ErrFlg:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
     
End Sub

Private Sub DataGrid1_Click()
    Me.cmdOK(1).Default = True
End Sub

Private Sub DataGrid1_LostFocus()
    Me.cmdOK(3).Default = True
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    pemain.CursorLocation = adUseClient
    InitGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090217_1 = Nothing
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1.Text <> "" Then
      If CheckIsTaiwanDate(Text1.Text) = False Then
          Text1.SetFocus
          Text1.SelStart = 0
          Text1.SelLength = Len(Text1)
          Cancel = True
      Else
          Cancel = False
      End If
   End If
End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)
End Sub

Private Sub Text2_LostFocus()

   If Val(Text1) > Val(Text2) Then
       MsgBox "輸入資料範圍錯誤,請重新輸入", vbInformation
       Text1.SetFocus
       Text1.SelStart = 0
       Text1.SelLength = Len(Text1)
       Exit Sub
   End If

End Sub

Sub Process()
   If pemain.State = adStateOpen Then pemain.Close
   strExc(0) = "SELECT PC20, PC01||'-'||PC02||'-'||PC03||'-'||PC04 AS PC01020304, DECODE(PC19,'0','判決','1','決定書','2','其他') PC19C, PC05||'-'||PC06||'-'||PC07||'-'||PC08 PC05060708, PC09, PC10, PC11" & _
      ", PC01, PC02, PC03, PC04, PC05, PC06, PC07, PC08" & _
      " FROM PATENTCASE WHERE PC20 BETWEEN " & ChangeTStringToWString(Text1.Text) & " AND " & ChangeTStringToWString(Text2.Text) & " and (PC18<>'1' OR PC18 IS NULL) ORDER BY TO_Number(PC20), PC01, PC02, PC03, PC04"
   pemain.CursorLocation = adUseClient
   pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If pemain.RecordCount <> 0 Then
      Set Adodc1.Recordset = pemain
      Adodc1.Recordset.ReQuery
   Else
      Set Adodc1.Recordset = pemain
      Screen.MousePointer = vbDefault
      ShowNoData
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2.Text <> "" Then
      If CheckIsTaiwanDate(Text1.Text) = False Then
         Text1.SetFocus
         Text1.SelStart = 0
         Text1.SelLength = Len(Text1)
         Cancel = True
         Exit Sub
      Else
         Cancel = False
      End If
   End If
End Sub
'Add By Cheng 2003/03/03
Private Sub InitGrid()
    With Me.DataGrid1
        .Columns(0).Width = 800
        .Columns(1).Width = 1150
        .Columns(2).Width = 800
        .Columns(3).Width = 1200
        .Columns(4).Width = 1500
        .Columns(5).Width = 1500
        .Columns(6).Width = 3000
        
    End With
End Sub
