VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm090209_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "公報簡訊資料彙整作業"
   ClientHeight    =   5715
   ClientLeft      =   90
   ClientTop       =   990
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
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
      Bindings        =   "frm090209_1.frx":0000
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "BB07"
         Caption         =   "公告日期"
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
         DataField       =   "BB01"
         Caption         =   "公告頁數"
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
         DataField       =   "BB02"
         Caption         =   "公告號數"
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
         DataField       =   "BB03"
         Caption         =   "國際分類1"
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
         DataField       =   "BB04"
         Caption         =   "國際分類2"
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
         DataField       =   "BB05"
         Caption         =   "國際分類3"
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
         DataField       =   "BB06"
         Caption         =   "索引"
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
         DataField       =   "BB08"
         Caption         =   "內容"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   945.071
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
         BeginProperty Column07 
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
      Caption         =   "公告日期：        　　　　　 －"
      Height          =   180
      Left            =   252
      TabIndex        =   0
      Top             =   540
      Width           =   2772
   End
End
Attribute VB_Name = "frm090209_1"
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
Dim s As Integer

'Add By Cheng 2003/03/03
'初始化列表
InitGrid
Select Case Index
       Case 0 '全部彙整
         'Modify By Cheng 2002/03/04
'         If Val(Text1) < Val(Text2) Then
         If Val(Text1) > Val(Text2) Then
             MsgBox "輸入資料範圍錯誤,請重新輸入", vbInformation
             Text1.Text = "": Text2.Text = ""
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
               Screen.MousePointer = vbDefault
             Else
               s = MsgBox("請輸入條件！！", , "USER 輸入錯誤")
               Text1.SetFocus
               Text1.SelStart = 0
               Text1.SelLength = Len(Text1)
               Exit Sub
             End If
         End If

'911107 nickchen
On Error GoTo CheckingErr

         If Adodc1.Recordset.RecordCount <> 0 Then
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         pemain.MoveFirst
         
         '911107 nickchen
         cnnConnection.BeginTrans
         
         Do While Not pemain.EOF
            cnnConnection.Execute "UPDATE BULLETINBRIEF SET BB09='1' WHERE bb02='" & CheckStr(pemain.Fields(1).Value) & "' and bb01='" & CheckStr(pemain.Fields(0).Value) & "' "
            pemain.MoveNext
         Loop
         
         '911107 nickchen
         cnnConnection.CommitTrans
         
         MsgBox "彙整完畢", vbInformation
         'Process
          Adodc1.Recordset.ReQuery
          DataGrid1.Refresh
          Me.Enabled = True
          Screen.MousePointer = vbDefault
         End If
       Case 1 '修改刪除
         If Adodc1.Recordset.RecordCount <> 0 Then
            frm090209_2.Show
            frm090209_2.Process CheckStr(pemain.Fields(0).Value), CheckStr(pemain.Fields(1).Value)
            frm090209_1.Hide
         End If
       Case 2 '結束
         Unload Me
       Case 3 '搜尋
'         Text2_LostFocus
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
                    'Add By Cheng 2003/04/10
                    '若無資料
                    If pemain.RecordCount <= 0 Then
                        Me.Text1.SetFocus
                        Text1_GotFocus
                    End If
                    Screen.MousePointer = vbDefault
                End If
            End If
End Select

 '911107 nick transation
     Exit Sub
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
End Sub

Private Sub DataGrid1_Click()
    'Modify By Cheng 2003/03/03
    '預設修改/刪除按錄
    'Me.cmdOK(1).SetFocus
    Me.CMDOK(1).Default = True
End Sub

Private Sub DataGrid1_LostFocus()
    'Add By Cheng 2003/03/03
    '預設搜尋按鈕
    Me.CMDOK(3).Default = True
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    pemain.CursorLocation = adUseClient
    'Modify By Cheng 2002/03/04
'    cmdOK(1).Default = True
    'Add By Cheng 2003/03/03
    InitGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090209_1 = Nothing
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
Else
    'Modify By Cheng 2003/04/10
    '取消在此搜尋資料
'    If Len(Trim(Text2)) <> 0 Then
'      Screen.MousePointer = vbHourglass
'      Me.Enabled = False
'      Process
'      Me.Enabled = True
'      Screen.MousePointer = vbDefault
'    End If
End If

End Sub

Sub Process()
   Screen.MousePointer = vbHourglass
    If pemain.State = adStateOpen Then pemain.Close
    'Modify By Cheng 2003/01/06
    '依公告日期及頁數由小至大排序
'    strExc(0) = "SELECT BB01,BB02,BB03,BB04,BB05,BB06,BB07,BB08 FROM BULLETINBRIEF WHERE BB07>=" & ChangeTStringToWString(Text1.Text) & " AND BB07 <=" & ChangeTStringToWString(Text2.Text) & " and (bb09<>'1' OR BB09 IS NULL) ORDER BY 1"
    'Modify By Cheng 2003/01/14
    '依公告日期及頁數及公告號數由小至大排序
'    strExc(0) = "SELECT BB01,BB02,BB03,BB04,BB05,BB06,BB07,BB08 FROM BULLETINBRIEF WHERE BB07>=" & ChangeTStringToWString(Text1.Text) & " AND BB07 <=" & ChangeTStringToWString(Text2.Text) & " and (bb09<>'1' OR BB09 IS NULL) ORDER BY BB07, BB01 "
    'Modify by Morgan 2004/11/29 公告號會有英文，排序不轉數字
    'strExc(0) = "SELECT BB01,BB02,BB03,BB04,BB05,BB06,BB07,BB08 FROM BULLETINBRIEF WHERE BB07>=" & ChangeTStringToWString(Text1.Text) & " AND BB07 <=" & ChangeTStringToWString(Text2.Text) & " and (bb09<>'1' OR BB09 IS NULL) ORDER BY TO_Number(BB07), TO_Number(BB01), TO_Number(BB02) "
    strExc(0) = "SELECT BB01,BB02,BB03,BB04,BB05,BB06,BB07,BB08 FROM BULLETINBRIEF WHERE BB07>=" & ChangeTStringToWString(Text1.Text) & " AND BB07 <=" & ChangeTStringToWString(Text2.Text) & " and (bb09<>'1' OR BB09 IS NULL) ORDER BY TO_Number(BB07), TO_Number(BB01), BB02"
    pemain.CursorLocation = adUseClient
    pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
    If pemain.RecordCount <> 0 Then
        Set Adodc1.Recordset = pemain
        Adodc1.Recordset.ReQuery
    Else
        Set Adodc1.Recordset = pemain
        Screen.MousePointer = vbDefault
        ShowNoData
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
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
        .Columns(1).Width = 500
        .Columns(2).Width = 700
        .Columns(3).Width = 450
        .Columns(4).Width = 250
        .Columns(5).Width = 250
        .Columns(6).Width = 300
        .Columns(7).Width = 6000
    End With
End Sub
