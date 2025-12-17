VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1194 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "銷帳退費作業(分錄資料維護)"
   ClientHeight    =   4340
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4340
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1194.frx":0000
      Height          =   1740
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   8445
      _ExtentX        =   14887
      _ExtentY        =   3069
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a0102"
         Caption         =   "會計科目"
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
         DataField       =   "a1p07"
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
         DataField       =   "a1p08"
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
         DataField       =   "a1p14"
         Caption         =   "摘要"
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
         DataField       =   "a1p23"
         Caption         =   "轉出單號"
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
            ColumnWidth     =   3339.78
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1310.173
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1280.126
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   5559.875
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1539.78
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6615
      Picture         =   "Frmacc1194.frx":0015
      Style           =   1  '圖片外觀
      TabIndex        =   26
      ToolTipText     =   "取消"
      Top             =   1905
      Width           =   350
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6780
      MaxLength       =   9
      TabIndex        =   7
      Top             =   3390
      Width           =   1572
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1185
      MaxLength       =   3
      TabIndex        =   5
      Top             =   3390
      Width           =   615
   End
   Begin VB.TextBox Text33 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      MaxLength       =   12
      TabIndex        =   13
      Top             =   1965
      Width           =   855
   End
   Begin VB.TextBox Text34 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5055
      TabIndex        =   12
      Top             =   1965
      Width           =   1308
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6795
      MaxLength       =   14
      TabIndex        =   2
      Top             =   2655
      Width           =   1572
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4875
      MaxLength       =   14
      TabIndex        =   1
      Top             =   2655
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1785
      TabIndex        =   9
      Top             =   2655
      Width           =   2652
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      MaxLength       =   6
      TabIndex        =   0
      Top             =   2655
      Width           =   1572
   End
   Begin VB.TextBox Text43 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   1965
      Width           =   1332
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1185
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3015
      Width           =   720
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3915
      MaxLength       =   12
      TabIndex        =   6
      Top             =   3390
      Width           =   1572
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3765
      Width           =   1665
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   7380
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   2117
      _ExtentY        =   564
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
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Left            =   3915
      TabIndex        =   4
      Top             =   3000
      Width           =   4455
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7858;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text17 
      Height          =   330
      Left            =   1890
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3015
      Width           =   1020
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   8
      Size            =   "1799;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label35 
      BackStyle       =   0  '透明
      Caption         =   "對沖(客)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5835
      TabIndex        =   24
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label37 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   23
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label38 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   22
      Top             =   3045
      Width           =   975
   End
   Begin VB.Label Label39 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   21
      Top             =   1965
      Width           =   855
   End
   Begin VB.Label Label40 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6795
      TabIndex        =   20
      Top             =   2415
      Width           =   1575
   End
   Begin VB.Label Label41 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2955
      TabIndex        =   19
      Top             =   3045
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      Height          =   1830
      Left            =   90
      Top             =   2355
      Width           =   8475
   End
   Begin VB.Label Label48 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4875
      TabIndex        =   18
      Top             =   2415
      Width           =   1575
   End
   Begin VB.Label Label49 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   17
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label Label50 
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3030
      TabIndex        =   16
      Top             =   1965
      Width           =   495
   End
   Begin VB.Label Label36 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2955
      TabIndex        =   15
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label34 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   14
      Top             =   3795
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1194"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Add by Morgan 2006/12/25
Option Explicit

Dim strSerialNo As String
Dim strA1P04 As String '銷帳單號
Dim strA1P23 As String '轉出單號
Dim strA1P23_1 As String '轉出單號2
Dim strA1P01 As String '公司別 Added by Morgan 2014/8/6
Dim bolIsReceipt As Boolean '是否為收據銷退 Added by Morgan 2015/8/12

Private Sub Combo2_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
End Sub

Private Sub Combo2_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Command2_Click()
   If Not bolIsReceipt Then
      If strSerialNo <= "001" Then
         MsgBox "前兩筆分錄資料只能修改不可刪除！"
         Exit Sub
      End If
   End If
   Adodc1.Recordset.Find "a1p03=" & strSerialNo, 0, adSearchForward, 0
   If Not Adodc1.Recordset.EOF Then
      Adodc1.Recordset.Delete
      Adodc1.Recordset.UpdateBatch
   End If
   DataGrid1.Refresh
   SumShow
   ClearInput
End Sub

Private Sub DataGrid1_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Add By Sindy 2024/8/1 會出現錯誤訊息
   '可能是 BOF 或 EOF 的值為 True，或目前的資料錄已被刪除。所要求的操作需要目前的資料錄。
   If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
   '2024/8/1 END
      strSerialNo = Adodc1.Recordset.Fields("a1p03").Value
      AdodcShow
   End If
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(0, KeyCode) 'Added by Morgan 2022/7/19
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, 8760, 4665, strBackPicPath1
   DataGrid1.ScrollBars = dbgBoth
   strA1P04 = Frmacc1190.Text2
   strA1P23 = Frmacc1190.Text22
   strA1P23_1 = Frmacc1190.Text28
   AdodcRefresh
   
   'Added by Morgan 2015/8/12
   If Frmacc1190.Option1.Value = True Then
      bolIsReceipt = True
      Text6.Enabled = True
   Else
      'Modified by Morgan 2022/4/29
      'bolIsReceipt = True
      bolIsReceipt = False
   End If
   'end 2015/8/12
End Sub

Private Function FormSave() As Boolean
   
   Dim a1p(1 To 31) As String
   Dim strA1P18 As String '入帳日期
   Dim stra1p22 As String '傳票編號
   Dim stra1p27 As String '是否更新
   Dim rsACC1P0 As New ADODB.Recordset
   Dim lngA0T08 As Long
   
   adoTaie.BeginTrans
   
On Error GoTo ErrHnd
   '刪除舊資料
   'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
   strSql = "delete from acc1p0 where a1p02 = 'Z' and a1p04 = '" & strA1P04 & "'"
   adoTaie.Execute strSql, intI
   strSql = "select * from acc1p0 where a1p02 = 'Z' and a1p04 = '" & strA1P04 & "'"
   rsACC1P0.CursorLocation = adUseClient
   rsACC1P0.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   With Adodc1.Recordset
      If .RecordCount > 0 Then
      .MoveFirst
      strA1P18 = "" & .Fields("A1P18")
      stra1p22 = "" & .Fields("A1P22")
      stra1p27 = "" & .Fields("A1P27")
      Do While Not .EOF
         rsACC1P0.AddNew
         rsACC1P0.Fields("a1p01") = .Fields("a1p01")
         rsACC1P0.Fields("a1p02") = .Fields("a1p02")
         rsACC1P0.Fields("a1p03") = .Fields("a1p03")
         rsACC1P0.Fields("a1p04") = strA1P04
         rsACC1P0.Fields("a1p05") = .Fields("a1p05")
         rsACC1P0.Fields("a1p06") = .Fields("a1p06")
         rsACC1P0.Fields("a1p07") = .Fields("a1p07")
         rsACC1P0.Fields("a1p08") = .Fields("a1p08")
         rsACC1P0.Fields("a1p14") = .Fields("a1p14")
         rsACC1P0.Fields("a1p15") = .Fields("a1p15")
         rsACC1P0.Fields("a1p16") = .Fields("a1p16")
         rsACC1P0.Fields("a1p17") = .Fields("a1p17")
         rsACC1P0.Fields("a1p18") = strA1P18
         rsACC1P0.Fields("a1p22") = stra1p22
         '暫收轉暫收
         'Modified by Morgan 2022/4/29 +檢查不是收據的銷退
         'If .Fields("a1p05") = "2401" And .Fields("a1p08") > 0 Then
         If Not bolIsReceipt And .Fields("a1p05") = "2401" And .Fields("a1p08") > 0 Then
         'end 2022/4/29
            '原先無轉出單號2
            If strA1P23_1 = "" Then
               '新增
               strA1P23_1 = AutoNo(MsgText(806), 5, 1)
               strSql = "insert into acc0t0 (a0t01, a0t02, a0t03, a0t04, a0t08, a0t13, a0t11, a0t12, a0t05, a0t06)" & _
                  " values ('" & strA1P23_1 & "', '3', " & strA1P18 & ", " & strA1P18 & "," & .Fields("a1p08") & ", '" & strUserNum & "', " & strSrvDate(2) & ",TO_CHAR(SYSDATE,'HH24MISS'), '" & .Fields("a1p16") & "', '" & .Fields("a1p15") & "')"
               adoTaie.Execute strSql, intI
               strSql = "Update acc0s0 set a0s23='" & strA1P23_1 & "' where a0s01='" & strA1P04 & "'"
               adoTaie.Execute strSql, intI
               Frmacc1190.Text28 = strA1P23_1
               lngA0T08 = .Fields("a1p08")
            Else
               '修改
               lngA0T08 = lngA0T08 + .Fields("a1p08")
               strSql = "UPDATE acc0t0 set a0t08=" & lngA0T08 & ", a0t05='" & .Fields("a1p16") & "',a0t06='" & .Fields("a1p15") & "',a0t16='" & strUserNum & "',a0t14=" & strSrvDate(2) & ",a0t15=TO_CHAR(SYSDATE,'HH24MISS') where a0t01='" & strA1P23_1 & "'"
               adoTaie.Execute strSql, intI
            End If
            rsACC1P0.Fields("a1p23") = strA1P23_1
            If InStr(.Fields("a1p14"), strA1P23_1) = 0 Then
               rsACC1P0.Fields("a1p14") = .Fields("a1p14") & "/" & strA1P23_1
            End If
         Else
            rsACC1P0.Fields("a1p23") = .Fields("a1p23")
            rsACC1P0.Fields("a1p30") = .Fields("a1p30")  '2012/9/26 ADD BY SONIA
         End If
         
         
         'Modified by Morgan 2017/4/18 有傳票號時改分錄是否更新要上Y
         'rsACC1P0.Fields("a1p27") = stra1p27
         If stra1p22 <> "" Then
            rsACC1P0.Fields("a1p27") = "Y"
         Else
            rsACC1P0.Fields("a1p27") = stra1p27
         End If
         'end 2017/4/18
         .MoveNext
      Loop
      '若轉暫收款金額為0時刪除轉出單號2
      If strA1P23_1 <> "" And lngA0T08 = 0 Then
         strSql = "delete from acc0t0 where a0t01='" & strA1P23_1 & "'"
         adoTaie.Execute strSql, intI
         strA1P23_1 = ""
         If Frmacc1190.Text28 <> "" Then
            strSql = "Update acc0s0 set a0s23=null where a0s01='" & strA1P04 & "'"
            adoTaie.Execute strSql, intI
            Frmacc1190.Text28 = ""
         End If
      Else
      
      End If
      rsACC1P0.UpdateBatch
      End If
   End With
   adoTaie.CommitTrans
   FormSave = True
   With Frmacc1190
      If .Option2.Value = True Then
         .adoacc0t0.Requery
         .adoacc0t0.Find "a0t01 = '" & .Text1 & "'", 0, adSearchForward, 1
      End If
   End With
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description
   Set rsACC1P0 = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   If Text43.Text <> Text34.Text Then
      MsgBox "借貸方金額不平衡，請再確認...", vbExclamation
      Cancel = 1
      Exit Sub
   End If
   
   'Added by Morgan 2015/8/12
   Adodc1.Recordset.MoveFirst
   Adodc1.Recordset.Find "a1p07=0"
   'Modified by Morgan 2022/4/28 有可能是新增借方
   'If Not Adodc1.Recordset.EOF Then
   '   Adodc1.Recordset.Find "a1p08=0"
   '   If Not Adodc1.Recordset.EOF Then
   '      Cancel = 1
   '      DataGrid1_Click
   '      MsgBox "請輸入" & Text5 & "的金額...", vbExclamation
   '      Exit Sub
   '   End If
   'End If
   With Adodc1.Recordset
   Do While Not .EOF
      If .Fields("a1p08") = 0 And .Fields("a1p07") = 0 Then
         Cancel = 1
         DataGrid1_Click
         MsgBox "請輸入" & Text5 & "的金額...", vbExclamation
         Exit Sub
      End If
      .Find "a1p08=0", 1
   Loop
   End With
   'end 2022/4/28
   'end 2015/8/12
   
   If FormSave = False Then
      Cancel = 1
      Exit Sub
   End If
   strTrackMode = "" 'Added by Morgan 2022/7/19
   tool1_enabled
   Frmacc1190.Enabled = True
   Set Frmacc1194 = Nothing
End Sub

Private Sub KeyDefine(KeyCode As Integer)
   Call PUB_SaveTrackMode(1, KeyCode) 'Added by Morgan 2022/7/19
   Select Case KeyCode
      Case vbKeyInsert
         If PUB_ChkTrackMode = False Then Exit Sub 'Added by Morgan 2022/7/19
         GridUpdate
      Case Else
         KeyEnter KeyCode
   End Select
End Sub

Private Sub AdodcRefresh()
   'Modified by Morgan 2014/8/6 會有J公司,取消 a1p01='1' 條件
   strExc(0) = "select a0102,acc1p0.* from acc1p0, acc010 where a1p05 = a0101  and a1p02 = 'Z'" & _
      " and a1p04 = '" & strA1P04 & "' order by a1p03 asc"
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Added by Morgan 2014/8/6
   If intI = 1 Then
      strA1P01 = RsTemp.Fields("a1p01")
   Else
      strA1P01 = "1"
   End If
   'end 2014/8/6
   'Modify by Amy 2014/06/24 +FormName 改暫存TB
   Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   SumShow
End Sub

'*************************************************
'  計算並顯示總計
'
'*************************************************
Private Sub SumShow()
   
   Text33 = "0"
   Text43 = "0"
   Text34 = "0"
   Set RsTemp = Adodc1.Recordset.Clone
   With RsTemp
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         Text33 = Val(Text33) + 1
         Text43 = Val(Text43) + Val("" & .Fields("a1p07"))
         Text34 = Val(Text34) + Val("" & .Fields("a1p08"))
         .MoveNext
      Loop
      Text43 = Format(Text43, FDollar)
      Text34 = Format(Text34, FDollar)
      Adodc1.Recordset.MoveFirst
      Adodc1.Recordset.Find "a1p03='" & strSerialNo & "'"
   End If
   End With
End Sub

'*************************************************
'  顯示分錄檔資料
'
'*************************************************
Private Sub AdodcShow()
   Text4 = Adodc1.Recordset.Fields("a1p05").Value
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a1p08").Value
   End If

   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text13 = MsgText(601)
      Text17 = ""
   Else
      Text13 = Adodc1.Recordset.Fields("a1p16").Value
      Text17 = StaffQuery(Text13)
   End If
   Combo2 = "" & Adodc1.Recordset.Fields("a1p14").Value
   Text14 = "" & Adodc1.Recordset.Fields("a1p06").Value
   Text15 = "" & Adodc1.Recordset.Fields("a1p17").Value
   Text16 = "" & Adodc1.Recordset.Fields("a1p15").Value
   Text18 = "" & Adodc1.Recordset.Fields("a1p30").Value
   
   '第一筆分錄固定為暫收款借方且科目金額不可改，其他為貸方
   If Not bolIsReceipt And strSerialNo = "001" Then
      Text4.Enabled = False
      Text11.Enabled = False
      'modify by sonia 2025/5/9
      'Text13.SetFocus
      If Text13.Enabled = True Then Text13.SetFocus
      'end 2025/5/9
   Else
      Text4.Enabled = True
      Text11.Enabled = True
      Text4.SetFocus
   End If
   
   'add by sonia 2025/5/9 2401且有暫收單號時鎖住業務編號及客戶編號
   If Text4 = "2401" And Left(Text18, 1) = "J" Then
      Text13.Enabled = False
      Text16.Enabled = False
   Else
      Text13.Enabled = True
      Text16.Enabled = True
   End If
   'end 2025/5/9
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   'Add by Morgan 2007/2/5 員工已離職要提醒
   If Text13 <> "" Then
      If PUB_GetStaffState(Text13.Text, strExc(1), True) = 0 Then
         Text13.SetFocus
         Cancel = True
         TextInverse Text13
      Else
         Text17 = strExc(1)
      End If
   End If
   'end 2007/2/5
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
   If Text14 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text14, Label37) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Text4, Text14) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text15_Validate(Cancel As Boolean)

   If Text15 <> MsgText(601) Then
      Text15 = CaseNoZero(Text15)
      If Len(Text15) < 10 Then
         MsgBox MsgText(28) & Label36, , MsgText(5)
         Cancel = True
         Exit Sub
      End If
      strExc(0) = "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and pa02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and pa03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and pa04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
         "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and tm02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and tm03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and tm04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
         "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and lc02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and lc03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and lc04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
         "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and hc02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and hc03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and hc04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
         "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and sp02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and sp03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and sp04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 1 Then
         MsgBox MsgText(28) & Label36, , MsgText(5)
         Cancel = True
         Exit Sub
      End If
   End If
   
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Add by Morgan 2007/3/1
Private Sub Text16_Validate(Cancel As Boolean)
   If Text16 <> MsgText(601) Then
      If Len(Text16) = 6 Then
         Text16 = AfterZero(Text16)
      ElseIf Len(Text16) = 8 Then
         Text16 = Text16 & "0"
      End If
      If ExistCheck("customer", "cu01", Mid(Text16, 1, 8), Label35, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text16, Label35, False) = False Then
            If ExistCheck("staff", "st01", Text16, Label35, False) = False Then
               MsgBox MsgText(28) & Label35, , MsgText(5)
               Cancel = True
               Exit Sub
            End If
         End If
      End If
   End If
End Sub

Private Sub Text18_GotFocus()
   TextInverse Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'End 2007/3/1
Private Sub Text4_Change()
   Text5 = A0102Query(Text4)
End Sub

Private Sub GridUpdate()
   
   If Text4 = "" Then
      MsgBox "會計科目不可空白！"
      Text4.SetFocus
      Exit Sub
   End If
   If Val(Text6) = 0 And Val(Text11) = 0 Then
      MsgBox "金額不可全為0(空白)！"
      If Text11.Enabled = True Then
         Text11 = ""
         Text11.SetFocus
      End If
      Exit Sub
   End If
   If Text14 = "" Then
      MsgBox "部門別不可空白！"
      Text14.SetFocus
      Exit Sub
   End If
   
   If Text4 = "2401" Or Text4 = "2112" Then
      If Text16 = "" Then
         MsgBox "客戶不可空白！"
         Text16.SetFocus
         Exit Sub
      End If
      If Text13 = "" Then
         MsgBox "智權人員不可空白！"
         Text13.SetFocus
         Exit Sub
      End If
   End If
   
   'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(Text4, Val(FCDate(Frmacc1190.MaskEdBox1.Text)))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      Text4.SetFocus
      Exit Sub
   End If
   'end 2015/12/30
   'Add by Morgan 2007/2/2 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Text4, Text14, Text13)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Text4.SetFocus
      ElseIf intI = 2 Then
         Text14.SetFocus
      ElseIf intI = 3 Then
         Text13.SetFocus
      End If
      Exit Sub
   End If
   'end 2007/2/2
   
   'Modify by Amy 2014/06/24 +單引號 修正因為文字型態又為空值時產生錯誤(由前畫面按修改至此畫面新增)
   Adodc1.Recordset.Find "a1p03='" & strSerialNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF Then
      If Adodc1.Recordset.RecordCount > 0 Then
         Adodc1.Recordset.MoveLast
         strSerialNo = Format(Val(Adodc1.Recordset("a1p03")) + 1, "00#")
      Else
         strSerialNo = "001"
      End If
      Adodc1.Recordset.AddNew
      'Added by Morgan 2014/8/6
      'Adodc1.Recordset.Fields("a1p01") = "1"
      Adodc1.Recordset.Fields("a1p01") = strA1P01
      'end 2014/8/6
      Adodc1.Recordset.Fields("a1p02") = "Z"
      Adodc1.Recordset.Fields("a1p03") = strSerialNo
      Adodc1.Recordset.Fields("a1p04") = strA1P04
   End If
   
   Adodc1.Recordset.Fields("a1p05") = Text4
   Adodc1.Recordset.Fields("a0102") = Text5
   If Val(Text6) > 0 Then
      Adodc1.Recordset.Fields("a1p07").Value = Format(Val(Text6))
   Else
      Adodc1.Recordset.Fields("a1p07").Value = 0
   End If
   If Val(Text11) > 0 Then
      Adodc1.Recordset.Fields("a1p08").Value = Format(Val(Text11))
   Else
      Adodc1.Recordset.Fields("a1p08").Value = 0
   End If
   
   If Text13 <> "" Then
      Adodc1.Recordset.Fields("a1p16").Value = Text13
   Else
      Adodc1.Recordset.Fields("a1p16").Value = Null
   End If
   If Combo2 <> "" Then
      Adodc1.Recordset.Fields("a1p14").Value = Combo2
      Combo2.AddItem Combo2.Text
   Else
      Adodc1.Recordset.Fields("a1p14").Value = Null
   End If
   
   Adodc1.Recordset.Fields("a1p06").Value = Text14
   
   If Text15 <> "" Then
      Adodc1.Recordset.Fields("a1p17").Value = Text15
   Else
      Adodc1.Recordset.Fields("a1p17").Value = Null
   End If
   
   Adodc1.Recordset.Fields("a1p18").Value = Val(FCDate(Frmacc1190.MaskEdBox1.Text))  'add by sonia 2014/9/17 I10300773
   
   If Text16 <> "" Then
      Adodc1.Recordset.Fields("a1p15").Value = Text16
   Else
      Adodc1.Recordset.Fields("a1p15").Value = Null
   End If
   If Text18 <> "" Then
      Adodc1.Recordset.Fields("a1p30").Value = Text18
   Else
      Adodc1.Recordset.Fields("a1p30").Value = Null
   End If
   '暫收款
   If Text4 = "2401" Then
      '借方
      If Val(Text6) > 0 Then
         Adodc1.Recordset.Fields("a1p23").Value = strA1P23
         'Added by Morgan 2015/8/21 2401的其他對沖要放暫收款單號
         If IsNull(Adodc1.Recordset.Fields("a1p30").Value) And Left(Frmacc1190.Text1, 1) = "J" Then
            Adodc1.Recordset.Fields("a1p30").Value = Frmacc1190.Text1
         End If
         'end 2015/8/21
      '貸方
      Else
         Adodc1.Recordset.Fields("a1p23").Value = strA1P23_1
         'Added by Morgan 2015/8/21 2401的其他對沖要放暫收款單號
         If IsNull(Adodc1.Recordset.Fields("a1p30").Value) And Left(strA1P23_1, 1) = "J" Then
            Adodc1.Recordset.Fields("a1p30").Value = strA1P23_1
         End If
         'end 2015/8/21
      End If
   '應付款
   ElseIf Text4 = "2112" Then
      Adodc1.Recordset.Fields("a1p23").Value = strA1P23
   Else
      Adodc1.Recordset.Fields("a1p23").Value = Null
   End If
   
   Adodc1.Recordset.UPDATE
   SumShow
   ClearInput
End Sub

Private Sub ClearInput()
   strSerialNo = 0
   Text4 = ""
   Text6 = ""
   Text11 = ""
   Text13 = ""
   Combo2 = ""
   Text14 = ""
   Text15 = ""
   Text16 = ""
   Text18 = ""
   If Text4.Visible = True And Text4.Enabled = True Then Text4.SetFocus
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

