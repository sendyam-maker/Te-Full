VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc31d0 
   AutoRedraw      =   -1  'True
   Caption         =   "銀行往來調節作業"
   ClientHeight    =   5112
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5112
   ScaleWidth      =   8760
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   12
      Top             =   4608
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   11
      Top             =   4608
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc31d0.frx":0000
      Height          =   3100
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   5440
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "a0e02"
         Caption         =   "票據號碼"
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
         DataField       =   "a0e11"
         Caption         =   "票據金額"
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
      BeginProperty Column02 
         DataField       =   "a0g02"
         Caption         =   "收票銀行"
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
         DataField       =   "a0e07"
         Caption         =   "收票帳號"
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
         DataField       =   "a0e10"
         Caption         =   "到期日期"
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
      BeginProperty Column05 
         DataField       =   "a0e08"
         Caption         =   "票別"
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
            ColumnWidth     =   1476.284
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2615.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1476.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1008
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   587.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
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
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1080
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7920
      Picture         =   "Frmacc31d0.frx":0015
      Style           =   1  '圖片外觀
      TabIndex        =   4
      ToolTipText     =   "取消"
      Top             =   1030
      Width           =   492
   End
   Begin VB.TextBox Text2 
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
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text1 
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
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1320
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
   Begin MSForms.TextBox Text7 
      Height          =   300
      Left            =   5670
      TabIndex        =   7
      Top             =   240
      Width           =   2775
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "金額合計"
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
      Left            =   2280
      TabIndex        =   14
      Top             =   4608
      Width           =   972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      Left            =   240
      TabIndex        =   13
      Top             =   4608
      Width           =   492
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   852
      Left            =   240
      Top             =   120
      Width           =   8292
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4248
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
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
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "調整日期"
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
      Left            =   360
      TabIndex        =   9
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "開票帳號"
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
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "開票銀行"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc31d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/20 Form2.0已修改 Text7/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc0e0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset

Private Sub Command2_Click()
   AdodcDelete
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'Mark by Amy 2023/08/07 系統會自動帶下一票據號碼,有時需按二次[Insert]鍵才會有動作(目前測只會觸發一次)
''Add by Amy 2021/10/20
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
'End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2023/08/07 預帶開票帳號-瑞婷
   Text2 = "1756650"
   Call Text2_Validate(False)
   'end 2023/08/07
   MaskEdBox1.Mask = DFormat
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   strTrackMode = "" 'Add by Amy 2021/10/20 Form2.0 記錄鍵盤傳入順序(清除)
   Set Frmacc31d0 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   Else
      If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
         MsgBox Label2 & MsgText(63), , MsgText(5)
         Cancel = True
         MaskEdBox1.SetFocus
         Exit Sub
      End If
      AdodcRefresh
   End If
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Text7 = A0g02Query(Text1)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If ExistCheck("acc0g0", "a0g01", Text1, Label9) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2004/11/1 加 and rownum<1
   adoadodc1.Open "select * from acc0e0 where a0e01 = '" & Text1 & "' and a0e22 = " & Val(FCDate(MaskEdBox1.Text)) & " and rownum<1  order by a0e22 asc, a0e41 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
   'Add by Morgan 2004/11/1
   If Text1 = "" Or Text2 = "" Or MaskEdBox1.Text = "___/__/__" Then Exit Sub
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   '2011/4/11 modify by sonia 改變insert後的順序
   'adoadodc1.Open "select * from acc0e0, acc0g0 where a0e01 = a0g01 and a0e01 = '" & Text1 & "' and a0e07 = '" & Text2 & "' and a0e22 = " & Val(FCDate(MaskEdBox1.Text)) & " order by a0e22 asc, a0e41 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2011/5/5 modify by sonia再改一次,瑞婷希望要依輸入順序由後至前排列
   'adoadodc1.Open "select * from acc0e0, acc0g0 where a0e01 = a0g01 and a0e01 = '" & Text1 & "' and a0e07 = '" & Text2 & "' and a0e22 = " & Val(FCDate(MaskEdBox1.Text)) & " order by a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select * from acc0e0, acc0g0 where a0e01 = a0g01 and a0e01 = '" & Text1 & "' and a0e07 = '" & Text2 & "' and a0e22 = " & Val(FCDate(MaskEdBox1.Text)) & " order by a0e22 asc,a0e02 desc, a0e41 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   '2011/4/11 add by sonia 將剛insert那一筆放在grid的第一筆
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a0e02 = '" & Text4 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Adodc1.Recordset.MoveFirst
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   End If
   '20114/11 end
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(票據資料)
'
'*************************************************
Private Sub Acc0e0Save()
On Error GoTo Checking
   If Text4 = MsgText(601) Then
      MsgBox MsgText(10) & Label4, , MsgText(5)
      strControlButton = MsgText(602)
      Text4.SetFocus
      Exit Sub
   Else
      If Text2 = MsgText(601) Then
         MsgBox Label1 & MsgText(52), , MsgText(5)
         strControlButton = MsgText(602)
         Text2.SetFocus
         Exit Sub
      End If
      If Text1 = MsgText(601) Then
         MsgBox Label9 & MsgText(52), , MsgText(5)
         strControlButton = MsgText(602)
         Text1.SetFocus
         Exit Sub
      End If
      If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
         MsgBox Label2 & MsgText(52), , MsgText(5)
         strControlButton = MsgText(602)
         MaskEdBox1.SetFocus
         Exit Sub
      Else
         If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
            MsgBox Label2 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Sub
         End If
      End If

      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0h01, a0h02 from acc0h0 where a0h02 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label1
         strControlButton = MsgText(602)
         adocheck.Close
         Text2.SetFocus
         Exit Sub
      End If
      adocheck.Close
   End If
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Morgan 2004/11/1 加應付控制 and a0e04='P'
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text1 & "' and a0e25 = 0 and a0e10 <= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e02 = '" & Text4 & "' and (a0e22 = 0 or a0e22 is null)  and a0e04='P'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0e0.RecordCount = 0 Then
      MsgBox MsgText(167), , MsgText(5)
      adoacc0e0.Close
      Exit Sub
   End If
   If Text1 <> MsgText(601) Then
      adoacc0e0.Fields("a0e19").Value = Text1
   Else
      adoacc0e0.Fields("a0e19").Value = Null
   End If
   If Text2 <> MsgText(601) Then
      adoacc0e0.Fields("a0e20").Value = Text2
   Else
      adoacc0e0.Fields("a0e20").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adoacc0e0.Fields("a0e22").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc0e0.Fields("a0e22").Value = 0
   End If
   adoacc0e0.Fields("a0e41").Value = ServerTime
   adoacc0e0.Fields("a0e29").Value = Val(strSrvDate(2))
   adoacc0e0.Fields("a0e30").Value = ServerTime
   adoacc0e0.Fields("a0e31").Value = strUserNum
   adoacc0e0.UpdateBatch
   AdodcRefresh
   adoacc0e0.Close
   'Add By Sindy 2011/2/11
   Text4 = Right("0000000" & Val(Text4) + 1, 7)
   '2011/2/11 End
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Modified by Morgan 2025/11/76 +開票銀行及票號可能會重複，還要加開票銀行帳號a0e07 Ex:0006458
   adoTaie.Execute "update acc0e0 set a0e22 = 0 where a0e01 = '" & Adodc1.Recordset.Fields("a0e01").Value & "' and a0e02 = '" & Adodc1.Recordset.Fields("a0e02").Value & "' and a0e07='" & Adodc1.Recordset.Fields("a0e07").Value & "'"
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   'Mark by Amy 2023/08/07 系統會自動帶下一票據號碼,有時需按二次[Insert]鍵才會有動作(目前測只會觸發一次)
'   'Add by Amy 2021/10/20
'   Call PUB_SaveTrackMode(1, KeyCode)
'    'Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
'    If PUB_ChkTrackMode = False Then
'        Exit Sub
'    End If
 
   Select Case KeyCode
      Case vbKeyF12
         AdodcRefresh
      Case vbKeyInsert
         Acc0e0Save
         Text4_GotFocus 'Add By Sindy 2011/2/11
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0h01, a0h02 from acc0h0 where a0h02 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) = False Then
            Text1 = adocheck.Fields(0).Value
            adocheck.Close
            Exit Sub
         End If
      End If
      MessageShow Label1
      adocheck.Close
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  筆數及金額合計
'
'*************************************************
Private Sub SumShow()
Dim adoaccsum As New ADODB.Recordset

   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e01 = '" & Text1 & "' and a0e07 = '" & Text2 & "' and a0e22 = " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = Format(adoaccsum.Fields(1).Value, DDollar)
      End If
   Else
      Text8 = MsgText(601)
      Text3 = MsgText(601)
   End If
   adoaccsum.Close
End Sub
