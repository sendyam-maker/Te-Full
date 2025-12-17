VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc12b0 
   AutoRedraw      =   -1  'True
   Caption         =   "翻譯費查詢"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4618.544
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   5580
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      MaxLength       =   1
      TabIndex        =   2
      Top             =   450
      Width           =   300
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      MaxLength       =   6
      TabIndex        =   3
      Top             =   810
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc12b0.frx":0000
      Height          =   3285
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   5794
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "C01"
         Caption         =   "身分"
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
         DataField       =   "C02"
         Caption         =   "姓名"
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
         DataField       =   "C03"
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
      BeginProperty Column03 
         DataField       =   "C04"
         Caption         =   "入帳日期"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1154.835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   1920
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1215
      TabIndex        =   0
      Top             =   105
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
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
      Height          =   300
      Left            =   3120
      TabIndex        =   1
      Top             =   105
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
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
   Begin MSForms.Label lblStaffName 
      Height          =   255
      Left            =   2970
      TabIndex        =   8
      Top             =   840
      Width           =   1740
      VariousPropertyBits=   19
      Caption         =   "lblStaffName"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "身分              (1:內翻  2:外翻  空白:全部)"
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
      Left            =   180
      TabIndex        =   7
      Top             =   480
      Width           =   4215
   End
   Begin VB.Line Line2 
      X1              =   2835
      X2              =   3065
      Y1              =   248.729
      Y2              =   248.729
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "員工代號"
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
      Left            =   180
      TabIndex        =   6
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      Left            =   180
      TabIndex        =   5
      Top             =   150
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc12b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 lblStaffName/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Add by Morgan 2007/6/1
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck = True Then
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
            Screen.MousePointer = vbHourglass
            AdodcRefresh
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
      Case Else
         KeyEnter KeyCode
   End Select
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Function FormCheck() As Boolean
   If MaskEdBox1.Text = MsgText(29) Then
      MsgBox "入帳日期不可空白！"
      MaskEdBox1.SetFocus
      Exit Function
   ElseIf MaskEdBox2.Text = MsgText(29) Then
      MsgBox "入帳日期不可空白！"
      MaskEdBox2.SetFocus
      Exit Function
   Else
      FormCheck = True
   End If
End Function

Private Sub AdodcRefresh()
   Dim strCon As String, strCon1 As String
   strCon = ""
   '入帳日期
   If MaskEdBox1.Text <> MsgText(29) Then
      strCon = strCon & " and a1p18>=" & Val(FCDate(MaskEdBox1.Text))
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      strCon = strCon & " and a1p18<=" & Val(FCDate(MaskEdBox2.Text))
   End If
   '員工代號
   If Text2 <> "" Then
      strCon = strCon & " and a1p15='" & Text2 & "'"
   End If
   '內翻
   If Text4 = "1" Then
      strCon1 = strCon1 & " and s2.st04='1'"
   '外翻
   ElseIf Text4 = "2" Then
      strCon1 = strCon1 & " and nvl(s2.st04,'2')='2'"
   End If
   
   strExc(0) = "select decode(s2.st04,'1','內翻','外翻') C01,a1p15||' '||s1.st02 C02, pay C03,a1p18 C04 from (" & _
      " select a1p15,a1p18,sum(a1p07) pay from acc1p0 where a1p05='6130'" & strCon & _
      " group by a1p15,a1p18),staff s1,staff_idmap,staff s2" & _
      " where s1.st01(+)=a1p15 and sim02(+)=a1p15 and s2.st01(+)=sim01 and s2.st04(+)='1'" & strCon1 & " order by 1,2"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      DataGrid1.Enabled = True
   Else
      DataGrid1.Enabled = False
      MsgBox "查無資料！"
   End If
   'Modify by Amy 2014/06/26 改不用離線資料集，避免資料多時新增至暫存檔慢
   'Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp)
   Set Adodc1.Recordset = RsTemp
   
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 5700, 5010
   '畫面初值設定
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text2 = ""
   Text4 = ""
   lblStaffName = ""
   DataGrid1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc12b0 = Nothing
End Sub

Private Sub MaskEdBox1_GotFocus()
   If MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox1.SelStart = 0
      MaskEdBox1.SelLength = MaskEdBox1.MaxLength
   End If
End Sub

Private Sub MaskEdBox2_GotFocus()
   If MaskEdBox2.Text = MsgText(29) And MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox2 = MaskEdBox1
      MaskEdBox2.SelStart = 0
      MaskEdBox2.SelLength = MaskEdBox2.MaxLength
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = "" Then
      lblStaffName = ""
   Else
      lblStaffName = GetStaffName(Text2, True)
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

