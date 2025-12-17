VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11m0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "拆收據作業"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8640
   Begin VB.TextBox Text11 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1035
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
      Height          =   330
      Left            =   1305
      MaxLength       =   15
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2505
      Picture         =   "Frmacc11m0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "拆收據"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7245
      TabIndex        =   2
      Top             =   240
      Width           =   1185
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   372
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1800
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1800
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11m0.frx":0102
      Height          =   2385
      Left            =   120
      TabIndex        =   15
      Top             =   2220
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   4207
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "a0j02"
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
      BeginProperty Column01 
         DataField       =   "a0j07"
         Caption         =   "合併"
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
         DataField       =   "cp10N"
         Caption         =   "案件性質"
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
         DataField       =   "na03"
         Caption         =   "申請國家"
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
         DataField       =   "a0j09"
         Caption         =   "服務費"
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
      BeginProperty Column05 
         DataField       =   "a0j10"
         Caption         =   "規費"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1154.835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   180
      Top             =   2070
      Visible         =   0   'False
      Width           =   975
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   1500
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   14737632
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
   Begin MSForms.TextBox Text10 
      Height          =   330
      Left            =   5310
      TabIndex        =   6
      Top             =   720
      Width           =   1575
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   30
      Size            =   "2778;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   5850
      TabIndex        =   9
      Top             =   1080
      Width           =   2535
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "4471;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   330
      Left            =   1500
      TabIndex        =   10
      Top             =   1440
      Width           =   4335
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   30
      Size            =   "7646;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text12 
      Height          =   330
      Left            =   3615
      TabIndex        =   3
      Top             =   210
      Width           =   3450
      VariousPropertyBits=   -1466941409
      BackColor       =   14737632
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6085;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
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
      Left            =   300
      TabIndex        =   26
      Top             =   270
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   3300
      TabIndex        =   25
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Left            =   300
      TabIndex        =   24
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
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
      Left            =   300
      TabIndex        =   23
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1.不可扣繳 2.可扣繳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6420
      TabIndex        =   22
      Top             =   1500
      Width           =   1755
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "規費合計"
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
      Left            =   3300
      TabIndex        =   21
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "服務費合計"
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
      Left            =   300
      TabIndex        =   20
      Top             =   1830
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "總計"
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
      Left            =   6060
      TabIndex        =   19
      Top             =   1830
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   180
      Top             =   150
      Width           =   6990
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   300
      TabIndex        =   18
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   3300
      TabIndex        =   17
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   3060
      TabIndex        =   16
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "Frmacc11m0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
Option Explicit


Private Sub ReadData()
   
   strExc(0) = "select * from acc0k0 where a0k01 = '" & Text1 & "' and nvl(a0k09,0)=0 order by a0k01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "收據編號輸入錯誤!!"
      Text1.SetFocus
      Exit Sub
   End If
   
   Text1.Tag = Text1.Text
   Command1.Enabled = True
   
   With RsTemp
   If IsNull(.Fields("a0k08").Value) Then
      Text12 = MsgText(601)
   Else
      Text12 = .Fields("a0k08").Value
   End If
   If IsNull(.Fields("a0k11").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = .Fields("a0k11").Value
   End If
   If IsNull(.Fields("a0k20").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = .Fields("a0k20").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(.Fields("a0k02").Value) Or .Fields("a0k02").Value = 0 Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(.Fields("a0k02").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(.Fields("a0k03").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = .Fields("a0k03").Value
   End If
   If IsNull(.Fields("a0k04").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = .Fields("a0k04").Value
   End If
   If IsNull(.Fields("a0k05").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = .Fields("a0k05").Value
   End If
   If IsNull(.Fields("a0k07").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = .Fields("a0k07").Value
   End If
   If IsNull(.Fields("a0k06").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = .Fields("a0k06").Value
   End If
   Text8 = Val(Text4) + Val(Text7)
   End With
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   strExc(0) = "select a.*,getcp10desc(cp01,cp10,a0j04) cp10N,na03 from acc0j0 a,caseprogress,nation where a0j13 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 order by a0j01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set Adodc1.Recordset = RsTemp.Clone
End Sub

Private Sub Command1_Click()
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   If CheckData = True Then
      Me.Enabled = False
      Set Frmacc1125.frmCallForm = Me
      Frmacc1125.m_OldNo = Me.Text1
      Frmacc1125.Caption = "拆收據作業"
      Frmacc1125.Show
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
End Sub

Private Sub Command2_Click()
   ReadData
End Sub

Private Sub Form_Activate()
   If Text1 <> "" Then
      strFormName = Me.Name
      ReadData
   End If
   tool3_enabled
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   AdodcClear
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11m0 = Nothing
End Sub

Private Sub Text1_Change()
   Command1.Enabled = False
   If Text1.Tag <> "" Then
      AdodcClear
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_Change()
   Text10 = StaffQuery(Text11)
End Sub

Private Sub Text2_Change()
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Text3 = CustomerQuery(Text2, 1)
End Sub

'*************************************************
'  清除查詢資料
'
'*************************************************
Private Sub AdodcClear()
   Text12 = ""
   Text9 = ""
   Text11 = ""
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text2 = ""
   Text3 = ""
   Text6 = ""
   Text5 = ""
   Text4 = ""
   Text7 = ""
   Text8 = ""
   Text1.Tag = ""
End Sub

Private Function CheckData() As Boolean
   strExc(0) = "select * from acc1u0 where a1u02='" & Text1 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "收據號碼【" & Text1 & "】已收款或已銷帳，不可拆收據！", vbExclamation
      Exit Function
   End If
   
   'Add By Sindy 2013/12/31
   strExc(0) = "select * from acc430,acc431 where axc02='" & Text1 & "' and axc01=a4301"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "J 公司請款單【" & Text1 & "】已開立發票，不可拆請款單！", vbExclamation
      Exit Function
   End If
   '2013/12/31 END
   
   CheckData = True
End Function
