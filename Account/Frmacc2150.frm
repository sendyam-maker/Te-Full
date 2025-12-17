VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2150 
   AutoRedraw      =   -1  'True
   Caption         =   "帳單輸入"
   ClientHeight    =   5364
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5364
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H008080FF&
      Caption         =   "？"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3924
      Style           =   1  '圖片外觀
      TabIndex        =   40
      Top             =   2265
      Width           =   285
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   60
      Top             =   3420
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2150.frx":0000
      Height          =   1635
      Left            =   210
      TabIndex        =   16
      Top             =   2625
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2879
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
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
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "axf03"
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
         DataField       =   "axf02"
         Caption         =   "總收文號"
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
         DataField       =   "cpm03"
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
         DataField       =   "axf04"
         Caption         =   "帳單金額"
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
         DataField       =   "axf14"
         Caption         =   "案件盈虧"
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
      BeginProperty Column05 
         DataField       =   "axf12"
         Caption         =   "案件名稱"
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
         DataField       =   "axf13"
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
      BeginProperty Column07 
         DataField       =   "axf16"
         Caption         =   "是否含公開費"
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
         Size            =   284
         BeginProperty Column00 
            ColumnWidth     =   1404.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1284.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1332.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1404.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3636.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4356.284
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1451.906
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2310
      Style           =   2  '單純下拉式
      TabIndex        =   37
      Top             =   2263
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5850
      MaxLength       =   6
      TabIndex        =   36
      Top             =   2220
      Width           =   1170
   End
   Begin VB.TextBox Text15 
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
      Height          =   330
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2220
      Width           =   1428
   End
   Begin VB.CheckBox Check3 
      Caption         =   "急件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   2280
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "獨立水單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4005
      TabIndex        =   33
      Top             =   60
      Width           =   1140
   End
   Begin VB.CheckBox Check1 
      Caption         =   "紙本會簽"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4005
      TabIndex        =   32
      Top             =   360
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "電子檔"
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
      Left            =   3015
      TabIndex        =   31
      Top             =   210
      Width           =   855
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3165
      TabIndex        =   13
      Top             =   4845
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2820
      TabIndex        =   12
      Top             =   4845
      Width           =   348
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2565
      TabIndex        =   11
      Top             =   4845
      Width           =   240
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   10
      Top             =   4845
      Width           =   780
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   6
      Top             =   1410
      Width           =   1572
   End
   Begin VB.TextBox Text14 
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
      Height          =   330
      Left            =   4485
      TabIndex        =   28
      Top             =   4335
      Width           =   1428
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2670
      Picture         =   "Frmacc2150.frx":0015
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   195
      Width           =   350
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   6840
      TabIndex        =   27
      Top             =   1425
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6840
      Picture         =   "Frmacc2150.frx":0117
      Style           =   1  '圖片外觀
      TabIndex        =   15
      ToolTipText     =   "取消"
      Top             =   4305
      Width           =   450
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6780
      MaxLength       =   14
      TabIndex        =   14
      Top             =   4860
      Width           =   1572
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   9
      Top             =   4845
      Width           =   480
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4164
      TabIndex        =   7
      Top             =   1410
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1035
      Width           =   3816
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   2
      Top             =   660
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Height          =   330
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   6840
      TabIndex        =   5
      Top             =   1035
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   593
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "付款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   1356
      TabIndex        =   39
      Top             =   2299
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "額外通知人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4470
      TabIndex        =   38
      Top             =   2280
      Width           =   1350
   End
   Begin MSForms.TextBox Text8 
      Height          =   405
      Left            =   1320
      TabIndex        =   8
      Top             =   1770
      Width           =   7095
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "12515;714"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   2910
      TabIndex        =   3
      Top             =   660
      Width           =   5520
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      MaxLength       =   50
      Size            =   "9737;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳單台幣金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5490
      TabIndex        =   30
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3930
      TabIndex        =   29
      Top             =   4380
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4980
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   510
      Left            =   225
      Top             =   4770
      Width           =   8295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "帳單金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5820
      TabIndex        =   26
      Top             =   4875
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   25
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "作廢日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   1455
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "帳單總金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3015
      TabIndex        =   22
      Top             =   1455
      Width           =   1185
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   1455
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "帳單日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人D/N No."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   345
      TabIndex        =   18
      Top             =   705
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "帳單編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   225
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text3、Text8
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc150 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public strDocNo As String
Public strYes As String
Public strFMP As Boolean   'add by sonia 2017/9/13
Public strCFPC As Boolean  'add by sonia 2017/10/11
Public m_a1512 As String   '2010/4/2 add by sonia
Dim RQstr As String 'Add by Lydia 2014/10/31 控制是否可看寰華案的SQL條件(先經過Form_Load預設,再依帳單的狀況修改為本程式的區域變數)
'Added by Morgan 2016/6/27 電子檔名,本所案號,母表單
Public m_eFileName As String, m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, m_ParentForm As Form
'Add By Sindy 2018/2/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_RDate As String
Dim m_Done As Boolean
'2018/2/22 END
Public m_Comp As String '帳單公司別 Added by Morgan 2023/4/7

Private Sub Check3_Click()
   If Check3.Value = vbChecked Then
      Combo2.Enabled = True
      Text13.Enabled = True
      SetPayDay
   Else
      Combo2.Clear
      Combo2.Tag = ""
      Combo2.ListIndex = -1
      Combo2.Enabled = False
      Text13 = "": Text15 = ""
      Text13.Enabled = False
   End If
End Sub

'Added by Morgan 2023/6/12 --斯閔
Private Sub cmdHelp_Click()
   strExc(0) = "USD: 本月第一週" & vbCrLf & _
         "EUR等多幣: 本月第二週" & vbCrLf & _
         "RMB: 本月第三週" & vbCrLf & _
         "智權公司: 本月第四週"
   MsgBox strExc(0), vbInformation, "幣別與付款規則"
End Sub

Private Sub Combo1_Click()
   'Modified by Morgan 2024/12/27
   'If adoacc150.State <> adStateOpen Then
   '   Exit Sub
   'End If
   'If adoacc150.RecordCount = 0 Then
   '   Exit Sub
   'End If
   'If Combo1 <> adoacc150.Fields("a1505").Value Then
   '   If Adodc1.Recordset.RecordCount <> 0 Then
   '      MsgBox MsgText(206), , MsgText(5)
   '   End If
   'End If
   CurCheck
   'end 2024/12/27
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo1, Label5) = False Then
      Cancel = True
      Combo1.SetFocus
   End If
   
   'Added by Morgan 2024/12/27
   If Cancel = False Then
      CurCheck
   End If
   'end 2024/12/27
End Sub

Private Sub Command1_Click()
   AdodcDelete
   AdodcClear
End Sub

Private Sub Command2_Click()
   If Text2 <> "" Then
      strExc(0) = "select a1501 from acc150 where a1501='" & Text2 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strItemNo = Text2
         Frmacc2154.Show vbModal
         strItemNo = ""
         strFormName = Me.Name
      Else
         MsgBox "帳單不存在！", vbCritical
      End If
   Else
      MsgBox "請先輸入帳單編號！", vbExclamation
   End If
End Sub

Private Sub Command3_Click()
   Acc150Refresh
   AdodcClear
   If adoacc150.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   Else
      'Add by Lydia 2014/10/31 提示訊息
      'Modified by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入
      'If FMP2open = True Then
      If FMP2openSQL <> "" And Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M31" Then
        MsgBox "權限不足或查無符合資料 !", vbInformation
      Else
        MsgBox MsgText(33), , MsgText(5)
      End If
   End If
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcShow
End Sub

Private Sub Form_Activate()
Static bolActivated As Boolean 'Added by Morgan 2016/4/20
   
   Dim Cancel As Boolean  '2005/4/14 ADD BY SONIA
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
   strFormName = Name
   
   'Added by Morgan 2016/6/27
   If Not bolActivated Then
      strCon9 = "" 'Added by Morgan 2021/2/23 若有殘留會在下面當成語法執行，導致錯誤
      If m_eFileName <> "" Then
         KeyDefine vbKeyF2 '新增
         Text10 = m_CP01
         'Modify By Sindy 2021/1/21
         If m_CP01 = "TF" Then
            Call Text10_Validate(False)
            Text5 = Left(m_CP02, 5)
            Text7 = Right(m_CP02, 1)
            Text9 = m_CP03
            Text12 = m_CP04
         Else
         '2021/1/21 END
            Text5 = m_CP02
            Text7 = m_CP03
            Text9 = m_CP04
         End If
         'Added by Morgan 2019/11/14
         'P案預設人民幣(RMB)，但若該案最後發文代理人為 Y53374, Y20821則仍預設美金(USD)--玲玲
         If Text10 = "P" And Pub_StrUserSt03 <> "F22" Then
            strExc(0) = "select cp44 from caseprogress where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' and cp27>19221111 and cp44 is not null order by cp27 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modified by Morgan 2019/11/15 +Y52404 唯源--玲玲
               'Modified by Morgan 2020/3/19 改抓代理人檔設定,不例外
               'If RsTemp(0) = "Y53374000" Or RsTemp(0) = "Y20821000" Or RsTemp(0) = "Y52404000" Then
               '   Combo1 = "USD"
               'End If
               Text1 = "" & RsTemp(0)
               Text1_Validate False
               'end 2020/3/19
            End If
            Text4.SetFocus
            If Combo1 = "" Then Combo1 = "RMB"
            Combo1.Tag = Combo1 'Added by Morgan 2024/12/27
         End If
         'end 2019/11/14
         strSql = "update acc152 set ayf01='" & Text2 & "' where ayf01='U' and ayf02='" & ChgSQL(m_eFileName) & "'"
         adoTaie.Execute strSql, intI
         m_eFileName = ""
      End If
      bolActivated = True
   End If
   'end 2016/6/27
   
   'Added by Sindy 2018/2/22
   If m_strIR01 <> "" And m_Done = False Then
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "＜" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "＞）"
   End If
   '2018/2/22 END
   
   Label8 = ""  '2010/4/13 ADD BY SONIA
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      If strItemNo = MsgText(601) Then
         Exit Sub
      End If
      If adoacc150.RecordCount <> 0 Then
         adoacc150.MoveFirst
      End If
      Text2 = strItemNo
      Acc150Refresh
      If adoacc150.RecordCount <> 0 Then
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
      strItemNo = MsgText(601)
      Exit Sub
   End If
   If strCon9 <> "" Then
      If strControlButton <> MsgText(602) Then
         adoTaie.Execute strCon9
         adoTaie.Execute strCon10
         '2015/7/22 ADD BY SONIA 做完即清除,否則開其他畫面時會觸發此FORM_ACTIVATE,即會出現違反唯一的限制條件
         strCon9 = ""
         strCon10 = ""
         '2015/7/22 END
      End If
      If strControlButton <> MsgText(602) Then
         AdodcRefresh
         AdodcClear
         Text10.SetFocus
      End If
      If strCustNo <> MsgText(601) Then
         Text1 = strCustNo
         Text3 = FagentQuery(Text1, 2)
         '2005/4/14 ADD BY SONIA
         Cancel = False
         Text1_Validate Cancel
         '2005/4/14 END
      End If
'      Frmacc2150_Save
      strControlButton = MsgText(601)
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Do While Adodc1.Recordset.EOF = False
      adoquery.CursorLocation = adUseClient
      '2007/3/2 modify by sonia 加入cp87,cp88
      'adoquery.Open "select cp61, cp62, cp63, a0k04 from caseprogress, acc0k0 where cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("axf02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select cp61, cp62, cp63, cp87, cp88, a0k04 from caseprogress, acc0k0 where cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("axf02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      '2007/3/2 end
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            adoTaie.Execute "update caseprogress set cp61 = '" & Text2 & "' where cp09 = '" & Adodc1.Recordset.Fields("axf02").Value & "'"
         Else
            If IsNull(adoquery.Fields(1).Value) And adoquery.Fields(0).Value <> Text2 Then
               adoTaie.Execute "update caseprogress set cp62 = '" & Text2 & "' where cp09 = '" & Adodc1.Recordset.Fields("axf02").Value & "'"
            Else
               If IsNull(adoquery.Fields(2).Value) And adoquery.Fields(0).Value <> Text2 And adoquery.Fields(1).Value <> Text2 Then
                  adoTaie.Execute "update caseprogress set cp63 = '" & Text2 & "' where cp09 = '" & Adodc1.Recordset.Fields("axf02").Value & "'"
               '2007/3/2 add by sonia 加入cp87,cp88
               Else
                  If IsNull(adoquery.Fields(2).Value) And adoquery.Fields(0).Value <> Text2 And adoquery.Fields(1).Value <> Text2 Then
                     adoTaie.Execute "update caseprogress set cp87 = '" & Text2 & "' where cp09 = '" & Adodc1.Recordset.Fields("axf02").Value & "'"
                  Else
                     If IsNull(adoquery.Fields(2).Value) And adoquery.Fields(0).Value <> Text2 And adoquery.Fields(1).Value <> Text2 Then
                        adoTaie.Execute "update caseprogress set cp88 = '" & Text2 & "' where cp09 = '" & Adodc1.Recordset.Fields("axf02").Value & "'"
                     End If
                  End If
              '2007/3/2 end
               End If
            End If
         End If
         If IsNull(adoquery.Fields("a0k04").Value) Then
            adoTaie.Execute "update acc151 set axf13 = null where axf01 = '" & Text2 & "' and axf02 = '" & Adodc1.Recordset.Fields("axf02").Value & "'"
         Else
            'Modified by Morgan 2012/1/10 收據抬頭會有單引號
            adoTaie.Execute "update acc151 set axf13 = '" & ChgSQL("" & adoquery.Fields("a0k04").Value) & "' where axf01 = '" & Text2 & "' and axf02 = '" & Adodc1.Recordset.Fields("axf02").Value & "'"
         End If
      End If
      adoquery.Close
      'modify by sonia 2017/9/13 FMP案不判斷虧損(U10607678)
      'modify by sonia 2017/10/11 CFP之C類不檢查收文號及案號之虧損
      If Val("" & Adodc1.Recordset.Fields("axf14").Value) < 0 And Not strFMP And Not strCFPC Then
         strYes = MsgText(603) 'N
      'Modify By Sindy 2021/1/21 + 林經理提,TF都要審核
      ElseIf m_CP01 = "TF" Then
         strYes = "W"
      '2021/1/21 END
      End If
      Adodc1.Recordset.MoveNext
   Loop
   
   'Added by Morgan 2023/4/21
   m_Comp = GetComp()
   If m_Comp <> Combo2.Tag And (m_Comp = "J" Or Combo2.Tag = "J") Then
      If Combo2.ListIndex >= 0 Then
         MsgBox "目前設定的急件付款日期並非帳單公司別的預定付款日，請重新設定！", vbInformation
      End If
      SetPayDay
   End If
   'end 2023/4/21
End Sub

'Added by Lydia 2021/12/07
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
      
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   Me.Height = 5500
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 8850, 5800, strBackPicPath1
   'end 2021/12/07
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   'Added by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入
   If Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M31" Then
        FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05, "INVERSE_SQL")
        '內專才控制回傳
        If UCase(App.EXEName) <> "PATPRO" And UCase(App.EXEName) <> "TEPATPRO" Then
            FMP2openSQL = ""
        'Added by Lydia 2019/10/14 CFP案也會有寰華帳單但不必檢查 Ex:CFP-31173; CFP程序有時會支援P程序
        ElseIf InStr("83,85", Pub_strUserST05) > 0 Then
            'Modified by Morgan 2020/10/7 +CPS
            FMP2openSQL = "and ( f0.cp01='CFP' or f0.cp01='CPS' or (f0.cp01='P' " & FMP2openSQL & ") ) "
        End If
   Else
   'end 2019/09/10
        FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   End If
   
   OpenTable
   If adoacc150.RecordCount <> 0 Then
      adoacc150.MoveLast
      adoacc150.MoveFirst
      RecordShow
   End If
   FormDisabled
   
   'Added by Morgan 2016/6/28
   'Add By Sindy 2021/1/18 + or (T商標電子化第2階段啟用日 <= Val(strSrvDate(1)) And Pub_StrUserSt03 = "P22")
   'Modified by Morgan 2023/4/19 +F11外商承辦
   'If (內專全面電子化啟用日 <= Val(strSrvDate(1)) And _
         (Pub_StrUserSt03 = "P12" Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31")) Or _
      (T商標電子化第2階段啟用日 <= Val(strSrvDate(1)) And Pub_StrUserSt03 = "P22") Then
    'Removed by Morgan 2025/8/13 上傳不必再限制
    'If InStr("M51,M31,P12,P22,F11,F22", Pub_StrUserSt03) > 0 Then
    'end 2025/8/13
    'end 2023/4/19
      Command2.Visible = True
      'Modified by Morgan 2023/4/19
      ''Add By Sindy 2021/1/18
      'If Pub_StrUserSt03 = "P22" Then '內商
      '   Check1.Visible = False
      '   Check2.Visible = False
      'Else
      ''2021/1/18 END
      '   Check1.Visible = True 'Added by Morgan 2019/3/12
      '   Check2.Visible = True 'Added by Morgan 2019/3/15
      'End If
      If InStr("M51,M31,P12", Pub_StrUserSt03) > 0 Then
         Check1.Visible = True
         Check2.Visible = True
      Else
         Check1.Visible = False
         Check2.Visible = False
      End If
      'end 2023/4/19
   'Removed by Morgan 2025/8/13 上傳不必再限制
   'Else
   '   Command2.Visible = False
   '   Check1.Visible = False 'Added by Morgan 2019/3/12
   '   Check2.Visible = False 'Added by Morgan 2019/3/15
   'End If
   'end 2025/8/13
   'end 2016/6/28
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2019/8/2
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   
   'Added by Morgan 2016/6/27
   If Not m_ParentForm Is Nothing Then
      'Added by Morgan 2018/2/6
      '回待處理區
      If UCase(m_ParentForm.Name) = UCase("frm210149") Then
         m_ParentForm.Show
         m_ParentForm.PubShowNextData
      'Add By Sindy 2018/2/23
      ElseIf m_strIR01 <> "" Then
         If UCase(m_ParentForm.Name) = UCase("frm04010519") Then
            If Not m_ParentForm Is Nothing Then
               Call m_ParentForm.GoNext
            End If
         '整批帳單
         Else
            m_ParentForm.Enabled = True
            tool3_enabled
            Unload m_ParentForm
         End If
         If Not m_ParentForm Is Nothing Then
            Set m_ParentForm = Nothing
         End If
         '2018/2/23 END
      Else
      'end 2018/2/6
         m_ParentForm.Enabled = True
         m_ParentForm.Show
         tool3_enabled
      End If 'Added by Morgan 2018/2/6
   End If
   'end 2016/6/27
   
   Set Frmacc2150 = Nothing
End Sub

Private Sub MaskEdBox1_GotFocus()
   CloseIme
   If Mid(MaskEdBox1, 5, 2) = "__" Then
      MaskEdBox1.SelStart = 4
      MaskEdBox1.SelLength = 0
   End If
End Sub

Private Sub MaskEdBox1_LostFocus()
    'Add By Cheng 2003/12/25
    '若為新增狀態
    If strSaveConfirm = MsgText(3) Then
    
'Mem by Morgan 2009/10/2 移到存檔時才檢查

'        '檢查帳單資料是否重覆
'        'Modify By Sindy 2009/06/17 若為專利處只須以代理人+代理人D/N No.做重覆檢核
'        If Left(Trim(GetStaffDepartment(strUserNum)), 2) <> "P1" Then
'            'If PUB_ChkDNDup(Me.MaskEdBox1.Text, Text1.Text, Text4.Text, "", , 0) = True Then
'            If PUB_ChkDNDup(Me.MaskEdBox1.Text, Text1.Text, Text4.Text, Text2.Text, , 0) = True Then
'               Text4.SetFocus
'               Text4_GotFocus
'               Exit Sub
'            End If
'        End If
'        If ChkDataRepaet(Me.MaskEdBox1.Text, Me.Text1.Text, Me.Text4.Text) = True Then
'            'Modify By Cheng 2004/02/12
'            'Focus設在代理人D/N No.欄位
''            Me.Text1.SetFocus
''            Text1_GotFocus
'            Me.Text4.SetFocus
'            Text4_GotFocus
'            'End
'            Exit Sub
'        End If
    End If
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label4 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label4 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   'Ken 92/01/03 改為編號依系統年度編號
'   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
'      Text2 = UpdateNo("acc150", "a1501", 5, MaskEdBox1.Text, Mid(Text2, 1, 1))
'   Else
      'Text2 = AutoNo(MsgText(812), 5)
'      Text2 = strDocNo
'   End If
End Sub

Private Sub Text1_Change()
   Dim strMsg As String
   If Len(Text1) = 9 And (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) Then
      'Modify by Morgan 2008/6/6 有代理人帳單備註提醒
      If Left(Pub_StrUserSt03, 2) = "P1" Then
         strExc(0) = "select nvl(fa05,fa04),fa92 from fagent where fa01='" & Left(Text1, 8) & "' and fa02='" & Mid(Text1, 9) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Text3 = "" & RsTemp(0)
            If Not IsNull(RsTemp(1)) Then
               MsgBox "" & RsTemp(1), vbExclamation, "代理人帳單備註"
            End If
         End If
      End If
   End If
End Sub

Private Sub Text1_GotFocus()
   CloseIme
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text1)
      Case 6
        Text1 = AfterZero(Text1)
      Case 8
        Text1 = Text1 & "0"
   End Select
   'Modified by Lydia 2019/09/10 新增或修改才檢查
   'If Text1 <> MsgText(601) Then
   If Text1 <> MsgText(601) And (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) Then
      'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
      If FMP2open = True Then
         If InStr(1, FMP2openSQL, Trim(Text1)) = 0 Then  '限定特定代理人
            MsgBox "權限不足 !", vbInformation
            Cancel = True
            Text1.SetFocus
            TextInverse Text1
            Exit Sub
         End If
      'Added by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入
      'Remove by Lydia 2019/10/14 改到按Insert再檢查
      'ElseIf Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M31" And FMP2openSQL <> "" Then
      '   If InStr(1, FMP2openSQL, Trim(Text1)) > 0 Then  '限定特定代理人(反向)
      '      MsgBox "此案為FCP自行連繫，請交FCP程序處理！", vbCritical, "寰華案控制"
      '      Cancel = True
      '      Text1.SetFocus
      '      TextInverse Text1
      '      Exit Sub
      '   End If
      'end 2019/09/10
      End If
      
      If ExistCheck("fagent", "fa01", Mid(Text1, 1, 8), Label2) = False Then
         Cancel = True
         Text1.SetFocus
         TextInverse Text1
         Exit Sub
      End If
   End If
   
   'Modified by Morgan 2019/9/18
   'Text3 = FagentQuery(Text1, 2)
   'If Text3 = MsgText(601) Then
   '   Text3 = FagentQuery(Text1, 1)
   'End If
   Text3 = ""
   If ClsPDGetAgent(Text1, strExc(0)) = True Then
      Text3 = strExc(0)
   End If
   'end 2019/9/18
   
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   '2005/4/28 MODIFY BY SONIA
   'adoquery.CursorLocation = adUseClient
   'adoquery.Open "select na52 from fagent, nation where fa10 = na01 (+) and fa01 = '" & Mid(Text1, 1, 8) & "' and fa02 = '" & Mid(Text1, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   'If adoquery.RecordCount <> 0 Then
   '   If IsNull(adoquery.Fields("na52").Value) = False Then
   '      Combo1 = adoquery.Fields("na52").Value
   '   End If
   'End If
   'adoquery.Close
   If Combo1 = "" Then
      'Add By Sindy 2012/6/5 改為先抓該代理人是否有設定帳單幣別,若有,則抓代理人的帳單幣別,若沒有才抓NA52
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select fa113 from fagent where fa01='" & Mid(Text1, 1, 8) & "' and fa02='" & Mid(Text1, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields("fa113").Value) = False Then
            Combo1 = adoquery.Fields("fa113").Value
            adoquery.Close
            Exit Sub
         End If
      End If
      adoquery.Close
      '2012/6/5 End
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select na52 from fagent, nation where fa10 = na01 (+) and fa01 = '" & Mid(Text1, 1, 8) & "' and fa02 = '" & Mid(Text1, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields("na52").Value) = False Then
            Combo1 = adoquery.Fields("na52").Value
         End If
      End If
      adoquery.Close
      
      Combo1.Tag = Combo1 'Added by Morgan 2024/12/27
   End If
   '2005/4/28 END
End Sub

Private Sub Text10_GotFocus()
   CloseIme
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 = "TF" Then
      Text12.Visible = True
      Text5.MaxLength = 5 '2010/7/2 ADD BY SONIA
   Else
      Text12.Visible = False
      Text5.MaxLength = 6 '2010/7/2 ADD BY SONIA
   End If
End Sub

Private Sub Text11_GotFocus()
   CloseIme
   TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_GotFocus()
   CloseIme
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   If Text13 <> "" Then
      Text15 = GetStaffName(Text13)
      If Text15 = "" Then
         MsgBox Label14 & "輸入錯誤！", vbCritical
         Cancel = True
      End If
   Else
      Text15 = ""
   End If
End Sub

Private Sub Text2_GotFocus()
   CloseIme
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'Modified by Lydia 2021/12/07 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Text3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   CloseIme
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   
   adoacc150.CursorLocation = adUseClient
   adoacc150.MaxRecords = intMax
   'adoacc150.Open "select * from acc150 where a1501 >= '" & Text2 & "' order by a1501 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim midSql As String
   midSql = " select m0.* from acc150 m0 where a1501>='" & Text2 & "' "
   'Modified by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入
   'If FMP2open = True Then
   If FMP2openSQL <> "" And Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M31" Then
      RQstr = " select m1.axf01 from acc151 m1,caseprogress f0 where m0.a1501=m1.axf01(+) and m1.axf02=f0.cp09(+) " & FMP2openSQL
      midSql = midSql & " and a1501 in (" & RQstr & ") "
      'Added by Lydia 2019/12/10 考慮尚未輸入明細的狀況;
      midSql = midSql & "union all select m0.* from acc150 m0 where a1501>='" & Text2 & "' and a1501 not in (select axf01 from acc151 where axf01>='" & Text2 & "' )"
   End If
   midSql = midSql & " order by 1 asc "
   adoacc150.Open midSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc151 where axf01 = '" & Text2 & "' order by axf02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo1.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   '2005/4/14 MODIFY BY SONIA
   'Combo1 = "USD"
   Combo1 = ""
   Combo1.Tag = Combo1 'Added by Morgan 2024/12/27
   '2005/4/14 END
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2008/2/14 +axf16
   '2009/9/11 modify by sonia 帳單申請國家一定非台灣故不抓CPM03改抓CPM04,U09807576大陸復審答辯406會出現參加訴願
   'adoadodc1.Open "select axf01, axf02, axf03, axf04, axf05, axf06, axf07, axf08, axf09, axf10, axf11, axf12, axf13, axf14, axf15, decode(cpm03, '（無）', cpm04, cpm03) as cpm03, cp09,axf16 from acc151, caseprogress, casepropertymap where axf02 = cp09 and cp01 = cpm01 and cp10 = cpm02 and axf01 = '" & Text2 & "' order by axf02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2009/10/8 modify by sonia FCP之舜禹翻譯帳單的案件性質會出現（無）U09808512
   'adoadodc1.Open "select axf01, axf02, axf03, axf04, axf05, axf06, axf07, axf08, axf09, axf10, axf11, axf12, axf13, axf14, axf15, cpm04 as cpm03, cp09,axf16 from acc151, caseprogress, casepropertymap where axf02 = cp09 and cp01 = cpm01 and cp10 = cpm02 and axf01 = '" & Text2 & "' order by axf02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select axf01, axf02, axf03, axf04, axf05, axf06, axf07, axf08, axf09, axf10, axf11, axf12, axf13, axf14, axf15, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as cpm03, cp09,axf16 " & _
          "from acc151, caseprogress, casepropertymap where axf02 = cp09 and cp01 = cpm01 and cp10 = cpm02 and axf01 = '" & Text2 & "' order by axf02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.ReQuery
   SumShow
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text2 = adoacc150.Fields("a1501").Value
   If IsNull(adoacc150.Fields("a1503").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc150.Fields("a1503").Value
   End If
'   If Len(Text1) = 6 Then
'      Text3 = FagentQuery(AfterZero(Text1), 2)
'   Else
'      Text3 = FagentQuery(Text1, 2)
'   End If
'Add by Lydia 2014/11/13 改變讀取代理人名稱的方式
   If ClsPDGetAgent(Text1, strExc(0)) = True Then
      Text3 = strExc(0)
   Else
      Text3 = ""
   End If
   
   If IsNull(adoacc150.Fields("a1504").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc150.Fields("a1504").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc150.Fields("a1502").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc150.Fields("a1502").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc150.Fields("a1505").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = adoacc150.Fields("a1505").Value
   End If
   Combo1.Tag = Combo1 'Added by Morgan 2024/12/27
   '2010/4/2 add by sonia
   If IsNull(adoacc150.Fields("a1512").Value) Then
      m_a1512 = MsgText(601)
   Else
      m_a1512 = adoacc150.Fields("a1512").Value
   End If
   '2010/4/2 end
   If IsNull(adoacc150.Fields("a1506").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Format(adoacc150.Fields("a1506").Value, FAmount)
      Text6_LostFocus  '2010/4/2 ADD BY SONIA
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc150.Fields("a1507").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc150.Fields("a1507").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc150.Fields("a1509").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc150.Fields("a1509").Value
   End If
   '94.1.5 ADD BY SONIA
   If IsNull(adoacc150.Fields("a1521").Value) Then
      strYes = MsgText(601)
   Else
      strYes = adoacc150.Fields("a1521").Value
      If strYes = "R" Then strYes = "N" 'Added by Morgan 2018/2/6
   End If
   '94.1.5 END
   
   'Added by Morgan 2019/3/12
   If adoacc150.Fields("a1525").Value = "Y" Then
      Check1.Value = vbChecked
   Else
      Check1.Value = vbUnchecked
   End If
   'end 2019/3/12
   
   'Added by Morgan 2019/3/15
   If adoacc150.Fields("a1526").Value = "Y" Then
      Check2.Value = vbChecked
   Else
      Check2.Value = vbUnchecked
   End If
   'end 2019/3/15
   
   'Added by Morgan 2023/4/6
   m_Comp = ""
   If Not IsNull(adoacc150.Fields("a1527").Value) Then
      Check3.Value = vbChecked
      strExc(0) = CFDate(adoacc150.Fields("a1527").Value)
      For intI = 0 To Combo2.ListCount - 1
         If Combo2.List(intI) = strExc(0) Then
            Combo2.ListIndex = intI
            Exit For
         End If
      Next
      If intI = Combo2.ListCount Then
         Combo2.AddItem strExc(0), 0
         Combo2.ListIndex = 0
      End If
      Text13 = "" & adoacc150.Fields("a1528").Value
      Text15 = GetStaffName(Text13, True)
   Else
      Check3.Value = vbUnchecked
   End If
   'end 2023/4/6
   
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modified by Lydia 2019/10/02 已產生Acc170結匯資料,就不可修改: ex.U10808026在9/25已產生結匯,在9/26早上先修改帳單金額-100後,財務才執行結匯資料產生付款單
   'adoquery.Open "select * from acc170 where a1702 = '" & Text2 & "' and a1709 is not null", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select * from acc170 where a1702 = '" & Text2 & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      tool15_enabled
   Else
      '2008/11/6 MODIFY BY SONIA 已作廢不可修改
      'tool1_enabled
      'modify by sonia 2017/4/17 +已抵帳不可修改m_a1512 <> ""
      If MaskEdBox2.Text <> MsgText(29) Or m_a1512 <> "" Then
         tool15_enabled
      Else
         tool1_enabled
      End If
      '2008/11/6 END
   End If
   adoquery.Close
   
End Sub

'*************************************************
'  顯示 Adodc 之資料
'
'*************************************************
Private Sub AdodcShow()
   Text10 = Mid(Adodc1.Recordset.Fields("axf03").Value, 1, Len(Adodc1.Recordset.Fields("axf03").Value) - 9)
   Select Case Text10
      Case "TF"
         '2010/7/2 ADD BY SONIA
         Text5 = Mid(Adodc1.Recordset.Fields("axf03").Value, Len(Adodc1.Recordset.Fields("axf03").Value) - 8, 5)
         Text7 = Mid(Adodc1.Recordset.Fields("axf03").Value, Len(Adodc1.Recordset.Fields("axf03").Value) - 7, 1)
         Text9 = Mid(Adodc1.Recordset.Fields("axf03").Value, Len(Adodc1.Recordset.Fields("axf03").Value) - 2, 1)
         Text12 = Mid(Adodc1.Recordset.Fields("axf03").Value, Len(Adodc1.Recordset.Fields("axf03").Value) - 1, 2)
         Text12.Visible = True
         '2010/7/2 END
      Case Else
         Text5 = Mid(Adodc1.Recordset.Fields("axf03").Value, Len(Adodc1.Recordset.Fields("axf03").Value) - 8, 6)
         Text7 = Mid(Adodc1.Recordset.Fields("axf03").Value, Len(Adodc1.Recordset.Fields("axf03").Value) - 2, 1)
         Text9 = Mid(Adodc1.Recordset.Fields("axf03").Value, Len(Adodc1.Recordset.Fields("axf03").Value) - 1, 2)
         Text12.Visible = False   '2010/7/2 ADD BY SONIA
   End Select
   If IsNull(Adodc1.Recordset.Fields("axf04").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Format(Adodc1.Recordset.Fields("axf04").Value, FAmount)
   End If
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Public Sub AdodcClear()
   Text10 = ""
   Text5 = ""
   Text7 = ""
   Text9 = ""
   Text11 = ""
   Text12 = ""
   'edit by nickc 2007/02/08
   'Text13 = ""
End Sub
'2008/4/9 add by sonia 為避免錯誤U09702029及U09702240同一帳單重覆問題,代理人D/N No.不可空白
Private Sub Text4_LostFocus()
Dim strTemp1 As String 'Added by Lydia 2019/03/18
   If strSaveConfirm = MsgText(3) And Text4 = "" Then
      MsgBox "為免帳單重覆輸入,代理人D/N No.不可空白,請依貴單位統一方式輸入", vbExclamation + vbOKOnly
      Me.Text4.SetFocus
      Text4_GotFocus
   'Added by Lydia 2019/03/18 檢查不可輸入中文字
   ElseIf strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
       strTemp1 = PUB_GetSimpleName(Text4.Text, True, True)
       If Text4.Text <> strTemp1 Then
           MsgBox "代理人D/N No.只可輸入英數字、半形空白及""_"",""-""，請修改！", vbCritical
           Me.Text4.SetFocus
           Text4_GotFocus
       End If
   'end 2019/03/18
   End If
   
'Mem by Morgan 2009/10/2 移到存檔時才檢查
   
'   'Add By Sindy 2009/06/17
'   '檢查抵帳單資料是否重覆
'   '若為專利處只須以代理人+代理人D/N No.做重覆檢核
'   If Left(Trim(GetStaffDepartment(strUserNum)), 2) = "P1" Then
'      If PUB_ChkDNDup("", Text1.Text, Text4.Text, Text2.Text, , 0) = True Then
'         Text4.SetFocus
'         Text4_GotFocus
'         Exit Sub
'      End If
'   End If
End Sub
'2008/4/9 end

Private Sub Text5_GotFocus()
   CloseIme
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Select Case Text10
      Case "TF"
         Text9 = "0"
         Text12 = "00"
      Case Else
         Text7 = "0"
         Text9 = "00"
   End Select
End Sub

Private Sub Text6_GotFocus()
   CloseIme
   TextInverse Text6
End Sub

'2010/4/2 ADD BY SONIA 加帳單台幣金額
Private Sub Text6_LostFocus()
   If Text6 <> MsgText(601) Then
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select decode(" & Val(Replace(MaskEdBox2, "/", "")) & ", 0, nvl(NVL(NVL(" & Text6 & "*decode(A1906,0,null,A1906)," & Text6 & "*A1G03),0), 0), 0) as PayAmt," & Text6 & "*NVL(A2103,1) Payable " & _
                    "from (select 1 x1,a2103 from acc210 where a2102 = '" & Combo1 & "' and a2101 = (select max(a2101) from acc210 where a2102 = '" & Combo1 & "' and a2101 <= " & strSrvDate(2) & ")) x, " & _
                    "     (select 1 y1,a1906 from acc190 where A1902='" & Text2 & "') y, " & _
                    "     (select 1 z1,a1g03 from acc1g0 where '" & m_a1512 & "'=A1G01) z" & _
                    " where x.x1=y.y1(+) and x.x1=z.z1(+)", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If Val(adoquery.Fields("PayAmt")) <> 0 Then
            Label8 = "實際結匯台幣金額：" & adoquery.Fields("PayAmt")
         Else
            Label8 = "預估結匯台幣金額：" & adoquery.Fields("Payable")
         End If
      End If
      adoquery.Close
   End If
End Sub
'2010/4/2 END

Private Sub Text7_GotFocus()
   CloseIme
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
   TextInverse Text8
End Sub

'*************************************************
'  儲存資料表(國外帳單資料(交易檔))
'
'*************************************************
'Private Sub Acc151Save()
'On Error GoTo Checking
'      If Text10 = MsgText(601) Then
'         MsgBox MsgText(10) & Label11, , MsgText(5)
'         strControlButton = MsgText(602)
'         Text10.SetFocus
'         Exit Sub
'      End If
'      If Val(Text11) = 0 Then
'         MsgBox MsgText(186), , MsgText(5)
'         strControlButton = MsgText(602)
'         Text11.SetFocus
'         Exit Sub
'      End If
'      If Adodc1.Recordset.RecordCount <> 0 Then
'         Adodc1.Recordset.Find "axf02 = '" & Text9 & "'", 0, adSearchForward, 1
'         If Adodc1.Recordset.EOF Then
'            Adodc1.Recordset.AddNew
'         End If
'      Else
'         Adodc1.Recordset.AddNew
'      End If
'      Adodc1.Recordset.Fields("axf01").Value = Text2
'      Adodc1.Recordset.Fields("axf02").Value = Text9
'      If Text10 <> MsgText(601) Then
'         Adodc1.Recordset.Fields("axf03").Value = Text10
'      Else
'         Adodc1.Recordset.Fields("axf03").Value = Null
'      End If
'      If Text7 <> MsgText(601) Then
'         Adodc1.Recordset.Fields("axf12").Value = Text7
'      Else
'         Adodc1.Recordset.Fields("axf12").Value = Null
'      End If
'      If Text11 <> MsgText(601) Then
'         Adodc1.Recordset.Fields("axf04").Value = Val(Text11)
'      Else
'         Adodc1.Recordset.Fields("axf04").Value = Null
'      End If
'      If Text12 <> MsgText(601) Then
'         Adodc1.Recordset.Fields("axf05").Value = Text12
'      Else
'         Adodc1.Recordset.Fields("axf05").Value = Null
'      End If
'      Adodc1.Recordset.Fields("axf06").Value = Val(strSrvDate(2))
'      Adodc1.Recordset.Fields("axf07").Value = ServerTime
'      Adodc1.Recordset.Fields("axf08").Value = strUserNum
'      Adodc1.Recordset.UpdateBatch
'      adoquery.CursorLocation = adUseClient
'      '2007/3/2 modify by sonia 加入cp87,cp88
'      'adoquery.Open "select cp61, cp62, cp63, a0k04 from caseprogress, acc0k0 where cp60 = a0k01 (+) and cp09 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      adoquery.Open "select cp61, cp62, cp63, cp87, cp88, a0k04 from caseprogress, acc0k0 where cp60 = a0k01 (+) and cp09 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      '2007/3/2 end
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields(0).Value) = True Or adoquery.Fields(0).Value = "" Or adoquery.Fields(0).Value = Text2 Then
'            adoTaie.Execute "update caseprogress set cp61 = '" & Text2 & "' where cp09 = '" & Text9 & "'"
'         Else
'            If IsNull(adoquery.Fields(1).Value) = True Or adoquery.Fields(1).Value = "" Or adoquery.Fields(0).Value = Text2 Then
'               adoTaie.Execute "update caseprogress set cp62 = '" & Text2 & "' where cp09 = '" & Text9 & "'"
'            Else
'               If IsNull(adoquery.Fields(2).Value) = True Or adoquery.Fields(2).Value = "" Or adoquery.Fields(0).Value = Text2 Then
'                  adoTaie.Execute "update caseprogress set cp63 = '" & Text2 & "' where cp09 = '" & Text9 & "'"
'               '2007/3/2 add by sonia 加入cp87,cp88
'               Else
'                  If IsNull(adoquery.Fields(2).Value) = True Or adoquery.Fields(2).Value = "" Or adoquery.Fields(0).Value = Text2 Then
'                     adoTaie.Execute "update caseprogress set cp87 = '" & Text2 & "' where cp09 = '" & Text9 & "'"
'                  Else
'                     If IsNull(adoquery.Fields(2).Value) = True Or adoquery.Fields(2).Value = "" Or adoquery.Fields(0).Value = Text2 Then
'                        adoTaie.Execute "update caseprogress set cp88 = '" & Text2 & "' where cp09 = '" & Text9 & "'"
'                     End If
'                  End If
'               '2007/3/2 end
'               End If
'            End If
'         End If
'         If IsNull(adoquery.Fields("a0k04").Value) Then
'            adoTaie.Execute "update acc151 set axf13 = null where axf01 = '" & Text2 & "' and axf02 = '" & Text9 & "'"
'         Else
'            'Modified by Morgan 2012/1/10 收據抬頭會有單引號
'            adoTaie.Execute "update acc151 set axf13 = '" & ChgSQL("" & adoquery.Fields("a0k04").Value) & "' where axf01 = '" & Text2 & "' and axf02 = '" & Text9 & "'"
'         End If
'      End If
'      adoquery.Close
'      AdodcRefresh
'Checking:
'   If Err.NUMBER = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)

   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         'Added by Lydia 2021/12/07 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'end 2021/12/07
         '2009/11/25 add by sonia
         If Combo1 = MsgText(601) Then
            MsgBox "請輸入帳單幣別, 否則無法計算損益...", , MsgText(5)
            strControlButton = MsgText(602)
            Combo1.SetFocus
            Exit Sub
         End If
         '2009/11/25 end
        'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
         If FMP2open = True Then
            If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text10, Text5, Text7, Text9) = False Then
              Text5.SetFocus
              Exit Sub
            End If
         'Added by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入並出現訊息告知USER「此案為FCP自行連繫，請交FCP程序處理」。
         ElseIf Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M31" Then
              If PUB_FMPtoCheck(1, 2, Pub_strUserST05, Text10, Text5, Text7, Text9) = True Then
                   MsgBox "此案為FCP自行連繫，請交FCP程序處理！", vbCritical, "寰華案控制輸入"
                   Exit Sub
              End If
         'end 2019/09/10
         End If
     
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         If Val(Text11) = 0 Then
            MsgBox MsgText(162) & Label12, , MsgText(5)
            Text11.SetFocus
            Exit Sub
         End If
         
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         Select Case Text10
            Case "TF"
               strExc(1) = PUB_GetReceiptComp(Text10, Text5 & Text7, Text9, Text12) 'Added by Morgan 2019/9/18
               adoquery.Open "select cp09 from caseprogress where cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "'", adoTaie, adOpenStatic, adLockReadOnly
            Case Else
               strExc(1) = PUB_GetReceiptComp(Text10, Text5, Text7, Text9)  'Added by Morgan 2019/9/18
               adoquery.Open "select cp09 from caseprogress where cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
         End Select
         If adoquery.RecordCount = 0 Then
            adoquery.Close
            MsgBox MsgText(188) & Label11, , MsgText(5)
            Text5.SetFocus
            Exit Sub
         End If
         adoquery.Close
         
         'Added by Morgan 2019/9/18 一張帳單不可不同公司別--禧佩,婉莘
         If Adodc1.Recordset.RecordCount > 0 Then
            If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
            strExc(0) = Adodc1.Recordset.Fields("axf03")
            strExc(2) = PUB_GetReceiptComp(Left(strExc(0), Len(strExc(0)) - 9), Mid(strExc(0), Len(strExc(0)) - 8, 6), Mid(strExc(0), Len(strExc(0)) - 2, 1), Right(strExc(0), 2))
            If strExc(1) <> strExc(2) Then
               MsgBox "不同公司別不可輸在同一帳單編號！", vbCritical
               Exit Sub
            End If
         End If
         'end 2019/9/18
         
         If Text1 <> MsgText(601) Then
            strCustNo = Text1
         Else
            strCustNo = ""
         End If
         If Text3 <> MsgText(601) Then
            strCon1 = Text3
         Else
            strCon1 = ""
         End If
         If Text10 <> MsgText(601) Then
            strCon2 = Text10
         Else
            strCon2 = ""
         End If
         If Text5 <> MsgText(601) Then
            strCon3 = Text5
         Else
            strCon3 = ""
         End If
         If Text7 <> MsgText(601) Then
            strCon4 = Text7
         Else
            strCon4 = ""
         End If
         If Text9 <> MsgText(601) Then
            strCon5 = Text9
         Else
            strCon5 = ""
         End If
         If Text12 <> MsgText(601) Then
            strCon6 = Text12
         Else
            strCon6 = ""
         End If
         If Text11 <> MsgText(601) Then
            strCon7 = Text11
         Else
            strCon7 = ""
         End If
         If Text2 <> MsgText(601) Then
            strCon8 = Text2
         Else
            strCon8 = ""
         End If
         If Combo1 <> MsgText(601) Then
            strTitle = Combo1
         Else
            strTitle = ""
         End If
         tool3_enabled
         Screen.MousePointer = vbHourglass
         Frmacc2152.Show
         Screen.MousePointer = vbDefault
         Me.Hide
'         Frmacc2150_Save
'         If strControlButton <> MsgText(602) Then
'            Acc151Save
'         End If
'         If strControlButton <> MsgText(602) Then
'            AdodcClear
'            Text10.SetFocus
'         End If
'         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text8_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text9_GotFocus()
   CloseIme
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
   '2007/3/2 modify by sonia 加入cp87,cp88
   'adoTaie.Execute "update caseprogress set cp61=decode(cp61, '" & Text2 & "', null, cp61), cp62=decode(cp62, '" & Text2 & "', null, cp62), cp63=decode(cp63, '" & Text2 & "', null, cp63) where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'"
   'Modified by Morgan 2017/1/17 一收文號若有多張帳單,前面的刪除後面的要往前順移
   'adoTaie.Execute "update caseprogress set cp61=decode(cp61, '" & Text2 & "', null, cp61), cp62=decode(cp62, '" & Text2 & "', null, cp62), cp63=decode(cp63, '" & Text2 & "', null, cp63), cp87=decode(cp87, '" & Text2 & "', null, cp87), cp88=decode(cp88, '" & Text2 & "', null, cp88) where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'"
   adoTaie.Execute "update caseprogress set cp61=cp62,cp62=cp63,cp63=cp87,cp87=cp88,cp88=''  where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp61='" & Text2 & "'", intI
   If intI = 0 Then
      adoTaie.Execute "update caseprogress set cp62=cp63,cp63=cp87,cp87=cp88,cp88=''  where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp62='" & Text2 & "'", intI
      If intI = 0 Then
         adoTaie.Execute "update caseprogress set cp63=cp87,cp87=cp88,cp88=''  where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp63='" & Text2 & "'", intI
         If intI = 0 Then
            adoTaie.Execute "update caseprogress set cp87=cp88,cp88=''  where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp87='" & Text2 & "'", intI
            If intI = 0 Then
               adoTaie.Execute "update caseprogress set cp88=''  where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp88='" & Text2 & "'", intI
            End If
         End If
      End If
   End If
   'end 2017/1/17
   '2007/3/2 end
   adoTaie.Execute "delete from acc151 where axf01 = '" & Text2 & "' and axf02 = '" & Adodc1.Recordset.Fields("cp09").Value & "'"
   'Adodc1.Recordset.Delete
   'Adodc1.Recordset.UpdateBatch
   AdodcRefresh
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc150.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow adoacc150.Bookmark, adoacc150.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text2.Enabled = True
   Text10.Enabled = False
   Text5.Enabled = False
   Text7.Enabled = False
   Text9.Enabled = False
   Text12.Enabled = False
   Text11.Enabled = False
   Command1.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text2.Enabled = False
   Text10.Enabled = True
   Text5.Enabled = True
   Text7.Enabled = True
   Text9.Enabled = True
   Text12.Enabled = True
   Text11.Enabled = True
   Command1.Enabled = True
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(axf04) from acc151 where axf01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text14 = MsgText(601)
      Else
         Text14 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
   Else
      Text14 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  重新整理國外帳單資料
'
'*************************************************
Public Sub Acc150Refresh()
On Error GoTo Checking
   
   If adoacc150.State = adStateOpen Then
      adoacc150.Close
   End If
   adoacc150.CursorLocation = adUseClient
   adoacc150.MaxRecords = intMax
  ' adoacc150.Open "select * from acc150 where a1501 >= '" & Text2 & "' order by a1501 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。=>RQstr
   Dim midSql As String
   midSql = " select m0.* from acc150 m0 where a1501 >= '" & Text2 & "' "
   'Modified by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入
   'If FMP2open = True Then
   If FMP2openSQL <> "" And Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M31" Then
      midSql = midSql & " and a1501 in (" & RQstr & ") "
      'Added by Lydia 2019/12/10 考慮尚未輸入明細的狀況;
      midSql = midSql & "union all select m0.* from acc150 m0 where a1501>='" & Text2 & "' and a1501 not in (select axf01 from acc151 where axf01>='" & Text2 & "' )"
   End If
   midSql = midSql & " order by 1 asc "
   adoacc150.Open midSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Public Function ChkDataRepaet(strA1502 As String, strA1503 As String, strA1504 As String) As Boolean
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'
''Add By Cheng 2004/02/12
'strControlButton = MsgText(601)
''End
'ChkDataRepaet = False
''2005/11/1 MODIFY BY SONIA strA1504為NULL時不可如此判斷
''StrSQLa = "Select * From Acc150 Where A1502=" & Val(Replace(strA1502, "/", "")) & " And A1503='" & strA1503 & "' And A1504='" & strA1504 & "' "
''Modify by Morgan 2009/4/6 加判斷未作廢
'If strA1504 = "" Then
'   'StrSQLa = "Select * From Acc150 Where A1502=" & Val(Replace(strA1502, "/", "")) & " And A1503='" & strA1503 & "' "
'   StrSQLa = "Select * From Acc150 Where A1502=" & Val(Replace(strA1502, "/", "")) & " And A1503='" & strA1503 & "' and a1507 is null"
'Else
'   'StrSQLa = "Select * From Acc150 Where A1502=" & Val(Replace(strA1502, "/", "")) & " And A1503='" & strA1503 & "' And A1504='" & strA1504 & "' "
'   StrSQLa = "Select * From Acc150 Where A1502=" & Val(Replace(strA1502, "/", "")) & " And A1503='" & strA1503 & "' And A1504='" & strA1504 & "' and a1507 is null"
'End If
''2005/11/1 END
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    ChkDataRepaet = True
'    If strA1504 = "" Then
'      MsgBox "此帳單資料重覆，請確認!!!", vbExclamation + vbOKOnly
'    Else
'      MsgBox "代理人此帳單D/N No.重覆，請確認!!!", vbExclamation + vbOKOnly
'    End If
'    strControlButton = MsgText(602)
'    'End
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function

'Added by Lydia 2017/05/17 新增帳單檢查翻譯費的完稿字數是否超過原來預估
Public Sub ChkMailTransFee()
Dim inA As Integer
Dim rsAD As New ADODB.Recordset
Dim strA1 As String
Dim bolChk1 As Boolean 'Added by Lydia 2018/01/05
Dim strWordVal As String, inB As Integer 'Added by Lydia 2019/05/16 因為備註自動加註訊息(ex.此收文號有虧損;) ,所以先判斷有無";"
   
   'Added by Lydia 2021/12/29 增加檢查:沒有異常訊息，但是ACC151有存，進度檔的CP61也有存，可是就是沒有ACC150; ex.CFP-031156-0-35
   strA1 = "SELECT DISTINCT axf01,AXF02,axf03 FROM ACC151,ACC150 WHERE AXF02<>'000000000' AND AXF01=A1501(+) AND A1501 IS NULL " & _
               IIf(Text2.Text <> "", " and axf01=" & CNULL(Trim(Text2.Text)), "")
   inA = 1
   Set rsAD = ClsLawReadRstMsg(inA, strA1)
   If inA = 1 Then
        strExc(6) = "帳單編號" & String(10, " ") & "收文號" & String(10, " ") & "本所案號"
        rsAD.MoveFirst
        Do While Not rsAD.EOF
             strExc(6) = strExc(6) & vbCrLf & convForm(rsAD.Fields("axf01"), 18) & convForm(rsAD.Fields("axf02"), 16) & rsAD.Fields("axf03")
             rsAD.MoveNext
        Loop
        PUB_SendMail strUserNum, "83002", "", "帳單沒有ACC150", strExc(6)
   End If
   'end 2021/12/29
   If rsAD.State <> adStateClosed Then rsAD.Close 'Added by Morgan 2022/1/11 沒close後面再用會錯
   
   'Added by Lydia 2019/05/16  抓原文字數
   If Trim(Text8) <> "" Then
        inA = InStr(Text8, ";")
        inB = InStr(Text8, "/")
        If inB > inA Then '訊息在前
            If inA > 0 Then
               strWordVal = Val(Mid(Text8, inA + 1))
            Else
               strWordVal = Val(Text8)
            End If
        Else '無訊息或訊息在後
            strWordVal = Val(Text8)
        End If
   End If
    
   '判斷代理人為舜禹(Y53541)或捷恩凱(Y52268)
   'Modified by Lydia 2017/10/16 改成共用變數
   'If Trim(Text2) <> "" And Val(Text8) > 0 And InStr("Y5354100,Y5226800", Left(Text1, 8)) > 0 Then
   'Modified by Lydia 2019/05/16 Text8=> strWordVal
   'Modified by Lydia 2025/03/13 改用模組取得
   'If Trim(Text2) <> "" And Val(strWordVal) > 0 And InStr(外翻Y編號, Left(Text1, 8)) > 0 Then
   If Trim(Text2) <> "" And Val(strWordVal) > 0 And InStr(Pub_SetF51Order("Y", ""), Left(Text1, 8)) > 0 Then
      '抓翻譯費有原文字數和相似度
      strA1 = "SELECT AXF01,AXF02,AXF03,B.*,CP10,NVL(CPM03,CPM04) CPM03 FROM ACC151,TRANSFEE B,CASEPROGRESS,CASEPROPERTYMAP " & _
                 "WHERE AXF01='" & Trim(Text2) & "' AND AXF02=TF01(+) AND TF01=CP09(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP10 IN ('201','927') AND NVL(TF23,0) > 0 AND NVL(TF19,0) > 0 "
      rsAD.CursorLocation = adUseClient
      rsAD.Open strA1, cnnConnection, adOpenStatic, adLockReadOnly
      If rsAD.RecordCount > 0 Then
            bolChk1 = True 'Added by Lydia 2018/01/05
            'Modified by Lydia 2019/05/16 Text8=> strWordVal
            strA1 = Format(Val(strWordVal) / (Val("" & rsAD.Fields("TF23")) * (1 - Val("" & rsAD.Fields("TF19")) / 100)), "0.00")
            '比對完稿字數(備註前固定輸入數字)和原文字數(預估值),超出5%,發email通知
            'Modified by Lydia 2018/04/02 原文字數比對請排除P案設定,因為無法要求內專輸入原文字數(ex.P119594)
            'If Val(strA1) > 1.05 Then
            If Val(strA1) > 1.05 And Mid("" & rsAD.Fields("AXF03"), 1, 3) = "FCP" Then
                Call ChgCaseNo("" & rsAD.Fields("AXF03"), strExc)
                strExc(5) = IIf(strExc(3) & strExc(4) = "000", strExc(1) & strExc(2), strExc(1) & strExc(2) & strExc(3) & strExc(4))
                '內文
                'Modified by Lydia 2019/05/16 Text8=> strWordVal
                strExc(6) = vbCrLf & "本所案號：" & strExc(5) & vbCrLf & _
                           "收  文  號：" & rsAD.Fields("AXF02") & "　　" & rsAD.Fields("CPM03") & vbCrLf & _
                           "原文字數：" & PUB_StrToStr("" & rsAD.Fields("TF23"), 6, True, True) & " 字" & "　　" & _
                           "相  似  度：" & PUB_StrToStr("" & rsAD.Fields("TF19"), 3, True, True) & " %" & _
                           IIf("" & rsAD.Fields("TF20") <> "", "　　相似案號：" & rsAD.Fields("TF20"), "") & vbCrLf & _
                           "完稿字數：" & PUB_StrToStr(Val(strWordVal), 6, True, True) & " 字" & vbCrLf & _
                           "完稿字數比對結果：" & (Val(strA1) - 1) * 100 & " %"
                PUB_SendMail strUserNum, "86013", "", strExc(5) & "翻譯字數異常超過5%", strExc(6)
            End If
      End If
      If rsAD.State <> adStateClosed Then rsAD.Close
      'Added by Lydia 2018/01/05 新案翻譯-舜禹,捷恩凱,迅達需比對完稿字數,備註欄位/前數字與原字數計算,
      If bolChk1 = False Then
           'Modified by Lydi 2018/01/08 +CP66
            strA1 = "SELECT AXF01,AXF02,AXF03,B.*,CP10,NVL(CPM03,CPM04) CPM03,CP66 FROM ACC151,TRANSFEE B,CASEPROGRESS,CASEPROPERTYMAP " & _
                    "WHERE AXF01='" & Trim(Text2) & "' AND AXF02=TF01(+) AND TF01=CP09(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP10 IN ('201') "
            rsAD.CursorLocation = adUseClient
            rsAD.Open strA1, cnnConnection, adOpenStatic, adLockReadOnly
            If rsAD.RecordCount > 0 Then
                'Modified by Lydia 2018/04/02 原文字數比對請排除P案設定,因為無法要求內專輸入原文字數(ex.P119594)
                'If "" & rsAD.Fields("CP66") >= "20180101" Then  'Added by Lydia 2018/01/08 從107/1/1開始控管 by Sharon
                If "" & rsAD.Fields("CP66") >= "20180101" And Mid("" & rsAD.Fields("AXF03"), 1, 3) = "FCP" Then
                    Call ChgCaseNo("" & rsAD.Fields("AXF03"), strExc)
                    strExc(5) = IIf(strExc(3) & strExc(4) = "000", strExc(1) & strExc(2), strExc(1) & strExc(2) & strExc(3) & strExc(4))
                     '若無原文字數可比對, 自動發一Email給程序管制人員,cc:Sharon
                     If Val("" & rsAD.Fields("TF23")) = 0 Or Val("" & rsAD.Fields("TF23")) < 0 Then
                           strExc(6) = PUB_GetFCPHandler(strExc(1), strExc(2), strExc(3), strExc(4))
                           If strExc(6) <> "" Then
                                PUB_SendMail strUserNum, strExc(6), "", strExc(5) & "無原文字數可比對,請後續追蹤交稿字數是否為正確", "同主旨", , , , , , "86013"
                           End If
                     Else
                     '完稿字數大於原文字數250字, 自動發一Email至Sharon
                           'Modified by Lydia 2018/04/24 若完稿字數大於原文字數5%(原文10000字以下)或3%(原文超過10000字)
                           'If Val(Text8) > 0 And Val(Text8) > Val("" & rsAD.Fields("TF23")) + 250 Then
                           'Modified by Lydia 2019/05/16 Text8=> strWordVal
                           If Val(strWordVal) > 0 And Val(strWordVal) > Val("" & rsAD.Fields("TF23")) + Format(IIf(Val("" & rsAD.Fields("TF23")) <= 10000, Val("" & rsAD.Fields("TF23")) * 0.05, Val("" & rsAD.Fields("TF23")) * 0.03), "0") Then
                                strExc(6) = vbCrLf & "本所案號：" & strExc(5) & vbCrLf & _
                                           "收  文  號：" & rsAD.Fields("AXF02") & "　　" & rsAD.Fields("CPM03") & vbCrLf & _
                                           "原文字數：" & PUB_StrToStr("" & rsAD.Fields("TF23"), 6, True, True) & " 字" & "　　" & _
                                           "完稿字數：" & PUB_StrToStr(Val(strWordVal), 6, True, True) & " 字" & vbCrLf & _
                                           "完稿字數比對結果：大於" & Val(strWordVal) - Val("" & rsAD.Fields("TF23")) & " 字"
                                'Modified by Lydia 2018/04/24 改主旨
                                'PUB_SendMail strUserNum, "86013", "", strExc(5) & "翻譯字數異常大於250字", strExc(6)
                                PUB_SendMail strUserNum, "86013", "", strExc(5) & "翻譯字數異常大於" & IIf(Val("" & rsAD.Fields("TF23")) <= 10000, "５％", "３％"), strExc(6)
                           End If
                     End If
                End If 'end 2018/01/08
            End If
      End If
      'end 2018/01/05
      Set rsAD = Nothing
   End If
End Sub
'end 2017/05/17

'Added by Morgan 2023/4/7 設定可付款日期
'Modified by Morgan 2023/12/20 改週三付款,週五不可選次週不變--斯閔
'智慧所第1-3個週二(若有5個週二則為第2-4個週二)，遇假日順延到下一工作天
'智權公司最後一個週二，遇假日順延到下一工作天
'週五開始不可選次週二
Private Sub SetPayDay()
   
   Dim bGood As Boolean, bLstWk As Boolean
   Dim stDate1 As String, stDate2 As String, stDate3 As String, stWkDate As String
   Dim adoRst As ADODB.Recordset
      
   m_Comp = GetComp()
   Combo2.Clear
   Combo2.Tag = ""
   
   '週五開始不可選次週二:dd>sysdate+4
   'Modified by Morgan 2023/12/20 改週三付款,週五不可選次週不變--斯閔
   'strExc(0) = "select sqldatet(to_char(dd,'yyyymmdd')) pd,wd01" & _
      " from (select sysdate+rownum dd from workday where rownum<=365) x,workday" & _
      " where to_char(dd,'D')=3 and wd01(+)=to_char(dd,'yyyymmdd') and dd>sysdate+4"
   strExc(0) = "select sqldatet(to_char(dd,'yyyymmdd')) pd,wd01" & _
      " from (select sysdate+rownum dd from workday where rownum<=365) x,workday" & _
      " where to_char(dd,'D')=4 and wd01(+)=to_char(dd,'yyyymmdd') and dd>sysdate+5"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      Do While Not .EOF
         stDate1 = DBDATE(.Fields("pd"))
         stDate2 = CompDate(2, 7, stDate1)
         
         bGood = True
         '是否最後一個週二
         If Left(stDate2, 6) = Left(stDate1, 6) Then
            bLstWk = False
         Else
            bLstWk = True
         End If
         
         '智權公司最後一個週二
         If m_Comp = "J" Then
            If bLstWk Then
               bGood = True
            Else
               bGood = False
            End If
            
         '智慧所
         '排除最後一個週二
         ElseIf bLstWk Then
            bGood = False
            
         Else
            '有5個週二的第1個週二不付款
            If Right(stDate1, 2) <= "03" Then
               stDate3 = CompDate(2, 28, stDate1)
               If Left(stDate3, 6) = Left(stDate1, 6) Then
                  bGood = False
               End If
            End If
         End If
         
         If bGood Then
            '遇假日順延到下一工作天
            If IsNull(.Fields("wd01")) Then
               stWkDate = PUB_GetWorkDay1(stDate1, False)
               '若下一工作天超過7天則略過(如過年)
               If stWkDate >= stDate2 Then
                  bGood = False
               End If
            Else
               stWkDate = stDate1
            End If
         End If
         If bGood Then
            stWkDate = stWkDate - 19110000
            '檢查是否有付款日期調整紀錄
            strExc(0) = "select a1f02 from acc1F0 where a1f01=" & stWkDate
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               stWkDate = RsTemp.Fields(0)
            End If
            Combo2.AddItem CFDate(stWkDate)
         End If
         .MoveNext
      Loop
      Combo2.Tag = m_Comp
      End With
   End If
   Set adoRst = Nothing
End Sub

'Added by Morgan 2023/4/21
Private Function GetComp() As String
   Dim strCaseNo As String
   
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      strCaseNo = Adodc1.Recordset.Fields("axf03")
      GetComp = PUB_GetReceiptComp(Left(strCaseNo, Len(strCaseNo) - 9), Left(Right(strCaseNo, 9), 6), Left(Right(strCaseNo, 3), 1), Right(strCaseNo, 2))
   End If
End Function

'Added by Morgan 2024/12/27
'幣別檢查
Private Sub CurCheck()
   If Combo1.Tag <> "" And (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) Then
      If Combo1.Tag <> Combo1 Then
         If Adodc1.Recordset.RecordCount > 0 Then
            If MsgBox("修改幣別時, 必須重新輸入明細資料, 以重新計算正確損益, " & vbCrLf & "原明細將清除, 是否要繼續?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
               Do While (Adodc1.Recordset.RecordCount <> 0)
                  AdodcDelete
               Loop
            Else
               Combo1 = Combo1.Tag
            End If
         End If
      End If
   End If
   Combo1.Tag = Combo1
End Sub
