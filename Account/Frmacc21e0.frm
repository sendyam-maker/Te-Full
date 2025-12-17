VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21e0 
   AutoRedraw      =   -1  'True
   Caption         =   "付款後退費作業"
   ClientHeight    =   5556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5556
   ScaleWidth      =   8760
   Begin VB.TextBox Text19 
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
      Left            =   1296
      MaxLength       =   8
      TabIndex        =   15
      Top             =   4980
      Width           =   1572
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6828
      MaxLength       =   10
      TabIndex        =   14
      Top             =   4655
      Width           =   1572
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      MaxLength       =   9
      TabIndex        =   13
      Top             =   4655
      Width           =   1572
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1296
      MaxLength       =   12
      TabIndex        =   12
      Top             =   4655
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1296
      MaxLength       =   3
      TabIndex        =   10
      Top             =   4300
      Width           =   528
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   6816
      TabIndex        =   3
      Top             =   96
      Width           =   1572
   End
   Begin VB.TextBox Text15 
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
      Height          =   315
      Left            =   6816
      MaxLength       =   14
      TabIndex        =   9
      Top             =   3960
      Width           =   1572
   End
   Begin VB.TextBox Text20 
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
      Left            =   1296
      MaxLength       =   12
      TabIndex        =   35
      Top             =   3300
      Width           =   855
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
      Height          =   300
      Left            =   5088
      TabIndex        =   34
      Top             =   3300
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   5616
      Picture         =   "Frmacc21e0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   108
      Width           =   350
   End
   Begin VB.TextBox Text13 
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
      Height          =   315
      Left            =   6816
      TabIndex        =   33
      Top             =   1188
      Width           =   1572
   End
   Begin VB.TextBox Text12 
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
      Left            =   3576
      TabIndex        =   31
      Top             =   3300
      Width           =   1452
   End
   Begin VB.TextBox Text9 
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
      Height          =   315
      Left            =   1296
      MaxLength       =   13
      TabIndex        =   4
      Top             =   468
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21e0.frx":0102
      Height          =   1500
      Left            =   210
      TabIndex        =   17
      Top             =   1695
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2646
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "a0102"
         Caption         =   "會計科目"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   3072.189
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1476.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1548.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   7020.284
         EndProperty
      EndProperty
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
      Height          =   350
      Left            =   7656
      Picture         =   "Frmacc21e0.frx":0117
      Style           =   1  '圖片外觀
      TabIndex        =   16
      ToolTipText     =   "取消"
      Top             =   3300
      Width           =   350
   End
   Begin VB.TextBox Text7 
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
      Height          =   315
      Left            =   5016
      MaxLength       =   14
      TabIndex        =   8
      Top             =   3960
      Width           =   1572
   End
   Begin VB.TextBox Text5 
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
      Left            =   336
      MaxLength       =   6
      TabIndex        =   7
      Top             =   3960
      Width           =   1572
   End
   Begin VB.TextBox Text4 
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
      Height          =   300
      Left            =   4416
      MaxLength       =   14
      TabIndex        =   5
      Top             =   468
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Left            =   1296
      TabIndex        =   21
      Top             =   828
      Width           =   1572
   End
   Begin VB.TextBox Text1 
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
      Left            =   4416
      MaxLength       =   15
      TabIndex        =   1
      Top             =   108
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1296
      TabIndex        =   0
      Top             =   108
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
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
      Height          =   315
      Left            =   210
      Top             =   1575
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
   Begin MSForms.TextBox Text11 
      Height          =   465
      Left            =   1296
      TabIndex        =   6
      Top             =   1170
      Width           =   3975
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "7011;820"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   6816
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   828
      Width           =   1575
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "2778;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   330
      Left            =   1920
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3945
      Width           =   2775
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text10 
      Height          =   330
      Left            =   2880
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   828
      Width           =   3165
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "5583;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   345
      Left            =   4020
      TabIndex        =   11
      Top             =   4285
      Width           =   4335
      VariousPropertyBits=   679495707
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "7646;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "對沖(業)"
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
      TabIndex        =   42
      Top             =   5003
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
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
      Left            =   5865
      TabIndex        =   41
      Top             =   4685
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "對沖(客)"
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
      Left            =   3075
      TabIndex        =   40
      Top             =   4685
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
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
      TabIndex        =   39
      Top             =   4685
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      TabIndex        =   38
      Top             =   4330
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
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
      Left            =   6810
      TabIndex        =   37
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label22 
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
      Height          =   255
      Left            =   330
      TabIndex        =   36
      Top             =   3300
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "退費台幣金額"
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
      Left            =   5376
      TabIndex        =   32
      Top             =   1188
      Width           =   1452
   End
   Begin VB.Label Label12 
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
      Height          =   255
      Left            =   2850
      TabIndex        =   30
      Top             =   3300
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
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
      Left            =   3075
      TabIndex        =   29
      Top             =   4320
      Width           =   750
   End
   Begin VB.Label Label10 
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
      Height          =   252
      Left            =   336
      TabIndex        =   28
      Top             =   1188
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "匯率"
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
      Left            =   336
      TabIndex        =   27
      Top             =   468
      Width           =   972
   End
   Begin VB.Label Label8 
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
      Height          =   252
      Left            =   6096
      TabIndex        =   26
      Top             =   108
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4656
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1650
      Left            =   225
      Top             =   3675
      Width           =   8295
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
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
      Left            =   5010
      TabIndex        =   25
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   24
      Top             =   3720
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "退費外幣金額"
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
      Left            =   2976
      TabIndex        =   23
      Top             =   468
      Width           =   1452
   End
   Begin VB.Label Label4 
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
      Height          =   252
      Left            =   6096
      TabIndex        =   22
      Top             =   828
      Width           =   732
   End
   Begin VB.Label Label3 
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
      Height          =   252
      Left            =   336
      TabIndex        =   20
      Top             =   828
      Width           =   972
   End
   Begin VB.Label Label2 
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
      Height          =   252
      Left            =   2976
      TabIndex        =   19
      Top             =   108
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "退費日期"
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
      Left            =   336
      TabIndex        =   18
      Top             =   108
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc21e0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Combo1、Text3、Text6、Text10、Text11
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc1e0 As New ADODB.Recordset
Public adoacc150 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strSerialNo As String

Private Sub Combo1_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
TextInverse Combo1  'Added by Lydia 2021/12/14 Form 2.0的ComboBox的GotFocus不會全選反白
End Sub

Private Sub Combo1_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo2, Label8) = False Then
      Cancel = True
      Combo2.SetFocus
   End If
End Sub

Private Sub Command1_Click()
On Error GoTo HandErr
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text5.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   AdodcDelete
   DataGrid1.Refresh
   'Add by Amy 2014/11/06 新增-acc1p0沒資料可以改退費日
   If strSaveConfirm = MsgText(3) Then
        If Adodc1.Recordset.RecordCount = 0 Then
            '當Acc1p0沒資料時,Acc1e0也需刪掉否則資再Insert 會出現資料已存在
            adoTaie.Execute "Delete From Acc1e0 Where a1e01 = '" & Text1 & "' and a1e02 = '" & Val(FCDate(MaskEdBox1.Text)) & "' "
            adoacc1e0.Requery
            MaskEdBox1.Enabled = True
        Else
            MaskEdBox1.Enabled = False
        End If
   End If
   'end 2014/11/06
   SumShow
   AdodcClear
   
HandErr:
    If Err.Number = 0 Then
      Exit Sub
    End If
    MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Command5_Click()
   If adoacc1e0.RecordCount = 0 Or Text1 = MsgText(601) Or MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   adoacc1e0.Find "a1e01 = '" & Text1 & "'", 0, adSearchForward, 1
   If adoacc1e0.EOF = False Then
      adoacc1e0.Find "a1e02 = " & Val(FCDate(MaskEdBox1.Text)) & "", 0, adSearchForward, adoacc1e0.Bookmark
      If adoacc1e0.EOF = False Then
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
         Exit Sub
      End If
   End If
   MsgBox MsgText(33), , MsgText(5)
   adoacc1e0.MoveFirst
   Frmacc21e0_Clear
   AdodcRefresh
   SumShow
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSerialNo = Adodc1.Recordset.Fields("a1p03").Value
   AdodcShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc1e0.RecordCount <> 0 Then
      adoacc1e0.MoveFirst
   End If
   adoacc1e0.Find "a1e01 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoacc1e0.EOF = False Then
      adoacc1e0.Find "a1e02 = '" & strCustNo & "'", 0, adSearchForward, adoacc1e0.Bookmark
      If adoacc1e0.EOF = False Then
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
   End If
   strItemNo = MsgText(601)
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
   'Modify by Amy 2023/08/18 H5800
   PUB_InitForm Me, 8850, 6000, strBackPicPath1
   'end 2021/12/07
   
   MaskEdBox1.Mask = DFormat
   OpenTable
   If adoacc1e0.RecordCount <> 0 Then
      adoacc1e0.MoveLast
      adoacc1e0.MoveFirst
      RecordShow
   End If
   FormDisabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   CreDebCheck
   If CreDebCheck <> MsgText(602) Then
      tool1_enabled
      MsgBox MsgText(11), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc21e0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc1e0.CursorLocation = adUseClient
   adoacc1e0.Open "select * from acc1e0 order by a1e02 asc, a1e01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'J' and a1p04 = '" & Text1 & Val(FCDate(MaskEdBox1.Text)) & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo2.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   Combo2 = "USD"
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
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0, acc010 where a1p05 = a0101 and a1p01 = '1' and a1p02 = 'J' and a1p04 = '" & Text1 & Val(FCDate(MaskEdBox1.Text)) & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   MaskEdBox1.Mask = MsgText(601)
   MaskEdBox1.Text = CFDate(adoacc1e0.Fields("a1e02").Value)
   MaskEdBox1.Mask = DFormat
   Text1 = adoacc1e0.Fields("a1e01").Value
   adoacc150.CursorLocation = adUseClient
   adoacc150.Open "select axf02, axf03, a1503 from acc151, acc150 where axf01 = a1501 and axf01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc150.RecordCount <> 0 Then
      If IsNull(adoacc150.Fields("axf03").Value) Then
         Text2 = ""
         Text10 = ""
      Else
         Text2 = adoacc150.Fields("axf03").Value
         Text10 = CaseNameQuery(adoacc150.Fields("axf02").Value, 1)
         If Text10 = "" Then
            Text10 = CaseNameQuery(adoacc150.Fields("axf02").Value, 2)
         End If
      End If
      If IsNull(adoacc150.Fields("a1503").Value) Then
         Text3 = ""
      Else
         Text3 = adoacc150.Fields("a1503").Value
      End If
   Else
      Text2 = ""
      Text3 = ""
      Text10 = ""
   End If
   adoacc150.Close
   If IsNull(adoacc1e0.Fields("a1e04").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = adoacc1e0.Fields("a1e04").Value
   End If
   If IsNull(adoacc1e0.Fields("a1e05").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc1e0.Fields("a1e05").Value
   End If
   If IsNull(adoacc1e0.Fields("a1e03").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc1e0.Fields("a1e03").Value
   End If
   If IsNull(adoacc1e0.Fields("a1e06").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc1e0.Fields("a1e06").Value
   End If
End Sub

'*************************************************
'  顯示查詢資料表(國外帳單資料)
'
'*************************************************
Private Sub Acc150Query()
   adoacc150.CursorLocation = adUseClient
   adoacc150.Open " select axf03, a1503, axf02, a1903, a1906, a1904, a1905 from acc151, acc150, acc190 where axf01 = a1501 and axf01 = a1902 and axf01 = '" & Text1 & "' order by axf02 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc150.RecordCount <> 0 Then
      If IsNull(adoacc150.Fields("axf03").Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = adoacc150.Fields("axf03").Value
      End If
      Text10 = CaseNameQuery(adoacc150.Fields("axf02").Value, 1)
      If Text10 = "" Then
         Text10 = CaseNameQuery(adoacc150.Fields("axf02").Value, 2)
      End If
      If IsNull(adoacc150.Fields("a1503").Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = adoacc150.Fields("a1503").Value
      End If
      If IsNull(adoacc150.Fields("a1903").Value) Then
         Combo2 = ""
      Else
         Combo2 = adoacc150.Fields("a1903").Value
      End If
      If IsNull(adoacc150.Fields("a1906").Value) Then
         Text9 = ""
      Else
         Text9 = adoacc150.Fields("a1906").Value
      End If
      If IsNull(adoacc150.Fields("a1904").Value) Then
         Text4 = ""
      Else
         Text4 = adoacc150.Fields("a1904").Value
      End If
      If IsNull(adoacc150.Fields("a1905").Value) Then
         Text13 = ""
      Else
         Text13 = Format(adoacc150.Fields("a1905").Value, FAmount)
      End If
   Else
      Text2 = MsgText(601)
      Text10 = MsgText(601)
      Text3 = MsgText(601)
      Text8 = ""
      Text9 = ""
      Text4 = ""
      Text13 = ""
   End If
   adoacc150.Close
End Sub

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
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         'Added by Lydia 2021/12/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
             Exit Sub
         End If
         'end 2021/12/07
         
         Frmacc21e0_Save
         'Add by Amy 2014/11/06 +新增按過insert 鎖住退費日
         If strSaveConfirm = MsgText(3) And MaskEdBox1.Enabled = True And strControlButton <> MsgText(602) Then
            MaskEdBox1.Enabled = False
         End If
         'end 2014/11/06
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            SumShow
            AdodcClear
            Text5.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox MsgText(10) & Label1, , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Then
      Acc150Query
   End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Text11_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
'add by nickc 2007/02/08
Dim adocase As New ADODB.Recordset
Set adocase = New ADODB.Recordset

On Error GoTo Checking
   If Text16 <> MsgText(601) Then
      Text16 = CaseNoZero(Text16)
      adocase.CursorLocation = adUseClient
      adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and pa02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and pa03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and pa04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and tm02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and tm03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and tm04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and lc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and lc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and lc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and hc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and hc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and hc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and sp02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and sp03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and sp04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase.RecordCount = 0 Then
         MsgBox MsgText(28) & Label17, , MsgText(5)
         Cancel = True
         adocase.Close
         Exit Sub
      End If
      adocase.Close
      'add by sonia 2021/1/28 以本所案號以判別FCP,FCT英日文組
      If AccNoToSalesNo(Text5, Text16) <> "" Then
         Text19 = AccNoToSalesNo(Text5, Text16)
      End If
      'end 2021/1/28
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   If Text17 <> MsgText(601) Then
      If Len(Text17) = 6 Then
         Text17 = AfterZero(Text17)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text17) = 8 Then
         Text17 = Text17 & "0"
      'End 2007/3/1
      End If
      If ExistCheck("customer", "cu01", Mid(Text17, 1, 8), Label18, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text17, Label18, False) = False Then
            If ExistCheck("staff", "st01", Text17, Label18, False) = False Then
               'Add by Morgan 2006/7/20 加抓代理人檔
               If ExistCheck("fagent", "fa01", Mid(Text17, 1, 8), Label18, False) = False Then
                  MsgBox MsgText(28) & Label18, , MsgText(5)
                  Cancel = True
                  Exit Sub
               End If
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

Private Sub Text19_GotFocus()
   TextInverse Text19
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text19_Validate(Cancel As Boolean)
   If Text19 <> MsgText(601) Then
      If ExistCheck("staff", "st01", Text19, Label15) = False Then
         Cancel = True
         Exit Sub
      End If
      'add by sonia 2021/1/28
      If SalesNoCheckAccNo(Text5, Text19) = False Then
      End If
      'end 2021/1/28
   End If
End Sub

Private Sub Text4_Change()
   If Text4 = MsgText(601) Then
      Exit Sub
   End If
   Text13 = Format(Val(Text9) * Val(Text4), FAmount)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_Change()
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   Text6 = A0102Query(Text5)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc010", "a0101", Text5, Label6) = False Then
      Cancel = True
      Exit Sub
   End If
   Text16 = Text2
   'add by sonia 2021/1/28 加傳本所案號以判別FCP,FCT英日文
   If AccNoToSalesNo(Text5, Text16) <> "" Then
      Text19 = AccNoToSalesNo(Text5, Text16)
   End If
   'end 2021/1/28
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text8, Label16) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Text5, Text8) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text9_Change()
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   Text13 = Format(Int(Val(Text9) * Val(Text4)), FAmount)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'*************************************************
'  儲存 Adodc 之資料
'
'*************************************************
Private Sub Acc1p0Save()
On Error GoTo Checking
      If Text5 = MsgText(601) Then
         MsgBox MsgText(10) & Label6, , MsgText(5)
         strControlButton = MsgText(602)
         Text5.SetFocus
         Exit Sub
      Else
         If ExistCheck("acc010", "a0101", Text5, Label6) = False Then
            strControlButton = MsgText(602)
            Text5.SetFocus
            Exit Sub
         End If
      End If
      If CheckDept(Text5, Text8) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text8.SetFocus
         Exit Sub
      End If
      If Text16 <> MsgText(601) Then
         Text16 = CaseNoZero(Text16)
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and pa02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and pa03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and pa04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                        "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and tm02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and tm03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and tm04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                        "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and lc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and lc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and lc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                        "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and hc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and hc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and hc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                        "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and sp02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and sp03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and sp04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount = 0 Then
            MsgBox MsgText(28) & Label17, , MsgText(5)
            strControlButton = MsgText(602)
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
      If Text17 <> MsgText(601) Then
         If Len(Text17) = 6 Then
            Text17 = AfterZero(Text17)
         'Add by Morgan 2007/3/1 八碼時要補'0'
         ElseIf Len(Text17) = 8 Then
            Text17 = Text17 & "0"
         'End 2007/3/1
         End If
         If ExistCheck("customer", "cu01", Mid(Text17, 1, 8), Label18, False) = False Then
            If ExistCheck("acc0i0", "a0i01", Text17, Label18, False) = False Then
               If ExistCheck("staff", "st01", Text17, Label18, False) = False Then
                  MsgBox MsgText(28) & Label18, , MsgText(5)
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
            End If
         End If
      End If
      If Text19 <> MsgText(601) Then
         If ExistCheck("staff", "st01", Text19, Label15) = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
      End If
      
      'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
      intI = PUB_AccNoEnable(Text5, Val(FCDate(MaskEdBox1.Text)))
      If intI <> 0 Then
         strControlButton = MsgText(602)
         Text5.SetFocus
         Exit Sub
      End If
      'end 2015/12/30
      'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
      intI = PUB_AccNoGood(Text5, Text8, Text19)
      If intI <> 0 Then
         strControlButton = MsgText(602)
         If intI = 1 Then
            Text5.SetFocus
         ElseIf intI = 2 Then
            Text8.SetFocus
         ElseIf intI = 3 Then
            Text19.SetFocus
         End If
         Exit Sub
      End If
      'end 2007/10/2
   
      If Adodc1.Recordset.RecordCount <> 0 Then
         If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               MsgBox MsgText(155), , MsgText(5)
               strControlButton = MsgText(602)
               Text5.SetFocus
               adoquery.Close
               Exit Sub
            End If
            adoquery.Close
         End If
      End If
      adoacc1p0.CursorLocation = adUseClient
      adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'J' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text1 & Val(FCDate(MaskEdBox1.Text)) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc1p0.RecordCount = 0 Then
         Adodc1.Recordset.AddNew
         Adodc1.Recordset.Fields("a1p01").Value = "1"
         Adodc1.Recordset.Fields("a1p02").Value = "J"
         Adodc1.Recordset.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'J' and a1p04 = '" & Text1 & Val(FCDate(MaskEdBox1.Text)) & "'", 3)
         Adodc1.Recordset.Fields("a1p04").Value = Text1 & Val(FCDate(MaskEdBox1.Text))
      End If
      adoacc1p0.Close
      Adodc1.Recordset.Fields("a1p05").Value = Text5
      If Text7 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a1p07").Value = Val(Text7)
      Else
         Adodc1.Recordset.Fields("a1p07").Value = 0
      End If
      If Text15 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a1p08").Value = Val(Text15)
      Else
         Adodc1.Recordset.Fields("a1p08").Value = 0
      End If
      If Combo1 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a1p14").Value = Combo1
         Combo1.AddItem Combo1
      Else
         Adodc1.Recordset.Fields("a1p14").Value = Null
      End If
      'modify by sonia 2021/1/27 加傳本所案號以判別FCP,FCT英日文組
      'If AccNoToSalesNo(Text5) = "" Then
      If AccNoToSalesNo(Text5, Text16) = "" Then
         Adodc1.Recordset.Fields("a1p16").Value = Null
      Else
         'modify by sonia 2021/1/27 加傳本所案號以判別FCP,FCT英日文組
         'Adodc1.Recordset.Fields("a1p16").Value = AccNoToSalesNo(Text5)
         Adodc1.Recordset.Fields("a1p16").Value = AccNoToSalesNo(Text5, Text16)
      End If
      If MaskEdBox1.Text <> MsgText(29) Then
         Adodc1.Recordset.Fields("a1p18").Value = Val(FCDate(MaskEdBox1.Text))
      End If
      If Text8 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a1p06").Value = Text8
      Else
         Adodc1.Recordset.Fields("a1p06").Value = MsgText(55)
      End If
      If Text16 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a1p17").Value = Text16
      Else
         Adodc1.Recordset.Fields("a1p17").Value = Null
      End If
      If Text17 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a1p15").Value = Text17
      Else
         Adodc1.Recordset.Fields("a1p15").Value = Null
      End If
      'Add by Amy 2014/11/06
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         Adodc1.Recordset.Fields("a1p27").Value = MsgText(602)
      End If
      'end 2014/11/06
      If Text18 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a1p30").Value = Text18
      Else
         Adodc1.Recordset.Fields("a1p30").Value = Null
      End If
      Adodc1.Recordset.UpdateBatch
      strSerialNo = MsgText(601)
      AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示 Adodc 之資料
'
'*************************************************
Public Sub AdodcShow()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Text5 = Adodc1.Recordset.Fields("a1p05").Value
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
      Text15 = MsgText(601)
   Else
      Text15 = Adodc1.Recordset.Fields("a1p08").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1p06").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) Then
      Text16 = MsgText(601)
   Else
      Text16 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text17 = MsgText(601)
   Else
      Text17 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text18 = MsgText(601)
   Else
      Text18 = Adodc1.Recordset.Fields("a1p30").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text19 = MsgText(601)
   Else
      Text19 = Adodc1.Recordset.Fields("a1p16").Value
   End If
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
   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'J' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text1 & Val(FCDate(MaskEdBox1.Text)) & "'"
   AdodcRefresh
   AdodcClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  清除 Adodc 之顯示資料
'
'*************************************************
Public Sub AdodcClear()
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text15 = ""
   Combo1 = ""
   Text8 = ""
   Text16 = ""
   Text17 = ""
   Text18 = ""
   Text19 = ""
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01 = '1' and a1p02 = 'J' and a1p04 = '" & Text1 & Val(FCDate(MaskEdBox1.Text)) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text14 = MsgText(601)
      Else
         Text14 = Format(adoaccsum.Fields(1).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = Format(adoaccsum.Fields(2).Value, DDollar)
      End If
   Else
      Text12 = MsgText(601)
      Text14 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc1e0.Bookmark & MsgText(35) & adoacc1e0.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text5.Enabled = False
   Text7.Enabled = False
   Combo1.Enabled = False
   Text8.Enabled = False
   Text16.Enabled = False
   Text17.Enabled = False
   Text18.Enabled = False
   Text19.Enabled = False
   Command1.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text5.Enabled = True
   Text7.Enabled = True
   Combo1.Enabled = True
   Text8.Enabled = True
   Text16.Enabled = True
   Text17.Enabled = True
   Text18.Enabled = True
   Text19.Enabled = True
   Command1.Enabled = True
End Sub

'*************************************************
'  借貸方檢核
'
'*************************************************
Public Function CreDebCheck() As String
   If Text12 = Text14 Then
      CreDebCheck = MsgText(602)
      Exit Function
   End If
   CreDebCheck = MsgText(603)
End Function

'Add by Amy 2014/11/04 由aacc_sav搬回
Public Sub Frmacc21e0_Save()
Dim strMsg As String 'Add by Amy 2014/11/06

    'Added by Lydia 2021/12/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        strControlButton = MsgText(602)
        Exit Sub
    End If
    'end 2021/12/07
    
On Error GoTo Checking
      If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
         MsgBox MsgText(10) & Label1, , MsgText(5)
         strControlButton = MsgText(602)
         MaskEdBox1.SetFocus
         Exit Sub
      Else
         If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
            MsgBox Label1 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Sub
         End If
         'Add by Amy 2014/11/06 +系統日檢查
         If MaskEdBox1.Enabled = True Then
            If ChkWorkData("1", DBDATE(MaskEdBox1), strMsg) = False Then
                MsgBox Label1 & strMsg, , MsgText(5)
                strControlButton = MsgText(602)
                MaskEdBox1.SetFocus
                Exit Sub
            End If
         End If
         'end 2014/11/06
         If Text1.Text = MsgText(601) Then
            MsgBox MsgText(10) & Label2, , MsgText(5)
            strControlButton = MsgText(602)
            Text1.SetFocus
            Exit Sub
         End If
      End If
      If strSaveConfirm = MsgText(3) Then
         If adoacc1e0.RecordCount <> 0 Then
           If Frmacc21e0.Text20 = "" Then '2011/4/26 ADD BY SONIA 第一項次時才檢查,否則其他項次也檢查則無法輸入
            adoacc1e0.Find "a1e01 = '" & Text1 & "'", 0, adSearchForward, 1
            If adoacc1e0.EOF = False Then
'2009/11/10 modify by sonia 一帳單不可退費二次 U09501730,否則退費日期不同會產生二筆傳票ACC1P0(U09501730981008,U09501730981007)
'               .adoacc1e0.Find "a1e02 = '" & Val(FCDate(.MaskEdBox1.Text)) & "'", 0, adSearchForward, .adoacc1e0.Bookmark
'               If .adoacc1e0.EOF = False Then
''                  MsgBox MsgText(9), , MsgText(5)
''                  strControlButton = MsgText(602)
''                  .MaskEdBox1.SetFocus
'                  Exit Sub
'               End If
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               MaskEdBox1.SetFocus
            '2011/4/26 ADD BY SONIA
            Else
               adoacc1e0.AddNew
            '2011/4/26 END
            End If
           End If
         End If
         '.adoacc1e0.AddNew   '2011/4/26 CANCEL BY SONIA
      End If
      adoacc1e0.Fields("a1e02").Value = Val(FCDate(MaskEdBox1.Text))
      adoacc1e0.Fields("a1e01").Value = Text1
      If Combo2 <> MsgText(601) Then
         adoacc1e0.Fields("a1e04").Value = Combo2
      Else
         adoacc1e0.Fields("a1e04").Value = Null
      End If
      If Text9 <> MsgText(601) Then
         adoacc1e0.Fields("a1e05").Value = Val(Text9)
      Else
         adoacc1e0.Fields("a1e05").Value = 0
      End If
      If Text4 <> MsgText(601) Then
         adoacc1e0.Fields("a1e03").Value = Val(Text4)
      Else
         adoacc1e0.Fields("a1e03").Value = 0
      End If
      If Text11 <> MsgText(601) Then
         adoacc1e0.Fields("a1e06").Value = Text11
      Else
         adoacc1e0.Fields("a1e06").Value = Null
      End If
      '2009/11/10 add by sonia
      If strSaveConfirm = MsgText(3) Then
         adoacc1e0.Fields("a1e07").Value = Val(strSrvDate(2))
         adoacc1e0.Fields("a1e08").Value = ServerTime
         adoacc1e0.Fields("a1e09").Value = strUserNum
      Else
         adoacc1e0.Fields("a1e10").Value = Val(strSrvDate(2))
         adoacc1e0.Fields("a1e11").Value = ServerTime
         adoacc1e0.Fields("a1e12").Value = strUserNum
      End If
      '2009/11/10 end
      adoacc1e0.UpdateBatch
      RecordShow
      Exit Sub
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2014/11/06
Public Sub SetData(Optional ByVal strKeyCode As String)
    Select Case strKeyCode
        Case "F2"
            Combo1.Clear
            FormEnabled
            MaskEdBox1.Enabled = True '只開放新增可以輸
        Case "F3"
            FormEnabled
            '因寫法關係 帳單編號+退費日期 寫入a1p04 所以鎖住退費日期
            MaskEdBox1.Enabled = False
        Case "F9", "F10"
            FormDisabled
            MaskEdBox1.Enabled = True
        Case Else
    End Select
End Sub
