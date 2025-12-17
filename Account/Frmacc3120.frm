VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc3120 
   AutoRedraw      =   -1  'True
   Caption         =   "開票作業"
   ClientHeight    =   5508
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5508
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
      Height          =   315
      Left            =   4068
      MaxLength       =   12
      TabIndex        =   17
      Top             =   4236
      Width           =   855
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
      Height          =   300
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   22
      Top             =   4908
      Width           =   540
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
      Height          =   300
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   21
      Top             =   4560
      Width           =   1572
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   1320
      TabIndex        =   20
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text23 
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
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   19
      Top             =   4224
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
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   49
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text21 
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
      Left            =   4440
      MaxLength       =   12
      TabIndex        =   48
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text22 
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
      Left            =   6000
      MaxLength       =   12
      TabIndex        =   47
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text9 
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
      Height          =   300
      Left            =   6840
      MaxLength       =   14
      TabIndex        =   15
      Top             =   3888
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2550
      Picture         =   "Frmacc3120.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   408
      Width           =   350
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   1320
      TabIndex        =   6
      Top             =   744
      Width           =   1572
   End
   Begin VB.TextBox Text16 
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
      Left            =   1920
      TabIndex        =   44
      Top             =   4224
      Width           =   972
   End
   Begin VB.TextBox Text15 
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
      MaxLength       =   3
      TabIndex        =   16
      Top             =   4224
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      Height          =   372
      Left            =   8000
      Picture         =   "Frmacc3120.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   3216
      Width           =   372
   End
   Begin VB.TextBox Text14 
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
      Height          =   300
      Left            =   5040
      MaxLength       =   14
      TabIndex        =   14
      Top             =   3888
      Width           =   1572
   End
   Begin VB.TextBox Text7 
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
      Left            =   360
      MaxLength       =   6
      TabIndex        =   13
      Top             =   3888
      Width           =   1572
   End
   Begin VB.TextBox Text6 
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
      Height          =   300
      Left            =   1320
      TabIndex        =   11
      Top             =   1416
      Width           =   1572
   End
   Begin VB.TextBox Text12 
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
      MaxLength       =   10
      TabIndex        =   0
      Top             =   72
      Width           =   1572
   End
   Begin VB.TextBox Text10 
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
      MaxLength       =   9
      TabIndex        =   7
      Top             =   744
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
      MaxLength       =   15
      TabIndex        =   4
      Top             =   408
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
      Height          =   315
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   1
      Top             =   72
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Height          =   300
      Left            =   4080
      MaxLength       =   14
      TabIndex        =   9
      Top             =   1080
      Width           =   1572
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   6480
      TabIndex        =   10
      Top             =   1080
      Width           =   1932
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   2
      Top             =   408
      Width           =   1220
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   6840
      TabIndex        =   5
      Top             =   408
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   16777215
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
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
      Left            =   1320
      Top             =   -48
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
      Bindings        =   "Frmacc3120.frx":076C
      Height          =   1392
      Left            =   240
      TabIndex        =   25
      Top             =   1800
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   2455
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a1p05"
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
         DataField       =   "a0102"
         Caption         =   "科目名稱"
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
         DataField       =   "a1p07"
         Caption         =   "借方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a1p08"
         Caption         =   "貸方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "a1p06"
         Caption         =   "部門別"
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
      BeginProperty Column06 
         DataField       =   "a1p17"
         Caption         =   "對沖代號(本所案號)"
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
            ColumnWidth     =   1284.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2580.095
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1548.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1548.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4559.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2064.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   312
      Left            =   240
      Top             =   1704
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
      Caption         =   "Adodc2"
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
   Begin MSForms.TextBox Text24 
      Height          =   315
      Left            =   4905
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4236
      Width           =   960
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   8
      Size            =   "4877;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   312
      Left            =   4080
      TabIndex        =   23
      Top             =   4920
      Width           =   4344
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7779;550"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   300
      Left            =   1920
      TabIndex        =   40
      Top             =   3888
      Width           =   2772
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4877;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text13 
      Height          =   300
      Left            =   2880
      TabIndex        =   38
      Top             =   72
      Width           =   2772
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4877;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text11 
      Height          =   300
      Left            =   5640
      TabIndex        =   37
      Top             =   744
      Width           =   2772
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4877;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   300
      Left            =   4080
      TabIndex        =   12
      Top             =   1416
      Width           =   4332
      VariousPropertyBits=   -1466941413
      MaxLength       =   35
      ScrollBars      =   2
      Size            =   "12524;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
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
      Height          =   252
      Left            =   3144
      TabIndex        =   56
      Top             =   4248
      Width           =   900
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "作帳公司"
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
      TabIndex        =   55
      Top             =   4932
      Width           =   960
   End
   Begin VB.Label Label19 
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
      Height          =   252
      Left            =   5928
      TabIndex        =   54
      Top             =   4572
      Width           =   912
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "款項類別"
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
      TabIndex        =   53
      Top             =   4584
      Width           =   972
   End
   Begin VB.Label Label24 
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
      Height          =   252
      Left            =   5916
      TabIndex        =   52
      Top             =   4236
      Width           =   900
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
      Height          =   252
      Left            =   360
      TabIndex        =   51
      Top             =   3240
      Width           =   852
   End
   Begin VB.Label Label23 
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
      Height          =   252
      Left            =   3480
      TabIndex        =   50
      Top             =   3240
      Width           =   732
   End
   Begin VB.Label Label17 
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
      Height          =   252
      Left            =   6960
      TabIndex        =   46
      Top             =   3648
      Width           =   1332
   End
   Begin VB.Label Label16 
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
      Height          =   252
      Left            =   3168
      TabIndex        =   45
      Top             =   4932
      Width           =   612
   End
   Begin VB.Label Label14 
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
      Height          =   252
      Left            =   360
      TabIndex        =   43
      Top             =   4224
      Width           =   972
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   1704
      Left            =   264
      Top             =   3588
      Width           =   8292
   End
   Begin VB.Label Label13 
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
      Height          =   252
      Left            =   5160
      TabIndex        =   42
      Top             =   3648
      Width           =   1332
   End
   Begin VB.Label Label12 
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
      Height          =   252
      Left            =   360
      TabIndex        =   41
      Top             =   3648
      Width           =   4332
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "手續費"
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
      TabIndex        =   39
      Top             =   1416
      Width           =   972
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "往來類別"
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
      TabIndex        =   36
      Top             =   744
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "單據號碼"
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
      TabIndex        =   35
      Top             =   408
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "開票日期"
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
      Left            =   5880
      TabIndex        =   34
      Top             =   408
      Width           =   972
   End
   Begin VB.Label Label5 
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
      Left            =   5880
      TabIndex        =   33
      Top             =   72
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "往來對象"
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
      TabIndex        =   32
      Top             =   744
      Width           =   972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
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
      TabIndex        =   31
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "票據金額"
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
      TabIndex        =   30
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "票別"
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
      Left            =   5880
      TabIndex        =   29
      Top             =   1080
      Width           =   612
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
      Left            =   360
      TabIndex        =   28
      Top             =   72
      Width           =   972
   End
   Begin VB.Label Label10 
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
      TabIndex        =   27
      Top             =   408
      Width           =   972
   End
   Begin VB.Label Label11 
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
      Left            =   3120
      TabIndex        =   26
      Top             =   1416
      Width           =   732
   End
End
Attribute VB_Name = "Frmacc3120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/09/01 Form2.0已修改 Text4/Text8/Text11/Text13/Text24/DataGrid1/Combo3
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc1p0 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adocase As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public strAccNumber As String
Dim strSerialNo As String
Public strLstAddDate As String 'Add by Morgan 2006/7/21

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo3_GotFocus()
OpenIme
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo3_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Command1_Click()
   If Adodc2.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc2.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc2.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc2.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text7.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   Adodc2Delete
   Adodc2Clear
End Sub

Private Sub Command2_Click()
   'Modify by Amy 2014/11/12
'   If Adodc1.Recordset.RecordCount = 0 Or Text12 = MsgText(601) Or Text5 = MsgText(601) Then
'      Exit Sub
'   End If
'   Adodc1.Recordset.Find "a0e01 = '" & Text12 & "'", 0, adSearchForward, 1
'   If Adodc1.Recordset.EOF = False Then
'      Adodc1.Recordset.Find "a0e02 = '" & Text5 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
'      If Adodc1.Recordset.EOF = False Then
'         FormShow
'         Adodc2Refresh
'         RecordShow
'      Else
'         MsgBox MsgText(33), , MsgText(5)
'         Adodc1.Recordset.MoveFirst
'      End If
'   Else
'      MsgBox MsgText(33), , MsgText(5)
'      Adodc1.Recordset.MoveFirst
'   End If
On Error GoTo Checking
    If Trim(Text12) = MsgText(601) Then
        MsgBox Label9.Caption & "必輸查詢條件", , MsgText(5)
        Exit Sub
    End If
    If Trim(Text5) = MsgText(601) Then
        MsgBox Label5.Caption & "必輸查詢條件", , MsgText(5)
        Exit Sub
    End If
    'Add by Amy 2020/07/14 +開票帳號 為必輸查詢條件
    If Trim(Text3) = MsgText(601) Then
        MsgBox Label10.Caption & "必輸查詢條件", , MsgText(5)
        Exit Sub
    End If
    adoadodc1.Close
    adoadodc1.CursorLocation = adUseClient
    'Modify by Amy 2020/07/14 +a0e07 因改為key
    strExc(0) = "Select * From acc0e0 Where a0e04 = '" & MsgText(19) & "' and a0e14 = 0 and a0e15 = 0 and a0e17 = 0 and a0e25 = 0 " & _
                     "and a0e01 = '" & Text12 & "' and  a0e02 = '" & Text5 & "' And a0e07='" & Text3 & "' " & "order by a0e01 asc, a0e02 asc"
    adoadodc1.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
    Adodc1.Recordset.Requery
     If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        FormShow
        Adodc2Refresh
        RecordShow
     Else
        MsgBox MsgText(33), , MsgText(5)
     End If
     
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
'end 2014/11/12
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command2_Click
         Exit Sub
   End Select
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSerialNo = Adodc2.Recordset.Fields("a1p03").Value
   Adodc2Show
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   'Modify by Morgan 2005/1/19 改重新抓資料庫
'   If Adodc1.Recordset.EOF Then
'      Adodc1.Recordset.MoveFirst
'   End If
'   Adodc1.Recordset.Find "a0e01 = '" & strCompanyNo & "'", 0, adSearchForward, 1
'   If Adodc1.Recordset.EOF = False Then
'      Adodc1.Recordset.Find "a0e02 = '" & strItemNo & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
'      If Adodc1.Recordset.EOF = False Then
'         FormShow
'         Adodc2Refresh
'         RecordShow
'      End If
'   End If
   'Modify by Amy 2020/07/14 +a0e07 因改為key
   strSql = "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e14 = 0 and a0e15 = 0 and a0e17 = 0 and a0e25 = 0 and a0e01 = '" & strCompanyNo & "' and a0e02 = '" & strItemNo & "' And a0e07='" & strBankAcc & "' order by a0e01 asc, a0e02 asc"
   If adoadodc1.State = adStateOpen Then adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.EOF = False Then
      FormShow
      Adodc2Refresh
      RecordShow
   End If
   '2005/1/19 end

   strCompanyNo = MsgText(601)
   strBankAcc = MsgText(601) 'Add by Amy 2020/07/16
End Sub

'Add by Amy 2021/09/01 避免於Form2.0 物件上按Insert 觸發2次 KeyDefine KeyCode事件
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   'Add by Morgan 2006/12/7
   Dim bCancel As Boolean
   If KeyCode = vbKeyInsert Then
      If Me.ActiveControl = Me.Text9 Then
         Text9_Validate bCancel
         If bCancel = True Then Exit Sub
      End If
   End If
   'end 2006/12/7
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
   Me.Height = 5950 'Modify by Amy 2023/08/16 原:5700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Combo1.AddItem ComboItem(131)
   Combo1.AddItem ComboItem(132)
   Combo1.AddItem ComboItem(133)
   Combo1.AddItem ComboItem(134)
   Combo2.AddItem ComboItem(11)
   Combo2.AddItem ComboItem(12)
   Combo2.AddItem ComboItem(13)
   Combo4.AddItem ComboItem(111)
   Combo4.AddItem ComboItem(112)
   Combo4.AddItem ComboItem(113)
   Combo4.AddItem ComboItem(114)
   Combo4.AddItem ComboItem(115)
   Combo4.AddItem ComboItem(116)
   Combo4.AddItem ComboItem(117)
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveLast
      Adodc1.Recordset.MoveFirst
      RecordShow
   End If
   FormDisabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   strTrackMode = "" 'Add by Amy 2021/09/01 Form2.0 記錄鍵盤傳入順序(清除)
   MenuEnabled
   Set Frmacc3120 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub
'Add by Morgan 2006/12/29 --瑞婷
Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   MaskEdBox2.Text = MaskEdBox1.Text
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      MsgBox Label7 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label7 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   Else
      'Modify by Morgan 2010/8/3 百年蟲
      'If FCDate(MaskEdBox2.Text) < FCDate(MaskEdBox1.Text) Then
      If Val(FCDate(MaskEdBox2.Text)) < Val(FCDate(MaskEdBox1.Text)) Then
         MsgBox MsgText(57), , MsgText(5)
         Cancel = True
         MaskEdBox2.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2005/1/19 改只抓一筆 and rownum<2
   adoadodc1.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e14 = 0 and a0e15 = 0 and a0e17 = 0 and a0e25 = 0 and rownum<2  order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   adoadodc2.CursorLocation = adUseClient
   'Modify by Amy 2020/07/14 a1p04 加開票帳號 因a0e07改為key,避免key重覆
   adoadodc2.Open "select * from acc1p0, acc010 where acc1p0.a1p05 = acc010.a0101 (+) and a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Text5 & Text12 & Text3 & "2" & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc2.Recordset = adoadodc2
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(票據資料--應付)
'
'*************************************************
Public Sub FormShow()
   If IsNull(Adodc1.Recordset.Fields("a0e03").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a0e03").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e13").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a0e13").Value)
   End If
   MaskEdBox1.Tag = "" & Adodc1.Recordset.Fields("a0e13").Value 'Add by Amy 2014/11/12
   MaskEdBox1.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e02").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("a0e02").Value
      Text5.Tag = Adodc1.Recordset.Fields("a0e02").Value   'add by sonia 2020/2/3
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e05").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Combo1.List(Val(Adodc1.Recordset.Fields("a0e05").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e06").Value) Then
      Text10 = MsgText(601)
      Text11 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a0e06").Value
      Select Case Mid(Combo1, 1, 1)
         Case Mid(ComboItem(131), 1, 1)
            If Len(Text10) = 6 Then
               Text11 = CustomerQuery(AfterZero(Text10), 1)
            Else
               Text11 = CustomerQuery(Text10, 1)
            End If
         Case Mid(ComboItem(132), 1, 1)
            Text11 = A0i02Query(Text10)
         Case Mid(ComboItem(133), 1, 1)
            Text11 = StaffQuery(Text10)
         Case Else
            Text11 = MsgText(601)
      End Select
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e36").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a0e36").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(Adodc1.Recordset.Fields("a0e10").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e11").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("a0e11").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e01").Value) Then
      Text12 = MsgText(601)
      Text13 = MsgText(601)
   Else
      Text12 = Adodc1.Recordset.Fields("a0e01").Value
      Text13 = A0g02Query(Text12)
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e08").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Combo2.List(Val(Adodc1.Recordset.Fields("a0e08").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e07").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0e07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e12").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0e12").Value
   End If
   'Add by Amy 2020/07/14 +記錄前次記錄
   Text12.Tag = Text12
   Text5.Tag = Text5
   Text3.Tag = Text3
   'end 2020/07/14
   Adodc2Clear
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Mid(Combo1, 1, 1) = Mid(ComboItem(134), 1, 1) Then
      Exit Sub
   End If
   If Text10 = MsgText(601) Then
      MsgBox Label6 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   Select Case Mid(Combo1, 1, 1)
      Case Mid(ComboItem(131), 1, 1)
         If ExistCheck("customer", "cu01", Mid(IIf(Len(Text10) = 6, AfterZero(Text10), Text10), 1, 8), Label6) = False Then
            Cancel = True
            Exit Sub
         End If
      Case Mid(ComboItem(132), 1, 1)
         If ExistCheck("acc0i0", "a0i01", Text10, Label6) = False Then
            Cancel = True
            Exit Sub
         End If
      Case Mid(ComboItem(133), 1, 1)
         If ExistCheck("staff", "st01", Text10, Label6) = False Then
            Cancel = True
            Exit Sub
         End If
   End Select
   Select Case Mid(Combo1, 1, 1)
      Case Mid(ComboItem(131), 1, 1)
         If Len(Text10) = 6 Then
            Text10 = AfterZero(Text10)
         ElseIf Len(Text10) = 8 Then
            Text10 = Text10 & "0"
         End If
         Text11 = CustomerQuery(Text10, 1)
      Case Mid(ComboItem(132), 1, 1)
         Text11 = A0i02Query(Text10)
      Case Mid(ComboItem(133), 1, 1)
         Text11 = StaffQuery(Text10)
      Case Else
         Text11 = MsgText(601)
   End Select
   Text4 = Text11
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text12, Label9) = False Then
      Cancel = True
      Exit Sub
   End If
   If strSaveConfirm = MsgText(3) Then
      adoaccsum.CursorLocation = adUseClient
      adoaccsum.Open "select a0h02 from acc0h0 where a0h01 = '" & Text12 & "' order by a0h02 asc", adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            Text3 = ""
         Else
            Text3 = adoaccsum.Fields(0).Value
            '2010/6/21 ADD BY SONIA 開票銀行011010075改預設0149951
            If Text12 = "011010075" Then
               Text3 = "1756650" 'Modify by Amy 2020/07/14 原:"0149951"   'modify by sonia 2020/6/19 改帳號原為0149951(1756650)
            End If
            '2010/6/21 END
         End If
      Else
         Text3 = ""
      End If
      adoaccsum.Close
   End If
   Text13 = A0g02Query(Text12)
   CheckCheck
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text15_Change()
   If Text15 = MsgText(601) Then
      Exit Sub
   End If
   Text16 = A0902Query(Text15)
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text15_Validate(Cancel As Boolean)
   If Text15 = "" Then
      Exit Sub
   End If
   If CheckDept(Text7, Text15) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   If ExistCheck("acc090", "a0901", Text15, Label14) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
'Add by Morgan 2007/12/31
Private Sub Text19_Validate(Cancel As Boolean)
   Text24.Text = ""
   If Text19 <> MsgText(601) Then
      If PUB_GetStaffState(Text19.Text, strExc(1), True) = 0 Then
         Cancel = True
         TextInverse Text19
      Else
         Text24.Text = strExc(1)
      End If
      'add by sonia 2021/1/28
      If SalesNoCheckAccNo(Text7, Text19) = False Then
      End If
      'end 2021/1/28
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Val(Text2) <= 0 Then
      MsgBox MsgText(58), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text23_GotFocus()
   TextInverse Text23
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text23_LostFocus()
   'add by nick 2004/07/01
On Error GoTo Checking
If Text23 <> MsgText(601) Then
         Dim strNation As String
      Text23 = CaseNoZero(Text23)
      adocase.CursorLocation = adUseClient
      adocase.Open "select pa01 as SystemNo,pa09,pa26 from patent where pa01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and pa02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and pa03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and pa04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo,tm10,tm23 from trademark where tm01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and tm02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and tm03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and tm04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo,lc15,lc11 from lawcase where lc01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and lc02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and lc03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and lc04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo,'000',hc07 from hirecase where hc01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and hc02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and hc03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and hc04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo,sp09,sp08 from servicepractice where sp01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and sp02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and sp03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and sp04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase.RecordCount > 0 Then
         '檢查當科目是 220112 220102 220111 220101 220103 220104 220105 220106 時，要檢查申請國家級系統別
         strNation = CheckStr(adocase.Fields(1).Value)
      End If
      adocase.Close
      
         Select Case Text7
         Case "220101"
                 'edit by nick 2004/07/07 加系統別
                 'If (Mid(Text23, 1, Len(Text23) - 9) = "T" Or Mid(Text23, 1, Len(Text23) - 9) = "TB") And strNation = "000" Then
                 If (Mid(Text23, 1, Len(Text23) - 9) = "T" Or Mid(Text23, 1, Len(Text23) - 9) = "TB" Or Mid(Text23, 1, Len(Text23) - 9) = "TS" Or Mid(Text23, 1, Len(Text23) - 9) = "TD" Or Mid(Text23, 1, Len(Text23) - 9) = "TM" Or Mid(Text23, 1, Len(Text23) - 9) = "TR" Or Mid(Text23, 1, Len(Text23) - 9) = "TT") And strNation = "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220102"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text23, 1, Len(Text23) - 9) = "P" And strNation = "000" Then
                 If (Mid(Text23, 1, Len(Text23) - 9) = "P" Or Mid(Text23, 1, Len(Text23) - 9) = "PS") And strNation = "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220103"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text23, 1, Len(Text23) - 9) = "FCT" Then
                 If (Mid(Text23, 1, Len(Text23) - 9) = "FCT" Or Mid(Text23, 1, Len(Text23) - 9) = "S") And strNation = "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220104"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text23, 1, Len(Text23) - 9) = "FCP" Then
                 If Mid(Text23, 1, Len(Text23) - 9) = "FCP" Or Mid(Text23, 1, Len(Text23) - 9) = "FG" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220105"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text23, 1, Len(Text23) - 9) = "CFT" Then
                 If (Mid(Text23, 1, Len(Text23) - 9) = "CFT" Or Mid(Text23, 1, Len(Text23) - 9) = "CFC" Or Mid(Text23, 1, Len(Text23) - 9) = "S") And strNation <> "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220106"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text23, 1, Len(Text23) - 9) = "CFP" Then
                 '2012/4/24 MODIFY BY SONIA 加系統類別 LIN
                 If Mid(Text23, 1, Len(Text23) - 9) = "CFP" Or Mid(Text23, 1, Len(Text23) - 9) = "FCL" Or Mid(Text23, 1, Len(Text23) - 9) = "LIN" Or Mid(Text23, 1, Len(Text23) - 9) = "CFL" Or Mid(Text23, 1, Len(Text23) - 9) = "CPS" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220107"
                 'add by nick 2004/07/07 加系統別
                 If Mid(Text23, 1, Len(Text23) - 9) = "TC" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220111"
                 'edit by nick 2004/07/07 加系統別
                 'If (Mid(Text23, 1, Len(Text23) - 9) = "T" Or Mid(Text23, 1, Len(Text23) - 8) = "TF") And strNation <> "000" Then
                 '2012/4/24 modify by sonia 加系統類別TC,TD,TM,TT,TB,TR
                 If (Mid(Text23, 1, Len(Text23) - 9) = "TS" Or Mid(Text23, 1, Len(Text23) - 9) = "T" Or Mid(Text23, 1, Len(Text23) - 8) = "TF" Or Mid(Text23, 1, Len(Text23) - 8) = "TC" Or Mid(Text23, 1, Len(Text23) - 8) = "TD" Or Mid(Text23, 1, Len(Text23) - 8) = "TM" Or Mid(Text23, 1, Len(Text23) - 8) = "TT" Or Mid(Text23, 1, Len(Text23) - 8) = "TB" Or Mid(Text23, 1, Len(Text23) - 8) = "TT") And strNation <> "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220112"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text23, 1, Len(Text23) - 9) = "P" And strNation <> "000" Then
                 If (Mid(Text23, 1, Len(Text23) - 9) = "P" Or Mid(Text23, 1, Len(Text23) - 9) = "PS") And strNation <> "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         '2012/4/24 ADD BY SONIA
         Case "220108"
                 If (Mid(Text23, 1, Len(Text23) - 9) = "P" Or Mid(Text23, 1, Len(Text23) - 9) = "PS") Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case "220113"
                 If Mid(Text23, 1, Len(Text23) - 9) = "L" Or Mid(Text23, 1, Len(Text23) - 9) = "LA" Or Mid(Text23, 1, Len(Text23) - 9) = "FCL" Or Mid(Text23, 1, Len(Text23) - 9) = "LIN" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         '2012/4/24 END
         Case "610103"
                 'add by nick 2004/07/07 加系統別
                 '2012/4/24 MODIFY BY SONIA 加系統類別LIN
                 If Mid(Text23, 1, Len(Text23) - 9) = "L" Or Mid(Text23, 1, Len(Text23) - 9) = "LA" Or Mid(Text23, 1, Len(Text23) - 9) = "FCL" Or Mid(Text23, 1, Len(Text23) - 9) = "LIN" Or Mid(Text23, 1, Len(Text23) - 9) = "CFL" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text7.SetFocus
                       Text7.SelStart = 0
                       Text7.SelLength = Len(Text7)
                       Exit Sub
                 End If
         Case Else
         End Select
End If
Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text23_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text23 <> MsgText(601) Then
      Text23 = CaseNoZero(Text23)
      adocase.CursorLocation = adUseClient
      'nick edit by 2004/07/01
      'adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and pa02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and pa03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and pa04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and tm02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and tm03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and tm04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and lc02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and lc03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and lc04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and hc02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and hc03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and hc04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and sp02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and sp03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and sp04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      adocase.Open "select pa01 as SystemNo,pa09,pa26 from patent where pa01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and pa02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and pa03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and pa04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo,tm10,tm23 from trademark where tm01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and tm02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and tm03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and tm04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo,lc15,lc11 from lawcase where lc01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and lc02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and lc03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and lc04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo,'000',hc07 from hirecase where hc01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and hc02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and hc03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and hc04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo,sp09,sp08 from servicepractice where sp01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and sp02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and sp03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and sp04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase.RecordCount = 0 Then
                MsgBox MsgText(28) & Label24, , MsgText(5)
                Cancel = True
                adocase.Close
                Exit Sub
      Else
                If Combo3 = MsgText(601) Then
                   'Combo3 = CustomerQuery(CaseCustQuery(adocase.Fields("cp09").Value), 1)
                   Select Case Text7
                      Case "2111"
                         Combo3 = Text4
                   End Select
                Else
                   'Combo3 = Combo3 & "/" & CustomerQuery(CaseCustQuery(adocase.Fields("cp09").Value), 1)
                   Select Case Text7
                      Case "2111"
                         Combo3 = Combo3 & "/" & Text4
                   End Select
                End If
               'add by sonia 2021/1/28 以本所案號以判別FCP,FCT英日文組
               If AccNoToSalesNo(Text7, Text23) <> "" Then
                  Text19 = AccNoToSalesNo(Text7, Text23)
                  Text19_Validate True
               End If
               'end 2021/1/28
      
                '2004/07/01 nick
                '針對   P  T  TF  CFT  CFP  加入客戶名稱
                '  FCT  FCP 加入本所案號
                '這些案號的摘要 , 要在前面加入資訊
                Select Case Mid(Text23, 1, Len(Text23) - 9)
                'Modify by Morgan 2005/8/11 加TC--瑞婷
                '2006/10/27 MODIFY BY SONIA 加入所有T字頭系統類別及PS,CFC,CPS--瑞婷
                Case "P", "PS", "T", "TF", "CFT", "CFC", "CFP", "CPS", "TC", "TB", "TD", "TM", "TR", "TS", "TT"
                       Dim strCustomer As String
                       strCustomer = CheckStr(adocase.Fields(2).Value)
                       CheckOC3
                       AdoRecordSet3.CursorLocation = adUseClient
                       AdoRecordSet3.Open "SELECT cu04 FROM Customer " & _
                      "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                            "CU02 = '" & Mid(strCustomer, 9, 1) & "'", cnnConnection, adOpenStatic, adLockReadOnly
                       If AdoRecordSet3.RecordCount > 0 Then
                              Combo3 = CheckStr(AdoRecordSet3.Fields(0).Value) & "/" & Combo3
                       End If
                       AdoRecordSet3.Close
                Case "FCT", "FCP", "FG"    '2006/10/27 加入FG
                        Combo3 = Text23 & "/" & Combo3
                Case Else
                End Select
      End If
      adocase.Close
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'Added by Morgan 2011/11/4 -- 瑞婷
Private Sub CheckCheck()
   If Text12 <> "" And Text3 <> "" Then
      'modify by sonia 2020/6/19 改帳號原為0149951(1756650)
      'Modify by Amy 2020/07/14 原:Text3 <> "0149951"
      If Text12 <> "011010075" Or Text3 <> "1756650" Then
         MsgBox "目前為開票作業但您輸入的銀行帳號不為" & vbCrLf & vbCrLf & "【011010075】【1756650】，請檢核!!", vbExclamation
      End If
   End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   Dim strQ As String 'Add by Amy 2020/07/14
   
   If Text3 = MsgText(601) Then
      MsgBox Label10 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   If CheckData("select a0h01 from acc0h0 where a0h01 = '" & Text12 & "' and a0h02 = '" & Text3 & "'", Label10) = False Then
      Cancel = True
      Exit Sub
   End If
   CheckCheck
   
   'Add by Amy 從Text5_Validate搬過來,+修改也判斷
   If Text5 = MsgText(601) Then
      MsgBox MsgText(10) & Label5, , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If Text12 & Text5 & Text3 <> Text12.Tag & Text5.Tag & Text3.Tag Then
        If adoquery.State = adStateOpen Then
           adoquery.Close
        End If
        adoquery.CursorLocation = adUseClient
        adoquery.Open "select a0e01 from acc0e0 where a0e01 = '" & Text12 & "' and a0e02 = '" & Text5 & "' And a0e07='" & Text3 & "' ", adoTaie, adOpenStatic, adLockReadOnly
        If adoquery.RecordCount <> 0 Then
           '此處應不必控制否則存檔後會工具列無法回復
           'strControlButton = MsgText(602)
           MsgBox MsgText(9), , MsgText(5)
           adoquery.Close
           Cancel = True
           Text5.SetFocus
           Exit Sub
        End If
        adoquery.Close
      End If
   End If
   UpdateNo
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text4_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2020/07/14 改至 開票帳號 判斷
'Private Sub Text5_Validate(Cancel As Boolean)
'   If Text5 = MsgText(601) Then
'      MsgBox MsgText(10) & Label5, , MsgText(5)
'      Cancel = True
'      Exit Sub
'   End If
'   If strSaveConfirm = MsgText(3) Then
'      If adoquery.State = adStateOpen Then
'         adoquery.Close
'      End If
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select a0e01 from acc0e0 where a0e01 = '" & Text12 & "' and a0e02 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         'Modify by Morgan 2005/11/4 此處應不必控制否則存檔後會工具列無法回復
'         'strControlButton = MsgText(602)
'         MsgBox MsgText(9), , MsgText(5)
'         adoquery.Close
'         Cancel = True
'         Text5.SetFocus
'         Exit Sub
'      End If
'      adoquery.Close
'   End If
'   UpdateNo
'End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  重新整理 Adodc1 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2005/1/19 改只抓一筆 and rownum<2
   adoadodc1.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e14 = 0 and a0e15 = 0 and a0e17 = 0 and a0e25 = 0 and rownum<2 order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

'*************************************************
'  重新整理 Adodc2 之資料
'
'*************************************************
Public Sub Adodc2Refresh()
On Error GoTo Checking
   adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   'Modify by Amy 2020/07/13 a1p04 加開票帳號 因a0e07改為key,避免key重覆
   adoadodc2.Open "select * from acc1p0, acc010 where acc1p0.a1p05 = acc010.a0101 (+) and a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Text5 & Text12 & Text3 & "2" & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc2.Recordset.Requery
   'Add by Amy 2014/11/12 +鎖/開放 開票日
   If Adodc2.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc2.Recordset.Fields("a1p22").Value) Then
         strAccNumber = "null"
         MaskEdBox1.Enabled = True
      Else
         strAccNumber = "'" & Adodc2.Recordset.Fields("a1p22").Value & "'"
         MaskEdBox1.Enabled = False
      End If
   Else
      strAccNumber = "null"
      MaskEdBox1.Enabled = True
   End If
   'end 2014/11/12
   TotalShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(國內收據資料(分錄檔))
'
'*************************************************
Private Sub Acc1p0Save()
Dim strCombo3 As String

On Error GoTo Checking
   If Text7 = MsgText(601) Then
      MsgBox MsgText(10) & Label12, , MsgText(5)
      strControlButton = MsgText(602)
      Text7.SetFocus
      Exit Sub
   Else
      If Val(Text14) <> 0 And Val(Text9) <> 0 Then
         MsgBox MsgText(47) & MsgText(46), , MsgText(5)
         strControlButton = MsgText(602)
         Text14.SetFocus
         Exit Sub
      End If
      If ExistCheck("acc010", "a0101", Text7, Label12) = False Then
         strControlButton = MsgText(602)
         Text7.SetFocus
         Exit Sub
      End If
      If CheckDept(Text7, Text15) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text15.SetFocus
         Exit Sub
      End If
      If Text15 <> MsgText(601) Then
         If ExistCheck("acc090", "a0901", Text15, Label14) = False Then
            strControlButton = MsgText(602)
            Text15.SetFocus
            Exit Sub
         End If
      End If
      'Add by Amy 2021/09/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
            strControlButton = MsgText(602)
            Exit Sub
      End If
      If Val(Text14) <= 0 And Val(Text9) <= 0 Then
         MsgBox MsgText(58), , MsgText(5)
         strControlButton = MsgText(602)
         Text14.SetFocus
         Exit Sub
      End If
      If Text23 <> MsgText(601) Then
         adocase.CursorLocation = adUseClient
         adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and pa02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and pa03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and pa04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and tm02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and tm03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and tm04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and lc02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and lc03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and lc04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and hc02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and hc03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and hc04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text23, 1, Len(Text23) - 9) & "' and sp02 = '" & Mid(Text23, Len(Text23) - 8, 6) & "' and sp03 = '" & Mid(Text23, Len(Text23) - 2, 1) & "' and sp04 = '" & Mid(Text23, Len(Text23) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocase.RecordCount = 0 Then
            MessageShow Label24
            strControlButton = MsgText(602)
            adocase.Close
            Text23.SetFocus
            Exit Sub
         End If
         adocase.Close
      End If
   End If
   
   'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(Text7, Val(FCDate(MaskEdBox1.Text)))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      Text7.SetFocus
      Exit Sub
   End If
   'end 2015/12/30
   '2007/8/8 ADD BY SONIA 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Text7, Text15, Text19)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Text7.SetFocus
      ElseIf intI = 2 Then
         Text15.SetFocus
      ElseIf intI = 3 Then
         Text19.SetFocus
      End If
      Exit Sub
   End If
   '2007/8/8 END
   If Adodc2.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc2.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc2.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc2.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            strControlButton = MsgText(602)
            Text7.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   adoacc1p0.CursorLocation = adUseClient
   'Modify by Amy 2020/07/14 a1p04 加開票帳號 因a0e07改為key,避免key重覆
   adoacc1p0.Open "select * from acc1p0, acc010 where acc1p0.a1p05 = acc010.a0101 and a1p01 = '1' and a1p02 = 'L' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text5 & Text12 & Text3 & "2" & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1p0.RecordCount = 0 Then
      adoacc1p0.AddNew
      adoacc1p0.Fields("a1p01").Value = "1"
      adoacc1p0.Fields("a1p02").Value = "L"
      adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Text5 & Text12 & Text3 & "2" & "' ", 3)
      strSerialNo = adoacc1p0.Fields("a1p03").Value
      adoacc1p0.Fields("a1p04").Value = Text5 & Text12 & Text3 & "2"
   End If
   'end 2020/07/14
   adoacc1p0.Fields("a1p05").Value = Text7
   If Text14 <> MsgText(601) Then
      adoacc1p0.Fields("a1p07").Value = Val(Text14)
   Else
      adoacc1p0.Fields("a1p07").Value = 0
   End If
   If Text9 <> MsgText(601) Then
      adoacc1p0.Fields("a1p08").Value = Val(Text9)
   Else
      adoacc1p0.Fields("a1p08").Value = 0
   End If
   If Text15 <> MsgText(601) Then
      adoacc1p0.Fields("a1p06").Value = Text15
   Else
      adoacc1p0.Fields("a1p06").Value = MsgText(55)
   End If
   adoacc1p0.Fields("a1p09").Value = Text5
   adoacc1p0.Fields("a1p10").Value = Text12
   If Text3 <> MsgText(601) Then
      adoacc1p0.Fields("a1p11").Value = Text3
   Else
      adoacc1p0.Fields("a1p11").Value = Null
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p12").Value = Val(FCDate(MaskEdBox2.Text))
   Else
      adoacc1p0.Fields("a1p12").Value = Null
   End If
   If Combo3 <> MsgText(601) Then
      adoacc1p0.Fields("a1p14").Value = Combo3
      strCombo3 = Combo3
      Combo3.Clear
      Combo3.AddItem strCombo3
   Else
      adoacc1p0.Fields("a1p14").Value = Null
   End If
   If Text10 <> MsgText(601) Then
      adoacc1p0.Fields("a1p15").Value = Text10
   End If
   'modify by sonia 2021/1/28 加傳本所案號以判別FCP,FCT英日文組
   'If AccNoToSalesNo(Text7) <> "" Then
   '   adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text7)
   If AccNoToSalesNo(Text7, Text23) <> "" Then
      adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text7, Text23)
   'end 2021/1/28
   Else
      If Text19 <> MsgText(601) Then
         adoacc1p0.Fields("a1p16").Value = Text19
      Else
         adoacc1p0.Fields("a1p16").Value = Null
      End If
   End If
   If Text23 <> MsgText(601) Then
      adoacc1p0.Fields("a1p17").Value = Text23
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p18").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc1p0.Fields("a1p18").Value = Null
   End If
   If Combo4 <> MsgText(601) Then
      adoacc1p0.Fields("a1p26").Value = Mid(Combo4, 1, 1)
   Else
      adoacc1p0.Fields("a1p26").Value = Null
   End If
   If IsNull(adoacc1p0.Fields("a1p22").Value) = False Then
      adoacc1p0.Fields("a1p27").Value = MsgText(602)
   End If
   If Text17 <> MsgText(601) Then
      adoacc1p0.Fields("a1p30").Value = Text17
   Else
      adoacc1p0.Fields("a1p30").Value = Null
   End If
   If Text18 <> MsgText(601) Then
      adoacc1p0.Fields("a1p31").Value = Text18
   Else
      adoacc1p0.Fields("a1p31").Value = Null
   End If
   adoacc1p0.UpdateBatch
   Adodc2Refresh
   Adodc2.Recordset.Find "a1p03 = '" & strSerialNo & "'", 0, adSearchForward, 1
   If Adodc2.Recordset.EOF Then
      Adodc2.Recordset.MoveFirst
   End If
   strSerialNo = MsgText(601)
   adoacc1p0.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示 Adodc2 之資料
'
'*************************************************
Private Sub Adodc2Show()
   Text7 = Adodc2.Recordset.Fields("a1p05").Value
   If IsNull(Adodc2.Recordset.Fields("a1p07").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = Adodc2.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc2.Recordset.Fields("a1p08").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc2.Recordset.Fields("a1p08").Value
   End If
   If IsNull(Adodc2.Recordset.Fields("a1p06").Value) Then
      Text15 = MsgText(601)
   Else
      Text15 = Adodc2.Recordset.Fields("a1p06").Value
   End If
   If IsNull(Adodc2.Recordset.Fields("a1p14").Value) Then
      Combo3 = MsgText(601)
   Else
      Combo3 = Adodc2.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc2.Recordset.Fields("a1p26").Value) Then
      Combo4 = MsgText(601)
   Else
      Combo4 = Combo4.List(Val(Adodc2.Recordset.Fields("a1p26").Value) - 1)
   End If
   If IsNull(Adodc2.Recordset.Fields("a1p17").Value) Then
      Text23 = MsgText(601)
   Else
      Text23 = Adodc2.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc2.Recordset.Fields("a1p30").Value) Then
      Text17 = MsgText(601)
   Else
      Text17 = Adodc2.Recordset.Fields("a1p30").Value
   End If
   If IsNull(Adodc2.Recordset.Fields("a1p31").Value) Then
      Text18 = MsgText(601)
   Else
      Text18 = Adodc2.Recordset.Fields("a1p31").Value
   End If
   If IsNull(Adodc2.Recordset.Fields("a1p16").Value) Then
      Text19 = MsgText(601)
   Else
      Text19 = Adodc2.Recordset.Fields("a1p16").Value
   End If
End Sub

'*************************************************
'  刪除 Adodc2 之資料
'
'*************************************************
Private Sub Adodc2Delete()
On Error GoTo Checking
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Modify by Amy 2020/07/14 a1p04 加開票帳號 因a0e07改為key,避免key重覆
   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text5 & Text12 & Text3 & "2" & "' and a1p05 = '" & Text7 & "' "
   Adodc2Refresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Public Sub Adodc2Clear()
   Text7 = ""
   Text8 = ""
   Text14 = ""
   Text9 = ""
   Text15 = ""
   Text16 = ""
   Combo3 = ""
   Text23 = ""
   Text17 = ""
   Text18 = ""
   Text19 = ""
   Combo4 = ""
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Add by Amy 2021/09/01 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         'Add by Amy 2021/09/01 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'Frmacc3120_Save
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            Adodc2Clear
            Text7.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text7_Change()
   If Text7 = MsgText(601) Then
      Exit Sub
   End If
   Text8 = A0102Query(Text7)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text7.Enabled = False
   Text14.Enabled = False
   Text9.Enabled = False
   Text15.Enabled = False
   Combo3.Enabled = False
   Text17.Enabled = False
   Text18.Enabled = False
   Text19.Enabled = False
   Text23.Enabled = False
   Command1.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text7.Enabled = True
   Text14.Enabled = True
   Text9.Enabled = True
   Text15.Enabled = True
   Combo3.Enabled = True
   Text17.Enabled = True
   Text18.Enabled = True
   Text19.Enabled = True
   Text23.Enabled = True
   Command1.Enabled = True
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc010", "a0101", Text7, Label12) = False Then
      Cancel = True
      Exit Sub
   End If
'   If Mid(Text7, 1, 1) = "2" Then
'      Combo3 = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text3 & "/" & Text13
'      Text4 = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text3 & "/" & Text13
'      Exit Sub
'   End If
   Select Case Text7
      Case "2111"
         Combo3 = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text4
'         Text4 = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text11
         Exit Sub
   End Select
'   If Text7 = "110201" Or Text7 = "110214" Then
'      Combo3 = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text11
'      Text4 = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text11
'      Exit Sub
'   End If
   'add by sonia 2021/1/28 以本所案號以判別FCP,FCT英日文組
   If AccNoToSalesNo(Text7, Text23) <> "" Then
      Text19 = AccNoToSalesNo(Text7, Text23)
      Text19_Validate True
   End If
   'end 2021/1/28
End Sub

'*************************************************
'  合計計算並判斷是否等於票據金額
'
'*************************************************
Public Function SumShow() As String
   adoaccsum.CursorLocation = adUseClient
   'Modify by Amy 2020/07/14 a1p04 加開票帳號 因a0e07改為key,避免key重覆
   adoaccsum.Open "select sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Text5 & Text12 & Text3 & "2" & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         If Val(adoaccsum.Fields(0).Value) = (Val(Text2) + Val(Text6)) Then
            SumShow = MsgText(602)
            adoaccsum.Close
            Exit Function
         End If
      End If
   End If
   MsgBox MsgText(59), , MsgText(5)
   SumShow = MsgText(603)
   adoaccsum.Close
End Function

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub TotalShow()
   adoaccsum.CursorLocation = adUseClient
   'Modify by Amy 2020/07/14 a1p04 加開票帳號 因a0e07改為key,避免key重覆
   adoaccsum.Open "select count(*), sum(a1p07), sum(a1p08) from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Text5 & Text12 & Text3 & "2" & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text21 = MsgText(601)
      Else
         Text21 = Format(adoaccsum.Fields(1).Value, FDollar)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text22 = MsgText(601)
      Else
         Text22 = Format(adoaccsum.Fields(2).Value, FDollar)
      End If
   Else
      Text20 = MsgText(601)
      Text21 = MsgText(601)
      Text22 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = "" Then
      Exit Sub
   End If
   Select Case Text7
      Case "2111"
         Combo3 = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text4
'         Text4 = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text11
         'add by nick 2004/07/05 檢查當 2111 時，金額要跟貸方金額相同
         If Val(Text2.Text) <> Val(Text9.Text) Then
              MsgBox "貸方金額輸入錯誤!!", , "User 輸入錯誤!!"
              Cancel = True
         End If
   End Select
End Sub

'*************************************************
'  更改票號
'
'*************************************************
Public Sub UpdateNo()
   If strSaveConfirm = MsgText(4) Then
      If Adodc2.Recordset.RecordCount <> 0 Then
         Adodc2.Recordset.MoveFirst
         Do While Adodc2.Recordset.EOF = False
            'Modify by Amy 2020/07/14 a1p04 加開票帳號 因a0e07改為key,避免key重覆
            Adodc2.Recordset.Fields("a1p04").Value = Text5 & Text12 & Text3 & "2"
            Select Case Text7
               Case "2111"
                  Adodc2.Recordset.Fields("a1p14").Value = FCDate(MaskEdBox2.Text) & "/" & Text5 & "/" & Text4
            End Select
            Adodc2.Recordset.UpdateBatch
            Adodc2.Recordset.MoveNext
         Loop
         Adodc2.Recordset.MoveFirst
      End If
   End If
End Sub

'Add by Amy 2014/11/05 由aacc_sav搬回
Public Sub Frmacc3120_Save()
Dim adocheck As New ADODB.Recordset
'Add by Amy 2014/11/12
Dim strUpd As String
Dim strMsg As String

   On Error GoTo Checking
   With Frmacc3120
      If .Text12 = MsgText(601) Then
         MsgBox MsgText(10) & .Label9, , MsgText(5)
         strControlButton = MsgText(602)
         .Text12.SetFocus
         Exit Sub
      Else
         If .Text5 = MsgText(601) Then
            MsgBox MsgText(10) & .Label5, , MsgText(5)
            strControlButton = MsgText(602)
            .Text5.SetFocus
            Exit Sub
         End If
         If .Text21 <> .Text22 Then
            MsgBox MsgText(11), , MsgText(5)
            strControlButton = MsgText(602)
            Exit Sub
         End If
         If .Text3 = MsgText(601) Then
            MsgBox .Label10 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .Text3.SetFocus
            Exit Sub
         Else
            adocheck.CursorLocation = adUseClient
            adocheck.Open "select a0h01, a0h02 from acc0h0 where a0h01 = '" & .Text12 & "' and a0h02 = '" & .Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adocheck.RecordCount = 0 Then
               MessageShow .Label10
               strControlButton = MsgText(602)
               adocheck.Close
               .Text3.SetFocus
               Exit Sub
            End If
            adocheck.Close
         End If
         If ExistCheck("acc0g0", "a0g01", .Text12, .Label9) = False Then
            strControlButton = MsgText(602)
            .Text12.SetFocus
            Exit Sub
         End If
         adocheck.CursorLocation = adUseClient
         'Modify by Amy 2020/07/14 a1p04 加開票帳號 因a0e07改為key,避免key重覆
         adocheck.Open "select ax210 from acc1p0, acc021 where a1p22 = ax202 and a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text5 & .Text12 & Text3 & "6" & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount <> 0 Then
            MsgBox MsgText(180), , MsgText(5)
            strControlButton = MsgText(602)
            .Text12.SetFocus
            adocheck.Close
            Exit Sub
         End If
         adocheck.Close
         If .Combo1 <> MsgText(601) Then
            Select Case Mid(.Combo1, 1, 1)
               Case Mid(ComboItem(131), 1, 1)
                  If ExistCheck("customer", "cu01 || cu02", IIf(Len(.Text10) = 6, AfterZero(.Text10), .Text10), .Label6) = False Then
                     strControlButton = MsgText(602)
                     .Text10.SetFocus
                     Exit Sub
                  End If
               Case Mid(ComboItem(132), 1, 1)
                  If ExistCheck("acc0i0", "a0i01", .Text10, .Label6) = False Then
                     strControlButton = MsgText(602)
                     .Text10.SetFocus
                     Exit Sub
                  End If
               Case Mid(ComboItem(133), 1, 1)
                  If ExistCheck("staff", "st01", .Text10, .Label6) = False Then
                     strControlButton = MsgText(602)
                     .Text10.SetFocus
                     Exit Sub
                  End If
            End Select
         End If
         'Add by Amy 2014/11/12 +開票日檢查
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label2 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox .Label2 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
            If MaskEdBox1.Enabled = True Then
                If ChkWorkData("1", DBDATE(MaskEdBox1), strMsg) = False Then
                    MsgBox Label2 & strMsg, , MsgText(5)
                    strControlButton = MsgText(602)
                    MaskEdBox1.SetFocus
                    Exit Sub
                End If
            End If
         End If
         'end 2014/11/12
         If .MaskEdBox2.Text = MsgText(601) Or .MaskEdBox2.Text = MsgText(29) Then
            MsgBox .Label7 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox2.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox2.Text) = MsgText(603) Then
               MsgBox .Label7 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox2.SetFocus
               Exit Sub
            Else
               If FCDate(.MaskEdBox2.Text) < FCDate(.MaskEdBox1.Text) Then
                  MsgBox MsgText(57), , MsgText(5)
                  strControlButton = MsgText(602)
                  .MaskEdBox2.SetFocus
                  Exit Sub
               End If
            End If
         End If
         'Add by Amy 2021/09/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me) = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
         If Val(.Text2) <= 0 Then
            MsgBox MsgText(58), , MsgText(5)
            strControlButton = MsgText(602)
            .Text2.SetFocus
            Exit Sub
         End If
      End If
      If strSaveConfirm = MsgText(3) Then
         If .adoquery.State = adStateOpen Then
            .adoquery.Close
         End If
         .adoquery.CursorLocation = adUseClient
         'Modify by Amy 2020/07/14 +a0e07 因改為key
         .adoquery.Open "select a0e01 from acc0e0 where a0e01 = '" & .Text12 & "' and a0e02 = '" & .Text5 & "' And a0e07='" & .Text3 & "' ", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount <> 0 Then
            strControlButton = MsgText(602)
            MsgBox MsgText(9), , MsgText(5)
            .adoquery.Close
            Exit Sub
         End If
         .adoquery.Close
'         .Adodc1.Recordset.Requery
'         If .Adodc1.Recordset.RecordCount <> 0 Then
'            .Adodc1.Recordset.Find "a0e01 = '" & .Text12 & "'", 0, adSearchForward, 1
'            If .Adodc1.Recordset.EOF = False Then
'               .Adodc1.Recordset.Find "a0e02 = '" & .Text5 & "'", 0, adSearchForward, .Adodc1.Recordset.Bookmark
'               If .Adodc1.Recordset.EOF = False Then
'                  strControlButton = MsgText(602)
'                  Exit Sub
'               End If
'            End If
'         End If
         .Adodc1.Recordset.AddNew
      End If
      If .Text1 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0e03").Value = .Text1
      Else
         .Adodc1.Recordset.Fields("a0e03").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .Adodc1.Recordset.Fields("a0e13").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .Adodc1.Recordset.Fields("a0e13").Value = Null
      End If
      .Adodc1.Recordset.Fields("a0e02").Value = .Text5
      .Adodc1.Recordset.Fields("a0e23").Value = "1"     'ADD BY SONIA 2014/1/20
      If .Combo1 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0e05").Value = Mid(.Combo1, 1, 1)
      Else
         .Adodc1.Recordset.Fields("a0e05").Value = Null
      End If
      If .Text10 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0e06").Value = .Text10
      Else
         .Adodc1.Recordset.Fields("a0e06").Value = Null
      End If
      If .Text6 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0e36").Value = Val(.Text6)
      Else
         .Adodc1.Recordset.Fields("a0e36").Value = 0
      End If
      If .MaskEdBox2.Text <> MsgText(601) And .MaskEdBox2.Text <> MsgText(29) Then
         .Adodc1.Recordset.Fields("a0e10").Value = Val(FCDate(.MaskEdBox2.Text))
      Else
         .Adodc1.Recordset.Fields("a0e10").Value = Null
      End If
      If .Text2 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0e11").Value = Val(.Text2)
      Else
         .Adodc1.Recordset.Fields("a0e11").Value = Null
      End If
      .Adodc1.Recordset.Fields("a0e01").Value = .Text12
      If .Combo2 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0e08").Value = Mid(.Combo2, 1, 1)
      Else
         .Adodc1.Recordset.Fields("a0e08").Value = Null
      End If
      If .Text3 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0e07").Value = .Text3
      Else
         .Adodc1.Recordset.Fields("a0e07").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0e12").Value = .Text4
      Else
         .Adodc1.Recordset.Fields("a0e12").Value = Null
      End If
      .Adodc1.Recordset.Fields("a0e04").Value = MsgText(19)
      If strSaveConfirm = MsgText(3) Then
         .Adodc1.Recordset.Fields("a0e14").Value = 0
         .Adodc1.Recordset.Fields("a0e15").Value = 0
         .Adodc1.Recordset.Fields("a0e16").Value = 0
         .Adodc1.Recordset.Fields("a0e17").Value = 0
         .Adodc1.Recordset.Fields("a0e21").Value = 0
         .Adodc1.Recordset.Fields("a0e22").Value = 0
         .Adodc1.Recordset.Fields("a0e25").Value = 0
      End If
      If strSaveConfirm = MsgText(3) Then
         .Adodc1.Recordset.Fields("a0e26").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a0e27").Value = ServerTime
         .Adodc1.Recordset.Fields("a0e28").Value = strUserNum
      Else
         .Adodc1.Recordset.Fields("a0e29").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a0e30").Value = ServerTime
         .Adodc1.Recordset.Fields("a0e31").Value = strUserNum
         'add by sonia 2020/2/3 改票據號碼要清列印次數
         If Text5 <> Text5.Tag Then
            .Adodc1.Recordset.Fields("a0e46").Value = 0
         End If
         'end 2020/2/3
      End If
      'Modfy by Amy 2014/11/12 +更新a1p18
      'Modify by Amy 2020/07/14 a1p04 加開票帳號 因a0e07改為key,避免key重覆
      If strSaveConfirm = MsgText(3) Then
         strUpd = "Update acc1p0 set a1p18=" & Val(FCDate(MaskEdBox1)) & " where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text5 & .Text12 & Text3 & "2' "
         adoTaie.Execute strUpd
      ElseIf strSaveConfirm = MsgText(4) Then
         If Val(MaskEdBox1.Tag) <> Val(FCDate(MaskEdBox1)) Then strUpd = ",a1p18='" & Val(FCDate(MaskEdBox1)) & "' "
         adoTaie.Execute "update acc1p0 set a1p22 = " & .strAccNumber & strUpd & " where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text5 & .Text12 & Text3 & "2' "
         adoTaie.Execute "update acc1p0 set a1p27 = decode(a1p22, null, null, '" & MsgText(602) & "') where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text5 & .Text12 & Text3 & "2' "
      End If
      .Adodc1.Recordset.UpdateBatch
      .UpdateNo
      .RecordShow
      adoTaie.CommitTrans
      'Add by Morgan 2006/7/21
      If strSaveConfirm = MsgText(3) Then
         .strLstAddDate = .MaskEdBox1.Text
      End If
Checking:
   If Err.Number = 0 Or Err.Number = -2147168242 Then
      Exit Sub
   End If
   MsgBox Err.Number
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'Add by Amy 2020/07/14 從aacc_var搬回修改
Public Function DelCheck() As Boolean
    DelCheck = False
    'Add by Morgan 2005/5/13 控制有託收日期,退票日期,抽票日期,貼現日期,兌現日期,兌領日期,轉出日期,調整日期
    If Val("" & adoadodc1.Fields("A0E14")) > 0 Then
        MsgBox "已有託收日期，不可刪除！", vbExclamation
        Exit Function
    ElseIf Val("" & adoadodc1.Fields("A0E15")) > 0 Then
        MsgBox "已有退票日期，不可刪除！", vbExclamation
        Exit Function
    ElseIf Val("" & adoadodc1.Fields("A0E16")) > 0 Then
        MsgBox "已有抽票日期，不可刪除！", vbExclamation
        Exit Function
    ElseIf Val("" & adoadodc1.Fields("A0E17")) > 0 Then
        MsgBox "已有貼現日期，不可刪除！", vbExclamation
        Exit Function
    ElseIf Val("" & adoadodc1.Fields("A0E21")) > 0 Then
        MsgBox "已有兌現日期，不可刪除！", vbExclamation
        Exit Function
    ElseIf Val("" & adoadodc1.Fields("A0E37")) > 0 Then
        MsgBox "已有兌領日期，不可刪除！", vbExclamation
        Exit Function
    ElseIf Val("" & adoadodc1.Fields("A0E34")) > 0 Then
        MsgBox "已有轉出日期，不可刪除！", vbExclamation
        Exit Function
    ElseIf Val("" & adoadodc1.Fields("A0E22")) > 0 Then
        MsgBox "已有調整日期，不可刪除！", vbExclamation
        Exit Function
    End If
    '2005/5/13 END
            
    If Adodc2.Recordset.RecordCount <> 0 Then
        If IsNull(Adodc2.Recordset.Fields("a1p22").Value) = False Then
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc2.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc2.Recordset.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
                MsgBox MsgText(158), , MsgText(5)
                Text12.SetFocus
                adoquery.Close
                Exit Function
            End If
            adoquery.Close
        End If
    End If
    
    DelCheck = True
End Function

'從aacc_del搬回修改
Public Sub Frmacc3120_Delete()
On Error GoTo Checking
    'Modify by Amy 2020/07/14 +a0e07 改為key
    If DeleteCheck("select a0e01 from acc0e0 where a0e01 = '" & Text12 & "' and a0e02 = '" & Text5 & "' And a0e07='" & Text3 & "' ") = MsgText(603) Then
        Exit Sub
    End If

'Add by Morgan 2004/9/23 加transaction
On Error GoTo Checking2

    adoTaie.BeginTrans
    '2004/9/23 end
    
    'Modify by Amy 2020/07/14 a1p04 加開票帳號及a0e07 因a0e07改為key,避免key重覆
    adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Text5 & Text12 & Text3 & "2" & "'"
    adoTaie.Execute "delete from acc0e0 where a0e01 = '" & Text12 & "' and a0e02 = '" & Text5 & "' And a0e07='" & Text3 & "' "
    'end 2020/07/14
      
    'Add by Morgan 2004/9/23 若是開給智慧局的票同時要清除送件資料
    If Text10.Text = "V0001" Then
        'Add by Morgan 2005/7/26 改更新 AppList
        strSql = "update applist set al06=null" & _
            " where al01=" & ChangeTStringToWString(ChangeTDateStringToTString(MaskEdBox1)) & _
            " and al06='" & Text5.Text & "'"
        
        adoTaie.Execute strSql
         
        strSql = "update applistE set al06=null" & _
            " where al01=" & ChangeTStringToWString(ChangeTDateStringToTString(MaskEdBox1)) & _
            " and al06='" & Text5.Text & "'"
    
        adoTaie.Execute strSql
         
    End If
      
    adoTaie.CommitTrans
      
On Error GoTo Checking3

    '2004/9/23 end

    AdodcRefresh
    Adodc2Refresh
    If Adodc1.Recordset.RecordCount <> 0 Then
        Adodc1.Recordset.MoveFirst
        RecordShow
    Else
        StatusClear
    End If
   
   
'Add by Morgan 2004/9/23
Checking2:
   If Err.Number <> 0 Then
      adoTaie.RollbackTrans
      MsgBox Err.Description, , MsgText(5)
      Err.Clear
   End If
Checking3:
'2004/9/23 end

Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   
End Sub

'從aacc_cls搬回修改
Public Sub Frmacc3120_Clear()
    Text1 = ""
    'Modify by Morgan 2006/7/21 新增帶上一筆的開票日期--瑞婷
    MaskEdBox1.Mask = ""
    If strSaveConfirm = MsgText(3) And strLstAddDate <> "" Then
        MaskEdBox1.Text = strLstAddDate
    Else
         MaskEdBox1.Text = CFDate(ACDate(ServerDate))
    End If
    MaskEdBox1.Mask = DFormat
    Text5 = ""
    Combo1 = ""
    Text10 = ""
    Text6 = ""
    Text11 = ""
    MaskEdBox2.Mask = ""
    'Modify by Morgan 2005/7/25 預設當天--瑞婷
    '.MaskEdBox2.Text =""
    MaskEdBox2.Text = CFDate(ACDate(ServerDate))
    MaskEdBox2.Mask = DFormat
    Text2 = ""
    Text12 = "011010075"
    Text13 = ""
    'Modify by Morgan 2004/12/3 預設支票
    '.Combo2 = ""
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
    Text3 = ""
    Text4 = ""
    Adodc2Refresh
    Adodc2Clear
    Text12.SetFocus
   End Sub


