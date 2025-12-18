VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21t0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "客製化請款項目資料維護"
   ClientHeight    =   6144
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6144
   ScaleWidth      =   8760
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   6255
      Picture         =   "Frmacc21t0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   450
      Width           =   350
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8100
      MaxLength       =   5
      TabIndex        =   3
      Top             =   450
      Width           =   405
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   300
      Left            =   5895
      Picture         =   "Frmacc21t0.frx":015A
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   450
      Width           =   350
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   450
      Top             =   3930
      Visible         =   0   'False
      Width           =   1200
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
      Bindings        =   "Frmacc21t0.frx":025C
      Height          =   2604
      Left            =   96
      TabIndex        =   18
      Top             =   3456
      Width           =   8568
      _ExtentX        =   15113
      _ExtentY        =   4593
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "a2602"
         Caption         =   "系統類別"
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
         DataField       =   "a2603"
         Caption         =   "項目代號"
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
         DataField       =   "a2604"
         Caption         =   "TASK_CODE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "####-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a2605"
         Caption         =   "EXPENSE_CODE"
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
         DataField       =   "a2606"
         Caption         =   "ACTIVITY_CODE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "####-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "a2616"
         Caption         =   "對方代碼"
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
         DataField       =   "EngDesc"
         Caption         =   "英文名稱"
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
         DataField       =   "flag"
         Caption         =   "客製"
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
            ColumnWidth     =   515.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   984.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   887.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   984.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2592
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   552.189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "TASK_CODE(分割案用)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   19
      Left            =   5112
      TabIndex        =   42
      Top             =   3096
      Width           =   2124
   End
   Begin MSForms.TextBox Text1 
      Height          =   312
      Index           =   23
      Left            =   7344
      TabIndex        =   41
      Top             =   3048
      Width           =   1176
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "2064;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Classification"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   18
      Left            =   459
      TabIndex        =   40
      Top             =   3108
      Width           =   1176
   End
   Begin MSForms.TextBox Text1 
      Height          =   312
      Index           =   22
      Left            =   1668
      TabIndex        =   17
      Top             =   3048
      Width           =   1200
      VariousPropertyBits=   671107099
      MaxLength       =   10
      Size            =   "2117;550"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   21
      Left            =   3330
      TabIndex        =   13
      Top             =   2370
      Width           =   495
      VariousPropertyBits=   671107099
      Size            =   "873;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   20
      Left            =   4815
      TabIndex        =   14
      Top             =   2370
      Width           =   3690
      VariousPropertyBits=   671107099
      Size            =   "6509;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   19
      Left            =   5625
      TabIndex        =   16
      Top             =   2700
      Width           =   2880
      VariousPropertyBits=   671107099
      Size            =   "5080;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   18
      Left            =   1665
      TabIndex        =   15
      Top             =   2700
      Width           =   2880
      VariousPropertyBits=   671107099
      Size            =   "5080;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   17
      Left            =   1665
      TabIndex        =   12
      Top             =   2370
      Width           =   1215
      VariousPropertyBits=   671107099
      Size            =   "2143;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   315
      Left            =   4590
      TabIndex        =   34
      Top             =   1110
      Width           =   3900
      VariousPropertyBits=   671107103
      BackColor       =   14737632
      Size            =   "6879;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   16
      Left            =   1665
      TabIndex        =   32
      Top             =   1110
      Width           =   1845
      VariousPropertyBits=   671107099
      Size            =   "3254;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1665
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      VariousPropertyBits=   671107099
      BackColor       =   12648447
      MaxLength       =   8
      Size            =   "2778;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   7200
      TabIndex        =   8
      Top             =   780
      Width           =   1305
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "2302;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   4320
      TabIndex        =   7
      Top             =   780
      Width           =   1350
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "2381;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   1665
      TabIndex        =   6
      Top             =   780
      Width           =   1170
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "2064;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   300
      Left            =   3240
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   120
      Width           =   5250
      VariousPropertyBits=   671107103
      BackColor       =   14737632
      Size            =   "9260;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   1665
      TabIndex        =   11
      Top             =   2040
      Width           =   6825
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "12039;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   1665
      TabIndex        =   10
      Top             =   1740
      Width           =   6825
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "12039;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   1665
      TabIndex        =   9
      Top             =   1440
      Width           =   6825
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "12039;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   4305
      TabIndex        =   2
      Top             =   450
      Width           =   1575
      VariousPropertyBits=   671107099
      BackColor       =   12648447
      MaxLength       =   6
      Size            =   "2778;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1665
      TabIndex        =   1
      Top             =   450
      Width           =   1575
      VariousPropertyBits=   671107099
      BackColor       =   12648447
      MaxLength       =   3
      Size            =   "2778;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   2925
      TabIndex        =   39
      Top             =   2430
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   3870
      TabIndex        =   38
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   4635
      TabIndex        =   37
      Top             =   2760
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   675
      TabIndex        =   36
      Top             =   2760
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Timerkeeper Id"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   225
      TabIndex        =   35
      Top             =   2400
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "英文名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   3690
      TabIndex        =   33
      Top             =   1170
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "對方代碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   225
      TabIndex        =   31
      Top             =   1140
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客製"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   7545
      TabIndex        =   30
      Top             =   510
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款對象"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   29
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "EXPENSE_CODE"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2880
      TabIndex        =   28
      Top             =   825
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "ACTIVITY_CODE"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5715
      TabIndex        =   27
      Top             =   825
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "TASK_CODE"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   225
      TabIndex        =   26
      Top             =   810
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   3384
      Left            =   96
      Top             =   24
      Width           =   8568
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   24
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   23
      Top             =   1740
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "特殊英文名稱"
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
      Index           =   6
      Left            =   225
      TabIndex        =   21
      Top             =   1470
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "項目代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   20
      Top             =   450
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "系統類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   19
      Top             =   450
      Width           =   840
   End
End
Attribute VB_Name = "Frmacc21t0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (DataGrid1,Text2,txtDesc,Text1)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Create by Morgan 2010/11/19
Option Explicit
Dim iCurState As Integer '0:無資料,1:查詢,2:修改
Dim lstCol As String, lstAsc As String

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If iCurState = 1 Then
      If pRecordset.BOF Then
         MsgBox "已經是最後一筆!"
         Adodc1.Recordset.MoveFirst
      ElseIf pRecordset.EOF Then
         MsgBox "已經是第一筆!"
         Adodc1.Recordset.MoveLast
      Else
         RecordShow
      End If
   End If
End Sub

Private Sub Command1_Click()
   If Text1(1) = "" Then
      MsgBox "請輸入請款對象！"
      Text1(1).SetFocus
      Exit Sub
   Else
      AdodcRefresh
   End If
End Sub

Private Sub Command2_Click()
   FormClear True
   FormLock False
   Text1(1).SetFocus
   tool3_enabled
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Dim strAsc As String
   If iCurState = 1 Then
      If lstCol = DataGrid1.Columns(ColIndex).DataField Then
         If lstAsc = " Asc" Then
            strAsc = " Desc"
         Else
            strAsc = " Asc"
         End If
         lstAsc = strAsc
      Else
         lstCol = DataGrid1.Columns(ColIndex).DataField
         lstAsc = strAsc
      End If
      DataGrid1.Visible = False
      Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & strAsc & "," & DataGrid1.Columns(0).DataField & "," & DataGrid1.Columns(1).DataField
      DataGrid1.Visible = True
   End If
End Sub

Private Sub Form_Activate()
   SetTool
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyHome, vbKeyEnd
      
      Case Else
         KeyEnter KeyCode
   End Select
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   Set Frmacc21t0 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 2, 3, 4, 5, 6
         KeyAscii = UpperCase(KeyAscii)
      Case 21 'Rate
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Dim stTmp As String, stTmp1 As String
   Select Case Index
      Case 1
         If Text1(1) <> "" Then
            stTmp = Text1(1)
            If Len(stTmp) = 6 Then
               'Modified by Morgan 2023/5/16
               'stTmp = stTmp & "000"
               stTmp = stTmp & "00"
               'end 2023/5/16
            End If
            If ChkNo(stTmp, stTmp1) Then
               Text1(1) = stTmp
               Text2 = stTmp1
            Else
               MsgBox "請款對象輸入錯誤!!"
               Cancel = True
            End If
         End If
      Case 2
         If Text1(2) <> "" Then
            If InStr("," & Systemkind_g, "," & Text1(2) & ",") = 0 Then
               MsgBox "您沒有使用該系統類別的權限"
               Cancel = True
            End If
         End If
   End Select
End Sub
'檢查請款對象
Private Function ChkNo(pNo As String, Optional pName As String) As Boolean
   'Modify by Morgan 2011/9/21 改只輸入 8 碼(更名視為相同)
   If Left(pNo, 1) = "Y" Then
      strExc(0) = "select nvl(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),nvl(fa04,fa06))" & _
         " from fagent where fa01='" & Left(pNo, 8) & "' and fa02='0'"
   Else
      strExc(0) = "select nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),nvl(cu04,cu06))" & _
         " from customer where cu01='" & Left(pNo, 8) & "' and cu02='0'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ChkNo = True
      pName = "" & RsTemp(0)
   End If
End Function

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Private Sub AdodcRefresh()
   Dim stCon As String
   
   iCurState = 0
   SetTool
   
   If Systemkind_g <> "" Then
      stCon = " and a1j01 in (" & "'" & Join(Split(Left(Systemkind_g, Len(Systemkind_g) - 1), ","), "','") & "'" & ")"
   End If
   '系統類別
   If Text1(2) <> "" Then
      stCon = stCon & " and a1j01='" & Text1(2) & "'"
   End If
   '項目代號
   If Text1(3) <> "" Then
      stCon = stCon & " and a1j02='" & Text1(3) & "'"
   End If
'   'TASK_CODE
'   If Text1(4) <> "" Then
'      stCon = stCon & " and decode(a2601,null,a1j18,a2604)='" & Text1(4) & "'"
'   End If
'   'EXPENSE_CODE
'   If Text1(5) <> "" Then
'      stCon = stCon & " and decode(a2601,null,a1j19,a2605)='" & Text1(5) & "'"
'   End If
'   'ACTIVITY_CODE
'   If Text1(6) <> "" Then
'      stCon = stCon & " and decode(a2601,null,a1j20,a2606)='" & Text1(6) & "'"
'   End If
'   If Text1(7) <> "" Then
'      stCon = stCon & " and decode(a2601,null,a1j04,a2607)='" & ChgSQL(Text1(7)) & "'"
'   End If
'   If Text1(8) <> "" Then
'      stCon = stCon & " and decode(a2601,null,a1j05,a2608)='" & ChgSQL(Text1(8)) & "'"
'   End If
'   If Text1(9) <> "" Then
'      stCon = stCon & " and decode(a2601,null,a1j06,a2609)='" & ChgSQL(Text1(9)) & "'"
'   End If

   '客製
   If Text3 <> "" Then
      stCon = stCon & " and a2601 is not null"
   End If
   
   'Modified by Morgan 2014/2/20 英文名稱不同時才要寫特殊英文名稱,否則若有調整時會不同步
   'Modified by Morgan 2018/8/28 +,a2617,a2618,a2619,a2620,a2621
   'Modified by Morgan 2023/6/16 +a2622
   'Modified by Morgan 2025/1/16 +a2623
   strExc(0) = "select '" & Text1(1) & "' a2601,a1j01 a2602,a1j02 a2603" & _
      ",decode(a2601,null,a1j18,a2604) a2604" & _
      ",decode(a2601,null,a1j19,a2605) a2605" & _
      ",decode(a2601,null,a1j20,a2606) a2606" & _
      ",a2607,a2608,a2609,decode(a2607,null,rtrim(a1j04||' '||a1j05||' '||a1j06),rtrim(a2607||' '||a2608||' '||a2609)) EngDesc,rtrim(a1j04||' '||a1j05||' '||a1j06) NormDesc" & _
      ",decode(a2601,null,null,'Y') flag,a2616,a2617,a2618,a2619,a2620,a2621,a2622,a2623" & _
      " from acc1j0,acc260 where a2602(+)=a1j01 and a2603(+)=a1j02" & _
      " and a2601(+)='" & Text1(1) & "'" & stCon & " order by 1,2,3"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/25 +FormName 改暫存TB
   Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   If RsTemp.RecordCount > 0 Then
      'RecordShow
      CountShow 0, Adodc1.Recordset.RecordCount
      iCurState = 1
      tool6_enabled
      Adodc1.Recordset.MoveFirst
   Else
      MsgBox "無符合資料！"
      CountShow 0, 0
   End If
End Sub

Public Sub RecordShow()
   Dim txt As Object
   
   FormClear
   
   If Adodc1.Recordset Is Nothing Then Exit Sub
   
   With Adodc1.Recordset
   If .RecordCount > 0 Then
      For Each txt In Text1
         txt = "" & .Fields("a26" & Format(txt.Index, "00"))
      Next
      Text3 = "" & .Fields("flag")
      txtDesc = "" & .Fields("NormDesc") 'Added by Morgan 2014/2/20
      Text1_Validate 1, False
      CountShow Adodc1.Recordset.Bookmark, Adodc1.Recordset.RecordCount
      DataGrid1.Enabled = True
   Else
      CountShow 0, 0
   End If
   End With
   SetTool
   Command1.Enabled = True
   Command2.Enabled = True
   FormLock False
End Sub

Private Sub FormClear(Optional bAll As Boolean)
   Dim txt As Object
   For Each txt In Text1
      Select Case txt.Index
         Case 1
'Removed by Morgan 2024/10/21 請款對象都不要清除,這樣查詢較方便
'            If bAll Then
'               txt = ""
'               Text2 = ""
'            End If
         Case 2, 3
            If bAll Then
               txt = ""
            End If
         Case Else
            txt = ""
      End Select
   Next
   Text3 = ""
   txtDesc = ""
End Sub

Public Sub FormLock(bLock As Boolean, Optional bEdit As Boolean)
   Dim txt As Object
   For Each txt In Text1
      txt.Locked = bLock
   Next
   Text3.Enabled = Not bLock
   If bEdit Then
      Text1(1).Locked = True
      Text1(2).Locked = True
      Text1(3).Locked = True
      Text3.Enabled = False
   End If
End Sub

Private Sub SetTool()
   ToolShow
   Select Case iCurState
      Case 0
         tool3_enabled
      Case 1 '查詢
         tool16_enabled
      Case 2 '修改
         tool2_enabled
   End Select
End Sub

Public Sub SetEdit()
   FormLock False, True
   DataGrid1.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = False
   Text1(4).SetFocus
End Sub

Public Sub CancelEdit()
   RecordShow
End Sub

Public Function DeleteRec() As Boolean
   If Text3 <> "Y" Then
      MsgBox "非客製化資料，無需刪除！"
      Exit Function
   End If
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   strSql = "delete acc260  where a2601='" & Text1(1) & "' and a2602='" & Text1(2) & "' and a2603='" & Text1(3) & "'"
   cnnConnection.Execute strSql, intI
   
   strSql = "select a1j18 a2604,a1j19 a2605,a1j20 a2606,'' a2607,'' a2608,'' a2609 from acc1j0 where a1j01='" & Text1(2) & "' and a1j02='" & Text1(3) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
      For intI = 0 To .Fields.Count - 1
         Adodc1.Recordset.Fields(.Fields(intI).Name).Value = .Fields(intI)
      Next
      End With
      Adodc1.Recordset.Fields("flag") = ""
      'Modify by Amy 2014/06/25
      'Adodc1.Recordset.UpdateBatch
      Adodc1.Recordset.UPDATE
   End If
   
   cnnConnection.CommitTrans
  
   DeleteRec = True
   RecordShow
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Public Function SaveRec() As Boolean
   Dim txt As Object
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd
   'Modified by Morgan 2018/8/28 +a2617,a2618,a2619,a2620,a2621
   'Modified by Morgan 2023/6/16 +a2622
   'Modified by Morgan 2025/1/16 +a2623
   strSql = "update acc260 set a2604='" & Text1(4) & "',a2605='" & Text1(5) & "',a2606='" & Text1(6) & "',a2607='" & ChgSQL(Text1(7)) & "',a2608='" & ChgSQL(Text1(8)) & "',a2609='" & ChgSQL(Text1(9)) & "',a2616='" & Text1(16) & "',a2617='" & ChgSQL(Text1(17)) & "',a2618='" & ChgSQL(Text1(18)) & "',a2619='" & ChgSQL(Text1(19)) & "',a2620='" & ChgSQL(Text1(20)) & "',a2621=" & Val(Text1(21)) & ",a2622='" & Text1(22) & "',a2623='" & Text1(23) & "' where a2601='" & Text1(1) & "' and a2602='" & Text1(2) & "' and a2603='" & Text1(3) & "'"
   cnnConnection.Execute strSql, intI
   If intI = 0 Then
      strSql = "insert into acc260 (a2601,a2602,a2603,a2604,a2605,a2606,a2607,a2608,a2609,a2616,a2617,a2618,a2619,a2620,a2621,a2622,a2623) values ('" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & "','" & Text1(5) & "','" & Text1(6) & "','" & ChgSQL(Text1(7)) & "','" & ChgSQL(Text1(8)) & "','" & ChgSQL(Text1(9)) & "','" & Text1(16) & "','" & ChgSQL(Text1(17)) & "','" & ChgSQL(Text1(18)) & "','" & ChgSQL(Text1(19)) & "','" & ChgSQL(Text1(20)) & "'," & Val(Text1(21)) & ",'" & Text1(22) & "','" & Text1(23) & "')"
      cnnConnection.Execute strSql, intI
   End If
   
   For Each txt In Text1
      Adodc1.Recordset.Fields("a26" & Format(txt.Index, "00")).Value = txt
   Next
   Adodc1.Recordset.Fields("flag") = "Y"
   'Modify by Amy 2014/06/25
   'Adodc1.Recordset.UpdateBatch
   Adodc1.Recordset.UPDATE
   
   cnnConnection.CommitTrans

   SaveRec = True
   RecordShow
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   strControlButton = MsgText(602)
   
End Function

Public Sub MoveFirst()
   Adodc1.Recordset.MoveFirst
End Sub

Public Sub MoveNext()
   Adodc1.Recordset.MoveNext
End Sub

Public Sub MoveLast()
   Adodc1.Recordset.MoveLast
End Sub

Public Sub MovePrevious()
   Adodc1.Recordset.MovePrevious
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
