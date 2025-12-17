VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4170 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "每月固定傳票資料"
   ClientHeight    =   5424
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5424
   ScaleWidth      =   8760
   Begin VB.TextBox Text9 
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
      MaxLength       =   9
      TabIndex        =   13
      Top             =   4890
      Width           =   1572
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
      Height          =   330
      Left            =   6840
      TabIndex        =   7
      Top             =   948
      Width           =   1500
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
      Height          =   324
      Left            =   4428
      TabIndex        =   6
      Top             =   936
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2880
      Picture         =   "Frmacc4170.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   600
      Width           =   350
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
      Height          =   600
      Left            =   480
      Picture         =   "Frmacc4170.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   14
      ToolTipText     =   "清除畫面"
      Top             =   3120
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4170.frx":09CC
      Height          =   1650
      Left            =   240
      TabIndex        =   16
      Top             =   1410
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2900
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
         DataField       =   "a0d03"
         Caption         =   "項次"
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
         DataField       =   "a0d05"
         Caption         =   "科目代號"
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
      BeginProperty Column03 
         DataField       =   "a0d06"
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
      BeginProperty Column04 
         DataField       =   "a0d07"
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
      BeginProperty Column05 
         DataField       =   "a0d08"
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
      BeginProperty Column06 
         DataField       =   "a0d10"
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
      BeginProperty Column07 
         DataField       =   "a0d11"
         Caption         =   "對沖代號(其它)"
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
            Alignment       =   2
            ColumnWidth     =   515.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1152
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2724.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   684.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4452.095
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   1290
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1736
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
      Height          =   600
      Left            =   1200
      Picture         =   "Frmacc4170.frx":09E1
      Style           =   1  '圖片外觀
      TabIndex        =   15
      ToolTipText     =   "取消"
      Top             =   3120
      Width           =   612
   End
   Begin VB.TextBox Text6 
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
      Left            =   240
      TabIndex        =   17
      Top             =   4125
      Width           =   492
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      MaxLength       =   14
      TabIndex        =   9
      Top             =   4125
      Width           =   1572
   End
   Begin VB.TextBox Text5 
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
      Left            =   5040
      MaxLength       =   14
      TabIndex        =   10
      Top             =   4125
      Width           =   1572
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4860
      TabIndex        =   27
      Top             =   3120
      Width           =   1368
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6240
      TabIndex        =   26
      Top             =   3120
      Width           =   1356
   End
   Begin VB.TextBox Text14 
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   8
      Top             =   4125
      Width           =   972
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
      Left            =   1800
      TabIndex        =   25
      Top             =   4125
      Width           =   1452
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
      Height          =   330
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   11
      Top             =   4125
      Width           =   612
   End
   Begin VB.TextBox Text17 
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
      Left            =   7320
      TabIndex        =   24
      Top             =   4125
      Width           =   1212
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
      Height          =   330
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   2
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   19
      Top             =   240
      Width           =   3372
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
      Height          =   330
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   0
      Top             =   240
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   6840
      TabIndex        =   1
      Top             =   228
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   593
      _Version        =   393216
      MaxLength       =   2
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
      Height          =   330
      Left            =   6840
      TabIndex        =   23
      Top             =   600
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   593
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   330
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   1000
      _ExtentX        =   1757
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   330
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   1005
      _ExtentX        =   1778
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   345
      Left            =   840
      TabIndex        =   12
      Top             =   4500
      Width           =   7695
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13573;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(其它)"
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
      Left            =   240
      TabIndex        =   40
      Top             =   4935
      Width           =   1500
   End
   Begin VB.Label Label16 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "PS : 貸方 暫收款(2401) 科目請放在貸方的頭一個項次"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   39
      Top             =   3495
      Width           =   5415
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "餘額"
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
      Left            =   6216
      TabIndex        =   38
      Top             =   972
      Width           =   516
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "總額"
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
      Left            =   3828
      TabIndex        =   37
      Top             =   960
      Width           =   516
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   36
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "有效期間"
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
      TabIndex        =   35
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label10 
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
      Left            =   240
      TabIndex        =   34
      Top             =   4485
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "項次"
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
      Left            =   240
      TabIndex        =   33
      Top             =   3885
      Width           =   495
   End
   Begin VB.Label Label5 
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
      Left            =   1560
      TabIndex        =   32
      Top             =   3885
      Width           =   975
   End
   Begin VB.Label Label6 
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
      Left            =   3600
      TabIndex        =   31
      Top             =   3885
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Left            =   5280
      TabIndex        =   30
      Top             =   3885
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
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
      Left            =   7200
      TabIndex        =   29
      Top             =   3885
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1575
      Left            =   120
      Top             =   3765
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4845
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label15 
      Alignment       =   2  '置中對齊
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
      Left            =   4080
      TabIndex        =   28
      Top             =   3165
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "上次處理日期"
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
      Left            =   5400
      TabIndex        =   22
      Top             =   600
      Width           =   1452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "流水號"
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
      TabIndex        =   21
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "每月傳票日"
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
      Left            =   5400
      TabIndex        =   20
      Top             =   240
      Width           =   1212
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      TabIndex        =   18
      Top             =   240
      Width           =   732
   End
End
Attribute VB_Name = "Frmacc4170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/01 改成Form2.0 ; Combo1、DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0d1 As New ADODB.Recordset
Public adoacc0d0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim mPrevForm As Form 'Added by Lydia 2017/03/01 前一畫面表單
Dim mA2b01 As String 'Added by Lydia 2017/03/01 財產目錄編號
Dim bolCheck As Boolean 'Added by Lydia 2017/03/09 是否詢問過與財產目錄不同
Public bolA4170Jump As Boolean 'Added by Lydia 2021/12/ 22 跳過一次KeyF9檢查; 因為從frmacc41i0新增時自動呼叫frmacc4170會對frmacc4170再執行一次F9

'Added by Lydia 2017/03/01 呼叫表單
Public Sub SetFmForm(ByRef fm As Form, Optional ByVal pNo As String, Optional ByVal bJump As Boolean = False)
    Set mPrevForm = fm
    If pNo <> "" Then mA2b01 = pNo
    bolA4170Jump = bJump 'Added by Lydia 2021/12/22 跳過一次KeyF9檢查
    
End Sub

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo1_GotFocus()
   OpenIme
End Sub

'Modified by Lydia 2021/12/01 改成Form 2.0; KeyCode As Integer=>MSForms.ReturnInteger
Private Sub Combo1_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   'Modifie by Lydia 2021/12/01 +val()
   KeyDefine Val(KeyCode)
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo1_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Command1_Click()
Dim BookThisRec 'Add by Amy 2014/01/07
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Add by Amy 2014/01/07
   Adodc1.Recordset.MovePrevious
   If Adodc1.Recordset.BOF Then
        Adodc1.Recordset.MoveFirst
        BookThisRec = Adodc1.Recordset.Bookmark
   Else
        BookThisRec = Adodc1.Recordset.Bookmark
        Adodc1.Recordset.MoveNext
   End If
   'end 2014/01/07
   adoTaie.Execute "delete from acc0d0 where a0d01 = '" & Text1 & "' and a0d02 = '" & Val(Text3) & "' and a0d03 = '" & Text6 & "'"
   adoacc0d0.Close
   adoacc0d0.CursorLocation = adUseClient
   'Modified by Lydia 2017/05/10 依流水號排序
   'adoacc0d0.Open "select * from acc0d0 order by a0d01 asc, a0d02 asc, a0d03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0d0.Open "select * from acc0d0 order by a0d02 asc, a0d03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   AdodcRefresh
   'Add by Amy 2014/01/07
   If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.Bookmark = BookThisRec
   End If
   'end 2014/01/07
   AdodcClear
   Text6 = GetSeqNo(Text1, Text3) 'Added by Lydia 2017/01/18 重抓流水
   
   If adoacc0d0.RecordCount = 0 Then
      StatusClear
   Else
      SumShow 'Added by Lydia 2017/01/18
      RecordShow
   End If
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command2_Click()
'Remvove by Lydia 2017/01/18
'Modified by Lydia 2017/01/17 改成模組
'Dim adoaccmax As New ADODB.Recordset
'   If Adodc1.Recordset.RecordCount = 0 Then
'      Text6 = ZeroBeforeNo(0, 3)
'      Text14.SetFocus
'      Exit Sub
'   End If
'   adoaccmax.CursorLocation = adUseClient
'   adoaccmax.Open "select max(a0d03) from acc0d0 where a0d01 = '" & Text1 & "' and a0d02 = " & Val(Text3) & "", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccmax.RecordCount <> 0 Then
'      If IsNull(adoaccmax.Fields(0).Value) Then
'         Text6 = ZeroBeforeNo(0, 3)
'      Else
'         Text6 = ZeroBeforeNo(adoaccmax.Fields(0).Value, 3)
'      End If
'   End If
'   adoaccmax.Close
   AdodcClear
   Text6 = GetSeqNo(Text1, Text3)
   Text14.SetFocus
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command3_Click()
   If adoacc0d1.RecordCount = 0 Or Text1 = MsgText(601) Or Text3 = MsgText(601) Then
      Exit Sub
   End If
   adoacc0d1.Find "axd01 = '" & Text1 & "'", 0, adSearchForward, 1
   If adoacc0d1.EOF = False Then
      adoacc0d1.Find "axd02 = '" & Text3 & "'", 0, adSearchForward, adoacc0d1.Bookmark
      If adoacc0d1.EOF Then
         MsgBox MsgText(33), , MsgText(5)
         adoacc0d1.MoveFirst
      End If
   Else
      MsgBox MsgText(33), , MsgText(5)
      adoacc0d1.MoveFirst
   End If
   FormShow
   AdodcRefresh
   SumShow
   RecordShow
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyDefine KeyCode
   KeyEnter KeyCode
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
   strFormName = Name
   'Modified by Lydia 2017/12/05 增加判斷strItemNo,避免adoacc0d1.Find出錯
   'If strCompanyNo = MsgText(601) Then
   If strCompanyNo = MsgText(601) Or strItemNo = MsgText(601) Then
      Exit Sub
   End If
   adoacc0d1.Find "axd01 = '" & strCompanyNo & "'", 0, adSearchForward, 1
   If adoacc0d1.EOF = False Then
      adoacc0d1.Find "axd02 = '" & strItemNo & "'", 0, adSearchForward, adoacc0d1.Bookmark
      If adoacc0d1.EOF = False Then
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
   End If
   strCompanyNo = MsgText(601)
End Sub

'Added by Lydia 2021/12/01
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/01 Form2.0 記錄鍵盤傳入順序
End Sub

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
   Me.Height = 5850 'Modify by Amy 2014/05/13 原:5500 'Modified by Lydia 2021/12/01 5700=>5850
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = Mid(DFormat, 1, 6)
   MaskEdBox4.Mask = Mid(DFormat, 1, 6)
   OpenTable
   'Text1 = "1" 'Modify by Amy 2013/12/20 不預帶
   If adoacc0d1.RecordCount <> 0 Then
      adoacc0d1.MoveLast
      adoacc0d1.MoveFirst
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
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/01 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc4170 = Nothing
   
   'Added by Lydia 2017/03/01 回到前一畫面
   If TypeName(mPrevForm) = "Frmacc41i0" Then
      strItemNo = mA2b01
      Frmacc41i0.Show
      tool1_enabled
   End If

End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   'Modify By Sindy 2011/3/8 改為只能輸入到25, 因2月26,27,28有可能放假日, 無法產生傳票
   'If Val(MaskEdBox1.Text) < 1 Or Val(MaskEdBox1.Text) > 28 Then
   If Val(MaskEdBox1.Text) < 1 Or Val(MaskEdBox1.Text) > 25 Then
      'MsgBox Label1 & MsgText(56), , MsgText(5)
      MsgBox Label1 & "限於1至25...;因2月26,27,28有可能為放假日, 無法產生傳票!", , MsgText(5)
      Cancel = True
   End If
End Sub

'2009/12/7 add by sonia
Private Sub MaskEdBox4_Validate(Cancel As Boolean)
   If MaskEdBox4 <> MsgText(601) Then
      If Not MaskEdBox3.Text < MaskEdBox4.Text Then
         MsgBox "有效期間範圍不正確 !", vbCritical
         MaskEdBox4.SetFocus
         Cancel = True
      End If
   End If
End Sub
'2009/12/7 end

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   'Add by Amy 2020/04/07
   If InStr(GetBookKeepCmp, Text1) = 0 Then
     Text2 = ""
     Exit Sub
   End If
   Text2 = A0802Query(Text1)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

'Add by Amy 2013/12/20
Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2013/12/30

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0d1.CursorLocation = adUseClient
   'Modified by Lydia 2017/05/10 依流水號排序
   'adoacc0d1.Open "select * from acc0d1 order by axd01 asc, axd02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0d1.Open "select * from acc0d1 order by axd02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0d0.CursorLocation = adUseClient
   'Modified by Lydia 2017/05/10 依流水號排序
   'adoacc0d0.Open "select * from acc0d0 where a0d01 = '" & Text1 & "' and a0d02 = " & Val(Text3) & " and a0d03 = '" & Text6 & "' order by a0d01 asc, a0d02 asc, a0d03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0d0.Open "select * from acc0d0 where a0d01 = '" & Text1 & "' and a0d02 = " & Val(Text3) & " and a0d03 = '" & Text6 & "' order by a0d02 asc, a0d03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0d0, acc010 where a0d05 = a0101 (+) and a0d01 = '" & Text1 & "' and a0d02 = '" & Text3 & "' order by a0d03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料(自動產生傳票資料)
'
'*************************************************
Public Sub FormShow()
   Text1 = adoacc0d1.Fields("axd01").Value
   If IsNull(adoacc0d1.Fields("axd03").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = adoacc0d1.Fields("axd03").Value
   End If
   Text3 = adoacc0d1.Fields("axd02").Value
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0d1.Fields("axd04").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0d1.Fields("axd04").Value)
   End If
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(adoacc0d1.Fields("axd11").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      'Modify By Sindy 2011/3/8
      'MaskEdBox3.Text = IIf(Len(adoacc0d1.Fields("axd11").Value) < 5, "0" & Mid(adoacc0d1.Fields("AXD11").Value, 1, 2) & "/" & Mid(adoacc0d1.Fields("AXD11").Value, 3, 2), Mid(adoacc0d1.Fields("AXD11").Value, 1, 2) & "/" & Mid(adoacc0d1.Fields("AXD11").Value, 3, 2))
      MaskEdBox3.Text = IIf(Len(adoacc0d1.Fields("axd11").Value) < 5, "0" & Mid(adoacc0d1.Fields("AXD11").Value, 1, 2) & "/" & Mid(adoacc0d1.Fields("AXD11").Value, 3, 2), Mid(adoacc0d1.Fields("AXD11").Value, 1, 3) & "/" & Mid(adoacc0d1.Fields("AXD11").Value, 4, 2))
   End If
   MaskEdBox3.Mask = Mid(DFormat, 1, 6)
   MaskEdBox4.Mask = MsgText(601)
   If IsNull(adoacc0d1.Fields("axd12").Value) Then
      MaskEdBox4.Text = MsgText(601)
   Else
      MaskEdBox4.Text = IIf(Len(adoacc0d1.Fields("axd12").Value) < 5, "0" & Mid(adoacc0d1.Fields("AXD12").Value, 1, 2) & "/" & Mid(adoacc0d1.Fields("AXD12").Value, 3, 2), Mid(adoacc0d1.Fields("AXD12").Value, 1, 3) & "/" & Mid(adoacc0d1.Fields("AXD12").Value, 4, 2))
   End If
   MaskEdBox4.Mask = Mid(DFormat, 1, 6)
   If IsNull(adoacc0d1.Fields("axd13").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = adoacc0d1.Fields("axd13").Value
   End If
   If IsNull(adoacc0d1.Fields("axd14").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc0d1.Fields("axd14").Value
   End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   'Add by Amy 2014/01/02
   'Modify by Amy 2020/04/07
   'If Text1 <> "1" And Text1 <> "J" Then
   If InStr(GetBookKeepCmp, Text1) = 0 Then
         MsgBox Label4 & MsgText(63), , MsgText(5) '原:"公司別只可輸入 1 或 J"
   'end 2020/04/07
         Cancel = True
         Text1.SetFocus
         Exit Sub
   End If
   'end 2014/01/02
   If ExistCheck("acc080", "a0801", Text1, Label4) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text14_Change()
   If Text14 = MsgText(601) Then
      Exit Sub
   End If
   Text15 = A0102Query(Text14)
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
   CloseIme 'Added by Lydia 2017/05/10
End Sub

Private Sub Text14_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
   'Add by Amy 2014/02/11
   If strSaveConfirm = MsgText(601) Then
        Exit Sub
   End If
   'end 2014/02/11
   If Text14 <> MsgText(601) Then
      'Modify by Amy 204/01/07 +公司別確認
'      If ExistCheck("acc010", "a0101", Text14, Label5) = False Then
'         Cancel = True
'         Exit Sub
'      End If
      If PUB_CheckCompany(Text14, Text1) = False Then
         Cancel = True
         Exit Sub
      End If
      'end 2014/01/07
   End If
End Sub

Private Sub Text16_Change()
   'Remove by Lydia 2017/03/01
'   If CheckDept(Text14, Text16) = False Then
'      MsgBox MsgText(103), , MsgText(5)
'      'edit by nickc 2007/02/08
'      'Cancel = True
'      Exit Sub
'   End If
'   If Text16 = MsgText(601) Then
'      Exit Sub
'   End If
   Text17 = A0902Query(Text16)
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
   CloseIme
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   'Memo by Lydia 2017/03/01 從下面移上來
   If CheckDept(Text14, Text16) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   'end 2017/03/01
   If Text16 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text16, Label8) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text3 = MsgText(601) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text14.SetFocus
         Exit Sub
   End Select
   KeyDefine (KeyCode)
End Sub

'*************************************************
'  清除欄位資料
'
'*************************************************
Public Sub AdodcClear()
   Text14 = ""
   Text15 = ""
   Text4 = ""
   Text5 = ""
   Text16 = ""
   Text17 = ""
   Combo1 = ""
   Text9 = "" 'Add by Amy 2014/05/13
   Text6 = "" 'Added by Lydia 2017/01/18
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)

   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/01 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         'Added by Lydia 2021/12/01 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'end 2021/12/01
         'Added by Lydia 2021/12/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
             Exit Sub
         End If
         'end 2021/12/01
         
         Frmacc4170_Save
         If strControlButton <> MsgText(602) Then
            acc0d0Save
         End If
         If strControlButton <> MsgText(602) Then
            'Remove by Lydia 2017/01/18
            'If Text6 = MsgText(601) Then
            '   Text6 = MsgText(16)
            'Else
            '   Text6 = ZeroBeforeNo(Text6, 3)
            'End If
            Combo1.AddItem Combo1
            AdodcClear
            Text6 = GetSeqNo(Text1, Text3) 'Added by Lydia 2017/01/18
            SumShow
            Text1.Locked = True
            Text14.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  顯示Grid資料(自動產生傳票資料)
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0d0, acc010 where a0d05 = a0101 (+) and a0d01 = '" & Text1 & "' and a0d02 = '" & Text3 & "' order by a0d03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   'Added by Lydia 2017/01/18 Grid 移動到現在項次
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a0d03 = '" & Text6 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Adodc1.Recordset.MoveFirst
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   End If
   'end 2017/01/18
   
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(Grid資料)
'
'*************************************************
Private Sub AdodcShow()
   If IsNull(Adodc1.Recordset.Fields("a0d03").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a0d03").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0d05").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = Adodc1.Recordset.Fields("a0d05").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0d06").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0d06").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0d07").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("a0d07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0d08").Value) Then
      Text16 = MsgText(601)
   Else
      If Adodc1.Recordset.Fields("a0d08").Value = MsgText(55) Then
         Text16 = MsgText(601)
      Else
         Text16 = Adodc1.Recordset.Fields("a0d08").Value
      End If
   End If
   If IsNull(Adodc1.Recordset.Fields("a0d10").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("a0d10").Value
   End If
   'Add by Amy 2014/05/13
   If IsNull(Adodc1.Recordset.Fields("a0d11").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a0d11").Value
   End If
   'end 2014/05/13
End Sub

'*************************************************
'  計算並顯示合計資料
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0d06), sum(a0d07) from acc0d0 where a0d01 = '" & Text1 & "' and a0d02 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text11 = MsgText(601)
      Else
         Text11 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = Format(adoaccsum.Fields(1).Value, DDollar)
      End If
   Else
      Text11 = MsgText(601)
      Text12 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc0d1.RecordCount <> 0 Then
      Frmacc0000.StatusBar1.Panels(2).Text = adoacc0d1.Bookmark & MsgText(35) & adoacc0d1.RecordCount
   End If
End Sub

'*************************************************
'  儲存欄位資料(傳票資料--交易檔)
'
'*************************************************
Public Sub acc0d0Save()
On Error GoTo Checking
   If Text14 = MsgText(601) Then
      MsgBox MsgText(10) & Label5, , MsgText(5)
      strControlButton = MsgText(602)
      Text14.SetFocus
      Exit Sub
   Else
      'Modify by Amy 2014/01/07 +公司別確認
'      If ExistCheck("acc010", "a0101", Text14, Label5) = False Then
'         strControlButton = MsgText(602)
'         Text14.SetFocus
'         Exit Sub
'      End If
      If PUB_CheckCompany(Text14, Text1) = False Then
         strControlButton = MsgText(602)
         Text14.SetFocus
         Exit Sub
      End If
      'end 2014/01/07
      
      If Val(Text4) <> 0 And Val(Text5) <> 0 Then
         MsgBox MsgText(47) & MsgText(46), , MsgText(5)
         strControlButton = MsgText(602)
         Text4.SetFocus
         Exit Sub
      End If
      If CheckDept(Text14, Text16) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text16.SetFocus
         Exit Sub
      End If
      If Text16 <> MsgText(601) Then
         If ExistCheck("acc090", "a0901", Text16, Label8) = False Then
            strControlButton = MsgText(602)
            Text16.SetFocus
            Exit Sub
         End If
      End If
   End If
   
   'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Text14, Text16)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Text14.SetFocus
      ElseIf intI = 2 Then
         Text16.SetFocus
      End If
      Exit Sub
   End If
   'end 2007/10/2
   
    'Added by Lydia 2017/01/18 重抓流水
    If Text6 = MsgText(601) Then
       Text6 = GetSeqNo(Text1, Text2) '解按修改直接insert的錯誤
    End If
      
   adoacc0d0.Close
   adoacc0d0.CursorLocation = adUseClient
   'Modified by Lydia 2017/05/10 依流水號排序
   'adoacc0d0.Open "select * from acc0d0 where a0d01 = '" & Text1 & "' and a0d02 = " & Val(Text3) & " and a0d03 = '" & Text6 & "' order by a0d01 asc, a0d02 asc, a0d03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0d0.Open "select * from acc0d0 where a0d01 = '" & Text1 & "' and a0d02 = " & Val(Text3) & " and a0d03 = '" & Text6 & "' order by a0d02 asc, a0d03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0d0.RecordCount = 0 Then
      adoacc0d0.AddNew
   End If
   adoacc0d0.Fields("a0d01").Value = Text1
   adoacc0d0.Fields("a0d02").Value = Val(Text3)
   adoacc0d0.Fields("a0d03").Value = Text6
   If Text14 <> MsgText(601) Then
      adoacc0d0.Fields("a0d05").Value = Text14
   Else
      adoacc0d0.Fields("a0d05").Value = Null
   End If
   If Text4 <> MsgText(601) Then
      adoacc0d0.Fields("a0d06").Value = Val(Text4)
   Else
      adoacc0d0.Fields("a0d06").Value = 0
   End If
   If Text5 <> MsgText(601) Then
      adoacc0d0.Fields("a0d07").Value = Val(Text5)
   Else
      adoacc0d0.Fields("a0d07").Value = 0
   End If
   If Text16 <> MsgText(601) Then
      adoacc0d0.Fields("a0d08").Value = Text16
   Else
      adoacc0d0.Fields("a0d08").Value = MsgText(55)
   End If
   If Combo1 <> MsgText(601) Then
      adoacc0d0.Fields("a0d10").Value = Combo1
   Else
      adoacc0d0.Fields("a0d10").Value = Null
   End If
   'Add by Amy 2014/05/13 +對沖代號(其它)
   If Text9 <> MsgText(601) Then
      adoacc0d0.Fields("a0d11").Value = Text9
   Else
      adoacc0d0.Fields("a0d11").Value = Null
   End If
   adoacc0d0.UpdateBatch
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  借貸方檢核
'
'*************************************************
Public Function CreDebCheck() As String
   If Text11 = Text12 Then
      CreDebCheck = MsgText(602)
   End If
End Function

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   'Add by Amy 2013/12/24
   Text1.Locked = False
   Text3.Locked = False
    Command3.Enabled = True
   'end 2013/12/24
   Text14.Enabled = False
   Text4.Enabled = False
   Text5.Enabled = False
   Text16.Enabled = False
   Combo1.Enabled = False
   Text9.Enabled = False 'Add by Amy 2014/05/13
   Command1.Enabled = False
   Command2.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   'Add by Amy 2013/12/20
   If strSaveConfirm = MsgText(3) Then
        Text1.Locked = False
        Text3.Locked = True
        Command3.Enabled = False
   ElseIf strSaveConfirm = MsgText(4) Then
        Text1.Locked = True
        Text3.Locked = True
        Command3.Enabled = False
   End If
   'end 2013/12/20
   Text14.Enabled = True
   Text4.Enabled = True
   Text5.Enabled = True
   Text16.Enabled = True
   Combo1.Enabled = True
   Text9.Enabled = True 'Add by Amy 2014/05/13
   Command1.Enabled = True
   Command2.Enabled = True
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(601) Then
      Text8 = Text7
   End If
End Sub

'Add by Amy 2013/12/20 將Frmacc4170_Clear搬回
Public Sub Frmacc4170_Clear()
Dim adoautono As New ADODB.Recordset
    
      Text1 = "" 'Modify by 2013/12/20 原:"1"
      MaskEdBox1.Text = ""
      adoautono.CursorLocation = adUseServer
      'Modify by Amy 2013/12/20 改只抓axd02最大值不需判斷公司別
      'adoautono.Open "select max(axd02) from acc0d1 where axd01 = '1'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      adoautono.Open "select max(axd02) from acc0d1 ", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoautono.RecordCount <> 0 Then
         If IsNull(adoautono.Fields(0).Value) Then
            Text3 = Val(ZeroBeforeNo(0, 3))
         Else
            Text3 = Val(ZeroBeforeNo(adoautono.Fields(0).Value, 3))
         End If
'      Else
'         .Text3 = Val(ZeroBeforeNo(0, 3))
      End If

      adoautono.Close
      MaskEdBox2.Mask = ""
      MaskEdBox2.Text = ""
      MaskEdBox2.Mask = DFormat
      MaskEdBox3.Mask = ""
      MaskEdBox3.Text = ""
      MaskEdBox3.Mask = Mid(DFormat, 1, 6)
      MaskEdBox4.Mask = ""
      MaskEdBox4.Text = ""
      MaskEdBox4.Mask = Mid(DFormat, 1, 6)
      Text7 = ""
      Text8 = ""
      Text1.SetFocus
End Sub

'Modify by Amy 2014/05/13 由acc_sav搬回
Public Sub Frmacc4170_Save()
Dim rsAD As New ADODB.Recordset 'Added by Lydia 2016/12/09
Dim bolCancel As Boolean 'Add by Amy 2020/04/07

   On Error GoTo Checking
  
      If Text1 = MsgText(601) Then
         MsgBox MsgText(10) & Label4, , MsgText(5)
         strControlButton = MsgText(602)
         Text1.SetFocus
         Exit Sub
      Else
         'Add by Amy 2014/01/02
         'Modify by Amy 2020/04/07
         'If Text1 <> "1" And Text1 <> "J" Then
             'MsgBox "公司別只可輸入 1 或 J", , MsgText(5)
         Call Text1_Validate(bolCancel)
         If bolCancel = True Then
         'end 2020/04/07
            strControlButton = MsgText(602)
            Text1.SetFocus
            Exit Sub
         End If
         'end 2014/01/02
         If Text3 = MsgText(601) Then
            MsgBox MsgText(10) & Label2, , MsgText(5)
            strControlButton = MsgText(602)
            Text3.SetFocus
            Exit Sub
         End If
         If Val(MaskEdBox1.Text) < 1 Or Val(MaskEdBox1.Text) > 28 Then
            MsgBox MsgText(48), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Sub
         End If
'         'Add By Sindy 2010/12/20
'         Dim star_month As String, end_month As String
'         Dim intMonth As Integer
'         star_month = CStr(Val(Left(.MaskEdBox3, 3)) + 1911) & Mid(.MaskEdBox3, 4, Len(.MaskEdBox3))
'         end_month = CStr(Val(Left(.MaskEdBox4, 3)) + 1911) & Mid(.MaskEdBox4, 4, Len(.MaskEdBox4))
'         intMonth = DateDiff("m", star_month, end_month) + 1
'         If Val(Format(.Text7, strPercent)) / Val(Format(.Text11, strPercent)) <> intMonth Then
'            MsgBox "總額除以每月合計不等於有效月數！", , MsgText(5)
'            strControlButton = MsgText(602)
'            .Text7.SetFocus
'            Exit Sub
'         End If
'         '2010/12/20 End
         If ExistCheck("acc080", "a0801", Text1, Label4) = False Then
            strControlButton = MsgText(602)
            Text1.SetFocus
            Exit Sub
         End If
         'Added by Lydia 2022/02/07 檢查總額不可為零; ex.452一開始輸入為零or空白
         If Val(Text7) <= 0 Then
            MsgBox MsgText(10) & Label13, , MsgText(5)
            strControlButton = MsgText(602)
            Text7.SetFocus
            Exit Sub
         End If
         'end 2022/02/07
         'Added by Lydia 2024/09/12 在新增傳票時總額=餘額；ex.548一開始輸300000後來總額修改金額後直接按存檔,餘額未一併更新
         If strSaveConfirm = MsgText(3) And Val(Text7) <> Val(Text8) Then
            Text8 = Val(Text7)
         End If
         'end 2024/09/12
         
         'Added by Lydia 2022/12/30 在固定傳票frmacc4170直接新增612601~612604彈提醒不可新增 from 辜 (ex. 固定傳票489,493)
         If strSaveConfirm = MsgText(3) And mA2b01 = "" And InStr("612601,612602,612603,612604", Text14) > 0 And Text14.Text <> "" Then
            MsgBox "請從財產目錄新增資料！", vbCritical, MsgText(5)
            strControlButton = MsgText(602)
            Text14.SetFocus
            Exit Sub
         End If
         'end 2022/12/30
         
         'Added by Lydia 2016/12/09 檢查總額、已列金額、餘額、借貸方合計的關係
         '以公司別+流水號抓ACC1P0之A1P01='公司別' AND A1P02='U' AND A1P04>=流水號||有效期間起始年月 AND A1P04<=流水號||上次處理日期的借方總額SUM(A1P07)
         'Modified by Lydia 2017/01/06 新增時,上次處理日期為___/__/__
         'strSql = "select sum(a1p07) s1 from acc1p0 where a1p01='1' and a1p02='U' and A1P04>='" & Text3 & Replace(MaskEdBox3.Text, "/", "") & "' AND A1P04<='" & Text3 & Replace(MaskEdBox2.Text, "/", "") & "' "
         'Modified by Lydia 2017/02/23 改成模組
         'strExc(3) = Replace(Replace(MaskEdBox3.Text, "/", ""), "_", "")
         'strExc(2) = Replace(Replace(MaskEdBox2.Text, "/", ""), "_", "")
         'strSql = "select nvl(sum(a1p07),0) s1 from acc1p0 where a1p01='1' and a1p02='U' and A1P04>='" & Text3 & IIf(Trim(strExc(3)) <> "", strExc(3), "00000") & "' AND A1P04<='" & Text3 & IIf(Trim(strExc(2)) <> "", strExc(2), "00000") & "' "
         ''end 2017/01/06
         'intI = 1
         'Set rsAD = ClsLawReadRstMsg(intI, strSql)
         'If intI = 1 Then
         strExc(8) = PUB_SumA1PtoU(Text1, Text3, MaskEdBox2, MaskEdBox3)
         'Modified by Lydia 2022/02/07
         'If strExc(8) <> "" Then
         If Val(strExc(8)) <> 0 Then
         'end 2017/02/23
            '若總額－已列金額<>餘額，則顯示"餘額有問題，是否更新為….."，讓使用者選擇是否要更新，若更新則改正確後存檔；若不更新則游標停在總額欄，不可存檔；
            'Modified by Lydia 2017/02/23
            'strExc(0) = Val(Text7) - Val("" & rsAD(0))
            strExc(0) = Val(Text7) - Val(strExc(8))
            If Val(strExc(0)) <> Val(Text8) Then
               If MsgBox("餘額有問題，是否更新為" & strExc(0) & "？", vbYesNo + vbDefaultButton2, "檢查總額、已列金額和餘額") = vbYes Then
                  Text8.Text = strExc(0)
               Else
                  strControlButton = MsgText(602)
                  Text7.SetFocus
                  Exit Sub
               End If
            End If
         End If
         'end 2016/12/09
      End If
      
      'Added by Lydia 2017/05/11 與財產目錄不同
      If bolCheck = False And mA2b01 <> "" Then
         'Modified by Lydia 2017/05/22 指定會計科目
         'strSql = "select a2b01,a2b05,a2b06,a2b18,a2b20,a2b21,a2b22,nvl(sum(a1p07),0) amt1,nvl(sum(ax206),0) amt2 from acc2b0,acc1p0,acc021 " & _
                  "where a2b01='" & mA2b01 & "' and a2b16=a1p01(+) and a2b01||a2b05=a1p04(+) and a1p02(+)='M' and a2b16=ax201(+) and a2b22=ax202(+) " & _
                  "group by a2b01,a2b05,a2b06,a2b18,a2b20,a2b21,a2b22"
         strSql = "select a2b01,a2b05,a2b06,a2b18,a2b20,a2b21,a2b22,nvl(sum(a1p07),0) amt1,nvl(sum(ax206),0) amt2 from acc2b0,acc1p0,acc021 " & _
                  "where a2b01='" & mA2b01 & "' and a2b16=a1p01(+) and a2b01||a2b05=a1p04(+) and a1p02(+)='M' and (a1p05 is null or substr(a1p05,1,4)='6126') " & _
                  "and a2b16=ax201(+) and a2b22=ax202(+) and (ax205 is null or substr(ax205,1,4)='6126') " & _
                  "group by a2b01,a2b05,a2b06,a2b18,a2b20,a2b21,a2b22"
         intI = 1
         Set rsAD = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'Modified by Lydia 2017/05/23 未折減餘額若為負數,顯示為零
            'If Val(rsAd.Fields("a2b06") - IIf("" & rsAd.Fields("a2b22") <> "", rsAd.Fields("amt2"), rsAd.Fields("amt1"))) <> Val(Text7.Text) Then
            '   MsgBox "總額與財產目錄的未折減餘額 " & Val(rsAd.Fields("a2b06") - IIf("" & rsAd.Fields("a2b22") <> "", rsAd.Fields("amt2"), rsAd.Fields("amt1"))) & " 不一致，請洽電腦中心協助 !", vbCritical
            strExc(1) = Val("" & rsAD.Fields("a2b06")) - IIf("" & rsAD.Fields("a2b22") <> "", rsAD.Fields("amt2"), rsAD.Fields("amt1"))
            strExc(1) = IIf(Val(strExc(1)) < 0, "0", strExc(1))
            If Val(strExc(1)) <> Val(Text7.Text) Then
               MsgBox "總額與財產目錄的未折減餘額 " & strExc(1) & " 不一致，請洽電腦中心協助 !", vbCritical
            'end 2017/05/23
            End If
            If rsAD.Fields("a2b20") <> Val(Replace(MaskEdBox3.Text, "/", "")) Or rsAD.Fields("a2b21") <> Val(Replace(MaskEdBox4.Text, "/", "")) Then
               MsgBox "有效期間與財產目錄的攤提期間 " & rsAD.Fields("a2b20") & "－" & rsAD.Fields("a2b21") & " 不一致 !"
            End If
            If rsAD.Fields("a2b18") <> Val(Replace(MaskEdBox1.Text, "/", "")) Then
               MsgBox "每月傳票日與財產目錄的每月攤提日期 " & rsAD.Fields("a2b18") & " 不一致 !"
            End If
            bolCheck = True
         End If
      End If
      'end 2017/05/11
      
      If strSaveConfirm = MsgText(3) Then
         If adoacc0d1.RecordCount <> 0 Then
            adoacc0d1.Find "axd01 = '" & Text1 & "'", 0, adSearchForward, 1
            If adoacc0d1.EOF Then
               adoacc0d1.AddNew
            Else
               adoacc0d1.Find "axd02 = " & Val(Text3) & "", 0, adSearchForward, adoacc0d1.Bookmark
               If adoacc0d1.EOF Then
                  adoacc0d1.AddNew
               End If
            End If
         Else
            adoacc0d1.AddNew
         End If
      End If
      adoacc0d1.Fields("axd01").Value = Text1
      If MaskEdBox1.Text <> MsgText(601) Then
         adoacc0d1.Fields("axd03").Value = Val(MaskEdBox1.Text)
      Else
         adoacc0d1.Fields("axd03").Value = Null
      End If
      adoacc0d1.Fields("axd02").Value = Val(Text3)
      If MaskEdBox3.Text <> Mid(MsgText(29), 1, 6) Then
         adoacc0d1.Fields("AXD11").Value = Val(Mid(MaskEdBox3.Text, 1, 3) & Mid(MaskEdBox3.Text, 5, 2))
      Else
         adoacc0d1.Fields("AXD11").Value = Null
      End If
      If MaskEdBox4.Text <> Mid(MsgText(29), 1, 6) Then
         adoacc0d1.Fields("AXD12").Value = Val(Mid(MaskEdBox4.Text, 1, 3) & Mid(MaskEdBox4.Text, 5, 2))
      Else
         adoacc0d1.Fields("AXD12").Value = Null
      End If
      If Text7 <> MsgText(601) Then
         adoacc0d1.Fields("axd13").Value = Val(Text7)
      Else
         adoacc0d1.Fields("axd13").Value = 0
      End If
      If Text8 <> MsgText(601) Then
         adoacc0d1.Fields("axd14").Value = Val(Text8)
      Else
         adoacc0d1.Fields("axd14").Value = 0
      End If
      If strSaveConfirm = MsgText(3) Then
         adoacc0d1.Fields("axd06").Value = Val(strSrvDate(2))
         adoacc0d1.Fields("axd07").Value = ServerTime
         adoacc0d1.Fields("axd05").Value = strUserNum
      Else
         adoacc0d1.Fields("axd09").Value = Val(strSrvDate(2))
         adoacc0d1.Fields("axd10").Value = ServerTime
         adoacc0d1.Fields("axd08").Value = strUserNum
      End If
      adoacc0d1.UpdateBatch
      RecordShow
      
      Set rsAD = Nothing 'Added by Lydia 2016/12/09
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2014/05/13 +對沖代號(其它)
Private Sub Text9_GotFocus()
   TextInverse Text9
   OpenIme
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyDefine KeyCode
End Sub
'end 2014/05/13

'Added by Lydia 2017/01/18 取得最大明細項次
Private Function GetSeqNo(stra0d01 As String, stra0d02 As String) As String
    '取得項次
    Dim adoaccmax As New ADODB.Recordset
    
    If adoaccmax.State = adStateOpen Then
         adoaccmax.Close
    End If
    adoaccmax.CursorLocation = adUseClient
    'Modified by Lydia 2017/03/09 max(a0d03) -> nvl(max(a0d03),0) mno
    adoaccmax.Open "select nvl(max(a0d03),0) mno from acc0d0 where a0d01 = '" & stra0d01 & "' and a0d02 = " & Val(stra0d02), adoTaie, adOpenStatic, adLockReadOnly
    If adoaccmax.RecordCount <> 0 Then
       GetSeqNo = ZeroBeforeNo(Val(adoaccmax.Fields(0).Value), 3)
    Else
       GetSeqNo = ZeroBeforeNo(0, 3)
    End If
    adoaccmax.Close
End Function
