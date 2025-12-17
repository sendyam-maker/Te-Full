VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1140 
   AutoRedraw      =   -1  'True
   Caption         =   "收據抬頭修改"
   ClientHeight    =   5424
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9132
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5424
   ScaleWidth      =   9132
   Begin VB.TextBox txtPrintNo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   4
      Top             =   450
      Width           =   345
   End
   Begin VB.CommandButton CmdAddM 
      Caption         =   "加註備註欄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7665
      TabIndex        =   54
      Top             =   2280
      Width           =   1170
   End
   Begin VB.ComboBox CboClass 
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
      Left            =   7110
      TabIndex        =   51
      Text            =   "CboClass"
      Top             =   1560
      Width           =   1605
   End
   Begin VB.ComboBox cboA0S03 
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
      Left            =   3765
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   1890
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   210
      TabIndex        =   42
      Top             =   2640
      Width           =   8175
      Begin VB.CheckBox Check2 
         Caption         =   "3.代理人請款之匯款日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   5130
         TabIndex        =   17
         Top             =   0
         Width           =   2355
      End
      Begin VB.CheckBox Check2 
         Caption         =   "2.代理人請款日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3300
         TabIndex        =   16
         Top             =   0
         Width           =   1725
      End
      Begin VB.CheckBox Check2 
         Caption         =   "1.送件日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2100
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "收據自動列印時間點"
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
         Left            =   120
         TabIndex        =   43
         Top             =   60
         Width           =   1890
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1140.frx":0000
      Height          =   1845
      Left            =   180
      TabIndex        =   18
      Top             =   3405
      Width           =   8445
      _ExtentX        =   14901
      _ExtentY        =   3260
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
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
      ColumnCount     =   7
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
         DataField       =   "a0j10"
         Caption         =   "規費"
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
         DataField       =   "cp27t"
         Caption         =   "發文日"
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
            ColumnWidth     =   1848.189
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   552.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1188.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1019.906
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "收據暫不列印"
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
      Left            =   4980
      TabIndex        =   12
      Top             =   1920
      Width           =   1605
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   5
      Top             =   810
      Width           =   1425
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6600
      TabIndex        =   14
      Top             =   2250
      Width           =   1000
   End
   Begin VB.ComboBox cboA0M01 
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
      Left            =   1125
      Style           =   2  '單純下拉式
      TabIndex        =   10
      Top             =   1890
      Width           =   1605
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      MaxLength       =   3
      TabIndex        =   8
      Top             =   1170
      Width           =   945
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1170
      Width           =   372
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2310
      Picture         =   "Frmacc1140.frx":0015
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   90
      Width           =   350
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6030
      MaxLength       =   6
      TabIndex        =   36
      Top             =   810
      Width           =   795
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1170
      Width           =   612
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   3270
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   614
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
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3990
      TabIndex        =   20
      Top             =   90
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4470
      TabIndex        =   24
      Top             =   3030
      Width           =   1572
   End
   Begin VB.TextBox Text5 
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
      Left            =   5040
      MaxLength       =   1
      TabIndex        =   3
      Top             =   450
      Width           =   315
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1590
      TabIndex        =   23
      Top             =   3030
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6870
      TabIndex        =   22
      Top             =   3030
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1050
      MaxLength       =   15
      TabIndex        =   0
      Top             =   90
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   1050
      TabIndex        =   19
      Top             =   810
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   572
      _Version        =   393216
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   7470
      TabIndex        =   53
      Top             =   1920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   593
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Left            =   1500
      TabIndex        =   9
      Top             =   1560
      Width           =   1530
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2699;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   1050
      TabIndex        =   2
      Top             =   450
      Width           =   4005
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7064;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text9 
      Height          =   480
      Left            =   660
      TabIndex        =   13
      Top             =   2190
      Width           =   4365
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "7699;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   5550
      TabIndex        =   25
      Top             =   90
      Width           =   3285
      VariousPropertyBits=   679493657
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "5794;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblA0k20_N 
      Height          =   288
      Left            =   6876
      TabIndex        =   58
      Top             =   840
      Width           =   1260
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2222;508"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label28 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   330
      Left            =   6645
      TabIndex        =   57
      Top             =   1170
      Width           =   960
   End
   Begin VB.Label Label27 
      BackStyle       =   0  '透明
      Caption         =   "付款週期月份"
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
      Left            =   5280
      TabIndex        =   56
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印統編       (Y:印)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   7260
      TabIndex        =   55
      Top             =   516
      Width           =   1584
   End
   Begin VB.Label Label25 
      BackStyle       =   0  '透明
      Caption         =   "控管日期"
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
      Left            =   6630
      TabIndex        =   52
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label23 
      BackStyle       =   0  '透明
      Caption         =   "控管類別"
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
      Left            =   6150
      TabIndex        =   50
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   345
      Left            =   4950
      TabIndex        =   49
      Top             =   1530
      Width           =   1005
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "介紹獎金可發放日期"
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
      Left            =   3030
      TabIndex        =   48
      Top             =   1560
      Width           =   1995
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "介紹案源同仁"
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
      Left            =   150
      TabIndex        =   47
      Top             =   1590
      Width           =   1515
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "銷退日期"
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
      Left            =   2820
      TabIndex        =   46
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   8610
      TabIndex        =   45
      Top             =   1050
      Width           =   1200
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "預訂收款日"
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
      Left            =   7470
      TabIndex        =   44
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
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
      Left            =   2730
      TabIndex        =   41
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "最近修改抬頭日"
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
      Left            =   5070
      TabIndex        =   40
      Top             =   2250
      Width           =   1695
   End
   Begin VB.Label Label24 
      BackStyle       =   0  '透明
      Caption         =   "收款單號"
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
      Left            =   150
      TabIndex        =   39
      Top             =   1890
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
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
      Left            =   3300
      TabIndex        =   38
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   150
      TabIndex        =   37
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   5160
      TabIndex        =   35
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "列印次數"
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
      Left            =   1620
      TabIndex        =   34
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   150
      TabIndex        =   33
      Top             =   2250
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2985
      Left            =   60
      Top             =   30
      Width           =   8820
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   3030
      TabIndex        =   32
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Left            =   150
      TabIndex        =   31
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
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
      Left            =   150
      TabIndex        =   30
      Top             =   840
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
      Left            =   5400
      TabIndex        =   29
      Top             =   510
      Width           =   1755
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "規費合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3390
      TabIndex        =   28
      Top             =   3030
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "服務費合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   390
      TabIndex        =   27
      Top             =   3030
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "總計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6270
      TabIndex        =   26
      Top             =   3030
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
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
      Left            =   150
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -90
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc1140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/01 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoacc0j0 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public m_dbla0k17 As Double '已收服務費
'Add By Cheng 2003/12/04
Public m_dbla0k18 As Double '已收規費
Public m_CP09 As String 'Add By Sindy 2012/12/6
Public m_ShowMsg As String 'Add By Sindy 2013/12/27
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, m_CP31 As String
Dim m_CP10 As String 'Add By Sindy 2023/9/5
Public TmpFrmacc1420 As Form 'Added by Lydia 2016/04/14
Dim bolUpdData As Boolean 'Add By Sindy 2017/11/2
Dim HasA4319 As Boolean, HasA4321 As Boolean 'Add by Amy 2021/01/28 有上傳發票/有作廢
Public ProState As String '權限: 1.全所 2.該所 add by sonia 2023/5/26
Dim m_strA0k32 As String 'Added by Lydia 2023/12/12

Private Sub cboA0M01_Click()
    'Add By Cheng 2004/02/02
    '顯示發票號碼
    If Text12 <> "J" Then 'Add By Sindy 2013/12/27 +if
      If Me.cboA0M01.Text <> "" Then
          ShowInvoiceNo Me.cboA0M01.Text, Me.Text1.Text
      Else
          Me.Text15.Text = ""
      End If
    End If
    'End
End Sub

'Add by Amy 2017/05/24
Private Sub cboClass_Click()
    '若「預訂收款日」有值,且選「預計收款」則「控管日期」=預訂收款日-瑞婷
    If (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) And (MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29)) Then
        If Label18 <> MsgText(601) Then
            MaskEdBox1 = Label18
        End If
    End If
End Sub

Private Sub CboClass_Validate(Cancel As Boolean)
    If CboClass = "" Then Exit Sub
    With CboClass
        Select Case CboClass
            'Modify by Amy 2015/08/24 +催款中/預計收款
            'Modify by Amy 2017/08/29 待收款 改為 請款中
            Case "待銷帳", "請款中", "未送件", "依流程請款", "其他", "會稿中", "尚未辦理", "催款中", "預計收款"
            Case Else
                ShowMsg Label23 & "錯誤,請以下拉方式點選 !"
                Cancel = True
        End Select
    End With
End Sub

'Add By Sindy 2013/12/25
Private Sub Check2_GotFocus(Index As Integer)
   Select Case Index
      Case 0
         Check2(1).Value = 0
         Check2(2).Value = 0
      Case 1
         Check2(0).Value = 0
         Check2(2).Value = 0
      Case 2
         Check2(0).Value = 0
         Check2(1).Value = 0
   End Select
End Sub

Private Sub CmdAddM_Click()
    If CboClass = MsgText(601) Or MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        MsgBox Label23 & " 及 " & Label25 & "不可有空值!!"
        If CboClass = MsgText(601) Then
            CboClass.SetFocus
        Else
            MaskEdBox1.SetFocus
        End If
        Exit Sub
    End If
    'Modify by Amy 2016/08/23 取消日期中的/及"控管"字樣;"預計//日收款"改為"預計年月日收款"
    If CboClass = "預計收款" Then
        'Add by Amy 2015/08/24 +若選「預計收款」增加「預計//日收款」文字
        Text9 = "預計" & FCDate(MaskEdBox1) & "收款" & ";" & Text9
    Else
        'Modify by Amy 2015/07/14 改寫至最前 原:最後
        'Text9 = MaskEdBox1 & "控管" & CboClass & ";" & Text9
        Text9 = FCDate(MaskEdBox1) & CboClass & ";" & Text9
     End If
     'end 2016/08/23
End Sub

Private Sub Combo1_GotFocus()
   StatusView MsgText(65) & "100"
End Sub

Private Sub Combo1_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

Private Sub Combo1_LostFocus()
Dim m_CU173 As String

   'Modify By Sindy 2017/3/24
   If Combo1.Tag <> Combo1.Text Then
      m_CU173 = ""
      'Modify By Sindy 2019/5/22 + Text2
      Call GetTitleCustData(Combo1.Text, Text2, "", , , , , , , , , , , , , , , , , , , , , , , , , , , , m_CU173)
      txtPrintNo.Text = m_CU173
   End If
   StatusView MsgText(601)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If CheckLen(Label4, Combo1.Text, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Command2_Click()
   'If adoacc0k0.RecordCount = 0 Or Text1 = MsgText(601) Then
   '   Exit Sub
   'End If
   'adoacc0k0.Find "a0k01 = '" & Text1 & "'", 0, adSearchForward, 1
   'If adoacc0k0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   RecordShow
   'Else
   '   Text1 = ""
   '   MsgBox MsgText(33), , MsgText(5)
   '   adoacc0k0.MoveFirst
   '   AdodcRefresh
   'End If
   Acc0k0Refresh
   If adoacc0k0.RecordCount <> 0 Then
      'Modify by Amy 2014/09/24 若為境外公司 只能為1.個人且不可改
      If PUB_GetTaxNo(Combo1, 1) = "Y" Then
          Text5 = "1"
          Text5.Locked = True
      Else
          Text5.Locked = False
      End If
      'end 2014/09/24
      
      FormShow
      RD06Show 'Add By Amy 2013/06/14 增加顯示預訂收款日
      AdodcRefresh
      RecordShow
   End If
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command2_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   'If adoacc0k0.RecordCount <> 0 Then
   '   adoacc0k0.MoveFirst
   '   AdodcRefresh
   'End If
   'adoacc0k0.Find "a0k01 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adoacc0k0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   RecordShow
   'End If
   Text1 = strItemNo
   Acc0k0Refresh
   If adoacc0k0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      RecordShow
      RD06Show 'Added by Lydia 2016/04/28
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/06 原W:9048／Ｈ:5700
   Me.Width = 9255 '8850
   Me.Height = 5900 '5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'strItemNo = MsgText(601) 'Remove by Lydia 2016/01/19
   'Add by Amy 2015/04/07 +控管類別及日期
   SetCombo CboClass
   MaskEdBox1.Mask = DFormat
   'end 2015/04/07

    OpenTable
    If adoacc0k0.RecordCount <> 0 Then
       adoacc0k0.MoveLast
       adoacc0k0.MoveFirst
       RecordShow
    End If
    
    'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
    Label17.Visible = False
    Label18.Visible = False
    
    LblA0k20_N = "" 'Add By Sindy 2021/5/10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   'Added by Lydia 2016/01/19 點選呼叫收據抬頭修改
   Select Case strTitle
      Case "Frmacc1220"
         tool3_enabled
         Frmacc1220.Enabled = True
      Case "Frmacc1230"
         tool3_enabled
         Frmacc1230.Enabled = True
      'Added by Lydia 2016/01/20
      Case "Frmacc1240"
         tool3_enabled
         Frmacc1240.Enabled = True
      'Added by Lydia 2016/01/20
      Case "Frmacc1211"
         tool3_enabled
         Frmacc1211.Enabled = True
      'Added by Lydia 2016/04/13
      Case "Frmacc1420"
         tool3_enabled
         'Modified by Lydia 2016/04/14
         'Frmacc1420.Enabled = True
         TmpFrmacc1420.Enabled = True
      'Add By Sindy 2016/6/8
      Case "Frmacc12d0"
         tool3_enabled
         Frmacc12d0.Enabled = True
      'Add By Sindy 2017/11/2
      Case "Frmacc11c0"
         tool3_enabled
         '有異動資料,重新查詢
         If bolUpdData = True Then
            Frmacc11c0.cmdSearch_Click
         End If
         Frmacc11c0.Enabled = True
   End Select
   strItemNo = ""
   strCustNo = ""
   strTitle = ""
   'end 2016/01/19
   
   Set TmpFrmacc1420 = Nothing 'Added by Lydia 2016/04/14
   
   MenuEnabled
   
   Set Frmacc1140 = Nothing
End Sub

'控管日期
Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        Exit Sub
    End If
    'Add by Amy 2015/05/22 +日期判斷否則判斷是否為工作日時會錯
    If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
        MsgBox Label25 & MsgText(63), , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    'end 2015/05/22
    'Mark by Amy 2016/08/23 智權人員提供的日期可能為非工作日,所以要可輸非工作日-瑞婷
'    If ChkWorkDay(FCDate(MaskEdBox1.Text) + 19110000) = False Then
'        MsgBox Label25 & "請輸入工作日！", vbExclamation, "日期錯誤！"
'        Cancel = True
'        MaskEdBox1.SetFocus
'        Exit Sub
'    End If
End Sub

'Add By Sindy 2015/8/26
'收據日期
Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      Exit Sub
   End If
   '日期判斷否則判斷是否為工作日時會錯
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label5 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   If ChkWorkDay(FCDate(MaskEdBox2.Text) + 19110000) = False Then
      MsgBox Label5 & "請輸入工作日！", vbExclamation, "日期錯誤！"
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   'Add by Amy 2023/08/16 +不可小於830101
   If Val(FCDate(MaskEdBox2.Text) + 19110000) < 19940101 Then
      MsgBox Label5 & "不可小於83/01/01！", vbExclamation, "日期錯誤！"
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.MaxRecords = intMax
   adoacc0k0.Open "select * from acc0k0 where (A0K09 IS NULL OR a0k09 = 0) and a0k01 >= '" & Text1 & "' order by a0k01", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'Modify By Sindy 2012/11/12 +cp151,cp05
   'Modified by Lydia 2018/09/12+cp27
   adoadodc1.Open "select a.*,getcp10desc(cp01,cp10,a0j04) cp10N,na03,cp151,cp05,cp31,sqldatet(cp27) cp27t from acc0j0 a,caseprogress,nation where a0j13 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 order by a0j01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國內收據資料)
'
'*************************************************
Public Sub FormShow()
Dim dblA1V06 As Double 'Add By Sindy 2016/6/17
Dim m_CU173 As String
    
   HasA4319 = False: HasA4321 = False 'Add by Amy 2021/01/28
   Text1 = adoacc0k0.Fields("a0k01").Value
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox1.Mask = MsgText(601)
   'add by sonia 2023/5/26   分所操作者檢查
   If ProState = "2" Then
      '檢查所操作者與收據智權人員的所別
      If PUB_GetST06(strUserNum) <> "1" And PUB_GetST06(strUserNum) <> PUB_GetST06("" & adoacc0k0.Fields("a0k20").Value) Then
         MsgBox "不可跨所修改收據資料！", , MsgText(5)
         MaskEdBox2.Text = MsgText(601): Text10 = MsgText(601): LblA0k20_N = MsgText(601): Text9 = MsgText(601): Text2 = MsgText(601)
         Combo1.Text = MsgText(601): Text5 = MsgText(601): Text4 = MsgText(601): Text7 = MsgText(601): Text8 = Val(Text4) + Val(Text7)
         Text12.Text = MsgText(601): Me.Text13.Text = MsgText(601): Text14.Text = MsgText(601): Combo2.Text = MsgText(601): Label21 = MsgText(601)
         MaskEdBox1.Text = MsgText(601)
         Forms(0).Toolbar1.Buttons.Item(5).Enabled = False
         Exit Sub
      End If
      '檢查已收款不可修改
      If PUB_GetST06(strUserNum) <> "1" And Val("" & adoacc0k0.Fields("a0k17").Value) + Val("" & adoacc0k0.Fields("a0k18").Value) > 0 Then
         MsgBox "已收款不可修改收據資料！", , MsgText(5)
         MaskEdBox2.Text = MsgText(601): Text10 = MsgText(601): LblA0k20_N = MsgText(601): Text9 = MsgText(601): Text2 = MsgText(601)
         Combo1.Text = MsgText(601): Text5 = MsgText(601): Text4 = MsgText(601): Text7 = MsgText(601): Text8 = Val(Text4) + Val(Text7)
         Text12.Text = MsgText(601): Me.Text13.Text = MsgText(601): Text14.Text = MsgText(601): Combo2.Text = MsgText(601): Label21 = MsgText(601)
         MaskEdBox1.Text = MsgText(601)
         Forms(0).Toolbar1.Buttons.Item(5).Enabled = False
         Exit Sub
      End If
   End If
   Forms(0).Toolbar1.Buttons.Item(5).Enabled = True
   'end 2023/5/26
      
   If IsNull(adoacc0k0.Fields("a0k02").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0k0.Fields("a0k02").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0k0.Fields("a0k19").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc0k0.Fields("a0k19").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
      Text11 = MsgText(601)
      LblA0k20_N = MsgText(601) 'Add By Sindy 2021/5/10
   Else
      Text11 = adoacc0k0.Fields("a0k20").Value
      LblA0k20_N = GetPrjSalesNM(Text11) 'Add By Sindy 2021/5/10
   End If
   If IsNull(adoacc0k0.Fields("a0k08").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc0k0.Fields("a0k08").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0k0.Fields("a0k03").Value
   End If
   SetReceiptTitle Text2.Text, Combo1 'Add by Morgan 2006/12/7 --瑞婷
   If IsNull(adoacc0k0.Fields("a0k04").Value) Then
      Combo1.Text = MsgText(601)
   Else
      Combo1.Text = adoacc0k0.Fields("a0k04").Value
   End If
   '記錄原收據抬頭
   Combo1.Tag = Combo1.Text
   'End
   'Add By Sindy 2013/12/27
   m_ShowMsg = ""
   Combo1.Enabled = True
   If "" & adoacc0k0.Fields("a4301").Value <> "" Then '已開發票
      '發票已申報,收據抬頭鎖住
      If Val(GetInvDataA4111(adoacc0k0.Fields("a4302").Value)) > 0 Then
         Combo1.Enabled = False
      End If
      '有未收款沖帳傳票
      If "" & adoacc0k0.Fields("A4317").Value <> "" Then
         strSql = "select * From acc021 where ax201='J' and ax202='" & adoacc0k0.Fields("A4317").Value & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            '傳票已過帳,收據抬頭鎖住
            If Val("" & RsTemp.Fields("ax210")) > 0 Then
               Combo1.Enabled = False
            Else
               m_ShowMsg = "已產生未收款沖帳傳票"
            End If
         End If
      End If
   End If
   '2013/12/27 END
   If IsNull(adoacc0k0.Fields("a0k05").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc0k0.Fields("a0k05").Value
   End If
   'Add by Amy 2021/01/28
   Text5.Tag = Text5
   If Not IsNull(adoacc0k0.Fields("a4319")) Then
      HasA4319 = True
   End If
   If Not IsNull(adoacc0k0.Fields("a4321")) Then
      HasA4321 = True
   End If
   'end 2021/01/28
   If IsNull(adoacc0k0.Fields("a0k07").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0k0.Fields("a0k07").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k06").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = adoacc0k0.Fields("a0k06").Value
   End If
   Text8 = Val(Text4) + Val(Text7)
   'Add By Cheng 2003/04/21
   '公司別
   Me.Text12.Text = "" & adoacc0k0.Fields("a0k11").Value
   Me.Text12.Tag = Me.Text12 'Add by Amy 2020/04/24
   'Add By Sindy 2013/12/27 J公司若已開發票則不可改公司別
   Me.Text12.Enabled = True
   If Me.Text12.Text = "J" Then
      If "" & adoacc0k0.Fields("a4301").Value <> "" Then
         Me.Text12.Enabled = False
      End If
   End If
   '2013/12/27 END
   '扣繳年度
   Me.Text13.Text = "" & adoacc0k0.Fields("a0k16").Value
   'Add by Amy 2013/07/22 讀出的a0k16先暫存 for 儲存比較
   Me.Text13.Tag = "" & Me.Text13.Text
   
   'Modify By Sindy 2014/2/17 該收據已收到扣繳資料,不必RunJ公司的控管,因J公司沒有扣繳問題
   'Add By Cheng 2004/03/15
   '若該收據已收到扣繳資料, 則扣繳年度欄位鎖住
   'Me.Text13.Enabled = Not ChkA1V15(Me.Text1.Text)
   'Modify By Sindy 2016/6/17 + dblA1V06
   If ChkA1V15(Me.Text1.Text, dblA1V06) = True Then
      Me.Text13.Enabled = False
      'add by sonia 2019/5/23
      Combo1.Enabled = False
      MsgBox "已輸扣繳憑單，不可修改收據抬頭及扣繳年度！", , MsgText(5)
      'end 2019/5/23
   Else
   'End
      'Add By Sindy 2013/12/27 J公司沒有扣繳問題,扣繳年度欄要鎖住
      Me.Text13.Enabled = True
      Combo1.Enabled = True   'add by sonia 2019/5/23
      If Me.Text12.Text = "J" Then
         Me.Text13.Text = ""
         Me.Text13.Enabled = False
      End If
   End If
   '2014/2/17 END
   'Add By Sindy 2016/6/17 以收據編號檢查acc1v0,若a1v06>0並且收據別為公司時,收據別欄位要鎖住
   If Text5 = "2" And dblA1V06 > 0 Then
      Me.Text5.Enabled = False
   Else
      Me.Text5.Enabled = True
   End If
   '2016/6/17 END
   
   '2013/12/27 END
   '記錄已收服務費
   m_dbla0k17 = Val("" & adoacc0k0.Fields("a0k17").Value)
   'Add By Cheng 2003/12/04
   '記錄已收規費
   m_dbla0k18 = Val("" & adoacc0k0.Fields("a0k18").Value)
   'End
   'Add By Cheng 2003/12/09
   '顯示最近修改抬頭日期
   If Val("" & adoacc0k0.Fields("a0k31").Value) = 0 Then
       Me.Text14.Text = ""
   Else
       Me.Text14.Text = "" & adoacc0k0.Fields("a0k31").Value
   End If
   
   'Add By Sindy 2010/5/5
   If Trim(adoacc0k0.Fields("a0k32").Value) = "N" Then
      Check1.Value = 1
      '2013/11/18 add by sonia
      Text10.Locked = True
      Text10.Enabled = False
      '2013/11/18 end
   Else
     Check1.Value = 0
      '2013/11/18 add by sonia
      Text10.Locked = False
      Text10.Enabled = True
      '2013/11/18 end
   End If
   '2010/5/5 End
   'Added by Lydia 2023/12/12
   m_strA0k32 = "" & adoacc0k0.Fields("a0k32").Value
   'Z=確定不印，鎖住「收據暫不列印」，存檔時回存讀出的值。
   If m_strA0k32 = "Z" Then
      Check1.Enabled = False
   Else
      Check1.Enabled = True
   End If
   'end 2023/12/12
   
   'Add By Sindy 2013/12/27
   If IsNull(adoacc0k0.Fields("a0k34").Value) Then
      'Modify By Sindy 2014/12/29
      'txtSales = MsgText(601): lblSales = ""
      Combo2.Text = ""
      '2014/12/29 END
   Else
      'Modify By Sindy 2014/12/29
      'txtSales = adoacc0k0.Fields("a0k34").Value
      'Call txtSales_Validate(False)
      Combo2 = adoacc0k0.Fields("a0k34").Value
      Call Combo2_LostFocus
      '2014/12/29 END
   End If
   If IsNull(adoacc0k0.Fields("a0k36").Value) Then
      Label21 = MsgText(601)
   Else
      Label21 = CFDate(adoacc0k0.Fields("a0k36").Value)
   End If
   '2013/12/27 END
   'Add by Amy 2015/04/07 +控管日期及類別
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0k0.Fields("a0k38").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0k0.Fields("a0k38").Value)
   End If
   MaskEdBox1.Tag = "" & adoacc0k0.Fields("a0k38").Value 'Add by Amy 2015/04/17
   MaskEdBox1.Mask = DFormat
   CboClass = "" & adoacc0k0.Fields("a0k39").Value
   CboClass.Tag = "" & adoacc0k0.Fields("a0k39").Value 'Add by Amy 2015/04/17
   'Add By Sindy 2017/3/17 是否列印統一編號
   'Modify By Sindy 2017/3/24
   m_CU173 = ""
   'Modify By Sindy 2019/5/22 + Text2
   Call GetTitleCustData(Combo1.Text, Text2, "", , , , , , , , , , , , , , , , , , , , , , , , , , , , m_CU173)
   txtPrintNo.Text = m_CU173
'   If IsNull(adoacc0k0.Fields("a0k40").Value) Then
'      txtPrintNo.Text = ""
'   Else
'      txtPrintNo.Text = "" & adoacc0k0.Fields("a0k40").Value
'   End If
   '2017/3/24 END
   '2017/3/17 END
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'ADD BY SONIA 2014/5/28 J公司不可暫不列印-婷
Private Sub Text12_Change()
   'cancel by sonia 2017/12/6婷要求開放
   'If Text12 = "J" Then
   '   Check1.Value = 0
   '   Check1.Enabled = False
   '   'add by sonia 2015/11/26
   '   Check2(0).Value = 0: Check2(1).Value = 0: Check2(2).Value = 0
   '   Check2(0).Enabled = False: Check2(1).Enabled = False: Check2(2).Enabled = False
   '   'end 2015/11/26
   'Else
   'end 2017/12/6
      If m_strA0k32 <> "Z" Then 'Added by Lydia 2023/12/12 判斷Z.確定不列印
        Check1.Enabled = True
      End If 'Added by Lydia 2023/12/12
      'add by sonia 2015/11/26
      Check2(0).Enabled = True: Check2(1).Enabled = True: Check2(2).Enabled = True
      'end 2015/11/26
   'End If  'cancel by sonia 2017/12/6婷要求開放
End Sub
   'END 2014/5/28

Private Sub Text12_GotFocus()
   TextInverse Me.Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)
    'Add By Cheng 2003/05/12
    KeyEnter KeyCode
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   Dim strMsg As String 'Add by Amy 2020/03/24
   
   If Text12 = MsgText(601) Then
      MsgBox MsgText(188) & Label12, , MsgText(5)
      Cancel = True
      Text12.SetFocus
      Exit Sub
   Else
     'Moidfy by Amy 2020/03/24 寫法改一致,從Frmacc1140_Save搬過來
'      If adocheck.State = adStateOpen Then
'         adocheck.Close
'      End If
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select * from acc080 where a0801 = '" & Text12 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount = 0 Then
'         MsgBox MsgText(188) & Label12, , MsgText(5)
'         adocheck.Close
'         Cancel = True
'         Text12.SetFocus
'         Exit Sub
'      End If
'      adocheck.Close
      If ExistCheck("acc080", "a0801", Text12, "", False) = False Then
            MsgBox MsgText(45) & Label12, , MsgText(5)
            Cancel = True
            Text12.SetFocus
            Exit Sub
      End If
     'end 2020/03/24
      'Add by Amy 2020/03/24
      If Text12.Enabled = True Then
            If strSrvDate(1) >= 事務所合併日 Then
                'Moidfy by Amy 2020/04/24 +if 有修改才判斷 ex:E10906599 原:1公司 只改扣繳欄會無法存檔
                If Text12.Tag <> Text12 Then
                    If ChkAccReceiptComp(0, Text12, strMsg) = False Then
                        MsgBox Replace(strMsg, "收據公司別", "公司別"), , MsgText(5)
                        Cancel = True
                        Text12.SetFocus
                        Exit Sub
                    End If
                End If
            ElseIf Text12 <> "1" And Text12 <> "2" And Text12 <> "9" And Text12 <> "J" Then
                MsgBox "公司別輸入有誤，請確認！", , MsgText(5)
                Cancel = True
                Text12.SetFocus
                Exit Sub
            End If
      End If
      
      'ADD BY SONIA 2014/5/28 J公司不可暫不列印-婷
      'cancel by sonia 2017/12/6婷要求開放
      'If Text12 = "J" Then
      '   Check1.Value = 0
      '   Check1.Enabled = False
      '   'add by sonia 2015/11/26
      '   Check2(0).Value = 0: Check2(1).Value = 0: Check2(2).Value = 0
      '   Check2(0).Enabled = False: Check2(1).Enabled = False: Check2(2).Enabled = False
      '   'end 2015/11/26
      'Else
      'end 2017/12/6
         If m_strA0k32 <> "Z" Then 'Added by Lydia 2023/12/12 判斷Z.確定不列印
            Check1.Enabled = True
         End If  'Added by Lydia 2023/12/12
         'add by sonia 2015/11/26
         Check2(0).Enabled = True: Check2(1).Enabled = True: Check2(2).Enabled = True
         'end 2015/11/26
      'End If  'cancel by sonia 2017/12/6婷要求開放
      'END 2014/5/28
   End If
End Sub

Private Sub Text13_GotFocus()
    'Add By Cheng 2003/04/24
    TextInverse Me.Text13
End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)
    'Add By Cheng 2003/05/12
    KeyEnter KeyCode
End Sub

Private Sub Text15_GotFocus()
    TextInverse Me.Text15
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text2_Change()
   If Text2 = MsgText(601) Then
      Text3 = ""
      Exit Sub
   End If
   Text3 = CustomerQuery(Text2, 1)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Text9_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strAppl As String 'Add By Sindy 2015/6/22
   
On Error GoTo Checking
   
   '92.6.28 ADD BY SONIA
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select DISTINCT(A0M01) from acc0M0 where a0M02 = '" & Text1 & "' order by a0M01 asc", adoTaie, adOpenStatic, adLockReadOnly
   cboA0M01.Clear
   If adoadodc1.RecordCount <> 0 Then
      Do While adoadodc1.EOF = False
         cboA0M01.AddItem adoadodc1.Fields("A0M01")
         adoadodc1.MoveNext
      Loop
   End If
   If cboA0M01.ListCount > 0 Then cboA0M01.ListIndex = 0
   '92.6.28 END
    'Add By Cheng 2004/02/02
    '顯示發票號碼
    If Text12 = "J" Then 'Add By Sindy 2013/12/27 +if
      ShowInvoiceNo Me.cboA0M01.Text, Me.Text1.Text
    Else
    '2013/12/27 END
      If Me.cboA0M01.Text <> "" Then
          ShowInvoiceNo Me.cboA0M01.Text, Me.Text1.Text
      Else
          Me.Text15.Text = ""
      End If
    End If
    'End
   'ADD BY SONIA 2013/11/14 加銷退日期
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select DISTINCT(A0S03) from acc0S0 where a0S02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   cboA0S03.Clear
   If adoadodc1.RecordCount <> 0 Then
      Do While adoadodc1.EOF = False
         cboA0S03.AddItem CFDate(adoadodc1.Fields("A0S03"))
         adoadodc1.MoveNext
      Loop
   End If
   If cboA0S03.ListCount > 0 Then cboA0S03.ListIndex = 0
   '2013/11/14 END
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   '2005/6/16 MODIFY BY SONIA 加收文號排序
   'adoadodc1.Open "select * from acc0j0 where a0j13 = '" & Text1 & "' order by a0j02 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'Modify By Sindy 2012/11/12 +cp151,cp05
   'Modify By Sindy 2015/6/22 +cp01,cp02,cp03,cp04
   'Modified by Lydia 2018/09/12+cp27
   'Modify By Sindy 2023/9/5 +cp10
   adoadodc1.Open "select a.*,getcp10desc(cp01,cp10,a0j04) cp10N,na03,cp151,cp05,cp31,sqldatet(cp27)cp27t,cp01,cp02,cp03,cp04,cp10 " & _
                             "from acc0j0 a,caseprogress,nation where a0j13 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 order by a0j02,A0J01 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2005/6/16 END
   Adodc1.Recordset.Requery
   'Add By Sindy 2012/11/12
   Me.Check2(0).Value = 0
   Me.Check2(1).Value = 0
   Me.Check2(2).Value = 0
   If adoadodc1.RecordCount >= 1 Then
      adoadodc1.MoveFirst
      If Check1.Value = 1 Then 'Add By Sindy 2014/1/14 +if 有暫不列印時,才需要帶出值
         If "" & adoadodc1.Fields("cp151") = "1" Then Me.Check2(0).Value = 1
         If "" & adoadodc1.Fields("cp151") = "2" Then Me.Check2(1).Value = 1
         If "" & adoadodc1.Fields("cp151") = "3" Then Me.Check2(2).Value = 1
      End If
      m_CP31 = "" 'Add By Sindy 2014/1/9
      If "" & adoadodc1.Fields("a0j01") <> "" Then
         m_CP09 = adoadodc1.Fields("a0j01") 'Add By Sindy 2012/12/6
         m_CP31 = "" & adoadodc1.Fields("cp31") 'Add By Sindy 2014/1/9
      End If
      
      'Add By Sindy 2020/4/28 (E10910184) L案號，以收文號抓法律所案源資料的LOS06，若其案源案件類型LOS02為C類時，
      '收據客戶編號自動設定為智慧所X03072010，抬頭為此客戶編號之名稱，同時將收據抬頭鎖住。
      '例L-006203(收文號AA9014217)
      intI = 0
      If m_CP09 <> "" Then
         strExc(0) = "select * from lawofficesource where los06='" & m_CP09 & "' and los02='C'"
         intI = 1
         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 And adoadodc1.Fields("cp01") = "L" Then
            '不檢查申請人1
            Combo1.Enabled = False '收據抬頭要鎖住
         Else
            intI = 0
         End If
      End If
      If intI = 0 Then
      '2020/4/28 END
         'Add By Sindy 2015/6/22 檢查案件的第一申請人編號的前八碼若與收據a0k03的前八碼是否不同
         strAppl = GetPrjPeopleNum1(adoadodc1.Fields("cp01") & "-" & adoadodc1.Fields("cp02") & "-" & adoadodc1.Fields("cp03") & "-" & adoadodc1.Fields("cp04"))
         If Left(strAppl, 8) <> Left(Text2, 8) And Combo1.Text <> MsgText(601) Then
            MsgBox "此案件已換申請人，請作廢收據重開！", , MsgText(5)
         End If
         '2015/6/22 END
      End If
      
      'Add By Sindy 2015/7/22
      m_CP01 = Left(adoadodc1.Fields("a0j02"), Len(adoadodc1.Fields("a0j02")) - 9)
      m_CP02 = Mid(adoadodc1.Fields("a0j02"), Len(adoadodc1.Fields("a0j02")) - 8, 6)
      m_CP03 = Mid(adoadodc1.Fields("a0j02"), Len(adoadodc1.Fields("a0j02")) - 2, 1)
      m_CP04 = Right(adoadodc1.Fields("a0j02"), 2)
      'Modified by Morgan 2015/12/15
      'If ChkPatentNameCompany(m_CP01, m_CP02, m_CP03, m_CP04) = "J" Then
      If ChkPatentNameCompany(m_CP01, m_CP02, m_CP03, m_CP04) = Text12 Then
      'end 2015/12/15
         Text12.Enabled = False
      'Add by Amy 2020/03/24
      ElseIf strSrvDate(1) >= 智慧所更名日 Then
        '法務/顧問案不可改收據公司別
        'Modify by Amy 2020/04/15 原:False
        If ChkAccReceiptComp(1, m_CP01) = True Then
            Text12.Enabled = False
        End If
      End If
      '2015/7/22 END
   End If
   '2012/11/12 End
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理國內收據資料
'
'*************************************************
Public Sub Acc0k0Refresh()
On Error GoTo Checking
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.MaxRecords = intMax
   'Modify By Sindy 2013/12/27
   'adoacc0k0.Open "select * from acc0k0 where (a0k09 IS NULL OR A0K09 = 0) and a0k01 >= '" & Text1 & "' order by a0k01", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0k0.Open "select * from acc0k0,acc430,acc431 where (a0k09 IS NULL OR A0K09 = 0) and a0k01 >= '" & Text1 & "' and a0k01=axc02(+) and axc01=a4301(+) order by a0k01", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2013/12/27 END
'   If adoacc0k0.RecordCount <> 0 Then
'      If Text1 <> MsgText(601) Then
'         adoacc0k0.Find "a0k01 = '" & Text1 & "'", 0, adSearchForward, 1
'         If adoacc0k0.EOF = False Then
'            FormShow
'            AdodcRefresh
'            RecordShow
'         End If
'      End If
'   End If
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
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0k0.Bookmark & MsgText(35) & adoacc0k0.RecordCount
End Sub

'Add By Cheng 2004/02/02
'顯示發票號碼
Private Sub ShowInvoiceNo(strA0M01 As String, strA0M02 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'Add By Sindy 2013/12/27
If Text12 = "J" Then
   StrSQLa = "Select AXC01 From ACC431 Where AXC02='" & strA0M02 & "'"
Else
'2013/12/27 END
   StrSQLa = "Select A0M03 From ACC0M0 Where A0M01='" & strA0M01 & "' And A0M02='" & strA0M02 & "' "
End If
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    Me.Text15.Text = "" & rsA.Fields(0).Value
Else
    Me.Text15.Text = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

'Add By Cheng 2004/03/15
'檢查收據是否已收到扣繳憑單
'Modify By Sindy 2016/6/17 Optional ByRef dblA1V06 As Double = 0
Private Function ChkA1V15(strA1V02 As String, Optional ByRef dblA1V06 As Double = 0) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select * From ACC1V0 Where A1V02='" & strA1V02 & "' And A1V15 Is Not Null "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   ChkA1V15 = True
Else
   ChkA1V15 = False
End If
rsA.Close

'Add By Sindy 2016/6/17
StrSQLa = "Select * From ACC1V0 Where A1V02='" & strA1V02 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   dblA1V06 = CDbl("" & rsA.Fields("A1V06"))
End If
rsA.Close
'2016/6/17 END

Set rsA = Nothing
End Function

'Add By Morgan 2006/12/7
'設定此客戶及其關係企業曾經開過的收據抬頭
Private Sub SetReceiptTitle(p_CustNo As String, p_Combo As Object)
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'strExc(0) = "Select Distinct A0K04 From ACC0K0 Where A0K03>='" & Left(p_CustNo, 6) & "000' and A0K03<='" & Left(p_CustNo, 6) & "999' and A0K04 IS NOT NULL Order By 1 "
   strExc(0) = "Select Distinct A0K04 From ACC0K0 Where A0K03>='" & Left(p_CustNo, 6) & "000' and A0K03<='" & Left(p_CustNo, 6) & "ZZZ' and A0K04 IS NOT NULL Order By 1 "
   intI = 1
   'edit by nickc 2007/02/08 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      p_Combo.Clear
      Do While Not RsTemp.EOF
         p_Combo.AddItem RsTemp(0)
         RsTemp.MoveNext
      Loop
   End If
End Sub

Private Sub Text9_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

'Add by Amy 2013/06/14
'以收據號碼抓ACC0J0之A0J13(可能多筆),再以A0J01抓預定收款日期異動記錄ReceivablesDay之RD01,
'抓出所有資料異動日期+序號最大那一筆的預定收款日期顯示,若該筆RD06有值則畫面上預定收款日清空.(瑞婷)
Private Sub RD06Show()
    Dim strSql As String, strSQLR As String
    Dim rsR As New ADODB.Recordset
    'Modified by Lydia 2018/09/12 改為付款週期月份
'    strSQLR = "Select A0J01 From Acc0J0 Where A0J13='" & Text1 & "'"
'    strSql = "Select NVL(RD05,0) RD05,NVL(RD06,'NULL') RD06 From ReceivablesDay,(Select Max(RD02||RD03) RDMax From ReceivablesDay Where RD01 IN(" & strSQLR & ")) Where RD01 IN(" & strSQLR & ") And RD02=SubStr(RDMax,1,8) And RD03=SubStr(RDMax,9,1)"
    strSql = "select a0j11,decode(substr(a0j11,1,1),'X',nvl(cu175,2),null) cu175 from acc0j0,customer " & _
                     "where A0J13='" & Text1 & "' and substr(a0j11,1,8) = cu01(+) and substr(a0j11,9,1) = cu02(+) "
    'end 2018/09/12
    rsR.CursorLocation = adUseClient
    rsR.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If Not rsR.EOF And Not rsR.BOF Then
    'Modified by Lydia 2018/09/12
'        If rsR.Fields("RD06") = "Y" Then
'            Label18.Caption = ""
'        Else
'            Label18.Caption = IIf(rsR.Fields("RD05") = 0, "", CFDate(ACDate(rsR.Fields("RD05"))))
'        End If
        Label28.Caption = "" & rsR.Fields("cu175")
    End If
    If rsR.State <> adStateClosed Then rsR.Close
    Set rsR = Nothing
End Sub

''Add By Sindy 2013/12/15
'Private Sub txtSales_GotFocus()
'   txtSales.SelStart = 0
'   txtSales.SelLength = Len(txtSales.Text)
'   '儲存未修改前之值至Tag中,供再確認時使用
'   txtSales.Tag = txtSales
'   '切換輸入法
'   CloseIme
'End Sub
'
''Add By Sindy 2013/12/15
'Private Sub txtSales_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''Add By Sindy 2013/12/15
'Private Sub txtSales_Validate(Cancel As Boolean)
'Dim strTemp As String, strTemp1 As String
'
'   lblSales.Caption = ""
'   If txtSales.Text <> "" Then
'      If Not ClsPDGetStaff(txtSales.Text, strTemp, strTemp1) Then
'         Cancel = True
'         Exit Sub
'      End If
'      lblSales.Caption = strTemp
'   End If
'End Sub
'Add By Sindy 2014/12/29
Private Sub Combo2_GotFocus()
   InverseTextBox Combo2
End Sub
Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo2_LostFocus()
   If Combo2.Text > "" And Len(Trim(Combo2.Text)) = 5 Then
      '抓取員工姓名
      Combo2.Text = SetCboStaffName(Combo2.Text)
   End If
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(Combo2, 5)) = True Then
         Call Combo2_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
'2014/12/29 END

Public Sub Frmacc1140_Save()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim bolIs000 As Boolean
Dim bolCancel As Boolean
Dim strCP05 As String
Dim m_strChkCompany As String, m_strCaseNo As String, strSpecCompany As String
Dim strCU11 As String
Dim bolSaveCtrlTxt As Boolean 'Add by Amy 2015/04/07 備註欄是否存控管日期及類別
Dim bolIsACSTIPS As Boolean 'Add By Sindy 2023/9/5
Dim stMsg As String 'Add by Amy 2025/07/16

On Error GoTo Checking
   
   bolIsACSTIPS = False 'Add By Sindy 2023/9/5
   With Frmacc1140
  
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
            'Add By Cheng 2003/04/21
            '檢查欄位--公司別
            If .Text12.Text = MsgText(601) Then
                MsgBox MsgText(10) & .Label12, , MsgText(5)
                strControlButton = MsgText(602)
                .Text12.SetFocus
                Exit Sub
            End If
            'Modify by Amy 2020/03/24 ExistCheck搬至 Text12_Validate
            bolCancel = False
            Call Text12_Validate(bolCancel)
            If bolCancel = True Then
               strControlButton = MsgText(602)
               Exit Sub
            End If
            'end 2020/03/24
            'Add by Amy 2021/01/28 +發票未作廢判斷
            If .Text12 = "J" Then
                '有開發票上傳過發票,又修改抬頭或扣繳且未作廢,彈訊息且不可操作
                If Text15 <> MsgText(601) And HasA4319 = True And HasA4321 = False Then
                    If Combo1.Tag <> Combo1 Then
                        MsgBox "修改收據抬頭前需先將此發票作廢上傳至盟立!!!", vbExclamation + vbOKOnly
                        strControlButton = MsgText(602)
                        Combo1.SetFocus
                        Exit Sub
                    ElseIf Text5.Tag <> Text5 Then
                        MsgBox "修改扣繳前需先將此發票作廢上傳至盟立!!!", vbExclamation + vbOKOnly
                        strControlButton = MsgText(602)
                        Text5.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            'end 2021/01/28
            If .Text12.Text <> "J" Then 'Add By Sindy 2013/12/27 +if
               If .m_dbla0k17 <> 0 And .Text13.Text = "" Then
                   MsgBox "若已收服務費時，扣繳年度一定要輸!!!", vbExclamation + vbOKOnly
                   strControlButton = MsgText(602)
                   .Text13.SetFocus
                   Exit Sub
               End If
            End If
            'Modify by Amy 2014/09/24 為境外公司 只能為1.個人
            If PUB_GetTaxNo(.Combo1, 1) = "Y" And .Text5 <> "1" Then
                MsgBox "此收據抬頭為境外公司不可設 2.公司", vbExclamation + vbOKOnly
                strControlButton = MsgText(602)
                '.Text5.SetFocus
                Exit Sub
            End If
            'end 2014/09/24
            'Modify by Amy 2015/04/07 +控管日期及類別
            bolCancel = False
            Call CboClass_Validate(bolCancel)
            Call MaskEdBox1_Validate(bolCancel)
            Call MaskEdBox2_Validate(bolCancel) 'Add By Sindy 2015/8/26
            If bolCancel = True Then strControlButton = MsgText(602): Exit Sub
            If CboClass <> MsgText(601) And (MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29)) Then
                MsgBox Label23 & "有資料," & Label25 & "不可為空!!"
                strControlButton = MsgText(602)
                .MaskEdBox1.SetFocus
                Exit Sub
            End If
            If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) And CboClass = MsgText(601) Then
                MsgBox Label25 & "有資料," & Label23 & "不可為空!!"
                strControlButton = MsgText(602)
                .CboClass.SetFocus
                Exit Sub
            End If
            'Add by Amy 2015/04/17 +if 判斷不為空值且有改沒存才彈訊息
            If MaskEdBox1.Text <> MsgText(601) Or MaskEdBox1.Text <> MsgText(29) Or CboClass <> MsgText(601) Then
                If (Val(MaskEdBox1.Tag) & CboClass.Tag) <> (Val(FCDate(MaskEdBox1)) & CboClass) Then
                    'Modify by Amy 2016/08/23 取消"控管"字樣
                    'If InStr(Text9, MaskEdBox1 & "控管" & CboClass) = 0 Then
                    'Modify by Amy 2016/08/29 修正因2016/08/23修改 預計年月日收款 造成bug
                    If (CboClass <> "預計收款" And InStr(Text9, FCDate(MaskEdBox1) & CboClass) = 0) Or (CboClass = "預計收款" And InStr(Text9, "預計" & FCDate(MaskEdBox1) & "收款") = 0) Then
                        strExc(0) = MsgBox("此次控管內容未加入備註欄,是否要存入備註欄？", vbYesNoCancel + vbQuestion + vbDefaultButton3)
                        If strExc(0) = vbCancel Then
                           strControlButton = MsgText(602)
                          Exit Sub
                       ElseIf strExc(0) = vbYes Then
                          bolSaveCtrlTxt = True
                       End If
                    End If
                End If
            End If
            'end 2015/04/17
            'end 2015/04/07
            'Add By Sindy 2012/11/12
            bolIs000 = True
            strCP05 = ""
            m_strChkCompany = "": m_strCaseNo = "" 'Add By Sindy 2014/1/9
            If .adoadodc1.RecordCount >= 1 Then
               .adoadodc1.MoveFirst
               Do While Not .adoadodc1.EOF
                  If .adoadodc1.Fields("a0j04") <> "000" Then
                     bolIs000 = False
                  End If
                  strCP05 = .adoadodc1.Fields("CP05")
                  'Add By Sindy 2014/1/9
                  m_CP01 = Left(.adoadodc1.Fields("a0j02"), Len(.adoadodc1.Fields("a0j02")) - 9)
                  m_CP02 = Mid(.adoadodc1.Fields("a0j02"), Len(.adoadodc1.Fields("a0j02")) - 8, 6)
                  m_CP03 = Mid(.adoadodc1.Fields("a0j02"), Len(.adoadodc1.Fields("a0j02")) - 2, 1)
                  m_CP04 = Right(.adoadodc1.Fields("a0j02"), 2)
                  m_CP10 = adoadodc1.Fields("cp10") 'Add By Sindy 2023/9/5
                  If InStr(m_strCaseNo, .adoadodc1.Fields("a0j02")) = 0 Then
                     strSpecCompany = ChkPatentNameCompany(m_CP01, m_CP02, m_CP03, m_CP04)
                     If strSpecCompany <> "" And (strSpecCompany = m_strChkCompany Or m_strChkCompany = "") Then
                        m_strChkCompany = strSpecCompany
                        If m_strCaseNo <> "" Then m_strCaseNo = m_strCaseNo & ","
                        m_strCaseNo = m_strCaseNo & .adoadodc1.Fields("a0j02")
                     End If
                  End If
                  '2014/1/9 END
                  
                  'Add By Sindy 2023/9/5 檢查是否為ACS案件的TIPS
                  If PUB_ChkACSforTIPS(m_CP01 & m_CP02 & m_CP03 & m_CP04) = True Then
                     '僅一筆文號且為代收代付時,不能鎖定收據暫不列印
                     If Not (.adoadodc1.RecordCount = 1 And m_CP10 = "706") Then
                        bolIsACSTIPS = True
                     End If
                  End If
                  '2023/9/5 END
                  
                  .adoadodc1.MoveNext
               Loop
            End If
            'Add By Sindy 2013/12/17
            If m_strChkCompany = "T" And Text12 <> "1" And m_CP31 = "Y" Then
               MsgBox "專利案" & m_strCaseNo & "有設定以專利商標出名不可開立其他公司別，請與專業部確認!!", vbCritical, "收據公司別提醒"
               strControlButton = MsgText(602)
               Text12.SetFocus
               Exit Sub
            ElseIf m_strChkCompany = "J" And Text12 <> "J" And m_CP31 = "Y" Then
               MsgBox m_strCaseNo & "有設定以智權公司出名不可開立其他公司別，請與專業部確認!!", vbCritical, "收據公司別提醒"
               strControlButton = MsgText(602)
               Text12.SetFocus
               Exit Sub
            End If
            '2013/12/17 END
            
            'Add By Sindy 2023/9/5
            If bolIsACSTIPS = False Then
            '2023/9/5 END
               '台灣案,則為送件日
               If bolIs000 = True And (.Check2(1).Value = 1 Or .Check2(2).Value = 1) Then
                  MsgBox "非台灣案時, 收據自動列印時間點才可選擇2 或 3 !!!", vbExclamation + vbOKOnly
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
               If .Check1.Value = 1 And _
                  (.Check2(0).Value = 0 And .Check2(1).Value = 0 And .Check2(2).Value = 0) And _
                  Val(strCP05) >= 20121115 Then
                  MsgBox "勾選收據暫不列印時, 收據自動列印時間點不可空白!!!", vbExclamation + vbOKOnly
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
               '2012/11/12 End
               'add by Sindy 2013/12/25
               If .Check1.Value = 0 And _
                  (.Check2(0).Value = 1 Or .Check2(1).Value = 1 Or .Check2(2).Value = 1) Then
                  MsgBox "點選收據自動列印時間點, 收據暫不列印一定要勾選!!!", vbExclamation + vbOKOnly
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
               '2013/12/25 end
            End If
            
            If .Check1.Value = 1 Then 'Added by Morgan 2013/3/21 收據暫不列印才要檢查
               'Add By Sindy 2012/12/6
               '檢查是否可上收據自動列印時間點
               If PUB_ChkAccIsUpdCP151(.m_CP09, IIf(.Check2(0).Value = 1, "1", IIf(.Check2(1).Value = 1, "2", IIf(.Check2(2).Value = 1, "3", "")))) = False Then
                  .Check2(0).Value = 0
                  .Check2(1).Value = 0
                  .Check2(2).Value = 0
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
               '2012/12/6 End
            End If
         If CheckLen(.Label4, .Combo1.Text, 100) = MsgText(603) Then
            strControlButton = MsgText(602)
            .Combo1.SetFocus
            Exit Sub
         End If
         If .adoquery.State = adStateOpen Then
            .adoquery.Close
         End If
         .adoquery.CursorLocation = adUseClient
         .adoquery.Open "select * from acc0m0 where a0m02 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount <> 0 Then
            MsgBox MsgText(194), , MsgText(21)
         End If
         .adoquery.Close
      End If
      
      If .Combo1.Text <> MsgText(601) Then
         .adoacc0k0.Fields("a0k04").Value = .Combo1.Text
      Else
         .adoacc0k0.Fields("a0k04").Value = Null
      End If
      If .Text5 <> MsgText(601) Then
         .adoacc0k0.Fields("a0k05").Value = .Text5
      Else
         .adoacc0k0.Fields("a0k05").Value = Null
      End If
      'Modify by Amy 2014/04/07 +控管日期及類別
      If bolSaveCtrlTxt = True Then
        'Modify by Amy 2016/08/23 取消日期中的/及"控管"字樣;"預計//日收款"改為 "預計年月日收款"
        If CboClass = "預計收款" Then
            'Add by Amy 2015/08/24 +若選「預計收款」增加「預計//日收款」文字
            Text9 = "預計" & FCDate(MaskEdBox1) & "收款" & ";" & Text9
        Else
            'Modify by Amy 2015/07/14 改寫至最前 原:最後
            Text9 = FCDate(MaskEdBox1) & CboClass & ";" & Text9
        End If
        'end 2016/08/23
      End If
      If .Text9 <> MsgText(601) Then
         .adoacc0k0.Fields("a0k08").Value = .Text9
      Else
         .adoacc0k0.Fields("a0k08").Value = Null
      End If
      'end 2014/04/07
      
      'Modify By Sindy 2013/12/27 取消
'      '公司別
'      'Add by Morgan 2007/3/28 若有發票號碼且不為'E'字頭時更新收據公司為9
'      If .Combo1.Text <> "" And .Text15.Text <> "" And Left(.Text15.Text, 1) <> "E" Then
'         .Text12 = "9"
'      End If
      'end 2007/3/28
      
      .adoacc0k0.Fields("a0k02").Value = FCDate(MaskEdBox2.Text) 'Add By Sindy 2015/12/31
      .adoacc0k0.Fields("a0k11").Value = .Text12
        '扣繳年度
        'Modify By Cheng 2003/04/24
'        .adoacc0k0.Fields("a0k16").Value = .Text13
        .adoacc0k0.Fields("a0k16").Value = Val(.Text13.Text)
        
        'Modify by Amy 2013/07/22  +判斷有改扣扣繳年度A0K16時(讀出的與畫面不同)才更新
        If Val(.Text13.Tag) <> Val(.Text13.Text) Then
            'Modify By Amy 2013/06/20 修改扣繳年度A0K16時，同時更新A0K15為系統日-瑞婷
            .adoacc0k0.Fields("a0k15").Value = Val(strSrvDate(2))
            'Add by Amy 2025/07/16 有修改[扣繳年度]欄,寫入[收據扣繳年度異動檔]
            If UpdateAcc290(stMsg) = False Then
               GoTo Checking
            End If
            'end 2025/07/16
        End If
        
      If .Text10 <> MsgText(601) Then
         .adoacc0k0.Fields("a0k19").Value = Val(.Text10)
      Else
         .adoacc0k0.Fields("a0k19").Value = 0 'Null Modify By Sindy 2024/7/5 列印次數不可存null預設值為0
      End If
      .adoacc0k0.Fields("a0k27").Value = Val(strSrvDate(2))
      .adoacc0k0.Fields("a0k28").Value = ServerTime
      .adoacc0k0.Fields("a0k29").Value = strUserNum
        'Add By Cheng 2003/12/04
        '若修改抬頭且已收款, 記錄最後修改抬頭日期
        If .Combo1.Text <> .Combo1.Tag And .m_dbla0k17 + .m_dbla0k18 <> 0 Then
            .adoacc0k0.Fields("a0k31").Value = Val(strSrvDate(2))
            'm_ShowMsg = m_ShowMsg & IIf(m_ShowMsg = "", "", "及") & "已收款" 'Add By Sindy 2013/12/27
        End If
        'End
        
        'Added by Lydia 2023/12/12 Z=確定不印，鎖住「收據暫不列印」，存檔時回存讀出的值。
        If m_strA0k32 = "Z" Then
           .adoacc0k0.Fields("a0k32").Value = m_strA0k32
        Else
        'end 2023/12/12
           'Add By Sindy 2010/5/5
           If .Check1.Value = 1 Then
               .adoacc0k0.Fields("a0k32").Value = "N"
               .adoacc0k0.Fields("a0k19").Value = 0   '2013/11/18 add by sonia
               .Text10 = 0                             '2013/11/18 add by sonia
           Else
               .adoacc0k0.Fields("a0k32").Value = Null
           End If
           '2010/5/5 End
        End If 'Added by Lydia 2023/12/12
      'Add By Sindy 2013/12/27
      'Modify By Sindy 2014/12/29
      'If .txtSales = "" Then
      If .Combo2 = "" Then
      '2014/12/29 END
         .adoacc0k0.Fields("a0k34").Value = Null
      Else
         'Modify By Sindy 2014/12/29
         '.adoacc0k0.Fields("a0k34").Value = .txtSales
         .adoacc0k0.Fields("a0k34").Value = Left(Trim(.Combo2.Text), 5)
         '2014/12/29 END
      End If
      '2013/12/27 END
      'Add by Amy 2015/04/07 +控管日期及類別
      If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
           .adoacc0k0.Fields("a0k38").Value = Null
      Else
          .adoacc0k0.Fields("a0k38").Value = Val(FCDate(MaskEdBox1.Text))
      End If
      If CboClass = MsgText(601) Then
          .adoacc0k0.Fields("a0k39").Value = Null
      Else
          .adoacc0k0.Fields("a0k39").Value = CboClass
      End If
      
'      'Add By Sindy 2017/3/17
'      If .txtPrintNo = "" Then
'         .adoacc0k0.Fields("a0k40").Value = Null
'      Else
'         .adoacc0k0.Fields("a0k40").Value = txtPrintNo.Text
'      End If
'      '2017/3/17 END
      
      .adoacc0k0.UpdateBatch
      .RecordShow
      
      'Add By Sindy 2016/11/4 高所接洽單事後收到才輸介紹案源同仁,導至收款當下沒有更新到介紹獎金可發放日期 ex.幸福緻皂珈 T206253延展(謝秀珠)
      If .Combo2 <> "" And Trim(Label21.Caption) = "" Then
         PUB_UpdateReceiptStatus Text1
      End If
      '2016/11/4 END
      
        'Add By Cheng 2003/10/17
        '更新國內收款資料的扣繳年度
        'Modify By Cheng 2004/02/02
        '更新發票號碼
'        strSQLA = "Update ACC0M0 Set A0M07=" & Val(.Text13.Text) & " Where A0M02='" & .Text1 & "' "
      'Add By Sindy 2013/12/27 +if
      If .Text12 <> "J" Then
        StrSQLa = "Update ACC0M0 Set A0M03='" & .Text15.Text & "', A0M07=" & Val(.Text13.Text) & " Where A0M02='" & .Text1 & "' "
        cnnConnection.Execute StrSQLa
      End If
        'End
        '92.11.20 add by sonia
        '更新國內收款資料的扣繳年度
'        strSQLA = "Update ACC1V0 Set A1V09=" & Val(.Text13.Text) & " Where A1V02='" & .Text1 & "' "
        'Modify By Cheng 2004/02/02
        '更新發票號碼
        '93.2.9 modify by sonia 更新收據公司別
        'strSQLA = "Update ACC1V0 Set A1V17='" & .Text15.Text & "', A1V09=" & Val(.Text13.Text) & " Where A1V02='" & .Text1 & "' "
        'modify by sonia 2021/1/18 更新收據公司別A1V03有L公司不能用VAL
        'StrSQLa = "Update ACC1V0 Set A1V17='" & .Text15.Text & "', A1V09=" & Val(.Text13.Text) & ", A1V03='" & Val(.Text12.Text) & "' Where A1V02='" & .Text1 & "' "
        StrSQLa = "Update ACC1V0 Set A1V17='" & .Text15.Text & "', A1V09=" & Val(.Text13.Text) & ", A1V03='" & .Text12.Text & "' Where A1V02='" & .Text1 & "' "
        '93.2.9 end
        cnnConnection.Execute StrSQLa
        '92.11.20 End
        
        bolUpdData = True 'Add By Sindy 2017/11/2 有異動資料,重新查詢
        
        'Add By Sindy 2012/11/12 收據自動列印時間點
        If .Check1.Value = 1 Then
            StrSQLa = "update caseprogress " & _
                      "set cp151=" & CNULL(IIf(.Check2(0).Value = 1, "1", IIf(.Check2(1).Value = 1, "2", IIf(.Check2(2).Value = 1, "3", "")))) & " " & _
                      "where cp09 in(select a0j01 from acc0j0 where a0j13='" & .Text1 & "')"
            cnnConnection.Execute StrSQLa
        End If
        '2012/11/12 End
        
      'Add By Sindy 2013/12/27
      '若修改收據抬頭,存檔時自動帶出發票資料畫面
      If .Combo1.Text <> .Combo1.Tag Then
         'Add By Sindy 2014/1/10 修改發票資料裡的統一編號
         'Modify By Sindy 2015/8/26 增加顯示訊息提醒
         'Modify By Sindy 2016/8/22 + or cu05||' '||cu88||' '||cu89||' '||cu90='" & cboTitle & "' or cu06='" & cboTitle & "'
         strCU11 = ""
'         strSql = "select cu11" & _
'                  " From customer" & _
'                  " where (upper(cu04)=upper('" & ChgSQL(.Combo1.Text) & "') or upper(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))=upper('" & ChgSQL(.Combo1.Text) & "') or upper(cu06)=upper('" & ChgSQL(.Combo1.Text) & "'))" & _
'                  " and cu15<>'0'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 1 Then
'            strCU11 = "" & RsTemp.Fields("cu11")
''         End If
''         If strCU11 = "" Then
'         Else
'            'Modify By Sindy 2014/3/25 若A4202='04150022'者視為空值
'            'Modify By Sindy 2017/4/18 and A4202<>'04150022'==>and (A4202<>'04150022' or A4202 is null) 改語法不然抓不到資料
'            strSql = "select a4202" & _
'                     " From acc420" & _
'                     " where upper(a4201)=upper('" & ChgSQL(.Combo1.Text) & "') and (A4202<>'04150022' or A4202 is null)"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               strCU11 = "" & RsTemp.Fields("a4202")
'            Else
'               'Modify By Sindy 2015/12/1
'               'MsgBox "此為新的收據抬頭，請聯絡智權同仁提供基本資料以利建檔!!", vbInformation
'               'Modify By Sindy 2016/8/22 + or cu05||' '||cu88||' '||cu89||' '||cu90='" & cboTitle & "' or cu06='" & cboTitle & "'
'               strSql = "select cu11" & _
'                        " From customer" & _
'                        " where (upper(cu04)=upper('" & ChgSQL(.Combo1.Text) & "') or upper(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))=upper('" & ChgSQL(.Combo1.Text) & "') or upper(cu06)=upper('" & ChgSQL(.Combo1.Text) & "'))"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'               If intI = 0 Then
'                  MsgBox "此為新的收據抬頭，請聯絡智權同仁提供基本資料以利建檔!!", vbInformation
'               End If
'               '2015/12/1 END
'            End If
'         End If
'         '2015/8/26 END
         'Add By Sindy 2017/6/19 改呼叫函數 : 檢查收據抬頭是否存在
         'Modified by Sindy 2018/9/18 拿掉chgsql
         strCU11 = PUB_ChkTitleNmExist(.Combo1.Text)
         '2017/6/19 END
         'Add By Sindy 2019/9/18
         If strCU11 = "無統編" Then
            strCU11 = ""
            'MsgBox "此收據抬頭無統編資料，請確認!!", vbInformation
         End If
         '2019/9/18 END
         
         'Modify by Amy 2021/01/28 發票未上傳才更新 a4303(電子發票上線後,已上傳之發票不可再修改)
         If "" & .adoacc0k0.Fields("a4301").Value <> "" And Text12 = "J" And IsNull(.adoacc0k0.Fields("a4319")) Then '已開發票
            StrSQLa = "update acc430" & _
                      " set a4303=" & CNULL(strCU11) & _
                      " where a4301='" & .adoacc0k0.Fields("a4301").Value & "'"
            cnnConnection.Execute StrSQLa
            '2014/1/10 END
            
            '開啟frmacc1127畫面
            strItemNo = .Text1
            strCustNo = .Text2
            strTitle = Me.Name
            Me.Enabled = False
            Screen.MousePointer = vbHourglass
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
            Frmacc1127.Text1.Enabled = False
            'Frmacc1127.cmdSave.Visible = False 'Modify By Sindy 2014/3/31 Mark
            Frmacc1127.Show
            Screen.MousePointer = vbDefault
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
            If m_ShowMsg <> "" Then
               MsgBox m_ShowMsg & "，請自行調整傳票內容!!", vbInformation
            End If
         End If
      End If
      '2013/12/27 END
      
      Call PUB_ChkJCompanyRecv_Mail(.Text1, .Text12) 'Add By Sindy 2014/1/29 若收據開J公司,但案件的特殊出名公司未輸入時,同時發E-MAIL
      
      'Add By Sindy 2016/11/8
      If .Combo1.Text <> .Combo1.Tag Then
         strSql = "select a1v02" & _
                  " From acc1v0" & _
                  " where a1v02='" & .Text1 & "' and a1v06>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            PUB_SendMail strUserNum, strUserNum, "", .Text1 & "已扣繳，收據修改抬頭，請聯絡扣繳應收回事！", _
               "智權人員：" & Text11 & GetPrjSalesNM(Text11) & vbCrLf & _
               "收據號碼：" & .Text1 & vbCrLf & _
               "原收據抬頭：" & .Combo1.Tag & vbCrLf & _
               "更改後收據抬頭：" & .Combo1.Text & vbCrLf & _
               "本張收據已收款並已扣繳；請與客戶或智權人員連絡本收據抬頭是否確定要修改！" & vbCrLf _
               , , , , , , , , , , True
         End If
      End If
      '2016/11/8 END
      
      .Combo1.Tag = .Combo1.Text 'Add By Sindy 2014/1/10
      .Text5.Tag = Text5 'Add by Amy 2021/01/28
Checking:
   'Modify by Amy 2025/07/16 +stMsg
   If Err.Number = 0 And stMsg = "" Then
      Exit Sub
   End If
   If Err.Number <> 0 Then
      If stMsg <> "" Then stMsg = stMsg & vbCrLf
      stMsg = stMsg & Err.Description
   End If
   'MsgBox Err.Description, , MsgText(5)
   MsgBox stMsg, , MsgText(5)
   'end 2025/07/16
   End With
End Sub

'檢查專利案是否已專利商標出名
Private Function ChkPatentNameCompany(pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String) As String
   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer
   ChkPatentNameCompany = ""
   stSQL = "select pa161 from patent where pa01='" & pPA01 & "' and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "' and pa161 is not null" & _
           " union select tm130 from trademark where tm01='" & pPA01 & "' and tm02='" & pPA02 & "' and tm03='" & pPA03 & "' and tm04='" & pPA04 & "' and tm130 is not null" & _
           " union select sp85 from servicepractice where sp01='" & pPA01 & "' and sp02='" & pPA02 & "' and sp03='" & pPA03 & "' and sp04='" & pPA04 & "' and sp85 is not null" & _
           " union select lc48 from lawcase where lc01='" & pPA01 & "' and lc02='" & pPA02 & "' and lc03='" & pPA03 & "' and lc04='" & pPA04 & "' and lc48 is not null"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      ChkPatentNameCompany = Trim("" & adoRst.Fields(0).Value)
   End If
End Function

'Add by Amy 2015/04/07
Private Sub SetCombo(oCombo As Object)
   With oCombo
      .Clear
      .AddItem ""
      'Modify by Amy 2016/08/23 修改顯示順序-瑞婷
      .AddItem "預計收款" 'Add by Amy 2015/08/24
      .AddItem "待銷帳"
      .AddItem "未送件"
      .AddItem "依流程請款"
      .AddItem "會稿中"
      .AddItem "尚未辦理"
      .AddItem "催款中" 'Add by Amy 2015/08/24
      .AddItem "請款中" 'Modify by Amy 2017/08/29 原:待收款
      .AddItem "其他"
      'end 2016/08/23
   End With
End Sub

'Add by Amy 2015/04/17 從acc_cls搬回
Public Sub Frmacc1140_Clear()
    Text1 = "E"
    TextInverse Text1
    MaskEdBox2.Mask = ""
    MaskEdBox2.Mask = DFormat
    Text10 = ""
    Text11 = ""
    LblA0k20_N = "" 'Add By Sindy 2021/5/10
    Text2 = ""
    Text3 = ""
    'Modify by Morgan 2006/12/7
    '.Text6 = ""
      Combo1.Clear
    'end
    Text5 = ""
    'Add byAmy 2021/01/28
    Text5.Tag = ""
    HasA4319 = False
    HasA4321 = False
    'end 2021/01/28
    Text9 = ""
    Text4 = ""
    Text7 = ""
    Text8 = ""
    'Add By Cheng 2003/04/21
    Text12 = ""
    Text12.Tag = "" 'Add by Amy 2020/04/24
    Text13 = ""
    AdodcRefresh
    Text1.SetFocus
    'Add By Sindy 2010/5/5
    Check1.Value = 0
    'Add By Sindy 2012/11/12
    Check2(0).Value = False
    Check2(1).Value = False
    Check2(2).Value = False
    'Add by Amy 2013/06/14
    Label18.Caption = ""
    'Add by Amy 2015/04/17
    MaskEdBox1.Tag = ""
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = ""
    MaskEdBox1.Mask = DFormat
    CboClass = ""
    CboClass.Tag = ""
End Sub

'Add by Amy 2025/07/16 取得收據扣繳年度異動檔之收據流水號
Private Function UpdateAcc290(ByRef stMsg) As String
   Dim RsQ As New ADODB.Recordset, intQ As Integer, strQ As String
   Dim strSeq As String, strCmd As String, intR As Integer
   
On Error GoTo ErrHnd

   UpdateAcc290 = False
   stMsg = ""
   
   'Mark by Amy 2025/07/18 有修改都記錄,才能讓A0k16 一致-秀玲
'   '確認同一天,同一人是否輸相同資料
'   strQ = "Select * From Acc290 Where A2901='" & Text1 & "' And A2903=" & Val(Text13) & _
'                " And A2904='" & strUserNum & "' And A2905=" & strSrvDate(2)
'   intQ = 1
'   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'   If intQ = 1 Then
'      UpdateAcc290 = True '同一天已有資料不寫入
'      Set RsQ = Nothing
'      Exit Function
'   End If
   
   '取目前收據編號最大流水號
   strQ = "Select Nvl(Max(A2902),'00') as SeqNo From Acc290 Where A2901='" & Text1 & "' "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      strSeq = Val(RsQ.Fields("SeqNo"))
   End If
   strSeq = Format(Val(strSeq) + 1, "00")
   strCmd = "Insert into Acc290 (a2901,a2902,a2903,a2904,a2905) " & _
                                 "Values ('" & Text1 & "','" & strSeq & "','" & Val(Text13.Text) & "','" & strUserNum & "'," & strSrvDate(2) & ")"
   cnnConnection.Execute strCmd, intR
   If intR = 0 Then
      stMsg = "資料未寫入[收據扣繳年度異動檔]" & vbCrLf & _
                     "請洽電腦中心！"
   Else
      UpdateAcc290 = True
   End If
   
   Set RsQ = Nothing
   Exit Function
   
ErrHnd:
   stMsg = Err.Description
End Function

'避免存檔完成後不會更新Tag(ex:Text13.Tag),再修改可能不會更新資料
Public Sub ReadData()
   adoacc0k0.Find "a0k01 = '" & Text1 & "'", 0, adSearchForward, 1
   If adoacc0k0.EOF = False Then
      If PUB_GetTaxNo(Combo1, 1) = "Y" Then
         Text5 = "1"
         Text5.Locked = True
      Else
         Text5.Locked = False
      End If
         
      FormShow
      RD06Show '增加顯示預訂收款日
      AdodcRefresh
      RecordShow
   End If
   
End Sub
'end 2025/07/16
