VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_17 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶/代理人聯絡人資料查詢"
   ClientHeight    =   6380
   ClientLeft      =   420
   ClientTop       =   4420
   ClientWidth     =   9160
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6380
   ScaleWidth      =   9160
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束"
      Height          =   400
      Index           =   1
      Left            =   8250
      TabIndex        =   34
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   0
      Left            =   6960
      TabIndex        =   33
      Top             =   60
      Width           =   1230
   End
   Begin VB.Frame fraContact 
      Height          =   3735
      Left            =   225
      TabIndex        =   18
      Top             =   2550
      Width           =   8610
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "上傳相片"
         Height          =   276
         Left            =   1920
         Style           =   1  '圖片外觀
         TabIndex        =   44
         Top             =   150
         Width           =   948
      End
      Begin VB.CommandButton CmdOk1 
         Caption         =   "寄發信函-往來記錄"
         Height          =   400
         Index           =   2
         Left            =   4440
         TabIndex        =   40
         Top             =   1530
         Width           =   1845
      End
      Begin VB.TextBox txtPCC20 
         BackColor       =   &H8000000F&
         Height          =   264
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   810
         Width           =   2055
      End
      Begin VB.ListBox lstTitle 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         ItemData        =   "frm100101_17.frx":0000
         Left            =   1080
         List            =   "frm100101_17.frx":0007
         MultiSelect     =   1  '簡易多重選取
         Sorted          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1530
         Width           =   3180
      End
      Begin VB.ListBox lstDept 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         ItemData        =   "frm100101_17.frx":0015
         Left            =   1080
         List            =   "frm100101_17.frx":001C
         MultiSelect     =   1  '簡易多重選取
         Sorted          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1110
         Width           =   3180
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   555
         Index           =   1
         Left            =   1080
         TabIndex        =   43
         Top             =   2580
         Width           =   1290
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2275;979"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   24
         Left            =   6870
         TabIndex        =   12
         Top             =   2580
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   26
         Left            =   3480
         TabIndex        =   13
         Top             =   3180
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   26
         Size            =   "503;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   25
         Left            =   5640
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1965
         Width           =   2055
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "3625;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   5
         Top             =   795
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "5609;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   3
         Top             =   495
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "5609;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   4
         Left            =   5310
         TabIndex        =   4
         Top             =   495
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "5609;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   10
         Left            =   6870
         TabIndex        =   11
         Top             =   2265
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   8
         Left            =   1080
         TabIndex        =   8
         Top             =   1965
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5609;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   630
         Index           =   13
         Left            =   5490
         TabIndex        =   14
         Top             =   3030
         Width           =   3060
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "5397;1111"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   11
         Left            =   1080
         TabIndex        =   9
         Top             =   2265
         Width           =   1035
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1826;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   9
         Left            =   4365
         TabIndex        =   10
         Top             =   2265
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   2
         Left            =   1245
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   150
         Width           =   600
         VariousPropertyBits=   671105055
         Size            =   "1058;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID1 
         Height          =   300
         Left            =   2976
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   132
         Width           =   5508
         VariousPropertyBits=   -2147467233
         BackColor       =   16777215
         Size            =   "9716;529"
         Caption         =   "LblFM2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "（W:待回覆 Y/N:同意/不同意）"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   39
         Top             =   3480
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄專利雙週報：      （N:不寄)"
         Height          =   180
         Index           =   25
         Left            =   5280
         TabIndex        =   38
         Top             =   2610
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否同意歐盟通用資料保護規範(GDPR)："
         Height          =   180
         Index           =   37
         Left            =   135
         TabIndex        =   37
         Top             =   3180
         Width           =   3270
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "名片臨時編號："
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   14
         Left            =   4350
         TabIndex        =   36
         Top             =   2025
         Width           =   1260
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "名稱( 中 )："
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   32
         Top             =   855
         Width           =   930
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "名稱( 英 )："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   31
         Top             =   555
         Width           =   930
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "名稱( 日 )："
         Height          =   180
         Index           =   2
         Left            =   4365
         TabIndex        =   30
         Top             =   555
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報：      （Y:寄／ N:不寄)"
         Height          =   180
         Index           =   11
         Left            =   5640
         TabIndex        =   29
         Top             =   2325
         Width           =   2865
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "職稱："
         Height          =   180
         Index           =   3
         Left            =   135
         TabIndex        =   28
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "部門："
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   27
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發人員："
         Height          =   180
         Index           =   12
         Left            =   135
         TabIndex        =   26
         Top             =   2610
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   14
         Left            =   4950
         TabIndex        =   25
         Top             =   3060
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄台一雜誌：     （N:不寄)"
         Height          =   180
         Index           =   7
         Left            =   2925
         TabIndex        =   24
         Top             =   2325
         Width           =   2430
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人編號："
         Height          =   180
         Index           =   7
         Left            =   135
         TabIndex        =   23
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "E-MAIL："
         Height          =   180
         Index           =   5
         Left            =   135
         TabIndex        =   22
         Top             =   2025
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發日期：                         ( 西元 )"
         Height          =   180
         Index           =   9
         Left            =   135
         TabIndex        =   21
         Top             =   2325
         Width           =   2595
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "相關聯絡人編號："
         Height          =   180
         Index           =   8
         Left            =   4365
         TabIndex        =   20
         Top             =   855
         Width           =   1440
      End
   End
   Begin VB.TextBox txtPCU 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   555
      Width           =   255
   End
   Begin VB.TextBox txtPCU 
      Height          =   300
      Index           =   1
      Left            =   1005
      MaxLength       =   8
      TabIndex        =   0
      Top             =   555
      Width           =   1092
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7425
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm100101_17.frx":0029
      Height          =   1095
      Left            =   240
      TabIndex        =   16
      Top             =   1200
      Width           =   8625
      _ExtentX        =   15222
      _ExtentY        =   1923
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "X1"
         Caption         =   "編號"
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
         DataField       =   "PCC03"
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
      BeginProperty Column02 
         DataField       =   "PCC04"
         Caption         =   "日文名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PCC05"
         Caption         =   "中文名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PCC06"
         Caption         =   "部門"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PCC07"
         Caption         =   "職稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "PCC08"
         Caption         =   "EMail"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "PCC09"
         Caption         =   "寄台一雜誌"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "PCC10"
         Caption         =   "寄電子報"
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
      BeginProperty Column09 
         DataField       =   "PCC11"
         Caption         =   "開發日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "PCC12"
         Caption         =   "開發人員"
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
         Size            =   315
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2580.095
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1610.079
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1269.921
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1340.221
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1599.874
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   890.079
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            ColumnWidth     =   920.126
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox textName 
      Height          =   615
      Left            =   2400
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   540
      Width           =   6405
      VariousPropertyBits=   -2139078625
      BackColor       =   16777215
      Size            =   "11298;1085"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   165
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "＊：聯絡人已離職"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   10
      Left            =   330
      TabIndex        =   17
      Top             =   2340
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "編號："
      Height          =   210
      Index           =   0
      Left            =   315
      TabIndex        =   15
      Top             =   555
      Width           =   585
   End
End
Attribute VB_Name = "frm100101_17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/07 改成Form2.0 ; textCUID1、txtPCC(index)、lstUsers(index)、DataGrid1改字型=新細明體-ExtB、textName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
'Create by Morgan 2007/12/20
Option Explicit

Public cmdState As Integer
Dim strTmp As String

Dim rsContact As ADODB.Recordset
Dim m_bReadGrid As Boolean
Dim oText As Object
Dim idx As Integer
   
Private Sub DataGrid1_Click()
   '點選同一列可能不會觸發RowColChange
   If DataGrid1.col = -1 Then
      ReadContact
   End If
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If m_bReadGrid = True Then
      ReadContact
   End If
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
   textName.BackColor = &H8000000F
   textCUID1.BackColor = &H8000000F
   'Add by Amy 2023/08/29
   lstDept.Clear
   lstTitle.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_17 = Nothing
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
      'Add By Sindy 2019/10/7
      Case 2 '寄發信函-往來記錄
         Me.Hide
         Set frm880022.m_PrevF = Me
         frm880022.m_strNo = txtPCU(1) & "0"
         frm880022.m_PCC02 = txtPCC(2)
         If frm880022.QueryData = True Then
            frm880022.Show 'vbModal
         End If
      '2019/10/7 END
   End Select
End Sub

Sub StrMenu()
   Dim strKey  As String, strKey1 As String
   Dim adoRst As New ADODB.Recordset
   
   If Mid(Me.Tag, 10, 1) = "-" Then
      strKey = Left(Me.Tag, 8) & "0"
      strKey1 = Mid(Me.Tag, 11)
   Else
      strKey = Me.Tag
   End If
   
   'Add By Sindy 2011/01/03 檢查國內外權限
   If CheckSR12(strKey) = False Then
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
   pub_QL05 = pub_QL05 & IIf(PUB_CheckQL05("編號：" & strKey & "(聯絡人資料)") = "", "", ";編號：" & strKey & "(聯絡人資料)") 'Add By Sindy 2025/8/13
   
   If Left(strKey, 1) = "X" Then
      strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN" & _
         " FROM Customer WHERE CU01 = '" & Left(strKey, 8) & "' AND CU02 = '" & Mid(strKey, 9) & "'"
   Else
      strExc(0) = "SELECT FA01 NO,FA04 CN,rtrim(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
         " FROM FAgent WHERE FA01 = '" & Left(strKey, 8) & "' AND FA02 = '" & Mid(strKey, 9) & "'"
   End If
         
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ShowRecord adoRst
      If strKey1 <> "" Then
         ReadContact strKey1
      End If
   Else
      ShowNoData
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ShowRecord(ByRef p_Rst As ADODB.Recordset)
   Dim rsPCU As ADODB.Recordset
   Dim CUID(1 To 6) As String
   
   ClearField
   SetCtrlReadOnly True
   Set rsPCU = p_Rst.Clone
   With rsPCU
      txtPCU(1) = "" & .Fields("NO")
      textName = "中: " & .Fields("CN") & _
      vbCrLf & "英: " & .Fields("EN") & _
      vbCrLf & "日: " & .Fields("JN")
      OpenContactTable
   End With
End Sub

Private Sub ClearField()
   For Each oText In txtPCU
      oText.Text = Empty
   Next
   ClearField1
End Sub

Private Sub ClearField1()
   For Each oText In txtPCC
      oText.Text = Empty
   Next
   lstDept.Clear
   lstTitle.Clear
   lstUsers(1).Clear
   textCUID1 = ""
   'Added by Lydia 2024/05/10
   Command1.Visible = False
   Command1.Caption = "上傳相片"
   Command1.BackColor = &H8080FF     '紅色
   Command1.Tag = ""
   'end 2024/05/10
End Sub

Private Sub OpenContactTable()
'On Error GoTo Checking
   If txtPCU(1) <> "" Then
      strExc(0) = "select PCC.*,decode(pcc20,null,'　','＊')||pcc02 X1,decode(pcc20,null,'',substr(pcc20,1,8)||'-'||substr(pcc20,9)) X2 from PotCustCont PCC where pcc01='" & txtPCU(1) & "' order by pcc02"
   Else
      strExc(0) = "select PCC.*,decode(pcc20,null,'　','＊')||pcc02 X1,decode(pcc20,null,'',substr(pcc20,1,8)||'-'||substr(pcc20,9)) X2 from PotCustCont PCC where rownum<1"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/10 +FormName 改暫存TB
   'Set rsContact = PUB_CreateRecordset(RsTemp)
   Set rsContact = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set Adodc1.Recordset = rsContact
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   If rsContact.RecordCount > 0 Then
      If pub_QL04 <> "" Then InsertQueryLog (rsContact.RecordCount) 'Add By Sindy 2025/8/13
      ReadContact
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
   End If
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
End Sub

Private Sub ReadContact(Optional stPCC02 As String)
   Dim CUID(1 To 6) As String
   
   ClearField1
   With rsContact
   If Not (.EOF Or .BOF) Then
      If stPCC02 <> "" Then
         .MoveFirst
         .Find "PCC02='" & stPCC02 & "'"
      End If
      For Each oText In txtPCC
         oText = "" & .Fields("PCC" & Format(oText.Index, "00"))
      Next
      CUID(1) = "" & .Fields("PCC14")
      CUID(2) = "" & .Fields("PCC15")
      CUID(3) = "" & .Fields("PCC16")
      CUID(4) = "" & .Fields("PCC17")
      CUID(5) = "" & .Fields("PCC18")
      CUID(6) = "" & .Fields("PCC19")
      txtPCC20 = "" & .Fields("X2")
      '部門
      If Not IsNull(.Fields("PCC06")) Then
         SetList lstDept, .Fields("pcc06")
      End If
      '職稱
      If Not IsNull(.Fields("PCC07")) Then
         SetList lstTitle, .Fields("pcc07")
      End If
      '開發人員
      If Not IsNull(.Fields("pcc12")) Then
         strExc(0) = "select st02 from staff where instr('" & .Fields("pcc12") & "',st01)>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strTmp = RsTemp.GetString(, , , ",")
            SetList lstUsers(1), strTmp
            lstUsers(1).ListIndex = 0 'Added by Lydia 2022/01/07
         End If
      End If
      UpdateCUID CUID, textCUID1
      
      'Added by Lydia 2024/05/10 聯絡人相片
      If Trim(txtPCU(1)) <> "" And Trim(txtPCC(2)) <> "" Then
         Command1.Visible = True
         Call Pub_GetPCCtoIBF_2(Trim(txtPCU(1)), Trim(txtPCC(2)), Command1)
      Else
         Command1.Visible = False
      End If
      'end 2024/05/10
   End If
   End With
End Sub

' 更新 Create 及 Update 的人
'Modified by Lydia 2022/01/07 As TextBox=> Object
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   'Modified by Lydia 2024/05/10 String(10 -> 6
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(6, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

'Modified by Lydia 2022/01/07  As ListBox=> Object
Private Sub SetList(oList As Object, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         If Trim(arrID(intI)) <> "" Then 'Added by Lydia 2022/01/07
             oList.AddItem arrID(intI), 0
         End If 'Added by Lydia 2022/01/07
      Next
   End If
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtPCU
      oText.Locked = bLocked
   Next
   For Each oText In txtPCC
      oText.Locked = bLocked
   Next
End Sub

'Added by Lydia 2024/05/10
Private Sub Command1_Click()
   frmPic001.oCP01 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "1")
   frmPic001.oCP02 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "2")
   frmPic001.oCP03 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "3")
   frmPic001.oCP04 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "4")
   frmPic001.strWorkType = "1"
   frmPic001.Label11 = "聯絡人相片"
   frmPic001.bolQuery = True '只查詢
   frmPic001.StrMenu
   frmPic001.SetSeekCmdok
   frmPic001.Show vbModal
End Sub

