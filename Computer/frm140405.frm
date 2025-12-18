VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140405 
   BorderStyle     =   1  '單線固定
   Caption         =   "互惠代理人資料維護"
   ClientHeight    =   5532
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   9156
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5532
   ScaleWidth      =   9156
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製前半年資料"
      Height          =   405
      Left            =   7650
      TabIndex        =   24
      Top             =   120
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6840
      Top             =   4740
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
   Begin VB.Frame fraContact 
      Height          =   3075
      Left            =   72
      TabIndex        =   13
      Top             =   1032
      Width           =   9015
      Begin VB.TextBox textCountry 
         BorderStyle     =   0  '沒有框線
         Height          =   240
         Left            =   900
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   840
         Width           =   2652
      End
      Begin MSForms.TextBox txtFC 
         Height          =   300
         Index           =   18
         Left            =   7296
         TabIndex        =   2
         Top             =   840
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "656;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "關聯編號的案件往來須做電郵通知：　　　(Y:是)"
         Height          =   180
         Index           =   6
         Left            =   4368
         TabIndex        =   31
         Top             =   870
         Width           =   3888
      End
      Begin MSForms.Label lblSName 
         Height          =   225
         Left            =   4770
         TabIndex        =   30
         Top             =   2145
         Width           =   1185
         Caption         =   "lblSName"
         Size            =   "2090;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFC 
         Height          =   300
         Index           =   17
         Left            =   4050
         TabIndex        =   9
         Top             =   2100
         Width           =   630
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1111;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "提出人員："
         Height          =   180
         Index           =   5
         Left            =   3120
         TabIndex        =   29
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "提出年月："
         Height          =   180
         Index           =   4
         Left            =   3120
         TabIndex        =   28
         Top             =   1830
         Width           =   900
      End
      Begin MSForms.TextBox txtFC 
         Height          =   300
         Index           =   16
         Left            =   4050
         TabIndex        =   8
         Top             =   1770
         Width           =   4830
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "8520;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFC 
         Height          =   300
         Index           =   6
         Left            =   1575
         TabIndex        =   6
         Top             =   1784
         Width           =   600
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1058;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFC 
         Height          =   300
         Index           =   3
         Left            =   900
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1155
         Width           =   600
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1058;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFC 
         Height          =   300
         Index           =   5
         Left            =   3060
         TabIndex        =   5
         Top             =   1462
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFC 
         Height          =   300
         Index           =   4
         Left            =   900
         TabIndex        =   4
         Top             =   1462
         Width           =   600
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1058;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFC 
         Height          =   300
         Index           =   7
         Left            =   900
         TabIndex        =   7
         Top             =   2106
         Width           =   600
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1058;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboContact 
         Height          =   330
         Left            =   1530
         TabIndex        =   27
         Top             =   1140
         Width           =   7395
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "13044;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFC 
         Height          =   555
         Index           =   8
         Left            =   900
         TabIndex        =   10
         Top             =   2430
         Width           =   8010
         VariousPropertyBits=   -1466941413
         MaxLength       =   400
         ScrollBars      =   2
         Size            =   "14129;979"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textName 
         Height          =   630
         Left            =   900
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   7935
         VariousPropertyBits=   -2147467233
         BackColor       =   16777215
         Size            =   "13996;1111"
         Caption         =   "LblFM2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國籍："
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   22
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "給案系統類別："
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   1830
         Width           =   1260
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "期間：           ( 1:上半年 2:下半年 )"
         Height          =   180
         Index           =   3
         Left            =   2520
         TabIndex        =   18
         Top             =   1522
         Width           =   2640
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "年度：                    (民國年)"
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   17
         Top             =   1522
         Width           =   2100
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "給案量："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   2130
         Width           =   720
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人："
         Height          =   180
         Index           =   7
         Left            =   135
         TabIndex        =   15
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   14
         Left            =   135
         TabIndex        =   14
         Top             =   2430
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8550
      Top             =   600
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140405.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9156
      _ExtentX        =   16150
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm140405.frx":20F4
      Height          =   1305
      Left            =   90
      TabIndex        =   12
      Top             =   4170
      Width           =   8985
      _ExtentX        =   15854
      _ExtentY        =   2307
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
      ColumnCount     =   8
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
         DataField       =   "X2"
         Caption         =   "名稱"
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
         DataField       =   "FC04"
         Caption         =   "年度"
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
         DataField       =   "X3"
         Caption         =   "期間"
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
         DataField       =   "FC06"
         Caption         =   "系統別"
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
         DataField       =   "FC07"
         Caption         =   "給案量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "FC08"
         Caption         =   "備註"
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
         DataField       =   "FC16"
         Caption         =   "提出年月"
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
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2352.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   468.283
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   455.811
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   648
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   612.284
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1008
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox txtFC 
      Height          =   300
      Index           =   2
      Left            =   2295
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Width           =   285
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "503;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFC 
      Height          =   300
      Index           =   1
      Left            =   1215
      TabIndex        =   0
      Top             =   660
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1931;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   2610
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   660
      Width           =   5865
      VariousPropertyBits=   -2147467233
      BackColor       =   16777215
      Size            =   "10345;529"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   20
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "frm140405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/11 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、textCUID、textName、cboContact、txtFC08
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'Create by Morgan 2008/1/30
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

Dim m_FieldList() As FIELDITEM

Dim oText As Control
Dim TF_FC As Integer
Dim m_strFC01 As String
Dim m_strKey As String
Dim idx As Integer
Dim m_bReadGrid As Boolean '是否要讀取被點選聯絡人資料

Public adofagent As New ADODB.Recordset    '20140224ADD By eric

Private Sub cboContact_Click()
Dim strTmp As String 'Added by Lydia 2022/01/11
   idx = cboContact.ListIndex
   If idx >= 0 Then
      'Modified by Lydia 2022/01/11
      'If cboContact.ITEMDATA(idx) = 0 Then
      strTmp = Val(PUB_GetItemData(cboContact.Tag, idx))
      If strTmp = "0" Then
         txtFC(3) = ""
      Else
         'Modified by Lydia 2022/01/11
         'txtFC(3) = Format(cboContact.ITEMDATA(idx), "00")
         'Modified by Lydia 2023/01/19 +Format
         txtFC(3) = Format(strTmp, "00")
      End If
   End If
End Sub

'Modified by Lydia 2022/01/11 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub cboContact_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmdCopy_Click()
   Dim stTitle As String
   Dim iYear As Integer, iPeriod As Integer, iMonth As Integer
   Dim stYear As String, stPeriod As String
   
   stTitle = "複製前半年資料"
   iYear = Val(Left(strSrvDate(1), 4)) - 1911
   iMonth = Val(Mid(strSrvDate(1), 5, 2))
   If iMonth >= 2 And iMonth <= 8 Then
      iPeriod = 2
   Else
      iPeriod = 1
      If iMonth >= 8 Then
         iYear = iYear + 1
      End If
   End If
   Do
      stYear = InputBox("請輸入欲複製到的年度？", stTitle, iYear)
      If stYear = "" Then Exit Sub
      If Val(stYear) >= iYear Then
         Exit Do
      Else
         If MsgBox("年度輸入錯誤！", vbRetryCancel, stTitle) = vbCancel Then
            Exit Sub
         End If
      End If
   Loop
   
   Do
      stPeriod = InputBox("請輸入欲複製到的期間？( 1:上半年 2:下半年 )", stTitle, iPeriod)
      If stPeriod = "" Then Exit Sub
      If stPeriod = "1" Or stPeriod = "2" Then
         Exit Do
      Else
         If MsgBox("期間輸入錯誤！", vbRetryCancel, stTitle) = vbCancel Then
            Exit Sub
         End If
      End If
   Loop
   
   'Added by Lydia 2020/09/25 複製前半年資料，詢問完 欲複製到的年度和期間後，若新的期間已有該系統類別資料，則顯示訊息：複製到的109年下半年已有資料，不可再複製。
                                            '=> 只提供複製一次的機會,後面新增之記錄人工自行補上
   If Left(Pub_StrUserSt03, 2) = "F1" Then
      strExc(0) = "CFT"
   ElseIf Left(Pub_StrUserSt03, 2) = "F2" Then
      strExc(0) = "CFP"
   Else
      strExc(0) = ""
      MsgBox "無系統別不可複製!!!"
      Exit Sub
   End If
   If strExc(0) <> "" Then
       strSql = "select count(*) cnt from fagentconfig where fc06='" & strExc(0) & "' and fc04=" & stYear & " and fc05=" & stPeriod & " "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
       If intI = 1 Then
           If Val("" & RsTemp.Fields(0)) > 0 Then
               strExc(1) = "複製到的" & stYear & "年" & IIf(stPeriod = "1", "上", "下") & "半年已有資料，不可再複製。"
               MsgBox strExc(1), vbInformation, "檢核資料"
               Exit Sub
           End If
       End If
   End If
   'end 2020/09/25
   If doCopy(stYear, stPeriod, intI) = True Then
      If intI > 0 Then
     '    MsgBox "資料已複製到 " & stYear & " 年 " & IIf(stPeriod = "1", "上半年", "下半年")
         ShowRecord 3
      Else
         MsgBox "前半年無資料可供複製！"
      End If
   End If
End Sub

'Add by Morgan 2008/7/22 複製前半年資料
Private Function doCopy(p_stToYear As String, p_stToPeriod As String, Optional p_iRec As Integer) As Boolean
   Dim iFrYear As Integer, iFrPeriod As Integer
   Dim strFC06 As String
   
   'Modified by Lydia 2022/05/12 adofgt(1 To 9)=> adofgt(1 To 11)
   'Modified by Lydia 2024/02/06 adofgt(1 To 11)=> adofgt(1 To 12)
   Dim adofgt(1 To 12) As String                    '20140225ADD By eric
   Dim strFG As String                             '20140225ADD By eric
      
   If p_stToPeriod = "2" Then
      iFrPeriod = 1
      iFrYear = p_stToYear
   Else
      iFrPeriod = 2
      iFrYear = p_stToYear - 1
   End If
   'Add By Sindy 2013/5/22
   If Left(Pub_StrUserSt03, 2) = "F1" Then
      strFC06 = "CFT"
   ElseIf Left(Pub_StrUserSt03, 2) = "F2" Then
      strFC06 = "CFP"
   Else
      MsgBox "無系統別不可複製!!!"
      Exit Function
   End If
   '2013/5/22 End

'20140225START Modify By eric
'   strSql = "insert into fagentconfig(fc01,fc02,fc03,fc04,fc05,fc06,fc07,fc08,fc09,fc10,fc11,fc15)" & _
'      " select fc01,fc02,fc03," & p_stToYear & "," & p_stToPeriod & ",fc06,fc07,fc08" & _
'      ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24mi')" & _
'      ",fc01||fc02||fc03||'" & p_stToYear & "'||'" & p_stToPeriod & "'||fc06" & _
'      " from fagentconfig a where fc04=" & iFrYear & " and fc05=" & iFrPeriod & _
'      " and not exists(select * from fagentconfig b where b.fc15=a.fc01||a.fc02||a.fc03||'" & p_stToYear & "'||'" & p_stToPeriod & "'||a.fc06)" & _
'      " and fc06='" & strFC06 & "'"
'   cnnConnection.Execute strSql, p_iRec
'   doCopy = True

'    strFC06 = "CFP"         'FOR TEST
    
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   If adofagent.State = adStateOpen Then
      adofagent.Close
   End If
   
   adofagent.CursorLocation = adUseClient
   'Modified by Lydia 2022/05/12 +a.fc16,a.fc17
   'Modified by Lydia 2024/02/06 +FC18
   adofagent.Open "select a.fc01,a.fc02,a.fc03,a.fc04,a.fc05,a.fc06,a.fc07,a.fc08,a.fc09,a.fc10,a.fc11,a.fc15,d.fa69,a.fc16,a.fc17,a.fc18 " & _
                 " from fagentconfig a,fagent d " & _
                 " WHERE a.fc04='" & iFrYear & "' and a.fc05='" & iFrPeriod & "' " & _
                 " and a.fc01=d.fa01 and a.fc02=d.fa02" & _
                 " and not exists(select * from fagentconfig b where b.fc15=a.fc01||a.fc02||a.fc03||'" & p_stToYear & "'||'" & p_stToPeriod & "'||a.fc06) " & _
                 " and a.fc06='" & strFC06 & "' ", adoTaie, adOpenStatic, adLockReadOnly
                 
   If adofagent.RecordCount <> 0 Then
      adofagent.MoveFirst
      
      Do While Not adofagent.EOF
         p_iRec = 0
         
         adofgt(1) = "'" & adofagent.Fields("fc01") & "'"
         adofgt(2) = "'" & adofagent.Fields("fc02") & "'"
         adofgt(3) = "'" & adofagent.Fields("fc03") & "'"
         adofgt(4) = "'" & adofagent.Fields("fc04") & "'"
         adofgt(5) = "'" & adofagent.Fields("fc05") & "'"
         adofgt(6) = "'" & adofagent.Fields("fc06") & "'"
         adofgt(7) = "'" & adofagent.Fields("fc07") & "'"
         'Modified by Lydia 2024/01/29 備註內容不要複製--Sharon提出, 外商陳金蓮同意
         'adofgt(8) = "'" & adofagent.Fields("fc08") & "'"
         adofgt(8) = "NULL"
         adofgt(9) = "'" & adofagent.Fields("fa69") & "'"
         'Added by Lydia 2022/05/12
         adofgt(10) = "'" & adofagent.Fields("fc16") & "'"
         adofgt(11) = "'" & adofagent.Fields("fc17") & "'"
         'end 2022/05/12
         'Added by Lydia 2024/02/06
         adofgt(12) = "'" & adofagent.Fields("fc18") & "'"
         
         If adofgt(9) = "''" Then
            'Modified by Lydia 2022/05/12 +fc16,fc17
            'Modified by Lydia 2024/02/06 +FC18
            strSql = "insert into fagentconfig(fc01,fc02,fc03,fc04,fc05,fc06,fc07,fc08,fc09,fc10,fc11,fc15,fc16,fc17,fc18)" & _
                    " values(" & adofgt(1) & "," & adofgt(2) & "," & adofgt(3) & "," & p_stToYear & "," & p_stToPeriod & "," & adofgt(6) & "," & _
                             adofgt(7) & "," & adofgt(8) & ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24mi')," & _
                             "'" & adofagent.Fields("fc01") & "'||'" & adofagent.Fields("fc02") & "'||'" & adofagent.Fields("fc03") & "'||'" & p_stToYear & "'||'" & p_stToPeriod & "'||'" & adofagent.Fields("fc06") & "'," & _
                             adofgt(10) & " ," & adofgt(11) & "," & adofgt(12) & " )"
            cnnConnection.Execute strSql, p_iRec
            
            If p_iRec = 0 Then
               MsgBox "資料異常或已存在，請確認！", vbInformation, "複製資料"
               adofagent.Close
               GoTo ErrHnd
            End If
         Else
            strFG = strFG & "," & adofagent.Fields("fc01") & adofagent.Fields("fc02")
         End If
         adofagent.MoveNext
      
      Loop
    
   Else
      adofagent.Close
      cnnConnection.RollbackTrans
      GoTo ErrHnd
   End If
   
   cnnConnection.CommitTrans
   If p_iRec > 0 Then
      MsgBox "資料已複製到 " & p_stToYear & " 年 " & IIf(p_stToPeriod = "1", "上半年", "下半年") & IIf(strFG = "", "", "；代理人" & strFG & " 不再使用，資料不複製！"), vbInformation, "複製資料"
   End If
   doCopy = True
   
ErrHnd:
      If Err.Number <> 0 Then
         cnnConnection.RollbackTrans
         MsgBox Err.Description, vbCritical
         Err.Clear
      Else
         Exit Function
      End If
      Screen.MousePointer = vbDefault
'20140225END
   
End Function

Private Sub DataGrid1_Click()
   '點選同一列可能不會觸發RowColChange
   If DataGrid1.col = -1 Then
      ReadGrid
   End If
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If m_bReadGrid = True Then
      ReadGrid
   End If
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub Form_Initialize()
   strExc(0) = "select * from FagentConfig where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_FC = RsTemp.Fields.Count
   ReDim m_FieldList(TF_FC) As FIELDITEM
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         KeyCode = 0
         If TBar1.Buttons(1).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction vbKeyF2
            End If
         End If
      
      Case vbKeyF3 ' 修改
         KeyCode = 0
         If TBar1.Buttons(2).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction vbKeyF3
            End If
         End If
         
      Case vbKeyF5 ' 刪除
         KeyCode = 0
         If TBar1.Buttons(3).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction vbKeyF5
            End If
         End If
      
      Case vbKeyF4 ' 查詢
         KeyCode = 0
         If TBar1.Buttons(4).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction vbKeyF4
            End If
         End If
      
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd
         If TBar1.Buttons(6).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction KeyCode
            End If
         End If
         KeyCode = 0
         
      Case vbKeyF9, vbKeyF10
         If TBar1.Buttons(11).Enabled = True Then
            If m_EditMode <> 0 Then
               OnAction KeyCode
            End If
         End If
         KeyCode = 0
         
      Case vbKeyEscape
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            KeyCode = 0
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction vbKeyEscape
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 KeyCode 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
         
   End Select
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   textName.BackColor = &H8000000F
   textCountry.BackColor = &H8000000F
   InitialField
   m_EditMode = 0
   ShowRecord -2
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140405 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         
      Case vbKeyF3 ' 修改
         If ChkModifyLimit = False Then Exit Sub 'Add By Sindy 2013/5/22
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         setContact
         
      Case vbKeyF5 ' 刪除
         If ChkModifyLimit = False Then Exit Sub 'Add By Sindy 2013/5/22
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         txtFC(1).Locked = False
         ClearField
         UpdateToolbarState
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
         
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  m_EditMode = 0
                  If m_strFC01 <> "" Then
                     txtFC(1) = m_strFC01
                     ShowRecord
                  Else
                     ClearField
                  End If
                  UpdateToolbarState
               End If
               
            Case Else
               m_EditMode = 0
               If m_strFC01 <> "" Then
                  txtFC(1) = m_strFC01
                  ShowRecord
               Else
                  ClearField
               End If
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

'Add By Sindy 2013/5/22
Private Function ChkModifyLimit() As Boolean
   ChkModifyLimit = False
   If Left(Pub_StrUserSt03, 2) = "F1" And txtFC(6) = "CFT" Then
      ChkModifyLimit = True
   ElseIf Left(Pub_StrUserSt03, 2) = "F2" And txtFC(6) = "CFP" Then
      ChkModifyLimit = True
   End If
   If Pub_StrUserSt03 = "M51" Then
      ChkModifyLimit = True
   Else
      If ChkModifyLimit = False Then
         MsgBox "無權限!!!"
      End If
   End If
End Function

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtFC
      oText.Locked = bLocked
   Next
   txtFC(2).Locked = True
   txtFC(3).Locked = True
   'Add by Morgan 2008/6/23
   '修改狀態不可改代理人編號
   If m_EditMode = 2 Then
      txtFC(1).Locked = True
   End If
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
            CmdCopy.Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
            CmdCopy.Enabled = False
         End If
         If m_bUpdate And txtFC(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtFC(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtFC(1) <> "" Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
         CmdCopy.Enabled = False
   End Select
   
End Sub

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0, Optional ByVal p_bGridOnly As Boolean) As Boolean
   Dim stFC01 As String

   stFC01 = Left(txtFC(1) & "000", 8)
   
   Select Case p_iWay
      Case 0, 3 '當筆
         strExc(0) = "SELECT F.*,FA04 CN,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
            ",PCC05 CCN,PCC03 CEN,PCC04 CJN,FA31,fa10,na03" & _
            " FROM FAgentConfig F,Fagent,PotCustCont,nation WHERE FC01='" & stFC01 & "'" & _
            " AND FA01(+)=FC01 AND FA02(+)=FC02 AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
            " and na01(+)=fa10"
         
      Case -2 '首筆
         strExc(0) = "SELECT F.*,FA04 CN,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
            ",PCC05 CCN,PCC03 CEN,PCC04 CJN,FA31,fa10,na03" & _
            " FROM FAgentConfig F,Fagent,PotCustCont,nation WHERE FC01=(SELECT MIN(F1.FC01) FROM FAgentConfig F1)" & _
            " AND FA01(+)=FC01 AND FA02(+)=FC02 AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
            " and na01(+)=fa10"

      Case -1 '前筆
         strExc(0) = "SELECT F.*,FA04 CN,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
            ",PCC05 CCN,PCC03 CEN,PCC04 CJN,FA31,fa10,na03" & _
            " FROM FAgentConfig F,Fagent,PotCustCont,nation WHERE FC01=(SELECT MAX(F1.FC01) FROM FAgentConfig F1 WHERE F1.FC01<'" & stFC01 & "')" & _
            " AND FA01(+)=FC01 AND FA02(+)=FC02 AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
            " and na01(+)=fa10"

      Case 1 '後筆
         strExc(0) = "SELECT F.*,FA04 CN,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
            ",PCC05 CCN,PCC03 CEN,PCC04 CJN,FA31,fa10,na03" & _
            " FROM FAgentConfig F,Fagent,PotCustCont,nation WHERE FC01=(SELECT MIN(F1.FC01) FROM FAgentConfig F1 WHERE F1.FC01>'" & stFC01 & "')" & _
            " AND FA01(+)=FC01 AND FA02(+)=FC02 AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
            " and na01(+)=fa10"

      Case 2 '末筆
         strExc(0) = "SELECT F.*,FA04 CN,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
            ",PCC05 CCN,PCC03 CEN,PCC04 CJN,FA31,fa10,na03" & _
            " FROM FAgentConfig F,Fagent,PotCustCont,nation WHERE FC01=(SELECT MAX(F1.FC01) FROM FAgentConfig F1)" & _
            " AND FA01(+)=FC01 AND FA02(+)=FC02 AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
            " and na01(+)=fa10"
   End Select
   
   strExc(0) = "SELECT X.*,FC01||FC02||DECODE(FC03,NULL,'','-'||FC03) X1" & _
      ",DECODE(FC03,NULL,NVL(EN,NVL(JN,CN)),NVL(CEN,NVL(CJN,CCN))) X2" & _
      ",DECODE(FC05,'1','上半','下半') X3" & _
      " FROM (" & strExc(0) & ") X "
   '2010/12/2 modify by sonia 因100年改排序條件
   'strExc(0) = strExc(0) & "ORDER BY FC15 DESC"
   'Modified by Lydia 2024/07/23 改成依年度排; ex.Y54091有代理人和聯絡人不同設定
   'strExc(0) = strExc(0) & "ORDER BY fc01,fc02,fc03,fc04 desc,fc05,fc06"
   strExc(0) = strExc(0) & "ORDER BY fc04 desc,fc05 desc,fc01,fc02,fc03,fc06"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If p_bGridOnly = True Then
      Set Adodc1.Recordset = RsTemp
   Else
      If intI = 1 Then
         Set Adodc1.Recordset = RsTemp
         ClearField
         If p_iWay = 3 And m_strKey <> "" Then
            RsTemp.Find "FC15='" & m_strKey & "'"
            If RsTemp.EOF Then
               RsTemp.MoveFirst
            End If
         End If
         UpdateCtrlData RsTemp
         ShowRecord = True
      Else
         If p_iWay = -1 Then
            MsgBox "已經是第一筆！", vbInformation
         ElseIf p_iWay = 1 Then
            MsgBox "已經是最後筆！", vbInformation
         Else
            Set Adodc1.Recordset = RsTemp
            MsgBox "無資料！", vbInformation
            ClearField
         End If
      End If
      
      If m_EditMode = 0 Then
         SetCtrlReadOnly True
      End If
      
      If Me.Visible = True Then
         txtFC(1).SetFocus
         txtFC_GotFocus 1
      End If
   End If
   
End Function

Private Sub ClearField()
   
   For Each oText In txtFC
      oText.Text = Empty
   Next
   
   cboContact.Text = "" 'Added by Lydia 2023/01/19
   cboContact.Clear
   For intI = 1 To TF_FC
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = Empty
   textName = Empty
   textCountry = Empty
   txtFC(1).Tag = Empty
   txtFC(2).Tag = Empty
   'Add by Morgan 2008/4/7 新增時要預設
   If m_EditMode = 1 Then
      txtFC(4) = Left(strSrvDate(1), 4) - 1911
      If Val(Mid(strSrvDate(1), 5, 2)) > 6 Then
         txtFC(5) = "2"
      Else
         txtFC(5) = "1"
      End If
      'Modify By Sindy 2013/5/22
      'txtFC(6) = "CFP"
      If Left(Pub_StrUserSt03, 2) = "F1" Then
         txtFC(6) = "CFT"
      ElseIf Left(Pub_StrUserSt03, 2) = "F2" Then
         txtFC(6) = "CFP"
      End If
      '2013/5/22 End
   End If
   
   lblSname.Caption = "" 'Added by Lydia 2022/05/12
   txtFC(17).Tag = "" 'Added by Lydia 2023/07/24
End Sub
Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord 3
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord 3
            End If
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            If Adodc1.Recordset.RecordCount > 1 Then
               ShowRecord
            Else
               ShowRecord 2
            End If
         End If
      
       Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtFC(1).SetFocus
               txtFC_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Sub txtFC_GotFocus(Index As Integer)
   'Remove by Lydia 2022/01/11
   'If Index = 8 Then
   '   OpenIme
   'Else
   '   CloseIme
   'End If
   'end 2022/01/11
   TextInverse txtFC(Index)
End Sub

Private Function TxtValidate() As Boolean
   
   Dim Cancel As Boolean, ii As Integer, jj As Integer

   For Each oText In txtFC
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         Cancel = False
         txtFC_Validate oText.Index, Cancel
         If Cancel = True Then
            oText.SetFocus
            txtFC_GotFocus oText.Index
            Exit Function
         End If
      End If
   Next
   'Added by Lydia 2022/01/11 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If txtFC(8) <> "" Then
        If PUB_ChkUniText(Me, , True, "TextBox") = False Then
            Exit Function
        End If
   End If
   'end 2022/01/11
   
   '20140224START ADD By ERIC
   '新增
   If m_EditMode = 1 Then
      If GetSPName(txtFC(1), txtFC(2)) = False Then
         ShowMsg "該代理人編號有特殊狀態設定 !"
         txtFC(1).SetFocus
         txtFC_GotFocus 1
         Exit Function
      End If
   End If
   '20140224END
   
   '查詢
   If m_EditMode = 4 Then
      If txtFC(1) = "" Then
         ShowMsg "請輸入欲查詢之代理人編號 !"
         txtFC(1).SetFocus
         txtFC_GotFocus 1
         Exit Function
      End If
   '維護
   Else
      If Val(txtFC(4).Text) < "96" Then
         ShowMsg "年度輸入錯誤 !"
         txtFC_GotFocus 4
         txtFC(4).SetFocus
         Exit Function
      End If
     
      If txtFC(5).Text = "" Then
         ShowMsg "期間不可為空白 !"
         txtFC(5).SetFocus
         Exit Function
      End If
            
      If txtFC(6).Text = "" Then
         ShowMsg "系統別不可為空白 !"
         txtFC(6).SetFocus
         Exit Function
      Else
         strExc(0) = "select sk01 from systemkind where sk01='" & txtFC(6) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI <> 1 Then
            ShowMsg "系統別輸入錯誤 !"
            txtFC_GotFocus 6
            txtFC(6).SetFocus
            Exit Function
         End If
      End If
      
      If Val(txtFC(7).Text) = 0 Then
         'Modified by Lydia 2022/07/19 給案量開放可以輸入0，但彈訊息提醒
         'ShowMsg "給案量必須大於 0 !"
         If MsgBox("給案量為零，是否繼續存檔？", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
            txtFC(7).Text = "0"
         Else
         'end 2022/07/19
            txtFC_GotFocus 7
            txtFC(7).SetFocus
            Exit Function
         End If  'Added by Lydia 2022/07/19
      End If

   End If
   
   TxtValidate = True
End Function

Private Sub UpdateFieldNewData()
   For Each oText In txtFC
      idx = oText.Index
      m_FieldList(idx).fiNewData = oText.Text
   Next
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
   Dim stSQL As String, stCols As String, stValues As String

   strExc(0) = "SELECT * FROM FAgentConfig WHERE FC15='" & txtFC(1) & txtFC(2) & txtFC(3) & txtFC(4) & txtFC(5) & txtFC(6) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "資料重複，請再確認！"
      Exit Function
   End If
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '畫面有的欄位才更新
   stCols = "FC15": stValues = "'" & txtFC(1) & txtFC(2) & txtFC(3) & txtFC(4) & txtFC(5) & txtFC(6) & "'"
   For Each oText In txtFC
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   stSQL = "INSERT INTO FAgentConfig (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   
   cnnConnection.Execute stSQL, intI

   cnnConnection.CommitTrans
   m_strKey = txtFC(1) & txtFC(2) & txtFC(3) & txtFC(4) & txtFC(5) & txtFC(6)
   AddRecord = True
   
   txtFC(1) = m_FieldList(1).fiNewData
   m_strKey = txtFC(1) & txtFC(2) & txtFC(3) & txtFC(4) & txtFC(5) & txtFC(6)
   
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtFC
            idx = oText.Index
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
            'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
            'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
               m_FieldList(idx).fiType = 0
            'Else
            '   m_FieldList(idx).fiType = 1
            'End If
            'end 2017/06/29
            oText.Text = m_FieldList(idx).fiOldData
         Next
         CUID(1) = "" & .Fields("FC09")
         CUID(2) = "" & .Fields("FC10")
         CUID(3) = "" & .Fields("FC11")
         CUID(4) = "" & .Fields("FC12")
         CUID(5) = "" & .Fields("FC13")
         CUID(6) = "" & .Fields("FC14")
         
         textName = "英: " & .Fields("EN") & _
            vbCrLf & "日: " & .Fields("JN") & _
            vbCrLf & "中: " & .Fields("CN")
         textCountry = .Fields("fa10") & " " & .Fields("na03")
         If Not IsNull(.Fields("FC03")) Then
            cboContact.Text = .Fields("CEN") & " " & .Fields("CJN") & " " & .Fields("CCN")
         End If
         
         m_strKey = .Fields("FC15")
         Call txtFC_Validate(17, False) 'Added by Lydia 2022/05/12
         txtFC(17).Tag = txtFC(17).Text 'Added by Lydia 2023/07/24
      End If
   End With
   UpdateCUID CUID, textCUID
   m_strFC01 = m_FieldList(1).fiOldData
   
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   For Each oText In txtFC
      idx = oText.Index
      m_FieldList(idx).fiName = "FC" & Format(idx, "00")
   Next
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Control)
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
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

'Modified by Lydia 2022/01/11 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtFC_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index = 8 Then
      If KeyAscii = vbKeyReturn Then
         KeyAscii = 0
         Beep
      End If
   Else
      KeyAscii = UpperCase(KeyAscii)
      Select Case Index
         Case 5
            If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
               KeyAscii = 0
               Beep
            End If
         Case 4
            If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
               KeyAscii = 0
               Beep
            End If
         'Added by Lydia 2024/02/06
         Case 18
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         'end 2024/02/06
      End Select
   End If
   
End Sub

Private Sub txtFC_Validate(Index As Integer, Cancel As Boolean)
      
   Dim iLen As Integer
   Select Case Index
      Case 1, 2
         If txtFC(1) <> "" Then
            If Len(txtFC(1)) > 5 Then
               txtFC(1) = Left(txtFC(1) & "00", 8)
               txtFC(2) = Left(txtFC(2) & "0", 1)
               If txtFC(1) & txtFC(2) <> txtFC(1).Tag & txtFC(2).Tag Then
                  If GetFAgentName(txtFC(1), txtFC(2)) = False Then
                     Cancel = True
                     MsgBox "代理人編號輸入錯誤！"
                     txtFC_GotFocus Index
                  Else
                     txtFC(1).Tag = txtFC(1)
                     txtFC(2).Tag = txtFC(2)
                     'Add by Morgan 2008/4/7 新增時要顯示相同代理人資料於下表
                     If m_EditMode = 1 Then
                        '20140224START ADD By eric
                        If GetSPName(txtFC(1), txtFC(2)) = False Then
                           MsgBox "該代理人編號有特殊狀態設定", vbInformation, "新增代理人"
                           Exit Sub
                        End If
                        '20140224END
                        ShowRecord 0, True
                     End If
                  End If
               End If
            Else
               Cancel = True
               MsgBox "代理人編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
               txtFC_GotFocus Index
            End If
         End If
      'Added by Lydia 2022/05/12
      Case 17   '提出人員
         'Modified by Lydia 2023/07/24
         'lblSName.Caption = ""
         'If txtFC(Index).Text <> "" Then
         If txtFC(Index) <> txtFC(Index).Tag Then
            strExc(1) = ""
            strExc(0) = GetStaffName(txtFC(Index), True, , , strExc(1))
            'Modified by Lyida 2023/07/24 新增時必須在職、修改時有改到此欄位時也必須在職。
            'If m_EditMode = 1 And strExc(1) = "2" Then
            '   If MsgBox("該員工已離職，是否繼續？", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
            If (m_EditMode = 1 Or m_EditMode = 2) And ((txtFC(Index) <> "" And strExc(0) = "") Or strExc(1) = "2") Then
               MsgBox "員工編號已離職或不存在！", vbCritical + vbOKOnly
               Cancel = True
               txtFC(Index).SetFocus
               Call txtFC_GotFocus(Index)
               Exit Sub
               'End If
            End If
            lblSname.Caption = strExc(0)
         End If
      'end 2022/05/12
   End Select
   
   If Cancel = False Then
      If txtFC(Index).MaxLength > 0 Then
         If Not CheckLengthIsOK(txtFC(Index), iLen) Then
            Cancel = True
         End If
      End If
   End If
End Sub
'讀取代理人與聯絡人名稱
Private Function GetFAgentName(p_FA01 As String, p_FA02 As String) As Boolean
   Dim idx As Integer
   
   '代理人
   strExc(0) = "SELECT FA04 CN,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN,fa10,na03" & _
      " FROM Fagent,Nation WHERE FA01='" & p_FA01 & "' AND FA02='" & p_FA02 & "' and na01(+)=fa10"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         textName = "英: " & .Fields("EN") & _
            vbCrLf & "日: " & .Fields("JN") & _
            vbCrLf & "中: " & .Fields("CN")
            
         textCountry = .Fields("fa10") & " " & .Fields("na03")
      End With
      If m_EditMode = 1 Or m_EditMode = 2 Then
         setContact
      End If
      GetFAgentName = True
   End If
End Function

'20140224ADD By eric
Private Function GetSPName(p_FA01 As String, p_FA02 As String) As Boolean
  
   strExc(0) = "SELECT FA01 FROM Fagent WHERE FA01='" & p_FA01 & "' AND FA02='" & p_FA02 & "' and FA69 IS NOT NULL"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetSPName = False
      Exit Function
   End If
  
   GetSPName = True
   
End Function


' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '刪除資料
   stSQL = "delete from FAgentConfig where FC15='" & m_strKey & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   m_strFC01 = ""
   m_strKey = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub ReadGrid()
   With Adodc1.Recordset
      If Not (.EOF Or .BOF) Then
         ClearField
         UpdateCtrlData Adodc1.Recordset
         If m_EditMode = 1 Or m_EditMode = 2 Then
            setContact
         End If
      End If
   End With
End Sub

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
   If m_strKey <> txtFC(1) & txtFC(2) & txtFC(3) & txtFC(4) & txtFC(5) & txtFC(6) Then
      strExc(0) = "SELECT * FROM FAgentConfig WHERE FC15='" & txtFC(1) & txtFC(2) & txtFC(3) & txtFC(4) & txtFC(5) & txtFC(6) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "資料重複，請再確認！"
         Exit Function
      End If
   End If
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE FAgentConfig SET "
   stSet = ""
   For Each oText In txtFC
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & ",FC15='" & txtFC(1) & txtFC(2) & txtFC(3) & txtFC(4) & txtFC(5) & txtFC(6) & "' where FC15='" & m_strKey & "'; end; "
      Pub_SeekTbLog stSQL
      
      cnnConnection.Execute stSQL, intI
   End If
   
   cnnConnection.CommitTrans
   m_strKey = txtFC(1) & txtFC(2) & txtFC(3) & txtFC(4) & txtFC(5) & txtFC(6)
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical

End Function

Private Sub setContact()
   cboContact.Clear
   cboContact.Tag = "" 'Added by Lydia 2022/01/11
   '聯絡人
   strExc(0) = "SELECT PCC02,PCC03,PCC04,PCC05" & _
      " FROM PotCustCont WHERE PCC01='" & txtFC(1) & "' order by PCC02 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Do While Not .EOF
            cboContact.AddItem .Fields("PCC03") & " " & .Fields("PCC04") & " " & .Fields("PCC05"), 0
            'Modified by Lydia 2022/01/11 改成Form 2.0沒有ItemData屬性
            'cboContact.ITEMDATA(0) = .Fields("PCC02")
            cboContact.Tag = .Fields("PCC02") & "," & cboContact.Tag
            .MoveNext
         Loop
      End With
   Else
      txtFC(3) = ""
   End If
   cboContact.AddItem "", 0
   'Modified by Lydia 2022/01/11
   'cboContact.ITEMDATA(0) = 0
   cboContact.Tag = " ," & cboContact.Tag
   For idx = 0 To cboContact.ListCount - 1
      'Modified by Lydia 2022/01/11
      'If Val(cboContact.ITEMDATA(idx)) = Val(txtFC(3)) Then
      If Val(PUB_GetItemData(cboContact.Tag, idx)) = Val(txtFC(3)) Then
         cboContact.ListIndex = idx
         Exit For
      End If
   Next
End Sub


