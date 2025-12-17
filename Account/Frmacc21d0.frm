VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21d0 
   AutoRedraw      =   -1  'True
   Caption         =   "匯票輸入"
   ClientHeight    =   6030
   ClientLeft      =   50
   ClientTop       =   280
   ClientWidth     =   8810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   8810
   Begin VB.TextBox Text22 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6780
      MaxLength       =   8
      TabIndex        =   24
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      MaxLength       =   10
      TabIndex        =   23
      Top             =   5520
      Width           =   1572
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "Ｌ"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   870
      TabIndex        =   62
      Top             =   1830
      Width           =   280
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "Ｊ"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   450
      TabIndex        =   61
      Top             =   1830
      Width           =   280
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "２"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   30
      TabIndex        =   10
      Top             =   1830
      Width           =   280
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "１"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      TabIndex        =   59
      Top             =   23
      Width           =   1545
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      MaxLength       =   12
      TabIndex        =   22
      Top             =   5535
      Width           =   1572
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      MaxLength       =   3
      TabIndex        =   20
      Top             =   5190
      Width           =   528
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6810
      MaxLength       =   14
      TabIndex        =   17
      Top             =   4515
      Width           =   1572
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6450
      TabIndex        =   52
      Top             =   3730
      Width           =   1788
   End
   Begin VB.CommandButton Command6 
      Caption         =   "全部還原"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   7950
      TabIndex        =   50
      Top             =   1425
      Width           =   612
   End
   Begin VB.TextBox Text20 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      MaxLength       =   12
      TabIndex        =   48
      Top             =   3730
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   8064
      Picture         =   "Frmacc21d0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   360
      Width           =   350
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      TabIndex        =   47
      Top             =   3730
      Width           =   1380
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      MaxLength       =   1
      TabIndex        =   13
      Top             =   4170
      Width           =   612
   End
   Begin VB.CommandButton Command4 
      Caption         =   "全部選取"
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
      Left            =   75
      TabIndex        =   8
      Top             =   1290
      Width           =   1005
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21d0.frx":0102
      Height          =   810
      Left            =   1485
      TabIndex        =   26
      Top             =   1335
      Width           =   2595
      _ExtentX        =   4568
      _ExtentY        =   1446
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "a1902"
         Caption         =   "單據編號"
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
         DataField       =   "a1917"
         Caption         =   "公司別"
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
            ColumnWidth     =   1239.874
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   709.795
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4230
      Picture         =   "Frmacc21d0.frx":0117
      Style           =   1  '圖片外觀
      TabIndex        =   11
      ToolTipText     =   "取消"
      Top             =   1305
      Width           =   612
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Frmacc21d0.frx":0559
      Left            =   4080
      List            =   "Frmacc21d0.frx":055B
      TabIndex        =   3
      Top             =   345
      Width           =   1572
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      MaxLength       =   13
      TabIndex        =   15
      Top             =   4515
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc21d0.frx":055D
      Height          =   1440
      Left            =   195
      TabIndex        =   28
      Top             =   2235
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   2540
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "a0102"
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
         DataField       =   "a1p21"
         Caption         =   "外幣金額"
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
         DataField       =   "a1p11"
         Caption         =   "銀行帳號"
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
         DataField       =   "a0g02"
         Caption         =   "銀行名稱"
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
         DataField       =   "a1p20"
         Caption         =   "銀存匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "a1p17"
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
      BeginProperty Column08 
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
            Locked          =   -1  'True
            ColumnWidth     =   3280.252
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1239.874
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1759.748
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1759.748
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   1599.874
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   5859.78
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2310
      Picture         =   "Frmacc21d0.frx":0572
      Style           =   1  '圖片外觀
      TabIndex        =   25
      ToolTipText     =   "取消"
      Top             =   3730
      Width           =   300
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3675
      TabIndex        =   40
      Top             =   3730
      Width           =   1332
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4035
      MaxLength       =   6
      TabIndex        =   14
      Top             =   4170
      Width           =   1572
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4035
      MaxLength       =   14
      TabIndex        =   16
      Top             =   4500
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      MaxLength       =   12
      TabIndex        =   18
      Top             =   4845
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4035
      MaxLength       =   10
      TabIndex        =   19
      Top             =   4830
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4230
      Picture         =   "Frmacc21d0.frx":0BDC
      Style           =   1  '圖片外觀
      TabIndex        =   12
      ToolTipText     =   "取消"
      Top             =   1800
      Width           =   612
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   14
      TabIndex        =   6
      Top             =   690
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      MaxLength       =   15
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   30
      Top             =   30
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   0
      Top             =   30
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
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
      Left            =   6840
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "Frmacc21d0.frx":101E
      Height          =   810
      Left            =   5430
      TabIndex        =   27
      Top             =   1335
      Width           =   2295
      _ExtentX        =   4039
      _ExtentY        =   1393
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "a1c03"
         Caption         =   "單據編號"
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
            ColumnWidth     =   1280.126
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   1200
      Top             =   1350
      Visible         =   0   'False
      Width           =   960
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   315
      Left            =   5160
      Top             =   1305
      Visible         =   0   'False
      Width           =   960
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   315
      Left            =   -15
      Top             =   2235
      Visible         =   0   'False
      Width           =   960
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
      Caption         =   "Adodc3"
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
   Begin MSForms.TextBox Text16 
      Height          =   330
      Left            =   5610
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   4155
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
      Left            =   5610
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4815
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
   Begin MSForms.TextBox Text5 
      Height          =   435
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   4305
      VariousPropertyBits=   -1467989989
      Size            =   "7594;767"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   345
      Left            =   4035
      TabIndex        =   21
      Top             =   5175
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
   Begin VB.Label Label25 
      BackStyle       =   0  '透明
      Caption         =   "選公司："
      Height          =   255
      Left            =   300
      TabIndex        =   60
      Top             =   1650
      Width           =   735
   End
   Begin VB.Label Label24 
      BackStyle       =   0  '透明
      Caption         =   "傳票號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   58
      Top             =   30
      Width           =   975
   End
   Begin VB.Label Label23 
      BackStyle       =   0  '透明
      Caption         =   "對沖(業)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5835
      TabIndex        =   57
      Top             =   5565
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3075
      TabIndex        =   56
      Top             =   5565
      Width           =   975
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   55
      Top             =   5565
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   54
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "台幣金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5850
      TabIndex        =   53
      Top             =   4530
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3105
      TabIndex        =   51
      Top             =   5220
      Width           =   855
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   49
      Top             =   3730
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "借1/貸2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   46
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "結匯單據"
      Height          =   855
      Left            =   5190
      TabIndex        =   45
      Top             =   1335
      Width           =   255
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "未結匯單據"
      Height          =   975
      Left            =   1230
      TabIndex        =   44
      Top             =   1290
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   -75
      X2              =   8673
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -60
      X2              =   8688
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "付款方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3090
      TabIndex        =   43
      Top             =   383
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "作業日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   42
      Top             =   180
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "銀存匯率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   41
      Top             =   4530
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2835
      TabIndex        =   39
      Top             =   3730
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3075
      TabIndex        =   38
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "外幣金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3075
      TabIndex        =   37
      Top             =   4530
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   36
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "銀行代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3075
      TabIndex        =   35
      Top             =   4860
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   1800
      Left            =   210
      Top             =   4110
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -45
      Top             =   5010
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1275
      Left            =   240
      Top             =   -30
      Width           =   8295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   34
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "手續費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "匯票號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5850
      TabIndex        =   32
      Top             =   383
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "結匯日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   31
      Top             =   375
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   30
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc21d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text5、Text16、Text10、Combo1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc1b0 As New ADODB.Recordset
Public adoacc190 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoadodc3 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strA1P01 As String     'add by sonia 2014/3/26 存傳票公司別(ACC190的A1917) 'Memo by Lydia 2020/10/27 CreditGen收據公司別=>轉傳票公司別
'ADD BY SONIA 2014/6/18
Dim strAutoGen As String          '借方資料之傳票公司別
Dim strCreditGen As String        '借方資料之傳票公司別
'END 2014/6/18
Dim strSerialNo As String
Dim strCurrency As String
Dim strNo As String
Dim douTAmount As Double
Dim douLAmount As Double
Dim strAccNo As String
Dim strYes As String
Dim strDYes As String
Dim strDocuNo As String
Dim strOriDocNo As String
Dim strA1812 As String            'add by sonia 2014/8/7 存是否獨立水單
Dim strA1917 As String            'add by sonia 2017/6/16 存收據公司別 'Memo by Lydia 2020/10/27 AutoGen收據公司別=>轉傳票公司別別

Private Sub Combo1_GotFocus()
   TextInverse Combo1  'Added by Lydia 2021/12/14 Form 2.0的ComboBox的GotFocus不會全選反白
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   Select Case Mid(Combo2, 1, 1)
      Case "1"
      Case Else
         If strSaveConfirm = MsgText(3) Then
            strAccNo = AccAutoNo("A", 5)
            strYes = AccSaveAutoNo(Mid(strAccNo, 1, 1), Mid(strAccNo, 5, 5))
            Text3 = strAccNo
         End If
   End Select
End Sub

Private Sub Command1_Click()
   If Adodc3.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc3.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc3.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text6.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   Adodc3Delete
   SumShow
   Adodc3Clear
   DataGrid2.Refresh
End Sub

'Add by Morgan 2007/8/23 p_iAct:1=加,-1=減
Private Sub UpdateA150(p_A1501 As String, Optional p_iAct As Integer = 1)

   'Modify By Sindy 2011/01/06 加入p_A1501的第一碼判斷
   '第一碼為 U 者才更新 ACC150
   '第一碼為 V 者要更新 ACC160, 更新A1607=畫面上的結匯日期
   If Left(p_A1501, 1) = "V" Then
      If p_iAct = 1 Then
         strSql = "update acc160 set a1607 = " & Val(FCDate(MaskEdBox1.Text)) & " where a1601 = '" & p_A1501 & "'"
      Else
         strSql = "update acc160 set a1607 = null where a1601 = '" & p_A1501 & "'"
      End If
   '2011/01/06 End
   Else
      If p_iAct = 1 Then
         strSql = "update acc150 set a1520 = a1506 where a1501 = '" & p_A1501 & "'"
      Else
         strSql = "update acc150 set a1520 = 0 where a1501 = '" & p_A1501 & "'"
      End If
   End If
   adoTaie.Execute strSql, intI
End Sub

Private Sub Command2_Click()
   If Adodc3.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc3.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc3.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text1.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   Screen.MousePointer = vbHourglass
   strOriDocNo = strDocuNo
   Adodc2Delete
   'ADD BY SONIA 2014/6/18
   adoTaie.Execute "delete from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and substr(a1p23,1,9)='" & Adodc2.Recordset.Fields("a1c03").Value & "'"
   Adodc3.Recordset.Requery
   'END 2014/6/18
   Adodc1.Recordset.Requery
   Adodc2.Recordset.Requery
   'AutoGen  'CANCEL BY SONIA 2014/6/18
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
   If Adodc3.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc3.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc3.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text1.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   Screen.MousePointer = vbHourglass
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Adodc2Save
   Adodc1.Recordset.Requery
   Adodc2.Recordset.Requery
   If Text1 <> MsgText(601) And Text3 <> MsgText(601) Then
      'Modified by Lydia 2022/03/15
      'If IsNull(Adodc2.Recordset.Fields("a1505").Value) Then
      If "" & Adodc2.Recordset.Fields("a1505").Value <> "" Then
         strCurrency = Adodc2.Recordset.Fields("a1505").Value
      Else
         If Mid(Adodc2.Recordset.Fields("a1505").Value, 1, 2) = "US" Then
            strCurrency = Adodc2.Recordset.Fields("a1505").Value
         Else
            strCurrency = Adodc2.Recordset.Fields("a1505").Value
         End If
      End If
      AutoGen
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()
   If Adodc3.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc3.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc3.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text1.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   Screen.MousePointer = vbHourglass
   Do While Adodc1.Recordset.EOF = False
      Adodc2Save
      Adodc1.Recordset.Requery
      If strControlButton = MsgText(602) Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   Loop
   Adodc2.Recordset.Requery
   If IsNull(Adodc2.Recordset.Fields("a1505").Value) Then
      strCurrency = "USD"
   Else
      strCurrency = Adodc2.Recordset.Fields("a1505").Value
   End If
   AutoGen
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command5_Click()
   If adoacc1b0.RecordCount = 0 Or Text1 = MsgText(601) Or Text3 = MsgText(601) Then
      Exit Sub
   End If
   adoacc1b0.Find "a1b01 = '" & Text3 & "'", 0, adSearchForward, 1
   If adoacc1b0.EOF = False Then
'      adoacc1b0.Find "a1b02 = '" & Text1 & "'", 0, adSearchForward, adoacc1b0.Bookmark
      adoacc1b0.Find "a1b02 = '" & ChgSQL(Text1) & "'", 0, adSearchForward, adoacc1b0.Bookmark
      If adoacc1b0.EOF = False Then
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      Else
         MsgBox MsgText(33), , MsgText(5)
         adoacc1b0.MoveFirst
      End If
   Else
      MsgBox MsgText(33), , MsgText(5)
      adoacc1b0.MoveFirst
   End If
End Sub

Private Sub Command6_Click()
   If Adodc3.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc3.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc3.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text1.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   Screen.MousePointer = vbHourglass
   strOriDocNo = strDocuNo
   Do While Adodc2.Recordset.EOF = False
      Adodc2Delete
      'AutoGen
      Adodc2.Recordset.MoveNext
   Loop
   Adodc1.Recordset.Requery
   Adodc2.Recordset.Requery
'   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'"
   '2014/3/26 modify by sonia 取消a1p01
   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
   adoTaie.Execute "delete from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
   Adodc3.Recordset.Requery
   Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid2_SelChange(Cancel As Integer)
   If Adodc3.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSerialNo = Adodc3.Recordset.Fields("a1p03").Value
   Adodc3Show
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc1b0.RecordCount <> 0 Then
      adoacc1b0.MoveFirst
   End If
   adoacc1b0.Find "a1b01 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoacc1b0.EOF = False Then
'      adoacc1b0.Find "a1b02 = '" & strCustNo & "'", 0, adSearchForward, adoacc1b0.Bookmark
      adoacc1b0.Find "a1b02 = '" & ChgSQL(strCustNo) & "'", 0, adSearchForward, adoacc1b0.Bookmark
      If adoacc1b0.EOF = False Then
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
'   Me.Width = 8850
'   Me.Height = 5985
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
   'Modify by Amy 2023/08/18 W8850 H6270
   PUB_InitForm Me, 8900, 6480, strBackPicPath1
   'end 2021/12/07
   
   Combo2.AddItem ComboItem(71)
   Combo2.AddItem ComboItem(72)
   Combo2.AddItem ComboItem(73)
   Combo2.AddItem ComboItem(74)
   Combo2.AddItem ComboItem(75)
   Combo2.AddItem ComboItem(76)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   If adoacc1b0.RecordCount <> 0 Then
      adoacc1b0.MoveLast
      adoacc1b0.MoveFirst
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
   Set Frmacc21d0 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
Dim strMsg As String 'Add by Amy 2014/11/04
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   'Add by Amy 2014/11/04 +系統日檢查
   If ChkWorkData("1", DBDATE(MaskEdBox1), strMsg) = False Then
        MsgBox Label2 & strMsg, , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      MsgBox Label6 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label6 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   '2005/7/19 CANCEL BY SONIA 因有輸代理人名稱之情形 例 : Fortu
   'KeyAscii = UpperCase(KeyAscii)
   'add by sonia 2013/8/7 第一碼才轉大寫
   If Text1.SelStart = 0 Then KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc1b0.CursorLocation = adUseClient
   'modify by sonia 2017/1/17 進入有點慢
   'Modified by Morgan 2018/3/22 婉莘反應需要能看舊資料先改2年
   'adoacc1b0.Open "select * from acc1b0 order by a1b01 asc, a1b02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1b0.Open "select * from acc1b0 where a1b03>='" & strSrvDate(2) - 20000 & "' order by a1b01 asc, a1b02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2017/1/17
   adoadodc1.CursorLocation = adUseClient
'   adoadodc1.Open "select * from acc190, acc180 where acc190.a1901 = acc180.a1801 and a1915 = " & Val(FCDate(MaskEdBox2.Text)) & " and a1803 = '" & Text1 & "' and a1908 is null order by a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc190, acc180 where acc190.a1901 = acc180.a1801 and a1915 = " & Val(FCDate(MaskEdBox2.Text)) & " and a1803 = '" & ChgSQL(Text1) & "' and a1908 is null order by a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   adoadodc2.CursorLocation = adUseClient
'   adoadodc2.Open "select * from (select a1505, a1c03 from acc1c0, acc150, acc190 where a1c03 = a1501 and a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & Text1 & "%" & "' union " & _
'                  "select a1605 as a1505, a1c03 from acc1c0, acc160, acc190 where a1c03 = a1601 and a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & Text1 & "%" & "' union " & _
'                  "select a1903 as a1505, a1c03 from acc1c0, acc190 where a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & Text1 & "%" & "') new order by decode(substr(a1c03, 1, 1), 'U', 1, 'V', 2, 'O', 3, 'B', 4) asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc2.Open "select * from (select a1505, a1c03 from acc1c0, acc150, acc190 where a1c03 = a1501 and a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & ChgSQL(Text1) & "%" & "' union " & _
                  "select a1605 as a1505, a1c03 from acc1c0, acc160, acc190 where a1c03 = a1601 and a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & ChgSQL(Text1) & "%" & "' union " & _
                  "select a1903 as a1505, a1c03 from acc1c0, acc190 where a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & ChgSQL(Text1) & "%" & "') new order by decode(substr(a1c03, 1, 1), 'U', 1, 'V', 2, 'O', 3, 'B', 4) asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc2.Recordset = adoadodc2
   adoadodc3.CursorLocation = adUseClient
'   adoadodc3.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc3.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc3.Recordset = adoadodc3
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
Dim StrSQLa As String

On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
'        strSQLA = "select * from acc190, acc180 where a1901 = a1801 and a1803 = '" & Text1 & "' and (a1908 is null or a1908 = '') order by a1917 asc, a1902 asc "
        StrSQLa = "select * from acc190, acc180 where a1901 = a1801 and a1803 = '" & ChgSQL(Text1) & "' and (a1908 is null or a1908 = '') order by a1917 asc, a1902 asc "
'      adoadodc1.Open "select * from acc190, acc180 where a1901 = a1801 and a1803 = '" & Text1 & "' and (a1908 is null or a1908 = '') order by a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
        adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   Else
'        strSQLA = "select a1902, Decode(a0k11,Null,Decode(SK02,1,'2','5','2','1'),a0k11) as a0k11 from acc190, acc180, fagent, nation, (Select CP01, CP61, CP60 From acc190, acc180, CaseProgress Where a1901=a1801 and a1902=CP61 And a1803='A' Group By CP01, CP61, CP60 ) C1 , acc0k0, acc170, Systemkind where a1701='1' and a1702=a1902 and C1.CP01=SK01 and a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null And a1902=C1.CP61 And C1.CP60=a0k01(+) And a1803='A'  group by a1902, Decode(a0k11,Null,Decode(SK02,1,'2','5','2','1'),a0k11) "
'        strSQLA = strSQLA & " Union select a1902, Decode(a0k11,Null,Decode(SK02,1,'2','5','2','1'),a0k11) as a0k11 from acc190, acc180, fagent, nation, (Select CP01, CP62, CP60 From acc190, acc180, CaseProgress Where a1901=a1801 and a1902=CP62 And a1803='A' Group By CP01, CP62, CP60 ) C1 , acc0k0, acc170, Systemkind where a1701='1' and a1702=a1902 and C1.CP01=SK01 and a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null And a1902=C1.CP62 And C1.CP60=a0k01(+) And a1803='A' group by a1902, Decode(a0k11,Null,Decode(SK02,1,'2','5','2','1'),a0k11) "
'        strSQLA = strSQLA & " Union select a1902, Decode(a0k11,Null,Decode(SK02,1,'2','5','2','1'),a0k11) as a0k11 from acc190, acc180, fagent, nation, (Select CP01, CP63, CP60 From acc190, acc180, CaseProgress Where a1901=a1801 and a1902=CP63 And a1803='A' Group By CP01, CP63, CP60 ) C1 , acc0k0, acc170, Systemkind where a1701='1' and a1702=a1902 and C1.CP01=SK01 and a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null And a1902=C1.CP63 And C1.CP60=a0k01(+) And a1803='A' group by a1902, Decode(a0k11,Null,Decode(SK02,1,'2','5','2','1'),a0k11) "
'        strSQLA = strSQLA & " Union select a1902, Decode(a0k11,Null,Decode(SK02,1,'2','5','2','1'),a0k11) as a0k11 from acc190, acc180, fagent, nation, acc161, acc160, CaseProgress, acc0k0, acc170, Systemkind where a1701='2' and a1702=a1902 and CP01=SK01 and a1901 = a1801 and a1902=axg01 and axg01=a1601 and axg02=cp09 and cp60=a0k01(+) and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null And a1803='A' group by a1902, Decode(a0k11,Null,Decode(SK02,1,'2','5','2','1'),a0k11) "
'        strSQLA = strSQLA & " Union select a1902, '2' as a0k11 from acc190, acc180, fagent, nation, acc170 where a1701='3' and a1702=a1902 and a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null And a1803='A' group by a1902 "
'        strSQLA = strSQLA & " Union select a1902, '2' as a0k11 from acc190, acc180, fagent, nation, acc170 where a1701='4' and a1702=a1902 and a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null And a1803='A' group by a1902 "
'        strSQLA = strSQLA & " Order By 2, 1 "
        StrSQLa = "select * from acc190, acc180 where a1901 = a1801 and a1803 = 'A' and (a1908 is null or a1908 = '') order by a1917 asc, a1902 asc"
'      adoadodc1.Open "select * from acc190, acc180 where a1901 = a1801 and a1803 = 'A' and (a1908 is null or a1908 = '') order by a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
        adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   End If
   Adodc1.Recordset.Requery
   'add by sonia 2014/8/7
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("A1812").Value) Then
         strA1812 = ""
      Else
         strA1812 = Adodc1.Recordset.Fields("A1812").Value
      End If
   Else
      strA1812 = ""
   End If
   'end 2014/8/7

   adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
'   adoadodc2.Open "select * from (select a1505, a1c03 from acc1c0, acc150, acc190 where a1c03 = a1501 and a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & Text1 & "%" & "' union " & _
'                  "select a1605 as a1505, a1c03 from acc1c0, acc160, acc190 where a1c03 = a1601 and a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & Text1 & "%" & "' union " & _
'                  "select a1903 as a1505, a1c03 from acc1c0, acc190 where a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & Text1 & "%" & "') new order by decode(substr(a1c03, 1, 1), 'U', 1, 'V', 2, 'O', 3, 'B', 4) asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc2.Open "select * from (select a1505, a1c03, a1812 from acc1c0, acc150, acc190, acc180 where a1c03 = a1501 and a1c03 = a1902 and a1901=a1801(+) and a1c01 = '" & Text3 & "' and a1c02 like '" & ChgSQL(Text1) & "%" & "' union " & _
                  "select a1605 as a1505, a1c03, a1812 from acc1c0, acc160, acc190, acc180 where a1c03 = a1601 and a1c03 = a1902 and a1901=a1801(+) and a1c01 = '" & Text3 & "' and a1c02 like '" & ChgSQL(Text1) & "%" & "' union " & _
                  "select a1903 as a1505, a1c03, a1812 from acc1c0, acc190, acc180 where a1c03 = a1902 and a1901=a1801(+) and a1c01 = '" & Text3 & "' and a1c02 like '" & ChgSQL(Text1) & "%" & "') new order by decode(substr(a1c03, 1, 1), 'U', 1, 'V', 2, 'O', 3, 'B', 4) asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc2.Recordset.Requery
   If Adodc2.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc2.Recordset.Fields("a1505").Value) Then
         strCurrency = "USD"
      Else
         strCurrency = Adodc2.Recordset.Fields("a1505").Value
      End If
      'add by sonia 2017/1/18
      If IsNull(Adodc2.Recordset.Fields("A1812").Value) Then
         strA1812 = ""
      Else
         strA1812 = Adodc2.Recordset.Fields("A1812").Value
      End If
      'end 2017/1/18
   Else
      strCurrency = "USD"
      strA1812 = ""     'add by sonia 2017/1/18
   End If
   adoadodc3.Close
   adoadodc3.CursorLocation = adUseClient
'   adoadodc3.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2014/3/26 modify by sonia 取消a1p01
   'adoadodc3.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc3.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc3.Recordset.Requery
   If Adodc3.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p22").Value) Then
         strDocuNo = "null"
         strDYes = "null"
      Else
         strDocuNo = "'" & Adodc3.Recordset.Fields("a1p22").Value & "'"
         strDYes = "'Y'"
      End If
   Else
      If strOriDocNo <> "" And strOriDocNo <> "null" Then
         strDocuNo = strOriDocNo
         strDYes = "'Y'"
         strOriDocNo = ""
      Else
         strDocuNo = "null"
         strDYes = "null"
      End If
   End If
'   If Adodc3.Recordset.RecordCount <> 0 Then
'      Adodc3.Recordset.Find "a1p03 = '" & strSerialNo & "'", 0, adSearchForward, 1
'      If Adodc3.Recordset.EOF Then
'         Exit Sub
'      Else
'         DataGrid2.SelBookmarks.Add Adodc3.Recordset.Bookmark
'      End If
'   End If
   'Mark by Amy 2015/07/22 若先選某筆資料才按修改鈕,導致新增一筆Record
   'strSerialNo = MsgText(601)
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
   Text1 = adoacc1b0.Fields("a1b02").Value
'   If Len(Text1) = 6 Then
'      Text2 = FagentQuery(AfterZero(Text1), 2)
'      Text1 = AfterZero(Text1)
'   Else
   Text2 = FagentQuery(Text1, 2)
   If Text2 = MsgText(601) Then
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select a1810 from acc180 where a1803 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select a1810 from acc180 where a1803 = '" & ChgSQL(Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         Text2 = MsgText(601)
      Else
         If IsNull(adoquery.Fields("a1810").Value) Then
            Text2 = MsgText(601)
         Else
            Text2 = adoquery.Fields("a1810").Value
         End If
      End If
      adoquery.Close
   End If
'   End If
   Text3 = adoacc1b0.Fields("a1b01").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc1b0.Fields("a1b03").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc1b0.Fields("a1b03").Value)
   End If
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc1b0.Fields("a1b05").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc1b0.Fields("a1b05").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc1b0.Fields("a1b06").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Combo2.List(Val(adoacc1b0.Fields("a1b06").Value) - 1)
   End If
   If IsNull(adoacc1b0.Fields("a1b04").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc1b0.Fields("a1b04").Value
   End If
   If IsNull(adoacc1b0.Fields("a1b07").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc1b0.Fields("a1b07").Value
   End If
    'Add By Cheng 2004/02/02
    '顯示傳票號碼
    ShowA1P22 Me.Text3.Text, Me.Text1.Text, Me.MaskEdBox1.Text
    'End
    'Add by Amy 2014/11/04 a1p22有值不可修改結匯日
    If Text23 = "" Then
        MaskEdBox1.Enabled = True
    Else
        MaskEdBox1.Enabled = False
    End If
    'end 2014/11/04
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text1)
      Case 6
         Text1 = Text1 & "000"
      Case 8
         Text1 = Text1 & "0"
   End Select
   If ExistCheck("fagent", "fa01", Mid(Text1, 1, 8), Label1, False) = False Then
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select a1810 from acc180 where a1803 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select a1810 from acc180 where a1803 = '" & ChgSQL(Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         MsgBox MsgText(45) & Label1, , MsgText(5)
         Cancel = True
         adoquery.Close
         Exit Sub
      Else
         If IsNull(adoquery.Fields("a1810").Value) Then
            Text2 = MsgText(601)
         Else
            Text2 = adoquery.Fields("a1810").Value
         End If
         adoquery.Close
         '2012/2/14 modify by sonia
         'Exit Sub
         If Text2 <> "" Then Exit Sub
      End If
   End If
   Text2 = FagentQuery(Text1, 2)
   'Add by Morgan 2010/10/28
   If Text2 = "" Then
      Text2 = FagentQuery(Text1, 1)
      If Text2 = "" Then
         Text2 = FagentQuery(Text1, 3)
      End If
   End If
   '2012/2/14 add by sonia
   If Text2 = "" Then
      Text2 = GetCustomerName(Text1, 1)
   End If
   '2012/2/14 end
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   Select Case Text6
      Case "1"
         '2005/5/4 MODIFY BY SONIA
         'Text17 = Val(Format(Val(Text11) * Val(Text7), DAmount))
         Text17 = Val(Format(Val(Text11) * Val(Text7), FAmount))
         '2005/5/4 END
      Case "2"
   End Select
End Sub

Private Sub Text13_Change()
   Text16 = A0102Query(Text13)
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   If Text13 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc010", "a0101", Text13, Label13) = False Then
      Cancel = True
      Exit Sub
   End If
   If Text6 = "2" Then
'      Select Case Text13
'         Case "110205", "110206", "110208", "110218"
'            adoquery.CursorLocation = adUseClient
'            adoquery.Open "select a1x02 from acc1x0 where a1x01 = '" & strCurrency & "'", adoTaie, adOpenStatic, adLockReadOnly
'            If adoquery.RecordCount <> 0 Then
'               If IsNull(adoquery.Fields("a1x02").Value) Then
'                  Text7 = "1"
'               Else
'                  Text7 = adoquery.Fields("a1x02").Value
'               End If
'            Else
'               Text7 = "1"
'            End If
'            adoquery.Close
'         Case Else
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select * from acc1x0 where a1x01 = '" & strCurrency & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               If IsNull(adoquery.Fields("a1x02").Value) Then
                  Text7 = "1"
               Else
                  Text7 = adoquery.Fields("a1x02").Value
               End If
            Else
               Text7 = "1"
            End If
            adoquery.Close
'      End Select
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a0h01, a0h02 from acc0h0 where a0h08 = '" & Text13 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("a0h02").Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = adoquery.Fields("a0h02").Value
      End If
      If IsNull(adoquery.Fields("a0h01").Value) Then
         Text9 = MsgText(601)
      Else
         Text9 = adoquery.Fields("a0h01").Value
      End If
   Else
      Text8 = MsgText(601)
      Text9 = MsgText(601)
   End If
   adoquery.Close
   If Text6 = "2" Then
      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select sum(a1904) from acc190, acc1c0 where a1902 = a1c03 and a1c01 = '" & Text3 & "' and a1c02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select sum(a1904) from acc190, acc1c0 where a1902 = a1c03 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            Text11 = "0"
         Else
            Text11 = adoquery.Fields(0).Value
         End If
      Else
         Text11 = "0"
      End If
      adoquery.Close
   End If
End Sub

Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text17_LostFocus()
    'Added by Lydia 2016/08/02 規費只能輸入整數
    If Left(Trim(Text13), 4) = "2201" And Text17 <> "" And Text17 <> Format(Val(Text17), DAmount) Then
        MsgBox "規費只能輸入整數!", vbCritical
        Text17.SetFocus
    End If
    'end 2016/08/02
End Sub

Private Sub Text18_GotFocus()
   TextInverse Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text18_Validate(Cancel As Boolean)
   If Text18 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text18, Label19) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Text13, Text18) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text19_GotFocus()
   TextInverse Text19
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text21_GotFocus()
   TextInverse Text21
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text21_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text21 <> MsgText(601) Then
      Text21 = CaseNoZero(Text21)
      adoquery.CursorLocation = adUseClient
      'modify by sonia 2017/9/26 +特殊出名公司
      adoquery.Open "select pa01 as SystemNo,pa161 from patent where pa01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and pa02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and pa03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and pa04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo,tm130 pa161 from trademark where tm01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and tm02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and tm03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and tm04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo,lc48 pa161 from lawcase where lc01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and lc02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and lc03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and lc04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo,'' pa161 from hirecase where hc01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and hc02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and hc03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and hc04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo,sp85 pa161 from servicepractice where sp01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and sp02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and sp03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and sp04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         MsgBox MsgText(28) & Label21, , MsgText(5)
         Cancel = True
         adoquery.Close
         Exit Sub
      'add by sonia 2017/9/26
      Else
         If "" & adoquery.Fields("pa161") = "J" Then
           MsgBox "請注意！此為智權公司出名案件！", , MsgText(5)
         End If
      'end 2017/9/26
      End If
      adoquery.Close
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text22_GotFocus()
   TextInverse Text22
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text22_Validate(Cancel As Boolean)
   If Text22 <> MsgText(601) Then
      If ExistCheck("staff", "st01", Text22, Label23) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Then
      If adoquery.State = adStateOpen Then
         adoquery.CursorLocation = adUseClient
      End If
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select a1b01 from acc1b0 where a1b01 = '" & Text3 & "' and a1b03 = " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         MsgBox MsgText(215), , MsgText(5)
         Cancel = True
         Text3.SetFocus
         adoquery.Close
         Exit Sub
      End If
      adoquery.Close
   End If
   Screen.MousePointer = vbHourglass
   AdodcRefresh
   Screen.MousePointer = vbDefault
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
   'edit by nickc 2007/06/11  切換輸入法改用API
   'OpenIme 'Mark by Amy 2016/06/28 瑞婷:不要切中文輸入
End Sub

'*************************************************
'  儲存資料表(國外匯票資料(交易檔))
'
'*************************************************
Private Sub Adodc2Save()
On Error GoTo Checking
   strControlButton = MsgText(601)
   If Text1 = MsgText(601) Then
      MsgBox MsgText(10) & Label1, , MsgText(5)
      strControlButton = MsgText(602)
      Text1.SetFocus
      Exit Sub
   Else
      If Text3 = MsgText(601) Then
         MsgBox MsgText(10) & Label3, , MsgText(5)
         strControlButton = MsgText(602)
         Text3.SetFocus
         Exit Sub
      End If
   End If
   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select * from acc1c0 where a1c01 = '" & Text3 & "' and a1c02 = '" & Text1 & "' and a1c03 = '" & Adodc1.Recordset.Fields("a1902").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select * from acc1c0 where a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' and a1c03 = '" & Adodc1.Recordset.Fields("a1902").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      MsgBox MsgText(9), , MsgText(5)
      strControlButton = MsgText(602)
      adoquery.Close
      Text1.SetFocus
      Exit Sub
   End If
   adoquery.Close
'   adoTaie.Execute "insert into acc1c0 (a1c01, a1c02, a1c03, a1c04, a1c05, a1c06) values ('" & Text3 & "', '" & Text1 & "', '" & Adodc1.Recordset.Fields("a1902").Value & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "')"
   adoTaie.Execute "insert into acc1c0 (a1c01, a1c02, a1c03, a1c04, a1c05, a1c06) values ('" & Text3 & "', '" & ChgSQL(Text1) & "', '" & Adodc1.Recordset.Fields("a1902").Value & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "')"
   adoTaie.Execute "update acc190 set a1908 = '" & Text3 & "' where a1902 = '" & Adodc1.Recordset.Fields("a1902").Value & "'"
   UpdateA150 Adodc1.Recordset.Fields("a1902").Value 'Add by Morgan 2007/8/23
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除資料表(國外匯票資料(交易檔))
'
'*************************************************
Private Sub Adodc2Delete()
On Error GoTo Checking
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   UpdateA150 Adodc2.Recordset.Fields("a1c03"), -1 'Add by Morgan 2007/8/23
   adoTaie.Execute "update acc190 set a1908 = null where a1902 = '" & Adodc2.Recordset.Fields("a1c03").Value & "'"
   adoTaie.Execute "delete from acc1c0 where a1c03 = '" & Adodc2.Recordset.Fields("a1c03").Value & "'"
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存 Adodc3 之資料
'
'*************************************************
Private Sub Adodc3Save()
Dim m_strA1P23 As String      'ADD BY SONIA 2025/7/24 結匯帳單編號

On Error GoTo Checking
   If Text13 = MsgText(601) Then
      MsgBox MsgText(10) & Label13, , MsgText(5)
      strControlButton = MsgText(602)
      Text13.SetFocus
      Exit Sub
   Else
      If ExistCheck("acc010", "a0101", Text13, Label13) = False Then
         strControlButton = MsgText(602)
         Text13.SetFocus
         Exit Sub
      End If
      If CheckDept(Text13, Text18) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text18.SetFocus
         Exit Sub
      End If
      If Text9 <> MsgText(601) Then
         If ExistCheck("acc0g0", "a0g01", Text9, Label10) = False Then
            strControlButton = MsgText(602)
            Text9.SetFocus
            Exit Sub
         End If
      End If
   End If
   
   'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(Text13, Val(FCDate(MaskEdBox1.Text)))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      Text13.SetFocus
      Exit Sub
   End If
   'end 2015/12/30
   'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Text13, Text18, Text22)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Text13.SetFocus
      ElseIf intI = 2 Then
         Text18.SetFocus
      ElseIf intI = 3 Then
         Text22.SetFocus
      End If
      Exit Sub
   End If
   'end 2007/10/2
   
    'Added by Lydia 2016/08/02 規費只能輸入整數
    If Left(Trim(Text13), 4) = "2201" And Text17 <> "" And Text17 <> Format(Val(Text17), DAmount) Then
        MsgBox "規費只能輸入整數!", vbCritical
        Text17.SetFocus
        Exit Sub
    End If
    'end 2016/08/02
    
   If Text21 <> MsgText(601) Then
      Text21 = CaseNoZero(Text21)
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and pa02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and pa03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and pa04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and tm02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and tm03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and tm04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and lc02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and lc03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and lc04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and hc02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and hc03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and hc04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text21, 1, Len(Text21) - 9) & "' and sp02 = '" & Mid(Text21, Len(Text21) - 8, 6) & "' and sp03 = '" & Mid(Text21, Len(Text21) - 2, 1) & "' and sp04 = '" & Mid(Text21, Len(Text21) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         MsgBox MsgText(28) & Label21, , MsgText(5)
         strControlButton = MsgText(602)
         adoquery.Close
         Exit Sub
      End If
      adoquery.Close
   End If
   If Text22 <> MsgText(601) Then
      If ExistCheck("staff", "st01", Text22, Label23) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If
   End If
   If Adodc3.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc3.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc3.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            strControlButton = MsgText(602)
            Text6.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   strA1P01 = "" & Adodc3.Recordset.Fields("a1p01").Value  '2014/4/21 ADD BY SONIA
   adoacc1p0.CursorLocation = adUseClient
'   adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text3 & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   '2014/3/26 modify by sonia 取消a1p01
   'adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoacc1p0.Open "select * from acc1p0 where a1p02 = 'I' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   m_strA1P23 = ""   'add by sonia 2025/7/24
   If adoacc1p0.RecordCount = 0 Then
      Adodc3.Recordset.AddNew
      '2014/3/26 modify by sonia 加入J公司
      'Adodc3.Recordset.Fields("a1p01").Value = "1"
      '2014/4/21 MODIFY BY SONIA Adodc3.Recordset.Fields("a1p01").Value 改為strA1P01
      Adodc3.Recordset.Fields("a1p01").Value = strA1P01
      Adodc3.Recordset.Fields("a1p02").Value = "I"
'      Adodc3.Recordset.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", 3)
      '2014/3/26 modify by sonia 取消a1p01
      'Adodc3.Recordset.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
      Adodc3.Recordset.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
      strSerialNo = Adodc3.Recordset.Fields("a1p03").Value
      Adodc3.Recordset.Fields("a1p04").Value = Text3 & Text1
   'add by sonia 2025/7/24
   Else
      m_strA1P23 = "" & Adodc3.Recordset.Fields("a1p23").Value
   'end 2025/7/24
   End If
   adoacc1p0.Close
   Adodc3.Recordset.Fields("a1p05").Value = Text13
   If Text7 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p20").Value = Val(Text7)
   Else
      Adodc3.Recordset.Fields("a1p20").Value = 0
   End If
   If Text11 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p21").Value = Val(Text11)
   Else
      Adodc3.Recordset.Fields("a1p21").Value = 0
   End If
   If Text17 <> MsgText(601) Then
      Select Case Text6
         Case "1"
            Adodc3.Recordset.Fields("a1p07").Value = Val(Text17)
            Adodc3.Recordset.Fields("a1p08").Value = 0
         Case "2"
            Adodc3.Recordset.Fields("a1p08").Value = Val(Text17)
            Adodc3.Recordset.Fields("a1p07").Value = 0
         Case Else
            Adodc3.Recordset.Fields("a1p07").Value = 0
            Adodc3.Recordset.Fields("a1p08").Value = 0
      End Select
   Else
      Adodc3.Recordset.Fields("a1p07").Value = 0
      Adodc3.Recordset.Fields("a1p08").Value = 0
   End If
   If Text9 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p10").Value = Text9
   Else
      Adodc3.Recordset.Fields("a1p10").Value = Null
   End If
   If Text8 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p11").Value = Text8
   Else
      Adodc3.Recordset.Fields("a1p11").Value = Null
   End If
   If Combo1 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p14").Value = Combo1
   Else
      Adodc3.Recordset.Fields("a1p14").Value = Null
   End If
   If Text18 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p06").Value = Text18
   Else
      Adodc3.Recordset.Fields("a1p06").Value = MsgText(55)
   End If
   If Text21 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p17").Value = Text21
   Else
      Adodc3.Recordset.Fields("a1p17").Value = Null
   End If
   If Text19 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p30").Value = Text19
   Else
      Adodc3.Recordset.Fields("a1p30").Value = Null
   End If
   If Text22 <> MsgText(601) Then
      Adodc3.Recordset.Fields("a1p16").Value = Text22
   Else
      Adodc3.Recordset.Fields("a1p16").Value = Null
   End If
   If strDocuNo = "null" Then
      Adodc3.Recordset.Fields("a1p22").Value = Null
   Else
      Adodc3.Recordset.Fields("a1p22").Value = Replace(strDocuNo, "'", "")
   End If
   If strDYes = "null" Then
      Adodc3.Recordset.Fields("a1p27").Value = Null
   Else
      Adodc3.Recordset.Fields("a1p27").Value = "Y"
   End If
      
   Adodc3.Recordset.Fields("a1p18").Value = Val(FCDate(MaskEdBox1.Text)) 'Added by Morgan 2022/6/29
   
   Adodc3.Recordset.UpdateBatch
   'CreditGen
   Adodc3.Recordset.Requery
   If Adodc3.Recordset.RecordCount <> 0 Then
      Adodc3.Recordset.Find "a1p03 = '" & strSerialNo & "'", 0, adSearchForward, 1
      If Adodc3.Recordset.EOF Then
         Exit Sub
      Else
         DataGrid2.SelBookmarks.add Adodc3.Recordset.Bookmark
      End If
   End If
   strSerialNo = MsgText(601)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示 Adodc3 之資料
'
'*************************************************
Private Sub Adodc3Show()
   If Adodc3.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p07").Value) Or Adodc3.Recordset.Fields("a1p07").Value = 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p08").Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = "2"
      End If
   Else
      Text6 = "1"
   End If
   Select Case Text6
      Case "1"
         If IsNull(Adodc3.Recordset.Fields("a1p07").Value) Then
            Text17 = MsgText(601)
         Else
            Text17 = Adodc3.Recordset.Fields("a1p07").Value
         End If
      Case "2"
         If IsNull(Adodc3.Recordset.Fields("a1p08").Value) Then
            Text17 = MsgText(601)
         Else
            Text17 = Adodc3.Recordset.Fields("a1p08").Value
         End If
   End Select
   Text13 = Adodc3.Recordset.Fields("a1p05").Value
   If IsNull(Adodc3.Recordset.Fields("a1p20").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc3.Recordset.Fields("a1p20").Value
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p21").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc3.Recordset.Fields("a1p21").Value
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p10").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc3.Recordset.Fields("a1p10").Value
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p11").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc3.Recordset.Fields("a1p11").Value
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p14").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc3.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p06").Value) Then
      Text18 = MsgText(601)
   Else
      Text18 = Adodc3.Recordset.Fields("a1p06").Value
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p17").Value) Then
      Text21 = MsgText(601)
   Else
      Text21 = Adodc3.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p30").Value) Then
      Text19 = MsgText(601)
   Else
      Text19 = Adodc3.Recordset.Fields("a1p30").Value
   End If
   If IsNull(Adodc3.Recordset.Fields("a1p16").Value) Then
      Text22 = MsgText(601)
   Else
      Text22 = Adodc3.Recordset.Fields("a1p16").Value
   End If
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
         
         'Frmacc21d0_Save
         If strControlButton <> MsgText(602) Then
            Adodc3Save
         End If
         If strControlButton <> MsgText(602) Then
            DataGrid2.Refresh
            SumShow
            Adodc3Clear
            Text6.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Public Sub Adodc3Clear()
   Text6 = ""
   Text13 = ""
   Text16 = ""
   Text7 = ""
   Text11 = ""
   Text17 = ""
   Text9 = ""
   Text10 = ""
   Text8 = ""
   Combo1 = ""
   Text18 = ""
   Text21 = ""
   Text19 = ""
   Text22 = ""
End Sub

Private Sub Text5_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
'CloseIme 'Mark by Amy 2022/01/11 莘:改完Form2.0 有時跳至此欄位會自動變中文
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text9_Change()
   Text10 = A0g02Query(Text9)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub Adodc3Delete()
On Error GoTo Checking
   If Adodc3.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
'   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' and a1p05 = '7128'"
   '2014/3/26 modify by sonia 取消a1p01
   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and a1p05 = '7128'"
   adoTaie.Execute "delete from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and a1p05 = '7128'"
'   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text3 & Text1 & "'"
   '2014/3/26 modify by sonia 取消a1p01
   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text3 & ChgSQL(Text1) & "'"
   adoTaie.Execute "delete from acc1p0 where a1p02 = 'I' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text3 & ChgSQL(Text1) & "'"
   AdodcRefresh
   Adodc3Clear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*), sum(decode(substr(a1p05, 1, 1), '2', a1p21, 0)) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   '2012/9/17 modify by sonia 外幣金額V單號應減非加Y51350000匯票A10101226
   'adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*), sum(decode(substr(a1p05, 1, 1), '2', a1p21, 0)) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   '2014/3/26 modify by sonia 取消a1p01
   'adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*), sum(decode(substr(a1p05, 1, 1), '2', decode(substr(a1p23,1,1),'V',a1p21*-1,a1p21), 0)) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
  'Modify by Amy 2017/10/24 E_Fail Err
   adoaccsum.Open "select Nvl(sum(a1p07),0), Nvl(sum(a1p08),0), count(*), Nvl(sum(decode(substr(a1p05, 1, 1), '2', decode(substr(a1p23,1,1),'V',a1p21*-1,a1p21), 0)),0) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.EOF = False And adoaccsum.BOF = False Then
      If adoaccsum.Fields(0).Value = 0 And Text1 = "" And Text3 = "" Then
         Text14 = MsgText(601)
      Else
         Text14 = Format(Val(adoaccsum.Fields(0).Value), FAmount)
      End If
      If adoaccsum.Fields(1).Value = 0 And Text1 = "" And Text3 = "" Then
         Text12 = MsgText(601)
      Else
         Text12 = Format(Val(adoaccsum.Fields(1).Value), FAmount)
      End If
      If adoaccsum.Fields(2).Value = 0 And Text1 = "" And Text3 = "" Then
         Text20 = MsgText(601)
      Else
         Text20 = Format(adoaccsum.Fields(2).Value, DAmount)
      End If
      If adoaccsum.Fields(3).Value = 0 And Text1 = "" And Text3 = "" Then
         Text15 = MsgText(601)
      Else
         Text15 = Format(adoaccsum.Fields(3).Value, FAmount)
      End If
   Else
      Text14 = MsgText(601)
      Text12 = MsgText(601)
      Text20 = MsgText(601)
      Text15 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc1b0.Bookmark & MsgText(35) & adoacc1b0.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text6.Enabled = False
   Text13.Enabled = False
   Text7.Enabled = False
   Text11.Enabled = False
   Text17.Enabled = False
   Text8.Enabled = False
   Text9.Enabled = False
   Combo1.Enabled = False
   Text18.Enabled = False
   Text21.Enabled = False
   Text19.Enabled = False
   Text22.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command6.Enabled = False
   'Added by Lydia 2017/11/01
   cmdC(0).Enabled = False
   cmdC(1).Enabled = False
   'Added by Lydia 2020/08/31
   cmdC(2).Enabled = False
   cmdC(3).Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text6.Enabled = True
   Text13.Enabled = True
   Text7.Enabled = True
   Text11.Enabled = True
   Text17.Enabled = True
   Text8.Enabled = True
   Text9.Enabled = True
   Combo1.Enabled = True
   Text18.Enabled = True
   Text21.Enabled = True
   Text19.Enabled = True
   Text22.Enabled = True
   Command1.Enabled = True
   Command2.Enabled = True
   Command3.Enabled = True
   Command4.Enabled = True
   Command6.Enabled = True
   'Added by Lydia 2017/11/01
   cmdC(0).Enabled = True
   cmdC(1).Enabled = True
   'Added by Lydia 2020/08/31
   cmdC(2).Enabled = True
   cmdC(3).Enabled = True
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text9, Label10) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

'*************************************************
'  借貸方檢核
'
'*************************************************
Public Function CreDebCheck() As String
   If Text14 = Text12 Then
      CreDebCheck = MsgText(602)
   End If
End Function

'*************************************************
'  自動產生借方科目
'
'*************************************************
Public Sub AutoGen()
Dim douRate As Double
Dim douAmount As Double
Dim StrSQLa As String
Dim m_strDomAmt  As String    ' 國內收款金額 Add By Cheng 2003/07/28
Dim lngEff As Long            '2005/10/26 ADD BY SONIA
Dim m_strA1P14 As String      '2011/11/23 ADD BY SONIA 結匯傳票摘要
Dim m_strA1P06 As String      'add by sonia 2016/2/17 傳票部門
Dim m_strA1L02 As String, m_strA0K04 As String 'Added by Lydia 2020/10/27 模組取得:取得國內收款日、國內收據抬頭
Dim m_strA1P16 As String      'add by sonia 2021/1/20
Dim m_Amt1 As Double, m_TotAmt As Double 'Added by Lydia 2021/07/26 平均分攤手續費的金額、累計金額
Dim stracc1a0 As String       'add by sonia 2025/8/5

   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' and a1p07 <> 0"
   '2014/3/26 modify by sonia 取消a1p01
   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
   adoTaie.Execute "delete from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
   adoquery.CursorLocation = adUseClient
    'Modify By Cheng 2003/07/23
    '加收據抬頭, 本所案號, 收文號
   '2005/5/4 MODIFY BY SONIA 加暫收款單號A1303,另暫收款之AMOUNT要算至小數二位不取整數
   'strSQLA = "select a1c03, axf03 as Caseno, nvl(round(axf04 * a1906), 0) as Amount, '2201' as Accno, axf04 as Famount, nvl(a1906, 0) as Rate, a1907, axf02, a1901 from acc151, acc1c0, acc190 where axf01 = a1c03 and axf01 = a1902 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' union " & _
   '              "select a1c03, '' as Caseno, nvl(round(a1904 * a1906), 0) as Amount, '2201' as Accno, a1904 as Famount, nvl(a1906, 0) as Rate, a1907, '' as axf02, a1901 from acc1c0, acc190, acc170 where a1c03 = a1902 and a1c03 = a1702 and a1701 = '4' and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' union " & _
   '              "select a1c03, a1208 as Caseno, nvl(round(a1307 * a1906), 0) as Amount, '2401' as Accno, a1307 as Famount, nvl(a1906, 0) as Rate, a1907, '' as axf02, a1901 from acc130, acc120, acc1c0, acc190 where a1303 = a1201 and a1301 = a1c03 and a1301 = a1902 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' order by Accno asc, a1901 asc, Caseno asc, axf02 asc"
   '2012/2/14 modify by sonia 加暫收款之a1210(暫收款退費摘要要用)
   'MODIFY BY SONIA 2014/3/26 加A1917讀取收據公司別
   'modify by sonia 2025/8/5 因為A1906在匯票存檔時會被更新，故此處不可再用A1906改抓ACC1A0之A1A04
   'StrSQLa = "select a1c03, axf03 as Caseno, nvl(round(axf04 * a1906), 0) as Amount, '2201' as Accno, axf04 as Famount, nvl(a1906, 0) as Rate, a1907, axf02, a1901, '' AS A1303, axf01,'' a1210,A1917 from acc151, acc1c0, acc190 where axf01 = a1c03 and axf01 = a1902 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' union " & _
             "select a1c03, '' as Caseno, nvl(round(a1904 * a1906), 0) as Amount, '2201' as Accno, a1904 as Famount, nvl(a1906, 0) as Rate, a1907, '' as axf02, a1901, '' AS A1303, a1902 axf01,'' a1210,A1917 from acc1c0, acc190, acc170 where a1c03 = a1902 and a1c03 = a1702 and a1701 = '4' and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' union " & _
             "select a1c03, a1208 as Caseno, nvl(a1307 * a1906, 0) as Amount, '2401' as Accno, a1307 as Famount, nvl(a1906, 0) as Rate, a1907, '' as axf02, a1901, a1303,a1902 axf01,a1210,A1917 from acc130, acc120, acc1c0, acc190 where a1303 = a1201 and a1301 = a1c03 and a1301 = a1902 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' order by Accno asc, a1901 asc, Caseno asc, axf02 asc"
   stracc1a0 = "(SELECT a1a03,min(a1a01||a1a02||a1a03) aa FROM Acc1a0 WHERE A1A01>=" & Val(FCDate(MaskEdBox1.Text)) & " AND NVL(A1A04,0)>0 group by a1a03) a2"
   'modify by sonia 2025/8/25 2401暫收款退費還是要用A1906(原暫收款匯率)，另台幣金額Amount改抓原傳票貸方金額，否則O11200013(N11200086)會有誤差
   StrSQLa = "select a1c03, axf03 as Caseno, nvl(round(axf04 * A1.A1a04), 0) as Amount, '2201' as Accno, axf04 as Famount, nvl(A1.A1a04, 0) as Rate, a1907, axf02, a1901, '' AS A1303, axf01,'' a1210,A1917 from acc151, acc1c0, acc190,Acc1a0 A1," & stracc1a0 & " where axf01 = a1c03 and axf01 = a1902 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' AND A1903=a2.A1A03(+) and substr(aa,1,7)=A1.a1a01(+) and substr(aa,8,7)=A1.a1a02(+) and substr(aa,15)=A1.a1a03(+) union " & _
             "select a1c03, '' as Caseno, nvl(round(a1904 * A1.A1a04), 0) as Amount, '2201' as Accno, a1904 as Famount, nvl(A1.A1a04, 0) as Rate, a1907, '' as axf02, a1901, '' AS A1303, a1902 axf01,'' a1210,A1917 from acc1c0, acc190, acc170,Acc1a0 A1," & stracc1a0 & " where a1c03 = a1902 and a1c03 = a1702 and a1701 = '4' and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' AND A1903=a2.A1A03(+) and substr(aa,1,7)=A1.a1a01(+) and substr(aa,8,7)=A1.a1a02(+) and substr(aa,15)=A1.a1a03(+) union " & _
             "select a1c03, a1208 as Caseno, nvl(a1p08,nvl(a1307 * a1906, 0)) as Amount, '2401' as Accno, a1307 as Famount, nvl(a1906, 0) as Rate, a1907, '' as axf02, a1901, a1303,a1902 axf01,a1210,A1917 from acc130, acc120, acc1c0, acc190, acc1p0 where a1303 = a1201 and a1301 = a1c03 and a1301 = a1902 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' and a1303=a1p30(+) and a1p08>0 order by Accno asc, a1901 asc, Caseno asc, axf02 asc"
   '2005/5/4 END
   adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   
   strA1P01 = ""   '2014/3/26 ADD BY SONIA
   strAutoGen = "" 'ADD BY SONIA 2014/6/18
   strA1917 = " "  'add by sonia 2017/6/16
   Do While adoquery.EOF = False
'      strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", 3)
      '2014/3/26 modify by sonia 取消a1p01
      'strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
      strA1917 = IIf(adoquery.Fields("A1917").Value = "J", "J", IIf(adoquery.Fields("A1917").Value = "L", "L", "1")) 'Added by Lydia 2020/10/27 收據公司別=>轉傳票公司別
      strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
      If Mid(adoquery.Fields("a1c03").Value, 1, 1) = MsgText(812) Then
         '2011/11/23 ADD BY SONIA 抓各語法的傳票摘要抓出來改一次即可不必每句改,同時把收款日/收款金額改到摘要的最前面,以便外帳核對資料
         'Modified by Lydia 2020/10/27 分開處理
         'm_strA1P14 = GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value, m_strDomAmt) & "/" & m_strDomAmt & " " & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar)
         ''2011/11/23 END
         m_strA1L02 = GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value, m_strDomAmt)
         m_strA0K04 = GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value)
         '若申請人公司名稱為 株式會社(株式�B社)，請修正傳票摘要公司名；請略過 "株式會社(株式�B社) ，直接抓第五個字起的公司名
         If Left(m_strA0K04, 4) = "株式會社" Or Left(m_strA0K04, 4) = "株式�B社" Then
              m_strA0K04 = Mid(m_strA0K04, 5, 4)
         Else
              m_strA0K04 = Left(m_strA0K04, 4)
         End If
         m_strA1P14 = m_strA1L02 & "/" & m_strDomAmt & " " & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & m_strA0K04 & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar)
         'end 2020/10/27
 
         'add by sonia 2016/2/17 案號之系統類別若非總帳做帳部門則要改,FG-001043
         m_strA1P06 = Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9)
         'strA1917 = adoquery.Fields("A1917").Value    'add by sonia 2017/6/16 'Remove by Lydia 2020/10/27
         Select Case m_strA1P06
            Case "CFL", "LIN", "LA"
               m_strA1P06 = "L"
            Case "TB", "TC", "TF", "TM", "TR", "TS", "TT"
               m_strA1P06 = "T"
            Case "CFC", "S"
               m_strA1P06 = "CFT"
            Case "PS"
               m_strA1P06 = "P"
            Case "CPS"
               m_strA1P06 = "CFP"
            Case "FG"
               m_strA1P06 = "FCP"
            Case Else
         End Select
         'end 2016/2/17
         If Len(adoquery.Fields("Caseno").Value) = 12 Then
            If Mid(adoquery.Fields("Caseno").Value, 1, 3) = "CFT" Or Mid(adoquery.Fields("Caseno").Value, 1, 3) = "CFC" Then
               '2005/11/4 MODIFY BY SONIA 收據抬頭原抓a1907改抓收據或客戶名稱,因一帳單多案號時收據抬頭可能不同
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & adoquery.Fields("a1907").Value, 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               'Modify by Morgan 2006/3/29 加a1p23
               'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf02")) & ", " & strDYes & ")"
               '2011/8/19 modify by sonia GetA1l02加傳收款金額m_strDomAmt
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               'Modifed by Lydia 2020/10/27 改變數
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
            'ADD BY SONIA 2015/6/30 再加Y53374北京寰華的翻譯費仍為6130,匯票號碼A10400913之FCP051292000(婧瑄說控制北京寰華的FCP案)
            'modify by sonia 2017/8/7 Y53374北京寰華的CFP案帳單也做翻譯費6130
            'modify by sonia 2019/7/17 Y53374北京寰華的CFP案帳單再改回規費(專利郭經理向財務反應)
            'ElseIf Text1 = "Y53374000" And (Mid(adoquery.Fields("Caseno").Value, 1, 3) = "FCP" Or Mid(adoquery.Fields("Caseno").Value, 1, 3) = "CFP") Then
            'Modify by Amy 2025/01/16 +Y00043 百靈不限定系統別(若以後有OA委外翻譯再討論)-秀玲  ex:A11400142
            'ElseIf Text1 = "Y53374000" And (Mid(adoquery.Fields("Caseno").Value, 1, 3) = "FCP") Then
            'modify by sonia 2025/10/13 Y00043百靈改編號為Y56151，此處取消Y00043改至下方Pub_SetF51Order加入Y56151
            ElseIf Text1 = "Y53374000" And (Mid(adoquery.Fields("Caseno").Value, 1, 3) = "FCP") Then
                'Modifed by Lydia 2020/10/27 改變數
                'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) " & _
                    " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) " & _
                    " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
           '2008/6/25 add by sonia 代理人為Y52268江蘇舜禹翻譯有限公司則科目用6130,部門用系統類別
            '2010/5/4 MODIFY BY SONIA 加 Y53035江蘇通用信息
            '2012/3/2 MODIFY BY SONIA 加 Y53374北京寰華知識產權代理有限公司
            '2012/9/5 MODIFY BY SONIA 加 Y53541南京捷恩凱信息技術有限公司
            '2014/12/30 modify by sonia 取消Y53374匯票號碼A10301835之TC010759,TC010743
            'Modified by Lydia 2017/10/16 改成共用變數
            'ElseIf Text1 = "Y52268000" Or Text1 = "Y53035000" Or Text1 = "Y53541000" Then
            'Modified by Lydia 2025/03/13 改用模組取得
            'ElseIf InStr(外翻Y編號 & ",Y53035000", Text1) > 0 Then
            ElseIf InStr(Pub_SetF51Order("Y", "") & ",Y53035000", Text1) > 0 Then
               '2009/10/20 MODIFY BY SONIA Y52268江蘇舜禹摘要不必帶申請人名稱及收款日,收款金額
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) " & _
                    " values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               '2011/11/23 MODIFY BY SONIA 取消2009/10/20的修改,因為下面程式沒改到,財務說改回來
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) " & _
                    " values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) " & _
                    " values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               'add by sonia 2016/8/18 CFP案件更改入帳科目為應付規費220106
               If Mid(adoquery.Fields("Caseno").Value, 1, 3) = "CFP" Or Mid(adoquery.Fields("Caseno").Value, 1, 3) = "CPS" Then
                  'Modifed by Lydia 2020/10/27 改變數
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               Else
               'end 2016/8/18
                  'modify by sonia 2017/4/14 +a1p30對沖-其他
                  'modify by sonia FCP及FMP案B類收文927其他翻譯且承辦人為外翻編號且相關總收文號為C類之結匯,改借方規費為收入(扣業務點數),例FCP-050279(U10608558)
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                                   " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                  'add by sonia 2017/10/17 FCP及FMP案B類收文927其他翻譯且承辦人為外翻編號且相關總收文號為C類之結匯,改借方規費為收入(扣業務點數),例FCP-050279(U10608558)
                  adoacc190.CursorLocation = adUseClient
                  adoacc190.Open "select cpm24 from acc190,acc151,caseprogress,casepropertymap where a1908='" & Text3 & "' and a1902='" & adoquery.Fields("a1c03") & "' and a1902=axf01(+) and substr(axf02,1,1)='B' " & _
                                 "and axf02=cp09(+) and cp01='FCP' and cp10='927' and substr(cp14,1,1)='F' and substr(cp43,1,1)='C' and cp01=cpm01(+) and cp10=cpm02(+)", adoTaie, adOpenStatic, adLockReadOnly
                  If adoacc190.RecordCount <> 0 Then
                     m_strA1P14 = adoquery.Fields("Caseno").Value & "/OA委外翻譯"
                     'Modifed by Lydia 2020/10/27 改變數
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                     " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & "" & adoacc190.Fields("cpm24") & "', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','F4102')"
                     'modify by sonia 2021/1/20 F4102依案號判別改為F4104,F4105
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                     " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & "" & adoacc190.Fields("cpm24") & "', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','F4102')"
                     'modify by sonia 2021/3/12 加傳日期
                     m_strA1P16 = SalesNoToAccSales("F4102", adoacc190.Fields("cpm24"), adoquery.Fields("Caseno").Value, Val(FCDate(MaskEdBox1.Text)))
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                     " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & "" & adoacc190.Fields("cpm24") & "', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','" & m_strA1P16 & "')"
                     'end 2021/1/20
                  Else
                     'Modifed by Lydia 2020/10/27 改變數
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                                     " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                                     " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                  End If
                  adoacc190.Close
                  'end 2017/10/17
               End If  'add by sonia 2016/8/18
               '2009/10/20 END
            '2008/6/25 end
            '2008/3/21 add by sonia D097030255結匯FCT,故加入FCT,FCL,FCP
            ElseIf Mid(adoquery.Fields("Caseno").Value, 1, 3) = "FCT" Then
               '2012/3/3 modify by sonia 瑞婷說要帶部門及智權人員a1p16
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417201', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & m_strA1P14 & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417201', 'FCT', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", 'F4103') "
               'Modifed by Lydia 2020/10/27 改變數
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417201', 'FCT', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", 'F4103') "
               'modify by sonia 2021/1/20 F4103依案號判別改為F4106,F4107
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417201', 'FCT', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", 'F4103') "
               'modify by sonia 2021/3/12 加傳日期
               m_strA1P16 = SalesNoToAccSales("F4103", "417201", adoquery.Fields("Caseno").Value, Val(FCDate(MaskEdBox1.Text)))
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417201', 'FCT', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & m_strA1P16 & "') "
               'end 2021/1/20
            ElseIf Mid(adoquery.Fields("Caseno").Value, 1, 3) = "FCP" Then
               '2009/4/28 modify by sonia
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '4171', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               '2012/3/3 modify by sonia 瑞婷說要帶部門及智權人員a1p16
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417101', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & m_strA1P14 & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417101', 'FCP', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", 'F4102' )"
               'Modifed by Lydia 2020/10/27 改變數
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417101', 'FCP', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", 'F4102' )"
               'modify by sonia 2021/1/20 F4102依案號判別改為F4104,F4105
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417101', 'FCP', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", 'F4102' )"
               'modify by sonia 2021/3/12 加傳日期
               m_strA1P16 = SalesNoToAccSales("F4102", "417101", adoquery.Fields("Caseno").Value, Val(FCDate(MaskEdBox1.Text)))
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417101', 'FCP', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & m_strA1P16 & "' )"
               'end 2021/1/20
            '2012/3/3 modify by sonia 再加LIN
            ElseIf Mid(adoquery.Fields("Caseno").Value, 1, 3) = "FCL" Or Mid(adoquery.Fields("Caseno").Value, 1, 3) = "LIN" Then
               '2012/3/3 modify by sonia 瑞婷說要帶部門及智權人員a1p16
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '4161', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & m_strA1P14 & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '4161', 'FCL', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", 'F4101' )"
               'modify by sonia 2016/2/17 105年起法務收入改其他部門收入(傳CP09以判斷案件性質及收文人員)
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p16) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '4161', 'FCL', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", 'F4101' )"
               'Modifed by Lydia 2020/10/27 改變數
               'InsertLawACC1P0 IIf(adoquery.Fields("A1917").Value = "J", "J", "1"), "I", strNo, ChgSQL(Text3 & Text1), "4161", "FCL", Val(Format(adoquery.Fields("Amount").Value, DAmount)), 0, "", "", "", "", "", ChgSQL(m_strA1P14), "", "F4101", adoquery.Fields("Caseno").Value, Val(FCDate(MaskEdBox1.Text)), strCurrency, Val(adoquery.Fields("Rate").Value), Val(adoquery.Fields("Famount").Value), IIf(strDocuNo = "null", "", Replace(strDocuNo, "'", "")), adoquery("axf01") & adoquery("axf02"), "", "", "", IIf(strDYes = "null", "", Replace(strDYes, "'", "")), "", "", adoquery("axf02")
               InsertLawACC1P0 strA1917, "I", strNo, ChgSQL(Text3 & Text1), "4161", "FCL", Val(Format(adoquery.Fields("Amount").Value, DAmount)), 0, "", "", "", "", "", ChgSQL(m_strA1P14), "", "F4101", adoquery.Fields("Caseno").Value, Val(FCDate(MaskEdBox1.Text)), strCurrency, Val(adoquery.Fields("Rate").Value), Val(adoquery.Fields("Famount").Value), IIf(strDocuNo = "null", "", Replace(strDocuNo, "'", "")), adoquery("axf01") & adoquery("axf02"), "", "", "", IIf(strDYes = "null", "", Replace(strDYes, "'", "")), "", "", adoquery("axf02")
            Else
            '2008/3/21 END
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & adoquery.Fields("a1907").Value, 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               'Modify by Morgan 2006/3/29 加a1p23
               'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf02")) & ", " & strDYes & ")"
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               'Modifed by Lydia 2020/10/27 改變數
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
            End If
         Else
            If Mid(adoquery.Fields("Caseno").Value, 1, 1) = "T" Then
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & adoquery.Fields("a1907").Value, 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               'Modify by Morgan 2006/3/29 加a1p23
               'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf02")) & ", " & strDYes & ")"
               '2008/6/25 add by sonia 代理人為Y52268江蘇舜禹翻譯有限公司則科目用6130,部門用系統類別
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               '2010/5/4 MODIFY BY SONIA 加 Y53035江蘇通用信息
               '2012/3/2 MODIFY BY SONIA 加 Y53374北京寰華知識產權代理有限公司
               '2012/9/5 MODIFY BY SONIA 加 Y53541南京捷恩凱信息技術有限公司
               '2014/12/30 modify by sonia 取消Y53374匯票號碼A10301835之TC010759,TC010743
               'Modified by Lydia 2017/10/16 改成共用變數
               'If Text1 = "Y52268000" Or Text1 = "Y53035000" Or Text1 = "Y53541000" Then
               'Modified by Lydia 2025/03/13 改用模組取得
               'If InStr(外翻Y編號 & ",Y53035000", Text1) > 0 Then
               If InStr(Pub_SetF51Order("Y", "") & ",Y53035000", Text1) > 0 Then
                  '2014/3/26 modify by sonia 加入J公司
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) " & _
                       " values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  'modify by sonia 2017/4/14 +a1p30對沖-其他
                  'Modifed by Lydia 2020/10/27 改變數
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                       " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                       " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
               Else
                  '2014/3/26 modify by sonia 加入J公司
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  'Modifed by Lydia 2020/10/27 改變數
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
               End If
               '2008/6/25 end
            Else
               If Mid(adoquery.Fields("Caseno").Value, 1, 1) = "S" Then
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & adoquery.Fields("a1907").Value, 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
                  'Modify by Morgan 2006/3/29 加a1p23
                  'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf02")) & ", " & strDYes & ")"
                  '2008/6/25 add by sonia 代理人為Y52268江蘇舜禹翻譯有限公司則科目用6130,部門用系統類別
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  '2010/5/4 MODIFY BY SONIA 加 Y53035江蘇通用信息
                  '2012/3/2 MODIFY BY SONIA 加 Y53374北京寰華知識產權代理有限公司
                  '2012/9/5 MODIFY BY SONIA 加 Y53541南京捷恩凱信息技術有限公司
                  '2014/12/30 modify by sonia 取消Y53374匯票號碼A10301835之TC010759,TC010743
                  'Modified by Lydia 2017/10/16 改成共用變數
                  'If Text1 = "Y52268000" Or Text1 = "Y53035000" Or Text1 = "Y53541000" Then
                  'Modified by Lydia 2025/03/13 改用模組取得
                  'If InStr(外翻Y編號 & ",Y53035000", Text1) > 0 Then
                  If InStr(Pub_SetF51Order("Y", "") & ",Y53035000", Text1) > 0 Then
                     '2014/3/26 modify by sonia 加入J公司
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) " & _
                          " values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     'modify by sonia 2017/4/14 +a1p30對沖-其他
                     'Modifed by Lydia 2020/10/27 改變數
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                          " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                          " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                  Else
                     '2014/3/26 modify by sonia 加入J公司
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     'Modifed by Lydia 2020/10/27 改變數
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  End If
                  '2008/6/25 end
               Else
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & adoquery.Fields("a1907").Value, 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
                  'Modify by Morgan 2006/3/29 加a1p23
                  'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf02")) & ", " & strDYes & ")"
                  '2008/6/25 add by sonia 代理人為Y52268江蘇舜禹翻譯有限公司則科目用6130,部門用系統類別
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  '2010/5/4 MODIFY BY SONIA 加 Y53035江蘇通用信息
                  '2012/3/2 MODIFY BY SONIA 加 Y53374北京寰華知識產權代理有限公司
                  '2012/9/5 MODIFY BY SONIA 加 Y53541南京捷恩凱信息技術有限公司
                  '2012/11/19 modify by sonia 取消Y53374,大部分是台->大案件(匯票號碼A10101595)
                  'Modified by Lydia 2017/10/16 改成共用變數
                  'If Text1 = "Y52268000" Or Text1 = "Y53035000" Or Text1 = "Y53541000" Then
                  'Modified by Lydia 2025/03/13 改用模組取得
                  'If InStr(外翻Y編號 & ",Y53035000", Text1) > 0 Then
                  If InStr(Pub_SetF51Order("Y", "") & ",Y53035000", Text1) > 0 Then
                     '2014/3/26 modify by sonia 加入J公司
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) " & _
                          " values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     'modify by sonia 2017/4/14 +a1p30對沖-其他
                     'modify by sonia FCP及FMP案B類收文927其他翻譯且承辦人為外翻編號且相關總收文號為C類之結匯,改借方規費為收入(扣業務點數),例FCP-050279(U10608558)
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                                      " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                     'add by sonia 2017/10/23 FCP及FMP案B類收文927其他翻譯且承辦人為外翻編號且相關總收文號為C類之結匯,改借方規費為收入(扣業務點數),例FCP-050279(U10608558)
                     adoacc190.CursorLocation = adUseClient
                     adoacc190.Open "select cpm24,cp01,cp02,cp03,cp04 from acc190,acc151,caseprogress,casepropertymap where a1908='" & Text3 & "' and a1902='" & adoquery.Fields("a1c03") & "' and a1902=axf01(+) and substr(axf02,1,1)='B' " & _
                                    "and axf02=cp09(+) and cp01='P' and cp10='927' and substr(cp14,1,1)='F' and substr(cp43,1,1)='C' and cp01=cpm01(+) and cp10=cpm02(+)", adoTaie, adOpenStatic, adLockReadOnly
                     If adoacc190.RecordCount <> 0 Then
                        m_strA1P14 = adoquery.Fields("Caseno").Value & "/OA委外翻譯"
                        If PUB_FMPtoCheck(1, 2, "", adoacc190.Fields("cp01"), adoacc190.Fields("cp02"), adoacc190.Fields("cp03"), adoacc190.Fields("cp04")) = False Then   '非寰華案要依拆20%扣專利處收入
                          'Modifed by Lydia 2020/10/27 改變數
                          'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '411106', '" & m_strA1P06 & "', " & Round(Val(Format(adoquery.Fields("Amount").Value, DAmount)) * 0.2, 0) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) * 0.2 & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','F4102')"
                          'modify by sonia 2021/1/20 F4102依案號判別改為F4104,F4105
                          'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '411106', '" & m_strA1P06 & "', " & Round(Val(Format(adoquery.Fields("Amount").Value, DAmount)) * 0.2, 0) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) * 0.2 & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','F4102')"
                          'modify by sonia 2021/3/12 加傳日期
                          m_strA1P16 = SalesNoToAccSales("F4102", "411106", adoquery.Fields("Caseno").Value, Val(FCDate(MaskEdBox1.Text)))
                          adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '411106', '" & m_strA1P06 & "', " & Round(Val(Format(adoquery.Fields("Amount").Value, DAmount)) * 0.2, 0) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) * 0.2 & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','" & m_strA1P16 & "')"
                          'end 2021/1/21
                          strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
                          'Modifed by Lydia 2020/10/27 改變數
                          'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417102', 'FCP', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) - Round(Val(Format(adoquery.Fields("Amount").Value, DAmount)) * 0.2, 0) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) - (Val(adoquery.Fields("Famount").Value) * 0.2) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','F4102')"
                          'modify by sonia 2021/1/20 F4102依案號判別改為F4104,F4105
                          'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417102', 'FCP', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) - Round(Val(Format(adoquery.Fields("Amount").Value, DAmount)) * 0.2, 0) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) - (Val(adoquery.Fields("Famount").Value) * 0.2) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','F4102')"
                          'modify by sonia 2021/3/12 加傳日期
                          m_strA1P16 = SalesNoToAccSales("F4102", "417102", adoquery.Fields("Caseno").Value, Val(FCDate(MaskEdBox1.Text)))
                          adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417102', 'FCP', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) - Round(Val(Format(adoquery.Fields("Amount").Value, DAmount)) * 0.2, 0) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) - (Val(adoquery.Fields("Famount").Value) * 0.2) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','" & m_strA1P16 & "')"
                          'end 2021/1/21
                        Else
                          'Modifed by Lydia 2020/10/27 改變數
                          'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417102', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','F4102')"
                          'modify by sonia 2021/1/20 F4102依案號判別改為F4104,F4105
                          'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417102', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','F4102')"
                          'modify by sonia 2021/3/12 加傳日期
                          m_strA1P16 = SalesNoToAccSales("F4102", "417102", adoquery.Fields("Caseno").Value, Val(FCDate(MaskEdBox1.Text)))
                          adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30, a1p16) " & _
                                          " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '417102', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "','" & m_strA1P16 & "')"
                          'end 2021/1/21
                        End If
                     Else
                        'Modifed by Lydia 2020/10/27 改變數
                        'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                                        " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                        adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                                        " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                     End If
                     adoacc190.Close
                     'end 2017/10/23
                  '2011/10/17 add by sonia FG-705的結匯D100072982
                  ElseIf Mid(adoquery.Fields("Caseno").Value, 1, 1) = "FG" Then
                     '2014/3/26 modify by sonia 加入J公司
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220104', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     'Modifed by Lydia 2020/10/27 改變數
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220104', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220104', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  '2011/10/17 end
                  '2014/3/4 add by sonia L-005237的結匯
                  ElseIf Mid(adoquery.Fields("Caseno").Value, 1, 1) = "L" Then
                     '2014/3/26 modify by sonia 加入J公司
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     'Modifed by Lydia 2020/10/27 改變數
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                  '2014/3/4 end
                  Else
                     '2014/3/26 modify by sonia 加入J公司
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     'modify by sonia 2017/8/7 P台灣案改用6130 P-116355
                     'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     adoacc190.CursorLocation = adUseClient
                     adoacc190.Open "select pa09 as pa09 from patent where pa01 = '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "' and pa02 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6) & "' and pa03 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 1) & "' and pa04 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 1, 2) & "' union " & _
                                    "select tm10 as pa09 from trademark where tm01 = '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "' and tm02 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6) & "' and tm03 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 1) & "' and tm04 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 1, 2) & "' union " & _
                                    "select lc15 as pa09 from lawcase where lc01 = '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "' and lc02 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6) & "' and lc03 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 1) & "' and lc04 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 1, 2) & "' union " & _
                                    "select '000' as pa09 from hirecase where hc01 = '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "' and hc02 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6) & "' and hc03 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 1) & "' and hc04 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 1, 2) & "' union " & _
                                    "select sp09 as pa09 from servicepractice where sp01 = '" & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "' and sp02 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6) & "' and sp03 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 1) & "' and sp04 = '" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
                     If adoacc190.RecordCount = 0 Then
                        'Modifed by Lydia 2020/10/27 改變數
                        'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                        adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                     Else
                        If adoacc190.Fields("pa09") = "000" Then
                           'Modifed by Lydia 2020/10/27 改變數
                           'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                                   " values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27, a1p30) " & _
                                   " values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & m_strA1P06 & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ", '" & ChgSQL(Text1) & "')"
                        Else
                           'Modifed by Lydia 2020/10/27 改變數
                           'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & CNULL(adoquery("axf01") & adoquery("axf02")) & ", " & strDYes & ")"
                        End If
                     End If
                     adoacc190.Close
                     'end 2017/8/7
                  End If
                  '2008/6/25 end
               End If
            End If
         End If
         'add by sonia 2017/4/17 北京寰華介紹案源(a1803='Y53374000' & a0k34='F5639'之結匯,改借方規費為收入(扣點數)
         adoacc190.CursorLocation = adUseClient
         adoacc190.Open "select axf03,a0k20,sn01,cpm24,cp10,cp09,a0k03,substr(nvl(nvl(cu04,cu05),cu06),6) cu04 from acc190,acc180,acc151,acc0j0,acc0k0,customer,salesno,caseprogress,casepropertymap " & _
                        "where a1908='" & Text3 & "' and a1902='" & adoquery.Fields("a1c03") & "' and a1901=a1801(+) and a1803='Y53374000' and a1902=axf01(+) and axf02=a0j01(+) and a0j13=a0k01(+) and a0k34='F5639' " & _
                        "and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and a0k20=sn02(+) and axf02=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+)", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc190.RecordCount <> 0 Then
            m_strA1P14 = "" & adoacc190.Fields("sn01") & "/" & adoacc190.Fields("cu04") & "/" & "北京寰華介紹案源" & "/" & adoacc190.Fields("axf03")
            'Modifed by Lydia 2020/10/27 改變數
            'adoTaie.Execute "update acc1p0 set a1p05='" & "" & adoacc190.Fields("cpm24") & "',a1p15='" & "" & adoacc190.Fields("a0k03") & "',a1p16='" & "" & adoacc190.Fields("a0k20") & "',a1p17='" & "" & adoacc190.Fields("axf03") & "',a1p14='" & m_strA1P14 & "' where a1p01='" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "' and a1p02='I' and a1p03='" & strNo & "' and a1p04='" & ChgSQL(Text3 & Text1) & "'"
            adoTaie.Execute "update acc1p0 set a1p05='" & "" & adoacc190.Fields("cpm24") & "',a1p15='" & "" & adoacc190.Fields("a0k03") & "',a1p16='" & "" & adoacc190.Fields("a0k20") & "',a1p17='" & "" & strA1917 & "' and a1p02='I' and a1p03='" & strNo & "' and a1p04='" & ChgSQL(Text3 & Text1) & "'"
         End If
         adoacc190.Close
         'end 2017/4/17
      Else
         '2011/11/23 ADD BY SONIA 抓各語法的傳票摘要抓出來改一次即可不必每句改,同時把收款日/收款金額改到摘要的最前面,以便外帳核對資料
         'Modified by Lydia 2020/10/27 分開處理
         'm_strA1P14 = GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value, m_strDomAmt) & "/" & m_strDomAmt & " " & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar)
         ''2011/11/23 END
         m_strA1L02 = GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value, m_strDomAmt)
         m_strA0K04 = GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value)
         '若申請人公司名稱為 株式會社(株式�B社)，請修正傳票摘要公司名；請略過 "株式會社(株式�B社) ，直接抓第五個字起的公司名
         If Left(m_strA0K04, 4) = "株式會社" Or Left(m_strA0K04, 4) = "株式�B社" Then
              m_strA0K04 = Mid(m_strA0K04, 5, 4)
         Else
              m_strA0K04 = Left(m_strA0K04, 4)
         End If
         m_strA1P14 = m_strA1L02 & "/" & m_strDomAmt & " " & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & m_strA0K04 & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar)
         'end 2020/10/27
         
         '2012/2/14 ADD BY SONIA 暫收款退費的摘要改同收到暫收款時的傳票摘要(代理人英文名稱/幣別 外幣金額/暫收單號)或收款溢收款摘要(代理人英文名稱/外幣金額/收款單號/暫收單號)D100110128
         If Mid(adoquery.Fields("a1c03").Value, 1, 1) <> "B" Then
            m_strA1P14 = Text2 & "/"
            If adoquery.Fields("a1210").Value <> "" Then
               '收款溢收款
               m_strA1P14 = m_strA1P14 & Val(adoquery.Fields("Famount").Value) & "/" & adoquery.Fields("a1210").Value & "/" & adoquery.Fields("A1303").Value
            Else
               '收到暫收款
               m_strA1P14 = m_strA1P14 & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "/" & adoquery.Fields("A1303").Value
            End If
         End If
         '2012/2/14 END
         
         If Mid(adoquery.Fields("a1c03").Value, 1, 1) = "B" Then
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & adoquery.Fields("a1907").Value, 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
            '2014/3/26 modify by sonia 加入J公司
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & strDYes & ")"
            'modify by sonia 2015/4/13 配合Frmacc2450國外付款明細表,財務結匯用代理人要發e-mail給副所長秘書,a1p14摘要加放單號a1c03
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & strDYes & ")"
            'Modifed by Lydia 2020/10/27 改變數
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "(單號:" & adoquery.Fields("a1c03").Value & ")', " & strDocuNo & ", " & strDYes & ")"
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "(單號:" & adoquery.Fields("a1c03").Value & ")', " & strDocuNo & ", " & strDYes & ")"
         Else
            '2005/5/4 MODIFY BY SONIA 加A1P30另AMOUNT改為小數二位
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27,A1P30) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '2401', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, FAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & adoquery.Fields("a1907").Value, 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ", '" & adoquery.Fields("A1303").Value & "')"
            '2014/3/26 modify by sonia 加入J公司
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27,A1P30) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '2401', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, FAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & strDYes & ", '" & adoquery.Fields("A1303").Value & "')"
            'Modifed by Lydia 2020/10/27 改變數
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27,A1P30) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '2401', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, FAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & strDYes & ", '" & adoquery.Fields("A1303").Value & "')"
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27,A1P30) values ('" & strA1917 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '2401', '" & MsgText(55) & "', " & Val(Format(adoquery.Fields("Amount").Value, FAmount)) & ", '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", 0, '" & ChgSQL(m_strA1P14) & "', " & strDocuNo & ", " & strDYes & ", '" & adoquery.Fields("A1303").Value & "')"
         End If
      End If
      'Remove by Morgan 2007/8/23 移到新增acc1c0處
      'adoTaie.Execute "update acc150 set a1520 = NVL(a1520,0)+" & Val(Format(adoquery.Fields("Famount").Value, FAmount)) & " where a1501 = '" & adoquery.Fields("a1c03").Value & "'"
      
      If IsNull(adoquery.Fields("Rate").Value) Then
         douRate = 1
      Else
         douRate = adoquery.Fields("Rate").Value
      End If
      'Modified by Lydia 2020/10/27
      'strA1P01 = IIf(adoquery.Fields("A1917").Value = "J", "J", "1") '2014/3/26 ADD BY SONIA
      strA1P01 = strA1917
      strAutoGen = strA1P01                        'ADD BY SONIA 2014/6/18
      adoquery.MoveNext
   Loop
   adoquery.Close
   If strA1P01 = "" Then strA1P01 = strCreditGen   'ADD BY SONIA 2014/6/18
   
   Adodc3.Recordset.Requery
' Ken 92/07/31 所有幣別都將手續費併入第一筆之應付規費
'   If Mid(strCurrency, 1, 2) <> "US" And strCurrency <> "EUR" Then
      If Adodc3.Recordset.RecordCount <> 0 Then
         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select sum(a1p21), sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' and a1p07 > 0", adoTaie, adOpenStatic, adLockReadOnly
         '2014/3/26 modify by sonia 取消a1p01
         'adoquery.Open "select sum(a1p21), sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and a1p07 > 0", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select sum(a1p21), sum(a1p07) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and a1p07 > 0", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               douAmount = 0
            Else
               'Memo by Lydia 2021/07/26 抓匯差
               douAmount = Val(Format(adoquery.Fields(0).Value * douRate, DAmount)) - adoquery.Fields(1).Value
            End If
         Else
            douAmount = 0
         End If
         'Added by Lydia 2021/07/26 平均分攤手續費的金額
         m_Amt1 = 0: m_TotAmt = 0
         Dim m_GrpNo As String 'Added by Lydia 2022/03/04
         m_GrpNo = ""          'add by sonia 2025/7/24
         If Val(Text4) <> 0 Then
            adoquery.Close
            adoquery.CursorLocation = adUseClient
            'Modified by Lydia 2022/03/04 借方科目"6120"也要分攤; 因為不會有2201 , 6120同時出現的狀況, 所以個別分攤即可(先分2201XX) ; ex.匯票號碼A11100397+代理人Y00007000
            'adoquery.Open "select count(*) cnt from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and substr(a1p05,1,4)='2201' and a1p07 > 0", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2025/8/5 6130翻譯費也要分攤
            strSql = "select substr(a1p05,1,4) grpno, count(*) cnt from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and (substr(a1p05,1,4)='2201' or substr(a1p05,1,4)='6120' or substr(a1p05,1,4)='6130') and a1p07 > 0 group by substr(a1p05,1,4) order by substr(a1p05,1,4) "
            adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
            'Modifoed by Lydia 2022/03/15 debug
            'If Val("" & adoquery.Fields("cnt")) > 0 Then
            StrSQLa = ""
            If adoquery.RecordCount > 0 Then
                StrSQLa = "" & adoquery.Fields("cnt")
            End If
            If Val(StrSQLa) > 0 Then
            'end 2022/03/15
                m_GrpNo = "" & adoquery.Fields("grpno") 'Added by Lydia 2022/03/04
                m_Amt1 = Format(Val(Text4) / adoquery.Fields("cnt"), DAmount)
                '手續費(四捨五入進位到個位數)分攤到2201開頭且借方金額>0的所有項次，差額(不管增還是減)放在2201開頭且借方金額>0的最後一項
                adoquery.Close
                adoquery.CursorLocation = adUseClient
                'Modified by Lydia 2022/03/04 借方科目"6120"也要分攤
                'adoquery.Open "select * from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and substr(a1p05,1,4)='2201' and a1p07 > 0 order by a1p03 ", adoTaie, adOpenStatic, adLockReadOnly
                adoquery.Open "select * from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and substr(a1p05,1,4)='" & m_GrpNo & "' and a1p07 > 0 order by a1p03 ", adoTaie, adOpenStatic, adLockReadOnly
                If adoquery.RecordCount <> 0 Then
                   adoquery.MoveFirst
                   Do While Not adoquery.EOF
                       If adoquery.AbsolutePosition = adoquery.RecordCount Then
                           StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & (Val(Text4) - m_TotAmt) & " where a1p02 = 'I' and a1p03 ='" & adoquery.Fields("a1p03") & "' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
                       Else
                           StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & Val(m_Amt1) & " where a1p02 = 'I' and a1p03 ='" & adoquery.Fields("a1p03") & "' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
                           m_TotAmt = m_TotAmt + m_Amt1
                       End If
                       adoTaie.Execute StrSQLa, lngEff
                       adoquery.MoveNext
                   Loop
                End If
            Else  '無2201XX科目時更新借方最大項次
                'Memo by Lydia 2022/03/04 同時沒有借方科目"6120"; ex.匯票號碼A11100397+代理人Y00007000
                StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & Val(Text4) & " where a1p02 = 'I' and a1p03 = (select max(a1p03) from acc1p0 where a1p02 = 'I' AND SUBSTR(A1P05,1,4)<>'2401' and a1p07 > 0 and a1p04 = '" & ChgSQL(Text3 & Text1) & "') and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
                adoTaie.Execute StrSQLa, lngEff
            End If
         End If
         'end 2021/07/26
         adoquery.Close
          '2005/5/17 MODIFY BY SONIA 只更新 2201XX 的科目
          '2005/10/26 MODIFY BY SONIA 無2201XX科目時更新借方最大項次
          'adoTaie.Execute "update acc1p0 set a1p07 = a1p07 + " & Val(Text4) + douAmount & " where a1p01 = '1' and a1p02 = 'I' and a1p03 = (select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' AND SUBSTR(A1P05,1,4)='2201' and a1p07 > 0 and a1p04 = '" & ChgSQL(Text3 & Text1) & "') and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
         adoquery.CursorLocation = adUseClient
         '2014/3/26 modify by sonia 取消a1p01
         'StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & Val(Text4) + douAmount & " where a1p01 = '1' and a1p02 = 'I' and a1p03 = (select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' AND SUBSTR(A1P05,1,4)='2201' and a1p07 > 0 and a1p04 = '" & ChgSQL(Text3 & Text1) & "') and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
         If douAmount <> 0 Then 'Added by Lydia 2021/07/26 增加判斷
            'Modified by Lydia 2021/07/26 匯差放在「2201開頭且借方金額>0的最後一項」,前面已有平均分攤手續費=>拿掉Val(Text4)
            'Modified by Lydia 2022/03/04 借方科目"6120"也要分攤
            'StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & douAmount & " where a1p02 = 'I' and a1p03 = (select max(a1p03) from acc1p0 where a1p02 = 'I' AND SUBSTR(A1P05,1,4)='2201' and a1p07 > 0 and a1p04 = '" & ChgSQL(Text3 & Text1) & "') and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
            StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & douAmount & " where a1p02 = 'I' and a1p03 = (select max(a1p03) from acc1p0 where a1p02 = 'I' AND SUBSTR(A1P05,1,4)='" & m_GrpNo & "' and a1p07 > 0 and a1p04 = '" & ChgSQL(Text3 & Text1) & "') and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
            adoTaie.Execute StrSQLa, lngEff
            If lngEff = 0 Then   'Memo by Lydia 2021/07/26 無2201XX科目時更新借方最大項次 'Memo by Lydia 2022/03/04 同時沒有借方科目"6120"
               '2006/3/21 MODIFY BY SONIA 但暫收款退費不能更新
               'StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & Val(Text4) + douAmount & " where a1p01 = '1' and a1p02 = 'I' and a1p03 = (select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p07 > 0 and a1p04 = '" & ChgSQL(Text3 & Text1) & "') and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
               '2014/3/26 modify by sonia 取消a1p01
               'StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & Val(Text4) + douAmount & " where a1p01 = '1' and a1p02 = 'I' and a1p03 = (select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' AND SUBSTR(A1P05,1,4)<>'2401' and a1p07 > 0 and a1p04 = '" & ChgSQL(Text3 & Text1) & "') and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
               'Modified by Lydia 2021/07/26 前面已有平均分攤手續費=>拿掉Val(Text4)
               StrSQLa = "update acc1p0 set a1p07 = a1p07 + " & douAmount & " where a1p02 = 'I' and a1p03 = (select max(a1p03) from acc1p0 where a1p02 = 'I' AND SUBSTR(A1P05,1,4)<>'2401' and a1p07 > 0 and a1p04 = '" & ChgSQL(Text3 & Text1) & "') and a1p04 = '" & ChgSQL(Text3 & Text1) & "'"
               '2006/3/21 END
               adoTaie.Execute StrSQLa, lngEff
            End If
         End If 'Added by Lydia 2021/07/26 增加判斷
         'add by sonia 2025/7/24 2201開頭科目要重算匯率並更新至A1906，以利後續國外案件帳目查詢時能帶出較接近結匯傳票的金額
         'modify by sonia 2025/8/5 6130翻譯費也要更新
         'StrSQLa = "update acc190 set a1906 =(Select Sum(A1p07)/Sum(A1p21) From Acc1p0,Acc180 where a1901=a1801 and a1908||a1803=a1p04 And A1p05 Like '2201%' And A1p07>0 and a1902=substr(a1p23,1,9) ) " & _
                   "where a1902 in (select Distinct Substr(A1p23,1,9) From Acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and SUBSTR(A1P05,1,4)='2201' and a1p07 > 0 )"
         StrSQLa = "update acc190 set a1906 =(Select Sum(A1p07)/Sum(A1p21) From Acc1p0,Acc180 where a1901=a1801 and a1908||a1803=a1p04 And SUBSTR(A1P05,1,4) in ('2201','6130') And A1p07>0 and a1902=substr(a1p23,1,9) ) " & _
                   "where a1902 in (select Distinct Substr(A1p23,1,9) From Acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and SUBSTR(A1P05,1,4) in ('2201','6130') and a1p07 > 0 )"
         adoTaie.Execute StrSQLa, lngEff
         'end 2025/7/24
         
      End If
'   End If
   CreditGen
   Adodc3.Recordset.Requery
   SumShow
End Sub

'*************************************************
'  自動產生貸方科目
'
'*************************************************
Public Sub CreditGen()
Dim douRate As Double
Dim strBankAccNo As String
Dim strBankNo As String
Dim douAmount As Double
Dim strAccNo As String
Dim douFAmount As Double
Dim StrSQLa As String
Dim strA1306 As String    '2010/6/23 ADD BY SONIA 若為台幣暫收款退費時手續費不可扣除

   strA1306 = MsgText(601) '2010/6/23 ADD BY SONIA
   strSerialNo = MsgText(601)
'   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' and a1p08 <> 0"
   '2014/3/26 modify by sonia 取消a1p01
   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and a1p08 <> 0"
   adoTaie.Execute "delete from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and a1p08 <> 0"
   adoquery.CursorLocation = adUseClient
'   strSQLA = "select a1c03, axg03 as Caseno, round(axg04 * a1906) as Amount, '2201' as Accno, axg04 as Famount, a1906 as Rate from acc161, acc1c0, acc190 where axg01 = a1c03 and axg01 = a1902 and a1c01 = '" & Text3 & "' and a1c02 = '" & Text1 & "' order by axg03 asc"
   '2007/11/5 modify by sonia 加cp10
   'strSQLA = "select a1c03, axg03 as Caseno, round(axg04 * a1906) as Amount, '2201' as Accno, axg04 as Famount, a1906 as Rate,axg02,axg01 from acc161, acc1c0, acc190 where axg01 = a1c03 and axg01 = a1902 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' order by axg03 asc"
   '2014/6/18 MODIFY BY SONIA 加A1917讀取收據公司別 A10300670
   StrSQLa = "select a1c03, axg03 as Caseno, NVL(round(axg04 * a1906), 0) as Amount, '2201' as Accno, axg04 as Famount, NVL(a1906, 0) as Rate,axg02,axg01,cp10,A1917 from acc161, acc1c0, acc190, caseprogress where axg01 = a1c03 and axg01 = a1902 and axg02=cp09(+) and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' order by axg03 asc"
   adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   
   strA1P01 = "": strCreditGen = "" 'ADD BY SONIA 2014/6/18
   Do While adoquery.EOF = False
'      strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", 3)
      '2014/3/26 modify by sonia 取消a1p01
      'strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
      strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
      strA1P01 = IIf(adoquery.Fields("A1917").Value = "J", "J", IIf(adoquery.Fields("A1917").Value = "L", "L", "1")) 'Added by Lydia 2020/10/27 轉傳票公司別
      If Mid(adoquery.Fields("a1c03").Value, 1, 1) = MsgText(813) Then
         If Len(adoquery.Fields("Caseno").Value) = 12 Then
            '2005/9/22 MODIFY BY SONIA CFP抵帳單之摘要帶 退公開費
            'Modify by Morgan 2006/4/6 加a1p23
            'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
            '2007/11/5 modify by sonia CFP抵帳單須為領證或公開費且金額大於us100則才加註 退公開費
            'If Mid(adoquery.Fields("Caseno").Value, 1, 3) = "CFT" Or Mid(adoquery.Fields("Caseno").Value, 1, 3) = "CFC" Then
            '   adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNO & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
            'Else
            '   adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNO & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "/退公開費', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
            'End If
            If Mid(adoquery.Fields("Caseno").Value, 1, 3) = "CFP" Then
               If (adoquery.Fields("cp10").Value = "601" Or adoquery.Fields("cp10").Value = "217") And adoquery.Fields("FAMOUNT").Value > 100 Then
                  '2014/3/26 modify by sonia 加入J公司
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "/退公開費', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
                  'Modifed by Lydia 2020/10/27 改變數
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "/退公開費', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "/退公開費', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               Else
                  '2014/3/26 modify by sonia 加入J公司
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
                  'Modifed by Lydia 2020/10/27 改變數
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220106', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               End If
            Else
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               'Modifed by Lydia 2020/10/27 改變數
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               'modify by sonia 2025/10/14 翻譯社改用6130科目D114031390之FCP072720000、FCP072787000、FCP072788000
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               If InStr(Pub_SetF51Order("Y", "") & ",Y53035000", Text1) > 0 Then
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6130', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               Else
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               End If
               'end 2025/10/14
            End If
            '2007/11/5 end
         Else
            If Mid(adoquery.Fields("Caseno").Value, 1, 1) = "T" Then
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '220111', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               'Modify by Morgan 2006/4/6 加a1p23
               'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg02")) & ", " & strDYes & ")"
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               'Modifed by Lydia 2020/10/27 改變數
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220111', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
            Else
               If adoquery.Fields("Caseno").Value = "S" Then
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
                  'Modify by Morgan 2006/4/6 加a1p23
                  'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg02")) & ", " & strDYes & ")"
                  '2014/3/26 modify by sonia 加入J公司
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
                  'Modifed by Lydia 2020/10/27 改變數
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220105', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               Else
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '220112', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
                  'Modify by Morgan 2006/4/6 加a1p23
                  'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg02")) & ", " & strDYes & ")"
                  '2014/3/26 modify by sonia 加入J公司
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
                  'Modifed by Lydia 2020/10/27 改變數
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '220112', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
               End If
            End If
         End If
      Else
         If Mid(adoquery.Fields("a1c03").Value, 1, 1) = "B" Then
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '6120', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
            'Modify by Morgan 2006/4/6 加a1p23
            'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg02")) & ", " & strDYes & ")"
            '2014/3/26 modify by sonia 加入J公司
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
            'Modifed by Lydia 2020/10/27 改變數
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '6120', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
         Else
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '2401', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
            'Modify by Morgan 2006/4/6 加a1p23
            'Modify by Morgan 2006/4/19 a1p23改放"帳單號+收文號"
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '2401', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg02")) & ", " & strDYes & ")"
            '2014/3/26 modify by sonia 加入J公司
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '2401', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
            'Modifed by Lydia 2020/10/27 改變數
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & IIf(adoquery.Fields("A1917").Value = "J", "J", "1") & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '2401', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p23, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '2401', '" & MsgText(55) & "', 0, '" & adoquery.Fields("Caseno").Value & "', " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & Val(adoquery.Fields("Rate").Value) & ", " & Val(adoquery.Fields("Famount").Value) & ", " & Val(Format(adoquery.Fields("Amount").Value, DAmount)) & ", '" & adoquery.Fields("Caseno").Value & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar) & "', " & strDocuNo & ", " & CNULL(adoquery("axg01") & adoquery("axg02")) & ", " & strDYes & ")"
         End If
      End If
      If IsNull(adoquery.Fields("Rate").Value) Then
         douRate = 1
      Else
         douRate = adoquery.Fields("Rate").Value
      End If
      'strA1P01 = IIf(adoquery.Fields("A1917").Value = "J", "J", "1") 'ADD BY SONIA 2014/6/18 A10300670 'Remove by Lydia 2020/10/27
      strCreditGen = strA1P01                  'ADD BY SONIA 2014/6/18
      adoquery.MoveNext
   Loop
   If strA1P01 = "" Then strA1P01 = strAutoGen 'ADD BY SONIA 2014/6/18
   
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1x0 where a1x01 = '" & strCurrency & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("a1x02").Value) Then
         douRate = 1
      Else
         douRate = Val(adoquery.Fields("a1x02").Value)
      End If
   Else
      douRate = 1
   End If
   adoquery.Close
   '2010/6/23 ADD BY SONIA 判斷是否為台幣暫收款退費
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a1306 from acc1c0, acc130 where a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "' AND a1c03=a1301 order by a1c03 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If "" & adoquery.Fields("A1306").Value = "NTD" Then
         strA1306 = "Y"
      End If
   End If
   adoquery.Close
   '2010/6/23 END
   adoquery.CursorLocation = adUseClient
   'add by sonia 2014/3/26 J公司用110304
   If strA1P01 = "J" Then
      adoquery.Open "select a0h01, a0h02 from acc0h0, acc0g0 where a0h01 = a0g01 and a0h08 = '110304'", adoTaie, adOpenStatic, adLockReadOnly
   Else
   '2014/3/26 end
      Select Case strCurrency
         Case "USD"
            'modify by sonia 2014/8/7 1公司獨立水單以新臺幣結購的科目固定用110204
            'adoquery.Open "select a0h01, a0h02 from acc0h0, acc0g0 where a0h01 = a0g01 and a0h08 = '110205'", adoTaie, adOpenStatic, adLockReadOnly
            If strA1812 = "Y" Then
               adoquery.Open "select a0h01, a0h02 from acc0h0, acc0g0 where a0h01 = a0g01 and a0h08 = '110204'", adoTaie, adOpenStatic, adLockReadOnly
            Else
               'modify by sonia 2017/6/16 1公司美金改用110228科目
               'adoquery.Open "select a0h01, a0h02 from acc0h0, acc0g0 where a0h01 = a0g01 and a0h08 = '110205'", adoTaie, adOpenStatic, adLockReadOnly
               'Modified by Lydia 2020/09/11 已不用1公司(商標)美金帳戶110228
               'adoquery.Open "select a0h01, a0h02 from acc0h0, acc0g0 where a0h01 = a0g01 and a0h08 = '" & IIf(strA1917 = "1", "110228", "110205") & "'", adoTaie, adOpenStatic, adLockReadOnly
               adoquery.Open "select a0h01, a0h02 from acc0h0, acc0g0 where a0h01 = a0g01 and a0h08 = '110205'", adoTaie, adOpenStatic, adLockReadOnly
               'end 2017/6/16
            End If
            'end 2014/8/7
         '2010/4/29 CANCEL BY SONIA 婧瑄因無歐元存款故取消
         'Case "EUR"
         '   adoquery.Open "select a0h01, a0h02 from acc0h0, acc0g0 where a0h01 = a0g01 and a0h08 = '110222'", adoTaie, adOpenStatic, adLockReadOnly
         '2010/4/29 END
         Case Else
            adoquery.Open "select a0h01, a0h02 from acc0h0, acc0g0 where a0h01 = a0g01 and a0h08 = '110204'", adoTaie, adOpenStatic, adLockReadOnly
      End Select
   End If
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("a0h02").Value) Then
         strBankAccNo = MsgText(601)
      Else
         strBankAccNo = adoquery.Fields("a0h02").Value
      End If
      If IsNull(adoquery.Fields("a0h01").Value) Then
         strBankNo = MsgText(601)
      Else
         strBankNo = adoquery.Fields("a0h01").Value
      End If
   Else
      strBankAccNo = MsgText(601)
      strBankNo = MsgText(601)
   End If
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   '2010/5/19 MODIFY BY SONIA 婧瑄因無歐元存款故取消
   'If Mid(strCurrency, 1, 2) = "US" Or strCurrency = "EUR" Then
   If Mid(strCurrency, 1, 2) = "US" Then
'       adoquery.Open "select nvl(sum(a1904), 0) from acc190, acc1c0 where a1902 = a1c03 and a1c01 = '" & Text3 & "' and a1c02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
       adoquery.Open "select nvl(sum(a1904), 0) from acc190, acc1c0 where a1902 = a1c03 and a1c01 = '" & Text3 & "' and a1c02 = '" & ChgSQL(Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   Else
'       adoquery.Open "select nvl(sum(a1p07 - a1p08), 0), nvl(sum(decode(a1p07, 0, a1p21*(-1), a1p21)), 0) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
       '2014/3/26 modify by sonia 取消a1p01
       'adoquery.Open "select nvl(sum(a1p07 - a1p08), 0), nvl(sum(decode(a1p07, 0, a1p21*(-1), a1p21)), 0) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
       adoquery.Open "select nvl(sum(a1p07 - a1p08), 0), nvl(sum(decode(a1p07, 0, a1p21*(-1), a1p21)), 0) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   End If
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         douAmount = 0
      Else
         douAmount = Val(adoquery.Fields(0).Value)
      End If
      '2010/5/19 MODIFY BY SONIA 婧瑄因無歐元存款故取消
      'If Mid(strCurrency, 1, 2) <> "US" And strCurrency <> "EUR" Then
      If Mid(strCurrency, 1, 2) <> "US" Then
         If IsNull(adoquery.Fields(1).Value) Then
            douFAmount = 0
         Else
            douFAmount = Val(adoquery.Fields(1).Value)
         End If
      End If
   Else
      douAmount = 0
      douFAmount = 0
   End If
   adoquery.Close
'   strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", 3)
   '2014/3/26 modify by sonia 取消a1p01
   'strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
   strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
   '2010/5/19 MODIFY BY SONIA 婧瑄因無歐元存款故取消
   'If Mid(strCurrency, 1, 2) = "US" Or strCurrency = "EUR" Then
   If Mid(strCurrency, 1, 2) = "US" Then
      '92.6.13 CANCEL BY SONIA
      ''Ken 92/06/05 加入商標案一律帶110204科目
      'adoquery.CursorLocation = adUseClient
      'adoquery.Open "select a1p17 from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and substr(a1p17, 1, 3) in ('CFT', 'FCT')", adoTaie, adOpenStatic, adLockReadOnly
      'If adoquery.RecordCount <> 0 Then
      '   strAccNo = "110204"
      'Else
      'END 92.6.13
      Select Case strCurrency
         Case "USD"
            'modify by sonia 2014/8/7 1公司獨立水單以新臺幣結購的科目固定用110204
            'strAccNo = "110205"
            If strA1812 = "Y" Then
               strAccNo = "110204"
            'add by sonia 2017/6/16 1公司美金改用110228科目
            'Remove by Lydia 2020/09/11 已不用1公司(商標)美金帳戶110228
            'ElseIf strA1917 = "1" Then
            '   strAccNo = "110228"
            'end 2020/09/11
            Else
               strAccNo = "110205"
            End If
            'END 2014/8/7
         '2010/4/29 CANCEL BY SONIA 婧瑄因無歐元存款故取消
         'Case "EUR"
         '   strAccNo = "110222"
         '2010/4/29 END
         Case Else
            'modify by sonia 2017/1/17 發現錯誤,上面為110204,此處不知為何寫110205
            strAccNo = "110204"
      End Select
      If Mid(Combo2, 1, 1) = "5" Then
         strAccNo = "110207"
      End If
      'add by sonia 2014/3/26 J公司用110304
      If strA1P01 = "J" Then
         strAccNo = "110304"
         
         'Added by Morgan 2022/8/23 J公司商務卡用110303瑞興銀行長安乙存(智權)
         If Mid(Combo2, 1, 1) = "5" Then
            strAccNo = "110303"
         End If
         'end 2022/8/23
      End If
      '2014/3/26 end
      
      'End If
      'adoquery.Close  '92.6.13 CANCEL BY SONIA
'      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount * douRate, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & Text1 & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
      '2012/4/24 MODIFY BY SONIA 摘要加代理人編號
      'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount * douRate, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
      '2012/5/30 MODIFY BY SONIA 摘要代理人編號改位置
      'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount * douRate, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douAmount, FDollar) & "/" & Text1 & "', " & strDocuNo & ", " & strDYes & ")"
      '2014/3/26 modify by sonia 加入J公司
      'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount * douRate, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
      'modify by sonia 2014/11/5 智權公司不列手續費科目併入結匯科目110304也不會有匯差故抓借貸差額,傳票D103090070
      'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount * douRate, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
      If strA1P01 = "J" Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select sum(a1p07)-sum(a1p08) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               douLAmount = 0
            Else
               douLAmount = Val(adoquery.Fields(0).Value)
            End If
         Else
            douLAmount = 0
         End If
         adoquery.Close
         'modify by sonia 2018/7/19 商務卡結匯,摘要的"結匯"二字請修改成"商務卡"(商務卡不會有手續費)
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(douLAmount) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(douLAmount) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & IIf(Mid(Combo2, 1, 1) = "5", "商務卡", MsgText(127)) & "/" & Text1 & "/" & strCurrency & " " & Format(douAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
      Else
         'modify by sonia 2018/7/19 商務卡結匯,摘要的"結匯"二字請修改成"商務卡"(商務卡不會有手續費)
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount * douRate, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount * douRate, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & IIf(Mid(Combo2, 1, 1) = "5", "商務卡", MsgText(127)) & "/" & Text1 & "/" & strCurrency & " " & Format(douAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
      End If
      'end 2014/11/5
      '2012/4/24 END
      If Val(Text4) > 0 Then
'         strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", 3)
         '2014/3/26 modify by sonia 取消a1p01
         'strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
         strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
'         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '110204', '" & MsgText(55) & "', 0, " & Val(Text4) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & Text1 & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/手續費" & "', " & strDocuNo & ", " & strDYes & ")"
         '2012/4/24 MODIFY BY SONIA 摘要加代理人編號
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110204', '" & MsgText(55) & "', 0, " & Val(Text4) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/手續費" & "', " & strDocuNo & ", " & strDYes & ")"
         '2014/3/26 modify by sonia 加入J公司且銀存科目110304
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110204', '" & MsgText(55) & "', 0, " & Val(Text4) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/手續費/" & strCurrency & "', " & strDocuNo & ", " & strDYes & ")"
         If strA1P01 = "J" Then
            'cancel by sonia 2014/11/5 智權公司不列手續費科目,傳票D103090070
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110304', '" & MsgText(55) & "', 0, " & Val(Text4) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/手續費/" & strCurrency & "', " & strDocuNo & ", " & strDYes & ")"
         Else
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110204', '" & MsgText(55) & "', 0, " & Val(Text4) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/手續費/" & strCurrency & "', " & strDocuNo & ", " & strDYes & ")"
         End If
         '2014/3/26 end
         '2012/4/24 END
      End If
   Else
      '2010/6/17 ADD BY SONIA 非美金之暫收款退費有手續費時Y45814010匯票T00128000
      If Val(Text4) > 0 And strA1306 = "Y" Then
         douAmount = douAmount + Val(Text4)
         douFAmount = douFAmount + Val(Text4)
      End If
      '2010/6/17 END
      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      '2014/3/26 modify by sonia 取消a1p01
      'adoquery.Open "select sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select sum(a1p07) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         
         '2014/3/26 add by sonia J公司銀存科目110304
         If strA1P01 = "J" Then
            'Modified by Morgan 2022/8/23 非美金也會用商務卡,J公司商務卡用110303瑞興銀行長安乙存(智權)--婉莘
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('J', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110304', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, DAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('J', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & IIf(Mid(Combo2, 1, 1) = "5", "110303", "110304") & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, DAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douFAmount & ", " & Val(Text4) & ", '" & IIf(Mid(Combo2, 1, 1) = "5", "商務卡", MsgText(127)) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
         Else
         '2014/3/26 end
      
            If IsNull(adoquery.Fields(0).Value) Then
               'douAmount = 0
   '            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & Text1 & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               '2012/4/24 MODIFY BY SONIA 摘要加代理人編號
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               '2012/5/30 MODIFY BY SONIA 摘要代理人編號改位置
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "/" & Text1 & "', " & strDocuNo & ", " & strDYes & ")"
               'Modifed by Lydia 2020/11/05 公司別改變數
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               ''2012/4/24 END
               'Modified by Morgan 2022/8/23 非美金也會用商務卡--婉莘
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & IIf(Mid(Combo2, 1, 1) = "5", "110207", "110205") & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & IIf(Mid(Combo2, 1, 1) = "5", "商務卡", MsgText(127)) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               
            Else
               'douAmount = Val(adoquery.Fields(0).Value)
   '            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '110204', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, DAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & Text1 & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               '2012/4/24 MODIFY BY SONIA 摘要加代理人編號
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110204', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, DAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               '2012/5/30 MODIFY BY SONIA 摘要代理人編號改位置
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110204', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, DAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "/" & Text1 & "', " & strDocuNo & ", " & strDYes & ")"
               'Modifed by Lydia 2020/11/05 公司別改變數
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110204', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, DAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               ''2012/4/24 END
               'Modified by Morgan 2022/8/23 非美金也會用商務卡--婉莘
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110204', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, DAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & IIf(Mid(Combo2, 1, 1) = "5", "110207", "110204") & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, DAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douFAmount & ", " & Val(Text4) & ", '" & IIf(Mid(Combo2, 1, 1) = "5", "商務卡", MsgText(127)) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
            End If
         End If
      Else
         'douAmount = 0
'         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & Text1 & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
         '2012/4/24 MODIFY BY SONIA 摘要加代理人編號
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
         '2012/5/30 MODIFY BY SONIA 摘要代理人編號改位置
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "/" & Text1 & "', " & strDocuNo & ", " & strDYes & ")"
         '2014/3/26 add by sonia J公司銀存科目110304
         If strA1P01 = "J" Then
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('J', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110304', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
         Else
         '2014/3/26 end
            'Modifed by Lydia 2020/11/05 公司別改變數
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
            'Modified by Morgan 2022/8/23 非美金也會用商務卡--婉莘
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '110205', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & MsgText(127) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p25, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '" & IIf(Mid(Combo2, 1, 1) = "5", "110207", "110205") & "', '" & MsgText(55) & "', 0, " & Val(Format(douAmount, FAmount)) & ", '" & strBankNo & "', '" & strBankAccNo & "', '" & ChgSQL(Text1) & "', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & douFAmount & ", " & Val(Text4) & ", '" & IIf(Mid(Combo2, 1, 1) = "5", "商務卡", MsgText(127)) & "/" & Text1 & "/" & strCurrency & " " & Format(douFAmount, FDollar) & "', " & strDocuNo & ", " & strDYes & ")"
         End If
         '2012/4/24 END
      End If
      adoquery.Close
      '2010/6/17 ADD BY SONIA 非美金之暫收款退費有手續費時Y45814010匯票T00128000
      If Val(Text4) > 0 And strA1306 = "Y" Then
         '2014/3/26 modify by sonia 取消a1p01
         'strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
         strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
         '2014/3/26 modify by sonia 加入J公司 取消a1p01
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '611301', '" & MsgText(55) & "', " & Val(Text4) & ", 0, '" & strBankNo & "', '" & strBankAccNo & "', '', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & Val(Text4) & ", '" & "手續費" & "', " & strDocuNo & ", " & strDYes & ")"
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p10, a1p11, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '611301', '" & MsgText(55) & "', " & Val(Text4) & ", 0, '" & strBankNo & "', '" & strBankAccNo & "', '', null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', " & douRate & ", " & Val(Text4) & ", '" & "手續費" & "', " & strDocuNo & ", " & strDYes & ")"
         '2014/3/26 end
      End If
      '2010/6/17 END
   End If
   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select sum(a1p07), sum(a1p08) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   '2014/3/26 modify by sonia 取消a1p01
   'adoquery.Open "select sum(a1p07), sum(a1p08) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select sum(a1p07), sum(a1p08) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         douLAmount = 0
      Else
         douLAmount = Val(adoquery.Fields(0).Value)
      End If
      If IsNull(adoquery.Fields(1).Value) Then
         douTAmount = 0
      Else
         douTAmount = Val(adoquery.Fields(1).Value)
      End If
   Else
      douLAmount = 0
      douTAmount = 0
   End If
   adoquery.Close
'2010/5/19 CANCEL BY SONIA因上下並無差異故調整
'   Select Case strCurrency
'      Case "USD", "EUR"
''         strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", 3)
'         strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
'         If douLAmount > douTAmount Then
''            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p08, a1p17, a1p18, a1p19, a1p20, a1p21, a1p07, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '7128', '" & MsgText(55) & "', " & douLAmount - douTAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p08, a1p17, a1p18, a1p19, a1p20, a1p21, a1p07, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '7128', '" & MsgText(55) & "', " & douLAmount - douTAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
'         Else
'            If douLAmount < douTAmount Then
''               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '7128', '" & MsgText(55) & "', " & douTAmount - douLAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '7128', '" & MsgText(55) & "', " & douTAmount - douLAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
'            End If
'         End If
'      Case Else
'         strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", 3)
         '2014/3/26 modify by sonia 取消a1p01
         'strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
         strNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "'", 3)
         If douLAmount > douTAmount Then
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p08, a1p17, a1p18, a1p19, a1p20, a1p21, a1p07, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '7128', '" & MsgText(55) & "', " & douLAmount - douTAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
            '2014/3/26 modify by sonia 加入J公司
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p08, a1p17, a1p18, a1p19, a1p20, a1p21, a1p07, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '7128', '" & MsgText(55) & "', " & douLAmount - douTAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p08, a1p17, a1p18, a1p19, a1p20, a1p21, a1p07, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '7128', '" & MsgText(55) & "', " & douLAmount - douTAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
         Else
            If douLAmount < douTAmount Then
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & Text3 & Text1 & "', '7128', '" & MsgText(55) & "', " & douTAmount - douLAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
               '2014/3/26 modify by sonia 加入J公司
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('1', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '7128', '" & MsgText(55) & "', " & douTAmount - douLAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p17, a1p18, a1p19, a1p20, a1p21, a1p08, a1p14, a1p22, a1p27) values ('" & strA1P01 & "', 'I', '" & strNo & "', '" & ChgSQL(Text3 & Text1) & "', '7128', '" & MsgText(55) & "', " & douTAmount - douLAmount & ", null, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strCurrency & "', 1, " & douTAmount - douLAmount & ", 0, '" & MsgText(127) & "', " & strDocuNo & ", " & strDYes & ")"
            End If
         End If
'   End Select
'   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' and a1p07 = 0 and a1p08 = 0"
   '2014/3/26 modify by sonia 取消a1p01
   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and a1p07 = 0 and a1p08 = 0"
   adoTaie.Execute "delete from acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and a1p07 = 0 and a1p08 = 0"
   'Frmacc21d0_Save
End Sub

'Add By Cheng 2004/02/02
'抓傳票號碼
Private Sub ShowA1P22(strA1B01 As String, strA1B02 As String, strA1B03 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   'modify by sonia 2014/11/5 加抓傳票公司別
   StrSQLa = "Select A1P22,A1P01 From ACC1P0 Where A1P04='" & strA1B01 & strA1B02 & "' And A1P18=" & Val(Replace(strA1B03, "/", ""))
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       Me.Text23.Text = "" & rsA.Fields(0).Value
       strA1P01 = "" & rsA.Fields(1).Value       'ADD BY SONIA 2014/11/5
   Else
       Me.Text23.Text = ""
       strA1P01 = ""                             'ADD BY SONIA 2014/11/5
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Sub

'Add by Amy 2014/11/03 從aacc_sav搬回
Public Sub Frmacc21d0_Save()
Dim strMsg As String 'Add by Amy 21014/11/04
'add by sonia 2017/4/11
Dim strReceiverID As String
Dim strSubject As String, strContent As String
'end 2017/4/11
   
    'Added by Lydia 2021/12/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        strControlButton = MsgText(602)
        Exit Sub
    End If
    'end 2021/12/07
   On Error GoTo Checking
   With Frmacc21d0
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If .Text3 = MsgText(601) Then
            MsgBox MsgText(10) & .Label3, , MsgText(5)
            strControlButton = MsgText(602)
            .Text3.SetFocus
            Exit Sub
         End If
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
            'Add by Amy 2014/11/04 +系統日檢查
            If MaskEdBox1.Enabled = True And ChkWorkData("1", DBDATE(.MaskEdBox1.Text), strMsg) = False Then
                MsgBox .Label2 & strMsg, , MsgText(5)
                strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
                Exit Sub
            End If
            'end 2014/11/04
         End If
         
         'add by sonia 2013/11/14
         If .Combo2 = MsgText(601) Then
            MsgBox "付款方式不可空白！", , MsgText(5)
            strControlButton = MsgText(602)
            .Combo2.SetFocus
            Exit Sub
         End If
         '2013/11/14 end
         'If .MaskEdBox2.Text = MsgText(601) Or .MaskEdBox2.Text = MsgText(29) Then
         '   MsgBox .Label6 & MsgText(52), , MsgText(5)
         '   strControlButton = MsgText(602)
         '   .MaskEdBox2.SetFocus
         '   Exit Sub
         'Else
         '   If DateCheck(.MaskEdBox2.Text) = MsgText(603) Then
         '      MsgBox .Label6 & MsgText(63), , MsgText(5)
         '      strControlButton = MsgText(602)
         '      .MaskEdBox2.SetFocus
         '      Exit Sub
         '   End If
         'End If
         If ExistCheck("fagent", "fa01 || fa02", .Text1, .Label1, False) = False Then
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
'            .adoquery.Open "select a1810 from acc180 where a1803 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
            .adoquery.Open "select a1810 from acc180 where a1803 = '" & ChgSQL(.Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount = 0 Then
               MsgBox MsgText(45) & .Label1, , MsgText(5)
               strControlButton = MsgText(602)
               .adoquery.Close
               .Text1.SetFocus
               Exit Sub
            End If
            .adoquery.Close
         End If
         If .adoquery.State = adStateOpen Then
            .adoquery.Close
         End If
         .adoquery.CursorLocation = adUseClient
'         .adoquery.Open "select * from acc1c0 where a1c01 = '" & .Text3 & "' and a1c02 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
         .adoquery.Open "select * from acc1c0 where a1c01 = '" & .Text3 & "' and a1c02 = '" & ChgSQL(.Text1) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount = 0 Then
            MsgBox MsgText(216), , MsgText(5)
            strControlButton = MsgText(602)
            .adoquery.Close
            Exit Sub
         End If
         .adoquery.Close
      End If
      If strSaveConfirm = MsgText(3) Then
         If .adoacc1b0.RecordCount <> 0 Then
            .adoacc1b0.Find "a1b01 = '" & .Text3 & "'", 0, adSearchForward, 1
            If .adoacc1b0.EOF = False Then
'               .adoacc1b0.Find "a1b02 = '" & .Text1 & "'", 0, adSearchForward, .adoacc1b0.Bookmark
               .adoacc1b0.Find "a1b02 = '" & ChgSQL(.Text1) & "'", 0, adSearchForward, .adoacc1b0.Bookmark
               If .adoacc1b0.EOF = False Then
                  Exit Sub
               End If
            End If
         End If
         .adoacc1b0.AddNew
      End If
      .adoacc1b0.Fields("a1b01").Value = .Text3
      .adoacc1b0.Fields("a1b02").Value = .Text1
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc1b0.Fields("a1b03").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc1b0.Fields("a1b03").Value = Null
      End If
      If .MaskEdBox2.Text <> MsgText(601) And .MaskEdBox2.Text <> MsgText(29) Then
         .adoacc1b0.Fields("a1b05").Value = Val(FCDate(.MaskEdBox2.Text))
      Else
         .adoacc1b0.Fields("a1b05").Value = Null
      End If
      If .Combo2 <> MsgText(601) Then
         .adoacc1b0.Fields("a1b06").Value = Mid(.Combo2, 1, 1)
      Else
         .adoacc1b0.Fields("a1b06").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc1b0.Fields("a1b04").Value = Val(.Text4)
      Else
         .adoacc1b0.Fields("a1b04").Value = 0
      End If
      If .Text5 <> MsgText(601) Then
         .adoacc1b0.Fields("a1b07").Value = .Text5
      Else
         .adoacc1b0.Fields("a1b07").Value = Null
      End If
      If strSaveConfirm = MsgText(3) Then
         .adoacc1b0.Fields("a1b08").Value = Val(strSrvDate(2))
         .adoacc1b0.Fields("a1b09").Value = ServerTime
         .adoacc1b0.Fields("a1b10").Value = strUserNum
      Else
         .adoacc1b0.Fields("a1b11").Value = Val(strSrvDate(2))
         .adoacc1b0.Fields("a1b12").Value = ServerTime
         .adoacc1b0.Fields("a1b13").Value = strUserNum
      End If
      'add by sonia 2025/8/5 同時更新A1906否則案件系統與傳票的結匯金額會差很多CFP-021904
      adoTaie.Execute "update acc190 set a1906 =(Select Sum(A1p07)/Sum(A1p21) From Acc1p0,Acc180 where a1901=a1801 and a1908||a1803=a1p04 And SUBSTR(A1P05,1,4) in ('2201','6130') And A1p07>0 and a1902=substr(a1p23,1,9) ) " & _
                      "where a1902 in (select Distinct Substr(A1p23,1,9) From Acc1p0 where a1p02 = 'I' and a1p04 = '" & ChgSQL(Text3 & Text1) & "' and SUBSTR(A1P05,1,4) in ('2201','6130') and a1p07 > 0 )"
      'end 2025/8/5
'      adoTaie.Execute "update acc1p0 set a1p18 = " & Val(FCDate(.MaskEdBox1.Text)) & ", a1p27 = decode(a1p22, null, null, 'Y') where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & .Text3 & .Text1 & "'"
      adoTaie.Execute "update acc1p0 set a1p18 = " & Val(FCDate(.MaskEdBox1.Text)) & ", a1p27 = decode(a1p22, null, null, 'Y') where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & ChgSQL(.Text3 & .Text1) & "'"
      .adoacc1b0.UpdateBatch
      .RecordShow

'cancel by sonia 2017/4/12 瑞婷需求,婧瑄及辜反對
'      'add by sonia 2017/4/11 若案件有未輸帳單且未列印且收據自動列印時間點設在'代理人請款之匯款日'的收據要通知財務處總帳人員(P116373)
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select axf03,cp13,st02,sum(axf04) amt from acc151,caseprogress,staff,acc190 where a1908='" & .Text3 & "' and a1902=axf01(+) " & _
'                    "and cp01(+)=substr(axf03,1,Length(axf03)-9) and cp02(+)=substr(axf03, Length(axf03) - 8, 6) and cp03(+) = substr(axf03, Length(axf03) - 2, 1) and cp04(+) = substr(axf03, Length(axf03) - 1, 2) " & _
'                    "and cp60 is not null and cp61 is null and cp151='3' and cp13=st01(+) group by axf03,cp13,st02"
'      If adoquery.RecordCount <> 0 Then
'         strReceiverID = Pub_GetSpecMan("財務處總帳人員")
'         With adoquery
'         .MoveFirst
'         Do While Not .EOF
'            strSubject = "" & adoquery.Fields("axf03") & "帳單已結匯,尚有未列印收據 ！"
'            strContent = "本所案號：" & "" & adoquery.Fields("axf03") & vbCrLf
'            strContent = strContent & "智權人員：" & "" & adoquery.Fields("cp13") & "　" & adoquery.Fields("st02") & vbCrLf
'            strContent = strContent & "結匯外幣：" & "" & adoquery.Fields("amt") & vbCrLf & vbCrLf
'            strContent = strContent & "請確認未列印收據是否要改金額或列印收據 ！"
'            PUB_SendMail strUserNum, strReceiverID, "", strSubject, strContent, , , , , , , , , , True
'            .MoveNext
'         Loop
'         End With
'      End If
'      adoquery.Close
'      'end 2017/4/11
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'Modify by Amy 2014/11/04 Frmacc21d0_Clear()由aacc_cls搬回
Public Sub Frmacc21d0_Clear()
   With Frmacc21d0
      .Text1 = ""
      .Text2 = ""
      .Text3 = ""
      If .MaskEdBox1.Text = MsgText(29) Or .MaskEdBox1.Text = MsgText(601) Then
         .MaskEdBox1.Mask = ""
         .MaskEdBox1.Text = ""
         .MaskEdBox1.Mask = DFormat
      End If
      .MaskEdBox1.Enabled = True 'Add by Amy 2014/11/04
      'Modify by Morgan 2004/11/10 資料要清除
      'If .MaskEdBox2.Text = MsgText(29) Or .MaskEdBox2.Text = MsgText(601) Then
         .MaskEdBox2.Mask = ""
         .MaskEdBox2.Text = CFDate(ACDate(ServerDate))
         .MaskEdBox2.Mask = DFormat
      'End If
      .Combo2 = ""
      .Text4 = ""
      .Text5 = ""
      .AdodcRefresh
      .SumShow
      .Text1.SetFocus
   End With
End Sub

Public Sub SetData(ByVal strKeyCode As String)
    Select Case strKeyCode
        Case "F3"
            
        Case "F9"
            '解改日期存檔再修改不會存acc1p0 (因tag只記錄前一次改前資料)
            MaskEdBox1.Tag = Val(FCDate(MaskEdBox1))
        Case Else
    End Select
End Sub
'end 2014/11/04

'Added by Lydia 2017/11/01 依公司別選取
Private Sub cmdC_Click(Index As Integer)
Dim Strindex As String  'Added by Lydia 2020/08/31
   
   'Added by Lydia 2020/08/31 9/1取消"1公司的國外結匯"; 匯票輸入, 選公司部分改成:2 J L
   Select Case Index
       Case 0  '1公司=>已隱藏
          Strindex = "1"
       Case 1   '2公司
          Strindex = "2"
       Case 2  'J公司
          Strindex = "J"
       Case 3   'L公司
          Strindex = "L"
   End Select
   'end 2020/08/31
   If Adodc3.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc3.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc3.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc3.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text1.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   Screen.MousePointer = vbHourglass
   Do While Adodc1.Recordset.EOF = False
      'Modified by Lydia 2020/08/31
      'If Val("" & Adodc1.Recordset.Fields("a1917").Value) = Index + 1 Then
      If "" & Adodc1.Recordset.Fields("a1917").Value = Strindex Then
        Adodc2Save
        If strControlButton = MsgText(602) Then
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
      End If
      Adodc1.Recordset.MoveNext
   Loop
   Adodc1.Recordset.Requery
   
   Adodc2.Recordset.Requery
   If Adodc2.Recordset.RecordCount > 0 Then
      If IsNull(Adodc2.Recordset.Fields("a1505").Value) Then
         strCurrency = "USD"
      Else
         strCurrency = Adodc2.Recordset.Fields("a1505").Value
      End If
      AutoGen
   End If
   Screen.MousePointer = vbDefault
End Sub
