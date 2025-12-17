VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41l0 
   AutoRedraw      =   -1  'True
   Caption         =   "ACS待分潤"
   ClientHeight    =   4704
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8772
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4704
   ScaleWidth      =   8772
   Begin VB.CommandButton CmdSearch 
      Height          =   300
      Left            =   3990
      Picture         =   "Frmacc41l0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   45
      Width           =   350
   End
   Begin VB.TextBox txtInsTime 
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
      Left            =   7890
      TabIndex        =   37
      Top             =   3330
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox txtCusNo 
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
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1080
      Width           =   1200
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2820
      Width           =   2000
   End
   Begin VB.CommandButton CmdClear 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   7290
      Picture         =   "Frmacc41l0.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   12
      ToolTipText     =   "清除畫面"
      Top             =   4020
      Width           =   550
   End
   Begin VB.CommandButton CmdDel 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   7935
      Picture         =   "Frmacc41l0.frx":09CC
      Style           =   1  '圖片外觀
      TabIndex        =   13
      ToolTipText     =   "取消"
      Top             =   4020
      Width           =   550
   End
   Begin VB.TextBox txtSystem 
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   21
      Text            =   "ACS"
      Top             =   45
      Width           =   612
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   2
      Left            =   3465
      MaxLength       =   2
      TabIndex        =   2
      Top             =   45
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   1
      Left            =   3075
      MaxLength       =   1
      TabIndex        =   1
      Top             =   45
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   0
      Left            =   1830
      MaxLength       =   6
      TabIndex        =   0
      Top             =   45
      Width           =   1212
   End
   Begin VB.TextBox txtCmp 
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   4
      Top             =   405
      Width           =   612
   End
   Begin VB.TextBox TxtSum 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1800
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
      Height          =   315
      Index           =   2
      Left            =   5940
      MaxLength       =   5
      TabIndex        =   8
      Top             =   3330
      Width           =   1100
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
      Height          =   315
      Index           =   1
      Left            =   1650
      MaxLength       =   5
      TabIndex        =   7
      Top             =   3330
      Width           =   1000
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
      Height          =   315
      Index           =   3
      Left            =   1650
      MaxLength       =   9
      TabIndex        =   9
      Top             =   3690
      Width           =   2000
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
      Height          =   315
      Index           =   4
      Left            =   5940
      MaxLength       =   3
      TabIndex        =   10
      Top             =   3690
      Width           =   612
   End
   Begin VB.CommandButton CmdSaveAcc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生傳票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7560
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   420
      Width           =   1150
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   750
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41l0.frx":1036
      Height          =   1305
      Left            =   120
      TabIndex        =   36
      Top             =   1470
      Width           =   8535
      _ExtentX        =   15050
      _ExtentY        =   2307
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
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
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "SalesNo"
         Caption         =   "對沖代號(業)"
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
         DataField       =   "Ohther"
         Caption         =   "對沖代號(其他)"
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
         DataField       =   "Amt"
         Caption         =   "金　　額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "AcDept"
         Caption         =   "部 門"
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
         DataField       =   "Memo"
         Caption         =   "摘　　要"
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
         DataField       =   "InsTime"
         Caption         =   "時間"
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
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1607.811
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   708.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   708.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   1410
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
   Begin MSForms.TextBox txtOrgSales 
      Height          =   300
      Left            =   5580
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   750
      Width           =   1900
      VariousPropertyBits=   679493663
      BackColor       =   16777215
      Size            =   "3351;529"
      BorderColor     =   -2147483643
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCmpName 
      Height          =   315
      Left            =   1830
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   390
      Width           =   3200
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "5644;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "最早智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3990
      TabIndex        =   34
      Top             =   780
      Width           =   1700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "公 司 別："
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
      Left            =   30
      TabIndex        =   33
      Top             =   405
      Width           =   1200
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期："
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
      Left            =   30
      TabIndex        =   32
      Top             =   750
      Width           =   1200
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "客　　戶："
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
      Left            =   30
      TabIndex        =   31
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "本所案號："
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
      Left            =   30
      TabIndex        =   30
      Top             =   45
      Width           =   1200
   End
   Begin MSForms.TextBox txtCusName 
      Height          =   315
      Left            =   2445
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1080
      Width           =   5055
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8908;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseName 
      Height          =   300
      Left            =   4500
      TabIndex        =   22
      Top             =   45
      Width           =   4205
      VariousPropertyBits=   679493663
      BackColor       =   16777215
      Size            =   "7417;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "目前餘額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   180
      TabIndex        =   28
      Top             =   2880
      Width           =   1200
   End
   Begin MSForms.TextBox txtNote 
      Height          =   495
      Left            =   1650
      TabIndex        =   11
      Top             =   4050
      Width           =   5505
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "9710;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtSalesName 
      Height          =   315
      Left            =   2685
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3330
      Width           =   960
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "金額合計："
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
      Left            =   4800
      TabIndex        =   26
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "前月未過帳,不可輸當月"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   24
      Top             =   465
      Width           =   2505
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "摘　  　　要："
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
      Left            =   150
      TabIndex        =   20
      Top             =   4020
      Width           =   1500
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "金           額："
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
      Left            =   150
      TabIndex        =   19
      Top             =   3690
      Width           =   1500
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(其他)："
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
      Left            =   4200
      TabIndex        =   18
      Top             =   3330
      Width           =   1800
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "(W)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6630
      TabIndex        =   17
      Top             =   3690
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(業)："
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
      Left            =   180
      TabIndex        =   16
      Top             =   3330
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "部　　  　　門："
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
      Left            =   4140
      TabIndex        =   15
      Top             =   3690
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1455
      Left            =   120
      Top             =   3210
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc41l0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2023/02/17
Option Explicit

Dim ado41L0 As New ADODB.Recordset
Dim i As Integer
Dim bolPreHasAx210 As Boolean, strPreAxb(4 To 8) As String '系統-1個月是否已過帳/系統-1個月傳票起迄
Dim strPreYM As String, strMaxSP01 As String, strA0b01 As String, strA0b05 As String '系統年月-1個月/目前智權點數輸入年月/目前過帳日/目前業績輸入關閉年月
Dim strDefDate As String, strMaxDate As String, m_strLC11 As String '登入時預設傳票日(西元)/允許的最大傳票日/案件請人1
Dim stA0202 As String, stAx212 As String, stAx213 As String '傳票號/預設摘要/4191 會計科目之傳票號最小的對沖-其他
Dim strNowSys As String, strNowCaseNo(2) As String '目前畫面案號資料

Private Sub cmdClear_Click()
    Call FormClear(2)
    Text1(1).SetFocus
End Sub

Private Sub cmdDel_Click()
    Dim ThisRec, stCmd As String
    
    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF = True Then
        Adodc1.Recordset.MoveFirst
        ThisRec = Adodc1.Recordset.Bookmark
    Else
        ThisRec = Adodc1.Recordset.Bookmark
        Adodc1.Recordset.MoveNext
    End If
    stCmd = "Delete AccRpt41L0 Where InsTime=" & Adodc1.Recordset.Fields("InsTime").Value
    adoTaie.Execute stCmd, intI
    
    Call QueryData
    If Adodc1.Recordset.BOF = False Then Adodc1.Recordset.Bookmark = ThisRec
    Call FormClear(2)
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSaveAcc_Click()
    Dim stCmd As String
    
    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    
    If FormCheck(1) = False Then
        Exit Sub
    End If
    
    If SaveVoucher = True Then
        stCmd = "Delete AccRpt41L0"
        adoTaie.Execute stCmd
        
        QueryData
        Call FormClear(0)
        '開傳票輸入核對
        With Frmacc4120
            .MaskEdBox1 = MaskEdBox1 'Add by Amy 2024/10/08 避免彈無傳票日訊息
            .Tag = Me.Name
            Me.Hide
            .Text1 = txtCmp
            .Text2 = stA0202
            .bolF3 = True
            .Command3_Click
            .bolF3 = False
        End With
        strNowSys = txtSystem: strNowCaseNo(0) = txtCode(0): strNowCaseNo(1) = txtCode(1): strNowCaseNo(2) = txtCode(2)
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim bolData As Boolean
    
    If txtCode(0) = MsgText(601) Then
        MsgBox "本所案號不可為空！"
        Exit Sub
    End If
    
    '第3及4碼案號未輸,補0
    If txtCode(1) = MsgText(601) Then txtCode(1) = "0"
    If txtCode(2) = MsgText(601) Then txtCode(2) = "00"
    
    If Adodc1.Recordset.RecordCount = 0 Then
        strNowSys = txtSystem: strNowCaseNo(0) = txtCode(0): strNowCaseNo(1) = txtCode(1): strNowCaseNo(2) = txtCode(2)
    ElseIf txtCode(0) & txtCode(1) & txtCode(2) <> strNowCaseNo(0) & strNowCaseNo(1) & strNowCaseNo(2) Then
        '避免DataGrid有資料,又改案號,導致之前寫入的資料會有問題
        MsgBox "未產生傳票,不可修改案號！"
        txtCode(0) = strNowCaseNo(0): txtCode(1) = strNowCaseNo(1): txtCode(2) = strNowCaseNo(2)
        Exit Sub
    End If
 
    'DataGrid沒資料,才彈智權或客戶多筆 提醒 ex:ACS-000091 Grid有資料,會再彈
    If Adodc1.Recordset.RecordCount = 0 Then
        Call FormClear(1)
            
        bolData = SetData(0)
        If bolData = True Then
            txtCaseName = CaseNameShow(txtSystem, txtCode(0), txtCode(1), txtCode(2), 1)
        End If
        QueryData
        Call SetInvDate
    End If
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
    If Not Adodc1.Recordset.EOF Then
        AdodcShow
    End If
End Sub

Private Sub Form_Activate()
    tool3_enabled
    txtCode(0).SetFocus
End Sub

Private Sub Form_Load()
    Dim stTP(1) As String
    Dim intX As Integer, intY As Integer
    Dim sglWidth As Single, sglHeight As Single
  
    strFormName = Name
    Me.Width = 8895
    Me.Height = 5115
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Call FormClear(0)
    
    strPreYM = Val(Mid(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 1, 6)) - 191100 '系統前一個月
    strA0b01 = GetA0b01(strA0b05)
    strMaxSP01 = Val(GetMaxSP01(True)) - 191100
   
    txtCmp = "1" '預設1公司
    Call txtCmp_Validate(False)
    Call SetInvDate
    strDefDate = FCDate(MaskEdBox1)
    If bolPreHasAx210 = True Then
        '系統-1月[已]過帳,最大傳票日只能為系統當日
        strMaxDate = strSrvDate(2)
    Else
        '系統-1月[未]過帳,最大傳票日只能為上個月最後一個工作日
        strMaxDate = GetPreMonLastDate(strSrvDate(1), True)
    End If
    Call QueryData
    
    '當掉再進入
    If Adodc1.Recordset.RecordCount > 0 Then
        txtCode(0) = Adodc1.Recordset.Fields("CaseNo2").Value
        txtCode(1) = Adodc1.Recordset.Fields("CaseNo3").Value
        txtCode(2) = Adodc1.Recordset.Fields("CaseNo4").Value
        txtCmp = Adodc1.Recordset.Fields("Cmp").Value: Call txtCmp_Validate(False)
        MaskEdBox1.Mask = ""
        MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("InvDate").Value)
        MaskEdBox1.Mask = DFormat
        Call SetData(1)
        
        txtCusNo = Adodc1.Recordset.Fields("CusNo").Value: Call txtCusNo_Validate(False)
        strNowSys = txtSystem
        strNowCaseNo(0) = txtCode(0): strNowCaseNo(1) = txtCode(1): strNowCaseNo(2) = txtCode(2)
        txtCaseName = CaseNameShow(txtSystem, txtCode(0), txtCode(1), txtCode(2), 1)
        Call GetSum
        
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode) 'Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(1, KeyCode) 'Form2.0
    KeyDefine KeyCode
End Sub

'intChoose:0-全部 / 1-除案號其他都清 / 2-只清Grid及明細
Private Sub FormClear(ByVal intChoose As Integer)
    Dim oTxt
    
    If intChoose = 0 Then
        For Each oTxt In txtCode
            oTxt.Text = ""
        Next
    End If
    If intChoose <> 2 Then
        txtCaseName.Text = ""
        txtCmp.Text = "1"
        txtOrgSales.Text = ""
        MaskEdBox1.Tag = ""
        txtCusNo.Text = ""
        txtCusName.Text = ""
        txtAmt.Text = "" '目前餘額
        TxtSum.Text = "" '金額合計
    End If
    For Each oTxt In Text1
        oTxt.Text = ""
        oTxt.Tag = ""
    Next
    txtSalesName.Text = ""
    txtNote.Text = "" '摘要
    txtInsTime.Text = "" '新增時間
End Sub

'intChoose:1-只控制案號 / 2-控制公司~客戶(上半部) / 3-只控制明細
Private Sub TxtLock(ByVal intChoose As Integer, ByVal bolLock As Boolean)
    Dim oTxt
    
    If intChoose = 1 Then
        For Each oTxt In txtCode
            oTxt.Locked = bolLock
        Next
    End If
    If intChoose = 2 Then
        'txtCmp.Locked = bolLock '公司別目前鎖住不可改
        If bolLock = False Then
            MaskEdBox1.Tag = ""
        Else
            MaskEdBox1.Tag = "Lock"
        End If
        txtCusNo.Locked = bolLock '客戶
    End If
    
    If intChoose = 3 Then
        For Each oTxt In Text1
            oTxt.Locked = bolLock
        Next
        txtNote.Locked = bolLock '摘要
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Adodc1.Recordset.RecordCount > 0 Then
        If MsgBox("未產傳票要離開？" & vbCrLf & _
                          "是：離開,資料會清空" & vbCrLf & _
                          "否：不離開", vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = True
            Exit Sub
        Else
            strExc(1) = "Delete Accrpt41l0 "
            adoTaie.Execute strExc(1)
        End If
    End If
    
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    strTrackMode = "" 'Form2.0 記錄鍵盤傳入順序(清除)
    Call PUB_GetLock("", "Frmacc41l0")
    
    Set Frmacc41l0 = Nothing
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
    If MaskEdBox1.Tag = "Lock" Then KeyAscii = 0: Exit Sub
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    '不寫在MaskEdBox1_LostFocus,MaskEdBox1跳離開不會觸發
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then Exit Sub
    
    If FormCheck = False Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Dim stTmp As String
    
    If Index = 2 Then Exit Sub
    If (Index = 3 Or Index = 4) And Trim(Text1(Index)) = MsgText(601) Then Exit Sub
    
    Select Case Index
        Case 1 '對沖代號-業
            If Trim(Text1(Index)) <> MsgText(601) Then
                If FormCheck = False Then
                    Exit Sub
                End If
                
                '預帶摘要
                If txtCode(1) <> "0" And Len(Trim(txtCode(1))) > 0 Then stTmp = stTmp & txtCode(1)
                If txtCode(2) <> "00" And Len(Trim(txtCode(2))) > 0 Then stTmp = stTmp & txtCode(2)
                Text1(4) = "W" '預設W部門
                '摘要為空 or 對沖代號-業 有修改
                If (txtNote = MsgText(601) Or Text1(1).Text <> Text1(1).Tag) And txtSalesName <> MsgText(601) And txtCusName <> MsgText(601) And txtCaseName <> MsgText(601) Then
                    txtNote = PUB_GetShortName(Text1(1)) & "/" & Left(txtCusName, 6) & "/" & txtSystem & txtCode(0) & txtCode(1) & txtCode(2)
                End If
            End If
        Case 3 '金額
            If FormCheck = False Then
                Exit Sub
            End If
        Case 4 '部門
            If FormCheck = False Then
                Exit Sub
            End If
    End Select
End Sub

Private Sub txtCmp_GotFocus()
    TextInverse txtCmp
End Sub

Private Sub txtCmp_Validate(Cancel As Boolean)
    If txtCmp = MsgText(601) Then Exit Sub
    
    txtCmpName = A0802Query(txtCmp, True)
    
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
    TextInverse txtCode(Index)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then Exit Sub
    
    '第3及4碼案號未輸,補0
    If txtCode(Index - 1) <> MsgText(601) And txtCode(Index) = MsgText(601) Then
        If Index = 1 Then txtCode(Index) = "0"
        If Index = 2 Then txtCode(Index) = "00"
    End If
 
End Sub

Private Sub txtCusNo_GotFocus()
    TextInverse txtCusNo
End Sub

Private Sub txtCusNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCusNo_Validate(Cancel As Boolean)
    If txtCusNo = MsgText(601) Then Exit Sub
     
    txtCusNo = GetNewFagent(txtCusNo)
    txtCusName = GetCustomerName(txtCusNo)
End Sub

Private Sub txtNote_GotFocus()
    TextInverse txtNote
End Sub

'顯示資料
Private Sub QueryData()
    If ado41L0.State = adStateOpen Then ado41L0.Close
    ado41L0.CursorLocation = adUseClient
    strSql = "Select * From AccRpt41L0 Order by InsTime "
    ado41L0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly

    Set Adodc1.Recordset = ado41L0
    Adodc1.Recordset.Requery
    
    If Adodc1.Recordset.RecordCount = 0 Then
        TxtSum = ""
        CmdSaveAcc.Enabled = False
    Else
        Call GetSum
        CmdSaveAcc.Enabled = True
    End If
    
    If Adodc1.Recordset.RecordCount = 0 Then
        Call TxtLock(2, False)
        If txtCaseName = MsgText(601) And txtAmt = MsgText(601) Then
            Call TxtLock(3, True)
        Else
            Call TxtLock(3, False)
        End If
    Else
        Call TxtLock(2, True)
        Call TxtLock(3, False)
    End If
End Sub

Private Sub AdodcShow()
    Call FormClear(2)
    If Adodc1.Recordset.RecordCount > 0 Then
        '對沖代號-業
        Text1(1) = "" & Adodc1.Recordset.Fields("SalesNo")
        Text1(1).Tag = Text1(1)
        txtSalesName = GetStaffName(Text1(1), True)
        '對沖代號-其他
        Text1(2) = "" & Adodc1.Recordset.Fields("Other")
        Text1(2).Tag = Text1(2)
        '金額
        Text1(3) = PUB_ChgFormat("" & Adodc1.Recordset.Fields("Amt"), True)
        Text1(3).Tag = Text1(3)
        '部門
        Text1(4) = "" & Adodc1.Recordset.Fields("AcDept")
        txtNote = "" & Adodc1.Recordset.Fields("Memo")
        txtInsTime = Adodc1.Recordset.Fields("InsTime")
    End If
End Sub


Private Function FormIns() As Boolean
    Dim stCmd As String, stField As String, stVal As String
    
    '新增
    If txtInsTime = MsgText(601) Then
        stField = "CaseNo1,CaseNo2,CaseNo3,CaseNo4"
        stVal = "'" & txtSystem & "','" & txtCode(0) & "','" & txtCode(1) & "','" & txtCode(2) & "' "
        
        stField = stField & ",Cmp,InvDate,CusNo"
        stVal = stVal & ",'" & txtCmp & "'," & Val(FCDate(MaskEdBox1)) & ",'" & txtCusNo & "' "
        
        stField = stField & ",SalesNo,Other,Amt,AcDept,Memo,InsTime"
        stVal = stVal & ",'" & Text1(1) & "'," & CNULL(ChgSQL(Text1(2))) & "," & Val(Text1(3)) & ",'" & Text1(4) & "'," & CNULL(ChgSQL(txtNote)) & "," & ServerTime
        
        stCmd = "Insert Into AccRpt41L0 (" & stField & ") Values(" & stVal & ")"
    '修改
    Else
        stField = "SalesNo='" & Text1(1) & "',Other=" & CNULL(ChgSQL(Text1(2))) & ",Amt=" & Val(Replace(Text1(3), ",", "")) & ",AcDept='" & Text1(4) & "',Memo=" & CNULL(ChgSQL(txtNote)) & " "
        stCmd = "Update AccRpt41L0 Set " & stField & " Where InsTime=" & txtInsTime
    End If
    adoTaie.Execute stCmd
End Function

Private Sub KeyDefine(KeyCode As Integer)
    Dim strMsg As String
    Dim nResponse
    
    Select Case KeyCode
        Case vbKeyInsert
            'Form2.0控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
            If PUB_ChkTrackMode = False Then Exit Sub
            DataGrid1.SetFocus
            If FormCheck(2) = False Then Exit Sub
            
            FormIns
            QueryData
            Call FormClear(2)
            Text1(1).SetFocus
        Case vbKeyDown
            If Adodc1.Recordset.EOF = False Then
                Adodc1.Recordset.MoveNext
                If Adodc1.Recordset.EOF = False Then
                    AdodcShow
                End If
            End If
       Case vbKeyUp
            If Adodc1.Recordset.BOF = False Then
                Adodc1.Recordset.MovePrevious
                If Adodc1.Recordset.BOF = False Then
                    AdodcShow
                End If
            End If
    End Select
    KeyEnter KeyCode
End Sub

'取得案件客戶及餘額資料
'intChoose:0-查案號/1-當掉抓資料/2-FormCheck只確認案號是否餘額正確(未鎖案號)
Private Function SetData(intChoose As Integer) As Boolean
    Dim RsQ As ADODB.Recordset, intQ As Integer, rsA As ADODB.Recordset, intA As Integer
    Dim stQ As String, stBase1 As String, stBase2 As String, stMsg As String, stTP(2) As String
    Dim stChkAmt As String, stChkAx208 As String, stChkAx213 As String
    Dim stCusNo As String, stSalesNo As String, stAmt As String
    
    SetData = False
    stBase1 = "And SubStr(ax214, 1, length(ax214) - 9)='" & txtSystem & "' And SubStr(ax214, length(ax214)- 8, 6)='" & txtCode(0) & "' " & _
                  "And SubStr(ax214, length(ax214)- 2,1)='" & txtCode(1) & "' And SubStr(ax214, length(ax214)- 1,length(ax214))='" & txtCode(2) & "' "
    stBase2 = stBase1
    stBase1 = GetACSData("9", Me.Name, "", "", stBase1) '抓案號 2492 餘額

    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stBase1)
    If intQ = 1 Then
        RsQ.MoveFirst
        SetData = True
        If intChoose < 2 Then
            stChkAmt = PUB_ChgFormat("" & RsQ.Fields("AMT"), True)
            stTP(1) = "Y"
            '抓取最小傳票日 4191 「借方」資料
            stTP(0) = GetACSData("2", Me.Name, "", ",Acc020", stBase2 & "And ax206>0 ", stTP(1))
            'ax208對沖-客戶
            stChkAx208 = Mid(stTP(1), 1, Val(InStr(stTP(1), ";") - 1))
            'ax213(智權)
            stTP(1) = Replace(stTP(1), stChkAx208 & ";", "")
            stChkAx213 = Mid(stTP(1), 1, Val(InStr(stTP(1), ";") - 1))
            
            txtAmt = stChkAmt
            txtCusNo = stChkAx208
            stAx213 = stChkAx213
            txtOrgSales = stAx213 & " " & GetPrjSalesNM(stAx213)
            'ax212摘要
            stAx212 = Replace(stTP(1), stAx213 & ";", "")
            
            If intChoose = 0 Then
                '抓取 4191,ax208(客戶)及ax213(智權),若多筆彈訊息
                stQ = GetACSData("3", Me.Name, "", "", stBase2 & "And ax206>0 ")
                intA = 1
                Set rsA = ClsLawReadRstMsg(intA, stQ)
                If intQ = 1 Then
                    rsA.MoveFirst
                    '超過一筆才顯示
                    If rsA.RecordCount > 1 Then
                        stMsg = ""
                        Do While rsA.EOF = False
                            If IsNull(rsA.Fields("ax208")) Then
                                stCusNo = stCusNo & ",空白"
                            ElseIf InStr(stCusNo, "" & rsA.Fields("ax208")) = 0 Then
                                stCusNo = stCusNo & "," & rsA.Fields("ax208")
                            End If
                            If IsNull(rsA.Fields("ax213")) Then
                                stSalesNo = stSalesNo & ",空白"
                            ElseIf InStr(stSalesNo, "" & rsA.Fields("ax213")) = 0 Then
                                stSalesNo = stSalesNo & "," & rsA.Fields("ax213") & " " & GetPrjSalesNM(rsA.Fields("ax213"))
                            End If
                            
                            rsA.MoveNext
                        Loop
                        stCusNo = Mid(stCusNo, 2)
                        stSalesNo = Mid(stSalesNo, 2)
                        
                        If stCusNo <> MsgText(601) And InStr(stCusNo, ",") > 0 Then
                            txtCusNo = "": txtCusName = "" '多筆不預帶
                            stMsg = stMsg & "客戶為" & stCusNo & vbCrLf
                        End If
                        
                        If stSalesNo <> MsgText(601) And InStr(stSalesNo, ",") > 0 Then
                            stMsg = stMsg & "對沖其他(業務)為" & stSalesNo & vbCrLf
                        End If
                        If stMsg <> MsgText(601) Then
                            MsgBox "目前傳票" & vbCrLf & stMsg & _
                                            "請確認！"
                            txtCusNo.SetFocus
                        End If
                    End If
                End If
                
            End If
            If txtCusNo <> MsgText(601) Then
                Call txtCusNo_Validate(False)
            End If
            
        End If
    Else
        stAx212 = "": stAx213 = ""
        MsgBox "無此案號資料！"
    End If
    
    Set RsQ = Nothing
End Function

Private Sub SetInvDate()
    MaskEdBox1.Mask = ""
    
    '系統-1月,傳票資料是否已寫入
    Call bolAcc0b1(1, strPreYM, strPreAxb())
    '系統-1月,實績傳票是否已過帳
    If strPreAxb(4) <> MsgText(601) Then
        bolPreHasAx210 = Pub_ChkAxbPost(strPreAxb(4), strPreAxb(5))
    End If
    
    '系統-1月,已過帳,可輸當月
    If bolPreHasAx210 = True Then
        '因可一直輸,避免輸錯月份,控制系統前一個月傳票已過帳,才可輸當月
        '預設畫面公司別當月最大傳票日
        MaskEdBox1.Text = CFDate(Pub_GetMaxA0205(txtCmp, Val(Left(strSrvDate(1), 6)) - 191100))
    '業績輸入未開及業績輸入已開尚未關閉都只能輸系統-1個月的資料
    Else
        MaskEdBox1.Text = CFDate(Pub_GetMaxA0205(txtCmp, strPreYM))
    End If
    
    MaskEdBox1.Mask = DFormat
End Sub

'取得目前Grid 合計
Private Function GetSum() As String
    Dim RsSum As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    stQ = "Select Sum(Amt) as Amt From AccRpt41L0 "
    intQ = 1
    Set RsSum = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        TxtSum = PUB_ChgFormat("" & RsSum.Fields("Amt"), True)
    End If
    Set RsSum = Nothing
End Function

'intChoose=0:/1:產生傳票 鈕 /2:Insert 鈕
Private Function FormCheck(Optional ByVal intChoose As Integer = 0) As Boolean
    Dim intIdx As Integer, bCancel As Boolean
    Dim strMsg As String, stTmp As String, stTmp2 As String, stTmp3 As String
    
    FormCheck = False
    
    If intChoose = 1 Then
        '產生之傳票只有一張只需檢查,不需更改a0b10值
        stTmp = Pub_GetField("Acc0b0", "a0b04='1'", "a0b10")
        If stTmp = "01" Then
            MsgBox MsgText(197), , MsgText(5)
            Exit Function
        End If
    End If
    
    '產生傳票/Insert 鈕
    If intChoose <> 0 Then
        '本所案號
        stTmp = Mid(Label15, 1, Val(Len(Label15)) - 1)
        If txtCode(0) = MsgText(601) Then
            MsgBox stTmp & "不可為空！", vbExclamation, "錯誤！"
            txtCode(0).SetFocus
            Exit Function
        Else
            '避免輸入明細又改案號,帶的資料會錯
            If txtCode(0) & txtCode(1) & txtCode(2) <> strNowCaseNo(0) & strNowCaseNo(1) & strNowCaseNo(2) Then
                If Adodc1.Recordset.RecordCount = 0 Then
                    MsgBox "目前資料為" & strNowSys & strNowCaseNo(0) & strNowCaseNo(1) & strNowCaseNo(2) & vbCrLf & _
                                    "若要修改案號,請輸完案號後按查詢鈕！"
                Else
                    MsgBox "未產生傳票,不可修改案號！"
                End If
                txtCode(0) = strNowCaseNo(0): txtCode(1) = strNowCaseNo(1): txtCode(2) = strNowCaseNo(2)
                Exit Function
            End If
            If SetData(2) = False Then
                txtCode(0).SetFocus
                Exit Function
            End If
        End If
        stTmp = "公司別"
        If txtCmp = MsgText(601) Then
            MsgBox stTmp & "不可為空！", vbExclamation, "錯誤！"
            txtCmp.SetFocus
            Exit Function
        ElseIf txtCmp <> "1" Then
            MsgBox stTmp & "只能輸 1公司！", vbExclamation, "錯誤！"
            txtCmp.SetFocus
            Exit Function
        End If
        stTmp = "客戶"
        If txtCusNo = MsgText(601) Then
            MsgBox stTmp & "不可為空！", vbExclamation, "錯誤！"
            txtCusNo.SetFocus
            Exit Function
        Else
            Call txtCusNo_Validate(False)
            If txtCusName = MsgText(601) Then
                MsgBox stTmp & "輸入有誤！", vbExclamation, "錯誤！"
                txtCusNo.SetFocus
                Exit Function
            End If
        End If
    End If
    
    
    '傳票日期
    If Me.ActiveControl.Name = "MaskEdBox1" Or intChoose = 1 Or intChoose = 2 Then
        stTmp = Mid(Label10, 1, Val(Len(Label10)) - 1) & vbCrLf
        If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
            If intChoose = 0 Then
                Exit Function
            Else
                MsgBox stTmp & "不可為空！", vbExclamation, "日期錯誤！"
                MaskEdBox1.SetFocus
                Exit Function
            End If
        End If
        'Memo 檢查參考「傳票輸入-傳票日期」
        If DateCheck(MaskEdBox1) = MsgText(603) Then
            MsgBox stTmp & MsgText(63), vbExclamation, "日期錯誤！"
            MaskEdBox1.SetFocus
            Exit Function
        End If
        If ChkWorkDay(FCDate(MaskEdBox1) + 19110000) = False Then
            MsgBox stTmp & "請輸入工作日！", vbExclamation, "日期錯誤！"
            MaskEdBox1.SetFocus
            Exit Function
        End If
        '只能輸預設傳票日當月
        If Left(Val(strDefDate + 19110000), 6) <> Left(Val(FCDate(MaskEdBox1)) + 19110000, 6) Then
            MsgBox stTmp & "只可輸" & Val(Left(Val(strDefDate + 19110000), 6)) - 191100 & "月的日期", vbExclamation, "日期錯誤！"
            Exit Function
        End If
        '系統-1月[已]過帳,只能輸公司別 當月最大傳票日~系統日
        '系統-1月[未]過帳,只能輸公司別 上個月最大傳票日~上個月月底前的工作日
        If Not (Val(FCDate(MaskEdBox1)) >= Val(strDefDate) And Val(FCDate(MaskEdBox1)) <= Val(strMaxDate)) Then
            stTmp = stTmp & "只可輸入"
            If Val(strDefDate) = Val(strMaxDate) Then
                stTmp = stTmp & CFDate(strDefDate) & "！"
            Else
                stTmp = stTmp & CFDate(strDefDate) & "~" & CFDate(strMaxDate) & "間的工作日！"
            End If
            MsgBox stTmp, vbExclamation, "日期錯誤！"
            MaskEdBox1.SetFocus
            Exit Function
        End If
    End If
    
    'Text1(0)不使用
    If Me.ActiveControl.Name = "Text1" Then
        intIdx = Me.ActiveControl.Index
    End If
    
    '對沖代號-業(intIdx=1)
    If intIdx = 1 Or intChoose = 2 Then
        stTmp = Mid(Label2, 1, Val(Len(Label2)) - 1) & vbCrLf
        If Text1(1) = MsgText(601) Then
            If intChoose = 0 Then
                Exit Function
            Else
                MsgBox stTmp & "不可為空！", vbExclamation, "錯誤！"
                Text1(1).SetFocus
                Exit Function
            End If
        End If
        stTmp3 = GetStaffName(Text1(1), True, , , stTmp2)
        If Text1(1) <> MsgText(601) Then
            If stTmp3 = Empty Then
                MsgBox stTmp & "輸入錯誤,無此員工！", vbExclamation, "錯誤！"
                Text1(1).SetFocus
                Exit Function
            ElseIf intIdx = 1 Then
                '員工離職,且 有修改 或 新增
                If stTmp2 = "2" And (Text1(1).Text <> Text1(1).Tag Or Text1(1).Tag = MsgText(601) And Text1(1).Text <> MsgText(601)) Then
                    MsgBox stTmp & "輸入的員工已離職！", vbExclamation, "提醒！"
                End If
            End If
        End If
        txtSalesName = stTmp3
    End If
    
    If intIdx = 3 Or intChoose = 2 Then
        stTmp = "金額"
        If Text1(3) = MsgText(601) Or Text1(3) = "0" Then
            If intChoose = 0 Then
                Exit Function
            Else
                If Text1(3) = MsgText(601) Then
                    strMsg = "不可為空！"
                ElseIf Text1(3) = "0" Then
                    strMsg = "不可為 0！"
                End If
                If strMsg <> MsgText(601) Then
                    MsgBox stTmp & strMsg, vbExclamation, "錯誤！"
                    Text1(3).SetFocus
                    Exit Function
                End If
            End If
        End If
        If IsNumeric(Replace(Text1(3), ",", "")) = False Then
            MsgBox stTmp & "只可輸入數字！", vbExclamation, "錯誤！"
            Text1(3).SetFocus
            Exit Function
        End If
    End If
    
    If intIdx = 4 Or intChoose = 2 Then
        stTmp = "部門"
        If Text1(4) = MsgText(601) Then
            If intChoose = 0 Then
                Exit Function
            Else
                MsgBox stTmp & "不可為空！", vbExclamation, "錯誤！"
                Text1(4).SetFocus
                Exit Function
            End If
        End If
        If Text1(4) <> "W" Then
            MsgBox stTmp & "輸入錯誤,只可輸W！", vbExclamation, "錯誤！"
            Text1(4).SetFocus
            Exit Function
        End If
    End If
    
    '總金額
    If intChoose = 1 Then
        If Val(Replace(txtAmt, ",", "")) < Val(Replace(TxtSum, ",", "")) Then
            MsgBox "金額合計超過目前餘額" & vbCrLf & _
                          "請確認！", vbExclamation, "錯誤！"
            Exit Function
        End If
        '傳票日為已過帳年月彈訊息
        If Val(strA0b05) + 191100 = Left(Val(FCDate(MaskEdBox1)) + 19110000, 6) Then
            '避免財務開好幾支操作,故再抓一次,避免日期與傳票號不連號
            '系統-1月,傳票資料是否已寫入
            Call bolAcc0b1(1, strPreYM, strPreAxb())
            '系統-1月,實績傳票是否已過帳
            If strPreAxb(4) <> MsgText(601) Then
                bolPreHasAx210 = Pub_ChkAxbPost(strPreAxb(4), strPreAxb(5))
                If bolPreHasAx210 = True Then
                    MsgBox strPreYM & "月已過帳,不可新增" & vbCrLf & _
                                    "請確認！", vbExclamation, "錯誤！"
                    Exit Function
                End If
            End If
        End If
    ElseIf intChoose = 2 Then
        stTmp = ""
        '明細狀態
        If txtInsTime = MsgText(601) Then
            '新增
            stTmp = Val(Replace(TxtSum, ",", "")) + Val(Replace(Text1(3), ",", ""))
        Else
            '修改
            stTmp = Pub_GetField("Accrpt41l0", "InsTime='" & txtInsTime & "'", "Amt")
            '合計金額-修改前+修改後
            stTmp = Val(Replace(TxtSum, ",", "")) - Val(stTmp) + Val(Replace(Text1(3), ",", ""))
        End If
        If Val(Replace(txtAmt, ",", "")) < Val(stTmp) Then
            MsgBox "金額合計超過目前餘額" & vbCrLf & _
                          "請確認！", vbExclamation, "錯誤！"
            Text1(3).SetFocus
            Exit Function
        End If
    End If
    
    FormCheck = True
End Function

Private Function SaveVoucher() As Boolean
    Dim stCmd As String, stFixCmd As String, stFixCmdSP As String, stAcDate As String, stAx203 As String
    Dim stCaseNo As String, stAmt As String, stTmp(2) As String
    
On Error GoTo Checking
    
    SaveVoucher = False
    adoTaie.BeginTrans
    
    stAcDate = FCDate(MaskEdBox1.Text)
    stA0202 = AccAutoNo(MsgText(801), 4, Val(Left(stAcDate, 3)), Val(Mid(stAcDate, 4, 2)))
    stCmd = AccSaveAutoNo(MsgText(801), Right(stA0202, 4), Val(Left(stAcDate, 3)), Val(Mid(stAcDate, 4, 2)))
   
    '傳票主檔
    stCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                  "Values('" & txtCmp & "','" & stA0202 & "', " & Val(stAcDate) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
    adoTaie.Execute stCmd
    '傳票明細
    With Adodc1.Recordset
        .MoveFirst
        stCaseNo = txtSystem & txtCode(0) & txtCode(1) & txtCode(2)
        stAmt = Replace(TxtSum, ",", "")
        stAx203 = GetSeqNo(txtCmp, stA0202) '流水號
        stFixCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax208,ax209,ax212,ax213,ax214) "
        stFixCmdSP = "Insert Into SalesPoint (sp01,sp02,sp48) "
        
        '借方
        stTmp(0) = String("0", 2) & stAx203
        stCmd = "Values('" & txtCmp & "','" & stA0202 & "','" & stTmp(0) & "','TOT','2492'," & stAmt & ",0,'" & txtCusNo & "','M0101'," & CNULL(ChgSQL(stAx212)) & ",'顧服組','" & stCaseNo & "')"
        adoTaie.Execute stFixCmd & stCmd
        Do While .EOF = False
            stTmp(1) = "" & .Fields("Memo") '摘要
            stTmp(2) = "" & .Fields("Other") '對沖-其他
            '貸方
            stAx203 = ZeroBeforeNo(stAx203, 3)
            stCmd = "Values('" & txtCmp & "','" & stA0202 & "','" & stAx203 & "','" & .Fields("AcDept") & "','420101',0," & .Fields("Amt") & ",'" & txtCusNo & "','" & "" & .Fields("SalesNo") & "'," & CNULL(stTmp(1)) & "," & CNULL(stTmp(2)) & ",'" & stCaseNo & "')"
            adoTaie.Execute stFixCmd & stCmd
            
            '對沖-業務有值且系統-1個月點數輸入已開放(不論是否已關閉都要加此人員)且未過帳
            If "" & .Fields("SalesNo") <> MsgText(601) And strMaxSP01 = strPreYM And bolPreHasAx210 = False Then
                '點數輸入資料表無此人員,需新增
                If ExistCheck("SalesPoint", "sp01||sp02", Val(strPreYM) + 191100 & "" & .Fields("SalesNo"), strExc(1), False) = False Then
                    stCmd = "Values(" & Val(strPreYM) + 191100 & ",'" & .Fields("SalesNo") & "','" & GetST15(.Fields("SalesNo")) & "')"
                    adoTaie.Execute stFixCmdSP & stCmd
                End If
            End If
            
            .MoveNext
        Loop
        '借方
        stAx203 = ZeroBeforeNo(stAx203, 3)
        stCmd = "Values('" & txtCmp & "','" & stA0202 & "','" & stAx203 & "','W','420101'," & stAmt & ",0,'" & txtCusNo & "','M0101'," & CNULL(ChgSQL(stAx212)) & ",Null,'" & stCaseNo & "')"
        adoTaie.Execute stFixCmd & stCmd
        
        '貸方
        stAx203 = ZeroBeforeNo(stAx203, 3)
        stCmd = "Values('" & txtCmp & "','" & stA0202 & "','" & stAx203 & "','W','4191',0," & stAmt & ",'" & txtCusNo & "','M0101'," & CNULL(ChgSQL(stAx212)) & "," & CNULL(ChgSQL(stAx213)) & ",'" & stCaseNo & "')"
        adoTaie.Execute stFixCmd & stCmd
    End With
    adoTaie.CommitTrans
    
    SaveVoucher = True
    Exit Function
    
Checking:
    If Err.Number <> 0 Then
        MsgBox "存檔失敗: " & Err.Description, vbCritical
        adoTaie.RollbackTrans
    End If
End Function



