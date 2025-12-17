VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2111 
   AutoRedraw      =   -1  'True
   Caption         =   "收款資料"
   ClientHeight    =   5460
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   8760
   Begin VB.TextBox Text13 
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
      Left            =   7230
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1590
      Width           =   516
   End
   Begin VB.TextBox Text11 
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
      Left            =   4680
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1590
      Width           =   1572
   End
   Begin VB.OptionButton Option3 
      Height          =   255
      Left            =   4050
      TabIndex        =   23
      Top             =   818
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   4050
      TabIndex        =   22
      Top             =   458
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Left            =   4050
      TabIndex        =   21
      Top             =   98
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.TextBox Text10 
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
      TabIndex        =   20
      Top             =   5055
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2111.frx":0000
      Height          =   2040
      Left            =   270
      TabIndex        =   8
      Top             =   2970
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   3598
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "a0x02"
         Caption         =   "請款編號"
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
      BeginProperty Column01 
         DataField       =   "A0X16"
         Caption         =   "請款台幣"
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
      BeginProperty Column02 
         DataField       =   "a0x08"
         Caption         =   "請款幣別"
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
         DataField       =   "a0x05"
         Caption         =   "請款外幣"
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
         DataField       =   "a0x06"
         Caption         =   "外幣折讓"
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
         DataField       =   "a0x11"
         Caption         =   "外幣金額"
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
      BeginProperty Column06 
         DataField       =   "a0x09"
         Caption         =   "收款金額(台)"
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
      BeginProperty Column07 
         DataField       =   "a0x10"
         Caption         =   "結清"
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
         DataField       =   "a1k34"
         Caption         =   "注意事項"
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
         DataField       =   "a0x03"
         Caption         =   "規費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "a0x04"
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
      BeginProperty Column11 
         DataField       =   "A0X07"
         Caption         =   "已收外幣"
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
      BeginProperty Column12 
         DataField       =   "a0x12"
         Caption         =   "扣繳金額(台幣)"
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
      BeginProperty Column13 
         DataField       =   "a1k35"
         Caption         =   "請款單抬頭"
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
      BeginProperty Column14 
         DataField       =   "a0x13"
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
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   792
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   972.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   875.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1175.811
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   768.189
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
            Alignment       =   2
            ColumnWidth     =   659.906
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
      Height          =   405
      Left            =   7830
      Picture         =   "Frmacc2111.frx":0015
      Style           =   1  '圖片外觀
      TabIndex        =   9
      ToolTipText     =   "取消"
      Top             =   1560
      Width           =   612
   End
   Begin VB.TextBox Text8 
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
      Left            =   4290
      MaxLength       =   9
      TabIndex        =   5
      Top             =   780
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "轉暫收款"
      Enabled         =   0   'False
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
      Left            =   3708
      TabIndex        =   10
      Top             =   5100
      Width           =   1332
   End
   Begin VB.TextBox Text6 
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
      Left            =   4290
      MaxLength       =   9
      TabIndex        =   4
      Top             =   420
      Width           =   1455
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
      Left            =   4290
      MaxLength       =   9
      TabIndex        =   3
      Top             =   60
      Width           =   1455
   End
   Begin VB.TextBox Text3 
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
      Left            =   1320
      TabIndex        =   15
      Top             =   210
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Left            =   1320
      TabIndex        =   13
      Top             =   5055
      Width           =   1572
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
      Left            =   1290
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1590
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   150
      Top             =   2760
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
   Begin MSForms.TextBox txtFM2 
      Height          =   315
      Left            =   450
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   2295
      VariousPropertyBits=   671105051
      BackColor       =   16761087
      Size            =   "4048;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text14 
      Height          =   330
      Left            =   1770
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1140
      Width           =   6765
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      MaxLength       =   80
      Size            =   "11933;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA1K35 
      Height          =   330
      Left            =   1290
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2340
      Width           =   7185
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      MaxLength       =   100
      Size            =   "12674;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text12 
      Height          =   330
      Left            =   1290
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1965
      Width           =   7185
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "12674;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text9 
      Height          =   330
      Left            =   5760
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   780
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
   Begin MSForms.TextBox Text7 
      Height          =   330
      Left            =   5760
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   420
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
      Height          =   330
      Left            =   5760
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   60
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
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "請款單抬頭"
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
      Left            =   120
      TabIndex        =   29
      Top             =   2378
      Width           =   1155
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "大陸收據抬頭"
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
      TabIndex        =   28
      Top             =   1178
      Width           =   1365
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "※請注意, 若有修改下列欄位資料, 修改完再按Enter鍵以示確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1350
      TabIndex        =   27
      Top             =   2700
      Width           =   7065
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   6405
      TabIndex        =   26
      Top             =   1628
      Width           =   810
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "注意事項"
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
      Left            =   300
      TabIndex        =   25
      Top             =   2003
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "扣繳金額(台幣)"
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
      Left            =   3090
      TabIndex        =   24
      Top             =   1628
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "轉入暫收款單號"
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
      Left            =   5160
      TabIndex        =   19
      Top             =   5100
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   5145
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1515
      Left            =   30
      Top             =   30
      Width           =   8655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "代理人3"
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
      Left            =   3210
      TabIndex        =   18
      Top             =   818
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "代理人2"
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
      Left            =   3210
      TabIndex        =   17
      Top             =   458
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "代理人1"
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
      Left            =   3210
      TabIndex        =   16
      Top             =   98
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "收款單號"
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
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "溢收金額"
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
      TabIndex        =   12
      Top             =   5100
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款編號"
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
      Left            =   300
      TabIndex        =   11
      Top             =   1628
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text5、Text7、Text9、Text12、Text14、txtA1K35
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0y0 As New ADODB.Recordset
Public adoacc1k0 As New ADODB.Recordset
Public adoacc0x0 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoacc0z0 As New ADODB.Recordset
Public adoacc120 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim douAmount As Double
Dim strYes As String
'2005/5/3 ADD BY SONIA
Dim m_NAME As String
Dim m_Year As String
Dim m_bolAlert As Boolean '檢查分錄提醒
Dim m_strAlertMsg As String 'Add by Morgan 2010/6/11
Public m_Currency As String '記錄前一畫面的幣別 Add By Sindy 2014/12/10
Dim Cancel As Boolean 'Add By Sindy 2015/4/29
Dim strLOS02 As String, bolB2NeeCourt As Boolean, strLCaseNo As String, strA1P22_L As String 'Added by Morgan 2021/4/9
Dim m_A1K37 As String 'Add By Sindy 2022/3/8
Dim lngCurtFee As Long 'Added by Morgan 2022/12/13

Private Sub Command1_Click()
   AdodcDelete
   SumShow
End Sub

Private Sub Command2_Click()
   If Text10 <> MsgText(601) Then
      Exit Sub
   End If
   If Val(Text2) <= 0 Then
      Exit Sub
   End If
   Acc120Save
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
On Error GoTo Checking
   If ColIndex <= 6 Then 'Add by Morgan 2010/6/15
      '2008/7/15 MODIFY BY SONIA 收RMB故改為以台幣判斷
      'If Adodc1.Recordset.Fields("A0X11").Value >= Adodc1.Recordset.Fields("A0X05").Value Then
      If Adodc1.Recordset.Fields("A0X09").Value >= Adodc1.Recordset.Fields("A1K11").Value Then
      '2008/7/15 END
         'Modify by Morgan 2010/6/15 若外幣金額小於請款金額時允許清除結清欄位
         If Not (ColIndex = 6 And Adodc1.Recordset.Fields("A0X11").Value < Adodc1.Recordset.Fields("A0X05").Value) Then
            Adodc1.Recordset.Fields("A0X10").Value = "Y"
         End If
      End If
   End If
   SendKeys "{ENTER}"
   Select Case ColIndex
'      Case 6
'      Case 3 '外幣金額
      Case 5 '外幣金額
         Adodc1.Recordset.Fields("A0X09").Value = Adodc1.Recordset.Fields("A0X11").Value * Val(strCon3)
      'add by sonia 2025/8/28 M111403799收X11404811~X11404817台幣金額A0X09<規費A0X03時改預設規費A0X03
      If Adodc1.Recordset.Fields("A0X09").Value < Adodc1.Recordset.Fields("A0X03").Value Then
         Adodc1.Recordset.Fields("A0X09").Value = Adodc1.Recordset.Fields("A0X03").Value
      End If
      'end 2025/8/28
      'Added by Lydia 2017/03/07 是否結清,限制輸入Y
      Case 7
         If IsNull(Adodc1.Recordset.Fields("A0X10").Value) Then
            Adodc1.Recordset.Fields("A0X10").Value = ""
         Else
            Adodc1.Recordset.Fields("A0X10").Value = "Y"
         End If
      'end 2017/03/07
      'Added by Lydia 2021/12/03 注意事項=>檢查Unicode字(自動更換)
      Case 8
         txtFM2 = Adodc1.Recordset.Fields("a1k34")
         If PUB_ChkUniText(Me, , , "TextBox", , True) = False Then
         End If
         Adodc1.Recordset.Fields("a1k34") = txtFM2
      'end 2021/12/03
   End Select
   Adodc1.Recordset.UpdateBatch
   SumShow
Checking:
   Exit Sub
End Sub

Private Sub DataGrid1_GotFocus()
Dim intCounter As Integer

   DataGrid1.col = 0
'   For intCounter = 1 To 6
   For intCounter = 1 To 7
      SendKeys "{RIGHT}"
   Next intCounter
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid1.col
'            Case 6
            Case 7
               SendKeys "{RIGHT}"
'            Case 7
            Case 8
               SendKeys "{RIGHT}"
'            Case 8
            Case 9
               SendKeys "{DOWN}"
               SendKeys "{LEFT}"
               SendKeys "{LEFT}"
          End Select
   End Select
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Text1 = Adodc1.Recordset.Fields("a0x02").Value
   'Add By Sindy 2015/10/20
   Text11 = "" & Adodc1.Recordset.Fields("a0x12").Value
   Text13 = "" & Adodc1.Recordset.Fields("a0x13").Value
   '2015/10/20 END
End Sub

'Added by Lydia 2021/12/03
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
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
'   Me.Width = 8880
'   Me.Height = 5700
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
   PUB_InitForm Me, 8880, 5900, strBackPicPath1
   'end 2021/12/07
   
   If strItemNo <> MsgText(601) Then
      Text3 = strItemNo
   Else
      Text3 = MsgText(601)
   End If
   Text1 = "X"
   Acc0x0Show
   OpenTable
   If adoacc0y0.RecordCount <> 0 Then
      FormShow
      SumShow
   End If
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(107)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strFagentNo As String    '2010/6/22 ADD BY SONIA
Dim bolSave As Boolean 'Added by Lydia 2015/10/19

   bolSave = False 'Added by Lydia 2015/10/19
   
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   If Option3.Value Then
      If Text8 = MsgText(601) Then
         tool3_enabled
         MsgBox MsgText(202), , MsgText(5)
         Cancel = 1
         Text8.SetFocus
         Exit Sub
      End If
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      'Added by Morgan 2020/9/1
      '檢查請款幣別是否與收款幣別不同
      intI = 0
      With Adodc1.Recordset
      .MoveFirst
      Do While Not .EOF
         If .Fields("a0x08") <> strCon2 Then
            'Modify by Amy 2021/02/01 改訊息 原:幣別不同提醒！請確認是否修改？
            intI = MsgBox("幣別不同提醒！請確認是否離開此畫面？" & vbCrLf & "按「是 」 : 離開畫面" & vbCrLf & "按「否」: 停留在原畫面", _
                    vbExclamation + vbYesNo + vbDefaultButton1)
            Exit Do
         End If
         .MoveNext
      Loop
      End With
      'Modify by Amy 2021/02/01 原:intI = vbYes
      If intI = vbNo Then
         tool3_enabled
         Cancel = 1
         Exit Sub
      End If
      'end 2020/9/1
      
      If Val(Text2) < 0 Then
         tool3_enabled
         MsgBox MsgText(89), , MsgText(5)
         Cancel = 1
         Exit Sub
      Else
         If Val(Text2) > 0 And Text10 = "" Then
            tool3_enabled
            MsgBox MsgText(132), , MsgText(5)
            Cancel = 1
            Exit Sub
         End If
      End If

     'Added by Lydia 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
     If PUB_ChkUniText(Me, , True, "TextBox", , True) = False Then '因為彈訊息後表單操作不順,改成自動更換文字不彈訊息
         Exit Sub
     End If
     'end 2021/12/03
    
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
      Screen.MousePointer = vbHourglass
      Frmacc2111_Save
      Acc0z0Save
      If Text10 <> MsgText(601) Then
         Acc120Save
      End If
      
      bolSave = True 'Added by Lydia 2015/10/19
      
      DeliverInform 'Added by Morgan 2015/11/30
      Screen.MousePointer = vbDefault
   End If
   '2010/6/22 ADD BY SONIA
   If Option1.Value = True Then
      strFagentNo = Text4
   Else
      If Option2.Value = True Then
         strFagentNo = Text6
      Else
         strFagentNo = Text8
      End If
   End If
   If strFagentNo = "Y46505000" Then
      MsgBox "廣東省商標的收據金額請更改為實際匯款的金額 !" & Err.Description, vbCritical
   End If
   '2010/6/22 END
   
   'Added by Lydia 2015/10/19 較早帳款未付即時催款
   If bolSave Then
      'Modified by Lydia 2015/11/05 +判斷是否寄發催款單
'      strSql = "select a1k01,a1k02 from acc1k0 where a1k28='" & strFagentNo & "' " & _
'               "and a1k02 < (select min(a1k02) from acc0z0,acc1k0 where a0z01='" & Text3.Text & "' and a0z02=a1k01(+)) " & _
'               "and a1k29 is null and a1k12||a1k17||a1k25 is null order by a1k02 "
      'modify by sonia 2021/11/8 原a1k02 < (...條件改為<=  收款單M11004785沒抓出X11013246(與X11013258同一天)
      strSql = "select a1k01,a1k02,decode(fa01,null,cu140,fa101) YN from acc1k0,fagent,customer " & _
               "where substr(a1k28,1,8)=fa01(+) and substr(a1k28,9,1)=fa02(+) and substr(a1k28,1,8)=cu01(+) and substr(a1k28,9,1)=cu02(+) " & _
               "and a1k28='" & strFagentNo & "' " & _
               "and a1k02 <= (select min(a1k02) from acc0z0,acc1k0 where a0z01='" & Text3.Text & "' and a0z02=a1k01(+)) " & _
               "and a1k29 is null and a1k12||a1k17||a1k25 is null order by a1k02 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Modified by Lydia 2015/11/05
         'Frmacc2110.m_A1k28 = strFagentNo
         'Modify by Amy 2020/09/18 原:"" & RsTemp.Fields("YN") <> "N",因不寄催款單改輸1-3
         'Modify by Amy 2020/10/05 不論是否有設「不寄催款單」都要彈
         'If IsNull(RsTemp.Fields("YN")) Then
         Frmacc2110.m_A1k28 = strFagentNo
         'end 2020/10/05
      End If
   End If
   'end 2015/10/19
   
   StatusClear
   strCon1 = "Y"
   strTrackMode = "" 'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序(清除)
   StatusClear
   tool1_enabled
   Frmacc2110.Show
   Set Frmacc2111 = Nothing
End Sub

Private Sub Text1_Change()
   If Len(Text1) <> 9 Then
      Exit Sub
   End If
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   '2013/10/18 MODIFY BY SONIA 改A1K05為A1K34
   'Modify By Sindy 2015/4/20 +a1k35
   'Modify By Sindy 2015/8/25 +a1k30,a1k29
   'Modify By Sindy 2020/12/9 +a1k37
   adoaccsum.Open "select a1k34,a1k35,a1k30,a1k29,a1k37 from acc1k0 where a1k01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   m_A1K37 = "2" 'Add By Sindy 2022/3/8 預設為2
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields("a1k34").Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = adoaccsum.Fields("a1k34").Value
      End If
      'Add By Sindy 2015/4/20
      If IsNull(adoaccsum.Fields("a1k35").Value) Then
         txtA1K35 = MsgText(601)
      Else
         txtA1K35 = adoaccsum.Fields("a1k35").Value
      End If
      '2015/4/20 END
      'Add By Sindy 2015/8/25
      If Val("" & adoaccsum.Fields("a1k30").Value) > 0 And "" & adoaccsum.Fields("a1k29").Value = "" Then
         MsgBox "已部分收款！", , MsgText(5)
      End If
      '2015/8/25 END
      'Add By Sindy 2020/12/9 帶入公司別
      If "" & adoaccsum.Fields("A1K37") <> "" Then
         Text13 = adoaccsum.Fields("A1K37")
         m_A1K37 = adoaccsum.Fields("A1K37") 'Add By Sindy 2022/3/8
      End If
      '2020/12/9
   Else
      Text12 = MsgText(601)
      txtA1K35 = MsgText(601) 'Add By Sindy 2015/4/20
   End If
   adoaccsum.Close
End Sub

Private Sub Text1_GotFocus()
   'MODIFY BY SONIA 2014/3/19
   'TextInverse Text1
   If Len(Text1) > 0 Then
      Text1.SelStart = 1
      Text1.SelLength = Len(Text1) - 1
   End If
   '2014/3/19 END
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2015/4/28
Private Sub Text11_LostFocus()
Dim strChkID As String
   
   Cancel = False
   '有輸入扣繳金額時,預設”請款單抬頭”為代理人1,2,3的名稱(中->英->日)
   If Val(Text11) > 0 And txtA1K35 = "" Then
      If Option1.Value = True Then strChkID = Text4
      If Option2.Value = True Then strChkID = Text6
      If Option3.Value = True Then strChkID = Text8
      If strChkID = "" Then
         MsgBox "代理人編號不可空白！", , MsgText(5)
         If Option1.Value = True Then Text4.SetFocus
         If Option2.Value = True Then Text6.SetFocus
         If Option3.Value = True Then Text8.SetFocus
         Cancel = True
         Exit Sub
      End If
      strExc(0) = "select nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) from customer" & _
                  " where substr('" & strChkID & "',1,8)=cu01 and substr('" & strChkID & "',9)=cu02" & _
                  " Union" & _
                  " select nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)) from fagent" & _
                  " where substr('" & strChkID & "',1,8)=fa01 and substr('" & strChkID & "',9)=fa02"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtA1K35 = RsTemp.Fields(0)
      End If
   End If
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2013/1/11
Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text4_Change()
   If Text4 = MsgText(601) Then
      Exit Sub
   End If
   Text5 = FagentQuery(Text4, 2)
   If Text5 = "" Then
      Text5 = CustomerQuery(Text4, 2)
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 = MsgText(601) Then
      Exit Sub
   End If
   If Len(Text4) = 6 Then
      Text4 = AfterZero(Text4)
   'Add by Morgan 2007/3/1 八碼時要補'0'
   ElseIf Len(Text4) = 8 Then
      Text4 = Text4 & "0"
   'End 2007/3/1
   End If
   
   Text5 = FagentQuery(Text4, 2)
   '2005/5/20 ADD BY SONIA
   If Text5 = "" Then
      Text5 = FagentQuery(Text4, 1)
   End If
   '2005/5/20 END
   If Text5 = "" Then
      Text5 = CustomerQuery(Text4, 2)
   End If
   '2005/5/20 ADD BY SONIA
   If Text5 = "" Then
      Text5 = CustomerQuery(Text4, 1)
   End If
   
   If ExistCheck("fagent", "fa01", Mid(Text4, 1, 8), Label4, False) = False Then
      If ExistCheck("customer", "cu01", Mid(Text4, 1, 8), Label4) = False Then
         Cancel = True
         Text4.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text6_Change()
   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   Text7 = FagentQuery(Text6, 2)
   If Text7 = "" Then
      Text7 = CustomerQuery(Text6, 2)
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   If Len(Text6) = 6 Then
      Text6 = AfterZero(Text6)
   'Add by Morgan 2007/3/1 八碼時要補'0'
   ElseIf Len(Text6) = 8 Then
      Text6 = Text6 & "0"
   'End 2007/3/1
   End If
   
   Text7 = FagentQuery(Text6, 2)
   '2005/5/20 ADD BY SONIA
   If Text7 = "" Then
      Text7 = FagentQuery(Text6, 1)
   End If
   '2005/5/20 END
   If Text7 = "" Then
      Text7 = CustomerQuery(Text6, 2)
   End If
   '2005/5/20 ADD BY SONIA
   If Text7 = "" Then
      Text7 = CustomerQuery(Text6, 1)
   End If
   '2005/5/20 END
   
   If ExistCheck("fagent", "fa01", Mid(Text6, 1, 8), Label5, False) = False Then
      If ExistCheck("customer", "cu01", Mid(Text6, 1, 8), Label5) = False Then
         Cancel = True
         Text6.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
'Add By Cheng 2003/09/03
Dim StrSQLa As String

On Error GoTo Checking
   adoacc0y0.CursorLocation = adUseClient
   adoacc0y0.Open "select * from acc0y0 where a0y01 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   'Modify By Cheng 2003/09/03
'   strSQLA = "select * from acc0x0 where a0x01 = '" & strItemNo & "' and a0x15 = '" & strUserNum & "' order by a0x14 desc"
   'Modified by Morgan 2021/4/21 因修改時資料的時間可能相同，增加 a0x02 asc 排序
   StrSQLa = "select * from acc0x0, acc1k0 where a0x02=a1k01(+) And a0x01 = '" & strItemNo & "' and a0x15 = '" & strUserNum & "' order by a0x14 desc,a0x02 asc"
   adoadodc1.Open StrSQLa, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   'Add By Sindy 2020/6/4 帶入公司別
   m_A1K37 = "2" 'Add By Sindy 2022/3/8 預設為2
   If adoadodc1.RecordCount > 0 Then
      If "" & adoadodc1.Fields("A1K37") <> "" Then
         Text13 = adoadodc1.Fields("A1K37")
         m_A1K37 = adoadodc1.Fields("A1K37") 'Add By Sindy 2022/3/8
      End If
   End If
   '2020/6/4
   If IsNull(adoacc0y0.Fields("a0y18").Value) = False Then
      Select Case adoacc0y0.Fields("a0y18").Value
         Case 1
            Option1.Value = True
         Case 2
            Option2.Value = True
         Case 3
            Option3.Value = True
      End Select
   End If
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
'Add By Cheng 2003/09/03
Dim StrSQLa As String

On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify By Cheng 2003/09/03
   'strSQLA = "select * from acc0x0 where a0x01 = '" & strItemNo & "' and a0x15 = '" & strUserNum & "' order by a0x14 desc"
   'Modified by Morgan 2021/4/21 因修改時資料的時間可能相同，增加 a0x02 asc 排序
   StrSQLa = "select * from acc0x0, acc1k0 where a0x02=a1k01(+) And a0x01 = '" & strItemNo & "' and a0x15 = '" & strUserNum & "' order by a0x14 desc,a0x02 asc"
   adoadodc1.Open StrSQLa, adoTaie, adOpenDynamic, adLockBatchOptimistic
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
   If IsNull(adoacc0y0.Fields("a0y07").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0y0.Fields("a0y07").Value
      If Len(Text4) = 6 Then
         Text4 = AfterZero(Text4)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text4) = 8 Then
         Text4 = Text4 & "0"
      'End 2007/3/1
      End If
      
      Text5 = FagentQuery(Text4, 2)
      '2005/5/20 ADD BY SONIA
      If Text5 = "" Then
         Text5 = FagentQuery(Text4, 1)
      End If
      '2005/5/20 END
      If Text5 = "" Then
         Text5 = CustomerQuery(Text4, 2)
      End If
      '2005/5/20 ADD BY SONIA
      If Text5 = "" Then
         Text5 = CustomerQuery(Text4, 1)
      End If
   End If
   If IsNull(adoacc0y0.Fields("a0y08").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc0y0.Fields("a0y08").Value
      If Len(Text6) = 6 Then
         Text6 = AfterZero(Text6)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text6) = 8 Then
         Text6 = Text6 & "0"
      'End 2007/3/1
      End If
      
      Text7 = FagentQuery(Text6, 2)
      '2005/5/20 ADD BY SONIA
      If Text7 = "" Then
         Text7 = FagentQuery(Text6, 1)
      End If
      '2005/5/20 END
      If Text7 = "" Then
         Text7 = CustomerQuery(Text6, 2)
      End If
      '2005/5/20 ADD BY SONIA
      If Text7 = "" Then
         Text7 = CustomerQuery(Text6, 1)
      End If
      '2005/5/20 END
   End If
   If IsNull(adoacc0y0.Fields("a0y09").Value) Then
      Text8 = MsgText(601)
      Text9 = MsgText(601)
   Else
      Text8 = adoacc0y0.Fields("a0y09").Value
      If Len(Text8) = 6 Then
         Text8 = AfterZero(Text8)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text8) = 8 Then
         Text8 = Text8 & "0"
      'End 2007/3/1
      End If
      
      Text9 = FagentQuery(Text8, 2)
      '2005/5/20 ADD BY SONIA
      If Text9 = "" Then
         Text9 = FagentQuery(Text8, 1)
      End If
      '2005/5/20 END
      If Text9 = "" Then
         Text9 = CustomerQuery(Text8, 2)
      End If
      '2005/5/20 ADD BY SONIA
      If Text9 = "" Then
         Text9 = CustomerQuery(Text8, 1)
      End If
      
   End If
   If IsNull(adoacc0y0.Fields("a0y10").Value) Then
      Command2.Enabled = True
      Text10 = MsgText(601)
   Else
      Command2.Enabled = False
      Text10 = adoacc0y0.Fields("a0y10").Value
   End If
   'Add By Sindy 2013/1/11
   If IsNull(adoacc0y0.Fields("a0y19").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = adoacc0y0.Fields("a0y19").Value
   End If
   '2013/1/11 End
End Sub

'*************************************************
'  儲存資料表( 請款單收款記錄資料)
'
'*************************************************
Private Sub Acc0x0Save()
Dim strMsg As String 'Added by Lydia 2025/01/06

On Error GoTo Checking
   adoacc1k0.CursorLocation = adUseClient
   '93.12.15 MODIFY BY SONIA 未銷帳才抓
   'adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & Text1 & "' and a1k12 is null and a1k29 is null", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & Text1 & "' and a1k12 is null and a1k25 is null and a1k29 is null", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '93.12.15 END
   If adoacc1k0.RecordCount <> 0 Then
      adoacc0x0.CursorLocation = adUseClient
      adoacc0x0.Open "select * from acc0x0 where a0x01 = '" & Text3 & "' and a0x02 = '" & Text1 & "' and a0x15 = '" & strUserNum & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc0x0.RecordCount <> 0 Then
         MsgBox MsgText(9), , MsgText(5)
         adoacc1k0.Close
         adoacc0x0.Close
         Exit Sub
      Else
         adoacc0x0.Close
         adoacc0z0.CursorLocation = adUseClient
         adoacc0z0.Open "select * from acc0z0 where a0z01 = '" & Text3 & "' and a0z02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc0z0.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            adoacc1k0.Close
            adoacc0z0.Close
            Exit Sub
         End If
         adoacc0z0.Close
      End If
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields("a0x01").Value = Text3
      Adodc1.Recordset.Fields("a0x02").Value = Text1
      If IsNull(adoacc1k0.Fields("a1k09").Value) Then
         Adodc1.Recordset.Fields("a0x03").Value = 0
      Else
         Adodc1.Recordset.Fields("a0x03").Value = adoacc1k0.Fields("a1k09").Value
      End If
      adocaseprogress.CursorLocation = adUseClient
      adocaseprogress.Open "select nvl(cpm03, cpm04) from caseprogress, casepropertymap where cp01 = cpm01 and cp10 = cpm02 and cp60 = '" & adoacc1k0.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocaseprogress.RecordCount <> 0 Then
         If IsNull(adocaseprogress.Fields(0).Value) Then
            Adodc1.Recordset.Fields("a0x04").Value = Null
         Else
            Adodc1.Recordset.Fields("a0x04").Value = adocaseprogress.Fields(0).Value
         End If
      Else
         Adodc1.Recordset.Fields("a0x04").Value = Null
      End If
      adocaseprogress.Close
      If IsNull(adoacc1k0.Fields("a1k08").Value) Then
         Adodc1.Recordset.Fields("a0x05").Value = 0
      Else
         Adodc1.Recordset.Fields("a0x05").Value = adoacc1k0.Fields("a1k08").Value
      End If
      'Modify By Sindy 2012/12/6 外幣折讓
'      If IsNull(adoacc1k0.Fields("a1k06").Value) Then
'         Adodc1.Recordset.Fields("a0x06").Value = 0
'      Else
'         Adodc1.Recordset.Fields("a0x06").Value = adoacc1k0.Fields("a1k06").Value
'      End If
      If IsNull(adoacc1k0.Fields("a1k31").Value) Then
         Adodc1.Recordset.Fields("a0x06").Value = 0
      Else
         Adodc1.Recordset.Fields("a0x06").Value = adoacc1k0.Fields("a1k31").Value
      End If
      '2012/12/6 End
      adoacc0y0.Close
      adoacc0y0.CursorLocation = adUseClient
      adoacc0y0.Open "select * from acc0y0 where a0y01 = '" & Text3 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc0y0.RecordCount <> 0 Then
         'Modify by Morgan 2010/4/26 改用畫面上的欄位判斷
         'If IsNull(adoacc0y0.Fields("a0y07").Value) Then
         '   adoacc0y0.Fields("a0y07").Value = adoacc1k0.Fields("a1k28").Value
         If Text4 = "" Then
            Text4 = adoacc1k0.Fields("a1k28").Value
         'end 2010/4/26
            adoacc0y0.Fields("a0y18").Value = 1
            'Modified by Lydia 2025/01/06 +Y55913000
            'If Text4 = "Y51371000" Then Text4 = adoacc1k0.Fields("a1k27").Value       '2012/7/25 add by sonia 婧瑄要求
            If Text4 = "Y51371000" Or Text4 = "Y55913000" Then
               Text4 = adoacc1k0.Fields("a1k27").Value
               strMsg = Text4
            End If
            'end 2025/01/06
         Else
            'Modify by Morgan 2010/4/26 改用畫面上的欄位判斷
            'If adoacc0y0.Fields("a0y07").Value <> adoacc1k0.Fields("a1k28").Value Then
            '   If IsNull(adoacc0y0.Fields("a0y08").Value) Then
            '      adoacc0y0.Fields("a0y08").Value = adoacc1k0.Fields("a1k28").Value
            If Text4 <> adoacc1k0.Fields("a1k28").Value Then
               If Text6 = "" Then
                   Text6 = adoacc1k0.Fields("a1k28").Value
            'end 2010/4/26
                  adoacc0y0.Fields("a0y18").Value = 2
                  'Modified by Lydia 2025/01/06 +Y55913000
                  'If Text6 = "Y51371000" Then Text6 = adoacc1k0.Fields("a1k27").Value      '2012/7/25 add by sonia 婧瑄要求
                  If Text6 = "Y51371000" Or Text6 = "Y55913000" Then
                     Text6 = adoacc1k0.Fields("a1k27").Value
                     strMsg = Text6
                  End If
                  'end 2025/01/06
               'Add by Morgan 2010/4/26
               ElseIf Text6 <> adoacc1k0.Fields("a1k28").Value Then
                  If Text8 = "" Then
                     Text8 = adoacc1k0.Fields("a1k28").Value
                     adoacc0y0.Fields("a0y18").Value = 3
                     'Modified by Lydia 2025/01/06 +Y55913000
                     'If Text8 = "Y51371000" Then Text8 = adoacc1k0.Fields("a1k27").Value      '2012/7/25 add by sonia 婧瑄要求
                     If Text8 = "Y51371000" Or Text8 = "Y55913000" Then
                        Text8 = adoacc1k0.Fields("a1k27").Value
                        strMsg = Text8
                     End If
                     'end 2025/01/06
                  End If
               'end 2010/4/26
               End If
            End If
         End If
         
         If Text4 = "Y51345000" Then Text4 = "Y51345010" 'add by sonia 2025/2/24 北京正理知識產權代理有限公司收據固定改Y51345010
         
         'Add by Morgan 2010/4/26
         adoacc0y0.Fields("a0y07").Value = Text4
         adoacc0y0.Fields("a0y08").Value = Text6
         'end 2010/4/26
         
         'Add By Sindy 2013/1/11
         adoacc0y0.Fields("a0y19").Value = Text14
         '2013/1/11 End
         
         If Text8 <> MsgText(601) Then
            adoacc0y0.Fields("a0y09").Value = Text8
         Else
            adoacc0y0.Fields("a0y09").Value = Null
         End If
         adoacc0y0.UpdateBatch
         adoacc0y0.Requery
      End If
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select sum(a0z04) from acc0z0 where a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "' and a0z01 <> '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            Adodc1.Recordset.Fields("a0x07").Value = 0
         Else
            Adodc1.Recordset.Fields("a0x07").Value = Val(Format(adoquery.Fields(0).Value, FAmount))
         End If
      Else
         Adodc1.Recordset.Fields("a0x07").Value = 0
      End If
      adoquery.Close
      'Modify by Amy 2021/01/29 列印幣別格式為 4.外幣+美金合計,且前畫面幣別與原始幣別不同,則「外幣金額」改顯示為「美金請款金額」
      If "" & adoacc1k0.Fields("A1K33") = "4" And m_Currency <> "" & adoacc1k0.Fields("a1k18") Then
            'Modified by Morgan 2023/3/15 整批列印或未單筆列印時a1k38不會更新，改用計算的(與請款單實際美金金額會有誤差)
            'Adodc1.Recordset.Fields("A0X11") = adoacc1k0.Fields("a1k38")
            Adodc1.Recordset.Fields("A0X11") = Trunc((adoacc1k0.Fields("a1k08") - Val("" & adoacc1k0.Fields("a1k31"))) * PUB_GetDNRate(adoacc1k0.Fields("a1k02"), adoacc1k0.Fields("a1k18")))
            'end 2023/3/15
      '2006/3/27 MODIFY BY SONIA 台幣收款時,外幣收款金額預設為台幣請款金額
      'Adodc1.Recordset.Fields("A0X11").Value = Val(Adodc1.Recordset.Fields("A0X05").Value) - Val(Adodc1.Recordset.Fields("A0X06").Value) - Val(Adodc1.Recordset.Fields("A0X07").Value)
      ElseIf strCon2 = "NTD" Then
         'Modify By Sindy 2012/12/6
         'Adodc1.Recordset.Fields("A0X11").Value = Val(adoacc1k0.Fields("a1k11").Value) - Val(Adodc1.Recordset.Fields("A0X06").Value) * Val(adoacc1k0.Fields("a1k10").Value) - Val(Adodc1.Recordset.Fields("A0X07").Value)
         Adodc1.Recordset.Fields("A0X11").Value = Val(adoacc1k0.Fields("a1k11").Value) - Val("" & adoacc1k0.Fields("a1k06").Value) - Val(Adodc1.Recordset.Fields("A0X07").Value)
         '2012/12/6 End
      '2010/6/18 ADD BY SONIA 人民幣收款時,人民幣的請款單外幣收款預設為人民幣請款金額
      ElseIf strCon2 = "RMB" And adoacc1k0.Fields("a1k18").Value = "RMB" Then
         '抓請款匯率
         'dblRate = PUB_GetUSXRate_1(Val(adoacc1k0.Fields("a1k02").Value), adoacc1k0.Fields("a1k18").Value)
         '計算請款幣別合計
         'Modify By Sindy 2012/12/6
         'Adodc1.Recordset.Fields("A0X11").Value = Format(((Val(adoacc1k0.Fields("a1k11").Value) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
         Adodc1.Recordset.Fields("A0X11").Value = Val(adoacc1k0.Fields("a1k08").Value)
         '2012/12/6 End
         '扣除折讓
         If Not IsNull(adoacc1k0.Fields("a1k31").Value) Then
            Adodc1.Recordset.Fields("A0X11").Value = Adodc1.Recordset.Fields("A0X11").Value - Val(adoacc1k0.Fields("a1k31").Value)
         End If
      '2010/6/18 END
      Else
         'Modify By Sindy 2012/12/6
         'Adodc1.Recordset.Fields("A0X11").Value = Val(Adodc1.Recordset.Fields("A0X05").Value) - Val(Adodc1.Recordset.Fields("A0X06").Value) - Val(Adodc1.Recordset.Fields("A0X07").Value)
         Adodc1.Recordset.Fields("A0X11").Value = Val(adoacc1k0.Fields("a1k08").Value) - Val(Adodc1.Recordset.Fields("A0X06").Value) - Val(Adodc1.Recordset.Fields("A0X07").Value)
         '2012/12/6 End
      End If
      '2006/3/27 END
      '2010/3/15 add by sonia 請款台幣應扣除折讓 X09813703, 故暫存檔加A0X16,畫面A1k11改為A0X16
      'Modify By Sindy 2012/12/6
      'Adodc1.Recordset.Fields("A0X16").Value = Val(adoacc1k0.Fields("a1k11").Value) - Val(Adodc1.Recordset.Fields("A0X06").Value) * Val(adoacc1k0.Fields("a1k10").Value)
      Adodc1.Recordset.Fields("A0X16").Value = Val(adoacc1k0.Fields("a1k11").Value) - Val("" & adoacc1k0.Fields("a1k06").Value)
      '2012/12/6 End
      '2010/3/15 END
      Adodc1.Recordset.Fields("A0X09").Value = Val(Format((Val(Adodc1.Recordset.Fields("A0X11").Value)) * Val(strCon3), FAmount))
      'add by sonia 2025/8/28 M111403799收X11404811~X11404817台幣金額A0X09<規費A0X03時改預設規費A0X03
      If Adodc1.Recordset.Fields("A0X09").Value < Adodc1.Recordset.Fields("A0X03").Value Then
         Adodc1.Recordset.Fields("A0X09").Value = Adodc1.Recordset.Fields("A0X03").Value
      End If
      'end 2025/8/28
      Adodc1.Recordset.Fields("A0X10").Value = "Y"
      If strCon2 = MsgText(601) Then
         Adodc1.Recordset.Fields("a0x08").Value = Null
      Else
         '2013/8/19 modify by sonia
         'Adodc1.Recordset.Fields("a0x08").Value = strCon2
         If strCon2 = "NTD" Then
            Adodc1.Recordset.Fields("a0x08").Value = strCon2
         Else
            Adodc1.Recordset.Fields("a0x08").Value = adoacc1k0.Fields("a1k18").Value
         End If
         '2013/8/19 end
      End If
      If Text11 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a0x12").Value = Val(Text11)
      Else
         Adodc1.Recordset.Fields("a0x12").Value = 0
      End If
      If Text13 <> MsgText(601) Then
         Adodc1.Recordset.Fields("a0x13").Value = Text13
      Else
         Adodc1.Recordset.Fields("a0x13").Value = Null
      End If
      Adodc1.Recordset.Fields("a0x14").Value = ServerTime
      Adodc1.Recordset.Fields("a0x15").Value = strUserNum
      'Modified by Morgan 2023/8/1
      'adoacc1k0.Fields("a1k35").Value = Text12 'Added by Lydia 2021/12/03 注意事項
      adoacc1k0.Fields("a1k34").Value = Text12
      'end 2023/8/1
      'Add By Sindy 2015/4/28
      adoacc1k0.Fields("a1k35").Value = txtA1K35
      adoacc1k0.UpdateBatch
      '2015/4/28 END
      Adodc1.Recordset.UpdateBatch
      AdodcRefresh
      'Added by Lydia 2025/01/06
      If strMsg <> "" Then
         MsgBox "請確認以下列印對象Email是否包含請款對象:" & vbCrLf & strMsg
      End If
      'end 2025/01/06
   Else
      MsgBox MsgText(28), , MsgText(5)
   End If
   adoacc1k0.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)

   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         'Added by Lydia 2021/12/03 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'Added by Lydia 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "TextBox") = False Then
             Exit Sub
         End If
         'end 2021/12/03
         If Text10 <> MsgText(601) Then
            MsgBox MsgText(80), , MsgText(5)
            Exit Sub
         End If
         'Add By Sindy 2015/4/28
         '有輸入扣繳金額時,請款單抬頭不可空白
         If Val(Text11) > 0 And Trim(txtA1K35) = "" Then
            MsgBox "因有扣繳金額「請款單抬頭」不可空白！", , MsgText(5)
            txtA1K35.SetFocus
            Exit Sub
         End If
         'Add By Sindy 2015/10/20
         '公司別不可輸入J公司
         If Trim(Text13) = "J" Then
            MsgBox "公司別不可輸入J公司！", , MsgText(5)
            Text13.SetFocus
            Exit Sub
         End If
         '2015/10/20 END
         Cancel = False
         Call Text11_LostFocus
         If Cancel = True Then
            Exit Sub
         End If
         Cancel = False
         Call txtA1K35_Validate(Cancel)
         If Cancel = True Then
            txtA1K35.SetFocus
            Exit Sub
         End If
         '2015/4/28 END
         Acc0x0Save
         FormShow
         SumShow
         Text1 = "X"
         Text11 = ""
         Text12 = ""
         Text13 = ""
         txtA1K35 = "" 'Add By Sindy 2015/4/20
         Text1.SetFocus
         'MODIFY BY SONIA 2014/3/19
         'TextInverse Text1
         Text1_GotFocus
         '2014/3/19 END
   End Select
   KeyEnter KeyCode
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(107)
End Sub

'*************************************************
'  刪除資料表
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   
   'MODIFY BY SONIA 2014/3/28 改同下面,否則若未重點GRID依當時停住的資料刪除,畫面刪除但資料卻沒更新
   'adoTaie.Execute "update acc1k0 set a1k29 = null, a1k30 = 0 where a1k01 = '" & Text1 & "'"
   'adoTaie.Execute "delete from acc0z0 where a0z01 = '" & Text3 & "' and a0z02 = '" & Text1 & "'"
   'Modify By Sindy 2021/3/15 a1k30不可以直接更新為0,因有可能有部分收款
   'adoTaie.Execute "update acc1k0 set a1k29 = null, a1k30 = 0 where a1k01 = '" & Adodc1.Recordset.Fields("a0x02").Value & "'"
   adoTaie.Execute "delete from acc0z0 where a0z01 = '" & Text3 & "' and a0z02 = '" & Adodc1.Recordset.Fields("a0x02").Value & "'"
   adoTaie.Execute "update acc1k0 set a1k29 = null where a1k01 = '" & Adodc1.Recordset.Fields("a0x02").Value & "'"
   'Add By Sindy 2021/3/29
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   '2021/3/29 END
   adoquery.Open "select sum(nvl(a0z04, 0) * nvl(a0y04, 0)),sum(nvl(a0z12, 0)) from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & Adodc1.Recordset.Fields("a0x02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      'Modify By Sindy 2021/3/29 Val("" & adoquery.Fields(0).Value) ==> + "" &
      adoTaie.Execute "update acc1k0 set a1k30 = " & Val("" & adoquery.Fields(0).Value) & " where a1k01 = '" & Adodc1.Recordset.Fields("a0x02").Value & "'"
   Else
      adoTaie.Execute "update acc1k0 set a1k30 = 0 where a1k01 = '" & Adodc1.Recordset.Fields("a0x02").Value & "'"
   End If
   adoquery.Close
   '2021/3/15 END
   '2014/3/28 END
   
'   Ken 92/09/09 變更刪除方式
'   Adodc1.Recordset.Delete
'   Adodc1.Recordset.UpdateBatch
   adoTaie.Execute "delete from acc0x0 where a0x02 = '" & Adodc1.Recordset.Fields("a0x02").Value & "' and a0x15 = '" & strUserNum & "'"
   'Add By Sindy 2015/12/31
   adoTaie.Execute "delete from acc1v0 where a1v02 = '" & Adodc1.Recordset.Fields("a0x02").Value & "'"
   '2015/12/31 END
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Morgan 2010/4/22
'依照點數分配資料產生分錄
Private Sub Acc0z0SaveNew(ByVal strA0x01 As String, ByVal strA0x02 As String, ByVal stra1p22 As String, ByVal stra1p27 As String, ByRef W_strSerialNo As String)
   Dim strCaseNo As String '本所案號
   Dim strCaseProperty As String '案件性質
   Dim strSalesMan As String '智權人員
   Dim strCurrency As String '請款幣別
   Dim strExchange As String '請款匯率
   Dim strSerialNo As String '分錄序次
   Dim strSystemType As String '系統別
   Dim strAccNo As String '科目
   Dim strDept As String '承辦人會計部門
   Dim strEngDept As String '承辦人部門
   Dim strSalesDept As String '智權人員部門
   Dim strCustNo As String '客戶編號
   Dim strProperty As String '案件性質碼
   Dim strR As String '收入科目
   Dim strF As String '規費科目
   Dim strNation As String '申請國家
   Dim bolXFee As Boolean '服務費是否含出庭費
   Dim bolXFeeDone As Boolean '出庭費是否已扣除
   Dim strCP09 As String '收文號
   Dim strA1p30x As String '對沖-其他
   Dim strAmt As String '分錄金額
   Dim strAmtTot As String '收款金額
   Dim strA1p14 As String '摘要
   Dim strA1p08 As String '借方金額
   Dim strAmtRest As String '未分配金額
   Dim adoAcc0x0_1 As ADODB.Recordset
   Dim adoacc1n0 As ADODB.Recordset
   Dim adoCP As ADODB.Recordset
   Dim strPtTot As String '請款單總點數
   Dim strA1p16s As String '智權人代碼清單
   Dim strSerialNoFrom As String '分錄序次起號
   Dim strNetAmount As String '可分配金額(收款點數會大於請款點數)
   Dim strShareP As String
   Dim strShareT As String
   Dim strShareL As String
   Dim strShareFCP As String
   Dim strShareFCT As String
   Dim strShareFCL As String
   Dim strSharePointMemo As String '跨部門點數分配摘要
   Dim strMemoFrom As String '本次更新項次
   'Add by Morgan 2011/10/7
   Dim bolLawyerGuei As Boolean '是否委任律師為桂律師
   Dim strLawyerName As String '律師
   Dim strA1P30 As String '其他對沖
   Dim strA1P16 As String '業務對沖
   Dim strA1p30_new As String 'Add By Sindy 2015/12/31
   Dim strCP09_Min As String 'Added by Morgan 2016/1/7 最小收文號
   Dim dblTotAmt As Double 'Add By Sindy 2018/11/2 台幣已收款金額
   Dim dblTotA0Z12 As Double 'Add By Sindy 2018/11/5 扣繳金額
   Dim strE As String, strG2 As String, strD As String, strG1 As String, strCustNo_L As String 'Added by Morgan 2021/4/12
   
   strExc(0) = "select * from acc0x0,acc1k0 where a0x01='" & strA0x01 & "' and a0x02='" & strA0x02 & "' and a1k01(+)=a0x02"
   intI = 1
   Set adoAcc0x0_1 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   With adoAcc0x0_1
      strCaseNo = "" & .Fields("a1k13").Value & .Fields("a1k14").Value & .Fields("a1k15").Value & .Fields("a1k16").Value
      strSystemType = "" & .Fields("a1k13").Value
              
      '設定請款匯率幣別
      strCurrency = "" & .Fields("a1k18").Value
      strExchange = "" & .Fields("a1k10").Value
      
      '新增acc0z0
      If adoacc0z0.State = adStateOpen Then adoacc0z0.Close
      adoacc0z0.CursorLocation = adUseClient
      adoacc0z0.Open "select * from acc0z0 where a0z01 = '" & .Fields("a0x01").Value & "' and a0z02 = '" & .Fields("a0x02").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc0z0.RecordCount = 0 Then
         adoacc0z0.AddNew
      End If
      adoacc0z0.Fields("a0z01").Value = .Fields("a0x01").Value
      adoacc0z0.Fields("a0z02").Value = .Fields("a0x02").Value
      'Modify By Sindy 2014/12/10 改存前一畫面(Frmacc2110)的幣別
      If m_Currency <> "" Then
         adoacc0z0.Fields("a0z03").Value = m_Currency
      Else
         adoacc0z0.Fields("a0z03").Value = Null
      End If
'      If IsNull(.Fields("a0x08").Value) Then
'         adoacc0z0.Fields("a0z03").Value = Null
'      Else
'         adoacc0z0.Fields("a0z03").Value = .Fields("a0x08").Value
'      End If
      If IsNull(.Fields("a0x11").Value) Then
         adoacc0z0.Fields("a0z04").Value = 0
      Else
         adoacc0z0.Fields("a0z04").Value = .Fields("a0x11").Value
      End If
      If IsNull(.Fields("a0x12").Value) Then
         adoacc0z0.Fields("a0z12").Value = 0
      Else
         adoacc0z0.Fields("a0z12").Value = .Fields("a0x12").Value
      End If
      If IsNull(.Fields("a0x13").Value) Then
         adoacc0z0.Fields("a0z13").Value = Null
      Else
         adoacc0z0.Fields("a0z13").Value = .Fields("a0x13").Value
      End If
      adoacc0z0.Fields("a0z06").Value = strSrvDate(2)
      adoacc0z0.Fields("a0z07").Value = ServerTime
      adoacc0z0.Fields("a0z08").Value = strUserNum
      adoacc0z0.UpdateBatch
      adoacc0z0.Close
      
      '更新台幣收款金額
      'Modify By Sindy 2018/11/5 + ,sum(nvl(a0z12, 0))
      strExc(0) = "select nvl(sum(nvl(a0z04, 0) * nvl(a0y04, 0)),0),sum(nvl(a0z12, 0)) from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & .Fields("a1k01").Value & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      dblTotAmt = 0 'Add By Sindy 2018/11/2 台幣已收款金額
      dblTotA0Z12 = 0 'Add By Sindy 2018/11/5 扣繳金額
      If intI = 1 Then
         dblTotAmt = Val(RsTemp.Fields(0).Value) 'Add By Sindy 2018/11/2 台幣已收款金額
         dblTotA0Z12 = Val(RsTemp.Fields(1).Value) 'Add By Sindy 2018/11/5 扣繳金額
         adoTaie.Execute "update acc1k0 set a1k30 = " & Val(RsTemp.Fields(0).Value) & " where a1k01 = '" & .Fields("a1k01").Value & "'"
      End If
      
      '設定分錄欄位預設值
      'Modify By Sindy 2015/5/1 nvl(cpm03, nvl(cpm10, cpm13)) ==> DECODE(PA09,'000',CPM03,CPM04)
      strExc(0) = "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(PA09,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, pa09 as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, patent, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and cp60 = '" & .Fields("a0x02").Value & "' union " & _
         "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(TM10,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, tm10 as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, trademark, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and cp60 = '" & .Fields("a0x02").Value & "' union " & _
         "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(LC15,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, lc15 as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, lawcase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and cp60 = '" & .Fields("a0x02").Value & "' union " & _
         "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, nvl(cpm03, cpm04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, null as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, hirecase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and cp60 = '" & .Fields("a0x02").Value & "' union " & _
         "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(SP09,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, sp09 as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, servicepractice, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and cp60 = '" & .Fields("a0x02").Value & "' order by cp09 asc"
      intI = 1
      Set adoCP = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With adoCP
         strCP09_Min = .Fields("cp09")  'Added by Morgan 2016/1/7
         strCustNo = "" & .Fields("Custno").Value
         strCaseProperty = "" & .Fields("Property").Value
         strSalesMan = "" & .Fields("cp13").Value
         'Modified by Morgan 2012/9/19 秀玲說會有調區問題改抓 cp12
         'strSalesDept = "" & .Fields("st03").Value
         strSalesDept = "" & .Fields("cp12").Value
         
         strNation = "" & .Fields("nation").Value
         If strNation <> "000" Then
            strR = "" & .Fields("cpm24").Value
            strF = "" & .Fields("cpm25").Value
         Else
            strR = "" & .Fields("cpm11").Value
            strF = "" & .Fields("cpm12").Value
         End If
         
         'Add by Morgan 2011/10/11
         bolLawyerGuei = False
         strLawyerName = ""
         strA1P30 = ""
         If InStr(.Fields("cpm03"), "委任律師") > 0 Then
            If .Fields("cp14") = "76012" Then bolLawyerGuei = True
            strA1P30 = "" & .Fields("cp14") 'A1P30對沖代號(其它)存承辦人編號
            strExc(0) = "select s1.st02,s2.st02 from caseprogress,staff s1,CaseLawer,staff s2" & _
               " where cp09='" & .Fields("cp09") & "' and s1.st01(+)=cp14 and cl01(+)=cp09 and s2.st01(+)=cl02"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(RsTemp.Fields(0)) Then
                  strLawyerName = "/" & RsTemp.Fields(0)
               End If
               If Not IsNull(RsTemp.Fields(1)) Then
                  strLawyerName = strLawyerName & "/" & RsTemp.Fields(1)
               End If
            End If
         End If
         'end 2011/10/11
         
         End With
      Else
         Exit Sub
      End If
      
      strA1p14 = strCaseNo & "/" & strCaseProperty
      'Modify By Sindy 2017/3/10 從有扣繳時程式段移出來,均帶此字樣
      'Modify By Sindy 2015/11/5
      If .Fields("a1k35").Value <> "" Then strA1p14 = Mid(.Fields("a1k35").Value, 1, 6) & "/" & strA1p14
      strA1p30x = Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4)
      strA1p30_new = Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & IIf(Trim(strA1P30) <> "", "/" & strA1P30, "")
      '2015/11/5 END
      '2017/3/10 END
      '有扣繳時產生預付稅捐分錄
      If .Fields("a0x12").Value <> 0 Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         If Option1.Value Then
            strCustNo = Text4
         ElseIf Option2.Value Then
            strCustNo = Text6
         Else
            strCustNo = Text8
         End If
         'Modify By Sindy 2015/11/5 取消特殊客戶
'         '特殊客戶-->對沖-其他
'         If PUB_CHKCUST(strCustNo) = True Then
'            strA1p30x = StrToStr(m_YEAR & Mid(m_NAME, 1, 4), 5)
'         Else
'            strA1p30x = ""
'         End If
         
'         'Modify By Sindy 2015/11/5
'         strA1p14 = Mid(.Fields("a1k35").Value, 1, 6) & "/" & strA1p14
'         strA1p30x = Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4)
'         strA1p30_new = Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & IIf(Trim(strA1P30) <> "", "/" & strA1P30, "")
'         '2015/11/5 END
         'modify by sonia 2021/2/25 +a1p23存A1K01
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                         ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                         ", A1P30, a1p23) values " & _
                         "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '1203', " & Val(.Fields("a0x12").Value) & ", 0" & _
                         ", '" & strA1p14 & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format((Val(.Fields("a0x12").Value) / Val(strExchange)), FAmount)) & _
                         ",'" & strA1p30x & "','" & .Fields("a1k01").Value & "')"
         
'         'Add By Sindy 2015/4/28 有輸入扣繳金額時,同時寫入acc1v0
'         'Modify By Sindy 2015/11/5 檢查資料是否已存在
'         strExc(0) = "select a1v02 from acc1v0 where a1v02='" & .Fields("a0x02").Value & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 0 Then
'            adoTaie.Execute "insert into ACC1V0 (a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v18,a1v12,a1v13)" & _
'                            " values('" & adoCP.Fields("cp09") & "','" & .Fields("a0x02").Value & "'," & IIf(Text13 <> "", Text13, "GetA0k11('" & adoCP.Fields("cp09") & "')") & _
'                            "," & Val(.Fields("a0x12").Value) & ",'" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'," & Val(.Fields("a0x12").Value) & ",0" & _
'                            "," & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & ",'1','" & strCaseProperty & "','" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "')"
'         Else
'            'Modify By Sindy 2018/11/2 多次收款:A1V04=(分次收款總額-規費)/10
'            '   Val(.Fields("a0x12").Value) => Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0)
'            adoTaie.Execute "update ACC1V0" & _
'                            " set a1v01='" & adoCP.Fields("cp09") & "'" & _
'                            ",a1v03=" & IIf(Text13 <> "", Text13, "GetA0k11('" & adoCP.Fields("cp09") & "')") & _
'                            ",a1v04=" & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0) & _
'                            ",a1v05='" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'" & _
'                            ",a1v06=" & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0) & _
'                            ",a1v07=0" & _
'                            ",a1v09=" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & _
'                            ",a1v18='1'" & _
'                            ",a1v12='" & strCaseProperty & "'" & _
'                            ",a1v13='" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "'" & _
'                            " where a1v02='" & .Fields("a0x02").Value & "'"
'         End If
'         '2015/4/28 END
'      Else
'         'Add By Sindy 2015/11/3 有輸入請款單抬頭時,不管是否有輸入扣繳金額都要寫入acc1v0
'         If Trim(.Fields("a1k35").Value) <> "" Then
'            strExc(0) = "select a1v02 from acc1v0 where a1v02='" & .Fields("a0x02").Value & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 0 Then
'               'modify by sonia 2017/7/7 收款有扣繳金額a1v18才存'1'
'               'adoTaie.Execute "insert into ACC1V0 (a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v18,a1v12,a1v13)" & _
'                               " values('" & adoCP.Fields("cp09") & "','" & .Fields("a0x02").Value & "'," & IIf(Text13 <> "", Text13, "GetA0k11('" & adoCP.Fields("cp09") & "')") & _
'                               "," & Round((Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value)) / 10, 0) & ",'" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "',0," & Round((Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value)) / 10, 0) & _
'                               "," & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & ",'1','" & strCaseProperty & "','" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "')"
'               'Modify By Sindy 2018/11/2 多次收款:A1V04=(分次收款總額-規費)/10
'               '   Round((Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value)) / 10, 0) => Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0)
'               adoTaie.Execute "insert into ACC1V0 (a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v12,a1v13)" & _
'                               " values('" & adoCP.Fields("cp09") & "','" & .Fields("a0x02").Value & "'," & IIf(Text13 <> "", Text13, "GetA0k11('" & adoCP.Fields("cp09") & "')") & _
'                               "," & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0) & ",'" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "',0," & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0) & _
'                               "," & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & ",'" & strCaseProperty & "','" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "')"
'            Else
'               'modify by sonia 2017/7/7 收款有扣繳金額a1v18才存'1'
'               'adoTaie.Execute "update ACC1V0" & _
'                            " set a1v01='" & adoCP.Fields("cp09") & "'" & _
'                            ",a1v03=" & IIf(Text13 <> "", Text13, "GetA0k11('" & adoCP.Fields("cp09") & "')") & _
'                            ",a1v04=" & Round((Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value)) / 10, 0) & _
'                            ",a1v05='" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'" & _
'                            ",a1v06=0" & _
'                            ",a1v07=" & Round((Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value)) / 10, 0) & _
'                            ",a1v09=" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & _
'                            ",a1v18='1'" & _
'                            ",a1v12='" & strCaseProperty & "'" & _
'                            ",a1v13='" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "'" & _
'                            " where a1v02='" & .Fields("a0x02").Value & "'"
'               'Modify By Sindy 2018/11/2 多次收款:A1V04=(分次收款總額-規費)/10
'               '   Round((Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value)) / 10, 0) => Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0)
'               adoTaie.Execute "update ACC1V0" & _
'                            " set a1v01='" & adoCP.Fields("cp09") & "'" & _
'                            ",a1v03=" & IIf(Text13 <> "", Text13, "GetA0k11('" & adoCP.Fields("cp09") & "')") & _
'                            ",a1v04=" & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0) & _
'                            ",a1v05='" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'" & _
'                            ",a1v06=0" & _
'                            ",a1v07=" & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0) & _
'                            ",a1v09=" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & _
'                            ",a1v12='" & strCaseProperty & "'" & _
'                            ",a1v13='" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "'" & _
'                            ",a1v18=null where a1v02='" & .Fields("a0x02").Value & "'"
'            End If
'         End If
'         '2015/11/3 END
      End If
      'Modify By Sindy 2018/11/5
      '有扣繳時產生預付稅捐分錄
      'Add By Sindy 2015/11/3 有輸入請款單抬頭時,不管是否有輸入扣繳金額都要寫入acc1v0
      If .Fields("a0x12").Value <> 0 Or dblTotA0Z12 > 0 Or Trim(.Fields("a1k35").Value) <> "" Then
         'Add By Sindy 2015/4/28 有輸入扣繳金額時,同時寫入acc1v0
         'Modify By Sindy 2015/11/5 檢查資料是否已存在
         strExc(0) = "select a1v02 from acc1v0 where a1v02='" & .Fields("a0x02").Value & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            adoTaie.Execute "insert into ACC1V0 (a1v01,a1v02)" & _
                            " values('" & adoCP.Fields("cp09") & "','" & .Fields("a0x02").Value & "')"
         End If
         'Modify By Sindy 2022/3/8
         '",a1v03=" & IIf(Text13 <> "", Text13, "GetA0k11('" & adoCP.Fields("cp09") & "')")
         '=> ",a1v03=" & m_A1K37
         adoTaie.Execute "update ACC1V0 set" & _
                         " a1v01='" & adoCP.Fields("cp09") & "'" & _
                         ",a1v03=" & m_A1K37 & _
                         ",a1v05='" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'" & _
                         ",a1v09=" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & _
                         ",a1v12='" & strCaseProperty & "'" & _
                         ",a1v13='" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "'" & _
                         " where a1v02='" & .Fields("a0x02").Value & "'"
         'Modify By Sindy 2018/11/2 多次收款:A1V04=(分次收款總額-規費)/10
         '   Val(.Fields("a0x12").Value) => Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0)
         '應該是A1V04 = (分次收款總額 - 規費) / 10
         'A1V06=分次收款A0Z12合計，
         '若A1V06>0則A1V07=0,A1V18='1'
         '  A1V06=0則A1V07=A1V04,A1V18=NULL
         If .Fields("a0x12").Value <> 0 Or dblTotA0Z12 > 0 Then
            'Modify By Sindy 2019/5/10 + Round(dblTotA0Z12, 0)
            ' a1v04=" & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0)
            ',a1v06=" & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0)
            adoTaie.Execute "update ACC1V0 set" & _
                            " a1v04=" & Round(dblTotA0Z12, 0) & _
                            ",a1v06=" & Round(dblTotA0Z12, 0) & _
                            ",a1v07=0" & _
                            ",a1v18='1'" & _
                            " where a1v02='" & .Fields("a0x02").Value & "'"
         Else
            adoTaie.Execute "update ACC1V0 set" & _
                            " a1v04=" & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0) & _
                            ",a1v06=0" & _
                            ",a1v07=" & Round((dblTotAmt - Val(.Fields("a0x03").Value)) / 10, 0) & _
                            ",a1v18=null" & _
                            " where a1v02='" & .Fields("a0x02").Value & "'"
         End If
      End If
      '2018/11/5 END
      
      strA1p14 = strA1p14 & strLawyerName 'Add by Morgan 2011/10/11
      
      '收入科目
      '先產生收入再產生規費
      If Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value) > 0 Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         W_strSerialNo = strSerialNo
         
         '台灣案專利商標出庭費控管
         If strNation = "000" And Val(adoacc0x0("a1k02")) >= 960815 And Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value) >= 10000 Then
            With adoCP
            .MoveFirst
            Do While Not .EOF
            strProperty = .Fields("cp10")
            strCP09 = .Fields("cp09")
            '專利
            If (strSystemType = "P" Or strSystemType = "FCP") Then
               If InStr("211,212", strProperty) > 0 Then
                  'Modified by Morgan 2018/12/13
                  'bolXFee = True
                  If PUB_ChkNoXFee("" & .Fields("cp09")) = False Then
                     bolXFee = True
                  End If
                  'end 2018/12/13
                  Exit Do
               ElseIf InStr("503,507,506", strProperty) > 0 Then
                  strExc(0) = "select * from caseprogress a where cp09='" & strCP09 & "'" & _
                     " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('211','212'))" & _
                     " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('211','212') and b.cp16>0)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolXFee = True
                     Exit Do
                  End If
               End If
            '商標
            'Modify by Morgan 2010/11/8 FCT不扣
            'ElseIf (strSystemType = "T" Or strSystemType = "FCT") Then
            ElseIf (strSystemType = "T") Then
               If InStr("204,205", strProperty) > 0 Then
                  bolXFee = True
                  '2013/8/19 ADD BY SONIA 葉經理說訴願的言詞辯論為商標處的人處理,故不扣出庭費T-182351
                  strExc(0) = "select b.cp10,c.cp10 from caseprogress a,caseprogress b,caseprogress c where a.cp09='" & strCP09 & "'" & _
                     " and a.cp43=b.cp09(+) and b.cp43=c.cp09(+)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If "" & RsTemp.Fields(0) = "401" Or "" & RsTemp.Fields(1) = "401" Then bolXFee = False
                  End If
                  '2013/8/19 END
                  Exit Do
               ElseIf InStr("403,408,407", strProperty) > 0 Then
                  strExc(0) = "select * from caseprogress a where cp09='" & strCP09 & "'" & _
                     " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('204','205'))" & _
                     " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('204','205') and b.cp16>0)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolXFee = True
                     Exit Do
                  End If
               End If
            End If
            .MoveNext
            Loop
            End With
         End If
         
         strAmtTot = Val(Format(Val(.Fields("a0x09").Value) - Val(.Fields("a0x03").Value), FAmount))
         strNetAmount = strAmtTot
         '是否要扣出庭費10000
         If bolXFee = True Then
            strAmtTot = Val(strAmtTot - 10000)
         End If
         strAmtRest = strAmtTot
         
         'Added by Morgan 2021/4/12 FCT B2類案源
         'Modified by Morgan 2021/4/29 行政訴訟上訴及行政訴訟上訴答辯不必出庭
         If strLOS02 = "B2" And Left(strCaseNo, 3) = "FCT" Then
            strCustNo_L = strCustNo
            'E:規費
            strE = Trunc(.Fields("a0x03").Value)
            If bolB2NeeCourt = True Then
               'G2:FCT收入-爭議 417202= A(收款金額) - B(商標處出庭費5000) - E(規費) - F(律師出庭費5000)
               'Modified by Morgan 2022/12/13 出庭費改抓設定，但暫不考慮多律師出庭狀況--婉莘
               'strG2 = .Fields("a0x09").Value - 5000 - strE - 5000
               strG2 = .Fields("a0x09").Value - 5000 - strE - lngCurtFee
               'end 2022/12/13
            Else
               'G2:FCT收入-爭議 417202= A(收款金額) - E(規費)
               strG2 = .Fields("a0x09").Value - strE
            End If
            
            'Added by Morgan 2021/10/25
            strAmtTot = strG2
            strNetAmount = strAmtTot 'FCT分配比例以扣除出庭費後金額計算 Ex:M11004546--婉莘
            strAmtRest = strAmtTot
            'end 2021/10/25
         End If
         'end 2021/4/12
         
      'Added by Morgan 2016/1/7
      '收款
      'Modified by Morgan 2020/4/15 請款單日期>=智慧所更名日者改回依案件性質表設定之科目收入；
      If DBDATE(.Fields("a1k02")) < 智慧所更名日 And Val(strCon1) > 1050000 And (Left(strR, 4) = "4141" Or Left(strR, 4) = "4161" Or Left(strR, 4) = "4181") Then
         'Added by Morgan 2016/2/17
         'modify by sonia 2021/3/12 加傳日期
         strSalesMan = SalesNoToAccSales(strSalesMan, strAccNo, strCaseNo, Val(strCon1))
         If strSalesMan = "" Then
            strSalesMan = "M0100"
         End If
         'end 2016/2/17
         
         InsertLawACC1P0 "1", "F", strSerialNo, strItemNo, strR, IIf(strDept = "", MsgText(55), strDept), 0, Val(strAmtTot), "", "", "", "", "", _
         strA1p14 & IIf(strR = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, "0.00"), ""), _
         strCustNo, strSalesMan, strCaseNo, strCon1, strCurrency, strExchange, "" & .Fields("a0x11").Value, Replace(stra1p22, "'", ""), "", "", "", "", Replace(stra1p27, "'", ""), strA1p30_new, "", strCP09_Min
      Else
      'Added by Morgan 2016/1/7
         
         '部門,點數
         'modify by sonia 2021/1/22 110年起法務案改依業務點數a1n02='1'帶科目
         'strExc(0) = "select a0910,max(st03) st03,sum(a1n05) pts from acc1n0,staff,acc090" & _
            " where a1n01='" & .Fields("a1k01") & "' and a1n02='2' and st01(+)=a1n04" & _
            " and a0901(+)=st15 group by a0910 order by 3 desc,2,1"
         If Val(strCon1) > 1100000 And InStr(Mid(strCaseNo, 1, Len(strCaseNo) - 9), "L") > 0 Then
            strExc(0) = "select a0910,max(st03) st03,sum(a1n05) pts from acc1n0,staff,acc090" & _
               " where a1n01='" & .Fields("a1k01") & "' and a1n02='1' and st01(+)=a1n04" & _
               " and a0901(+)=st15 group by a0910 order by 3 desc,2,1"
         Else
            strExc(0) = "select a0910,max(st03) st03,sum(a1n05) pts from acc1n0,staff,acc090" & _
               " where a1n01='" & .Fields("a1k01") & "' and a1n02='2' and st01(+)=a1n04" & _
               " and a0901(+)=st15 group by a0910 order by 3 desc,2,1"
         End If
         'end 2021/1/22
         intI = 1
         Set adoacc1n0 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strMemoFrom = strSerialNo
            '請款單點數
            'Modify By Sindy 2012/12/6
            'strPtTot = (Val("" & .Fields("a1k11")) - Val("" & .Fields("a1k09")) - Val("" & .Fields("a1k10")) * Val("" & .Fields("a1k06"))) / 1000
            strPtTot = (Val("" & .Fields("a1k11")) - Val("" & .Fields("a1k09")) - Val("" & .Fields("a1k06"))) / 1000
            '2012/12/6 End
            Do While Not adoacc1n0.EOF
               strAccNo = strR
               '只有一筆時不必分配
               If adoacc1n0.RecordCount = 1 Then
                  strAmt = strAmtTot
               '尚有未分配點數時
               Else
                  m_bolAlert = True '承辦點數有跨部門
                  If InStr(m_strAlertMsg, .Fields("a1k01")) = 0 Then m_strAlertMsg = m_strAlertMsg & vbCrLf & .Fields("a1k01")
                  
                  If adoacc1n0.AbsolutePosition > 1 Then
                     strSerialNo = Format(Val(strSerialNo) + 1, "000")
                  End If
                  
                  If strAmtRest > 0 Then
                     '若有出庭費固定扣第一筆(點數最多的)
                     If bolXFee = True And bolXFeeDone = False Then
                        strAmt = Round(strNetAmount * adoacc1n0.Fields("pts") / strPtTot, 2) - 10000
                        bolXFeeDone = True
                        
                     'Added by Morgan 2015/10/15
                     ElseIf adoacc1n0.AbsolutePosition = adoacc1n0.RecordCount Then
                        strAmt = strAmtRest
                     'end 2015/10/15
                     
                     'Modify by Morgan 2010/6/24 不足額也照比例分
                     'ElseIf 1000 * adoacc1n0.Fields("pts") > Val(strAmtRest) Then
                     '   strAmt = strAmtRest
                     Else
                        strAmt = Round(strNetAmount * adoacc1n0.Fields("pts") / strPtTot, 2)
                     End If
                  Else
                     strAmt = 0
                  End If
               End If
               
               strAmtRest = strAmtRest - strAmt
               
               '承辦人會計部門
               strDept = "" & adoacc1n0.Fields("a0910")
               '承辦人部門
               strEngDept = "" & adoacc1n0.Fields("st03")
               
               'FMP
               If strR = "411103" And Left(strSalesDept, 1) = "F" Then
                  If strDept = "FCP" Then
                     strAccNo = "417102"
                  Else
                     strAccNo = "411106"
                  End If
               'FMT
               ElseIf strR = "410103" And Left(strSalesDept, 1) = "F" Then
                  strAccNo = "410109"
            
               '417201 FCT收入,若為內商人員承辦時改科目為 417202 FCT爭議
               '417202 FCT爭議,若為國外部承辦時改科目為 417201 FCT收入
               'Modify by Morgan 2010/6/21 非 FCT,T 時要依跨部門規則
               'ElseIf strSystemType = "FCT" Then
               '   If strDept = "T" Then
               '      strAccNo = "417202"
               '   Else
               '      strAccNo = "417201"
               '   End If
               ElseIf strSystemType = "FCT" And strDept = "T" Then
                  strAccNo = "417202"
                  
               'cancel by sonia 2020/12/16 婧瑄說維持417202, T及FCT都可有此科目
               'ElseIf strSystemType = "FCT" And strDept = "FCT" Then
               '   strAccNo = "417201"
               'end 2020/12/16
               'end 2010/6/21
               
               'Add by Morgan 2010/10/8 CFT&FCT 或 CFL&FCL  不算跨部門
               ElseIf strSystemType = "CFT" And strDept = "FCT" Then
                  strDept = strSystemType
               ElseIf strSystemType = "CFL" And strDept = "FCL" Then
                  strDept = strSystemType
               'end 2010/10/8
               'add by sonia 2014/8/25 CFC&FCT 不算跨部門,但部門放CFT (X10305066,CFC-000795)
               ElseIf strSystemType = "CFC" And strDept = "FCT" Then
                  strDept = "CFT"
               'end 2014/8/25
               '2014/12/9 add by sonia S&FCT 不算跨部門,但部門依申請國家決定 (X10316550,S-003979)
               ElseIf strSystemType = "S" And strDept = "FCT" Then
                  If strNation <> "000" Then
                     strDept = "CFT"
                  Else
                     strDept = "FCT"
                  End If
               '2014/12/9 END
               
               'Add by Morgan 2010/10/27 CFP & P 不算跨部門
               ElseIf strSystemType = "CFP" And strDept = "P" Then
                  strDept = strSystemType
               'add by sonia 2019/5/3 FG&FCP 不算跨部門 M10801706
               ElseIf strSystemType = "FG" And strDept = "FCP" Then
               'end 2019/5/3
               '跨部門分點數
               ElseIf strSystemType <> strDept Then
                  Select Case strDept
                     Case "P"
                        strAccNo = "411101"
                        strShareP = Val(strShareP) + Val(strAmt)
                     Case "T"
                        strAccNo = "410101"
                        strShareT = Val(strShareT) + Val(strAmt)
                     Case "L"
                        strAccNo = "414101"
                        strShareL = Val(strShareL) + Val(strAmt)
                     Case "FCP"
                        'modify by sonia 2016/8/3
                        'strAccNo = "417101"
                        Select Case strSystemType
                           Case "FCL", "CFL", "LIN"
                              'modify by sonia 2022/2/17 由417103改為417109(M11100664同時改)
                              strAccNo = "417109"
                           Case Else
                              strAccNo = "417101"
                        End Select
                        'end 2016/8/3
                        strShareFCP = Val(strShareFCP) + Val(strAmt)
                     Case "FCT"
                        'modify by sonia 2016/8/3
                        'strAccNo = "417201"
                        Select Case strSystemType
                           Case "FCL", "CFL", "LIN"
                              'modify by sonia 2022/2/17 由417203改為417202(M11100664)
                              strAccNo = "417202"
                           Case Else
                              strAccNo = "417201"
                        End Select
                        'end 2016/8/3
                        strShareFCT = Val(strShareFCT) + Val(strAmt)
                     Case "FCL"
                        strAccNo = "416101"
                        strShareFCL = Val(strShareFCL) + Val(strAmt)
                  End Select
               End If
               
               strA1p16s = ""
               strSerialNoFrom = strSerialNo
               strExc(0) = "select a1n04,sum(a1n05) pts from acc1n0" & _
                  " where a1n01='" & .Fields("a1k01") & "' and a1n02='1'" & _
                  " group by a1n04 order by 2,1"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  Do While Not RsTemp.EOF
                  
                     'Added by Morgan 2021/4/12 FCT B2類案源
                     'G2:FCT收入-爭議
                     If strLOS02 = "B2" And Left(strCaseNo, 3) = "FCT" Then
                        strAccNo = "417202"
                     End If
                     'end 2021/4/12
                           
                     strSalesMan = "" & RsTemp("a1n04")
                     '作帳智權人代碼
                     'modify by sonia 2021/3/12 加傳日期
                     strSalesMan = SalesNoToAccSales(strSalesMan, strAccNo, strCaseNo, Val(strCon1))
                     If strSalesMan = "" Then
                        strSalesMan = "M0100"
                        
                     End If
                     '只有一筆時不必分配
                     If RsTemp.RecordCount = 1 Then
                        strA1p08 = strAmt
                        'Added by Morgan 2016/2/17
                        'Modified by Morgan 2020/4/9 請款單日期>=智慧所更名日者改回依案件性質表設定之科目收入；
                        If DBDATE(.Fields("a1k02")) < 智慧所更名日 And Val(strCon1) > 1050000 And (Left(strAccNo, 4) = "4141" Or Left(strAccNo, 4) = "4161" Or Left(strAccNo, 4) = "4181") Then
                           InsertLawACC1P0 "1", "F", strSerialNo, strItemNo, strAccNo, IIf(strDept = "", MsgText(55), strDept), 0, Val(strA1p08), "", "", "", "", "", _
                           strA1p14 & IIf(strR = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, "0.00"), ""), _
                           strCustNo, strSalesMan, strCaseNo, strCon1, strCurrency, strExchange, "" & .Fields("a0x11").Value, Replace(stra1p22, "'", ""), "", "", "", "", Replace(stra1p27, "'", ""), strA1p30_new, "", strCP09_Min
                           
'                        'add by sonia 2016/8/31 2016/9/1起FCP收入再細分科目
'                        ElseIf Val(strCon1) > 1050900 And Left(strAccNo, 4) = "4171" And strAccNo <> "417102" And strAccNo <> "417103" Then
'                           InsertFCPACC1P0 "1", "F", strSerialNo, strItemNo, strAccNo, IIf(strDept = "", MsgText(55), strDept), 0, Val(strA1p08), "", "", "", "", "", _
'                           strA1p14 & IIf(strR = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, "0.00"), ""), _
'                           strCustNo, strSalesMan, strCaseNo, strCon1, strCurrency, strExchange, "" & .Fields("a0x11").Value, Replace(stra1p22, "'", ""), "", "", "", "", Replace(stra1p27, "'", ""), strA1p30_new, "", .Fields("a1k01"), strNation
'                        'end 2016/8/31
                        
                        Else
                        'end 2016/2/17
                           
                           'Modify By Sindy 2015/12/31 strA1p30 ==> strA1p30_new
                           'modify by sonia 2021/2/25 +a1p23存A1K01
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                        ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                                        ",a1p30,a1p23) values " & _
                                        "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & strA1p08 & _
                                        ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, "0.00"), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                                        ",'" & strA1p30_new & "','" & .Fields("a1k01").Value & "')"
                        End If 'Added by Morgan 2016/2/17
                        
                     Else
                        strA1p08 = Round(strAmt * RsTemp("pts") / strPtTot, 2)
                        '若分配智權人的作帳智權人代碼已有資料時累加
                        If InStr(strA1p16s, strSalesMan) > 0 Then
                           adoTaie.Execute "update acc1p0 set a1p08=a1p08+" & strA1p08 & " where a1p01='1' and a1p02='F' and a1p03>='" & strSerialNoFrom & "' and a1p04='" & strItemNo & "' and a1p16='" & strSalesMan & "' and rownum<2"
                        Else
                           If RsTemp.AbsolutePosition > 1 Then
                              strSerialNo = Format(Val(strSerialNo) + 1, "000")
                           End If
                           
                           'Added by Morgan 2016/2/17
                           'Modified by Morgan 2020/4/9 請款單日期>=智慧所更名日者改回依案件性質表設定之科目收入；
                           If DBDATE(.Fields("a1k02")) < 智慧所更名日 And Val(strCon1) > 1050000 And (Left(strAccNo, 4) = "4141" Or Left(strAccNo, 4) = "4161" Or Left(strAccNo, 4) = "4181") Then
                              InsertLawACC1P0 "1", "F", strSerialNo, strItemNo, strAccNo, IIf(strDept = "", MsgText(55), strDept), 0, Val(strA1p08), "", "", "", "", "", _
                              strA1p14 & IIf(strR = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, "0.00"), ""), _
                              strCustNo, strSalesMan, strCaseNo, strCon1, strCurrency, strExchange, "" & .Fields("a0x11").Value, Replace(stra1p22, "'", ""), "", "", "", "", Replace(stra1p27, "'", ""), strA1p30_new, "", strCP09_Min
                              
'                           'add by sonia 2016/8/31 2016/9/1起FCP收入再細分科目
'                           ElseIf Val(strCon1) > 1050900 And Left(strAccNo, 4) = "4171" And strAccNo <> "417102" And strAccNo <> "417103" Then
'                              InsertFCPACC1P0 "1", "F", strSerialNo, strItemNo, strAccNo, IIf(strDept = "", MsgText(55), strDept), 0, Val(strA1p08), "", "", "", "", "", _
'                              strA1p14 & IIf(strR = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, "0.00"), ""), _
'                              strCustNo, strSalesMan, strCaseNo, strCon1, strCurrency, strExchange, "" & .Fields("a0x11").Value, Replace(stra1p22, "'", ""), "", "", "", "", Replace(stra1p27, "'", ""), strA1p30_new, "", .Fields("a1k01"), strNation
'                           'end 2016/8/31
                        
                           Else
                           'end 2016/2/17
                        
                              'Modify By Sindy 2015/12/31 strA1p30 ==> strA1p30_new
                              'modify by sonia 2021/2/25 +a1p23存A1K01
                              adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                              ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                                              ",a1p30,a1p23) values " & _
                                              "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & strA1p08 & _
                                              ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, "0.00"), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                                              ",'" & strA1p30_new & "','" & .Fields("a1k01").Value & "')"
                                              
                           End If 'Added by Morgan 2016/2/17
                           strA1p16s = strA1p16s & "," & strSalesMan
                        End If
                     End If
                     strAmt = strAmt - strA1p08
                     RsTemp.MoveNext
                  Loop
                  '差額放在最後一筆
                  If Val(strAmt) > 0 Then
                     adoTaie.Execute "update acc1p0 set a1p08=a1p08+" & strAmt & " where a1p01='1' and a1p02='F' and a1p03='" & strSerialNo & "' and a1p04='" & strItemNo & "'"
                  End If
                  
                  '若項次有增加表示智權人點數有做分配
                  If strSerialNo <> strSerialNoFrom Then
                     m_bolAlert = True
                     If InStr(m_strAlertMsg, .Fields("a1k01")) = 0 Then m_strAlertMsg = m_strAlertMsg & vbCrLf & .Fields("a1k01")
                  End If
               End If
               adoacc1n0.MoveNext
            Loop
            'add by sonia 2016/9/19 FCP收入再細分科目
            'modify by sonia 2023/1/18 +.Fields("a1k35").Value
            If Val(strCon1) > 1050900 And Left(strAccNo, 4) = "4171" And strAccNo <> "417102" And strAccNo <> "417103" Then
               UpdateFCPACC1P0 "1", "F", strSerialNo, strItemNo, strAccNo, .Fields("a1k01"), strNation, "" & .Fields("a1k35")
            End If
            'end 2016/9/19
            'Add by Morgan 2010/6/11
            '點數分配摘要
            If Val(strShareP) > 0 Then
               strSharePointMemo = strSharePointMemo & " P" & Round(100 * Val(strShareP) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareT) > 0 Then
               strSharePointMemo = strSharePointMemo & " T" & Round(100 * Val(strShareT) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareL) > 0 Then
               strSharePointMemo = strSharePointMemo & " L" & Round(100 * Val(strShareL) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareFCP) > 0 Then
               strSharePointMemo = strSharePointMemo & " FCP" & Round(100 * Val(strShareFCP) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareFCT) > 0 Then
               strSharePointMemo = strSharePointMemo & " FCT" & Round(100 * Val(strShareFCT) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareFCL) > 0 Then
               strSharePointMemo = strSharePointMemo & " FCL" & Round(100 * Val(strShareFCL) / Val(strAmtTot), 0) & "%"
            End If
            If strSharePointMemo <> "" Then
               adoTaie.Execute "update acc1p0 set a1p14 = a1p14 ||'/'||'" & Trim(strSharePointMemo) & "' where a1p01='1' and a1p02='F' and a1p03>='" & strMemoFrom & "' and a1p04='" & strItemNo & "' and substr(a1p05,1,1)='4' and a1p17='" & strCaseNo & "'"
            End If
            'end 2010/6/11
            
         'Added by Morgan 2016/4/8
         Else
         Debug.Print "test"
         'end 2016/4/8
         End If
         
      End If 'Added by Morgan 2016/1/7
         
         If Option1.Value Then
            strCustNo = Text4
         ElseIf Option2.Value Then
            strCustNo = Text6
         Else
            strCustNo = Text8
         End If
         
         'Modify By Sindy 2015/11/5 取消特殊客戶
'         '更新特定客戶之摘要
'         If PUB_CHKCUST(strCustNo) = True Then
'            adoTaie.Execute "UPDATE acc1p0 SET A1P14='" & Mid(m_NAME, 1, 4) & "/'||A1P14 WHERE A1P01='1' AND A1P02='F' AND A1P03='" & strSerialNo & "' AND A1P04='" & strItemNo & "'"
'         End If
      End If
      
      '規費
      strAccNo = strF
      
      Select Case strSystemType
         Case "S"
            If strNation = "000" Then
               strDept = "FCT"
            Else
               strDept = "CFT"
            End If
         Case "T", "TF"
            strDept = "T"
         Case "P", "PS"
            strDept = "P"
         Case "FCT"
            strDept = "FCT"
         Case "FCP", "FG"
            strDept = "FCP"
         Case "CFT", "CFC"
            strDept = "CFT"
         Case "CFP", "CPS"
            strDept = "CFP"
         Case "L"
            strDept = "L"
         'modify by sonia 2016/8/3 +LIN
         Case "FCL", "CFL", "LIN"
            strDept = "FCL"
         Case Else
            strDept = "T"
      End Select
      
      '台灣案專利商標出庭費控管
      If bolXFee = True Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         'modify by sonia 2021/2/25 +a1p23存A1K01
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                         ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                         "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, 10000" & _
                         ", '" & strA1p14 & "/出庭費', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(.Fields("a0x11").Value) & ", '" & .Fields("a1k01").Value & "')"
      End If
      
      'Added by Morgan 2021/4/12 FCT B2類案源
      'Modified by Morgan 2021/4/29 行政訴訟上訴及行政訴訟上訴答辯不必出庭
      If strLOS02 = "B2" And Left(strCaseNo, 3) = "FCT" Then
         If bolB2NeeCourt = True Then
            'B:應付規費－FCT 220103 = 商標處出庭費5000
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
            'Modified by Morgan 2024/11/13 改科目 220103 -> 220113 應付規費－律師庭費
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                      ", a1p14, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                      ",a1p23) values " & _
                      "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '220113', 0, 5000" & _
                      ", '" & strA1p14 & "(商標處出庭費)" & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", 'TOT', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                      ",'" & .Fields("a1k01").Value & "')"
                      
            'D:瑞興銀行乙存(智慧所) 110602 = A(收款金額) - B(商標處出庭費5000) - C(收款金額百位(含)以下)
            strD = Trunc(.Fields("a0x09").Value, -3) - 5000
         Else
            'D:瑞興銀行乙存(智慧所) 110602 = A(收款金額) - C(收款金額百位(含)以下)
            strD = Trunc(.Fields("a0x09").Value, -3)
         End If
         
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                   ", a1p14, a1p18, a1p19, a1p20, a1p06, a1p22, a1p27, a1p21" & _
                   ",a1p23) values " & _
                   "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '110602', 0, " & strD & _
                   ", '法律所/" & strA1p14 & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", 'TOT', " & stra1p22 & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                   ",'" & .Fields("a1k01").Value & "')"
                   
         
         'G1:應收帳款 1133 = G2 - C = G(法律所)
         strG1 = Trunc(Val(strG2), -3)
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                      ", a1p14, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                      ", a1p23) values " & _
                      "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '1133', " & strG1 & ", 0" & _
                      ", '法律所/" & strA1p14 & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", 'TOT', 'X82357000', " & stra1p22 & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                      ",'" & .Fields("a1k01").Value & "')"
         
         'D1:瑞興銀行乙存(法律所) 110502 = D(智慧所)
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                   ", a1p14, a1p18, a1p19, a1p20, a1p06,  a1p22, a1p27, a1p21" & _
                   ",a1p23) values " & _
                   "('L', 'F', '" & strSerialNo & "', '" & strItemNo & "', '110502', " & strD & ", 0" & _
                   ", '法律所/" & strA1p14 & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", 'TOT', " & strA1P22_L & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                   ",'" & .Fields("a1k01").Value & "')"
                   
         'G:代收款項-訴訟 240701 = G1(智慧所)
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                   ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                   ",a1p23) values " & _
                   "('L', 'F', '" & strSerialNo & "', '" & strItemNo & "', '240701', 0, " & strG1 & _
                   ", '" & strA1p14 & "(智慧所收款)/" & strLCaseNo & "', 'L0100', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", 'TOT', 'X82357000', " & strA1P22_L & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                   ",'" & .Fields("a1k01").Value & "')"
                   
         '客戶對沖放客戶編號
         'E:規費
         If Val(strE) > 0 Then
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
            'Modified by Morgan 2021/10/25 法律所的規費, 統一用"2403 代收代付"--婉莘
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                      ", a1p14, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                      ",a1p23) values " & _
                      "('L', 'F', '" & strSerialNo & "', '" & strItemNo & "', '2403', 0, " & strE & _
                      ", '" & strA1p14 & "(智慧所收款)/" & strLCaseNo & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", 'TOT', '" & strCustNo_L & "', " & strA1P22_L & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                      ",'" & .Fields("a1k01").Value & "')"
         End If
   
         If bolB2NeeCourt = True Then
            'F:應付規費－律師庭費 220113 = 5000
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
            'Modified by Morgan 2022/12/13 出庭費改抓設定，但暫不考慮多律師出庭狀況--婉莘
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                      ", a1p14, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                      ",a1p23) values " & _
                      "('L', 'F', '" & strSerialNo & "', '" & strItemNo & "', '220113', 0, " & lngCurtFee & _
                      ", '" & strA1p14 & "/" & strLCaseNo & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", 'TOT', '" & strCustNo_L & "', " & strA1P22_L & ", " & stra1p27 & ", " & .Fields("a0x11").Value & _
                      ",'" & .Fields("a1k01").Value & "')"
         End If
         
      'end 2021/4/12
      '規費>0
      ElseIf .Fields("a0x03").Value <> 0 Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         
         strA1P16 = ""
         'Add by Morgan 2011/10/7
         '承辦人為桂律師時規費科目項次(原為220113)改為414101,智權人員M0100
         If bolLawyerGuei Then
            strA1P16 = "M0100"
            strAccNo = "414101"
         End If
                           
         If Val(.Fields("a0x09").Value) >= Val(.Fields("a0x03").Value) Then
            If strAccNo = "220105" Or strAccNo = "220106" Then
                'CFT, CFP摘要帶的金額為總額(規費+服務費)
               'Modify By Sindy 2015/12/31 strA1p14 ==> strCaseNo & "/" & strCaseProperty
               'Modified by Lydia 2016/07/14 小數捨去Trunc
               'modify by sonia 2021/2/25 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                               ",a1p30,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(.Fields("a0x03").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "/" & IIf(strDept = "CFT" Or strDept = "CFP", Val(.Fields("a0x09").Value), Val(.Fields("a0x03").Value)) & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(.Fields("a0x11").Value) & _
                               ",'" & strA1P30 & "','" & .Fields("a1k01").Value & "')"
            ElseIf strAccNo = "220111" Or strAccNo = "220112" Then
               '大陸專利商標摘要帶的金額為總額(規費+服務費)
               'Modify By Sindy 2015/12/31 strA1p14 ==> strCaseNo & "/" & strCaseProperty
               'Modified by Lydia 2016/07/14 小數捨去Trunc
               'modify by sonia 2021/2/25 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                               ",a1p30,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(.Fields("a0x03").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "/" & Val(.Fields("a0x09").Value) & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(.Fields("a0x11").Value) & _
                               ",'" & strA1P30 & "','" & .Fields("a1k01").Value & "')"
            'Added by Lydia 2016/07/14 規費不要有小數
            ElseIf Left(strAccNo, 4) = "2201" Then
               'modify by sonia 2021/2/25 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                               ",a1p30,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(.Fields("a0x03").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(.Fields("a0x11").Value) & _
                               ",'" & strA1P30 & "','" & .Fields("a1k01").Value & "')"
            'end 2016/07/14
            Else
               'Modify By Sindy 2015/12/31 strA1p14 ==> strCaseNo & "/" & strCaseProperty
               'modify by sonia 2021/2/25 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                               ",a1p30,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(.Fields("a0x03").Value) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(.Fields("a0x11").Value) & _
                               ",'" & strA1P30 & "','" & .Fields("a1k01").Value & "')"
            End If
         Else
            If strAccNo = "220105" Or strAccNo = "220106" Or strAccNo = "220111" Or strAccNo = "220112" Then
               'Modify By Sindy 2015/12/31 strA1p14 ==> strCaseNo & "/" & strCaseProperty
               'Modified by Lydia 2016/07/14 小數捨去Trunc
               'modify by sonia 2021/2/25 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                               ",a1p30,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(.Fields("a0x09").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "/" & IIf(IsNull(.Fields("a0x09").Value), "", .Fields("a0x09").Value) & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(.Fields("a0x11").Value) & _
                               ",'" & strA1P30 & "','" & .Fields("a1k01").Value & "')"
            'Added by Lydia 2016/07/14 規費不要有小數
            ElseIf Left(strAccNo, 4) = "2201" Then
               'modify by sonia 2021/2/25 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                               ",a1p30,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(.Fields("a0x09").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(.Fields("a0x11").Value) & _
                               ",'" & strA1P30 & "','" & .Fields("a1k01").Value & "')"
            'end 2016/07/14
            Else
               'Modify By Sindy 2015/12/31 strA1p14 ==> strCaseNo & "/" & strCaseProperty
               'modify by sonia 2021/2/25 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                               ",a1p30,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(.Fields("a0x09").Value) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(.Fields("a0x11").Value) & _
                               ",'" & strA1P30 & "','" & .Fields("a1k01").Value & "')"
            End If
         End If
      End If
      
      'add by sonia 2024/11/15 外專OA委外翻譯之帳單費用，改用6130科目
      adocaseprogress.CursorLocation = adUseClient
      adocaseprogress.Open "select a1p06,a1p07,a1p19,a1p20,a1p21,a1p30 from acc1w0,caseprogress,acc1p0 where a1w01='" & .Fields("a1k01").Value & "' and substr(a1w02,1,1)='B' and a1w02=cp09(+) " & _
                           "and cp01 in ('P','FCP') and cp10='927' and substr(cp14,1,1)='F' and substr(cp43,1,1)='C' and cp61||a1w02=a1p23 and a1p07>0", adoTaie, adOpenStatic, adLockReadOnly
      If adocaseprogress.RecordCount <> 0 Then
         If .Fields("a0x03").Value <> 0 Then
            If .Fields("a0x03").Value > Val("" & adocaseprogress.Fields("a1p07")) Then
               adoTaie.Execute "update acc1p0 set a1p08=a1p08-" & Val("" & adocaseprogress.Fields("a1p07")) & " WHERE A1P01='1' AND A1P02='F' AND A1P03='" & strSerialNo & "' AND A1P04='" & strItemNo & "'"
            Else
               adoTaie.Execute "delete acc1p0 WHERE A1P01='1' AND A1P02='F' AND A1P03='" & strSerialNo & "' AND A1P04='" & strItemNo & "'"
            End If
         End If
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                         ", a1p14, a1p17, a1p18, a1p19, a1p20, a1p06, a1p22, a1p27, a1p21" & _
                         ",a1p30,a1p23) values " & _
                         "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '6130', 0, " & Val("" & adocaseprogress.Fields("a1p07")) & _
                         ", '" & strCaseNo & "/OA委外翻譯', '" & strCaseNo & "', " & Val(strCon1) & ", '" & adocaseprogress.Fields("a1p19") & "', " & adocaseprogress.Fields("a1p20") & ", '" & adocaseprogress.Fields("a1p06") & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adocaseprogress.Fields("a1p21").Value) & _
                         ",'" & adocaseprogress.Fields("a1p30") & "','" & .Fields("a1k01").Value & "')"
      End If
      adocaseprogress.Close
      'end 2024/11/15
   
   End With
   End If
   Set adoAcc0x0_1 = Nothing
   Set adoacc1n0 = Nothing
   Set adoCP = Nothing
End Sub

'*************************************************
'  儲存資料表(國外收款資料(交易檔))
'
'*************************************************
Private Sub Acc0z0Save()
Dim strCaseNo As String
Dim strCaseProperty As String
Dim strCustomer As String
Dim strSalesMan As String
Dim strCurrency As String
Dim strExchange As String
Dim strSerialNo As String
Dim strSystemType As String
Dim strAccNo As String
Dim intArgument As Integer
Dim douFCTamount As Double
Dim douFCTAamount As Double
Dim douAmount As Double
Dim strDept As String
Dim stra1p22 As String
Dim stra1p27 As String
Dim strCompany As String
Dim strMan As String
Dim strCustNo As String
Dim strProperty As String
Dim strR As String
Dim strF As String
Dim StrStaff As String
Dim strNation As String
Dim W_strSerialNo As String  '2005/8/12 ADD BY SONIA
'Add by Morgan 2007/9/19
Dim bolXFee As Boolean '服務費是否含出庭費
Dim bolXFeeDone As Boolean '出庭費是否已扣除
Dim strCP09 As String '收文號
Dim strA1p14 As String '摘要
Dim dblTotAmt As Double 'Add By Sindy 2018/11/2 台幣已收款金額
Dim dblTotA0Z12 As Double 'Add By Sindy 2018/11/5 扣繳金額

On Error GoTo Checking
   m_bolAlert = False
   m_strAlertMsg = ""
   W_strSerialNo = ""
   '紀錄傳票號,設定是否更新
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      stra1p22 = "'" & adoaccsum.Fields("a1p22").Value & "'"
      stra1p27 = "'" & "Y" & "'"
      adoTaie.Execute "update acc1p0 set a1p27 = 'Y' where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "'"
   Else
      stra1p22 = "null"
      stra1p27 = "null"
   End If
   adoaccsum.Close
   
   'Added by Morgan 2021/4/12
   '紀錄L公司傳票號
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select a1p22 from acc1p0 where a1p01 = 'L' and a1p02 = 'F' and a1p04 = '" & Text3 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      strA1P22_L = "'" & adoaccsum.Fields("a1p22").Value & "'"
      adoTaie.Execute "update acc1p0 set a1p27 = 'Y' where a1p01 = 'L' and a1p02 = 'F' and a1p04 = '" & Text3 & "'"
   Else
      strA1P22_L = "null"
   End If
   adoaccsum.Close
   'end 2021/4/12
   
   '刪除收款資料,貸方&預付稅捐分錄
   adoTaie.Execute "delete from acc0z0 where a0z01 = '" & Text3 & "'"
   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "' and a1p08 <> 0"
   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "' and a1p05 = '1203'"
   'Modify By Sindy 2015/11/5 Mark
'   'Add By Sindy 2015/4/29
'   adoTaie.Execute "delete from acc1v0 where a1v02 in(select a0x02 from acc0x0 where a0x01 = '" & Text3 & "' and a0x15 = '" & strUserNum & "')"
'   '2015/4/29 END

   'Added by Morgan 2021/4/12 刪除案源收款系統自動產生的分錄
   '智慧所借方分錄1133應收帳款
   'Modified by Morgan 2021/6/22 +a1p14 判斷
   'Modified by Morgan 2024/8/23 +可一律排除應收帳款(1133)，因摘要內容不確定(Ex:M11303689)，且秀玲說我們收款借方不會用此科目
   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "' and a1p05 = '1133' and instr(a1p14||' ','法律所/')=1 and a1p07<>0"
   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "' and a1p05 = '1133' and a1p07<>0"
   'end 2024/8/23
   'L公司分錄
   adoTaie.Execute "delete from acc1p0 where a1p01 = 'L' and a1p02 = 'F' and a1p04 = '" & Text3 & "'"
   'end 2021/4/12
   
   '讀取請款&收款暫存資料
   adoacc0x0.CursorLocation = adUseClient
   'Modified by Morgan 2021/4/21 為避免收款修改產生的分錄順序會跟原來不同，改固定依請款單號排序
   'adoacc0x0.Open "select * from acc1k0, acc0x0 where a1k01 = a0x02 and a0x15 = '" & strUserNum & "' order by a0x14 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0x0.Open "select * from acc1k0, acc0x0 where a1k01 = a0x02 and a0x15 = '" & strUserNum & "' order by a0x02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0x0.EOF = False
      SetLOSVar adoacc0x0("a1k01"), strLOS02, strLCaseNo, bolB2NeeCourt, lngCurtFee 'Added by Morgan 2021/4/12 案源變數設定

      'Add by Morgan 2010/5/21 若有分配資料則跑新程式
      strExc(0) = "select * from acc1n0 where a1n01='" & adoacc0x0("a1k01") & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Acc0z0SaveNew adoacc0x0("a0x01"), adoacc0x0("a0x02"), stra1p22, stra1p27, W_strSerialNo
         GoTo NextRec
      End If
      
      bolXFee = False: bolXFeeDone = False 'Add by Morgan 2007/9/19
      If IsNull(adoacc0x0.Fields("a1k13").Value) Then
         strCaseNo = ""
         strSystemType = ""
      Else
         strCaseNo = adoacc0x0.Fields("a1k13").Value & adoacc0x0.Fields("a1k14").Value & adoacc0x0.Fields("a1k15").Value & adoacc0x0.Fields("a1k16").Value
         strSystemType = adoacc0x0.Fields("a1k13").Value
      End If
      adoacc0z0.CursorLocation = adUseClient
      adoacc0z0.Open "select * from acc0z0 where a0z01 = '" & adoacc0x0.Fields("a0x01").Value & "' and a0z02 = '" & adoacc0x0.Fields("a0x02").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc0z0.RecordCount = 0 Then
         adoacc0z0.AddNew
      End If
      adoacc0z0.Fields("a0z01").Value = adoacc0x0.Fields("a0x01").Value
      adoacc0z0.Fields("a0z02").Value = adoacc0x0.Fields("a0x02").Value
      If IsNull(adoacc0x0.Fields("a0x08").Value) Then
         adoacc0z0.Fields("a0z03").Value = Null
      Else
         adoacc0z0.Fields("a0z03").Value = adoacc0x0.Fields("a0x08").Value
      End If
      If IsNull(adoacc0x0.Fields("a0x11").Value) Then
         adoacc0z0.Fields("a0z04").Value = 0
      Else
         adoacc0z0.Fields("a0z04").Value = adoacc0x0.Fields("a0x11").Value
      End If
      If IsNull(adoacc0x0.Fields("a0x12").Value) Then
         adoacc0z0.Fields("a0z12").Value = 0
      Else
         adoacc0z0.Fields("a0z12").Value = adoacc0x0.Fields("a0x12").Value
      End If
      If IsNull(adoacc0x0.Fields("a0x13").Value) Then
         adoacc0z0.Fields("a0z13").Value = Null
      Else
         adoacc0z0.Fields("a0z13").Value = adoacc0x0.Fields("a0x13").Value
      End If
      adoacc0z0.Fields("a0z06").Value = Val(strSrvDate(2))
      adoacc0z0.Fields("a0z07").Value = ServerTime
      adoacc0z0.Fields("a0z08").Value = strUserNum
      adoacc0z0.UpdateBatch
      adoacc0z0.Close
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
      adoquery.CursorLocation = adUseClient
      'Modify By Sindy 2018/11/5 + ,sum(nvl(a0z12, 0))
      adoquery.Open "select sum(nvl(a0z04, 0) * nvl(a0y04, 0)),sum(nvl(a0z12, 0)) from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & adoacc0x0.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      dblTotAmt = 0 'Add By Sindy 2018/11/2 台幣已收款金額
      dblTotA0Z12 = 0 'Add By Sindy 2018/11/5 扣繳金額
      If adoquery.RecordCount <> 0 Then
         dblTotAmt = Val(adoquery.Fields(0).Value) 'Add By Sindy 2018/11/2 台幣已收款金額
         dblTotA0Z12 = Val(adoquery.Fields(1).Value) 'Add By Sindy 2018/11/5 扣繳金額
         adoTaie.Execute "update acc1k0 set a1k30 = " & Val(adoquery.Fields(0).Value) & " where a1k01 = '" & adoacc0x0.Fields("a1k01").Value & "'"
      End If
      adoquery.Close
      adocaseprogress.CursorLocation = adUseClient
      'Modify By Sindy 2015/5/1 nvl(cpm03, nvl(cpm10, cpm13)) ==> DECODE(PA09,'000',CPM03,CPM04)
      adocaseprogress.Open "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(PA09,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, pa09 as nation from caseprogress, salesno, staff, casepropertyMap, patent, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' union " & _
                           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(TM10,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, tm10 as nation from caseprogress, salesno, staff, casepropertyMap, trademark, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' union " & _
                           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(LC15,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, lc15 as nation from caseprogress, salesno, staff, casepropertyMap, lawcase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' union " & _
                           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, nvl(cpm03, cpm04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, null as nation from caseprogress, salesno, staff, casepropertyMap, hirecase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' union " & _
                           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(SP09,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, sp09 as nation from caseprogress, salesno, staff, casepropertyMap, servicepractice, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
      If adocaseprogress.RecordCount <> 0 Then
         If IsNull(adocaseprogress.Fields("CustNo").Value) Then
            strCustNo = ""
         Else
            strCustNo = adocaseprogress.Fields("Custno").Value
         End If
         If IsNull(adocaseprogress.Fields("Company").Value) Then
            strCompany = ""
         Else
            strCompany = adocaseprogress.Fields("Company").Value
         End If
         If IsNull(adocaseprogress.Fields("Property").Value) Then
            strCaseProperty = ""
         Else
            strCaseProperty = adocaseprogress.Fields("Property").Value
         End If
         If IsNull(adocaseprogress.Fields("cp13").Value) Then
            strSalesMan = ""
         Else
            strSalesMan = adocaseprogress.Fields("cp13").Value
         End If
         If IsNull(adocaseprogress.Fields("Man").Value) Then
            strMan = ""
         Else
            strMan = adocaseprogress.Fields("Man").Value
         End If
         If IsNull(adocaseprogress.Fields("st03").Value) Then
            strDept = ""
         Else
            strDept = adocaseprogress.Fields("st03").Value
         End If
         If IsNull(adocaseprogress.Fields("cp10").Value) Then
            strProperty = ""
         Else
            strProperty = adocaseprogress.Fields("cp10").Value
         End If
         If IsNull(adocaseprogress.Fields("cpm11").Value) Then
            strR = ""
         Else
            strR = adocaseprogress.Fields("cpm11").Value
         End If
         If IsNull(adocaseprogress.Fields("cpm12").Value) Then
            strF = ""
         Else
            strF = adocaseprogress.Fields("cpm12").Value
         End If
         If IsNull(adocaseprogress.Fields("cp14").Value) Then
            StrStaff = ""
         Else
            StrStaff = adocaseprogress.Fields("cp14").Value
         End If
         If IsNull(adocaseprogress.Fields("nation").Value) Then
            strNation = ""
         Else
            strNation = adocaseprogress.Fields("nation").Value
         End If
         strCP09 = "" & adocaseprogress.Fields("cp09")
      Else
         strCaseProperty = ""
         strSalesMan = ""
         strDept = ""
         strCompany = ""
         strMan = ""
         strProperty = ""
         strR = ""
         strF = ""
         StrStaff = ""
         strNation = ""
         strCP09 = ""
      End If
      adocaseprogress.Close
      If strSystemType = "" Then
         If IsNull(adoacc0x0.Fields("a1k13").Value) = False Then
            strSystemType = adoacc0x0.Fields("a1k13").Value
         End If
      End If
      
'Remove by Morgan 2010/4/22 因目前1j09 沒有設定且相關計算也有問題故取消
'      '設定FCT案收入
'      If strSystemType = "FCT" Then
'         adoacc1k0.CursorLocation = adUseClient
'         adoacc1k0.Open "select sum(a1l05) from acc1l0, acc1j0 where a1l03 = a1j01 and a1l04 = a1j02 and a1l01 = '" & adoacc0x0.Fields("a0x02").Value & "' and a1j09 = '417201'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoacc1k0.RecordCount <> 0 Then
'            If IsNull(adoacc1k0.Fields(0).Value) Then
'               douFCTamount = 0
'            Else
'               douFCTamount = adoacc1k0.Fields(0).Value
'            End If
'         Else
'             douFCTamount = 0
'         End If
'         adoacc1k0.Close
'         adoacc1k0.CursorLocation = adUseClient
'         adoacc1k0.Open "select sum(a1l05) from acc1l0, acc1j0 where a1l03 = a1j01 and a1l04 = a1j02 and a1l01 = '" & adoacc0x0.Fields("a0x02").Value & "' and a1j09 = '417202'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoacc1k0.RecordCount <> 0 Then
'            If IsNull(adoacc1k0.Fields(0).Value) Then
'               douFCTAamount = 0
'            Else
'               douFCTAamount = adoacc1k0.Fields(0).Value
'            End If
'         Else
'             douFCTAamount = 0
'         End If
'         adoacc1k0.Close
'      End If
'end 2010/4/22

      '設定匯率幣別
      adoacc1k0.CursorLocation = adUseClient
      adoacc1k0.Open "select a1k18, a1k10 from acc1k0 where a1k01 = '" & adoacc0x0.Fields("a0x02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc1k0.RecordCount <> 0 Then
         If IsNull(adoacc1k0.Fields(0).Value) Then
            strCurrency = ""
         Else
            strCurrency = adoacc1k0.Fields(0).Value
         End If
         If IsNull(adoacc1k0.Fields(1).Value) Then
            strExchange = ""
         Else
            strExchange = adoacc1k0.Fields(1).Value
         End If
      Else
         strCurrency = ""
         strExchange = ""
      End If
      adoacc1k0.Close
      strAccNo = ""
      
      strA1p14 = strCaseNo & "/" & strCaseProperty 'Add By Sindy 2015/11/5 摘要
      
      '有扣繳時產生預付稅捐分錄
      If adoacc0x0.Fields("a0x12").Value <> 0 Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
         '                "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '1203', " & Val(adoacc0x0.Fields("a0x12").Value) & ", 0, '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format((Val(adoacc0x0.Fields("a0x12").Value) / Val(strExchange)), FAmount)) & ")"
         '2005/5/3 MODIFY BY SONIA
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
         '                "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '1203', " & Val(adoacc0x0.Fields("a0x12").Value) & ", 0, '" & strCaseNo & "/" & strCaseProperty & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format((Val(adoacc0x0.Fields("a0x12").Value) / Val(strExchange)), FAmount)) & ")"
         '2005/10/25 MODIFY BY SONIA
         'If PUB_CHKCUST(Text4) = True Then
         If Option1.Value Then
            strCustNo = Text4
         Else
            If Option2.Value Then
               strCustNo = Text6
            Else
               strCustNo = Text8
            End If
         End If
         
         'Modify By Sindy 2015/11/5 取消特殊客戶
'         If PUB_CHKCUST(strCustNo) = True Then
'         '2005/10/25 END
            strA1p14 = Mid(adoacc0x0.Fields("a1k35").Value, 1, 6) & "/" & strA1p14 'Add By Sindy 2015/11/5
            'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
            '                          StrToStr(m_YEAR & Mid(m_NAME, 1, 4), 5) ==> Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4)
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                            ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                            ", A1P30) values " & _
                            "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '1203', " & Val(adoacc0x0.Fields("a0x12").Value) & ", 0" & _
                            ", '" & strA1p14 & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format((Val(adoacc0x0.Fields("a0x12").Value) / Val(strExchange)), FAmount)) & _
                            ",'" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & "')"
'         Else
''            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
''                            "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '1203', " & Val(adoacc0x0.Fields("a0x12").Value) & ", 0, '" & strCaseNo & "/" & strCaseProperty & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format((Val(adoacc0x0.Fields("a0x12").Value) / Val(strExchange)), FAmount)) & ")"
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21, A1P30) values " & _
'                            "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '1203', " & Val(adoacc0x0.Fields("a0x12").Value) & ", 0, '" & strA1p14 & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format((Val(adoacc0x0.Fields("a0x12").Value) / Val(strExchange)), FAmount)),'')"
'         End If
'         '2005/5/3 END
         
'         'Add By Sindy 2015/4/28 有輸入扣繳金額時,同時寫入acc1v0
'         'Modify By Sindy 2015/11/5 檢查資料是否已存在
'         strExc(0) = "select a1v02 from acc1v0 where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 0 Then
'            adoTaie.Execute "insert into ACC1V0 (a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v18,a1v12,a1v13)" & _
'                            " values('" & strCP09 & "','" & adoacc0x0.Fields("a0x02").Value & "'," & IIf(Text13 <> "", Text13, "GetA0k11('" & strCP09 & "')") & _
'                            "," & Val(adoacc0x0.Fields("a0x12").Value) & ",'" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'," & Val(adoacc0x0.Fields("a0x12").Value) & ",0" & _
'                            "," & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & ",'1','" & strCaseProperty & "','" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "')"
'         Else
'            'Modify By Sindy 2018/11/2 多次收款:A1V04=(分次收款總額-規費)/10
'            '   Val(adoacc0x0.Fields("a0x12").Value) => Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0)
'            adoTaie.Execute "update ACC1V0" & _
'                            " set a1v01='" & strCP09 & "'" & _
'                            ",a1v03=" & IIf(Text13 <> "", Text13, "GetA0k11('" & strCP09 & "')") & _
'                            ",a1v04=" & Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
'                            ",a1v05='" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'" & _
'                            ",a1v06=" & Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
'                            ",a1v07=0" & _
'                            ",a1v09=" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & _
'                            ",a1v18='1'" & _
'                            ",a1v12='" & strCaseProperty & "'" & _
'                            ",a1v13='" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "'" & _
'                            " where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
'         End If
'         '2015/4/28 END
'      Else
'         'Add By Sindy 2015/11/3 有輸入請款單抬頭時,無輸入扣繳金額也要寫入acc1v0
'         If Trim(adoacc0x0.Fields("a1k35").Value) <> "" Then
'            'Modify By Sindy 2015/11/5 檢查資料是否已存在
'            strExc(0) = "select a1v02 from acc1v0 where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 0 Then
'               'modify by sonia 2017/7/7 收款有扣繳金額a1v18才存'1'
''               adoTaie.Execute "insert into ACC1V0 (a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v18,a1v12,a1v13)" & _
''                               " values('" & strCP09 & "','" & adoacc0x0.Fields("a0x02").Value & "'," & IIf(Text13 <> "", Text13, "GetA0k11('" & strCP09 & "')") & _
''                               "," & Round((Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & ",'" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "',0," & Round((Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
''                               "," & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & ",'1','" & strCaseProperty & "','" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "')"
'               adoTaie.Execute "insert into ACC1V0 (a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v12,a1v13)" & _
'                               " values('" & strCP09 & "','" & adoacc0x0.Fields("a0x02").Value & "'," & IIf(Text13 <> "", Text13, "GetA0k11('" & strCP09 & "')") & _
'                               "," & Round((Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & ",'" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "',0," & Round((Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
'                               "," & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & ",'" & strCaseProperty & "','" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "')"
'            Else
'               'modify by sonia 2017/7/7 收款有扣繳金額a1v18才存'1'
''               adoTaie.Execute "update ACC1V0" & _
''                               " set a1v01='" & strCP09 & "'" & _
''                               ",a1v03=" & IIf(Text13 <> "", Text13, "GetA0k11('" & strCP09 & "')") & _
''                               ",a1v04=" & Round((Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
''                               ",a1v05='" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'" & _
''                               ",a1v06=0" & _
''                               ",a1v07=" & Round((Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
''                               ",a1v09=" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & _
''                               ",a1v18='1'" & _
''                               ",a1v12='" & strCaseProperty & "'" & _
''                               ",a1v13='" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "'" & _
''                               " where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
'               'Modify By Sindy 2018/11/2 多次收款:A1V04=(分次收款總額-規費)/10
'               '   Round((Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) => Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0)
'               adoTaie.Execute "update ACC1V0" & _
'                               " set a1v01='" & strCP09 & "'" & _
'                               ",a1v03=" & IIf(Text13 <> "", Text13, "GetA0k11('" & strCP09 & "')") & _
'                               ",a1v04=" & Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
'                               ",a1v05='" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'" & _
'                               ",a1v06=0" & _
'                               ",a1v07=" & Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
'                               ",a1v09=" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & _
'                               ",a1v12='" & strCaseProperty & "'" & _
'                               ",a1v13='" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "'" & _
'                               ",a1v18=null where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
'            End If
'         End If
'         '2015/11/3 END
      End If
      'Modify By Sindy 2018/11/5
      '有扣繳時產生預付稅捐分錄
      'Add By Sindy 2015/11/3 有輸入請款單抬頭時,無輸入扣繳金額也要寫入acc1v0
      If adoacc0x0.Fields("a0x12").Value <> 0 Or dblTotA0Z12 > 0 Or Trim(adoacc0x0.Fields("a1k35").Value) <> "" Then
         'Add By Sindy 2015/4/28 有輸入扣繳金額時,同時寫入acc1v0
         'Modify By Sindy 2015/11/5 檢查資料是否已存在
         strExc(0) = "select a1v02 from acc1v0 where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            adoTaie.Execute "insert into ACC1V0 (a1v01,a1v02)" & _
                            " values('" & strCP09 & "','" & adoacc0x0.Fields("a0x02").Value & "')"
         End If
         'Modify By Sindy 2022/3/8
         '",a1v03=" & IIf(Text13 <> "", Text13, "GetA0k11('" & strCP09 & "')")
         '=> ",a1v03=" & m_A1K37
         adoTaie.Execute "update ACC1V0 set" & _
                         " a1v01='" & strCP09 & "'" & _
                         ",a1v03=" & m_A1K37 & _
                         ",a1v05='" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "Y", "N") & "'" & _
                         ",a1v09=" & Left(adoacc0y0.Fields("a0y02").Value, Len(adoacc0y0.Fields("a0y02").Value) - 4) & _
                         ",a1v12='" & strCaseProperty & "'" & _
                         ",a1v13='" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "'" & _
                         " where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
         'Modify By Sindy 2018/11/2 多次收款:A1V04=(分次收款總額-規費)/10
         '   Val(adoacc0x0.Fields("a0x12").Value) => Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0)
         '應該是A1V04 = (分次收款總額 - 規費) / 10
         'A1V06=分次收款A0Z12合計，
         '若A1V06>0則A1V07=0,A1V18='1'
         '  A1V06=0則A1V07=A1V04,A1V18=NULL
         'Modify By Sindy 2019/5/10 + Round(dblTotA0Z12, 0)
         ' a1v04=" & Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0)
         ',a1v06=" & Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0)
         If adoacc0x0.Fields("a0x12").Value <> 0 Or dblTotA0Z12 > 0 Then
            adoTaie.Execute "update ACC1V0 set" & _
                            " a1v04=" & Round(dblTotA0Z12, 0) & _
                            ",a1v06=" & Round(dblTotA0Z12, 0) & _
                            ",a1v07=0" & _
                            ",a1v18='1'" & _
                            " where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
         Else
            adoTaie.Execute "update ACC1V0 set" & _
                            " a1v04=" & Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
                            ",a1v06=0" & _
                            ",a1v07=" & Round((dblTotAmt - Val(adoacc0x0.Fields("a0x03").Value)) / 10, 0) & _
                            ",a1v18=null" & _
                            " where a1v02='" & adoacc0x0.Fields("a0x02").Value & "'"
         End If
      End If
      '2018/11/5 END
      
      '收入科目
      '2005/8/10 SONIA 原為先產生規費再產生收入, 改成先產生收入再產生規費
      If Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value) > 0 Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         '2005/8/12 ADD BY SONIA
         W_strSerialNo = strSerialNo
         '2005/8/12 END
         If strR <> "" Then
            strAccNo = strR
         Else
            Select Case strSystemType
               Case "S"
                  If strNation = "000" Then
                     strAccNo = "417201"
                  Else
                     strAccNo = "4121"
                  End If
               Case "P", "PS"
                  If strNation = "000" Then
                     strAccNo = "411101"
                  Else
                     strAccNo = "411103"
                  End If
               Case "T", "TF"
                  If strNation = "000" Then
                     strAccNo = "410101"
                  Else
                     strAccNo = "410103"
                  End If
               Case "TB"
                  strAccNo = "410105"
                  '2008/9/23 ADD BY SONIA
                  If strNation <> "000" Then
                     strAccNo = "410103"
                  End If
                  '2008/9/23 END
               Case "TC"
                  strAccNo = "415101"
                  '2008/9/23 ADD BY SONIA
                  If strNation <> "000" Then
                     strAccNo = "410103"
                  End If
                  '2008/9/23 END
               Case "TD"
                  strAccNo = "410108"
                  '2008/9/23 ADD BY SONIA
                  If strNation <> "000" Then
                     strAccNo = "410103"
                  End If
                  '2008/9/23 END
               Case "TM"
                  strAccNo = "410106"
                  '2008/9/23 ADD BY SONIA
                  If strNation <> "000" Then
                     strAccNo = "410103"
                  End If
                  '2008/9/23 END
               Case "CFT", "CFC"
                  strAccNo = "4121"
               Case "CFP", "CPS"
                  strAccNo = "4131"
               Case "L"
                  strAccNo = "4141"
               'modify by sonia 2016/8/3 +LIN
               Case "FCL", "LIN"
                  strAccNo = "416101"
               'add by sonia 2016/8/3
               Case "CFL"
                  strAccNo = "416102"
               'end 2016/8/3
               Case "FCP", "FG"
                  '2009/4/17 MODIFY BY SONIA
                  'strAccNo = "4171"
                  strAccNo = "417101"
               Case "FCT"
                  strAccNo = "4172"
               Case Else
                  strAccNo = "410101"
                  '2008/9/23 ADD BY SONIA
                  If strNation <> "000" Then
                     strAccNo = "410103"
                  End If
                  '2008/9/23 END
            End Select
         End If
        'Modify By Cheng 2004/05/11
        'SalesNoToAccSales執行一次就好
'         If SalesNoToAccSales(strSalesMan, strAccNo) <> "" Then
'        strSalesMan = SalesNoToAccSales(strSalesMan, strAccNo)
         'modify by sonia 2021/3/12 加傳日期
         strSalesMan = SalesNoToAccSales(strSalesMan, strAccNo, strCaseNo, Val(strCon1))
         If strSalesMan <> "" Then
'            strSalesMan = SalesNoToAccSales(strSalesMan, strAccNo)
         '92.10.6 ADD BY SONIA
         Else
            strSalesMan = "M0100"
         '92.10.6 END
         End If
         
         Select Case strSystemType
            Case "S"
               '92.4.18 MODIFY BY SONIA
               'strDept = "S"
               If strNation = "000" Then
                  strDept = "FCT"
               Else
                  strDept = "CFT"
                  strAccNo = "4121"
               End If
               '92.4.18 END
            Case "P", "PS"
               strDept = "P"
            Case "T", "TF"
               strDept = "T"
            Case "CFT", "CFC"
               strDept = "CFT"
            Case "CFP", "CPS"
               strDept = "CFP"
            Case "L"
               strDept = "L"
            'modify by sonia 2016/8/3 +LIN
            Case "FCL", "CFL", "LIN"
               strDept = "FCL"
            Case "FCP", "FG"
               strDept = "FCP"
            Case "FCT"
               strDept = "FCT"
               If adoaccsum.State = adStateOpen Then
                  adoaccsum.Close
               End If
               adoaccsum.CursorLocation = adUseClient
               adoaccsum.Open "select st03 from staff where st01 = '" & StrStaff & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoaccsum.RecordCount <> 0 Then
                  If IsNull(adoaccsum.Fields("st03").Value) = False Then
                     If Mid(adoaccsum.Fields("st03").Value, 1, 2) = "P2" Then
                        strDept = "T"
                     End If
                  End If
               End If
               adoaccsum.Close
            Case Else
               '92.10.9 MODIFY BY SONIA
               'strDept = "TOT"
               strDept = "T"
               '92.10.9 END
         End Select
         
         'Add by Morgan 2007/9/12 台灣案專利商標出庭費控管
         If strNation = "000" And Val(adoacc0x0("a1k02")) >= 960815 And Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value) >= 10000 Then
            '專利
            If (strSystemType = "P" Or strSystemType = "FCP") Then
               If InStr("211,212", strProperty) > 0 Then
                  bolXFee = True
               ElseIf InStr("503,507,506", strProperty) > 0 Then
                  strExc(0) = "select * from caseprogress a where cp09='" & strCP09 & "'" & _
                     " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('211','212'))" & _
                     " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('211','212') and b.cp16>0)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolXFee = True
                  End If
               End If
            '商標
            'Modify by Morgan 2010/11/8 FCT不扣
            'ElseIf (strSystemType = "T" Or strSystemType = "FCT") Then
            ElseIf (strSystemType = "T") Then
               If InStr("204,205", strProperty) > 0 Then
                  bolXFee = True
                  '2013/8/19 ADD BY SONIA 葉經理說訴願的言詞辯論為商標處的人處理,故不扣出庭費T-182351
                  strExc(0) = "select b.cp10,c.cp10 from caseprogress a,caseprogress b,caseprogress c where a.cp09='" & strCP09 & "'" & _
                     " and a.cp43=b.cp09(+) and b.cp43=c.cp09(+)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If "" & RsTemp.Fields(0) = "401" Or "" & RsTemp.Fields(1) = "401" Then bolXFee = False
                  End If
                  '2013/8/19 END
               ElseIf InStr("403,408,407", strProperty) > 0 Then
                  strExc(0) = "select * from caseprogress a where cp09='" & strCP09 & "'" & _
                     " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('204','205'))" & _
                     " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('204','205') and b.cp16>0)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolXFee = True
                  End If
               End If
            End If
         End If
         'end 2007/9/19
         
         If strSystemType = "FCT" Then
         
'Remove by Morgan 2010/4/22 因目前1j09 沒有設定且相關計算也有問題故取消
'            If douFCTamount <> 0 Then
''               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
''                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - Val(adoacc0x0.Fields("a0x03").Value) - douFCTAamount) * Val(strCon3), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'               If AccNoToSalesNo("417201") <> "" Then
'                  strSalesMan = AccNoToSalesNo("417201")
'               End If
'               'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
'               If bolXFee = True And bolXFeeDone = False Then
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - Val(adoacc0x0.Fields("a0x03").Value) - douFCTAamount) * Val(strCon3) - 10000, FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'                  bolXFeeDone = True
'               Else
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - Val(adoacc0x0.Fields("a0x03").Value) - douFCTAamount) * Val(strCon3), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'               End If
'               'end 2007/9/19
'               douAmount = Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - douFCTAamount) * Val(strCon3), FAmount))
'               '2005/5/3 ADD BY SONIA
'               '2005/10/25 MODIFY BY SONIA
'               'If PUB_CHKCUST(Text4) = True Then
'               If Option1.Value Then
'                  strCustNo = Text4
'               Else
'                  If Option2.Value Then
'                     strCustNo = Text6
'                  Else
'                     strCustNo = Text8
'                  End If
'               End If
'               If PUB_CHKCUST(strCustNo) = True Then
'               '2005/10/25 END
'                  adoTaie.Execute "UPDATE acc1p0 SET A1P14='" & Mid(m_NAME, 1, 4) & "/'||A1P14 WHERE A1P01='1' AND A1P02='F' AND A1P03='" & strSerialNo & "' AND A1P04='" & strItemNo & "'"
'               End If
'               '2005/5/3 END
'            End If
'end 2010/4/22

            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
            
'Remove by Morgan 2010/4/22 因目前1j09 沒有設定且相關計算也有問題故取消
'            If douFCTAamount <> 0 Then
''               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
''                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417202', 0, " & Val(Format(douFCTAamount * Val(strCon3), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'               If AccNoToSalesNo("417202") <> "" Then
'                  strSalesMan = AccNoToSalesNo("417202")
'               End If
'               'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
'               If bolXFee = True And bolXFeeDone = False Then
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417202', 0, " & Val(Format(douFCTAamount * Val(strCon3) - 10000, FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'                  bolXFeeDone = True
'               Else
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417202', 0, " & Val(Format(douFCTAamount * Val(strCon3), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'               End If
'               'end 2007/9/19
'
'               douAmount = Val(Format(douFCTAamount * Val(strCon3), FAmount))
'            Else
'end 2010/4/22

               adoaccsum.CursorLocation = adUseClient
               adoaccsum.Open "select cpm11, cpm12 from casepropertymap where cpm01 = '" & strSystemType & "' and cpm02 = '" & strProperty & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoaccsum.RecordCount <> 0 Then
                  If IsNull(adoaccsum.Fields("cpm11").Value) Then
'                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                                     "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
                     If AccNoToSalesNo("417201") <> "" Then
                        strSalesMan = AccNoToSalesNo("417201")
                     End If
                     'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
                     If bolXFee = True And bolXFeeDone = False Then
                        'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                        'modify by sonia 2021/3/22 +a1p23存A1K01
                        adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                        ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21, a1p23) values " & _
                                        "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value) - 10000, FAmount)) & _
                                        ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                        bolXFeeDone = True
                     Else
                        'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                        'modify by sonia 2021/3/22 +a1p23存A1K01
                        adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                        ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                        "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & _
                                        ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                     End If
                     'end 2007/9/19
                  Else
'                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                                     "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & adoaccsum.Fields("cpm11").Value & "', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
                     If adoaccsum.Fields("cpm11").Value = "417202" Then
                        If adoquery.State = adStateOpen Then
                           adoquery.Close
                        End If
                        adoquery.CursorLocation = adUseClient
                        adoquery.Open "select st03 from staff where st01 = '" & StrStaff & "'", adoTaie, adOpenStatic, adLockReadOnly
                        If adoquery.RecordCount <> 0 Then
                           If IsNull(adoquery.Fields("st03").Value) = False Then
                              If Mid(adoquery.Fields("st03").Value, 1, 1) = "F" Then
                                 strAccNo = "417201"
                              End If
                           End If
                        End If
                        adoquery.Close
                        If AccNoToSalesNo(strAccNo) <> "" Then
                           strSalesMan = AccNoToSalesNo(strAccNo)
                        End If
                        
                        'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
                        If bolXFee = True And bolXFeeDone = False Then
                           'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                           'modify by sonia 2021/3/22 +a1p23存A1K01
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                           ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                           "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value) - 10000, FAmount)) & _
                                           ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                           bolXFeeDone = True
                        Else
                           'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                           'modify by sonia 2021/3/22 +a1p23存A1K01
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                           ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                           "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & _
                                           ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                        End If
                        'end 2007/9/19
                        
                     '2006/8/7 ADD BY SONIA 科目雖為417201但承辦人仍為內商人員
                     Else
                        If adoaccsum.Fields("cpm11").Value = "417201" And strDept = "T" Then
                           strAccNo = "417202"
                           If AccNoToSalesNo(strAccNo) <> "" Then
                              strSalesMan = AccNoToSalesNo(strAccNo)
                           End If
                           
                           'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
                           If bolXFee = True And bolXFeeDone = False Then
                              'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                              'modify by sonia 2021/3/22 +a1p23存A1K01
                              adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                              ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                              "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value) - 10000, FAmount)) & _
                                              ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                              bolXFeeDone = True
                           Else
                              'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                              'modify by sonia 2021/3/22 +a1p23存A1K01
                              adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                              ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                              "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & _
                                              ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                           End If
                           'end 2007/9/19
                     '2006/8/7 END
                        Else
                           If strDept = "S" Then
                              'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                              'modify by sonia 2021/3/22 +a1p23存A1K01
                              adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                              ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                              "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '4121', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & _
                                              ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                           Else
                              'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
                              If bolXFee = True And bolXFeeDone = False Then
                                 'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                                 'modify by sonia 2021/3/22 +a1p23存A1K01
                                 adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                                 ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                                 "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & adoaccsum.Fields("cpm11").Value & "', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value) - 10000, FAmount)) & _
                                                 ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                                 bolXFeeDone = True
                              Else
                                 'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                                 'modify by sonia 2021/3/22 +a1p23存A1K01
                                 adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                                 ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                                 "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & adoaccsum.Fields("cpm11").Value & "', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & _
                                                 ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                              End If
                              'end 2007/9/19
                           End If
                        End If
                     End If
                     
                  End If
                  
               Else
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                                  "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
                  If AccNoToSalesNo("417201") <> "" Then
                     strSalesMan = AccNoToSalesNo("417201")
                  End If
                  
                  'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
                  If bolXFee = True And bolXFeeDone = False Then
                     'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                     'modify by sonia 2021/3/22 +a1p23存A1K01
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                     ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                     "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value) - 10000, FAmount)) & _
                                     ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                     bolXFeeDone = True
                  Else
                     'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                     'modify by sonia 2021/3/22 +a1p23存A1K01
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                     ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                     "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & _
                                     ", '" & strA1p14 & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                  End If
                  'end 2007/9/19
                  
               End If
               adoaccsum.Close
               douAmount = Val(Format(Val(adoacc0x0.Fields("a0x09").Value), FAmount))
               
'            End If'Remove by Morgan 2010/4/22 因目前1j09 沒有設定且相關計算也有問題故取消

         Else
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
            If strDept = "P" Then
               '2007/6/11 MODIFY BY SONIA 改判斷非台灣
               'If strNation = "020" Or strNation = "013" Then
               If strNation <> "000" Then
                  strAccNo = "411103"
               End If
            End If
            If strDept = "T" Then
               '2007/6/11 MODIFY BY SONIA 改判斷非台灣
               'If strNation = "020" Then
               If strNation <> "000" Then
                  strAccNo = "410103"
               End If
            End If
            
            'Added by Morgan 2016/4/8 沒有點數(acc1n0)時規則也要一樣 Ex.P-56989(X10408536->M10501189)
            'FMP
            If strAccNo = "411103" And Left(strSalesMan, 1) = "F" Then
               If strDept = "FCP" Then
                  strAccNo = "417102"
               Else
                  strAccNo = "411106"
               End If
            'FMT
            ElseIf strAccNo = "410103" And Left(strSalesMan, 1) = "F" Then
               strAccNo = "410109"
            End If
            'end 2016/4/8
                  
            If strAccNo <> "" And Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - (Val(adoacc0x0.Fields("a0x03").Value) / IIf(Val(strCon3) = 0, 1, Val(strCon3)))) * Val(strCon3), DAmount)) > 0 Then
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - (Val(adoacc0x0.Fields("a0x03").Value) / IIf(Val(strCon3) = 0, 1, Val(strCon3)))) * Val(strCon3), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
               'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
               If bolXFee = True And bolXFeeDone = False Then
                  'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                  'modify by sonia 2021/3/22 +a1p23存A1K01
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                  ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                  "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - (Val(adoacc0x0.Fields("a0x03").Value) / IIf(Val(strCon3) = 0, 1, Val(strCon3)))) * Val(strCon3) - 10000, FAmount)) & _
                                  ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & adoacc0x0.Fields("A1K01").Value & "/" & strCurrency & Format("" & adoacc0x0.Fields("A1K08").Value, "0.00"), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                  bolXFeeDone = True
               Else
                  'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                  'modify by sonia 2021/3/22 +a1p23存A1K01
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                  ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                  "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - (Val(adoacc0x0.Fields("a0x03").Value) / IIf(Val(strCon3) = 0, 1, Val(strCon3)))) * Val(strCon3), FAmount)) & _
                                  ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & adoacc0x0.Fields("A1K01").Value & "/" & strCurrency & Format("" & adoacc0x0.Fields("A1K08").Value, "0.00"), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
               End If
               'end 2007/9/19
               
               douAmount = Val(Format((Val(adoacc0x0.Fields("a0x11").Value)) * Val(strCon3), FAmount))
            Else
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("A0X09").Value - adoacc0x0.Fields("A0X03").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("A0X11").Value & ")"
               
               'Modify by Morgan 2007/9/19 加判斷是否扣出庭費10000
               If bolXFee = True And bolXFeeDone = False Then
                  'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                  'modify by sonia 2021/3/22 +a1p23存A1K01
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                  ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                  "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("A0X09").Value - adoacc0x0.Fields("A0X03").Value) - 10000 & _
                                  ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & adoacc0x0.Fields("A1K01").Value & "/" & strCurrency & Format("" & adoacc0x0.Fields("A1K08").Value, "0.00"), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("A0X11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                  bolXFeeDone = True
               Else
                  'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
                  'modify by sonia 2021/3/22 +a1p23存A1K01
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                                  ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                                  "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("A0X09").Value - adoacc0x0.Fields("A0X03").Value) & _
                                  ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & adoacc0x0.Fields("A1K01").Value & "/" & strCurrency & Format("" & adoacc0x0.Fields("A1K08").Value, "0.00"), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("A0X11").Value & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
               End If
               'end 2007/9/19
               
               douAmount = Val(adoacc0x0.Fields("A0X09").Value - adoacc0x0.Fields("A0X03").Value)
            End If
         End If
         '2005/5/3 ADD BY SONIA
         '2005/10/25 MODIFY BY SONIA
         'If PUB_CHKCUST(Text4) = True Then
         If Option1.Value Then
            strCustNo = Text4
         Else
            If Option2.Value Then
               strCustNo = Text6
            Else
               strCustNo = Text8
            End If
         End If
         
         'Modify By Sindy 2015/11/5 取消特殊客戶
'         '更新特定客戶之摘要
'         If PUB_CHKCUST(strCustNo) = True Then
'         '2005/10/25 END
'            adoTaie.Execute "UPDATE acc1p0 SET A1P14='" & Mid(m_NAME, 1, 4) & "/'||A1P14 WHERE A1P01='1' AND A1P02='F' AND A1P03='" & strSerialNo & "' AND A1P04='" & strItemNo & "'"
'         End If
'         '2005/5/3 END
         
      End If
      '2005/8/10 END
      
      '規費科目
      If strF <> "" Then
         strAccNo = strF
         Select Case strSystemType
            Case "S"
               Select Case strNation
                  Case "000"
                     strAccNo = "220103"
                  Case "020"
                     strAccNo = "220112"
                  Case Else
                     strAccNo = "220105"
               End Select
            Case "T", "TF", "TS"  '94.1.20 加入TS
               Select Case strNation
                  Case "000"
                  Case Else
                     strAccNo = "220111"
               End Select
            '93.12.17 add BY SONIA
            Case "P", "PS"
               Select Case strNation
                  Case "000"
                     strAccNo = "220102"
                  Case Else
                     strAccNo = "220112"
               End Select
            '93.12.17 END
         End Select
      Else
         Select Case strSystemType
            Case "S"
               Select Case strNation
                  Case "000"
                     strAccNo = "220103"
                  Case "020"
                     strAccNo = "220112"
                  Case Else
                     strAccNo = "220105"
               End Select
            Case "T", "TF", "TS"  '94.1.20 加入TS
               Select Case strNation
                  Case "000"
                     strAccNo = "220101"
                  Case Else
                     strAccNo = "220111"
               End Select
            Case "P", "PS"
               '93.12.17 MODIFY BY SONIA
               'strAccNo = "220102"
               Select Case strNation
                  Case "000"
                     strAccNo = "220102"
                  Case Else
                     strAccNo = "220112"
               End Select
               '93.12.17 END
            Case "FCT"
               strAccNo = "220103"
            Case "FCP", "FG"
               strAccNo = "220104"
            Case "CFT", "CFC"
               strAccNo = "220105"
            Case "CFP", "CPS"
               strAccNo = "220106"
            Case Else
               strAccNo = "220101"
         End Select
         'Modify By Cheng 2003/09/15
         'Begin
'            If strAccNo = "2201" Then
'               strAccNo = strF
'            End If
         'End
      End If
      
      Select Case strSystemType
         Case "S"
            '92.4.18 MODIFY BY SONIA
            'strDept = "S"
            If strNation = "000" Then
               strDept = "FCT"
            Else
               strDept = "CFT"
            End If
            '92.4.18 END
         Case "T", "TF"
            strDept = "T"
         Case "P", "PS"
            strDept = "P"
         Case "FCT"
            strDept = "FCT"
            'If adoaccsum.State = adStateOpen Then
            '   adoaccsum.Close
            'End If
            'adoaccsum.CursorLocation = adUseClient
            'adoaccsum.Open "select st03 from staff where st01 = '" & strStaff & "'", adoTaie, adOpenStatic, adLockReadOnly
            'If adoaccsum.RecordCount <> 0 Then
            '   If IsNull(adoaccsum.Fields("st03").Value) = False Then
            '      If Mid(adoaccsum.Fields("st03").Value, 1, 2) = "P2" Then
            '         strDept = "T"
            '      End If
            '   End If
            'End If
            'adoaccsum.Close
         Case "FCP", "FG"
            strDept = "FCP"
         Case "CFT", "CFC"
            strDept = "CFT"
         Case "CFP", "CPS"
            strDept = "CFP"
         Case "L"
            strDept = "L"
         'modify by sonia 2016/8/3 +LIN
         Case "FCL", "CFL", "LIN"
            strDept = "FCL"
         Case Else
            '92.10.9 MODIFY BY SONIA
            'strDept = "TOT"
            strDept = "T"
            '92.10.9 END
      End Select
      
      'Add by Morgan 2007/9/19 台灣案專利商標出庭費控管
      If bolXFee = True And bolXFeeDone = True Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         'Modify By Sindy 2015/11/5 strCaseNo & "/" & strCaseProperty ==> strA1p14
         'modify by sonia 2021/3/22 +a1p23存A1K01
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                         ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                         "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, 10000" & _
                         ", '" & strA1p14 & "/出庭費', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
      End If
      'end 2007/9/19
      
      '規費>0
      If adoacc0x0.Fields("a0x03").Value <> 0 Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         
         If Val(adoacc0x0.Fields("a0x09").Value) >= Val(adoacc0x0.Fields("a0x03").Value) Then
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                            "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x03").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ")"
            '92.10.21 MODIFY BY SONIA
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
            '                "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x03").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ")"
            If strAccNo = "220105" Or strAccNo = "220106" Then
                'Modify By Cheng 2004/03/31
                'CFT, CFP摘要帶的金額為總額(規費+服務費)
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x03").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "/" & Val(adoacc0x0.Fields("a0x03").Value) & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ")"
               'Modified by Lydia 2016/07/14 小數捨去Trunc
               'modify by sonia 2021/3/22 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(adoacc0x0.Fields("a0x03").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "/" & IIf(strDept = "CFT" Or strDept = "CFP", Val(adoacc0x0.Fields("a0x09").Value), Val(adoacc0x0.Fields("a0x03").Value)) & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
                'End
            ElseIf strAccNo = "220111" Or strAccNo = "220112" Then
               '2007/6/26 modify by sonia 大陸專利商標摘要帶的金額為總額(規費+服務費)
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
               '                "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x03").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "/" & Val(adoacc0x0.Fields("a0x03").Value) & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ")"
               'Modified by Lydia 2016/07/14 小數捨去Trunc
               'modify by sonia 2021/3/22 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(adoacc0x0.Fields("a0x03").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "/" & Val(adoacc0x0.Fields("a0x09").Value) & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
            'Added by Lydia 2016/07/14 規費不要有小數
            ElseIf Left(strAccNo, 4) = "2201" Then
               'modify by sonia 2021/3/22 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(adoacc0x0.Fields("a0x03").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
            'end 2016/07/14
            Else
               'modify by sonia 2021/3/22 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x03").Value) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
            End If
            '92.10.21 END
         Else
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                            "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x09").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ")"
            '92.10.21 MODIFY BY SONIA
            If strAccNo = "220105" Or strAccNo = "220106" Or strAccNo = "220111" Or strAccNo = "220112" Then
               'Modified by Lydia 2016/07/14 小數捨去Trunc
               'modify by sonia 2021/3/22 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(adoacc0x0.Fields("a0x09").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "/" & IIf(IsNull(adoacc0x0.Fields("a0x09").Value), "", adoacc0x0.Fields("a0x09").Value) & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
            'Added by Lydia 2016/07/14 規費不要有小數
            ElseIf Left(strAccNo, 4) = "2201" Then
               'modify by sonia 2021/3/22 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Trunc(Val(adoacc0x0.Fields("a0x09").Value)) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
            'end 2016/07/14
            Else
               'modify by sonia 2021/3/22 +a1p23存A1K01
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                               ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21,a1p23) values " & _
                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x09").Value) & _
                               ", '" & strCaseNo & "/" & strCaseProperty & "', null, '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ", '" & adoacc0x0.Fields("a1k01").Value & "')"
            End If
            '92.10.21 END
         End If
         douAmount = Val(adoacc0x0.Fields("a0x03").Value)
      End If
      
NextRec:

      '更新是否結清
      '2005/10/17 MODIFY BY SONIA
      'adoTaie.Execute "update acc1k0 set a1k29 = '" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "", adoacc0x0.Fields("a0x10").Value) & "', a1k30 = " & Val(Format(douAmount, FAmount)) & " where a1k01 = '" & adoacc0x0.Fields("a0x02").Value & "'"
      'Modified by Lydia 2017/03/07 是否結清,限制輸入Y
      'adoTaie.Execute "update acc1k0 set a1k29 = '" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "", adoacc0x0.Fields("a0x10").Value) & "' where a1k01 = '" & adoacc0x0.Fields("a0x02").Value & "'"
      adoTaie.Execute "update acc1k0 set a1k29 = '" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "", "Y") & "' where a1k01 = '" & adoacc0x0.Fields("a0x02").Value & "'"
      '2005/10/17 END
      adoacc0x0.MoveNext
   Loop
   adoacc0x0.Close
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1p08), sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(1).Value) = False Then
         douAmount = adoaccsum.Fields(1).Value
      Else
         douAmount = 0
      End If
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         douAmount = douAmount - adoaccsum.Fields(0).Value
      End If
   Else
      douAmount = Val(strCon4)
   End If
   adoaccsum.Close
   '沒溢收
   If Val(Text2) = 0 Then
       '2005/8/12 MODIFY BY SONIA 差額更新在收入科目
       'adoTaie.Execute "update acc1p0 set a1p08 = a1p08 + " & douAmount & " where a1p01 = '1' and a1p02 = 'F' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text3 & "'"
       adoTaie.Execute "update acc1p0 set a1p08 = a1p08 + " & douAmount & " where a1p01 = '1' and a1p02 = 'F' and a1p03 = '" & W_strSerialNo & "' and a1p04 = '" & Text3 & "'"
       '2005/8/12 END
   '有溢收
   Else
      '借-貸>0
      If douAmount <> 0 Then
         If Option1.Value Then
            strCustNo = Text4
            strCompany = Text5
         Else
            If Option2.Value Then
               strCustNo = Text6
               strCompany = Text7
            Else
               strCustNo = Text8
               strCompany = Text9
            End If
         End If
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
         '2005/4/6 MODIFY BY SONIA
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
         '                "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '2401', 0, " & douAmount & ", '" & ChgSQL(strCompany) & "/" & Text2 & "/" & strItemNo & "/" & Text10 & "', null, null, " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strCon3) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format(douAmount / IIf(Val(strCon3) = 0, 1, Val(strCon3)), FAmount)) & ")"
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08" & _
                         ", a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21" & _
                         ", A1P30) values " & _
                         "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '2401', 0, " & douAmount & _
                         ", '" & ChgSQL(strCompany) & "/" & Text2 & "/" & strItemNo & "/" & Text10 & "', null, null, " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strCon3) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format(douAmount / IIf(Val(strCon3) = 0, 1, Val(strCon3)), FAmount)) & _
                         ", '" & Text10 & "' )"
         '2005/4/6 END
      End If
   End If
   
   'Added by Morgan 2023/5/22 出庭費220113要放在代收款項2407xx後面，統一放在最後
   'Removed by Morgan 2023/5/23 國外收款目前順序正確可以不用
   'strSql = "update acc1p0 a set a1p03=(select lpad(mxsn+sn,3,'0')" & _
      " from (select a1p03 odsn,rownum sn from acc1p0 where a1p04=a.a1p04 and a1p05=a.a1p05 order by a1p03) x" & _
      ",(select max(a1p03) mxsn from acc1p0 b where  a1p04=a.a1p04 and a1p05<>a.a1p05) y where odsn=a.a1p03)" & _
      " where a1p04='" & Text3 & "' and a1p01='L' and a1p05='220113'" & _
      " and exists(select * From acc1p0 b where a1p01=a.a1p01 and a1p04=a.a1p04 and a1p05 like '2407%' and a1p03>a.a1p03)"
   'adoTaie.Execute strSql, intI
   'end 2023/5/23
   'end 2023/5/22
      
   adoTaie.Execute "delete from acc0x0 where a0x15 = '" & strUserNum & "'"
Checking:
   If Err.Number = 0 Then
      'Removed by Morgan 2012/5/31 不要彈訊息--婧瑄
      'If m_bolAlert Then MsgBox "下列請款單號有分配點數，請確認結果是否無誤！" & vbCrLf & vbCrLf & m_strAlertMsg
      'end 2012/5/31
      Exit Sub
   End If
   adoTaie.Execute "delete from acc0x0 where a0x15 = '" & strUserNum & "'"
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(國外收款資料(交易檔))
'
'*************************************************
'Private Sub Acc0z0Save()
'Dim strCaseNo As String
'Dim strCaseProperty As String
'Dim strCustomer As String
'Dim strSalesMan As String
'Dim strCurrency As String
'Dim strExchange As String
'Dim strSerialNo As String
'Dim strSystemType As String
'Dim strAccNo As String
'Dim intArgument As Integer
'Dim douFCTamount As Double
'Dim douFCTAamount As Double
'Dim douAmount As Double
'Dim strDept As String
'Dim stra1p22 As String
'Dim stra1p27 As String
'Dim strCompany As String
'Dim strMan As String
'Dim strCustNo As String
'Dim strProperty As String

'On Error GoTo Checking

'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      stra1p22 = "'" & adoaccsum.Fields("a1p22").Value & "'"
'      stra1p27 = "'" & "Y" & "'"
'      adoTaie.Execute "update acc1p0 set a1p27 = 'Y' where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "'"
'   Else
'      stra1p22 = "null"
'      stra1p27 = "null"
'   End If
'   adoaccsum.Close
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & strItemNo & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoTaie.Execute "delete from acc0x0"
'      Exit Sub
'   End If
'   adoaccsum.Close
'   adoTaie.Execute "delete from acc0z0 where a0z01 = '" & Text3 & "'"
'   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text3 & "' and a1p08 <> 0"
'   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & TEXT3 & "' and a1p05 = '2401' and a1p08 <> 0"
'   adoacc0x0.CursorLocation = adUseClient
'   adoacc0x0.Open "select * from acc1k0, acc0x0 where a1k01 = a0x02 order by a1k01 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc0x0.EOF = False
'      If IsNull(adoacc0x0.Fields("a1k13").Value) Then
'         strCaseNo = ""
'         strSystemType = ""
'      Else
'         strCaseNo = adoacc0x0.Fields("a1k13").Value & adoacc0x0.Fields("a1k14").Value & adoacc0x0.Fields("a1k15").Value & adoacc0x0.Fields("a1k16").Value
'         strSystemType = adoacc0x0.Fields("a1k13").Value
'      End If
'      adoacc0z0.CursorLocation = adUseClient
'      adoacc0z0.Open "select * from acc0z0 where a0z01 = '" & adoacc0x0.Fields("a0x01").Value & "' and a0z02 = '" & adoacc0x0.Fields("a0x02").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'      If adoacc0z0.RecordCount = 0 Then
'         adoacc0z0.AddNew
'      End If
'      adoacc0z0.Fields("a0z01").Value = adoacc0x0.Fields("a0x01").Value
'      adoacc0z0.Fields("a0z02").Value = adoacc0x0.Fields("a0x02").Value
'      If IsNull(adoacc0x0.Fields("a0x08").Value) Then
'         adoacc0z0.Fields("a0z03").Value = Null
'      Else
'         adoacc0z0.Fields("a0z03").Value = adoacc0x0.Fields("a0x08").Value
'      End If
'      If IsNull(adoacc0x0.Fields("a0x11").Value) Then
'         adoacc0z0.Fields("a0z04").Value = 0
'      Else
'         adoacc0z0.Fields("a0z04").Value = adoacc0x0.Fields("a0x11").Value
'      End If
'      adoacc0z0.Fields("a0z06").Value = Val(ACDate(ServerDate))
'      adoacc0z0.Fields("a0z07").Value = ServerTime
'      adoacc0z0.Fields("a0z08").Value = strUserNum
'      adoacc0z0.UpdateBatch
'      adoacc0z0.Close
'      adocaseprogress.CursorLocation = adUseClient
'      adocaseprogress.Open "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10 from caseprogress, salesno, staff, casepropertyMap, patent, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' union " & _
'                           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10 from caseprogress, salesno, staff, casepropertyMap, trademark, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' union " & _
'                           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10 from caseprogress, salesno, staff, casepropertyMap, lawcase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' union " & _
'                           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10 from caseprogress, salesno, staff, casepropertyMap, hirecase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "' union " & _
'                           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu04, nvl(cu05||cu88||cu89||cu90, cu06)) as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10 from caseprogress, salesno, staff, casepropertyMap, servicepractice, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and cp60 = '" & adoacc0x0.Fields("a0x02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocaseprogress.RecordCount <> 0 Then
'         If IsNull(adocaseprogress.Fields("CustNo").Value) Then
'            strCustNo = ""
'         Else
'            strCustNo = adocaseprogress.Fields("Custno").Value
'         End If
'         If IsNull(adocaseprogress.Fields("Company").Value) Then
'            strCompany = ""
'         Else
'            strCompany = adocaseprogress.Fields("Company").Value
'         End If
'         If IsNull(adocaseprogress.Fields("Property").Value) Then
'            strCaseProperty = ""
'         Else
'            strCaseProperty = adocaseprogress.Fields("Property").Value
'         End If
'         If IsNull(adocaseprogress.Fields("cp13").Value) Then
'            strSalesMan = ""
'         Else
'            strSalesMan = adocaseprogress.Fields("cp13").Value
'         End If
'         If IsNull(adocaseprogress.Fields("Man").Value) Then
'            strMan = ""
'         Else
'            strMan = adocaseprogress.Fields("Man").Value
'         End If
'         If IsNull(adocaseprogress.Fields("st03").Value) Then
'            strDept = ""
'         Else
'            strDept = adocaseprogress.Fields("st03").Value
'         End If
'         If IsNull(adocaseprogress.Fields("cp10").Value) Then
'            strProperty = ""
'         Else
'            strProperty = adocaseprogress.Fields("cp10").Value
'         End If
'      Else
'         strCaseProperty = ""
'         strSalesMan = ""
'         strDept = ""
'         strCompany = ""
'         strMan = ""
'         strProperty = ""
'      End If
'      adocaseprogress.Close
'      If strSystemType = "" Then
'         If IsNull(adoacc0x0.Fields("a1k13").Value) = False Then
'            strSystemType = adoacc0x0.Fields("a1k13").Value
'         End If
'      End If
'      If strSystemType = "FCT" Then
'         adoacc1k0.CursorLocation = adUseClient
'         adoacc1k0.Open "select sum(a1l05) from acc1l0, acc1j0 where a1l03 = a1j01 and a1l04 = a1j02 and a1l01 = '" & adoacc0x0.Fields("a0x02").Value & "' and a1j09 = '417201'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoacc1k0.RecordCount <> 0 Then
'            If IsNull(adoacc1k0.Fields(0).Value) Then
'               douFCTamount = 0
'            Else
'               douFCTamount = adoacc1k0.Fields(0).Value
'            End If
'         Else
'             douFCTamount = 0
'         End If
'         adoacc1k0.Close
'         adoacc1k0.CursorLocation = adUseClient
'         adoacc1k0.Open "select sum(a1l05) from acc1l0, acc1j0 where a1l03 = a1j01 and a1l04 = a1j02 and a1l01 = '" & adoacc0x0.Fields("a0x02").Value & "' and a1j09 = '417202'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoacc1k0.RecordCount <> 0 Then
'            If IsNull(adoacc1k0.Fields(0).Value) Then
'               douFCTAamount = 0
'            Else
'               douFCTAamount = adoacc1k0.Fields(0).Value
'            End If
'         Else
'             douFCTAamount = 0
'         End If
'         adoacc1k0.Close
'      End If
'      adoacc1k0.CursorLocation = adUseClient
'      adoacc1k0.Open "select a1k18, a1k10 from acc1k0 where a1k01 = '" & adoacc0x0.Fields("a0x02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc1k0.RecordCount <> 0 Then
'         If IsNull(adoacc1k0.Fields(0).Value) Then
'            strCurrency = ""
'         Else
'            strCurrency = adoacc1k0.Fields(0).Value
'         End If
'         If IsNull(adoacc1k0.Fields(1).Value) Then
'            strExchange = ""
'         Else
'            strExchange = adoacc1k0.Fields(1).Value
'         End If
'      Else
'         strCurrency = ""
'         strExchange = ""
'      End If
'      adoacc1k0.Close
'      strAccNo = ""
'      If adoacc0x0.Fields("a0x03").Value <> 0 Then
'         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
'         Select Case strSystemType
'            Case "T"
'               strAccNo = "220101"
'               strDept = "T"
'            Case "P"
'               strAccNo = "220102"
'               strDept = "P"
'            Case "FCT"
'               strAccNo = "220103"
'               strDept = "FCT"
'            Case "FCP"
'               strAccNo = "220104"
'               strDept = "FCP"
'            Case "CFT"
'               strAccNo = "220105"
'               strDept = "CFT"
'            Case "CFP"
'               strAccNo = "220106"
'               strDept = "CFP"
'            Case Else
'               strAccNo = "2201"
'               strDept = "TOT"
'         End Select
'         If Val(adoacc0x0.Fields("a0x09").Value) >= Val(adoacc0x0.Fields("a0x03").Value) Then
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                            "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x03").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ")"
'         Else
'            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                            "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("a0x09").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(adoacc0x0.Fields("a0x11").Value) & ")"
'         End If
'         douAmount = Val(adoacc0x0.Fields("a0x03").Value)
'      End If
'      If Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value) > 0 Then
'         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
'         Select Case strSystemType
'            Case "P"
'               strAccNo = "411103"
'               strDept = "P"
'            Case "T"
'               strAccNo = "410103"
'               strDept = "T"
'            Case "CFT"
'               strAccNo = "4121"
'               strDept = "CFT"
'            Case "CFP"
'               strAccNo = "4131"
'               strDept = "CFP"
'            Case "L"
'               strAccNo = "4141"
'               strDept = "L"
'            Case "FCL"
'               strAccNo = "4161"
'               strDept = "FCL"
'            Case "FCP"
'               strAccNo = "4171"
'               strDept = "FCP"
'            Case "FCT"
'               strAccNo = "4172"
'               strDept = "FCT"
'            Case "FL"
'               strAccNo = "4161"
'               strDept = "FL"
'            Case Else
'               strAccNo = "41"
'               strDept = "TOT"
'         End Select
'         If strSystemType = "FCT" Then
'            If douFCTamount <> 0 Then
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - Val(adoacc0x0.Fields("a0x03").Value) - douFCTAamount) * Val(strCon3), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'               douAmount = Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - douFCTAamount) * Val(strCon3), FAmount))
'            End If
'            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
'            If douFCTAamount <> 0 Then
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417202', 0, " & Val(Format(douFCTAamount * Val(strCon3), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'               douAmount = Val(Format(douFCTAamount * Val(strCon3), FAmount))
'            Else
'               adoaccsum.CursorLocation = adUseClient
'               adoaccsum.Open "select cpm11, cpm12 from casepropertymap where cpm01 = '" & strDept & "' and cpm02 = '" & strProperty & "'", adoTaie, adOpenStatic, adLockReadOnly
'               If adoaccsum.RecordCount <> 0 Then
'                  If IsNull(adoaccsum.Fields("cpm11").Value) Then
'                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                                     "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'                  Else
'                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                                     "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & adoaccsum.Fields("cpm11").Value & "', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'                  End If
'               Else
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                                  "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '417201', 0, " & Val(Format(Val(adoacc0x0.Fields("a0x09").Value) - Val(adoacc0x0.Fields("a0x03").Value), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'               End If
'               adoaccsum.Close
'               douAmount = Val(Format(Val(adoacc0x0.Fields("a0x09").Value), FAmount))
'            End If
'         Else
'            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
'            If strAccNo <> "" And Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - (Val(adoacc0x0.Fields("a0x03").Value) / IIf(Val(strCon3) = 0, 1, Val(strCon3)))) * Val(strCon3), DAmount)) > 0 Then
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(Format((Val(adoacc0x0.Fields("a0x11").Value) - (Val(adoacc0x0.Fields("a0x03").Value) / IIf(Val(strCon3) = 0, 1, Val(strCon3)))) * Val(strCon3), FAmount)) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("a0x11").Value & ")"
'               douAmount = Val(Format((Val(adoacc0x0.Fields("a0x11").Value)) * Val(strCon3), FAmount))
'            Else
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                               "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '" & strAccNo & "', 0, " & Val(adoacc0x0.Fields("A0X09").Value - adoacc0x0.Fields("A0X03").Value) & ", '" & strCaseNo & "/" & strCaseProperty & "', '" & strSalesMan & "', '" & strCaseNo & "', " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & adoacc0x0.Fields("A0X11").Value & ")"
'               douAmount = Val(adoacc0x0.Fields("A0X09").Value - adoacc0x0.Fields("A0X03").Value)
'            End If
'         End If
'      End If
'      adoTaie.Execute "update acc1k0 set a1k29 = '" & IIf(IsNull(adoacc0x0.Fields("a0x10").Value), "", adoacc0x0.Fields("a0x10").Value) & "', a1k30 = " & Val(Format(douAmount, FAmount)) & " where a1k01 = '" & adoacc0x0.Fields("a0x02").Value & "'"
'      adoacc0x0.MoveNext
'   Loop
'   adoacc0x0.Close
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(a1p08) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         douAmount = Val(strCon4)
'      Else
'         douAmount = Val(strCon4) - adoaccsum.Fields(0).Value
'      End If
'   Else
'      douAmount = Val(strCon4)
'   End If
'   adoaccsum.Close
'   If Val(Text2) = 0 Then
'       adoTaie.Execute "update acc1p0 set a1p08 = a1p08 + " & douAmount & " where a1p01 = '1' and a1p02 = 'F' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text3 & "'"
'   Else
'      If douAmount <> 0 Then
'         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & strItemNo & "'", 3)
'         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p22, a1p27, a1p21) values " & _
'                         "('1', 'F', '" & strSerialNo & "', '" & strItemNo & "', '2401', 0, " & douAmount & ", '" & strItemNo & "', null, null, " & Val(strCon1) & ", '" & strCurrency & "', " & Val(strCon3) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & stra1p22 & ", " & stra1p27 & ", " & Val(Format(douAmount / IIf(Val(strCon3) = 0, 1, Val(strCon3)), FAmount)) & ")"
'      End If
'   End If
'   adoTaie.Execute "delete from acc0x0"
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   adoTaie.Execute "delete from acc0x0"
'   MsgBox Err.Description, , MsgText(5)
'End Sub

'*************************************************
'  計算並顯示溢收金額
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0x11), sum(a0x12), sum(a0x07) from acc0x0 where a0x01 = '" & Text3 & "' and a0x15 = '" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text2 = dblTotal
      Else
         'Text2 = Format(dblTotal - (Val(adoaccsum.Fields(0).Value) - Val(adoaccsum.Fields(2).Value)), FAmount)
         Text2 = Format(dblTotal - (Val(adoaccsum.Fields(0).Value)), FAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) = False Then
         Text2 = Val(Text2) + Val(Format((Val(adoaccsum.Fields(1).Value) / Val(strCon3)), FAmount))
      End If
   Else
      Text2 = dblTotal
   End If
   adoaccsum.Close
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 = MsgText(601) Then
      Exit Sub
   End If
   If Len(Text8) = 6 Then
      Text8 = AfterZero(Text8)
   'Add by Morgan 2007/3/1 八碼時要補'0'
   ElseIf Len(Text8) = 8 Then
      Text8 = Text8 & "0"
   'End 2007/3/1
   End If
   
   Text9 = FagentQuery(Text8, 2)
   '2005/5/20 ADD BY SONIA
   If Text9 = "" Then
      Text9 = FagentQuery(Text8, 1)
   End If
   '2005/5/20 END
   If Text9 = "" Then
      Text9 = CustomerQuery(Text8, 2)
   End If
   '2005/5/20 ADD BY SONIA
   If Text9 = "" Then
      Text9 = CustomerQuery(Text8, 1)
   End If
   
   If ExistCheck("fagent", "fa01", Mid(Text8, 1, 8), Label6, False) = False Then
      If ExistCheck("customer", "cu01", Mid(Text8, 1, 8), Label6) = False Then
         Cancel = True
         Text8.SetFocus
         Exit Sub
      End If
   End If
End Sub

'*************************************************
'  儲存資料表(請款單收款記錄資料)
'
'*************************************************
Private Sub Acc0x0Show()
Dim strA1K18 As String
Dim adoCP As ADODB.Recordset 'Add By Sindy 2015/10/21
Dim strCP09 As String, strGetA0K11 As String 'Add By Sindy 2015/10/21
   
   adoacc0x0.CursorLocation = adUseClient
   adoacc0x0.Open "select * from acc0x0 where a0x01 = '" & Text3 & "' and a0x15 = '" & strUserNum & "' order by a0x02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0z0.CursorLocation = adUseClient
   'Modified by Morgan 2021/4/21 + order by a0z02 asc
   adoacc0z0.Open "select * from acc0z0 where a0z01 = '" & Text3 & "' order by a0z02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0z0.EOF = False
      adoacc0x0.AddNew
      
      'Add By Sindy 2015/10/21
      '取得總收文號
      strExc(0) = "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(PA09,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, pa09 as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, patent, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and cp60 = '" & adoacc0z0.Fields("a0z02").Value & "' union " & _
         "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(TM10,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, tm10 as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, trademark, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and cp60 = '" & adoacc0z0.Fields("a0z02").Value & "' union " & _
         "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(LC15,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, lc15 as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, lawcase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and cp60 = '" & adoacc0z0.Fields("a0z02").Value & "' union " & _
         "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, nvl(cpm03, cpm04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, null as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, hirecase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and cp60 = '" & adoacc0z0.Fields("a0z02").Value & "' union " & _
         "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, nvl(cu05||cu88||cu89||cu90, nvl(cu04, cu06)) as Company, DECODE(SP09,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, st03, cp01, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cp10, cpm11, cpm12, cp14, sp09 as nation,cpm24,cpm25,cpm03,cp12 from caseprogress, salesno, staff, casepropertyMap, servicepractice, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and cp60 = '" & adoacc0z0.Fields("a0z02").Value & "'" & _
         " order by cp09 asc"
      intI = 1
      Set adoCP = ClsLawReadRstMsg(intI, strExc(0))
      strCP09 = "": strGetA0K11 = ""
      If intI = 1 Then
         '取得公司別
         strCP09 = adoCP.Fields("cp09")
         strExc(0) = "select cp09,GetA0k11('" & strCP09 & "') from caseprogress where cp09='" & strCP09 & "'"
         intI = 1
         Set adoCP = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strGetA0K11 = adoCP.Fields(1)
         End If
      End If
      adoCP.Close
      '2015/10/21 END
      
      adoacc0x0.Fields("a0x01").Value = adoacc0z0.Fields("a0z01").Value
      adoacc0x0.Fields("a0x02").Value = adoacc0z0.Fields("a0z02").Value
      adoacc1k0.CursorLocation = adUseClient
      adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & adoacc0z0.Fields("a0z02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc1k0.RecordCount <> 0 Then
         If IsNull(adoacc1k0.Fields("a1k09").Value) Then
            adoacc0x0.Fields("a0x03").Value = 0
         Else
            adoacc0x0.Fields("a0x03").Value = adoacc1k0.Fields("a1k09").Value
'cancel by sonia 2024/11/15 改在Acc0z0SaveNew
'            'add by sonia 2017/11/20 FCP及FMP案B類收文927其他翻譯且承辦人為外翻編號且相關總收文號為C類之結匯金額,已在結匯扣收入及請款單扣點數(規費),此處收款要從規費扣掉結匯金額M10605730(X10617049(FCP-047593)
'            adocaseprogress.CursorLocation = adUseClient
'            adocaseprogress.Open "select a1p07,a1w01,a1w02,cp60,cp61 from acc1w0,caseprogress,acc1p0 where a1w01='" & adoacc0z0.Fields("a0z02").Value & "' and substr(a1w02,1,1)='B' and a1w02=cp09(+) " & _
'                                 "and cp01 in ('P','FCP') and cp10='927' and substr(cp14,1,1)='F' and substr(cp43,1,1)='C' and cp61||a1w02=a1p23 and a1p07>0", adoTaie, adOpenStatic, adLockReadOnly
'            If adocaseprogress.RecordCount <> 0 Then
'               adoacc0x0.Fields("a0x03").Value = adoacc0x0.Fields("a0x03").Value - Val("" & adocaseprogress.Fields("a1p07"))
'            End If
'            adocaseprogress.Close
'            'end 2017/11/20
'end 2024/11/15
         End If
         '2010/3/15 ADD BY SONIA
         If IsNull(adoacc1k0.Fields("a1k11").Value) Then
            adoacc0x0.Fields("a0x16").Value = 0
         Else
            If IsNull(adoacc1k0.Fields("a1k06").Value) Then
               adoacc0x0.Fields("a0x16").Value = adoacc1k0.Fields("a1k11").Value
            Else
               'Modify By Sindy 2012/12/6
               'adoacc0x0.Fields("a0x16").Value = adoacc1k0.Fields("a1k11").Value - Val(adoacc1k0.Fields("A1K06").Value) * Val(adoacc1k0.Fields("a1k10").Value)
               adoacc0x0.Fields("a0x16").Value = adoacc1k0.Fields("a1k11").Value - Val(adoacc1k0.Fields("A1K06").Value)
               '2012/12/6 End
            End If
         End If
         '2010/3/15 END
         adocaseprogress.CursorLocation = adUseClient
         adocaseprogress.Open "select nvl(cpm03, cpm04) from caseprogress, casepropertymap where cp01 = cpm01 and cp10 = cpm02 and cp60 = '" & adoacc1k0.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocaseprogress.RecordCount <> 0 Then
            If IsNull(adocaseprogress.Fields(0).Value) Then
               adoacc0x0.Fields("a0x04").Value = Null
            Else
               adoacc0x0.Fields("a0x04").Value = adocaseprogress.Fields(0).Value
            End If
         Else
            adoacc0x0.Fields("a0x04").Value = Null
         End If
         adocaseprogress.Close
         If IsNull(adoacc1k0.Fields("a1k10").Value) Then
            adoacc0x0.Fields("a0x05").Value = 0
            adoacc0x0.Fields("a0x06").Value = 0
         Else
            If IsNull(adoacc1k0.Fields("a1k08").Value) Then
               adoacc0x0.Fields("a0x05").Value = 0
            Else
               adoacc0x0.Fields("a0x05").Value = adoacc1k0.Fields("a1k08").Value
            End If
            'Modify By Sindy 2012/12/6 外幣折讓
'            If IsNull(adoacc1k0.Fields("a1k06").Value) Then
'               adoacc0x0.Fields("a0x06").Value = 0
'            Else
'               adoacc0x0.Fields("a0x06").Value = adoacc1k0.Fields("a1k06").Value
'            End If
            If IsNull(adoacc1k0.Fields("a1k31").Value) Then
               adoacc0x0.Fields("a0x06").Value = 0
            Else
               adoacc0x0.Fields("a0x06").Value = adoacc1k0.Fields("a1k31").Value
            End If
            '2012/12/6 End
         End If
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select sum(a0z04) from acc0z0 where a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "' and a0z01 <> '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               adoacc0x0.Fields("a0x07").Value = 0
            Else
               adoacc0x0.Fields("a0x07").Value = Val(Format(adoquery.Fields(0).Value, FAmount))
            End If
         Else
            adoacc0x0.Fields("a0x07").Value = 0
         End If
         adoquery.Close
         If IsNull(adoacc1k0.Fields("a1k29").Value) Then
            adoacc0x0.Fields("a0x10").Value = ""
         Else
            adoacc0x0.Fields("A0X10").Value = adoacc1k0.Fields("a1k29").Value
         End If
         strA1K18 = "" & adoacc1k0.Fields("a1k18").Value 'Added by Morgan 2013/8/26
      End If
      adoacc1k0.Close
      If strCon2 = MsgText(601) Then
         adoacc0x0.Fields("a0x08").Value = Null
      Else
         '2013/8/19 modify by sonia
         'adoacc0x0.Fields("a0x08").Value = strCon2
         If strCon2 = "NTD" Then
            adoacc0x0.Fields("a0x08").Value = strCon2
         Else
            'Modified by Morgan 2013/8/23
            'adoacc0x0.Fields("a0x08").Value = adoacc1k0.Fields("a1k18").Value
            adoacc0x0.Fields("a0x08").Value = strA1K18
         End If
         '2013/8/19 end
      End If
      If IsNull(adoacc0z0.Fields("a0z12").Value) Then
         adoacc0x0.Fields("a0x12").Value = 0
      Else
         adoacc0x0.Fields("a0x12").Value = adoacc0z0.Fields("a0z12").Value
      End If
      If IsNull(adoacc0z0.Fields("a0z13").Value) Then
         If strGetA0K11 <> "" Then
            adoacc0x0.Fields("a0x13").Value = strGetA0K11
         Else
            adoacc0x0.Fields("a0x13").Value = Null
         End If
      Else
         adoacc0x0.Fields("a0x13").Value = adoacc0z0.Fields("a0z13").Value
      End If
      If IsNull(adoacc0z0.Fields("a0z07").Value) Then
         adoacc0x0.Fields("a0x14").Value = Null
      Else
         adoacc0x0.Fields("a0x14").Value = adoacc0z0.Fields("a0z07").Value
      End If
      adoacc0x0.Fields("A0X11").Value = adoacc0z0.Fields("A0Z04").Value
      adoacc0x0.Fields("A0X09").Value = Val(Format(Val(adoacc0x0.Fields("A0X11").Value) * Val(strCon3), FAmount))
      'add by sonia 2018/12/26 M10706118收X10714005台幣金額A0X09<規費A0X03時改預設規費A0X03
      If adoacc0x0.Fields("A0X09").Value < adoacc0x0.Fields("A0X03").Value Then
         adoacc0x0.Fields("A0X09").Value = adoacc0x0.Fields("A0X03").Value
      End If
      'end 2018/12/26
      adoacc0x0.Fields("a0x15").Value = strUserNum
      adoacc0x0.UpdateBatch
      adoacc0z0.MoveNext
   Loop
   adoacc0z0.Close
   adoacc0x0.Close
End Sub

'*************************************************
'  儲存資料表(暫收款資料)
'
'*************************************************
Private Sub Acc120Save()
On Error GoTo Checking
   adoacc120.CursorLocation = adUseClient
   adoacc120.Open "select * from acc120 where a1201 = '" & Text10 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc120.RecordCount = 0 Then
      adoacc120.AddNew
   End If
   If Text10 <> "" Then
      adoacc120.Fields("a1201").Value = Text10
   Else
      adoacc120.Fields("a1201").Value = AutoNo(MsgText(809), 5)
   End If
   Text10 = adoacc120.Fields("a1201").Value
   If IsNull(adoacc0y0.Fields("a0y02").Value) Then
      adoacc120.Fields("a1202").Value = Null
   Else
      adoacc120.Fields("a1202").Value = adoacc0y0.Fields("a0y02").Value
   End If
   If Option1.Value Then
      adoacc120.Fields("a1203").Value = Text4
   Else
      If Option2.Value Then
         adoacc120.Fields("a1203").Value = Text6
      Else
         adoacc120.Fields("a1203").Value = Text8
      End If
   End If
   If IsNull(adoacc0y0.Fields("a0y03").Value) Then
      adoacc120.Fields("a1204").Value = Null
   Else
      adoacc120.Fields("a1204").Value = adoacc0y0.Fields("a0y03").Value
   End If
   If IsNull(adoacc0y0.Fields("a0y04").Value) Then
      adoacc120.Fields("a1205").Value = 0
   Else
      adoacc120.Fields("a1205").Value = adoacc0y0.Fields("a0y04").Value
   End If
   If Text2 <> MsgText(601) Then
      adoacc120.Fields("a1207").Value = Val(Text2)
   Else
      adoacc120.Fields("a1207").Value = 0
   End If
   adoacc120.Fields("a1209").Value = "2"
   If Text3 <> MsgText(601) Then
      adoacc120.Fields("a1210").Value = Text3
   End If
   adoacc120.Fields("a1214").Value = Val(strSrvDate(2))
   adoacc120.Fields("a1215").Value = ServerTime
   adoacc120.Fields("a1216").Value = strUserNum
   adoacc120.UpdateBatch
   adoacc120.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Modify By Sindy 2015/11/5 Mark
''2005/5/3 ADD BY SONIA
'Public Function PUB_CHKCUST(m_TEXT4 As String) As Boolean
'Dim adocheck As New ADODB.Recordset
'On Error GoTo Err
'
'   PUB_CHKCUST = False
'   m_NAME = "": m_YEAR = ""
'   '2005/5/17 MODIFY BY SONIA 加 X48334
'   '2005/5/20 MODIFY BY SONIA 加 X30072
'   '2005/6/10 MODIFY BY SONIA 加 X53230,X18923
'   '2005/7/25 MODIFY BY SONIA 加 X27575
'   '2005/9/14 MODIFY BY SONIA 加 X56559
'   If Mid(m_TEXT4, 1, 6) = "X19135" Or Mid(m_TEXT4, 1, 6) = "X16269" Or Mid(m_TEXT4, 1, 6) = "X49029" Or Mid(m_TEXT4, 1, 6) = "X54662" Or Mid(m_TEXT4, 1, 6) = "X24309" Or Mid(m_TEXT4, 1, 6) = "X11833" Or Mid(m_TEXT4, 1, 6) = "X49010" Or Mid(m_TEXT4, 1, 6) = "X52996" Or Mid(m_TEXT4, 1, 6) = "X48334" Or Mid(m_TEXT4, 1, 6) = "X30072" Or Mid(m_TEXT4, 1, 6) = "X53230" Or Mid(m_TEXT4, 1, 6) = "X18923" Or Mid(m_TEXT4, 1, 6) = "X27575" Or Mid(m_TEXT4, 1, 6) = "X56559" Then
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select * from CUSTOMER where CU01 = '" & ChgSQL(Mid(m_TEXT4, 1, 8)) & "' and CU02 = '" & ChgSQL(IIf(Mid(m_TEXT4, 9, 1) = "", "0", Mid(m_TEXT4, 9, 1))) & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount <> 0 Then
'         adocheck.MoveFirst
'            If IsNull(adocheck.Fields("CU04").Value) Then
'               If IsNull(adocheck.Fields("CU05").Value) Then
'                  If IsNull(adocheck.Fields("CU06").Value) Then
'                     m_NAME = MsgText(601)
'                  Else
'                     m_NAME = adocheck.Fields("CU06").Value
'                  End If
'               Else
'                  m_NAME = adocheck.Fields("CU05").Value
'               End If
'            Else
'               m_NAME = adocheck.Fields("CU04").Value
'            End If
'      End If
'      adocheck.Close
'      m_YEAR = strSrvDate(2) \ 10000
'      PUB_CHKCUST = True
'   End If
'
'   '2005/7/25 MODIFY BY SONIA 加 Y27575
'   '2005/9/14 MODIFY BY SONIA 加 Y51858
'   '2006/4/7  MODIFY BY SONIA 加 Y30072
'   If Mid(m_TEXT4, 1, 6) = "Y19135" Or Mid(m_TEXT4, 1, 6) = "Y51766" Or Mid(m_TEXT4, 1, 6) = "Y49010" Or Mid(m_TEXT4, 1, 6) = "Y51665" Or Mid(m_TEXT4, 1, 6) = "Y11833" Or Mid(m_TEXT4, 1, 6) = "Y27575" Or Mid(m_TEXT4, 1, 6) = "Y51858" Or Mid(m_TEXT4, 1, 6) = "Y30072" Then
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select * from fagent where fa01 = '" & ChgSQL(Mid(m_TEXT4, 1, 8)) & "' and fa02 = '" & ChgSQL(IIf(Mid(m_TEXT4, 9, 1) = "", "0", Mid(m_TEXT4, 9, 1))) & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount <> 0 Then
'         adocheck.MoveFirst
'            If IsNull(adocheck.Fields("fa04").Value) Then
'               If IsNull(adocheck.Fields("fa05").Value) Then
'                  If IsNull(adocheck.Fields("fa06").Value) Then
'                     m_NAME = MsgText(601)
'                  Else
'                     m_NAME = adocheck.Fields("fa06").Value
'                  End If
'               Else
'                  m_NAME = adocheck.Fields("fa05").Value
'               End If
'            Else
'               m_NAME = adocheck.Fields("fa04").Value
'            End If
'      End If
'      adocheck.Close
'      m_YEAR = strSrvDate(2) \ 10000
'      PUB_CHKCUST = True
'   End If
'
'   Exit Function
'Err:
'   MsgBox "錯誤 : " & Err.Description, vbCritical
'End Function
''2005/5/3 END

'Add By Sindy 2015/4/20
Private Sub txtA1K35_GotFocus()
   OpenIme
   InverseTextBox txtA1K35
End Sub
Private Sub txtA1K35_Validate(Cancel As Boolean)
   If txtA1K35.Enabled = False Then Exit Sub
   If txtA1K35.Text = "" Then Exit Sub
   
   '剔除跳行符號
   txtA1K35.Text = PUB_StringFilter(txtA1K35.Text)
   
   If Not CheckLengthIsOK(txtA1K35, txtA1K35.MaxLength) Then
      Cancel = True
   End If
End Sub
'2015/4/20 END

'Added by Morgan 2015/11/27
Private Sub DeliverInform()
   'Added by Morgan 2025/10/15 其他系統也要通知，改比照國內收款改公用函數--秀玲
   PUB_AccDeliverInform "1", Text3
   
'Removed by Morgan 2025/10/15
'   Dim stSubject As String, stContent As String
'   Dim adoRst As ADODB.Recordset
'
'   '所有T*案申請國家非台灣之非FMT案(即CP12非'F'字頭)，或是收款後送件CP141='2'之台灣案，於全額收款時發E-MAIL給承辦人
'   strExc(0) = "SELECT distinct A.*,NA03,CU04,DECODE(A3,'000',CPM03,CPM04) 案件性質,ST02" & _
'      " FROM (select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
'      ",TM05||SP05 案件名稱,SQLDATET(CP05) 收文日" & _
'      ",TM10||SP09 A3,TM23||SP08 A4,CP01,CP10,CP13,CP14" & _
'      " from acc0z0,acc1k0,CASEPROGRESS,trademark,servicepractice" & _
'      " where a0z01='" & Text3 & "' and a1k01(+)=a0z02 and a1k29='Y'" & _
'      " AND A1K13 LIKE 'T%' AND CP60(+)=A1K01 AND CP27||CP57 IS NULL" & _
'      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
'      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
'      " and ((nvl(tm10,sp09)<>'000' and cp12 not like 'F%') or cp141='2')" & _
'      ") A,STAFF,CASEPROPERTYMAP,nation,CUSTOMER" & _
'      " WHERE ST01(+)=CP13 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
'      " AND NA01(+)=A3 AND CU01(+)=SUBSTR(A4,1,8) AND CU02(+)=SUBSTR(A4,9)"
'
'   intI = 1
'   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With adoRst
'      .MoveFirst
'      Do While Not .EOF
'         '考慮會有拆收據情形,收款金額改不帶出
'         stSubject = .Fields("本所案號") & "之" & .Fields("案件性質") & "已收款，請送件！"
'         stContent = "本所案號：" & .Fields("本所案號") & vbCrLf & _
'            "案件名稱：" & .Fields("案件名稱") & vbCrLf & _
'            "申請國家：" & .Fields("NA03") & vbCrLf & _
'            "申請人　：" & .Fields("CU04") & vbCrLf & _
'            "智權人員：" & .Fields("ST02") & vbCrLf & _
'            "案件性質：" & .Fields("案件性質") & vbCrLf & _
'            "收文日　：" & .Fields("收文日") & vbCrLf & _
'            ""
'
'         PUB_SendMail strUserNum, "" & .Fields("CP14"), "", stSubject, stContent, , , , , , , , , , , False
'
'         .MoveNext
'      Loop
'      End With
'   End If
'   'end 2015/11/27
'
'ErrHnd:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description
'   End If
'   Set adoRst = Nothing
End Sub

'Added by Morgan 2021/4/9
'設定案源變數
'Modified by Morgan 2022/12/13 +pCourtFee出庭費
Private Sub SetLOSVar(pA1k01 As String, ByRef pLOS02 As String, ByRef pLCaseNo As String, ByRef pB2NeeCourt As Boolean, Optional ByRef pCourtFee As Long)
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   pLOS02 = "" '案源類別
   pLCaseNo = "" '法務案號
   pCourtFee = 0
   stSQL = "select los02,c2.cp01||c2.cp02||decode(c2.cp04,'00',decode(c2.cp03,'0','','-'||c2.cp03),'-'||c2.cp04) LCase,los15,c2.cp01" & _
      " from caseprogress c1,lawofficesource,caseprogress c2" & _
      " where c1.cp60='" & pA1k01 & "' and c1.cp162 is not null and los15(+)=c1.cp162" & _
      " and c2.cp09(+)=los06"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      pLOS02 = rsQuery.Fields("los02")
      pLCaseNo = rsQuery.Fields("LCase")
      'Added by Morgan 2021/4/29
      If pLOS02 = "B2" Then
         pB2NeeCourt = PUB_IsB2NeedCourt(rsQuery.Fields("los15"))
         If pB2NeeCourt Then
            stSQL = "select cl02,cl03 from (select decode(a.cp01||a.cp10,'L78',b.cp09,'FCL997',b.cp09,a.cp09) RNo" & _
               " from caseprogress a,caseprogress b  where a.cp162='" & rsQuery.Fields("los15") & "'" & _
               " and a.cp01='" & rsQuery.Fields("cp01") & "' and b.cp09(+)=a.cp43) X,caseprogress,caselawer a" & _
               " where cp09(+)=RNo and cl01(+)=cp09 and exists(select * from caselawer b where cl01=cp09 and cl02=cp14)"
            intQ = 1
            Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
            If intQ = 1 Then
               Do While Not rsQuery.EOF
                  pCourtFee = pCourtFee + rsQuery.Fields("cl03")
                  rsQuery.MoveNext
               Loop
            Else
               pCourtFee = 5000
            End If
         End If
      End If
   End If
End Sub

