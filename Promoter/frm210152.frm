VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210152 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月點數結算及查詢"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdOk 
      Caption         =   "結束(&X)"
      Height          =   345
      Left            =   8100
      TabIndex        =   87
      Top             =   30
      Width           =   800
   End
   Begin TabDlg.SSTab tabSP 
      Height          =   5520
      Left            =   0
      TabIndex        =   30
      Top             =   195
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   9737
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   10
      TabHeight       =   420
      TabCaption(0)   =   "個人資料"
      TabPicture(0)   =   "frm210152.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label21(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label21(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblAccept"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboEmp"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl21(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl21(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check1(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cboSalesArea(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame1(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdSearch(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdSave(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Check1(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame1(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Check1(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "全區資料"
      TabPicture(1)   =   "frm210152.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11(1)"
      Tab(1).Control(1)=   "Label11(2)"
      Tab(1).Control(2)=   "Label11(0)"
      Tab(1).Control(3)=   "Label10(1)"
      Tab(1).Control(4)=   "Label10(0)"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "grdDataList"
      Tab(1).Control(7)=   "cmdSave(1)"
      Tab(1).Control(8)=   "cmdSearch(1)"
      Tab(1).Control(9)=   "Text11(0)"
      Tab(1).Control(10)=   "Text11(1)"
      Tab(1).Control(11)=   "cboSalesArea(1)"
      Tab(1).ControlCount=   12
      Begin VB.CheckBox Check1 
         Height          =   280
         Index           =   2
         Left            =   4510
         TabIndex        =   20
         Top             =   1320
         Width           =   250
      End
      Begin VB.Frame Frame1 
         Height          =   4100
         Index           =   1
         Left            =   4460
         TabIndex        =   63
         Top             =   1320
         Width           =   4350
         Begin VB.TextBox Text5 
            Height          =   270
            Index           =   1
            Left            =   2320
            TabIndex        =   91
            Text            =   "Text5(1)"
            Top             =   0
            Width           =   1000
         End
         Begin VB.OptionButton Option2 
            Height          =   300
            Index           =   2
            Left            =   2100
            TabIndex        =   90
            Top             =   0
            Width           =   200
         End
         Begin VB.TextBox Text4 
            Height          =   270
            Index           =   1
            Left            =   1480
            TabIndex        =   22
            Text            =   "Text4(1)"
            Top             =   735
            Width           =   1000
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Ｈ結餘動用："
            Height          =   300
            Index           =   0
            Left            =   80
            TabIndex        =   21
            Top             =   735
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Ｉ期末結餘："
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   23
            Top             =   1335
            Width           =   1380
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   270
            Index           =   24
            Left            =   1480
            TabIndex        =   24
            Top             =   1335
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   270
            Index           =   36
            Left            =   1480
            TabIndex        =   27
            Top             =   2400
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   270
            Index           =   32
            Left            =   1480
            TabIndex        =   26
            Top             =   2160
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   270
            Index           =   40
            Left            =   1480
            TabIndex        =   28
            Top             =   2700
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   270
            Index           =   28
            Left            =   1480
            TabIndex        =   25
            Top             =   1920
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl21 
            Height          =   255
            Index           =   9
            Left            =   2550
            TabIndex        =   101
            Top             =   2700
            Width           =   1755
            VariousPropertyBits=   27
            Caption         =   "lbl21(9)"
            Size            =   "3096;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl21 
            Height          =   255
            Index           =   7
            Left            =   2550
            TabIndex        =   99
            Top             =   2430
            Width           =   1755
            VariousPropertyBits=   27
            Caption         =   "lbl21(7)"
            Size            =   "3096;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl21 
            Height          =   255
            Index           =   6
            Left            =   2550
            TabIndex        =   98
            Top             =   2190
            Width           =   1755
            VariousPropertyBits=   27
            Caption         =   "lbl21(6)"
            Size            =   "3096;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl21 
            Height          =   255
            Index           =   5
            Left            =   2550
            TabIndex        =   97
            Top             =   1920
            Width           =   1755
            VariousPropertyBits=   27
            Caption         =   "lbl21(5)"
            Size            =   "3096;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   525
            Index           =   41
            Left            =   1480
            TabIndex        =   29
            Top             =   2970
            Width           =   2805
            VariousPropertyBits=   -1466941413
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "4948;926"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label41 
            Caption         =   "Label41(12)"
            Height          =   270
            Index           =   12
            Left            =   2520
            TabIndex        =   82
            Top             =   1335
            Width           =   1000
         End
         Begin VB.Label Label41 
            Caption         =   "Label41(11)"
            Height          =   270
            Index           =   11
            Left            =   2520
            TabIndex        =   80
            Top             =   735
            Width           =   1000
         End
         Begin VB.Label Label41 
            Caption         =   "Label41(1)"
            Height          =   270
            Index           =   1
            Left            =   3390
            TabIndex        =   77
            Top             =   45
            Width           =   780
         End
         Begin VB.Label Lbl10 
            AutoSize        =   -1  'True
            Caption         =   "結餘：報出結餘點數＝"
            Height          =   180
            Index           =   1
            Left            =   315
            TabIndex        =   35
            Top             =   45
            Width           =   1800
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "                        －Ｉ期末結餘＋Ｊ轉撥增減"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   10
            Left            =   145
            TabIndex        =   76
            Top             =   3840
            Width           =   3060
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "　轉撥備註："
            Height          =   180
            Index           =   9
            Left            =   345
            TabIndex        =   75
            Top             =   3000
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "報出結餘點數＝Ｆ期初結餘＋Ｇ當月結餘"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   20
            Left            =   240
            TabIndex        =   74
            Top             =   3600
            Width           =   2970
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ｆ期初結餘："
            Height          =   180
            Index           =   27
            Left            =   360
            TabIndex        =   73
            Top             =   285
            Width           =   1080
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Label41(0)"
            Height          =   180
            Index           =   0
            Left            =   1480
            TabIndex        =   72
            Top             =   285
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ｇ當月結餘："
            Height          =   180
            Index           =   25
            Left            =   360
            TabIndex        =   71
            Top             =   525
            Width           =   1080
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Label41(6)"
            Height          =   180
            Index           =   6
            Left            =   1485
            TabIndex        =   70
            Top             =   525
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "＝Ｆ期初結餘＋Ｇ當月結餘－Ｉ期末結餘"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   345
            TabIndex        =   69
            Top             =   1080
            Width           =   2970
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "＝Ｆ期初結餘＋Ｇ當月結餘－Ｈ結餘動用"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   345
            TabIndex        =   68
            Top             =   1680
            Width           =   2970
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ｊ轉撥增減："
            Height          =   180
            Index           =   21
            Left            =   345
            TabIndex        =   67
            Top             =   2760
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "主管一調整："
            Height          =   180
            Index           =   19
            Left            =   345
            TabIndex        =   66
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "主管二調整："
            Height          =   180
            Index           =   18
            Left            =   345
            TabIndex        =   65
            Top             =   2160
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "總經理調整："
            Height          =   180
            Index           =   17
            Left            =   345
            TabIndex        =   64
            Top             =   2400
            Width           =   1080
         End
      End
      Begin VB.CheckBox Check1 
         Height          =   280
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   1320
         Width           =   250
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "存檔"
         Height          =   345
         Index           =   0
         Left            =   7920
         TabIndex        =   34
         Top             =   310
         Width           =   700
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查詢"
         Height          =   345
         Index           =   0
         Left            =   7080
         TabIndex        =   33
         Top             =   310
         Width           =   700
      End
      Begin VB.ComboBox cboSalesArea 
         Height          =   300
         Index           =   1
         Left            =   -74160
         TabIndex        =   0
         Text            =   "cboSalesArea(1)"
         Top             =   900
         Width           =   2000
      End
      Begin VB.TextBox Text11 
         Height          =   264
         Index           =   1
         Left            =   -66840
         MaxLength       =   2
         TabIndex        =   2
         Top             =   900
         Width           =   350
      End
      Begin VB.TextBox Text11 
         Height          =   264
         Index           =   0
         Left            =   -67800
         MaxLength       =   3
         TabIndex        =   1
         Top             =   900
         Width           =   600
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查詢"
         Height          =   345
         Index           =   1
         Left            =   -68160
         TabIndex        =   10
         Top             =   480
         Width           =   700
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "主管確認"
         Height          =   345
         Index           =   1
         Left            =   -67320
         TabIndex        =   11
         Top             =   480
         Width           =   1000
      End
      Begin VB.Frame Frame1 
         Height          =   4100
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   4350
         Begin VB.TextBox Text5 
            Height          =   270
            Index           =   0
            Left            =   2320
            TabIndex        =   89
            Text            =   "Text5(0)"
            Top             =   0
            Width           =   1000
         End
         Begin VB.OptionButton Option1 
            Height          =   300
            Index           =   2
            Left            =   2095
            TabIndex        =   88
            Top             =   0
            Width           =   200
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ｄ期末實績："
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   14
            Top             =   1320
            Width           =   1380
         End
         Begin VB.TextBox Text4 
            Height          =   270
            Index           =   0
            Left            =   1480
            TabIndex        =   13
            Text            =   "Text4(0)"
            Top             =   735
            Width           =   1000
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ｃ實績動用："
            Height          =   300
            Index           =   0
            Left            =   80
            TabIndex        =   12
            Top             =   720
            Width           =   1380
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   264
            Index           =   15
            Left            =   1480
            TabIndex        =   18
            Top             =   2400
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   264
            Index           =   11
            Left            =   1480
            TabIndex        =   17
            Top             =   2160
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   264
            Index           =   7
            Left            =   1480
            TabIndex        =   16
            Top             =   1920
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   264
            Index           =   19
            Left            =   1480
            TabIndex        =   19
            Top             =   2700
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   270
            Index           =   3
            Left            =   1480
            TabIndex        =   15
            Top             =   1335
            Width           =   1000
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TxtSP 
            Height          =   525
            Index           =   20
            Left            =   1470
            TabIndex        =   102
            Top             =   2970
            Width           =   2805
            VariousPropertyBits=   -1466941413
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "4948;926"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl21 
            Height          =   255
            Index           =   8
            Left            =   2550
            TabIndex        =   100
            Top             =   2700
            Width           =   1755
            VariousPropertyBits=   27
            Caption         =   "lbl21(8)"
            Size            =   "3096;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl21 
            Height          =   255
            Index           =   4
            Left            =   2550
            TabIndex        =   96
            Top             =   2430
            Width           =   1755
            VariousPropertyBits=   27
            Caption         =   "lbl21(4)"
            Size            =   "3096;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl21 
            Height          =   255
            Index           =   3
            Left            =   2550
            TabIndex        =   95
            Top             =   2190
            Width           =   1755
            VariousPropertyBits=   27
            Caption         =   "lbl21(3)"
            Size            =   "3096;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl21 
            Height          =   255
            Index           =   2
            Left            =   2550
            TabIndex        =   94
            Top             =   1920
            Width           =   1755
            VariousPropertyBits=   27
            Caption         =   "lbl21(2)"
            Size            =   "3096;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "(負數表示期末增加)"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   2520
            TabIndex        =   81
            Top             =   1080
            Width           =   1560
         End
         Begin VB.Label Label31 
            Caption         =   "12345.123"
            Height          =   270
            Index           =   12
            Left            =   2520
            TabIndex        =   79
            Top             =   1335
            Width           =   1000
         End
         Begin VB.Label Label31 
            Caption         =   "12345.123"
            Height          =   270
            Index           =   11
            Left            =   2520
            TabIndex        =   78
            Top             =   735
            Width           =   1000
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "　轉撥備註："
            Height          =   180
            Index           =   8
            Left            =   330
            TabIndex        =   62
            Top             =   3000
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "　　　　　　－Ｄ期末實績＋Ｅ轉撥增減"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   240
            TabIndex        =   61
            Top             =   3840
            Width           =   2970
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "報出實績點數＝Ａ期初實績＋Ｂ當月點數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   7
            Left            =   240
            TabIndex        =   54
            Top             =   3600
            Width           =   2970
         End
         Begin VB.Label Label31 
            Caption         =   "12345.123"
            Height          =   270
            Index           =   1
            Left            =   3390
            TabIndex        =   53
            Top             =   45
            Width           =   900
         End
         Begin VB.Label Lbl10 
            AutoSize        =   -1  'True
            Caption         =   "實績：報出實績點數＝"
            Height          =   180
            Index           =   0
            Left            =   320
            TabIndex        =   83
            Top             =   45
            Width           =   1800
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "總經理調整："
            Height          =   180
            Index           =   5
            Left            =   360
            TabIndex        =   52
            Top             =   2400
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "主管二調整："
            Height          =   180
            Index           =   4
            Left            =   360
            TabIndex        =   51
            Top             =   2160
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "主管一調整："
            Height          =   180
            Index           =   3
            Left            =   360
            TabIndex        =   50
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ｅ轉撥增減："
            Height          =   180
            Index           =   6
            Left            =   330
            TabIndex        =   49
            Top             =   2700
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "＝Ａ期初實績－Ｃ實績動用"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   360
            TabIndex        =   48
            Top             =   1680
            Width           =   1980
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "＝Ａ期初實績－Ｄ期末實績"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   345
            TabIndex        =   47
            Top             =   1080
            Width           =   1980
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "12345.123"
            Height          =   180
            Index           =   6
            Left            =   1480
            TabIndex        =   46
            Top             =   525
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ｂ當月點數："
            Height          =   180
            Index           =   2
            Left            =   360
            TabIndex        =   45
            Top             =   525
            Width           =   1080
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Label31"
            Height          =   180
            Index           =   0
            Left            =   1480
            TabIndex        =   44
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ａ期初實績："
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   43
            Top             =   285
            Width           =   1080
         End
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1330
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   6
         Top             =   720
         Width           =   350
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   6960
         MaxLength       =   3
         TabIndex        =   5
         Top             =   720
         Width           =   600
      End
      Begin VB.ComboBox cboSalesArea 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Text            =   "cboSalesArea(0)"
         Top             =   360
         Width           =   2500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "報出點數："
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   4285
         Left            =   -74880
         TabIndex        =   55
         Top             =   1220
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   7557
         _Version        =   393216
         BackColor       =   -2147483624
         Rows            =   16
         Cols            =   8
         BackColorFixed  =   -2147483624
         BackColorBkg    =   -2147483624
         BackColorUnpopulated=   -2147483624
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         MergeCells      =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.Label lbl21 
         Height          =   255
         Index           =   1
         Left            =   5220
         TabIndex        =   93
         Top             =   1110
         Width           =   3000
         VariousPropertyBits=   27
         Caption         =   "lbl21(1)"
         Size            =   "5292;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   165
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl21 
         Height          =   255
         Index           =   0
         Left            =   5220
         TabIndex        =   92
         Top             =   930
         Width           =   3000
         VariousPropertyBits=   27
         Caption         =   "lbl21(0)"
         Size            =   "5292;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   165
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboEmp 
         Height          =   300
         Left            =   4410
         TabIndex        =   4
         Top             =   360
         Width           =   2295
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4048;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblAccept 
         Caption         =   "主管已確認"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   4485
         TabIndex        =   86
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label9 
         Caption         =   "註:藍色尚未輸入實績與結餘資料"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   -72050
         TabIndex        =   85
         Top             =   960
         Width           =   2805
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "PS：１．期末實績及期末結餘欄，若主管有調整以主管調整數字顯示"
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   60
         Top             =   360
         Width           =   5400
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "        ２．主管確認後不可再修改數字，若需修改請財務處開放"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   59
         Top             =   600
         Width           =   4860
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "業務區："
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   58
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "年           月"
         Height          =   180
         Index           =   2
         Left            =   -67080
         TabIndex        =   57
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "業績年月：民國"
         Height          =   180
         Index           =   1
         Left            =   -69120
         TabIndex        =   56
         Top             =   900
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "達成率："
         Height          =   180
         Index           =   2
         Left            =   2520
         TabIndex        =   41
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Label21(1)％"
         Height          =   180
         Index           =   1
         Left            =   3240
         TabIndex        =   40
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "目　標："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Label21(0)"
         Height          =   180
         Index           =   0
         Left            =   840
         TabIndex        =   38
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年           月"
         Height          =   180
         Index           =   3
         Left            =   7680
         TabIndex        =   37
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "業績年月：民國"
         Height          =   180
         Index           =   2
         Left            =   5640
         TabIndex        =   36
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   84
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "＝報出實績點數＋報出結餘點數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   2370
         TabIndex        =   32
         Top             =   1020
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "業務區："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   390
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm210152"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、cboEmp、TxtSP(index)、lbl21(index)並且將主管名稱改為lbl21(參考SetCUID)
'Memo 2021/11/11 GetSPDept(stUser As String) 函數不使用已刪,若需再使用可改用GetST15
'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線
'原「當月　　實績」修改為「當月實績」。
'原「當月　　結餘」修改為「當月結餘」。
'end 2021/07/27
'Create by Amy 2015/12/08
'Memo 2016/02/19 文雄詢問過 目標>=當月實績,當月實績 要全報(201602上線時只有智權部需操作)
'Memo 2023/06/08 目前開放部分非智權部,也需操作報點數,[非智權部] 目標>=當月實績,當月實績要全報 也加此規定-秀玲與婉莘討論後,經總經理同意
'                                     且增加：當月實績>=目標時,報出點數必須>=目標 (需每月報點數者都判斷)-杜燕文協理 ex:11205月 87011林青祺 目標760,當月796.9,報出點數609.9
Option Explicit

Const intField = 51 'SalesPoint 欄位數
' 宣告欄位內容結構
Private Type FIELDITEM
    fiName As String
    fiOldData As String
    fiType As Integer
End Type
Dim m_FieldList(1 To intField) As FIELDITEM
Dim strRowN()
Dim dblTot(1 To 15) As Double
Dim SetColor '每人的Window外觀設定值不同,故先記錄原始設定顏色
Dim i As Integer
Dim bolOrgSet(1 To 2) As Boolean
Public strAreaManNo As String, strGlManNo As String '代理區主管編號/總經理員編
Dim strNowUser As String '目前User(可能為職代)
Dim intLimit As Integer  '輸入欄位權限設定:-1-ReadlyOnly主管/0-ReadyOlny個人/10-個人輸自己/2-主管一/20-主管一輸自己/3-主管二/4-總經理/40-總經理輸自己
Dim bolIsFirst As Boolean '第一次進入,未查詢
Dim bolGlMan As Boolean '是否為總經理權限
Dim bolIsAccept As Boolean, bolSave As Boolean '主管是否已確認/是否可存檔
Dim strArea1 As String, strArea2 As String  '特殊人員查詢區別
Dim strAreaList As String, strSt52List As String '區主管管理的區List/帶人主管帶的人員list
Dim stST05 As String 'Modify by Amy 2019/10/16 stST15 改Public
Dim strA0b01 As String, strA0b05 As String '會計過帳日/業績輸入關閉年月
Dim bolNoChkMod As Boolean
'Add by Amy 2016/02/16
Dim bolNowChk As Boolean 'TextBox LostFocus 避免無窮迴圈用
Dim bolNoMsg As Boolean '勾選報出是否show訊息
Dim bolLeave As Boolean 'Add by Amy 2016/04/07 是否離職
'Add by Amy 2016/05/09 SalesPoint是否無資料
Dim bolEmptyF1 As Boolean, bolEmptyF2 As Boolean '實績/結餘
Dim bolEmptyF3 As Boolean, bolemptyF4 As Boolean '實績轉撥/結餘轉撥
'Dim strUpdSB As String 'Mark by Amy 2022/06/09 不使用 'Add by Amy 2017/09/26 SP36有修改需更新SalesBalance語法
Dim strSP48 As String 'Add by Amy 2017/09/26 查詢人員部門(可能換部門)
Dim strMaxSP01 As String 'Add by Amy 2017/12/01 '記錄目前SalesPoint最大年月
'Modify by Amy 2019/10/16
Public stST15 As String '收文部門-登入操作人員(可能為區主管以區主管身份登入的職代或多重身份)
'Const stSpecEmpNo = "F4102;F4103;W1001;W2001;20091" ' 特殊員編 'Mark by Amy 2021/01/18 改為共用 智權點數實績與結餘特殊員編 +F4104~07 '2020/06/16 +20091
Public bolAreaMan As Boolean 'Modify by Amy 2021/07/16 是否為區主管 原:Dim
Dim strToSpecNo As String 'Add by Amy 2021/07/16 特殊發mail員編
'Add by Amy 2021/11/11
Public strInputEmp As String '非智權部操作之特殊員編
Dim strEmpList As String
Public IsAgentLimit As String 'Add by Amy 2023/02/02 是職代

Private Sub CboEmp_Click()
    Dim strDeptName As String, strDept As String
    
    If bolIsFirst = True Then Exit Sub 'Add by Amy 2021/01/04
    If CboEmp = MsgText(601) Then Exit Sub
    'Mark by Amy 2021/11/11 於From_Load將登入者可使用之部門設好,故不需重抓(避免未帶部門產生錯誤)
'    If cboSalesArea(0).Enabled = False Then
'        cboSalesArea(0).Clear
'        'Modify by Amy 2016/08/17 員編可能4碼(S142)或5碼
'        If Left(cboEmp, InStr(cboEmp, " ") - 1) <> strNowUser Then
'            'Modify by Amy 2017/09/26 +GetST15參數
'            strDept = GetST15(Left(cboEmp, InStr(cboEmp, " ") - 1), strDeptName, Val(Text1(0) & Text1(1)) + 191100)
'            cboSalesArea(0).AddItem strDept & " " & strDeptName
'            cboSalesArea(0) = strDept & " " & strDeptName
'        Else
'            'Modify by Amy 2017/09/26 +strSP48
'            Call SetcboSalesArea(strSP48, False)
'        End If
'    End If
    
    '選完自動查詢
    If Trim(CboEmp) <> MsgText(601) And Trim(Text1(0)) <> MsgText(601) And Trim(Text1(1)) <> MsgText(601) Then
        'Modify by Amy 2019/10/16 +bolNoChkMod 避免無窮迴圈
        bolNoChkMod = True
        Call cmdSearch_Click(0)
        bolNoChkMod = False
    End If
End Sub

'Add by Amy 2016/04 總經理權限可以輸入員編帶區
Private Sub CboEmp_GotFocus()
    If bolGlMan = True Then
        CloseIme
        CboEmp.SelStart = 0
        CboEmp.SelLength = Len(CboEmp)
    End If
End Sub

'Modified by Lydia 2022/01/03 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub CboEmp_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If bolGlMan = False Then KeyAscii = 0: Exit Sub
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboEmp_LostFocus()
    Dim strTemp(2) As String
    Dim strDept As String
    
    '只有總經理權限可輸入
    If bolGlMan = False Then Exit Sub
    'Modify by Amy 2016/08/17 員編可能4碼(S142)或5碼
    'Modify by Amy 2019/10/16 員編可能3碼(S29)
    If Trim(CboEmp) = MsgText(601) Or Len(Trim(CboEmp)) < 3 Then Exit Sub
    If Trim(CboEmp) = "M0101" Then Exit Sub 'Add by Amy 2023/01/04 ACS分潤程式未完成前先住
    
    strTemp(0) = Left(Trim(CboEmp), IIf(InStr(Trim(CboEmp), " ") > 0, InStr(Trim(CboEmp), " ") - 1, Len(Trim(CboEmp))))
    'end 2016/08/17
    strTemp(1) = GetPrjSalesNM(strTemp(0))
    If strTemp(1) = MsgText(601) Then
        MsgBox "員工編號輸入錯誤,請確認！", , MsgText(5)
        CboEmp.SetFocus
        CboEmp_GotFocus
    Else
        'Modify by Amy 2016/04/07 +輸入後下拉裡也出現同部門其他人員
        cboSalesArea(0) = ""
        'Modify by Amy 2017/09/26 中一高國碩/陳頌恩轉中二,修正原部門抓st15轉後看不到舊資料
        'strTemp(2) = PUB_GetStaffST15(strTemp(0), 1)
        'Modfiy by Amy 2021/11/23 原程式改至函數
        Call ChkSetDept(strTemp(0))
        'Modfiy by Amy 2021/11/11 重改SetEmp 函數
'        Call SetEmp(3)
'        cboEmp = strTemp(0) & " " & strTemp(1)
        Call SetEmp(1)
        '若設自動查詢當下拉選完後(cboEmp_click執行),游標離開後觸發cboEmp_LostFocus又會查一次
        '操作:改完資料未存檔直接下拉,此處若設會尋問兩次是否存檔
        'end 2016/04/07
    End If
End Sub
'end 2016/04/07

Private Sub cboSalesArea_Click(Index As Integer)
    If cboSalesArea(Index) = MsgText(601) Then Exit Sub
    '個人
    If Index = 0 Then
        'Moidfy by Amy 2019/10/16 +if 下拉選部門智權人員清空
        If cboSalesArea(Index).Tag <> Trim(cboSalesArea(Index)) Then CboEmp = ""
        
        'Modify by Amy 2020/06/16 操作者為柄佑 82026,智權人員不預設自己
        'Modfiy by Amy 2021/11/11 重改SetEmp 函數 原:SetEmp(IIf(strNowUser = "82026", 2, 1))
        Call SetEmp(1)
    '全區
    Else
        'Add by Amy 2016/03/25 +選完自動查
        If cboSalesArea(1) <> MsgText(601) And Trim(Text11(0)) <> MsgText(601) And Trim(Text11(1)) <> MsgText(601) Then
            cmdSearch_Click (1)
        End If
    End If
End Sub

Private Sub Check1_Click(Index As Integer)
    '勾報出點數
    If Check1(0).Value = 1 Then
        Select Case Index
            Case 0
                Call SetInputText(0)
                If intLimit > 0 Then Call SetTextCSS(Text2, 2)
                Check1(1).Value = 0
                Check1(2).Value = 0
                If intLimit > 0 Then
                    Check1(1).Enabled = True
                    Check1(2).Enabled = True
                End If
                Call SetChoose(False, 2)
                Call SetChoose(False, 3)
            Case 1, 2
                If Check1(1).Enabled = False And Check1(2).Enabled = False Then
                    '當勾選報出點數設check1(1).value時會觸發此事件
                    '為不run下方程式,故判斷兩者Enabled = False時
                Else
                    If Check1(Index).Value = vbChecked Then
                        If Index = 1 Then
                            Check1(2).Value = 0
                        Else
                            Check1(1).Value = 0
                        End If
                    Else
                        If Index = 1 Then
                            Check1(2).Value = 1
                        Else
                            Check1(1).Value = 1
                        End If
                    End If
                    Frame1(0).Enabled = IIf(Check1(1).Value = vbChecked, 1, 0)
                    Frame1(1).Enabled = IIf(Check1(2).Value = vbChecked, 1, 0)
                    Call SetChoose(IIf(Check1(1).Value = vbChecked, 1, 0), 2)
                    Call SetChoose(IIf(Check1(2).Value = vbChecked, 1, 0), 3)
                    Call SetInputText(98, Index)
                End If
        End Select
    '不勾報出點數
    Else
        If Index = 0 Then
            Call SetInputText(0)
            Call SetTextCSS(Text2, 3)
            Check1(1).Value = 1
            Check1(2).Value = 1
            If intLimit > 0 Then
                Check1(1).Enabled = False
                Check1(2).Enabled = False
                Frame1(0).Enabled = True
                Frame1(1).Enabled = True
            End If
            Call SetChoose(True, 2)
            Call SetChoose(True, 3)
        Else
            Call SetInputText(1)
            Call SetInputText(2)
        End If
    End If
    
    'Modify by Amy 2016/02/16 第一次讀取不show訊息
    If Index = 0 And Check1(0).Value = 1 And bolNoMsg = False Then
        MsgBox "請選擇輸入「實績部分」或「結餘部分」", , MsgText(21)
    End If
    
End Sub

Private Sub cmdok_Click()
    bolGlMan = False
    Unload Me
End Sub

Private Sub CmdSave_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    If FormCheck1(Index) = True Then
        If FormSave(Index) = True Then
            If Index = 0 Then
                bolNoChkMod = True
                Call cmdSearch_Click(0)
                bolNoChkMod = False
            Else
                '主管確認成功鎖住按鈕
                cmdSave(Index).Enabled = False
                '個人資料畫面欄位鎖住
                If cmdSave(0).Enabled = True And (Left(cboSalesArea(0), 3) = stST15 Or InStr(Left(cboSalesArea(0), 3), strAreaList) > 0) Then
                    If Left(CboEmp, 5) = strNowUser Then
                        intLimit = 0
                    Else
                        intLimit = -1
                    End If
                    If Check1(0).Value = 1 Then
                        Call SetChoose(False)
                        Call SetTextCSS(Text2, 3)
                        Call SetInputText(98)
                    Else
                        Call SetInputText(1)
                        Call SetInputText(2)
                    End If
                End If
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    Dim RsQ As New ADODB.Recordset
    Dim bolLocked As Boolean '記錄欄位目前Locked值
    Dim bolTarget As Boolean '是否有目標
    
    Screen.MousePointer = vbHourglass
    If Index = 0 Then
        If bolNoChkMod = False Then
            If intLimit > 0 And cmdSave(Index).Enabled = True Then
                '先判斷是否需存檔
                If bolIsFirst = False And ChkModify = True And CboEmp.Tag <> MsgText(601) Then
                    If MsgBox(GetStaffName(CboEmp.Tag, True) & " 資料有修改要存檔?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                        If FormCheck1(0) = False Then
                            CboEmp = CboEmp.Tag & GetStaffName(CboEmp.Tag, True)
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        Else
                            If FormSave(Index) = False Then
                                MsgBox "存檔失敗,請洽電腦中心！", , MsgText(5)
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
        CboEmp.Tag = Left(CboEmp, 5)
        bolLeave = ChkStaffST04(Left(CboEmp, 5), False)
        'Modify by Amy 2019/10/16
        Text1(0).Tag = Val(Text1(0)) & Text1(1) '從Text1_Validate搬過來(A2004登入輸A4023->改日期 無窮回圈)
        cboSalesArea(0).Tag = cboSalesArea(0)
        'end 2019/10/16
    End If
    
    If Index = 0 Then Call FormClear(1): LblAccept.Visible = False
    If FormCheck1(Index) = True Then
        bolIsAccept = IsAccept(Index, Left(cboSalesArea(Index), 3))
        If Index = 0 Then
            If bolIsAccept = True Then LblAccept.Visible = True
            '顯示目標/期初實績、結餘/當月實績、結餘(抓傳票資料)
            bolTarget = doQuery1(1)
            
            '抓智權點數與結餘檔資料
            Set RsQ = ReadSalesPoint(Val(Text1(0) & Text1(1)) + 191100, Left(CboEmp, 5))
            UpdateFieldOldData RsQ
            If bolTarget = False And RsQ.RecordCount = 0 Then
                'Mark by Amy 2017/09/26 拿掉總經理權限秀不同訊息(輸入員工若無資料需至隱藏版人員新增)
                'Add by Amy 2017/02/03 +總經理權限秀不同訊息
'                If bolGlMan = True Then
'                    If MsgBox("無符合資料！請問是否輸入此筆資料？", vbYesNo + vbDefaultButton2) = vbNo Then
'                        Screen.MousePointer = vbDefault
'                        Exit Sub
'                    End If
'                Else
                    MsgBox "無符合資料！", vbInformation
'                End If
            ElseIf RsQ.RecordCount > 0 Then
                RefreshRecord RsQ
            End If
            
            bolIsFirst = False
            cmdSave(0).Enabled = False
            'Modify by Amy 2016/02/16 先預設不勾選「報出點數」
            Check1(0).Value = 0: Check1(0).Enabled = True
            Option1(1).Value = True: Option2(1).Value = True
            Call SetTextCSS(Text2, 3)
            Call SetChoose(True, 2)
            Call SetChoose(True, 3)
            '設定輸入欄位
            Call SetInputText(99)
            'Memo 2023/08/07 原程式往下搬:Modify by Amy 2019/10/16 員工編號為區目標不輸資料 ex:S29/S212
            
            'Add by Amy 2016/02/23
            If m_FieldList(49).fiOldData = Empty Then
                bolLocked = TxtSP(3).Locked
                TxtSP(3).Locked = False
                Call TxtSP_LostFocus(3)
                If bolLocked <> TxtSP(3).Locked Then TxtSP(3).Locked = bolLocked
                bolLocked = TxtSP(24).Locked
                TxtSP(3).Locked = False
                Call TxtSP_LostFocus(24)
                If bolLocked <> TxtSP(24).Locked Then TxtSP(24).Locked = bolLocked
            End If
            'end 2016/02/23
            Call RunSum(, , True)
            '畫面記錄非選期末值需更新主管欄位顯示值
            If Not (Val(m_FieldList(49).fiOldData) = 0 And Val(m_FieldList(50).fiOldData) = 1 And Val(m_FieldList(51).fiOldData) = 1) Then
                Call SetViewVal
            End If
            '目標大於等於當月點數
            If Val(Label21(0)) >= Val(Label31(6)) And Text2.Locked = False Then Call Text2_GotFocus: Text2_LostFocus
            'end 2016/02/16
            'Add by Amy 2016/05/09 +判斷避免一直彈存檔訊息
            Call SetChkModifyField
            
            'Modify by Amy 2023/08/07 從上面搬下來,加W3001可能沒目標,且沒收款,故判斷沒目標且報出點數=0不用存檔-秀玲
            'Modify by Amy 2019/10/16 +Len(Mid(CboEmp, 1, Val(InStr(CboEmp, " ")) - 1)) >= 5 員工編號為區目標不輸資料 ex:S29/S212
            strExc(0) = CboEmp
            If InStr(strExc(0), " ") > 0 Then
                strExc(0) = Mid(CboEmp, 1, Val(InStr(CboEmp, " ")) - 1)
            End If
            If bolSave = True And (intLimit = 4 Or bolIsAccept = False) And Len(strExc(0)) >= 5 Then
                'Mark by Amy 2023/09/04 拿掉2023/08/07 目標且報出點數=0 條件,11208月 10052目標且報出點數=0(實績與結餘只有期初資料)無法按存檔鈕,且無法按主管確認
'                If Label21(0) = MsgText(601) And Val(Text2) = 0 Then
'                Else
                  cmdSave(0).Enabled = True
'               End If
            End If
            'end 2023/08/07
        Else
            Call SetDataListWidth
            If doQuery1(2) = False Then
                MsgBox "無符合資料！", vbInformation
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

'Add by Amy 2016/05/09 設定彈存檔訊息檢查欄位(SalesPoint對應欄位無資料)
Private Sub SetChkModifyField()
    bolEmptyF1 = False: bolEmptyF2 = False: bolEmptyF3 = False: bolemptyF4 = False
    '個人輸自己
    If intLimit = 10 Or intLimit = 20 Or intLimit = 40 Then
        '記錄預設資料
        If m_FieldList(3).fiOldData = Empty Then bolEmptyF1 = True: TxtSP(3).Tag = TxtSP(3)
        If m_FieldList(24).fiOldData = Empty Then bolEmptyF2 = True: TxtSP(24).Tag = TxtSP(24)
        If intLimit >= 20 Then
            If m_FieldList(19).fiOldData = Empty Then bolEmptyF3 = True
            If m_FieldList(40).fiOldData = Empty Then bolemptyF4 = True
        End If
    '主管操作他人
    Else
        If intLimit = 2 Then
            If m_FieldList(7).fiOldData = Empty Then bolEmptyF1 = True
            If m_FieldList(28).fiOldData = Empty Then bolEmptyF2 = True
        End If
        If intLimit = 3 Then
            If m_FieldList(11).fiOldData = Empty Then bolEmptyF1 = True
            If m_FieldList(32).fiOldData = Empty Then bolEmptyF2 = True
        End If
        If intLimit = 4 Then
            If m_FieldList(15).fiOldData = Empty Then bolEmptyF1 = True
            If m_FieldList(36).fiOldData = Empty Then bolEmptyF2 = True
        End If
        If intLimit = 2 Or intLimit = 3 Or intLimit = 4 Then
            If m_FieldList(19).fiOldData = Empty Then bolEmptyF3 = True
            If m_FieldList(40).fiOldData = Empty Then bolemptyF4 = True
        End If
    End If
End Sub

Private Sub Form_Load()
    '*** Memo by Amy 2023/02/03 有加Public變數,要確認 Account-Frmacc0000 是否會錯 ***
    
    MoveFormToCenter Me
    InitialField
    SetColor = Me.BackColor
    Call SetDataListWidth
    LblAccept.Visible = False
    FormClear (0)
    Check1_Click (0)
    cmdSave(0).Enabled = False 'Add by Amy 2017/09/26 避免一開始無資料就按「存檔」發生錯誤,先隱藏
    tabSP.TabVisible(1) = False
    cmdSave(1).Enabled = False
    
    'Modify by Amy 2017/09/26 中一高國碩/陳頌恩轉中二,修正部門預帶
    'Modify by Amy 2017/12/01 原預設系統日前一個月,改預設SelesPoint最大業績年月
    'strExc(0) = (CompDate(1, -1, strSrvDate(2)) - 19110000) '系統日前一個月
    strMaxSP01 = GetMaxSP01()
    strExc(0) = strMaxSP01 & "01"
    'end 2017/12/01
    '業績年月預設
    Text1(0) = Val(strExc(0)) \ 10000
    Text1(1) = Left(Right(strExc(0), 4), 2)
    Text1(0).Tag = Val(strExc(0)) \ 10000 & Left(Right(strExc(0), 4), 2)
    Text11(0) = Text1(0)
    Text11(1) = Text1(1)
    Text11(0).Tag = Text11(0) & Text11(1)
    
    '只開放可看特殊部門,部門下拉才可選 因中一高國碩/陳頌恩轉中二,避免抓錯(先不考慮st52跨區問題-秀玲)
    cboSalesArea(0).Enabled = False
    cboSalesArea(1).Enabled = False
    
    strA0b01 = GetA0b01(strA0b05)
    strGlManNo = Pub_GetSpecMan("總經理員工編號") '總經理員編
    
    bolIsFirst = True
    
'*** 權限設定 (此有修改需看 mdiMain.Frm210152Limit是否要改)***
    'Modify by Amy 2021/07/06 從上面搬下來整理,因外商由江郁仁98020確認,但由洪琬姿80030輸F4106,葉易雲 78011輸F4107
    stST05 = PUB_GetST05(strUserNum)
    '電腦中心,財務,總經理,主任秘書(等級08)
    If InStr("00,01,08", stST05) > 0 Then
        strNowUser = strUserNum
        bolGlMan = True '開放總經理權限
    ElseIf strAreaManNo = MsgText(601) Then
        strNowUser = strUserNum
    Else
        '區主管職代 or 區主管
         strNowUser = strAreaManNo
    End If
    strSt52List = GetST52List(strNowUser) '帶人主管 帶的人(文雄北四區登入可看到總經理,文雄為總經理帶人主管-1081024秀玲:沒關係)
    strSP48 = stST15
    strAreaList = strSP48
    If stST15 = "F11" Or bolGlMan = True Then strToSpecNo = Set98020Ag 'Modify by Amy 2021/07/20 外商or總經理權限登入要判斷江郁仁是否請假,for 發mail判斷

    'Mark by Amy 2021/07/06 改由 mdiMain.Frm210152Limit設定
'    'Modify by Amy 2019/08/01 +if 開放陳鳳英(F11)輸F4103及其職代A0914
'    'Modify by Amy 2021/06/21 陳鳳英退休改江郁仁(L01)操作
'    If Left(stST15, 2) = "F1" Or strUserNum = "98020" Then
'        strSP48 = GetSPDept("F4103;F4106;F4107")
'        strAreaList = strSP48 'Add by Amy 2019/11/04 陳鳳英登入部門為F10,但操作F11
'        stST15 = "F11" 'Add by Amy 2021/06/21 江郁仁進入要設為F11
'    'Add by Amy 2021/04/29 開放林純真(P20)輸P2005(P21)
'    ElseIf Left(stST15, 2) = "P2" Then
'        strSP48 = GetSPDept("P2005")
'        strAreaList = strSP48
'    'Add by Amy 2019/10/16 開放W部門可輸區編號
'    ElseIf Left(stST15, 1) = "W" Then
'        strSP48 = stST15
'    'Modify by Amy 2019/11/21  柄佑P11部門,但只能看S21~29,不可看自己在P11的資料
'    ElseIf strNowUser <> "82026" Then
'        strSP48 = GetSPDept(strNowUser)
'    End If
'    'end 2019/08/01
'    'Modify by Amy 2019/11/21  +82026柄佑P11部門,但只能看S21~29,不可看自己在P11的資料
'    If strSP48 = MsgText(601) And strNowUser <> "82026" Then
'        strSP48 = stST15
'    End If
'    'Modify by Amy 2020/06/16 +if 開放柄佑 82026 輸20091(S29部門),預設帶S21
'    If strNowUser = "82026" Then
'        stST15 = GetSPDept("20091"): strSP48 = "S21"
'    'Modify by Amy 2019/10/16 +if 已設定部門不再抓,文雄可以客服組區主管輸資料
'    ElseIf stST15 = MsgText(601) Then
'        stST15 = PUB_GetStaffST15(strNowUser, 1)
'    Else
'        strAreaList = stST15
'    End If
'    'Modifyb by Amy 2021/07/06
'    If strAreaList = MsgText(601) Then strAreaList = GetDeptList(3, strNowUser, stST15) '區主管管理區別List
'    'end 2019/10/16
    'end 2021/07/06
    
    strArea1 = strSP48
    strArea2 = strSP48
    
    'Modify by Amy 2023/02/02 +if  區主管職代(A0914)只能看到設定的該區(A0908)
    If IsAgentLimit = True Then
        'ex:杜協理(74018)請假時,79053蘇嫄媛只可以代為輸入台南所,A1033丁浚評只可以輸高雄所
    Else
        '下拉選單權限
        Select Case strNowUser
            '杜副總(68006)可看S部門全部
            'Mark by Amy 2021/09/02 改以部門判斷
    '        'Modify by Amy 2021/01/07 改為簡協理
    '        Case "69005"
    '            strArea1 = "S00"
    '            strArea2 = "S99"
    '            cboSalesArea(0).Enabled = True
    '            cboSalesArea(1).Enabled = True
    '            tabSP.TabVisible(1) = True
    '        '簡協理可看北所全部
    '        Case "69005"
    '            strArea1 = "S10"
    '            strArea2 = "S19"
    '            cboSalesArea(0).Enabled = True
    '            cboSalesArea(1).Enabled = True
    '            tabSP.TabVisible(1) = True
    '        '蘇特助可以看自己區及分所
    '        Case "69010"
    '            strArea1 = "S21"
    '            strArea2 = "S49"
    '            cboSalesArea(0).Enabled = True
    '            cboSalesArea(1).Enabled = True
    '            tabSP.TabVisible(1) = True
            'end 2021/01/07
            'Add by Amy 2016/11/18 林柄佑經理(P11)可以看自已和中所全部的個人資料,不可看P11的全區資料
            'Memo by Amy 2019/11/21  林柄佑經理需輸20091個人資料,因程式不好判斷,將原可看自己的資料拿掉
            Case "82026"
                strArea1 = "S21"
                strArea2 = "S29"
                cboSalesArea(0).Enabled = True
                cboSalesArea(1).Enabled = True
                tabSP.TabVisible(1) = True
            Case Else
                'Modify by Amy 2021/09/02 從上面搬下來,原判斷簡協理員編,改判斷部門(Pub_StrUserSt03),且加可看區別至W10
                'Modify by Amy 2022/05/24 原判斷:Pub_StrUserSt03 = "S00",strArea1="S00"
                If InStr(Pub_GetSpecMan("全所智權部主管"), strNowUser) > 0 Then
                    strArea1 = "S"
                    strArea2 = "W10"
                    cboSalesArea(0).Enabled = True
                    cboSalesArea(1).Enabled = True
                    tabSP.TabVisible(1) = True
                Else
                    Select Case stST05
                        '電腦中心,財務,總經理,主任秘書(等級08)可看全部看全部
                        Case "00", "01", "08"
                            strArea1 = ""
                            strArea2 = ""
                            cboSalesArea(0).Enabled = True
                            cboSalesArea(1).Enabled = True
                            tabSP.TabVisible(1) = True
                        Case Else
                    End Select
                End If
        End Select
    End If
    'end 2017/09/26
    '區主管才能看到「全區資料」頁籤
    If bolAreaMan = True Then
        tabSP.TabVisible(1) = True
    End If
    
    Call SetcboSalesArea(strSP48) 'Modify by Amy 2017/09/26 原: stST15
    'Add by Amy 2021/11/11 整理智權人員下拉選單設定
    CboEmp.Enabled = False
    '部門下拉選單可使用,智權人員下拉需開啟
    If cboSalesArea(0).Enabled = True Then CboEmp.Enabled = True
    
    '總經理權限
    If bolGlMan = True Then
    '非智權部下拉智權人員設定
    ElseIf Left(stST15, 1) <> "S" Then
        If strInputEmp <> MsgText(601) Then
            strEmpList = strInputEmp
        ElseIf GetInputPointST14Data(2, stST15, , , strEmpList) = True Then
        End If
    '智權部 82026 輸 20091
    ElseIf strInputEmp <> MsgText(601) Then
        strEmpList = strInputEmp
    'Add by Amy 2025/03/14 讓 鈺華A5005 下拉選單有30015(雖為30015的帶人主管,但GetST52List 只抓st01>'6')
    '              1140306 鈺華通知 1140305 可於智權人員下拉輸30015 ,但突然不行(Amy測一直都不行),秀玲說鈺華登入時,固定顯示於其下拉選單中
    Else
       strSt52List = strSt52List & ",'30015'"
       If Left(strSt52List, 1) = "," Then strSt52List = Mid(strSt52List, 2)
    'end 2025/03/14
    End If
    Call SetEmp
    'end 2021/11/11
    'Mark by Amy 2021/11/11 整理SetEmp,故下方式式不需使用
'    'Add by Amy 2017/0926 高國碩查10609以前資料只能看自己的(10610才為中二區主管)
'    If SpecData = False Then
'        'Modify by Amy 2020/06/16 +若操作者為 柄佑82026 則智權人員不預設自己
'        Call SetEmp(IIf(strNowUser = "82026", 2, 0))
'    End If
    '*** End 權限設定 (此有修改需看 mdiMain.Frm210152Limit是否要改)***
    
    Call SetInputText(0, 99)
End Sub
'intChoose:0-全部/1-查詢用
Private Sub FormClear(intChoose As Integer)
    Dim oText As Object
    Dim oLabel As Object
    Dim ii As Integer
    
    If intChoose = 0 Then
        For Each oText In Text1
            oText.Text = Empty
        Next
    End If
    
    For Each oLabel In lbl21
        oLabel.Caption = Empty
    Next
    
    For Each oLabel In Label21
        oLabel.Caption = Empty
    Next
    Check1(0).Value = 0
    Text2 = Empty
    TxtSP(19) = Empty
    TxtSP(20) = Empty
    TxtSP(40) = Empty
    TxtSP(41) = Empty
   
    '實績/結餘部分
    For Each oText In TxtSP
        oText.Text = Empty
        oText.Tag = Empty 'Add by Amy 2016/05/09
    Next
    Text5(0).Text = Empty
    Text5(1).Text = Empty
    Text4(0).Text = Empty
    Text4(1).Text = Empty
    For Each oLabel In Label31
        oLabel.Caption = Empty
    Next
    Option1(1).Value = True
    
    For Each oLabel In Label41
        oLabel.Caption = Empty
    Next
    Option2(1).Value = True
   
End Sub

Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim iCol As Integer, iRow As Integer
   
   ReDim strRowN(0 To 15)
   strRowN = Array("智權人員", "目標", "達成點數", "達 成 率", "期初實績", _
                            "當月實績", "實績動用", "期末實績", "轉撥實績增減", "報出實績點數", _
                            "期初結餘", "當月結餘", "結餘動用", "期末結餘", "轉撥結餘增減", "報出結餘點數")
                            
   With grdDataList
      .Visible = False
      If p_bolHeaderOnly = False Then
         .Clear
         .Cols = 8: .FixedRows = 1
         .ColAlignmentFixed = flexAlignCenterCenter
      End If
      .ColWidth(0) = 1300

      For iRow = 0 To UBound(strRowN)
          .TextMatrix(iRow, 0) = strRowN(iRow)
      Next
      .Visible = True
   End With
End Sub

'intQuery:1-個人/2.主管
Private Function doQuery1(intQuery As Integer) As Boolean
    Dim stCon As String, stConST As String, stConR1 As String, stConR2 As String, stConPE As String
    Dim strQ As String, strQDate As String
    Dim i As Integer, j As Integer
    Dim iRow As Integer, iCol As Integer
    Dim C_Input As Integer, C_InputOk As Integer '需要輸實績與結餘/已經輸實績與結餘
    Dim dblSumR As Double, dblPoint As Double
    Dim strYM As String 'Add by Amy 2020/01/03

    Erase dblTot
    stCon = "": stConST = "": stConPE = ""

On Error GoTo ErrHnd
    'Memo by Amy 2021/05/28 修改實績相關ex:期初保留->期初實績 文字
    If intQuery = 1 Then
        strQDate = Text1(0) & Text1(1)
        'Modfiy by Amy 2016/04/07 +bolGlMan參數(讓財務可輸入特殊編號版人員)
        strQ = GetPoint(1, strQDate, strQDate, , , Left(CboEmp, 5), , Me.Name, bolGlMan)
    Else
        cnnConnection.Execute "Delete From R210152 Where ID='" & strUserNum & "'"
        strQDate = Text11(0) & Text11(1)
       cnnConnection.Execute "Delete From R210152 Where ID='" & strUserNum & "'"
      strQDate = Text11(0) & Text11(1)
      Call InsertR210152(strQDate)
      
        'Modify by Amy 2016/03/25 若SalesPoint還沒輸期末資料另設
        'Modify by Amy 2021/05/06 離職人要輸 原:Decode(st04,2,'F0000',st01) ex:11004 A9007陳莫茗 離職且主管未輸,未全關閉前「智權點數實績與結餘分析表」報出值會錯
        'Modify by Amy 2024/07/10 區主管排前
        strQ = "Select ST02,R03 as PE04,0,0,R04 as C1,R06 as C3,0,Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,SP03,SP07),SP11),SP15) as C5,SP19," & _
                    "R05 as C2,R07 as C4,0,Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,SP24,SP28),SP32),SP36) as C6,SP40," & _
                    "st01 as ChkStNo,sp03||sp07||sp11||sp15||sp19||sp20||sp24||sp28||sp32||sp36||sp40||sp41 as ChkField " & _
                    "From R210152,(Select * From SalesPoint Where SP01=" & Val(strQDate) + 191100 & "),Staff,Acc090 " & _
                    "Where ID='" & strUserNum & "' And R01=SP02(+) And R01=ST01(+) And SP48=A0901(+) " & _
                    "Order by Decode(A0918,R01,1,Decode(A0908,R01,1,2)),R01 "
    End If

    intI = 1
    Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strQ)
    If intI = 1 Then
        '個人資料
        If intQuery = 1 Then
            With AdoRecordSet3
                If Val("" & .Fields("PE04")) > 0 Then Label21(0) = Format(Val("" & .Fields("PE04")), "0.000") '目標
                Label31(0) = Format(Val("" & .Fields("C1")), "0.000")  '期初實績
                Label31(6) = Format(Val("" & .Fields("C3")), "0.000")  '當月實績
                Label41(0) = Format(Val("" & .Fields("C2")), "0.000")  '期初結餘
                Label41(6) = Format(Val("" & .Fields("C4")), "0.000")  '當月結餘
            End With
        '全區資料
        Else
            cmdSave(1).Enabled = False
            grdDataList.Visible = False
            With AdoRecordSet3
                iCol = 1
                Erase dblTot
                Do While Not .EOF
                    dblSumR = 0
                    grdDataList.Cols = iCol + 1
                    '智權人員
                    iRow = GetValue("智權人員")
                    grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("ST02")
                    'Add by Amy 2021/01/18
                    If LenB("" & .Fields("ST02")) >= 12 Then
                        grdDataList.ColWidth(iCol) = 1400
                    Else
                        grdDataList.ColWidth(iCol) = 1000
                    End If
                    grdDataList.CellAlignment = flexAlignCenterCenter
                    '目　　標
                    iRow = GetValue("目標")
                    grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("PE04")
                    grdDataList.CellAlignment = flexAlignRightCenter
                    '達成點數
                    iRow = GetValue("達成點數")
                    dblPoint = Val("" & .Fields("C1")) + Val("" & .Fields("C2")) + Val("" & .Fields("C3")) + _
                                    Val("" & .Fields("C4")) - Val("" & .Fields("C5")) - Val("" & .Fields("C6")) + Val("" & .Fields("SP19")) + Val("" & .Fields("SP40"))
                    grdDataList.TextMatrix(iRow, iCol) = Round(dblPoint, 3)
                    dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
                    dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
                    '達成率
                    iRow = GetValue("達 成 率")
                    If Val(grdDataList.TextMatrix(GetValue("目標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) > 0 Then
                       grdDataList.TextMatrix(iRow, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目標"), iCol)), "0.000") & "%"
                    Else
                       grdDataList.TextMatrix(iRow, iCol) = "0%"
                    End If
                    '*** 記錄需要輸實績與結餘的個數 ***
                    'Modify by Amy 2019/10/16 開放F及W部門使用
                    'If Val("" & .Fields("C1")) + Val("" & .Fields("C3")) + Val("" & .Fields("C2")) + Val("" & .Fields("C4")) > 0 _
                      And (Not (.Fields("ChkStNo") < "6" Or .Fields("ChkStNo") > "F")) Then
                    'Modify by Amy 2020/06/16 開放柄佑輸20091 原:.Fields("ChkStNo") < "6"
                    'Modify by Amy 2021/02/03 +需輸點數之非S部門,若為正常員編不需輸點數,但主管需確認 ex:11001 A6034(陳蒲璇) 不需確認
                    'Modify by Amy 2021/06/03 財務登入strSP48為M31導致顏色未變色,加管理部門顯示判斷
                    If Val("" & .Fields("C1")) + Val("" & .Fields("C3")) + Val("" & .Fields("C2")) + Val("" & .Fields("C4")) > 0 _
                      And ((Left(strSP48, 1) = "S" And Not (.Fields("ChkStNo") < "1" Or .Fields("ChkStNo") > "F")) _
                      Or (Left(strSP48, 1) <> "S" And Left(strSP48, 1) <> "M" And Left(strSP48, 1) = Left(.Fields("ChkStNo"), 1)) _
                      Or (Left(strSP48, 1) = "M" And (Left(cboSalesArea(1), 1) = "S" And Not (.Fields("ChkStNo") < "1" Or .Fields("ChkStNo") > "F")) _
                                                                Or (Left(cboSalesArea(1), 1) <> "S" And Left(cboSalesArea(1), 1) = Left(.Fields("ChkStNo"), 1))) _
                      ) Then
                        C_Input = C_Input + 1
                        'Add by Amy 2016/05/09 記錄已輸實績與結餘的個數
                        '因新人期初期末沒資料可不輸SalesPoint但輸了,而該輸的沒輸,避免計算有誤,故寫於此
                        'Memo by Amy 2021/04/29 林純真可不輸自己(不需確認),但P2005必輸且需確認
                        If Not IsNull(.Fields("ChkField")) Then C_InputOk = C_InputOk + 1
                    End If
                    '*** End 記錄需要輸實績與結餘的個數 ***

                    '期初實績
                    iRow = GetValue("期初實績")
                    grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("C1")
                    '當月實績
                    iRow = GetValue("當月實績")
                    grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("C3")
                    '期末實績
                    iRow = GetValue("期末實績")
                    'Modify by Amy 2016/03/25 個人尚未輸入,期末實績=期初實績
                    grdDataList.TextMatrix(iRow, iCol) = IIf(IsNull(.Fields("C5")) = True, "" & .Fields("C1"), .Fields("C5"))
                    '實績增減
                    iRow = GetValue("實績動用")
                    grdDataList.TextMatrix(iRow, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初實績"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績"), iCol)), 3)
                    '轉撥實績增減
                    iRow = GetValue("轉撥實績增減")
                    grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("SP19")
                    '報出實績點數
                    iRow = GetValue("報出實績點數")
                    'Modify by Amy 2021/05/06 原:Val("" & .Fields("C5")) ex:11004 A9007陳莫茗 離職且主管未輸,顯示會錯
                    grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C1")) + Val("" & .Fields("C3")) - Val(grdDataList.TextMatrix(GetValue("期末實績"), iCol)) + Val("" & .Fields("SP19")), 3)
                    '期初結餘
                    iRow = GetValue("期初結餘")
                    grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("C2")
                    '當月結餘
                    iRow = GetValue("當月結餘")
                    grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("C4")
                    '期末結餘
                    iRow = GetValue("期末結餘")
                    'Modify by Amy 2016/03/25 個人尚未輸入,期末結餘=期初結餘+當月結餘
                    grdDataList.TextMatrix(iRow, iCol) = IIf(IsNull(.Fields("C6")) = True, Round(Val("" & .Fields("C2")) + Val("" & .Fields("C4")), 3), .Fields("C6"))
                    '結餘增減
                    iRow = GetValue("結餘動用")
                    grdDataList.TextMatrix(iRow, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初結餘"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘"), iCol)), 3)
                    '轉撥結餘增減
                    iRow = GetValue("轉撥結餘增減")
                    grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("SP40")
                    '報出實績點數
                    iRow = GetValue("報出結餘點數")
                    'Modify by Amy 2021/05/06 原:Val("" & .Fields("C6"))
                    grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C2")) + Val("" & .Fields("C4")) - Val(grdDataList.TextMatrix(GetValue("期末結餘"), iCol)) + Val("" & .Fields("SP40")), 3)

                    'Modify by Amy 2020/06/16 開放柄佑輸20091 原:.Fields("ChkStNo") < "6"
                    'Modify by Amy 2021/01/19 F4102拆成F4104,F4105/F4103拆成F4106,F4107,需確認
                    'Modify by Amy 2021/02/03 +需輸點數之非S部門,若為正常員編不需輸點數,但主管需確認 ex:11001 A6034(陳蒲璇) 不需確認
                    'Modify by Amy 2021/06/03 財務登入strSP48為M31導致顏色未變色,加管理部門顯示判斷
                    If IsNull(.Fields("ChkField")) _
                      And ((Left(strSP48, 1) = "S" And Not (.Fields("ChkStNo") < "1" Or .Fields("ChkStNo") > "F")) _
                      Or (Left(strSP48, 1) <> "S" And Left(strSP48, 1) <> "M" And Left(strSP48, 1) = Left(.Fields("ChkStNo"), 1)) _
                      Or (Left(strSP48, 1) = "M" And (Left(cboSalesArea(1), 1) = "S" And Not (.Fields("ChkStNo") < "1" Or .Fields("ChkStNo") > "F")) _
                                                                    Or (Left(cboSalesArea(1), 1) <> "S" And Left(cboSalesArea(1), 1) = Left(.Fields("ChkStNo"), 1))) _
                      ) Then
                        '未輸過實績與結餘變色
                        For j = 1 To grdDataList.Rows - 1
                            grdDataList.col = iCol
                            grdDataList.row = j
                            grdDataList.CellBackColor = &HFFC0C0
                        Next j
                    Else
                        For j = 1 To grdDataList.Rows - 1
                            grdDataList.col = iCol
                            grdDataList.row = j
                            grdDataList.CellBackColor = &H80000018
                        Next j
                    End If

                    .MoveNext
                    iCol = iCol + 1
                Loop
            End With

            grdDataList.Cols = iCol + 1
            grdDataList.TextMatrix(0, iCol) = "合　　計"
            '總目標需抓該區所有(因為有可能點數掛該區代碼如20021)
            'Modify by Amy 2017/09/26 因中一高國碩/陳頌恩轉中二部門增加抓SP48
            strExc(0) = "Select nvl(sum(PE04),0) PE04 From Staff,PerFormance,SalesPoint " & _
                 "Where PE01(+)=ST01 And PE02(+)='TOT' And Decode(SP48,null,ST15,SP48)='" & Left(cboSalesArea(1), 3) & "' And  PE03(+) = " & Val(Text11(0) & Text11(1)) + 191100 & _
                 " And ST01=SP02(+) And " & Val(Text11(0) & Text11(1)) + 191100 & "=SP01(+)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                dblTot(1) = RsTemp.Fields(0)
            End If
            For i = 1 To UBound(strRowN)
                Select Case i
                    Case GetValue("目標")
                        grdDataList.TextMatrix(i, iCol) = dblTot(i)
                    Case GetValue("達成點數")
                        grdDataList.TextMatrix(i, iCol) = dblTot(i)
                        'add by sonia 2016/1/26
                        For j = 1 To grdDataList.Cols - 1
                            grdDataList.col = j
                            grdDataList.row = i
                            grdDataList.CellBackColor = &HC000&
                        Next j
                        'end 2016/1/26

                    Case GetValue("達 成 率")
                        If Val(grdDataList.TextMatrix(GetValue("目標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) > 0 Then
                            grdDataList.TextMatrix(i, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目標"), iCol)), "0.000") & "%"
                        Else
                            grdDataList.TextMatrix(i, iCol) = "0%"
                        End If
                    Case Else
                        dblTot(i) = 0
                        For j = 1 To grdDataList.Cols - 1
                            dblTot(i) = Round(Val(dblTot(i)) + Val(grdDataList.TextMatrix(i, j)), 3)
                        Next j
                        grdDataList.TextMatrix(i, iCol) = dblTot(i)
                        'add by sonia 2016/1/26
                        If i = GetValue("報出實績點數") Or i = GetValue("報出結餘點數") Then
                           For j = 1 To grdDataList.Cols - 1
                               grdDataList.col = j
                               grdDataList.row = i
                               grdDataList.CellBackColor = &HC000&
                           Next j
                        End If
                        'end 2016/1/26
                End Select
            Next i
            grdDataList.Visible = True
            If C_Input = C_InputOk And bolIsAccept = False Then
                'Moidfy by Amy 2019/10/16 +日期判斷,開放W/F部門操作,開放後查之前資料未有人確認,不可按主管確認鈕
                'Modify by Amy 2020/01/03 原:Val(strMaxSP01) = Val(Left(strSrvDate(1), 6) - 191101), 年底會有問題
                strYM = Left(strSrvDate(1), 6)
                If Right(strYM, 2) = "01" Then
                    strYM = Val(Left(strYM, 4)) - 1912 & "12"
                Else
                    strYM = Val(strYM) - 191101
                End If
                'Memo by Amy 2021/07/05 瑞婷說看不到主管是否確認,因只能是「區主管」確認,避免財務按到,故仍鎖住
                If (bolAreaMan = True Or (Left(cboSalesArea(1), 3) = "S00" And stST15 = "M31")) _
                   And Val(Text11(0) & Text11(1)) = Val(strYM) And Val(strMaxSP01) = Val(strYM) Then
                'end 2020/01/03
                    'Modify by Amy 2019/08/01 10809(輸10808月資料)開放F4102王文安可操作,F4103陳鳳英及其職代A0914可操作,但王副總只能看
                    'Moidify by Amy 2019/10/16 10810月開放W部門操作 區員編 ex:W1001
'                    If Left(strSP48, 2) = "F1" Or Left(strSP48, 2) = "F2" Or Left(strSP48, 1) = "W" Then
'                        '開放F部門10808開始操作
'                        If Val(Text11(0) & Text11(1)) >= 10808 And Left(strSP48, 1) = "F" Then cmdSave(1).Enabled = True
'                    Else
                        'Modify by Amy 2022/05/24 bug-杜燕文協理只能操作自己部門的主管確認(杜協理接簡協理工作,簡協理非北五區主管不會有問題)
                        If stST15 <> "P10" And Left(cboSalesArea(1), 3) = stST15 Then cmdSave(1).Enabled = True
'                    End If
                End If
            End If
        End If
        doQuery1 = True
    End If

ErrHnd:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Function GetValue(pRowN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strRowN)
       If UCase(strRowN(jj)) = UCase(pRowN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function


'Modify by Amy 2017/09/26 中一高國碩/陳頌恩轉中二,轉前與轉後顯示部門需不同
Private Sub SetcboSalesArea(ByVal m_ST15 As String, Optional ByVal bolAll As Boolean = True)
    Dim strTp1, strTp2
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strWhere As String
    
    If Trim(Text1(0) & Text1(1)) = MsgText(60) Then
        MsgBox "請輸入業績年月", , MsgText(5)
        Text1(0).SetFocus
    End If
            
    cboSalesArea(0).Clear
    If bolAll = True Then cboSalesArea(1).Clear
    '只能看同一區且為個人或某區主管
    'Modify by Amy 2016/08/17 將Pub_StrUserSt15(登入者)改抓stST15(可能為職代操作區主管權限)
    If strArea1 = strArea2 And strArea1 <> MsgText(601) And (strAreaList = MsgText(601) Or Replace(strAreaList, "'", "") = strArea1) Then
        cboSalesArea(0) = m_ST15 & " " & A0902Query(m_ST15)
        'cboSalesArea(0).Enabled = False 'Mark by Amy 2017/09/26 因中一高國碩/陳頌恩轉中二,避免抓錯(先不考慮st52跨區問題-秀玲)
        If bolAll = True Then
            cboSalesArea(1) = m_ST15 & " " & A0902Query(m_ST15)
            'cboSalesArea(1).Enabled = False  'Mark by Amy 2017/09/26 因中一高國碩/陳頌恩轉中二,避免抓錯(先不考慮st52跨區問題-秀玲)
        End If
    Else
        'Modify by Amy 2016/04/07 縮減下部門下拉選單list (因SalesPoint 有新增201512月名單為2016年以前人員部門資料)
        '總經理登入自己部門放第一個選項
        If strGlManNo = strNowUser Then
            strQ = "Select a0901,a0902 From Acc090," & _
                    "(Select Distinct sp48,'1' as Sort From SalesPoint Where sp48='" & m_ST15 & "' And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & _
                    " Union Select Distinct sp48,'2' as Sort From SalesPoint Where SubStr(sp48,1,1)='S' And sp48<>'" & m_ST15 & "' And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & _
                    " Union Select Distinct sp48,'3' as Sort From SalesPoint Where SubStr(sp48,1,1)<>'S' And sp48<>'" & m_ST15 & "'  And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & ")" & _
                    " Where a0901=sp48 Order by Sort,a0901"
        '可看全部資料
        ElseIf strArea1 = MsgText(601) And strArea1 = strArea2 Then
                'Modify by Amy 2021/11/11 原sp01=" & Val(Text1(0) & Text1(1)) + 191100,因總經理權限可輸入,故設定時抓全部SalesPoint 201601月後有的部門資料
                strQ = "Select a0901,a0902,Sort From Acc090," & _
                    "(Select Distinct sp48,'1' as Sort From SalesPoint Where SubStr(sp48,1,1)='S' And sp01<>201512 " & _
                    " Union Select Distinct sp48,'2' as Sort From SalesPoint Where SubStr(sp48,1,1)<>'S' And sp01<>201512" & _
                    " ) Where a0901=sp48 Order by Sort,a0901"
           
        Else
            'Add by Amy 2016/11/18 林柄佑經理可以看自已和中所全部
            If strNowUser = "82026" Then
                strWhere = " And ( sp48=" & CNULL(m_ST15) & " Or (sp48>=" & CNULL(strArea1) & " And sp48<=" & CNULL(strArea2) & "))"
            'Add by Amy 2022/05/24 協理可看部門
            ElseIf Len(strArea1) = 1 Then
                strWhere = " And (Substr(sp48,1,1)='" & strArea1 & "'  Or sp48 In('" & Replace(strArea2, ",", "','") & "') )"
            ElseIf strAreaList <> MsgText(601) Then
                '區主管級人員
                If InStr(strAreaList, ",") > 0 Then
                    strWhere = "sp48 in (" & strAreaList & ")"
                Else
                    strWhere = "sp48=" & CNULL(strAreaList)
                End If
                strWhere = " And (" & strWhere & " Or (sp48>=" & CNULL(strArea1) & " And sp48<=" & CNULL(strArea2) & "))"
            Else
                strWhere = " And sp48>=" & CNULL(strArea1) & " And sp48<=" & CNULL(strArea2) & " "
            End If
            
            '其他人員,自已部門排前
            strQ = "Select Distinct a0901,a0902 From Acc090,SalesPoint Where sp48=a0901 And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & strWhere & _
                        " Order by Decode(SubStr(a0901,1,1),'" & Left(m_ST15, 1) & "',1,2),a0901"
        End If
        'end 2016/04/07
     
        RsQ.CursorLocation = adUseClient
        RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
        If RsQ.RecordCount > 0 Then
            cboSalesArea(0).AddItem "": cboSalesArea(1).AddItem ""
            cboSalesArea(0).ListIndex = 0: cboSalesArea(1).ListIndex = 0
            RsQ.MoveFirst
            Do While Not RsQ.EOF
                cboSalesArea(0).AddItem RsQ.Fields("a0901") & " " & RsQ.Fields("a0902")
                'Mark by Amy 2020/06/16 開放柄佑 輸20091,已設其st15為S29 故不會查到自己(P11)部門
                'Modify by Amy 2016/11/18 +林柄佑經理權限
'                If strNowUser = "82026" And rsQ.Fields("a0901") = stST15 Then
'                    '只能看中所全部的全區,不能看自己部門的全區
'                Else
                    cboSalesArea(1).AddItem RsQ.Fields("a0901") & " " & RsQ.Fields("a0902")
'                End If
                'end 2020/06/16
                RsQ.MoveNext
            Loop
        End If
        RsQ.Close
        
        If (stST05 <> "00" And stST05 <> "01" And stST05 <> "08") Or strGlManNo = strNowUser Then
              cboSalesArea(0) = m_ST15 & " " & A0902Query(m_ST15)
              'Modify by Amy 2016/11/18
              'Mark by Amy 2020/06/16 拿掉strNowUser <> "82026" ,讓柄佑輸20091,一開始預帶S29
              'If strNowUser <> "82026" Then
              cboSalesArea(1) = m_ST15 & " " & A0902Query(m_ST15)
              'end 2020/06/16
        End If
    End If
    'end 2016/08/17
End Sub

'intSet:0-一般/1:run search/2:不預設人員/3:特殊人員(for 高國碩 查舊資料)只顯示自己
Private Sub SetEmp_Old(Optional ByVal intSet As Integer = 0)
'    Dim strTemp
'    Dim strQ As String, strWhere As String
'    Dim bolSelf As Boolean
'    Dim stTP(2) As String, stDef As String 'Add by Amy 2019/10/16
'
'    If bolNowChk = True Then Exit Sub
'
'    bolNowChk = True
'    If bolGlMan = True And cboEmp <> MsgText(601) And cboEmp <> "cboEmp" Then
'         'Add by Amy 2019/10/23 原有資料不清
'        If InStr(cboEmp, " ") > 0 Then
'            stDef = Left(cboEmp, InStr(cboEmp, " ") - 1)
'        Else
'            stDef = cboEmp
'        End If
'    End If
'    cboEmp.Clear
'    'Modify by Amy 2016/02/23 改SalesPoint有資料才預設(加IsRecordExist(Val(Text1(0) & Text1(1)) + 191100, strNowUser) = True )
'    'Modify by Amy 2017/09/26 部門改抓strSP48 原stST15
'    'Modify by Amy 2017/12/01 拿掉SalesPoint有資料的判斷,因沒值人員會是空
'    'Modify by Amy 2019/08/01 +if 開放王文安(F21)輸F4102/陳鳳英(F11)輸F4103及其職代A0914
'    'Moidfy by Amy 2019/10/23 登入或下拉選  F1 or F2開頭部門
'    If (cboSalesArea(0).Enabled = False And (Left(stST15, 2) = "F2" Or Left(stST15, 2) = "F1")) _
'      Or (cboSalesArea(0).Enabled = True And (Left(cboSalesArea(0), 2) = "F2" Or Left(cboSalesArea(0), 2) = "F1")) Then
'        bolSelf = True
'        'Modify by Amy 2021/01/18 11001月F4102拆成F4104、F4105/F4103拆成F4106、F4107,原編號保留因要查舊資料
'        If Left(stST15, 2) = "F1" Or Left(cboSalesArea(0), 2) = "F1" Then
'            'Modify by Amy 2021/07/06 +if
'            If strNowUser <> "98020" Then
'                '洪琬姿80030輸F4106,May 78011輸F4107
'                stTP(0) = "ST01"
'                Call GetDeptList(1, strNowUser, , stTP(0))
'                stTP(0) = Replace(stTP(0), "'", "") '取代單引號
'                stTP(1) = GetStaffName(stTP(0), True)
'                cboEmp.AddItem stTP(0) & " " & stTP(1)
'                cboEmp = stTP(0) & " " & stTP(1)
'            Else
'                cboEmp.AddItem "F4106" & " " & GetStaffName("F4106", True)
'                cboEmp.AddItem "F4107" & " " & GetStaffName("F4107", True)
'                cboEmp.AddItem "F4103" & " " & GetStaffName("F4103", True)
'                cboEmp = "F4106" & " " & GetStaffName("F4106", True)
'            End If
'        Else
'            cboEmp.AddItem "F4104" & " " & GetStaffName("F4104", True)
'            cboEmp.AddItem "F4105" & " " & GetStaffName("F4105", True)
'            cboEmp.AddItem "F4102" & " " & GetStaffName("F4102", True)
'            cboEmp = "F4104" & " " & GetStaffName("F4104", True)
'        End If
'        'end 2021/01/18
'    'Add by Amy 2019/10/16 開放W部門區主管輸該區編號
'    ElseIf (cboSalesArea(0).Enabled = False And Left(stST15, 1) = "W") _
'      Or (cboSalesArea(0).Enabled = True And Left(cboSalesArea(0), 1) = "W") Then
'        bolSelf = True
'        'Modify by Amy 2020/07/03 bug-+if 財務/電腦中心 選W部門,會顯示財務/電腦中心 部門主管
'        If stST15 = "M51" Or stST15 = "M31" Or strGlManNo = strNowUser Then
'            stTP(0) = GetAreaEmpNo(Left(cboSalesArea(0), 3))
'        'Add by Amy 2021/09/02 +S00 且選 W10部門可看W1001
'        ElseIf Pub_StrUserSt03 = "S00" Then
'            stTP(0) = "W1001"
'        Else
'            stTP(0) = GetAreaEmpNo(stST15)
'        End If
'        'end 2020/07/03
'        stTP(1) = GetStaffName(stTP(0), True)
'        cboEmp.AddItem stTP(0) & " " & stTP(1)
'        cboEmp = stTP(0) & " " & stTP(1)
'    Else
'        '部門下拉選單前3碼與目前登入者部門相同(換區會不同)或總經理登入(總經理可看其他人的,且可用總經理權限操作)且登入者不是財務且不是電腦中心
'        'Modify by Amy 2020/06/16 柄佑 82026 選單內不可有自己
'        If (Left(cboSalesArea(0), 3) = strSP48 Or (strGlManNo = strNowUser And Left(cboSalesArea(0), 3) = strSP48)) And stST15 <> "M31" _
'          And stST15 <> "M51" And strNowUser <> "82026" Then
'            cboEmp.AddItem strNowUser & " " & GetStaffName(strNowUser, True)
'            bolSelf = True
'        End If
'
'        'Modify by Amy 2020/06/16 +if 82026 為中區的帶人主管,不需再顯示於智權人員下拉中
'        If strNowUser <> "82026" Then
'            If strSt52List <> MsgText(601) Then
'                If InStr(strSt52List, ",") > 0 Then
'                    strWhere = " And sp02 in (" & strSt52List & ")"
'                Else
'                    strWhere = " And sp02=" & strSt52List
'                End If
'            End If
'        End If
'        'end 2020/06/16
'
'        '一般個人或帶人主管
'        If cboSalesArea(0).Enabled = False And bolAreaMan = False Then
'            '帶人主管人員
'            If strSt52List <> MsgText(601) Then
'                strQ = "Select st01,st02 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & strWhere & _
'                            " Order by st01"
'            End If
'        '區主管以上人員
'        Else
'            '總經理權限(電腦中心,財務,總經理,主任秘書(等級08))
'            If bolGlMan = True And stDef <> MsgText(601) And Val(Text1(0) & Text1(1)) <> 0 Then
'                '以人員及日期重抓其部門,因查中一高國碩/陳頌恩轉/林青祺(轉區)
'                strQ = "Select st01,st02 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & " And sp48 in " & _
'                            "(Select sp48 From SalesPoint Where  sp01=" & Val(Text1(0) & Text1(1)) + 191100 & " And sp02='" & stDef & "' )" & _
'                          " Order by st01"
'            '王副總只能看自己和P1001
'            ElseIf strNowUser = "71011" Then
'                strQ = "Select st01,st02 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Val(Text1(0)) & Text1(1)) + 191100 & " And sp48='" & Left(cboSalesArea(0), 3) & "' " & _
'                            " And sp02='P1001' "
'            'Add by Amy 2016/11/18 林柄佑經理可以看自已和中所全部,不需加同部門其他人員
'            'Modify by Amy 2020/06/16 柄佑 要輸20091,可看中所全部,但不可看自己(2019/11/21改),故加語法
''            ElseIf strNowUser = "82026" And Left(cboSalesArea(0), 3) = strSP48 Then
''                '不需加同部門其他人員
'            ElseIf strNowUser = "82026" Then
'                strQ = "Select st01,st02 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Val(Text1(0)) & Text1(1)) + 191100 & " And sp48='" & Left(cboSalesArea(0), 3) & "' "
'            'end 2020/06/16
'            '帶人主管
'            ElseIf strSt52List <> MsgText(601) Then
'                'Modify by Amy 2017/10/13 中一高國碩/陳頌恩轉中二,修正原部門抓st15轉後看不到舊資料
'                strQ = "Select Distinct st01,st02 From (" & _
'                        "Select st01,st02 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Val(Text1(0)) & Text1(1)) + 191100 & " And sp48='" & Left(cboSalesArea(0), 3) & "' " & _
'                        IIf(Left(cboSalesArea(0), 3) = strSP48, "And sp02<> '" & strNowUser & "' ", "") & _
'                        "Union Select st01,st02 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & strWhere & _
'                        ") Order by st01"
'            '非帶人主管
'            Else
'                'Modify by Amy 2017/10/13 中一高國碩/陳頌恩轉中二,修正原部門抓st15轉後看不到舊資料
'                strQ = "Select st01,st02 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & " And sp48='" & Left(cboSalesArea(0), 3) & "' " & _
'                        IIf(Left(cboSalesArea(0), 3) = strSP48, "And sp02<> '" & strNowUser & "' ", "") & _
'                         " Order by st01"
'            End If
'        End If
'
'        If strQ <> MsgText(601) Then
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strQ)
'            If intI = 1 Then
'                RsTemp.MoveFirst
'                Do While Not RsTemp.EOF
'                    cboEmp.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'                    RsTemp.MoveNext
'                Loop
'            End If
'        End If
'
'        '若設定狀態 非不預設人員 且目前登錄者為員工自已,則智權人員預設登入者自已
'        If intSet <> 2 And bolSelf = True Then cboEmp = strNowUser & " " & GetStaffName(strNowUser, True)
'        '若設定狀態為「查詢」,則自動執行查詢
'        If intSet = 1 Then
'            If bolSelf = True Then
'                'Modify by Amy 2019/10/16 +bolNoChkMod 避免無窮迴圈
'                bolNoChkMod = True
'                call cmdSearch_Click(0)
'                bolNoChkMod = False
'            ElseIf bolGlMan = True And stDef <> MsgText(601) Then
'                cboEmp = stDef & " " & GetStaffName(stDef, True)
'            End If
'        End If
'    End If
'    'end 2019/08/01
'
'    cboEmp.Enabled = False
'    If cboEmp.ListCount > 1 Or cboSalesArea(0).Enabled = True Then
'        cboEmp.Enabled = True
'    End If
'    bolNowChk = False
End Sub

Private Function FormCheck1(Index As Integer) As Boolean
    Dim bCancel As Boolean
    Dim strVal As String
    Dim strMsg As String 'Add by Amy 2017/05/19
    
    If cboSalesArea(Index) = MsgText(601) Then
        MsgBox "業務區不可空白", , MsgText(5)
        cboSalesArea(Index).SetFocus
        Exit Function
    End If
    If Index = 0 Then
        If CboEmp = MsgText(601) Then
            MsgBox "智權人員不可空白", , MsgText(5)
            CboEmp.SetFocus
            Exit Function
        End If
        If Text1(0) = MsgText(601) Then
            MsgBox "業績年不可空白", , MsgText(5)
            Text1(0).SetFocus
            Exit Function
        End If
        If Text1(1) = MsgText(601) Then
            MsgBox "業績月份不可空白", , MsgText(5)
            Text1(1).SetFocus
            Exit Function
        End If
        Call Text1_Validate(1, bCancel)
        If bCancel = True Then
            Text1(1).SetFocus
            Exit Function
        End If
        'Add by Amy 2016/04/07 +可存檔才判斷
        If cmdSave(0).Enabled = True Then
            If Check1(0).Value = 1 And Trim(Text2) = MsgText(601) Then
                MsgBox "有勾選「報出點數」需輸入值", , MsgText(5)
                Exit Function
            End If
            If Check1(0).Value = 1 And Check1(1).Value = 0 And Check1(2).Value = 0 Then
                MsgBox "請點要輸入「實績部分」或「結餘部分」", , MsgText(5)
                Exit Function
            End If
            '2017/05/19 從FormSave 搬過來 'Add by Amy 2016/04/07 總經理權限可以輸負值
            If bolGlMan = False Then
                If ChkMinus(strMsg) = True Then
                    MsgBox strMsg & "數值不可為負數"
                    Exit Function
                End If
            'Add by Amy 2017/09/26 總經理權限者檢查期末不可為負數
            Else
                If Val(Label31(12)) < 0 Then
                    MsgBox "期末實績不可為負數"
                    Exit Function
                End If
                If Val(Label41(12)) < 0 Then
                    MsgBox "期末結餘不可為負數"
                    Exit Function
                End If
            End If
            'Add by Amy 2016/02/16
            If CboEmp.Tag <> MsgText(601) Then
                'Modify by Amy 2016/04/07 +只有智權需判斷且在職才判斷
                strExc(0) = PUB_GetStaffST15(CboEmp.Tag, 1)
                If Text5(0) <> MsgText(601) Then strVal = Text5(0) '[個人]報出實績點數
                If Label31(1) <> MsgText(601) Then strVal = Label31(1) '[主管]報出實績點數
                'Modify by Amy 2023/06/09 拿掉Left(strExc(0), 1) = "S",需要輸的人員都判斷
                'If Left(strExc(0), 1) = "S" And bolLeave = False Then
                'end 2016/04/07
                If bolLeave = False Then
                'end 2023/06/09
                    '目標>=當月實績且報出實績<當月實績 ex:10501月 A2025 黃誠安 目標>=當月實績,無法存檔
                    If Val(Label21(0)) >= Val(Label31(6)) And Val(strVal) < Val(Label31(6)) Then
                        MsgBox "目標 >= 當月實績" & vbCrLf & "報出實績點數 不可小於 當月實績", , MsgText(5)
                        If Text5(0).Locked = False Then Text5(0).SetFocus
                        Exit Function
                    End If
                End If
                'Add by Amy 2023/06/07 只要需輸入都要判斷 ex:11205月 87011林青祺 目標760,當月796.9,報出點數609.9 應該不可存-杜協理
                '有 目標 且 當月實績>=目標,檢查 報出點數<目標 ,不可存
                If Val(Label21(0)) <> 0 And Val(Label31(6)) >= Val(Label21(0)) And Val(Text2) < Val(Label21(0)) Then
                     MsgBox "當月實績 >=目標" & vbCrLf & "報出點數 不可小於 目標"
                     If Text5(0).Locked = False Then Text5(0).SetFocus
                     Exit Function
                End If
            End If
        End If
        'Added by Lydia 2022/01/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
        If TxtSP(20) <> "" Or TxtSP(41) <> "" Then
          If PUB_ChkUniText(Me, , True, "TextBox") = False Then
              Exit Function
          End If
        End If
        'end 2022/01/03
    '主管確認
    Else
         If Text11(0) = MsgText(601) Then
            MsgBox "業績年不可空白", , MsgText(5)
            Text11(0).SetFocus
            Exit Function
        End If
        If Text11(1) = MsgText(601) Then
            MsgBox "業績月份不可空白", , MsgText(5)
            Text11(1).SetFocus
            Exit Function
        End If
        Call Text11_Validate(1, bCancel)
        If bCancel = True Then
            Text11(1).SetFocus
            Exit Function
        End If
    End If
    
    FormCheck1 = True
End Function

'檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, Optional ByRef stGetData As String = "*") As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select " & stGetData & " From SalesPoint Where SP01=" & Val(strKEY01) & " And SP02='" & strKEY02 & "'"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    ' 檢查讀取的資料筆數
    If rsTmp.RecordCount > 0 Then
        IsRecordExist = True
        If stGetData = "*" Then
            stGetData = ""
        Else
            stGetData = "" & rsTmp.Fields(0)
        End If
    Else
        IsRecordExist = False
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

'主管是否已確認
Private Function IsAccept(ByVal idx As Integer, ByVal stDeptCode As String) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, stDate As String
    
    IsAccept = False
    
    If idx = 0 Then
        stDate = Val(Text1(0) & Text1(1)) + 191100
    Else
        stDate = Val(Text11(0) & Text11(1)) + 191100
    End If
    strSql = "Select sp45 From SalesPoint,Staff Where sp01=" & stDate & " And sp02=st01(+) And st15='" & stDeptCode & "' And SP45 is not null"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly

    If rsTmp.RecordCount > 0 Then
        IsAccept = True
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

Private Function ReadSalesPoint(stSP01 As String, stSP02 As String) As ADODB.Recordset
    Dim rsRecordset As New ADODB.Recordset
    Dim stQ As String
    
    stQ = "Select * From SalesPoint Where SP01=" & Val(stSP01) & " And SP02='" & stSP02 & "'"
    rsRecordset.CursorLocation = adUseClient
    rsRecordset.Open stQ, cnnConnection, adOpenStatic, adLockReadOnly
    
    Set ReadSalesPoint = rsRecordset
    Set rsRecordset = Nothing
    
End Function

Private Function FormSave(Index As Integer) As Boolean
    Dim strSql As String, strMsg As String
    Dim intR As Integer
    'Add by Amy 2016/02/16
    Dim strContent As String, strTo As String
    Dim intOpt1 As Integer, intOpt2 As Integer
    Dim strST14 As String 'Add by Amy 2016/04/07 +內部郵件收件員工編號
    Dim strCC As String 'Add by Amy 2016/11/18
    'Add by Amy 2017/02/03
    Dim bolExist As Boolean 'SalesPoint是否有資料
    'Dim strIns As String'Modify by Amy 2017/09/26
    Dim bolAccept As Boolean 'Add by Amy 2021/01/07
    Dim bolSpecMail As Boolean, strTmp As String 'Add by Amy 2021/07/16 特殊mail(職代不寄)/strtmp
    
    'strUpdSB = "" 'Mark by Amy 2022/06/09 不使用
    
    FormSave = False
    If Index = 0 Then
        'Mark by Amy 2017/09/26 輸入員工若無資料需至隱藏版人員那支新增其傳票資料
        'Modify by Amy 2017/02/03 瑞婷輸入的員工SalesPoint沒資料需新增
'        bolNotExist = IsRecordExist(Val(Text1(0) & Text1(1)) + 191100, cboEmp.Tag)
'        If bolNotExist = False Then
'            If bolGlMan = False Then
'                MsgBox "當月無資料！"
'                Exit Function
'            ElseIf cboSalesArea(0) = MsgText(601) Then
'                MsgBox "業務區不可為空！"
'                Exit Function
'            Else
'                strIns = "Insert Into SalesPoint (SP01,SP02,SP48) Values(" & Val(Text1(0) & Text1(1)) + 191100 & ",'" & cboEmp.Tag & "','" & Left(cboSalesArea(0), 3) & "')"
'            End If
'        End If
        'end 2017/02/03
        'end 2017/09/26
        'Modify 原 總經理權限可以輸負值搬至FormCheck1
        strSql = SqlUpd(intOpt1, intOpt2)
    
    '主管確認
    Else
        'Modify by Amy 2021/08/04 原抓st15,11008月中三區人員調至其他各區,因以st15部門抓更新其主管確認欄位,導致11007月的資料以新部門抓資料更新
        strSql = "Update SalesPoint Set sp45='" & strUserNum & "',sp46=to_number(to_char(sysdate,'YYYYMMDD')),sp47=to_number(to_char(sysdate,'HH24MISS')) " & _
                    "Where sp01=" & Val(Text11(0) & Text11(1)) + 191100 & " And sp02 in (" & _
                    "Select sp02 From SalesPoint,Staff Where sp01=" & Val(Text11(0) & Text11(1)) + 191100 & " And sp02=st01(+) And sp48='" & Left(cboSalesArea(1), 3) & "' " & _
                    ")"
    End If
    If strSql <> MsgText(601) Then
On Error GoTo ErrTran:
        cnnConnection.BeginTrans
        'Mark by Amy 2017/09/26
'        'Add by Amy 2017/02/03 瑞婷輸入的員工SalesPoint沒資料需新增
'        If strIns <> MsgText(601) Then
'            cnnConnection.Execute strIns
'        End If
'        'end 2017/02/03
        'Mark by Amy 2022/06/09  不使用
'        'Add by Amy 2017/09/26 增加更新SalesBalance
'        If strUpdSB <> MsgText(601) Then
'            cnnConnection.Execute strUpdSB
'        End If
        cnnConnection.Execute strSql, intR
        'Add by Amy 2019/10/16 F4102/F4103/W1001/W2001,個人未輸而由總經理權限輸入,則直接上主管已確認
        'Modify by Amy 2021/01/18 原stSpecEmpNo改為智權點數實績與結餘特殊員編
        If intLimit = 4 And InStr(智權點數實績與結餘特殊員編, CboEmp.Tag) > 0 And intR >= 1 Then
            Call ChkUpdSP45(CboEmp.Tag)
        End If
        cnnConnection.CommitTrans
        FormSave = True
        If intR >= 1 Then
            If Index = 0 Then
                'Modify by Amy 2016/03/02 簡協理輸自己存檔後選周哲丞彈是否存檔選「是」後發錯對象,發給了周哲丞 原:Left(CboEmp, 5)
                'Modify by Amy 2016/04/07 +內部郵件收件員工編號若為 9997 則不發mail ex:F4102
                strST14 = PUB_GetST14(Left(CboEmp.Tag, 5))
                'Add by Amy 2021/07/16 +if F1(外商)改由 洪琬姿(80030)與葉易雲(78011)輸F4106/F4107 個人,江郁仁(98020)為區主管輸主管欄
                'Modify by Amy 2021/12/06 P2005 改由沈佳穎(96003)輸個人,江郁仁(98020)輸區主管欄
                If InStr("F4103,F4106,F4107,P2005", Left(CboEmp.Tag, 5)) > 0 Then
                    strTo = ""
                    If (m_FieldList(3).fiOldData <> Empty Or m_FieldList(24).fiOldData <> Empty) Then
                        '總經理修改通知區主管
                        If bolGlMan = True Then
                            If strToSpecNo = MsgText(601) Then
                                strTo = ";98020"
                            Else
                                bolSpecMail = True
                                strTo = strToSpecNo
                            End If
                        '江郁仁修改通知st14(洪琬姿(80030)與葉易雲(78011)改個人欄位不通知)
                        ElseIf (Left(CboEmp.Tag, 5) = "F4106" Or Left(CboEmp.Tag, 5) = "F4107") And strNowUser = "98020" Then
                            '江郁仁修改通知st14
                            strTo = ";" & strST14
                            '職代操作發mail 通知江郁仁
                            If strToSpecNo <> MsgText(601) And strUserNum <> strNowUser Then
                                bolSpecMail = True
                                strTo = strTo & Replace(strToSpecNo, ";78011", "") '職代操作(78011)不需再發(78011)
                            End If
                        '江郁仁修改通知st14
                        ElseIf Left(CboEmp.Tag, 5) = "P2005" And strNowUser = "98020" Then
                            strTo = ";" & strST14
                        End If
                    End If
                'Modify by Amy 2020/07/03 員編為F/W開頭及20091自已修改不發mail,總經理權限調整發部門主管
                'Modify by Amy 2021/04/29 員編為F/W開頭及20091改抓智權點數實績與結餘特殊員編
                ElseIf InStr(智權點數實績與結餘特殊員編, Left(CboEmp.Tag, 5)) > 0 Then
                    strTo = ""
                    If bolGlMan = True And (m_FieldList(3).fiOldData <> Empty Or m_FieldList(24).fiOldData <> Empty) Then
                        strTo = ";" & GetDeptMan(PUB_GetStaffST15(Left(CboEmp.Tag, 5), 1))
                    End If
                'Modify by Amy 2017/02/03 員編大於等於5字頭才發mail ex:20091不發
                ElseIf strST14 <> "99997" And Left(CboEmp.Tag, 1) >= "5" Then
                    '總經理權限人員有修改,通知個人及區主管,若帶人主管有輸加發帶人主管
                    If bolGlMan = True And CboEmp.Tag <> strNowUser Then
                        'Modify by Amy 2016/11/18 改收件者顯示個人,cc主管
                        strCC = ";" & GetDeptMan(PUB_GetStaffST15(Left(CboEmp.Tag, 5), 1))
                        '判斷是否有帶人主管
                        If (m_FieldList(8).fiOldData <> Empty And m_FieldList(8).fiOldData <> Mid(strCC, 2)) Or (m_FieldList(12).fiOldData <> Empty And m_FieldList(12).fiOldData <> Mid(strCC, 2)) Then
                            If m_FieldList(8).fiOldData <> Mid(strCC, 2) Then
                                strCC = strCC & ";" & m_FieldList(8).fiOldData
                            Else
                                strCC = strCC & ";" & m_FieldList(12).fiOldData
                            End If
                        End If
                        strTo = ";" & Left(CboEmp.Tag, 5)
                        If InStr(strCC, Left(CboEmp.Tag, 5)) > 0 Then strCC = Replace(strCC, ";" & Left(CboEmp.Tag, 5), "")
                        'end 2016/11/18
                    '個人輸完,主管有調整,通知個人
                    ElseIf (m_FieldList(3).fiOldData <> Empty Or m_FieldList(24).fiOldData <> Empty) _
                       And ((bolAreaMan = True Or strSt52List <> MsgText(601)) And m_FieldList(2).fiOldData <> strNowUser) Then
                        strTo = strTo & ";" & Left(CboEmp.Tag, 5)
                    End If
                End If
                'end 2016/04/07
                'end 2016/03/02
                'Modify by Amy 2021/07/19 +if bolSpecMail = True
                If bolSpecMail = True Then
                    strContent = GetViewData(intOpt1, intOpt2, Left(CboEmp.Tag, 5))
                    strTmp = "因收件人江郁仁請假，請副本收件人處理此郵件。"
                    If InStr(strTo, ";78011") > 0 Then strContent = strContent & vbCrLf & strTmp
                    PUB_SendMail "QPGMR", Mid(strTo, 2), "", Me.Caption & "有更新", strContent, , , , , , , , , , True
                ElseIf strTo <> MsgText(601) Then
                    strContent = GetViewData(intOpt1, intOpt2, Left(CboEmp.Tag, 5))
                    'Modify by Amy 2016/11/18 改收件者顯示個人,cc主管
                    'Modify by Amy 2021/07/16 原:發信者為strNowUser 改為strUserNum 才合理,因職代進入雖strNowUser=區主管編員,但發信應要顯示真實登入者
                    If strCC <> MsgText(601) Then
                        'Modified by Lydia 2019/08/08 frm210152智權點數實績與結餘輸入在承辦人系統更名為「每月點數查詢／輸入」，與財務系統不同
                        'PUB_SendMail strNowUser, Mid(strTo, 2), "", "智權點數實績與結餘有更新", strContent, , , , , , Mid(strCC, 2)
                        PUB_SendMail strUserNum, Mid(strTo, 2), "", Me.Caption & "有更新", strContent, , , , , , Mid(strCC, 2)
                    Else
                        'Modified by Lydia 2019/08/08
                        'PUB_SendMail strNowUser, Mid(strTo, 2), "", "智權點數實績與結餘有更新", strContent
                        PUB_SendMail strUserNum, Mid(strTo, 2), "", Me.Caption & "有更新", strContent
                    End If
                    'end 2016/11/18
                '個人修改發信給部門主管
                ElseIf Left(CboEmp.Tag, 5) = strNowUser And bolGlMan = False And bolAreaMan = False Then
                    strTo = GetDeptMan(stST15)
                    'Modified by Lydia 2019/08/08
                    'strContent = StaffQuery(strNowUser) & "已" & IIf(m_FieldList(49).fiOldData = Empty, "新增", "修改") & "智權點數實績與結餘資料！"
                    strContent = StaffQuery(strNowUser) & "已" & IIf(m_FieldList(49).fiOldData = Empty, "新增", "修改") & Me.Caption & "資料！"
                    If strTo <> MsgText(601) Then
                        'Modify by Amy 2021/07/16 原:發信者為strNowUser 改為strUserNum 才合理,因職代進入雖strNowUser=區主管編員,但發信應要顯示真實登入者
                        PUB_SendMail strUserNum, strTo, "", strContent, "如主旨"
                    End If
                End If
                'Add by Amy 2021/07/16 判斷F11(外商)資料已輸入,發 mail通知 98020(江郁仁協理)
                'Modify by Amy 2021/12/06 +P2005-MCTF(P20部門)
                If (stST15 = "F11" Or stST15 = "P20") And strNowUser <> "98020" Then
                    strTo = "": strTmp = "如主旨"
                    If ChkNoInput("F4106,F4107,P2005", "sp03||sp24") = False Then
                        strContent = Me.Caption & ",外商資料已輸入完成，請至系統查詢並確認"
                        strTo = "98020"
                        If strToSpecNo <> MsgText(601) Then
                            bolSpecMail = True
                            strTo = Mid(strToSpecNo, 2)
                            If InStr(strTo, ";78011") > 0 Then strTmp = "因收件人江郁仁請假，請副本收件人處理此郵件。"
                        End If
                        '若98020請假不發其人事職代
                        PUB_SendMail "QPGMR", strTo, "", strContent, strTmp, , , , , , , , , , bolSpecMail
                    End If
                ElseIf stST15 = "F21" And strNowUser <> GetDeptMan("F21", 1) Then
                    strTo = "": strTmp = "如主旨"
                    If ChkNoInput("F4104,F4105", "sp03||sp24") = False Then
                        strContent = Me.Caption & "," & GetPrjSalesNM("F4104") & "/" & GetPrjSalesNM("F4105") & "資料已輸入完成，請至系統查詢並確認"
                        strTo = GetDeptMan("F21", 1)
                        PUB_SendMail "QPGMR", strTo, "", strContent, "如主旨"
                    End If
                End If
                'end 2021/12/06
                'end 2021/07/16
            'Add by Amy 2016/11/03 +按下主管確認鈕
            Else
                'Modify by Amy 2020/01/07 +bolAccept
                '判斷必需輸入之部門是否都已確認,都已確認發 mail通知財務-瑞婷
                bolAccept = False: strMsg = ""
                bolAccept = ChkAllAccept(strMsg)
                If strMsg <> MsgText(601) Then
                    MsgBox "判斷所有主管是否已確認有誤" & vbCrLf & strMsg
                    Exit Function
                ElseIf bolAccept = True Then
                    strTo = Pub_GetSpecMan("財務處總帳人員")
                    strContent = "智權點數實績與結餘輸入,所有智權部門主管皆已確認"
                    PUB_SendMail "QPGMR", strTo, "", strContent, "如主旨"
                End If
                '判斷S2部門是否都已確認,都已確認發 mail通知 82026(林柄佑經理)
                bolAccept = False: strMsg = ""
                bolAccept = ChkAllAccept(strMsg, 2)
                If strMsg <> MsgText(601) Then
                    MsgBox "判斷「中所」所有主管是否已確認有誤" & vbCrLf & strMsg
                    Exit Function
                ElseIf Mid(stST15, 1, 2) = "S2" And bolAccept = True Then
                    strTo = "82026"
                    'Modified by Lydia 2019/08/08
                    'strContent = "智權點數實績與結餘輸入,中所各區已輸入完成，請至系統查詢"
                    'strContent = "智權點數實績與結輸入,中所各區已輸入完成，請至系統查詢"
                    strContent = Me.Caption & ",中所各區已輸入完成，請至系統查詢"
                    PUB_SendMail "QPGMR", strTo, "", strContent, "如主旨"
                End If
            End If
            'end 2016/02/23
            MsgBox IIf(Index = 0, "資料已存檔！", "主管已確認！")
        End If
    End If
    Exit Function
    
ErrTran:
    If FormSave = False Then cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
End Function

Private Sub Form_Unload(Cancel As Integer)
    strAreaManNo = MsgText(601)
    If UCase(App.EXEName) = "TEACCOUNT" Or UCase(App.EXEName) = "ACCOUNT" Then
        MenuEnabled
    End If
    'Add by Amy 2017/05/19
    If Pub_StrUserSt03 = "M31" Or Pub_StrUserSt03 = "M51" Then
        If bolAxbNotNull = True Then
            MsgBox "此月期末傳票已產生,修改報出點數後," & _
                           "請重新關閉輸入並更正傳票再過帳！"
        End If
    End If
    IsAgentLimit = False 'Add by Amy 2023/02/02 職代
    'Add by Amy 2019/10/16
    stST15 = ""
    strAreaList = ""
    'end 2019/10/16
     strInputEmp = "" 'Add by Amy 2021/11/11
    bolAreaMan = False 'Add by Amy 2021/07/06
    Set frm210152 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
    'Add by Amy 2016/02/23 +勾選報出點數,結餘部分的點選設相同
    If Check1(0).Value = 1 And bolNowChk = False Then
        bolNowChk = True
        Option2(Index).Value = True
        bolNowChk = False
    End If
    'Modify by Amy 2020/02/04 +非S 部門操作
    If strNowUser = Left(CboEmp, 5) Or (Left(strSP48, 1) <> "S" And intLimit = 20) Then
        If Check1(0).Value = 1 And intLimit > 0 Then
            Call SetInputText(98)
        Else
            Call SetInputText(1, Index)
        End If
    End If
    
End Sub

Private Sub Option2_Click(Index As Integer)
    'Add by Amy 2016/02/23 +勾選報出點數,實績部分的點選設相同
    If Check1(0).Value = 1 And bolNowChk = False Then
        bolNowChk = True
        Option1(Index).Value = True
        bolNowChk = False
    End If
    'Modify by Amy 2020/02/04 +非S 部門操作
    If strNowUser = Left(CboEmp, 5) Or (Left(strSP48, 1) <> "S" And intLimit = 20) Then
        If Check1(0).Value = 1 And intLimit > 0 Then
            Call SetInputText(98)
        Else
            Call SetInputText(2, Index)
        End If
    End If
   
End Sub

'intSet:0-全部鎖住/1-未勾選報出點數-實績/2-勾選未報出點數-結餘/98-勾選報出點數-勾實績or結餘用/99-查詢
'idx:Option1、2 index/99其他用途
Private Sub SetInputText(intSet As Integer, Optional ByVal inDx As Integer = 0)
    Dim oText As Object
    'Add by Amy 2017/02/03
    Dim bolRest1Day As Boolean '是否請一整天
    Dim strAbs001 As String, strAbs002 As String, strAbs003 As String
    Dim strDeputy As String '是否為職代/職代人員
    
    If intSet = 0 Then
        '設定所有輸入欄位鎖住
        For Each oText In TxtSP
            If inDx = 99 Or bolIsFirst = True Then
                '灰階
                Call SetTextCSS(oText, 3)
            Else
                '只鎖住
                Call SetTextCSS(oText, 0)
            End If
        Next
        For Each oText In Text5
            If inDx = 99 Or bolIsFirst = True Then
                '灰階
                Call SetTextCSS(oText, 3)
            Else
                '只鎖住
                Call SetTextCSS(oText, 0)
            End If
        Next
        For Each oText In Text4
            If inDx = 99 Or bolIsFirst = True Then
                '灰階
                Call SetTextCSS(oText, 3)
            Else
                '只鎖住
                Call SetTextCSS(oText, 0)
            End If
        Next
        Exit Sub
    End If
    
    If intSet = 99 Then
        bolSave = False
        '財務已過帳或業績輸入已關閉,全部都不可修改
        If Val(Val(Text1(0)) & Text1(1)) <= Val(Left(strA0b01, 5)) Or Val(Val(Text1(0)) & Text1(1)) <= Val(strA0b05) Then
            If strNowUser = Left(CboEmp, 5) Then
                intLimit = 0 '操作個人
            Else
                intLimit = -1 '操作他人
            End If
        '財務或總經理或主祕輸轉撥增減/備註或總經理欄位已輸入/主管已確認/非區主管則其他人不可再修改
        ElseIf ((m_FieldList(21).fiOldData <> Empty And InStr("00,01,08", PUB_GetST05(m_FieldList(21).fiOldData)) > 0) _
          Or (m_FieldList(42).fiOldData <> Empty And InStr("00,01,08", PUB_GetST05(m_FieldList(42).fiOldData)) > 0) _
          Or (m_FieldList(16).fiOldData <> Empty Or m_FieldList(37).fiOldData <> Empty And InStr("00,01,08", stST05) = 0) _
          Or bolIsAccept = True Or (bolAreaMan = True And InStr(strAreaList, Left(cboSalesArea(0), 3)) = 0)) And bolGlMan = False Then
            If strNowUser = Left(CboEmp, 5) Then
                intLimit = 0 '操作個人
            Else
                intLimit = -1 '操作他人
            End If
        'Add by Amy 2021/04/29 非輸入部門者只能看 ex:王副總只能看
        'Modify by Amy 2021/05/13 +bolGlMan = False  bug財務不能改資料
        'Modify by Amy 2022/04/01 +stST15<>"P12",開放雅娟輸P1005個人欄
        ElseIf (InStr(智權點數實績與結餘輸入部門, Left(stST15, 1)) = 0 Or Left(stST15, 2) = "P1") And stST15 <> "P12" And bolGlMan = False Then
            If strNowUser = Left(CboEmp, 5) Then
                intLimit = 0 '操作個人
            Else
                intLimit = -1 '操作他人
            End If
        '**操作個人 欄
        'Modify by Amy 2019/08/01 開放王文安(F21)以個人輸F4102/陳鳳英(F11)以個人輸F4103及其職代A0914,判斷登入者部門為F1或F2
        'Modify by Amy 2019/10/16 開放W部門區主管輸該區編號
        'Modify by Amy 2020/06/16 開放柄佑 82026 輸20091
        'Modify by Amy 2021/04/29 開放林純真(P20)輸P2005
        'Modify by Amy 2021/07/06 +And strNowUser <> "98020" 外商由江協理98020確認輸「區主管」欄,洪琬姿80030輸F4106,May 78011輸F4107輸「個人」欄
        'Modify by Amy 2021/08/03 以W1001 登入輸個人欄(bolAreaMan=False)/W1001區主管(bolAreaMan=True)
        'Modify by Amy 2021/12/06 P2005 由沈佳穎96003輸個人欄,江協理98020確認「區主管」欄
        'Modify by Amy 2022/04/01 開放雅娟輸P1005個人欄
        'Modify by Amy 2022/12/02 因王文安協理退休,F4104改鄭光益(87003)輸/F4105改簡偉倫(99037)輸,顏裕洋(77015 F23部門)確認(F4104/05 F21部門)
        'Modify by Amy 2023/08/07 +W3001 王秀娟經理 輸個人欄位
        ElseIf strNowUser = Left(CboEmp, 5) Or ((Left(stST15, 2) = "F1" Or Left(stST15, 2) = "P2") And strNowUser <> "98020") Or strNowUser = "82026" _
          Or ((Left(stST15, 2) = "W1" Or Left(stST15, 2) = "F2") And bolAreaMan = False) Or (Left(stST15, 2) = "W2" And bolAreaMan = True) Or (Left(stST15, 2) = "W3" And bolAreaMan = True) _
          Or (stST15 = "P12" And bolAreaMan = True) Then
            intLimit = 0
            '總經理輸自己且財務未輸
            If strGlManNo = strNowUser And (m_FieldList(16).fiOldData = Empty Or m_FieldList(37).fiOldData = Empty) Then
                intLimit = 40: bolSave = True
            'Modify by Amy 2016/05/09 修正簡協理(69005)請假周哲丞(82015)以簡協理權限進入操作周哲丞,周哲丞再以自己權限進入看自己Label報出實績點數顯示成2個TextBox且可輸入
            '改先判斷「個人輸自己且主管沒輸過資料or主管未確認才可修改」若主管1和2沒資料才個人才能輸入(即使周哲丞資料的主管1為周哲丞自己仍新增不可改)
            '個人輸自己且主管沒輸過資料or主管未確認才可修改
            ElseIf bolIsAccept = False And Not (m_FieldList(8).fiOldData <> Empty Or m_FieldList(12).fiOldData <> Empty _
                    Or m_FieldList(29).fiOldData <> Empty Or m_FieldList(33).fiOldData <> Empty) Then
                    intLimit = 10: bolSave = True
                'Modify by Amy 2016/05/09 原判斷strNowUser
                '個人且為主管輸自己
                'Modify by Amy 2019/08/01 開放王文安(F21)以個人輸F4102/陳鳳英(F11)以個人輸F4103及其職代A0914,判斷登入者部門為F1或F2
                'Modify by Amy 2019/10/16 開放W部門區主管輸該區編號
                'Modify by Amy 2020/06/16 開放柄佑 82026 輸20091
                'Modify by Amy 2021/01/18 原stSpecEmpNo改為智權點數實績與結餘特殊員編
                'Modify by amy 2021/04/29 開放林純真(P20)輸P2005
                'Modify by Amy 2021/07/06 拿掉Left(stST15, 2) = "F1"  讓洪琬姿80030輸F4106,May 78011輸F4107輸「個人」欄
                'Memo by Amy 2021/12/06 沈佳穎96003(P2部門)輸P2005輸「個人」欄
                'Modify by Amy 2022/04/01 開放雅娟輸P1005個人欄
                If (bolAreaMan = True Or strSt52List <> MsgText(601) Or Left(stST15, 2) = "F2" Or Left(stST15, 2) = "P2" Or Left(stST15, 1) = "W" Or stST15 = "P12" Or (strNowUser = "82026" And InStr(智權點數實績與結餘特殊員編, Left(CboEmp, 5)) > 0)) _
                  And Not ( _
                    (m_FieldList(12).fiOldData = Empty And m_FieldList(33).fiOldData = Empty _
                    And ((m_FieldList(8).fiOldData <> Empty And m_FieldList(8).fiOldData <> strUserNum) _
                    Or (m_FieldList(29).fiOldData <> Empty And m_FieldList(29).fiOldData <> strUserNum))) _
                    Or (m_FieldList(12).fiOldData <> Empty And m_FieldList(12).fiOldData <> strUserNum) _
                    Or (m_FieldList(33).fiOldData <> Empty And m_FieldList(33).fiOldData <> strUserNum) _
                    ) And m_FieldList(16).fiOldData = Empty And m_FieldList(37).fiOldData = Empty Then
                        intLimit = 20
                End If
            End If
            'end 2016/05/09
        '**操作他人
        Else
            intLimit = -1 '其他人只能看
            If bolGlMan = True Then
                '主管確認才能輸資料(因郭少鈞和莊宏宇只能總經理權限輸)
                'If  bolIsAccept = True Then
                    intLimit = 4
                    bolSave = True
                'End If
            '不是輸自己且主管未確認
            ElseIf bolIsAccept = False Then
                'Mark by Amy 2025/02/03 郭少鈞77043 於11401月調至S23,秀玲與杜燕文協理確認其 區主管 詹偉呈98024要可以修改
'                'Modify by Amy 2017/02/03 郭少鈞77043或莊宏宇80010請一整天時,開放職代本人可以區主管身份輸其資料(原只有總經理權限可輸)
'                '郭少鈞與莊宏宇只有總經理及財務和主秘可以輸,區主管只能看
'                If Left(cboEmp, 5) = "77043" And CheckIsPersonRest("77043", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day) = True Then
'                    Call GetABS001_1(Left(cboEmp, 5), strAbs001, strAbs002, strAbs003, True)
'                    '職代一請假則職代二可填,職代二請假職代三可填
'                    If strAbs001 <> MsgText(601) Then
'                        If CheckIsPersonRest(strAbs001, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day) = False And strUserNum = strAbs001 Then
'                            strDeputy = strAbs001
'                        ElseIf strAbs002 <> MsgText(601) Then
'                            If CheckIsPersonRest(strAbs002, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day) = False And strUserNum = strAbs002 Then
'                              strDeputy = strAbs002
'                            ElseIf strAbs003 <> MsgText(601) Then
'                                If CheckIsPersonRest(strAbs003, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day) = False And strUserNum = strAbs003 Then
'                                    strDeputy = strAbs003
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
                'end 2025/02/03
                '莊宏宇80010請一整天時,開放職代本人可以區主管身份輸其資料(原只有總經理權限可輸)
                If Left(CboEmp, 5) = "80010" And CheckIsPersonRest("80010", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day) = True Then
                    Call GetABS001_1(Left(CboEmp, 5), strAbs001, strAbs002, strAbs003, True)
                    '職代一請假則職代二可填,職代二請假職代三可填
                    If strAbs001 <> MsgText(601) Then
                        If CheckIsPersonRest(strAbs001, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day) = False And strUserNum = strAbs001 Then
                            strDeputy = strAbs001
                        ElseIf strAbs002 <> MsgText(601) Then
                            If CheckIsPersonRest(strAbs002, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day) = False And strUserNum = strAbs002 Then
                              strDeputy = strAbs002
                            ElseIf strAbs003 <> MsgText(601) Then
                                If CheckIsPersonRest(strAbs003, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day) = False And strUserNum = strAbs003 Then
                                    strDeputy = strAbs003
                                End If
                            End If
                        End If
                    End If
                End If
                
                'Modify by Amy 2025/02/03 郭少鈞77043 於11401月調至S23,秀玲與杜燕文協理確認其 區主管 詹偉呈98024要可以修改
                'If (Left(cboEmp, 5) = "77043" Or Left(cboEmp, 5) = "80010") And strDeputy = MsgText(601) Then
                If Left(CboEmp, 5) = "80010" And strDeputy = MsgText(601) Then
                'end 2017/02/03
                '區主管只能輸自己區,帶人主管只能輸帶的人
                ElseIf (bolAreaMan = True Or InStr(strAreaList, Left(cboSalesArea(0), 3)) > 0 Or InStr(strSt52List, Left(CboEmp, 5)) > 0) Or strDeputy <> MsgText(601) Then
                    'Modify by Amy 2016/05/09 原判斷strNowUser(周哲丞以簡協理權限新增誠安主管1資料存檔後不應鎖住主管1開放主管2)
                    '主管一調整
                    If ((m_FieldList(8).fiOldData = Empty And m_FieldList(29).fiOldData = Empty) _
                        Or m_FieldList(8).fiOldData = strUserNum Or m_FieldList(29).fiOldData = strUserNum) _
                        And m_FieldList(12).fiOldData = Empty And m_FieldList(33).fiOldData = Empty Then
                        intLimit = 2: bolSave = True
                    '主管二調整
                    ElseIf (m_FieldList(12).fiOldData = Empty And m_FieldList(33).fiOldData = Empty) _
                        Or m_FieldList(12).fiOldData = strUserNum Or m_FieldList(33).fiOldData = strUserNum Then
                        intLimit = 3: bolSave = True
                    End If
                    'end 2016/05/09
                End If
            End If
        End If
        
        
        '設定 動用及期未 欄位顯示方式
'MODIFY BY SONIA 2016/1/26
'        If InStr("00,01,08", stST05) > 0 And intLimit <> 40 Then
'            bolOrgSet = False
'            Call SetChgTxtLbl(bolOrgSet)
'        '主管級以上欄位有值,則個人輸入欄位Text及Label交換顯示
'        ElseIf m_FieldList(8).fiOldData <> Empty Or m_FieldList(29).fiOldData <> Empty Or m_FieldList(12).fiOldData <> Empty Or m_FieldList(33).fiOldData <> Empty _
'          Or m_FieldList(16).fiOldData <> Empty Or m_FieldList(37).fiOldData <> Empty Or m_FieldList(21).fiOldData <> Empty Or m_FieldList(42).fiOldData <> Empty Then
'            bolOrgSet = False
'        ElseIf Left(cboEmp, 5) = strNowUser Then
'            bolOrgSet = True
'        Else
'            bolOrgSet = False
'        End If
        bolOrgSet(1) = True: bolOrgSet(2) = True '是個人設定
        If (TxtSP(7) <> Empty Or TxtSP(11) <> Empty Or TxtSP(15) <> Empty) _
          And (Label31(1) <> Empty Or Label31(11) <> Empty Or Label31(12) <> Empty) Then
            bolOrgSet(1) = False
        End If
        If (TxtSP(28) <> Empty Or TxtSP(32) <> Empty Or TxtSP(36) <> Empty) _
          And (Label41(1) <> Empty Or Label41(11) <> Empty Or Label41(12) <> Empty) Then
            bolOrgSet(2) = False
        End If
'END 2016/1/26
        Call SetInputText(0)
        Call SetChgTxtLbl(bolOrgSet(1), 1)
        Call SetChgTxtLbl(bolOrgSet(2), 2)
        Call SetInputText(1)
        Call SetInputText(2)
        Exit Sub
    End If
    
    ' *** 實績 ***
    If intSet = 1 Then
        Select Case intLimit
            Case -1, 0 'ReadyOnly
                SetChoose (False)
            Case 10 '個人輸自己
                Call SetTextCSS(Text5(0), IIf(Option1(2).Value = True, 2, 3))
                Call SetTextCSS(Text4(0), IIf(Option1(0).Value = True, 2, 3))
                Call SetTextCSS(TxtSP(3), IIf(Option1(1).Value = True, 2, 3))
            Case 2 '主管1
                Call SetTextCSS(TxtSP(7), 2)
            Case 20, 40 '主管1、2輸自己/總經理輸自己
                Call SetTextCSS(Text5(0), IIf(Option1(2).Value = True, 2, 3))
                Call SetTextCSS(Text4(0), IIf(Option1(0).Value = True, 2, 3))
                Call SetTextCSS(TxtSP(3), IIf(Option1(1).Value = True, 2, 3))
            Case 3 '主管2
                Call SetTextCSS(TxtSP(11), 2)
            Case 4 '總經理
                Call SetTextCSS(TxtSP(15), 2)
        End Select
        
        '區主管才開放輸入轉撥
        If intLimit > 0 And (bolAreaMan = True Or bolGlMan = True) Then
            Call SetTextCSS(TxtSP(19), 2)
            Call SetTextCSS(TxtSP(20), 2)
        End If
        Exit Sub
    End If
    
    '*** 結餘 ***
    If intSet = 2 Then
        Select Case intLimit
            Case 10 '個人輸自己
                Call SetTextCSS(Text5(1), IIf(Option2(2).Value = True, 2, 3))
                Call SetTextCSS(Text4(1), IIf(Option2(0).Value = True, 2, 3))
                Call SetTextCSS(TxtSP(24), IIf(Option2(1).Value = True, 2, 3))
            Case 2 '主管1
                Call SetTextCSS(TxtSP(28), 2)
            Case 20, 40 '主管1、2輸自己/總經理輸自己
                '點「結餘」
                Call SetTextCSS(Text5(1), IIf(Option2(2).Value = True, 2, 3))
                Call SetTextCSS(Text4(1), IIf(Option2(0).Value = True, 2, 3))
                Call SetTextCSS(TxtSP(24), IIf(Option2(1).Value = True, 2, 3))
            Case 3 '主管2
                Call SetTextCSS(TxtSP(32), 2)
            Case 4 '總經理
                Call SetTextCSS(TxtSP(36), 2)
        End Select
        
        '區主管才開放輸入轉撥
        If intLimit > 0 And (bolAreaMan = True Or bolGlMan = True) Then
            Call SetTextCSS(TxtSP(40), 2)
            Call SetTextCSS(TxtSP(41), 2)
        End If
        Exit Sub
    End If
    
    '*** 勾選設定 ***
    If intSet = 98 Then
        If Check1(0).Value = 1 Then
            Call SetTextCSS(Text4(0), 0)
            Call SetTextCSS(TxtSP(3), 0)
            Call SetTextCSS(TxtSP(7), 0)
            Call SetTextCSS(TxtSP(11), 0)
            Call SetTextCSS(TxtSP(15), 0)
            Call SetTextCSS(TxtSP(19), 0)
            Call SetTextCSS(TxtSP(20), 0)

            Call SetTextCSS(Text4(1), 0)
            Call SetTextCSS(TxtSP(24), 0)
            Call SetTextCSS(TxtSP(28), 0)
            Call SetTextCSS(TxtSP(32), 0)
            Call SetTextCSS(TxtSP(36), 0)
            Call SetTextCSS(TxtSP(40), 0)
            Call SetTextCSS(TxtSP(41), 0)

            '實績
            If Check1(1).Value = 1 Then
                Check1(2) = 0
                Select Case intLimit
                    Case 10
                        Call SetTextCSS(Text5(1), 3)
                        Call SetTextCSS(Text5(0), IIf(Option1(2).Value = True, 2, 3))
                        Call SetTextCSS(Text4(0), IIf(Option1(0).Value = True, 2, 3))
                        Call SetTextCSS(TxtSP(3), IIf(Option1(1).Value = True, 2, 3))
                    Case 2
                        Call SetTextCSS(TxtSP(7), 2)
                    Case 20, 40
                        Call SetTextCSS(Text5(1), 3)
                        Call SetTextCSS(Text5(0), IIf(Option1(2).Value = True, 2, 3))
                        Call SetTextCSS(Text4(0), IIf(Option1(0).Value = True, 2, 3))
                        Call SetTextCSS(TxtSP(3), IIf(Option1(1).Value = True, 2, 3))
                    Case 3
                        Call SetTextCSS(TxtSP(11), 2)
                    Case 4
                        Call SetTextCSS(TxtSP(15), 2)
                End Select
                
                '區主管/總經理權限才開放輸入轉撥
                If intLimit > 0 And (bolAreaMan = True Or bolGlMan = True) Then
                    Call SetTextCSS(TxtSP(19), 2)
                    Call SetTextCSS(TxtSP(20), 2)
                End If
            End If

           '結餘
            If Check1(2).Value = 1 Then
                Check1(1) = 0
                Select Case intLimit
                    Case 10
                        Call SetTextCSS(Text5(0), 3)
                        Call SetTextCSS(Text5(1), IIf(Option2(2).Value = True, 2, 3))
                        Call SetTextCSS(Text4(1), IIf(Option2(0).Value = True, 2, 3))
                        Call SetTextCSS(TxtSP(24), IIf(Option2(1).Value = True, 2, 3))
                    Case 2
                        Call SetTextCSS(TxtSP(28), 2)
                    Case 20, 40
                        Call SetTextCSS(Text5(0), 3)
                        Call SetTextCSS(Text5(1), IIf(Option2(2).Value = True, 2, 3))
                        Call SetTextCSS(Text4(1), IIf(Option2(0).Value = True, 2, 3))
                        Call SetTextCSS(TxtSP(24), IIf(Option2(1).Value = True, 2, 3))
                    Case 3
                        Call SetTextCSS(TxtSP(32), 2)
                    Case 4
                        Call SetTextCSS(TxtSP(36), 2)
                End Select
                 
                '區主管/總經理權限才開放輸入轉撥
                If intLimit > 0 And (bolAreaMan = True Or bolGlMan = True) Then
                    Call SetTextCSS(TxtSP(40), 2)
                    Call SetTextCSS(TxtSP(41), 2)
                End If

            End If
        End If
        Exit Sub
    End If
End Sub

'Modifed by Lydia 2022/01/03 objTxt As TextBox=>objTxt As object
Private Sub SetTextCSS(objTxt As Object, ByVal intChoose As Integer, Optional ByVal intStyle As Integer = 1)
    Select Case intChoose
        Case 0 'Locked
            objTxt.Locked = True
            If UCase(objTxt.Name) = "TEXT5" And (objTxt.Index = 0 Or objTxt.Index = 1) Then
                '實績
                If objTxt.Index = 0 Then
                    objTxt.BorderStyle = IIf(bolOrgSet(1) = True, 1, 0)
                    objTxt.BackColor = IIf(bolOrgSet(1) = True, QBColor(7), SetColor)
                    '                            主管用Label       是個人         False-不設color            是個人              0-無框
                    Call SetLabelCSS(Label31(1), IIf(bolOrgSet(1) = True, False, True), IIf(bolOrgSet(1) = True, 0, 1))
                '結餘
                Else
                    objTxt.BorderStyle = IIf(bolOrgSet(2) = True, 1, 0)
                    objTxt.BackColor = IIf(bolOrgSet(2) = True, QBColor(7), SetColor)
                    Call SetLabelCSS(Label41(1), IIf(bolOrgSet(2) = True, False, True), IIf(bolOrgSet(2) = True, 0, 1))
                End If
            ElseIf UCase(objTxt.Name) = "TEXT4" And (objTxt.Index = 0 Or objTxt.Index = 1) Then
                If objTxt.Index = 0 Then
                    objTxt.BorderStyle = IIf(bolOrgSet(1) = True, 1, 0)
                    objTxt.BackColor = IIf(bolOrgSet(1) = True, QBColor(7), SetColor)
                    Call SetLabelCSS(Label31(11), IIf(bolOrgSet(1) = True, False, True), IIf(bolOrgSet(1) = True, 0, 1))
                Else
                    objTxt.BorderStyle = IIf(bolOrgSet(2) = True, 1, 0)
                    objTxt.BackColor = IIf(bolOrgSet(2) = True, QBColor(7), SetColor)
                    Call SetLabelCSS(Label41(11), IIf(bolOrgSet(2) = True, False, True), IIf(bolOrgSet(2) = True, 0, 1))
                End If
            ElseIf UCase(objTxt.Name) = "TXTSP" And (objTxt.Index = 3 Or objTxt.Index = 24) Then
                If objTxt.Index = 3 Then
                    'Modify by Amy 2022/01/04 改Form2.0 後設定BorderStyle 無效,改設BackStyle及SpecialEffect
                    'objTxt.BorderStyle = IIf(bolOrgSet(1) = True, 1, 0)
                    objTxt.BackStyle = IIf(bolOrgSet(1) = True, 1, 0)
                    objTxt.SpecialEffect = IIf(bolOrgSet(1) = True, 2, 0)
                    'end 2022/01/04
                    objTxt.BackColor = IIf(bolOrgSet(1) = True, QBColor(7), SetColor)
                    Call SetLabelCSS(Label31(12), IIf(bolOrgSet(1) = True, False, True), IIf(bolOrgSet(1) = True, 0, 1))
                Else
                    'Modify by Amy 2022/01/04 改Form2.0 後設定BorderStyle 無效,改設BackStyle及SpecialEffect
                    'objTxt.BorderStyle = IIf(bolOrgSet(2) = True, 1, 0)
                    objTxt.BackStyle = IIf(bolOrgSet(1) = True, 1, 0)
                    objTxt.SpecialEffect = IIf(bolOrgSet(1) = True, 2, 0)
                    'end 2022/01/04
                    objTxt.BackColor = IIf(bolOrgSet(2) = True, QBColor(7), SetColor)
                    Call SetLabelCSS(Label41(12), IIf(bolOrgSet(2) = True, False, True), IIf(bolOrgSet(2) = True, 0, 1))
                End If
            Else
                objTxt.BackColor = QBColor(7) '灰色
            End If
        Case 2 '原樣-可輸
            objTxt.Locked = False
            objTxt.BackColor = QBColor(15) '白色
        Case 3 '原樣-灰階Locked
            objTxt.Locked = True
            objTxt.BackColor = QBColor(7) '灰色
        Case 4 '仿Label
            objTxt.BorderStyle = 0
            objTxt.BackColor = SetColor
    End Select
End Sub

Private Sub SetLabelCSS(objLbl As LABEL, bolSetColor As Boolean, intStyle As Integer)
    objLbl.BorderStyle = intStyle '0-無框/1-有框
    If bolSetColor = True Then
        objLbl.BackColor = QBColor(7) '灰色
    Else
        objLbl.BackColor = SetColor
    End If
End Sub

Private Sub SetChgTxtLbl(bolReset As Boolean, intChoose As Integer)
    '實績
    If intChoose = 1 Then
        If bolReset = True Then
            '實績動用/期末實績 位置還原個人輸入位置
            Text5(0).Left = 2320: Text5(0).Width = 1000
            Label31(1).Left = 3390: Label31(1).Width = 900
            Text4(0).Left = 1480
            Label31(11).Left = 2520
            TxtSP(3).Left = 1480
            Label31(12).Left = 2520
        Else
            Text5(0).Left = 3390: Text5(0).Width = 900
            Label31(1).Left = 2320: Label31(1).Width = 1000
            Text4(0).Left = 2520
            Label31(11).Left = 1480
            TxtSP(3).Left = 2520
            Label31(12).Left = 1480
        End If
        Exit Sub
    End If
    '結餘
    If intChoose = 2 Then
        If bolReset = True Then
            '結餘動用/期末結餘 位置還原個人輸入位置
            Text5(1).Left = 2320: Text5(0).Width = 1000
            Label41(1).Left = 3390: Label41(1).Width = 900
            Text4(1).Left = 1480
            Label41(11).Left = 2520
            TxtSP(24).Left = 1480
            Label41(12).Left = 2520
        Else
            Text5(1).Left = 3390: Text5(0).Width = 900
            Label41(1).Left = 2320: Label41(1).Width = 1000
            Text4(1).Left = 2520
            Label41(11).Left = 1480
            TxtSP(24).Left = 2520
            Label41(12).Left = 1480
        End If
        Exit Sub
    End If
End Sub

Private Sub RunSum(Optional ByVal oTextN As String = "", Optional ByVal idx As Integer, Optional ByVal bolFirst As Boolean = False)
    Dim dblFinal(1) As Double, dblTot(1) As Double, dblTemp(1) As Double
    Dim bolPerson(1) As Boolean
    Dim intTxtIdx As Integer, intOptIdx As Integer
    Dim bolText2 As Boolean, bolText5 As Boolean, bolE As Boolean, bolJ As Boolean 'For Text2/Text5/轉撥E/轉撥J  使用
    
    If oTextN <> MsgText(601) Then
        Select Case oTextN
            Case "TEXT2"
                '於最下方計算
                bolText2 = True
            Case "TEXT4"
                Select Case idx
                    Case 0
                        TxtSP(3) = Round(Val(Label31(0)) - Val(Text4(0)), 3)
                        Text5(0) = Round(Val(Label31(0)) + Val(Label31(6)) - Val(TxtSP(3)) + Val(TxtSP(19)), 3)
                    Case 1
                        TxtSP(24) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(Text4(1)), 3)
                        Text5(1) = Round(Val(Label41(0)) + Val(Label41(6)) - dblFinal(1) + Val(TxtSP(40)), 3)
                End Select
            Case "TEXT5"
                bolText5 = True
                Select Case idx
                    Case 0
                        TxtSP(3) = Round(Val(Label31(0)) + Val(Label31(6)) + Val(TxtSP(19)) - Val(Text5(0)), 3)
                        Text4(0) = Round(Val(Label31(0)) - Val(TxtSP(3)), 3)
                    Case 1
                        TxtSP(24) = Round(Val(Label41(0)) + Val(Label41(6)) + Val(TxtSP(40)) - Val(Text5(1)), 3)
                        Text4(1) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(TxtSP(24)), 3)
                End Select
            Case "TXTSP"
                Select Case idx
                    Case 3
                        '沒加Round導致690-693.9會出現-3.89999999999998
                        Text4(0) = Round(Val(Label31(0)) - Val(TxtSP(idx)), 3)
                        Text5(0) = Round(Val(Label31(0)) + Val(Label31(6)) - Val(TxtSP(3)) + Val(TxtSP(19)), 3)
                    Case 7, 11, 15
                        If Option1(0).Value = True Then
                            Label31(11) = TxtSP(idx)
                            Label31(12) = Round(Val(Label31(0)) - Val(Label31(11)), 3)
                        ElseIf Option1(1).Value = True Then
                            Label31(12) = TxtSP(idx)
                            Label31(11) = Round(Val(Label31(0)) - Val(Label31(12)), 3)
                        Else
                            Label31(1) = TxtSP(idx)
                            Label31(12) = Round(Val(Label31(0)) + Val(Label31(6)) + Val(TxtSP(19)) - Val(Label31(1)), 3)
                            Label31(11) = Round(Val(Label31(0)) - Val(Label31(12)), 3)
                        End If
                    Case 24
                        Text4(1) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(TxtSP(idx)), 3)
                        Text5(1) = Round(Val(Label41(0)) + Val(Label41(6)) - dblFinal(1) + Val(TxtSP(40)), 3)
                    Case 28, 32, 36
                        If Option2(0).Value = True Then
                            Label41(11) = TxtSP(idx)
                            Label41(12) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(Label41(11)), 3)
                        ElseIf Option2(1).Value = True Then
                            Label41(12) = TxtSP(idx)
                            Label41(11) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(Label41(12)), 3)
                        Else
                            Label41(1) = TxtSP(idx)
                            Label41(12) = Round(Val(Label41(0)) + Val(Label41(6)) + Val(TxtSP(40)) - Val(Label41(1)), 3)
                            Label41(11) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(Label41(12)), 3)
                        End If
                    Case 19
                        bolE = True
                    Case 40
                        bolJ = True
                End Select
            Case Else
        End Select
    End If
      
    '計算:總經理>主管二>主管一>個人
    '實績
    'Modify by Amy 2016/12/05 用Val(m_FieldList(xx).fiOldData) <> 0 導致資料若=0,某些值算錯 exp:10511 P1001 期末結餘=0
    If m_FieldList(15).fiOldData <> Empty Or Val(TxtSP(15)) <> 0 Then
        dblFinal(0) = Val(Label31(12))
    ElseIf m_FieldList(11).fiOldData <> Empty Or Val(TxtSP(11)) <> 0 Then
        dblFinal(0) = Val(Label31(12))
    ElseIf m_FieldList(7).fiOldData <> Empty Or Val(TxtSP(7)) <> 0 Then
        dblFinal(0) = Val(Label31(12))
    'Add by Amy 2021/12/01 尚未輸入
    ElseIf m_FieldList(15).fiOldData = Empty And m_FieldList(11).fiOldData = Empty And m_FieldList(7).fiOldData = Empty And m_FieldList(3).fiOldData = Empty And m_FieldList(19).fiOldData = Empty _
              And TxtSP(15) = Empty And TxtSP(11) = Empty And TxtSP(7) = Empty And TxtSP(3) = Empty And TxtSP(19) = Empty Then
        'Modified by Lydia 2022/01/03 無值會出錯; ex. S152北五目標
        'dblFinal(0) = Val(Round(Label31(0), 2)): bolPerson(0) = True
        dblFinal(0) = Val(Round(Val(Label31(0).Caption), 2)): bolPerson(0) = True
    Else
        dblFinal(0) = Val(TxtSP(3)): bolPerson(0) = True
    End If
    
    '結餘
    If m_FieldList(36).fiOldData <> Empty Or Val(TxtSP(36)) <> 0 Then
        dblFinal(1) = Val(Label41(12))
    ElseIf m_FieldList(32).fiOldData <> Empty Or Val(TxtSP(32)) <> 0 Then
        dblFinal(1) = Val(Label41(12))
    ElseIf m_FieldList(28).fiOldData <> Empty Or Val(TxtSP(28)) <> 0 Then
        dblFinal(1) = Val(Label41(12))
    'Add by Amy 2021/12/01
    ElseIf m_FieldList(28).fiOldData = Empty And m_FieldList(32).fiOldData = Empty And m_FieldList(36).fiOldData = Empty And m_FieldList(24).fiOldData = Empty And m_FieldList(20).fiOldData = Empty _
              And TxtSP(28) = Empty And TxtSP(32) = Empty And TxtSP(36) = Empty And TxtSP(24) = Empty And TxtSP(20) = Empty Then
        'Modified by Lydia 2022/01/03 無值會出錯; ex. S152北五目標
        'dblFinal(1) = Val(Round(Label41(0), 2)): bolPerson(1) = True
        dblFinal(1) = Val(Round(Val(Label41(0).Caption), 2)): bolPerson(1) = True
    Else
        dblFinal(1) = Val(TxtSP(24)): bolPerson(1) = True
    End If
    'end 2016/12/05
    
    '未勾選報出點數
    If Check1(0).Value = 0 Then
        '報出實績點數=期初+當月-期末+轉撥
        dblTot(0) = Round(Val(Label31(0)) + Val(Label31(6)) - dblFinal(0) + Val(TxtSP(19)), 3)
        If m_FieldList(15).fiOldData <> Empty Or TxtSP(15) <> Empty Or m_FieldList(11).fiOldData <> Empty Or TxtSP(11) <> Empty _
          Or m_FieldList(7).fiOldData <> Empty Or TxtSP(7) <> Empty Then
            Text5(0) = Round(Val(Label31(0)) + Val(Label31(6)) - Val(TxtSP(3)) + Val(TxtSP(19)), 3)
            Label31(1) = dblTot(0)
        Else
            Text5(0) = dblTot(0)
        End If
        '報出結餘點數=期初+當月-期末+轉撥
        dblTot(1) = Round(Val(Label41(0)) + Val(Label41(6)) - dblFinal(1) + Val(TxtSP(40)), 3)
        If m_FieldList(36).fiOldData <> Empty Or TxtSP(36) <> Empty Or m_FieldList(32).fiOldData <> Empty Or TxtSP(32) <> Empty _
          Or m_FieldList(28).fiOldData <> Empty Or TxtSP(28) <> Empty Then
            Text5(1) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(TxtSP(24)) + Val(TxtSP(40)), 3)
            Label41(1) = dblTot(1)
            'Label41(11) = dblTot(1) 'Mark by Amy 2024/07/10 bug-存檔後會顯示0
        Else
            Text5(1) = dblTot(1)
        End If
        '未勾選「報出點數」計算(報出點數 鎖住)
        Text2 = Round(Val(dblTot(0)) + Val(dblTot(1)), 3)
         If bolFirst = True Then
             If m_FieldList(4).fiOldData = Empty Then
                Text4(0) = "0"
                Text4(1) = "0"
                '未輸入過 sp3資料需計算
                If TxtSP(3) = MsgText(601) Then TxtSP(3) = Round(Val(Label31(0)) + Val(Text4(0)), 3)
                '未輸入過 sp24資料需計算
                If TxtSP(24) = MsgText(601) Then TxtSP(24) = Round(Val(Label41(0)) + Val(Label41(6)) + Val(Text4(1)), 3)
            Else
                If TxtSP(3) = MsgText(601) Then TxtSP(3) = "0"
                If TxtSP(24) = MsgText(601) Then TxtSP(24) = "0"
                If Text4(0) = MsgText(601) Then Text4(0) = Round(Val(Label31(0)) - Val(TxtSP(3)), 3)
                If Text4(1) = MsgText(601) Then Text4(1) = (Round(Val(Label41(0)) + Val(Label41(6)) - Val(TxtSP(24)), 3))
            End If
            If bolPerson(0) = False Then
                Label31(11) = Round(Val(Label31(0)) - Val(Label31(12)), 3)
                Label31(12) = Round(Val(Label31(0)) - Val(Label31(11)), 3)
            End If
            If bolPerson(1) = False Then
                Label41(11) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(Label41(12)), 3)
                Label41(12) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(Label41(11)), 3)
            End If
         End If
    '勾選報出點數
    Else
        If (Check1(1).Value = 1 And Check1(2).Value = 0) Or (Check1(1).Value = 1 And Check1(2).Value = 0) Then
            '記錄目前輸入項目
            intTxtIdx = 0
            If Check1(1).Value = 1 And Check1(2).Value = 0 Then intTxtIdx = 1
            If m_FieldList(15).fiOldData <> Empty Or TxtSP(15) <> Empty Or m_FieldList(36).fiOldData <> Empty Or TxtSP(36) <> Empty Then
                intTxtIdx = 15
                If Check1(1).Value = 1 And Check1(2).Value = 0 Then intTxtIdx = 36
            ElseIf m_FieldList(11).fiOldData <> Empty Or TxtSP(11) <> Empty Or m_FieldList(32).fiOldData <> Empty Or TxtSP(32) <> Empty Then
                intTxtIdx = 11
                If Check1(1).Value = 1 And Check1(2).Value = 0 Then intTxtIdx = 32
            ElseIf m_FieldList(7).fiOldData <> Empty Or TxtSP(7) <> Empty Or m_FieldList(28).fiOldData <> Empty Or TxtSP(28) <> Empty Then
                intTxtIdx = 7
                If Check1(1).Value = 1 And Check1(2).Value = 0 Then intTxtIdx = 28
            ElseIf m_FieldList(3).fiOldData <> Empty Or TxtSP(3) <> Empty Or m_FieldList(24).fiOldData <> Empty Or TxtSP(24) <> Empty Then
                intTxtIdx = 3
                If Check1(1).Value = 1 And Check1(2).Value = 0 Then intTxtIdx = 24
            End If
        End If
        
        '** 勾選實績 **
        If Check1(1).Value = 1 And Check1(2).Value = 0 Then
            '實績主管欄未輸
            If m_FieldList(15).fiOldData = Empty And TxtSP(15) = Empty And m_FieldList(11).fiOldData = Empty And TxtSP(11) = Empty _
              And m_FieldList(7).fiOldData = Empty And TxtSP(7) = Empty Then
                Label31(1) = MsgText(601)
                Label31(11) = MsgText(601)
                Label31(12) = MsgText(601)
            End If
            If bolText2 = True Then
                '更新個人報出實績點數(若主管輸轉撥)
                If Label31(1) <> "" And intLimit > 0 And intLimit < 10 Then
                     Text5(0) = Round(Val(Label31(0)) + Val(Label31(6)) - Val(TxtSP(3)) + Val(TxtSP(19)), 3)
                End If
                dblTot(0) = IIf(Label31(1) <> "", Val(Label31(1)), Val(Text5(0)))
            ElseIf Not (bolText5 = True Or Option1(2).Value = True) Or bolE = True Then
                dblTot(0) = Round(Val(Label31(0)) + Val(Label31(6)) - dblFinal(0) + Val(TxtSP(19)), 3)
            End If
            If Not (bolText5 = True Or Option1(2).Value = True) Or bolE = True Then
                If m_FieldList(15).fiOldData <> Empty Or TxtSP(15) <> Empty Or m_FieldList(11).fiOldData <> Empty Or TxtSP(11) <> Empty _
                  Or m_FieldList(7).fiOldData <> Empty Or TxtSP(7) <> Empty Then
                    Text5(0) = Round(Val(Label31(0)) + Val(Label31(6)) - Val(TxtSP(3)) + Val(TxtSP(19)), 3)
                    Label31(1) = dblTot(0)
                Else
                    If bolText5 = False Then Text5(0) = dblTot(0)
                End If
            End If
            
            intOptIdx = 1
            If Option1(0).Value = True Then
                intOptIdx = 0
            ElseIf Option1(2).Value = True Then
                intOptIdx = 2
            End If
            Option2(intOptIdx).Value = True
            
            '輸入報出點數重新計算報出結餘點數
            '報出結餘點數
            dblTot(1) = Round(Val(Text2) - IIf(Label31(1) <> "", Val(Label31(1)), Val(Text5(0))), 3)
            '期末結餘
            dblTemp(0) = Round(Val(Label41(0)) + Val(Label41(6)) + Val(TxtSP(40)) - Val(dblTot(1)), 3)
            '結餘動用
            dblTemp(1) = Round(Val(Label41(0)) + Val(Label41(6)) - dblTemp(0), 3)
           
            '結餘主管欄未輸
            If m_FieldList(36).fiOldData <> Empty Or TxtSP(36) <> Empty Or m_FieldList(32).fiOldData <> Empty Or TxtSP(32) <> Empty _
              Or m_FieldList(28).fiOldData <> Empty Or TxtSP(28) <> Empty Then
                Label41(1) = dblTot(1)       '報出結餘點數
                Label41(11) = dblTemp(1) '結餘動用
                Label41(12) = dblTemp(0) '期末結餘
                Select Case intOptIdx
                    Case 1
                        dblTemp(1) = Label41(12)
                    Case 2
                        dblTemp(1) = Label41(1)
                End Select
                If oTextN <> MsgText(601) And idx <> 19 And intTxtIdx <> 0 Then TxtSP(intTxtIdx) = dblTemp(1)
                If bolOrgSet(2) = True Then
                    bolOrgSet(2) = False
                    Call SetChgTxtLbl(bolOrgSet(2), 2)
                End If
            Else
                Text5(1) = dblTot(1)
                Text4(1) = dblTemp(1)
                TxtSP(24) = dblTemp(0)
            End If
        End If
        
        
        '** 勾選結餘 **
        If Check1(1).Value = 0 And Check1(2).Value = 1 Then
            If m_FieldList(36).fiOldData = Empty And TxtSP(36) = Empty And m_FieldList(32).fiOldData = Empty And TxtSP(32) = Empty _
              And m_FieldList(28).fiOldData = Empty And TxtSP(28) = Empty Then
                Label41(1) = MsgText(601)
                Label41(11) = MsgText(601)
                Label41(12) = MsgText(601)
            End If
            '報出結餘點數
            If bolText2 = True Then
                '更新個人報出結餘點數(若主管輸轉撥)
                If Label41(1) <> "" And intLimit > 0 And intLimit < 10 Then
                     Text5(1) = Round(Val(Label41(0)) + Val(Label41(6)) - dblFinal(1) + Val(TxtSP(40)), 3)
                End If
                dblTot(1) = IIf(Label41(0) <> "", Val(Text5(1)), Val(Label41(1)))
            ElseIf Not (bolText5 = True Or Option2(2).Value = True) Or bolJ = True Then
                dblTot(1) = Round(Val(Label41(0)) + Val(Label41(6)) - dblFinal(1) + Val(TxtSP(40)), 3)
            End If
            If Not (bolText5 = True Or Option2(2).Value = True) Or bolJ = True Then
                If m_FieldList(36).fiOldData <> Empty Or TxtSP(36) <> Empty Or m_FieldList(32).fiOldData <> Empty Or TxtSP(32) <> Empty _
                  Or m_FieldList(28).fiOldData <> Empty Or TxtSP(28) <> Empty Then
                    Text5(1) = Round(Val(Label41(0)) + Val(Label41(6)) - Val(TxtSP(24)) + Val(TxtSP(40)), 3)
                    Label41(1) = dblTot(1)
                Else
                    If bolText5 = False Then Text5(1) = dblTot(1)
                End If
            End If
            
             intOptIdx = 1
            If Option2(0).Value = True Then
                intOptIdx = 0
            ElseIf Option2(2).Value = True Then
                intOptIdx = 2
            End If
            Option1(intOptIdx).Value = True
            
            '輸入報出點數重新計算報出實績點數
            '報出實績
            dblTot(0) = Round(Val(Text2) - IIf(Label41(1) <> "", Val(Label41(1)), Val(Text5(1))), 3)
            '期末實績
            dblTemp(0) = Round(Val(Label31(0)) + Val(Label31(6)) + Val(TxtSP(19)) - Val(dblTot(0)), 3)
            '實績動用
            dblTemp(1) = Round(Val(Label31(0)) - dblTemp(0), 3)
            
            If m_FieldList(15).fiOldData <> Empty Or TxtSP(15) <> Empty Or m_FieldList(11).fiOldData <> Empty Or TxtSP(11) <> Empty _
              Or m_FieldList(7).fiOldData <> Empty Or TxtSP(7) <> Empty Then
                Label31(1) = dblTot(0)       '報出實績點數
                Label31(11) = dblTemp(1) '實績動用
                Label31(12) = dblTemp(0) '期末實績
                Select Case intOptIdx
                    Case 1
                        dblTemp(1) = Label31(12)
                    Case 2
                        dblTemp(1) = Label31(1)
                End Select
                If oTextN <> MsgText(601) And idx <> 40 And intTxtIdx <> 0 Then TxtSP(intTxtIdx) = dblTemp(1)
                
                If bolOrgSet(1) = True Then
                    bolOrgSet(1) = False
                    Call SetChgTxtLbl(bolOrgSet(1), 1)
                End If
            Else
                Text5(0) = dblTot(0)
                Text4(0) = dblTemp(1)
                TxtSP(3) = dblTemp(0)
            End If
        End If
    End If
    
    '達成率
    If Val(Label21(0)) > 0 Then
        Label21(1) = Format(Round(Val((Text2)) / Val(Label21(0)), 3), "0.000%")
    End If
End Sub

Private Sub tabSP_Click(PreviousTab As Integer)
    '點「全區資料」
    If PreviousTab = 0 Then
        If cboSalesArea(0) <> MsgText(601) Then
            'Modify by Amy 2016/11/18
            'Mark by Amy 2020/06/16 開放柄佑 輸20091(S29),故登入時設st15=S29
'            If strNowUser = "82026" And Left(cboSalesArea(0), 3) = stST15 Then
'                '林柄佑經理不可看自己的全區
'            Else
                cboSalesArea(1) = cboSalesArea(0)
'            End If
            'end 2020/06/16
            'end 2016/11/18
        End If
        If Trim(Text1(0) & Text1(1)) <> MsgText(601) And Val(Text1(0) & Text1(1)) <> Val(Text11(0) & Text11(1)) Then
            Text11(0) = Text1(0)
            Text11(1) = Text1(1)
        End If
        If cboSalesArea(1) <> MsgText(601) And Trim(Text11(0) & Text11(1)) <> MsgText(601) Then cmdSearch_Click (1)
    '點「個人資料」
    Else
        '於「全區資料」回前頁籤時若部門或日期年月與前頁不同時,設與全區相同,人員依部門別重抓但不預設
        If cboSalesArea(1) <> MsgText(601) And cboSalesArea(1) <> cboSalesArea(0) _
          Or Trim(Text11(0) & Text11(1)) <> MsgText(601) And Val(Text11(0) & Text11(1)) <> Val(Text1(0) & Text1(1)) Then
            If Trim(Text11(0) & Text11(1)) <> MsgText(601) And Val(Text11(0) & Text11(1)) <> Val(Text1(0) & Text1(1)) Then
                Text1(0) = IIf(Text11(0).Tag <> MsgText(601), Mid(Text11(0).Tag, 1, Len(Text11(0).Tag) - 2), Text11(0))
                Text1(1) = IIf(Text11(0).Tag <> MsgText(601), Right(Text11(0).Tag, 2), Text11(1))
            End If
            If cboSalesArea(1) <> cboSalesArea(0) Then
                Call FormClear(1)
                cboSalesArea(0) = cboSalesArea(1)
                Call SetEmp
                Call SetInputText(0, 99)
            End If
        End If
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
     TextInverse Text1(Index)
End Sub

'Modify by Amy 2021/11/11 重新整理
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    If Text1(0) = MsgText(601) Or Text1(1) = MsgText(601) Or bolNowChk = True Then Exit Sub
    
    Text1(1) = Format(Text1(1), "00")
    
    If ChkDate(Val(Text1(0)) & Text1(1) & "01") = False Then
        Text1_GotFocus Index
        Cancel = True
        bolNowChk = False
        Exit Sub
    End If
    If Val(Text1(0)) & Text1(1) < Val(業績輸入啟用年月) Then
        MsgBox "業績輸入上線日為" & Left(Val(業績輸入啟用年月), 3) & "年" & Right(Val(業績輸入啟用年月), 2) & _
                    "故系統無資料！", , MsgText(5)
        Text1_GotFocus Index
        Cancel = True
        bolNowChk = False
        Exit Sub
    End If
    '原判斷系統日,改判斷SP01最大值
    If Val(Text1(0)) & Text1(1) > Val(strMaxSP01) Then
        MsgBox "不可輸入尚未開放的業績日期！", , MsgText(5)
        Text1_GotFocus Index
        Cancel = True
        bolNowChk = False
        Exit Sub
    End If
    
    '修改業績年月
    If Val(Text1(0).Tag) <> Val(Text1(0)) & Text1(1) And Trim(Text1(0)) <> MsgText(601) And Trim(Text1(1)) <> MsgText(601) Then
        Call FormClear(1)
        Call SpecData '特殊轉區人員
        '自動查詢(避免無窮迴圈及有修改才執行)
        If Trim(cboSalesArea(0)) <> MsgText(601) And Trim(CboEmp) <> MsgText(601) And Val(Text1(0).Tag) <> Val(Trim(Text1(0) & Text1(1))) And Val(Trim(Text1(0) & Text1(1))) <> 0 Then
            bolNoChkMod = True
            Call cmdSearch_Click(0)
            bolNoChkMod = True = False
        End If
    End If
    Text1(0).Tag = Val(Text1(0)) & Text1(1)
    bolNowChk = False
End Sub

Private Sub Text1_Validate_OLD(Index As Integer, Cancel As Boolean)
    'Mark by Amy 2021/11/11 因重新整理SetEmp會影響此,故也重新整理
'    If Text1(0) = MsgText(601) Or Text1(1) = MsgText(601) Or bolNowChk = True Then Exit Sub
'
'    'Modfiy by Amy 2016/02/23
'    bolNowChk = True
'    Text1(1) = Format(Text1(1), "00")
'
'    If ChkDate(Val(Text1(0)) & Text1(1) & "01") = False Then
'        Text1_GotFocus Index
'        Cancel = True
'        bolNowChk = False
'        Exit Sub
'    End If
'    If Val(Text1(0)) & Text1(1) < Val(業績輸入啟用年月) Then
'        MsgBox "業績輸入上線日為" & Left(Val(業績輸入啟用年月), 3) & "年" & Right(Val(業績輸入啟用年月), 2) & _
'                    "故系統無資料！", , MsgText(5)
'        Text1_GotFocus Index
'        Cancel = True
'        bolNowChk = False
'        Exit Sub
'    End If
'    'Modify by Amy 2017/12/01 原判斷系統日,改判斷SP01最大值
'    If Val(Text1(0)) & Text1(1) > Val(strMaxSP01) Then
'        MsgBox "不可輸入尚未開放的業績日期！", , MsgText(5)
'        Text1_GotFocus Index
'        Cancel = True
'        bolNowChk = False
'        Exit Sub
'    End If
'    If Val(Text1(0).Tag) <> Val(Text1(0)) & Text1(1) Then
'        'Modify by Amy 2016/09/01 財務操作 10506 北五 誠安 改查10505時業務區會被清空
'        'Modify by Amy 2017/09/26 10610月中一高國碩/陳頌恩轉中二,中一主管9月可看兩人資料10月不可
'        'If cboSalesArea(0).Enabled = True And cboSalesArea(0) = MsgText(601) Then
'        'Modify by Amy 2019/10/16 當智權人員有資料時,修改日期員工編號會被清掉,若原本有員工編號不清
''        If cboSalesArea(0).Enabled = False Then
''            'Modify by Amy 2019/08/01 F4102王文安可操作,F4103陳鳳英及其職代A0914可操作
''            strExc(0) = GetSPDept(strNowUser) '因人員轉他區,故需抓條件月份下的區別
''            If strExc(0) <> "" Then strSP48 = strExc(0)
''            'end 2019/08/01
''        ElseIf CboEmp <> MsgText(601) Then
''            strSP48 = GetSPDept(CboEmp)
''        End If
'        If cboEmp <> MsgText(601) Then
'            strExc(1) = Left(cboEmp, InStr(cboEmp, " ") - 1)
'        Else
'            strExc(1) = strNowUser
'        End If
'        strExc(0) = GetSPDept(strExc(1)) '重抓目前人員日期條件下的部門
'        If strExc(0) = MsgText(601) Then
'            '總經理權限登入若輸的智權人員不為空,且不是財務處或不是M51薛經理,重抓部門
'            If bolGlMan = True And Trim(cboEmp) <> MsgText(601) And Left(cboEmp, InStr(cboEmp, " ") - 1) <> "74001" Then
'                SetcboSalesArea (strSP48)
'                strSP48 = ""
'            End If
'        ElseIf strSP48 <> strExc(0) Or Left(cboSalesArea(0), 3) <> strExc(0) Then
'             strSP48 = strExc(0) '因人員轉他區,故需抓條件月份下的區別
'        End If
'
'        'Modify by Amy 2019/10/16 下拉選單部門前三碼與智權人員部門不同,重抓部門
'        If Left(cboSalesArea(0), 3) <> strSP48 Then
'            'cboSalesArea(0).Clear 'Mark by Amy 2021/11/11
'            cboSalesArea(0) = strSP48 & " " & A0902Query(strSP48)
'        End If
'        'end 2019/10/16
'        'End If
'
'        Call SpecData '特殊轉區人員
'        'Mark by Amy 2019/10/16 避免輸入一直被清空,林青祺87011 10610不可看高國碩和陳頌恩 人員重抓改至SpecData
''        If cboEmp.Enabled = True Then 'And cboEmp = MsgText(601)
''            bolNowChk = False
''            Call SetEmp(1)
''            cboEmp = "" 'Add by Amy 2016/04/07 不預設人員
''        End If
'
'        'end 2016/09/01
'        '自動查詢
'        'Modify by Amy 2019/10/16 +bolNoChkMod 避免無窮迴圈及有修改才執行
'        If Trim(cboSalesArea(0)) <> MsgText(601) And Trim(cboEmp) <> MsgText(601) And Val(Text1(0).Tag) <> Val(Trim(Text1(0) & Text1(1))) And Val(Trim(Text1(0) & Text1(1))) <> 0 Then
'            bolNoChkMod = True
'            call cmdSearch_Click(0)
'            bolNoChkMod = True = False
'        Else
'            FormClear (1)
'        End If
'        'end 2017/09/26
'        Text1(0).Tag = Val(Text1(0)) & Text1(1)
'    End If
'    bolNowChk = False
'    'end 2016/02/23
End Sub

Private Sub Text11_GotFocus(Index As Integer)
     TextInverse Text11(Index)
End Sub

Private Sub Text11_Validate(Index As Integer, Cancel As Boolean)
    If Text11(0) = MsgText(601) Or Text11(1) = MsgText(601) Or bolNowChk = True Then Exit Sub
    
    bolNowChk = True
    Text11(1) = Format(Text11(1), "00")
    
    If ChkDate(Val(Text11(0)) & Text11(1) & "01") = False Then
        Text11_GotFocus Index
        Cancel = True
        bolNowChk = False
        Exit Sub
    End If
    If Val(Text11(0)) & Text11(1) < Val(業績輸入啟用年月) Then
        MsgBox "業績輸入上線日為" & Left(Val(業績輸入啟用年月), 3) & "年" & Right(Val(業績輸入啟用年月), 2) & _
                    "故系統無資料！", , MsgText(5)
        Text11_GotFocus Index
        Cancel = True
        bolNowChk = False
        Exit Sub
    End If
    'Modify by Amy 2017/12/01 原判斷系統日,改判斷SP01最大值
    If Val(Text11(0)) & Text11(1) > Val(strMaxSP01) Then
        MsgBox "不可輸入尚未開放的業績日期！", , MsgText(5)
        Text11_GotFocus Index
        Cancel = True
        bolNowChk = False
        Exit Sub
    End If
    bolNowChk = False
    Text11(0).Tag = Text11(0) & Text11(1)
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
End Sub

Private Sub Text2_LostFocus()
    If Text2.Locked = True Then Exit Sub
    
    Dim strMsg As String
    
    If Check1(0).Value = 1 And (Check1(1).Value = 1 Or Check1(2).Value = 1) Then
        Call RunSum(UCase("Text2"))
'        If bolNowChk = False Then
'            'Modify by Amy 2016/04/07 +if 總經理權限可以輸負值
'            If bolGlMan = False And cmdSave(0).Enabled = True Then
'                If ChkMinus(strMsg) = True Then
'                    bolNowChk = True
'                    MsgBox strMsg & "數值不可為負數"
'                    TextInverse Text2
'                    Text2.SetFocus
'                    Exit Sub
'                End If
'            End If
'        Else
'            bolNowChk = False
'        End If
    End If
End Sub

Private Sub Text4_GotFocus(Index As Integer)
    If Text4(Index).Locked = True Then
        Text4(Index).TabStop = False
        Exit Sub
    Else
        Text4(Index).TabStop = True
    End If
    
    TextInverse Text4(Index)
    If Index = 0 Then
        Option1_Click (0)
    Else
        Option2_Click (0)
    End If
End Sub

Private Sub Text4_LostFocus(Index As Integer)
    If Text4(Index).Locked = True Then Exit Sub
    
    Dim strMsg As String
    
    Call RunSum(UCase("Text4"), Index)
'    If bolNowChk = False Then
'        'Modify by Amy 2016/04/07 +if 總經理權限可以輸負值
'        If bolGlMan = False Then
'            If ChkMinus(strMsg) = True Then
'                bolNowChk = True
'                MsgBox strMsg & "數值不可為負數"
'                TextInverse Text4(Index)
'                Text4(Index).SetFocus
'                Exit Sub
'            End If
'        End If
'    Else
'        bolNowChk = False
'    End If
End Sub

Private Sub Text5_GotFocus(Index As Integer)
    If Text5(Index).Locked = True Then
        Text5(Index).TabStop = False
        Exit Sub
    Else
        Text5(Index).TabStop = True
    End If
    
    TextInverse Text5(Index)
    If Index = 0 Then
        Option1_Click (2)
    Else
        Option2_Click (2)
    End If
End Sub

Private Sub Text5_LostFocus(Index As Integer)
    If Text5(Index).Locked = True Then Exit Sub
    
    Dim strMsg As String
    Call RunSum(UCase("Text5"), Index)
'    If bolNowChk = False Then
'        'Modify by Amy 2016/04/07 +if 總經理權限可以輸負值
'        If bolGlMan = False Then
'            If ChkMinus(strMsg) = True Then
'                bolNowChk = True
'                MsgBox strMsg & "數值不可為負數"
'                TextInverse Text5(Index)
'                Text5(Index).SetFocus
'                Exit Sub
'            End If
'        End If
'    Else
'        bolNowChk = False
'    End If
End Sub

Private Sub TxtSP_GotFocus(Index As Integer)
    If TxtSP(Index).Locked = True Then
        TxtSP(Index).TabStop = False
        Exit Sub
    Else
        TxtSP(Index).TabStop = True
    End If
    
    TextInverse TxtSP(Index)
    
    'add by sonia
    'Modify by Amy 2016/02/16 主管未輸帶個人的值
    Select Case Index
      Case 7, 11, 15
         If Option1(0).Value = True Then
            If Val(Label31(11)) <> 0 Then
               TxtSP(Index) = Val(Label31(11))
            Else
               TxtSP(Index) = Val(Text4(0))
            End If
         ElseIf Option1(1).Value = True Then
            If Val(Label31(12)) <> 0 Then
               TxtSP(Index) = Val(Label31(12))
            Else
               TxtSP(Index) = Val(TxtSP(3))
            End If
         Else
            If Val(Label31(1)) <> 0 Then
               TxtSP(Index) = Val(Label31(1))
            Else
               TxtSP(Index) = Val(Text5(0))
            End If
         End If
      Case 28, 32, 36
         If Option2(0).Value = True Then
            If Val(Label41(11)) <> 0 Then
               TxtSP(Index) = Val(Label41(11))
            Else
               TxtSP(Index) = Val(Text4(1))
            End If
         ElseIf Option2(1).Value = True Then
            If Val(Label41(12)) <> 0 Then
               TxtSP(Index) = Val(Label41(12))
            Else
               TxtSP(Index) = Val(TxtSP(24))
            End If
         Else
            If Val(Label41(1)) <> 0 Then
               TxtSP(Index) = Val(Label41(1))
            Else
               TxtSP(Index) = Val(Text5(1))
            End If
         End If
    End Select
    'end sonia
    TextInverse TxtSP(Index)
    Select Case Index
        Case 7, 11, 15, 19
            bolOrgSet(1) = True
            If (TxtSP(7) <> Empty Or TxtSP(11) <> Empty Or TxtSP(15) <> Empty) _
              And (Label31(1) <> Empty Or Label31(11) <> Empty Or Label31(12) <> Empty) Then
                bolOrgSet(1) = False
            End If
            Call SetChgTxtLbl(bolOrgSet(1), 1)
        Case 28, 32, 36, 40
            bolOrgSet(2) = True
            If (TxtSP(28) <> Empty Or TxtSP(32) <> Empty Or TxtSP(36) <> Empty) _
              And (Label41(1) <> Empty Or Label41(11) <> Empty Or Label41(12) <> Empty) Then
                bolOrgSet(2) = False
            End If
            Call SetChgTxtLbl(bolOrgSet(2), 2)
        Case Else
    End Select
        Call SetInputText(0)
    If Check1(0).Value = 1 Then
        Call SetInputText(98)
    Else
        Call SetInputText(1)
        Call SetInputText(2)
    End If
End Sub

Private Sub TxtSP_LostFocus(Index As Integer)
    If TxtSP(Index).Locked = True Then Exit Sub
    
    Dim strMsg As String
    Call RunSum(UCase("TxtSP"), Index)
    
    'add by sonia 2016/1/26
    Select Case Index
        Case 7, 11, 15, 19
            bolOrgSet(1) = True
            If (TxtSP(7) <> Empty Or TxtSP(11) <> Empty Or TxtSP(15) <> Empty) _
              And (Label31(1) <> Empty Or Label31(11) <> Empty Or Label31(12) <> Empty) Then
                bolOrgSet(1) = False
            End If
            Call SetChgTxtLbl(bolOrgSet(1), 1)
        Case 28, 32, 36, 40
            bolOrgSet(2) = True
            If (TxtSP(28) <> Empty Or TxtSP(32) <> Empty Or TxtSP(36) <> Empty) _
              And (Label41(1) <> Empty Or Label41(11) <> Empty Or Label41(12) <> Empty) Then
                bolOrgSet(2) = False
            End If
            Call SetChgTxtLbl(bolOrgSet(2), 2)
        Case Else
    End Select
    Call SetInputText(0)
    If Check1(0).Value = 1 Then
        Call SetInputText(98)
    Else
        Call SetInputText(1)
        Call SetInputText(2)
    End If
    'end 2016/1/26
'    If bolNowChk = False Then
'        'Modify by Amy 2016/04/07 +if 總經理權限可以輸負值
'        If bolGlMan = False Then
'            If ChkMinus(strMsg) = True Then
'                bolNowChk = True
'                MsgBox strMsg & "數值不可為負數"
'                TextInverse TxtSP(Index)
'                TxtSP(Index).SetFocus
'                Exit Sub
'            End If
'        End If
'    Else
'        bolNowChk = False
'    End If
End Sub

Private Function SqlUpd(ByRef intOpt1 As Integer, ByRef intOpt2 As Integer) As String
    Dim j As Integer, sIdx As Integer, eIdx As Integer
    Dim strTmp As String
    Dim strUpd As String, strVal As String
    'Add by Amy 2016/02/23
    Dim strV As String
    Dim oOpt As OptionButton
    Dim strMsg As String 'Add by Amy 2022/06/09
    
    eIdx = 41
    Select Case intLimit
        Case 10, 20, 40 '個人
            sIdx = 3
        Case 2 '主管一
            sIdx = 7
            eIdx = 28
        Case 3 '主管二
            sIdx = 11
            eIdx = 32
        Case 4 '總經理
            sIdx = 15
            eIdx = 36
        Case Else
    End Select

    For j = sIdx To eIdx Step 21
        If intLimit = 10 Or intLimit = 20 Or intLimit = 40 Then
            strVal = TxtSP(j)
        Else
            Select Case j
                Case 7, 11, 15 '主管級 實績 欄
                    strVal = IIf(Label31(12) <> MsgText(601), Label31(12), TxtSP(3))
                Case 28, 32, 36 '主管級 結餘 欄
                    strVal = IIf(Label41(12) <> MsgText(601), Label41(12), TxtSP(24))
                Case Else
                    strVal = TxtSP(j)
            End Select
        End If
        strTmp = Empty
        '數值
        If m_FieldList(j).fiType = 1 Then
            'Modify by Amy 2016/05/09 周哲丞以簡協理權限進入新增誠安(點選「報出實績點數」點選「主管1」預帶值)無法存檔
            '2016/05/09因計算後的期末實績為0,而判斷條件無符合故不會存,秀玲說:判斷SalesPoint存欄位有改或0都存
            '有修改或未改但目標>=當月實績或未改但非智權部門就存
            'If Val(m_FieldList(j).fiOldData) <> Val(strVal) Or (Val(strVal) = 0 And Val(Label21(0)) >= Val(Label31(6))) _
              Or (Val(strVal) = 0 And Left(PUB_GetStaffST15(Left(CboEmp, 5), 1), 1) <> "S") Then
            If Val(m_FieldList(j).fiOldData) <> Val(strVal) Or Val(strVal) = 0 Then
                 strTmp = m_FieldList(j).fiName & "=" & Val(strVal) & ","
                 'Mark by Amy 2022/06/09 避免與智權期末結餘保留傳票產生規則不一致
'                 'Add by Amy 2017/09/26 期初+當月結餘且總經理結餘欄有修改更新SalesBalance
'                 strExc(0) = ""
'                 If j = 36 And Val(Label41(0)) + Val(Label41(1)) <> 0 And _
'                    ExistCheck("SalesBalance", "SB01", Val(Text1(0) & Text1(1)), strExc(0), False) = True Then
'                   strUpdSB = GetSBSql(m_FieldList(2).fiOldData, strVal)
'                 End If
'                 'end 2017/09/26
            End If
            'Add by Amy 2022/06/09  結餘欄位有修改 更新acc0b1.axb16 =Y
            strExc(0) = ""
            If ExistCheck("SalesBalance", "SB01", Val(Text1(0) & Text1(1)), strExc(0), False) = True And (j = 24 Or j = 28 Or j = 32 Or j = 36) Then
                If Val(m_FieldList(j).fiOldData) <> Val(strVal) Then
                    strMsg = "N"
                    If WirteAxb16(Val(Text1(0) & Text1(1)), "Y", strMsg) = False Then
                        PUB_SendMail strUserNum, "A2004", "", "每月點數結算及查詢 結餘 欄位修改寫 Axb16有問題", strMsg
                    End If
                End If
            End If
            'end 2022/06/09
        Else
            If m_FieldList(j).fiOldData <> strVal Then
                 strTmp = m_FieldList(j).fiName & "=" & "'" & strVal & "',"
            End If
        End If
        If strTmp <> Empty Then
           strUpd = strUpd & strTmp
        End If
    Next j
    
    '判斷轉撥增減欄及撥增備註是否有修改(有輸才存,不控制目標>=當月實績)
    If intLimit <> 10 Then
        For j = 19 To 20
            strTmp = Empty
            If m_FieldList(j).fiType = 1 Then
                If Val(m_FieldList(j).fiOldData) <> Val(TxtSP(j)) Then
                    strTmp = m_FieldList(j).fiName & "=" & Val(TxtSP(j)) & ","
                End If
            Else
                If m_FieldList(j).fiOldData <> TxtSP(j) Then
                    strTmp = m_FieldList(j).fiName & "=" & "'" & TxtSP(j) & "',"
                End If
            End If
            If strTmp <> Empty Then
                strUpd = strUpd & strTmp
            End If
        Next j
        For j = 40 To 41
            strTmp = Empty
            If m_FieldList(j).fiType = 1 Then
                If Val(m_FieldList(j).fiOldData) <> Val(TxtSP(j)) Then
                    strTmp = m_FieldList(j).fiName & "=" & Val(TxtSP(j)) & ","
                    'Add by Amy 2022/06/09  結餘轉撥欄位有修改 更新acc0b1.axb16 =Y
                    If ExistCheck("SalesBalance", "SB01", Val(Text1(0) & Text1(1)), strExc(0), False) = True Then
                        strMsg = "N"
                        If WirteAxb16(Val(Text1(0) & Text1(1)), "Y", strMsg) = False Then
                            PUB_SendMail strUserNum, "A2004", "", "每月點數結算及查詢 結餘轉撥 欄位修改寫 Axb16有問題", strMsg
                        End If
                    End If
                    'end 2022/06/09
                End If
            Else
                If m_FieldList(j).fiOldData <> TxtSP(j) Then
                    strTmp = m_FieldList(j).fiName & "=" & "'" & TxtSP(j) & "',"
                End If
            End If
            If strTmp <> Empty Then
                strUpd = strUpd & strTmp
            End If
        Next j
    End If
    'Add by Amy 2016/02/23 +記錄畫面勾選
    '勾「報出點數」
    If Check1(0).Value = 1 Then
        If Check1(1).Value = 1 Then
            strV = "1"
        Else
            strV = "2"
        End If
    Else
         strV = "0"
    End If
    For Each oOpt In Option1
        If oOpt.Value = True Then
            intOpt1 = oOpt.Index
            Exit For
        End If
    Next
    For Each oOpt In Option2
        If oOpt.Value = True Then
            intOpt2 = oOpt.Index
            Exit For
        End If
    Next
    For j = 49 To 51
        strTmp = Empty
        If m_FieldList(j).fiType = 1 Then
            If j = 50 Then
                strTmp = m_FieldList(j).fiName & "=" & Val(intOpt1) & ","
            ElseIf j = 51 Then
                strTmp = m_FieldList(j).fiName & "=" & Val(intOpt2) & ","
            End If
        Else
            'SP49
            If m_FieldList(j).fiOldData <> strV Then
                strTmp = m_FieldList(j).fiName & "=" & "'" & strV & "',"
            End If
        End If
        If strTmp <> Empty Then
            strUpd = strUpd & strTmp
        End If
    Next j
    'end 2016/02/23
    
    If strUpd <> Empty Then
        strExc(1) = Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))), 6)
        strUpd = "Update SalesPoint Set " & Mid(strUpd, 1, Len(strUpd) - 1) & _
                    " Where SP01=" & strExc(1) & _
                    " And SP02='" & CboEmp.Tag & "' "
        SqlUpd = strUpd
    End If
End Function

' 初始化欄位陣列
Private Sub InitialField()
    Dim nIndex As Integer
    Dim strTmp As String
       
    ' 初始化欄位陣列
    For nIndex = 1 To intField
        strTmp = Format(nIndex, "00")
        m_FieldList(nIndex).fiName = "SP" & strTmp
        m_FieldList(nIndex).fiOldData = Empty
        m_FieldList(nIndex).fiType = 1  '數值型態
        Select Case nIndex
            'Modify by Amy 2016/02/16 +SP49
            Case 2, 4, 8, 12, 16, 20, 21, 25, 29, 33, 37, 41, 42, 45, 49
                m_FieldList(nIndex).fiType = 0 '文字型態
        End Select
    Next nIndex
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
Dim bFirst As Boolean
    
    For nIndex = 1 To UBound(m_FieldList)
        If rsTmp.RecordCount > 0 Then
            If m_FieldList(nIndex).fiName <> Empty Then
                If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
                    m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
                Else
                    m_FieldList(nIndex).fiOldData = Empty
                End If
            End If
        Else
            m_FieldList(nIndex).fiOldData = Empty
        End If
    Next nIndex

End Sub


Private Sub SetCUID(intSP As Integer, ByVal stUser As String, ByVal stDate As String, ByVal stTime As String)
    Dim strTemp(1) As String
    
    If stUser = MsgText(601) Then Exit Sub
    
    If IsEmptyText(stDate) = False Then strTemp(0) = Format(TAIWANDATE(stDate), "###/##/##")
    If IsEmptyText(stTime) = False Then strTemp(1) = Format(Left(stTime, IIf(Len(stTime) = 6, 4, 3)), "##:##")
    
    strTemp(0) = GetStaffName(stUser, True) & " " & strTemp(0) & " " & strTemp(1)

    Select Case intSP
        Case 0
            lbl21(0) = "CREATE: " & strTemp(0)
        Case 1
            lbl21(1) = "UPDATE: " & strTemp(0)
        Case 7
            'Modified by Lydia Lydia 2022/01/03 Label31(7)=>lbl21(2)
            lbl21(2) = strTemp(0)
        Case 11
            'Modified by Lydia Lydia 2022/01/03 Label31(8)=>lbl21(3)
            lbl21(3) = strTemp(0)
        Case 15
            'Modified by Lydia Lydia 2022/01/03 Label31(9)=>lbl21(4)
            lbl21(4) = strTemp(0)
        Case 28
            'Modified by Lydia Lydia 2022/01/03 Label41(7)=>lbl21(5)
            lbl21(5) = strTemp(0)
        Case 32
            'Modified by Lydia Lydia 2022/01/03 Label41(7)=>lbl21(6)
            lbl21(6) = strTemp(0)
        Case 36
            'Modified by Lydia Lydia 2022/01/03 Label41(7)=>lbl21(7)
            lbl21(7) = strTemp(0)
        Case 21
            'Modified by Lydia Lydia 2022/01/03 Label31(2)=>lbl21(8)
            lbl21(8) = strTemp(0)
        Case 42
            'Modified by Lydia Lydia 2022/01/03 Label41(2)=>lbl21(9)
            lbl21(9) = strTemp(0)
    End Select
End Sub

'intSet:0-全部/1-勾選(報出點數/實績部分/結餘部分)/2-點選實績(動用/期末/報出)/3-點選結餘(動用留/期末/報出)
Private Sub SetChoose(ByVal bolEnabled As Boolean, Optional ByVal intSet As Integer = 0)
    Dim oCheck As CheckBox
    Dim oOpt As OptionButton
    
    If intSet = 0 Or intSet = 1 Then
        For Each oCheck In Check1
            oCheck.Enabled = bolEnabled
        Next
    End If
    
    If intSet = 0 Or intSet = 2 Then
        For Each oOpt In Option1
            oOpt.Enabled = bolEnabled
        Next
    End If
    
    If intSet = 0 Or intSet = 3 Then
        For Each oOpt In Option2
            oOpt.Enabled = bolEnabled
        Next
    End If
End Sub

Private Sub RefreshRecord(ByRef rsRec As ADODB.Recordset)
    
    If rsRec.RecordCount > 0 Then
        If Not IsNull(rsRec.Fields("SP03")) Then TxtSP(3) = Val("" & rsRec.Fields("SP03")): Call TxtSP_LostFocus(3)
        If Not IsNull(rsRec.Fields("SP04")) Then Call SetCUID(0, rsRec.Fields("SP04"), rsRec.Fields("SP05"), rsRec.Fields("SP06"))
        If Not IsNull(rsRec.Fields("SP07")) Then
            TxtSP(7) = Val("" & rsRec.Fields("SP07")) ': Call TxtSP_LostFocus(7)
            Call SetCUID(7, rsRec.Fields("SP08"), rsRec.Fields("SP09"), rsRec.Fields("SP10"))
            Label31(12) = TxtSP(7) ': Call TxtSP_LostFocus(7)
        End If
        If Not IsNull(rsRec.Fields("SP11")) Then
            TxtSP(11) = Val("" & rsRec.Fields("SP11")) ': Call TxtSP_LostFocus(11)
            Call SetCUID(11, rsRec.Fields("SP12"), rsRec.Fields("SP13"), rsRec.Fields("SP14"))
            Label31(12) = TxtSP(11)
        End If
        If Not IsNull(rsRec.Fields("SP15")) Then
            TxtSP(15) = Val("" & rsRec.Fields("SP15")) ': Call TxtSP_LostFocus(15)
            Call SetCUID(15, rsRec.Fields("SP16"), rsRec.Fields("SP17"), rsRec.Fields("SP18"))
            Label31(12) = TxtSP(15)
        End If
        If Not IsNull(rsRec.Fields("SP19")) Then TxtSP(19) = Val("" & rsRec.Fields("SP19"))
        TxtSP(20) = "" & rsRec.Fields("SP20")
        If Not IsNull(rsRec.Fields("SP21")) Then
            Call SetCUID(21, rsRec.Fields("SP21"), rsRec.Fields("SP22"), rsRec.Fields("SP23"))
        End If
            
        If Not IsNull(rsRec.Fields("SP24")) Then TxtSP(24) = Val("" & rsRec.Fields("SP24")): Call TxtSP_LostFocus(24)
        If Not IsNull(rsRec.Fields("SP25")) Then Call SetCUID(1, rsRec.Fields("SP25"), rsRec.Fields("SP26"), rsRec.Fields("SP27"))
        If Not IsNull(rsRec.Fields("SP28")) Then
            TxtSP(28) = Val("" & rsRec.Fields("SP28")) ': Call TxtSP_LostFocus(28)
            Call SetCUID(28, rsRec.Fields("SP29"), rsRec.Fields("SP30"), rsRec.Fields("SP31"))
            Label41(12) = TxtSP(28)
        End If
        If Not IsNull(rsRec.Fields("SP32")) Then
            TxtSP(32) = Val("" & rsRec.Fields("SP32")) ': Call TxtSP_LostFocus(32)
            Call SetCUID(32, rsRec.Fields("SP33"), rsRec.Fields("SP34"), rsRec.Fields("SP35"))
            Label41(12) = TxtSP(32)
        End If
        If Not IsNull(rsRec.Fields("SP36")) Then
            TxtSP(36) = Val("" & rsRec.Fields("SP36")): ' Call TxtSP_LostFocus(36)
            Call SetCUID(36, rsRec.Fields("SP37"), rsRec.Fields("SP38"), rsRec.Fields("SP39"))
            Label41(12) = TxtSP(36)
        End If
        If Not IsNull(rsRec.Fields("SP40")) Then TxtSP(40) = Val("" & rsRec.Fields("SP40"))
        TxtSP(41) = "" & rsRec.Fields("SP41")
        If Not IsNull(rsRec.Fields("SP42")) Then
            Call SetCUID(42, rsRec.Fields("SP42"), rsRec.Fields("SP43"), rsRec.Fields("SP44"))
        End If
    End If
    rsRec.Close
End Sub

'確認畫面是否有修改
Private Function ChkModify() As Boolean
    Dim j As Integer, sIdx As Integer, eIdx As Integer
    Dim strVal As String
    Dim bolEmpty As Boolean, strValTag As String 'Add by Amy 2016/05/09
    
    ChkModify = False
    If intLimit <= 0 Then Exit Function
        
    eIdx = 41
    Select Case intLimit
        Case 10, 20, 30, 40 '個人
            sIdx = 3
        Case 2 '主管一
            sIdx = 7
            eIdx = 28
        Case 3 '主管二
            sIdx = 11
            eIdx = 32
        Case 4 '總經理
            sIdx = 15
            eIdx = 36
        Case Else
    End Select
    
    '實績/結餘
    For j = sIdx To eIdx Step 21
        'Modify by Amy 2016/05/09 +避免一直彈存檔訊息
        '個人操作自己
        If intLimit = 10 Or intLimit = 20 Or intLimit = 40 Then
            strVal = TxtSP(j)
            If j = 3 Then
                bolEmpty = bolEmptyF1
            Else
                bolEmpty = bolEmptyF2
            End If
            'SalesPoint沒資料判斷預設值是否修改
            If bolEmpty = True Then strValTag = TxtSP(j).Tag
        Else
            Select Case j
                Case 7, 11, 15
                    strVal = Label31(12)
                    bolEmpty = bolEmptyF1
                Case 28, 32, 36
                    strVal = Label41(12)
                    bolEmpty = bolEmptyF2
                Case Else
                    strVal = TxtSP(j)
            End Select
            If bolEmpty = True Then strVal = TxtSP(j)
        End If
        
        '數值
        'Modify by Amy 2016/05/09 +避免一直彈存檔訊息
        If m_FieldList(j).fiType = 1 Then
            'SalesPoint沒資料
            If bolEmpty = True Then
                '個人操作
                If intLimit >= 10 Then
                    '操作自己的有改彈訊息(與預設值比)
                    If Val(strVal) <> Val(strValTag) Then ChkModify = True: Exit For
                '主管操作他人畫面欄位有值就彈訊息
                ElseIf strVal <> MsgText(601) Then
                    ChkModify = True: Exit For
                End If
            'SalesPoint有資料,判斷期末值有變就彈訊息
            ElseIf Val(m_FieldList(j).fiOldData) <> Val(strVal) Then
                ChkModify = True: Exit For
            End If
        End If
    Next j
    If ChkModify = True Then Exit Function
 
    '主管身份-判斷轉撥增減欄及撥增備註是否有修改
    If intLimit <> 10 Then
        If bolEmptyF3 = True Then
            'SalesPoint沒資料但畫面對應欄位有值就彈訊息
            If TxtSP(19) <> MsgText(601) Or TxtSP(20) <> MsgText(601) Then ChkModify = True
        Else
            If Val(m_FieldList(19).fiOldData) <> Val(TxtSP(19)) Then ChkModify = True
            If m_FieldList(20).fiOldData <> TxtSP(20) And m_FieldList(20).fiOldData <> MsgText(601) Then
                ChkModify = True
            End If
        End If
        If ChkModify = True Then Exit Function

        If bolemptyF4 = True Then
            'SalesPoint沒資料但畫面對應欄位有值就彈訊息
            If TxtSP(40) <> MsgText(601) Or TxtSP(41) <> MsgText(601) Then ChkModify = True
        Else
            If Val(m_FieldList(40).fiOldData) <> Val(TxtSP(40)) Then ChkModify = True
            If m_FieldList(41).fiOldData <> TxtSP(41) And m_FieldList(41).fiOldData <> MsgText(601) Then
                ChkModify = True
            End If
        End If
        'end 2016/05/09
        If ChkModify = True Then Exit Function
    End If
End Function

'確認畫面是否負值
Private Function ChkMinus(ByRef stMsg As String) As Boolean
    Dim strTp(1) As String
    
    strTp(0) = "實績": strTp(1) = "結餘"
    stMsg = MsgText(601): ChkMinus = False
    
    If Val(Text2) < 0 Then stMsg = stMsg & "/ " & Mid(Check1(0).Caption, 1, Len(Check1(0).Caption) - 1)
    
    '實績-總經理/主管二/主管一 調整
    If m_FieldList(15).fiOldData <> Empty Or TxtSP(15) <> Empty Or m_FieldList(11).fiOldData <> Empty Or TxtSP(11) <> Empty _
      Or m_FieldList(7).fiOldData <> Empty Or TxtSP(7) <> Empty Then
        '報出實績主管欄
        If Val(Label31(1)) < 0 Then stMsg = stMsg & "/ " & Mid(Lbl10(0), 4, 6)
    Else
        '報出實績個人欄
        If Val(Text5(0)) < 0 Then stMsg = stMsg & "/ " & Mid(Lbl10(0), 4, 6)
    End If
          
    '結餘-總經理/主管二/主管一 調整
    If m_FieldList(36).fiOldData <> Empty Or TxtSP(36) <> Empty Or m_FieldList(32).fiOldData <> Empty Or TxtSP(32) <> Empty _
      Or m_FieldList(28).fiOldData <> Empty Or TxtSP(28) <> Empty Then
        '報出結餘主管欄 ex:84045 10503 個人未輸簡協理調整後為負值
        If Val(Label41(1)) < 0 Then stMsg = stMsg & "/ " & Mid(Lbl10(1), 4, 6)
    Else
        '報出結餘個人欄
        If Val(Text5(1)) < 0 Then stMsg = stMsg & "/ " & Mid(Lbl10(1), 4, 6)
    End If
    
    If Val(TxtSP(3)) < 0 Or Val(Label31(12)) < 0 Then stMsg = stMsg & "/ " & strTp(0) & Mid(Option1(1).Caption, 1, Len(Option1(1).Caption) - 1)
    If Val(TxtSP(24)) < 0 Or Val(Label41(12)) < 0 Then stMsg = stMsg & "/ " & strTp(1) & Mid(Option2(1).Caption, 1, Len(Option2(1).Caption) - 1)
    If stMsg <> MsgText(601) Then ChkMinus = True: stMsg = Mid(stMsg, 3) & vbCrLf
End Function

'依記錄勾選值更新畫面設定
Private Sub SetViewVal()
    Dim j As Integer
    Dim bolLocked As Boolean '記錄欄位目前Locked值
    Dim strVal As String
    If m_FieldList(49).fiOldData = Empty Then
        bolNoMsg = True
        Call TxtSP_LostFocus(3)
        Check1(0).Value = 1
        If intLimit > 0 Then Check1(0).Enabled = True
        bolNoMsg = False
        Check1(1).Value = 1
        Option1(2).Value = True
        Option2(2).Value = True
    Else
        For j = 7 To 15 Step 4
            If TxtSP(j) <> Empty Then
                '實績
                bolLocked = TxtSP(j).Locked
                TxtSP(j).Locked = False
                Call TxtSP_LostFocus(j)
                If TxtSP(j).Locked <> bolLocked Then TxtSP(j).Locked = bolLocked
                Select Case Val(m_FieldList(50).fiOldData)
                    Case 0
                        strVal = Label31(11)
                    Case 1
                        strVal = TxtSP(j)
                    Case 2
                        strVal = Label31(1)
                End Select
                TxtSP(j) = strVal
                '結餘
                bolLocked = TxtSP(j + 21).Locked
                TxtSP(j + 21).Locked = False
                Call TxtSP_LostFocus(j + 21)
                If TxtSP(j + 21).Locked <> bolLocked Then TxtSP(j + 21).Locked = bolLocked
                Select Case Val(m_FieldList(51).fiOldData)
                    Case 0
                        strVal = Label41(11)
                    Case 1
                        strVal = TxtSP(j + 21)
                    Case 2
                        strVal = Label41(1)
                End Select
                TxtSP(j + 21) = strVal
            End If
        Next j
    End If
    
    '勾 實績 or 結餘 點數
    If m_FieldList(49).fiOldData <> Empty And m_FieldList(49).fiOldData <> "0" Then
        bolNoMsg = True
        Check1(0).Value = 1
        If intLimit > 0 Then Check1(0).Enabled = True
        bolNoMsg = False
        Check1(m_FieldList(49).fiOldData).Value = 1
    End If
    
    If m_FieldList(50).fiOldData <> Empty And Val(m_FieldList(50).fiOldData) <> 1 Then Option1(m_FieldList(50).fiOldData).Value = True
    If m_FieldList(51).fiOldData <> Empty And Val(m_FieldList(51).fiOldData) <> 1 Then Option2(m_FieldList(51).fiOldData).Value = True
End Sub

'調整過後的mail內容(因個人畫面點與主管可能不一樣,故只show主管報出最後的值-文雄)
'Modify by Amy 2016/11/18 +stEmpNo 參數
Private Function GetViewData(ByVal intOpt1 As Integer, ByVal intOpt2 As Integer, stEmpNo As String) As String
    Dim strTemp(1) As String
    
    'Modify by Amy 2016/11/18 +調整對象
    'Modified by Lydia 2019/08/08 frm210152智權點數實績與結餘輸入在承辦人系統更名為「每月點數查詢／輸入」，與財務系統不同
    'GetViewData = "敬啟者：" & vbCrLf & vbCrLf & _
                            "經主管調整後，目前智權點數實績與結餘資料更新如下：" & vbCrLf & vbCrLf & _
                            "調整　　對象：" & GetStaffName(stEmpNo, True) & vbCrLf & _
                            "報出　　點數：" & Text2 & vbCrLf
    GetViewData = "敬啟者：" & vbCrLf & vbCrLf & _
                            "經主管調整後，目前" & Me.Caption & "更新如下：" & vbCrLf & vbCrLf & _
                            "調整　　對象：" & GetStaffName(stEmpNo, True) & vbCrLf & _
                            "報出　　點數：" & Text2 & vbCrLf
                            
'    Select Case intOpt1
'        Case 0
'            strTemp(0) = "實績動用：" & IIf(Label31(11) <> MsgText(601), Label31(11), Text4(0)) & vbCrLf
'        Case 1
'            strTemp(0) = "期末實績：" & IIf(Label31(12) <> MsgText(601), Label31(12), TxtSP(3)) & vbCrLf
'        Case 2
            strTemp(0) = "報出實績點數：" & IIf(Label31(1) <> MsgText(601), Label31(1), Text5(0)) & vbCrLf
'    End Select
    GetViewData = GetViewData & strTemp(0)
    If Val(m_FieldList(19).fiOldData) <> Val(TxtSP(19)) Then
        GetViewData = GetViewData & "轉撥實績增減：" & TxtSP(19) & vbCrLf
        'Add by Amy 2016/04/07 有調整轉撥備註需顯示
        If TxtSP(20) <> MsgText(601) Then GetViewData = GetViewData & "轉撥實績備註：" & TxtSP(20) & vbCrLf
    End If
    
'    Select Case intOpt2
'        Case 0
'            strTemp(1) = "結餘動用：" & IIf(Label41(11) <> MsgText(601), Label41(11), Text4(1)) & vbCrLf
'        Case 1
'            strTemp(1) = "期末結餘：" & IIf(Label41(12) <> MsgText(601), Label41(12), TxtSP(24)) & vbCrLf
'        Case 2
            strTemp(1) = "報出結餘點數：" & IIf(Label41(1) <> MsgText(601), Label41(1), Text5(1)) & vbCrLf
'    End Select
    GetViewData = GetViewData & strTemp(1)
    If Val(m_FieldList(40).fiOldData) <> Val(TxtSP(40)) Then
        GetViewData = GetViewData & "轉撥結餘增減：" & TxtSP(40) & vbCrLf
        'Add by Amy 2016/04/07 有調整轉撥備註需顯示
        If TxtSP(41) <> MsgText(601) Then GetViewData = GetViewData & "轉撥結餘備註：" & TxtSP(41) & vbCrLf
    End If
End Function

'Add by Amy 2016/11/03 '確認所有S部門是否區主管已確認
'Modify by Amy 2016/11/18 +intChoose參數
'Modify by Amy 2021/07/06 +strWhere 參數
Private Function ChkAllAccept(strMsg As String, Optional ByVal intChoose As Integer = 0, Optional ByVal strWhere As String = "") As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String

    ChkAllAccept = False: strMsg = ""
    'Modify by Amy 2021/01/07 法語改抓共用Function 瑞婷過帳後才發現有69005結餘轉撥給10051沒輸隱藏版
    '                                             輸完後因北五區已確認,應再開放北五區主管確認
'    'Add by Amy 2016/11/18
'    Select Case intChoose
'        Case 0
'            strQ = "And SubStr(a.st15,1,1)='S' "
'        Case 2 '中所
'            strQ = "And SubStr(a.st15,1,2)='S2' "
'    End Select
'    'Modify by Amy 2017/01/04 SP01 帶錯,導致按確認就發
'    'Modify by Amy 2019/10/16 +F4102/F4103/W1001/W2001 也需確認
'    strQ = "Select Distinct sp48 From SalesPoint,Staff a " & _
'                "Where sp01=" & Val(Text1(0) & Text1(1)) + 191100 & " And sp02=a.st01(+) " & _
'                strQ & "And sp45 is null " & _
'                "And Decode(a.st04,2,'F0000',a.st01)>='6' And Decode(a.st04,2,'F0000',a.st01)<'F' " & _
'    "Union Select Distinct sp48 From SalesPoint Where sp01=" & Val(Text1(0) & Text1(1)) + 191100 & _
'                " And sp02 in('F4102','F4103','W1001','W2001') And sp45 is null " & _
'                "Order by sp48"
    'Modify by Amy 2021/07/06 +strWhere
    strQ = ChkPointAcceptSql(Val(Text1(0) & Text1(1)) + 191100, Me.Name, intChoose, strWhere)
    If InStr(strQ, "請洽電腦中心") > 0 Then strMsg = strQ: Exit Function
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount = 0 Then ChkAllAccept = True
    RsQ.Close
    Set RsQ = Nothing
End Function
'end 2021/01/07

Private Function bolAxbNotNull() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim intQ As Integer
    
    'Modify by Amy 2018/02/02 增加判斷是否已過帳, 否則當月尚未產生傳票也會彈
    strQ = "Select Distinct Axb04  From Acc0b1,Acc021 Where Axb01=(Select Max(Axb01) From Acc0b1 ) And Axb04=Ax202(+) And Ax201(+)='1' And Ax210 is null " & _
     "Union Select Distinct Axb09 From Acc0b1,Acc021 Where Axb01=(Select Max(Axb01) From Acc0b1 ) And Axb09=Ax202(+) And Ax201(+)='1' And Ax210 is null "
                   
    intQ = 1: bolAxbNotNull = False
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If Not IsNull(RsQ.Fields(0)) Then
            bolAxbNotNull = True
        End If
    End If
    'end 2018/02/02
    RsQ.Close
End Function

'Mark by Amy 2022/06/09  避免與智權期末結餘保留傳票產生不一致,不使用
'Add by Amy 2017/09/26 取得更新SalesBalance語法
Private Function GetSBSql(ByVal stAx209 As String, ByVal stSP36 As String) As String
'    Dim adoQ As New ADODB.Recordset
'    Dim strQ As String
'    Dim stDate_S As String, stDate_E As String, intQ As Integer
'    Dim stBegNow(4 To 7) As String, stAllSeq As String '期初+當月值/抓取stBS順序
'    Dim stSB_B(4 To 7) As String, stSB_L(10 To 13) As String 'SB報出/期末
'    Dim stSeq As Variant
'    Dim stTP As String, j As Integer
'
'    GetSBSql = ""
'    stDate_S = Text1(0) & Text1(1)
'    If Val(Right(stDate_S, 2)) = 12 Then
'        stDate_E = Val(Left(stDate_S, 3)) + 1 & "01"
'    Else
'        stDate_E = Left(stDate_S, 3) & IIf(Val(Right(stDate_S, 2)) + 1 <= 9, "0" & Val(Right(stDate_S, 2)) + 1, Val(Right(stDate_S, 2)) + 1)
'    End If
'
'    '*** 期初+本月 四部門值
'    strQ = "Select * From (" & _
'                    GetBalanceSQL(3, stDate_S, stDate_E, stAx209) & _
'    " Union All " & GetBalanceSQL(4, stDate_S, stDate_E, stAx209) & _
'    " Union All " & GetBalanceSQL(5, stDate_S, stDate_E, stAx209) & _
'    " Union All " & GetBalanceSQL(6, stDate_S, stDate_E, stAx209) & _
'                ") Where stVal is not null Order by StVal Asc"
'
'    If adoQ.State = adStateOpen Then adoQ.Close
'    adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'    If adoQ.RecordCount <> 0 Then
'        If adoQ.EOF = False Then adoQ.MoveFirst
'        Do While adoQ.EOF = False
'            stAllSeq = stAllSeq & "," & Val(adoQ.Fields("Ord")) + 1
'            If Val(stSP36) = 0 Then
'                '報出
'                stSB_B(Val(adoQ.Fields("Ord")) + 1) = adoQ.Fields("stVal")
'            Else
'                '期初+當月
'                stBegNow(Val(adoQ.Fields("Ord")) + 1) = adoQ.Fields("stVal")
'            End If
'            adoQ.MoveNext
'        Loop
'    End If
'    adoQ.Close
'
'    '期未為0=全報(報出值依四部門值更新,期未固定為0
'    '期初+本月-期末,依四部門合計小至大之餘額扣之,至扣完為止
'    If Not (Val(stSP36) = 0 And stAllSeq <> MsgText(601)) Then
'        stSeq = Split(Mid(stAllSeq, 2), ",")
'        stTP = (Val(Label41(0)) + Val(Label41(6)) - Val(stSP36)) * 1000
'        For j = LBound(stSeq) To UBound(stSeq)
'            If Val(stTP) = Val(stBegNow(Val(stSeq(j)))) Then
'                stSB_B(Val(stSeq(j))) = stBegNow(Val(stSeq(j)))
'                stTP = "0"
'            ElseIf Val(stTP) < Val(stBegNow(Val(stSeq(j)))) Then
'                stSB_B(Val(stSeq(j))) = stTP
'                stSB_L(Val(stSeq(j)) + 6) = Val(stBegNow(Val(stSeq(j)))) - Val(stTP)
'                stTP = "0"
'            Else
'                stSB_B(Val(stSeq(j))) = stBegNow(Val(stSeq(j)))
'                stTP = Val(stTP) - Val(stSB_B(Val(stSeq(j))))
'            End If
'        Next j
'    End If
'
'    GetSBSql = "Update SalesBalance " & _
'                 "Set SB04=" & Val(stSB_B(4)) & ",SB05=" & Val(stSB_B(5)) & ",SB06=" & Val(stSB_B(6)) & ",SB07=" & Val(stSB_B(7)) & "," & _
'                     "SB10=" & Val(stSB_L(10)) & ",SB11=" & Val(stSB_L(11)) & ",SB12=" & Val(stSB_L(12)) & ",SB13=" & Val(stSB_L(13)) & "," & _
'                     "SB08=" & Val(strSrvDate(2)) & ",SB09=" & ServerTime & _
'                " Where SB01=" & Val(stDate_S) & " And SB03='" & stAx209 & "'"
End Function

Private Function SpecData() As Boolean
    Dim stUser As String 'Add by Amy 2019/10/16
    Dim stTP(1) As String 'Add byAmy 2021/11/11
    SpecData = False
    
    Select Case strNowUser
        Case "86047" '高國碩
            '10609以前資料只能看自己的(10610才為中二區主管)
            If Val(Text1(0) & Text1(1)) <= 10609 Then
                SpecData = True
                CboEmp.Enabled = False
                tabSP.TabVisible(1) = False
                CboEmp = strNowUser & " " & GetStaffName(strNowUser, True)
            Else
                CboEmp.Enabled = True
                tabSP.TabVisible(1) = True
            End If
        'Add by Amy 2019/10/16 避免員編輸入一直被清空,搬過來設為特殊
        Case "87011" '林青祺
            '10610不可看高國碩86047和陳頌恩A3023
            bolNowChk = False
            'Modify by Amy 2021/11/11 重改SetEmp 函數 原:SetEmp(1),此人為10609月可看高國碩86047和陳頌恩A3023-測式時需注意
            Call SetEmp
        Case Else
            '總經理權限(電腦中心,財務,總經理,主任秘書(等級08))/簡協理/林炳佑 查特殊人員重抓部門
            stUser = CboEmp
            If InStr(stUser, " ") > 0 Then stUser = Left(stUser, InStr(stUser, " ") - 1)
            'Modify by Amy 2021/11/11 重改SetEmp
'            If bolGlMan = True And (stUser = "86047" Or stUser = "87011") Then
'                 bolNowChk = False
'                Call SetEmp(1)
'            End If
            If Val(Text1(0).Tag) <> Val(Text1(0)) & Text1(1) And Trim(Text1(0)) <> MsgText(601) And Trim(Text1(1)) <> MsgText(601) Then
                ChkSetDept (stUser)
            End If
            Call SetEmp(1)
        'end 2019/10/16
    End Select
End Function
'end 2017/09/26

'Add by Amy 2019/10/16 傳入編號判斷個人未輸且區主管尚未確認則更新 sp45
Private Sub ChkUpdSP45(ByVal stNo As String)
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim intQ As Integer
    
    strQ = "Select sp03||sp24 as stVal,sp45 From SalesPoint " & _
              "Where sp01=" & Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))), 6) & _
              " And sp02='" & stNo & "'"
              
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If IsNull(RsQ.Fields("stVal")) And IsNull(RsQ.Fields("sp45")) Then
            strQ = "Update SalesPoint Set sp45='" & strUserNum & "',sp46=to_number(to_char(sysdate,'YYYYMMDD')),sp47=to_number(to_char(sysdate,'HH24MISS')) " & _
                      "Where sp01=" & Left(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))), 6) & _
                      " And sp02='" & stNo & "' "
            cnnConnection.Execute strQ
        End If
    End If
    RsQ.Close
End Sub

'Add by Amy 2021/07/16 判斷欄位沒輸入回傳 True
Private Function ChkNoInput(ByVal stNo As String, ByVal stChkField As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim i As Integer, intQ As Integer, stQ As String, stWhere As String
    Dim arrTmp
    
    ChkNoInput = False
    If stChkField = "All" Then
                             '個人            主管1       主管2        經理        轉撥
        stChkField = "sp03||sp24||sp07||sp28||sp11||sp32||sp15||sp36||sp19||sp40"
    End If
    
    If InStr(stNo, ",") > 0 Then
        stWhere = " And sp02 in ('" & Replace(stNo, ",", "','") & "')"
    Else
        stWhere = " And sp02=" & CNULL(stNo)
    End If
    
    stQ = "Select  " & stChkField & " From SalesPoint Where sp01=" & Val(Text11(0) & Text11(1)) + 191100 & stWhere
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            If IsNull(RsQ.Fields(0)) Then
                ChkNoInput = True
            End If
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
End Function
 
'Add by Amy 2021/11/11 重新整理修改
Private Sub SetEmp(Optional ByVal intSet As Integer = 0)
    Dim strTemp
    Dim strQ As String, strWhere As String, stDef As String, stMsg As String, intC As Integer
    Dim bolHasData As Boolean
    Dim stTP(2) As String
    
    If bolNowChk = True Then Exit Sub
    
     bolNowChk = True
     '改「業績年月」,人員需先記錄再設回
     If intSet = 1 And CboEmp <> MsgText(601) And CboEmp <> "cboEmp" Then
        If InStr(CboEmp, " ") = 0 Then CboEmp = CboEmp & " " & GetStaffName(CboEmp, True)
        stDef = CboEmp
    End If
    '非智權部人員固定,不需清
    If Left(stST15, 1) = "S" Or bolGlMan = True Then CboEmp.Clear
    
    '總經理權限
    If bolGlMan = True Then
        CboEmp.Enabled = True
        strQ = "Select st01,st02 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & " And sp48='" & Left(cboSalesArea(0), 3) & "' " & _
                  " Order by st01"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strQ)
        If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
                 '查詢且人員於期間內有資料才預設
                If intSet = 1 And stDef <> MsgText(601) Then
                    If Left(stDef, 5) = RsTemp.Fields("st01") Then
                        bolHasData = True
                    End If
                End If
                CboEmp.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                RsTemp.MoveNext
            Loop
            '查詢時,此人員於期間內無資料,不預設 ex:張宜萱(A8007) 10906月以前都沒資料
            If intSet = 1 And bolHasData = False Then stDef = ""
            If stDef <> MsgText(601) Then CboEmp = stDef '預設
        End If
    '非智權部
    ElseIf Left(cboSalesArea(0), 1) <> "S" And Left(stST15, 1) <> "S" Then
        If strEmpList = MsgText(601) Then MsgBox "人員設定有問題，請洽電腦中心": Exit Sub
        
        stDef = ""
        If InStr(strEmpList, ";") > 0 Then
            If InStr(strEmpList, "F4102") > 0 Then
                strWhere = " And st01 In('" & Replace(Mid(Replace(";" & strEmpList, ";F4102", ""), 2), ";", "','") & "')"
            ElseIf InStr(strEmpList, "F4103") > 0 Then
                strWhere = " And st01 In('" & Replace(Mid(Replace(";" & strEmpList, ";F4103", ""), 2), ";", "','") & "')"
            Else
                strWhere = " And st01 in ('" & Replace(strEmpList, ";", "','") & "') "
            End If
        Else
            strWhere = " And st01='" & strEmpList & "' "
        End If
        '為避免有些人員某些月份於SalesPoint 無資料,故不串SalsePoint
        If InStr(strEmpList, "F4102") > 0 Then
            strQ = "Select 1 as Sort,st01,st02 From Staff Where 1=1 " & strWhere & _
            "Union Select 2 as Sort,st01,st02 From Staff Where st01='F4102' "
        ElseIf InStr(strEmpList, "F4103") > 0 Then
            strQ = "Select 1 as Sort,st01,st02 From Staff Where 1=1 " & strWhere & _
            "Union Select 2 as Sort,st01,st02 From Staff Where st01='F4103' "
        Else
            strQ = "Select 1 as Sort,st01,st02 From Staff Where 1=1 " & strWhere
        End If
        strQ = strQ & " Order by Sort,st01"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strQ)
        If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
                '初始設定時,非智權部不是輸自已員編,故預設第一筆
                If intSet = 0 And intC = 0 Then
                    stDef = RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                    intC = intC + 1
                End If
                CboEmp.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                RsTemp.MoveNext
            Loop
            If stDef <> MsgText(601) Then CboEmp = stDef '預設
        End If
        If CboEmp.ListCount > 1 Then
            CboEmp.Enabled = True
        End If
    '智權部
    Else
        If strSt52List <> MsgText(601) Then
            If InStr(strSt52List, ",") > 0 Then
                strWhere = " And sp02 in (" & strSt52List & ")"
            Else
                strWhere = " And sp02=" & strSt52List
            End If
        End If
        '林柄佑(82026)經理可以看中所全部,輸入20091(S29),不可看自己的(因登入時預設,若當期沒資料仍顯示,避免下拉其他部門後又選S29可能無20091, 故20091不論SalesPoint是否有資料都顯示)
        If strEmpList <> MsgText(601) Then
            strQ = "Select st01,st02,1 as Sort From Staff Where st15='" & Left(cboSalesArea(0), 3) & "' And st01='" & strEmpList & "' "
        '其他人員
        Else
            strQ = "Select st01,st02,1 as Sort From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & " And sp48='" & Left(cboSalesArea(0), 3) & "' And sp02='" & strNowUser & "' "
        End If
        
        '一般個人或帶人主管
        If cboSalesArea(0).Enabled = False And bolAreaMan = False Then
            If strSt52List <> MsgText(601) Then
                strQ = strQ & " Union " & _
                            "Select st01,st02,2 as Sort From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & "And sp02<>'" & strNowUser & "' " & strWhere & _
                            " Order by Sort,st01"
            End If
        '區主管
        Else
            '簡金泉(69005)協理可看S29,無法知道哪些部門哪些人一定要輸,故SalesPoint有資料才顯示 ex:20091 11010月無資料-不顯示
            'Modify by Amy 2022/05/04 11105月始 林柄佑協理ST15=S29 且為S20/S23/S29部門主管 原:IIf(Left(cboSalesArea(0), 3) = strSP48, "And sp02<> '" & strNowUser & "' ", "") 改為 And sp02<> '" & strNowUser & "' "
            '                                             避免以S20登入選S29時智權人員出現2筆 82026
            strQ = strQ & " Union " & _
                        "Select st01,st02,2 as Sort  From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1(0) & Text1(1)) + 191100 & " And sp48='" & Left(cboSalesArea(0), 3) & "' " & _
                        "And sp02<> '" & strNowUser & "' " & _
                         " Order by Sort,st01"
        End If
        If strQ <> MsgText(601) Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strQ)
            If intI = 1 Then
                RsTemp.MoveFirst
                Do While Not RsTemp.EOF
                    '初始設定時,預設自己
                    If intSet = 0 And RsTemp.Fields("Sort") = "1" Then
                        stDef = RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                    '查詢且人員於期間內有資料才預設
                    ElseIf intSet = 1 And stDef <> MsgText(601) Then
                        If Left(stDef, 5) = RsTemp.Fields("st01") Then
                            bolHasData = True
                        End If
                    End If
                    CboEmp.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                    RsTemp.MoveNext
                Loop
                '查詢時,此人員於期間內無資料,不預設 ex:張宜萱(A8007) 10906月以前都沒資料
                If intSet = 1 And bolHasData = False Then stDef = ""
                If stDef <> MsgText(601) Then CboEmp = stDef '預設
                If RsTemp.RecordCount > 0 Then
                    CboEmp.Enabled = True
                End If
            End If
        End If
    End If
    
    LblAccept.Visible = False
    If stMsg <> MsgText(601) Then
        MsgBox stMsg
        CboEmp = ""
    '若設定狀態為「查詢」,且部門、智權人員、業績年月不是空,則自動執行查詢
    ElseIf intSet >= 1 And Trim(cboSalesArea(0)) <> MsgText(601) And Trim(CboEmp) <> MsgText(601) And Trim(Text1(0) & Text1(1)) <> MsgText(601) Then
        '+bolNoChkMod 避免無窮迴圈
        bolNoChkMod = True
        Call cmdSearch_Click(0)
        bolNoChkMod = False
    End If

    bolNowChk = False
End Sub

Private Sub ChkSetDept(ByVal stNo As String)
    Dim strTp(1) As String
    
    If stNo = MsgText(601) Then Exit Sub
        
    strTp(1) = GetST15(stNo, strTp(0), Val(Text1(0) & Text1(1)) + 191100)
        
    For i = 1 To cboSalesArea(0).ListCount - 1
        If Mid(cboSalesArea(0).List(i), 1, InStr(cboSalesArea(0).List(i), " ") - 1) = strTp(1) Then
            cboSalesArea(0) = cboSalesArea(0).List(i)
            Exit For
        End If
    Next i
End Sub

'Add by Amy 2024/07/10 從doQuery1-全區資料搬過來修改
Private Sub InsertR210152(ByVal strQDate As String)
   Dim RsQ As ADODB.Recordset, strQ As String, intQ As Integer, stSQL As String
   
   'Modfiy by Amy 2016/04/07 +bolGlMan參數(讓財務可看特殊編號人員)
   'Modify by Amy 2019/08/01 原財務可看特殊編號人員,改F4102王文安可操作,F4103陳鳳英及其職代A0914可操作
   'Moidfy by Amy 2019/10/16 開放W部門區主管輸該區編號
   'Moidfy by Amy 2021/04/29 開放林純真可操作P2005
   'Modify by Amy 2021/06/03 財務登入strSP48為M31導致顏色未變色,加管理部門顯示判斷
   stSQL = "Insert Into R210152 (ID,R01,R02,R03,R04,R05,R06,R07,R08,R09) "
   If InStr(Replace(智權點數實績與結餘輸入部門, "S", ""), Left(cboSalesArea(1), 1)) > 0 _
      Or (Left(strSP48, 1) = "M" And InStr(Replace(智權點數實績與結餘輸入部門, "S", ""), Left(cboSalesArea(1), 1)) > 0) Then
      stSQL = stSQL & GetPoint(1, strQDate, strQDate, Left(cboSalesArea(1), 3), Left(cboSalesArea(1), 3), , , Me.Name, True)
   Else
      stSQL = stSQL & GetPoint(1, strQDate, strQDate, Left(cboSalesArea(1), 3), Left(cboSalesArea(1), 3), , , Me.Name, bolGlMan)
   End If
   'end 2019/08/01
   cnnConnection.Execute stSQL
   
   'Add by Amy 2024/07/10 由電腦中心以語法加人,全區會沒此人資料
   'ex:11306月75007魏天立 數點全轉給10011北一區,10011 當月無任何4字頭傳票資料,但又要輸轉撥,電腦中心以語法加10011
   strQ = "Select * From SalesPoint Where sp01=" & Val(strQDate) + 191100 & " And sp48='" & Left(cboSalesArea(1), 3) & "' " & _
                  " And sp02 not in (Select R01 From R210152 Where ID='" & strUserNum & "') "
   intQ = 1 'Add by Amy 2024/12/23 無資料會彈訊息 ex:杜協理進入此支直接切全區資料
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      Do While Not RsQ.EOF
         stSQL = "Insert Into R210152 (ID,R01,R03,R04,R05,R06,R07,R08,R09) " & _
                     "Values('" & strUserNum & "','" & RsQ.Fields("SP02") & "',0,0,0,0,0,0,0)"
         cnnConnection.Execute stSQL
         RsQ.MoveNext
      Loop
   End If
   
   'Add by Amy 2021/08/03 11007月有A6034資料,選擇F11部門時應該不出現(因不用輸)
   If Left(cboSalesArea(1), 1) <> "S" Then
      If Left(cboSalesArea(1), 3) = "P10" And strUserNum = "71011" Then
         'P11部門只顯示王副總和P1001(因王副總只能看自己和P1001)
         stSQL = "Delete R210152 Where ID='" & strUserNum & "' And R01 Not In('71011','P1001') "
      Else
         stSQL = "Delete R210152 Where ID='" & strUserNum & "' And R01 Not In('" & Replace(智權點數實績與結餘特殊員編, ";", "','") & "') "
      End If
      cnnConnection.Execute stSQL
   End If
End Sub




