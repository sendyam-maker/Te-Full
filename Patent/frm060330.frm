VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060330 
   BorderStyle     =   1  '單線固定
   Caption         =   "DHL列印"
   ClientHeight    =   5340
   ClientLeft      =   156
   ClientTop       =   1620
   ClientWidth     =   8460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8460
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "測試API連線"
      Height          =   375
      Left            =   2688
      Style           =   1  '圖片外觀
      TabIndex        =   77
      Top             =   72
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3675
      Left            =   90
      TabIndex        =   36
      Top             =   630
      Width           =   8325
      _ExtentX        =   14690
      _ExtentY        =   6477
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "本所案號/X/Y/R"
      TabPicture(0)   =   "frm060330.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblNation1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAdd(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAdd(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAdd(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAdd(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtAdd(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAdd(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdQN(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "人工填提單"
      TabPicture(1)   =   "frm060330.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdQN(1)"
      Tab(1).Control(1)=   "txtField(7)"
      Tab(1).Control(2)=   "txtField(6)"
      Tab(1).Control(3)=   "txtField(5)"
      Tab(1).Control(4)=   "txtField(4)"
      Tab(1).Control(5)=   "txtField(3)"
      Tab(1).Control(6)=   "txtField(2)"
      Tab(1).Control(7)=   "txtField(1)"
      Tab(1).Control(8)=   "txtField(0)"
      Tab(1).Control(9)=   "Label4(18)"
      Tab(1).Control(10)=   "lblNation2"
      Tab(1).Control(11)=   "Label8(7)"
      Tab(1).Control(12)=   "Label8(6)"
      Tab(1).Control(13)=   "Label4(14)"
      Tab(1).Control(14)=   "Label8(5)"
      Tab(1).Control(15)=   "Label8(4)"
      Tab(1).Control(16)=   "Label4(8)"
      Tab(1).Control(17)=   "Label8(3)"
      Tab(1).Control(18)=   "Label8(2)"
      Tab(1).Control(19)=   "Label8(1)"
      Tab(1).Control(20)=   "Label8(0)"
      Tab(1).ControlCount=   21
      Begin VB.CommandButton cmdQN 
         Caption         =   "？"
         Height          =   285
         Index           =   1
         Left            =   -74070
         TabIndex        =   75
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton cmdQN 
         Caption         =   "？"
         Height          =   285
         Index           =   0
         Left            =   5130
         TabIndex        =   74
         Top             =   3270
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame3 
         Caption         =   "申請人/代理人/潛在客戶"
         Height          =   2055
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   4095
         Begin VB.TextBox txtCNo 
            Height          =   300
            Left            =   1320
            MaxLength       =   9
            TabIndex        =   6
            Top             =   275
            Width           =   1215
         End
         Begin VB.OptionButton OptKind 
            Caption         =   "本所案號："
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton OptKind 
            Caption         =   "非案件說明："
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtNO 
            Height          =   300
            Index           =   3
            Left            =   3270
            MaxLength       =   2
            TabIndex        =   11
            Top             =   960
            Width           =   360
         End
         Begin VB.TextBox txtNO 
            Height          =   300
            Index           =   2
            Left            =   2925
            MaxLength       =   1
            TabIndex        =   10
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox txtNO 
            Height          =   300
            Index           =   1
            Left            =   2130
            MaxLength       =   6
            TabIndex        =   9
            Top             =   960
            Width           =   720
         End
         Begin VB.TextBox txtNO 
            Height          =   300
            Index           =   0
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   8
            Top             =   960
            Width           =   495
         End
         Begin MSForms.TextBox Text2 
            Height          =   300
            Left            =   1560
            TabIndex        =   13
            Top             =   1320
            Width           =   2415
            VariousPropertyBits=   671105051
            Size            =   "4260;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text3 
            Height          =   300
            Left            =   1560
            TabIndex        =   14
            Top             =   1635
            Width           =   2415
            VariousPropertyBits=   671105051
            Size            =   "4260;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblFM2 
            Height          =   225
            Left            =   1380
            TabIndex        =   76
            Top             =   615
            Width           =   2655
            Caption         =   "lblFM2"
            Size            =   "4683;397"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "編　　號："
            Height          =   180
            Index           =   0
            Left            =   315
            TabIndex        =   47
            Top             =   315
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "名　　稱："
            Height          =   180
            Index           =   1
            Left            =   315
            TabIndex        =   46
            Top             =   615
            Width           =   900
         End
         Begin VB.Line Line4 
            X1              =   2985
            X2              =   3405
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line5 
            X1              =   2055
            X2              =   2115
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label4 
            Caption         =   "聯絡人："
            Height          =   180
            Index           =   3
            Left            =   696
            TabIndex        =   45
            Top             =   1680
            Width           =   744
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "本所案號"
         Height          =   975
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   4095
         Begin VB.OptionButton Option1 
            Caption         =   "CF代理人"
            Height          =   255
            Index           =   1
            Left            =   2085
            TabIndex        =   5
            Top             =   600
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Caption         =   "FC代理人"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   4
            Top             =   600
            Width           =   1365
         End
         Begin VB.TextBox txt1 
            Height          =   300
            Index           =   3
            Left            =   2955
            MaxLength       =   2
            TabIndex        =   3
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txt1 
            Height          =   300
            Index           =   2
            Left            =   2610
            MaxLength       =   1
            TabIndex        =   2
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txt1 
            Height          =   300
            Index           =   1
            Left            =   1815
            MaxLength       =   6
            TabIndex        =   1
            Top             =   240
            Width           =   720
         End
         Begin VB.TextBox txt1 
            Height          =   300
            Index           =   0
            Left            =   1245
            MaxLength       =   3
            TabIndex        =   0
            Top             =   240
            Width           =   495
         End
         Begin VB.Line Line3 
            X1              =   2670
            X2              =   3090
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line2 
            X1              =   2490
            X2              =   2610
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line1 
            X1              =   1740
            X2              =   1800
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label Label1 
            Caption         =   "本所案號："
            Height          =   180
            Left            =   360
            TabIndex        =   43
            Top             =   270
            Width           =   915
         End
      End
      Begin MSForms.TextBox txtField 
         Height          =   330
         Index           =   7
         Left            =   -73800
         TabIndex        =   72
         Top             =   2865
         Width           =   600
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1058;582"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   330
         Index           =   6
         Left            =   -69780
         TabIndex        =   67
         Top             =   2475
         Width           =   2805
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "4948;582"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   330
         Index           =   5
         Left            =   -73800
         TabIndex        =   66
         Top             =   2475
         Width           =   2805
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "4948;582"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdd 
         Height          =   330
         Index           =   5
         Left            =   5370
         TabIndex        =   20
         Top             =   3240
         Width           =   600
         VariousPropertyBits=   671105055
         MaxLength       =   3
         Size            =   "1058;582"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdd 
         Height          =   330
         Index           =   4
         Left            =   5370
         TabIndex        =   19
         Top             =   2880
         Width           =   2805
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "4948;582"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdd 
         Height          =   330
         Index           =   3
         Left            =   4890
         TabIndex        =   18
         Top             =   2520
         Width           =   3315
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "5847;582"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdd 
         Height          =   330
         Index           =   2
         Left            =   4890
         TabIndex        =   17
         Top             =   2160
         Width           =   3315
         VariousPropertyBits=   671105051
         MaxLength       =   18
         Size            =   "5847;582"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdd 
         Height          =   330
         Index           =   1
         Left            =   4890
         TabIndex        =   16
         Top             =   1800
         Width           =   3315
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5847;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdd 
         Height          =   1335
         Index           =   0
         Left            =   4800
         TabIndex        =   15
         Top             =   420
         Width           =   3405
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "6006;2355"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   330
         Index           =   4
         Left            =   -73800
         TabIndex        =   41
         Top             =   3240
         Width           =   4755
         VariousPropertyBits=   671105051
         Size            =   "8387;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   330
         Index           =   3
         Left            =   -69780
         TabIndex        =   40
         Top             =   2100
         Width           =   2805
         VariousPropertyBits=   671105051
         MaxLength       =   18
         Size            =   "4948;582"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   330
         Index           =   2
         Left            =   -73800
         TabIndex        =   39
         Top             =   2100
         Width           =   2805
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4948;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   1200
         Index           =   1
         Left            =   -73800
         TabIndex        =   38
         Top             =   840
         Width           =   4755
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "8387;2117"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   330
         Index           =   0
         Left            =   -73800
         TabIndex        =   37
         Top             =   405
         Width           =   4755
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "8387;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "(必填)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   18
         Left            =   -69000
         TabIndex        =   73
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblNation2 
         Caption         =   "lblNation2"
         Height          =   180
         Left            =   -73050
         TabIndex        =   71
         Top             =   2934
         Width           =   4995
      End
      Begin VB.Label Label8 
         Caption         =   "收件國家:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   70
         Top             =   2934
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "郵遞區號:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   6
         Left            =   -70890
         TabIndex        =   69
         Top             =   2556
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "(必填)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   14
         Left            =   -69000
         TabIndex        =   68
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "收件城市:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   65
         Top             =   2556
         Width           =   1035
      End
      Begin VB.Label lblNation1 
         Caption         =   "lblNation1"
         Height          =   180
         Left            =   6120
         TabIndex        =   60
         Top             =   3330
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "收件國家:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   5
         Left            =   4320
         TabIndex        =   59
         Top             =   3322
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "郵遞區號:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   4
         Left            =   4320
         TabIndex        =   58
         Top             =   2970
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "收件城市:"
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   3
         Left            =   4320
         TabIndex        =   57
         Top             =   2520
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "電話:"
         Height          =   180
         Index           =   2
         Left            =   4320
         TabIndex        =   56
         Top             =   2220
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "收件聯絡人:"
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   1
         Left            =   4290
         TabIndex        =   55
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label6 
         Caption         =   "收件地址"
         Height          =   525
         Index           =   0
         Left            =   4320
         TabIndex        =   54
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "說　　明:"
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   53
         Top             =   3315
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "(必填，每行３６字元，最多５行)"
         ForeColor       =   &H000000FF&
         Height          =   660
         Index           =   8
         Left            =   -74880
         TabIndex        =   52
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "收件人電話:"
         Height          =   180
         Index           =   3
         Left            =   -70890
         TabIndex        =   51
         Top             =   2178
         Width           =   1125
      End
      Begin VB.Label Label8 
         Caption         =   "收件聯絡人:"
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   50
         Top             =   2178
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "收件地址:"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   49
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "收件公司:"
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.PictureBox tmpBCode 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Free 3 of 9 Extended"
         Size            =   39.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      ScaleHeight     =   672
      ScaleWidth      =   3684
      TabIndex        =   35
      Top             =   5520
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox tmpPic 
      Height          =   495
      Left            =   0
      ScaleHeight     =   444
      ScaleWidth      =   804
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      Begin VB.Image tmpBC 
         Height          =   375
         Left            =   960
         Top             =   0
         Width           =   495
      End
      Begin VB.Image tmpImg2 
         Height          =   375
         Left            =   480
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image tmpImg 
         Height          =   375
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "測試列印"
      Height          =   375
      Left            =   960
      TabIndex        =   24
      Top             =   73
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   240
      TabIndex        =   28
      Top             =   4440
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   21
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   30
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   27
      Top             =   6570
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   29
      Top             =   6870
      Width           =   705
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   5430
      TabIndex        =   23
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   4635
      TabIndex        =   22
      Top             =   60
      Width           =   756
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "3. 若收件地址不是英文，請改為英文地址。"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   64
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "2. 收件城市、郵遞區號、收件國家代碼：不可空白"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   63
      Top             =   4810
      Width           =   4215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "1. 收件聯絡人：若無法確認將會預設為收件公司名稱"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   62
      Top             =   4575
      Width           =   4215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "上述資料提供給DHL有助於寄送，請仔細填寫："
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   61
      Top             =   4350
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "請設定為雷射印表機"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   34
      Top             =   6240
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "畫面上印表機的X及Y偏移值。"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   3
      Left            =   1410
      TabIndex        =   32
      Top             =   8070
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "注意：新人第一次使用此作業功能，須設定"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   2
      Left            =   870
      TabIndex        =   31
      Top             =   7830
      Width           =   3420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   26
      Top             =   6630
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   25
      Top             =   6930
      Width           =   3240
   End
End
Attribute VB_Name = "frm060330"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2024/08/30 MyDHL API整合:資料格式為JSON，需要引用項目：Microsoft WinHTTP Services, version 5.1
'1. MyDHL API和MyDHL +不一樣：沒有UI介面，直接透過http模式request(送出提單)和response(取得提單)，後續追蹤也是透過http模式request / response；
'2. 提單列印：response取得base64字串，透過轉換取得提單pdf。
'--------------------------------
'Memo by Lydia 2022/04/22 DHL輸入資料功能修改過程：
'111/03/29 上線；111/4/11改成支援萬國碼；111/4/13 限制郵遞區號不可空白(與DHL聯繫)；111/4/22 增加DHL_Country_Code(Table)判斷該國是否要輸入郵遞區號
'end 2022/04/22
'Memo by Lydia 2022/04/06 改成Form2.0 ; Label4(2)=>lblFM2、Text2、Text3、txtAdd(index)、txtField(index)
                    '原本因為DHL不接收中文所以不修改Form 2.0 ; 但是111/4/1 May:需要列印萬國碼給人員和國外客戶查看，但是上傳DHL之資料依舊保持英數字
'Memo by Lydia 2022/03/28 增加DHL輸入資料DHL_Input_Data：配合DHL要求，提供介面給使用者設定"收件人城市"、"收件國代碼"，至於"收件人名稱"若無法確定可以重複"對方公司名稱"。
'Add by Lydia 2014/12/26 新增-DHL列印
Option Explicit
Dim strTemp1 As Variant, strNum(0 To 3) As String
Dim StrTmpNick As String, StrTmpNick1 As String, StrTmpNick2 As String
Dim iPrint As Integer, SeekPrint As Integer, SeekPrintL As Integer
Dim i As Integer, j As Integer, s As Integer, poliu As Integer

Dim m_dbl_LeftMargin  As Double '橫軸偏移值
Dim m_dbl_TopMargin  As Double '縱軸偏移值
Dim m_CP09 As String '總收文號
Dim printLH As Integer, printLW As Integer, posX As Integer, posY As Integer
'Added by Lydia 2016/08/01
Dim strPicFileName As String, strPicFileName2 As String '暫存圖檔路徑
Dim wDHL As Double, hDHL As Double '圖片大小
'Move by Lydia 2016/08/09 從PrintData移出
Private Const PayACNO As String = "620330312" '公司帳號
Private Const PayACName = "TAIE INT'L PATENT & LAW OFFICE" '公司名稱
Private Const PayACaddr01 = "9F NO 112 SEC.2" '公司地址1
Private Const PayACaddr02 = "CHANG AN E.ROAD, TAIPEI, TAIWAN" '公司地址2
Private Const PayACZip = "10491"  '公司Zip
Private Const PayACTel = "02 2506-1023" '公司電話
'Added by Lydia 2016/08/09 用於上傳到DHL FTP
Dim recAddr(1 To 3) As String '收件地址
Dim PayContact As String '寄件聯絡人
'Modified by Lydia 2024/08/30 MyDHL API整合: 收件地址長度45
'Private Const lenRA As Integer = 50   'FTP格式: 收件地址長度
Private Const lenRA As Integer = 45
'預設DHL接收內容 : SPS 2.03|SPS||||||N123456789||||DHL寄件帳號|Company Name|Shipper Name|Shipper Address line1|Shipper Address line2|Shipper Address line3|TAIPEI||10491|TAIWAN|TW|Shipper Reference|Shipper Tel||||TWD|N|||Receiver Company Name收件公司|Attention收件聯絡人|Receiver Address line 1收件地址1|Receiver Address line2收件地址2|Receiver Address line3收件地址3|收件城市填空白||收件郵遞區號填.|收件國|國代char(2)|收件人電話char(18)||||||||DOX|DOX|||||1|Y|0.50|||||0|||||||||||||||N||0|USD|||N||0||||||||2|doc ||0|1|N|USD|||||||||||||||||||TWD|||||||||||||||||||3|||||||1||0.5||||||KG|||||3|||||||
'Modified by Lydia 2022/03/28 增加DHL輸入資料DHL_Input_Data
'Private Const defCont = "SPS 2.03|SPS||||||N123456789||||" & PayACNO & "|" & PayACName & "|PayACMan|" & PayACaddr01 & "|" & PayACaddr02 & "||TAIPEI||10491|TAIWAN|TW|Shipper Reference|" & PayACTel & "||||TWD|N|||收件公司|收件聯絡人|收件地址1|收件地址2|收件地址3|收件城市填空白||收件郵遞區號填.|收件國|國代C(2)|收件人電話C(18)||||||||DOX|DOX|SYSD|SYST|||1|Y|0.50|||||0|||||||||||||||N||0|USD|||N||0||||||||2|doc ||0|1|N|USD|||||||||||||||||||TWD|||||||||||||||||||3|||||||1||0.5||||||KG|||||3|||||||"
Private Const defCont = "SPS 2.03|SPS||||||N123456789||||" & PayACNO & "|" & PayACName & "|PayACMan|" & PayACaddr01 & "|" & PayACaddr02 & "||TAIPEI||10491|TAIWAN|TW|Shipper Reference|" & PayACTel & "||||TWD|N|||收件公司|收件聯絡人|收件地址1|收件地址2|收件地址3|收件城市||收件郵遞區號填.|收件國|國代C(2)|收件人電話C(18)||||||||DOX|DOX|SYSD|SYST|||1|Y|0.50|||||0|||||||||||||||N||0|USD|||N||0||||||||2|doc ||0|1|N|USD|||||||||||||||||||TWD|||||||||||||||||||3|||||||1||0.5||||||KG|||||3|||||||"
Dim strFileName As String 'DHL提單資料檔名
Dim strFileDir As String  'DHL ftp 資料夾路徑
Private Const StartFtpDate As String = "20160830" '上傳到DHL FTP的啟用日
Dim bFTPok As Boolean
Dim pTime As String 'Added by Lydia 2016/08/31
Dim saveBCfile As String 'Added by Lydia 2016/08/31
'Memo by Lydia 2022/03/28
'strTemp(1)=收件聯絡人DID03, strTemp(2)=收件公司, strTemp(3~7)=收件地址, strTemp(8)=本所案號或其他參考, strTemp(9)=收件人電話DID06
'strTemp(10)=郵遞區號DID03, strTemp(11)=收件國家名稱DCC04
'Modified by Lydia 2022/03/28
'Dim strTemp(0 To 9) As String
Dim strTemp(0 To 11) As String
'end 2022/03/28
'Added by Lydia 2022/03/28
Dim strNowGrp As String '現在帶入資料的編號
Dim recNA89 As String 'DHL國家代碼
Dim recDID04 As String '收件城市
Dim m_strDCC04 As String 'Added by Lydia 2022/04/22 當地區域代號：描述內容有Postcode，代表有郵遞區號。
Dim m_strDCC02 As String 'Added by Lydia 2024/08/30 英文國家名稱
Dim oText As Control
Dim StrSQLa As String
Dim rsAD As New ADODB.Recordset
Dim intA As Integer
'Added by Lydia 2022/04/06 Word套印: Word需要顯示,不然BarCode圖片會偏移
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim bVisible As Boolean
Const msoFalse = 0 'Added by Lydia 2022/04/12
Const cntAuidPWD = "YXBZNnZGNWRPNGVMNnA6REA5cU1AMmJGXjh1Ql40dg=="  'Added by Lydia 2024/08/30 MyDHL API整合:從Dhl Developer申請到的UserName：apY6vF5dO4eL6p，PWD：D@9qM@2bF^8uB^4v再經過輸入(UserName:PWD)base64加密,取得的Authorization ID+PWD
Dim strPrinter As String 'Added by Lydia 2024/08/30 原本預設印表機

Private Sub cmdok_Click(Index As Integer)
Dim iErr As Integer 'Added by Lydia 2018/01/16
Dim strNoB As String, inTab As Integer  'Added by Lydia 2022/03/28
Dim bolChgPrt As Boolean 'Added by Lydia 2024/08/30

iErr = -1  'Added by Lydia 2022/03/28

Select Case Index
Case 0 '確定
  
  'Added by Lydia 2018/01/16 +人工填提單
  If Me.SSTab1.Tab = 1 Then
        inTab = 1 'Added by Lydia 2022/03/28
        'Modified by Lydia 2024/08/30 從36改到100
        If GetTextLength(txtField(0).Text) > 100 Then
              If MsgBox("收件公司超過100個字元，列印可能會超出欄位，是否要修改？", vbYesNo + vbDefaultButton1) = vbYes Then
                    GoTo JumpExitNum
              End If
        End If
        iErr = 1
        If Trim(txtField(1)) = "" Then
              s = MsgBox("收件地址不可空白!!", , "USER 輸入錯誤")
              GoTo JumpExitNum
        'Added by Lydia 2022/03/28
        ElseIf Trim(txtField(0)) = "" And Trim(txtField(2)) = "" Then
              s = MsgBox("收件公司或聯絡人不可空白!!", , "USER 輸入錯誤")
              If txtField(0) = "" Then
                 iErr = 0
              Else
                 iErr = 2
              End If
              GoTo JumpExitNum
        ElseIf Trim(txtField(5)) = "" Then
              s = MsgBox("收件城市不可空白!!", , "USER 輸入錯誤")
              iErr = 5
              GoTo JumpExitNum
        ElseIf Trim(txtField(7)) = "" Then
              s = MsgBox("收件國家不可空白!!", , "USER 輸入錯誤")
              iErr = 7
              GoTo JumpExitNum
        'end 2022/03/28
        Else
              strTemp1 = Empty
              strTemp1 = Split(txtField(1), vbNewLine)
              j = 0
              strExc(1) = ""
              For i = 0 To UBound(strTemp1)
                   If Trim(strTemp1(i)) <> "" Then
                       j = j + 1
                      If GetTextLength("" & strTemp1(i)) > 36 Then
                          strExc(1) = strExc(1) & IIf(strExc(1) <> "", "、", "") & "第" & j & "行"
                      End If
                   End If
              Next i
              If strExc(1) <> "" Then
                  If MsgBox(strExc(1) & "分別超過36個字元，列印可能會超出欄位，是否要修改？", vbYesNo + vbDefaultButton1) = vbYes Then
                        GoTo JumpExitNum
                  End If
              End If
              If j > 5 Then
                    s = MsgBox("收件地址不可超過5行!!", , "USER 輸入錯誤")
                    GoTo JumpExitNum
              End If
        End If
        iErr = 2
        If GetTextLength(txtField(2).Text) > 20 Then
            If MsgBox("收件聯絡人超過20個字元，列印可能會超出欄位，是否要修改？", vbYesNo + vbDefaultButton1) = vbYes Then
                  GoTo JumpExitNum
            End If
        End If
        iErr = 3
        If Trim(txtField(3)) = "" Then
              'Modified by Lydia 2022/03/28 收件人電話/Email=>收件人電話
              'Modified by Lydia 2022/07/20 改成詢問
              's = MsgBox("收件人電話不可空白!!", , "USER 輸入錯誤")
              If MsgBox("收件人電話空白，是否要修改？", vbYesNo + vbDefaultButton1) = vbYes Then
                   GoTo JumpExitNum
              End If 'Added by Lydia 2022/07/20
        Else
              If GetTextLength(txtField(3).Text) > 20 Then
                  'Modified by Lydia 2022/03/28 收件人電話/Email=>收件人電話
                  If MsgBox("收件人電話超過20個字元，列印可能會超出欄位，是否要修改？", vbYesNo + vbDefaultButton1) = vbYes Then
                        GoTo JumpExitNum
                  End If
              End If
        End If
        iErr = 4
        If GetTextLength(txtField(4).Text) > 30 Then
            If MsgBox("說明超過30個字元，列印可能會超出欄位，是否要修改？", vbYesNo + vbDefaultButton1) = vbYes Then
                  GoTo JumpExitNum
            End If
        'Added by Lydia 2024/09/09
        ElseIf Trim(txtField(4)) = "" Then
            MsgBox "說明不可空白!!", vbInformation
            GoTo JumpExitNum
        'end 2024/09/09
        End If
        'Added by Lydia 2022/03/28 判斷收件城市和郵遞區號是否正確
        iErr = 5
        strExc(2) = UCase(StringFilterr(txtField(1)))
        If InStr(strExc(2), UCase(txtField(5))) = 0 Then
            s = MsgBox("收件城市不在收件地址內，請檢查!!", , "USER 輸入錯誤")
            GoTo JumpExitNum
        End If
        'Added by Lydia 2024/08/30
        If InStr(UCase("HONG KONG,MACAU,SINGAPORE,MONACO,"), UCase(txtField(5))) = 0 Then '排除城邦國家(特別行政區):香港,澳門,新加坡,摩納哥
          If InStr(UCase(m_strDCC02), UCase(txtField(5))) > 0 Then
             If MsgBox("收件城市" & UCase(txtField(5)) & "與英文國家名稱相同，若城市名稱不正確請選擇「否」!", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                GoTo JumpExitNum
             End If
          End If
        Else
          If InStr(UCase(m_strDCC02), UCase(txtField(5))) = 0 Then
             MsgBox "以下收件城市必須國家名稱相同：" & vbCrLf & "城邦國家(特別行政區)：香港、澳門、新加坡、摩納哥", vbCritical
             GoTo JumpExitNum
          End If
        End If
        'end 2024/08/30
        'Added by Lydia 2025/02/19 瓜地馬拉首都(瓜地馬拉市Guatemala City)，DHL端會發生錯誤
        If UCase(Trim(txtField(5))) = UCase("Guatemala City") Then
           MsgBox "收件城市名稱不用包含CITY ！", vbCritical
           GoTo JumpExitNum
        End If
        'end 2025/02/19
        
        iErr = 6
        'Added by Lydia 2022/04/13 DHL: 為了運務遞送順暢，建議每筆貨件需提供郵遞區號
        'Modified by Lydia 2024/10/18 越南為非必要郵遞區號的國家,但是又輸入郵遞區號 --- from 玫音  +Or Trim(txtField(6)) <> ""
        If InStr(UCase(m_strDCC04), "POSTCODE") > 0 Or Trim(txtField(6)) <> "" Then  'Added by Lydia 2022/04/22 排除不用郵遞區號的國家
          If Trim(txtField(6)) = "" Or Trim(txtField(6)) = "." Then
              s = MsgBox("尚未輸入郵遞區號，請檢查!!", , "USER 輸入錯誤")
              GoTo JumpExitNum
          End If
          'end 2022/04/13
           'Added by Lydia 2024/08/30 檢查郵遞區號格式是否正確
           strExc(5) = GhgPostFormat(txtField(6))
           If strExc(5) <> "" Then
              strSql = "select POSTALCODEFORMAT from DHL_POSTCODE WHERE COUNTRYCODE='" & recNA89 & "' and POSTALCODEFORMAT='" & strExc(5) & "' "
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
              If intI = 0 Then
                 strSql = "select POSTALCODEFORMAT from DHL_POSTCODE WHERE COUNTRYCODE='" & recNA89 & "' "
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                 If intI = 1 Then
                     strSql = RsTemp.GetString(adClipString, , , ",")
                     MsgBox "請輸入正確郵遞區號格式如下：" & vbCrLf & Replace(strSql, ",", vbCrLf) & vbCrLf & "9代表輸入數字0~9，A代表輸入英文字母，有-或空白則要保留-或空白。"
                     GoTo JumpExitNum
                 Else
                     MsgBox "沒有郵遞區號格式，請洽電腦中心！"
                     GoTo JumpExitNum
                 End If
              End If
           End If
           'end 2024/08/30
          If InStr(strExc(2), UCase(txtField(6))) = 0 And txtField(6) <> "" And txtField(6) <> "." Then
              'Modified by Lydia 2022/04/13
              'If MsgBox("郵遞區號不在收件地址內，請問是否正確？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
              '    GoTo JumpExitNum
              'End If
              s = MsgBox("郵遞區號不在收件地址內，請檢查!!", , "USER 輸入錯誤")
              GoTo JumpExitNum
          End If
        'end 2022/03/28
        End If 'Added by Lydia 2022/04/22
        iErr = 0
  Else
  'end 2018/01/16
      
JumpReCheck: 'Added by Lydia 2022/03/28

        'Added by Lydia 2022/03/28 判斷重新輸入,需要再次檢查
        inTab = 0
        strNoB = strNowGrp
        Call GetNowGrp
        If strNoB <> strNowGrp Then
            GoTo JumpReCheck
        End If
        'end 2022/03/28
        If Option1(0).Value = True Or Option1(1).Value = True Then
           If Len(txt1(0)) = 0 Then
              s = MsgBox("第一欄位不可空白!!", , "USER 輸入錯誤")
              txt1(0).SetFocus
              Exit Sub
           Else
               strNum(0) = txt1(0)
               If Len(txt1(1)) = 0 Then
                   s = MsgBox("第二欄位不可空白!!", , "USER 輸入錯誤")
                   txt1(1).SetFocus
                   Exit Sub
               Else
                   strNum(1) = txt1(1)
                   If Len(txt1(2)) = 0 Then
                       strNum(2) = "0"
                   Else
                       strNum(2) = txt1(2)
                   End If
                   If Len(txt1(3)) = 0 Then
                       strNum(3) = "00"
                   Else
                       strNum(3) = txt1(3)
                   End If
               End If
           End If
        Else
           If OptKind(0).Value = True Or OptKind(1).Value = True Then
              If LTrim(RTrim(txtCNo)) = "" Then
                 MsgBox "申請人/代理人/潛在客戶只能為 X、Y 或 R !!", , "USER 輸入錯誤"
                 txtCNo_GotFocus
                 Exit Sub
              End If
              If OptKind(0).Value = True Then
                  If txtNo(2) = "" Then txtNo(2) = "0"
                  If txtNo(3) = "" Then txtNo(3) = "00"
                  If ClsPDCheckCaseCodeIsExist(txtNo(0), txtNo(1), txtNo(2), txtNo(3)) = False Then
                    txtNo(1).SetFocus
                    Exit Sub
                  End If
              End If
              
              If OptKind(1).Value = True Then
                 If Len(Text2) = 0 Then
                    'Modified by Lydia 2015/12/16
                    'MsgBox "非案件說明不可空白!!", , "USER 輸入錯誤"
                    MsgBox "非案件說明會列印在案號的位置,方便日後追蹤,所以不可空白!!", vbInformation
                    Text2.SetFocus
                    Exit Sub
                 End If
              End If
           Else
              MsgBox "請選擇條件範圍 !!", , "USER 輸入錯誤"
              If Len(txt1(0)) > 0 Then
                 txt1(0).SetFocus
              Else
                 txtCNo.SetFocus
              End If
              Exit Sub 'Added by Lydia 2022/03/28
           End If
        End If
        
        'Added by Lydia 2022/04/27 增加檢查字串長度; ex. 電話設MaxLeng=18, 但是讀取Y55419000的電話長度20字可完全顯示
        For intI = 1 To 4
            If CheckLengthIsOK(txtAdd(intI), txtAdd(intI).MaxLength, , , True) = False Then
                 iErr = intI
                 GoTo JumpExitNum
            End If
        Next intI
        'end 2022/04/27
        'Added by Lydia 2022/03/28 判斷收件城市和郵遞區號是否正確
        strExc(2) = UCase(StringFilterr(txtAdd(0)))
        If Left(strNowGrp, 1) = "X" Then
            strExc(3) = "〔客戶基本檔" & strNowGrp & "〕"
        ElseIf Left(strNowGrp, 1) = "Y" Then
            strExc(3) = "〔代理人基本檔" & strNowGrp & "〕"
        Else
            strExc(3) = "〔潛在客戶基本檔" & strNowGrp & "〕"
        End If
        If strExc(2) = "" Or txtAdd(5) = "" Then
            If strExc(2) = "" Then
                MsgBox "收件地址不可空白，請回到" & strExc(3) & "補上英文地址!!!", vbCritical
            Else
                MsgBox "地址國籍不可空白，請回到" & strExc(3) & "補上地址國籍!!!", vbCritical
                iErr = 5
            End If
            Exit Sub
        End If
        iErr = 1
        If Trim(txtAdd(3)) = "" Then
             s = MsgBox("收件城市不可空白!!", , "USER 輸入錯誤")
             GoTo JumpExitNum
        End If
        iErr = 3
        If InStr(strExc(2), UCase(txtAdd(3))) = 0 Then
            s = MsgBox("收件城市不在收件地址內，請檢查!!", , "USER 輸入錯誤")
            GoTo JumpExitNum
        End If
        'Added by Lydia 2024/08/30
        If InStr(UCase("HONG KONG,MACAU,SINGAPORE,MONACO,"), UCase(txtAdd(3))) = 0 Then '排除城邦國家(特別行政區)：香港、澳門、新加坡、摩納哥
          If InStr(UCase(m_strDCC02), UCase(txtAdd(3))) > 0 Then
             If MsgBox("收件城市" & UCase(txtAdd(3)) & "與英文國家名稱相同，若城市名稱不正確請選擇「否」!", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                GoTo JumpExitNum
             End If
          End If
        Else
          If InStr(UCase(m_strDCC02), UCase(txtAdd(3))) = 0 Then
             MsgBox "以下收件城市必須國家名稱相同：" & vbCrLf & "城邦國家(特別行政區)：香港、澳門、新加坡、摩納哥", vbCritical
             GoTo JumpExitNum
          End If
        End If
        'end 2024/08/30
        'Added by Lydia 2025/02/19 瓜地馬拉首都(瓜地馬拉市Guatemala City)，DHL端會發生錯誤
        If UCase(Trim(txtAdd(3))) = UCase("Guatemala City") Then
           MsgBox "收件城市名稱不用包含CITY ！", vbCritical
           GoTo JumpExitNum
        End If
        'end 2025/02/19
        
        iErr = 4
        'Added by Lydia 2022/04/13 DHL: 為了運務遞送順暢，建議每筆貨件需提供郵遞區號
        'Modified by Lydia 2024/10/18 越南為非必要郵遞區號的國家,但是又輸入郵遞區號 --- from 玫音 +Or Trim(txtAdd(4)) <> ""
        If InStr(UCase(m_strDCC04), "POSTCODE") > 0 Or Trim(txtAdd(4)) <> "" Then  'Added by Lydia 2022/04/22 排除不用郵遞區號的國家
           If Trim(txtAdd(4)) = "" Or Trim(txtAdd(4)) = "." Then
              s = MsgBox("尚未輸入郵遞區號，請檢查!!", , "USER 輸入錯誤")
              GoTo JumpExitNum
           End If
           'end 2022/04/13
           'Added by Lydia 2024/08/30 檢查郵遞區號格式是否正確
           strExc(5) = GhgPostFormat(txtAdd(4))
           If strExc(5) <> "" Then
              strSql = "select POSTALCODEFORMAT from DHL_POSTCODE WHERE COUNTRYCODE='" & recNA89 & "' and POSTALCODEFORMAT='" & strExc(5) & "' "
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
              If intI = 0 Then
                 strSql = "select POSTALCODEFORMAT from DHL_POSTCODE WHERE COUNTRYCODE='" & recNA89 & "' "
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                 If intI = 1 Then
                     strSql = RsTemp.GetString(adClipString, , , ",")
                     MsgBox "請輸入正確郵遞區號格式如下：" & vbCrLf & Replace(strSql, ",", vbCrLf) & vbCrLf & "9代表輸入數字0~9，A代表輸入英文字母，有-或空白則要保留-或空白。"
                     GoTo JumpExitNum
                 Else
                     MsgBox "沒有郵遞區號格式，請洽電腦中心！"
                     GoTo JumpExitNum
                 End If
              End If
           End If
           'end 2024/08/30
           If InStr(strExc(2), UCase(txtAdd(4))) = 0 And txtAdd(4) <> "" And txtAdd(4) <> "." Then
               'Modified by Lydia 2022/04/13
               'If MsgBox("郵遞區號不在收件地址內，請問是否正確？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
               '    GoTo JumpExitNum
               'End If
                s = MsgBox("郵遞區號不在收件地址內，請檢查!!", , "USER 輸入錯誤")
               GoTo JumpExitNum
           End If
        'Added by Lydia 2024/08/30
        Else
         
        End If  'Added by Lydia 2022/04/22
        iErr = 0
        'end 2022/03/28
  End If 'end 2018/01/16
      
      If strSrvDate(1) <= "20240908" Then 'Added by Lydia 2024/08/30
         'Modified by Lydia 2016/08/05 改成模組
         PUB_RestorePrinter Combo1
         'end 2016/08/05
         Printer.Orientation = 1
      'Added by Lydia 2024/08/30
         bolChgPrt = True
      End If
      'end 2024/08/30
      
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
    
      'Modified by Lydia 2022/03/28 增加DHL輸入資料功能: 先存檔,後列印
      'ClearQueryLog (Me.Name)
      'Added by Lydia 2018/01/16 +人工填提單
      'If Me.SSTab1.Tab = 1 Then
      '     ProcessNew2
      'Else
      ''end 2018/01/16
      '     ProcessNew
      'End If 'end 2018/01/16
      If SaveDID(inTab) = True Then
         'Added by Lydia 2024/08/30
         strExc(1) = "1"
         If Pub_StrUserSt03 = "M51" Then
JumpToReQ:
            strExc(1) = InputBox("請選擇1-DHL正式環境，2-DHL測試環境", "MyDHL API整合", "2")
            If strExc(1) <> "1" And strExc(1) <> "2" Then
               GoTo JumpToReQ
            End If
         End If
         If PrintDataFromAPI(strExc(1)) = False Then
            If strSrvDate(1) <= "20240908" Then  '考慮是否改用原本的FTP
         'end 2024/08/30
               PrintA4
            End If
         End If
      End If
      'end 2022/03/28
      
      '改成列印完就還原(原來Unload才做而在那之前有其他畫面要列印時會抓錯印表機)
      If strSrvDate(1) <= "20240908" Then  'Added by Lydia 2024/08/30
         Set Printer = Printers(SeekPrint)
         Printer.Orientation = SeekPrintL
      End If
      
      '初始化收文號
      m_CP09 = ""
      
      bolToEndByNick = True
      Me.Enabled = True
      Screen.MousePointer = vbDefault
     
Case 1 '結束
      '若有變動印表機或偏移值, 則更新列印設定
      If Me.Combo1.Text <> Me.Combo1.Tag Or Me.Text1(0).Text <> Me.Text1(0).Tag Or Me.Text1(1).Text <> Me.Text1(1).Tag Then
         PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, Me.Text1(0).Text, Me.Text1(1).Text, Me.Combo1.Text
      End If
      bolToEndByNick = True
      Unload Me
Case Else
End Select

'Added by Lydia 2018/01/16
Exit Sub

JumpExitNum:
    'Added by Lydia 2022/03/28
    If inTab = 0 Then
        Me.SSTab1.Tab = inTab
        If iErr >= 0 Then
           txtAdd(iErr).SetFocus
           txtAdd_GotFocus iErr
        End If
    ElseIf inTab = 1 Then
        Me.SSTab1.Tab = inTab
        If iErr >= 0 Then
    'end 2022/03/28
            txtField(iErr).SetFocus
            txtField_GotFocus iErr
        End If 'Added by Lydia 2022/03/28
    End If 'Added by Lydia 2022/03/28
'end 2018/01/16
End Sub

Private Sub SetCP09()

On Error GoTo ErrHnd

   strSql = "Select * From CaseProgress Where " & ChgCaseprogress(strNum(0) & strNum(1) & strNum(2) & strNum(3)) & " And CP09 < 'C' AND CP27>0 and not (CP01='CFT' and cp09>'B' and cp10='304') and CP44 is not null Order By CP27 Desc,CP09 Desc "
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         m_CP09 = "" & .Fields("CP09").Value
      Else
         m_CP09 = ""
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'規則
'  FC:聯絡人=>1.基本檔-->2.代理人檔(PA75,TM44)-->3.客戶檔(PA26,TM23)
'     名稱&地址=>1.代理人檔(PA75,TM44)-->2.客戶檔(PA26,TM23)
'  CF:聯絡人=>代理人檔(CP44)
'     名稱&地址=>代理人檔(CP44)
Private Sub ProcessNew()
   
'”申請人/代理人/潛在客戶”選項。
If Option1(0).Value = True Or Option1(1).Value = True Then
   If Option1(1).Value = True Then
      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption
      '若是直接下TNT列印, 非發文時列印TNT
      If m_CP09 = "" Then SetCP09
      If m_CP09 = "" Then
         MsgBox "找不到AB類的發文資料！", vbExclamation
         Exit Sub
      End If
      strSql = "SELECT NVL(FA08,FA53) C00" & _
         ", FA05||' '||FA63||' '||FA64||' '||FA65 C01" & _
         ", FA18 C02,FA19 C03,FA20 C04,FA21 C05,FA22||rtrim(' '||FA70) C06" & _
         ", NVL(FA12,FA13) C07" & _
         " FROM CASEPROGRESS,FAGENT" & _
         " WHERE CP09='" & m_CP09 & "'" & _
         " AND FA01(+)=SUBSTR(CP44,1,8) AND FA02(+)=SUBSTR(CP44,9,1)"
   Else
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption
      Select Case CheckSys(Me.txt1(0))

         Case "1"
            'TNT行數不夠英文地址5,6合併)
            strSql = "SELECT NVL(NVL(NVL(NVL(NVL(PA52,PA55),FA08),FA53),CU59),CU62) C00" & _
               ", DECODE(PA75,NULL,CU05||' '||CU88||' '||CU89||' '||CU90,FA05||' '||FA63||' '||FA64||' '||FA65) C01" & _
               ", DECODE(PA75,NULL,CU24,FA18) C02, DECODE(PA75,NULL,CU25,FA19) C03" & _
               ", DECODE(PA75,NULL,CU26,FA20) C04, DECODE(PA75,NULL,CU27,FA21) C05, DECODE(PA75,NULL,CU28||rtrim(' '||cu102),FA22||rtrim(' '||FA70)) C06" & _
               ", DECODE(PA75,NULL,NVL(CU16,CU17),NVL(FA12,FA13)) C07" & _
               " FROM PATENT,CUSTOMER,FAGENT" & _
               " WHERE PA01='" & strNum(0) & "' AND PA02='" & strNum(1) & "' AND PA03='" & strNum(2) & "' AND PA04='" & strNum(3) & "' " & _
               " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
               " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)"
            
         Case "5", "6", "7", "8"
            strSql = "SELECT DECODE(SP30,NULL,NVL(NVL(NVL(FA08,FA53),CU59),CU62),SP71||' '||SP30) C00" & _
               ", DECODE(SP26,NULL,CU05||' '||CU88||' '||CU89||' '||CU90,FA05||' '||FA63||' '||FA64||' '||FA65) C01" & _
               ", DECODE(SP26,NULL,CU24,FA18) C02, DECODE(SP26,NULL,CU25,FA19) C03" & _
               ", DECODE(SP26,NULL,CU26,FA20) C04, DECODE(SP26,NULL,CU27,FA21) C05, DECODE(SP26,NULL,CU28||rtrim(' '||cu102),FA22||rtrim(' '||FA70)) C06" & _
               ", DECODE(SP26,NULL,NVL(CU16,CU17),NVL(FA12,FA13)) C07" & _
               " FROM SERVICEPRACTICE,CUSTOMER,FAGENT" & _
               " WHERE SP01='" & strNum(0) & "' AND SP02='" & strNum(1) & "' AND SP03='" & strNum(2) & "' AND SP04='" & strNum(3) & "' " & _
               " AND CU01(+)=SUBSTR(SP08,1,8) AND CU02(+)=SUBSTR(SP08,9,1)" & _
               " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9,1)"
            
         Case "2"
            strSql = "SELECT NVL(NVL(NVL(NVL(NVL(TM39,TM42),FA08),FA53),CU59),CU62) C00" & _
               ", DECODE(TM44,NULL,CU05||' '||CU88||' '||CU89||' '||CU90,FA05||' '||FA63||' '||FA64||' '||FA65) C01" & _
               ", DECODE(TM44,NULL,CU24,FA18) C02, DECODE(TM44,NULL,CU25,FA19) C03" & _
               ", DECODE(TM44,NULL,CU26,FA20) C04, DECODE(TM44,NULL,CU27,FA21) C05, DECODE(TM44,NULL,CU28||rtrim(' '||cu102),FA22||rtrim(' '||FA70)) C06" & _
               ", DECODE(TM44,NULL,NVL(CU16,CU17),NVL(FA12,FA13)) C07" & _
               " FROM TRADEMARK,CUSTOMER,FAGENT" & _
               " WHERE TM01='" & strNum(0) & "' AND TM02='" & strNum(1) & "' AND TM03='" & strNum(2) & "' AND TM04='" & strNum(3) & "' " & _
               " AND CU01(+)=SUBSTR(TM23,1,8) AND CU02(+)=SUBSTR(TM23,9,1)" & _
               " AND FA01(+)=SUBSTR(TM44,1,8) AND FA02(+)=SUBSTR(TM44,9,1)"
      End Select
   End If
   pub_QL05 = pub_QL05 & ";" & Label1 & strNum(0) & "-" & strNum(1) & "-" & strNum(2) & "-" & strNum(3) 'Add By Sindy 2010/10/22
   
'申請人/代理人/潛在客戶
ElseIf OptKind(0).Value = True Or OptKind(1).Value = True Then

   strExc(0) = Left(LTrim(RTrim(txtCNo)) & "000000000", 9)
   strExc(1) = Left(txtCNo, 1)

   Select Case strExc(1)
        Case "X"
            pub_QL05 = pub_QL05 & ";申請人:" & LTrim(RTrim(txtCNo))
            'Modified by Lydia 2015/12/16 若無英文資料,改抓中文->日文
          '  strSql = "SELECT NVL(CU59,CU62) C00, CU05||' '||CU88||' '||CU89||' '||CU90 C01, CU24 C02,CU25 C03,CU26 C04,CU27 C05," & _
                     "CU28||rtrim(' '||CU102) C06, NVL(CU16,CU17) C07 FROM CUSTOMER WHERE " & _
                     "CU01=SUBSTR('" & strExc(0) & "',1,8) AND CU02(+)=SUBSTR('" & strExc(0) & "',9,1)"
            strSql = "SELECT DECODE(CU59||CU62,NULL,DECODE(CU58||CU61,NULL,NVL(CU60,CU63),NVL(CU58,CU61)) ,NVL(CU59,CU62)) C00, " & _
                     "DECODE(CU05||CU88||CU89||CU90,NULL,NVL(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90) C01, " & _
                     "NVL(CU24,NVL(SUBSTR(CU23,1,18),SUBSTR(CU29,1,18))) C02, " & _
                     "NVL(CU25,NVL(SUBSTR(CU23,19,36),SUBSTR(CU29,19,36))) C03, " & _
                     "NVL(CU26,NVL(SUBSTR(CU23,37,54),SUBSTR(CU29,37,54))) C04, " & _
                     "NVL(CU27,NVL(SUBSTR(CU23,55,72),SUBSTR(CU29,55,72))) C05, " & _
                     "NVL(CU28,NVL(SUBSTR(CU23,73,80),SUBSTR(CU29,73,80)))||rtrim(' '||CU102) C06, " & _
                     "NVL(CU16,CU17) C07 FROM CUSTOMER WHERE " & _
                     "CU01=SUBSTR('" & strExc(0) & "',1,8) AND CU02(+)=SUBSTR('" & strExc(0) & "',9,1)"
        Case "Y"
            pub_QL05 = pub_QL05 & ";代理人:" & LTrim(RTrim(txtCNo))
            'Modified by Lydia 2015/12/16 若無英文資料,改抓中文->日文
            'strSql = "SELECT NVL(FA08,FA53) C00, FA05||' '||FA63||' '||FA64||' '||FA65 C01, FA18 C02,FA19 C03,FA20 C04,FA21 C05," & _
                     "FA22||rtrim(' '||FA70) C06, NVL(FA12,FA13) C07 FROM FAGENT WHERE " & _
                     "FA01=SUBSTR('" & strExc(0) & "',1,8) AND FA02=SUBSTR('" & strExc(0) & "',9,1)"
            'Added by Lydia 2015/12/28 特定代理人指定抓中文
            'Modified by Lydia 2017/10/16 改成共用變數
            'If InStr("Y53541,Y52268", ChangeCustomerS(txtCNo)) > 0 Then
            'Remove by Lydia 2017/11/1 DHL要求用英文資料
            'If InStr(外翻Y編號, ChangeCustomerS(txtCNo)) > 0 Then
            '    strSql = "SELECT DECODE(FA07||FA52,NULL,DECODE(FA08||FA53,NULL,NVL(FA09,FA54),NVL(FA08,FA53)) ,NVL(FA07,FA52)) C00," & _
                         "DECODE(FA04,NULL,DECODE(FA05||FA63||FA64||FA65,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65),FA04) C01," & _
                         "DECODE(FA17,NULL,NVL(FA18,SUBSTR(FA23,1,18)),SUBSTR(FA17,1,18)) C02," & _
                         "DECODE(FA17,NULL,NVL(FA19,SUBSTR(FA23,19,36)),SUBSTR(FA17,19,36)) C03," & _
                         "DECODE(FA17,NULL,NVL(FA20,SUBSTR(FA23,37,54)),SUBSTR(FA17,37,54)) C04," & _
                         "DECODE(FA17,NULL,NVL(FA21,SUBSTR(FA23,55,72)),SUBSTR(FA17,55,72)) C05," & _
                         "DECODE(FA17,NULL,NVL(FA22,SUBSTR(FA23,73,80))||rtrim(' '||FA70),SUBSTR(FA17,73,80)) C06," & _
                         "NVL(FA12,FA13) C07 FROM FAGENT WHERE " & _
                         "FA01=SUBSTR('" & strExc(0) & "',1,8) AND FA02=SUBSTR('" & strExc(0) & "',9,1)"
            'Else
            'END 2015/12/28
                strSql = "SELECT DECODE(FA08||FA53,NULL,DECODE(FA07||FA52,NULL,NVL(FA09,FA54),NVL(FA07,FA52)) ,NVL(FA08,FA53)) C00, " & _
                         "DECODE(FA05||FA63||FA64||FA65,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65) C01, " & _
                         "NVL(FA18,NVL(SUBSTR(FA17,1,18),SUBSTR(FA23,1,18))) C02, " & _
                         "NVL(FA19,NVL(SUBSTR(FA17,19,36),SUBSTR(FA23,19,36))) C03, " & _
                         "NVL(FA20,NVL(SUBSTR(FA17,37,54),SUBSTR(FA23,37,54))) C04, " & _
                         "NVL(FA21,NVL(SUBSTR(FA17,55,72),SUBSTR(FA23,55,72))) C05, " & _
                         "NVL(FA22,NVL(SUBSTR(FA17,73,80),SUBSTR(FA23,73,80)))||rtrim(' '||FA70) C06, " & _
                         "NVL(FA12,FA13) C07 FROM FAGENT WHERE " & _
                         "FA01=SUBSTR('" & strExc(0) & "',1,8) AND FA02=SUBSTR('" & strExc(0) & "',9,1)"
            'End If 'Remove by Lydia 2017/11/01
        Case "R"
            pub_QL05 = pub_QL05 & ";潛在客戶:" & LTrim(RTrim(txtCNo))
            'Modified by Lydia 2015/12/16 若無英文資料,改抓中文->日文
            'strSql = "SELECT '' C00, PCU03||' '||PCU04||' '||PCU05||' '||PCU06 C01, PCU20 C02,PCU21 C03,PCU22 C04,PCU23 C05," & _
                     "PCU24||rtrim(' '||PCU25) C06, NVL(PCU13,PCU14) C07 FROM POTCUSTOMER WHERE " & _
                     "PCU01=SUBSTR('" & strExc(0) & "',1,8) AND PCU02(+)=SUBSTR('" & strExc(0) & "',9,1) " & _
                     "union SELECT '' C00, POC03 C01, POC10 C02,'' C03,'' C04,'' C05," & _
                     "'' C06, NVL(POC05,POC06) C07 FROM POTCUSTOMER1 WHERE " & _
                     "POC01=SUBSTR('" & strExc(0) & "',1,8) AND POC02(+)=SUBSTR('" & strExc(0) & "',9,1) "
            strSql = "SELECT '' C00, DECODE(PCU03||PCU04||PCU05||PCU06,NULL,NVL(PCU08,PCU07),PCU03||' '||PCU04||' '||PCU05||' '||PCU06) C01," & _
                     "NVL(PCU20,NVL(SUBSTR(PCU27,1,18),SUBSTR(PCU26,1,18))) C02," & _
                     "NVL(PCU21,NVL(SUBSTR(PCU27,19,36),SUBSTR(PCU26,19,36))) C03," & _
                     "NVL(PCU22,NVL(SUBSTR(PCU27,37,54),SUBSTR(PCU26,37,54))) C04," & _
                     "NVL(PCU23,NVL(SUBSTR(PCU27,55,72),SUBSTR(PCU26,55,72))) C05," & _
                     "NVL(PCU24,NVL(SUBSTR(PCU27,73,80),SUBSTR(PCU26,73,80)))||rtrim(' '||PCU25) C06," & _
                     "NVL(PCU13,PCU14) C07 FROM POTCUSTOMER WHERE " & _
                     "PCU01=SUBSTR('" & strExc(0) & "',1,8) AND PCU02(+)=SUBSTR('" & strExc(0) & "',9,1) " & _
                     "union SELECT '' C00, NVL(POC03,DECODE(POC23||POC24||POC25||POC26,NULL,POC27,POC23||' '||POC24||' '||POC25||' '||POC26)) C01," & _
                     "POC10 C02,'' C03,'' C04,'' C05,'' C06, NVL(POC05,POC06) C07 FROM POTCUSTOMER1 WHERE " & _
                     "POC01=SUBSTR('" & strExc(0) & "',1,8) AND POC02(+)=SUBSTR('" & strExc(0) & "',9,1) "
   End Select
   pub_QL05 = pub_QL05 & ";" & strExc(0)
End If

'On Error GoTo ErrHnd 'Remove by Lydia 2016/08/26

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount = 0 Then
         InsertQueryLog (0)
         If Option1(0).Value = True Or Option1(1).Value = True Then
            s = MsgBox("此本所案號搜尋不到!!", vbInformation)
         Else
            s = MsgBox("此申請人/代理人/潛在客戶號碼搜尋不到!!", vbInformation)
         End If
      Else
         InsertQueryLog (.RecordCount)
         '聯絡人
         strTemp(1) = "" & .Fields("C00")
         '英文名稱
         strTemp(2) = "" & .Fields("C01")
         '英文地址
         strTemp(3) = "" & .Fields("C02")
         strTemp(4) = "" & .Fields("C03")
         strTemp(5) = "" & .Fields("C04")
         strTemp(6) = "" & .Fields("C05")
         strTemp(7) = "" & .Fields("C06")
         '電話
         strTemp(9) = "" & (.Fields("C07"))
         '申請人/代理人/潛在客戶
         If Option1(0).Value = True Or Option1(1).Value = True Then
            strTemp(8) = strNum(0) & "-" & strNum(1) & "-" & strNum(2) & "-" & strNum(3)
         Else
            strTemp(8) = Left(LTrim(RTrim(txtCNo)) & "000000000", 9)
            If OptKind(0).Value = True Then
               strTemp(8) = Trim(txtNo(0)) & "-" & Trim(txtNo(1)) & (IIf(Len(txtNo(2)) > 0, "-" & txtNo(2), "-0")) & (IIf(Len(txtNo(3)) > 0, "-" & txtNo(3), "-00"))
            Else
               strTemp(8) = Trim(Text2)
            End If
            
            If Len(Text3) > 0 Then strTemp(1) = LTrim(RTrim(Text3))
            
         End If
         'Modified by Lydia 2016/08/18 上傳到DHL FTP的啟用日
         'Mark by Lydia 2022/03/28
         'If StartFtpDate > strSrvDate(1) Then
         '   PrintData
         'Else
         '   PrintA4
         'End If
         PrintA4
         'end 2022/03/28
      End If
   End With
'Remove by Lydia 2016/08/26 因為ftp失敗,仍要列印
'ErrHnd:

 '  If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Function PosSub(iSys As Integer, Optional aX1 As Integer, Optional aY1 As Integer)
If iSys = 1 Then
   '跳一行
   posY = posY + printLH
Else
   posX = posX + aX1
   posY = posY + aY1
End If
End Function

'Mark by Lydia 2022/03/28 DHL提供的覆寫提單套印
'Sub PrintData()
'Dim strSql As String
'
''poliu = 45 'TNT的列印長度
'poliu = 37
'If pub_OS = "1" Then
'    Printer.Height = 6 * 1440
'    Printer.Width = 12096
''NT 須先結束文件,否則紙張不會用喜好設定
'Else
'   Printer.PaperSize = PUB_GetPaperSize(11)
'End If
'
'strTemp(1) = StrToStr(strTemp(1), 81)
'Printer.Font.Name = "細明體"
'Printer.FontBold = False
'Printer.Font.Size = 10
'
'm_dbl_LeftMargin = CDbl(Me.Text1(0).Text) * 576: m_dbl_TopMargin = CDbl(Me.Text1(1).Text) * 576
'iPrint = 5300 + m_dbl_TopMargin '印寄件人資訊
'posX = 90 + 240 * 2 + m_dbl_LeftMargin
'posY = 950 + m_dbl_TopMargin
'
'Printer.CurrentX = posX + m_dbl_LeftMargin + 580
'Printer.CurrentY = posY + m_dbl_TopMargin - 170
'Printer.Print "v"
'
'Printer.CurrentX = posX + m_dbl_LeftMargin + 850
'Printer.CurrentY = posY + m_dbl_TopMargin
'Printer.Print PayACNO
'
'printLW = Printer.TextWidth("W")
'printLH = 240
'
'Call PosSub(2, printLW * 2, printLH * 5)
'Printer.CurrentX = posX + m_dbl_LeftMargin
'Printer.CurrentY = posY + m_dbl_TopMargin
'Printer.Print PayACNO 'DHL客戶編號
'
'Call PosSub(2, 0, printLH * 2)
'Printer.CurrentX = posX + m_dbl_LeftMargin
'Printer.CurrentY = posY + m_dbl_TopMargin
'Printer.Print strTemp(8) '本所案號
'
'Call PosSub(2, 0, printLH * 2)
'Printer.CurrentX = posX + m_dbl_LeftMargin
'Printer.CurrentY = posY + m_dbl_TopMargin
'Printer.Print PayACName '台一
'
'Call PosSub(2, 0, printLH * 2)
'Printer.CurrentX = posX + m_dbl_LeftMargin
'Printer.CurrentY = posY + m_dbl_TopMargin
'Printer.Print PayACaddr01  '台一地址1
'
'Call PosSub(1)
'Printer.CurrentX = posX + m_dbl_LeftMargin
'Printer.CurrentY = posY + m_dbl_TopMargin
'Printer.Print PayACaddr02  '台一地址2
'
'Call PosSub(2, 0, printLH * 3)
'Printer.CurrentX = posX + m_dbl_LeftMargin
'Printer.CurrentY = posY + m_dbl_TopMargin
'Printer.Print PayACZip  '台一郵遞區號
'
'Printer.CurrentX = 2650 + m_dbl_LeftMargin
'Printer.CurrentY = posY
'Printer.Print PayACTel  '台一電話
'
''英文名稱 strTemp(2) '英文名稱長度範例FCP-050103
'Printer.FontBold = True
'
''英文名稱超出長度,不折行
'Printer.CurrentX = 90 + m_dbl_LeftMargin + 340
'Printer.CurrentY = iPrint + m_dbl_TopMargin
'Printer.Print strTemp(2)
'iPrint = iPrint + 480 '跳行
'
'For j = 3 To 7 '英文地址　strTemp(3) ~ strTemp(7)
'   If LenB(StrConv(strTemp(j), vbFromUnicode)) > poliu Then
'       StrTmpNick = strTemp(j)
'       StrTmpNick2 = StrTmpNick
'       Do While Len(Trim(StrTmpNick)) <> 0
'          StrTmpNick1 = StrToStr(StrTmpNick, poliu / 2)
'          Printer.CurrentX = 90 + m_dbl_LeftMargin + 340
'          Printer.CurrentY = iPrint
'          Printer.Print StrTmpNick1
'          iPrint = iPrint + 240
'          StrTmpNick = Replace(StrTmpNick, StrTmpNick1, "")
'          If StrTmpNick = StrTmpNick2 Then
'             StrTmpNick = Replace(StrTmpNick, Left(StrTmpNick1, Len(StrTmpNick1) - 1), "")
'             StrTmpNick2 = StrTmpNick
'          Else
'             StrTmpNick2 = StrTmpNick
'          End If
'       Loop
'   Else
'      Printer.CurrentX = 90 + m_dbl_LeftMargin + 340
'      Printer.CurrentY = iPrint
'      Printer.Print strTemp(j)
'      iPrint = iPrint + 240
'   End If
'
'Next j
'
'
''列印ContactName
'posX = 90 + 240 + m_dbl_LeftMargin
'posY = 7910 + m_dbl_TopMargin
''strExc(0) = strTemp(1)
'Printer.Font.Size = 9
'Printer.FontBold = False
'Printer.CurrentX = posX
'Printer.CurrentY = posY
'Printer.Print Mid(strTemp(1), 1, 24) '限定ContactName長度
''列印TelNo (與連絡人同一行)
'Printer.Font.Size = 10
'Printer.CurrentX = 2800 + m_dbl_LeftMargin
'Printer.CurrentY = posY
'Printer.Print strTemp(9)
'
'Printer.EndDoc
'ShowPrintOk
'End Sub
'end 2022/03/28

Private Sub cmdQN_Click(Index As Integer)
   '人工填提單可選擇國家代號, 有X/Y/R編號採用帶入暫時不修改
   Me.Tag = ""
   StrSQLa = "select '' v, na01,na03,DCC02,na89 from nation,DHL_COUNTRY_CODE " & _
                   "where nvl(na89,'00') <>'00' and length(na01)='3' and na01> '010' and na89=dcc01(+) order by na02,na01"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, StrSQLa)
   If intA = 1 Then
       Set frm880012.grdDataList.Recordset = rsAD
       Set frm880012.fmParent = Me
       frm880012.iTyp = "5"
       frm880012.Show vbModal
       If Me.Tag <> "" Then
          If Index = 1 Then
             txtField(7).Text = Me.Tag
             Call txtField_Validate(7, False)
          End If
       End If
   End If
End Sub


Private Sub Form_Load()

MoveFormToCenter Me
'Modified by Lydia 2016/08/05 改成共用模組
   'Modified by Lydia 2024/08/30 + strPrinter
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , SeekPrint, Me.Text1(0), Me.Text1(1)
   SeekPrintL = Printer.Orientation
   
lblFM2.Caption = ""

'Added by Lydia 2016/08/18
If Pub_StrUserSt03 <> "M51" Then
   Command1.Visible = False
   Command2.Visible = False 'Added by Lydia 2024/08/30
End If

'Modified by Lydia 2018/01/16 增加人工填提單
'If StartFtpDate > strSrvDate(1) Then
'   Me.Height = 6465
'Else
'   Me.Height = 5210 '隱藏X,Y軸
'   Label3(0).Visible = False: Label3(1).Visible = False: Label3(2).Visible = False: Label3(3).Visible = False
'   Text1(0).Visible = False: Text1(1).Visible = False
'   '啟用日當天刪除點陣印表機的預設值
'   Label5.Visible = True
'   If Combo1.Tag = "" Then
'      MsgBox "請設定印表機為雷射印表機!"
'   End If
'End If
Me.SSTab1.Tab = 0
'end 2018/01/16

'Added by Lydia 2022/03/28
ClearAdd
ClearField
'Added by Lydia 2022/04/12 建立子資料夾
Call Pub_ChkExcelPath(App.path & "\" & strUserNum)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsAD = Nothing 'Added by Lydia 2022/03/28
   Set frm060330 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
    'Added by Lydia 2022/04/28 變更選項重新帶入資料; May: CFT-14769案要寄FC代理人
    If txt1(0) <> "" And txt1(1) <> "" Then
        Call GetNowGrp
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    '反白設定
    TextInverse Me.Text1(Index)
End Sub

Private Sub txt1_GotFocus(Index As Integer)

If txtCNo.Text <> "" Or (OptKind(0).Value = True Or OptKind(1).Value = True) Then
   ClearInputData 2
End If

txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))

CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
   Case 0

      If App.EXEName = "Patent" Then
        If txt1(Index).Enabled = True Then
        Select Case txt1(0)
        Case "FCP", "FG", "CFP", ""
        Case Else
             s = MsgBox("本所案號只能 FCP 或 FG 或 CFP !!!", , "USER 輸入錯誤")
             txt1(0).SetFocus
             txt1(0).SelStart = 0
             txt1(0).SelLength = Len(txt1(0))
             Exit Sub
        End Select
        End If
      ElseIf App.EXEName = "Trademark" Then
        If txt1(Index).Enabled = True Then
        Select Case txt1(0)
        Case "CFT", "FCT", ""
        Case Else
             s = MsgBox("本所案號只能 CFT 或 FCT !!", , "USER 輸入錯誤")
             txt1(0).SetFocus
             txt1(0).SelStart = 0
             txt1(0).SelLength = Len(txt1(0))
             Exit Sub
        End Select
        End If
      End If
   'Added by Lydia 2022/03/28
   Case 1
        Call GetNowGrp
   'end 2022/03/28
   Case Else
   End Select

   '預設代理人種類
   Select Case Me.txt1(0).Text
      Case "FCP", "FG", "FCT"
         Option1(0).Value = True
      Case Else
         Option1(1).Value = True
   End Select

End Sub

Private Sub txtCNo_GotFocus()
If txt1(0).Text <> "" Or (Option1(0).Value = True Or Option1(1).Value = True) Then
   ClearInputData 1
End If

txtCNo.SelStart = 0
txtCNo.SelLength = Len(txtCNo)
CloseIme
End Sub

Private Sub OptKind_Click(Index As Integer)
If Index = 0 Then
   Text2.Text = ""
   Text3.Text = ""
   txtNo(0).SetFocus
Else
   For intA = 0 To 3
       txtNo(intA).Text = ""
   Next intA
   Text2.SetFocus
End If

End Sub
Private Sub txtCNo_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCNo_LostFocus()
Dim clname As String

If LTrim(RTrim(txtCNo)) <> "" Then
   'Modified by Lydia 2022/03/28
   'clname = "求值"
   'If PUB_GetCustData(Left(LTrim(RTrim(txtCNo)) & "000000000", 9), , clname) = True Then
   '   Label4(2).Caption = clname
   'End If
   txtCNo = ChangeCustomerL(txtCNo)
   clname = Replace(Pub_GetNameBYnation(txtCNo, "1"), "&", "＆")
   lblFM2.Caption = clname
   'end 2022/03/28
End If

Call GetNowGrp  'Added by Lydia 2022/03/28
End Sub

Private Sub txtCNo_Validate(Cancel As Boolean)
If Trim(txtCNo) <> "" Then
   strExc(0) = Left(LTrim(RTrim(txtCNo)), 1)
   If Not (strExc(0) = "X" Or strExc(0) = "Y" Or strExc(0) = "R") Then
      MsgBox "申請人/代理人/潛在客戶只能為 X、Y 或 R !!", , "USER 輸入錯誤"
      txtCNo_GotFocus
   End If
End If
End Sub
Private Sub ClearInputData(cInX As Integer)
Dim rr As Integer

Select Case cInX
       Case 1 '清空本所案號條件範圍
            For rr = 0 To 3
                txt1(rr) = ""
                If rr < 2 Then Option1(rr).Value = False
            Next rr
       
       Case 2 '清空申請人/代理人/潛在客戶範圍
            For rr = 0 To 3
                txtNo(rr) = ""
                If rr < 2 Then OptKind(rr).Value = False
            Next rr
            txtCNo.Text = ""
            Text2.Text = ""
            Text3.Text = ""
End Select

Call ClearAdd  'Added by Lydia 2022/03/28

End Sub

Private Sub txtNO_LostFocus(Index As Integer)
   Select Case Index
   Case 0

      If App.EXEName = "Patent" Then
        If txtNo(Index).Enabled = True Then
        Select Case txtNo(0)
        Case "FCP", "FG", "CFP", ""
        Case Else
             s = MsgBox("本所案號只能 FCP 或 FG 或 CFP !!!", , "USER 輸入錯誤")
             txtNo(0).SetFocus
             txtNo(0).SelStart = 0
             txtNo(0).SelLength = Len(txtNo(0))
             Exit Sub
        End Select
        End If
      ElseIf App.EXEName = "Trademark" Then
        If txtNo(Index).Enabled = True Then
        Select Case txtNo(0)
        Case "CFT", "FCT", ""
        Case Else
             s = MsgBox("本所案號只能 CFT 或 FCT !!", , "USER 輸入錯誤")
             txtNo(0).SetFocus
             txtNo(0).SelStart = 0
             txtNo(0).SelLength = Len(txtNo(0))
             Exit Sub
        End Select
        End If
      End If
   Case Else
   End Select
End Sub
Private Sub txtNo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNo_GotFocus(Index As Integer)
txtNo(Index).SelStart = 0
txtNo(Index).SelLength = Len(txtNo(Index))
CloseIme
End Sub

'Added by Lydia 2016/08/01
Private Sub Command1_Click()

Debug.Print "測試列印-開始:" & Format(ServerTime, "000000")
  PUB_RestorePrinter Combo1
  Screen.MousePointer = vbHourglass
   'Remove by Lydia 2017/07/21 不使用
   ' strExc(9) = "Printer.DeviceName= " & Printer.DeviceName & vbCrLf & _
                "Printer.Orientation= " & Printer.Orientation & vbCrLf & _
                "Printer.ScaleMode= " & Printer.ScaleMode & vbCrLf & _
                "Printer.Height=" & Printer.Height & vbCrLf & _
                "Printer.ScaleHeight=" & Printer.ScaleHeight & vbCrLf & _
                "Printer.Width=" & Printer.Width & vbCrLf & _
                "Printer.ScaleWidth=" & Printer.ScaleWidth
   'MsgBox strExc(9), vbOKOnly, "印表機原始設定"
   'end 2017/07/21
   
   Call PrintA4(True)
  '列印完就還原
  
  PUB_RestorePrinter pub_OsPrinter
  Screen.MousePointer = vbDefault
  
Debug.Print "測試列印-結束:" & Format(ServerTime, "000000")

End Sub

'Added by Lydia 2016/08/09 DHL套印A4
'Modified by Lydia 2016/08/31 改版面,保留程式
Private Sub PrintA4_old(Optional ByVal bShow As Boolean = False)
Dim douExtRate As Double '圖片縮放比例
Dim d_Top As Double, d_Left As Double '印表機的最小輸出邊界
Dim lngTop As Double '預設列印的上邊界
Dim lngLeft As Double '預設列印的左邊界
Dim iL01 As Double, iL02 As Double      '固定第一、二欄的起始位置
Dim intTop As Double, intLeft As Double '列印資料的上、左邊界
Dim mBarCode As String '提單號碼=Barcode
Dim byTwips As Integer '每公分的單位 'twips ,每公分=567
Dim iHt As Integer '行高
Dim nPages As Integer '一頁2張(給DHL和本所保留)
Dim m_PrtOrientation As Integer '列印方向
Dim m_PrtScaleMode As Integer '列印座標單位
Dim mRFno As String '本所案號或其他參考
Dim mContent As String 'DHL 提單資料內容
Dim tmpBol As Boolean 'Added by Lydia 2016/08/30
'單位: Twips
iHt = 180
byTwips = 567

'單位: 公分
lngLeft = 1: lngTop = 1 '預設列印的邊界(圖片貼齊)

bFTPok = False
    '處理:收件地址
    Erase recAddr
    strExc(1) = ""
    If bShow = True Then
        mBarCode = "1234567890"
        mRFno = "FCP-043783-0-00"
        recAddr(1) = "收件地址第一行 aaaaaaaaaaa"
        recAddr(2) = "收件地址第二行 bbbbb bb"
        recAddr(3) = "收件地址第三行 cccc, ccccccc,123"
        PayContact = "M51"
    Else
        mRFno = strTemp(8)
        '取得提單號碼
        If GetDHLRec(mBarCode) = False Then
           Exit Sub
        End If
        '處理收件地址長度(1~3欄，長度50字元)
        For intA = 3 To 7
           If strTemp(intA) <> "" Then
              strExc(1) = strExc(1) & IIf(strExc(1) <> "", ", ", "") & strTemp(intA)
           Else
              If intA < 7 And strTemp(intA + 1) = "" Then Exit For
           End If
        Next
        If LenB(StrConv(strExc(1), vbFromUnicode)) <= lenRA Then
           recAddr(1) = strExc(1)
        Else
           For intA = 1 To 3
              recAddr(intA) = PUB_StrToStr(strExc(1), lenRA)
              '調整斷行
              If LenB(StrConv(recAddr(intA), vbFromUnicode)) = lenRA _
                  And Right(recAddr(intA), 1) <> "" And Right(recAddr(intA), 1) <> "," And Right(recAddr(intA), 1) <> "-" Then
                 recAddr(intA) = Mid(recAddr(intA), 1, InStrRev(recAddr(intA), " "))
              End If
              strExc(1) = Mid(strExc(1), Len(recAddr(intA)) + 1)
              If Trim(strExc(1)) = "" Then Exit For
           Next
        End If
        '取得使用者的英文名稱
        PayContact = ""
        strSql = " select ED01,ED04 from ExtensionData where ed02='" & strUserNum & "' "
        intA = 1
        Set rsAD = ClsLawReadRstMsg(intA, strSql)
        If intA = 1 Then
           PayContact = "" & rsAD.Fields("ED04")
        End If
        
        If GetFtpTxt(mBarCode, mContent) Then
          '上傳到FTP
          If GetDHLRec(mBarCode, mContent) = True Then
             '上傳失敗,列印時加註記
             bFTPok = True
          End If
        End If
    End If
    
   '刪除舊的暫存圖檔
   strExc(1) = App.path & "\$DHLTmp.jpg"
   Set tmpImg = LoadPicture("")
   'Added by Lydia 2016/08/30 第二頁的圖檔
   strExc(2) = App.path & "\$DHLTmp2.jpg"
   Set tmpImg2 = LoadPicture("")
   
    '取得預設印表機設定值
    m_PrtOrientation = Printer.Orientation
    m_PrtScaleMode = Printer.ScaleMode
   
    '檢查是否安裝條碼字型
    Printer.Font = "Free 3 of 9 Extended"
    If Printer.Font <> "Free 3 of 9 Extended" Then
        '下載條碼字型, 因為Win7 必須在控制台的字型中新增字型,所以不下載
        'strExc(4) = Dir(App.path & "\FRE3OF9X.TTF")
        'If strExc(4) <> "" Then Kill App.path & "\FRE3OF9X.TTF"
        
        'strExc(2) = Dir("C:\WINDOWS\Fonts\FRE3OF9X.TTF")
        'If strExc(2) = "" Then
        '   strExc(4) = App.path & "\FRE3OF9X.TTF"
        '   strExc(3) = "//PolyCOM/Setup/NewProc/DHL_Barcode/FRE3OF9X.TTF"
        '   If PUB_FtpGetFile(strExc(3), strExc(4)) Then
        '      FileCopy strExc(4), "C:\WINDOWS\Fonts\FRE3OF9X.TTF"
        '      SetAttr strExc(4), vbNormal
        '      Kill strExc(4)
        '   End If
        'End If
       'If Pub_StrUserSt03 <> "M51" Then
          'MsgBox "安裝條碼字型失敗,請使用程式管理員帳戶登入!"
          MsgBox "尚未安裝條碼字型,請使用程式管理員帳戶登入!"
          Printer.Font = "Times New Roman"
          Printer.EndDoc
          Exit Sub
       'End If
    End If
    
    '設定紙張和方向
    Printer.PaperSize = 9 'A4
    Printer.Orientation = 1 '直印
    Printer.ScaleMode = 1
                
    '印表機的輸出邊界 by 公分
    d_Top = Format((Printer.Height - Printer.ScaleHeight) / byTwips / 2, "0.000")
    d_Left = Format((Printer.Width - Printer.ScaleWidth) / byTwips / 2, "0.000")
    
    If strPicFileName = "" Then
        If Dir(strExc(1)) <> "" Then Kill strExc(1)
        strPicFileName = App.path & "\$DHLTmp.jpg"
        If PUB_ReadDB2File(strPicFileName, 48) = True Then
        Else
           MsgBox "圖片載入失敗!"
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    End If
    'Added by Lydia 2016/08/30 載入第二頁
    If strPicFileName2 = "" Then
        If Dir(strExc(2)) <> "" Then Kill strExc(2)
        strPicFileName2 = App.path & "\$DHLTmp2.jpg"
        If PUB_ReadDB2File(strPicFileName2, 50) = True Then
        Else
           MsgBox "圖片載入失敗!"
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    End If
    
    Set tmpImg.Picture = LoadPicture(strPicFileName)
    Set tmpImg2.Picture = LoadPicture(strPicFileName2) 'Added by Lydia 2016/08/30
    
    '調整圖片, 預留邊界
     strExc(2) = "": strExc(3) = ""
    
    If Printer.ScaleHeight < tmpImg.Height Then strExc(2) = Format((Printer.ScaleHeight - byTwips * (lngTop - d_Top) * 2) / tmpImg.Height, "0.000")
    If Printer.ScaleWidth < tmpImg.Width Then strExc(3) = Format((Printer.ScaleWidth - byTwips * (lngLeft - d_Left) * 2) / tmpImg.Width, "0.000")
    If Val(strExc(2)) > 0 And Val(strExc(3)) > 0 Then
       If Val(strExc(2)) <= Val(strExc(3)) Then
          douExtRate = Val(strExc(2))
       Else
          douExtRate = Val(strExc(3))
       End If
    Else
       If Val(strExc(2)) > 0 Then
          douExtRate = Val(strExc(2))
       ElseIf Val(strExc(3)) > 0 Then
          douExtRate = Val(strExc(3))
       End If
    End If
    If douExtRate > 0 Then
       tmpImg.Height = tmpImg.Height * douExtRate
       tmpImg.Width = tmpImg.Width * douExtRate
       'Added by Lydia 2016/08/30
       tmpImg2.Height = tmpImg2.Height * douExtRate
       tmpImg2.Width = tmpImg2.Width * douExtRate
    End If
    '預設列印->水平置中
    wDHL = Format(tmpImg.Width, "0.000"): hDHL = Format(tmpImg.Height, "0.000")
    lngLeft = Format((Printer.Width - wDHL) / byTwips / 2, "0.00")
    If bShow And Pub_StrUserSt03 = "M51" Then
        strExc(9) = "Printer.DeviceName= " & Printer.DeviceName & vbCrLf & _
                    "Printer.Orientation= " & Printer.Orientation & vbCrLf & _
                    "Printer.ScaleMode= " & Printer.ScaleMode & vbCrLf & _
                    "Printer.Height=" & Printer.Height & vbCrLf & _
                    "Printer.ScaleHeight=" & Printer.ScaleHeight & vbCrLf & _
                    "Printer.Width=" & Printer.Width & vbCrLf & _
                    "Printer.ScaleWidth=" & Printer.ScaleWidth & vbCrLf & _
                    "圖片縮放比例=" & douExtRate & " " & vbCrLf & _
                    "圖片高度/圖片寬度= " & wDHL & " / " & hDHL & vbCrLf & _
                    "印表機的輸出邊界(上/左)= " & d_Top & " / " & d_Left
        MsgBox strExc(9), vbOKOnly, "印表機列印設定"
    End If
    
    Printer.PaintPicture tmpImg, (lngLeft - d_Left) * byTwips, (lngTop - d_Top) * byTwips, wDHL, hDHL
    '還原預設印表機設值
    Printer.ScaleMode = 1 '以Twips計算在紙張的位置(1cm = 567 Twips)
     '固定第一欄: 寄件Zip code 位置 +與預設列印的左邊界的間隔 (公分)
    iL01 = lngLeft + 1.1
     '固定第二欄: 寄件電話位置  ,直接列印與pdf輸出的位置有差
    iL02 = lngLeft + 4.4 - d_Left

  'A4 寬21cm; 高29.7cm
  'Modified by Lydia 2016/08/30 改成A4二張印3份
  'For nPages = 1 To 2
  For nPages = 1 To 3
    '計算固定列印的上邊界(ex. 0.4cm) = 自定上邊界(ex: 1 cm)-印表機上邊界(ex. 0.6 cm)
    'Modified by Lydia 2016/08/30 改成A4二張印3份
    'If nPages = 1 Then
    If nPages Mod 2 = 1 Then
       If nPages > 1 Then
          Printer.NewPage
          Printer.PaintPicture tmpImg2, (lngLeft - d_Left) * byTwips, (lngTop - d_Top) * byTwips, wDHL, hDHL
       End If
       'end 2016/08/30
       intTop = (lngTop - d_Top) * byTwips
    Else
       intTop = intTop + Format(Printer.Height / 2, "0.000")
    End If
    
    intLeft = (iL01 - d_Left) * byTwips '預設:寄件第一欄邊界
    Printer.Font = "Times New Roman"

    Printer.FontSize = 11
    Printer.CurrentX = intLeft + 7.5 * byTwips '8 * byTwips '+與第一欄的間隔
    Printer.CurrentY = intTop + (1.2 - lngTop) * byTwips
    Printer.Print toDblFont(Mid(mBarCode, 1, 3) & " " & Mid(mBarCode, 4, 4) & " " & Mid(mBarCode, 8)) & IIf(bFTPok, "", "　TX error")
    
    Printer.FontSize = 9
    Printer.CurrentX = intLeft + 0.7 * byTwips
    Printer.CurrentY = intTop + (1.9 - lngTop) * byTwips
    Printer.Print "V"

    Printer.CurrentX = intLeft + 1.2 * byTwips
    Printer.CurrentY = intTop + (2.2 - lngTop) * byTwips
    Printer.Print PayACNO
        
    Printer.Font = "Free 3 of 9 Extended" '條碼: Free 3 of 9 Extended / Free 3 of 9(原有)
    Printer.FontSize = 44
    Printer.CurrentX = intLeft + 6.7 * byTwips
    Printer.CurrentY = intTop + (2 - lngTop) * byTwips
    'DHL條碼機在掃瞄時，需要在前後各加一個"*"符號
    Printer.Print "*" & mBarCode & "*"
    
    Printer.Font = "Times New Roman"
    Printer.FontSize = 9
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (4.2 - lngTop) * byTwips
    Printer.Print PayACNO
    '4. International Document
    Printer.CurrentX = intLeft + 8.2 * byTwips
    Printer.CurrentY = intTop + (4.2 - lngTop) * byTwips
    Printer.Print "V"
    
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (4.9 - lngTop) * byTwips
    Printer.FontBold = True
    Printer.Print mRFno
    Printer.FontBold = False

    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (5.6 - lngTop) * byTwips
    Printer.Print PayACName
    '5.Shipment details
    Printer.FontSize = 11
    Printer.CurrentX = intLeft + 7.5 * byTwips
    Printer.CurrentY = intTop + (5.4 - lngTop) * byTwips
    Printer.Print "1" '預設：1件
    Printer.CurrentX = intLeft + 9.5 * byTwips
    Printer.CurrentY = intTop + (5.4 - lngTop) * byTwips
    Printer.Print "0.5" '預設：0.5 kg
    Printer.FontSize = 9
    
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (6.5 - lngTop) * byTwips
    Printer.Print PayACaddr01
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (6.5 - lngTop) * byTwips + iHt
    Printer.Print PayACaddr02

    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (7.5 - lngTop) * byTwips
    Printer.Print "TAIPEI"
    Printer.CurrentX = iL02 * byTwips
    Printer.CurrentY = intTop + (7.5 - lngTop) * byTwips
    Printer.Print "TAIWAN"
    '6. Content
    Printer.FontSize = 11
    Printer.CurrentX = intLeft + 6.8 * byTwips
    Printer.CurrentY = intTop + (7.8 - lngTop) * byTwips
    Printer.Print "Document"
    Printer.FontSize = 9
    
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (8.2 - lngTop) * byTwips
    Printer.Print PayACZip
    Printer.CurrentX = iL02 * byTwips
    Printer.CurrentY = intTop + (8.2 - lngTop) * byTwips
    Printer.Print PayACTel
    
    intLeft = (iL01 - d_Left - 0.5) * byTwips '收件第一欄邊界,凸行
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (9.4 - lngTop) * byTwips
    Printer.Print IIf(bShow, "收件公司ABCDEFR", strTemp(2))

    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (10.2 - lngTop) * byTwips
    Printer.Print recAddr(1)
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (10.2 - lngTop) * byTwips + iHt
    Printer.Print recAddr(2)
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (10.2 - lngTop) * byTwips + iHt * 2
    Printer.Print recAddr(3)
    
    If bShow = True Then
        Printer.CurrentX = intLeft
        Printer.CurrentY = intTop + (11.8 - lngTop) * byTwips
        Printer.Print "收件國家-不提供"
        Printer.CurrentX = iL02 * byTwips
        Printer.CurrentY = intTop + (11.8 - lngTop) * byTwips
        Printer.Print "收件ZIP-不提供"
        
        Printer.CurrentX = intLeft
        Printer.CurrentY = intTop + (12.5 - lngTop) * byTwips
        Printer.Print "收件城市-不提供"
    End If
    
    Printer.CurrentX = intLeft '- 0.2 * byTwips
    Printer.CurrentY = intTop + (13.2 - lngTop) * byTwips
    Printer.Print IIf(bShow, "收件聯絡人", strTemp(1))
    Printer.CurrentX = iL02 * byTwips
    Printer.CurrentY = intTop + (13.2 - lngTop) * byTwips
    Printer.Print IIf(bShow, "收件人電話", strTemp(9))
    Printer.FontSize = 11
    '自動代入列印人員的英文名字
    Printer.CurrentX = intLeft + 7.3 * byTwips
    Printer.CurrentY = intTop + (13.2 - lngTop) * byTwips
    Printer.Print PayContact  '寄件聯絡;原本為人工簽名
    '代入列印日期
    Printer.CurrentX = intLeft + 12 * byTwips
    Printer.CurrentY = intTop + (13.2 - lngTop) * byTwips
    Printer.Print ChangeWStringToWDateString(strSrvDate(1))
    
  Next nPages
  
  Printer.EndDoc
  '還原
  Printer.Orientation = m_PrtOrientation
  Printer.ScaleMode = m_PrtScaleMode
  
  '刪除已完成上傳和列印的提單資料
  If bShow = False Then
     Kill strFileName
  End If
  ShowPrintOk
  
  Exit Sub
  
ErrHand01:
  If Err.Number <> 0 Then MsgBox Err.Description
End Sub

'Added by Lydia 2016/08/09 產生上傳FTP的txt檔
Private Function GetFtpTxt(ByVal SPno As String, ByRef SPcont As String) As Boolean
Dim fN As Integer

    strFileName = App.path & "\" & strSrvDate(1) & "_" & SPno & ".txt"
    If Dir(strFileName) <> "" Then Kill strFileName
    
    SPcont = defCont
    '提單號碼
    SPcont = Replace(SPcont, "N123456789", SPno)
    '寄件聯絡人
    SPcont = Replace(SPcont, "PayACMan", IIf(Trim(PayContact) = "", ".", PayContact))
    '本所案號
    SPcont = Replace(SPcont, "Shipper Reference", PUB_StrToStr(Trim(strTemp(8)), 35))
    '收件公司
    SPcont = Replace(SPcont, "收件公司", PUB_StrToStr(Trim(strTemp(2)), lenRA))
    'Added by Lydia 2016/08/31 +提單日期,時間
    SPcont = Replace(SPcont, "SYSD", strSrvDate(1))
    SPcont = Replace(SPcont, "SYST", Mid(pTime, 1, 4))
    '收件地址
    SPcont = Replace(SPcont, "收件地址1", IIf(Trim(recAddr(1)) = "", ".", recAddr(1)))
    SPcont = Replace(SPcont, "收件地址2", IIf(Trim(recAddr(2)) = "", ".", recAddr(2)))
    SPcont = Replace(SPcont, "收件地址3", recAddr(3))
    'Modified by Lydia 2022/03/28
    'SPcont = Replace(SPcont, "收件城市填空白", " ")
    'SPcont = Replace(SPcont, "收件郵遞區號填.", ".")
    'SPcont = Replace(SPcont, "收件國", " ")
    'SPcont = Replace(SPcont, "國代C(2)", " ")
    '--------------
    strExc(1) = PUB_GetSimpleName(recDID04, True, True)
    If Trim(strExc(1)) = "" Then 'DHL要求必填; 城市空白=>用國名
       SPcont = Replace(SPcont, "收件城市", PUB_StrToStr(strTemp(11), 35))
    Else
       SPcont = Replace(SPcont, "收件城市", PUB_StrToStr(strExc(1), 35))
    End If
    strExc(1) = PUB_GetSimpleName(strTemp(10), True, True)
    If Trim(strExc(1)) = "" Then
        SPcont = Replace(SPcont, "收件郵遞區號填.", ".")
    Else
        SPcont = Replace(SPcont, "收件郵遞區號填.", PUB_StrToStr(strExc(1), 12))
    End If
    SPcont = Replace(SPcont, "收件國", PUB_StrToStr(strTemp(11), 30))
    SPcont = Replace(SPcont, "國代C(2)", IIf(recNA89 = "", " ", recNA89)) 'DHL要求必填
    'end 2022/03/28
    '收件聯絡人
    'Modified by Lydia 2022/03/28
    'SPcont = Replace(SPcont, "收件聯絡人", IIf(Trim(strTemp(1)) = "", " ", PUB_StrToStr(Trim(strTemp(1)), lenRA)))
    strExc(1) = PUB_GetSimpleName(strTemp(1), True, True)
    If Trim(strExc(1)) = "" Then 'DHL要求必填; 可以重複公司名稱
        SPcont = Replace(SPcont, "收件聯絡人", PUB_StrToStr(Trim(strTemp(2)), lenRA))
    Else
        SPcont = Replace(SPcont, "收件聯絡人", PUB_StrToStr(Trim(strTemp(1)), lenRA))
    End If
    'end 2022/03/28
    '收件人電話
    strExc(1) = strTemp(9)
    If Len(strExc(1)) > 18 Then
       strExc(1) = Replace(strExc(1), " ", "")
       strExc(1) = Replace(strExc(1), "-", "")
    End If
    SPcont = Replace(SPcont, "收件人電話C(18)", PUB_StrToStr(strExc(1), 18))
    
    'Added by Lydia 2017/11/01 因為DHL反應非英數字會造成轉檔失敗,所以將txt內容的非英數字取代掉
    'Modified by Lydia 2022/03/28 改使用共用模組
    'strExc(1) = GetSimpleName(SPcont)
    strExc(1) = PUB_GetSimpleName(SPcont, True, True)
    SPcont = strExc(1)
    'end 2017/11/01
    
    fN = FreeFile
    Open strFileName For Output As fN
    Print #fN, "1|" & Space(6)
    Print #fN, SPcont
    
    If fN > 0 Then
       Close fN
       GetFtpTxt = True
    End If
   
End Function

'Added by Lydia 2016/08/09 取得提單號碼/上傳資料到FTP
Private Function GetDHLRec(ByRef RecNo As String, Optional ByRef RecCont) As Boolean
Dim intL As Integer
Dim tmpUs As String
Dim MaxRec As String '最大提單號碼
Dim mDate As String '建檔日=>順序
Dim tmpBol As Boolean
Dim bolWrite As Boolean
    
On Error GoTo ErrHand
    GetDHLRec = False
    '取得FTP資料夾路徑
    If strFileDir = "" Then
       strFileDir = Pub_GetSpecMan("FTP_DHL_Path")
       If strFileDir = "" Then
          MsgBox "無法取得DHL FTP資料夾路徑!'"
          Exit Function
       End If
    End If
    
    '列印前,先取得提單號碼
    If RecNo = "" Then
       tmpUs = "SELECT DSN01,DSN02,DSN05 FROM DHLSHIPMENTNO WHERE DSN03 IS NULL ORDER BY DSN05"
       intL = 1
       Set rsAD = ClsLawReadRstMsg(intL, tmpUs)
       If intL = 0 Then
          MsgBox "無可使用的DHL提單號碼!"
          Exit Function
       Else
          'DHL提單號碼=提單流水號(9碼)+檢查碼(1碼=流水號 mod 7)
          RecNo = "" & rsAD(0)
          MaxRec = "" & rsAD(1)
          mDate = "" & rsAD(2)
          tmpUs = "select nvl(max(dun01),'1') mno from dhluseno where dun01>='" & rsAD(0) & "' and dun01<='" & rsAD(1) & "' "
          intL = 1
          Set rsAD = ClsLawReadRstMsg(intL, tmpUs)
          If Trim(rsAD.Fields("MNO")) <> "1" Then
             RecNo = Val(Mid(rsAD.Fields("MNO"), 1, Len(rsAD.Fields("MNO")) - 1)) + 1
             If Mid(MaxRec, 1, Len(MaxRec) - 1) < RecNo Then
                MsgBox "DHL提單號碼使用完畢,請改回點陣印表機套印!"
                Exit Function
             Else
                RecNo = RecNo & Val(RecNo) Mod 7
             End If
          End If
          
          cnnConnection.BeginTrans
             bolWrite = True
             '用完提單號碼,整批記錄上不使用
             If Mid(MaxRec, 1, Len(MaxRec) - 1) = Mid(RecNo, 1, Len(RecNo) - 1) Then
                tmpUs = "Update DHLSHIPMENTNO set dsn03='N' where dsn05='" & mDate & "' and dsn02 like '" & Mid(MaxRec, 1, Len(MaxRec) - 1) & "%' "
                cnnConnection.Execute tmpUs, intL
             End If
             pTime = Format(ServerTime, "000000") 'Added by Lydia 2016/08/31
             tmpUs = "INSERT INTO DHLUSENO (DUN01,DUN02,DUN03,DUN04,DUN05,DUN06) VALUES ('" & RecNo & "',NULL,NULL,'" & strUserNum & "'," & CNULL(strSrvDate(1), True) & ",'" & pTime & "') "
             cnnConnection.Execute tmpUs, intL
          cnnConnection.CommitTrans
       End If
       Set rsAD = Nothing
    '上傳FTP
    ElseIf RecNo <> "" Then
          tmpUs = ""
          If RecCont <> "" Then tmpUs = ",dun02='" & ChgSQL(RecCont) & "' "
          If Pub_StrUserSt03 = "M51" Then
             If MsgBox("是否上傳資料到DHL FTP?", vbYesNo + vbDefaultButton2) = vbYes Then
               tmpUs = tmpUs & ",dun06=" & Format(ServerTime, "000000") '因為有詢問,所以create time要更新
               tmpBol = PUB_FtpPutFileDHL("FTP_DHL_IP", strFileName, strFileDir & "/" & Mid(strFileName, InStrRev(strFileName, "\") + 1))
               If tmpBol = True Then tmpUs = tmpUs & ",dun03=" & Format(ServerTime, "000000")  '上傳成功,更新時間
             Else
               tmpBol = True
             End If
          Else
             tmpBol = PUB_FtpPutFileDHL("FTP_DHL_IP", strFileName, strFileDir & "/" & Mid(strFileName, InStrRev(strFileName, "\") + 1))
             If tmpBol = True Then tmpUs = tmpUs & ",dun03=" & Format(ServerTime, "000000")  '上傳成功,更新時間
          End If
          
          If tmpUs <> "" Then
             cnnConnection.BeginTrans
                bolWrite = True
                tmpUs = "update dhluseno set " & Mid(tmpUs, 2) & " where dun01='" & RecNo & "' and dun05=" & CNULL(strSrvDate(1), True)
                cnnConnection.Execute tmpUs, intL
             cnnConnection.CommitTrans
          End If
          If tmpBol = False Then Exit Function
    Else
          Exit Function
    End If
    
    GetDHLRec = True
    Exit Function
    
ErrHand:

If Err.Number <> 0 Then
   If bolWrite Then cnnConnection.RollbackTrans
   MsgBox Err.Description
End If

End Function

'Modified by Lydia 2016/08/31 改版面,保留程式
Private Sub PrintA4(Optional ByVal bShow As Boolean = False)
Dim douExtRate As Double '圖片縮放比例
Dim d_Top As Double, d_Left As Double '印表機的最小輸出邊界
Dim lngTop As Double '預設列印的上邊界
Dim lngLeft As Double '預設列印的左邊界
Dim iL01 As Double, iL02 As Double      '固定第一、二欄的起始位置
Dim intTop As Double, intLeft As Double '列印資料的上、左邊界
Dim mBarCode As String '提單號碼=Barcode
Dim byTwips As Integer '每公分的單位 'twips ,每公分=567
Dim iHt As Integer '行高
Dim nPages As Integer '一頁2張(給DHL和本所保留)
Dim m_PrtOrientation As Integer '列印方向
Dim m_PrtScaleMode As Integer '列印座標單位
Dim mRFno As String '本所案號或其他參考
Dim mContent As String 'DHL 提單資料內容
Dim tmpBol As Boolean 'Added by Lydia 2016/08/30
'單位: Twips
iHt = 180
byTwips = 567

'單位: 公分
lngLeft = 1: lngTop = 1 '預設列印的邊界(圖片貼齊)

'Mark by Lydia 2016/09/01 保留上傳圖檔
'tmpBol = SaveImgByteFile("Z:\TaieNew\RptSample\M51-000048-0-00 DHL fullsize.JPG", "M51", "000048", "0", "00", "1", "2") 'M51-000048-0-00 DHL fullsize.JPG
'MsgBox "存檔=" & tmpBol
'tmpBol = SaveImgByteFile("Z:\TaieNew\RptSample\M51-000050-0-00 DHL half.JPG", "M51", "000050", "0", "00", "1", "2") 'M51-000048-0-00 DHL fullsize.JPG
'MsgBox "存檔=" & tmpBol
'end 2016/09/01

bFTPok = False
    '處理:收件地址
    Erase recAddr
    strExc(1) = ""
    If bShow = True Then
        mBarCode = "1234567890"
        mRFno = "FCP-043783-0-00"
        recAddr(1) = "收件地址第一行 aaaaaaaaaaa"
        recAddr(2) = "收件地址第二行 bbbbb bb"
        'Modified by Lydia 2022/03/28
        'recAddr(3) = "收件地址第三行 cccc, ccccccc,123"
        recAddr(3) = "上傳txt匯整為三欄,列印可分為5行"
        PayContact = "M51"
    Else
        mRFno = strTemp(8)
        '取得提單號碼
        If GetDHLRec(mBarCode) = False Then
           Exit Sub
        End If
        '處理收件地址長度(1~3欄，長度50字元)
        For intA = 3 To 7
           If strTemp(intA) <> "" Then
              strExc(1) = strExc(1) & IIf(strExc(1) <> "", ", ", "") & strTemp(intA)
           Else
              If intA < 7 And strTemp(intA + 1) = "" Then Exit For
           End If
        Next
        strExc(1) = Replace(strExc(1), ",,", ",")  'Added by Lydia 2022/04/22 原本地址就用,區隔；ex.Y20065
        'Added by Lydia 2022/04/06 先排除非英數字
        strExc(1) = Trim(PUB_GetSimpleName(strExc(1), True, True))
        If Right(strExc(1), 1) = "," Then
            strExc(1) = Mid(strExc(1), 1, Len(strExc(1)) - 1)
        End If
        'end 2022/04/06
        If LenB(StrConv(strExc(1), vbFromUnicode)) <= lenRA Then
           recAddr(1) = strExc(1)
        Else
           For intA = 1 To 3
              recAddr(intA) = PUB_StrToStr(strExc(1), lenRA)
              '調整斷行
              If LenB(StrConv(recAddr(intA), vbFromUnicode)) = lenRA _
                  And Right(recAddr(intA), 1) <> "" And Right(recAddr(intA), 1) <> "," And Right(recAddr(intA), 1) <> "-" Then
                 recAddr(intA) = Mid(recAddr(intA), 1, InStrRev(recAddr(intA), " "))
              End If
              strExc(1) = Mid(strExc(1), Len(recAddr(intA)) + 1)
              If Trim(strExc(1)) = "" Then Exit For
           Next
        End If
        
        '取得使用者的英文名稱
        PayContact = ""
        'Modified by Lydia 2024/06/26 因為虛編號沒有分機記錄，所以改用ST12
        'strSql = " select ED01,ED04 from ExtensionData where ed02='" & strUserNum & "' "
        strSql = " select st12 as ed04 from staff where st01='" & strUserNum & "' "
        intA = 1
        Set rsAD = ClsLawReadRstMsg(intA, strSql)
        If intA = 1 Then
           PayContact = "" & rsAD.Fields("ED04")
        End If
        'Added by Lydia 2024/06/26
        If PayContact = "" Then   '改抓分機表的英文別名
           strExc(1) = Pub_GetStaffExtn(strUserNum, PayContact)
        End If
        'end 2024/06/26
        
        If GetFtpTxt(mBarCode, mContent) Then
          '上傳到FTP
          If GetDHLRec(mBarCode, mContent) = True Then
             '上傳失敗,列印時加註記
             bFTPok = True
          End If
        End If
    End If
    
   '刪除舊的暫存圖檔
   strExc(1) = App.path & "\$DHLTmp.jpg"
   Set tmpImg = LoadPicture("")
   '第二頁的圖檔
   strExc(2) = App.path & "\$DHLTmp2.jpg"
   Set tmpImg2 = LoadPicture("")
   
   Call SaveBCbmp("*" & mBarCode & "*") 'DHL條碼機在掃瞄時，需要在前後各加一個"*"符號
   
    '取得預設印表機設定值
    m_PrtOrientation = Printer.Orientation
    m_PrtScaleMode = Printer.ScaleMode
   
    '檢查是否安裝條碼字型
    Printer.Font = "Free 3 of 9 Extended"
    If Printer.Font <> "Free 3 of 9 Extended" Then
        '因為Win7 必須在控制台的字型中新增字型,所以不下載
          MsgBox "尚未安裝條碼字型,請使用程式管理員帳戶登入!"
          Printer.Font = "Times New Roman"
          Printer.EndDoc
          Exit Sub
    End If
    
    'Added by Lydia 2022/04/06 改用Word套印
    Call PrintWordA4(mBarCode, mRFno, bShow)
    'Added by Lydia 2024/08/30 還原預設印表機
    PUB_SetOsDefaultPrinter strPrinter
    PUB_SetWordActivePrinter
    'end 2024/08/30
    
    Exit Sub
    'end 2022/04/06
    
    '設定紙張和方向
    Printer.PaperSize = 9 'A4
    Printer.Orientation = 1 '直印
    Printer.ScaleMode = 1
                
    '印表機的輸出邊界 by 公分
    d_Top = Format((Printer.Height - Printer.ScaleHeight) / byTwips / 2, "0.000")
    d_Left = Format((Printer.Width - Printer.ScaleWidth) / byTwips / 2, "0.000")
    
    '載入第一頁
    If strPicFileName = "" Then
        If Dir(strExc(1)) <> "" Then Kill strExc(1)
        strPicFileName = App.path & "\$DHLTmp.jpg"
        If PUB_ReadDB2File(strPicFileName, 48) = True Then
        Else
           MsgBox "圖片載入失敗!"
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    End If
    '載入第二頁
    If strPicFileName2 = "" Then
        If Dir(strExc(2)) <> "" Then Kill strExc(2)
        strPicFileName2 = App.path & "\$DHLTmp2.jpg"
        If PUB_ReadDB2File(strPicFileName2, 50) = True Then
        Else
           MsgBox "圖片載入失敗!"
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    End If
    
    Set tmpImg.Picture = LoadPicture(strPicFileName)
    Set tmpImg2.Picture = LoadPicture(strPicFileName2)
    
    '調整圖片, 預留邊界
     strExc(2) = "": strExc(3) = ""
    
    If Printer.ScaleHeight < tmpImg.Height Then strExc(2) = Format((Printer.ScaleHeight - byTwips * (lngTop - d_Top) * 2) / tmpImg.Height, "0.000")
    If Printer.ScaleWidth < tmpImg.Width Then strExc(3) = Format((Printer.ScaleWidth - byTwips * (lngLeft - d_Left) * 2) / tmpImg.Width, "0.000")
    If Val(strExc(2)) > 0 And Val(strExc(3)) > 0 Then
       If Val(strExc(2)) <= Val(strExc(3)) Then
          douExtRate = Val(strExc(2))
       Else
          douExtRate = Val(strExc(3))
       End If
    Else
       If Val(strExc(2)) > 0 Then
          douExtRate = Val(strExc(2))
       ElseIf Val(strExc(3)) > 0 Then
          douExtRate = Val(strExc(3))
       End If
    End If
    If douExtRate > 0 Then
       tmpImg.Height = tmpImg.Height * douExtRate
       tmpImg.Width = tmpImg.Width * douExtRate
       tmpImg2.Height = tmpImg2.Height * douExtRate
       tmpImg2.Width = tmpImg2.Width * douExtRate
    End If
    '預設列印->水平置中
    wDHL = Format(tmpImg.Width, "0.000"): hDHL = Format(tmpImg.Height, "0.000")
    lngLeft = Format((Printer.Width - wDHL) / byTwips / 2, "0.00")
    'Remove by Lydia 2017/07/21 不使用
'    If bShow And Pub_StrUserSt03 = "M51" Then
'        strExc(9) = "Printer.DeviceName= " & Printer.DeviceName & vbCrLf & _
'                    "Printer.Orientation= " & Printer.Orientation & vbCrLf & _
'                    "Printer.ScaleMode= " & Printer.ScaleMode & vbCrLf & _
'                    "Printer.Height=" & Printer.Height & vbCrLf & _
'                    "Printer.ScaleHeight=" & Printer.ScaleHeight & vbCrLf & _
'                    "Printer.Width=" & Printer.Width & vbCrLf & _
'                    "Printer.ScaleWidth=" & Printer.ScaleWidth & vbCrLf & _
'                    "圖片縮放比例=" & douExtRate & " " & vbCrLf & _
'                    "圖片高度/圖片寬度= " & wDHL & " / " & hDHL & vbCrLf & _
'                    "印表機的輸出邊界(上/左)= " & d_Top & " / " & d_Left
'        MsgBox strExc(9), vbOKOnly, "印表機列印設定"
'    End If
    'end 2017/07/21
    
    Printer.PaintPicture tmpImg, (lngLeft - d_Left) * byTwips, (lngTop - d_Top) * byTwips, wDHL, hDHL
    '還原預設印表機設值
    Printer.ScaleMode = 1 '以Twips計算在紙張的位置(1cm = 567 Twips)
     '固定第一欄: 寄件Zip code 位置 +與預設列印的左邊界的間隔 (公分)
    iL01 = lngLeft + 1.1
     '固定第二欄: 寄件電話位置  ,直接列印與pdf輸出的位置有差
    iL02 = lngLeft + 4.4 - d_Left

  'A4 寬21cm; 高29.7cm
  For nPages = 1 To 3
    '計算固定列印的上邊界(ex. 0.4cm) = 自定上邊界(ex: 1 cm)-印表機上邊界(ex. 0.6 cm)
    If nPages Mod 2 = 1 Then
       If nPages > 1 Then
          Printer.NewPage
          Printer.PaintPicture tmpImg2, (lngLeft - d_Left) * byTwips, (lngTop - d_Top) * byTwips, wDHL, hDHL
       End If
       intTop = (lngTop - d_Top) * byTwips
    Else
       intTop = intTop + Format(Printer.Height / 2, "0.000")
    End If
    
    intLeft = (iL01 - d_Left) * byTwips '預設:寄件第一欄邊界
    Printer.Font = "Times New Roman"
    
    'ORIGIN
    Printer.FontSize = 14
    Printer.CurrentX = intLeft + 12.85 * byTwips
    Printer.CurrentY = intTop + (1.2 - lngTop) * byTwips
    Printer.Print "TPE"
    
    Printer.Font = "Arial"
    Printer.FontSize = 19
    Printer.CurrentX = intLeft + 7.5 * byTwips '8 * byTwips '+與第一欄的間隔
    Printer.CurrentY = intTop + (1.4 - lngTop) * byTwips
    Printer.Print Mid(mBarCode, 1, 3) & " " & Mid(mBarCode, 4, 4) & " " & Mid(mBarCode, 8)
    If bFTPok = False Then
       Printer.FontSize = 10
       Printer.CurrentX = intLeft + 12.2 * byTwips
       Printer.CurrentY = intTop + (1.8 - lngTop) * byTwips
       Printer.Print "TX error"
    End If
    
    Printer.Font = "Times New Roman"
     
    'Products & service = International Document
    Printer.FontSize = 9
    Printer.CurrentX = intLeft + 14.85 * byTwips
    Printer.CurrentY = intTop + (2.9 - lngTop) * byTwips
    Printer.Print "V"
    
    'Added by Lydia 2017/07/21 DHL要求勾選Express/ Worldwide，預設為一般件
    Printer.FontSize = 8
    Printer.CurrentX = intLeft + 14.2 * byTwips
    Printer.CurrentY = intTop + (4.1 - lngTop) * byTwips
    Printer.Print "V"
    Printer.FontSize = 9
    'end 2017/07/21
    
    'Charge to Shipper
    Printer.CurrentX = intLeft + 0.75 * byTwips
    Printer.CurrentY = intTop + (2.1 - lngTop) * byTwips
    Printer.Print "V"
    'Payer
    Printer.CurrentX = intLeft + 1.2 * byTwips
    Printer.CurrentY = intTop + (2.4 - lngTop) * byTwips
    Printer.Print PayACNO
        
    '用圖片寫條碼 (40號字,高度*1.6)
    Printer.PaintPicture tmpBC, intLeft + 7 * byTwips, intTop + (2.3 - lngTop) * byTwips, tmpBC.Width, tmpBC.Height
     
    Printer.Font = "Times New Roman"
    Printer.FontSize = 9
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (4.3 - lngTop) * byTwips
    Printer.Print PayACNO
    
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (5 - lngTop) * byTwips
    Printer.FontBold = True
    Printer.Print mRFno
    Printer.FontBold = False

    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (5.8 - lngTop) * byTwips
    Printer.Print PayACName
    '5.Shipment details
    Printer.FontSize = 11
    Printer.CurrentX = intLeft + 7.3 * byTwips
    Printer.CurrentY = intTop + (5.4 - lngTop) * byTwips
    Printer.Print "1" '預設：1件
    Printer.CurrentX = intLeft + 8.7 * byTwips
    Printer.CurrentY = intTop + (5.4 - lngTop) * byTwips
    Printer.Print "0.5" '預設：0.5 kg
    Printer.FontSize = 9
    
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (6.6 - lngTop) * byTwips
    Printer.Print PayACaddr01
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (6.6 - lngTop) * byTwips + iHt
    Printer.Print PayACaddr02

    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (8.1 - lngTop) * byTwips
    Printer.Print PayACZip
    Printer.CurrentX = iL02 * byTwips
    Printer.CurrentY = intTop + (8.1 - lngTop) * byTwips
    Printer.Print PayACTel
    
    '6. Content
    Printer.FontSize = 11
    Printer.CurrentX = intLeft + 6.8 * byTwips
    Printer.CurrentY = intTop + (7.8 - lngTop) * byTwips
    Printer.Print "Document"
    
    Printer.FontSize = 9
    
    intLeft = (iL01 - d_Left - 0.5) * byTwips '收件第一欄邊界,凸行
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (9.2 - lngTop) * byTwips
    Printer.Print IIf(bShow, "收件公司ABCDEFR", strTemp(2))

   
    If bShow = True Then
        Printer.CurrentX = intLeft
        Printer.CurrentY = intTop + (10 - lngTop) * byTwips
        Printer.Print recAddr(1)
        Printer.CurrentX = intLeft
        Printer.CurrentY = intTop + (10 - lngTop) * byTwips + iHt
        Printer.Print recAddr(2)
        Printer.CurrentX = intLeft
        Printer.CurrentY = intTop + (10 - lngTop) * byTwips + iHt * 2
        Printer.Print recAddr(3)
        Printer.CurrentX = intLeft
        Printer.CurrentY = intTop + (12.6 - lngTop) * byTwips
        'Modified by Lydia 2022/03/28
        'Printer.Print "收件國家-不提供"
        Printer.Print "郵遞區號-DID03"
        Printer.CurrentX = iL02 * byTwips
        Printer.CurrentY = intTop + (12.6 - lngTop) * byTwips
        'Modified by Lydia 2022/03/28
        'Printer.Print "收件ZIP-不提供"
        Printer.Print "收件國家-DCC02"
    'Added by Lydia 2016/09/01
    Else '地址列印為5行，資料上傳處理合併為3行
        For intA = 3 To 7
            If strTemp(intA) <> "" Then
                Printer.CurrentX = intLeft
                Printer.CurrentY = intTop + (10 - lngTop) * byTwips + iHt * (intA - 3)
                Printer.Print strTemp(intA)
            End If
        Next
        'Added by Lydia 2022/03/28
        If strTemp(10) <> "" And strTemp(10) <> "." Then
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + (12.6 - lngTop) * byTwips
            Printer.Print strTemp(10)  '郵遞區號-DID03
        End If
        Printer.CurrentX = iL02 * byTwips
        Printer.CurrentY = intTop + (12.6 - lngTop) * byTwips
        Printer.Print strTemp(11) '收件國家-DCC02
        'end 2022/03/28
    End If
    'end 2016/09/01
    
    Printer.CurrentX = intLeft
    Printer.CurrentY = intTop + (13.4 - lngTop) * byTwips
    'Modified by Lydia 2022/03/28 DHL要求必填; 可以重複公司名稱
    'Printer.Print IIf(bShow, "收件聯絡人", strTemp(1))
    If bShow = True Then
       Printer.Print "收件聯絡人-DID05"
    Else
       strExc(1) = IIf(strTemp(1) = "", strTemp(2), strTemp(1))
       If Len(strExc(1)) < 24 Then
           Printer.Print strExc(1)
       Else
           Printer.Print Mid(strExc(1), 1, 24)
           '超過欄位,折行
           Printer.CurrentX = intLeft
           Printer.CurrentY = intTop + (13.8 - lngTop) * byTwips
           Printer.Print Mid(strExc(1), 25)
       End If
       
    End If
    'END 2022/03/28
    Printer.CurrentX = iL02 * byTwips
    Printer.CurrentY = intTop + (13.4 - lngTop) * byTwips
    'Modified by Lydia 2022/03/28
    'Printer.Print IIf(bShow, "收件人電話", strTemp(9))
    Printer.Print IIf(bShow, "收件人電話-DID06", strTemp(9))
    Printer.FontSize = 11
    '自動代入列印人員的英文名字
    Printer.CurrentX = intLeft + 8.2 * byTwips
    Printer.CurrentY = intTop + (13.4 - lngTop) * byTwips
    Printer.Print PayContact  '寄件聯絡;原本為人工簽名
    '代入列印日期
    Printer.FontSize = 9
    Printer.CurrentX = intLeft + 12.25 * byTwips
    Printer.CurrentY = intTop + (13.4 - lngTop) * byTwips
    Printer.Print Mid(strSrvDate(1), 1, 4)
    
    Printer.CurrentX = intLeft + 13.2 * byTwips
    Printer.CurrentY = intTop + (13.4 - lngTop) * byTwips
    Printer.Print Mid(strSrvDate(1), 5, 2)
    
    Printer.CurrentX = intLeft + 13.8 * byTwips
    Printer.CurrentY = intTop + (13.4 - lngTop) * byTwips
    Printer.Print Mid(strSrvDate(1), 7, 8)
        
        
  Next nPages
  
  Printer.EndDoc
  '還原
  Printer.Orientation = m_PrtOrientation
  Printer.ScaleMode = m_PrtScaleMode
  
  '刪除已完成上傳和列印的提單資料
  If bShow = False Then
     Kill strFileName
  End If
  ShowPrintOk
  
  Exit Sub
  
ErrHand01:
  If Err.Number <> 0 Then MsgBox Err.Description
End Sub

'Added by Lydia 2016/08/31 將Barcode轉圖片
Private Sub SaveBCbmp(ByVal pCode As String)
    
    Set tmpBC = LoadPicture("")
    
    'Modified by Lydia 2022/04/06
    'saveBCfile = App.path & "\$DHL_BC.BMP"
    saveBCfile = App.path & "\" & strUserNum & "\$DHL_BC.BMP"
    If Dir(saveBCfile) <> "" Then Kill saveBCfile
    
    tmpBCode.Cls
    tmpBCode.Print pCode
    
    SavePicture tmpBCode.Image, saveBCfile

    Set tmpBC = LoadPicture(saveBCfile)
    tmpBC.Height = tmpBC.Height * 1.8 '拉高
    
End Sub

'Added by Lydia 2017/11/01 DHL資料限英數字
'Modified by Lydia 2022/03/28 改使用共用模組
'Private Function GetSimpleName(oldName As String) As String
'   Dim stSimpleName As String
'   Dim iPos As Integer, ii As Integer, ICode As Integer
'
'   iPos = InStrRev(oldName, "/")
'   If iPos > 0 Then
'      stSimpleName = Left(oldName, iPos)
'   End If
'   For ii = iPos + 1 To Len(oldName)
'      ICode = Asc(Mid(oldName, ii, 1))
'      If ICode > 0 And ICode <= 255 Then
'        stSimpleName = stSimpleName & Mid(oldName, ii, 1)
'      Else
'        stSimpleName = stSimpleName & " "
'      End If
'   Next
'   GetSimpleName = stSimpleName
'
'End Function
'end 2022/03/28

'Added by Lydia 2018/01/16
Private Sub txtField_GotFocus(Index As Integer)
  If Index > 0 Then TextInverse txtField(Index)
End Sub

'Added by Lydia 2018/01/16 處理人工填提單的資料
Private Function StringFilterr(ByVal p_Str As String) As String
   p_Str = Replace(p_Str, Chr(11), " ")
   p_Str = Replace(p_Str, Chr(10), " ")
   p_Str = Replace(p_Str, Chr(13), " ")
   p_Str = PUB_RepToOneSpace(p_Str)
   StringFilterr = p_Str
End Function

'Added by Lydia 2018/01/16 人工填提單
Private Sub ProcessNew2()
   
    pub_QL05 = pub_QL05 & ";人工填提單;"
         
    '收件聯絡人
    strTemp(1) = StringFilterr(txtField(2).Text)
    If Trim(strTemp(1)) <> "" Then pub_QL05 = pub_QL05 & "收件聯絡人:" & strTemp(1) & ";"
    '收件公司
    strTemp(2) = StringFilterr(txtField(0).Text)
    If Trim(strTemp(2)) <> "" Then pub_QL05 = pub_QL05 & "收件公司:" & strTemp(2) & ";"
    '收件地址
    strTemp(3) = StringFilterr(txtField(1).Text)
    If Trim(strTemp(3)) <> "" Then pub_QL05 = pub_QL05 & "收件地址:" & strTemp(3) & ";"
    '收件地址分5行
    strTemp1 = Empty
    strTemp1 = Split(txtField(1), vbNewLine)
    j = 3
    For i = 0 To UBound(strTemp1)
         If Trim(strTemp1(i)) <> "" Then
             strTemp(j) = strTemp1(i)
             j = j + 1
         End If
    Next i
    If j <= 7 Then
        For i = j To 7
            strTemp(i) = ""
        Next i
    End If
    
    '收件人電話
   strTemp(9) = StringFilterr(txtField(3).Text)
   'Modified by Lydia 2022/03/28 收件人電話/Email=>收件人電話
   If Trim(strTemp(9)) <> "" Then pub_QL05 = pub_QL05 & "收件人電話:" & strTemp(9) & ";"
    '說明
   strTemp(8) = StringFilterr(txtField(4).Text)
   If Trim(strTemp(8)) <> "" Then pub_QL05 = pub_QL05 & "說明:" & strTemp(8) & ";"
   
   If Trim(strTemp(3) & strTemp(9)) = "" Then
        'Modified by Lydia 2022/03/28 收件人電話/Email=>收件人電話
        MsgBox "請輸入收件地址和電話 !!"
        Exit Sub
   End If
   
   Call InsertQueryLog("1")
   
   PrintA4
   
End Sub

Private Sub txtField_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index <> 1 Then txtField(Index).ToolTipText = Trim(StringFilterr(txtField(Index).Text))
End Sub

'Added by Lydia 2022/03/28
Private Sub txtAdd_GotFocus(Index As Integer)
  If Index > 0 Then TextInverse txtAdd(Index)
End Sub

Private Sub ClearAdd()
       
   For Each oText In txtAdd
       oText.Text = ""
       oText.Tag = ""
   Next
   lblNation1.Caption = ""
   m_strDCC04 = "" 'Added by Lydia 2022/04/22
   m_strDCC02 = "" 'Added by Lydia 2024/08/30
End Sub

Private Sub ClearField()
       
   For Each oText In txtField
       oText.Text = ""
       oText.Tag = ""
   Next
   lblNation2.Caption = ""
   m_strDCC04 = "" 'Added by Lydia 2022/04/22
   m_strDCC02 = "" 'Added by Lydia 2024/08/30
End Sub

'判斷X/Y/R編號變動,取得相關資料
Private Sub GetNowGrp()
Dim tmpGrp As String

    If SSTab1.Tab = 0 Then
        If Option1(1).Value = True Then  'CF代理人
            If txt1(0) <> "" And txt1(1) <> "" Then
                StrSQLa = "select cp09, cp27, cp44 from caseprogress, fagent where cp01='" & txt1(0) & "' and cp02='" & txt1(1) & "' and cp03='" & IIf(Trim(txt1(2)) = "", "0", Trim(txt1(2))) & "' and cp04='" & IIf(Trim(txt1(3)) = "", "00", Trim(txt1(3))) & "' " & _
                                 "And CP09 < 'C' AND CP27>0 and not (CP01='CFT' and cp09>'B' and cp10='304') and CP44 is not null " & _
                                 "and substr(cp44,1,8)=fa01(+) and substr(cp44,9,1)=fa02(+) and fa01 is not null Order By CP27 Desc,CP09 Desc "
                intA = 1
                Set rsAD = ClsLawReadRstMsg(intA, StrSQLa)
                If intA = 1 Then
                    tmpGrp = "" & rsAD.Fields("cp44")
                Else
                    MsgBox "找不到CF代理人！", vbExclamation
                End If
            End If
        ElseIf Option1(0).Value = True Then
            If txt1(0) <> "" And txt1(1) <> "" Then
               tmpGrp = PUB_GetA1K03(txt1(0), txt1(1), IIf(Trim(txt1(2)) = "", "0", Trim(txt1(2))), IIf(Trim(txt1(3)) = "", "00", Trim(txt1(3))))
               If tmpGrp = "" Then
                   MsgBox "找不到FC代理人！", vbExclamation
               End If
            End If
        ElseIf Trim(txtCNo) <> "" Then
             strExc(1) = Pub_GetNameBYnation(txtCNo, "1")
             If strExc(1) = "" Then
                MsgBox "找不到申請人/代理人/潛在客戶資料！", vbExclamation
             Else
                tmpGrp = ChangeCustomerL(txtCNo)
             End If
        End If
        If tmpGrp <> strNowGrp Then   '重新讀取資料
            ClearAdd
            If tmpGrp <> "" Then
                If GetNowData(tmpGrp) = True Then
                    strNowGrp = tmpGrp
                End If
            End If
        End If
    End If
End Sub

Private Function CheckDataValidate() As Boolean

   CheckDataValidate = True
End Function

Private Function GetNowData(ByVal pKeyNo As String, Optional ByVal bolLog As Boolean = True) As Boolean
   
   If bolLog = True Then ClearQueryLog (Me.Name)
   GetNowData = False
   Select Case Left(pKeyNo, 1)
        Case "X"
            If bolLog = True Then pub_QL05 = pub_QL05 & ";申請人:" & pKeyNo
            '若無英文資料,改抓中文->日文
            'Memo by Lydia 2022/04/22 增加DHL_Country_Code(Table)判斷該國是否要輸入郵遞區號
            strSql = "SELECT DECODE(CU59||CU62,NULL,DECODE(CU58||CU61,NULL,NVL(CU60,CU63),NVL(CU58,CU61)) ,NVL(CU59,CU62)) C00, " & _
                     "DECODE(CU05||CU88||CU89||CU90,NULL,NVL(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90) C01, " & _
                     "NVL(CU24,NVL(SUBSTR(CU23,1,18),SUBSTR(CU29,1,18))) C02, " & _
                     "NVL(CU25,NVL(SUBSTR(CU23,19,36),SUBSTR(CU29,19,36))) C03, " & _
                     "NVL(CU26,NVL(SUBSTR(CU23,37,54),SUBSTR(CU29,37,54))) C04, " & _
                     "NVL(CU27,NVL(SUBSTR(CU23,55,72),SUBSTR(CU29,55,72))) C05, " & _
                     "NVL(CU28,NVL(SUBSTR(CU23,73,80),SUBSTR(CU29,73,80)))||rtrim(' '||CU102) C06, " & _
                     "NVL(CU16,CU17) C07,SUBSTR(NA01,1,3) NA01,NA03,NA89,DCC02, DCC04, " & _
                     "DID03, DID04, DID05, DID06 " & _
                     "FROM CUSTOMER, NATION, DHL_COUNTRY_CODE , DHL_INPUT_DATA WHERE " & _
                     "CU01=SUBSTR('" & pKeyNo & "',1,8) AND CU02(+)=SUBSTR('" & pKeyNo & "',9,1) " & _
                     "AND NVL(CU87,CU10)=NA01(+) AND NA89=DCC01(+)AND CU01=DID01(+) AND CU02=DID02(+)"
        Case "Y"
            If bolLog = True Then pub_QL05 = pub_QL05 & ";代理人:" & pKeyNo
            '若無英文資料,改抓中文->日文
               'Memo by Lydia 2022/04/22 增加DHL_Country_Code(Table)判斷該國是否要輸入郵遞區號
                strSql = "SELECT DECODE(FA08||FA53,NULL,DECODE(FA07||FA52,NULL,NVL(FA09,FA54),NVL(FA07,FA52)) ,NVL(FA08,FA53)) C00, " & _
                         "DECODE(FA05||FA63||FA64||FA65,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65) C01, " & _
                         "NVL(FA18,NVL(SUBSTR(FA17,1,18),SUBSTR(FA23,1,18))) C02, " & _
                         "NVL(FA19,NVL(SUBSTR(FA17,19,36),SUBSTR(FA23,19,36))) C03, " & _
                         "NVL(FA20,NVL(SUBSTR(FA17,37,54),SUBSTR(FA23,37,54))) C04, " & _
                         "NVL(FA21,NVL(SUBSTR(FA17,55,72),SUBSTR(FA23,55,72))) C05, " & _
                         "NVL(FA22,NVL(SUBSTR(FA17,73,80),SUBSTR(FA23,73,80)))||rtrim(' '||FA70) C06, " & _
                         "NVL(FA12,FA13) C07,SUBSTR(NA01,1,3) NA01,NA03,NA89,DCC02, DCC04, " & _
                         "DID03, DID04, DID05, DID06 " & _
                         "FROM FAGENT, NATION, DHL_COUNTRY_CODE, DHL_INPUT_DATA WHERE " & _
                         "FA01=SUBSTR('" & pKeyNo & "',1,8) AND FA02=SUBSTR('" & pKeyNo & "',9,1) " & _
                         "AND NVL(FA55,FA10)=NA01(+) AND NA89=DCC01(+)AND FA01=DID01(+) AND FA02=DID02(+)"

        Case "R"  '只抓潛在客戶,不抓國內潛在客戶
            If bolLog = True Then pub_QL05 = pub_QL05 & ";潛在客戶:" & pKeyNo
            '若無英文資料,改抓中文->日文
            'Memo by Lydia 2022/04/22 增加DHL_Country_Code(Table)判斷該國是否要輸入郵遞區號
            strSql = "SELECT '' C00, DECODE(PCU03||PCU04||PCU05||PCU06,NULL,NVL(PCU08,PCU07),PCU03||' '||PCU04||' '||PCU05||' '||PCU06) C01," & _
                     "NVL(PCU20,NVL(SUBSTR(PCU27,1,18),SUBSTR(PCU26,1,18))) C02," & _
                     "NVL(PCU21,NVL(SUBSTR(PCU27,19,36),SUBSTR(PCU26,19,36))) C03," & _
                     "NVL(PCU22,NVL(SUBSTR(PCU27,37,54),SUBSTR(PCU26,37,54))) C04," & _
                     "NVL(PCU23,NVL(SUBSTR(PCU27,55,72),SUBSTR(PCU26,55,72))) C05," & _
                     "NVL(PCU24,NVL(SUBSTR(PCU27,73,80),SUBSTR(PCU26,73,80)))||rtrim(' '||PCU25) C06," & _
                     "NVL(PCU13,PCU14) C07,SUBSTR(NA01,1,3) NA01,NA03,NA89,DCC02, DCC04, " & _
                     "DID03, DID04, DID05, DID06 " & _
                     "FROM POTCUSTOMER, NATION, DHL_COUNTRY_CODE, DHL_INPUT_DATA WHERE " & _
                     "PCU01=SUBSTR('" & pKeyNo & "',1,8) AND PCU02(+)=SUBSTR('" & pKeyNo & "',9,1) " & _
                     "AND NVL(PCU28,PCU09)=NA01(+) AND NA89=DCC01(+)AND PCU01=DID01(+) AND PCU02=DID02(+)"
   End Select
   
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strSql)
   If intA = 0 Then
       If bolLog = True Then InsertQueryLog (0)
       MsgBox Mid(pub_QL05, 2) & "查無資料!", vbCritical
       Exit Function
   Else
       If bolLog = True Then InsertQueryLog (rsAD.RecordCount)
       '聯絡人
       strTemp(1) = "" & rsAD.Fields("C00")
       If "" & rsAD.Fields("DID05") <> "" Then strTemp(1) = "" & rsAD.Fields("DID05")
       txtAdd(1) = strTemp(1)
       '英文名稱
       strTemp(2) = "" & rsAD.Fields("C01")
       '英文地址=收件地址
       strTemp(3) = "" & rsAD.Fields("C02")
       strTemp(4) = "" & rsAD.Fields("C03")
       strTemp(5) = "" & rsAD.Fields("C04")
       strTemp(6) = "" & rsAD.Fields("C05")
       strTemp(7) = "" & rsAD.Fields("C06")
       StrSQLa = ""
       If "" & rsAD.Fields("C02") <> "" Then StrSQLa = StrSQLa & rsAD.Fields("C02") & vbCrLf
       If "" & rsAD.Fields("C03") <> "" Then StrSQLa = StrSQLa & rsAD.Fields("C03") & vbCrLf
       If "" & rsAD.Fields("C04") <> "" Then StrSQLa = StrSQLa & rsAD.Fields("C04") & vbCrLf
       If "" & rsAD.Fields("C05") <> "" Then StrSQLa = StrSQLa & rsAD.Fields("C05") & vbCrLf
       If "" & rsAD.Fields("C06") <> "" Then StrSQLa = StrSQLa & rsAD.Fields("C06") & vbCrLf
       txtAdd(0) = StrSQLa
       '電話
       strTemp(9) = "" & rsAD.Fields("C07")
       If "" & rsAD.Fields("DID06") <> "" Then
           strTemp(9) = "" & rsAD.Fields("DID06")  'DHL收件人電話
       End If
       txtAdd(2) = strTemp(9)
       '申請人/代理人/潛在客戶
       If Option1(0).Value = True Or Option1(1).Value = True Then
           strTemp(8) = strNum(0) & "-" & strNum(1) & "-" & strNum(2) & "-" & strNum(3)
       Else
           strTemp(8) = Left(LTrim(RTrim(txtCNo)) & "000000000", 9)
          If OptKind(0).Value = True Then
             strTemp(8) = Trim(txtNo(0)) & "-" & Trim(txtNo(1)) & (IIf(Len(txtNo(2)) > 0, "-" & txtNo(2), "-0")) & (IIf(Len(txtNo(3)) > 0, "-" & txtNo(3), "-00"))
          Else
             strTemp(8) = Trim(Text2)
          End If
          If Len(Text3) > 0 Then strTemp(1) = LTrim(RTrim(Text3))
       End If
       '郵遞區號DID03
       strTemp(10) = "" & rsAD.Fields("DID03")
       txtAdd(4) = strTemp(10)
       '收件國家名稱DCC02
       strTemp(11) = "" & rsAD.Fields("DCC02")
       txtAdd(5) = "" & rsAD.Fields("NA01")
       lblNation1.Caption = "" & rsAD.Fields("NA03")
       m_strDCC02 = "" & rsAD.Fields("DCC02") 'Added by Lydia 2024/08/30
       recNA89 = "" & rsAD.Fields("NA89")
       m_strDCC04 = "" & rsAD.Fields("DCC04") 'Added by Lydia 2022/04/22
       'Added by Lydia 2024/08/30 排除不用郵遞區號的國家>>API模式要求正確資料
       'Mark by Lydia 2024/10/18 越南為非必要郵遞區號的國家,但是又輸入郵遞區號 --- from 玫音
       'If InStr(UCase(m_strDCC04), "POSTCODE") = 0 Then
       '   strTemp(10) = ""
       '   txtAdd(4) = strTemp(10)
       'End If
       ''end 2024/08/30
       'end 2024/10/18
       
       '收件城市
       recDID04 = "" & rsAD.Fields("DID04")
       txtAdd(3) = recDID04
       For Each oText In txtAdd
            oText.Tag = oText.Text
       Next
   End If
   GetNowData = True
End Function

Private Sub txtAdd_Validate(Index As Integer, Cancel As Boolean)
Dim strTmpA As String

    If txtAdd(Index) = "" Then Exit Sub
    
    If txtAdd(Index).Tag <> txtAdd(Index).Text Then
        strTmpA = PUB_GetSimpleName(txtAdd(Index), True, True)
        If strTmpA <> txtAdd(Index).Text Then
           MsgBox "請注意" & Replace(Label6(Index).Caption, ":", "") & "包含非英數字！"
        End If
        If Index = 5 Then '國家代號=NA01=NA89
            lblNation1.Caption = ""
            'Modified by Lydia 2022/04/22
            'intI = ClsPDGetNation(txtAdd(Index), strTmpA)
            'Modified by Lydia 2024/08/30 + m_strDCC02
            'Modified by Lydia 2024/09/04 +recNA89
            strTmpA = GetNationName(txtAdd(Index), m_strDCC04, m_strDCC02, recNA89)
            If strTmpA = "" Then
                MsgBox "請輸入正確國家代號或是按？按鈕選擇！", vbExclamation
                Cancel = True
                Exit Sub
            Else
                lblNation1 = strTmpA
            End If
        End If
    End If
    
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
Dim strTmpA As String
    
    If Index = 7 And Trim(txtField(7)) = "" Then lblNation2.Caption = "" 'Added by Lydia 2022/04/22
    
    'Added by Lydia 2025/10/28
    If Index = 4 Then
      Cancel = False
      '備註只有35字元
      If CheckLengthIsOK(txtField(Index), 35) = False Then
         Cancel = True
         txtField_GotFocus Index
         Exit Sub
      End If
    End If
    'end 2025/10/28
    
    If txtField(Index) <> "" And txtField(Index).Tag <> txtField(Index).Text Then
        strTmpA = PUB_GetSimpleName(txtField(Index), True, True)
        If strTmpA <> txtField(Index).Text Then
           MsgBox "請注意" & Replace(Label8(Index).Caption, ":", "") & "包含非英數字！"
        End If
        If Index = 7 Then '國家代號=NA01=NA89
            lblNation2.Caption = ""
            'Modified by Lydia 2022/04/22
            'intI = ClsPDGetNation(txtAdd(Index), strTmpA)
            'Modified by Lydia 2024/08/30 + m_strDCC02
            'Modified by Lydia 2024/09/04 +recNA89
            strTmpA = GetNationName(txtField(Index), m_strDCC04, m_strDCC02, recNA89)
            If strTmpA = "" Then
                MsgBox "請輸入正確國家代號或是按？按鈕選擇！", vbExclamation
                Cancel = True
                Exit Sub
            Else
                lblNation2 = strTmpA
            End If
        End If
    End If
    txtField(Index).Tag = txtField(Index).Text
End Sub

Private Function SaveDID(ByVal inTab As Integer) As Boolean

    SaveDID = False
    
    If inTab = 0 Then
        If strNowGrp <> "" Then
            StrSQLa = "select did01 from dhl_input_data where did01='" & Mid(strNowGrp, 1, 8) & "'  and did02='" & Mid(strNowGrp, 9, 1) & "' "
            intA = 1
            strSql = ""
            Set rsAD = ClsLawReadRstMsg(intA, StrSQLa)
            If intA = 0 Then
                strSql = "Insert Into DHL_INPUT_DATA (DID01,DID02,DID03,DID04,DID05,DID06) Values " & _
                                 "('" & Mid(strNowGrp, 1, 8) & "','" & Mid(strNowGrp, 9, 1) & "',"
                '郵遞區號DID03, 收件城市DID04, 收件聯絡人DID05, 收件人電話DID06
                strSql = strSql & CNULL(ChgSQL(Trim(txtAdd(4)))) & "," & CNULL(ChgSQL(Trim(txtAdd(3)))) & "," & CNULL(ChgSQL(Trim(txtAdd(1)))) & "," & CNULL(ChgSQL(Trim(txtAdd(2))))
                strSql = strSql & ") "
            Else
                If txtAdd(1).Tag <> txtAdd(1).Text Then strSql = strSql & ", did05=" & CNULL(ChgSQL(Trim(txtAdd(1))))
                If txtAdd(2).Tag <> txtAdd(2).Text Then strSql = strSql & ", did06=" & CNULL(ChgSQL(Trim(txtAdd(2))))
                If txtAdd(3).Tag <> txtAdd(3).Text Then strSql = strSql & ", did04=" & CNULL(ChgSQL(Trim(txtAdd(3))))
                If txtAdd(4).Tag <> txtAdd(4).Text Then strSql = strSql & ", did03=" & CNULL(ChgSQL(Trim(txtAdd(4))))
                If strSql <> "" Then
                   strSql = "Update DHL_INPUT_DATA SET " & Mid(strSql, 2)
                   strSql = strSql & " where did01='" & Mid(strNowGrp, 1, 8) & "'  and did02='" & Mid(strNowGrp, 9, 1) & "' "
                End If
            End If
            If strSql <> "" Then cnnConnection.Execute strSql
            '重設變數
            Call GetNowData(strNowGrp, False)
        End If
    ElseIf inTab = 1 Then '人工填提單
        ClearQueryLog (Me.Name)
        
        pub_QL05 = pub_QL05 & ";人工填提單;"
        '收件聯絡人
        strTemp(1) = StringFilterr(txtField(2).Text)
        If Trim(strTemp(1)) <> "" Then pub_QL05 = pub_QL05 & "收件聯絡人:" & strTemp(1) & ";"
        '收件公司
        strTemp(2) = StringFilterr(txtField(0).Text)
        If Trim(strTemp(2)) <> "" Then pub_QL05 = pub_QL05 & "收件公司:" & strTemp(2) & ";"
        '收件地址
        strTemp(3) = StringFilterr(txtField(1).Text)
        If Trim(strTemp(3)) <> "" Then pub_QL05 = pub_QL05 & "收件地址:" & strTemp(3) & ";"
        '收件地址分5行
        strTemp1 = Empty
        strTemp1 = Split(txtField(1), vbNewLine)
        j = 3
        For i = 0 To UBound(strTemp1)
             If Trim(strTemp1(i)) <> "" Then
                 strTemp(j) = strTemp1(i)
                 j = j + 1
             End If
        Next i
        If j <= 7 Then
            For i = j To 7
                strTemp(i) = ""
            Next i
        End If
        
        '收件人電話
       strTemp(9) = StringFilterr(txtField(3).Text)
       If Trim(strTemp(9)) <> "" Then pub_QL05 = pub_QL05 & "收件人電話:" & strTemp(9) & ";"
        '說明
       strTemp(8) = StringFilterr(txtField(4).Text)
       If Trim(strTemp(8)) <> "" Then pub_QL05 = pub_QL05 & "說明:" & strTemp(8) & ";"
       
           '收件國家名稱;
           strTemp(11) = "": lblNation2.Caption = "": recNA89 = ""
           strSql = "select na01,na03,na89,dcc02,dcc04 from nation,dhl_country_code where na01='" & txtField(7) & "' and NA89=DCC01(+)"
           intA = 1
           Set rsAD = ClsLawReadRstMsg(intA, strSql)
           If intA = 1 Then
              strTemp(11) = "" & rsAD.Fields("DCC02")
              lblNation2.Caption = "" & rsAD.Fields("NA03")
              recNA89 = "" & rsAD.Fields("NA89")
              m_strDCC04 = "" & rsAD.Fields("DCC04") 'Added by Lydia 2022/04/22
              m_strDCC02 = "" & rsAD.Fields("DCC02") 'Added by Lydia 2024/08/30
              'Added by Lydia 2024/08/30 排除不用郵遞區號的國家>>API模式要求正確資料
              If InStr(UCase(m_strDCC04), "POSTCODE") = 0 Then
                 txtField(6).Text = ""
              End If
              'end 2024/08/30
              If strTemp(11) <> "" Then pub_QL05 = pub_QL05 & "收件國家:" & txtField(7) & lblNation2 & ";"
           End If
           
           '郵遞區號DID03
           strTemp(10) = StringFilterr(txtField(6).Text)
           If strTemp(10) <> "" Then pub_QL05 = pub_QL05 & "郵遞區號:" & strTemp(8) & ";"

           '收件城市
           recDID04 = StringFilterr(txtField(5).Text)
           If recDID04 <> "" Then pub_QL05 = pub_QL05 & "收件城市:" & recDID04 & ";"
           
       Call InsertQueryLog("1")

    End If
    
    SaveDID = True
    Exit Function
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox "存檔失敗(DHL_Input_DHL): " & Err.Description
    End If
End Function
'end 2022/03/28

'Added by Lydia 2022/04/06 用Word範本套印 ; 2022/4/1 May-列印需要萬國碼給人員和國外客戶查看
Private Sub PrintWordA4(ByVal mBarCode As String, ByVal mRFno As String, Optional ByVal bShow As Boolean = False)
Dim strText
Dim intA As Integer, iRound As Integer
Dim m_FileName As String, m_TempFileName As String
Dim m_DefPath As String, strName As String
Dim oShape
Dim oWord
    
   m_DefPath = App.path & "\" & strUserNum
   Pub_ChkExcelPath m_DefPath
   
   '變更Word印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_SetWordActivePrinter
   
   '下載範本檔
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000048-0-01 DHL Word套印.docx", "M51", "000048", "0", "01", "4", "2")
   
   m_FileName = "$$" & strUserNum & "_DHL.docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000048-0-01", , m_DefPath) = False Then
        Exit Sub
   End If
   
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
   End If

   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False

   '(保留) 找出特定TextBox名稱
'   strExc(1) = ""
'   For intI = 1 To g_WordAp.ActiveDocument.Shapes.Count
'         If InStr(UCase(g_WordAp.ActiveDocument.Shapes(intI).Name), "BOX") > 0 Or InStr(UCase(g_WordAp.ActiveDocument.Shapes(intI).Name), "文字方塊") > 0 Then
'            strExc(1) = strExc(1) & "Name: " & g_WordAp.ActiveDocument.Shapes(intI).Name & vbCrLf & _
'                                  "     Text:" & g_WordAp.ActiveDocument.Shapes(intI).TextFrame.TextRange.Text
''            If InStr(g_WordAp.ActiveDocument.Shapes(intI).TextFrame.TextRange.Text, "PS118") > 0 Then
''                g_WordAp.ActiveDocument.Shapes(intI).Name = "Text Box 118"
''            End If
''            If InStr(g_WordAp.ActiveDocument.Shapes(intI).TextFrame.TextRange.Text, "PS218") > 0 Then
''                g_WordAp.ActiveDocument.Shapes(intI).Name = "Text Box 218"
''            End If
''            If InStr(g_WordAp.ActiveDocument.Shapes(intI).TextFrame.TextRange.Text, "PS318") > 0 Then
''                g_WordAp.ActiveDocument.Shapes(intI).Name = "Text Box 318"
''            End If
'         Else
'            strExc(1) = strExc(1) & "Name: " & g_WordAp.ActiveDocument.Shapes(intI).Name & vbCrLf
'         End If
'   Next intI
'   Debug.Print strExc(1)
   'end 2018/05/28
   
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For iRound = 1 To 3 '列印3份
          For intA = 0 To 18
             strName = "PS" & Format(iRound * 100 + intA, "000")
             strText = ""
             If intA = 0 Then
                  '提單號碼
                  strText = " " & Mid(mBarCode, 1, 3) & " " & Mid(mBarCode, 4, 4) & " " & Mid(mBarCode, 8)
             ElseIf intA = 1 Then
                  'FTP傳送是否正確
                  If bFTPok = False Then
                     strText = "TX err"
                  Else
                     strText = String(6, " ")
                  End If
             ElseIf intA = 2 Then
                  '提單號碼BarCode
             ElseIf intA = 3 Or intA = 4 Then '付款帳號
                    strText = PayACNO
             ElseIf intA = 5 Then '本所案號／說明
                    strText = mRFno
             ElseIf intA = 6 Then '本所公司名稱
                    strText = PayACName
             ElseIf intA = 7 Then '寄件地址1
                    strText = PayACaddr01
             ElseIf intA = 8 Then '寄件地址2
                    strText = PayACaddr02
             ElseIf intA = 9 Then '寄件郵遞區號+電話
                    strText = PayACZip & String(18, " ") & PayACTel
             ElseIf intA = 10 Then '收件公司
                    'Modified by Lydia 2022/06/07 因為過長會影響版面,所以限制長度PUB_StrToStr; ex.FCP-58208之FC代理人
                    strText = PUB_StrToStr(IIf(bShow, "收件公司ABCDEFR", strTemp(2)), lenRA)
             ElseIf intA = 11 Then '收件地址1
                    strText = IIf(bShow, recAddr(1), strTemp(3))
             ElseIf intA = 12 Then '收件地址2
                    strText = IIf(bShow, recAddr(2), strTemp(4))
             ElseIf intA = 13 Then '收件地址3
                    strText = IIf(bShow, recAddr(3), strTemp(5))
             ElseIf intA = 14 Then '收件地址4
                    strText = IIf(bShow, "　", strTemp(6))
             ElseIf intA = 15 Then '收件地址5
                    strText = IIf(bShow, "　", strTemp(7))
             ElseIf intA = 16 Then '收件郵遞區號+國家名稱
                    If bShow = True Then
                         strText = PUB_StrToStr("郵遞區號-DID03", 24, True) & "收件國家-DCC04"
                    Else
                         strText = PUB_StrToStr(IIf(strTemp(10) = "", " ", strTemp(10)), 24, True) & strTemp(11)
                    End If
             ElseIf intA = 17 Then '收件聯絡人+電話+ 寄件聯絡
                    If bShow = True Then
                         strText = PUB_StrToStr("收件聯絡人-DID05", 24, True) & PUB_StrToStr("收件人電話-DID06", 20, True) & String(8, " ") & PUB_StrToStr(PayContact, 18, True)
                    Else
                         strText = PUB_StrToStr(IIf(strTemp(1) = "", " ", strTemp(1)), 24, True) & PUB_StrToStr(IIf(strTemp(9) = "", " ", strTemp(9)), 20, True) & String(8, " ") & PUB_StrToStr(PayContact, 18, True)
                    End If
             ElseIf intA = 18 Then '列印年月日(文字方塊)
                    strText = Mid(strSrvDate(1), 1, 4) & "  " & Mid(strSrvDate(1), 5, 2) & "  " & Mid(strSrvDate(1), 7, 2)
                    .ActiveDocument.Shapes("Text Box " & Format(iRound * 100 + intA, "000")).Select
                    .Selection.TypeText Text:=strText
                    .Selection.HomeKey Unit:=wdStory
             End If
             
             If Trim(strName) <> "" And intA <> 18 Then
                .Selection.Find.ClearFormatting
                .Selection.Find.Text = "|#" & strName & "#|"
                .Selection.Find.Replacement.Text = ""
                .Selection.Find.Forward = True
                .Selection.Find.Wrap = wdFindContinue
                .Selection.Find.Format = False
                .Selection.Find.MatchCase = False
                .Selection.Find.MatchWholeWord = False
                .Selection.Find.MatchWildcards = False
                .Selection.Find.MatchSoundsLike = False
                .Selection.Find.MatchAllWordForms = False
                .Selection.Find.MatchByte = True
                .Selection.Find.Execute
                .Selection.Delete
                If intA = 0 Then  '提單號碼
                    .Selection.Font.Size = 19
                End If
                If intA = 2 Then '提單號碼BarCode
                    Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=saveBCfile, LinkToFile:=False, SaveWithDocument:=True)
                    oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    'Modified by Lydia 2022/04/12 使用者的圖片寬度會等比開放
                    'oShape.Height = .CentimetersToPoints(2)
                    oShape.LockAspectRatio = msoFalse
                    oShape.Height = .CentimetersToPoints(2)
                    oShape.Width = .CentimetersToPoints(6.48)
                    'end 2022/04/12
                    oShape.Top = .CentimetersToPoints(1.4)
                    oShape.LockAnchor = False
                    oShape.LayoutInCell = True
                    oShape.WrapFormat.AllowOverlap = True
                    oShape.WrapFormat.Side = wdWrapBoth
                    oShape.WrapFormat.Type = 3
                End If
                '.Selection.Font.ColorIndex = wdBlack
                .Selection.TypeText strText
             End If
          Next intA
      Next iRound  '列印3份
   End With
    
    g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=1, Pages:="1-2", Collate:=True
   
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop '還原Word位置
   
   '保留: 存檔
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
      
  '刪除已完成上傳和列印的提單資料
  If bShow = False Then
     Kill strFileName
  End If
  ShowPrintOk
  
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Number & ":" & Err.Description, , "錯誤 "
   End If
End Sub

'Added by Lydia 2022/04/22
'Modified by Lydia 2024/08/30 +pEngName
'Modified by Lydia 2024/09/04 +pNaCode
Private Function GetNationName(ByVal pNo As String, ByRef pZip As String, ByRef pEngName As String, ByRef pNaCode As String) As String
Dim strA1 As String
Dim rsA1 As New ADODB.Recordset
Dim intX As Integer
    
   GetNationName = ""
   pZip = ""
   If pNo <> "" Then
       strA1 = "SELECT NA01,NA03,DCC02,NA89,DCC04 FROM NATION,DHL_COUNTRY_CODE " & _
                  "WHERE NVL(NA89,'00') <>'00' AND LENGTH(NA01)='3' AND NA01> '010' AND NA89=DCC01(+) AND NA01 =" & CNULL(pNo)
       intX = 1
       Set rsA1 = ClsLawReadRstMsg(intX, strA1)
       If intX = 1 Then
           GetNationName = "" & rsA1.Fields("NA03")
           pZip = "" & rsA1.Fields("DCC04")
           pEngName = "" & rsA1.Fields("DCC02") 'Added by Lydia 2024/08/30
           pNaCode = "" & rsA1.Fields("na89") 'Added by Lydia 2024/09/04
       End If
       Set rsA1 = Nothing
   End If
End Function

'Added by Lydia 2024/08/30 MyDHL API整合:直接透過http模式request(送出提單)和response(取得提單)
Private Function PrintDataFromAPI(ByVal pStatus As String) As Boolean
'pStatus：1-DHL正式環境, 2-DHL測試環境, 3-測試API連線
On Error GoTo 0
Dim strRequest As String, strResponse As String, strStatus As String
Dim oDHttp As New WinHttp.WinHttpRequest
Dim strDTime As String, strShipNo As String
Dim strB01 As String, intRetry As Integer

   
   strShipNo = ""
   strDTime = Format(ServerTime, "00:00:00")
   strB01 = Replace(ChangeTStringToTDateString(strSrvDate(1)), "/", "-") & "T" & strDTime & " GMT+01:00" 'DTEST00>>寄件日期格式:2022-06-07T23:00:31GMT+01:00
   strDTime = Replace(strDTime, ":", "")
   If pStatus = "3" Then  '測試API連線
      '傳入字串用|區隔
      strB01 = strB01 & "|FCP-064642-0-00"    'DTEST01>>備註
      strB01 = strB01 & "|Test Name"          'DTEST02>>員工英文別名
      strB01 = strB01 & "|The Woodlands"      'DTEST03>>收件人-城市
      strB01 = strB01 & "|US"                 'DTEST04>>DHL國家代號
      strB01 = strB01 & "|77380"              'DTEST05>>收件郵遞區號
      strB01 = strB01 & "|21 Waterway Avenue, Suite 300," 'DTEST06>>收件地址01
      strB01 = strB01 & "|The Woodlands, Texas 77380, U. S. A" 'DTEST07>>收件地址02
      strB01 = strB01 & "|"                   'DTEST08>>收件地址03
      strB01 = strB01 & "|+1 281- 362-2839"   'DTEST09>>收件人電話
      strB01 = strB01 & "|Angelo IP"          'DTEST10>>收件公司
      strB01 = strB01 & "|Mr. Basil M. Angelo" 'DTEST11>>收件人/公司名稱
      strB01 = strB01 & "|FCP-064642-0-00"    'DTEST12>>備註
   Else
      '處理收件地址長度(1~3欄)
      strExc(1) = "" 'Added by Lydia 2024/09/09
      For intA = 3 To 7
         If strTemp(intA) <> "" Then
            strExc(1) = strExc(1) & IIf(strExc(1) <> "", ", ", "") & strTemp(intA)
         Else
            If intA < 7 And strTemp(intA + 1) = "" Then Exit For
         End If
      Next
      strExc(1) = Replace(strExc(1), ",,", ",")  '原本地址就用,區隔；ex.Y20065
      '先排除非英數字
      strExc(1) = Trim(PUB_GetSimpleName(strExc(1), True, True))
      If Right(strExc(1), 1) = "," Then
          strExc(1) = Mid(strExc(1), 1, Len(strExc(1)) - 1)
      End If
      If LenB(StrConv(strExc(1), vbFromUnicode)) <= lenRA Then
         recAddr(1) = strExc(1)
      Else
         For intA = 1 To 3
            recAddr(intA) = PUB_StrToStr(strExc(1), lenRA)
            '調整斷行
            If LenB(StrConv(recAddr(intA), vbFromUnicode)) = lenRA _
                And Right(recAddr(intA), 1) <> "" And Right(recAddr(intA), 1) <> "," And Right(recAddr(intA), 1) <> "-" Then
               recAddr(intA) = Mid(recAddr(intA), 1, InStrRev(recAddr(intA), " "))
            End If
            strExc(1) = Mid(strExc(1), Len(recAddr(intA)) + 1)
            If Trim(strExc(1)) = "" Then Exit For
         Next
      End If
      If GetTextLength(recAddr(1) & recAddr(2) & recAddr(3)) > lenRA * 3 Then
         '已問過DHL工程師，目前地址只有３行，請儘量減少字元，例如郵遞區號
         MsgBox "請將英文地址長度簡化到" & lenRA * 3 & "字元以內！", vbCritical, "MyDHL API整合"
         PrintDataFromAPI = True '避免用Word套印
         Exit Function
      End If
      
      '取得使用者的英文名稱; 若虛編號沒有分機記錄，改用ST12
      PayContact = ""
      strSql = " select st12 as ed04 from staff where st01='" & strUserNum & "' "
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strSql)
      If intA = 1 Then
         PayContact = "" & rsAD.Fields("ED04")
      End If
      If PayContact = "" Then   '改抓分機表的英文別名
         strExc(1) = Pub_GetStaffExtn(strUserNum, PayContact)
      End If
      '傳入字串用|區隔
      strB01 = strB01 & "|" & Trim(strTemp(8))     'DTEST01>>備註,本所案號／說明
      strB01 = strB01 & "|" & PayContact           'DTEST02>>員工英文別名
      strB01 = strB01 & "|" & recDID04             'DTEST03>>收件人-城市
      strB01 = strB01 & "|" & recNA89              'DTEST04>>DHL國家代號
      strB01 = strB01 & "|" & strTemp(10)          'DTEST05>>收件郵遞區號
      strB01 = strB01 & "|" & recAddr(1)           'DTEST06>>收件地址01
      strB01 = strB01 & "|" & recAddr(2)           'DTEST07>>收件地址02
      strB01 = strB01 & "|" & recAddr(3)           'DTEST08>>收件地址03
      strB01 = strB01 & "|" & Trim(strTemp(9))     'DTEST09>>收件人電話
      strB01 = strB01 & "|" & Trim(strTemp(2))     'DTEST10>>收件公司
      'Modified by Lydia 2024/10/22 debug
      'strB01 = strB01 & "|" & Trim(IIf(Trim(strTemp(1)) = "", strTemp(2), strTemp(2)))   'DTEST11>>收件人/公司名稱
      strB01 = strB01 & "|" & Trim(IIf(Trim(strTemp(1)) = "", strTemp(2), strTemp(1)))
      strB01 = strB01 & "|" & Trim(strTemp(8))     'DTEST12>>備註
      '去掉"符號(JSON專用符號)
      strB01 = Replace(strB01, """", "")
   End If
   
   '目前可能發生的錯誤，已問過DHL工程師回答如下:
   '國家的城市沒有特定對照檔，只好等傳送發生錯誤，再來檢查資料
   '聯絡人還是只能輸入英數字
   
   '先留下DHL使用記錄
   strSql = "insert into dhluseno (DUN01,DUN02,DUN03,DUN04,DUN05,DUN06) VALUES " & _
           "('" & Val(strShipNo) & "', '" & ChgSQL(strB01) & "', '" & IIf(pStatus = "1", "1", "2") & "'," & _
           "'" & strUserNum & "', '" & strSrvDate(1) & "', '" & strDTime & "') "
   cnnConnection.Execute strSql
   
   strRequest = MakeJSONforDHL(strB01)  '轉換資料格式為JSON
'--------------

   Set oDHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
   oDHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
   oDHttp.SetTimeouts 10000, 10000, 10000, 10000  'Resolve, Connect, Send and Receive
   'false = 非動態>>要等到有回覆
   If pStatus = "1" Then 'DHL正式環境
      oDHttp.Open "POST", "https://express.api.dhl.com/mydhlapi/shipments", False
   Else                  '測試環境
      oDHttp.Open "POST", "https://express.api.dhl.com/mydhlapi/test/shipments", False
   End If
   
   'P.S. DHL只會針對實際收到的提單=文件來計算費用，不用擔心開發中上傳的資料會被計入
   oDHttp.SetRequestHeader "Content-Type", "application/json" '先省略 ; charset=UTF-8"
   oDHttp.SetRequestHeader "Accept", "application/json"
   oDHttp.SetRequestHeader "Authorization", "Basic " + cntAuidPWD
JumpToRetry:
   oDHttp.Send "{" & strRequest & "}"
   
   strStatus = oDHttp.StatusText
   If oDHttp.Status = 201 Then  'shipment提單建立成功
      strResponse = oDHttp.ResponseText
      '取得提單號碼：通過計算位置
      intI = InStr(strResponse, "shipmentTrackingNumber") - 1
      strShipNo = Mid(strResponse, intI + 26, InStr(intI, strResponse, ",") - intI - 27)
      strSql = "Update dhluseno set dun01='" & strShipNo & "' where dun04 = '" & strUserNum & "' and dun05='" & strSrvDate(1) & "' and dun06='" & strDTime & "' "
      cnnConnection.Execute strSql
      '取得PDF：通過計算位置
      intI = InStr(strResponse, "content"":") - 1
      strExc(2) = Mid(strResponse, intI + 11, InStr(intI, strResponse, ",") - intI - 12)
      If PUB_FromBase64ToFile(strExc(2), App.path & "\" & strUserNum & "\$" & strShipNo & ".pdf") = True Then
         Sleep 1000
         PUB_PrintPDF App.path & "\" & strUserNum & "\$" & strShipNo & ".pdf", Combo1.Text
         PrintDataFromAPI = True
      Else
         MsgBox "產生PDF失敗，請洽電腦中心！", vbCritical, "MyDHL API整合"
      End If
   Else
      If oDHttp.Status = 408 Then  '逾時過期
         If intRetry < 3 Then
            Sleep 5000
            intRetry = intRetry + 1
            GoTo JumpToRetry
         End If
      End If
      If strSrvDate(1) <= "20240908" Then
         'FTP運作期間直接發Email通知
         PUB_SendMail strUserNum, "A3034", "", "傳送到DHL失敗: " & IIf(oDHttp.Status = 408, "逾時過期已達3次", strStatus), "內容如下：" & vbCrLf & String(20, "=") & vbCrLf & strB01
      Else
         MsgBox "傳送到DHL失敗: " & IIf(oDHttp.Status = 408, "逾時過期已達3次", strStatus) & vbCrLf & "請洽電腦中心！", vbCritical, "MyDHL API整合"
         PUB_SendMail strUserNum, "A3034", "", "傳送到DHL失敗: " & IIf(oDHttp.Status = 408, "逾時過期已達3次", strStatus), "內容如下：" & vbCrLf & String(20, "=") & vbCrLf & strB01
      End If
   End If
   
   Set oDHttp = Nothing
   Exit Function
   
ErrHandle:
   Set oDHttp = Nothing
   MsgBox "傳輸失敗：" & vbCrLf & Err.Description, vbCritical, "MyDHL API整合"

End Function

'Added by Lydia 2024/08/30
Private Sub Command2_Click()
   
Debug.Print "測試API連線-開始:" & Format(ServerTime, "000000")

   Call PrintDataFromAPI("3")
   
Debug.Print "測試API連線-結束:" & Format(ServerTime, "000000")
   
End Sub

'Added by Lydia 2024/08/30 依輸入的郵遞區號轉換成DHL的判斷格式
Private Function GhgPostFormat(ByVal pPostNo As String) As String
Dim intX As Integer, strMidFMT As String
   
   For intX = 1 To Len(pPostNo)
      If InStr("0123456789", Mid(pPostNo, intX, 1)) > 0 Then
         strMidFMT = strMidFMT & "9"
      ElseIf Mid(pPostNo, intX, 1) = "-" Then
         strMidFMT = strMidFMT & "-"
      ElseIf Mid(pPostNo, intX, 1) = " " Then
         strMidFMT = strMidFMT & " "
      Else
         strMidFMT = strMidFMT & "A"
      End If
   Next intX
   GhgPostFormat = strMidFMT
   
End Function

'Added by Lydia 2024/08/30
'***********傳入字串，產生JSON格式
Private Function MakeJSONforDHL(ByVal pDTest) As String
Dim intX As Integer, tmpArrX As Variant
Dim strDef01 As String
'因為DHL的JSON只需傳入部份資料 , 其餘為固定設定, 所以採用固定寫法, 以下為DHL在113/08/14提供的範例;113/8/27在
'******************************
'{
'   "productCode": "D",  //Global產品代號 P=包裹 D=文件
'   "localProductCode": "D",  //與Global產品代號一樣
'   "plannedShippingDateAndTime": "2022-06-07T23:00:31GMT+01:00",  //寄件日期>>傳入Postman網站,要注意不能晚於10天前的資料
'   "pickup": {  //預約取件
'      "isRequested": false //請帶範例值
'   },
'   "accounts": [
'      {
'         "number": "620330312",  //DHL付費帳號
'         "typeCode": "shipper"  //寄件者
'      }
'   ],
'   "outputImageProperties": {  //輸出圖檔設定
'      "splitInvoiceAndReceipt": true,  //請帶範例值
'      "printerDPI": 300,
'      "encodingFormat": "pdf",  //The reference type please check on SPEC
'      "imageOptions": [
'         {
'            "typeCode": "invoice",  //發票
'            "templateName": "COMMERCIAL_INVOICE_P_10",
'            "invoiceType": "commercial",  //請帶範例值 commercial, proforma
'            "languageCode": "eng",
'            "isRequested": false  //影響發票格式，請帶範例值
'         },
'         {
'            "typeCode": "waybillDoc",  //運務單(外務人員會取走)
'            "hideAccountNumber": true,  //是否隱藏帳號
'            "templateName": "ARCH_8X4_A4_002",  //請帶範例值 (參考 Spec:原本ARCH_8X4為標籤機用,ARCH_8X4_A4_002為A4用)
'            "numberOfCopies": 2,         //外務單印2份,1份DHL收走,1份退回給寄件者
'            "isRequested": true  //影響PLT，請帶範例值
'         },
'         {
'            "typeCode": "label",  //標籤提單
'            "templateName": "ECOM26_84_001"  //請帶範例值(參考 Spec:原本ECOM26_84_001為標籤機用,ECOM26_84_A4_001為A4用)
'         }
'      ]
'   },
'   "customerReferences": [  //貨件備註
'      {
'         "value": "FCP-064642-0-00",
'         "typeCode": "CU"  //The reference type please check on SPEC
'      }
'   ],
'   "getRateEstimates": false,  //預估運費 請帶範例值
'   "customerDetails": {
'      "shipperDetails": {  //寄件人相關資訊
'         "postalAddress": {
'            "cityName": "Taipei City",
'            "countryCode": "TW",
'            "postalCode": "104",
'            "addressLine1": "9F NO 112 SEC.2",
'            "addressLine2": "CHANG AN E.ROAD, TAIPEI, TAIWAN"
'         },
'         "contactInformation": {
'            "phone": "+88625061023",
'            "companyName": "TAI E INTERNATIONAL PATENT & TRADEMARK OFFICE",
'            "fullName": "Phoebe",
'            "email": "ipexpress@taie.com.tw"
'         },
'         "typeCode": "business" //請帶範例值
'      },
'      "receiverDetails": {  //收件人相關資訊
'         "postalAddress": {
'            "cityName": "The Woodlands",
'            "countryCode": "US",
'            "postalCode": "77380",   POSTALCODEFORMAT
'            "addressLine1": "21 Waterway Avenue, Suite 300,",
'            "addressLine2": "The Woodlands, Texas 77380, U. S. A"   //地址長度45,地址行最多3，地址1必須有值不可補.
'         },
'         "contactInformation": {
'            "phone": "+1 281- 362-2839",   //電話為必填欄位，空白補.
'            "companyName": "Angelo IP",   //公司名稱
'            "fullName": "Mr. Basil M. Angelo"  //收件人名稱
'         },
'         "typeCode": "business"  //請帶範例值
'      }
'   },
'    "content": {
'        "unitOfMeasurement": "metric",  //測量單位
'        "isCustomsDeclarable": false,  //海關目的 dutiable (true) or non dutiable (false) 文件需要選false
'        "incoterm": "DAP",   //國際商業貿易術語 The reference type please check on SPEC
'        "description": "Document", //貨件描述
'        "packages": [ //包裝信息
'            {
'                "customerReferences": [  //貨件1備註
'                    {
'                        "value": "FCP-064642-0-00",
'                        "typeCode": "CU"  //The reference type please check on SPEC
'                    }
'                ],
'                "weight": 0.5, //包裹總重量
'                "dimensions": { //材積
'                    "length": 35,
'                    "width": 28,
'                    "height": 1
'                }
'            }
'        ]
'    }
'}

'*********************
   tmpArrX = Split(pDTest, "|")
              strDef01 = """productCode"": ""D""," 'Global產品代號 P=包裹 D=文件
   strDef01 = strDef01 & Trim("""localProductCode"": ""D"",")  '與Global產品代號一樣
   strDef01 = strDef01 & Trim("""plannedShippingDateAndTime"": """ & tmpArrX(0) & """,") '寄件日期格式:2022-06-07T23:00:31GMT+01:00
   strDef01 = strDef01 & Trim("""pickup"": {") '預約取件
   strDef01 = strDef01 & Trim("      ""isRequested"": false") '請帶範例值
   strDef01 = strDef01 & Trim("},")
   strDef01 = strDef01 & Trim("""accounts"": [")
   strDef01 = strDef01 & Trim("   {")
   strDef01 = strDef01 & Trim("      ""number"": ""620330312"",")  'DHL付費帳號
   strDef01 = strDef01 & Trim("      ""typeCode"": ""shipper""")   '寄件者
   strDef01 = strDef01 & Trim("   }")
   strDef01 = strDef01 & Trim("],")
   strDef01 = strDef01 & Trim("""outputImageProperties"": {")    '輸出圖檔設定
   strDef01 = strDef01 & Trim("   ""splitInvoiceAndReceipt"": true,")  '請帶範例值
   strDef01 = strDef01 & Trim("   ""printerDPI"": 300,")
   strDef01 = strDef01 & Trim("   ""encodingFormat"": ""pdf"",")   'The reference type please check on SPEC
   strDef01 = strDef01 & Trim("   ""imageOptions"": [")
   strDef01 = strDef01 & Trim("      {")
   strDef01 = strDef01 & Trim("         ""typeCode"": ""invoice"",")  '發票
   strDef01 = strDef01 & Trim("         ""templateName"": ""COMMERCIAL_INVOICE_P_10"",")
   strDef01 = strDef01 & Trim("         ""invoiceType"": ""commercial"",")  '請帶範例值 commercial, proforma
   strDef01 = strDef01 & Trim("         ""languageCode"": ""eng"",")
   strDef01 = strDef01 & Trim("         ""isRequested"": false")  '影響發票格式，請帶範例值
   strDef01 = strDef01 & Trim("      },")
   strDef01 = strDef01 & Trim("      {")
   strDef01 = strDef01 & Trim("         ""typeCode"": ""waybillDoc"",") '運務單(外務人員會取走)
   strDef01 = strDef01 & Trim("         ""hideAccountNumber"": true,")  '是否隱藏帳號
   strDef01 = strDef01 & Trim("         ""templateName"": ""ARCH_8X4_A4_002"",")  '請帶範例值 (參考 Spec:原本ARCH_8X4為標籤機用,ARCH_8X4_A4_002為A4用)
   strDef01 = strDef01 & Trim("         ""numberOfCopies"": 2,")    '外務單印2份,1份DHL收走,1份簽回給寄件者; P.S.已經向DHL資訊人員確認A4格式只有單純列印Label或外務單，沒有合併列印成一頁的功能；
   strDef01 = strDef01 & Trim("         ""isRequested"": true")  '影響PLT，請帶範例值
   strDef01 = strDef01 & Trim("      },")
   strDef01 = strDef01 & Trim("      {")
   strDef01 = strDef01 & Trim("         ""typeCode"": ""label"",")  '標籤提單
   strDef01 = strDef01 & Trim("         ""templateName"": ""ECOM26_84_A4_001""")  '請帶範例值(參考 Spec:原本ECOM26_84_001為標籤機用,ECOM26_84_A4_001為A4用)
   strDef01 = strDef01 & Trim("      }")
   strDef01 = strDef01 & Trim("   ]")
   strDef01 = strDef01 & Trim("},")
   strDef01 = strDef01 & Trim("""customerReferences"": [") '貨件備註
   strDef01 = strDef01 & Trim("   {")
   'Modified by Lydia 2025/10/28 備註只有35字元
   'strDef01 = strDef01 & Trim("      ""value"": """ & tmpArrX(1) & """,")
   strDef01 = strDef01 & Trim("      ""value"": """ & convForm(tmpArrX(1), 35) & """,")
   strDef01 = strDef01 & Trim("      ""typeCode"": ""CU""") 'The reference type please check on SPEC
   strDef01 = strDef01 & Trim("   }")
   strDef01 = strDef01 & Trim("],")
   strDef01 = strDef01 & Trim("""getRateEstimates"": false,")  '預估運費 請帶範例值
   strDef01 = strDef01 & Trim("""customerDetails"": {")
   strDef01 = strDef01 & Trim("   ""shipperDetails"": {")  '寄件人相關資訊
   strDef01 = strDef01 & Trim("      ""postalAddress"": {")
   strDef01 = strDef01 & Trim("         ""cityName"": ""Taipei City"",")
   strDef01 = strDef01 & Trim("         ""countryCode"": ""TW"",")
   strDef01 = strDef01 & Trim("         ""postalCode"": ""104"",")
   strDef01 = strDef01 & Trim("         ""addressLine1"": ""9F NO 112 SEC.2"",")
   strDef01 = strDef01 & Trim("         ""addressLine2"": ""CHANG AN E.ROAD, TAIPEI, TAIWAN""")
   strDef01 = strDef01 & Trim("      },")
   strDef01 = strDef01 & Trim("      ""contactInformation"": {")
   strDef01 = strDef01 & Trim("         ""phone"": ""+88625061023"",")
   strDef01 = strDef01 & Trim("         ""companyName"": ""TAI E INTERNATIONAL PATENT & TRADEMARK OFFICE"",")
   strDef01 = strDef01 & Trim("         ""fullName"": """ & tmpArrX(2) & """,") '寄件人=員工英文別名
   strDef01 = strDef01 & Trim("         ""email"": ""ipexpress@taie.com.tw""")  '針對DHL寄件服務增加的Email信箱
   strDef01 = strDef01 & Trim("      },")
   strDef01 = strDef01 & Trim("      ""typeCode"": ""business""")  '請帶範例值
   strDef01 = strDef01 & Trim("   },")
   strDef01 = strDef01 & Trim("   ""receiverDetails"": {")    '收件人相關資訊
   strDef01 = strDef01 & Trim("      ""postalAddress"": {")
   strDef01 = strDef01 & Trim("         ""cityName"": """ & tmpArrX(3) & """,")  '城市
   strDef01 = strDef01 & Trim("         ""countryCode"": """ & tmpArrX(4) & """,")  'DHL國家代號
   strDef01 = strDef01 & Trim("         ""postalCode"": """ & tmpArrX(5) & """,")   '寄件郵遞區號
   '地址長度45,地址行最多3，地址1必須有值不可補.
   '判斷address有幾行,參考最初傳入FTP的SPS 2.0.3檔案格式,地址1,2為必填若空白補.;
   If Trim(tmpArrX(7) & tmpArrX(8)) <> "" Then
      strDef01 = strDef01 & Trim("         ""addressLine1"": """ & IIf(Trim(tmpArrX(6)) = "", ".", tmpArrX(6)) & """,")
   Else
      strDef01 = strDef01 & Trim("         ""addressLine1"": """ & IIf(Trim(tmpArrX(6)) = "", ".", tmpArrX(6)) & """")
   End If
   If Trim(tmpArrX(8)) <> "" Then
      strDef01 = strDef01 & Trim("         ""addressLine2"": """ & IIf(Trim(tmpArrX(7)) = "", ".", tmpArrX(7)) & """,")
   ElseIf Trim(tmpArrX(7)) <> "" Then
      strDef01 = strDef01 & Trim("         ""addressLine2"": """ & IIf(Trim(tmpArrX(7)) = "", ".", tmpArrX(7)) & """")
   End If
   If Trim(tmpArrX(8)) <> "" Then  '地址最多3，必須有值不可補.
      strDef01 = strDef01 & Trim("         ""addressLine3"": """ & tmpArrX(8) & """")
   End If
   strDef01 = strDef01 & Trim("      },")
   strDef01 = strDef01 & Trim("      ""contactInformation"": {")
   strDef01 = strDef01 & Trim("         ""phone"": """ & IIf(Trim(tmpArrX(9)) = "", ".", tmpArrX(9)) & """,")  '電話為必填欄位，空白補.
   strDef01 = strDef01 & Trim("         ""companyName"": """ & tmpArrX(10) & """,")
   strDef01 = strDef01 & Trim("         ""fullName"": """ & tmpArrX(11) & """")   '收件人/公司名稱
   strDef01 = strDef01 & Trim("      },")
   strDef01 = strDef01 & Trim("      ""typeCode"": ""business""")  '請帶範例值
   strDef01 = strDef01 & Trim("   }")
   strDef01 = strDef01 & Trim("},")
   strDef01 = strDef01 & Trim(" ""content"": {")
   strDef01 = strDef01 & Trim("     ""unitOfMeasurement"": ""metric"",")  '測量單位
   strDef01 = strDef01 & Trim("     ""isCustomsDeclarable"": false,")  '海關目的 dutiable (true) or non dutiable (false) 文件需要選false
   strDef01 = strDef01 & Trim("     ""incoterm"": ""DAP"",")   '國際商業貿易術語 The reference type please check on SPEC
   strDef01 = strDef01 & Trim("     ""description"": ""Document"",") '貨件描述
   strDef01 = strDef01 & Trim("     ""packages"": [") '包裝信息
   strDef01 = strDef01 & Trim("         {")
   strDef01 = strDef01 & Trim("             ""customerReferences"": [")  '貨件1備註
   strDef01 = strDef01 & Trim("                 {")
   'Modified by Lydia 2025/10/28 備註只有35字元
   'strDef01 = strDef01 & Trim("                     ""value"": """ & tmpArrX(12) & """,")
   strDef01 = strDef01 & Trim("                     ""value"": """ & convForm(tmpArrX(12), 35) & """,")
   strDef01 = strDef01 & Trim("                     ""typeCode"": ""CU""")  'The reference type please check on SPEC
   strDef01 = strDef01 & Trim("                 }")
   strDef01 = strDef01 & Trim("             ],")
   strDef01 = strDef01 & Trim("             ""weight"": 0.5,") '包裹總重量KG
   strDef01 = strDef01 & Trim("             ""dimensions"": {") '材積
   strDef01 = strDef01 & Trim("                 ""length"": 35,") '量信封取得
   strDef01 = strDef01 & Trim("                 ""width"": 28,")
   strDef01 = strDef01 & Trim("                 ""height"": 1")
   strDef01 = strDef01 & Trim("             }")
   strDef01 = strDef01 & Trim("         }")
   strDef01 = strDef01 & Trim("     ]")
   strDef01 = strDef01 & Trim("}")
   
   MakeJSONforDHL = strDef01
   
End Function

'Added by Lydia 2025/10/28
Private Sub Text2_Validate(Cancel As Boolean)
   Cancel = False
   '備註只有35字元
   If CheckLengthIsOK(Text2, 35) = False Then
      Cancel = True
      Text2_GotFocus
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   Cancel = False
   '備註只有35字元
   If CheckLengthIsOK(Text3, 35) = False Then
      Cancel = True
      Text3_GotFocus
   End If
End Sub
'end 2025/10/28

