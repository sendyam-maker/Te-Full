VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_14 
   BackColor       =   &H80000004&
   BorderStyle     =   1  '單線固定
   Caption         =   "潛在客戶資料查詢"
   ClientHeight    =   6550
   ClientLeft      =   1440
   ClientTop       =   2320
   ClientWidth     =   9070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6550
   ScaleWidth      =   9070
   Begin VB.CommandButton CmdOk1 
      Caption         =   "被介紹者"
      Height          =   400
      Index           =   3
      Left            =   6120
      Style           =   1  '圖片外觀
      TabIndex        =   152
      Top             =   90
      Width           =   950
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   0
      Left            =   7120
      TabIndex        =   56
      Top             =   90
      Width           =   1130
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束"
      Height          =   400
      Index           =   1
      Left            =   8330
      TabIndex        =   57
      Top             =   90
      Width           =   700
   End
   Begin TabDlg.SSTab tabCustomer 
      Height          =   5700
      Left            =   90
      TabIndex        =   58
      Top             =   810
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   8
      TabHeight       =   420
      TabCaption(0)   =   "基本"
      TabPicture(0)   =   "frm100101_14.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(8)"
      Tab(0).Control(1)=   "Label1(22)"
      Tab(0).Control(2)=   "Label1(17)"
      Tab(0).Control(3)=   "Label1(21)"
      Tab(0).Control(4)=   "Label1(20)"
      Tab(0).Control(5)=   "Label1(15)"
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(7)=   "Label1(1)"
      Tab(0).Control(8)=   "Label1(2)"
      Tab(0).Control(9)=   "Label1(4)"
      Tab(0).Control(10)=   "Label1(5)"
      Tab(0).Control(11)=   "Label1(6)"
      Tab(0).Control(12)=   "lbl1(1)"
      Tab(0).Control(13)=   "Label1(13)"
      Tab(0).Control(14)=   "Label1(16)"
      Tab(0).Control(15)=   "Label1(19)"
      Tab(0).Control(16)=   "Label1(24)"
      Tab(0).Control(17)=   "Label1(23)"
      Tab(0).Control(18)=   "Label1(18)"
      Tab(0).Control(19)=   "Label1(107)"
      Tab(0).Control(20)=   "Label1(36)"
      Tab(0).Control(21)=   "cboPCU11"
      Tab(0).Control(22)=   "txtPCU47N"
      Tab(0).Control(23)=   "txtPCU(10)"
      Tab(0).Control(24)=   "txtPCU(34)"
      Tab(0).Control(25)=   "txtPCU(35)"
      Tab(0).Control(26)=   "txtPCU(39)"
      Tab(0).Control(27)=   "txtPCU(12)"
      Tab(0).Control(28)=   "txtPCU(3)"
      Tab(0).Control(29)=   "txtPCU(6)"
      Tab(0).Control(30)=   "txtPCU(5)"
      Tab(0).Control(31)=   "txtPCU(4)"
      Tab(0).Control(32)=   "txtPCU(7)"
      Tab(0).Control(33)=   "txtPCU(8)"
      Tab(0).Control(34)=   "txtPCU(9)"
      Tab(0).Control(35)=   "txtPCU(37)"
      Tab(0).Control(36)=   "txtPCU(36)"
      Tab(0).Control(37)=   "txtPCU(40)"
      Tab(0).Control(38)=   "txtPCU(48)"
      Tab(0).Control(39)=   "txtPCU(50)"
      Tab(0).Control(40)=   "txtPCU(47)"
      Tab(0).Control(41)=   "txtCity"
      Tab(0).Control(42)=   "lstUsers(0)"
      Tab(0).Control(43)=   "lstUsers(2)"
      Tab(0).Control(44)=   "Label1(34)"
      Tab(0).Control(45)=   "Label1(35)"
      Tab(0).Control(46)=   "lblXYS02_N"
      Tab(0).Control(47)=   "txtPCU(55)"
      Tab(0).Control(48)=   "Label1(41)"
      Tab(0).Control(49)=   "txtXYS03"
      Tab(0).Control(50)=   "txtXYS02"
      Tab(0).ControlCount=   51
      TabCaption(1)   =   "通訊"
      TabPicture(1)   =   "frm100101_14.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPCU(25)"
      Tab(1).Control(1)=   "txtPCU(27)"
      Tab(1).Control(2)=   "txtPCU(28)"
      Tab(1).Control(3)=   "txtPCU(24)"
      Tab(1).Control(4)=   "txtPCU(23)"
      Tab(1).Control(5)=   "txtPCU(22)"
      Tab(1).Control(6)=   "txtPCU(21)"
      Tab(1).Control(7)=   "txtPCU(20)"
      Tab(1).Control(8)=   "txtPCU(26)"
      Tab(1).Control(9)=   "txtPCU(17)"
      Tab(1).Control(10)=   "txtPCU(18)"
      Tab(1).Control(11)=   "txtPCU(16)"
      Tab(1).Control(12)=   "txtPCU(15)"
      Tab(1).Control(13)=   "txtPCU(14)"
      Tab(1).Control(14)=   "txtPCU(13)"
      Tab(1).Control(15)=   "txtPCU(19)"
      Tab(1).Control(16)=   "txtPCU(33)"
      Tab(1).Control(17)=   "txtPCU(32)"
      Tab(1).Control(18)=   "txtPCU(31)"
      Tab(1).Control(19)=   "txtPCU(30)"
      Tab(1).Control(20)=   "txtPCU(29)"
      Tab(1).Control(21)=   "Label41(20)"
      Tab(1).Control(22)=   "Label41(37)"
      Tab(1).Control(23)=   "Label41(36)"
      Tab(1).Control(24)=   "Label41(35)"
      Tab(1).Control(25)=   "Label41(34)"
      Tab(1).Control(26)=   "Label41(33)"
      Tab(1).Control(27)=   "Label41(32)"
      Tab(1).Control(28)=   "lbl1(2)"
      Tab(1).Control(29)=   "Label41(28)"
      Tab(1).Control(30)=   "Label41(22)"
      Tab(1).Control(31)=   "Label41(21)"
      Tab(1).Control(32)=   "Label63(15)"
      Tab(1).Control(33)=   "Label63(13)"
      Tab(1).Control(34)=   "Label63(12)"
      Tab(1).Control(35)=   "Label63(11)"
      Tab(1).Control(36)=   "Label63(10)"
      Tab(1).Control(37)=   "Label63(9)"
      Tab(1).Control(38)=   "Label63(6)"
      Tab(1).Control(39)=   "Label41(27)"
      Tab(1).Control(40)=   "Label41(26)"
      Tab(1).Control(41)=   "Label41(25)"
      Tab(1).Control(42)=   "Label41(24)"
      Tab(1).Control(43)=   "Label41(23)"
      Tab(1).ControlCount=   44
      TabCaption(2)   =   "聯絡人"
      TabPicture(2)   =   "frm100101_14.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(10)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DataGrid1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Adodc1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraContact"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "關聯企業"
      TabPicture(3)   =   "frm100101_14.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Combo1"
      Tab(3).Control(1)=   "List1"
      Tab(3).Control(2)=   "txtPCU(49)"
      Tab(3).Control(3)=   "Label1(26)"
      Tab(3).Control(4)=   "Label1(55)"
      Tab(3).Control(5)=   "Label1(27)"
      Tab(3).Control(6)=   "Label1(28)"
      Tab(3).Control(7)=   "Label1(29)"
      Tab(3).Control(8)=   "Label1(30)"
      Tab(3).Control(9)=   "Label1(31)"
      Tab(3).Control(10)=   "Label1(32)"
      Tab(3).Control(11)=   "lbl2(0)"
      Tab(3).Control(12)=   "lbl2(1)"
      Tab(3).Control(13)=   "Label1(33)"
      Tab(3).Control(14)=   "lbl2(2)"
      Tab(3).Control(15)=   "lbl2(3)"
      Tab(3).Control(16)=   "lbl2(4)"
      Tab(3).Control(17)=   "lbl2(5)"
      Tab(3).ControlCount=   18
      Begin VB.TextBox txtXYS02 
         Height          =   285
         Left            =   -73668
         MaxLength       =   8
         TabIndex        =   156
         Top             =   4980
         Width           =   1000
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   -70080
         TabIndex        =   121
         Text            =   "Combo1"
         Top             =   1980
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Height          =   760
         Left            =   -74160
         TabIndex        =   120
         Top             =   1980
         Width           =   2175
      End
      Begin VB.Frame fraContact 
         Height          =   3885
         Left            =   135
         TabIndex        =   101
         Top             =   1710
         Width           =   8610
         Begin VB.CommandButton Command1 
            BackColor       =   &H008080FF&
            Caption         =   "上傳相片"
            Height          =   276
            Left            =   1896
            Style           =   1  '圖片外觀
            TabIndex        =   153
            Top             =   192
            Width           =   948
         End
         Begin VB.CommandButton CmdOk1 
            Caption         =   "寄發信函-往來記錄"
            Height          =   400
            Index           =   2
            Left            =   4500
            TabIndex        =   144
            Top             =   1620
            Width           =   1845
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
            ItemData        =   "frm100101_14.frx":0070
            Left            =   1080
            List            =   "frm100101_14.frx":0077
            MultiSelect     =   1  '簡易多重選取
            Sorted          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1155
            Width           =   3180
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
            ItemData        =   "frm100101_14.frx":0084
            Left            =   1080
            List            =   "frm100101_14.frx":008B
            MultiSelect     =   1  '簡易多重選取
            Sorted          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1600
            Width           =   3180
         End
         Begin MSForms.ListBox lstUsers 
            Height          =   585
            Index           =   1
            Left            =   1080
            TabIndex        =   150
            Top             =   2700
            Width           =   1290
            VariousPropertyBits=   746586139
            ScrollBars      =   2
            DisplayStyle    =   2
            Size            =   "2275;1032"
            MatchEntry      =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   26
            Left            =   3480
            TabIndex        =   53
            Top             =   3345
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
            Left            =   5850
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2045
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
            Index           =   24
            Left            =   7335
            TabIndex        =   52
            Top             =   2910
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
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            VariousPropertyBits=   671105055
            Size            =   "1058;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   9
            Left            =   4410
            TabIndex        =   50
            Top             =   2370
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
            Index           =   11
            Left            =   1080
            TabIndex        =   49
            Top             =   2370
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
            Height          =   615
            Index           =   13
            Left            =   5490
            TabIndex        =   54
            Top             =   3210
            Width           =   3060
            VariousPropertyBits=   -1466941413
            MaxLength       =   2000
            ScrollBars      =   2
            Size            =   "5397;1085"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   8
            Left            =   1080
            TabIndex        =   47
            Top             =   2045
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
            Height          =   300
            Index           =   10
            Left            =   6975
            TabIndex        =   51
            Top             =   2370
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
            Index           =   4
            Left            =   5310
            TabIndex        =   43
            Top             =   505
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
            TabIndex        =   42
            Top             =   505
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
            Index           =   5
            Left            =   1080
            TabIndex        =   44
            Top             =   830
            Width           =   3180
            VariousPropertyBits=   671105051
            MaxLength       =   10
            Size            =   "5609;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC20 
            Height          =   300
            Left            =   5850
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   830
            Width           =   2055
            VariousPropertyBits=   671105055
            Size            =   "3625;529"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCUID1 
            Height          =   300
            Left            =   2904
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   180
            Width           =   5604
            VariousPropertyBits=   -2147467233
            BackColor       =   16777215
            Size            =   "9885;529"
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
            Index           =   38
            Left            =   960
            TabIndex        =   143
            Top             =   3555
            Width           =   2430
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否同意歐盟通用資料保護規範(GDPR)："
            Height          =   180
            Index           =   37
            Left            =   120
            TabIndex        =   142
            Top             =   3345
            Width           =   3270
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "名片臨時編號："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   14
            Left            =   4560
            TabIndex        =   139
            Top             =   2105
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否寄專利雙週報：      （N:不寄)"
            Height          =   180
            Index           =   25
            Left            =   5670
            TabIndex        =   119
            Top             =   2970
            Width           =   2655
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "相關聯絡人編號："
            Height          =   180
            Index           =   8
            Left            =   4365
            TabIndex        =   114
            Top             =   890
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "開發日期：                         ( 西元 )"
            Height          =   180
            Index           =   9
            Left            =   135
            TabIndex        =   113
            Top             =   2430
            Width           =   2595
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "E-MAIL："
            Height          =   180
            Index           =   5
            Left            =   135
            TabIndex        =   112
            Top             =   2105
            Width           =   780
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "聯絡人編號："
            Height          =   180
            Index           =   7
            Left            =   135
            TabIndex        =   111
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否寄台一雜誌：     （N:不寄)"
            Height          =   180
            Index           =   7
            Left            =   2970
            TabIndex        =   110
            Top             =   2430
            Width           =   2430
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "備註："
            Height          =   180
            Index           =   14
            Left            =   4950
            TabIndex        =   109
            Top             =   3210
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "開發人員："
            Height          =   180
            Index           =   12
            Left            =   120
            TabIndex        =   108
            Top             =   2760
            Width           =   900
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "部門："
            Height          =   180
            Index           =   4
            Left            =   135
            TabIndex        =   107
            Top             =   1155
            Width           =   540
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "職稱："
            Height          =   180
            Index           =   3
            Left            =   135
            TabIndex        =   106
            Top             =   1600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否寄電子報：      （N:不寄)"
            Height          =   180
            Index           =   11
            Left            =   5676
            TabIndex        =   105
            Top             =   2436
            Width           =   2316
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "名稱( 日 )："
            Height          =   180
            Index           =   2
            Left            =   4365
            TabIndex        =   104
            Top             =   565
            Width           =   930
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "名稱( 英 )："
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   103
            Top             =   565
            Width           =   930
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "名稱( 中 )："
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   102
            Top             =   890
            Width           =   930
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   7290
         Top             =   900
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
         Bindings        =   "frm100101_14.frx":0099
         Height          =   1155
         Left            =   135
         TabIndex        =   138
         Top             =   360
         Width           =   8625
         _ExtentX        =   15222
         _ExtentY        =   2028
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
         ColumnCount     =   12
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
         BeginProperty Column11 
            DataField       =   "PCC25"
            Caption         =   "名片臨時編號"
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
               ColumnWidth     =   1129.89
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
            EndProperty
         EndProperty
      End
      Begin MSForms.TextBox txtXYS03 
         Height          =   552
         Left            =   -69480
         TabIndex        =   160
         Top             =   4680
         Width           =   3200
         VariousPropertyBits=   -1466941413
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "5644;974"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "其他     說明："
         Height          =   492
         Index           =   41
         Left            =   -70080
         TabIndex        =   159
         Top             =   4680
         Width           =   588
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   285
         Index           =   55
         Left            =   -73668
         TabIndex        =   158
         Top             =   4680
         Width           =   2750
         VariousPropertyBits=   671105051
         Size            =   "4851;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblXYS02_N 
         Height          =   285
         Left            =   -72648
         TabIndex        =   157
         Top             =   5000
         Width           =   2556
         VariousPropertyBits=   27
         Caption         =   "blXYS02_N"
         Size            =   "4498;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹者編號："
         Height          =   180
         Index           =   35
         Left            =   -74760
         TabIndex        =   155
         Top             =   4980
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "來所原因："
         Height          =   180
         Index           =   34
         Left            =   -74760
         TabIndex        =   154
         Top             =   4680
         Width           =   900
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   1128
         Index           =   2
         Left            =   -68100
         TabIndex        =   149
         Top             =   3420
         Width           =   1812
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "3201;1984"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   588
         Index           =   0
         Left            =   -70824
         TabIndex        =   148
         Top             =   3840
         Width           =   1296
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2275;1032"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCity 
         Height          =   300
         Left            =   -73668
         TabIndex        =   10
         Top             =   2520
         Width           =   5412
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   47
         Left            =   -71520
         TabIndex        =   9
         Top             =   2208
         Width           =   1308
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "2302;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   50
         Left            =   -71400
         TabIndex        =   18
         Top             =   4400
         Width           =   336
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   48
         Left            =   -69240
         TabIndex        =   117
         Top             =   3492
         Width           =   336
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   384
         Index           =   40
         Left            =   -73668
         TabIndex        =   19
         Top             =   5280
         Width           =   7380
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13017;677"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   36
         Left            =   -73668
         TabIndex        =   15
         Top             =   3492
         Width           =   336
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   37
         Left            =   -73668
         TabIndex        =   16
         Top             =   3816
         Width           =   852
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1508;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   9
         Left            =   -73668
         TabIndex        =   8
         Top             =   2208
         Width           =   852
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1508;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   8
         Left            =   -73668
         TabIndex        =   7
         Top             =   1884
         Width           =   7332
         VariousPropertyBits=   671105051
         MaxLength       =   79
         Size            =   "12938;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   7
         Left            =   -73668
         TabIndex        =   6
         Top             =   1560
         Width           =   7332
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12938;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   4
         Left            =   -73668
         TabIndex        =   3
         Top             =   588
         Width           =   5412
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   5
         Left            =   -73668
         TabIndex        =   4
         Top             =   912
         Width           =   5412
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   6
         Left            =   -73668
         TabIndex        =   5
         Top             =   1236
         Width           =   5412
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   3
         Left            =   -73668
         TabIndex        =   2
         Top             =   276
         Width           =   5412
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   25
         Left            =   -69960
         TabIndex        =   32
         Top             =   2376
         Width           =   3630
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6403;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   465
         Index           =   27
         Left            =   -73845
         TabIndex        =   34
         Top             =   3183
         Width           =   7545
         VariousPropertyBits=   -1466941413
         MaxLength       =   70
         ScrollBars      =   2
         Size            =   "13309;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   28
         Left            =   -73845
         TabIndex        =   35
         Top             =   3669
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1085;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   24
         Left            =   -73710
         TabIndex        =   31
         Top             =   2376
         Width           =   3555
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6271;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   23
         Left            =   -69960
         TabIndex        =   30
         Top             =   2055
         Width           =   3630
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6403;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   22
         Left            =   -73710
         TabIndex        =   29
         Top             =   2055
         Width           =   3555
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6271;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   21
         Left            =   -69960
         TabIndex        =   28
         Top             =   1734
         Width           =   3630
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6403;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   20
         Left            =   -73710
         TabIndex        =   27
         Top             =   1734
         Width           =   3555
         VariousPropertyBits=   -1466941413
         MaxLength       =   30
         ScrollBars      =   2
         Size            =   "6271;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   465
         Index           =   26
         Left            =   -73845
         TabIndex        =   33
         Top             =   2697
         Width           =   7545
         VariousPropertyBits=   -1466941413
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "13309;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   17
         Left            =   -73845
         TabIndex        =   24
         Top             =   1092
         Width           =   2955
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   18
         Left            =   -69300
         TabIndex        =   25
         Top             =   1092
         Width           =   2955
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5212;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   16
         Left            =   -69300
         TabIndex        =   23
         Top             =   771
         Width           =   2955
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   15
         Left            =   -73845
         TabIndex        =   22
         Top             =   771
         Width           =   2955
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   14
         Left            =   -69300
         TabIndex        =   21
         Top             =   450
         Width           =   2955
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   13
         Left            =   -73845
         TabIndex        =   20
         Top             =   450
         Width           =   2955
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   12
         Left            =   -69840
         TabIndex        =   12
         Top             =   2844
         Width           =   852
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1508;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   19
         Left            =   -73845
         TabIndex        =   26
         Top             =   1413
         Width           =   7500
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "13229;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   33
         Left            =   -73845
         TabIndex        =   40
         Top             =   5280
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   32
         Left            =   -73845
         TabIndex        =   39
         Top             =   4953
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   31
         Left            =   -73845
         TabIndex        =   38
         Top             =   4632
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   30
         Left            =   -73845
         TabIndex        =   37
         Top             =   4311
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   29
         Left            =   -73845
         TabIndex        =   36
         Top             =   3990
         Width           =   3735
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6588;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   39
         Left            =   -73668
         TabIndex        =   17
         Top             =   4140
         Width           =   1308
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "2302;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   35
         Left            =   -69552
         TabIndex        =   14
         Top             =   3168
         Width           =   336
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   34
         Left            =   -73164
         TabIndex        =   13
         Top             =   3168
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   10
         Left            =   -66984
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2520
         Width           =   588
         VariousPropertyBits=   671105049
         MaxLength       =   3
         Size            =   "1032;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   285
         Index           =   49
         Left            =   -74760
         TabIndex        =   122
         Top             =   2400
         Visible         =   0   'False
         Width           =   1305
         VariousPropertyBits=   671105051
         MaxLength       =   18
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU47N 
         Height          =   300
         Left            =   -69144
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   2208
         Width           =   2796
         VariousPropertyBits=   16415
         BackColor       =   16777215
         Size            =   "4921;529"
         Caption         =   "LblFM2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   165
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboPCU11 
         Height          =   300
         Left            =   -73668
         TabIndex        =   11
         Top             =   2844
         Width           =   1716
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3016;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否同意歐盟通用資料保護規範(GDPR)：        （W:待回覆 Y/N:同意/不同意）"
         Height          =   180
         Index           =   36
         Left            =   -74772
         TabIndex        =   141
         Top             =   4450
         Width           =   6060
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "與                              為關係企業"
         Height          =   180
         Index           =   107
         Left            =   -71760
         TabIndex        =   140
         Top             =   2268
         Width           =   2520
      End
      Begin VB.Label Label1 
         Caption         =   "關聯代號："
         Height          =   180
         Index           =   26
         Left            =   -71040
         TabIndex        =   137
         Top             =   2040
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "關聯企業："
         Height          =   180
         Index           =   55
         Left            =   -74760
         TabIndex        =   136
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "關聯："
         Height          =   180
         Index           =   27
         Left            =   -74760
         TabIndex        =   135
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名　　稱："
         Height          =   180
         Index           =   28
         Left            =   -71760
         TabIndex        =   134
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名 稱(中)："
         Height          =   180
         Index           =   29
         Left            =   -74745
         TabIndex        =   133
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(英)："
         Height          =   180
         Index           =   30
         Left            =   -74340
         TabIndex        =   132
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國　　籍："
         Height          =   180
         Index           =   31
         Left            =   -74760
         TabIndex        =   131
         Top             =   900
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(日)："
         Height          =   180
         Index           =   32
         Left            =   -74340
         TabIndex        =   130
         Top             =   1620
         Width           =   480
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   0
         Left            =   -73800
         TabIndex        =   129
         Top             =   900
         Width           =   375
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "661;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   1
         Left            =   -73320
         TabIndex        =   128
         Top             =   900
         Width           =   1455
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2566;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "狀　　態："
         Height          =   180
         Index           =   33
         Left            =   -71760
         TabIndex        =   127
         Top             =   900
         Width           =   900
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   2
         Left            =   -70800
         TabIndex        =   126
         Top             =   900
         Width           =   1575
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2778;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   3
         Left            =   -73800
         TabIndex        =   125
         Top             =   1140
         Width           =   6735
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11880;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   4
         Left            =   -73800
         TabIndex        =   124
         Top             =   1380
         Width           =   6735
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11880;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   5
         Left            =   -73800
         TabIndex        =   123
         Top             =   1620
         Width           =   6735
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11880;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄專利雙週報：       （N:不寄）"
         Height          =   180
         Index           =   18
         Left            =   -70860
         TabIndex        =   118
         Top             =   3552
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最近日期"
         Height          =   180
         Index           =   23
         Left            =   -67128
         TabIndex        =   116
         Top             =   3228
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽同仁"
         Height          =   180
         Index           =   24
         Left            =   -68064
         TabIndex        =   115
         Top             =   3228
         Width           =   720
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
         Left            =   135
         TabIndex        =   60
         Top             =   1530
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   19
         Left            =   -74772
         TabIndex        =   99
         Top             =   5280
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文：             （1:中文   2:英文   3:日文）"
         Height          =   180
         Index           =   16
         Left            =   -74772
         TabIndex        =   98
         Top             =   3552
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發日期：                         ( 西元 )"
         Height          =   180
         Index           =   13
         Left            =   -74772
         TabIndex        =   97
         Top             =   3876
         Width           =   2592
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "lbl1"
         Height          =   180
         Index           =   1
         Left            =   -72780
         TabIndex        =   96
         Top             =   2268
         Width           =   996
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "城市："
         Height          =   180
         Index           =   6
         Left            =   -74772
         TabIndex        =   95
         Top             =   2580
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國籍："
         Height          =   180
         Index           =   5
         Left            =   -74772
         TabIndex        =   94
         Top             =   2268
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名稱（中）："
         Height          =   180
         Index           =   4
         Left            =   -74772
         TabIndex        =   93
         Top             =   1944
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名稱（英）："
         Height          =   180
         Index           =   2
         Left            =   -74772
         TabIndex        =   92
         Top             =   336
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名稱（日）："
         Height          =   180
         Index           =   1
         Left            =   -74772
         TabIndex        =   91
         Top             =   1620
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發人員："
         Height          =   180
         Index           =   3
         Left            =   -71712
         TabIndex        =   90
         Top             =   3876
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "日文地址："
         Height          =   180
         Index           =   20
         Left            =   -74790
         TabIndex        =   89
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   37
         Left            =   -70095
         TabIndex        =   88
         Top             =   2436
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   36
         Left            =   -70095
         TabIndex        =   87
         Top             =   2115
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   35
         Left            =   -70095
         TabIndex        =   86
         Top             =   1794
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   34
         Left            =   -73845
         TabIndex        =   85
         Top             =   2436
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   33
         Left            =   -73845
         TabIndex        =   84
         Top             =   2115
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   32
         Left            =   -73845
         TabIndex        =   83
         Top             =   1794
         Width           =   90
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "lbl1"
         Height          =   180
         Index           =   2
         Left            =   -73200
         TabIndex        =   82
         Top             =   3750
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "中文地址："
         Height          =   180
         Index           =   28
         Left            =   -74775
         TabIndex        =   81
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "地址國籍："
         Height          =   180
         Index           =   22
         Left            =   -74775
         TabIndex        =   80
         Top             =   3729
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "英文地址："
         Height          =   180
         Index           =   21
         Left            =   -74775
         TabIndex        =   79
         Top             =   1794
         Width           =   900
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "行動電話："
         Height          =   180
         Index           =   15
         Left            =   -74775
         TabIndex        =   78
         Top             =   1152
         Width           =   900
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "傳真2："
         Height          =   180
         Index           =   13
         Left            =   -70140
         TabIndex        =   77
         Top             =   831
         Width           =   630
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "電話2："
         Height          =   180
         Index           =   12
         Left            =   -70140
         TabIndex        =   76
         Top             =   510
         Width           =   630
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "E-MAIL："
         Height          =   180
         Index           =   11
         Left            =   -70140
         TabIndex        =   75
         Top             =   1152
         Width           =   780
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "傳真1："
         Height          =   180
         Index           =   10
         Left            =   -74775
         TabIndex        =   74
         Top             =   831
         Width           =   630
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "電話1："
         Height          =   180
         Index           =   9
         Left            =   -74775
         TabIndex        =   73
         Top             =   510
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "類別： "
         Height          =   180
         Index           =   15
         Left            =   -74772
         TabIndex        =   72
         Top             =   2904
         Width           =   588
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "成立日期：                         ( 西元 )"
         Height          =   180
         Index           =   20
         Left            =   -70824
         TabIndex        =   71
         Top             =   2904
         Width           =   2592
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "網址："
         Height          =   180
         Index           =   6
         Left            =   -74775
         TabIndex        =   70
         Top             =   1473
         Width           =   540
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB5："
         Height          =   180
         Index           =   27
         Left            =   -74775
         TabIndex        =   69
         Top             =   5340
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB4："
         Height          =   180
         Index           =   26
         Left            =   -74775
         TabIndex        =   68
         Top             =   5017
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB3："
         Height          =   180
         Index           =   25
         Left            =   -74775
         TabIndex        =   67
         Top             =   4695
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB2："
         Height          =   180
         Index           =   24
         Left            =   -74775
         TabIndex        =   66
         Top             =   4373
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB1："
         Height          =   180
         Index           =   23
         Left            =   -74775
         TabIndex        =   65
         Top             =   4051
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "狀態："
         Height          =   180
         Index           =   21
         Left            =   -74772
         TabIndex        =   64
         Top             =   4200
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報：        （N:不寄）"
         Height          =   180
         Index           =   17
         Left            =   -70860
         TabIndex        =   63
         Top             =   3228
         Width           =   2508
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄台一雜誌：           （N:不寄）"
         Height          =   180
         Index           =   22
         Left            =   -74772
         TabIndex        =   62
         Top             =   3228
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "城市代碼："
         Height          =   180
         Index           =   8
         Left            =   -67980
         TabIndex        =   61
         Top             =   2580
         Width           =   900
      End
   End
   Begin VB.Label SpecCU 
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   2280
      TabIndex        =   151
      Top             =   60
      Width           =   3465
   End
   Begin MSForms.TextBox txtPCU 
      Height          =   300
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   90
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1931;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPCU 
      Height          =   300
      Index           =   2
      Left            =   1920
      TabIndex        =   1
      Top             =   90
      Width           =   255
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   225
      Left            =   150
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   510
      Width           =   8745
      VariousPropertyBits=   -2139078625
      BackColor       =   16777215
      Size            =   "15425;397"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "編號："
      Height          =   210
      Index           =   0
      Left            =   150
      TabIndex        =   100
      Top             =   150
      Width           =   590
   End
End
Attribute VB_Name = "frm100101_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/07 改成Form2.0 ; textCUID、textCUID1、txtPCU(index)、txtPCC(index)、lstUsers(index)、txtPCU47N、lbl2(index)、DataGrid1改字型=新細明體-ExtB、txtCity
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
'Create by Morgan 2007/12/13
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
   textCUID.BackColor = &H8000000F
   textCUID1.BackColor = &H8000000F
   
   'Add by Sindy 2021/6/28
   Call PUB_SetComboPCU11(cboPCU11, "") '設定國外潛在客戶類別選項
   
   Pub_SetFTypeList Me.Combo1, 10 'Added by Lydia 2016/11/29
   
   tabCustomer.Tab = 0
   
   'Added by Lydia 2020/08/04 關聯企業：改用啟用日控制
   'Modified by Lydia 2021/01/06 改回原名
'   If strSrvDate(1) >= 國外部關聯企業啟用日 Then
'        Label1(107).Visible = False
'        tabCustomer.TabVisible(3) = True
'   Else
'   'end 2020/08/04
'        tabCustomer.TabVisible(3) = False  'Added by Lydia 2018/05/24 隱藏關聯企業頁籤
'   End If 'Added by Lydia 2020/08/04
   tabCustomer.TabVisible(3) = False
   'end 2021/01/06
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_14 = Nothing
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
      'Add by Amy 2023/07/12
      Case 3 '被介紹者
         If CmdOk1(3).BackColor <> &HFFFF80 Then
            MsgBox "無被介紹者資料"
            Exit Sub
         End If
         If PUB_CheckFormExist("frm050705_1") Then
              MsgBox "請先關閉〔被介紹資料〕的畫面！", vbInformation
              Exit Sub
         End If
         cmdState = -1
         Me.Enabled = False
         If fnSaveParentForm(Me) = False Then
            Me.Enabled = True
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         Call ShowFrm050705_1
         Me.Enabled = True
         Screen.MousePointer = vbDefault
         Exit Sub
   End Select
End Sub

'Modify By Sindy 2025/8/28
'sType As String: 0=主檔 1=聯絡人
Sub StrMenu(Optional ByVal sType As String = "0")
   Dim strKey  As String, strKey1 As String
   Dim strTp(2) As String 'Add by Amy 2025/02/14
   
   If Mid(Me.Tag, 10, 1) = "-" Then
      strKey = Left(Me.Tag, 9)
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
   pub_QL05 = ";潛在客戶編號：" & strKey & IIf(sType = "0", "(國外潛在客戶基本資料)", "(聯絡人)") 'Add By Sindy 2025/8/13
   
   'Added by Lydia 2023/01/18 往來紀錄中有「A14客戶名稱資訊不得宣傳」者，在申請人/代理人資料查詢首頁提示
   strExc(0) = "SELECT ac03 as memo FROM allcode where AC01='11' and ac02='A14' and exists (select * from contactrecord where instr(cr05,'A14')>0 and substr(cr03,1,8)='" & Mid(strKey, 1, 8) & "' and substr(cr03,9,1)='" & Mid(strKey, 9, 1) & "') "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       SpecCU.Caption = SpecCU.Caption & IIf(Trim(SpecCU.Caption) <> "", "；", "") & RsTemp.Fields("memo")
       SpecCU.Font.Size = 14
       SpecCU.AutoSize = True
   End If
   'end 2023/01/18
   
   'Modify by Amy 2025/02/14 +AllCode 顯示來所原因
   'Modify by Amy 2025/02/21 +ac01(+) ex:代理人查詢->R1896600 ->點代理人鈕 會顯示無資料
   strExc(0) = "select * from potcustomer,allCode where pcu01='" & Left(strKey, 8) & "' and pcu02='" & Mid(strKey, 9) & "' " & _
                        "And ac01(+)='13' And pcu55=ac02(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If pub_QL04 <> "" And sType = "0" Then InsertQueryLog (RsTemp.RecordCount)  'Add By Sindy 2025/8/13
      ShowRecord RsTemp, sType
      'Add by Amy 2025/02/14 國外潛在客戶需顯示 來所資料
      txtXYS02 = "": lblXYS02_N = "": txtXYS03 = ""
      If Pub_GetXYSource(1, Left(strKey, 8), strTp(0), strTp(1), strTp(2)) = True Then
         txtXYS02 = strTp(0)
         lblXYS02_N = strTp(1)
         txtXYS03 = strTp(2)
      End If
      If strKey1 <> "" Then
         tabCustomer.Tab = 2
         ReadContact strKey1
      End If
   Else
      If pub_QL04 <> "" And sType = "0" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
      ShowNoData
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
   'Add By Sindy 2025/8/28 按聯絡人鍵時,切到此Tab
   If sType = "1" Then
      tabCustomer.Tab = 2
   End If
   '2025/8/28 END
   
   'Add by Amy 2023/07/12 被介紹者
   CmdOk1(3).BackColor = &H8000000F
   If Pub_GetXYSource(2, Left(strKey, 8)) = True Then
      CmdOk1(3).BackColor = &HFFFF80
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
'Modify By Sindy 2025/8/28 +, sType As String
Private Sub ShowRecord(ByRef p_Rst As ADODB.Recordset, sType As String)
   Dim rsPCU As ADODB.Recordset
   Dim CUID(1 To 6) As String
   
   ClearField
   SetCtrlReadOnly True
   Set rsPCU = p_Rst.Clone
   With rsPCU
      If .RecordCount > 0 Then
         For Each oText In txtPCU
            idx = oText.Index
            oText.Text = "" & .Fields("PCU" & Format(idx, "0#"))
            'Add by Amy 2025/02/14 +來所原因名稱
            If idx = 55 Then
               oText.Text = oText.Text & " " & .Fields("AC03")
            End If
         Next
         
         'Add by Sindy 2021/6/28
         Call PUB_SetComboPCU11(cboPCU11, "" & .Fields("PCU11")) '類別
         '2021/6/28 END
         
         CUID(1) = "" & .Fields("PCU41")
         CUID(2) = "" & .Fields("PCU42")
         CUID(3) = "" & .Fields("PCU43")
         CUID(4) = "" & .Fields("PCU44")
         CUID(5) = "" & .Fields("PCU45")
         CUID(6) = "" & .Fields("PCU46")
         
         Call Pub_ShowSelectList(Combo1, List1, txtPCU(49).Text) 'Added by Lydia 2016/11/29
         
         '國籍
         If ClsPDGetNation(Left(txtPCU(9), 3), strTmp) = True Then
            lbl1(1).Caption = strTmp
         End If
         '城市
         If txtPCU(9) <> "" And txtPCU(10) <> "" Then
            strExc(0) = "select ct03 from city where ct01='" & Left(txtPCU(9), 3) & "' and ct02='" & txtPCU(10) & "' order by ct03 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               txtCity = "" & RsTemp.Fields(0)
            End If
         End If
         '開發人員
         If Not IsNull(.Fields("pcu38")) Then
            'Modify by Amy 2020/03/19 '改與其他開發人員一樣的function 照原順序排
'            strExc(0) = "select st02 from staff where instr('" & .Fields("pcu38") & "',st01)>0"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strTmp = RsTemp.GetString(, , , ",")
'
'               SetList lstUsers(0), strTmp
'               PUB_SetUserList lstUsers(0), .Fields("pcu38")
'            End If
            'Modified by Lydia 2022/01/07 改成Form 2.0 + True
            PUB_SetUserList lstUsers(0), .Fields("pcu38"), True
            'end 2020/03/19
         End If
         'Add By Sindy 2009/06/24
         '關係企業
         If Not IsNull(.Fields("PCU47")) And Trim(.Fields("PCU47")) <> "" Then
            Call GetCustData(.Fields("PCU47"))
         Else
            txtPCU47N.Text = ""
         End If
         
         'Add by Morgan 2009/2/11
         lstUsers(2).Clear
         strExc(0) = "SELECT MAX(CR02) CR02,RPAD(ST02,9,' '),ST01 FROM CONTACTRECORD,STAFF" & _
            " WHERE CR03='" & txtPCU(1) & txtPCU(2) & "' AND INSTR(CR19,ST01(+))>0 GROUP BY ST01,ST02 ORDER BY CR02 ASC,ST01 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
               lstUsers(2).AddItem RsTemp.Fields(1) & " " & RsTemp.Fields(0), 0
               RsTemp.MoveNext
            Loop
         End If
         
         '地址國籍
         If ClsPDGetNation(Left(txtPCU(28), 3), strTmp) = True Then
            lbl1(2).Caption = strTmp
         End If
         Call OpenContactTable(sType) 'Modify By Sindy 2025/8/28 +sType
      End If
   End With
   UpdateCUID CUID, textCUID
End Sub

Private Sub ClearField()
   Dim oLabel As Object
   For Each oText In txtPCU
      oText.Text = Empty
   Next
   For Each oLabel In lbl1
      oLabel.Caption = Empty
   Next
   'Add By Sindy 2009/06/24
   txtPCU47N = ""
   '2009/06/24 End
   txtCity.Text = ""
   textCUID = ""
   lstUsers(0).Clear
   
   'Added by Lydia 2016/11/29
   Combo1.ListIndex = -1
   List1.Clear
   For Each oLabel In lbl2
      oLabel.Caption = Empty
   Next
   'end 2016/11/29
   
   cboPCU11.ListIndex = -1 'Add By Sindy 2021/6/28
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

'Modify By Sindy 2025/8/28 +sType As String
Private Sub OpenContactTable(sType As String)
   
On Error GoTo Checking
   
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
      If pub_QL04 <> "" And sType = "1" Then InsertQueryLog (rsContact.RecordCount) 'Add By Sindy 2025/8/28
      ReadContact
   Else
      If pub_QL04 <> "" And sType = "1" Then InsertQueryLog (0) 'Add By Sindy 2025/8/28
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
            If .EOF Then Exit Sub
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
         If Not IsNull(.Fields("PCC12")) Then
            SetlstUsers 1, .Fields("PCC12")
         End If
         UpdateCUID CUID, textCUID1
      End If
      'Added by Lydia 2024/05/10 聯絡人相片
      If Trim(txtPCU(1)) <> "" And Trim(txtPCC(2)) <> "" Then
         Command1.Visible = True
         Call Pub_GetPCCtoIBF_2(Trim(txtPCU(1)), Trim(txtPCC(2)), Command1)
      Else
         Command1.Visible = False
      End If
      'end 2024/05/10
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
   'Modified by Lydia 2021/01/10 判斷textCUID1不使用換行
   'Modified by Lydia 2023/01/19 改成一行顯示，拿掉vbCrLf=> ""
   'Modified by Lydia 2024/05/10 String(10 -> 6
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(6, " ") & IIf(oText.Name = "textCUID1", "      ", "") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime

End Sub

Private Sub SetList(oList As ListBox, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID

   lstUsers(p_idx).Clear
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         arrID = Split(p_stNums, ",")
         With RsTemp
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("st01") = arrID(intI) Then
                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtPCU
      oText.Locked = bLocked
   Next
   For Each oText In txtPCC
      oText.Locked = bLocked
   Next
   txtCity.Locked = bLocked
   'Add By Sindy 2009/06/24
   txtPCU47N.Enabled = Not bLocked
   
   cboPCU11.Locked = bLocked 'Add By Sindy 2021/6/28
End Sub

'Add By Sindy 2009/06/23
Private Function GetCustData(p_stCust As String) As Boolean
   Dim aiOrder(1 To 3) As Integer
   Select Case Left(p_stCust, 1)
      Case "X"
         'Modified by Lydia 2016/11/29
         'strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 Na01,NA03,cu80 s1 from customer,nation where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "' and cu10=na01(+)"
      Case "Y"
         'Modified by Lydia 2016/11/29
         'strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 N3 from fagent where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "'"
         strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 Na01,na03,fa69 s1 from fagent,nation where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "' and fa10=na01(+)"
      Case Else
        'Modified by Lydia 2016/11/29 關係企業=>關聯企業
        'Modified by Lydia 2021/01/06 關聯企業=>關係企業
         MsgBox "關係企業必須為 X 或 Y 開頭", vbCritical + vbOKOnly, "檢核資料"
         Exit Function
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   txtPCU47N.Text = ""
   If intI = 1 Then
      For intI = 1 To 3
         'Modified by Lydia 2016/11/29 關聯企業名稱
         'If Not IsNull(RsTemp(intI)) Then
         '   txtPCU47N.Text = RsTemp(intI)
         '   Exit For
         'End If
         If txtPCU47N.Text = "" And "" & RsTemp(intI) <> "" Then txtPCU47N.Text = RsTemp(intI)
         lbl2(2 + intI) = "" & RsTemp(intI)
         'end 2016/11/19
      Next
      'Added by Lydia 2016/11/29 關聯企業名稱,國籍和狀態
      lbl2(0) = "" & RsTemp.Fields("na01")
      lbl2(1) = Trim(Mid("" & RsTemp.Fields("na03"), 1, 4))
      lbl2(2) = "" & RsTemp.Fields("s1")
      'end 2016/11/29
      GetCustData = True
   Else
      'Modified by Lydia 2016/11/29 關係企業=>關聯企業
      'Modified by Lydia 2021/01/06 關聯企業=>關係企業
      MsgBox "關係企業輸入錯誤！"
   End If
End Function

'Add by Amy 2023/07/12 顯示被介紹者資料
Private Sub ShowFrm050705_1()
   Dim stName As String
   
   If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
    '英->中->日
    If Trim(txtPCU(3)) = MsgText(601) Then
        If Trim(txtPCU(8)) = MsgText(601) Then
            stName = txtPCU(7) '日
        Else
            stName = txtPCU(8) '中
        End If
    Else
        stName = txtPCU(3)
        If Trim(txtPCU(4)) <> MsgText(601) Then
            stName = stName & " " & txtPCU(3)
        End If
        If Trim(txtPCU(5)) <> MsgText(601) Then
            stName = stName & " " & txtPCU(5)
         End If
         If Trim(txtPCU(6)) <> MsgText(601) Then
            stName = stName & " " & txtPCU(6)
         End If
    End If
    frm050705_1.txtNo = Left(txtPCU(1), 8)
    frm050705_1.lbl1(0) = txtPCU(9) '國籍code
    frm050705_1.lbl1(1) = lbl1(1) '國籍
    frm050705_1.lbl1(3) = stName
    frm050705_1.SetParent Me
    frm050705_1.QueryData
    frm050705_1.Show
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

