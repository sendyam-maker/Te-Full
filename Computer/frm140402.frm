VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140402 
   BorderStyle     =   1  '單線固定
   Caption         =   "潛在客戶資料維護"
   ClientHeight    =   6672
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
   ScaleHeight     =   6672
   ScaleWidth      =   9156
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8505
      Top             =   990
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
            Picture         =   "frm140402.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140402.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   79
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
   Begin TabDlg.SSTab tabCustomer 
      Height          =   5650
      Left            =   120
      TabIndex        =   83
      Top             =   960
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   9970
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   8
      TabHeight       =   420
      OLEDropMode     =   1
      TabCaption(0)   =   "基本"
      TabPicture(0)   =   "frm140402.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtXYS02"
      Tab(0).Control(1)=   "cboSource"
      Tab(0).Control(2)=   "cmdIntroduce"
      Tab(0).Control(3)=   "Option1(1)"
      Tab(0).Control(4)=   "Option1(0)"
      Tab(0).Control(5)=   "txtSameCnt"
      Tab(0).Control(6)=   "cmdTransfer"
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(8)=   "cboCity"
      Tab(0).Control(9)=   "Label1(41)"
      Tab(0).Control(10)=   "txtXYS03"
      Tab(0).Control(11)=   "Label1(35)"
      Tab(0).Control(12)=   "LblSourceN"
      Tab(0).Control(13)=   "Label1(34)"
      Tab(0).Control(14)=   "lstUsers(2)"
      Tab(0).Control(15)=   "lstUsers(0)"
      Tab(0).Control(16)=   "txtPCU47N"
      Tab(0).Control(17)=   "txtPCU(47)"
      Tab(0).Control(18)=   "txtPCU(9)"
      Tab(0).Control(19)=   "txtPCU(8)"
      Tab(0).Control(20)=   "txtPCU(7)"
      Tab(0).Control(21)=   "txtPCU(6)"
      Tab(0).Control(22)=   "txtPCU(5)"
      Tab(0).Control(23)=   "txtPCU(4)"
      Tab(0).Control(24)=   "txtPCU(3)"
      Tab(0).Control(25)=   "txtPCU(51)"
      Tab(0).Control(26)=   "txtPCU(50)"
      Tab(0).Control(27)=   "txtPCU(48)"
      Tab(0).Control(28)=   "txtPCU(38)"
      Tab(0).Control(29)=   "txtPCU(10)"
      Tab(0).Control(30)=   "txtPCU(34)"
      Tab(0).Control(31)=   "txtPCU(35)"
      Tab(0).Control(32)=   "txtPCU(39)"
      Tab(0).Control(33)=   "txtPCU(12)"
      Tab(0).Control(34)=   "txtPCU(37)"
      Tab(0).Control(35)=   "txtPCU(36)"
      Tab(0).Control(36)=   "txtPCU(40)"
      Tab(0).Control(37)=   "cboPCU11"
      Tab(0).Control(38)=   "lblTitle"
      Tab(0).Control(39)=   "Label1(39)"
      Tab(0).Control(40)=   "Label1(36)"
      Tab(0).Control(41)=   "Label1(107)"
      Tab(0).Control(42)=   "Label1(18)"
      Tab(0).Control(43)=   "Label1(24)"
      Tab(0).Control(44)=   "Label1(23)"
      Tab(0).Control(45)=   "Label1(8)"
      Tab(0).Control(46)=   "Label1(22)"
      Tab(0).Control(47)=   "Label1(17)"
      Tab(0).Control(48)=   "Label1(21)"
      Tab(0).Control(49)=   "Label1(20)"
      Tab(0).Control(50)=   "Label1(15)"
      Tab(0).Control(51)=   "Label1(3)"
      Tab(0).Control(52)=   "Label1(1)"
      Tab(0).Control(53)=   "Label1(2)"
      Tab(0).Control(54)=   "Label1(4)"
      Tab(0).Control(55)=   "Label1(5)"
      Tab(0).Control(56)=   "Label1(6)"
      Tab(0).Control(57)=   "lbl1(1)"
      Tab(0).Control(58)=   "Label1(13)"
      Tab(0).Control(59)=   "Label1(16)"
      Tab(0).Control(60)=   "Label1(19)"
      Tab(0).ControlCount=   61
      TabCaption(1)   =   "通訊"
      TabPicture(1)   =   "frm140402.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label41(20)"
      Tab(1).Control(1)=   "Label41(37)"
      Tab(1).Control(2)=   "Label41(36)"
      Tab(1).Control(3)=   "Label41(35)"
      Tab(1).Control(4)=   "Label41(34)"
      Tab(1).Control(5)=   "Label41(33)"
      Tab(1).Control(6)=   "Label41(32)"
      Tab(1).Control(7)=   "lbl1(2)"
      Tab(1).Control(8)=   "Label41(28)"
      Tab(1).Control(9)=   "Label41(22)"
      Tab(1).Control(10)=   "Label41(21)"
      Tab(1).Control(11)=   "Label63(15)"
      Tab(1).Control(12)=   "Label63(13)"
      Tab(1).Control(13)=   "Label63(12)"
      Tab(1).Control(14)=   "Label63(11)"
      Tab(1).Control(15)=   "Label63(10)"
      Tab(1).Control(16)=   "Label63(9)"
      Tab(1).Control(17)=   "Label63(6)"
      Tab(1).Control(18)=   "Label41(27)"
      Tab(1).Control(19)=   "Label41(26)"
      Tab(1).Control(20)=   "Label41(25)"
      Tab(1).Control(21)=   "Label41(24)"
      Tab(1).Control(22)=   "Label41(23)"
      Tab(1).Control(23)=   "txtPCU(25)"
      Tab(1).Control(24)=   "txtPCU(27)"
      Tab(1).Control(25)=   "txtPCU(28)"
      Tab(1).Control(26)=   "txtPCU(24)"
      Tab(1).Control(27)=   "txtPCU(23)"
      Tab(1).Control(28)=   "txtPCU(22)"
      Tab(1).Control(29)=   "txtPCU(21)"
      Tab(1).Control(30)=   "txtPCU(20)"
      Tab(1).Control(31)=   "txtPCU(26)"
      Tab(1).Control(32)=   "txtPCU(17)"
      Tab(1).Control(33)=   "txtPCU(18)"
      Tab(1).Control(34)=   "txtPCU(16)"
      Tab(1).Control(35)=   "txtPCU(15)"
      Tab(1).Control(36)=   "txtPCU(14)"
      Tab(1).Control(37)=   "txtPCU(13)"
      Tab(1).Control(38)=   "txtPCU(19)"
      Tab(1).Control(39)=   "txtPCU(33)"
      Tab(1).Control(40)=   "txtPCU(32)"
      Tab(1).Control(41)=   "txtPCU(31)"
      Tab(1).Control(42)=   "txtPCU(30)"
      Tab(1).Control(43)=   "txtPCU(29)"
      Tab(1).ControlCount=   44
      TabCaption(2)   =   "聯絡人"
      TabPicture(2)   =   "frm140402.frx":212C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(10)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DataGrid1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraContact"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Adodc1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdContact(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdContact(3)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdContact(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "關聯企業"
      TabPicture(3)   =   "frm140402.frx":2148
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdRemPCU49"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdAddPCU49"
      Tab(3).Control(2)=   "List1"
      Tab(3).Control(3)=   "Combo1"
      Tab(3).Control(4)=   "txtPCU(49)"
      Tab(3).Control(5)=   "Label1(40)"
      Tab(3).Control(6)=   "lblPCU47N"
      Tab(3).Control(7)=   "lblPCU47"
      Tab(3).Control(8)=   "lbl2(5)"
      Tab(3).Control(9)=   "lbl2(4)"
      Tab(3).Control(10)=   "lbl2(3)"
      Tab(3).Control(11)=   "lbl2(2)"
      Tab(3).Control(12)=   "Label1(33)"
      Tab(3).Control(13)=   "lbl2(1)"
      Tab(3).Control(14)=   "lbl2(0)"
      Tab(3).Control(15)=   "Label1(32)"
      Tab(3).Control(16)=   "Label1(31)"
      Tab(3).Control(17)=   "Label1(30)"
      Tab(3).Control(18)=   "Label1(29)"
      Tab(3).Control(19)=   "Label1(28)"
      Tab(3).Control(20)=   "Label1(27)"
      Tab(3).Control(21)=   "Label1(26)"
      Tab(3).Control(22)=   "Label1(55)"
      Tab(3).ControlCount=   23
      Begin VB.TextBox txtXYS02 
         Height          =   264
         Left            =   -73668
         MaxLength       =   8
         TabIndex        =   27
         Top             =   4620
         Width           =   1000
      End
      Begin VB.ComboBox cboSource 
         Height          =   276
         ItemData        =   "frm140402.frx":2164
         Left            =   -73860
         List            =   "frm140402.frx":2166
         Style           =   2  '單純下拉式
         TabIndex        =   26
         Top             =   4344
         Width           =   2750
      End
      Begin VB.CommandButton cmdIntroduce 
         Caption         =   "被介紹者"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71064
         Style           =   1  '圖片外觀
         TabIndex        =   190
         Top             =   4320
         Width           =   1000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "國外"
         Height          =   195
         Index           =   1
         Left            =   -72960
         TabIndex        =   183
         Top             =   5376
         Width           =   705
      End
      Begin VB.OptionButton Option1 
         Caption         =   "國內"
         Height          =   195
         Index           =   0
         Left            =   -73680
         TabIndex        =   182
         Top             =   5376
         Width           =   705
      End
      Begin VB.CommandButton cmdRemPCU49 
         Caption         =   "→"
         Height          =   285
         Left            =   -71895
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   2340
         Width           =   600
      End
      Begin VB.CommandButton cmdAddPCU49 
         Caption         =   "←"
         Height          =   285
         Left            =   -71895
         TabIndex        =   187
         Top             =   2040
         Width           =   600
      End
      Begin VB.ListBox List1 
         Height          =   228
         Left            =   -74160
         TabIndex        =   80
         Top             =   1980
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   -70080
         TabIndex        =   186
         Text            =   "Combo1"
         Top             =   1980
         Width           =   2895
      End
      Begin VB.TextBox txtSameCnt 
         Height          =   270
         Left            =   -67890
         MaxLength       =   6
         TabIndex        =   154
         Top             =   1056
         Width           =   1215
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "轉客戶或代理人"
         Height          =   405
         Left            =   -68040
         TabIndex        =   149
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton cmdContact 
         Caption         =   "加入"
         Height          =   285
         Index           =   2
         Left            =   7200
         TabIndex        =   76
         Top             =   1380
         Width           =   735
      End
      Begin VB.CommandButton cmdContact 
         Caption         =   "刪除"
         Height          =   285
         Index           =   3
         Left            =   7965
         TabIndex        =   77
         Top             =   1380
         Width           =   735
      End
      Begin VB.CommandButton cmdContact 
         Caption         =   "新增"
         Height          =   285
         Index           =   1
         Left            =   6435
         TabIndex        =   51
         Top             =   1380
         Width           =   735
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   7290
         Top             =   900
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
         Height          =   3945
         Left            =   135
         TabIndex        =   113
         Top             =   1600
         Width           =   8610
         Begin VB.CommandButton Command1 
            BackColor       =   &H008080FF&
            Caption         =   "上傳相片"
            Height          =   276
            Left            =   1896
            Style           =   1  '圖片外觀
            TabIndex        =   189
            Top             =   168
            Width           =   948
         End
         Begin VB.TextBox txtPCC20 
            BackColor       =   &H8000000F&
            Height          =   270
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   780
            Width           =   2055
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
            Height          =   540
            IntegralHeight  =   0   'False
            ItemData        =   "frm140402.frx":2168
            Left            =   1080
            List            =   "frm140402.frx":216F
            MultiSelect     =   1  '簡易多重選取
            Sorted          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1080
            Width           =   3180
         End
         Begin VB.Frame Frame4 
            Height          =   675
            Left            =   4275
            TabIndex        =   147
            Top             =   990
            Width           =   4155
            Begin VB.ComboBox cboDept 
               Height          =   300
               ItemData        =   "frm140402.frx":217C
               Left            =   810
               List            =   "frm140402.frx":217E
               TabIndex        =   56
               Text            =   "cboDept"
               Top             =   120
               Width           =   3285
            End
            Begin VB.CommandButton cmdAddDept 
               Caption         =   "<- 新增"
               Height          =   255
               Left            =   45
               TabIndex        =   57
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton cmdRemoveDept 
               Caption         =   "移除 ->"
               Height          =   255
               Left            =   45
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   390
               Width           =   735
            End
            Begin MSForms.TextBox txtPCC 
               Height          =   300
               Index           =   6
               Left            =   810
               TabIndex        =   148
               TabStop         =   0   'False
               Top             =   420
               Visible         =   0   'False
               Width           =   3285
               VariousPropertyBits=   671107099
               MaxLength       =   70
               Size            =   "5794;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
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
            Height          =   525
            IntegralHeight  =   0   'False
            ItemData        =   "frm140402.frx":2180
            Left            =   1080
            List            =   "frm140402.frx":2187
            MultiSelect     =   1  '簡易多重選取
            Sorted          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1650
            Width           =   3180
         End
         Begin VB.Frame Frame2 
            Height          =   675
            Left            =   4275
            TabIndex        =   142
            Top             =   1590
            Width           =   4155
            Begin VB.ComboBox cboTitle 
               Height          =   300
               ItemData        =   "frm140402.frx":2195
               Left            =   810
               List            =   "frm140402.frx":2197
               TabIndex        =   60
               Text            =   "cboTitle"
               Top             =   120
               Width           =   3300
            End
            Begin VB.CommandButton cmdRemoveTit 
               Caption         =   "移除 ->"
               Height          =   255
               Left            =   45
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   390
               Width           =   735
            End
            Begin VB.CommandButton cmdAddTit 
               Caption         =   "<- 新增"
               Height          =   255
               Left            =   45
               TabIndex        =   61
               Top             =   120
               Width           =   735
            End
            Begin MSForms.TextBox txtPCC 
               Height          =   240
               Index           =   7
               Left            =   810
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   420
               Visible         =   0   'False
               Width           =   3285
               VariousPropertyBits=   671107099
               MaxLength       =   70
               Size            =   "5794;423"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
         End
         Begin VB.Frame Frame3 
            Height          =   675
            Left            =   2115
            TabIndex        =   121
            Top             =   2760
            Width           =   2760
            Begin VB.TextBox txtUserNo 
               Height          =   264
               Index           =   1
               Left            =   810
               MaxLength       =   7
               TabIndex        =   67
               Top             =   90
               Width           =   945
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "<- 新增"
               Height          =   255
               Index           =   1
               Left            =   45
               TabIndex        =   68
               Top             =   90
               Width           =   735
            End
            Begin VB.CommandButton cmdRemove 
               Caption         =   "移除 ->"
               Height          =   255
               Index           =   1
               Left            =   45
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   360
               Width           =   735
            End
            Begin MSForms.TextBox txtPCC 
               Height          =   300
               Index           =   12
               Left            =   810
               TabIndex        =   139
               TabStop         =   0   'False
               Top             =   360
               Visible         =   0   'False
               Width           =   1890
               VariousPropertyBits=   671107099
               MaxLength       =   70
               Size            =   "3016;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblName 
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   122
               Top             =   120
               Width           =   900
               VariousPropertyBits=   27
               Caption         =   "lblName"
               Size            =   "1587;450"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
         End
         Begin MSForms.ListBox lstUsers 
            Height          =   600
            Index           =   1
            Left            =   1080
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   2805
            Width           =   1035
            VariousPropertyBits=   746586139
            ScrollBars      =   2
            DisplayStyle    =   2
            Size            =   "1826;1058"
            MatchEntry      =   0
            MultiSelect     =   1
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCUID1 
            Height          =   276
            Left            =   2976
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   156
            Width           =   5508
            VariousPropertyBits=   671105055
            Size            =   "9716;487"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   26
            Left            =   3600
            TabIndex        =   75
            Top             =   3450
            Width           =   285
            VariousPropertyBits=   671107099
            MaxLength       =   26
            Size            =   "503;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   25
            Left            =   5850
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   2250
            Width           =   2055
            VariousPropertyBits=   671107103
            MaxLength       =   20
            Size            =   "3625;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   24
            Left            =   7335
            TabIndex        =   73
            Top             =   2790
            Width           =   330
            VariousPropertyBits=   671107099
            MaxLength       =   1
            Size            =   "582;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   2
            Left            =   1245
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
            VariousPropertyBits=   671107103
            Size            =   "1058;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   9
            Left            =   4500
            TabIndex        =   71
            Top             =   2460
            Width           =   285
            VariousPropertyBits=   671107099
            MaxLength       =   1
            Size            =   "503;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   11
            Left            =   1080
            TabIndex        =   66
            Top             =   2505
            Width           =   1035
            VariousPropertyBits=   671107099
            MaxLength       =   8
            Size            =   "1826;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   645
            Index           =   13
            Left            =   5520
            TabIndex        =   74
            Top             =   3150
            Width           =   3060
            VariousPropertyBits=   -1466941413
            MaxLength       =   500
            ScrollBars      =   2
            Size            =   "5397;1138"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   8
            Left            =   1080
            TabIndex        =   64
            Top             =   2190
            Width           =   3180
            VariousPropertyBits=   671107099
            MaxLength       =   50
            Size            =   "5609;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   10
            Left            =   7020
            TabIndex        =   72
            Top             =   2520
            Width           =   285
            VariousPropertyBits=   671107099
            MaxLength       =   1
            Size            =   "503;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   4
            Left            =   5310
            TabIndex        =   54
            Top             =   480
            Width           =   3180
            VariousPropertyBits=   671107099
            MaxLength       =   60
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
            TabIndex        =   53
            Top             =   480
            Width           =   3180
            VariousPropertyBits=   671107099
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
            TabIndex        =   55
            Top             =   780
            Width           =   3180
            VariousPropertyBits=   671107099
            MaxLength       =   30
            Size            =   "5609;529"
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
            Left            =   1080
            TabIndex        =   180
            Top             =   3630
            Width           =   2430
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否同意歐盟通用資料保護規範(GDPR)： "
            Height          =   180
            Index           =   37
            Left            =   120
            TabIndex        =   179
            Top             =   3435
            Width           =   3315
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "名片臨時編號："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   14
            Left            =   4560
            TabIndex        =   173
            Top             =   2280
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否寄專利雙週報：       （N:不寄）"
            Height          =   180
            Index           =   25
            Left            =   5715
            TabIndex        =   156
            Top             =   2835
            Width           =   2820
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "相關聯絡人編號："
            Height          =   180
            Index           =   8
            Left            =   4365
            TabIndex        =   145
            Top             =   780
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "開發日期：                         ( 西元 )"
            Height          =   180
            Index           =   9
            Left            =   135
            TabIndex        =   125
            Top             =   2550
            Width           =   2595
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "E-MAIL："
            Height          =   180
            Index           =   5
            Left            =   135
            TabIndex        =   120
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "聯絡人編號："
            Height          =   180
            Index           =   7
            Left            =   135
            TabIndex        =   140
            Top             =   180
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否寄台一雜誌：     （N:不寄)"
            Height          =   180
            Index           =   7
            Left            =   3060
            TabIndex        =   138
            Top             =   2565
            Width           =   2430
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "備註："
            Height          =   180
            Index           =   14
            Left            =   4920
            TabIndex        =   124
            Top             =   3150
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "開發人員："
            Height          =   180
            Index           =   12
            Left            =   135
            TabIndex        =   123
            Top             =   2805
            Width           =   900
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "部門："
            Height          =   180
            Index           =   4
            Left            =   135
            TabIndex        =   119
            Top             =   1050
            Width           =   540
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "職稱："
            Height          =   180
            Index           =   3
            Left            =   135
            TabIndex        =   118
            Top             =   1680
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "是否寄電子報：      （ N:不寄)"
            Height          =   180
            Index           =   11
            Left            =   5712
            TabIndex        =   117
            Top             =   2556
            Width           =   2364
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "名稱( 日 )："
            Height          =   180
            Index           =   2
            Left            =   4365
            TabIndex        =   116
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "名稱( 英 )："
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   115
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "名稱( 中 )："
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   114
            Top             =   780
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -69960
         TabIndex        =   111
         Top             =   3336
         Width           =   1815
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   21
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox txtUserNo 
            Height          =   264
            Index           =   0
            Left            =   810
            MaxLength       =   6
            TabIndex        =   20
            Top             =   120
            Width           =   945
         End
         Begin MSForms.Label lblName 
            Height          =   255
            Index           =   0
            Left            =   855
            TabIndex        =   112
            Top             =   450
            Width           =   585
            VariousPropertyBits=   27
            Caption         =   "lblName"
            Size            =   "1032;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.ComboBox cboCity 
         Height          =   276
         Left            =   -73665
         TabIndex        =   11
         Text            =   "cboCity"
         Top             =   2220
         Width           =   5415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm140402.frx":2199
         Height          =   1005
         Left            =   135
         TabIndex        =   177
         Top             =   360
         Width           =   8625
         _ExtentX        =   15219
         _ExtentY        =   1778
         _Version        =   393216
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   14
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
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
               ColumnWidth     =   1607.811
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1272.189
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1344.189
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1595.906
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   887.811
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               ColumnWidth     =   1128.189
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "其他     說明："
         Height          =   492
         Index           =   41
         Left            =   -69840
         TabIndex        =   194
         Top             =   4332
         Width           =   588
      End
      Begin MSForms.TextBox txtXYS03 
         Height          =   550
         Left            =   -69204
         TabIndex        =   28
         Top             =   4344
         Width           =   2950
         VariousPropertyBits=   -1466941413
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "5203;970"
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
         TabIndex        =   193
         Top             =   4620
         Width           =   1080
      End
      Begin MSForms.Label LblSourceN 
         Height          =   288
         Left            =   -72640
         TabIndex        =   192
         Top             =   4636
         Width           =   2550
         VariousPropertyBits=   27
         Caption         =   "LblSourceN"
         Size            =   "4498;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "來所原因："
         Height          =   180
         Index           =   34
         Left            =   -74760
         TabIndex        =   191
         Top             =   4332
         Width           =   900
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   996
         Index           =   2
         Left            =   -68064
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1752
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "3096;1746"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   600
         Index           =   0
         Left            =   -71004
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3456
         Width           =   1032
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "1826;1058"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU47N 
         Height          =   300
         Left            =   -69120
         TabIndex        =   10
         Top             =   1920
         Width           =   2796
         VariousPropertyBits=   671107099
         MaxLength       =   20
         Size            =   "4921;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   47
         Left            =   -71460
         TabIndex        =   9
         Top             =   1920
         Width           =   1308
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "2302;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   9
         Left            =   -73668
         TabIndex        =   8
         Top             =   1920
         Width           =   852
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   8
         Left            =   -73668
         TabIndex        =   7
         Top             =   1656
         Width           =   7332
         VariousPropertyBits=   671107099
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
         Top             =   1368
         Width           =   7332
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "12938;529"
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
         Top             =   1080
         Width           =   5412
         VariousPropertyBits=   671107099
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
         Top             =   816
         Width           =   5412
         VariousPropertyBits=   671107099
         MaxLength       =   30
         Size            =   "9551;529"
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
         Top             =   540
         Width           =   5412
         VariousPropertyBits=   671107099
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
         VariousPropertyBits=   671107099
         MaxLength       =   30
         Size            =   "9551;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   51
         Left            =   -74100
         TabIndex        =   184
         Top             =   5016
         Visible         =   0   'False
         Width           =   336
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   50
         Left            =   -71520
         TabIndex        =   25
         Top             =   4032
         Width           =   336
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   49
         Left            =   -74760
         TabIndex        =   158
         Top             =   2400
         Visible         =   0   'False
         Width           =   1305
         VariousPropertyBits=   671107099
         MaxLength       =   18
         Size            =   "2302;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   48
         Left            =   -69372
         TabIndex        =   17
         Top             =   3060
         Width           =   336
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   38
         Left            =   -71988
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   3756
         Visible         =   0   'False
         Width           =   948
         VariousPropertyBits=   671107099
         MaxLength       =   70
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   10
         Left            =   -66984
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2220
         Width           =   588
         VariousPropertyBits=   671105055
         MaxLength       =   3
         Size            =   "1032;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   34
         Left            =   -73212
         TabIndex        =   15
         Top             =   2820
         Width           =   372
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   35
         Left            =   -69684
         TabIndex        =   16
         Top             =   2856
         Width           =   336
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   39
         Left            =   -73668
         TabIndex        =   24
         Top             =   3756
         Width           =   1308
         VariousPropertyBits=   671107099
         MaxLength       =   12
         Size            =   "2302;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   29
         Left            =   -73845
         TabIndex        =   46
         Top             =   3690
         Width           =   3735
         VariousPropertyBits=   671107099
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
         TabIndex        =   47
         Top             =   3930
         Width           =   3735
         VariousPropertyBits=   671107099
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
         TabIndex        =   48
         Top             =   4170
         Width           =   3735
         VariousPropertyBits=   671107099
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
         TabIndex        =   49
         Top             =   4410
         Width           =   3735
         VariousPropertyBits=   671107099
         MaxLength       =   30
         Size            =   "6588;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   33
         Left            =   -73845
         TabIndex        =   50
         Top             =   4650
         Width           =   3735
         VariousPropertyBits=   671107099
         MaxLength       =   30
         Size            =   "6588;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   19
         Left            =   -73845
         TabIndex        =   36
         Top             =   1350
         Width           =   7500
         VariousPropertyBits=   671107099
         MaxLength       =   50
         Size            =   "13229;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   12
         Left            =   -70044
         TabIndex        =   14
         Top             =   2520
         Width           =   852
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   13
         Left            =   -73845
         TabIndex        =   30
         Top             =   450
         Width           =   2955
         VariousPropertyBits=   671107099
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   14
         Left            =   -69300
         TabIndex        =   31
         Top             =   450
         Width           =   2955
         VariousPropertyBits=   671107099
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   15
         Left            =   -73845
         TabIndex        =   32
         Top             =   750
         Width           =   2955
         VariousPropertyBits=   671107099
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   16
         Left            =   -69300
         TabIndex        =   33
         Top             =   750
         Width           =   2955
         VariousPropertyBits=   671107099
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   18
         Left            =   -69300
         TabIndex        =   35
         Top             =   1050
         Width           =   2955
         VariousPropertyBits=   671107099
         MaxLength       =   50
         Size            =   "5212;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   17
         Left            =   -73845
         TabIndex        =   34
         Top             =   1050
         Width           =   2955
         VariousPropertyBits=   671107099
         MaxLength       =   20
         Size            =   "5212;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   420
         Index           =   26
         Left            =   -73845
         TabIndex        =   43
         Top             =   2490
         Width           =   7545
         VariousPropertyBits=   -1467989989
         MaxLength       =   70
         ScrollBars      =   2
         Size            =   "13309;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   20
         Left            =   -73710
         TabIndex        =   37
         Top             =   1650
         Width           =   3555
         VariousPropertyBits=   671107099
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
         TabIndex        =   38
         Top             =   1650
         Width           =   3630
         VariousPropertyBits=   671107099
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
         TabIndex        =   39
         Top             =   1920
         Width           =   3555
         VariousPropertyBits=   671107099
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
         TabIndex        =   40
         Top             =   1920
         Width           =   3630
         VariousPropertyBits=   671107099
         MaxLength       =   30
         Size            =   "6403;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   24
         Left            =   -73710
         TabIndex        =   41
         Top             =   2190
         Width           =   3555
         VariousPropertyBits=   671107099
         MaxLength       =   30
         Size            =   "6271;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   28
         Left            =   -73845
         TabIndex        =   45
         Top             =   3375
         Width           =   615
         VariousPropertyBits=   671107099
         MaxLength       =   3
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   420
         Index           =   27
         Left            =   -73845
         TabIndex        =   44
         Top             =   2925
         Width           =   7545
         VariousPropertyBits=   -1467989989
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "13309;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   25
         Left            =   -69960
         TabIndex        =   42
         Top             =   2190
         Width           =   3630
         VariousPropertyBits=   671107099
         MaxLength       =   30
         Size            =   "6403;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   37
         Left            =   -73668
         TabIndex        =   19
         Top             =   3420
         Width           =   852
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   300
         Index           =   36
         Left            =   -73668
         TabIndex        =   18
         Top             =   3120
         Width           =   336
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCU 
         Height          =   384
         Index           =   40
         Left            =   -73668
         TabIndex        =   29
         Top             =   4910
         Width           =   7402
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13056;677"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboPCU11 
         Height          =   288
         Left            =   -73668
         TabIndex        =   13
         Top             =   2520
         Width           =   1716
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3016;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "正式啟用要將第一頁的txtPCU(47),txtPCU47N移到下方”關聯企業”"
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Index           =   40
         Left            =   -74790
         TabIndex        =   188
         Top             =   300
         Visible         =   0   'False
         Width           =   5265
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "關係企業"
         Height          =   180
         Left            =   -69864
         TabIndex        =   185
         Top             =   1980
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "查詢權限：                                     部門"
         Height          =   180
         Index           =   39
         Left            =   -74772
         TabIndex        =   181
         Top             =   5376
         Width           =   2928
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否同意歐盟通用資料保護規範(GDPR)：        （W:待回覆 Y:同意  N:不同意）"
         Height          =   180
         Index           =   36
         Left            =   -74772
         TabIndex        =   178
         Top             =   4080
         Width           =   6108
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "與                              為"
         Height          =   180
         Index           =   107
         Left            =   -71724
         TabIndex        =   176
         Top             =   1980
         Width           =   1800
      End
      Begin MSForms.Label lblPCU47N 
         Height          =   180
         Left            =   -70800
         TabIndex        =   175
         Top             =   540
         Width           =   1350
         ForeColor       =   16711680
         VariousPropertyBits=   27
         Caption         =   "第一頁txtPCU47N"
         Size            =   "2381;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPCU47 
         AutoSize        =   -1  'True
         Caption         =   "第一頁txtPCU(47)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   -73800
         TabIndex        =   174
         Top             =   540
         Width           =   975
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   5
         Left            =   -73800
         TabIndex        =   172
         Top             =   1620
         Width           =   6735
         VariousPropertyBits=   27
         Caption         =   "lbl2"
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
         TabIndex        =   171
         Top             =   1380
         Width           =   6735
         VariousPropertyBits=   27
         Caption         =   "lbl2"
         Size            =   "11880;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   3
         Left            =   -73800
         TabIndex        =   170
         Top             =   1140
         Width           =   6735
         VariousPropertyBits=   27
         Caption         =   "lbl2"
         Size            =   "11880;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   2
         Left            =   -70800
         TabIndex        =   169
         Top             =   870
         Width           =   1575
         VariousPropertyBits=   27
         Caption         =   "lbl2"
         Size            =   "2778;450"
         FontName        =   "新細明體-ExtB"
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
         TabIndex        =   168
         Top             =   870
         Width           =   900
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   1
         Left            =   -73320
         TabIndex        =   167
         Top             =   870
         Width           =   1455
         VariousPropertyBits=   27
         Caption         =   "lbl2"
         Size            =   "2566;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   0
         Left            =   -73800
         TabIndex        =   166
         Top             =   870
         Width           =   375
         VariousPropertyBits=   27
         Caption         =   "lbl2"
         Size            =   "661;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(日)："
         Height          =   180
         Index           =   32
         Left            =   -74340
         TabIndex        =   165
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國　　籍："
         Height          =   180
         Index           =   31
         Left            =   -74760
         TabIndex        =   164
         Top             =   870
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(英)："
         Height          =   180
         Index           =   30
         Left            =   -74340
         TabIndex        =   163
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名 稱(中)："
         Height          =   180
         Index           =   29
         Left            =   -74745
         TabIndex        =   162
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名　　稱："
         Height          =   180
         Index           =   28
         Left            =   -71760
         TabIndex        =   161
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "關聯："
         Height          =   180
         Index           =   27
         Left            =   -74760
         TabIndex        =   160
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "關聯代號："
         Height          =   180
         Index           =   26
         Left            =   -71040
         TabIndex        =   159
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "關聯企業："
         Height          =   180
         Index           =   55
         Left            =   -74760
         TabIndex        =   157
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄專利雙週報：       （N:不寄）"
         Height          =   180
         Index           =   18
         Left            =   -70992
         TabIndex        =   155
         Top             =   3156
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽同仁"
         Height          =   180
         Index           =   24
         Left            =   -68064
         TabIndex        =   153
         Top             =   2856
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最近日期"
         Height          =   180
         Index           =   23
         Left            =   -67128
         TabIndex        =   152
         Top             =   2856
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "＊：聯絡人已離職"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   146
         Top             =   1440
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "城市代碼："
         Height          =   180
         Index           =   8
         Left            =   -67980
         TabIndex        =   144
         Top             =   2256
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄台一雜誌：           （N:不寄）"
         Height          =   180
         Index           =   22
         Left            =   -74772
         TabIndex        =   137
         Top             =   2820
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報：        （N:不寄）"
         Height          =   180
         Index           =   17
         Left            =   -70992
         TabIndex        =   136
         Top             =   2856
         Width           =   2508
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "狀態："
         Height          =   180
         Index           =   21
         Left            =   -74772
         TabIndex        =   134
         Top             =   3756
         Width           =   540
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB1："
         Height          =   180
         Index           =   23
         Left            =   -74775
         TabIndex        =   133
         Top             =   3690
         Width           =   600
      End
      Begin VB.Label Label41 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "2："
         Height          =   180
         Index           =   24
         Left            =   -74445
         TabIndex        =   132
         Top             =   3930
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3："
         Height          =   180
         Index           =   25
         Left            =   -74445
         TabIndex        =   131
         Top             =   4170
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4："
         Height          =   180
         Index           =   26
         Left            =   -74445
         TabIndex        =   130
         Top             =   4410
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5："
         Height          =   180
         Index           =   27
         Left            =   -74445
         TabIndex        =   129
         Top             =   4650
         Width           =   270
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "網址："
         Height          =   180
         Index           =   6
         Left            =   -74775
         TabIndex        =   128
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "成立日期：                     ( 西元 )"
         Height          =   180
         Index           =   20
         Left            =   -70992
         TabIndex        =   127
         Top             =   2556
         Width           =   2412
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "類別："
         Height          =   180
         Index           =   15
         Left            =   -74772
         TabIndex        =   126
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "電話1："
         Height          =   180
         Index           =   9
         Left            =   -74775
         TabIndex        =   110
         Top             =   450
         Width           =   630
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "傳真1："
         Height          =   180
         Index           =   10
         Left            =   -74775
         TabIndex        =   109
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "E-MAIL："
         Height          =   180
         Index           =   11
         Left            =   -70140
         TabIndex        =   108
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "電話2："
         Height          =   180
         Index           =   12
         Left            =   -70140
         TabIndex        =   107
         Top             =   450
         Width           =   630
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "傳真2："
         Height          =   180
         Index           =   13
         Left            =   -70140
         TabIndex        =   106
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "行動電話："
         Height          =   180
         Index           =   15
         Left            =   -74775
         TabIndex        =   105
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "英文地址："
         Height          =   180
         Index           =   21
         Left            =   -74775
         TabIndex        =   104
         Top             =   1650
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "地址國籍："
         Height          =   180
         Index           =   22
         Left            =   -74775
         TabIndex        =   103
         Top             =   3375
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "中文地址："
         Height          =   180
         Index           =   28
         Left            =   -74775
         TabIndex        =   102
         Top             =   2925
         Width           =   900
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "lbl1"
         Height          =   180
         Index           =   2
         Left            =   -73200
         TabIndex        =   101
         Top             =   3375
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   32
         Left            =   -73845
         TabIndex        =   100
         Top             =   1650
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   33
         Left            =   -73845
         TabIndex        =   99
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   34
         Left            =   -73845
         TabIndex        =   98
         Top             =   2190
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   35
         Left            =   -70095
         TabIndex        =   97
         Top             =   1650
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   36
         Left            =   -70095
         TabIndex        =   96
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   37
         Left            =   -70095
         TabIndex        =   95
         Top             =   2190
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "日文地址："
         Height          =   180
         Index           =   20
         Left            =   -74775
         TabIndex        =   94
         Top             =   2490
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發人員："
         Height          =   180
         Index           =   3
         Left            =   -71940
         TabIndex        =   93
         Top             =   3456
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名稱（日）："
         Height          =   180
         Index           =   1
         Left            =   -74772
         TabIndex        =   92
         Top             =   1368
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名稱（英）："
         Height          =   180
         Index           =   2
         Left            =   -74772
         TabIndex        =   91
         Top             =   276
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名稱（中）："
         Height          =   180
         Index           =   4
         Left            =   -74772
         TabIndex        =   90
         Top             =   1656
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國籍："
         Height          =   180
         Index           =   5
         Left            =   -74772
         TabIndex        =   89
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "城市："
         Height          =   180
         Index           =   6
         Left            =   -74772
         TabIndex        =   88
         Top             =   2256
         Width           =   540
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "lbl1"
         Height          =   180
         Index           =   1
         Left            =   -72780
         TabIndex        =   87
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發日期：                         ( 西元 )"
         Height          =   180
         Index           =   13
         Left            =   -74772
         TabIndex        =   86
         Top             =   3456
         Width           =   2592
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文：             （1:中文   2:英文   3:日文）"
         Height          =   180
         Index           =   16
         Left            =   -74772
         TabIndex        =   85
         Top             =   3120
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   19
         Left            =   -74772
         TabIndex        =   84
         Top             =   4910
         Width           =   540
      End
   End
   Begin MSForms.TextBox textCUID 
      Height          =   270
      Left            =   2745
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   660
      Width           =   6150
      VariousPropertyBits=   671105055
      Size            =   "11695;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPCU 
      Height          =   300
      Index           =   2
      Left            =   2085
      TabIndex        =   1
      Top             =   645
      Width           =   255
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPCU 
      Height          =   300
      Index           =   1
      Left            =   1005
      TabIndex        =   0
      Top             =   645
      Width           =   1095
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "編號："
      Height          =   210
      Index           =   0
      Left            =   315
      TabIndex        =   82
      Top             =   675
      Width           =   585
   End
End
Attribute VB_Name = "frm140402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/5 改成Form2.0 (DataGrid1,txtPCU,txtPCC,textCUID,textCUID1,lblPCU47N,lbl2,lblName,lstUsers)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2007/11/2
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

Dim m_PCU41 As String 'Add By Sindy 2009/04/30
Dim TF_PCU As Integer
Dim TF_PCC As Integer
Dim strTmp As String
Dim oText As Object
Dim oLabel As Object
Dim idx As Integer
Dim rsContact As ADODB.Recordset
Dim rsContactOld As ADODB.Recordset
Dim rsContactSim As ADODB.Recordset
Dim m_sDupeKey As String
'Add By Sindy 2014/1/27
Dim m_sDupeKey_c As String
Dim m_sDupeKey_j As String
'2014/1/27 END
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Dim m_bSaveCheck As Boolean
Dim m_arrConRefList() As String '相關聯絡人資料
Dim m_iConEditMode As Integer '聯絡人狀態 1:新增 2:修改
Dim m_bReadGrid As Boolean '是否要讀取被點選聯絡人資料
Dim RsQ As New ADODB.Recordset 'Add by Amy 2021/08/16
Dim bCancel As Boolean, stMsg As String  'Add by Amy 2024/11/29

Private Sub cboCity_Change()
   txtPCU(10) = ""
   If cboCity <> "" Then
      For intI = 0 To cboCity.ListCount - 1
         If cboCity = cboCity.List(intI) Then
            cboCity.ListIndex = intI
            cboCity.SelStart = Len(cboCity)
         End If
      Next
   End If
End Sub

Private Sub cboCity_Click()
   If cboCity.ListIndex >= 0 Then
      txtPCU(10) = Right("000" & cboCity.ItemData(cboCity.ListIndex), 3)
   End If
End Sub

Private Sub cboCity_GotFocus()
   CloseIme
   If cboCity.Locked = False Then
      SendMessage cboCity.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cboCity_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub AddNewCity()
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHand

   strExc(0) = "select lpad(nvl(max(ct02),0)+1,3,'0') from city where ct01='" & Left(txtPCU(9), 3) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strSql = "insert into City(CT01,CT02,CT03) values('" & Left(txtPCU(9), 3) & "','" & RsTemp(0) & "','" & ChgSQL(cboCity.Text) & "')"
      cnnConnection.Execute strSql
      cnnConnection.CommitTrans
      txtPCU(10) = RsTemp(0)
   Else
      GoTo ErrHand
   End If
   Exit Sub
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Sub

Private Sub cboCity_LostFocus()
   SendMessage cboCity.hWnd, CB_SHOWDROPDOWN, 0, 0
End Sub

Private Sub cboCity_Validate(Cancel As Boolean)
   Dim idx As Integer
   If cboCity.Locked = True Then Exit Sub
   If cboCity <> "" And txtPCU(9) <> "" And txtPCU(10) = "" Then
      For idx = 0 To cboCity.ListCount - 1
         If cboCity = cboCity.List(idx) Then
            txtPCU(10) = Format(cboCity.ItemData(idx), "000")
            Exit For
         End If
      Next
      If txtPCU(10) = "" Then
         If MsgBox("您所輸入的城市尚未建檔，是否要新增該筆城市資料？", vbYesNo + vbQuestion) = vbYes Then
            AddNewCity
            SetCity
         Else
            Cancel = True
         End If
      End If
   End If
   If Cancel = False Then
      If DupeCustCheck(False) = True Then
         cboCity.Text = ""
         Cancel = True
      End If
   End If
   '預設英文地址6
   If Cancel = False And m_EditMode = 1 And txtPCU(10) <> "" Then
      txtPCU(25) = PUB_GetNationEngName(txtPCU(9))
      '國名與城市名稱不同時才要預設英文地址5
      If cboCity <> txtPCU(25) Then
         txtPCU(24) = cboCity
      End If
   End If
End Sub

Private Sub cboDept_GotFocus()
   If cboDept.Locked = False Then
      CloseIme
      SendMessage cboDept.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cboPCU11_Click()
   setPCU36
End Sub

'Add by Amy 2024/11/29 來所原因,規則同代理人維護
Private Sub cboSource_Click()
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   'Memo 與代理人維護 不同-可允許不輸-Widen
   If cboSource = MsgText(601) Then
      txtXYS02.Text = "": LblSourceN.Caption = ""
      Exit Sub
   End If
   
   txtXYS02.Text = "": LblSourceN.Caption = ""
   Call Pub_SetCboComeSource(9, Me.Name, cboSource, , txtXYS02, txtXYS03)
   
End Sub

Private Sub cboTitle_GotFocus()
   If cboTitle.Locked = False Then
      CloseIme
      SendMessage cboTitle.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

'新增開發人員
Private Sub cmdAdd_Click(Index As Integer)
   AddlstUsers Index
   If Index = 0 Then
      'Modified by Morgan 2022/1/7
      'txtPCU(38) = ComposeListX(Index)
      txtPCU(38) = ComposeListX(Index)
   Else
      txtPCC(12) = ComposeListX(Index)
   End If
   txtUserNo(Index).SetFocus
End Sub

'新增部門
Private Sub cmdAddDept_Click()
   If InStr(cboDept, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      cboDept.SetFocus
      Exit Sub
   End If
   AddLstFrmCbo cboDept, lstDept
   txtPCC(6) = ComposeList(lstDept)
   cboDept.SetFocus
End Sub

'新增職稱
Private Sub cmdAddTit_Click()
   If InStr(cboTitle, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      cboTitle.SetFocus
      Exit Sub
   End If
   AddLstFrmCbo cboTitle, lstTitle
   txtPCC(7) = ComposeList(lstTitle)
   cboTitle.SetFocus
End Sub

Private Sub cmdIntroduce_Click()
   Dim stName As String
    
   If cmdIntroduce.BackColor = &HFFFF80 Then
      '英->中->日
      If txtPCU(3) = MsgText(601) Then
         If txtPCU(8) = MsgText(601) Then
            stName = stName & " " & txtPCU(7)
         Else
            stName = stName & " " & txtPCU(8)
         End If
      Else
         stName = txtPCU(3)
         If txtPCU(4) <> MsgText(601) Then stName = stName & " " & txtPCU(4)
         If txtPCU(5) <> MsgText(601) Then stName = stName & " " & txtPCU(5)
         If txtPCU(6) <> MsgText(601) Then stName = stName & " " & txtPCU(6)
      End If
      frm050705_1.txtNo = txtPCU(1)
      frm050705_1.lbl1(0) = txtPCU(9) '國籍編號
      frm050705_1.lbl1(1) = lbl1(1)
      frm050705_1.lbl1(3) = stName

      frm050705_1.SetParent Me
      frm050705_1.QueryData
      frm050705_1.Show
      Me.Hide
   End If
End Sub

'移除開發人員
Private Sub cmdRemove_Click(Index As Integer)
   RemovelstUsers Index
   If Index = 0 Then
      txtPCU(38) = ComposeListX(Index)
   Else
      txtPCC(12) = ComposeListX(Index)
   End If
   txtUserNo(Index).SetFocus
End Sub

'聯絡人
Private Sub cmdContact_Click(Index As Integer)
   Dim bDupeCheck As Boolean, sPCC(1 To 5) As String
   Select Case Index
      Case 1 '新增
         ClearField1
         txtPCC(2).Text = getNewNo
         txtPCC(3).SetFocus
         txtPCC(9) = "N" 'Add by Morgan 2007/12/19 預設不寄台一雜誌
         '新增時開發日期預設當天
         'Modify by Morgan 2007/12/19 打字室預設20070101
         If Pub_StrUserSt03 = "M13" Then
            txtPCC(11) = 20070101
         Else
            txtPCC(11) = strSrvDate(1)
         End If
         
         '新增時預設開發人員與客戶的相同
         'Modify by Morgan 2007/12/19 都預設
         'If m_EditMode = 1 Then
            SetlstUsers 1, txtPCU(38)
         'End If
         
         txtPCC(24) = txtPCU(48) 'Added by Morgan 2012/2/2
         m_iConEditMode = 1
      Case 2 '加入
         If TxtValidate1 = True Then
            UpdateContact
            DataGrid1.Refresh
            '檢查是否有名稱相同的聯絡人
            bDupeCheck = False
            If m_iConEditMode = 1 Then
               bDupeCheck = True
            ElseIf txtPCC20 = "" Then
               bDupeCheck = ContNameChanged
            End If
            If bDupeCheck = True Then
               sPCC(1) = txtPCU(1)
               sPCC(2) = txtPCC(2)
               sPCC(3) = txtPCC(3)
               sPCC(4) = txtPCC(4)
               sPCC(5) = txtPCC(5)
               If PUB_DupeContactCheck(sPCC, rsContactSim) = True Then
                  Me.Tag = ""
                  Set frm140402_1.grdDataList.Recordset = rsContactSim
                  Set frm140402_1.fmParent = Me
                  frm140402_1.Show vbModal
                  UpdateConRefList txtPCC(2), Me.Tag
               End If
            End If
            ClearField1
         End If
         
      Case 3 '刪除
         If txtPCC(2) <> "" Then
            If Not (rsContact.EOF Or rsContact.BOF) Then
               If PUB_PCCDelCheck(txtPCU(1), txtPCU(2), txtPCC(2)) = True Then
                  rsContact.Delete
                  rsContact.UpdateBatch
                  RemoveConRefList
                  ClearField1
               End If
            End If
         End If
   End Select
End Sub

Private Sub cmdRemoveDept_Click()
   RemoveList lstDept
   txtPCC(6) = ComposeList(lstDept)
End Sub

Private Sub cmdRemoveTit_Click()
   RemoveList lstTitle
   txtPCC(7) = ComposeList(lstTitle)
End Sub

Private Sub cmdTransfer_Click()
   Dim bolOpen As Boolean, stMsg As String 'Add by Amy 2024/01/22
   
   'Add by Amy 2024/01/22 避免同時開啟維護畫面
   If PUB_CheckFormExist("frm050705") = True Then
      bolOpen = True
      stMsg = stMsg & "國外代理人資料維護" & vbCrLf
   End If
   If PUB_CheckFormExist("frm140401") = True Then
      bolOpen = True
      stMsg = stMsg & "客戶基本資料維護" & vbCrLf
   End If
   If bolOpen = True Then
      MsgBox stMsg & "已開啟,請關閉後再操作"
      Exit Sub
   End If
   
   'Added by Lydia 2020/08/27 預設國外關聯企業
   'Remove by Lydia 2021/01/06 潛在客戶不使用關聯企業設定
'   If strSrvDate(1) >= 國外部關聯企業啟用日 And txtPCU(47) <> "" Then
'       If Trim(txtPCU(49)) = "" Then
'          MsgBox "請先設定關聯企業的關係！"
'          tabCustomer.Tab = 3
'          Exit Sub
'       Else
'          frm140402_2.m_PCU47 = Trim(Left(ChangeCustomerL(txtPCU(47)), 8))
'          frm140402_2.m_PCU49 = Trim(txtPCU(49))
'       End If
'   End If
'   'end 2020/08/27
   'end 2021/01/06
   
   frm140402_2.Label2(1) = txtPCU(1) & "0"
   frm140402_2.Label2(3) = txtPCU(3)
   frm140402_2.Label2(4) = txtPCU(4)
   frm140402_2.Label2(5) = txtPCU(5)
   frm140402_2.Label2(6) = txtPCU(6)
   frm140402_2.Label2(7) = txtPCU(7)
   frm140402_2.Label2(2) = txtPCU(8)
   'Added by Lydia 2018/05/16 傳開發日期,開發人員
   frm140402_2.m_PCU37 = txtPCU(37)
   frm140402_2.m_PCU38 = txtPCU(38)
   'end 2018/05/16
   
   frm140402_2.Frame1.Visible = True 'Add by Amy 2023/05/08 +來所原因(原:代理人來源)
   'Add by Amy 2024/11/29 帶入來所原因 相關欄位(R編號轉X編號也要將資料寫入客戶檔備註中)
   If cboSource <> MsgText(601) Then frm140402_2.cboSource = cboSource
   frm140402_2.cboSource.Tag = m_FieldList(54).fiOldData
   frm140402_2.txtXYS02 = txtXYS02
   frm140402_2.txtXYS03 = txtXYS03
   '2009/6/26 ADD BY SONIA
   If txtPCU(47) <> "" Then
      frm140402_2.TextOldNo = txtPCU(47)
   End If
   '2009/6/26 END
   'modify by sonia 2021/12/29 txtPCU(11)-->Left(cboPCU11.Text, 1)
   If Left(cboPCU11.Text, 1) = "1" Then
      frm140402_2.Option1(1).Visible = True
      frm140402_2.Option1(2).Visible = True
      '2009/6/26 MODIFY BY SONIA
      'frm140402_2.Option1(1).Value = True
      If txtPCU(47) = "" Then
         frm140402_2.Option1(1).Value = True
      Else
         frm140402_2.Option1(2).Value = True
      End If
      '2009/6/26 END
      frm140402_2.Frame1.Visible = False 'Add by Amy 2023/05/08 +來所原因(原:代理人來源)
      frm140402_2.Frame1.Enabled = False
      'end 2024/11/29
      frm140402_2.Option1(3).Visible = False
      frm140402_2.Option1(4).Visible = False
   'modify by sonia 2021/12/29 txtPCU(11)-->Left(cboPCU11.Text, 1)
   ElseIf Left(cboPCU11.Text, 1) = "2" Then
      frm140402_2.Option1(1).Visible = False
      frm140402_2.Option1(2).Visible = False
      frm140402_2.Option1(3).Visible = True
      frm140402_2.Option1(4).Visible = True
      '2009/6/26 MODIFY BY SONIA
      'frm140402_2.Option1(3).Value = True
      If txtPCU(47) = "" Then
         frm140402_2.Option1(3).Value = True
      Else
         frm140402_2.Option1(4).Value = True
      End If
      '2009/6/26 END
   End If
   
   'Add By Sindy 2014/7/10
   '按 轉客戶或代理人 按鈕時, 若開發人員內有81040的資料,則CU153要設定為 'N'
   For idx = 0 To lstUsers(0).ListCount - 1
      'Modified by Morgan 2022/1/7
      'If lstUsers(0).ITEMDATA(idx) = PUB_Id2Num("81040") Then '閻副所長
      If PUB_GetItemData(lstUsers(0).Tag, idx) = "81040" Then
      'end 2022/1/7
         frm140402_2.txtCU153 = "N"
      End If
   Next idx
   '2014/7/10 END
   
   frm140402_2.Show
End Sub

Public Sub AfterTransfer()
   m_EditMode = 0
   SetCtrlReadOnly True
   UpdateToolbarState
   SetInputEntry
   ShowRecord 1, True
End Sub

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
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   'Add by Sindy 2021/6/28
   Call PUB_SetComboPCU11(cboPCU11, "") '設定國外潛在客戶類別選項
   
   'Add By Sindy 2009/06/23
   txtSameCnt.Visible = False
   
   textCUID.BackColor = &H8000000F
   textCUID1.BackColor = &H8000000F
   txtPCU(10).BackColor = &H8000000F
   InitialField
   'Add by Amy 2024/11/29 來所原因下拉選單
   cboSource.ListIndex = -1
   LblSourceN = ""
   Call Pub_SetCboComeSource(0, Me.Name, cboSource)
   'end 2024/11/29
   ShowRecord -2
   m_EditMode = 0
   SetInputEntry
   UpdateToolbarState
   tabCustomer.Tab = 0
   AddCombo 1
   AddCombo 2
   
   Pub_SetFTypeList Me.Combo1, 10 'Added by Lydia 2016/11/29 關聯企業：關聯代號下拉選單
      
   'Added by Lydia 2020/05/07 關聯企業：改用啟用日控制
   lblTitle.BackColor = &H8000000F
   'Modified by Lydia 2021/01/06 改回原名
'   If strSrvDate(1) >= 國外部關聯企業啟用日 Then
'        lblTitle.Caption = "關聯企業"
'        tabCustomer.TabVisible(3) = True
'        lblPCU47.Visible = False: lblPCU47N.Visible = False: Label1(40).Visible = False
'        lblTitle.Visible = False
'        Label1(107).Visible = False
'        txtPCU47N.Locked = True
'        'Memo by Lydia 2020/05/07 原程式在輸入PCU47後，預設下一步cboCity.SetFocus，但是新程式兩者在不同頁籤
'   Else
'   'end 2020/05/07
'        tabCustomer.TabVisible(3) = False  'Added by Lydia 2018/05/24 隱藏關聯企業頁籤
'   End If 'Added by Lydaia 2020/05/07
   tabCustomer.TabVisible(3) = False
   lblTitle.Caption = "關係企業"
   'end 2021/01/06
End Sub

Private Sub Form_Initialize()
   strExc(0) = "select * from PotCustomer where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_PCU = RsTemp.Fields.Count
   
   strExc(0) = "select * from PotCustCont where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_PCC = RsTemp.Fields.Count
   
   ReDim m_FieldList(TF_PCU) As FIELDITEM
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         'Add by Amy 2015/01/15 +備註可輸入換行
         'Modify by Amy 2024/11/29 不是txtPCC陣列元件會錯,並加[其他說明] 欄
         If KeyCode = vbKeyReturn Then
            If UCase(Me.ActiveControl.Name) = UCase("txtXYS03") Then
               Exit Sub
            ElseIf UCase(Me.ActiveControl.Name) = UCase("txtPCU") Then
               If Me.ActiveControl.Index = 40 Or Me.ActiveControl.Index = 13 Then
                  Exit Sub
               End If
            End If
         End If
         'end 2015/01/15
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyInsert
         If cmdContact(2).Enabled = True Then
            cmdContact_Click 2
         End If
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140402 = Nothing
End Sub


'Add By Sindy 2019/7/25
Private Sub Option1_Click(Index As Integer)
   If Option1(0).Value = True Then
      txtPCU(51) = "C"
   ElseIf Option1(1).Value = True Then
      txtPCU(51) = "F"
   End If
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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   
   strExc(0) = "SELECT * FROM PotCustomer " & _
            "WHERE PCU01 = '" & strKEY01 & "' AND PCU02 = '" & strKEY02 & "' "
                  
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      IsRecordExist = True
   End If
   
End Function

Private Sub txtPCC_GotFocus(Index As Integer)
   Select Case Index
      Case 4, 5, 13
         OpenIme
         
      Case Else
         CloseIme
         
   End Select
   TextInverse txtPCC(Index)
End Sub

Private Sub txtPCC_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 8
         PUB_EMailFilter CInt(KeyAscii) 'Added by Morgan 2011/11/30 Email輸入字元檢查
      'Modified by Morgan 2012/2/2 +24
      'Modified by Lydia 2020/11/25 拿掉”是否寄電子報”
      'Case 9, 10, 24
      'Modified by Lydia 2025/05/15 只有代理人才可輸入Y
      'Case 9, 24
      Case 9, 10, 24
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      'Modify By Sindy 2018/5/30 +25
      Case 25
         KeyAscii = UpperCase(KeyAscii)
      'Added by Lydia 2018/06/28 是否同意歐盟通用資料保護規範(GDPR)
      Case 26
          KeyAscii = UpperCase(KeyAscii)
          'Modified by Lydia 2018/08/10 GDPR增加輸入W(KeyAscii = 87)
          If KeyAscii <> 87 And KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
             KeyAscii = 0
             Beep
          End If
      'Added by Lydia 2020/11/25 是否寄電子報欄位，請新增Y:寄
      'Mark by Lydia 2025/05/15 只有代理人才可輸入Y
      'Case 10
       '  KeyAscii = UpperCase(KeyAscii)
       '  If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
       '     KeyAscii = 0
       '     Beep
       '  End If
       'end 2025/05/15
   End Select
End Sub

Private Sub txtPCC_Validate(Index As Integer, Cancel As Boolean)
   Dim iLen As Integer
   If txtPCC(Index).Locked = True Then Exit Sub
   Select Case Index
      Case 8
         If txtPCC(Index) <> "" Then
            If InStr(1, txtPCC(Index), "@") = 0 Then
                MsgBox "Mail 必需要有 @ 符號！"
                txtPCC(Index).SetFocus
                Cancel = True
            'Modify by Amy 2017/12/14 +?
            ElseIf InStr(1, txtPCC(Index), ",") > 0 Or InStr(1, txtPCC(Index), "[") > 0 Or InStr(1, txtPCC(Index), "]") > 0 Or InStr(1, txtPCC(Index), "!") > 0 Or InStr(1, txtPCC(Index), "(") > 1 Or InStr(1, txtPCC(Index), ")") > 0 Or InStr(1, txtPCC(Index), "=") > 0 Or InStr(1, txtPCC(Index), "\") > 0 Or InStr(1, txtPCC(Index), "/") > 0 Or InStr(1, txtPCC(Index), "<") > 0 Or InStr(1, txtPCC(Index), ">") > 0 _
              Or InStr(1, txtPCC(Index), "~") > 0 Or InStr(1, txtPCC(Index), "$") > 0 Or InStr(1, txtPCC(Index), "%") > 0 Or InStr(1, txtPCC(Index), "^") > 0 Or InStr(1, txtPCC(Index), "&") > 0 Or InStr(1, txtPCC(Index), "*") > 0 Or InStr(1, txtPCC(Index), "?") > 0 Then
                MsgBox "Mail 不允許有下列符號！" & vbCrLf & ",、[、]、!、(、)、=、\、/、<、>、~、$、%、^、&、* 、?"
                txtPCC(Index).SetFocus
                Cancel = True
            End If
         End If
         
      Case 11
         If txtPCC(Index) <> "" Then
            If CheckIsDate(txtPCC(Index)) = False Then
               txtPCC_GotFocus Index
               Cancel = True
            End If
         End If
   End Select
   If Cancel = False Then
      '欄位長度檢查
      iLen = txtPCC(Index).MaxLength
      'Added by Lydia 2018/02/22 英中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
      'Modified by Lydia 2018/07/04 日文名稱、中文名稱、部門、職稱、備註可輸入造字
      'If Index = 3 Or Index = 4 Or Index = 5 Then
      If Index = 4 Or Index = 5 Or Index = 6 Or Index = 7 Or Index = 13 Then
         iLen = iLen - 1
      End If
      'end 2018/02/22
      If Not CheckLengthIsOK(txtPCC(Index), iLen) Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtPCU_Change(Index As Integer)
   Select Case Index
      Case 9
         If Left(txtPCU(9).Tag, 3) <> Left(txtPCU(9), 3) Then
            txtPCU(10) = ""
            cboCity.Clear
            lbl1(1) = ""
            If Len(txtPCU(9)) >= 3 Then
               If ClsPDGetNation(Left(txtPCU(9), 3), strTmp) = True Then
                  lbl1(1).Caption = strTmp
                  SetCity
               End If
            End If
         End If
         txtPCU(9).Tag = txtPCU(9)
         SetPCU48
      Case 28
         lbl1(2) = ""
         If Len(txtPCU(28)) = 3 Then
            If ClsPDGetNation(txtPCU(28), strTmp) = True Then
               lbl1(2).Caption = strTmp
            End If
         End If
   End Select
End Sub

'Added by Morgan 2011/12/30
Private Sub SetPCU48()
   If txtPCU(9) = "020" Then
      txtPCU(48) = ""
   Else
      txtPCU(48) = "N"
   End If
End Sub

Private Sub txtPCU_GotFocus(Index As Integer)
   Select Case Index
      Case 7, 8, 26, 27, 40
         OpenIme
         
      Case Else
         CloseIme
         
   End Select
   '國籍第4碼檢查錯誤時
   If m_bSaveCheck = True Then
      txtPCU(Index).SelStart = 3
      txtPCU(Index).SelLength = 1
      m_bSaveCheck = False
      Exit Sub
   End If
   
   TextInverse txtPCU(Index)
   
End Sub

Private Sub txtPCU_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      'Modify by Amy 2024/11/29 +9國籍 ex:011t不會自動轉大寫
      Case 1, 2, 9, 47
         KeyAscii = UpperCase(KeyAscii)
'      Case 11
'         If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'            KeyAscii = 0
'            Beep
'         End If
      Case 18
         PUB_EMailFilter CInt(KeyAscii) 'Added by Morgan 2011/11/30 Email輸入字元檢查
      'Modified by Morgan 2012/2/2 +48
      Case 34, 35, 48
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      Case 36
         If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
            KeyAscii = 0
            Beep
         End If
      'Added by Lydia 2018/06/28 是否同意歐盟通用資料保護規範(GDPR)
      Case 50
          KeyAscii = UpperCase(KeyAscii)
          'Modified by Lydia 2018/08/10 GDPR增加輸入W(KeyAscii = 87)
          If KeyAscii <> 87 And KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
             KeyAscii = 0
             Beep
          End If
   End Select
End Sub

Private Sub txtPCU_Validate(Index As Integer, Cancel As Boolean)
   If txtPCU(Index).Locked = True Then Exit Sub
   Dim iLen As Integer
   Select Case Index
      Case 1
         If Not IsEmptyText(txtPCU(1)) Then
            If Mid(txtPCU(1), 1, 1) <> "R" Then
               Cancel = True
               MsgBox "客戶編號必須為R開頭", vbCritical + vbOKOnly, "檢核資料"
               txtPCU(1).Text = ""
               txtPCU_GotFocus 1
               Exit Sub
            End If
            
            If Len(txtPCU(1)) < 6 Then
               Cancel = True
               MsgBox "客戶編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
               txtPCU_GotFocus 1
               Exit Sub
            End If
            
            txtPCU(1) = Left(txtPCU(1) & "00", 8)
            txtPCU(2) = Left(txtPCU(2) & "0", 1)
            If m_EditMode = 1 Then
               If IsRecordExist(txtPCU(1), txtPCU(2)) = True Then
                  Cancel = True
                  MsgBox "該筆客戶已存在! ", vbCritical + vbOKOnly, "檢核資料"
                  txtPCU_GotFocus 1
                  Exit Sub
               End If
               If IsOverAutoNumber("R", Empty, Mid(txtPCU(1), 2, 5)) = True Then
                  Cancel = True
                  MsgBox "客戶代碼超過自動編號! ", vbCritical + vbOKOnly, "檢核資料"
                  txtPCU_GotFocus 1
                  Exit Sub
               End If
            End If
          End If
      
      'Add By Sindy 2012/4/9
      Case 3, 4, 5, 6, 7, 8 '名稱
         If txtPCU(Index) <> "" Then
            If DupeCustCheck(True, Index) = True Then
               Cancel = True
               txtPCU_GotFocus Index
               Exit Sub
            End If
         End If
         
      Case 9 '國籍
         If txtPCU(9) <> "" Then
            If txtPCU(9) = 台灣國家代號 Then
               Cancel = True
               'modify by sonia 2016/11/24
               'ShowMsg MsgText(9153)
               MsgBox "台灣國籍, 請依地址改輸 001 ~ 008 ！"
               'end 2016/11/24
            Else
               If lbl1(1).Caption = "" Then
                  Cancel = True
               Else
                  '英文地址、地址國籍都空白時地址國籍預設客戶國籍
                  'Remove by Morgan 2007/12/19 改存檔前若空白才設定
                  'If txtPCU(20) = "" And txtPCU(28) = "" Then
                  '   txtPCU(28) = Left(txtPCU(9), 3)
                  '   lbl1(2).Caption = lbl1(1)
                  'End If
                  'end 2007/12/19
                  
                  'Modified by Morgan 2023/7/6
                  'If txtPCU(36) = "" Then
                  '   If Val(txtPCU(9)) < 9 Or txtPCU(9) = "013" Or txtPCU(9) = "020" Then
                  '      txtPCU(36) = 1
                  '   Else
                  '      txtPCU(36) = 2
                  '   End If
                  'End If
                  setPCU36
                  'end 2023/7/6
               End If
            End If
         End If
         
      Case 10 '城市
         ShowCity
         
      Case 12, 37
         If txtPCU(Index) <> "" Then
            If CheckIsDate(txtPCU(Index)) = False Then
               tabCustomer.Tab = 0
               txtPCU(Index).SetFocus
               txtPCU_GotFocus Index
               Cancel = True
            End If
         End If
         
      Case 18 'e-mail
         If txtPCU(Index) <> "" Then
            'Modified by Lydia 2019/08/19 Email輸入字元檢查(與客戶檔frm140401一致)
'            If InStr(1, txtPCU(Index), "@") = 0 Then
'                MsgBox "Mail 必需要有 @ 符號！"
'                Cancel = True
'            'Modify by Amy 2017/12/14 +?
'            ElseIf InStr(1, txtPCU(Index), ",") > 0 Or InStr(1, txtPCU(Index), "[") > 0 Or InStr(1, txtPCU(Index), "]") > 0 Or InStr(1, txtPCU(Index), "!") > 0 Or InStr(1, txtPCU(Index), "(") > 1 Or InStr(1, txtPCU(Index), ")") > 0 Or InStr(1, txtPCU(Index), "=") > 0 Or InStr(1, txtPCU(Index), "\") > 0 Or InStr(1, txtPCU(Index), "/") > 0 Or InStr(1, txtPCU(Index), "<") > 0 Or InStr(1, txtPCU(Index), ">") > 0 _
'              Or InStr(1, txtPCU(Index), "~") > 0 Or InStr(1, txtPCU(Index), "$") > 0 Or InStr(1, txtPCU(Index), "%") > 0 Or InStr(1, txtPCU(Index), "^") > 0 Or InStr(1, txtPCU(Index), "&") > 0 Or InStr(1, txtPCU(Index), "*") > 0 Or InStr(1, txtPCU(Index), "?") > 0 Then
'                MsgBox "Mail 不允許有下列符號！" & vbCrLf & ",、[、]、!、(、)、=、\、/、<、>、~、$、%、^、&、* 、?"
            If PUB_CheckMail(txtPCU(Index)) = False Then
                txtPCU(Index).SetFocus
                txtPCU_GotFocus Index
            'end 2019/08/19
                Cancel = True
            End If
         End If
      
      'Add by Amy 2017/12/14 不可有?
      Case 19 '網址
        If txtPCU(Index) <> "" Then
            If InStr(1, txtPCU(Index), "?") > 0 Then
                MsgBox "網址 不允許有問號！"
                Cancel = True
            End If
        End If
      Case 27 '地址-中
        If txtPCU(Index) <> "" Then
            If InStr(1, txtPCU(Index), "?") > 0 Then
                MsgBox "中文地址 不允許有問號！"
                Cancel = True
            End If
        End If
       'end 2017/12/14
       
      Case 28 '地址國籍
         If txtPCU(28) <> "" Then
            If lbl1(2).Caption = "" Then
               Cancel = True
            End If
         End If
      'Add by Amy 2017/12/14 不可有?
      Case 29, 30, 31, 32, 33 'POB
         If txtPCU(Index) <> "" Then
            If InStr(1, txtPCU(Index), "?") > 0 Then
                MsgBox "POB" & Index - 28 & " 不允許有問號！"
                Cancel = True
            End If
         End If
      'end 2017/12/14
      'Add By Sindy 2009/06/23
      Case 47 '關係企業
         If txtPCU(Index) <> "" Then
            If Len(txtPCU(Index)) > 5 Then
               txtPCU(Index) = Left(txtPCU(Index) & "000", 9)
               'modify by sonia 2021/12/29 txtPCU(11)-->Left(cboPCU11.Text, 1)
               If GetCustData(txtPCU(Index), Left(cboPCU11.Text, 1)) = False Then
                  If m_EditMode = "1" Or m_EditMode = "2" Then
                     Cancel = True
                     txtPCU_GotFocus Index
                  End If
               End If
            Else
               Cancel = True
               'Modified by Lydia 2020/05/07 關係企業=>lblTitle.Caption
               MsgBox lblTitle.Caption & "編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
               txtPCU_GotFocus Index
            End If
         End If
         
      'Added by Morgan 2012/2/2
      Case 48
         If txtPCU(9) <> "020" And txtPCU(48) = "" Then
            Cancel = True
            MsgBox "潛在客戶國籍不是大陸，不可設為要寄專利雙週報 !"
            txtPCU_GotFocus Index
         End If
   End Select
   
   If Cancel = False Then
      '欄位長度檢查
      Select Case Index
         '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
         Case 7, 8, 26, 27, 40
            iLen = txtPCU(Index).MaxLength - 1
         Case Else
            iLen = txtPCU(Index).MaxLength
      End Select
      
      'Added by Lydia 2021/01/07 名稱長度直接以TextBox.MaxLength控制; ex.R15419的英文名稱"Patentanwaelte · Rechtsanwaelt"字數30字，中英文長度31
      If Index >= 3 And Index <= 8 Then
      Else
      'end 2021/01/07
            If Not CheckLengthIsOK(txtPCU(Index), iLen) Then
               Cancel = True
            End If
      End If 'Added by Lydia 2021/01/07
   End If
   
End Sub

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim strTp As String 'Add by Amy 2023/07/12
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         m_sDupeKey = ""
         'Add By Sindy 2014/1/27
         m_sDupeKey_c = ""
         m_sDupeKey_j = ""
         '2014/1/27 END
         'Add by Morgan 2007/12/19 預設不寄台一雜誌
         txtPCU(34) = "N"
         'Modify by Morgan 2007/12/19
         '打字室預設 1.類別=2,2.開發日=20070101,3.開發人=81040
         If Pub_StrUserSt03 = "M13" Then
            'modify by sonia 2021/12/29 txtPCU(11)-->cboPCU11.Text
            cboPCU11.Text = "2"
            txtPCU(37) = 20070101
            txtPCU(38) = "81040"
            SetlstUsers 0, txtPCU(38)
         Else
            '開發日期預設當天
            txtPCU(37) = strSrvDate(1)
         End If
         tabCustomer.Tab = 0
         'Add By Sindy 2019/7/24 國內外權限
         If strUserNum = "67002" Then
            Option1(0).Value = True
         Else
            Option1(1).Value = True
         End If
         '2019/7/24 END
         
      Case vbKeyF3 ' 修改
         'Modified by Lydia 2019/06/18
         'If CheckModifyLimit(m_PCU41, True) = False Then Exit Sub 'Add By Sindy 2009/04/30
         'Modify By Sindy 2019/7/25
         'If PUB_CheckModifyLimit_frm140402(m_PCU41, "M") = False Then Exit Sub
         'Modify By Sindy 2024/10/1 傳入建檔人
         If PUB_CheckModifyLimit_frm140402(txtPCU(51), m_PCU41) = False Then Exit Sub
         '2019/7/25 END
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         'Modify By Sindy 2019/7/25
         'If PUB_CheckModifyLimit_frm140402(m_PCU41, "D") = False Then Exit Sub 'Added by Lydia 2019/06/17
         'Modify By Sindy 2024/10/1 傳入建檔人
         If PUB_CheckModifyLimit_frm140402(txtPCU(51), m_PCU41) = False Then Exit Sub
         '2019/7/25 END
         'Add by Amy 2023/07/12 若存在XYS02介紹來源編號,則不可刪
         'Modify by Amy 2024/11/29 考慮多筆,改訊息至共用
         If txtPCU(2) = "0" And Pub_GetXYSource(2, txtPCU(1), , , , Me.Name, strTp) = True Then
            MsgBox strTp, vbOKOnly, "注意"
            Exit Sub
         End If
            
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
         SetCtrlReadOnly True
         ClearField
         UpdateToolbarState
         SetInputEntry
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
         
      Case vbKeyF9 ' 確定
         'Add By Sindy 2009/06/23
         'If txtSameCnt <> "" Then Exit Sub
         'Modify By Amy 2015/01/15 +不過濾的文字框.name
         'Modify by Amy 2024/11/29 +txtXYS03
         PUB_FilterFormText Me, "txtPCU(40),txtPCC(13),txtXYS03"
         
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SetInputEntry
         
      Case vbKeyF10 ' 取消
         'Add By Sindy 2009/06/23
         'If txtSameCnt <> "" Then Exit Sub
         
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  txtPCU(1) = txtPCU(1).Tag
                  txtPCU(2) = txtPCU(2).Tag
                  m_EditMode = 0
                  SetInputEntry
                  ShowRecord
                  UpdateToolbarState
               End If
            Case Else
               txtPCU(1) = txtPCU(1).Tag
               txtPCU(2) = txtPCU(2).Tag
               m_EditMode = 0
               SetInputEntry
               ShowRecord
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
      'tabCustomer.Tab = 0
   End If
End Sub

Private Sub ClearField()
   For Each oText In txtPCU
      oText.Text = Empty
   Next
   For Each oLabel In lbl1
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_PCU
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   'Add By Sindy 2009/06/24
   txtPCU47N = ""
   '2009/06/24 End
   cboCity.Clear
   textCUID = ""
   txtUserNo(0) = ""
   lblName(0) = ""
   lstUsers(0).Clear
   lstUsers(0).Tag = "" 'Added by Morgan 2022/1/7
   lstUsers(2).Clear
   'Added by Lydia 2016/11/29 關聯企業
   Combo1.ListIndex = -1
   List1.Clear
   For Each oLabel In Lbl2
      oLabel.Caption = Empty
   Next
   lblPCU47.Caption = ""
   lblPCU47N.Caption = ""
   'end 2016/11/29
   
   'Add By Sindy 2019/7/24 國內外權限
   Option1(0).Value = False
   Option1(1).Value = False
   '2019/7/24 END
   cboPCU11.ListIndex = -1 'Add By Sindy 2021/6/28
   'Add by Amy 2024/11/29
   cboSource.ListIndex = -1
   txtXYS02 = ""
   txtXYS03 = ""
   LblSourceN.Caption = "" 'X or Y or R編號 名稱
   txtXYS02.Tag = ""
   txtXYS03.Tag = ""
   cmdIntroduce.BackColor = &H8000000F
   'end 2024/11/29
   OpenContactTable
   ClearField1
End Sub

Private Sub ClearField1()
   For Each oText In txtPCC
      oText.Text = Empty
   Next
   lstDept.Clear
   lstTitle.Clear
   txtUserNo(1) = ""
   lblName(1) = ""
   lstUsers(1).Clear
   lstUsers(1).Tag = "" 'Added by Morgan 2022/1/7
   cboDept = ""
   cboTitle = ""
   textCUID1 = ""
   txtPCC20 = ""
   'Added by Lydia 2024/05/14
   Command1.Visible = False
   Command1.Caption = "上傳相片"
   Command1.BackColor = &H8080FF     '紅色
   'end 2024/05/14
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   Dim intSourceState As Integer 'Add by Amy 2024/11/29
   
   For Each oText In txtPCU
      oText.Locked = bLocked
   Next
   cboCity.Locked = bLocked
   Frame1.Visible = Not bLocked
   
   'Add By Sindy 2009/06/24
   txtPCU47N.Enabled = Not bLocked
   
   cmdContact(1).Enabled = Not bLocked
   cmdContact(2).Enabled = Not bLocked
   cmdContact(3).Enabled = Not bLocked
   If m_EditMode = 2 And Pub_StrUserSt03 <> "M51" Then
      cmdContact(3).Enabled = False
   End If
   'fraContact.Enabled = Not bLocked
   For Each oText In txtPCC
      oText.Locked = bLocked
   Next
   txtPCC(2).Locked = True
   cboPCU11.Locked = bLocked 'Add By Sindy 2021/6/28
   
   Frame2.Visible = Not bLocked
   Frame3.Visible = Not bLocked
   Frame4.Visible = Not bLocked
   
   If bLocked = False Then
      SetCity
   End If
   'Added by Lydia 2016/11/29
   Combo1.Locked = bLocked
   List1.Enabled = Not bLocked
   cmdAdd(0).Enabled = Not bLocked
   cmdRemove(0).Enabled = Not bLocked
   cmdAdd(1).Enabled = Not bLocked
   cmdRemove(1).Enabled = Not bLocked
   cmdAddPCU49.Enabled = Not bLocked
   cmdRemPCU49.Enabled = Not bLocked
   'end 2016/11/29
   'Add by Amy 2024/11/29 來所原因 相關欄位設定
   'Memo 潛在客戶不會有更名問題,故不需判斷txtpcu(2)="0"
   If m_EditMode = 1 Then
      intSourceState = 6
   ElseIf m_EditMode = 2 Then
      intSourceState = 7
   Else
      intSourceState = 8
   End If
   Call Pub_SetCboComeSource(intSourceState, Me.Name, cboSource, , txtXYS02, txtXYS03)
   'end 2024/11/29
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate And txtPCU(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtPCU(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtPCU(1) <> "" Then
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
         cmdTransfer.Enabled = False      '2008/12/4 add by sonia
      
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
         '2008/12/4 add by sonia 有修改權限且非打字室M13者才可轉客戶或代理人
         If m_bUpdate And txtPCU(1) <> "" And GetStaffDepartment(strUserNum) <> "M13" Then
            cmdTransfer.Enabled = True
         Else
            cmdTransfer.Enabled = False
         End If
         '2008/12/4 end
   End Select
   
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      Select Case m_EditMode
         Case 1
            txtPCU(1).Locked = False
            txtPCU(2).Locked = False
            txtPCU(3).SetFocus
            SetCity
         Case 2
            txtPCU(1).Locked = True
            txtPCU(2).Locked = True
            txtPCU(3).SetFocus
            SetCity
         Case 4
            txtPCU(1).Locked = False
            txtPCU(2).Locked = False
            txtPCU(1).SetFocus
         Case Else
            txtPCU(1).Locked = True
            txtPCU(2).Locked = True
            txtPCU(1).SetFocus
      End Select
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean, ii As Integer, jj As Integer
   Dim iRtn As Integer, stMsg As String 'Add by Amy 2021/11/29
   
   'Added by Morgan 2022/1/5 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2022/1/5
   
   '查詢
   If m_EditMode = 4 Then
      If txtPCU(1) = "" Then
         ShowMsg "請輸入欲查詢之客戶編號 !"
         txtPCU(1).SetFocus
         txtPCU_GotFocus 1
         Exit Function
      End If
   '維護
   Else
      If txtPCU(3) = "" And txtPCU(7) = "" And txtPCU(8) = "" Then
         ShowMsg "客戶中、英、日文名稱不可同時為空白 !"
         txtPCU(3).SetFocus
         tabCustomer.Tab = 0
         Exit Function
      End If
      'Add by Amy 2021/11/29 國內同業判斷
      If Left(Trim(txtPCU(9)), 3) < "010" And InStr(txtPCU(7), "事務所") > 0 And txtPCU(39) <> "國內同業" Then
        '取消
        If iRtn = 2 Then
            Exit Function
        '是
        ElseIf iRtn = 6 Then
            txtPCU(39) = "國內同業"
        End If 'iRtn
      End If
      If txtPCU(39) = "國內同業" Then
        '電子報要設定不寄
        If txtPCU(34) <> "N" Then
            ShowMsg "此為國內同業, 不可寄台一雜誌 ！"
            txtPCU(34).SetFocus
            txtPCU_GotFocus (34)
            tabCustomer.Tab = 0
            Exit Function
        End If
        If txtPCU(35) <> "N" Then
            ShowMsg "此為國內同業, 不可寄電子報！"
            txtPCU(35).SetFocus
            txtPCU_GotFocus (35)
            tabCustomer.Tab = 0
            Exit Function
        End If
        If txtPCU(48) <> "N" Then
            ShowMsg "此為國內同業, 不可專利雙週報！"
            txtPCU(48).SetFocus
            txtPCU_GotFocus (48)
            tabCustomer.Tab = 0
            Exit Function
        End If
        '不可寄mail
        If txtPCU(18) <> MsgText(601) Then
            ShowMsg "此為國內同業,不可輸入E-Mail以免誤發電子郵件, 如有需要請加註於備註欄 ！"
            txtPCU(18).SetFocus
            txtPCU_GotFocus (18)
            Exit Function
        End If
        stMsg = ChkSameTradePCU
        If stMsg <> MsgText(601) Then
            MsgBox "聯絡人資料需修改如下:" & vbCrLf & stMsg
            tabCustomer.Tab = 2
            Exit Function
        End If
      End If
      'emd 2021/11/29
      
      '有輸入客戶狀態時，不寄雜誌、電子報
      If txtPCU(39) <> "" Then
         txtPCU(34) = "N"
         txtPCU(35) = "N"
      End If
      
      If txtPCU(9).Text = "" Then
         ShowMsg "客戶國籍不可為空白 !"
         txtPCU(9).SetFocus
         tabCustomer.Tab = 0
         Exit Function
      End If
      
      '檢查英文名稱第一碼
      m_bSaveCheck = True
'edit by nickc 2008/05/08 改共用 function
'      If Mid(txtPCU(9), 1, 3) = "101" Then
'          '2008/1/4 MODIFY BY SONIA 原為A~I為101,J~Z為1011,2008年改為分四段
'          If Mid(UCase(LTrim(txtPCU(3))), 1, 1) >= "A" And Mid(UCase(LTrim(txtPCU(3))), 1, 1) <= "E" Then
'               If Trim(txtPCU(9)) <> "101" Then
'                  ShowMsg "客戶英文名稱第一碼介於 A~E 之間，客戶國籍應該為 101 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          ElseIf Mid(UCase(LTrim(txtPCU(3))), 1, 1) >= "F" And Mid(UCase(LTrim(txtPCU(3))), 1, 1) <= "I" Then
'               If Trim(txtPCU(9)) <> "1011" Then
'                  ShowMsg "客戶英文名稱第一碼介於 F~I 之間，客戶國籍應該為 1011 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          ElseIf Mid(UCase(LTrim(txtPCU(3))), 1, 1) >= "J" And Mid(UCase(LTrim(txtPCU(3))), 1, 1) <= "N" Then
'               If Trim(txtPCU(9)) <> "1012" Then
'                  ShowMsg "客戶英文名稱第一碼介於 J~N 之間，客戶國籍應該為 1012 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          ElseIf Mid(UCase(LTrim(txtPCU(3))), 1, 1) >= "O" And Mid(UCase(LTrim(txtPCU(3))), 1, 1) <= "Z" Then
'               If Trim(txtPCU(9)) <> "1013" Then
'                  ShowMsg "客戶英文名稱第一碼介於 O~Z 之間，客戶國籍應該為 1013 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          '2008/1/9 add by sonia
'          Else
'               If Trim(txtPCU(9)) <> "1013" Then
'                  ShowMsg "客戶英文名稱第一碼非英文字母或無英文名稱，客戶國籍應該為 1013 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          '2008/1/9 end
'          End If
'      ElseIf Mid(txtPCU(9), 1, 3) = "011" Then
'          '2008/4/21 MODIFY BY SONIA 原為A~L為011,M~Z為0111,2008/4/22改為分三段(將M~Z再細分成二段)
'          If Mid(UCase(LTrim(txtPCU(3))), 1, 1) >= "A" And Mid(UCase(LTrim(txtPCU(3))), 1, 1) <= "L" Then
'               If Trim(txtPCU(9)) <> "011" Then
'                  ShowMsg "客戶英文名稱第一碼介於 A~L 之間，客戶國籍應該為 011 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          ElseIf Mid(UCase(LTrim(txtPCU(3))), 1, 1) >= "M" And Mid(UCase(LTrim(txtPCU(3))), 1, 1) <= "O" Then
'               If Trim(txtPCU(9)) <> "0111" Then
'                  ShowMsg "客戶英文名稱第一碼介於 M~O 之間，客戶國籍應該為 0111 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          ElseIf Mid(UCase(LTrim(txtPCU(3))), 1, 1) >= "P" And Mid(UCase(LTrim(txtPCU(3))), 1, 1) <= "Z" Then
'               If Trim(txtPCU(9)) <> "0112" Then
'                  ShowMsg "客戶英文名稱第一碼介於 P~Z 之間，客戶國籍應該為 0112 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          '2008/1/9 modify by sonia
'          'ElseIf Trim(txtPCU(3)) = "" Then
'          Else
'               If Trim(txtPCU(9)) <> "0112" Then
'                  ShowMsg "客戶英文名稱第一碼非英文字母或無英文名稱，客戶國籍應該為 0112 !"
'                  If Me.ActiveControl = txtPCU(9) Then
'                     txtPCU_GotFocus 9
'                  Else
'                     txtPCU(9).SetFocus
'                  End If
'                  tabCustomer.Tab = 0
'                  Exit Function
'               End If
'          End If
'      End If
      If Trim(txtPCU(9)) <> pub_NationByName(txtPCU(3) & txtPCU(4) & txtPCU(5) & txtPCU(6), Trim(txtPCU(9)), True, "客戶") Then
          If Me.ActiveControl = txtPCU(9) Then
             txtPCU_GotFocus 9
          Else
             txtPCU(9).SetFocus
          End If
          tabCustomer.Tab = 0
          Exit Function
      End If

      m_bSaveCheck = False
      
      If cboCity.Text = "" Then
         ShowMsg "客戶城市不可為空白 !"
         cboCity.SetFocus
         tabCustomer.Tab = 0
         Exit Function
      ElseIf txtPCU(10) = "" Then
         cboCity_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      
      'Modify By Sindy 2021/6/28
      'If txtPCU(11).Text = "" Then
      If cboPCU11.Text = "" Then
         ShowMsg "類別不可為空白 !"
         cboPCU11.SetFocus
         tabCustomer.Tab = 0
         Exit Function
      End If
      
      If txtPCU(20).Text = "" And txtPCU(26).Text = "" And txtPCU(27).Text = "" Then
         ShowMsg "中、英、日文地址不可同時為空白 !"
         txtPCU(20).SetFocus
         tabCustomer.Tab = 1
         Exit Function
      End If
      
      'Remove by Morgan 2007/12/19 改沒有時預設客戶國籍
      'If txtPCU(20).Text <> "" And txtPCU(28).Text = "" Then
      '   ShowMsg "英文地址不為空白時，地址國籍不可為空白 !"
      '   txtPCU(20).SetFocus
      '   tabCustomer.Tab = 1
      '   Exit Function
      'End If
            
      If m_EditMode = 1 And txtPCU(1) <> "" Then
         strExc(0) = "select count(*) from potcustomer where PCU01='" & Left(txtPCU(1) & "000", 8) & "' and PCU02='" & Left(txtPCU(2) & "0", 1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp(0) > 0 Then
               ShowMsg "客戶編號重覆，請重新輸入 !"
               txtPCU(1).SetFocus
               txtPCU_GotFocus 1
               tabCustomer.Tab = 0
               Exit Function
            End If
         End If
      End If
      
      If txtPCU(37).Text = "" Then
         ShowMsg "開發日期不可空白 !"
         txtPCU(37).SetFocus
         tabCustomer.Tab = 0
         Exit Function
      End If
      
      If lstUsers(0).ListCount = 0 Then
         ShowMsg "開發人員不可空白!"
         txtUserNo(0).SetFocus
         txtUserNo_GotFocus 0
         tabCustomer.Tab = 0
         Exit Function
      End If
      
      If DupeCustCheck(False) = True Then
         Exit Function
      End If
      
      'Add By Sindy 2014/7/10
      '新增存檔時,若智權人員為81040的資料,則'是否寄電子報'及'是否寄專利雙週報'都不可設為要寄
      If m_EditMode = 1 Then '新增
         For idx = 0 To lstUsers(0).ListCount - 1
            'Modified by Morgan 2022/1/7
            'If lstUsers(0).ITEMDATA(idx) = PUB_Id2Num("81040") Then '閻副所長
            If PUB_GetItemData(lstUsers(0).Tag, idx) = "81040" Then
            'end 2022/1/7
               If txtPCU(35) = "" Or txtPCU(48) = "" Then
                  'modify by sonia 2019/3/19
                  'ShowMsg "若要寄發則先存檔再修改！"
                  ShowMsg "新增時,若開發人員為閻副所長,電子報及專利雙週報都不可以設為要寄, 若要寄發則先存檔再修改！"
                  If txtPCU(35) = "" Then
                     txtPCU(35).SetFocus
                     txtPCU_GotFocus 35
                  ElseIf txtPCU(48) = "" Then
                     txtPCU(48).SetFocus
                     txtPCU_GotFocus 48
                  End If
                  Exit Function
               End If
            End If
         Next idx
      End If
      '2014/7/10 END
   End If
   
   For Each oText In txtPCU
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         Cancel = False
         'Add By Sindy 2012/4/9 +if
         If idx > 8 Or idx < 3 Then
         '2012/4/9 End
            txtPCU_Validate idx, Cancel
         End If
         If Cancel = True Then
            txtPCU(idx).SetFocus
            txtPCU_GotFocus idx
            Select Case idx
               Case 3 To 12, 34 To 40
                  tabCustomer.Tab = 0
               Case 13 To 33
                  tabCustomer.Tab = 1
            End Select
            Exit Function
         End If
      End If
   Next
      
   'Added by Morgan 2023/7/6
   If m_EditMode = "1" Or m_EditMode = "2" Then
      If txtPCU(36) = "" Then
         ShowMsg "定稿語文不可為空白 !"
         tabCustomer.Tab = 0
         txtPCU(36).SetFocus
         Exit Function
      Else
         If Left(txtPCU(9), 3) = "011" And txtPCU(36) <> "3" Then
             If MsgBox("國籍為「日本」，定稿語文確定「不是」日文？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                 tabCustomer.Tab = 0
                 txtPCU(36).SetFocus
                 Exit Function
             End If
         ElseIf Left(txtPCU(9), 3) <> "011" And txtPCU(36) = "3" Then
             If MsgBox("國籍為「不是」日本，定稿語文確定是「日文」？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                 tabCustomer.Tab = 0
                 txtPCU(36).SetFocus
                 Exit Function
             End If
         End If
      End If
   End If
   'end 2023/7/6
   'Add by Amy 2024/11/29 來所原因
   stMsg = ChkXYSourceReason(0, Me.Name, m_EditMode, cboSource, txtXYS02, _
                           , txtPCU(37), m_FieldList(41).fiOldData, txtXYS03, txtPCU(1))
   If stMsg <> MsgText(601) Then
      MsgBox stMsg, vbInformation
      tabCustomer.Tab = 0
      '來所原因 不可為空->20241114 Widen:允許可不輸,程式先寫,避免之後又改要輸,只需改共用
      If InStr(stMsg, "來所原因 不可為空") > 0 Then
         cboSource.SetFocus
      ElseIf InStr(stMsg, "介紹者編號") > 0 Then
         txtXYS02.SetFocus
      ElseIf InStr(stMsg, "其他說明") > 0 Then
         txtXYS03.SetFocus
      End If
      Exit Function
   End If
   stMsg = ChkXYSourceReason(2, Me.Name, m_EditMode, cboSource, txtXYS02, m_FieldList(54).fiOldData)
   If stMsg <> MsgText(601) Then
      If MsgBox(stMsg, vbYesNo + vbCritical) = vbNo Then
         tabCustomer.Tab = 0
         txtXYS02.SetFocus
         Exit Function
      End If
   End If
   stMsg = ""
   'end 2024/11/29
   
   TxtValidate = True
   
   'Add by Morgan 2007/12/19
   '地址國籍空白時設定為客戶國籍
   If txtPCU(28) = "" Then
      txtPCU(28) = Left(txtPCU(9), 3)
      lbl1(2).Caption = lbl1(1)
   End If
   'end 2007/12/19

   '整理英文地址
   For ii = 20 To 24
      If txtPCU(ii) = "" Then
         For jj = ii + 1 To 25
            If txtPCU(jj) <> "" Then
               txtPCU(ii) = txtPCU(jj)
               txtPCU(jj) = ""
               Exit For
            End If
         Next
      End If
   Next
   
End Function

Private Sub UpdateFieldNewData()
   For Each oText In txtPCU
      idx = oText.Index
      Select Case idx
         Case 12, 37
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
   'Add by Amy 2024/11/29 來所原因
   m_FieldList(54).fiNewData = ""
   If cboSource.Text <> MsgText(601) Then
      m_FieldList(54).fiNewData = Left(cboSource.Text, 2)
   End If
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
   Dim stSQL As String, stCols As String, stValues As String
   Dim intR As Integer 'Added by Lydia 2024/01/05
   
   'Move by Lydia 2024/01/05 從cnnConnection.BeginTrans下面搬上來
   If txtPCU(1) = "" Then
JumpToReNo: 'Added by Lydia 2024/01/05
      If ClsPDGetAutoNumber("R", strTmp, True, False) Then
         strTmp = "R" + Right(strTmp, 5) & "00"
         'Added by Lydia 2024/01/05 防止編號重覆
         stSQL = "select pcu01 from potcustomer where pcu01='" & strTmp & "' " & _
                 "union all select poc01 from potcustomer1 where poc01='" & strTmp & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            intR = intR + 1
            If intR > 3 Then
               MsgBox "存檔失敗!", vbCritical, "流水號給號"
               Exit Function
            End If
            GoTo JumpToReNo
         End If
         txtPCU(1) = strTmp
         'end 2024/01/05
         m_FieldList(1).fiNewData = strTmp
         m_FieldList(2).fiNewData = "0"
      End If
   End If
   'end --- Move by Lydia 2024/01/05 從cnnConnection.BeginTrans下面搬上來
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans

   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtPCU
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
   'Add By Sindy 2021/6/28
   stCols = stCols & ",PCU11"
   stValues = stValues & ",'" & Left(cboPCU11.Text, 1) & "'"
   '2021/6/28 END
   'Add by Amy 2024/11/29
   stCols = stCols & ",PCU55"
   stValues = stValues & ",'" & Left(cboSource.Text, 2) & "'"
   'end 2024/11/29
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO PotCustomer (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   'Add by Amy 2024/11/29 +客戶代理人來源資料檔
   stMsg = SaveXYNoSource(1, Me.Name, txtPCU(1), txtXYS02, txtXYS03, Left(cboSource, 2))
   If Len(stMsg) > 1 Then
      GoTo ErrHand
   End If
   stMsg = ""
   
   '新增聯絡人資料
   With rsContact
   If .RecordCount > 0 Then
      .MoveFirst
      .Sort = "PCC02 asc"
      Do While Not .EOF
         stCols = "PCC01"
         stValues = "'" & m_FieldList(1).fiNewData & "'"
         For idx = 1 To .Fields.Count - 3
            If Not IsNull(.Fields(idx)) Then
               stCols = stCols & "," & .Fields(idx).Name
               If .Fields(idx).Name = "PCC11" Then
                  stValues = stValues & "," & .Fields(idx)
               Else
                  stValues = stValues & ",'" & ChgSQL(.Fields(idx)) & "'"
               End If
            End If
         Next
         stSQL = "INSERT INTO PotCustCont (" & stCols & ") Values (" & stValues & ")"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL
         .MoveNext
      Loop
   End If
   End With
   '更新聯絡人相關編號
   UpdateRefContact m_FieldList(1).fiNewData
   cnnConnection.CommitTrans
   AddRecord = True
   
   txtPCU(1) = m_FieldList(1).fiNewData
   txtPCU(2) = m_FieldList(2).fiNewData
   
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    'Modify by Amy 2024/11/29 SaveXYNoSource有誤回傳其錯誤
    If stMsg = MsgText(601) Then
      stMsg = Err.Description
    End If
    MsgBox stMsg, vbCritical
    'end 2024/11/29
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '清除相關聯絡人資料
   stSQL = "update potcustcont set pcc20=null where substr(pcc20,1,8)='" & txtPCU(1) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   '刪除潛在聯絡人資料
   stSQL = "delete from PotCustCont where pcc01='" & txtPCU(1) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   '刪除潛在客戶資料
   stSQL = "delete from PotCustomer where pcu01='" & txtPCU(1) & "' and pcu02='" & txtPCU(2) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   'Add by Amy 2024/11/29 刪除 客戶代理人來源資料檔 的 被介紹 資料(log 記錄寫於SaveXYNoSource)
   stMsg = SaveXYNoSource(3, Me.Name, txtPCU(1))
   If Len(stMsg) > 1 Then
      GoTo ErrHand
   End If
   stMsg = ""
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtPCU(1).Tag = ""
   txtPCU(2).Tag = ""
   
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    'Modify by Amy 2024/11/29 SaveXYNoSource有誤回傳其錯誤
    If stMsg = MsgText(601) Then
      stMsg = Err.Description
    End If
    MsgBox stMsg, vbCritical
    'end 2024/11/29
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
  
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE PotCustomer SET "
   stSet = ""
   For Each oText In txtPCU
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
   'Add By Sindy 2021/6/28
   idx = 11
   If Left(cboPCU11.Text, 1) <> m_FieldList(idx).fiOldData Then
      bDifference = True
      '文字
      If m_FieldList(idx).fiType = 0 Then
         stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(Left(cboPCU11.Text, 1)))
      '數字
      Else
         stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(Left(cboPCU11.Text, 1), True)
      End If
   End If
   '2021/6/28 END
   'Add by Amy 2024/11/29 +來所原因
   idx = 54
   If Left(cboSource.Text, 2) <> m_FieldList(idx).fiOldData Then
      bDifference = True
      '文字
      stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(Left(cboSource.Text, 2)))
   End If
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where pcu01='" & txtPCU(1) & "' and pcu02='" & txtPCU(2) & "'; end; "
     
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   
   'Add by Amy 2024/11/29 +客戶代理人來源資料檔
   stMsg = SaveXYNoSource(2, Me.Name, txtPCU(1), txtXYS02, txtXYS03, Left(cboSource, 2), m_FieldList(54).fiOldData)
   If Len(stMsg) > 1 Then
      GoTo ErrHand
   End If
   stMsg = ""
      
   '更新聯絡人資料
   If rsContact.RecordCount = 0 Then
      If rsContactOld.RecordCount > 0 Then
         '清除相關聯絡人資料
         stSQL = "update potcustcont set pcc20=null where substr(pcc20,1,8)='" & txtPCU(1) & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
         'Added by Lydia 2024/05/14 刪除聯絡人相片
         strExc(0) = "select imgbytefile.* from potcustcont,imgbytefile where pcc01='" & txtPCU(1) & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               PUB_DelFtpFile2 RsTemp.Fields("IBF01") & "-" & RsTemp.Fields("IBF02") & "-" & RsTemp.Fields("IBF03") & "-" & RsTemp.Fields("IBF04") & "-" & RsTemp.Fields("IBF05"), , UCase("ImgByteFile")
               stSQL = "DELETE FROM IMGBYTEFILE WHERE IBF01='" & RsTemp.Fields("IBF01") & "' AND IBF02='" & RsTemp.Fields("IBF02") & "' AND IBF03='" & RsTemp.Fields("IBF03") & "' AND IBF04='" & RsTemp.Fields("IBF04") & "' AND IBF05='" & RsTemp.Fields("IBF05") & "' "
               cnnConnection.Execute stSQL
               RsTemp.MoveNext
            Loop
         End If
         'end 2024/05/14
         '刪除聯絡人資料
         stSQL = "delete from potcustcont where pcc01='" & txtPCU(1) & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
   Else
      '刪除聯絡人(原來的編號在新的聯絡人資料中找不到的)
      With rsContactOld
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            rsContact.MoveFirst
            rsContact.Find "PCC02='" & .Fields("PCC02") & "'"
            If rsContact.EOF Then
               '清除相關聯絡人資料
               stSQL = "update potcustcont set pcc20=null where pcc20='" & txtPCU(1) & .Fields("PCC02") & "'"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL, intI
               'Added by Lydia 2024/05/14 刪除聯絡人相片
               strExc(0) = "select imgbytefile.* from potcustcont,imgbytefile where pcc01='" & txtPCU(1) & "' and pcc02='" & .Fields("PCC02") & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     PUB_DelFtpFile2 RsTemp.Fields("IBF01") & "-" & RsTemp.Fields("IBF02") & "-" & RsTemp.Fields("IBF03") & "-" & RsTemp.Fields("IBF04") & "-" & RsTemp.Fields("IBF05"), , UCase("ImgByteFile")
                     stSQL = "DELETE FROM IMGBYTEFILE WHERE IBF01='" & RsTemp.Fields("IBF01") & "' AND IBF02='" & RsTemp.Fields("IBF02") & "' AND IBF03='" & RsTemp.Fields("IBF03") & "' AND IBF04='" & RsTemp.Fields("IBF04") & "' AND IBF05='" & RsTemp.Fields("IBF05") & "' "
                     cnnConnection.Execute stSQL
                     RsTemp.MoveNext
                  Loop
               End If
               'end 2024/05/14
               '刪除聯絡人資料
               stSQL = "delete from potcustcont where pcc01='" & txtPCU(1) & "' and pcc02='" & .Fields("PCC02") & "'"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL, intI
            End If
            .MoveNext
         Loop
      End If
      End With
      '更新(新增)連絡人
      With rsContact
      .MoveFirst
      Do While Not .EOF
         If rsContactOld.RecordCount = 0 Then
            bAddNew = True
         Else
            rsContactOld.MoveFirst
            rsContactOld.Find "PCC02='" & .Fields("PCC02") & "'"
            If rsContactOld.EOF Then
               bAddNew = True
            Else
               bAddNew = False
            End If
         End If
         
         '新增
         If bAddNew = True Then
            stCols = "PCC01"
            stValues = "'" & m_FieldList(1).fiNewData & "'"
            For idx = 1 To .Fields.Count - 3
               If .Fields(idx) <> "" Then
                  stCols = stCols & "," & .Fields(idx).Name
                  If .Fields(idx).Name = "PCC11" Then
                     stValues = stValues & "," & .Fields(idx)
                  Else
                     stValues = stValues & ",'" & ChgSQL(.Fields(idx)) & "'"
                  End If
               End If
            Next
            stSQL = "INSERT INTO PotCustCont (" & stCols & ") Values (" & stValues & ")"
            Pub_SeekTbLog stSQL
            cnnConnection.Execute stSQL, intI
         '修改
         Else
            bDifference = False
            stSet = ""
            For idx = 2 To .Fields.Count - 3
               If "" & .Fields(idx) <> "" & rsContactOld.Fields(idx) Then
                  bDifference = True
                  If .Fields(idx).Name = "PCC11" Then
                     stSet = stSet & "," & .Fields(idx).Name & "=" & CNULL(.Fields(idx), True)
                  Else
                     stSet = stSet & "," & .Fields(idx).Name & "=" & CNULL(ChgSQL(.Fields(idx)))
                  End If
               End If
            Next
            If bDifference = True Then
               stSet = Mid(stSet, 2)
               stSQL = "begin user_data.user_enabled:=1; Update PotCustCont set " & stSet & " where PCC01='" & m_FieldList(1).fiNewData & "' and PCC02='" & .Fields("PCC02") & "'; end;"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL
            End If
         End If
         .MoveNext
      Loop
      End With
   End If
   '更新聯絡人相關編號
   UpdateRefContact m_FieldList(1).fiNewData
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    'Modify by Amy 2024/11/29 SaveXYNoSource有誤回傳其錯誤
    If stMsg = MsgText(601) Then
      stMsg = Err.Description
    End If
    MsgBox stMsg, vbCritical

End Function

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 3: '刪除
         'Modify by Morgan 2008/1/24 +刪除檢查
         If PUB_PCCDelCheck(txtPCU(1), txtPCU(2)) = True Then
            If DelRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord 2
            End If
         End If
      
       Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtPCU(1).SetFocus
               txtPCU_GotFocus 1
            End If
         End If
         
   End Select
End Function

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
'Modify by Amy 2024/01/22 +IsTranBack
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0, Optional ByVal IsTranBack As Boolean = False) As Boolean
   
   Dim stPCU01 As String
   Dim stPCU02 As String
   Dim adoRst As New ADODB.Recordset
   Dim RsQ As New ADODB.Recordset 'Add by Amy 2024/01/22
   
   stPCU01 = Left(txtPCU(1) & "000", 8)
   stPCU02 = Left(txtPCU(2) & "0", 1)

   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM PotCustomer" & _
            " WHERE PCU01 = '" & stPCU01 & "' AND PCU02 = '" & stPCU02 & "'"
      Case -2
         strExc(0) = "SELECT * FROM PotCustomer order by PCU01 ASC,PCU02 ASC"
      Case -1
         strExc(0) = "SELECT * FROM PotCustomer" & _
            " WHERE PCU01||PCU02 <'" & stPCU01 & stPCU02 & "' order by PCU01 DESC,PCU02 DESC"
      Case 1
         strExc(0) = "SELECT * FROM PotCustomer" & _
            " WHERE PCU01||PCU02 >'" & stPCU01 & stPCU02 & "' order by PCU01 ASC,PCU02 ASC"
      Case 2
         strExc(0) = "SELECT * FROM PotCustomer order by PCU01 DESC,PCU02 DESC"
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      'Add By Sindy 2009/06/23
      If Not IsNull(adoRst.Fields("PCU47")) Then
         Call GetCustData(adoRst.Fields("PCU47"), adoRst.Fields("PCU11"))
      End If
      '2009/06/23 End
      UpdateCtrlData adoRst
      OpenContactTable
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         'Modify by Amy 2024/01/22 +if 代理人/客戶檔返回者,若轉的資料為最後一筆,要重抓最後一筆資料(否則顯示的會是已轉代理人or客戶的那筆資料)
         If IsTranBack = True Then
            strExc(0) = "SELECT * FROM PotCustomer order by PCU01 DESC,PCU02 DESC"
            intI = 1
            RsQ.MaxRecords = 1
            Set RsQ = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               ClearField
               If Not IsNull(RsQ.Fields("PCU47")) Then
                  Call GetCustData(RsQ.Fields("PCU47"), RsQ.Fields("PCU11"))
               End If
               UpdateCtrlData RsQ
               OpenContactTable
               Set RsQ = Nothing
               ShowRecord = True
            End If
         Else
            MsgBox "已經是最後筆！", vbInformation
         End If
         'end 2024/01/22
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtPCU(1).SetFocus
      txtPCU_GotFocus 1
   End If
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   Dim strTp(3) As String 'Add by Amy 2024/11/29
   
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtPCU
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
         
         'Add by Sindy 2021/6/28
         Call PUB_SetComboPCU11(cboPCU11, "" & .Fields("PCU11")) '類別
         m_FieldList(11).fiOldData = "" & .Fields("PCU11")
         m_FieldList(11).fiNewData = m_FieldList(11).fiOldData
         '2021/6/28 END
         'Add by Amy 2024/11/29 來所原因 相關
         m_FieldList(54).fiOldData = "" & .Fields("PCU55")
         m_FieldList(54).fiNewData = m_FieldList(54).fiOldData
         If IsNull(.Fields("PCU55")) Then
            cboSource.ListIndex = -1
         ElseIf "" & .Fields("PCU55") < 10 Then
            cboSource.ListIndex = .Fields("PCU55")
         'Memo by Amy 不會有PCU55=10,超過11以上用cboSource.ListIndex 會錯
         Else
            strExc(9) = "And AC02='" & .Fields("PCU55") & "'"
            Call Pub_SetCboComeSource(1, Me.Name, , strExc(9))
            cboSource = strExc(9)
         End If
         Call Pub_GetXYSource(1, txtPCU(1), strTp(0), strTp(1), strTp(2))
         txtXYS02 = strTp(0)
         LblSourceN.Caption = strTp(1)
         txtXYS03.Text = strTp(2)
         txtXYS02.Tag = txtXYS02
         txtXYS03.Tag = txtXYS03
         cmdIntroduce.BackColor = &H8000000F
         If Pub_GetXYSource(2, txtPCU(1)) = True Then
            cmdIntroduce.BackColor = &HFFFF80
         End If
         'end 2024/11/29
         
         m_PCU41 = "" & .Fields("PCU41") 'Add By Sindy 2009/04/30
         CUID(1) = "" & .Fields("PCU41")
         CUID(2) = "" & .Fields("PCU42")
         CUID(3) = "" & .Fields("PCU43")
         CUID(4) = "" & .Fields("PCU44")
         CUID(5) = "" & .Fields("PCU45")
         CUID(6) = "" & .Fields("PCU46")
         SetCity
         SetlstUsers 0, txtPCU(38)
         'Add by Morgan 2009/2/6
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
         
         Call Pub_ShowSelectList(Combo1, List1, txtPCU(49).Text) 'Added by Lydia 2016/11/29
         lblPCU47.Caption = txtPCU(47).Text 'Added by Lydia 2020/05/07 關聯企業：暫時帶入
         
         'Add By Sindy 2019/7/24 國內外權限
         If "" & .Fields("PCU51") = "C" Then
            Option1(0).Value = True
         ElseIf "" & .Fields("PCU51") = "F" Then
            Option1(1).Value = True
         End If
         '2019/7/24 END
      End If
   End With
   UpdateCUID CUID, textCUID
   txtPCU(1).Tag = m_FieldList(1).fiOldData
   txtPCU(2).Tag = m_FieldList(2).fiOldData
   m_sDupeKey = txtPCU(3) & txtPCU(4) & txtPCU(5) & txtPCU(6) & txtPCU(9) & txtPCU(10)
   'Add By Sindy 2014/1/27
   m_sDupeKey_c = txtPCU(8) & txtPCU(9)
   m_sDupeKey_j = txtPCU(7) & txtPCU(9)
   '2014/1/27 END
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   For Each oText In txtPCU
      idx = oText.Index
      m_FieldList(idx).fiName = "PCU" & Format(idx, "00")
   Next
   m_FieldList(11).fiName = "PCU" & Format(11, "00") 'Add By Sindy 2021/6/28
   m_FieldList(54).fiName = "PCU" & Format(55, "00") 'Add by Amy 2024/11/29
End Sub

' 更新 Create 及 Update 的人
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
   'Modified by Lydia 2024/05/14 String(10->String(6
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(6, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub SetCity()
   cboCity.Clear
   If txtPCU(9) <> "" Then
      strExc(0) = "select ct03,ct02 from city where ct01='" & Left(txtPCU(9), 3) & "' order by ct03 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            cboCity.AddItem .Fields("ct03"), 0
            cboCity.ItemData(0) = .Fields("ct02")
            .MoveNext
         Loop
         End With
      End If
   End If
   ShowCity
End Sub

Private Sub ShowCity()
   If txtPCU(10) <> "" And cboCity.ListCount > 0 Then
      For intI = 0 To cboCity.ListCount - 1
         If cboCity.ItemData(intI) = Val(txtPCU(10)) Then
            cboCity.ListIndex = intI
            Exit For
         End If
      Next
   End If
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = "" 'Added by Morgan 2022/1/7
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
                  'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
                  'Modified by Morgan 2022/1/7
                  'lstUsers(p_idx).ITEMDATA(0) = PUB_Id2Num(.Fields(0)) '員工編號
                  lstUsers(p_idx).Tag = .Fields(0) & "," & lstUsers(p_idx).Tag
                  'end 2022/1/7
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub

Private Sub SetList(oList As ListBox, p_stData As String)
   Dim arrID
   oList.Clear
   If p_stData <> "" Then
      arrID = Split(p_stData, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

Private Sub AddlstUsers(p_idx As Integer)
   Dim idx As Integer, bFound As Boolean
   
   If txtUserNo(p_idx) <> "" And lblName(p_idx) <> "" Then
      'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
      'Modified by Morgan 2022/1/7
      'For idx = 0 To lstUsers(p_idx).ListCount - 1
      '   If lstUsers(p_idx).ITEMDATA(idx) = PUB_Id2Num(txtUserNo(p_idx)) Then
      '      MsgBox "員工已存在於開發人員清單中！"
      '      txtUserNo(p_idx).SetFocus
      '      txtUserNo_GotFocus p_idx
      '      bFound = True
      '      Exit For
      '   End If
      'Next
      If InStr(lstUsers(p_idx).Tag, txtUserNo(p_idx)) > 0 Then
         MsgBox "員工已存在於開發人員清單中！"
         txtUserNo(p_idx).SetFocus
         txtUserNo_GotFocus p_idx
         bFound = True
      End If
      'end 2022/1/7
      If bFound = False Then
         lstUsers(p_idx).AddItem lblName(p_idx), 0
         'Modified by Morgan 2022/1/7
         'lstUsers(p_idx).ITEMDATA(0) = PUB_Id2Num(txtUserNo(p_idx))
         lstUsers(p_idx).Tag = txtUserNo(p_idx) & "," & lstUsers(p_idx).Tag
         'end 2022/1/7
         txtUserNo(p_idx) = ""
         lblName(p_idx) = ""
      End If
   End If
End Sub

Private Sub RemovelstUsers(p_idx As Integer)
   'Modified by Morgan 2022/1/7
   'Dim idx As Integer, ii As Integer
   'If lstUsers(p_idx).ListCount > 0 Then
   '   ii = 0
   '   For idx = 0 To lstUsers(p_idx).ListCount - 1
   '      If lstUsers(p_idx).Selected(ii) = True Then
   '         lstUsers(p_idx).RemoveItem ii
   '         ii = ii - 1
   '      End If
   '      ii = ii + 1
   '   Next
   'End If
   lstUsers(p_idx).Tag = PUB_RemoveListBox2(lstUsers(p_idx), lstUsers(p_idx).Tag)
   'end 2022/1/7
End Sub

Private Sub AddLstFrmCbo(oCombo As ComboBox, oList As ListBox)
   Dim idx As Integer, bFound As Boolean
   
   If oCombo <> "" Then
      For idx = 0 To oList.ListCount - 1
         If oList.List(idx) = oCombo Then
            MsgBox "資料已存在！"
            oCombo.SetFocus
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem oCombo, 0
         oCombo = ""
      End If
   End If
End Sub

Private Sub RemoveList(oList As ListBox)
   Dim idx As Integer, ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      For idx = 0 To oList.ListCount - 1
         If oList.Selected(ii) = True Then
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Next
   End If
End Sub

Private Sub txtUserNo_Change(Index As Integer)
   Dim strTempName As String
   If Len(txtUserNo(Index)) = 5 Then
      If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
         lblName(Index) = strTempName
      End If
   Else
      lblName(Index) = ""
   End If
End Sub

Private Sub txtUserNo_GotFocus(Index As Integer)
   TextInverse txtUserNo(Index)
End Sub

'Add By Sindy 2010/11/26
Private Sub txtUserNo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_Validate(Index As Integer, Cancel As Boolean)
   Dim strTempName As String
   If txtUserNo(Index).Visible = True Then
      If txtUserNo(Index) <> "" And lblName(Index) = "" Then
         If Len(txtUserNo(Index)) = 5 Then
            If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
               lblName(Index) = strTempName
            End If
         End If
         If lblName(Index) = "" Then
            MsgBox "員工編號輸入錯誤！", vbExclamation
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub OpenContactTable()
   
On Error GoTo Checking
   
   If txtPCU(1) <> "" Then
      strExc(0) = "select PCC.*,decode(pcc20,null,'　','＊')||pcc02 X1,decode(pcc20,null,'',substr(pcc20,1,8)||'-'||substr(pcc20,9)) X2 from PotCustCont PCC where pcc01='" & txtPCU(1) & "' order by pcc02"
   Else
      strExc(0) = "select PCC.*,decode(pcc20,null,'　','＊')||pcc02 X1,decode(pcc20,null,'',substr(pcc20,1,8)||'-'||substr(pcc20,9)) X2 from PotCustCont PCC where rownum<1"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/17 +FormName 改暫存TB
   Set rsContact = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set rsContactOld = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   'end 2014/06/17
   Set Adodc1.Recordset = rsContact
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   ReDim m_arrConRefList(1, 0)
   If rsContact.RecordCount > 0 Then
      ReadContact
   End If
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub

Private Function getNewNo() As String
   Dim myTemp As ADODB.Recordset
   Dim iUsableNo As Integer
   
   Set myTemp = rsContact.Clone
   With myTemp
      .Sort = "PCC02 asc"
      iUsableNo = 1
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            If iUsableNo = Val("" & .Fields(1)) Then
               iUsableNo = iUsableNo + 1
            Else
               Exit Do
            End If
            .MoveNext
         Loop
      End If
      getNewNo = Format(iUsableNo, "00")
   End With
   Set myTemp = Nothing
End Function

Private Sub UpdateContact()
   With rsContact
   If txtPCC(2) = "" Then
      m_iConEditMode = 1
      txtPCC(2) = getNewNo
      .AddNew
   Else
      If .RecordCount > 0 Then
         .MoveFirst
         .Find "PCC02='" & txtPCC(2) & "'"
         If .EOF Then
            .AddNew
         End If
      Else
         .AddNew
      End If
   End If
   For Each oText In txtPCC
      If oText.Index = 11 Then
         .Fields("PCC" & Format(oText.Index, "00")) = DBDATE(oText.Text)
      Else
         .Fields("PCC" & Format(oText.Index, "00")) = oText.Text
      End If
   Next
   .Fields("X2") = txtPCC20
   If txtPCC20 <> "" Then
      .Fields("X1") = "＊" & .Fields("PCC02")
   Else
      .Fields("X1") = "　" & .Fields("PCC02")
   End If
   .UPDATE
   End With
End Sub

Private Sub ReadContact()
   Dim CUID(1 To 6) As String
   ClearField1
   With rsContact
      If Not (.EOF Or .BOF) Then
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
         If Not IsNull(.Fields("PCC15")) Then
            m_iConEditMode = 2
         End If
         If txtPCC(6) <> "" Then
            SetList lstDept, txtPCC(6)
         End If
         If txtPCC(7) <> "" Then
            SetList lstTitle, txtPCC(7)
         End If
         If txtPCC(12) <> "" Then
            SetlstUsers 1, txtPCC(12)
         End If
         UpdateCUID CUID, textCUID1
         
      End If
   End With
   
   
   'Added by Lydia 2024/05/14 聯絡人相片
   If Trim(txtPCU(1)) <> "" And Trim(txtPCC(2)) <> "" Then
      Command1.Visible = True
      Call Pub_GetPCCtoIBF_2(Trim(txtPCU(1)), Trim(txtPCC(2)), Command1)
   Else
      Command1.Visible = False
   End If
   'end 2024/05/14
   
End Sub

'聯絡人檢查
Private Function TxtValidate1() As Boolean
   Dim Cancel As Boolean
   
   'Added by Morgan 2022/1/5 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2022/1/5
   
   For Each oText In txtPCC
   
      If oText.Locked = False Then
         idx = oText.Index
         Cancel = False
         txtPCC_Validate idx, Cancel
         If Cancel = True Then
            txtPCC_GotFocus idx
            Exit Function
         End If
      End If
   Next
   
   If txtPCC(3) = "" And txtPCC(4) = "" And txtPCC(5) = "" Then
      ShowMsg "聯絡人中、英、日文名稱不可同時為空白 !"
      txtPCC(3).SetFocus
      Exit Function
   End If
   
   If txtPCC(11).Text = "" Then
      ShowMsg "開發日期不可空白 !"
      txtPCC(11).SetFocus
      Exit Function
   End If
      
   If lstUsers(1).ListCount = 0 Then
      ShowMsg "開發人員不可空白!"
      txtUserNo(1).SetFocus
      txtUserNo_GotFocus 1
      Exit Function
   Else
      'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
      'Modified by Morgan 2022/1/7
      'strExc(1) = ""
      'strExc(1) = PUB_Num2Id(lstUsers(1).ITEMDATA(0))
      'For intI = 1 To lstUsers(1).ListCount - 1
      '   strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(1).ITEMDATA(intI))
      'Next
      'txtPCC(12).Text = strExc(1)
      'Added by Lydia 2022/04/18 去掉多餘的,
      If Right(lstUsers(1).Tag, 1) = "," Then
          txtPCC(12).Text = Mid(lstUsers(1).Tag, 1, Len(lstUsers(1).Tag) - 1)
      Else
      'end 2022/04/18
          txtPCC(12).Text = lstUsers(1).Tag
          'end 2022/1/7
      End If 'Added by Lydia 2022/04/18
   End If
   
   'Add by Amy 2021/11/29
   If txtPCU(39) = "國內同業" Then
        '不可寄mail
        'Modify by Amy 2023/02/01 bug-抓錯欄位 原:txtPCU(8)
        If txtPCC(8) <> MsgText(601) Then
            ShowMsg "此為國內同業,不可輸入E-Mail以免誤發電子郵件, 如有需要請加註於備註欄 ！"
            txtPCC(8).SetFocus
            txtPCC_GotFocus (8)
            tabCustomer.Tab = 2
            Exit Function
        End If
        '電子報要設定不寄
        If txtPCC(9) <> "N" Then
            ShowMsg "此為國內同業, 不可寄台一雜誌 ！"
            txtPCC(9).SetFocus
            txtPCC_GotFocus (9)
            tabCustomer.Tab = 2
            Exit Function
        End If
        If txtPCC(10) <> "N" Then
            ShowMsg "此為國內同業, 不可寄電子報 ！"
            txtPCC(10).SetFocus
            txtPCC_GotFocus (10)
            tabCustomer.Tab = 2
            Exit Function
        End If
        If txtPCC(24) <> "N" Then
            ShowMsg "此為國內同業, 不可寄專利雙週報！"
            txtPCC(24).SetFocus
            txtPCC_GotFocus (24)
            tabCustomer.Tab = 2
            Exit Function
        End If
    End If
    'end 2021/11/29
   
   '檢查聯絡人是否重複
   With rsContact
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         If .Fields("PCC02") <> txtPCC(2) And ( _
            (UCase("" & .Fields("PCC03")) = UCase(txtPCC(3)) And txtPCC(3) <> "") Or _
            (UCase("" & .Fields("PCC04")) = UCase(txtPCC(4)) And txtPCC(4) <> "") Or _
            (UCase("" & .Fields("PCC05")) = UCase(txtPCC(5)) And txtPCC(5) <> "")) Then
            If MsgBox("名稱與聯絡人[" & .Fields("PCC02") & "]重複！是否仍然要更新？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Exit Function
            End If
         End If
         .MoveNext
      Loop
   End If
   End With
   TxtValidate1 = True
End Function

'Modify By Sindy 2012/4/9 +bolMustChk, index
Private Function DupeCustCheck(bolMustChk As Boolean, Optional Index As Integer) As Boolean
    Dim strCheckWay As String, strFind As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String 'Add by Amy 2021/08/13
    Dim strNo As String, strTmp As String 'Add by Amy 2021/08/30
    Dim strRCLSql As String 'Add by Amy 2024/05/21 +風險檢查

    'Add by Amy 2021/08/16 +檢查對造,屬於對造[以等於比對,若xxx股份有限公司和xxx(股)公司為不同可存檔]不可存檔-秀玲
    If bolMustChk = True Then
        strSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
        strSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
        StrSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
        StrSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
        strSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
        strCheckWay = "="
        If Index = 8 Then
            strFind = UCase(txtPCU(8))
        ElseIf Index = 7 Then
            strExc(0) = UCase(txtPCU(7))
        ElseIf Index >= 3 And Index <= 6 Then
             strFind = txtPCU(3)
             If Trim(txtPCU(4)) <> MsgText(601) Then strFind = strFind & " " & txtPCU(4)
             If Trim(txtPCU(5)) <> MsgText(601) Then strFind = strFind & " " & txtPCU(5)
             If Trim(txtPCU(6)) <> MsgText(601) Then strFind = strFind & " " & txtPCU(6)
             strFind = UCase(strFind)
        End If
        Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, ChgSQL(strFind), strCheckWay, True)
        'Add by Amy 2024/05/21 +風險檢查(判斷同對造,用=)
        Call ChkRiskData(2, Me.Name, , , strFind, strRCLSql)
        If ShowData(strRCLSql) = True Then
        'end 2024/05/21
            Set frm210128_1.grdDataList.Recordset = RsQ
            Set frm210128_1.m_PrevForm = Me
            frm210128_1.Label1.Caption = "客戶名稱：" & strFind
            frm210128_1.Show vbModal
            DupeCustCheck = True
            Exit Function
        End If
    End If
    'end 2021/08/16
    
   '檢查英文名稱
   If txtPCU(3) <> "" And txtPCU(9) <> "" And txtPCU(10) <> "" And m_sDupeKey <> txtPCU(3) & txtPCU(4) & txtPCU(5) & txtPCU(6) & txtPCU(9) & txtPCU(10) Then
      'Modify by Morgan 2007/12/27 名稱的欄位數可能不同，要加空白比較
      'strExc(0) = "select pcu01||pcu02 from potcustomer where pcu09>='" & Left(txtPCU(9), 3) & "' and pcu09<='" & Left(txtPCU(9), 3) & "z' and upper(pcu03||pcu04||pcu05||pcu06)='" & ChgSQL(UCase(txtPCU(3)  & txtPCU(4)  & txtPCU(5) & txtPCU(6))) & "' and pcu10='" & txtPCU(10) & "' and rownum<2"
      '國外潛在客戶
       'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select pcu01||pcu02 from potcustomer where pcu09>='" & Left(txtPCU(9), 3) & "' and pcu09<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(pcu03||' '||pcu04||' '||pcu05||' '||pcu06))='" & ChgSQL(UCase(RTrim(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6)))) & "' and pcu10='" & txtPCU(10) & "' and rownum<2"
      strExc(0) = "select pcu01||pcu02 from potcustomer where pcu09>='" & Left(txtPCU(9), 3) & "' and pcu09<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(pcu03||' '||pcu04||' '||pcu05||' '||pcu06))=Rtrim('" & ChgSQL(UCase(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6))) & "') and pcu10='" & txtPCU(10) & "' and rownum<2"
      If txtPCU(1) <> "" Then
         strExc(0) = strExc(0) & " and pcu01||pcu02<>'" & Left(txtPCU(1) & txtPCU(2) & "000", 9) & "'"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與國外潛在客戶[" & RsTemp.Fields(0) & "]的英文名稱、國籍、城市皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      'Modify by Morgan 2007/12/27 名稱的欄位數可能不同，要加空白比較
      'strExc(0) = "select cu01||cu02 from customer where cu10>='" & Left(txtPCU(9), 3) & "' and cu10<='" & Left(txtPCU(9), 3) & "z' and upper(cu05||cu88||cu89||cu90)='" & ChgSQL(UCase(txtPCU(3) & txtPCU(4) & txtPCU(5) & txtPCU(6))) & "' and instr(upper(cu24||cu25||cu26||cu27||cu28||cu102),'" & UCase(cboCity) & "')>0  and rownum<2"
      '客戶
       'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select cu01||cu02 from customer where cu10>='" & Left(txtPCU(9), 3) & "' and cu10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(cu05||' '||cu88||' '||cu89||' '||cu90))='" & ChgSQL(UCase(RTrim(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6)))) & "' and instr(upper(cu24||cu25||cu26||cu27||cu28||cu102),'" & UCase(cboCity) & "')>0 and rownum<2"
      strExc(0) = "select cu01||cu02 from customer where cu10>='" & Left(txtPCU(9), 3) & "' and cu10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(cu05||' '||cu88||' '||cu89||' '||cu90))=Rtrim('" & ChgSQL(UCase(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6))) & "') and instr(upper(cu24||cu25||cu26||cu27||cu28||cu102),'" & UCase(cboCity) & "')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與客戶[" & RsTemp.Fields(0) & "]的英文名稱、國籍、城市皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      'Modify by Morgan 2007/12/27 名稱的欄位數可能不同，要加空白比較
      'strExc(0) = "select fa01||fa02 from fagent where fa10>='" & Left(txtPCU(9), 3) & "' and fa10<='" & Left(txtPCU(9), 3) & "z' and upper(fa05||fa63||fa64||fa65)='" & ChgSQL(UCase(txtPCU(3) & txtPCU(4) & txtPCU(5) & txtPCU(6))) & "' and instr(upper(fa18||fa19||fa20||fa21||fa22||fa70),'" & UCase(cboCity) & "')>0  and rownum<2"
      '代理人
       'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select fa01||fa02 from fagent where fa10>='" & Left(txtPCU(9), 3) & "' and fa10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(fa05||' '||fa63||' '||fa64||' '||fa65))='" & ChgSQL(UCase(RTrim(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6)))) & "' and instr(upper(fa18||fa19||fa20||fa21||fa22||fa70),'" & UCase(cboCity) & "')>0 and rownum<2"
      strExc(0) = "select fa01||fa02 from fagent where fa10>='" & Left(txtPCU(9), 3) & "' and fa10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(fa05||' '||fa63||' '||fa64||' '||fa65))=Rtrim('" & ChgSQL(UCase(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6))) & "') and instr(upper(fa18||fa19||fa20||fa21||fa22||fa70),'" & UCase(cboCity) & "')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與代理人[" & RsTemp.Fields(0) & "]的英文名稱、國籍、城市皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      'Add By Sindy 2014/1/27
      '國內潛在客戶
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select poc01||poc02 from potcustomer1 where poc04>='" & Left(txtPCU(9), 3) & "' and poc04<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(poc23||' '||poc24||' '||poc25||' '||poc26))='" & ChgSQL(UCase(RTrim(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6)))) & "' and rownum<2"
      strExc(0) = "select poc01||poc02 from potcustomer1 where poc04>='" & Left(txtPCU(9), 3) & "' and poc04<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(poc23||' '||poc24||' '||poc25||' '||poc26))=Rtrim('" & ChgSQL(UCase(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與國內潛在客戶[" & RsTemp.Fields(0) & "]的英文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '法務開拓
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select ecd02||'-'||LPAD(ecd01,6,'0') from expandcusdetail where ecd10>='" & Left(txtPCU(9), 3) & "' and ecd10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(ecd03||' '||ecd04))='" & ChgSQL(UCase(RTrim(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6)))) & "' and instr(upper(ecd05||ecd06||ecd07||ecd08||ecd09),'" & UCase(cboCity) & "')>0 and rownum<2"
      strExc(0) = "select ecd02||'-'||LPAD(ecd01,6,'0') from expandcusdetail where ecd10>='" & Left(txtPCU(9), 3) & "' and ecd10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(ecd03||' '||ecd04))=Rtrim('" & ChgSQL(UCase(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6))) & "') and instr(upper(ecd05||ecd06||ecd07||ecd08||ecd09),'" & UCase(cboCity) & "')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與法務開拓[" & RsTemp.Fields(0) & "]的英文名稱、國籍、城市皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '2014/1/27 END
      'Add by Amy 2021/08/30
      '國內開拓函特定公司不列印者
      If HasTMBulletinnp(RTrim(UCase(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6)))) = True Then
         If MsgBox("本客戶資料與[ 國內開拓函特定公司不列印者]名稱一樣" & _
          vbCrLf & "是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '聯絡人
      strNo = GetPotCustCont(2, RTrim(UCase(txtPCU(3) & " " & txtPCU(4) & " " & txtPCU(5) & " " & txtPCU(6))), strTmp)
      If strNo <> MsgText(601) Then
         If MsgBox("本客戶資料與" & strTmp & "[" & strNo & "]的英文名稱一樣是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      'end 2021/08/30
      m_sDupeKey = txtPCU(3) & txtPCU(4) & txtPCU(5) & txtPCU(6) & txtPCU(9) & txtPCU(10)
   End If
   'Add By Sindy 2014/1/27
   '檢查日文名稱
   If txtPCU(7) <> "" And txtPCU(9) <> "" And m_sDupeKey_j <> txtPCU(7) & txtPCU(9) Then
      '國外潛在客戶
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select pcu01||pcu02 from potcustomer where pcu09>='" & Left(txtPCU(9), 3) & "' and pcu09<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(pcu07))='" & ChgSQL(UCase(RTrim(txtPCU(7)))) & "' and rownum<2"
      strExc(0) = "select pcu01||pcu02 from potcustomer where pcu09>='" & Left(txtPCU(9), 3) & "' and pcu09<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(pcu07))=Rtrim('" & ChgSQL(UCase(txtPCU(7))) & "') and rownum<2"
      If txtPCU(1) <> "" Then
         strExc(0) = strExc(0) & " and pcu01||pcu02<>'" & Left(txtPCU(1) & txtPCU(2) & "000", 9) & "'"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與國外潛在客戶[" & RsTemp.Fields(0) & "]的日文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '客戶
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select cu01||cu02 from customer where cu10>='" & Left(txtPCU(9), 3) & "' and cu10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(cu06))='" & ChgSQL(UCase(RTrim(txtPCU(7)))) & "' and rownum<2"
      strExc(0) = "select cu01||cu02 from customer where cu10>='" & Left(txtPCU(9), 3) & "' and cu10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(cu06))=Rtrim('" & ChgSQL(UCase(txtPCU(7))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與客戶[" & RsTemp.Fields(0) & "]的日文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '代理人
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select fa01||fa02 from fagent where fa10>='" & Left(txtPCU(9), 3) & "' and fa10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(fa06))='" & ChgSQL(UCase(RTrim(txtPCU(7)))) & "' and rownum<2"
      strExc(0) = "select fa01||fa02 from fagent where fa10>='" & Left(txtPCU(9), 3) & "' and fa10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(fa06))=Rtrim('" & ChgSQL(UCase(txtPCU(7))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與代理人[" & RsTemp.Fields(0) & "]的日文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '國內潛在客戶
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select poc01||poc02 from potcustomer1 where poc04>='" & Left(txtPCU(9), 3) & "' and poc04<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(poc27))='" & ChgSQL(UCase(RTrim(txtPCU(7)))) & "' and rownum<2"
      strExc(0) = "select poc01||poc02 from potcustomer1 where poc04>='" & Left(txtPCU(9), 3) & "' and poc04<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(poc27))=Rtrim('" & ChgSQL(UCase(txtPCU(7))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與國內潛在客戶[" & RsTemp.Fields(0) & "]的日文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '法務開拓
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select ecd02||'-'||LPAD(ecd01,6,'0') from expandcusdetail where ecd10>='" & Left(txtPCU(9), 3) & "' and ecd10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(ecd03||' '||ecd04))='" & ChgSQL(UCase(RTrim(txtPCU(7)))) & "' and rownum<2"
      strExc(0) = "select ecd02||'-'||LPAD(ecd01,6,'0') from expandcusdetail where ecd10>='" & Left(txtPCU(9), 3) & "' and ecd10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(ecd03||' '||ecd04))=Rtrim('" & ChgSQL(UCase(txtPCU(7))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與法務開拓[" & RsTemp.Fields(0) & "]的日文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      'Add by Amy 2021/08/30
      '國內開拓函特定公司不列印者
      If HasTMBulletinnp(RTrim(UCase(txtPCU(7)))) = True Then
         If MsgBox("本客戶資料與[ 國內開拓函特定公司不列印者]名稱一樣" & _
          vbCrLf & "是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '聯絡人
      strNo = GetPotCustCont(3, RTrim(UCase(txtPCU(7))), strTmp)
      If strNo <> MsgText(601) Then
         If MsgBox("本客戶資料與" & strTmp & "[" & strNo & "]的英文名稱一樣是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      'end 2021/08/30
      m_sDupeKey_j = txtPCU(7) & txtPCU(9)
   End If
   '檢查中文名稱
   If txtPCU(8) <> "" And txtPCU(9) <> "" And m_sDupeKey_c <> txtPCU(8) & txtPCU(9) Then
      '國外潛在客戶
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select pcu01||pcu02 from potcustomer where pcu09>='" & Left(txtPCU(9), 3) & "' and pcu09<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(pcu08))='" & ChgSQL(UCase(RTrim(txtPCU(8)))) & "' and rownum<2"
      strExc(0) = "select pcu01||pcu02 from potcustomer where pcu09>='" & Left(txtPCU(9), 3) & "' and pcu09<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(pcu08))=Rtrim('" & ChgSQL(UCase(txtPCU(8))) & "') and rownum<2"
      If txtPCU(1) <> "" Then
         strExc(0) = strExc(0) & " and pcu01||pcu02<>'" & Left(txtPCU(1) & txtPCU(2) & "000", 9) & "'"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與國外潛在客戶[" & RsTemp.Fields(0) & "]的中文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '客戶
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select cu01||cu02 from customer where cu10>='" & Left(txtPCU(9), 3) & "' and cu10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(cu04))='" & ChgSQL(UCase(RTrim(txtPCU(8)))) & "' and rownum<2"
      strExc(0) = "select cu01||cu02 from customer where cu10>='" & Left(txtPCU(9), 3) & "' and cu10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(cu04))=Rtrim('" & ChgSQL(UCase(txtPCU(8))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與客戶[" & RsTemp.Fields(0) & "]的中文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '代理人
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select fa01||fa02 from fagent where fa10>='" & Left(txtPCU(9), 3) & "' and fa10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(fa04))='" & ChgSQL(UCase(RTrim(txtPCU(8)))) & "' and rownum<2"
      strExc(0) = "select fa01||fa02 from fagent where fa10>='" & Left(txtPCU(9), 3) & "' and fa10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(fa04))=Rtrim('" & ChgSQL(UCase(txtPCU(8))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與代理人[" & RsTemp.Fields(0) & "]的中文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '國內潛在客戶
      'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select poc01||poc02 from potcustomer1 where poc04>='" & Left(txtPCU(9), 3) & "' and poc04<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(poc03))='" & ChgSQL(UCase(RTrim(txtPCU(8)))) & "' and rownum<2"
      strExc(0) = "select poc01||poc02 from potcustomer1 where poc04>='" & Left(txtPCU(9), 3) & "' and poc04<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(poc03))=Rtrim('" & ChgSQL(UCase(txtPCU(8))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與國內潛在客戶[" & RsTemp.Fields(0) & "]的中文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '法務開拓
       'Modified by Lydia 2018/02/22 名稱有造字不可以下RTrim函數
      'strExc(0) = "select ecd02||'-'||LPAD(ecd01,6,'0') from expandcusdetail where ecd10>='" & Left(txtPCU(9), 3) & "' and ecd10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(ecd03||' '||ecd04))='" & ChgSQL(UCase(RTrim(txtPCU(8)))) & "' and rownum<2"
      strExc(0) = "select ecd02||'-'||LPAD(ecd01,6,'0') from expandcusdetail where ecd10>='" & Left(txtPCU(9), 3) & "' and ecd10<='" & Left(txtPCU(9), 3) & "z' and upper(RTRIM(ecd03||' '||ecd04))=rtrim('" & ChgSQL(UCase(txtPCU(8))) & "') and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與法務開拓[" & RsTemp.Fields(0) & "]的中文名稱、國籍皆相同，是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      'Add by Amy 2021/08/30
      '國內開拓函特定公司不列印者
      If HasTMBulletinnp(RTrim(UCase(txtPCU(8)))) = True Then
         If MsgBox("本客戶資料與[ 國內開拓函特定公司不列印者]名稱一樣" & _
          vbCrLf & "是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      '聯絡人
      strNo = GetPotCustCont(1, RTrim(UCase(txtPCU(8))), strTmp)
      If strNo <> MsgText(601) Then
         If MsgBox("本客戶資料與" & strTmp & "[" & strNo & "]的英文名稱一樣是否仍舊要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
      'end 2021/08/30
      m_sDupeKey_c = txtPCU(8) & txtPCU(9)
   End If
   '2014/1/27 END
   'Add By Sindy 2012/4/9
   If m_EditMode = 1 And bolMustChk = True Then
      If Index = 8 Then
         strExc(0) = "SELECT NT01,NT02,'中' FROM NotAgent WHERE NT02='" & txtPCU(8) & "' "
      ElseIf Index = 7 Then
         strExc(0) = "SELECT NT01,NT07,'日' FROM NotAgent WHERE NT07='" & txtPCU(7) & "' "
      ElseIf Index >= 3 And Index <= 6 Then
         strExc(0) = "SELECT NT01,NT03||' '||NT04||' '||NT05||' '||NT06,'英' FROM NotAgent WHERE Upper(NT03||' '||NT04||' '||NT05||' '||NT06)='" & UCase(Trim(txtPCU(3)) & " " & Trim(txtPCU(4)) & " " & Trim(txtPCU(5)) & " " & Trim(txtPCU(6))) & "' "
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("本客戶資料與不得代理案件之編號[" & RsTemp.Fields(0) & "]的" & RsTemp.Fields(2) & "文名稱相同，請詳細確認！是否仍要建立此筆客戶資料？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            DupeCustCheck = True
            Exit Function
         End If
      End If
   End If
   '2012/4/9 END
End Function

Private Sub AddCombo(p_iID As Integer)
   Select Case p_iID
      Case 1
         PUB_AddDeptCombo cboDept
         
      Case 2
         PUB_AddTitleCombo cboTitle
   End Select
End Sub

Private Function ComposeListX(p_index As Integer) As String
   'Modified by Morgan 2022/1/7
   'strExc(1) = ""
   'If lstUsers(p_index).ListCount > 0 Then
   '   'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
   '   strExc(1) = PUB_Num2Id(lstUsers(p_index).ITEMDATA(0))
   '   For intI = 1 To lstUsers(p_index).ListCount - 1
   '      strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(p_index).ITEMDATA(intI))
   '   Next
   'End If
   'ComposeListX = strExc(1)
   ComposeListX = lstUsers(p_index).Tag
   'end 2022/1/7
End Function

Private Function ComposeList(oList As ListBox) As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      strExc(1) = oList.List(0)
      For intI = 1 To oList.ListCount - 1
         strExc(1) = strExc(1) & "," & oList.List(intI)
      Next
   End If
   ComposeList = strExc(1)
End Function

Private Function ContNameChanged() As Boolean
   With rsContactOld
      .MoveFirst
      .Find "pcc02='" & txtPCC(2) & "'"
      If Not .EOF Then
         If UCase("" & .Fields("pcc03")) <> UCase(txtPCC(3)) Or UCase("" & .Fields("pcc04")) <> UCase(txtPCC(4)) Or UCase("" & .Fields("pcc05")) <> UCase(txtPCC(5)) Then
            ContNameChanged = True
         End If
      End If
   End With
End Function

'更新相關聯絡人資料
Private Sub UpdateConRefList(p_stContNo1, Optional p_stContNo2 As String)
   Dim ii As Integer, idx As Integer
   idx = 0
   For ii = 1 To UBound(m_arrConRefList, 2)
      If m_arrConRefList(0, ii) = p_stContNo1 Then
         idx = ii
         Exit For
      End If
   Next
   '若尚無本客戶聯絡人資料時新增
   If idx = 0 Then
      idx = UBound(m_arrConRefList, 2) + 1
      ReDim Preserve m_arrConRefList(1, idx)
      m_arrConRefList(0, idx) = p_stContNo1
   End If
   '紀錄相關聯絡人編號
   m_arrConRefList(1, idx) = p_stContNo2
End Sub

'移除相關聯絡人資料
Private Sub RemoveConRefList()
   Dim ii As Integer
   For ii = 1 To UBound(m_arrConRefList, 2)
      If m_arrConRefList(0, ii) = txtPCC(2) Then
         m_arrConRefList(1, ii) = ""
         m_arrConRefList(0, ii) = ""
         Exit For
      End If
   Next
End Sub

'更新相關聯絡人
Private Sub UpdateRefContact(p_stPCC01 As String)
   Dim ii As Integer
   For ii = 1 To UBound(m_arrConRefList, 2)
      If m_arrConRefList(0, ii) <> "" Then
         If m_arrConRefList(1, ii) = "" Then
            strSql = "update potcustcont set pcc20=Null where pcc20='" & p_stPCC01 & m_arrConRefList(0, ii) & "'"
         Else
            strSql = "update potcustcont set pcc20='" & p_stPCC01 & m_arrConRefList(0, ii) & "' where pcc01='" & Left(m_arrConRefList(1, ii), 8) & "' and pcc02='" & Mid(m_arrConRefList(1, ii), 9) & "'"
         End If
         adoTaie.Execute strSql, intI
      End If
   Next
End Sub

''Add By Sindy 2009/04/30
''檢查維護權限
''Modified by Lydia 2019/06/18
''Private Function CheckModifyLimit(strChkID As String, bModify As Boolean) As Boolean
'Private Function CheckModifyLimit(strChkID As String, strModify As String) As Boolean
'   CheckModifyLimit = True
'   If Trim(strUserNum) = "" Or Trim(strChkID) = "" Then Exit Function
'
'   'Added by Lydia 2019/06/18 修改為67002新增的資料只能67002維護, 其他人新增的資料除了67002以外都可以維護 ;
'                                           '另外99047從國外部調到管理部,先前的規則會令國外部人員無法修改99047新增的資料
'   If Pub_StrUserSt03 = "M51" Or (strUserNum = "67002" And strChkID = strUserNum) Or (strChkID <> "67002" And strUserNum <> "67002") Then
'        Exit Function
'   End If
'   If strModify = "M" Then
'        MsgBox "無修改權限 !!!", vbInformation
'   ElseIf strModify = "D" Then
'        MsgBox "無刪除權限 !!!", vbInformation
'   End If
'   CheckModifyLimit = False
'   Exit Function
'   'end 2019/06/18
'
'   '依LoginUser和輸入人員之部門第一碼判斷部門權限, 不同部門資料不可修改, 但可查詢
'   '但M51不受限制
'   strExc(0) = "SELECT A.ST03,B.ST03 FROM STAFF A,STAFF B " & _
'                      "WHERE A.ST01 = '" & strUserNum & "' " & _
'                           "AND B.ST01 = '" & Trim(strChkID) & "' "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If Trim(RsTemp(0)) = "M51" Then Exit Function
''      If bModify = True Then
''         If Trim(RsTemp(0)) = Trim(RsTemp(1)) Then Exit Function
''      Else
'         If Left(Trim(RsTemp(0)), 1) = Left(Trim(RsTemp(1)), 1) Then Exit Function
''      End If
'   Else
'      Exit Function
'   End If
'
'   CheckModifyLimit = False
'   'Remove by Lydia 2019/06/18
'   'If bModify = True Then
'   '   MsgBox "無修改權限 !!!", vbInformation
'   'End If
'End Function

'Add By Sindy 2009/06/23
Private Function GetCustData(p_stCust As String, p_stPCU11 As String) As Boolean
Dim aiOrder(1 To 3) As Integer
   'Added by Lydia 2016/11/29
   For Each oLabel In Lbl2
      oLabel.Caption = Empty
   Next
   'end 2016/11/29
   
   Select Case Left(p_stCust, 1)
      Case "X"
         If p_stPCU11 = "1" Then
            'Modified by Lydia 2016/11/29
            'strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
            strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 Na01,NA03,cu80 s1 from customer,nation where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "' and cu10=na01(+)"
         Else
            'Modified by Lydia 2020/05/07 關係企業=>lblTitle.Caption
            MsgBox "類別為2事務所時，" & lblTitle.Caption & "必須為 Y 開頭", vbCritical + vbOKOnly, "檢核資料"
            Exit Function
         End If
      Case "Y"
         If p_stPCU11 = "2" Then
            'Modified by Lydia 2016/11/29
            'strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 N3 from fagent where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "'"
            strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 Na01,na03,fa69 s1 from fagent,nation where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "' and fa10=na01(+)"
         Else
            'Modified by Lydia 2020/05/07 關係企業=>lblTitle.Caption
            MsgBox "類別為１廠商時，" & lblTitle.Caption & "必須為 X 開頭", vbCritical + vbOKOnly, "檢核資料"
            Exit Function
         End If
      Case Else
           'Modified by Lydia 2020/05/07 關係企業=>lblTitle.Caption
           MsgBox lblTitle.Caption & "必須為 X 或 Y 開頭", vbCritical + vbOKOnly, "檢核資料"
           Exit Function
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   txtPCU47N.Text = ""
   lblPCU47N.Caption = ""
   If intI = 1 Then
      For intI = 1 To 3
         'Modified by Lydia 2016/11/29 關聯企業：名稱帶入關聯企業之頁籤
         'If Not IsNull(RsTemp(intI)) Then
         '   txtPCU47N.Text = RsTemp(intI)
         '   Exit For
         'End If
         If txtPCU47N.Text = "" And "" & RsTemp(intI) <> "" Then txtPCU47N.Text = "" & RsTemp(intI)
         lblPCU47N.Caption = txtPCU47N
         Lbl2(2 + intI) = "" & RsTemp(intI)
         'end 2016/11/29
      Next
      'Added by Lydia 2016/11/29 關聯企業名稱,國籍和狀態
      Lbl2(0) = Left("" & RsTemp.Fields("na01"), 3)
      Lbl2(1) = Trim(Mid("" & RsTemp.Fields("na03"), 1, 4))
      Lbl2(2) = "" & RsTemp.Fields("s1")
      'end 2016/11/29
      
      GetCustData = True
   Else
      'Modified by Lydia 2020/05/07 關係企業=>lblTitle.Caption
      MsgBox lblTitle.Caption & "輸入錯誤！"
   End If
End Function

'Add By Sindy 2009/06/23
Private Sub txtPCU47N_GotFocus()
   OpenIme
   TextInverse txtPCU47N
End Sub

'Add By Sindy 2009/06/23
Private Sub txtPCU47N_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2009/06/23
Public Sub txtPCU47N_Validate(Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If txtPCU(47).Text = "" And txtPCU47N.Text <> "" Then
'         If txtSameCnt = "Y" Then
'            Me.Show
'            Me.cboCity.SetFocus
'            txtSameCnt = ""
'         ElseIf txtSameCnt = "E" Then
'            Me.Show
'            Me.txtPCU47N.SetFocus
'            Cancel = True
'            txtSameCnt = ""
'            Exit Sub
'         Else
'            Me.Enabled = False
'            Screen.MousePointer = vbHourglass
'            frm210128_2.Show
'            Call frm210128_2.StrMenu("0", Me.txtPCU47N.Text)
'            Screen.MousePointer = vbDefault
'            Me.Enabled = True
'            Me.Hide
'            If txtSameCnt.Text = 0 Then txtSameCnt.Text = ""
'            If Trim(txtPCU(47).Text) <> "" And Not IsNull(txtPCU(47)) Then
'               Me.Show
'               Me.cboCity.SetFocus
'            ElseIf txtSameCnt.Text = "" Then
'               Me.Show
'               Me.txtPCU47N.SetFocus
'               Cancel = True
'               Exit Sub
'            End If
'         End If
         'Modify By Sindy 2014/2/27
         txtSameCnt = ""
         If frm210128_2.StrMenu("0", Me.txtPCU47N.Text) = True Then
            Me.Hide
            'Modified by Lydia 2020/05/07
            'frm210128_2.Caption = "關聯企業名稱" 'Added by Lydia 2016/11/29
            frm210128_2.Caption = lblTitle.Caption & "名稱"
            frm210128_2.Show vbModal
            If txtSameCnt = "Y" Then
               Me.Show
               Me.cboCity.SetFocus
               'Added by Lydia 2016/11/29 取得關聯企業的資料
               'modify by sonia 2021/12/29 txtPCU(11)-->Left(cboPCU11.Text, 1)
               If GetCustData(txtPCU(47), Left(cboPCU11.Text, 1)) = False Then
               End If
               'end 2016/11/29
               txtSameCnt = ""
            ElseIf txtSameCnt = "E" Then
               Me.Show
               Me.txtPCU47N.SetFocus
               Cancel = True
               txtSameCnt = ""
               Exit Sub
            End If
         Else
            Unload frm210128_2
            Me.Show
            If Trim(txtPCU(47).Text) <> "" And Not IsNull(txtPCU(47)) Then
               Me.cboCity.SetFocus
            ElseIf txtSameCnt.Text = "" Then
               Me.txtPCU47N.SetFocus
               Cancel = True
               Exit Sub
            End If
         End If
         '2014/2/27 END
      End If
   End If
End Sub

'Added by Lydia 2016/11/29
Private Sub cmdAddPCU49_Click()
'Modified by Lydia 2017/06/28 從數字(intA)改成文字(strAns)
'Dim intA As Integer
Dim strAns As String

  If Combo1.ListIndex >= 0 And (m_EditMode = 1 Or m_EditMode = 2) Then
     If txtPCU(47) = "" Or txtPCU47N = "" Then
         MsgBox "請輸入關聯企業!", vbCritical
         txtPCU(47).SetFocus
         Exit Sub
     End If
     'Modified by Lydia 2017/06/28 從數字(intA)改成文字(strAns)
     Call Pub_AddFRelationList(Me.Combo1, Me.List1, strAns)
     If strAns <> "" Then txtPCU(49).Text = txtPCU(49).Text & strAns & ","
     'end 2017/06/28
  End If
End Sub

Private Sub cmdRemPCU49_Click()
  If m_EditMode = 1 Or m_EditMode = 2 Then
     If txtPCU(47) = "" Or txtPCU47N = "" Then
         MsgBox "請輸入關聯企業!", vbCritical
         txtPCU(47).SetFocus
         Exit Sub
     End If
     
     Call Pub_RemSelectList(List1, strExc(1))
     If strExc(1) <> "" Then
        txtPCU(49).Text = Replace(txtPCU(49).Text, strExc(1) & ",", "")
     End If
  End If
End Sub
'end 2016/11/29

'Added by Lydia 2018/02/22 去掉跳行符號
Private Sub txtPCC_Change(Index As Integer)
  'Modified by Lydia 2022/12/07 排除備註( 當時是為了"修正無法輸入造字的問題---Gill")
  'If m_EditMode = 1 Or m_EditMode = 2 Then txtPCC(Index) = PUB_StringFilter(txtPCC(Index))
  If (m_EditMode = 1 Or m_EditMode = 2) And Index <> 13 Then
       txtPCC(Index) = PUB_StringFilter(txtPCC(Index))
  End If
  'end 2022/12/07
End Sub

'Add by Amy 2021/08/16 資料會於 frm210128_1(國內潛在客戶維護)顯示
'Modify by Amy 2024/05/21 +風險檢查語法
Private Function ShowData(ByVal stRCLSql As String) As Boolean
    Dim strQ As String, intQ As Integer
    
    ShowData = False
     'Add by Amy 2024/05/21 +風險檢查語法
    If stRCLSql <> MsgText(601) Then
      strQ = "Union Select Distinct RCL01 AS 編號,Decode(RCLFIeld,'英',rtrim(RCL03||' '||RCL04||' '||RCL05||' '||RCL06),'中',RCL02,RCL07) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員" & _
                      ",Decode(RCL10,null,Decode(RCL09,null,RCL16,RCL09),rtrim(RCL10||' '||RCL11||' '||RCL12||' '||RCL13||' '||RCL14||' '||RCL15)) AS 地址,'' AS 電話,'' AS 傳真,'' AS 狀態,RCL23 AS 備註 " & _
                   "From (" & stRCLSql & "),Nation Where SubStr(RCL08,1,3)=NA01(+) "
    End If
    strQ = "Select Distinct R021001 AS 編號,R021002 AS 名稱,'' AS 國籍,'' AS 智權人員,'' AS 地址,'' AS 電話,'' AS 傳真,Decode(R021004,'1','對造','其他相關人') AS 狀態,'' AS 備註 " & _
                "From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' And R021004<3 " & strQ
    'end 2024/05/21
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If RsQ.RecordCount <= 0 Then
            Exit Function
        Else
            ShowData = True
        End If
    End If
End Function

'Add by Amy 2021/08/30 有 TMBulletinnp 資料
Private Function HasTMBulletinnp(ByVal stFindS As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    HasTMBulletinnp = False
    stQ = "Select * From TMBulletinnp Where Upper(RTrim(TBNP01))='" & stFindS & "' and RowNum<2"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        HasTMBulletinnp = True
    End If
    Set RsQ = Nothing
End Function

Private Function GetPotCustCont(ByVal intChoose, ByVal stFindS As String, ByRef stBackTBN As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    GetPotCustCont = "": stBackTBN = ""
    Select Case intChoose
        Case 1 '中
            stQ = "Select '客戶' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where pcc05='" & stFindS & "') A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' "
            stQ = stQ & " Union All Select '國內潛在客戶' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where pcc05='" & stFindS & "') A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) "
            stQ = stQ & " Union All Select '國外潛在客戶' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where pcc05='" & stFindS & "') A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
            stQ = stQ & " Union All Select '代理人' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where pcc05='" & stFindS & "') A,Fagent Where FA01(+)=PCC01 And FA02='0' "
        Case 2 '英
            stQ = " Select '客戶' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where Upper(pcc03)='" & stFindS & "') A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' "
            stQ = stQ & " Union All Select '國內潛在客戶' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where Upper(pcc03)='" & stFindS & "') A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) "
            stQ = stQ & " Union All Select '國外潛在客戶' as TBN,PCC01||'0-'||PCC02 as PNo  From (Select * From PotCustCont Where Upper(pcc03)='" & stFindS & "') A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
            stQ = stQ & " Union All Select '代理人' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where Upper(pcc03)='" & stFindS & "') A,Fagent Where FA01(+)=PCC01 And FA02='0' "
        Case 3 '日
            stQ = "Select '客戶' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where pcc04='" & stFindS & "') A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' "
            stQ = stQ & " Union All Select '國內潛在客戶' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where pcc04='" & stFindS & "') A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) "
            stQ = stQ & " Union All Select '國外潛在客戶' as TBN,PCC01||'0-'||PCC02 as PNo  From (Select * From PotCustCont Where pcc04='" & stFindS & "') A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
            stQ = stQ & " Union All Select '代理人' as TBN,PCC01||'0-'||PCC02 as PNo From (Select * From PotCustCont Where pcc04='" & stFindS & "') A,Fagent Where FA01(+)=PCC01 And FA02='0' "
    End Select
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        stBackTBN = RsQ.Fields("TBN")
        GetPotCustCont = RsQ.Fields("PNo")
    End If
    Set RsQ = Nothing
End Function

'Add by Amy 2021/11/29 潛在客戶為國內同業聯絡人欄位設定判斷
Private Function ChkSameTradePCU_Old() As String
'    Dim RsQ As New ADODB.Recordset
'    Dim intQ As Integer, stQ As String, stTmp As String
'
'    ChkSameTradePCU = ""
'    stQ = "Select Pcc02,Pcc08,Pcc09,Pcc10,Pcc24 From PotCustCont Where Pcc01='" & txtPCU(1) & "' " & _
'            "And (Pcc08 is not null Or Nvl(Pcc09,'0')<>'N' Or Nvl(Pcc10,'0')<>'N' Or Nvl(Pcc24,'0')<>'N') "
'    intQ = 1
'    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
'    If intQ = 1 Then
'        RsQ.MoveFirst
'        Do While Not RsQ.EOF
'            If Not IsNull(RsQ.Fields("Pcc08")) Then
'                If stTmp <> MsgText(601) Then stTmp = stTmp & vbCrLf
'                stTmp = stTmp & "不可輸入E-Mail以免誤發電子郵件, 如有需要請加註於備註欄 ！"
'            End If
'            If "" & RsQ.Fields("Pcc09") <> "N" Then
'                If stTmp <> MsgText(601) And Right(stTmp, 1) = "！" Then stTmp = stTmp & vbCrLf
'                stTmp = stTmp & "不可寄台一雜誌、"
'            End If
'            If "" & RsQ.Fields("Pcc10") <> "N" Then
'                If stTmp <> MsgText(601) And Right(stTmp, 1) = "！" Then stTmp = stTmp & vbCrLf
'                stTmp = stTmp & "不可寄電子報、"
'            End If
'            If "" & RsQ.Fields("Pcc24") <> "N" Then
'                If stTmp <> MsgText(601) And Right(stTmp, 1) = "！" Then stTmp = stTmp & vbCrLf
'                stTmp = stTmp & "不可寄專利雙週報、"
'            End If
'            If stTmp <> MsgText(601) Then
'                ChkSameTradePCU = ChkSameTradePCU & "@@編號「" & RsQ.Fields("Pcc02") & "」：" & vbCrLf
'                If Right(stTmp, 1) = "、" Then
'                    ChkSameTradePCU = ChkSameTradePCU & Mid(stTmp, 1, Len(stTmp) - 1) & "！"
'                Else
'                    ChkSameTradePCU = ChkSameTradePCU & stTmp
'                End If
'                stTmp = ""
'            End If
'            RsQ.MoveNext
'        Loop
'        If ChkSameTradePCU <> MsgText(601) Then
'            ChkSameTradePCU = Replace(Mid(ChkSameTradePCU, 3), "@@", vbCrLf & vbCrLf)
'        End If
'    End If
'    Set RsQ = Nothing
End Function

'Add by Amy 2023/02/01 潛在客戶為國內同業聯絡人欄位設定判斷(畫面上)
Private Function ChkSameTradePCU() As String
    Dim stTmp As String
    
    With rsContact
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If Trim("" & .Fields("Pcc08")) <> MsgText(601) Then
                    If stTmp <> MsgText(601) Then stTmp = stTmp & vbCrLf
                    stTmp = stTmp & "不可輸入E-Mail以免誤發電子郵件, 如有需要請加註於備註欄 ！"
                End If
                If "" & .Fields("Pcc09") <> "N" Then
                    If stTmp <> MsgText(601) And Right(stTmp, 1) = "！" Then stTmp = stTmp & vbCrLf
                    stTmp = stTmp & "不可寄台一雜誌、"
                End If
                If "" & .Fields("Pcc10") <> "N" Then
                    If stTmp <> MsgText(601) And Right(stTmp, 1) = "！" Then stTmp = stTmp & vbCrLf
                    stTmp = stTmp & "不可寄電子報、"
                End If
                If "" & .Fields("Pcc24") <> "N" Then
                    If stTmp <> MsgText(601) And Right(stTmp, 1) = "！" Then stTmp = stTmp & vbCrLf
                    stTmp = stTmp & "不可寄專利雙週報、"
                End If
                If stTmp <> MsgText(601) Then
                    ChkSameTradePCU = ChkSameTradePCU & "@@編號「" & .Fields("Pcc02") & "」：" & vbCrLf
                    If Right(stTmp, 1) = "、" Then
                        ChkSameTradePCU = ChkSameTradePCU & Mid(stTmp, 1, Len(stTmp) - 1) & "！"
                    Else
                        ChkSameTradePCU = ChkSameTradePCU & stTmp
                    End If
                    stTmp = ""
                End If
                .MoveNext
            Loop
            If ChkSameTradePCU <> MsgText(601) Then
                ChkSameTradePCU = Replace(Mid(ChkSameTradePCU, 3), "@@", vbCrLf & vbCrLf)
            End If
        End If
    End With
End Function

'Added by Morgan 2023/7/6
'預設定稿語文
Private Sub setPCU36()
   If txtPCU(36) = "" Then
      If txtPCU(9) <> "" Then
         '台灣、大陸
         If Val(txtPCU(9)) < 9 Or txtPCU(9) = "020" Then
            txtPCU(36) = 1
         '香港
         ElseIf txtPCU(9) = "013" Then
            If cboPCU11 <> "" Then
               '事務所
               If Left(cboPCU11, 1) = "2" Then
                  txtPCU(36) = 2
               '非事務所
               Else
                  txtPCU(36) = 1
               End If
            End If
         '日本
         ElseIf Left(txtPCU(9), 3) = "011" Then
            txtPCU(36) = 3
         '其他
         Else
            txtPCU(36) = 2
         End If
      End If
   End If
End Sub

'Added by Lydia 2024/05/14
Private Sub Command1_Click()
   If Trim(txtPCC(2)) = "" Then
      MsgBox "請輸入聯絡人編號！"
      Exit Sub
   Else
      If Left(txtPCU(1), 1) <> "R" Then
         MsgBox "請輸入潛在客戶編號！"
         txtPCU(1).SetFocus
         txtPCU_GotFocus 1
         Exit Sub
      End If
      strExc(0) = "select pcc01,pcc02 from potcustcont where pcc01='" & txtPCU(1) & "' and pcc02='" & txtPCC(2) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         MsgBox "請先將聯絡人建檔完成！"
         Exit Sub
      End If
   End If
   
   frmPic001.oCP01 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "1")
   frmPic001.oCP02 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "2")
   frmPic001.oCP03 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "3")
   frmPic001.oCP04 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "4")
   frmPic001.strWorkType = "1"
   frmPic001.Label11 = "聯絡人相片上傳"
   If m_EditMode <> 1 And m_EditMode <> 2 Then
      frmPic001.bolQuery = True '只查詢
   Else
      frmPic001.bolQuery = False '可存檔
   End If
   frmPic001.StrMenu
   frmPic001.SetSeekCmdok
   frmPic001.Show vbModal
   
   Call Pub_GetPCCtoIBF_2(Trim(txtPCU(1)), Trim(txtPCC(2)), Command1)

End Sub

'Add by Amy 2024/11/29
Private Sub txtXYS02_GotFocus()
   InverseTextBox txtXYS02
End Sub

Private Sub txtXYS02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtXYS02_Validate(Cancel As Boolean)
   Dim stName As String, stMsg As String
   
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtXYS02 = MsgText(601) Then LblSourceN.Caption = "": Exit Sub
   
   bCancel = False
   LblSourceN.Caption = ""
   txtXYS02 = Left(ChangeCustomerL(txtXYS02), 8) '補滿8碼
   stMsg = ChkXYSourceReason(1, Me.Name, m_EditMode, cboSource, txtXYS02, , , , , txtPCU(1), stName)
   If stMsg <> MsgText(601) Then
      MsgBox stMsg, vbInformation
      tabCustomer.Tab = 0
      'Memo 使用bCancel避免彈訊息後無法跳離 ex:來源選04 輸了Y編號,需刪Y編號,再重選
      bCancel = True
      txtXYS02_GotFocus
      Exit Sub
   End If
   LblSourceN.Caption = stName
End Sub

Private Sub txtXYS03_GotFocus()
   InverseTextBox txtXYS03
End Sub
'end 2024/11/29
