VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090202_4_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "待送件區-承辦歷程"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8955
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8955
   Tag             =   "加班資料"
   Begin VB.CommandButton CmdCalendar 
      Caption         =   "行事曆"
      Height          =   320
      Left            =   7110
      Style           =   1  '圖片外觀
      TabIndex        =   112
      Top             =   720
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdOutlook 
      Caption         =   "匯出Outlook"
      Height          =   320
      Left            =   6090
      Style           =   1  '圖片外觀
      TabIndex        =   87
      Top             =   720
      Visible         =   0   'False
      Width           =   1010
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Height          =   320
      Index           =   0
      Left            =   7815
      Style           =   1  '圖片外觀
      TabIndex        =   80
      Top             =   720
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   320
      Index           =   1
      Left            =   8130
      Style           =   1  '圖片外觀
      TabIndex        =   79
      Top             =   390
      Width           =   765
   End
   Begin VB.CommandButton cmdFlow 
      Caption         =   "聯絡"
      Height          =   320
      Index           =   1
      Left            =   5700
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   390
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton cmdFlow 
      Caption         =   "送判"
      Height          =   320
      Index           =   0
      Left            =   6510
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   390
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtLpNote 
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   2370
      TabIndex        =   70
      Text            =   "(共X筆)"
      Top             =   30
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3750
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   6120
      Width           =   4875
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3060
      Locked          =   -1  'True
      MousePointer    =   1  '箭號形狀
      TabIndex        =   67
      Text            =   "存卷資料"
      Top             =   720
      Width           =   1305
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   6930
      TabIndex        =   66
      Top             =   6480
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lstNameAgent 
      Height          =   240
      ItemData        =   "frm090202_4_1.frx":0000
      Left            =   7770
      List            =   "frm090202_4_1.frx":0007
      TabIndex        =   65
      Top             =   6480
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3750
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   5790
      Width           =   4875
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      Height          =   320
      Index           =   4
      Left            =   7290
      TabIndex        =   5
      Top             =   390
      Width           =   830
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "接洽單"
      Height          =   320
      Index           =   2
      Left            =   7320
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   30
      Width           =   765
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1980
      TabIndex        =   58
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   6480
      Width           =   4395
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "退件(&B)"
      Height          =   320
      Left            =   6510
      TabIndex        =   1
      Top             =   30
      Width           =   765
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   320
      Left            =   8130
      TabIndex        =   6
      Top             =   30
      Width           =   765
   End
   Begin VB.PictureBox pic1 
      Height          =   420
      Left            =   8430
      ScaleHeight     =   360
      ScaleWidth      =   420
      TabIndex        =   48
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5050
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   8930
      _ExtentX        =   15743
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "簽辦流程"
      TabPicture(0)   =   "frm090202_4_1.frx":0019
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdM51"
      Tab(0).Control(1)=   "txtEEP02"
      Tab(0).Control(2)=   "txtEEP10"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(4)=   "txtEEP03"
      Tab(0).Control(5)=   "CboEEP04"
      Tab(0).Control(6)=   "GRD1"
      Tab(0).Control(7)=   "CommonDialog1"
      Tab(0).Control(8)=   "Winsock1"
      Tab(0).Control(9)=   "Label25(0)"
      Tab(0).Control(10)=   "CboEEP05"
      Tab(0).Control(11)=   "txtEEP03_2"
      Tab(0).Control(12)=   "txtEEP08"
      Tab(0).Control(13)=   "txtEEP10_2"
      Tab(0).Control(14)=   "lblCM10"
      Tab(0).Control(15)=   "lblSendMailDt"
      Tab(0).Control(16)=   "lblEApp"
      Tab(0).Control(17)=   "Label4"
      Tab(0).Control(18)=   "Label1(3)"
      Tab(0).Control(19)=   "Label2"
      Tab(0).Control(20)=   "Label15"
      Tab(0).Control(21)=   "Label1(0)"
      Tab(0).Control(22)=   "Label10(0)"
      Tab(0).Control(23)=   "Label3(0)"
      Tab(0).Control(24)=   "Label5"
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "承辦單內容"
      TabPicture(1)   =   "frm090202_4_1.frx":0035
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(4)"
      Tab(1).Control(1)=   "txt1(0)"
      Tab(1).Control(2)=   "txt1(5)"
      Tab(1).Control(3)=   "txt1(6)"
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(7)=   "Label7"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "歷程備註"
      TabPicture(2)   =   "frm090202_4_1.frx":0051
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame945"
      Tab(2).Control(1)=   "Frame201"
      Tab(2).Control(2)=   "Label20"
      Tab(2).Control(3)=   "txtEP12"
      Tab(2).Control(4)=   "txt3(7)"
      Tab(2).Control(5)=   "Label18"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "存卷資料"
      TabPicture(3)   =   "frm090202_4_1.frx":006D
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "LblinfoNote"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame945 
         Caption         =   "【管制下一程序期限】"
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
         Height          =   1870
         Left            =   -74910
         TabIndex        =   107
         Top             =   420
         Visible         =   0   'False
         Width           =   5350
         Begin MSForms.TextBox txtEED14 
            Height          =   300
            Left            =   2820
            TabIndex        =   111
            Top             =   300
            Width           =   1020
            VariousPropertyBits=   680542239
            MaxLength       =   7
            ScrollBars      =   2
            Size            =   "1799;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "追蹤客戶指示【約定期限】："
            Height          =   180
            Index           =   11
            Left            =   90
            TabIndex        =   110
            Top             =   360
            Width           =   2700
         End
         Begin MSForms.TextBox txtEED15 
            Height          =   290
            Left            =   2820
            TabIndex        =   109
            Top             =   630
            Width           =   1020
            VariousPropertyBits=   680542239
            MaxLength       =   7
            ScrollBars      =   2
            Size            =   "1799;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "委員指定送件日期【本所期限】："
            Height          =   180
            Index           =   12
            Left            =   90
            TabIndex        =   108
            Top             =   690
            Width           =   2700
         End
      End
      Begin VB.Frame Frame201 
         Height          =   2350
         Left            =   -74910
         TabIndex        =   89
         Top             =   420
         Width           =   8770
         Begin VB.Frame Frame7 
            Height          =   280
            Left            =   150
            TabIndex        =   92
            Top             =   0
            Width           =   3940
            Begin VB.Label LblEED10_N 
               AutoSize        =   -1  'True
               Caption         =   "LblEED10_N"
               Height          =   180
               Left            =   1470
               TabIndex        =   96
               Top             =   60
               Width           =   910
            End
            Begin VB.Label LblEED10 
               AutoSize        =   -1  'True
               Caption         =   "譯者："
               Height          =   180
               Left            =   210
               TabIndex        =   95
               Top             =   30
               Width           =   540
            End
            Begin MSForms.TextBox txt3 
               Height          =   320
               Index           =   3
               Left            =   750
               TabIndex        =   94
               Top             =   0
               Width           =   690
               VariousPropertyBits=   -1466941413
               MaxLength       =   6
               ScrollBars      =   2
               Size            =   "1217;564"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label LblEED10_N_2 
               AutoSize        =   -1  'True
               Caption         =   "LblEED10_N_2"
               Height          =   180
               Left            =   2550
               TabIndex        =   93
               Top             =   60
               Width           =   1070
            End
         End
         Begin VB.CheckBox ChkEED13 
            Caption         =   "轉檔後送件（程序發文）"
            ForeColor       =   &H000000C0&
            Height          =   220
            Left            =   5640
            TabIndex        =   91
            Top             =   660
            Width           =   2710
         End
         Begin VB.ComboBox CmbFL 
            Height          =   260
            Index           =   3
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   90
            Text            =   "CmbFL"
            Top             =   300
            Width           =   6030
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "打字室："
            Height          =   180
            Index           =   10
            Left            =   180
            TabIndex        =   106
            Top             =   690
            Width           =   690
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "中說備註："
            Height          =   180
            Left            =   0
            TabIndex        =   105
            Top             =   1020
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "管制人："
            Height          =   180
            Index           =   9
            Left            =   2820
            TabIndex        =   104
            Top             =   690
            Width           =   690
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   5
            Left            =   3540
            TabIndex        =   103
            Top             =   630
            Width           =   660
            VariousPropertyBits=   -1466941413
            MaxLength       =   6
            ScrollBars      =   2
            Size            =   "1164;564"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   6
            Left            =   900
            TabIndex        =   102
            Top             =   630
            Width           =   660
            VariousPropertyBits=   -1466941413
            MaxLength       =   6
            ScrollBars      =   2
            Size            =   "1164;564"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label LblEED06_N 
            AutoSize        =   -1  'True
            Caption         =   "LblEED06_N"
            Height          =   180
            Left            =   1590
            TabIndex        =   101
            Top             =   690
            Width           =   880
         End
         Begin VB.Label LblEED09_N 
            AutoSize        =   -1  'True
            Caption         =   "LblEED09_N"
            Height          =   180
            Left            =   4290
            TabIndex        =   100
            Top             =   690
            Width           =   880
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   8
            Left            =   900
            TabIndex        =   99
            Top             =   300
            Width           =   7890
            VariousPropertyBits=   -1466941413
            MaxLength       =   50
            ScrollBars      =   2
            Size            =   "13917;564"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "檔案名稱："
            Height          =   180
            Index           =   8
            Left            =   0
            TabIndex        =   98
            Top             =   360
            Width           =   840
         End
         Begin MSForms.TextBox txt3 
            Height          =   1320
            Index           =   4
            Left            =   900
            TabIndex        =   97
            Top             =   960
            Width           =   7890
            VariousPropertyBits=   -1466941413
            MaxLength       =   400
            ScrollBars      =   2
            Size            =   "13917;2328"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.CommandButton cmdM51 
         BackColor       =   &H00C0FFC0&
         Caption         =   "電腦中心取消歸卷"
         Height          =   320
         Left            =   -72360
         Style           =   1  '圖片外觀
         TabIndex        =   88
         Top             =   4710
         Visible         =   0   'False
         Width           =   1670
      End
      Begin VB.Frame Frame5 
         Height          =   3075
         Left            =   150
         TabIndex        =   71
         Top             =   300
         Width           =   8655
         Begin VB.ListBox lstAtt 
            Height          =   2400
            Index           =   1
            ItemData        =   "frm090202_4_1.frx":0089
            Left            =   60
            List            =   "frm090202_4_1.frx":0090
            MultiSelect     =   1  '簡易多重選取
            Sorted          =   -1  'True
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   180
            Width           =   8490
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   345
            Index           =   1
            Left            =   210
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "下載"
            Height          =   345
            Index           =   1
            Left            =   1710
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "新增"
            Height          =   345
            Index           =   1
            Left            =   2460
            TabIndex        =   74
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "刪除"
            Height          =   345
            Index           =   1
            Left            =   3210
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            Height          =   345
            Index           =   1
            Left            =   960
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
      End
      Begin VB.TextBox txtEEP02 
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   -71010
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1980
         Width           =   645
      End
      Begin VB.TextBox txtEEP10 
         Height          =   270
         Left            =   -74910
         TabIndex        =   47
         Top             =   3450
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "歸檔附件區：(歷程順序 0)"
         ForeColor       =   &H000000C0&
         Height          =   3075
         Left            =   -70290
         TabIndex        =   14
         Top             =   1950
         Width           =   4155
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "刪除"
            Height          =   345
            Index           =   0
            Left            =   3480
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1950
            Width           =   675
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "新增"
            Height          =   345
            Index           =   0
            Left            =   2790
            TabIndex        =   20
            Top             =   1950
            Width           =   675
         End
         Begin VB.CommandButton cmdPrintAtt 
            Caption         =   "列印"
            Height          =   345
            Index           =   0
            Left            =   2100
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1950
            Width           =   675
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "下載"
            Height          =   345
            Index           =   0
            Left            =   1410
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1950
            Width           =   675
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            Height          =   345
            Index           =   0
            Left            =   720
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1950
            Width           =   675
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   345
            Index           =   0
            Left            =   30
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1950
            Width           =   675
         End
         Begin VB.CommandButton CmdOpen 
            BackColor       =   &H00FFFFC0&
            Caption         =   "<->"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   2310
            Style           =   1  '圖片外觀
            TabIndex        =   123
            Top             =   -60
            Width           =   460
         End
         Begin VB.TextBox TextFCPNote 
            Appearance      =   0  '平面
            BackColor       =   &H8000000F&
            BorderStyle     =   0  '沒有框線
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   500
            Index           =   0
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   86
            Text            =   "frm090202_4_1.frx":009C
            Top             =   1470
            Width           =   4010
         End
         Begin VB.ListBox lstAtt 
            Height          =   1140
            Index           =   0
            ItemData        =   "frm090202_4_1.frx":00FF
            Left            =   60
            List            =   "frm090202_4_1.frx":0106
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   210
            Width           =   4020
         End
         Begin VB.ComboBox Combo1 
            Height          =   260
            Left            =   840
            Style           =   2  '單純下拉式
            TabIndex        =   25
            Top             =   2760
            Width           =   3270
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  '沒有框線
            Height          =   410
            Left            =   30
            TabIndex        =   56
            Top             =   2280
            Width           =   4095
            Begin VB.CommandButton cmdTrans 
               Caption         =   "轉檔完成"
               Height          =   345
               Left            =   3210
               Style           =   1  '圖片外觀
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   60
               Visible         =   0   'False
               Width           =   885
            End
            Begin VB.CommandButton cmdSelAllPrt 
               Caption         =   "列印全部PDF"
               Height          =   345
               Left            =   1710
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   60
               Width           =   1275
            End
            Begin VB.CommandButton cmdPrintAllPDF 
               Caption         =   "產生承辦單及歸檔"
               Height          =   345
               Left            =   -30
               Style           =   1  '圖片外觀
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   60
               Width           =   1695
            End
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '透明
            Caption         =   "印表機："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   90
            TabIndex        =   57
            Top             =   2790
            Width           =   800
         End
      End
      Begin VB.TextBox txtEEP03 
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   -73980
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1980
         Width           =   645
      End
      Begin VB.ComboBox CboEEP04 
         Height          =   300
         Left            =   -73980
         TabIndex        =   11
         Text            =   "CboEEP04"
         Top             =   2220
         Width           =   2115
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   1395
         Left            =   -74940
         TabIndex        =   26
         Top             =   540
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2461
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "順序|發送者 |流程狀態 |收受者 | 送出時間 |  意見內容"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -74880
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   -74850
         Top             =   4260
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "\\typing2\電子送件暫存區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   -74640
         TabIndex        =   85
         Top             =   4770
         Visible         =   0   'False
         Width           =   1910
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "作業備註："
         Height          =   180
         Left            =   -74910
         TabIndex        =   84
         Top             =   2820
         Width           =   900
      End
      Begin MSForms.TextBox txtEP12 
         Height          =   800
         Left            =   -74010
         TabIndex        =   83
         Top             =   2790
         Width           =   7890
         VariousPropertyBits=   -1466941413
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "13917;1411"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt3 
         Height          =   1275
         Index           =   7
         Left            =   -74010
         TabIndex        =   82
         Top             =   3600
         Width           =   7890
         VariousPropertyBits=   -1466941413
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "13917;2249"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "請款備註："
         Height          =   180
         Left            =   -74910
         TabIndex        =   81
         Top             =   3630
         Width           =   870
      End
      Begin VB.Label LblinfoNote 
         Caption         =   "註：檔名則以P000000.info.pdf，多個以上資料檔則加序號（例：P000000.info2.pdf，P000000.info3.pdf）"
         ForeColor       =   &H000000C0&
         Height          =   230
         Left            =   150
         TabIndex        =   78
         Top             =   4620
         Width           =   8600
      End
      Begin MSForms.ComboBox CboEEP05 
         Height          =   300
         Left            =   -73980
         TabIndex        =   13
         Top             =   2550
         Width           =   2115
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3731;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   2250
         Index           =   4
         Left            =   -73710
         TabIndex        =   45
         Top             =   2520
         Width           =   4710
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "8308;3969"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   750
         Index           =   0
         Left            =   -73710
         TabIndex        =   41
         Top             =   1770
         Width           =   4710
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "8308;1323"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   600
         Index           =   5
         Left            =   -73710
         TabIndex        =   40
         Top             =   570
         Width           =   4710
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "8308;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   600
         Index           =   6
         Left            =   -73710
         TabIndex        =   39
         Top             =   1170
         Width           =   4710
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "8308;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEEP03_2 
         Height          =   285
         Left            =   -73290
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1605
         VariousPropertyBits=   671105055
         Size            =   "2831;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEEP08 
         Height          =   1515
         Left            =   -74370
         TabIndex        =   9
         Top             =   3180
         Width           =   4035
         VariousPropertyBits=   -1466941413
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "7117;2672"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEEP10_2 
         Height          =   285
         Left            =   -73860
         TabIndex        =   8
         Top             =   2880
         Width           =   3525
         VariousPropertyBits=   671105051
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCM10 
         Caption         =   "一案兩請"
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
         Left            =   -69690
         TabIndex        =   62
         Top             =   90
         Visible         =   0   'False
         Width           =   830
      End
      Begin VB.Label lblSendMailDt 
         AutoSize        =   -1  'True
         Caption         =   "寄件日期:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -68730
         TabIndex        =   61
         Top             =   90
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblEApp 
         Caption         =   "電子送件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -70500
         TabIndex        =   60
         Top             =   90
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "註:電子送件，請先加入下載的檔案後，再執行產生承辦單。"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -74940
         TabIndex        =   51
         Top             =   4740
         Width           =   4785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "順序："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   -71580
         TabIndex        =   50
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Left            =   -74790
         TabIndex        =   46
         Top             =   2595
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主旨："
         Height          =   180
         Index           =   2
         Left            =   -74790
         TabIndex        =   44
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "受文者："
         Height          =   180
         Left            =   -74790
         TabIndex        =   43
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "副本收受者："
         Height          =   180
         Left            =   -74790
         TabIndex        =   42
         Top             =   1245
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "(註:雙擊選取時,下方顯示歷程資料)"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -74940
         TabIndex        =   32
         Top             =   330
         Width           =   2895
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "收  受  者："
         Height          =   180
         Left            =   -74910
         TabIndex        =   31
         Top             =   2610
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發  送  者："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   30
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "流程狀態："
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   29
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "內容："
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   28
         Top             =   3210
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "副本收受者："
         Height          =   180
         Left            =   -74910
         TabIndex        =   27
         Top             =   2910
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "發文(&S)"
      Default         =   -1  'True
      Height          =   320
      Left            =   5700
      TabIndex        =   0
      Top             =   30
      Width           =   765
   End
   Begin VB.Frame Frame1Big 
      Caption         =   "附件區："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4660
      Left            =   0
      TabIndex        =   113
      Top             =   720
      Visible         =   0   'False
      Width           =   8920
      Begin VB.TextBox TextFCPNote 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   560
         Index           =   1
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   122
         Text            =   "frm090202_4_1.frx":0112
         Top             =   3600
         Width           =   4010
      End
      Begin VB.CommandButton cmdPrintAtt 
         Caption         =   "列印"
         Height          =   345
         Index           =   2
         Left            =   2490
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   4230
         Width           =   675
      End
      Begin VB.ListBox lstAtt 
         Height          =   3120
         Index           =   2
         ItemData        =   "frm090202_4_1.frx":017F
         Left            =   60
         List            =   "frm090202_4_1.frx":0186
         MultiSelect     =   2  '進階多重選取
         Sorted          =   -1  'True
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   270
         Width           =   8790
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   345
         Index           =   2
         Left            =   300
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   4230
         Width           =   675
      End
      Begin VB.CommandButton cmdSaveAtt 
         Caption         =   "下載"
         Height          =   345
         Index           =   2
         Left            =   1770
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   4230
         Width           =   675
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "新增"
         Height          =   345
         Index           =   2
         Left            =   3210
         TabIndex        =   117
         Top             =   4230
         Width           =   675
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "刪除"
         Height          =   345
         Index           =   2
         Left            =   3930
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   4230
         Width           =   675
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全選"
         Height          =   345
         Index           =   2
         Left            =   1050
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   4230
         Width           =   675
      End
      Begin VB.CommandButton CmdClose 
         BackColor       =   &H00FFFFC0&
         Caption         =   "-><-"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1050
         Style           =   1  '圖片外觀
         TabIndex        =   114
         Top             =   -30
         Width           =   640
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "\\typing2\電子送件暫存區"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   5040
         TabIndex        =   124
         Top             =   3810
         Visible         =   0   'False
         Width           =   1910
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "POA暫存區路徑："
      Height          =   180
      Index           =   3
      Left            =   2280
      TabIndex        =   69
      Top             =   6180
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TE暫存區路徑："
      Height          =   180
      Index           =   2
      Left            =   2430
      TabIndex        =   64
      Top             =   5850
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   59
      Top             =   6480
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   180
      Index           =   5
      Left            =   3180
      TabIndex        =   55
      Top             =   240
      Width           =   930
   End
   Begin VB.Label lblCP10 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4140
      TabIndex        =   54
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   180
      Index           =   4
      Left            =   0
      TabIndex        =   53
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblCP09 
      Height          =   180
      Left            =   990
      TabIndex        =   52
      Top             =   240
      Width           =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   19
      Left            =   0
      TabIndex        =   38
      Top             =   30
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   18
      Left            =   0
      TabIndex        =   37
      Top             =   450
      Width           =   960
   End
   Begin VB.Label lblCaseNo 
      Height          =   180
      Left            =   990
      TabIndex        =   36
      Top             =   30
      Width           =   1350
   End
   Begin MSForms.Label lblCaseName 
      Height          =   270
      Left            =   990
      TabIndex        =   35
      Top             =   450
      Width           =   5420
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "9560;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblPA08 
      Height          =   180
      Left            =   4140
      TabIndex        =   34
      Top             =   30
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "專利種類："
      Height          =   180
      Index           =   1
      Left            =   3180
      TabIndex        =   33
      Top             =   30
      Width           =   930
   End
End
Attribute VB_Name = "frm090202_4_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/24 Form2.0已修改
'Create by Sindy 2013/5/1
Option Explicit

'變數宣告區
Public m_EEP01 As String '總收文號
Public m_AttEEP02 As String '序號
Public m_ProState As String 'P,CFP,T,A(其他),FCP,FCT,CFT
Public m_NPManKind As String '程序人員種類：1.台灣案 2.非台灣案 3.非台灣案歸檔 4.待轉檔區
                             '              空白.CFP 或 其他
Dim strSubject As String, strContent As String
Dim ii As Integer, jj As Integer
Dim dblPrevRow As Double
Dim m_PrevForm As Form '前一畫面
Dim m_EPMan As String '承辦人
Dim strEEP10_Err As String, strEEP05_Err As String
Dim m_CP06 As String, m_CP07 As String, m_CP10 As String, m_CP13 As String, m_CP14 As String, m_CP18 As String
Dim m_EP06 As String, m_EP09 As String, m_EP07 As String, m_EP08 As String, m_EP35 As String
Dim m_CP14_2 As String 'Add By Sindy 2017/6/15 外翻人員在所內處理的人員可能一個以上,第2個以上為副本收受者
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
Dim m_CP118 As String, m_EP01 As String
Dim m_PA09 As String, m_PA11 As String
Dim strPrinter As String
Dim bolhaveEfile As Boolean '有承辦單
Dim bolStarHasWorkSheet As Boolean

'附件宣告區
Dim m_AttachPath As String
Dim m_FilesRemoved() As String
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Dim m_CPM26 As String 'Add By Sindy 2018/11/20

Dim oStrA06 As String                   '本所期限
Dim oStrA07 As String                   '法定期限
Dim oStrAFile As String                 '檔名
Dim oStrA10 As String                   '支援承辦時數
Dim oStrA11 As String                   '智權人員
Dim oStrA14 As String                   '工程師代碼
Dim oStrEP06 As String                  '齊備日
Dim oStrEP09 As String                  '完稿日
Dim oStrEP07 As String                  '會稿日
Dim oStrEP08 As String                  '會回日
Dim m_EED06 As String, m_EED07 As String, oStrEED07 As String

'列印承辦單
Dim IsHaveTaieLogo As Boolean
Dim i As Integer
Dim DrawCount As Integer
Dim DrawLeftMove As Integer   '左邊位移
Dim DrawRightMove As Integer    '右邊位移
Dim dblLine As Double, dblStarLine As Double
Dim dblMaxLine As Double
Dim PrintPage As Integer

'Dim str_SendAttEEP02 As String 'Add By Sindy 2014/1/17 記錄發文的附件順序為那一筆
Public cmdState As Integer, bolQuery As Boolean '紀錄作用按鍵
Dim m_CP140 As String 'Add By Sindy 2015/6/24
Dim m_PA26 As String
Dim m_PA149 As String
Dim m_PA75 As String
Dim m_CP44 As String, m_CP116 As String '代理人,聯絡人編號 Add By Sindy 2015/9/22
Dim m_CP12 As String 'Add By Sindy 2015/10/15
'-----------------------------------------------------------------------
Dim bolPAFlow As Boolean, bolTMFlow As Boolean 'Add By Sindy 2018/5/2
Dim bolOtherFlow As Boolean 'Add By Sindy 2021/9/2
Dim bolFCPFlow As Boolean 'Add By Sindy 2023/9/12
Dim bolCFTFlow As Boolean, bolFCTFlow As Boolean 'Add By Sindy 2024/8/14
'-----------------------------------------------------------------------
Dim bolFMP As Boolean 'Add By Sindy 2023/10/6
Dim bolOurFMP As Boolean '是否寰華案件 Add By Sindy 2023/10/6
Dim pa() As String, sp() As String, tm() As String, cp() As String 'Add By Sindy 2018/11/19
Dim lC() As String, hc() As String 'Add By Sindy 2021/9/2
Dim m_strFolder As String, m_strCaseNo As String 'Add By Sindy 2018/11/21
Dim m_strPOAFolder As String 'Add By Sindy 2020/3/18
Dim bolHadPOAfile As Boolean '有委任書
Dim m_bolShowEng As Boolean
Dim m_EEP15 As String 'Add By Sindy 2020/9/30 多案總收文號
'Add By Sindy 2023/11/9
Dim m_EEP11 As String '系統備註
Dim m_EEP04 As String '目前的歷程狀態
'2023/11/9END
Dim m_CP157 As String '北所分案日期 Add By Sindy 2023/3/29
'Add By Sindy 2023/9/12
Const intTab_承辦單 As Integer = 1
Const intTab_外專承辦單 As Integer = 2
Const intTab_存卷資料 As Integer = 3
'2023/9/12 END
Const 不需操作程序送判的案件性質 As String = "211,212,408" 'Add By Sindy 2025/6/12 211=準備程序,212=言詞辯論,408=面詢
Dim m_bolFirst As Boolean 'Add By Sindy 2023/12/6
Dim m_EP04 As String 'Add By Sindy 2023/12/6
Dim m_EP12 As String 'Add By Sindy 2024/1/23
Dim m_EP05 As String 'Add By Sindy 2024/8/14
'Add By Sindy 2025/1/15
Public m_FlowUserNum As String '案件流程所屬人員
Dim m_EEP12 As String '代理註明
Dim m_EEP16 As String '原收受者
'2025/1/15 END
Dim strSqlwhere As String 'Add By Sindy 2025/7/29


Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2015/6/15 +ac03
   '                        0       1        2         3        4           5        6         7           8             9           10       11         12      13       14
   arrGridHeadText = Array("順序", "EEP03", "發送者", "EEP04", "流程狀態", "EEP05", "收受者", "送出時間", "副本收受者", "意見內容", "EEP10", "c1.CP43", "ac03", "eep15", "系統備註")
   arrGridHeadWidth = Array(400, 0, 950, 0, 800, 0, 700, 1300, 1000, 3300, 0, 0, 0, 0, 800)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strText As String
Dim bolNotAdd As Boolean
Dim arrID As Variant 'Add By Sindy 2017/6/15
Dim m_strSys As String
   
   QueryData = True
   TextFCPNote(0).Visible = False 'Add By Sindy 2023/11/22
'   bolNotAdd = False
   '清空及預設欄位值
   GRD1.Clear
   m_CP118 = Empty
   m_PA09 = Empty '申請國家
   m_PA11 = Empty '申請案號
   SetGrd
   lblCaseNo.Caption = Empty
   lblPA08.Caption = Empty
   lblCaseName.Caption = Empty
   Call ClearData
   Call SetCtrlReadOnly(False)
   
   'Modify By Sindy 2018/11/19
   '進度檔
   cp(9) = m_EEP01
   Call PUB_ReadCaseProgressDatabase(cp(), 國外_CF)
   '2018/11/19 END
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   'Modify By Sindy 2023/12/6 +,EP04,EP05
   '案件資料
   'Modify By Sindy 2024/1/23 +,EP12
   strSql = "Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱,NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,PA09,CP140,pa26,pa149,pa75,PA11,CP44,CP116,CP12,CP157,EP04,EP05,EP12" & _
            " From CaseProgress,EngineerProgress,Patent,staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap" & _
            " Where CP09='" & m_EEP01 & "' And CP09=EP02(+) And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And PA09=NA01(+) And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '1'=PTM01(+) AND PA08=PTM02(+)"
   'Add By Sindy 2015/10/21 +服務
   strSql = strSql & "union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱,NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,SP09,CP140,SP08,SP78,SP26,SP11,CP44,CP116,CP12,CP157,EP04,EP05,EP12" & _
            " From CaseProgress,EngineerProgress,Servicepractice,staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "' And CP09=EP02(+) And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And SP09=NA01(+) And CP01=CPM01(+) And CP10=CPM02(+)"
   'Add By Sindy 2018/5/2 +商標
   'Modify By Sindy 2018/10/11 Decode(tm10,'000',PTM03,PTM04) ==> TM09: 顯示商品類別
   strSql = strSql & "union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,tm05||tm06||tm07 as 案件名稱,NA03 as 國家,TM09 as 種類,Decode(tm10,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,tm10,CP140,tm23,tm123,tm44,tm12,CP44,CP116,CP12,CP157,EP04,EP05,EP12" & _
            " From CaseProgress,EngineerProgress,Trademark,staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap" & _
            " Where CP09='" & m_EEP01 & "' And CP09=EP02(+) And CP01=tm01 And CP02=tm02 And CP03=tm03 And CP04=tm04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And tm10=NA01(+) And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '2'=PTM01(+) AND tm08=PTM02(+)"
   'Add By Sindy 2021/9/2 + 法務
   strSql = strSql & "union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,LC05||LC06||LC07 as 案件名稱,NA03 as 國家,'' as 種類,Decode(LC15,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,LC15,CP140,LC11,LC42,LC22,'' as SP11,CP44,CP116,CP12,CP157,EP04,EP05,EP12" & _
            " From CaseProgress,EngineerProgress,Lawcase,staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "' And CP09=EP02(+) And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And LC15=NA01(+) And CP01=CPM01(+) And CP10=CPM02(+)"
   '+ 顧問
   strSql = strSql & "union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,HC06 as 案件名稱,NA03 as 國家,'' as 種類,Decode('000','000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,'000',CP140,HC05,HC23,'' as SP26,'' as SP11,CP44,CP116,CP12,CP157,EP04,EP05,EP12" & _
            " From CaseProgress,EngineerProgress,Hirecase,staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "' And CP09=EP02(+) And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And '000'=NA01(+) And CP01=CPM01(+) And CP10=CPM02(+)"
   '2021/9/2 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_CP01 = Empty
      If Not IsNull(rsTmp.Fields("CP01")) Then m_CP01 = rsTmp.Fields("CP01")
      m_CP02 = Empty
      If Not IsNull(rsTmp.Fields("CP02")) Then m_CP02 = rsTmp.Fields("CP02")
      m_CP03 = Empty
      If Not IsNull(rsTmp.Fields("CP03")) Then m_CP03 = rsTmp.Fields("CP03")
      m_CP04 = Empty
      If Not IsNull(rsTmp.Fields("CP04")) Then m_CP04 = rsTmp.Fields("CP04")
      m_EP12 = "" & rsTmp.Fields("EP12") 'Add By Sindy 2024/1/23
      
      'Add By Sindy 2018/5/2
      m_strSys = CheckSys(m_CP01)
      If InStr("1", m_strSys) > 0 Then '專利
         pa(1) = m_CP01
         pa(2) = m_CP02
         pa(3) = m_CP03
         pa(4) = m_CP04
         If ClsPDReadPatentDatabase(pa(), 國外_CF) = True Then
            bolPAFlow = True
         End If
      ElseIf InStr("2", m_strSys) > 0 Then '商標
         tm(1) = m_CP01
         tm(2) = m_CP02
         tm(3) = m_CP03
         tm(4) = m_CP04
         If ClsPDReadTrademarkDatabase(tm(), 國外_CF) = True Then
            'bolTMFlow = True
            'Add By Sindy 2024/8/14
            If m_CP01 = "CFT" Then
               bolCFTFlow = True
            ElseIf m_CP01 = "FCT" And Left(PUB_GetST03(cp(14)), 1) = "F" Then '因FCT爭議案件是內商人員在承辦的
               bolFCTFlow = True
            Else
               bolTMFlow = True
            End If
            '2024/8/14 END
         End If
      ElseIf InStr("5,6", m_strSys) > 0 Then
         sp(1) = m_CP01
         sp(2) = m_CP02
         sp(3) = m_CP03
         sp(4) = m_CP04
         If ClsPDReadServicePracticeDatabase(sp(), 國外_CF) = True Then
            If m_strSys = "6" Then '商標:服務
               'bolTMFlow = True
               'Add By Sindy 2024/8/14
               If m_CP01 = "CFC" Or (m_CP01 = "S" And sp(9) <> "000") Then
                  bolCFTFlow = True
               ElseIf (m_CP01 = "S" And sp(9) = "000") Then
                  bolFCTFlow = True
               Else
                  bolTMFlow = True
               End If
               '2024/8/14 END
            Else
               bolPAFlow = True
            End If
         End If
      'Add By Sindy 2021/9/2
      ElseIf InStr("3,7", m_strSys) > 0 Then '法務
         lC(1) = m_CP01
         lC(2) = m_CP02
         lC(3) = m_CP03
         lC(4) = m_CP04
         If ClsPDReadLawCaseDatabase(lC()) = True Then
            bolOtherFlow = True
         End If
      ElseIf InStr("4,8", m_strSys) > 0 Then '顧問
         hc(1) = m_CP01
         hc(2) = m_CP02
         hc(3) = m_CP03
         hc(4) = m_CP04
         If ClsPDReadHireCaseDatabase(hc()) = True Then
            bolOtherFlow = True
         End If
      '2021/9/2 END
      Else 'If InStr("3,4,7,8", m_strSys) > 0 Then '其他
         Screen.MousePointer = vbDefault
         MsgBox "讀取系統類別有誤，請洽電腦中心！", vbExclamation
         QueryData = False
         rsTmp.Close
         Set rsTmp = Nothing
         Call cmdExit_Click
         Exit Function
      End If
      '2018/5/2 END
      
      If Not IsNull(rsTmp.Fields("本所案號")) Then lblCaseNo.Caption = rsTmp.Fields("本所案號")
      If Not IsNull(rsTmp.Fields("種類")) Then lblPA08.Caption = rsTmp.Fields("種類")
'      If Not IsNull(rsTmp.Fields("國家")) Then lblPA09.Caption = rsTmp.Fields("國家")
      If Not IsNull(rsTmp.Fields("案件名稱")) Then lblCaseName.Caption = rsTmp.Fields("案件名稱")
      
      If Not IsNull(rsTmp.Fields("PA09")) Then m_PA09 = rsTmp.Fields("PA09")
      If Not IsNull(rsTmp.Fields("PA11")) Then m_PA11 = rsTmp.Fields("PA11")
      If Not IsNull(rsTmp.Fields("CP44")) Then m_CP44 = rsTmp.Fields("CP44") '代理人 Add By Sindy 2015/9/22
      If Not IsNull(rsTmp.Fields("CP116")) Then m_CP116 = rsTmp.Fields("CP116") '聯絡人編號 Add By Sindy 2015/9/23
      If Not IsNull(rsTmp.Fields("CP157")) Then m_CP157 = rsTmp.Fields("CP157") '北所分案日期 Add By Sindy 2023/3/29
      If Not IsNull(rsTmp.Fields("EP04")) Then m_EP04 = "" & rsTmp.Fields("EP04") 'Add By Sindy 2023/12/6
      If Not IsNull(rsTmp.Fields("EP05")) Then m_EP05 = "" & rsTmp.Fields("EP05") 'Add By Sindy 2024/8/14
      
      m_CP140 = "" & rsTmp.Fields("CP140") 'Add By Sindy 2015/6/24
      '電子送件
      If Not IsNull(rsTmp.Fields("CP118")) Then
         m_CP118 = rsTmp.Fields("CP118")
      'Add by Sindy 2013/9/27
         lblEApp.Visible = True
      Else
         lblEApp.Visible = False
      End If
      '2013/9/27 END
            
      If bolPAFlow = True Then
         Label4.Visible = True '註:電子送件，請先加入下載的檔案後，再執行產生承辦單。
         'Add By Sindy 2016/10/18 一案兩請
         strSql = "select cm05,cm06,cm07,cm08 from casemap where cm01='" & m_CP01 & "' and cm02='" & m_CP02 & "' and cm03='" & m_CP03 & "' and cm04='" & m_CP04 & "' and cm10='3'" & _
                  " Union select cm01,cm02,cm03,cm04 from casemap where cm05='" & m_CP01 & "' and cm06='" & m_CP02 & "' and cm07='" & m_CP03 & "' and cm08='" & m_CP04 & "' and cm10='3'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         lblCM10.Tag = ""
         If intI = 1 Then
            lblCM10.Visible = True
            lblCM10.Tag = RsTemp.Fields(0) & RsTemp.Fields(1) & IIf(RsTemp.Fields(2) & RsTemp.Fields(3) <> "000", "-" & RsTemp.Fields(2) & "-" & RsTemp.Fields(3), "")
         Else
            lblCM10.Visible = False
         End If
         '2016/10/18 END
         
         'Add By Sindy 2023/9/12
         bolFMP = PUB_ChkIsFMP(m_CP01, m_CP02, m_CP03, m_CP04, pa(9))
         If bolFMP = True Then
            bolOurFMP = PUB_FMPtoCheck(1, 2, PUB_GetST05(cp(14)), m_CP01, m_CP02, m_CP03, m_CP04) '是否寰華案件
         Else
            bolOurFMP = False
         End If
         If m_CP01 = "FCP" Or _
            m_CP01 = "FG" Or _
            (bolFMP = True And Left(PUB_GetST03(cp(14)), 1) = "F") Then
            bolPAFlow = False
            bolFCPFlow = True '外專Flow
            SSTab1.TabVisible(intTab_承辦單) = False
            SSTab1.TabVisible(intTab_外專承辦單) = True
            'Add By Sindy 2024/4/17
            If Pub_StrUserSt03 = "F22" Or Pub_StrUserSt03 = "M51" Then
               cmdOutlook.Visible = True
               CmdCalendar.Visible = True 'Add By Sindy 2025/10/15
            End If
            '2024/4/17 END
         Else
            SSTab1.TabVisible(intTab_外專承辦單) = False
         End If
         '2023/9/12 END
         
      'Add By Sindy 2018/5/10
      'Modify By Sindy 2021/9/2
      Else
         Label4.Visible = False
         SSTab1.TabVisible(intTab_承辦單) = False
         SSTab1.TabVisible(intTab_外專承辦單) = False 'Add By Sindy 2023/9/12
         'Modify By Sindy 2024/8/14 + Or bolCFTFlow = True Or bolFCTFlow = True
         If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
            Label1(1).Caption = "類別："
         Else
            Label1(1).Visible = False
            '2021/9/2 END
         End If
      '2018/5/10 END
      End If
      
      'Add By Sindy 2015/9/4 P非台灣案
      'Modify By Sindy 2015/9/30
      'If m_CP01 = "P" And m_PA09 <> "000" Then
      Me.Height = 6165 'Add By Sindy 2018/11/19
      'Modify By Sindy 2021/5/13 + Or m_CP01 = "PS"
      If (m_CP01 = "P" Or m_CP01 = "PS") And (m_NPManKind = "2" Or m_NPManKind = "3") Then
      '2015/9/30 END
         'Add By Sindy 2025/2/13 針對以下三道程序修改為工程師跑歷程送件後，程序人員發文時，樞紐僅做歸檔即可
         '   209(檢視中說)
         '   942(檢視PCT公開本與FCP相異處)
         '   201(新案翻譯)
         If bolFCPFlow = True And InStr("201,209,942", cp(10)) > 0 Then
            cmdPrintAllPDF.Caption = "產生承辦單及歸檔"
         Else
         '2025/2/13 END
            cmdPrintAllPDF.Caption = "E-Mail及歸檔"
         End If
      'Add By Sindy 2018/11/19 + 4.待轉檔區
      ElseIf m_CP01 = "P" And m_NPManKind = "4" Then
         '待轉檔區資料夾
         'Modify By Sindy 2021/6/3 ex:P-82513
'         m_strCaseNo = Trim(pa(1)) & Val(Trim(pa(2))) & _
'                       IIf(Val(Trim(pa(3))) = 0 And Val(Trim(pa(4))) = 0, "", "-" & pa(3)) & _
'                       IIf(Val(Trim(pa(4))) = 0, "", "-" & Format(pa(4), "00"))
         m_strCaseNo = Trim(pa(1)) & Trim(pa(2)) & _
                       IIf(Val(Trim(pa(3))) = 0 And Val(Trim(pa(4))) = 0, "", "-" & pa(3)) & _
                       IIf(Val(Trim(pa(4))) = 0, "", "-" & Format(pa(4), "00"))
         '2021/6/3 END
         If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
            m_strFolder = PUB_Getdesktop
            m_strPOAFolder = PUB_Getdesktop & "\POA\"
         Else
            'Modify By Sindy 2022/10/25 改用常變數 str_P_台灣電子送件檔案路徑
            m_strFolder = str_P_台灣電子送件檔案路徑 'Text1.Text
            'Modify By Sindy 2022/10/25 改用常變數 str_P_OrderPath
            m_strPOAFolder = str_P_OrderPath & "\POA" & "\" 'Text3.Text & "\"
         End If
         m_strFolder = m_strFolder & "\" & m_strCaseNo & "\"
         'Add By Sindy 2021/11/17
         Text1.Text = m_strFolder
         Text3.Text = m_strPOAFolder
         '2021/11/17 END
         
         cmdPrintAllPDF.Caption = "產生送件資料夾"
         Me.Height = 6825
      Else
         If bolPAFlow = True Then
            cmdPrintAllPDF.Caption = "產生承辦單及歸檔"
         Else
            cmdPrintAllPDF.Caption = "歸　檔"
            'Add By Sindy 2025/10/28
            If m_ProState = "FCP" Then
               Me.Height = 6520
            End If
            '2025/10/28 END
         End If
      End If
      m_PA26 = Empty
      If Not IsNull(rsTmp.Fields("PA26")) Then m_PA26 = rsTmp.Fields("PA26")
      m_PA149 = Empty
      If Not IsNull(rsTmp.Fields("PA149")) Then m_PA149 = rsTmp.Fields("PA149")
      m_PA75 = Empty
      If Not IsNull(rsTmp.Fields("PA75")) Then m_PA75 = rsTmp.Fields("PA75")
      '2015/9/4 END
      
      'Add By Sindy 2015/10/15
      m_CP12 = Empty
      If Not IsNull(rsTmp.Fields("CP12")) Then m_CP12 = rsTmp.Fields("CP12")
      '2015/10/15 END
      m_CP10 = Empty
      If Not IsNull(rsTmp.Fields("CP10")) Then m_CP10 = rsTmp.Fields("CP10")
      m_CP14 = Empty
      If Not IsNull(rsTmp.Fields("CP14")) Then m_CP14 = rsTmp.Fields("CP14")
      m_CP06 = Empty
      If Not IsNull(rsTmp.Fields("CP06")) Then m_CP06 = rsTmp.Fields("CP06")
      If m_CP06 <> "" Then
         oStrA06 = "  " & Val(Left(m_CP06, 4)) - 1911 & "年  " & Mid(m_CP06, 5, 2) & "月  " & Right(m_CP06, 2) & "日"
      Else
         oStrA06 = "    年    月    日"
      End If
      m_CP07 = Empty
      If Not IsNull(rsTmp.Fields("CP07")) Then m_CP07 = rsTmp.Fields("CP07")
      If m_CP07 <> "" Then
         oStrA07 = "  " & Val(Left(m_CP07, 4)) - 1911 & "年  " & Mid(m_CP07, 5, 2) & "月  " & Right(m_CP07, 2) & "日"
      Else
         oStrA07 = "    年    月    日"
      End If
      '檔名
      oStrAFile = ""
      strExc(0) = "SELECT cpm26 FROM casepropertymap WHERE cpm01='" & rsTmp.Fields("CP01") & "' and cpm02='" & m_CP10 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If "" & RsTemp.Fields("cpm26") <> "" Then
            oStrAFile = rsTmp.Fields("CP01") & Trim(Val(rsTmp.Fields("CP02"))) & "." & Trim(RsTemp.Fields("cpm26"))
            m_CPM26 = Trim(RsTemp.Fields("cpm26")) 'Add By Sindy 2018/11/20
         End If
      End If
      
      '智權人員前加所別
      m_CP13 = Empty
      If Not IsNull(rsTmp.Fields("CP13")) Then m_CP13 = rsTmp.Fields("CP13")
      oStrA11 = ""
      If m_CP13 <> "" Then
         Select Case PUB_GetST06(m_CP13)
            Case "2"
               oStrA11 = "中所"
            Case "3"
               oStrA11 = "南所"
            Case "4"
               oStrA11 = "高所"
         End Select
         oStrA11 = oStrA11 & rsTmp.Fields("智權人員")
      End If
      '收文點數
      m_CP18 = Empty
      If Not IsNull(rsTmp.Fields("CP18")) Then m_CP18 = rsTmp.Fields("CP18")
      '齊備日
      m_EP06 = Empty
      If Not IsNull(rsTmp.Fields("EP06")) Then m_EP06 = rsTmp.Fields("EP06")
      If m_EP06 <> "" Then
         oStrEP06 = "  " & Val(Left(m_EP06, 4)) - 1911 & "年  " & Mid(m_EP06, 5, 2) & "月  " & Right(m_EP06, 2) & "日"
      Else
         oStrEP06 = "    年    月    日"
      End If
      '完稿日
      m_EP09 = Empty
      If Not IsNull(rsTmp.Fields("EP09")) Then m_EP09 = rsTmp.Fields("EP09")
      If m_EP09 <> "" Then
         oStrEP09 = "  " & Val(Left(m_EP09, 4)) - 1911 & "年  " & Mid(m_EP09, 5, 2) & "月  " & Right(m_EP09, 2) & "日"
      Else
         oStrEP09 = "    年    月    日"
      End If
      '會稿日
      m_EP07 = Empty
      If Not IsNull(rsTmp.Fields("EP07")) Then m_EP07 = rsTmp.Fields("EP07")
      If m_EP07 <> "" Then
         oStrEP07 = "  " & Val(Left(m_EP07, 4)) - 1911 & "年  " & Mid(m_EP07, 5, 2) & "月  " & Right(m_EP07, 2) & "日"
      Else
         oStrEP07 = "    年    月    日"
      End If
      '會回日
      m_EP08 = Empty
      If Not IsNull(rsTmp.Fields("EP08")) Then m_EP08 = rsTmp.Fields("EP08")
      If m_EP08 <> "" Then
         oStrEP08 = "  " & Val(Left(m_EP08, 4)) - 1911 & "年  " & Mid(m_EP08, 5, 2) & "月  " & Right(m_EP08, 2) & "日"
      Else
         oStrEP08 = "    年    月    日"
      End If
      '承辦天數
      m_EP35 = Empty
      oStrA10 = "  天"
      If Not IsNull(rsTmp.Fields("EP35")) Then
         m_EP35 = rsTmp.Fields("EP35")
         oStrA10 = m_EP35 & " 天"
      End If
      
      lblCP09 = "" & rsTmp.Fields("CP09")
      lblCP10 = "" & rsTmp.Fields("案件性質")
      
      m_EP01 = "" & rsTmp.Fields("EP01")
      '承辦人
      m_EPMan = ""
      m_CP14_2 = ""
      If m_CP14 <> "" Then
         'Modify By Sindy 2024/8/14
         If bolCFTFlow = True Or bolFCTFlow = True Then
            m_EPMan = m_EP05 & " " & GetPrjSalesNM(m_EP05)
         '2024/8/14 END
         'Modify By Sindy 2023/12/6
         ElseIf bolFCPFlow = True And cp(10) = "201" And Left(m_CP14, 1) = "F" And m_EP04 <> "" Then '新案翻譯
            m_EPMan = m_EP04 & " " & GetPrjSalesNM(m_EP04)
         '2023/12/6 END
         '若為外翻人員
         ElseIf Left(m_CP14, 1) = "F" Then
            strText = Trim(PUB_GetST14(m_CP14))
            If strText <> "" Then
               'Remove by Lydia 2017/03/28
               'm_EPMan = strText & " " & GetPrjSalesNM(strText)
            Else
               strText = Trim(Pub_GetSpecMan("H"))
               'Remove by Lydia 2017/03/28
               'If strText <> "" Then
               '   m_EPMan = strText & " " & GetPrjSalesNM(strText)
               'End If
            End If
            'Added by Lydia 2017/03/28 ST14改成多個編號,所以只抓第一位
            If strText <> "" Then
               If InStr(strText, ",") > 0 Or InStr(strText, ";") > 0 Then
                  'Add By Sindy 2017/6/15
                  strText = Replace(strText, ";", ",")
                  '檢查人員是否在職
                  If InStr(strText, ",") > 0 Then
                     arrID = Split(strText, ",")
                     strText = ""
                     For intI = 0 To UBound(arrID)
                        If ChkStaffST04(CStr(arrID(intI)), False) = False Then
                           strText = strText & "," & arrID(intI)
                        End If
                     Next intI
                     strText = Mid(strText, 2)
                  End If
                  m_CP14_2 = strText '記錄原資料
                  '2017/6/15 END
                  'Added by Lydia 2017/03/28
                  'strText = Replace(Replace(Mid(strText, 1, 6), ",", ""), ";", "")
                  strText = Mid(strText, 1, 5)
                  '2017/03/28 END
                  'Add By Sindy 2017/6/15 解析第一個人員之後的人員資料為副本收受者
                  m_CP14_2 = Mid(m_CP14_2, 7)
                  If InStr(m_CP14_2, strUserNum) > 0 Then
                     m_CP14_2 = Replace(m_CP14_2, strUserNum, "")
                     m_CP14_2 = Replace(m_CP14_2, ",,", ",")
                     If Left(m_CP14_2, 1) = "," Then m_CP14_2 = Mid(m_CP14_2, 2)
                     If Right(m_CP14_2, 1) = "," Then m_CP14_2 = Mid(m_CP14_2, 1, Len(m_CP14_2) - 1)
                  End If
                  '2017/6/15 END
               End If
               'end 2017/03/28
               m_EPMan = strText & " " & GetPrjSalesNM(strText)
            End If
            'end 2017/03/28
         Else
            m_EPMan = m_CP14 & " " & GetPrjSalesNM(m_CP14)
         End If
      End If
            
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      QueryData = False
      rsTmp.Close
      Set rsTmp = Nothing
      Call cmdExit_Click
      Exit Function
   End If
   rsTmp.Close
   
   'Add By Sindy 2025/1/15 代理註明
   If m_FlowUserNum <> strUserNum Then
      m_EEP12 = "(代)"
      m_EEP16 = m_FlowUserNum
   Else
      m_EEP12 = ""
      m_EEP16 = ""
   End If
   '2025/1/15 END
   
   lstAtt(0).Height = 1660 'Add By Sindy 2025/2/13
   'Modify By Sindy 2024/8/14
   '外專 及 外商FC 程序人員操作時
   If m_ProState = "FCP" Or bolFCTFlow = True Then
      'If bolFCPFlow = True Or bolFCTFlow = True Then
         Me.cmdFlow(0).Visible = True '程序送判
         Me.cmdFlow(1).Visible = True 'Add By Sindy 2023/12/6 聯絡
      'End If
      TextFCPNote(0).Visible = True '*** 重要 ***
      'Add By Sindy 2023/12/6
      If m_bolFirst = False Then
         lstAtt(0).Height = lstAtt(0).Height - (TextFCPNote(0).Height - 100)
         m_bolFirst = True
      End If
      '2023/12/6 END
      LblinfoNote.Visible = False '存卷區的加註不顯示,不鎖info
      
   'Add By Sindy 2025/1/22 CFT但由程序人員要操作發文的案件
   ElseIf PUB_GetST03(strUserNum) = "F12" Then
      Me.cmdFlow(0).Visible = True '程序送判
      Me.cmdFlow(1).Visible = True '聯絡
   End If
   '2024/8/14 END
   
   Call ReadSmailBackup 'Add By Sindy 2015/9/15 寄件日期
   
   Call QueryEmpElectronData '承辦單 Add By Sindy 2023/11/3
   
   'Add By Sindy 2015/1/21 一進此作業時,是否已有承辦單
   'Modify By Sindy 2020/9/29 + EMP_多案承辦單
   bolStarHasWorkSheet = False
   'Modify By Sindy 2023/4/25
   'If Left(m_CP01, 1) = "T" Or m_CP01 = "FCT" Then
   'Modify By Sindy 2024/8/14 + Or bolCFTFlow = True Or bolFCTFlow = True
   If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
      strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & m_EEP01 & "'" & _
                  " and (instr(upper(cpp02),upper('" & EMP_承辦單 & ".menu'))>0 or instr(upper(cpp02),upper('" & EMP_多案承辦單 & ".menu'))>0)"
   ElseIf bolPAFlow = True Then
   '2023/4/25 END
      strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & m_EEP01 & "'" & _
                  " and instr(upper(cpp02),upper('" & EMP_承辦單 & "'))>0"
   'Add By Sindy 2023/11/9
   Else 'If bolFCPFlow = True Then
      strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & m_EEP01 & "'" & _
                  " and instr(upper(cpp02),upper('" & EMP_承辦單 & ".menu'))>0"
      '2023/11/9 END
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      bolStarHasWorkSheet = True
   End If
   '2015/1/21 END
   
   '承辦電子簽核資料
   '待送件區時,只讀取最後一筆判發資料
   strSql = "Select distinct EEP02 as 順序,EEP03,s1.ST02||eep12 as 發送者,EEP04,decode(eep04,'" & EMP_附加流程 & "',decode(c2.CP43,'',ac03,Decode(" & m_PA09 & ",'000',CPM03,CPM04)),ac03) as 流程狀態,EEP05,decode(s2.ST02,null,eep05,s2.ST02) as 收受者,sqldatet(EEP06)||' ' ||sqltime(EEP07) as 送出時間,EEP10 as 副本收受者,EEP08 as 意見內容,EEP10,c1.CP43,ac03,eep15,eep11 as 系統備註" & _
            " From EmpElectronProcess,staff s1,staff s2,allcode,caseprogress c1,caseprogress c2,casepropertymap" & _
            " Where EEP01='" & m_EEP01 & "'" & _
            " And EEP03=s1.ST01(+) And EEP05=s2.ST01(+)" & _
            " And ac01='09' And EEP04=ac02(+)" & _
            " And eep01=c2.cp43(+) And eep06=c2.cp05(+)" & _
            " And c2.cp01=cpm01(+) And c2.cp10=cpm02(+)" & _
            " And EEP01=c1.cp09(+)" & _
            " order by EEP02 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      For ii = 1 To GRD1.Rows - 1
         txtEEP10_2 = GRD1.TextMatrix(ii, 10)
         'Add By Sindy 2020/9/30 多案總收文號
         If GRD1.TextMatrix(ii, 0) = m_AttEEP02 Then
            m_EEP15 = GRD1.TextMatrix(ii, 13)
            'Add By Sindy 2023/11/9
            m_EEP11 = GRD1.TextMatrix(ii, 14)
            m_EEP04 = GRD1.TextMatrix(ii, 3)
            '2023/11/9 END
         End If
         '2020/9/30 END
         Call txtEEP10_2_LostFocus
         GRD1.TextMatrix(ii, 8) = txtEEP10_2
         '判斷有相關總收文號才做案件性質轉換
         If GRD1.TextMatrix(ii, 4) = "附加流程" Then
            If GRD1.TextMatrix(ii, 11) <> "" Then
               GRD1.TextMatrix(ii, 4) = Trim(lblCP10) & PUB_GetRelateCasePropertyName(m_EEP01, "1")
            End If
         End If
      Next ii
'      '若有資料游標停在第一筆
'      GRD1.Visible = False
'      GRD1.col = 0
'      GRD1.row = 1
'      dblPrevRow = GRD1.row
'      If rsTmp.RecordCount > 0 Then
'         For ii = 0 To GRD1.Cols - 1
'            GRD1.col = ii
'            GRD1.CellBackColor = &HFFC0C0
'         Next ii
      
'      'Add By Sindy 2014/1/17 記錄發文的附件順序為那一筆
'      '待送件區時,附件檔則讀取判發最後一筆有附件的資料
'      'Modify By Sindy 2013/9/10 +EMP_退件重送
'      strExc(0) = "select eep01,eep02,eep03,eep04 from EmpElectronProcess,EmpElectronFile" & _
'                  " where eep01='" & m_EEP01 & "'" & _
'                  " and eep04 in('" & EMP_判發 & "','" & EMP_退件重送 & "')" & _
'                  " and eep01=eef01(+) and eep02=eef02(+) " & _
'                  " order by eep02 desc "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         If RsTemp.RecordCount > 0 Then
'            str_SendAttEEP02 = "" & RsTemp.Fields("eep02")
'         End If
'      End If
'      '2014/1/17 END
      
         Call ReadData(True)
'      End If
'      GRD1.Visible = True
   End If
   rsTmp.Close
   
   Call ReadAttachFile_other(m_EEP01) 'Add By Sindy 2013/9/25 查詢存卷區
   
   'Add By Sindy 2018/5/23
   If Label4.Visible = True Then  '註:電子送件，請先加入下載的檔案後，再執行產生承辦單。
      txtEEP08.Height = 1515
   'Modify By Sindy 2024/2/1
   ElseIf bolFCPFlow = True Then
      txtEEP08.Height = 1515
      Label25(0).Visible = True
      Label25(1).Visible = True 'Add By Sindy 2025/10/28
      cmdOK(4).Caption = "完整卷宗"
   '2024/2/1 END
   Else
      txtEEP08.Height = 1785
   End If
   '2018/5/23 END
   
   Call SetTxtLpNote(True) 'Add By Sindy 2020/9/29
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2023/11/3
'讀取承辦單
Private Sub QueryEmpElectronData()
Dim rsTmp As New ADODB.Recordset
Dim objText As Object
Dim Rs As New ADODB.Recordset 'Add By Sindy 2024/1/17
   
   '外專
   If bolFCPFlow = True Then
      '先清除承辦單內容
      For Each objText In Me.txt3
         objText.Text = ""
         objText.Tag = ""
      Next
      txtEP12.Text = ""
      txtEP12.Tag = ""
      ChkEED13.Tag = ""
      LblEED10_N.Caption = "": LblEED10_N_2.Caption = ""
      LblEED06_N.Caption = ""
      LblEED09_N.Caption = ""
      'Modify By Sindy 2024/1/17
      CmbFL(3).Clear
      CmbFL(3).Visible = False
      '2024/1/17 END
      
      'Add By Sindy 2025/4/7 945=電話聯絡單
'      ChkEED08.Visible = False
'      ChkEED08.Tag = ""
      Me.Frame945.Visible = False: Frame945.Tag = "" '預設值
      Me.Frame201.Visible = True '預設值
'      If cp(10) = "945" And _
'         (PUB_ChkEmpFlowExists(m_EEP01, EMP_送判) = True Or PUB_ChkEmpFlowExists(m_EEP01, EMP_發文歸檔) = True) Then
'         ChkEED08.Visible = True
      '當告代掛相關收文號為電話連絡單,增加可以輸入【管制下一程序期限】
'      Else
      If cp(10) = 告知代理人 And cp(43) <> "" Then
         strSql = "Select *" & _
                  " From caseprogress" & _
                  " Where cp09='" & cp(43) & "' and cp10='945'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            Me.Frame945.Visible = True: Frame945.Tag = "945"
            Me.Frame201.Visible = False
         End If
      End If
      '2025/4/7 END
      'Add By Sindy 2025/8/20
      If Frame945.Tag = "" Then
         If cp(10) = 告知代理人 Or cp(10) = 回覆代理人 Then
            Me.Frame945.Visible = True: Frame945.Tag = 告知代理人
            Me.Frame201.Visible = False
            Frame945.Caption = "【管制行事曆期限】"
            Label1(11).Caption = "追蹤客戶指示【管制日期】："
            Label1(12).Visible = False
            txtEED15.Visible = False
         End If
      End If
      '2025/8/20 END
      
      '讀取承辦單內容
      strSql = "Select *" & _
               " From EmpElectronData" & _
               " Where EED01='" & m_EEP01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         '譯者
         If Not IsNull(RsTemp.Fields("EED10")) Then
            txt3(3).Text = RsTemp.Fields("EED10")
            txt3(3).Tag = RsTemp.Fields("EED10")
            Call TXT3_LostFocus(3)
         Else
            Frame7.Visible = False '譯者
         End If
         '打字室
         If Not IsNull(RsTemp.Fields("EED06")) Then
            txt3(6).Text = RsTemp.Fields("EED06")
            txt3(6).Tag = RsTemp.Fields("EED06")
            Call TXT3_LostFocus(6)
         End If
         If Not IsNull(RsTemp.Fields("EED05")) Then
            txt3(4).Text = RsTemp.Fields("EED05")
            txt3(4).Tag = RsTemp.Fields("EED05")
         End If
         '管制人
         If Not IsNull(RsTemp.Fields("EED09")) Then
            txt3(5).Text = RsTemp.Fields("EED09")
            txt3(5).Tag = RsTemp.Fields("EED09")
            Call TXT3_LostFocus(5)
         End If
         If Not IsNull(RsTemp.Fields("EED11")) Then '請款備註
            txt3(7).Text = RsTemp.Fields("EED11")
            txt3(7).Tag = RsTemp.Fields("EED11")
         End If
         '檔案名稱
         If Not IsNull(RsTemp.Fields("EED12")) Then
            txt3(8).Text = RsTemp.Fields("EED12")
            txt3(8).Tag = RsTemp.Fields("EED12")
         End If
         '轉檔後送回
         If Not IsNull(RsTemp.Fields("EED13")) Then
            ChkEED13.Value = 1
         Else
            ChkEED13.Value = 0
         End If
         ChkEED13.Tag = ChkEED13.Value
         'Add By Sindy 2025/4/7
         '需收文告代
'         If Not IsNull(RsTemp.Fields("EED08")) Then
'            ChkEED08.Value = 1
'         Else
'            ChkEED08.Value = 0
'         End If
'         ChkEED08.Tag = ChkEED08.Value
         If Not IsNull(RsTemp.Fields("EED14")) Then
            Me.txtEED14.Text = ChangeWStringToTString(RsTemp.Fields("EED14"))
         Else
            Me.txtEED14.Text = ""
         End If
         If Not IsNull(RsTemp.Fields("EED15")) Then
            Me.txtEED15.Text = ChangeWStringToTString(RsTemp.Fields("EED15"))
         Else
            Me.txtEED15.Text = ""
         End If
         '2025/4/7 END
      End If
      
      '檔案名稱
      'Modify By Sindy 2024/1/17 Bobbie:帶客戶提供文件(中說)的檔案名稱
      'Modify By Sindy 2024/1/30 +Or cp(10) = "235" 核對中說格式
      If m_CP10 = "209" Or m_CP10 = "235" Then
         strSql = "Select *" & _
                  " From CustSupportDoc" & _
                  " Where csd01='" & cp(1) & "' and csd02='" & cp(2) & "' and csd03='" & cp(3) & "' and csd04='" & cp(4) & "'" & _
                  " and csd20 is not null order by CSD05 desc" '簡(繁)體中說檔案名稱
         intI = 1
         Set Rs = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            txt3(8).Visible = False
            CmbFL(3).Visible = True
            Call PUB_SetCmbList(CmbFL(3), Rs.Fields("csd20"))
         End If
      End If
      '2024/1/17 END
      
      'Add By Sindy 2024/1/23 承辦備註改顯示作業備註
      txtEP12.Text = m_EP12
      txtEP12.Enabled = True
      txtEP12.Locked = True
      '2024/1/23 END
   '內專
   ElseIf bolPAFlow = True Then
'      '先清除承辦單內容
'      For Each objText In Me.txt1
'         objText.Text = ""
'         objText.Tag = ""
'      Next
'      Me.lblFa.Caption = ""
'      If m_Country = "000" Then '台灣案不顯示代理人
'         Me.Frame2.Visible = False
'      Else
'         Me.Frame2.Visible = True
'      End If
'      '讀取承辦單內容
'      strSql = "Select *" & _
'               " From EmpElectronData" & _
'               " Where EED01='" & m_EEP01 & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If Not IsNull(RsTemp.Fields("EED02")) Then
'            txt1(5).Text = RsTemp.Fields("EED02")
'            txt1(5).Tag = RsTemp.Fields("EED02")
'         End If
'         If Not IsNull(RsTemp.Fields("EED03")) Then
'            txt1(6).Text = RsTemp.Fields("EED03")
'            txt1(6).Tag = RsTemp.Fields("EED03")
'         End If
'         If Not IsNull(RsTemp.Fields("EED04")) Then
'            txt1(0).Text = RsTemp.Fields("EED04")
'            txt1(0).Tag = RsTemp.Fields("EED04")
'         End If
'         If Not IsNull(RsTemp.Fields("EED05")) Then
'            txt1(4).Text = RsTemp.Fields("EED05")
'            txt1(4).Tag = RsTemp.Fields("EED05")
'         End If
'      End If
      
      '讀取承辦單內容
      txt1(5) = Empty
      txt1(6) = Empty
      txt1(0) = Empty
      txt1(4) = Empty
      m_EED06 = Empty
      m_EED07 = Empty
      strSql = "Select *" & _
               " From EmpElectronData,staff" & _
               " Where EED01='" & m_EEP01 & "' and EED06=ST01(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If Not IsNull(rsTmp.Fields("EED02")) Then txt1(5).Text = rsTmp.Fields("EED02")
         If Not IsNull(rsTmp.Fields("EED03")) Then txt1(6).Text = rsTmp.Fields("EED03")
         
         '主旨
         'If Not IsNull(rsTmp.Fields("EED04")) Then txt1(0).Text = rsTmp.Fields("EED04")
         Call GetPaperMain 'Modify By Sindy 2013/10/1 承辦單主旨直接抓DB資料
         
         If Not IsNull(rsTmp.Fields("EED05")) Then txt1(4).Text = rsTmp.Fields("EED05")
         If Not IsNull(rsTmp.Fields("ST02")) Then m_EED06 = rsTmp.Fields("ST02")
         If Not IsNull(rsTmp.Fields("EED07")) Then m_EED07 = rsTmp.Fields("EED07")
         If m_EED07 <> "" Then
            oStrEED07 = "   " & Val(Left(m_EED07, 4)) - 1911 & "年  " & Mid(m_EED07, 5, 2) & "月  " & Right(m_EED07, 2) & "日"
         Else
            oStrEED07 = "     年    月    日"
         End If
      End If
      rsTmp.Close
   End If
   
   Set rsTmp = Nothing
   Set Rs = Nothing
End Sub

'Add By Sindy 2020/10/12
Private Sub SetTxtLpNote(bolQueryStar As Boolean)
   If bolQueryStar = True Then
      txtLpNote.Visible = False
      txtLpNote = ""
      If m_EEP15 <> "" Then
         txtLpNote.Visible = True
         txtLpNote = "(共" & UBound(Split(m_EEP15, ",")) + 1 & "筆)"
      End If
   Else
      txtLpNote.Visible = False
      txtLpNote = ""
   End If
End Sub

'Add By Sindy 2015/9/15 寄件日期
Private Sub ReadSmailBackup()
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String, strCDate As String, strCTime As String
   
   'Modify By Sindy 2018/9/20 + and smb11 is null:非歷程的寄件備份
   'Modify By Sindy 2018/12/24 and smb11 is null => 改用建立人員的部門做判斷 P-121732
   strSql = "Select *" & _
            " From smailbackup,staff" & _
            " Where smb01='" & m_EEP01 & "' and smb12=st01(+)" & _
            " and st03='P12'" & _
            " order by smb02 desc,smb03 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   lblSendMailDt.Visible = False
   If rsTmp.RecordCount > 0 Then
      lblSendMailDt.Visible = True
      strTemp = TAIWANDATE(rsTmp.Fields("smb02"))
      strCDate = Format(strTemp, "###/##/##")
      strTemp = rsTmp.Fields("smb03")
      strCTime = Format(strTemp, "##:##:##")
      lblSendMailDt.Caption = "寄件日期:" & strCDate & " " & strCTime
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

'抓取主旨
Private Sub GetPaperMain()
   If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "105" _
      Or m_CP10 = "109" Or m_CP10 = "110" Or m_CP10 = "112" Or m_CP10 = "113" _
      Or m_CP10 = "114" Or m_CP10 = "115" Or m_CP10 = "118" Or m_CP10 = "301" _
      Or m_CP10 = "302" Or m_CP10 = "303" Or m_CP10 = "304" Or m_CP10 = "305" _
      Or m_CP10 = "306" Or m_CP10 = "307" Or m_CP10 = "803" Then
      Me.txt1(0) = "為「" & Trim(lblCaseName.Caption) & "」" & GetNationName(m_PA09, 0) & lblCP10 & "專利案提出申請。"
   Else
      Me.txt1(0).Text = "「" & Trim(lblCaseName.Caption) & "」" & GetNationName(m_PA09, 0) & Trim(lblPA08.Caption) & "專利之" & lblCP10
   End If
End Sub

Private Sub ClearData()
   'm_AttEEP02 = Empty
   txtEEP02 = Empty
   txtEEP03 = Empty
   txtEEP03_2 = Empty
   CboEEP04.Clear
   CboEEP05.Clear
   txtEEP10 = Empty
   txtEEP10_2 = Empty
   txtEEP08 = Empty
   lstAtt(0).Clear
   Me.cmdOpenAtt(0).Enabled = False
   Me.cmdSelect(0).Enabled = False
   Me.cmdSaveAtt(0).Enabled = False
   Me.cmdAddAtt(0).Visible = False
   Me.cmdRemAtt(0).Visible = False
   Me.cmdPrintAtt(0).Enabled = False
   Me.cmdSelAllPrt.Enabled = False
   Me.cmdBack.Visible = False
   Me.Frame2.Visible = False
   Me.cmdSend.Visible = False: bolhaveEfile = False
   Me.cmdFlow(0).Visible = False 'Add By Sindy 2023/11/9 程序送判
   Me.cmdFlow(1).Visible = False 'Add By Sindy 2023/12/6 聯絡
End Sub

'退件
Private Sub cmdBack_Click()
Dim strUpdTime As String
Dim intMaxEEP02 As Integer ', strEEP05 As String
Dim strEEP08 As String
Dim rsA As New ADODB.Recordset
   
   'Add by Sindy 2021/12/24 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   'Add By Sindy 2025/10/21
   If Frame1Big.Visible = True Then
      Call cmdClose_Click
      Exit Sub
   End If
   '2025/10/21 END
   
   If ChkRevStatus("1") = False Then Exit Sub 'Add By Sindy 2014/6/4
   'Add By Sindy 2015/9/11
   If m_EPMan = "" Then
      MsgBox "無承辦人不可退件！"
      Exit Sub
   End If
   'Modify By Sindy 2015/10/22 Mark
'   If Left(Trim(m_EPMan), 5) = strUserNum Then
'      MsgBox "退件對象不可為自己！"
'      Exit Sub
'   End If
   '2015/9/11 END
   
   'Modify By Sindy 2013/10/3 玲玲說一定要輸入理由
'   If MsgBox("是否確定要退件？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'      Exit Sub
'   Else
   strEEP08 = InputBox("是否確定要退件？請輸入退件原因：" & vbCrLf & "（註：一定要輸入原因）" & _
                  IIf(bolhaveEfile = True, vbCrLf & "注意：「卷宗區」及「原始檔區」若有電子檔需留存請先下載，因做了「退件」後會一併刪除該區電子檔！", ""))
   If Trim(strEEP08) = "" Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   strUpdTime = Right("000000" & ServerTime, 6)
   
   '取得最大序號
   intMaxEEP02 = 0
   strSql = "select eep02 From empelectronprocess where eep01='" & m_EEP01 & "' order by eep02 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      If RsTemp.RecordCount > 0 Then
         intMaxEEP02 = RsTemp.Fields(0)
      End If
   End If
   'Add By Sindy 2020/10/30
   If m_EEP15 <> "" Then
      strEEP08 = strEEP08 & "(" & m_EEP15 & "一併退件)"
   End If
   '2020/10/30 END
'   strEEP05 = Trim(Left(CboEEP05.Text, 5))
'   If strEEP05 = "" Then strEEP05 = strUserNum
   'Modify By Sindy 2013/10/24 EEP09改為不需回覆(取消Y)
   'Modify By Sindy 2017/6/15 + eep10=m_CP14_2
   'Modify By Sindy 2020/12/11 + eep15=m_EEP15
   'Modify By Sindy 2025/1/15 + ,eep12,eep16
   strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05" & _
            ",eep06,eep07,eep08,eep09,eep10,eep15,eep12,eep16) values(" & _
            CNULL(m_EEP01) & "," & (intMaxEEP02 + 1) & ",'" & strUserNum & "'," & _
            CNULL(EMP_退件) & "," & CNULL(Left(Trim(m_EPMan), 5)) & "," & strSrvDate(1) & "," & _
            strUpdTime & "," & CNULL(ChgSQL(strEEP08)) & ",null," & CNULL(m_CP14_2) & _
            ",'" & m_EEP15 & "','" & m_EEP12 & "','" & m_EEP16 & "')"
   cnnConnection.Execute strSql

   '將附件檔一併移到退件流程中
   strSql = "update empelectronfile" & _
            " set eef02=" & (intMaxEEP02 + 1) & _
            " where eef01='" & m_EEP01 & "' and eef02=" & CInt(m_AttEEP02)
   cnnConnection.Execute strSql
   
   'Modify By Sindy 2020/9/29 + EMP_多案承辦單
   PUB_DelFtpFile2 m_EEP01, " and eef02=" & CInt(m_AttEEP02) & _
                            " and (instr(upper(eef03),upper('" & EMP_承辦單 & "'))>0 or instr(upper(eef03),upper('" & EMP_多案承辦單 & "'))>0) ", "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
   '附件區是否已產生承辦單,若是,先刪除承辦單及卷宗區和原始檔區
   'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
   'Modify By Sindy 2020/9/29 + EMP_多案承辦單
   strSql = "delete from empelectronfile" & _
            " where eef01='" & m_EEP01 & "' and eef02=" & CInt(m_AttEEP02) & " and (instr(upper(eef03),upper('" & EMP_承辦單 & "'))>0 or instr(upper(eef03),upper('" & EMP_多案承辦單 & "'))>0)"
   cnnConnection.Execute strSql
   'Add By Sindy 2014/7/22 有歸檔過才需要清檔案
   If bolhaveEfile = True Then
   '2014/7/22 END
      PUB_DelFtpFile2 m_EEP01, " and cpf11='S' ", "CASEPAPERFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
      '刪除原始檔區
      'Modify By Sindy 2014/11/24
      'strSql = "delete from casepaperfile where cpf01='" & m_EEP01 & "'"
      'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
      strSql = "delete from casepaperfile where cpf01='" & m_EEP01 & "' and cpf11='S'"
      '2014/11/24 END
      cnnConnection.Execute strSql
      
      '刪除卷宗區
      strSql = "select cpp01,cpp10 from casepaperpdf where cpp01='" & m_EEP01 & "' and cpp12='S'"
      If rsA.State = adStateOpen Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
'         If rsA.Fields("cpp10") = "Y" Then
'            '已合併時,必須刪除整份卷宗
'            strSql = "delete from casepaperpdf where cpp01='000000000' and upper(cpp02)='" & UCase(m_CP01 & m_CP02 & m_CP03 & m_CP04 & ".pdf") & "'"
'            cnnConnection.Execute strSql
'            '並且要將此本所案號的全部附件已合併欄位值改成X,晚上批次作業必須再全部合併一次
'            strSql = "update casepaperpdf set cpp10='X' where cpp01 in(" & _
'                     "select cp09 from caseprogress where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "')" & _
'                     " and cpp10='Y'"
'            cnnConnection.Execute strSql
'         End If
         
         PUB_DelFtpFile2 m_EEP01, " and cpp12='S'" 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
         'Modify By Sindy 2014/11/24
         'strSql = "delete from casepaperpdf where cpp01='" & m_EEP01 & "'"
         'Memo by Morgan 2015/4/28 刪除條件要和刪除FTP檔的同步
         strSql = "delete from casepaperpdf where cpp01='" & m_EEP01 & "' and cpp12='S'"
         '2014/11/24 END
         cnnConnection.Execute strSql
      End If
      rsA.Close
      
      'Add By Sindy 2024/6/25
      '檢查是否為多案歷程
      strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp10" & _
                  " From caseprogress" & _
                  " where cp163='" & m_EEP01 & "' and cp163<>cp09"
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         rsA.MoveFirst
         Do While Not rsA.EOF
            '刪除原始檔區
            PUB_DelFtpFile2 rsA.Fields("cp09"), " and cpf11='S' ", "CASEPAPERFILE"
            strSql = "delete from casepaperfile where cpf01='" & rsA.Fields("cp09") & "' and cpf11='S'"
            cnnConnection.Execute strSql
            
            '刪除卷宗區
            PUB_DelFtpFile2 rsA.Fields("cp09"), " and cpp12='S'"
            strSql = "delete from casepaperpdf where cpp01='" & rsA.Fields("cp09") & "' and cpp12='S'"
            cnnConnection.Execute strSql
            
            rsA.MoveNext
         Loop
      End If
      rsA.Close
      '2024/6/25 END
   End If
   
'   'Add By Sindy 2020/9/30 若退件是多案總收文號,也要清除進度檔的多案總收文號
'   If m_EEP15 <> "" Then
'      strSql = "update caseprogress set" & _
'               " cp163=null" & _
'               " where cp09 in('" & Replace(m_EEP15, ",", "','") & "')"
'      cnnConnection.Execute strSql, intI
'   End If
'   '2020/9/30 END
   
   cnnConnection.CommitTrans
   
   'Modify By Sindy 2023/12/15 杜燕文協理請作,主旨和內文加申請國家
   strSubject = Replace(lblCaseNo, "-0-00", "") & "(" & GetPrjNation(lblCaseNo) & ")(核會流程)-->退件"
   strContent = "當月目次：" & m_EP01 & vbCrLf & _
                "本所案號：" & lblCaseNo & vbCrLf & _
                "案件名稱：" & lblCaseName & vbCrLf & _
                "申請國家：" & GetPrjNation(lblCaseNo) & vbCrLf & _
                "案件性質：" & lblCP10 & vbCrLf & _
                "流程狀態：退件" & vbCrLf & _
                "原　　因：" & strEEP08
   '發給工程師
   'Modify By Sindy 2022/8/5 + m_CP14_2
   PUB_SendMail strUserNum, Left(Trim(m_EPMan), 5), m_EEP01, strSubject, strContent, , , , , , m_CP14_2
   
   Screen.MousePointer = vbDefault
   cmdExit_Click
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 退件失敗！" & vbCrLf & Err.Description
End Sub

'Add By Sindy 2025/10/15
Private Sub CmdCalendar_Click()
   'Add By Sindy 2025/10/21
   If Frame1Big.Visible = True Then
      Call cmdClose_Click
      Exit Sub
   End If
   '2025/10/21 END
   
   If CheckUse("frm060209", strExec) = True Then
      'Added by Lydia 2025/09/10 傳入本所案號
      If PUB_CheckFormExist("frm060209") Then
         MsgBox "請先關閉〔行事曆提醒通知〕！", vbCritical + vbOKOnly
         Exit Sub
      End If
      Call frm060209.SetParent(Me, Replace(lblCaseNo, "-", ""))
      'end 2025/09/10
      frm060209.Show
   End If
End Sub

'結束
Public Sub cmdExit_Click()
   m_PrevForm.Hide
   If UCase(m_PrevForm.Name) = UCase("frm090202_4") Or _
      UCase(m_PrevForm.Name) = UCase("frm090202_7") Then
      m_PrevForm.QueryData
   End If
   m_PrevForm.Show
   Unload Me
End Sub

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Function DeleteFile(strFileName As String, intEEP02 As Integer) As Boolean
Dim stReName As String
   
On Error GoTo ErrHand
   
   DeleteFile = True
   Screen.MousePointer = vbHourglass
   
   'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
   PUB_DelFtpFile2 m_EEP01, " and eef02=" & intEEP02 & " and upper(eef03)='" & UCase(strFileName) & "'", "EMPELECTRONFILE"
   'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
   strSql = "delete from EmpElectronFile where eef01='" & m_EEP01 & "' and eef02=" & intEEP02 & " and upper(eef03)='" & UCase(strFileName) & "'"
   cnnConnection.Execute strSql
   'Pub_SeekTbLog strSql
   Pub_SaveLog strUserNum, "刪除歷程附件：順序(" & intEEP02 & ")" & strFileName, m_CP01, m_CP02, m_CP03, m_CP04, m_EEP01
'   If UCase(m_PrevForm.Name) = UCase("frm090202_5") Then '承辦單打字登錄作業
'      strSql = "delete from casepaperpdf where cpp01='" & m_EEP01 & "' and upper(cpp02)='" & UCase(strFileName) & "'"
'      cnnConnection.Execute strSql
'      Pub_SeekTbLog strSql
      '更名
      Call PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, strFileName, stReName)
      'If InStr(UCase(stReName), ".PDF") > 0 Then
      If Right(Trim(UCase(stReName)), 4) = ".PDF" Then
         If DelAttFile_PDF(lblCaseNo.Caption, m_EEP01, stReName) = False Then GoTo ErrHand1
      Else
         If DelAttFile_File(lblCaseNo.Caption, m_EEP01, stReName) = False Then GoTo ErrHand1
      End If
'   End If
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   DeleteFile = False
   Screen.MousePointer = vbDefault
   MsgBox " 刪除檔案（" & strFileName & "）失敗！" & vbCrLf & Err.Description
   Exit Function
   
ErrHand1:
   DeleteFile = False
   Screen.MousePointer = vbDefault
End Function

'歸卷
Private Function InsertFileData(isFileNameNoSave As String, Index As Integer) As Boolean
   Dim stFileName As String, stReName As String, stFileName2 As String
   Dim strTableName As String
   Dim UpdModifyDate As Double, UpdModifyTime As Double
   Dim bolFileSave As Boolean 'Add By Sindy 2013/11/6
   'Add By Sindy 2023/6/21
   Dim strUpdCP01 As String
   Dim strUpdCP02 As String
   Dim strUpdCP03 As String
   Dim strUpdCP04 As String
   Dim strUpdCP10 As String
   Dim strUpdEEP01 As String
   Dim rsA As New ADODB.Recordset
   Dim strUpdRecv As String '記錄有新增到卷宗區的文號
   Dim strSaveCaseNo1 As String, strSaveCaseNo2 As String, strSaveCaseNo3 As String, strSaveCaseNo4 As String
   Dim arrID As Variant, intCnt As Integer
   '2023/6/21 END
   Dim bolSavePDFtoOrg As Boolean 'Add By Sindy 2025/10/28
   
On Error GoTo ErrHand
   
   InsertFileData = True
   For ii = 0 To lstAtt(Index).ListCount - 1
      stFileName = lstAtt(Index).List(ii)
      If InStrRev(stFileName, " (") > 0 Then
         stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
      End If
      bolFileSave = True
      bolSavePDFtoOrg = False 'Add By Sindy 2025/10/28
      'Modify By Sindy 2013/11/6
      '檢查是否有要踢除的檔案
      If InStr(UCase(isFileNameNoSave), UCase("DWG")) > 0 Then
         'Modify By Sindy 2014/3/11 +dwg.7z
         If Right(UCase(stFileName), 4) = UCase(".DWG") Or _
            Right(UCase(stFileName), 7) = UCase("DWG.ZIP") Or _
            Right(UCase(stFileName), 6) = UCase("DWG.7Z") Or _
            Right(UCase(stFileName), 7) = UCase("DWG.PDF") Then
            bolFileSave = False
         End If
      End If
      'Modify By Sindy 2024/8/29 附件區剔除.MSG檔
      If InStr(UCase(isFileNameNoSave), UCase("MSG")) > 0 Then
         If Right(UCase(stFileName), 4) = UCase(".MSG") Then
            bolFileSave = False
            If m_ProState <> "FCP" And m_ProState <> "FCT" And m_ProState <> "CFT" Then
               PUB_SendMail strUserNum, "97038", "", _
                         "待送件附件區增加剔除.MSG檔 [觀察其他單位會放入MSG檔嗎?]", _
                         lblCP09.Caption & "(" & lblCP10.Caption & ")：" & stFileName, , , , , , , , , , True, False, , , False, , , False
            End If
         End If
      '2024/8/29 END
      End If
      
      'Modify By Sindy 2018/7/24
      stFileName2 = Right(stFileName, Len(stFileName) - InStrRev(stFileName, ".") + 1) '副檔名
      '排除 承辦單.menu
      If UCase(stFileName2) = UCase(".menu") Then
         bolFileSave = False
      'Add By Sindy 2025/10/21 外專"附件區"PDF不歸,但有例外 .FIG.PDF、.RES.PDF、.SEP.PDF
      ElseIf m_ProState = "FCP" And Index = 0 Then
         If UCase(stFileName2) = UCase(".PDF") Then
            If Right(UCase(stFileName), 8) = UCase(".FIG.PDF") _
               Or Right(UCase(stFileName), 8) = UCase(".RES.PDF") _
               Or Right(UCase(stFileName), 8) = UCase(".SEP.PDF") Then
               bolSavePDFtoOrg = True '要存原始檔區
            Else
               bolFileSave = False '不可存檔
            End If
         ElseIf UCase(stFileName2) = UCase(".MSG") Then
            bolFileSave = False '不可存檔
         End If
      '2025/10/21 END
      End If
      '2018/7/24 END
      
      'If InStr(UCase(stFileName), UCase(isFileNameNoSave)) = 0 Then
      If bolFileSave = True Then
      '2013/11/6 END
         'Add By Sindy 2023/6/21 開放多案歷程可以歸卷到其他案號
         strUpdCP01 = m_CP01
         strUpdCP02 = m_CP02
         strUpdCP03 = m_CP03
         strUpdCP04 = m_CP04
         strUpdCP10 = m_CP10
         strUpdEEP01 = m_EEP01
         If Left(stFileName, Len(m_CP01)) = m_CP01 And InStr(stFileName, Val(m_CP02)) = 0 And InStr(stFileName, m_CP02) = 0 Then
            '檢查是否為多案歷程
            strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp10" & _
                        " From caseprogress" & _
                        " where cp163='" & m_EEP01 & "' and cp163<>cp09"
            rsA.CursorLocation = adUseClient
            rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               rsA.MoveFirst
               Do While Not rsA.EOF
                  '案號形式
                  strSaveCaseNo1 = Trim(rsA.Fields("cp01")) & CStr(Val(rsA.Fields("cp02"))) & IIf(rsA.Fields("cp03") <> "0" Or rsA.Fields("cp04") <> "00", "-" & rsA.Fields("cp03"), "") & IIf(rsA.Fields("cp04") <> "00", "-" & rsA.Fields("cp04"), "")
                  strSaveCaseNo2 = Trim(rsA.Fields("cp01")) & "-" & CStr(Val(rsA.Fields("cp02"))) & IIf(rsA.Fields("cp03") <> "0" Or rsA.Fields("cp04") <> "00", "-" & rsA.Fields("cp03"), "") & IIf(rsA.Fields("cp04") <> "00", "-" & rsA.Fields("cp04"), "")
                  strSaveCaseNo3 = Trim(rsA.Fields("cp01")) & CStr(rsA.Fields("cp02")) & IIf(rsA.Fields("cp03") <> "0" Or rsA.Fields("cp04") <> "00", "-" & rsA.Fields("cp03"), "") & IIf(rsA.Fields("cp04") <> "00", "-" & rsA.Fields("cp04"), "")
                  strSaveCaseNo4 = Trim(rsA.Fields("cp01")) & "-" & CStr(rsA.Fields("cp02")) & IIf(rsA.Fields("cp03") <> "0" Or rsA.Fields("cp04") <> "00", "-" & rsA.Fields("cp03"), "") & IIf(rsA.Fields("cp04") <> "00", "-" & rsA.Fields("cp04"), "")
                  If InStr(UCase(stFileName), strSaveCaseNo1) > 0 Or _
                     InStr(UCase(stFileName), strSaveCaseNo2) > 0 Or _
                     InStr(UCase(stFileName), strSaveCaseNo3) > 0 Or _
                     InStr(UCase(stFileName), strSaveCaseNo4) > 0 Then
                     strUpdCP01 = rsA.Fields("cp01")
                     strUpdCP02 = rsA.Fields("cp02")
                     strUpdCP03 = rsA.Fields("cp03")
                     strUpdCP04 = rsA.Fields("cp04")
                     strUpdCP10 = rsA.Fields("cp10")
                     strUpdEEP01 = rsA.Fields("cp09")
                     strUpdRecv = strUpdRecv & "," & rsA.Fields("cp09")
                     Exit Do
                  End If
                  rsA.MoveNext
               Loop
            End If
            rsA.Close
         End If
         '2023/6/21 END
         stReName = ""
         stFileName2 = Right(stFileName, Len(stFileName) - InStrRev(stFileName, ".") + 1) '副檔名
         '更名
         Call PUB_GetEmpFlowReNameFile(strUpdCP01, strUpdCP02, strUpdCP03, strUpdCP04, strUpdCP10, stFileName, stReName)
         'Add By Sindy 2021/3/24 存卷資料歸卷宗區補.info.附檔名 CFP-031067(訴願)
         If Index = 1 Then
            strExc(0) = "select * from efilecaption where efc03='存卷資料' and instr(upper('" & stReName & "'),'.'||upper(efc02)||'.')>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            'Modify By Sindy 2023/12/7 外專不更名 +And bolFCPFlow = False
            'Modify By Sindy 2024/8/14 外商FC同外專 + And bolFCTFlow = False
            If intI = 0 And bolFCPFlow = False And bolFCTFlow = False Then
               stReName = Left(stReName, Len(stReName) - Len(stFileName2)) & ".INFO" & stFileName2
            End If
         End If
         '2021/3/24 END
         UpdModifyDate = Mid(lstAtt(Index).List(ii), InStr(lstAtt(Index).List(ii), "#") + 1, 8)
         UpdModifyTime = Mid(lstAtt(Index).List(ii), InStr(lstAtt(Index).List(ii), "#") + 9, 6)
         
         '歸卷宗區:
         'Modify By Sindy 2023/11/10 +.MSG 也歸卷宗區
         'Modify By Sindy 2025/10/28 bolSavePDFtoOrg = True:指定要存原始檔區的檔案
         '                           增加排除這類的檔案 + And bolSavePDFtoOrg = False
         If (UCase(stFileName2) = UCase(".PDF") Or UCase(stFileName2) = UCase(".MSG")) _
            And bolSavePDFtoOrg = False Then
            strTableName = "CasePaperPDF"
            If SaveAttFile_PDF(strUpdEEP01, m_AttachPath & "\" & stFileName, stReName, UpdModifyDate, UpdModifyTime, False, "S") = False Then
               GoTo ErrHand
               Exit Function
            End If
         '歸原始檔區:
         Else
            strTableName = "CasePaperFile"
            If SaveAttFile_Org(strUpdEEP01, m_AttachPath & "\" & stFileName, stReName, UpdModifyDate, UpdModifyTime, "S") = False Then
               GoTo ErrHand
               Exit Function
            End If
         End If
      End If
   Next ii
   
   Set rsA = Nothing
   Exit Function
   
ErrHand:
   Set rsA = Nothing
   InsertFileData = False
'   Screen.MousePointer = vbDefault
'   cnnConnection.RollbackTrans
   
   If Err.Number > 0 Then
      MsgBox " 新增檔案（" & stFileName & "）至" & strTableName & "失敗！" & vbCrLf & Err.Description
   End If
   
   'Add By Sindy 2019/2/11 凡有歸不成功者,已歸入的先全部刪除,待重歸
   If DelAttFile_PDF(lblCaseNo.Caption, m_EEP01, "", "S") = False Then Exit Function
   If DelAttFile_File(lblCaseNo.Caption, m_EEP01, "", "S") = False Then Exit Function
   '2019/2/11 END
   'Add By Sindy 2023/6/21
   If strUpdRecv <> "" Then
      strUpdRecv = Mid(strUpdRecv, 2)
      arrID = Split(strUpdRecv, ",")
      For intCnt = 0 To UBound(arrID)
         strExc(0) = "select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as caseno,cp09 from caseprogress where cp09='" & arrID(intCnt) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If DelAttFile_PDF(RsTemp.Fields("caseno"), RsTemp.Fields("cp09"), "", "S") = False Then Exit Function
            If DelAttFile_File(RsTemp.Fields("caseno"), RsTemp.Fields("cp09"), "", "S") = False Then Exit Function
         End If
      Next intCnt
   End If
   '2023/6/21 END
   
   If Index = 0 Then
      Call ReadAttachFile_other(m_EEP01)
   End If
End Function

'Add By Sindy 2013/10/8 檢查檔案是否開啟中
Private Function ChkInsFileOpening(Index As Integer) As Boolean
Dim stFileName As String
Dim fs, f
   
   ChkInsFileOpening = True
   For ii = 0 To lstAtt(Index).ListCount - 1
      stFileName = GetFileName(lstAtt(Index).List(ii))
'      'Modify By Sindy 2015/9/10
'      If InStr(UCase(stFileName), UCase(EMP_承辦單 & ".menu")) = 0 Then
'      '2015/9/10 END
         'Modify By Sindy 2013/11/7 Mark因重覆過濾檔名,會導致檔名中有(符號的檔名會被截掉
   '         If InStrRev(stFileName, " (") > 0 Then
   '            stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
   '         End If
      'Modify By Sindy 2018/7/10
      If Right(UCase(stFileName), 5) <> UCase(".menu") Then
      '2018/7/10 END
         '檔案是否正在使用中
         If PUB_ChkFileOpening(m_AttachPath & "\" & stFileName) = True Then
            MsgBox m_AttachPath & "\" & stFileName & vbCrLf & "檔案正在使用中，請關閉後，請重新執行〔產生承辦單及歸檔〕！", vbExclamation
            Me.cmdSend.Visible = False
            ChkInsFileOpening = False
            Exit Function
         End If
         
         Set fs = CreateObject("Scripting.FileSystemObject")
         Set f = fs.GetFile(m_AttachPath & "\" & stFileName)
         '檔案大小為 0 KB 有誤
         If f.Size = 0 Then
            MsgBox m_AttachPath & "\" & stFileName & vbCrLf & "檔案歸檔有誤，因檔案大小為 0 KB！請重新執行〔產生承辦單及歸檔〕！", vbExclamation
            Me.cmdSend.Visible = False
            ChkInsFileOpening = False
            Exit Function
         End If
      End If
   Next ii
End Function

'產生承辦單
Public Function PrintWorkSheet() As Boolean
Dim i As Long
Dim PrinterIndex As Integer
Dim stFileName As String
Dim stFileTime As String 'Add By Sindy 2018/5/10
Dim fs, f
Dim rsA As New ADODB.Recordset
   
   '先刪除卷宗區及原始檔區
'   strSql = "delete from CasePaperPDF where cpp01='" & m_EEP01 & "'"
'   cnnConnection.Execute strSql
'   strSql = "delete from CasePaperFile where cpf01='" & m_EEP01 & "'"
'   cnnConnection.Execute strSql
   'Add By Sindy 2014/7/22 未有承辦單時，不需要先刪卷宗區資料再新增
   If bolhaveEfile = True Then
   '2014/7/22 END
      If DelAttFile_PDF(lblCaseNo.Caption, m_EEP01, "", "S") = False Then Exit Function
      If DelAttFile_File(lblCaseNo.Caption, m_EEP01, "", "S") = False Then Exit Function
      'Add By Sindy 2024/6/25 多案歷程附件也要一併刪掉,後面重歸才不會重覆
      If cp(163) <> "" Then
         strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp10" & _
                     " From caseprogress" & _
                     " where cp163='" & m_EEP01 & "' and cp163<>cp09"
         rsA.CursorLocation = adUseClient
         rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do While Not rsA.EOF
               '刪除原始檔區
               PUB_DelFtpFile2 rsA.Fields("cp09"), " and cpf11='S' ", "CASEPAPERFILE"
               strSql = "delete from casepaperfile where cpf01='" & rsA.Fields("cp09") & "' and cpf11='S'"
               cnnConnection.Execute strSql
               
               '刪除卷宗區
               PUB_DelFtpFile2 rsA.Fields("cp09"), " and cpp12='S'"
               strSql = "delete from casepaperpdf where cpp01='" & rsA.Fields("cp09") & "' and cpp12='S'"
               cnnConnection.Execute strSql
               
               rsA.MoveNext
            Loop
         End If
         rsA.Close
      End If
      '2024/6/25 END
   End If
   
   'Add By Sindy 2025/7/29
   strSqlwhere = " and eef02=" & CInt(m_AttEEP02) & " and (instr(upper(eef03),upper('" & EMP_承辦單 & "'))>0 or instr(upper(eef03),upper('" & EMP_多案承辦單 & "'))>0)"
   strExc(0) = "select * from EmpElectronFile where eef01='" & m_EEP01 & "'" & strSqlwhere
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stFileName = RsTemp.Fields("eef03")
   '2025/7/29 END
      'Add By Sindy 2025/1/21 因後面程式會重新產生新的承辦單,先刪舊的
      PUB_DelFtpFile2 m_EEP01, strSqlwhere, "EMPELECTRONFILE"
      strSql = "delete from empelectronfile where eef01='" & m_EEP01 & "'" & strSqlwhere
      cnnConnection.Execute strSql, intI
      Pub_SaveLog strUserNum, "刪除歷程附件：順序(" & m_AttEEP02 & ")" & stFileName, m_CP01, m_CP02, m_CP03, m_CP04, m_EEP01 'Add By Sindy 2025/7/29
      '2025/1/21 END
   End If
      
   PrintWorkSheet = False
   '承辦單檔案名稱
   Call PUB_ChkEmpFlowFNMRule(lblCaseNo, "", "Y", m_CP10, stFileName, , False)
   'Add By Sindy 2020/9/29
   If cp(163) <> "" Then
      stFileName = stFileName & "." & m_CP10 & "." & EMP_多案承辦單
   Else
   '2020/9/29 END
      stFileName = stFileName & "." & m_CP10 & "." & EMP_承辦單
   End If
   
   If bolPAFlow = True Then
      '檢查是否有安裝PDFCreator
      PrinterIndex = -1
      For i = 0 To Printers.Count - 1
       If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
        PrinterIndex = i
        Exit For
       End If
      Next i
      If PrinterIndex < 0 Then
         MsgBox "請通知電腦中心安裝PDFCreator !!!"
         Exit Function
      End If
      
      '產生承辦單PDF
      Load frmPDF
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      Else
         KillAttach
      End If
      frmPDF.StartProcess m_AttachPath, stFileName
      'Modify By Sindy 2025/7/29 mark:上頭已刪除,此處重覆了
'      '檢查是否已有承辦單.PDF若有則先刪除重新產生
'      If DeleteFile(stFileName & ".pdf", CInt(m_AttEEP02)) = False Then Exit Function
      '2025/7/29 END
      Call PrintData '列印承辦單
      frmPDF.EndtProcess
      Unload frmPDF
   
   'Add By Sindy 2018/5/10
   Else
      stFileName = stFileName & ".menu"
      stFileTime = Right("000000" & ServerTime, 6)
      '以防重覆歸卷
      strSql = "delete from CasePaperPDF where cpp01='" & m_EEP01 & "' and cpp02='" & stFileName & "'"
      cnnConnection.Execute strSql, intI
      '新增一筆承辦單.menu至卷宗區
      strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10,cpp12)" & _
               " values('" & m_EEP01 & "'," & _
                       "'" & stFileName & "',0,'" & strUserNum & "'," & _
                       strSrvDate(1) & "," & stFileTime & "," & _
                       strSrvDate(1) & "," & stFileTime & ",'Y','S')"
      cnnConnection.Execute strSql, intI
      'Add By Sindy 2018/7/9
      'Modify By Sindy 2025/7/29 mark:上頭已刪除,此處重覆了
'      '檢查是否已有承辦單.PDF若有則先刪除重新產生
'      If DeleteFile(stFileName, CInt(m_AttEEP02)) = False Then Exit Function
      '2025/7/29 END
      '新增一筆承辦單.menu至歷程附件區
      strSql = "insert into empelectronfile(eef01,eef02,eef03,eef04,eef09,eef10)" & _
               " values('" & m_EEP01 & "'," & m_AttEEP02 & _
                       ",'" & stFileName & "',0," & strSrvDate(1) & "," & stFileTime & ")"
      cnnConnection.Execute strSql, intI
      '2018/7/9 END
   End If
   '2018/5/10 END
   
   'Modify By Sindy 2013/9/14
   '新增承辦單.PDF至lstAtt及儲存至資料庫
   lstAtt(0).Clear
   If bolPAFlow = True Then
      stFileName = m_AttachPath & "\" & stFileName & ".pdf"
      Set fs = CreateObject("Scripting.FileSystemObject")
      Set f = fs.GetFile(stFileName)
      '檔案大小為 0 KB 有誤
      If f.Size = 0 Then
         ShowMsg stFileName & vbCrLf & MsgText(9221)
         Exit Function
      End If
      AddListX lstAtt(0), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(0)
      'Add By Sindy 2013/10/8 檢查檔案是否開啟中
      If ChkInsFileOpening(0) = False Then
         Exit Function
      End If
      '2013/10/8 END
      '固定儲存List裡第一筆電子檔
      If SaveAttFile(CInt(m_AttEEP02), 0) = False Then
         Exit Function
      End If
      If UCase(m_PrevForm.Name) = UCase("frm090202_5") Then '承辦單打字登錄作業
         If SaveAttFile_PDF(m_EEP01, stFileName, GetFileName(stFileName), Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False, "S") = False Then
            Exit Function
         End If
      End If
      '2013/9/14 END
   'Add By Sindy 2018/5/10
   Else
      AddListX lstAtt(0), stFileName & " (0 KB)" & " #" & strSrvDate(1) & stFileTime & "#", lstAtt(0)
   '2018/5/10 END
   End If
   
   PrintWorkSheet = True
   Set rsA = Nothing
End Function

'Add By Sindy 2023/11/9
Private Sub cmdFlow_Click(Index As Integer)
   'Add By Sindy 2025/10/21
   If Frame1Big.Visible = True Then
      Call cmdClose_Click
      Exit Sub
   End If
   '2025/10/21 END
   
   '檢查表單是否已開啟，若是，則關閉
   If PUB_ChkFormIsClose("frm090202_2") = False Then
      Me.Enabled = True
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   frm090202_2.Hide
   frm090202_2.m_EEP01 = m_EEP01 '總收文號
   frm090202_2.m_FlowUserNum = Me.m_FlowUserNum 'Add By Sindy 2025/1/15
   If Index = 0 Then
      frm090202_2.intReceiveKind = 5 '程序送判
   Else
      frm090202_2.intReceiveKind = 99 '聯絡
   End If
   frm090202_2.SetParent Me
   frm090202_2.cmdOK(0).Visible = False
   frm090202_2.cmdOK(1).Visible = False
   frm090202_2.Cmd1(0).Visible = False
   If frm090202_2.QueryData = True Then
      frm090202_2.ShowNextData = True
      frm090202_2.cmdAdd_Click
      frm090202_2.Show
      Me.Hide
      Screen.MousePointer = vbDefault
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2025/1/23 電腦中心取消歸卷
Private Sub cmdM51_Click()
Dim strUpdTime As String
Dim rsA As New ADODB.Recordset
   
   '檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   If ChkRevStatus("1") = False Then Exit Sub 'Add By Sindy 2014/6/4
   
   'Add By Sindy 2024/6/21
   If cmdFlow(0).Tag = "Y" And bolhaveEfile = True Then
      If MsgBox("確定要取消歸卷嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Sub
      End If
   End If
   '2024/6/21 END
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   strUpdTime = Right("000000" & ServerTime, 6)
   
   'Modify By Sindy 2020/9/29 + EMP_多案承辦單
   PUB_DelFtpFile2 m_EEP01, " and eef02=" & CInt(m_AttEEP02) & _
                            " and (instr(upper(eef03),upper('" & EMP_承辦單 & "'))>0 or instr(upper(eef03),upper('" & EMP_多案承辦單 & "'))>0) ", "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
   '附件區是否已產生承辦單,若是,先刪除承辦單及卷宗區和原始檔區
   'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
   'Modify By Sindy 2020/9/29 + EMP_多案承辦單
   strSql = "delete from empelectronfile" & _
            " where eef01='" & m_EEP01 & "' and eef02=" & CInt(m_AttEEP02) & " and (instr(upper(eef03),upper('" & EMP_承辦單 & "'))>0 or instr(upper(eef03),upper('" & EMP_多案承辦單 & "'))>0)"
   cnnConnection.Execute strSql
   'Add By Sindy 2014/7/22 有歸檔過才需要清檔案
   If bolhaveEfile = True Then
   '2014/7/22 END
      PUB_DelFtpFile2 m_EEP01, " and cpf11='S' ", "CASEPAPERFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
      '刪除原始檔區
      'Modify By Sindy 2014/11/24
      'strSql = "delete from casepaperfile where cpf01='" & m_EEP01 & "'"
      'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
      strSql = "delete from casepaperfile where cpf01='" & m_EEP01 & "' and cpf11='S'"
      '2014/11/24 END
      cnnConnection.Execute strSql
      
      '刪除卷宗區
      strSql = "select cpp01,cpp10 from casepaperpdf where cpp01='" & m_EEP01 & "' and cpp12='S'"
      If rsA.State = adStateOpen Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
'         If rsA.Fields("cpp10") = "Y" Then
'            '已合併時,必須刪除整份卷宗
'            strSql = "delete from casepaperpdf where cpp01='000000000' and upper(cpp02)='" & UCase(m_CP01 & m_CP02 & m_CP03 & m_CP04 & ".pdf") & "'"
'            cnnConnection.Execute strSql
'            '並且要將此本所案號的全部附件已合併欄位值改成X,晚上批次作業必須再全部合併一次
'            strSql = "update casepaperpdf set cpp10='X' where cpp01 in(" & _
'                     "select cp09 from caseprogress where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "')" & _
'                     " and cpp10='Y'"
'            cnnConnection.Execute strSql
'         End If
         
         PUB_DelFtpFile2 m_EEP01, " and cpp12='S'" 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
         'Modify By Sindy 2014/11/24
         'strSql = "delete from casepaperpdf where cpp01='" & m_EEP01 & "'"
         'Memo by Morgan 2015/4/28 刪除條件要和刪除FTP檔的同步
         strSql = "delete from casepaperpdf where cpp01='" & m_EEP01 & "' and cpp12='S'"
         '2014/11/24 END
         cnnConnection.Execute strSql
      End If
      rsA.Close
      
      'Add By Sindy 2024/6/25
      '檢查是否為多案歷程
      strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp10" & _
                  " From caseprogress" & _
                  " where cp163='" & m_EEP01 & "' and cp163<>cp09"
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         rsA.MoveFirst
         Do While Not rsA.EOF
            '刪除原始檔區
            PUB_DelFtpFile2 rsA.Fields("cp09"), " and cpf11='S' ", "CASEPAPERFILE"
            strSql = "delete from casepaperfile where cpf01='" & rsA.Fields("cp09") & "' and cpf11='S'"
            cnnConnection.Execute strSql
            
            '刪除卷宗區
            PUB_DelFtpFile2 rsA.Fields("cp09"), " and cpp12='S'"
            strSql = "delete from casepaperpdf where cpp01='" & rsA.Fields("cp09") & "' and cpp12='S'"
            cnnConnection.Execute strSql
            
            rsA.MoveNext
         Loop
      End If
      rsA.Close
      '2024/6/25 END
   End If
      
   cnnConnection.CommitTrans
   
   Screen.MousePointer = vbDefault
   cmdExit_Click
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 退件失敗！" & vbCrLf & Err.Description
End Sub

'Add By Sindy 2015/6/24
Private Sub cmdok_Click(Index As Integer)
cmdState = Index
bolQuery = True 'Add By Sindy 2024/1/17
PubShowNextData
Exit Sub
End Sub

Public Sub PubShowNextData()
Dim rsA As New ADODB.Recordset
Dim stFileName As String
Dim hLocalFile As Long

Select Case cmdState
Case 0 '基本資料
   If bolQuery = True Then
      Me.Enabled = False
'      For i = 1 To GrdDataList.Rows - 1
'         GrdDataList.col = 0
'         GrdDataList.row = i
'         If Trim(GrdDataList.Text) = "V" Then
           Dim Str01 As String
'           GrdDataList.col = 0
'           GrdDataList.Text = ""
'           For j = 0 To GrdDataList.Cols - 1
'               GrdDataList.col = j
'               GrdDataList.CellBackColor = QBColor(15)
'           Next j
'           GrdDataList.col = 1
           Str01 = SystemNumber(lblCaseNo, 1)
           If Mid(UCase(Str01), 1, 1) = "N" Then
               Str01 = Mid(Str01, 2, 3)
           End If
'           If Not IsNull(GrdDataList.Text) Then
               'Modified by Morgan 2016/3/24 排除母層是共同查詢
               If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
                  fnCloseAllFrm100 'Added by Morgan 2016/2/22
               End If
               'end 2016/3/24
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               bolQuery = False
               Select Case Pub_RplStr(Str01)
                   Case "CFP", "FCP", "P"   '專利
                         Screen.MousePointer = vbHourglass
                         frm100101_3.Show
                         frm100101_3.Tag = Pub_RplStr(lblCaseNo)
                         frm100101_3.StrMenu
                         Screen.MousePointer = vbDefault
                   Case "CFT", "FCT", "T", "TF"   '商標
                         Screen.MousePointer = vbHourglass
                         frm100101_4.Show
                         frm100101_4.Tag = Pub_RplStr(lblCaseNo)
                         frm100101_4.StrMenu
                         Screen.MousePointer = vbDefault
                   'Modify By Sindy 2009/07/24 增加LIN系統類別
                   'modify by sonia 2019/7/29 +ACS系統類別
                   Case "CFL", "FCL", "L", "LIN", "ACS"   '法務
                         Screen.MousePointer = vbHourglass
                         frm100101_5.Show
                         frm100101_5.Tag = Pub_RplStr(lblCaseNo)
                         frm100101_5.StrMenu
                         Screen.MousePointer = vbDefault
                   Case "LA"            '顧問
                         Screen.MousePointer = vbHourglass
                         frm100101_6.Show
                         frm100101_6.Tag = Pub_RplStr(lblCaseNo)
                         frm100101_6.StrMenu
                         Screen.MousePointer = vbDefault
                   Case Else                  '服務
                        Select Case Pub_RplStr(Str01)
                            Case "TB"    '條碼
                               Screen.MousePointer = vbHourglass
                               frm100101_7.Show
                               frm100101_7.Tag = Pub_RplStr(lblCaseNo)
                               frm100101_7.StrMenu
                               Screen.MousePointer = vbDefault
                            Case "TM"
                               Screen.MousePointer = vbHourglass
                               frm100101_8.Show
                               frm100101_8.Tag = Pub_RplStr(lblCaseNo)
                               frm100101_8.StrMenu
                               Screen.MousePointer = vbDefault
                            Case "TD"
                               Screen.MousePointer = vbHourglass
                               frm100101_9.Show
                               frm100101_9.Tag = Pub_RplStr(lblCaseNo)
                               frm100101_9.StrMenu
                               Screen.MousePointer = vbDefault
                            Case "TC", "CFC"
                               Screen.MousePointer = vbHourglass
                               frm100101_A.Show
                               frm100101_A.Tag = Pub_RplStr(lblCaseNo)
                               frm100101_A.StrMenu
                               Screen.MousePointer = vbDefault
                            Case Else
                               Screen.MousePointer = vbHourglass
                               frm100101_B.Show
                               frm100101_B.Tag = Pub_RplStr(lblCaseNo)
                               frm100101_B.StrMenu
                               Screen.MousePointer = vbDefault
                         End Select
               End Select
'           End If
           Me.Enabled = True
           Exit Sub
'         End If
'      Next i
      Me.Enabled = True
   End If
Case 1 '進度
   If bolQuery = True Then
      Me.Enabled = False
'      StrTag = ""
'      For i = 1 To grdDataList.Rows - 1
'         grdDataList.col = 0
'         grdDataList.row = i
'         If Trim(grdDataList.Text) = "V" Then
'            grdDataList.col = 0
'            grdDataList.Text = ""
'            For j = 0 To grdDataList.Cols - 1
'                grdDataList.col = j
'                grdDataList.CellBackColor = QBColor(15)
'            Next j
'             grdDataList.col = 1
'             If Not IsNull(grdDataList.Text) Then
                'Modified by Morgan 2016/3/24 排除母層是共同查詢
                If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
                   fnCloseAllFrm100 'Added by Morgan 2016/2/22
                End If
                'end 2016/3/24
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                bolQuery = False
                frm100101_2.Show
                frm100101_2.Tag = Pub_RplStr(lblCaseNo)
                frm100101_2.bolEmpFlow = True 'Add By Sindy 2020/9/25
                'frm100101_2.cmdOK(6).Visible = False
                frm100101_2.StrMenu
                Screen.MousePointer = vbDefault
                Me.Enabled = True
                Exit Sub
'             End If
'         End If
'      Next i
      Me.Enabled = True
   End If
Case 2 '接洽單
   Screen.MousePointer = vbHourglass
   If m_CP140 <> "" Then
      '查詢接洽記錄單
      'Modify By Sindy 2022/12/23 改用共用函數
      Call PUB_Queryfrm090801(m_CP140, cp(5), Me)
'      'Modify By Sindy 2022/9/5
'      If DBDATE(cp(5)) >= 接洽單電子收文啟用日 Then
'         frm090801_Q.SetParent Me
'         frm090801_Q.m_blnCallPrint = True
'         frm090801_Q.Text5 = m_CP140
'         Call frm090801_New.cmdOK_Click(4)
'         frm090801_Q.Show vbModal
'      Else
'      '2022/9/5 END
'         frm090801.SetParent Me
'         frm090801.m_blnCallPrint = True 'Add By Sindy 2022/10/19
'         frm090801.Text5 = m_CP140
'         frm090801.m_blnCallPrint_CRL119 = True '是否列印特殊收據頁
'         Call frm090801.cmdOK_Click(4)
'         frm090801.cmdOK(2).Visible = False
'         frm090801.cmdOK(0).Visible = False
'         frm090801.txtPCnt.Visible = False
'         Me.Hide
'      End If
      '2022/12/23 END
      cmdState = 99 '結束
   Else
      '檢查是否有接洽單.pdf
      'Modify By Sindy 2023/11/17 +接洽單.msg
      'Modify By Sindy 2024/1/17 + order by cpp17 desc,cpp18 desc 增加接洽單電子檔進入系統的新增日期和時間做排序。（最近的優先抓）
      strExc(0) = "select *" & _
                  " From casepaperpdf" & _
                  " where cpp01='" & m_EEP01 & "'" & _
                  " and (instr(upper(cpp02),upper('" & EMP_接洽單 & ".pdf'))>0 or instr(upper(cpp02),upper('" & EMP_接洽單 & ".msg'))>0)" & _
                  " and cpp10<>'D'" & _
                  " order by cpp17 desc,cpp18 desc"
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         '讀取檔案名稱
         stFileName = rsA.Fields("cpp02")
'         If GetAttachFile_CPP(m_EEP01, stFileName, m_AttachPath & "\" & stFileName) = False Then
'            MsgBox "無法儲存欲開啟的檔案[ " & stFileName & " ]！"
'         End If
         If PUB_GetAttachFile_CPP(m_EEP01, stFileName, m_AttachPath) = True Then
            '開啟檔案
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Else
         MsgBox "無接洽單！"
      End If
      rsA.Close
      Set rsA = Nothing
   End If
   Screen.MousePointer = vbDefault
   
Case 4 '卷宗區
   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2024/2/1
   If cmdOK(4).Caption = "完整卷宗" Then
      frm100101_L.m_strKey = lblCaseNo.Caption
   Else
   '2024/2/1 END
      frm100101_L.m_strKey = m_EEP01
   End If
   'frm100101_L.Hide
   frm100101_L.SetParent Me
   If frm100101_L.QueryData = True Then
      frm100101_L.Show
      Me.Hide
   Else
      Unload frm100101_L
   End If
   Screen.MousePointer = vbDefault
Case Else
End Select
End Sub

'Add By Sindy 2015/9/9
'E-Mail及歸檔
Private Sub EMailKeepFile()
Dim bolHadFile As Boolean
Dim pbolDone As Boolean
Dim pFiles As String
Dim bolSelFile As Boolean, stFileName As String
Dim intStar As Integer
Dim intEnd As Integer
   
   '沒點選附件,就預設是全部附件
   bolHadFile = False
   For ii = 0 To lstAtt(0).ListCount - 1
      If lstAtt(0).Selected(ii) Then
         bolHadFile = True
         Exit For
      End If
   Next ii
   If bolHadFile = False Then
      For ii = 0 To lstAtt(0).ListCount - 1
         lstAtt(0).Selected(ii) = True
      Next ii
   End If
   
   '檢查檔案:
   '檢查一定要有.Data.PDF 或 .Data.DOC (指示信)
   'Modify By Sindy 2015/10/14 檢查若有 Ltr.PDF 或 Ltr.DOC 亦可放行
   bolHadFile = False
   For ii = 0 To lstAtt(0).ListCount - 1
      If lstAtt(0).Selected(ii) Then
         stFileName = Left(lstAtt(0).List(ii), InStrRev(lstAtt(0).List(ii), " (") - 1)
         'Modify By Sindy 2015/11/13 調整控管方式
'         If InStr(UCase(stFileName), UCase(".Data.PDF")) > 0 Or _
'            InStr(UCase(stFileName), UCase(".Data.DOC")) > 0 Or _
'            InStr(UCase(stFileName), UCase(".Ltr.PDF")) > 0 Or _
'            InStr(UCase(stFileName), UCase(".Ltr.DOC")) > 0
         If (InStr(UCase(stFileName), UCase("Data.")) > 0 And Right(UCase(stFileName), 4) = UCase(".PDF")) Or _
            (InStr(UCase(stFileName), UCase("Data.")) > 0 And Right(UCase(stFileName), 4) = UCase(".DOC")) Or _
            (InStr(UCase(stFileName), UCase("Data.")) > 0 And Right(UCase(stFileName), 5) = UCase(".DOCX")) Or _
            (InStr(UCase(stFileName), UCase("Ltr.")) > 0 And Right(UCase(stFileName), 4) = UCase(".PDF")) Or _
            (InStr(UCase(stFileName), UCase("Ltr.")) > 0 And Right(UCase(stFileName), 4) = UCase(".DOC")) Or _
            (InStr(UCase(stFileName), UCase("Ltr.")) > 0 And Right(UCase(stFileName), 5) = UCase(".DOCX")) Then
         '2015/11/13 END
            bolHadFile = True
            Exit For
'         'Add By Sindy 2015/11/12 案件性質副檔名和檔案副檔名中間若純數字則可以通過 ex:(.data1.PDF)
'         Else
'            intStar = 0: intEnd = 0
'            If Right(UCase(stFileName), 4) = UCase(".PDF") Or _
'               Right(UCase(stFileName), 4) = UCase(".DOC") Then
'               If InStr(UCase(stFileName), UCase(".Data")) > 0 Then
'                  intStar = InStr(UCase(stFileName), UCase(".Data")) + Len(".Data")
'                  intEnd = Len(UCase(stFileName)) - Len(".Data")
'               ElseIf InStr(UCase(stFileName), UCase(".Ltr")) > 0 Then
'                  intStar = InStr(UCase(stFileName), UCase(".Ltr")) + Len(".Ltr")
'                  intEnd = Len(UCase(stFileName)) - Len(".Ltr")
'               End If
'            End If
'            If intStar > 0 And intEnd > 0 Then
'               bolHadFile = True
'               For jj = intStar To intEnd
'                  If IsNumeric(Mid(UCase(stFileName), jj, 1)) = False Then
'                     bolHadFile = False
'                     Exit For
'                  End If
'               Next jj
'               If bolHadFile = True Then Exit For
'            End If
'         '2015/11/12 END
         End If
      End If
   Next ii
   If bolHadFile = False Then
      MsgBox "無指示信，不可執行！", vbExclamation
      Exit Sub
   End If
   '案件性質：202，一定要有 Poa.PDF 或 Assign.PDF
   'Modify By Sindy 2016/5/6 開放也可以是放 DWG.PDF (補圖)
   If m_CP10 = "202" Then
      bolHadFile = False
      For ii = 0 To lstAtt(0).ListCount - 1
         If lstAtt(0).Selected(ii) Then
            If InStr(UCase(lstAtt(0).List(ii)), UCase("Poa.PDF")) > 0 Or _
               InStr(UCase(lstAtt(0).List(ii)), UCase("Assign.PDF")) > 0 Or _
               InStr(UCase(lstAtt(0).List(ii)), UCase("DWG.PDF")) > 0 Then
               bolHadFile = True
               Exit For
            End If
         End If
      Next ii
      If bolHadFile = False Then
         MsgBox "無(Poa.PDF)或(Assign.PDF)或(DWG.PDF)檔案，不可執行！", vbExclamation
         Exit Sub
      End If
   End If
   '案件性質：232，一定要有 Pri.PDF
   If m_CP10 = "232" Then
      bolHadFile = False
      For ii = 0 To lstAtt(0).ListCount - 1
         If lstAtt(0).Selected(ii) Then
            If InStr(UCase(lstAtt(0).List(ii)), UCase("Pri.PDF")) > 0 Then
               bolHadFile = True
               Exit For
            End If
         End If
      Next ii
      If bolHadFile = False Then
         MsgBox "無(Pri.PDF)檔案，不可執行！", vbExclamation
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   
   '更新意見內容
   If txtEEP08.Locked = False Then
      strSql = "update EmpElectronProcess set eep08='" & txtEEP08 & "' where eep01='" & m_EEP01 & "' and eep02=" & m_AttEEP02
      cnnConnection.Execute strSql
   End If
   
   If lblSendMailDt.Visible = True Then
      If MsgBox("已寄過信件，是否需要重新寄信？" & _
                vbCrLf & vbCrLf & "（按「否」，直接處理重新歸檔）", vbExclamation + vbYesNo) = vbNo Then
         GoTo MoveToPaperPDF
      End If
   End If
   
   Screen.MousePointer = vbDefault
'*************************************************************************************************
   'E-Mail呼叫 frm880019:要將寄信的內容及寄信的成功時間儲存在資料庫中，便於事後查詢。
   frm880019.m_bolSaveMail = True
   frm880019.m_CP01 = m_CP01
   frm880019.m_CP02 = m_CP02
   frm880019.m_CP03 = m_CP03
   frm880019.m_CP04 = m_CP04
   frm880019.m_CP09 = m_EEP01
   frm880019.m_CP10 = m_CP10
   '主旨
   frm880019.txtSubject = "委託 " & m_CP01 & "-" & m_CP02 & IIf(m_CP03 & m_CP04 = "000", "", "-" & m_CP03) & IIf(m_CP04 = "00", "", "-" & m_CP04) & " 案" & lblCP10 & "作業電子檔"
   'Added by Morgan 2024/1/15 主旨加彼號--品薇
   If m_CP44 <> "" And cp(45) <> "" Then
      frm880019.txtSubject = frm880019.txtSubject & " Y/R:" & cp(45)
   End If
   'end 2024/1/15
   '本文
   'Modified by Morgan 2019/8/26 改呼叫函數(P案指示信本文統一)--品薇
   'frm880019.txtContent = vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                          "台一國際專利法律事務所  / " & PUB_GetST07(strUserNum) & vbCrLf & _
                          "電　話：(02)25061023" & vbCrLf & _
                          "傳　真：(02)25011666" & vbCrLf & _
                          "URL:https://www.taie.com.tw" & vbCrLf & _
                          "*************保密警語********************" & vbCrLf & _
                          "本信件僅授權於指定之收信人取閱之用，信件中可能含有機密性資訊。" & vbCrLf & _
                          "如果您並非被指定之收信人，任何未經授權而擅自使用此信件所含之機密資訊的行為是被嚴格禁止的。" & vbCrLf & _
                          "如果您在任何未經授權的情形之下收到本信件，煩請您立即告知原發信人並將此信件回傳至以上地址。" & vbCrLf & _
                          "謝謝您的合作。"
   frm880019.txtContent = PUB_GetOrderLetterContent("P")
   frm880019.m_bolPLetter = True 'Added by Morgan 2019/11/21
   'end 2019/8/26
   
   '附件
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum & "\otherFile" 'Add By Sindy 2017/1/6 以防止上面寄信時有些檔案會被咬住,後面刪檔會有權限問題
   KillAttach 'Add By Sindy 2017/3/10
   bolSelFile = False
   pFiles = ""
   For ii = 0 To lstAtt(0).ListCount - 1
      If lstAtt(0).Selected(ii) Then
         bolSelFile = True
         stFileName = lstAtt(0).List(ii)
         If InStrRev(stFileName, " (") > 0 Then
            stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
         End If
         If InStr(stFileName, "\") = 0 Then
            If GetAttachFile(stFileName, CInt(m_AttEEP02)) = False Then Exit Sub
         End If
         pFiles = pFiles & ";" & stFileName
      End If
   Next ii
   If bolSelFile = False Then
      Call DownloadAllAttachFile(CInt(m_AttEEP02), 0, pFiles)
   Else
      If pFiles <> "" Then pFiles = Mid(pFiles, 2)
   End If
   frm880019.SetAttach pFiles
   'frm880019.SetEmail m_PA26, m_PA149, m_PA75
   'Add By Sindy 2015/9/23
   If m_CP44 = "" Then
      '抓AB類收文號的代理人，預設最後發文日最大收文號的代理人...同發文作業預設的代理人(AddAgent)
      '2008/2/21 加聯絡人
      '2010/2/23 香港案要排除421
      strExc(0) = "SELECT  CP44,cp116,CP45,Max(nvl(CP27,0)||CP09) Srt FROM CASEPROGRESS" & _
                  " WHERE CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "'" & _
                  " AND CP03='" & m_CP03 & "' AND CP04='" & m_CP04 & "'" & _
                  " AND CP44 IS NOT NULL AND CP09<'C'" & _
                  IIf(m_PA09 = "013", " AND CP10<>'421'", "") & _
                  " Group By CP44,cp116,CP45 Order By Srt desc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_CP44 = "" & RsTemp.Fields("cp44")
         m_CP116 = "" & RsTemp.Fields("CP116")
         'Added by Morgan 2024/1/15 主旨加彼號--品薇
         If Not IsNull(RsTemp.Fields("cp44")) Then
            frm880019.txtSubject = frm880019.txtSubject & " Y/R:" & RsTemp.Fields("cp45")
         End If
         'end 2024/1/15
      End If
   End If
   '2015/9/23 END
   frm880019.SetEmail "", "", m_CP44, m_CP116 'Modify By Sindy 2015/9/22 改抓進度代理人
   frm880019.txtCopy = "" 'Add By Sindy 2015/10/6
   'If m_PA75 = "" Then
'   If m_CP44 = "" Then
'      frm880019.txtReceiver = ""
'   End If
   'Add By Sindy 2015/10/15 FMP時列印寄件備份歸檔存卷
   If Left(m_CP12, 1) = "F" Then
      frm880019.chkPrint.Visible = True
      'frm880019.chkPrint.Value = 1 'Removed by Morgan 2022/9/26 卷宗區已自動存寄件備份，無需再印紙本。--陳品薇
   End If
   '2015/10/15 END
   frm880019.cmdAttach.Visible = True 'False
   frm880019.SetParent Me
   frm880019.Show vbModal
   pbolDone = frm880019.m_bolDone
   Unload frm880019
'*************************************************************************************************
   Screen.MousePointer = vbHourglass
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum 'Add By Sindy 2017/1/6 以防止上面寄信時有些檔案會被咬住,後面刪檔會有權限問題
   If pbolDone = True Then  '寄信成功
      Call ReadSmailBackup '寄件日期
MoveToPaperPDF:
      Me.Enabled = False
      '產生承辦單
      If PrintWorkSheet = False Then
         Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
         Me.Enabled = True
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      If ToKeepOnFiles = True Then '歸檔
'         '產生本所案號+案件性質+WorkSheet.menu
'         strExc(0) = "select eef01,eef09,eef10 from EmpElectronFile" & _
'                     " where eef01='" & m_EEP01 & "' and eef02=" & m_AttEEP02 & " and instr(upper(eef03),upper('" & EMP_承辦單 & ".menu'))>0"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 0 Then
'            strSql = "insert into EmpElectronFile(eef01,eef02,eef03,eef04,eef09,eef10)" & _
'                     " values('" & m_EEP01 & "'," & m_AttEEP02 & _
'                             ",'" & m_CP01 & Val(m_CP02) & IIf(m_CP03 = "0" And m_CP04 = "00", "", "-" & m_CP03) & IIf(m_CP04 = "00", "", "-" & m_CP04) & "." & m_CP10 & "." & EMP_承辦單 & ".menu',0," & _
'                             strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ")"
'            cnnConnection.Execute strSql
'         End If
'         '卷宗區也要產生本所案號+案件性質+WorkSheet.menu
'         strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,cpp08,cpp09,cpp10,cpp12)" & _
'                  " values('" & m_EEP01 & "'," & _
'                          "'" & m_CP01 & Val(m_CP02) & IIf(m_CP03 = "0" And m_CP04 = "00", "", "-" & m_CP03) & IIf(m_CP04 = "00", "", "-" & m_CP04) & "." & m_CP10 & "." & EMP_承辦單 & ".menu',0," & _
'                          strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",'Y','S')"
'         cnnConnection.Execute strSql
      End If
      Me.Enabled = True
   End If
   
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Sub

'Added by Morgan 2020/3/11
'檢查申請人與多國(含國內案)是否不同
Private Function ApplyerCheck(pa01 As String, pa02 As String, pa03 As String, pa04 As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   Dim stCaseNo As String
   
   stSQL = "select b.pa01||'-'||b.pa02||decode(b.pa03||b.pa04,'000','','-'||b.pa03||'-'||b.pa04)||'('||na03||')' CNo,b.pa09" & _
      " from (select cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08 from caserelation where cr01='" & pa01 & "' and cr02='" & pa02 & "' and cr03='" & pa03 & "' and cr04='" & pa04 & "'" & _
      " union select cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08 from casemap where cm01='" & pa01 & "' and cm02='" & pa02 & "' and cm03='" & pa03 & "' and cm04='" & pa04 & "' and cm10='0'" & _
      ") x, patent a, patent b,nation where a.pa01(+)=cr01 and a.pa02(+)=cr02 and a.pa03(+)=cr03 and a.pa04(+)=cr04" & _
      " and b.pa01(+)=cr05 and b.pa02(+)=cr06 and b.pa03(+)=cr07 and b.pa04(+)=cr08" & _
      " and b.pa26||b.pa27||b.pa28||b.pa29||b.pa30<>a.pa26||a.pa27||a.pa28||a.pa29||a.pa30" & _
      " and na01(+)=b.pa09 order by b.pa09"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With RsQ
      Do While Not .EOF
         stCaseNo = stCaseNo & .Fields(0) & vbCrLf
         .MoveNext
      Loop
      End With
      If MsgBox("本案申請人與下列多國案(含國內案)申請人不同，" & vbCrLf & "請確認申請人資料是否正確？" & vbCrLf & vbCrLf & stCaseNo, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
         ApplyerCheck = True
      End If
   Else
      ApplyerCheck = True
   End If
   Set RsQ = Nothing
End Function

'Add By Sindy 2025/10/20 變大的附件區
Private Sub cmdOpen_Click()
   TextFCPNote(1).Visible = TextFCPNote(0).Visible
   If TextFCPNote(1).Visible = True Then lstAtt(2).Height = lstAtt(2).Height + TextFCPNote(1).Height
   Call SetFrame1Big(False, 2, 0)
End Sub
Private Sub CmdOpen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.CmdOpen.ToolTipText = "可放大附件區"
End Sub
Private Sub cmdClose_Click()
   If Frame1Big.Visible = False Then Exit Sub
   lstAtt(2).Height = lstAtt(2).Tag
   Call SetFrame1Big(True, 0, 2)
End Sub
Private Sub SetFrame1Big(bolVal As Boolean, lstAttNew As Integer, lstAttOld As Integer)
   Text2.Visible = bolVal: SSTab1.Visible = bolVal
   cmdOpenAtt(2).Enabled = cmdOpenAtt(0).Enabled
   cmdSelect(2).Enabled = cmdSelect(0).Enabled
   cmdSaveAtt(2).Enabled = cmdSaveAtt(0).Enabled
   cmdPrintAtt(2).Enabled = cmdPrintAtt(0).Enabled
   cmdAddAtt(2).Enabled = cmdAddAtt(0).Enabled
   cmdRemAtt(2).Enabled = cmdRemAtt(0).Enabled
   lstAtt(lstAttNew).Clear
   For ii = 0 To lstAtt(lstAttOld).ListCount - 1
      lstAtt(lstAttNew).AddItem lstAtt(lstAttOld).List(ii)
      lstAtt(lstAttNew).Selected(ii) = lstAtt(lstAttOld).Selected(ii)
   Next ii
   If lstAtt(lstAttNew).ListCount > 0 Then SetListScroll lstAtt(lstAttNew)
   Frame1Big.Visible = Not bolVal
End Sub
'2025/10/20 END

'Add By Sindy 2024/4/17 匯出Outlook
Private Sub cmdOutlook_Click()
   'Add By Sindy 2025/10/21
   If Frame1Big.Visible = True Then
      Call cmdClose_Click
      Exit Sub
   End If
   '2025/10/21 END
   
   Call EmpFlowFCPOutlook(m_EEP01, m_CP01, m_CP02, m_CP03, m_CP04, lblCaseName, lblCP10)
End Sub

'產生承辦單及歸檔
Private Sub cmdPrintAllPDF_Click()
   
   'Add by Sindy 2021/12/24 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   'Added by Morgan 2025/3/14
   '台灣案自請撤回發文檢查
   If m_CP01 = "P" And m_PA09 = "000" And m_CP10 = "413" Then
      If PUB_ChkTW413(cp(43)) = False Then
         Exit Sub
      End If
   End If
   'end 2025/3/14

   'Added by Morgan 2020/3/11
   'CFP新案發文檢查申請人與多國案(含國內案)是否不同
   If m_CP01 = "CFP" And InStr(NewCasePtyList, m_CP10) > 0 Then
      If ApplyerCheck(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
         Exit Sub
      End If
   End If
   'end 2020/3/11
   
   'Modify By Sindy 2015/9/9
   If cmdPrintAllPDF.Caption = "E-Mail及歸檔" Then
      'Added by Morgan 2024/2/22 +檢查送件方式檢查
      If PUB_ChkCP141IsSend(m_EEP01, , "EMail指示信") = False Then Exit Sub
      'end 2024/2/22
      Call EMailKeepFile
      Exit Sub
   'Modify By Sindy 2018/11/19
   ElseIf cmdPrintAllPDF.Caption = "產生送件資料夾" Then
      'Added by Morgan 2023/12/28
      'Modified by Morgan 2024/2/5 改用函數檢查(含收款後送件檢查)
      'If cp(141) = "3" And (cp(164) = "1" Or cp(164) = "3") And strSrvDate(1) < cp(142) Then
      '   '「本案需於指定日方可發文」或「本案需於指定日之後方可發文」
      '   strExc(0) = ChangeWStringToTDateString(cp(142))
      '   If cp(164) = "1" Then
      '      MsgBox "本案需於指定日期(" & strExc(0) & ")方可送件！", vbExclamation, "指定日期檢查"
      '   ElseIf cp(164) = "3" Then
      '      MsgBox "本案需於指定日期(" & strExc(0) & ")之後方可送件！", vbExclamation, "指定日期檢查"
      '   End If
      '   Exit Sub
      'End If
      If PUB_ChkCP141IsSend(m_EEP01, , "送件") = False Then Exit Sub
      'end 2024/2/5
      'end 2023/12/28
      Call BegetAppFUpload
      Exit Sub
   End If
   '2015/9/9 END
   
   '產生承辦單
   Me.Enabled = False
   If PrintWorkSheet = False Then
      Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
      Me.Enabled = True
      Exit Sub
   End If
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   
   'Add By Sindy 2023/11/22 外專附件一律不歸卷，維持Backup回存信件機制
   '　　                    僅歸（存卷資料）區
   'If bolFCPFlow = True Then
   If TextFCPNote(0).Visible = True Then '*** 重要 ***
      'Add By Sindy 2025/10/20 增加可歸原始檔區的功能
      '適用案件性質(排除各類報告客戶的案件): 新申請案(101,102,103,104,105,109,110,112,113,114,115,118,120,122,125,307)
      '                                     、203主動修正、204修正、205申復、107再審（紀錄是否有修正）
      '                                     、433誤譯訂正、210製作中說、307分割、242製作外文提申本。
      'Modify By Sindy 2025/11/14 自動歸檔適用類型: 增加416.實體審查 且 承辦人為工程師
      If m_ProState = "FCP" And _
         (InStr(NewCasePtyList & ",203,204,205,107,433,210,242", cp(10)) > 0 _
            Or (cp(10) = "416" And PUB_GetST03(cp(14)) = "F21")) _
         And cp(118) <> "" _
         And (cp(1) = "FCP" Or cp(1) = "FG") Then
         '將桌面上電子檔匯入附件區
         If FCPSaveFileToListBox(0) = False Then GoTo RunExit
                  
         Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), False)
         Call DownloadAllAttachFile(CInt(m_AttEEP02), 0)
         '檢查檔案是否開啟中
         If ChkInsFileOpening(0) = False Then
            GoTo RunExit
         End If
         '將檔案存至卷宗區或原始檔區
         If InsertFileData("無須踢除的檔案", 0) = False Then
            GoTo RunExit
         End If
      End If
      '2025/10/20 END
      
      '下載存卷附件
      Call ReadAttachFile_other(m_EEP01)
      Call DownloadAllAttachFile(0, 1)
      'Add By Sindy 2013/10/8 檢查檔案是否開啟中
      If ChkInsFileOpening(1) = False Then
         GoTo RunExit
      End If
      '2013/10/8 END
      '將"存卷區"電子檔分別存至卷宗區及原始檔區
      If InsertFileData("無須踢除的檔案", 1) = True Then
         Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
      End If
      GoTo RunExit '結束 '*****
   End If
   
   '其他系統的歸卷處理
   If ToKeepOnFiles = True Then '歸檔
      'Modify By Sindy 2014/5/21 Mark 電子檔匯入時,會檢查新案是否已歸足,若未歸足會重新檢核
'         'Add By Sindy 2013/10/15 若為補文件時,檢查新申請案是否為電子送件,若是,電子檔是否全數歸檔
'         'If m_CP01 = "P" And m_PA09 = "000" And m_CP10 = "202" Then
'         'Add By Sindy 2013/10/15 不要控管案件性質
'         If m_CP01 = "P" And m_PA09 = "000" Then
'            strExc(0) = "select cp09,cp10,cpm26 from caseprogress,casepropertymap" & _
'                        " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
'                        " and cp10 in(" & NewCasePtyList & ")" & _
'                        " and cp57 is null and cp118 is not null and cp120='Y' and cp121 is null" & _
'                        " and cp01=cpm01(+) and cp10=cpm02(+)"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               Call UpdateCP121(RsTemp.Fields("cp09"), RsTemp.Fields("cp10"), "" & RsTemp.Fields("cpm26"), m_EEP01)
'            End If
'         End If
'         '2013/10/15 END
      
      '*************** 承辦單 ***************
      'Modify By Sindy 2016/11/29 P案均不列印承辦單
      'Modify By Sindy 2018/10/22 Mark : CFP也不列印承辦單了
'      If bolPAFlow = True Then 'Add By Sindy 2018/5/10 + if
'         If m_CP01 <> "P" Then
'   '      'Add By Sindy 2014/12/17
'   '      If strSrvDate(1) >= P台灣案電子化啟用日 And (m_CP01 = "P" And m_PA09 = "000") Then
'   '         'P台灣案從2015/1/1起不列印承辦單
'   '      Else
'   '      '2014/12/17 END
'            If bolStarHasWorkSheet = True Then '已產生過電子承辦單
'               If MsgBox("是否要重新列印承辦單？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'                  GoTo RunExit
'               End If
'            End If
'            For ii = 0 To lstAtt(0).ListCount - 1
'               If InStr(UCase(lstAtt(0).List(ii)), UCase(EMP_承辦單)) > 0 Then
'                  lstAtt(0).Selected(ii) = True
'                  Exit For
'               End If
'            Next ii
'            Call cmdPrintAtt_Click(0) '列印
'         End If
'      End If
      '2018/10/22 END
      
'      'Modify By Sindy 2013/9/30 P台灣案電子送件當智權人員是分所時,要列印2張承辦單
'      'Modify By Sindy 2014/5/7 玲玲:待送件區,承辦單的列印若為Ｂ類收文的案件性質請只列印一張承辦單 ==> + And Left(Trim(lblCP09), 1) <> "B"
'      If m_CP01 = "P" And m_PA09 = "000" And m_CP118 <> "" And PUB_GetST06(m_CP13) <> "1" And Left(Trim(lblCP09), 1) <> "B" Then
'         If strSrvDate(1) < P台灣案電子化啟用日 Then 'Add By Sindy 2014/12/17 P台灣案從2015/1/1起不列印承辦單
'            For ii = 0 To lstAtt(0).ListCount - 1
'               If InStr(UCase(lstAtt(0).List(ii)), UCase(EMP_承辦單)) > 0 Then
'                  lstAtt(0).Selected(ii) = True
'                  Exit For
'               End If
'            Next ii
'            Call cmdPrintAtt_Click(0) '列印
'         End If
'      End If
      '*************** 承辦單 END ***************
      
      'Modify By Sindy 2013/9/30 P台灣案非電子送件第一次執行此功能時,要列印全部PDF
      'Modify By Sindy 2016/11/29 941分析及所有C類的案件性質,請於產生承辦單及歸檔時列印.PDF檔
      If (m_CP01 = "P" And m_PA09 = "000" And m_CP118 = "") Or _
         (m_CP01 = "P" And (m_CP10 = "941" Or Left(Trim(lblCP09), 1) = "C")) Then
         If bolStarHasWorkSheet = False And PUB_ChkEmpFlowExists(m_EEP01, EMP_退件重送) = False Then
            Call cmdSelAllPrt_Click
         End If
      End If
   End If
   
RunExit:
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   'Add By Sindy 2018/9/20
   If cmdSend.Visible = True Then
      cmdSend.SetFocus
      cmdSend.Default = True
   End If
   '2018/9/20 END
   Exit Sub
   
ErrHand:
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   MsgBox Err.Description & vbCrLf & " 歸檔資料夾有誤(m_strFolder): " & m_strFolder
End Sub

'歸檔
Private Function ToKeepOnFiles() As Boolean
Dim intTempEEP02 As Integer
'Add By Sindy 2024/7/5
Dim bolSaveDwg As Boolean
Dim bolInsFile As Boolean
'2024/7/5 END
   
   ToKeepOnFiles = False
   bolSaveDwg = False 'Add By Sindy 2024/7/5
   
'   '先刪除卷宗區及原始檔區
''   strSql = "delete from CasePaperPDF where cpp01='" & m_EEP01 & "'"
''   cnnConnection.Execute strSql
''   strSql = "delete from CasePaperFile where cpf01='" & m_EEP01 & "'"
''   cnnConnection.Execute strSql
'   'Add By Sindy 2014/7/22 未有承辦單時，不需要先刪卷宗區資料再新增
'   If bolhaveEfile = True Then
'   '2014/7/22 END
'      If DelAttFile_PDF(lblCaseNo.Caption, m_EEP01, "", "S") = False Then Exit Function
'      If DelAttFile_File(lblCaseNo.Caption, m_EEP01, "", "S") = False Then Exit Function
'   End If
   
   '讀取最近一筆.dwg檔做歸檔 and instr(upper(eef03),'DWG')>0 and instr(upper(eef03),'.PDF')=0 ==> and substr(upper(eef03),-4)='.DWG'
   lstAtt(0).Clear
   'Modify By Sindy 2013/11/6
'   strExc(0) = "select eef02,eef03,eef04,eef09,eef10 from EmpElectronFile where eef01='" & m_EEP01 & "'" & _
'               " and (substr(upper(eef03),-4)='.DWG' or substr(upper(eef03),-7)='DWG.ZIP')" & _
'               " order by eef02 desc"
   'Modify By Sindy 2014/3/11 +dwg.7z
   'Modify By Sindy 2020/2/25 ex:CFP-31474,CFP-31540 還有人判發後還在用聯絡送圖檔,所以鎖住只做到判發(前)的歸卷
   '                          + and eef02<=" & m_AttEEP02 & "
   strExc(0) = "select eef02,eef03,eef04,eef09,eef10 from EmpElectronFile where eef01='" & m_EEP01 & "'" & _
               " and eef02=(select max(eef02) from empelectronfile where eef01='" & m_EEP01 & "' and eef02<=" & m_AttEEP02 & " and (substr(upper(eef03),-4)='.DWG' or substr(upper(eef03),-7)='DWG.ZIP' or substr(upper(eef03),-6)='DWG.7Z'))" & _
               " and (substr(upper(eef03),-4)='.DWG' or substr(upper(eef03),-7)='DWG.ZIP' or substr(upper(eef03),-6)='DWG.7Z')" & _
               " order by eef02 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   intTempEEP02 = 0
   If intI = 1 Then
      bolSaveDwg = True 'Add By Sindy 2024/7/5
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            intTempEEP02 = .Fields("eef02")
            lstAtt(0).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)" & " #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "#", 0
            lstAtt(0).ItemData(0) = 1
            .MoveNext
         Loop
      End With
   End If
   If intTempEEP02 > 0 Then
      Call DownloadAllAttachFile(intTempEEP02, 0) '下載附件
      'Add By Sindy 2013/10/8 檢查檔案是否開啟中
      If ChkInsFileOpening(0) = False Then
         Screen.MousePointer = vbDefault
         Exit Function
      End If
      '2013/10/8 END
      '將檔案存至卷宗區或原始檔區
      If InsertFileData("無須踢除的檔案", 0) = False Then
         Screen.MousePointer = vbDefault
         Exit Function
      End If
   End If
   '讀取最近一筆.dwg.pdf檔做歸檔 and instr(upper(eef03),'DWG.PDF')>0 ==> and substr(upper(eef03),-7)='DWG.PDF'
   'Modify By Sindy 2014/6/27 電子送件者不需留dwg.pdf
   'If m_CP118 = "" Then '非電子送件,才要存dwg.pdf
   'Modify By Sindy 2014/9/16 開放電子送件發明及新型的新申請案可放dwg.pdf
   If m_CP118 = "" Or _
      (m_CP118 <> "" And (Trim(lblPA08.Caption) = "發明" Or Trim(lblPA08.Caption) = "新型") And InStr(NewCasePtyList, m_CP10) > 0) Then
   '2014/9/16 END
   '2014/6/27 END
      lstAtt(0).Clear
      'Modify By Sindy 2013/11/6
   '   strExc(0) = "select eef02,eef03,eef04,eef09,eef10 from EmpElectronFile where eef01='" & m_EEP01 & "'" & _
   '               " and substr(upper(eef03),-7)='DWG.PDF'" & _
   '               " order by eef02 desc"
      'Modify By Sindy 2020/2/25 ex:CFP-31474,CFP-31540 還有人判發後還在用聯絡送圖檔,所以鎖住只做到判發(前)的歸卷
      '                          + and eef02<=" & m_AttEEP02 & "
      strExc(0) = "select eef02,eef03,eef04,eef09,eef10 from EmpElectronFile where eef01='" & m_EEP01 & "'" & _
                  " and eef02=(select max(eef02) from empelectronfile where eef01='" & m_EEP01 & "' and eef02<=" & m_AttEEP02 & " and substr(upper(eef03),-7)='DWG.PDF')" & _
                  " and substr(upper(eef03),-7)='DWG.PDF'" & _
                  " order by eef02 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      intTempEEP02 = 0
      If intI = 1 Then
         bolSaveDwg = True 'Add By Sindy 2024/7/5
         With RsTemp
            .MoveFirst
            Do While Not .EOF
               intTempEEP02 = .Fields("eef02")
               lstAtt(0).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)" & " #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "#", 0
               lstAtt(0).ItemData(0) = 1
               .MoveNext
            Loop
         End With
      End If
      If intTempEEP02 > 0 Then
         Call DownloadAllAttachFile(intTempEEP02, 0) '下載附件
         'Add By Sindy 2013/10/8 檢查檔案是否開啟中
         If ChkInsFileOpening(0) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
         End If
         '2013/10/8 END
         '將檔案存至卷宗區或原始檔區
         If InsertFileData("無須踢除的檔案", 0) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
         End If
      End If
   End If
   
   '下載申請書附件
   Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), False)
   Call DownloadAllAttachFile(CInt(m_AttEEP02), 0)
   'Add By Sindy 2013/10/8 檢查檔案是否開啟中
   If ChkInsFileOpening(0) = False Then
      Screen.MousePointer = vbDefault
      Exit Function
   End If
   '2013/10/8 END
   '將電子檔分別存至卷宗區及原始檔區
   'Add By Sindy 2024/7/5
   'If InsertFileData("DWG", 0) = True Then
   'Modify By Sindy 2024/8/29 +剔除MSG
'   If bolSaveDwg = True Then bolInsFile = InsertFileData("DWG", 0) '前面有歸DWG,所以此處排除
'   If bolSaveDwg = False Then bolInsFile = InsertFileData("無須踢除的檔案", 0) '前面未歸DWG,無需特別排除
   If bolSaveDwg = True Then bolInsFile = InsertFileData("DWG,MSG", 0)
   If bolSaveDwg = False Then bolInsFile = InsertFileData("MSG", 0)
   '2024/8/29 END
   If bolInsFile = True Then
   '2024/7/5 END
      '下載存卷附件
      Call ReadAttachFile_other(m_EEP01)
      Call DownloadAllAttachFile(0, 1)
      'Add By Sindy 2013/10/8 檢查檔案是否開啟中
      If ChkInsFileOpening(1) = False Then
         Screen.MousePointer = vbDefault
         Exit Function
      End If
      '2013/10/8 END
      '將"存卷區"電子檔分別存至卷宗區及原始檔區
      If InsertFileData("無須踢除的檔案", 1) = True Then
         Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
      End If
   End If
   
   ToKeepOnFiles = True
End Function

'下載全部附件
'Modify By Sindy 2015/9/11
Private Function DownloadAllAttachFile(intEEP02 As Integer, Index As Integer, _
                                       Optional ByRef pFiles As String)
Dim stFileName As String, stFileNameErr As String
   
   pFiles = "" 'Add By Sindy 2015/9/11 記錄附件區電子檔的完整路徑及檔名
   KillAttach
   stFileNameErr = ""
   For ii = 0 To lstAtt(Index).ListCount - 1
      'If lstAtt(index).Selected(ii) Then
         stFileName = lstAtt(Index).List(ii)
         If InStrRev(stFileName, " (") > 0 Then
            stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
         End If
         'Modify By Sindy 2018/7/10
         If Right(UCase(stFileName), 5) <> UCase(".menu") Then
         '2018/7/10 END
            If InStr(stFileName, "\") = 0 Then
               If Index = 1 Then '存卷資料
                  If GetAttachFile(stFileName, 0) = False Then
                     stFileNameErr = stFileNameErr & stFileName & " 和 "
                     'Exit Function
                  End If
               Else
                  If GetAttachFile(stFileName, intEEP02) = False Then
                     stFileNameErr = stFileNameErr & stFileName & " 和 "
                     'Exit Function
                  End If
               End If
               'stFileName = m_AttachPath & "\" & stFileName 'Add By Sindy 2015/9/11
            End If
         End If
         'Modify By Sindy 2015/9/10
         'Modify By Sindy 2020/9/29 + EMP_多案承辦單
         If InStr(UCase(stFileName), UCase(EMP_承辦單)) = 0 And InStr(UCase(stFileName), UCase(EMP_多案承辦單)) = 0 Then
            pFiles = pFiles & ";" & stFileName 'Add By Sindy 2015/9/11
         End If
         '2015/9/10 END
      'End If
   Next ii
   If pFiles <> "" Then pFiles = Mid(pFiles, 2) 'Add By Sindy 2015/9/11
   If stFileNameErr <> "" Then
      stFileNameErr = Left(Trim(stFileNameErr), Len(Trim(stFileNameErr)) - 1)
      MsgBox "下載附件檔有誤！(" & stFileNameErr & ")"
   End If
End Function

'全部取消
Private Sub CallSelCancel(Index As Integer)
   Dim ii As Integer, oList As ListBox
   
   Set oList = lstAtt(Index)
   For ii = 0 To oList.ListCount - 1
      lstAtt(Index).Selected(ii) = False
   Next
End Sub

'Add By Sindy 2025/10/20
Private Sub cmdPrintAllPDF_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_ProState = "FCP" And m_strFolder <> "" Then
      cmdPrintAllPDF.ToolTipText = m_strFolder
   End If
End Sub

'全選
Private Sub cmdSelect_Click(Index As Integer)
   Dim ii As Integer, oList As ListBox
   
   Set oList = lstAtt(Index)
   For ii = 0 To oList.ListCount - 1
      lstAtt(Index).Selected(ii) = True
   Next
End Sub

'Add By Sindy 2019/6/28
Private Sub cmdTrans_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmdTrans.ToolTipText = "電子檔只會抓回來 .Data.doc 或 .Data.docx "
End Sub

Private Sub Form_Activate()
   If lblCM10.Visible = True And lblCM10.Tag <> "" And _
      (m_CP10 = "101" Or m_CP10 = "102") Then
      MsgBox "本案為一案兩請，另一案為" & lblCM10.Tag & "，請確認是否已判發。", vbInformation
      lblCM10.Tag = ""
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   lstAtt(2).Tag = lstAtt(2).Height 'Add By Sindy 2025/10/14
   Me.txtEEP02.BackColor = &H8000000F
   Me.txtEEP03.BackColor = &H8000000F
   Me.txtEEP03_2.BackColor = &H8000000F
   ReDim m_FilesRemoved(0)
   'Modify By Sindy 2021\5\19
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath")
   'Modify By Sindy 2022/6/22
   If PUB_ChkDir(m_AttachPath) = False Then
      MkDir m_AttachPath
   End If
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   If PUB_ChkDir(m_AttachPath) = False Then
      MkDir m_AttachPath
   End If
   '2022/6/22 END
   '2021\5\19 END
   
'   m_UploadAttPath = "C"
   SSTab1.Tab = 0
   If m_FlowUserNum = "" Then m_FlowUserNum = strUserNum 'Add By Sindy 2025/1/15 案件流程所屬人員
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True 'Modified by Morgan 2020/10/30 +只顯示有效的印表機參數
   strPrinter = PUB_GetOsDefaultPrinter '抓控制台印表機
   
   'Modify By Sindy 2014/9/3
   'SetFileAssociation
   txtPDFPath = PUB_SetFileAssociation
   '2014/9/3 END
   
   ReDim pa(TF_PA) 'Add By Sindy 2018/11/19
   ReDim sp(tf_SP) 'Add By Sindy 2018/11/19
   ReDim tm(TF_TM) 'Add By Sindy 2018/11/19
   ReDim cp(TF_CP) 'Add By Sindy 2018/11/19
   ReDim lC(TF_LC) 'Add By Sindy 2021/9/2
   ReDim hc(TF_HC) 'Add By Sindy 2021/9/2
   
   'Add By Sindy 2019/6/6 設定在頂層
   Text2.ZOrder '存卷資料文字框
   Frame7.BorderStyle = 0 'Add By Sindy 2023/9/23
   
   'Add By Sindy 2025/1/23
   If Pub_StrUserSt03 = "M51" Then
      Me.cmdM51.Visible = True
   Else
      Me.cmdM51.Visible = False
   End If
   '2025/1/23 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '因為無法判斷出列印動作結束沒,所以印表機切回的時間點改到Form結束時
   PUB_SetOsDefaultPrinter strPrinter
   
   KillAttach
   
   'Add by Sindy 2022/12/17 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q") = True Then
      Unload frm090801_Q
   End If
   '2022/12/17 END
   
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   Set m_PrevForm = Nothing
   Set frm090202_4_1 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
      DoEvents 'Add By Sindy 2017/3/10
   End If
End Sub

Private Sub GRD1_DblClick()
   GRD1.Visible = False
   If GRD1.MouseRow <> 0 And GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
      '上一筆資料列清除反白
      If dblPrevRow > 0 Then
         GRD1.col = 2
         GRD1.row = dblPrevRow
         For ii = 0 To GRD1.Cols - 1
            GRD1.col = ii
            GRD1.CellBackColor = QBColor(15)
         Next ii
      End If
      '目前資料列反白
      GRD1.col = 0
      GRD1.row = GRD1.MouseRow
      dblPrevRow = GRD1.row
      For ii = 0 To GRD1.Cols - 1
         GRD1.col = ii
         GRD1.CellBackColor = &HFFC0C0
      Next ii
      Call ReadData(False)
      'Add By Sindy 2014/1/17
      '待送件區時才可以新增,刪除
      If (UCase(m_PrevForm.Name) = UCase("frm090202_4") Or UCase(m_PrevForm.Name) = UCase("frm090202_7")) And _
         Val(txtEEP02) = Val(m_AttEEP02) Then
         Me.cmdAddAtt(1).Visible = True
         Me.cmdRemAtt(1).Visible = True
      Else
         Me.cmdAddAtt(1).Visible = False
         Me.cmdRemAtt(1).Visible = False
      End If
      '2014/1/17 END
   End If
   GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

'顯示明細資料於畫面上
'bolFirst : True.一進入此作業的一開始查詢動作
Private Sub ReadData(bolFirst As Boolean)
Dim intCF10 As Integer
   
   'Call ClearData
   Call SetCtrlReadOnly(False)
   
   If bolFirst = True Then
      '若有資料游標停在第一筆
      GRD1.Visible = False
      GRD1.col = 0
      'Modify By Sindy 2017/8/14
      For ii = 0 To GRD1.Rows - 1
         If m_AttEEP02 = GRD1.TextMatrix(ii, 0) Then
            GRD1.row = ii
            Exit For
         End If
      Next ii
      '2017/8/14 END
      dblPrevRow = GRD1.row
      For ii = 0 To GRD1.Cols - 1
         GRD1.col = ii
         GRD1.CellBackColor = &HFFC0C0
      Next ii
      GRD1.Visible = True
   End If
   
   'Add By Sindy 2017/8/14 王副總提出歷程判發(中)後還是可以開放聯絡
   'm_AttEEP02 = grd1.TextMatrix(dblPrevRow, 0)
   txtEEP02 = GRD1.TextMatrix(dblPrevRow, 0)
   txtEEP03 = GRD1.TextMatrix(dblPrevRow, 1)
   txtEEP03_2 = GRD1.TextMatrix(dblPrevRow, 2)
   CboEEP04.Text = GRD1.TextMatrix(dblPrevRow, 3) & " " & GRD1.TextMatrix(dblPrevRow, 12) '4
   CboEEP05.Text = GRD1.TextMatrix(dblPrevRow, 5) & " " & GRD1.TextMatrix(dblPrevRow, 6)
   txtEEP10_2 = GRD1.TextMatrix(dblPrevRow, 8)
   txtEEP08 = GRD1.TextMatrix(dblPrevRow, 9)
   'Add By Sindy 2015/9/4 P非台灣案新增歷程作業進來時,開放內容可以輸入
   If UCase(m_PrevForm.Name) = UCase("frm090202_7") Then
      txtEEP08.Locked = False
   End If
   '2015/9/4 END
   txtEEP10 = GRD1.TextMatrix(dblPrevRow, 10)
   '讀取附件
   Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
   
   'Add By Sindy 2023/11/9 外專要檢查台灣案發文前是否有需要送判程序主管
   cmdFlow(0).Tag = "" 'Add By Sindy 2024/6/21
   If UCase(m_PrevForm.Name) = UCase("frm090202_4") And Val(txtEEP02) = Val(m_AttEEP02) Then
      'Modify By Sindy 2024/8/14 +外商判斷
      If (bolFCPFlow = True And m_PA09 = "000" And PUB_GetST03(m_CP14) <> "F22" _
            And InStr(不需操作程序送判的案件性質, cp(10)) = 0) _
         Or _
         (bolFCTFlow = True And m_PA09 = "000" And PUB_GetST03(m_EP05) <> "F12") Then
         'Modify By Sindy 2024/1/18 改共用函數
         intCF10 = PUB_ChkhadCF10forEMP_46(m_CP01, m_PA09, m_CP10, m_EEP01, m_AttEEP02)
         If intCF10 > 0 Then '有主管機關者
            cmdFlow(0).Enabled = True
            cmdFlow(0).Tag = "Y" 'Add By Sindy 2024/6/21
            cmdFlow(0).BackColor = &HC0C0FF '變色,需程序送判
            If intCF10 = 2 Then '送判過了
               cmdFlow(0).BackColor = &H8000000F
            Else 'If cmdFlow(0).BackColor = &HC0C0FF Then '需程序送判
               If bolhaveEfile = True Then
                  cmdFlow(0).Enabled = False '產生承辦單就不能送判了
               Else
                  cmdPrintAllPDF.Enabled = False '未送判,不能歸檔
               End If
            End If
         End If
         '2024/1/18 END
      'Add By Sindy 2025/1/22 +CFT外商程序人員操作時,要做送判
      ElseIf (bolCFTFlow = True And PUB_GetST03(Trim(Left(CboEEP05.Text, 5))) = "F12") Then
         cmdFlow(0).Enabled = True
         cmdFlow(0).Tag = "Y" 'Add By Sindy 2024/6/21
      Else
         cmdFlow(0).Enabled = False
      End If
   End If
   '2023/11/9 END
End Sub

'查詢附件檔
Private Sub ReadAttachFile(strEEP01 As String, intEEP02 As Integer, bolQuery As Boolean)
Dim intEEFCnt As Integer 'Add By Sindy 2013/11/18
Dim strChkFileName As String 'Add By Sindy 2021/6/3
Dim strCPP12 As String 'Add By Sindy 2023/4/17
   
   KillAttach
   lstAtt(0).Clear
   intEEFCnt = 0
   
   'Modify By Sindy 2023/11/22
   'If bolFCPFlow = True Then
   'Modify By Sindy 2024/8/29 + Or bolFCTFlow = True
   If m_ProState = "FCP" Or bolFCTFlow = True Then
      Frame1.Caption = "最終送件附件區：(歷程順序 " & intEEP02 & ")"
      'Add By Sindy 2025/10/20
      If m_ProState = "FCP" And cp(118) <> "" And pa(9) = "000" Then
         m_strFolder = GetFCPPathVal("送件區", pa(1), pa(2), Trim(lblCP10.Caption), True, PUB_Getdesktop)
         If Dir(m_strFolder, vbDirectory) = "" Then
            'm_strFolder = ""
         End If
         Text1.Text = m_strFolder
      End If
      '2025/10/20 END
   Else
   '2023/11/22 END
      Frame1.Caption = "歸檔附件區：(歷程順序 " & intEEP02 & ")" 'Add By Sindy 2020/7/31
   End If
   
   'sort : 1.承辦單 2.其他或說明書 3.圖
   'Modify By Sindy 2018/9/4 + and eef12 is not null ==> Sindy 2023/11/10 拿掉 and eef12 is not null: 因外專操作上才知按過歸卷按鈕了
   'Modify By Sindy 2020/9/29 + EMP_多案承辦單
   strExc(0) = "select eef03,eef04,eef09,eef10," & _
                      "decode(sign(instr(upper(eef03),upper('" & EMP_承辦單 & "'))),1,1,decode(sign(instr(upper(eef03),upper('" & EMP_多案承辦單 & "'))),1,1,decode(sign(instr(upper(eef03),upper('DWG'))),1,3,2))) as sort" & _
               " from EmpElectronFile where eef01='" & strEEP01 & "' and eef02=" & intEEP02 & _
               " order by sort desc,eef03 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         intEEFCnt = .RecordCount
         .MoveFirst
         Do While Not .EOF
            'Add By Sindy 2018/7/10
            If Right(UCase(.Fields("eef03")), 5) = UCase(".menu") Then
               lstAtt(0).AddItem .Fields("eef03") & " (0 KB)" & " #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "#", 0
            Else
            '2018/7/10 END
               lstAtt(0).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)" & " #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "#", 0
            End If
            lstAtt(0).ItemData(0) = 1
            'Add By Sindy 2024/11/27 附件區.MSG不歸卷,所以不算檔案數
            If Right(UCase(.Fields("eef03")), 4) = UCase(".MSG") Then
               intEEFCnt = intEEFCnt - 1
            End If
            '2024/11/27 END
            .MoveNext
         Loop
      End With
   End If
   'Modify By Sindy 2013/11/8 發現若像電子送件是有可能沒有附件檔,但需要產生承辦單
   If bolQuery = True Then
      Me.cmdOpenAtt(0).Enabled = True
      Me.cmdSelect(0).Enabled = True
      Me.cmdSaveAtt(0).Enabled = True
      '待送件區時才可以新增,刪除,產生承辦單電子檔
      'Modify By Sindy 2014/1/17
      'If UCase(m_PrevForm.Name) = UCase("frm090202_4") Then
      If (UCase(m_PrevForm.Name) = UCase("frm090202_4") Or UCase(m_PrevForm.Name) = UCase("frm090202_7")) And _
         Val(txtEEP02) = intEEP02 Then
      '2014/1/17 END
         Me.cmdAddAtt(0).Visible = True
         Me.cmdRemAtt(0).Visible = True
'            'Add By Sindy 2013/9/24
'            If Left(m_EEP01, 1) = "B" And m_CP10 = 延期 Then
'               Me.cmdBack.Visible = False 'B類延期不可執行退回
'            Else
'            '2013/9/24 END
            If UCase(m_PrevForm.Name) = UCase("frm090202_7") Then
               Me.cmdBack.Visible = False
            Else
               Me.cmdBack.Visible = True
            End If
'            End If
         Me.Frame2.Visible = True
         'Me.cmdSend.Visible = True: bolhaveEfile = True
         If m_PA09 = "000" And bolPAFlow = True Then
            Me.Label4.Visible = True '註:電子送件，請先加入下載的檔案後，再執行產生承辦單。
         Else
            Me.Label4.Visible = False
         End If
      Else
         Me.cmdAddAtt(0).Visible = False
         Me.cmdRemAtt(0).Visible = False
         Me.cmdBack.Visible = False
         Me.Frame2.Visible = False
         Me.cmdSend.Visible = False: bolhaveEfile = False
         Me.cmdPrintAtt(0).Enabled = True
         Me.Label4.Visible = False
      End If
      
      'Add By Sindy 2018/11/21 待轉檔區
      If m_NPManKind = "4" Then
         cmdTrans.Visible = False '轉檔完成
         If Dir(m_strFolder & m_strCaseNo & "*.data*.doc") <> "" Or _
            Dir(m_strFolder & m_strCaseNo & "*.data*.docx") <> "" Then
            cmdPrintAllPDF.BackColor = &H80C0FF '已產生
            cmdTrans.Visible = True '轉檔完成
         End If
      End If
      
      '列印及發文按鈕必須案件文件卷宗區有承辦單PDF,並且卷宗區+原始檔區的檔案數不可小於歷程附件區的檔案數,才會亮起來
      'Modify By Sindy 2020/9/29 + EMP_多案承辦單
      'Modify By Sindy 2021/6/3
      'Modify By Sindy 2023/4/25 T*或FCT增加.menu判斷,因外商信件也會用到承辦單KeyWord
      If cp(163) <> "" Then
         strChkFileName = EMP_多案承辦單 & ".menu"
      Else
         'Modify By Sindy 2023/11/10
         If bolPAFlow = True Then
            strChkFileName = EMP_承辦單
         Else
            strChkFileName = EMP_承辦單 & ".menu"
         End If
'         If Left(m_CP01, 1) = "T" Or m_CP01 = "FCT" Then
'            strChkFileName = EMP_承辦單 & ".menu"
'         Else
'            strChkFileName = EMP_承辦單
'         End If
         '2023/11/10 END
      End If
'      strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & strEEP01 & "'" & _
'                  " and (instr(upper(cpp02),upper('" & EMP_承辦單 & "'))>0 or instr(upper(cpp02),upper('" & EMP_多案承辦單 & "'))>0)"
      'Modify By Sindy 2023/4/17 +,CPP12
      strExc(0) = "select cpp02,CPP12 from casepaperpdf where cpp01='" & strEEP01 & "'" & _
                  " and instr(upper(cpp02),upper('" & strChkFileName & "'))>0"
      '2021/6/3 END
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strCPP12 = "" & RsTemp.Fields("CPP12") 'Add By Sindy 2023/4/17
         'Modify By Sindy 2024/6/25
         If cp(163) <> "" Then
            strExc(0) = "select cpp02 from casepaperpdf,caseprogress" & _
                        " where cp163='" & strEEP01 & "' and cpp01=cp09 and cpp12='S'" & _
                        " Union all select cpf02 from casepaperfile,caseprogress" & _
                        " where cp163='" & strEEP01 & "' and cpf01=cp09 and cpf11='S'"
         Else
         '2024/6/25 END
            strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & strEEP01 & "' and cpp12='S'" & _
                        " Union all select cpf02 from casepaperfile where cpf01='" & strEEP01 & "' and cpf11='S'"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modify By Sindy 2024/1/2 + Or bolFCPFlow = True : 外專附件區不歸卷,不需要控管件數
            'Modify By Sindy 2024/8/29 外商FC同外專
            If RsTemp.RecordCount >= intEEFCnt Or bolFCPFlow = True Or bolFCTFlow = True Then
   '            strExc(0) = "select cpf02 from casepaperfile where cpf01='" & strEEP01 & "'"
   '            intI = 1
   '            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '            If intI = 1 Then
               Me.cmdPrintAtt(0).Enabled = True
               'Add By Sindy 2018/11/21 + 非待轉檔區
               If Not m_NPManKind = "4" Then
                  'Add By Sindy 2018/7/13 有承辦單示為已歸檔
                  If lstAtt(0).ListCount > 0 Then
                     'Add By Sindy 2018/9/20
                     If cmdPrintAllPDF.Caption <> "E-Mail及歸檔" Then
                     '2018/9/20 END
                        cmdPrintAllPDF.Caption = "已歸檔" 'Add By Sindy 2018/7/23
                     End If
                     'Add By Sindy 2023/4/17 此承辦單若非S(待送件區)產生的,請程序人員做確認
                     If strCPP12 <> "S" Then
                        MsgBox "卷宗區已有(類似)承辦單電子檔存在，但非待送件區產生，" & vbCrLf & vbCrLf & "請程序人員做確認處理後，再操作此作業！", vbExclamation
                        cmdPrintAllPDF.Enabled = False
                        Exit Sub
                     End If
                     '2023/4/17 END
                     cmdPrintAllPDF.BackColor = &H80C0FF
                  End If
               End If
               '2018/7/13 END
               'Modify By Sindy 2014/1/17
               'If UCase(m_PrevForm.Name) = UCase("frm090202_4") Then
               If (UCase(m_PrevForm.Name) = UCase("frm090202_4") Or UCase(m_PrevForm.Name) = UCase("frm090202_7")) And _
                  Val(txtEEP02) = intEEP02 Then
               '2014/1/17 END
                  Me.cmdSelAllPrt.Enabled = True
                  'Modify By Sindy 2018/11/21 +  Or m_NPManKind = "4"
                  'Modified by Morgan 2025/2/7 改3號區也要能發文
                  'If UCase(m_PrevForm.Name) = UCase("frm090202_7") Or m_NPManKind = "3" Or m_NPManKind = "4" Then
                  If UCase(m_PrevForm.Name) = UCase("frm090202_7") Or m_NPManKind = "4" Then
                     Me.cmdSend.Visible = False
                  Else
                     cmdFlow(0).Enabled = False 'Add By Sindy 2023/11/9 產生承辦單就不能送判了
                     Me.cmdSend.Visible = True
                  End If
                  bolhaveEfile = True
               End If
'               Me.cmdAddAtt.Visible = False
'               Me.cmdRemAtt.Visible = False
'            End If
            End If
         End If
      End If
      'Add By Sindy 2021/10/15 ACS發文時，歷程中的所有附件都不放入卷宗區或原始檔區。
      '                        所以沒有中文檔名的問題。
      If m_CP01 = "ACS" Then
         Me.cmdSend.Visible = True
         Me.cmdPrintAllPDF.Visible = False
         Me.cmdSelAllPrt.Visible = False
         Me.cmdPrintAtt(0).Enabled = True
      End If
      '2021/10/15 END
   End If
   If lstAtt(0).ListCount > 0 Then SetListScroll lstAtt(0)
   'If m_CP01 = "P" Then Me.cmdSend.Visible = False 'Add By Sindy 2015/1/21 P案因都走電子送件較多,所以不會在此作業直接發文,因此發文鍵不用開放
End Sub

'查詢存卷區
Private Sub ReadAttachFile_other(strEEP01 As String)
   KillAttach
   lstAtt(1).Clear
   'Modify By Sindy 2018/9/4 + and eef12 is not null
   strExc(0) = "select eef03,eef04,eef09,eef10 from EmpElectronFile where eef01='" & strEEP01 & "' and eef02=0 and eef12 is not null order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Do While Not .EOF
            lstAtt(1).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)" & " #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "#", 0
            lstAtt(1).ItemData(0) = 1
            .MoveNext
         Loop
      End With
      'Modify By Sindy 2021/9/2 + Or bolOtherFlow = True
      'Modify By Sindy 2024/8/14 + Or bolCFTFlow = True Or bolFCTFlow = True
      If bolTMFlow = True Or bolOtherFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
         Text2.Left = 1600 '顯示存卷資料的紅色方塊
      End If
      Text2.Visible = True 'Add By Sindy 2019/6/6 有存卷附件時顯示字樣
   Else
      Text2.Visible = False 'Add By Sindy 2019/6/6
   End If
   '待送件區時才可以新增,刪除
   'Modify By Sindy 2014/1/17
   'If UCase(m_PrevForm.Name) = UCase("frm090202_4") Then
   If (UCase(m_PrevForm.Name) = UCase("frm090202_4") Or UCase(m_PrevForm.Name) = UCase("frm090202_7")) And _
      Val(txtEEP02) = Val(m_AttEEP02) Then
   '2014/1/17 END
      Me.cmdAddAtt(1).Visible = True
      Me.cmdRemAtt(1).Visible = True
   Else
      Me.cmdAddAtt(1).Visible = False
      Me.cmdRemAtt(1).Visible = False
   End If
   If lstAtt(1).ListCount > 0 Then SetListScroll lstAtt(1)
End Sub

Private Function SaveAttFile(intEEF02 As Integer, Index As Integer) As Boolean
Dim stFilePath As String
Dim iFileNo As Integer
Dim bytes() As Byte
Dim lngSize As Long '檔案大小
Dim adoRst As New ADODB.Recordset
Const BlockSize = 500000
Dim Numblocks As Integer
Dim LeftOver As Long
Dim UpdModifyDate As Double, UpdModifyTime As Double
Dim stFtpPath As String 'Added by Morgan 2015/4/28
   
'   cnnConnection.BeginTrans
'
On Error GoTo ErrHand
   
   SaveAttFile = True
'   For ii = 0 To lstAtt(Index).ListCount - 1
      If lstAtt(Index).ItemData(0) = 0 Then
         
         stFilePath = lstAtt(Index).List(0)
         stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
         UpdModifyDate = Mid(lstAtt(Index).List(0), InStr(lstAtt(Index).List(0), "#") + 1, 8)
         UpdModifyTime = Mid(lstAtt(Index).List(0), InStr(lstAtt(Index).List(0), "#") + 9, 6)
'         strSql = "delete from EmpElectronFile where eef01='" & m_EEP01 & "' and eef02=" & m_AttEEP02 & " and eef03='" & GetFileName(stFilePath) & "'"
'         cnnConnection.Execute strSql
      
         If iFileNo > 0 Then Close #iFileNo
         iFileNo = FreeFile
         Open stFilePath For Binary Access Read As #iFileNo
         lngSize = LOF(iFileNo)
         
         'Add By Sindy 2013/10/22
         If lngSize = 0 Then
            Close #iFileNo
            SaveAttFile = False
            ShowMsg stFilePath & MsgText(9221)
            Exit Function
         End If
         '2013/10/22 END
         
         With adoRst
            If adoRst.State = adStateClosed Then
               strExc(0) = "select * from EmpElectronFile where rownum<1"
               .CursorLocation = adUseClient
               .Open strExc(0), cnnConnection, adOpenStatic, adLockOptimistic
            End If
            .AddNew
            .Fields("eef01").Value = m_EEP01
            .Fields("eef02").Value = intEEF02 'm_AttEEP02
            .Fields("eef03").Value = GetFileName(stFilePath)
            .Fields("eef04").Value = lngSize

'Removed by Morgan 2015/5/22 不再存DB
'            Numblocks = lngSize / BlockSize
'            LeftOver = lngSize Mod BlockSize
'
'            ReDim bytes(LeftOver)
'            Get #iFileNo, , bytes()
'            .Fields("eef05").AppendChunk bytes()
'
'            ReDim bytes(BlockSize)
'            For jj = 1 To Numblocks
'                Get #iFileNo, , bytes()
'                .Fields("eef05").AppendChunk bytes()
'            Next jj
'end 2015/5/22

            .Fields("eef09").Value = UpdModifyDate
            .Fields("eef10").Value = UpdModifyTime
            Close #iFileNo
            
            'Added by Morgan 2015/4/28 檔案改放FTP
            PUB_PutFtpFile stFilePath, m_EEP01, GetFileName(stFilePath), stFtpPath, "EMPELECTRONFILE", CStr(intEEF02)
            If stFtpPath <> "" Then
               .Fields("eef11") = strSrvDate(1)
               .Fields("eef12") = stFtpPath
            End If
            'end 2015/4/28
      
            .UPDATE
         End With
      End If
'   Next ii
'   cnnConnection.CommitTrans
   Exit Function
   
ErrHand:
   SaveAttFile = False
   MsgBox " 儲存附件失敗！" & vbCrLf & Err.Description
End Function

'發文
Private Sub cmdSend_Click()
Dim strCaseNo As String
   
   'Add By Sindy 2025/10/21
   If Frame1Big.Visible = True Then
      Call cmdClose_Click
      Exit Sub
   End If
   '2025/10/21 END
   
   If ChkRevStatus("2") = False Then Exit Sub 'Add By Sindy 2014/6/4
   
   'Add By Sindy 2023/3/29
   '檢查無北所分案日期,不可發文
   If Mid(m_EEP01, 1, 1) = "A" And m_CP140 <> "" Then
      If Val(m_CP157) = "0" Then
         MsgBox "未分案！請通知程序人員查看及處理後，才能發文。"
         Exit Sub
      End If
   End If
   '2023/3/29 END
   
   strCaseNo = Trim(lblCaseNo.Caption)
   cmdExit_Click
   'Modify By Sindy 2018/5/2
   Call mdiMain.frm090202_4CallFrm(m_ProState, m_PA11, m_CP01, m_CP02, m_CP03, m_CP04, m_EEP01)
   '2018/5/2 END
End Sub

'Add By Sindy 2014/6/4 發生敏惠做退件,但有人做發文事件,重新檢查文的狀況
'1.退件 2.發文
Private Function ChkRevStatus(strType As String) As Boolean
   
   ChkRevStatus = True
   strExc(0) = "select cp09,cp27 from caseprogress" & _
               " where cp09='" & Trim(lblCP09) & "' and nvl(cp27,0)>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         ChkRevStatus = False
         MsgBox "已發文不可執行" & IIf(strType = "1", "退件", "發文") & "！", vbExclamation
         Exit Function
      End If
   End If
   
'   strExc(0) = "select eep01,eep02,eep04" & _
'               " From EmpElectronProcess" & _
'               " where eep01='" & Trim(lblCP09) & "'" & _
'               " order by eep02 desc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      If RsTemp.Fields("eep04") <> EMP_判發 And RsTemp.Fields("eep04") <> EMP_退件重送 Then
'         ChkRevStatus = False
'         MsgBox "已無判發權限，不可執行" & IIf(strType = "1", "退件", "發文") & "！", vbExclamation
'         Exit Function
'      End If
'   End If
End Function

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim objText As Object
   
   txtEEP03.Locked = Not bEnable
   txtEEP03_2.Locked = Not bEnable
   CboEEP04.Locked = Not bEnable
   CboEEP05.Locked = Not bEnable
   txtEEP10_2.Locked = Not bEnable
   txtEEP08.Locked = Not bEnable
   txt1(5).Locked = Not bEnable
   txt1(6).Locked = Not bEnable
   txt1(0).Locked = Not bEnable
   txt1(4).Locked = Not bEnable
   'Add By Sindy 2023/11/8
   For Each objText In Me.txt3
      objText.Locked = Not bEnable
   Next
   txtEP12.Locked = Not bEnable
   ChkEED13.Enabled = bEnable
   '2023/11/8 END
End Sub

'Add By Sindy 2015/9/17
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
GRD1.ToolTipText = ""
If GRD1.MouseRow <> 0 And GRD1.MouseCol > 0 Then
   If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
      GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
   End If
End If
End Sub

'Add By Sindy 2024/2/1
'\\typing2\電子送件暫存區\+案號(流水號必須足6碼)
Private Sub Label25_Click(Index As Integer)
Dim hLocalFile As Long

On Error GoTo ErrHnd 'Add by Sindy 2024/8/8
   
   strExc(10) = GetFCPPathVal(Label25(Index).Caption, m_CP01, m_CP02, Trim(lblCP10.Caption), True)
   If Dir(strExc(10), vbDirectory) <> "" Then
      ShellExecute hLocalFile, "explore", strExc(10), vbNullString, vbNullString, 1
   Else
      MsgBox "無此資料夾! " & strExc(10), vbExclamation
   End If

'Add by Sindy 2024/8/8
   Exit Sub
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Number & ": " & Err.Description, , Label25(Index).Caption
   End If
'2024/8/8 END
End Sub
'Add by Sindy 2025/10/20
Private Sub Label25_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   strExc(10) = GetFCPPathVal(Label25(Index).Caption, m_CP01, m_CP02, Trim(lblCP10.Caption), True)
   Label25(Index).ToolTipText = strExc(10)
End Sub
'2025/10/20 END

'Add By Sindy 2015/9/15
Private Sub lstAtt_Click(Index As Integer)
   If lstAtt(Index).List(lstAtt(Index).ListIndex) = "" Then Exit Sub
   If InStr(UCase(lstAtt(Index).List(lstAtt(Index).ListIndex)), UCase(EMP_承辦單 & ".menu")) > 0 Or _
      InStr(UCase(lstAtt(Index).List(lstAtt(Index).ListIndex)), UCase(EMP_多案承辦單 & ".menu")) > 0 Then
      lstAtt(Index).Selected(lstAtt(Index).ListIndex) = False
   End If
End Sub

'Add By Sindy 2025/10/8 點二下可以開啟附件檔案
Private Sub lstAtt_DblClick(Index As Integer)
   Call cmdOpenAtt_Click(Index)
End Sub

'Add By Sindy 2019/6/6
Private Sub Text2_Click()
   SSTab1.Tab = intTab_存卷資料
End Sub

'Add By Sindy 2023/9/19
Private Sub TXT3_GotFocus(Index As Integer)
   TextInverse txt3(Index)
End Sub
Private Sub TXT3_LostFocus(Index As Integer)
Dim strText As String
   
   Select Case Index
      Case 3, 5, 6
         If Index = 3 Then LblEED10_N.Caption = "": LblEED10_N_2.Caption = ""
         If Index = 5 Then LblEED09_N.Caption = ""
         If Index = 6 Then LblEED06_N.Caption = ""
         strText = GetPrjSalesNM(CStr(Trim(txt3(Index).Text)))
         If strText <> "" Then
            '檢查人員是否存在或離職
            If ChkStaffST04(Trim(txt3(Index).Text), False) = True Then
               'Modify By Sindy 2024/6/17
               If Index = 5 Then '管制人
                  If cp(158) = 0 And cp(159) = 0 Then '未發文時重新帶最新管制人
                     strExc(10) = Left(PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4)), 5)
                     If strExc(10) <> Trim(txt3(Index).Text) Then
                        txt3(Index).Text = strExc(10)
                        strText = GetPrjSalesNM(CStr(Trim(txt3(Index).Text)))
                     End If
                  End If
               Else
               '2024/6/17 END
                  If txt3(Index).Enabled = True Then txt3(Index).SetFocus
                  'Exit Sub
               End If
            End If
            If Index = 3 Then
               LblEED10_N.Caption = strText
               '翻譯人員（譯者）
               'Modified by Lydia 2025/03/13 改用模組取得
               'If InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, txt3(Index)) = 0 Then
               If InStr(Pub_SetF51Order("F", ""), txt3(Index)) = 0 Then
                  If Left(txt3(Index), 1) = "F" Then
                      LblEED10_N_2 = "-下班"
                  Else
                      LblEED10_N_2 = "-上班"
                  End If
               End If
            End If
            If Index = 5 Then LblEED09_N.Caption = strText
            If Index = 6 Then LblEED06_N.Caption = strText
         End If
   End Select
End Sub

Private Sub txtEEP10_2_LostFocus()
Dim strText As String
Dim arrID
Dim strTempName As String

   If (txtEEP10_2.Text > "" Or txtEEP10.Text > "") And (txtEEP10_2.Text <> txtEEP10_2.Tag Or strEEP10_Err <> "") Then
      txtEEP10 = "": strTempName = "": strEEP10_Err = ""
      arrID = Split(txtEEP10_2.Text, ",")
      For intI = 0 To UBound(arrID)
         If IsNumeric(Mid(Trim(arrID(intI)), 2, 4)) Then
            '依員工編號抓取員工姓名
            strText = GetPrjSalesNM(CStr(arrID(intI)))
            If strText <> "" Then
               txtEEP10 = txtEEP10 & arrID(intI) & ","
               strTempName = strTempName & strText & ","
            Else
               strTempName = strTempName & arrID(intI) & ","
               strEEP10_Err = strEEP10_Err & arrID(intI) & ","
            End If
         Else
            '依員工姓名抓取員工編號
            strText = GetPrjSalesNM_2(CStr(arrID(intI)), , , , , False) 'Modify By Sindy 2021/6/22 + , , , , , False
            If strText <> "" Then
               txtEEP10 = txtEEP10 & strText & ","
               strTempName = strTempName & arrID(intI) & ","
            Else
               strTempName = strTempName & arrID(intI) & ","
               strEEP10_Err = strEEP10_Err & arrID(intI) & ","
            End If
         End If
      Next intI
      txtEEP10 = Left(txtEEP10, IIf(Len(txtEEP10) - 1 < 0, 0, Len(txtEEP10) - 1))
      txtEEP10_2.Text = Left(strTempName, IIf(Len(strTempName) - 1 < 0, 0, Len(strTempName) - 1))
      txtEEP10_2.Tag = txtEEP10_2.Text
      strEEP10_Err = Left(strEEP10_Err, IIf(Len(strEEP10_Err) - 1 < 0, 0, Len(strEEP10_Err) - 1))
'      If Trim(strEEP10_Err) <> "" Then
'         MsgBox "副本收受者資料有誤！(" & strEEP10_Err & ")"
'         txtEEP10_2.SetFocus
'         Call txtEEP10_2_GotFocus
'         'Cancel = True
'         Exit Sub
'      End If
   End If
End Sub

Private Sub txtEEP08_GotFocus()
   InverseTextBox txtEEP08
End Sub

Private Sub txtEEP08_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(txtEEP08, txtEEP08.MaxLength) = False Then
      Cancel = True
      txtEEP08_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Function BrowseForFolder(Optional sCaption As String = "請選擇欲儲存的位置", Optional sDefault As String) As String
    Const BIF_RETURNONLYFSDIRS = 1
    Const MAX_PATH = 260
    Dim lPos As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, tBrowse As BrowseInfo

    With tBrowse
        'Set the owner window
        .hwndOwner = GetActiveWindow        'Me.hWnd in VB
        .lpszTitle = sCaption
        .ulFlags = BIF_RETURNONLYFSDIRS     'Return only if the user selected a directory
    End With

    'Show the dialog
    lpIDList = SHBrowseForFolder(tBrowse)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        lPos = InStr(sPath, vbNullChar)
        If lPos Then
            BrowseForFolder = Left$(sPath, lPos - 1)
            If Right$(BrowseForFolder, 1) <> "\" Then
                BrowseForFolder = BrowseForFolder & "\"
            End If
        End If
    Else
        'User cancelled, return default path
        BrowseForFolder = sDefault
    End If
End Function

'列印
Private Sub cmdPrintAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   Dim process_id As Long
   Dim process_handle As Long
   Dim process_handle_PDF As Long
   Dim bolPDF As Boolean
   Dim program_name As String
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   strAtt = lstAtt(Index).Text
   'Add By Sindy 2014/8/27
   program_name = txtPDFPath
   bolPDF = False
   '2014/8/27 END
   
   If strAtt = "" Then
      MsgBox "請選擇欲列印的附件！"
   Else
      PUB_SetOsDefaultPrinter Combo1
      For ii = 0 To lstAtt(Index).ListCount - 1
         If lstAtt(Index).Selected(ii) Then
            bolIsSelect = True
            
            stFileName = lstAtt(Index).List(ii)
            'stFileName = strAtt
            If InStrRev(stFileName, " (") > 0 Then
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
            
            'Add By Sindy 2014/8/27
            If bolPDF = False And Right(UCase(stFileName), 4) = ".PDF" Then
               '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
               process_id = SHELL(txtPDFPath, vbHide)
               process_handle_PDF = OpenProcess(PROCESS_TERMINATE, 0, process_id)
               bolPDF = True
            End If
            DoEvents
            '2014/8/27 END
            
            If InStr(stFileName, "\") = 0 Then
               If Index = 1 Then '存卷資料
                  If GetAttachFile(stFileName, 0) = False Then Exit Sub
               Else
                  If GetAttachFile(stFileName, CInt(m_AttEEP02)) = False Then Exit Sub
               End If
            End If
            
            If Right(UCase(stFileName), 4) = ".PDF" Then
               PrintOnePdf program_name, " /n /t """ & stFileName & """ """ & Combo1 & """"
            Else
               'Modify By Sindy 2013/9/13
               process_id = ShellExecute(Me.hWnd, "print", stFileName, vbNullString, vbNullString, 1)
               process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
               If process_handle <> 0 Then
               'If process_id <> 0 Then
                  WaitForSingleObject process_handle, INFINITE
                  CloseHandle process_handle
               End If
               '2013/9/13 END
            End If
            DoEvents
         End If
      Next ii
      'Add By Sindy 2014/8/27
      If bolPDF = True Then
         TerminateProcess process_handle_PDF, 0&
         CloseHandle process_handle_PDF
         DoEvents
      End If
      '2014/8/27 END
      
      '因為無法判斷出列印動作結束沒,所以印表機切回的時間點改到Form結束時
      'PUB_SetOsDefaultPrinter strPrinter
      
      If bolIsSelect = False Then
         MsgBox "請選擇欲列印的附件！"
      Else
         Call CallSelCancel(Index) '全部取消
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2014/8/27
Private Sub PrintOnePdf(ByVal program_name As String, parameters As String)
Dim process_id As Long
Dim process_handle As Long
   
   ' Start the program.
   On Error GoTo ShellError
   
   process_id = SHELL(program_name & parameters, vbHide)
   
   On Error GoTo 0

   ' Wait for the program to finish.
   ' Get the process handle.
   process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
   If process_handle <> 0 Then
       WaitForSingleObject process_handle, INFINITE
       CloseHandle process_handle
   End If
   
   Exit Sub

ShellError:
   MsgBox " " & _
         program_name & vbCrLf & _
         Err.Description, vbOKOnly Or vbExclamation, _
         "Error"
End Sub

Private Function GetAttachFile(ByRef pFileName As String, intEEP02 As Integer, Optional pSavePath As String) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   Dim fs As Object 'Add By Sindy 2020/3/13
   Dim stModifyDateTime As String 'Added by Morgan 2025/7/18

On Error GoTo ErrHnd
   
   Set fs = CreateObject("Scripting.FileSystemObject") 'Add By Sindy 2020/3/13
   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
      '檔案已存在時
      If Dir(stAttPath) <> "" Then
         '檢查檔案是否正在使用中
         If PUB_ChkFileOpening(stAttPath) = True Then
            MsgBox stAttPath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
            Exit Function
         End If
         
         'Modify By Sindy 2020/3/13
         'Kill stAttPath
         fs.DeleteFile stAttPath, True  '刪除檔案
         '2020/3/13 END
         
         '不必重新下載
'         pFileName = stAttPath
'         GetAttachFile = True
'         Exit Function
      End If
   Else
      stAttPath = pSavePath
   End If
   
   'Added by Morgan 2015/4/28
   'Modified by Morgan 2015/5/22 FTP上線
   'Modify By Sindy 2018/7/10 + and substr(upper(eef03),-5)<>upper('.menu')
   'Modify By Sindy 2018/9/4 + and eef12 is not null
   'Modified by Morgan 2025/7/18 +eef09,eef10
      strExc(0) = "select eef12,eef09,eef10 from EmpElectronFile where eef01='" & m_EEP01 & "' and eef02=" & intEEP02 & " and eef12 is not null" & _
               " and eef03='" & ChgSQL(pFileName) & "' and substr(upper(eef03),-5)<>upper('.menu')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp(0)) Then
            pFileName = stAttPath
            'Modified by Morgan 2025/7/18 +更新檔案修改時間
            stModifyDateTime = RsTemp("eef09") & Format("" & RsTemp("eef10"), "000000")
            GetAttachFile = PUB_GetFtpFile(RsTemp(0), stAttPath, "EMPELECTRONFILE", True, , , , stModifyDateTime)
            'end 2025/7/18
         End If
      End If
      
      Set fs = Nothing 'Add By Sindy 2020/3/13
      Exit Function
   'end 2015/4/28
   
'Removed by Morgan 2015/5/22 不再存DB
'   strExc(0) = "select * from EmpElectronFile where eef01='" & m_EEP01 & "' and eef02=" & intEEP02 & _
'               " and eef03='" & ChgSQL(pFileName) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If Dir(stAttPath) <> "" Then Kill stAttPath
'      With RsTemp
'      lngSize = Val(.Fields("eef04").Value)
'      ReDim bytes(lngSize)
'      If lngSize > 0 Then bytes() = .Fields("eef05").GetChunk(lngSize)
'      End With
'      iFileNo = FreeFile
'      Open stAttPath For Binary Access Write As #iFileNo
'      If lngSize > 0 Then Put #iFileNo, , bytes()
'      Close #iFileNo
'
'      pFileName = stAttPath
'      GetAttachFile = True
'   End If
'   Exit Function
'end 2015/5/22
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   strAtt = lstAtt(Index).Text
   
   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      For ii = 0 To lstAtt(Index).ListCount - 1
         If lstAtt(Index).Selected(ii) Then
            bolIsSelect = True
            stFileName = lstAtt(Index).List(ii)
            'stFileName = strAtt
            If InStrRev(stFileName, " (") > 0 Then
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
            
            If InStr(stFileName, "\") = 0 Then
               If Index = 1 Then '存卷資料
                  If GetAttachFile(stFileName, 0) = False Then Exit Sub
               Else
                  If GetAttachFile(stFileName, CInt(m_AttEEP02)) = False Then Exit Sub
               End If
            End If
            
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Next ii
      If bolIsSelect = False Then
         MsgBox "請選擇欲開啟的附件！"
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Sub

'列印全部PDF(不含承辦單)
Private Sub cmdSelAllPrt_Click()
'   If bolhaveEfile = True Then '已產生過電子承辦單
'      If MsgBox("是否要重新列印全部PDF？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         Exit Sub
'      End If
'   End If
   '列印全部PDF
   For ii = 0 To lstAtt(0).ListCount - 1
'            '電子送件時,不列印繪圖PDF只列印其他的PDF
'            If m_CP118 <> "" Then
'               If InStr(UCase(lstAtt(0).List(ii)), ".PDF") > 0 And InStr(UCase(lstAtt(0).List(ii)), "DWG") = 0 Then
'                  lstAtt(0).Selected(ii) = True
'               End If
'            Else
         'Modify By Sindy 2020/9/29 + EMP_多案承辦單
         If InStr(UCase(lstAtt(0).List(ii)), ".PDF") > 0 And _
            InStr(UCase(lstAtt(0).List(ii)), UCase(EMP_承辦單)) = 0 And _
            InStr(UCase(lstAtt(0).List(ii)), UCase(EMP_多案承辦單)) = 0 Then
            lstAtt(0).Selected(ii) = True
         Else
            lstAtt(0).Selected(ii) = False
         End If
'            End If
   Next ii
   
RunPrinter:
   Call cmdPrintAtt_Click(0) '列印
End Sub

'下載
Private Sub cmdSaveAtt_Click(Index As Integer)
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer, oList As ListBox
   
   'Add By Sindy 2020/1/16
   '讀取前次設定路徑
   stFolderPath = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
   If stFolderPath <> "" Then
      If PUB_ChkDir(stFolderPath) = False Then
         stFolderPath = PUB_Getdesktop
      End If
   Else
      stFolderPath = PUB_Getdesktop
   End If
   stFolderPath = PUB_GetFolder(Me.hWnd, stFolderPath, "請選取資料夾:")
   If Trim(stFolderPath) <> "" Then
      SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", stFolderPath
   Else
      Exit Sub
   End If
   If Right(Trim(stFolderPath), 1) <> "\" Then
      stFolderPath = Trim(stFolderPath) & "\"
   End If
   '2020/1/16 END
   
   Set oList = lstAtt(Index)
   stFileName = ""
   bMultiFile = False
   For ii = 0 To oList.ListCount - 1
      If oList.Selected(ii) Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = oList.List(ii)
         End If
      End If
   Next
   
   Screen.MousePointer = vbHourglass
   If stFileName = "" Then
      MsgBox "請選擇欲存檔的附件！"
   Else
      '多選
      If bMultiFile Then
         'stFolderPath = BrowseForFolder() 'Modify By Sindy 2020/1/16 Mark
         If stFolderPath <> "" Then
            For ii = 0 To oList.ListCount - 1
               If oList.Selected(ii) Then
                  stFileName = oList.List(ii)
                  If InStrRev(stFileName, " (") > 0 Then
                     stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                  End If
                  stFullName = stFolderPath & stFileName
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        If Index = 1 Then '存卷資料
                           If GetAttachFile(stFileName, 0, stFullName) = False Then
                              MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                              GoTo RunExit
                           End If
                        Else
                           If GetAttachFile(stFileName, CInt(m_AttEEP02), stFullName) = False Then
                              MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                              GoTo RunExit
                           End If
                        End If
                     End If
                  End If
               End If
            Next
         End If
      Else
         stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
         'Modify By Sindy 2020/1/16
         'stFullName = GetSaveName(stFileName, stFolderPath)
         stFullName = stFolderPath & stFileName
         '2020/1/16 END
         If stFullName <> "" Then
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               If Index = 1 Then '存卷資料
                  If GetAttachFile(stFileName, 0, stFullName) = False Then
                     MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                     GoTo RunExit
                  End If
               Else
                  If GetAttachFile(stFileName, CInt(m_AttEEP02), stFullName) = False Then
                     MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                     GoTo RunExit
                  End If
               End If
            End If
         End If
      End If
      If stFullName <> "" Then
         MsgBox "下載完成！"
      End If
   End If
RunExit:
   Screen.MousePointer = vbDefault
End Sub

'新增
Private Sub cmdAddAtt_Click(Index As Integer)
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim UpdModifyDate As Double, UpdModifyTime As Double
   Dim stFiName As String, stReName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.*"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
            For ii = 1 To UBound(sFile)
               'Add By Sindy 2013/10/9
               If InStr(CStr(sFile(ii)), "#") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                  Exit Sub
               End If
               '2013/10/9 END
               'Add By Sindy 2023/3/16
               If InStr(UCase(CStr(sFile(ii))), UCase(EMP_承辦單)) > 0 Or _
                  InStr(UCase(CStr(sFile(ii))), UCase(EMP_多案承辦單)) > 0 Then
                  MsgBox "不可新增承辦單，此為系統產出的檔案！", vbExclamation
                  Exit Sub
               End If
               '2023/3/16 END
               
               '檢查檔名規則
               'Modify By Sindy 2024/8/22 外商FC同外專不鎖中文,因附件區不進卷宗區
'               If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), "Y", m_CP10, , Index) = False Then
'                  Exit Sub
'               End If
               If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), "Y", m_CP10, , Index, _
                     IIf(bolFCPFlow = True Or bolFCTFlow = True, False, True), , , , _
                     , IIf(Index = 0 And ((m_ProState = "FCP" And bolFCPFlow = True) Or bolFCTFlow = True), True, False)) = False Then
                  Exit Sub
               End If
               '2024/8/22 END
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  Exit Sub
               End If
               '2013/9/6 END
               If AddListX(lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)) = True Then
                  '存檔
                  If Index = 1 Then '存卷資料
                     If SaveAttFile(0, Index) = False Then
                        Exit Sub
                     End If
                  Else
                     If SaveAttFile(CInt(m_AttEEP02), Index) = False Then
                        Exit Sub
                     End If
                  End If
                  If bolhaveEfile = True Then  '已有卷宗區或原始檔區
                     UpdModifyDate = Mid(Format(f.DateLastModified, "YYYYMMDDHHMMSS"), 1, 8)
                     UpdModifyTime = Mid(Format(f.DateLastModified, "YYYYMMDDHHMMSS"), 9, 6)
                     '更名
                     Call PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, CStr(sFile(ii)), stReName)
                     
                     'If InStr(UCase(stReName), ".PDF") > 0 Then
                     If Right(Trim(UCase(stReName)), 4) = ".PDF" Then
                        If SaveAttFile_PDF(m_EEP01, stFileName, stReName, UpdModifyDate, UpdModifyTime, False, "S") = False Then
                           Exit Sub
                        End If
'                        Pub_SaveLog strUserNum, "新增卷宗區附件：" & sFile(ii), m_CP01, m_CP02, m_CP03, m_CP04, m_EEP01
                     Else
                        If SaveAttFile_Org(m_EEP01, stFileName, stReName, UpdModifyDate, UpdModifyTime, "S") = False Then
                           Exit Sub
                        End If
'                        Pub_SaveLog strUserNum, "新增原始檔區附件：" & sFile(ii), m_CP01, m_CP02, m_CP03, m_CP04, m_EEP01
                     End If
                  End If
                  '重新顯示附件區
                  If Index = 1 Then '存卷資料
                     Call ReadAttachFile_other(m_EEP01)
                  Else
                     Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), False)
                  End If
               End If
            Next
         Else
            'stFileName = GetFileName(.FileName)
            'Modify By Sindy 2013/10/9
            'stFiName = GetFileName(.FileName) '不含路徑的檔名
            stFiName = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(stFiName, "#") > 0 Then
               MsgBox stFiName & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
               Exit Sub
            End If
            '2013/10/9 END
            'Add By Sindy 2023/3/16
            If InStr(UCase(stFiName), UCase(EMP_承辦單)) > 0 Or _
               InStr(UCase(stFiName), UCase(EMP_多案承辦單)) > 0 Then
               MsgBox "不可新增承辦單，此為系統產出的檔案！", vbExclamation
               Exit Sub
            End If
            '2023/3/16 END
            
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            '檢查檔名規則
            'Modify By Sindy 2024/8/22 外商FC同外專不鎖中文,因附件區不進卷宗區
'            If PUB_ChkEmpFlowFNMRule(lblCaseNo, stFiName, "Y", m_CP10, , Index) = False Then
'               Exit Sub
'            End If
            If PUB_ChkEmpFlowFNMRule(lblCaseNo, stFiName, "Y", m_CP10, , Index, _
                  IIf(bolFCPFlow = True Or bolFCTFlow = True, False, True), , , , _
                  , IIf(Index = 0 And ((m_ProState = "FCP" And bolFCPFlow = True) Or bolFCTFlow = True), True, False)) = False Then
               Exit Sub
            End If
            '2024/8/22 END
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg stFiName & MsgText(9221)
               Exit Sub
            End If
            '2013/9/6 END
            If AddListX(lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)) = True Then
               '存檔
               If Index = 1 Then '存卷資料
                  If SaveAttFile(0, Index) = False Then
                     Exit Sub
                  End If
               Else
                  If SaveAttFile(CInt(m_AttEEP02), Index) = False Then
                     Exit Sub
                  End If
               End If
               If bolhaveEfile = True Then '已有卷宗區或原始檔區
                  UpdModifyDate = Mid(Format(f.DateLastModified, "YYYYMMDDHHMMSS"), 1, 8)
                  UpdModifyTime = Mid(Format(f.DateLastModified, "YYYYMMDDHHMMSS"), 9, 6)
                  '更名
                  Call PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, stFiName, stReName)
                  
                  'If InStr(UCase(stReName), ".PDF") > 0 Then
                  If Right(Trim(UCase(stReName)), 4) = ".PDF" Then
                     If SaveAttFile_PDF(m_EEP01, stFileName, stReName, UpdModifyDate, UpdModifyTime, False, "S") = False Then
                        Exit Sub
                     End If
'                     Pub_SaveLog strUserNum, "新增卷宗區附件：" & stFiName, m_CP01, m_CP02, m_CP03, m_CP04, m_EEP01
                  Else
                     If SaveAttFile_Org(m_EEP01, stFileName, stReName, UpdModifyDate, UpdModifyTime, "S") = False Then
                        Exit Sub
                     End If
'                     Pub_SaveLog strUserNum, "新增原始檔區附件：" & stFiName, m_CP01, m_CP02, m_CP03, m_CP04, m_EEP01
                  End If
               End If
               '重新顯示附件區
               If Index = 1 Then '存卷資料
                  Call ReadAttachFile_other(m_EEP01)
               Else
                  Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), False)
               End If
            End If
         End If
      End If
   End With
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'刪除
Private Sub cmdRemAtt_Click(Index As Integer)
   RemoveList lstAtt(Index), Index
End Sub

Private Function GetSaveName(ByVal pFileName As String, ByVal pFilePath As String) As String
   
On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = pFilePath 'PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowSave
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

'Modify By Sindy 2018/12/4 + Optional bolShowMsg As Boolean = True
Private Function RemoveList(oList As ListBox, Index As Integer, _
   Optional bolShowMsg As Boolean = True) As Boolean
Dim ii As Integer
Dim bolDel As Boolean
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            
            'Add By Sindy 2015/9/15
            'Modify By Sindy 2020/9/29 + EMP_多案承辦單
            If InStr(UCase(GetFileName(oList.List(ii))), UCase(EMP_承辦單)) > 0 Or _
               InStr(UCase(GetFileName(oList.List(ii))), UCase(EMP_多案承辦單)) > 0 Then
               MsgBox "不可刪除承辦單！", vbExclamation
               Exit Function
            End If
            '2015/9/15 END
            
            'Add By Sindy 2018/12/4
            If bolShowMsg = True Then
            '2018/12/4 END
               If MsgBox("確定要刪除" & GetFileName(oList.List(ii)) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Function
            End If
            
            If oList.ItemData(ii) > 0 Then
               intI = UBound(m_FilesRemoved) + 1
               ReDim Preserve m_FilesRemoved(intI) As String
               m_FilesRemoved(intI) = GetFileName(oList.List(ii))
            End If
            
            '直接從資料庫刪除檔案
            If Index = 1 Then '存卷資料
               bolDel = DeleteFile(GetFileName(oList.List(ii)), 0)
            Else
               bolDel = DeleteFile(GetFileName(oList.List(ii)), CInt(m_AttEEP02))
            End If
            If bolDel = True Then
               oList.RemoveItem ii
               SetListScroll oList
               RemoveList = True
               ii = ii - 1
            End If
         End If
         ii = ii + 1
      Loop
   End If
End Function

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Function AddListX(oList As ListBox, stNewItem As String, oList1 As ListBox) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
      
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         'Modify By Sindy 2014/6/20
         'If GetFileName(stNewItem) = stFileName Then
         If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
         '2014/6/20 END
            MsgBox "附件 " & stFileName & " 已存在！"
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      
      If bFound = False Then
         For idx = 0 To oList1.ListCount - 1
            stFileName = GetFileName(oList1.List(idx))
            'Modify By Sindy 2014/6/20
            'If GetFileName(stNewItem) = stFileName Then
            If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
            '2014/6/20 END
               MsgBox "附件 " & stFileName & " 已存在！"
               AddListX = False
               bFound = True
               Exit For
            End If
         Next
      End If
      
      If bFound = False Then
         oList.AddItem stNewItem, 0
         SetListScroll oList
         AddListX = True
      End If
   End If
End Function

'列印承辦單
Private Sub PrintData()
Dim rsA As New ADODB.Recordset
Dim w2 As Integer
Dim PrintDetailTxt(100) As String
Dim PrintDetailTemp As Variant
Dim PrintWidthWord As Integer
Dim bolOK As Boolean
Dim dblRow As Double
Dim oJ As Integer
Dim strEEP04 As String
   
   DrawLeftMove = 1000
   DrawRightMove = 650
   '粗線深度
   DrawCount = 20
   IsHaveTaieLogo = False
   If Dir(Trim(App.path) & "\taie_logo.jpg") <> "" Then
      IsHaveTaieLogo = True
      pic1.Picture = LoadPicture(App.path & "\taie_logo.jpg")
   ElseIf Dir("c:\pics\taie_logo.jpg") <> "" Then
      IsHaveTaieLogo = True
      pic1.Picture = LoadPicture("c:\pics\taie_logo.jpg")
   End If
   
   oStrA14 = PUB_GetST07(m_CP14)
   
   '畫線 *****************************
   '裝訂線
   Printer.DrawStyle = 2
   Printer.Line (500, 0)-(500, 6300)
   Printer.Line (500, 6300 + Printer.TextHeight("裝"))-(500, 7400)
   Printer.Line (500, 7400 + Printer.TextHeight("訂"))-(500, 8600)
   Printer.Line (500, 8600 + Printer.TextHeight("線"))-(500, 17000)
   '打字
   Printer.CurrentX = 500 - (Printer.TextWidth("裝") / 2)
   Printer.CurrentY = 6300
   Printer.Print "裝"
   Printer.CurrentX = 500 - (Printer.TextWidth("訂") / 2)
   Printer.CurrentY = 7400
   Printer.Print "訂"
   Printer.CurrentX = 500 - (Printer.TextWidth("線") / 2)
   Printer.CurrentY = 8600
   Printer.Print "線"
   Printer.DrawStyle = 0
   '粗線
   For i = 1 To DrawCount
      '方格
      Printer.Line (500 + i + DrawLeftMove, 1500 + i)-(10500 + i + DrawRightMove, 2000 + i), , B
      Printer.Line (500 + i + DrawLeftMove, 2000 + i)-(10500 + i + DrawRightMove, 16100 + i), , B
      '橫線
      Printer.Line (6400 + i + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2450 + i)-(10500 + i + DrawRightMove, 2450 + i)
      Printer.Line (6400 + i + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3350 + i)-(10500 + i + DrawRightMove, 3350 + i)
      Printer.Line (6400 + i + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3800 + i)-(10500 + i + DrawRightMove, 3800 + i)
      Printer.Line (500 + i + DrawLeftMove, 8000 + i)-(10500 + i + DrawRightMove, 8000 + i) '打字繕寫日期
      '直線
      Printer.Line (6400 + i + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2450 + i)-(6400 + i + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3800 + i)
   Next i
   '細線
   '直
   Printer.Line (3300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 1500)-(3300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000)
   Printer.Line (4000 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 1500)-(4000 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000)
   Printer.Line (6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000)-(6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3800)
   Printer.Line (7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000)-(7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3800)
   Printer.Line (6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200)-(6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 8000)
   
   Printer.Line (1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 1500)-(1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200)
   
   Printer.Line (5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800)-(5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 8000)
   Printer.Line (6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800)-(6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 8000)
   Printer.Line (7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800)-(7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200)
   Printer.Line (8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800)-(8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 8000)
   '橫
   Printer.Line (500 + DrawLeftMove, 3800)-(10500 + DrawRightMove, 3800)
   Printer.Line (500 + DrawLeftMove, 4800)-(10500 + DrawRightMove, 4800)
   Printer.Line (500 + DrawLeftMove, 2900)-(10500 + DrawRightMove, 2900)
   Printer.Line (500 + DrawLeftMove, 5600)-(10500 + DrawRightMove, 5600) '備註標題欄
   Printer.Line (5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6000)-(8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6000) '齊備日
   Printer.Line (5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6400)-(8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6400) '完稿日
   Printer.Line (5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6800)-(8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6800) '會稿日
   Printer.Line (5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7200)-(8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7200) '會回日
   Printer.Line (6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7600)-(8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7600) '打字繕寫人員
   Printer.Line (500 + DrawLeftMove, 5200 + i)-(10500 + DrawRightMove, 5200 + i) '檔名
   Printer.Line (500 + DrawLeftMove, 8400)-(10500 + DrawRightMove, 8400) '承辦歷程標題欄
   '畫線結束 ***************************
   
   '抬頭
   If IsHaveTaieLogo = True Then
      Printer.PaintPicture pic1.Picture, 2710 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 240, 600, 600
   End If
   Printer.Font.Name = "標楷體"
   Printer.Font.Size = 22
   'Removed by Morgan 2020/3/30
   'Printer.CurrentX = 2800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove '3450
   'Printer.CurrentY = 360
   'Printer.Print "台一國際專利商標事務所"
   'end 2020/3/30
   Printer.CurrentX = 3930 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove '4380
   Printer.CurrentY = 920
   'Modify By 2025/7/10
   'Printer.Print "專利處承辦單"
   Printer.Print "承辦單"
   '2025/7/10 END
   Printer.Font.Size = 12
   'Add By Sindy 2013/9/13
   If m_CP118 <> "" Then
      Printer.CurrentX = 2000
      Printer.CurrentY = 920
      Printer.Print "（電子送件）"
   End If
   '2013/9/13 END
   PrintFontIntoBox "速別", 500 + DrawLeftMove, 1500, 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000
   PrintFontIntoBox "發文", 3300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 1500, 4000 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000
   'Modify By Sindy 2013/9/5
   'PrintFontIntoBox "     年    月    日(      )晉" & IIf(Trim(oStrA14) = "", "    ", oStrA14) & "字第        號", 4000 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 1500, 10500 + i + DrawRightMove, 2000
   PrintFontIntoBox Val(Left(strSrvDate(1), 4)) - 1911 & " 年 " & Mid(strSrvDate(1), 5, 2) & " 月    日 晉" & IIf(Trim(oStrA14) = "", "    ", oStrA14) & "字第                號", 4000 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 1500, 10500 + i + DrawRightMove, 2000
   '2013/9/5 END
   PrintFontIntoBox "受文者", 500 + DrawLeftMove, 2000, 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2900
   PrintFontIntoBox "副本|收受者", 500 + DrawLeftMove, 2900, 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3800
   PrintFontIntoBox "主旨", 500 + DrawLeftMove, 3800, 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800
   PrintFontIntoBox "檔名", 500 + DrawLeftMove, 4800, 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200
   PrintFontIntoBox "本所案號", 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000, 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2450
   PrintFontIntoBox "約定期限", 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2450, 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2900
   PrintFontIntoBox "本所期限", 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2900, 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3350
   PrintFontIntoBox "法定期限", 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3350, 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3800
   PrintFontIntoBox "收文點數", 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200
   PrintFontIntoBox "承辦天數", 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200
   PrintFontIntoBox "備　　　　註", 500 + DrawLeftMove, 5200, 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5600
   PrintFontIntoBox "智權人員", 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5600
   PrintFontIntoBox "監印", 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200, 10500 + i + DrawRightMove, 5600
   PrintFontIntoBox "齊備日", 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5600, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6000
   PrintFontIntoBox "完稿日", 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6000, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6400
   PrintFontIntoBox "會稿日", 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6400, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6800
   PrintFontIntoBox "會回日", 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6800, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7200
   PrintFontIntoBox "打字繕寫", 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7200, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7600
   PrintFontIntoBox "承　　　　辦　　　　歷　　　　程", 500 + DrawLeftMove, 8000, 10500 + i + DrawRightMove, 8400
   '印資料
   Printer.Font.Name = "標楷體"
   Printer.Font.Size = 12
   '受文者
   PrintFontIntoBox txt1(5), 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2900
   '本所案號
   PrintFontIntoBox lblCaseNo, 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2000, 10500 + DrawRightMove, 2450
   '約定期限
   PrintFontIntoBox "    年    月    日", 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2450, 10500 + DrawRightMove, 2900
   '本所期限
   PrintFontIntoBox oStrA06, 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2900, 10500 + DrawRightMove, 3350
   '法定期限
   PrintFontIntoBox oStrA07, 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3350, 10500 + DrawRightMove, 3800
   '副本收受者
   PrintFontIntoBox txt1(6), 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 2900, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3800
   '主旨
   PrintFontIntoBox txt1(0), 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 3800, 10500 + DrawRightMove, 4800, True, False
   '檔名
   PrintFontIntoBox oStrAFile, 1600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800, 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200
   '收文點數
   PrintFontIntoBox m_CP18, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800, 7600 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200
   '承辦天數
   PrintFontIntoBox oStrA10, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 4800, 10500 + DrawRightMove, 5200
   '智權人員
   PrintFontIntoBox oStrA11, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5200, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5600
   '齊備
   PrintFontIntoBox oStrEP06, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 5600, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6000
   '完稿
   PrintFontIntoBox oStrEP09, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6000, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6400
   '會稿
   PrintFontIntoBox oStrEP07, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6400, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6800
   '會回
   PrintFontIntoBox oStrEP08, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 6800, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7200
   '打字繕寫
   PrintFontIntoBox m_EED06, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7200, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7600
   PrintFontIntoBox oStrEED07, 6400 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 7600, 8800 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 8000
   '備註
   'PutSplitSmb(txt1(4).Text, 31, False)
   'PutSplitSmb(txt1(4).Text, 20)
   PrintFontIntoBox PutSplitSmb(txt1(4).Text, 31, False), 550 + DrawLeftMove, 5600, 5300 + ((DrawLeftMove - DrawRightMove) / 2) + DrawLeftMove, 8000, False, False
   
   '加印分所案號 (第一頁的最下面)
   strSql = " Select pa47 From CaseProgress,Patent Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP09='" & m_EEP01 & "' union " & _
            " Select tm34 From CaseProgress,TradeMark Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP09='" & m_EEP01 & "' union " & _
            " Select hc07 From CaseProgress,HireCase Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP09='" & m_EEP01 & "' union " & _
            " Select lc16 From CaseProgress,LawCase Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP09='" & m_EEP01 & "' union " & _
            " Select sp28 From CaseProgress,ServicePractice Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP09='" & m_EEP01 & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If rsA.Fields(0) <> "" Then
         Printer.CurrentX = 1600
         Printer.CurrentY = 16200
         Printer.Print "(分所案號：" & rsA.Fields(0) & ")"
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   
   '承辦歷程
   dblLine = 0
   dblMaxLine = 25
   dblStarLine = 8200
   dblRow = 0
   PrintPage = 1
   'Modify By Sindy 2013/10/16 +eep12
   'Modify By Sindy 2013/10/24 開放顯示聯絡 (" And EEP04 not in('" & EMP_聯絡 & "')")
   strSql = "Select distinct EEP02,decode(eep04,'" & EMP_附加流程 & "',decode(c2.CP43,'',ac03,Decode(" & m_PA09 & ",'000',CPM03,CPM04)),ac03) as 流程狀態,s1.ST02||eep12 as 發送者,sqldatet(EEP06)||' ' ||sqltime(EEP07) as 送出時間,nvl(EEP08,'') as 意見內容,c1.CP43 as CP43" & _
            " From EmpElectronProcess,staff s1,allcode,caseprogress c1,caseprogress c2,casepropertymap" & _
            " Where EEP01='" & m_EEP01 & "'" & _
            " And EEP03=s1.ST01(+) " & _
            " And ac01='09' And EEP04=ac02(+)" & _
            " And eep01=c2.cp43(+) And eep06=c2.cp05(+)" & _
            " And c2.cp01=cpm01(+) And c2.cp10=cpm02(+)" & _
            " And EEP01=c1.cp09(+)" & _
            " order by EEP02 asc"
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'For jj = 1 To 24
      rsA.MoveFirst
      Do While Not rsA.EOF
         dblRow = dblRow + 1
         If dblLine > dblMaxLine Then PrintNewPage
'         dblLine = dblLine + 1
'         Printer.CurrentX = 1600
'         Printer.CurrentY = dblStarLine + (dblLine * 300)
         strEEP04 = Trim(rsA.Fields("流程狀態"))
         '判斷有相關總收文號才做案件性質轉換
         If strEEP04 = "附加流程" Then
            If Trim(rsA.Fields("CP43")) <> "" Then
               strEEP04 = Trim(lblCP10) & PUB_GetRelateCasePropertyName(m_EEP01, "1")
            End If
         End If
'         Printer.Print convForm(CheckStr(strEEP04), 12) & "  " & _
'                       convForm(CheckStr(Trim(rsA.Fields("發送者"))), 10) & "  " & _
'                       Trim(rsA.Fields("送出時間"))
         '***** 折行處理 Star *****
         PrintWidthWord = 78
         '清除陣列值
         For ii = 0 To 100 '40
            PrintDetailTxt(ii) = ""
         Next ii
         w2 = 0
'         PrintDetailTemp = Split(Trim("" & rsA.Fields("意見內容")), vbCrLf)
         PrintDetailTemp = Split(convForm(CheckStr(strEEP04), 10) & "  " & _
                                 convForm(CheckStr(Trim(rsA.Fields("發送者"))), 12) & "  " & _
                                 Trim(rsA.Fields("送出時間")) & "  " & Trim(rsA.Fields("意見內容")), vbCrLf)
         For oJ = 0 To UBound(PrintDetailTemp)
            w2 = w2 + 1
            PrintDetailTxt(w2) = PrintDetailTemp(oJ)
            If PUB_StrToStr_byVal(PrintDetailTemp(oJ), PrintWidthWord + 2) <> PrintDetailTemp(oJ) Then
               PrintDetailTxt(w2) = "": w2 = w2 - 1
               bolOK = True
               Do While bolOK = True
                  w2 = w2 + 1
                  PrintDetailTxt(w2) = PrintDetailTxt(w2) & RTrim(PUB_StrToStr_byVal(PrintDetailTemp(oJ), PrintWidthWord) & Chr(13)) & Chr(10)
                  PrintDetailTemp(oJ) = Replace(PrintDetailTemp(oJ), PUB_StrToStr_byVal(PrintDetailTemp(oJ), PrintWidthWord), "")
                  If PUB_StrToStr_byVal(PrintDetailTemp(oJ), PrintWidthWord) = PrintDetailTemp(oJ) Then
                     w2 = w2 + 1
                     PrintDetailTxt(w2) = PrintDetailTxt(w2) & PrintDetailTemp(oJ)
                     bolOK = False
                  End If
              Loop
            Else
               PrintDetailTxt(w2) = RTrim(PrintDetailTemp(oJ))
            End If
         Next oJ
         For ii = 1 To w2
            dblLine = dblLine + 1
            If dblLine > dblMaxLine Then PrintNewPage
            Printer.CurrentX = 1600
            Printer.CurrentY = dblStarLine + (dblLine * 300)
            Printer.Print PrintDetailTxt(ii)
         Next ii
         '***** 折行處理 End *****
         If dblRow < rsA.RecordCount Then
         'If dblRow > 24 Then
            dblLine = dblLine + 1
            If dblLine > dblMaxLine Then PrintNewPage
            Printer.CurrentX = 1600
            Printer.CurrentY = dblStarLine + (dblLine * 300)
            Printer.Print String(78, "-")
         End If
         rsA.MoveNext
      Loop
      'Next jj
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   
   Set rsA = Nothing
   Printer.EndDoc
   'ShowPrintOk
End Sub

Private Sub PrintNewPage()
   Printer.NewPage
   PrintPage = PrintPage + 1
   DrawLeftMove = 1000
   DrawRightMove = 650
   '粗線深度
   DrawCount = 20
   IsHaveTaieLogo = False
   If Dir(Trim(App.path) & "\taie_logo.jpg") <> "" Then
      IsHaveTaieLogo = True
      pic1.Picture = LoadPicture(App.path & "\taie_logo.jpg")
   ElseIf Dir("c:\pics\taie_logo.jpg") <> "" Then
      IsHaveTaieLogo = True
      pic1.Picture = LoadPicture("c:\pics\taie_logo.jpg")
   End If
   
   oStrA14 = PUB_GetST07(m_CP14)
   
   '畫線 *****************************
   '裝訂線
   Printer.DrawStyle = 2
   Printer.Line (500, 0)-(500, 6300)
   Printer.Line (500, 6300 + Printer.TextHeight("裝"))-(500, 7400)
   Printer.Line (500, 7400 + Printer.TextHeight("訂"))-(500, 8600)
   Printer.Line (500, 8600 + Printer.TextHeight("線"))-(500, 17000)
   '打字
   Printer.CurrentX = 500 - (Printer.TextWidth("裝") / 2)
   Printer.CurrentY = 6300
   Printer.Print "裝"
   Printer.CurrentX = 500 - (Printer.TextWidth("訂") / 2)
   Printer.CurrentY = 7400
   Printer.Print "訂"
   Printer.CurrentX = 500 - (Printer.TextWidth("線") / 2)
   Printer.CurrentY = 8600
   Printer.Print "線"
   Printer.DrawStyle = 0
   '粗線
   For i = 1 To DrawCount
      '方格
      Printer.Line (500 + i + DrawLeftMove, 2000 + i)-(10500 + i + DrawRightMove, 16100 + i), , B
   Next i
   
   dblLine = 0
   dblMaxLine = 46
   dblStarLine = 2000
End Sub

'放字進去框框中
'oStr 要放的字    要分行就要加 |
'oLeft, oUp  框框左上角
'oRight, 0Down 框框右下角
'IsCenter 是否置中
'IsAvgFont 是否自平均分配  與 IsCenterW = true  同時使用 無效
'Sub PrintFontIntoBox(ByVal oStr As String, oLeft As Integer, oUp As Integer, oRight As Integer, oDown As Integer, Optional IsCenterH As Boolean = True, Optional IsCenterW As Boolean = True)
Sub PrintFontIntoBox(ByVal oStr As String, oLeft As Integer, oUp As Integer, oRight As Integer, oDown As Integer, Optional IsCenterH As Boolean = True, Optional IsCenterW As Boolean = True, Optional IsAvgFont As Boolean = False)
Dim BoxHeight As Integer
Dim BoxWidth As Integer
Dim FontHeight As Integer
Dim FontWidth As Double
Dim ArrStr As Variant
Dim oIntI As Integer
Dim oIntJ As Integer
Dim FontTop As Integer
Dim FontLeft As Integer
Dim FontAllHeight As Integer
Dim SingleFontWidth As Integer
Dim CalFontWidth As Integer
Dim SingleFont As String
Dim TmpFont As String         '暫存的單字
Dim TmpAllFont As String         '暫存的整格字
Dim TmpLineFont As String
'add by nickc 2007/03/01
Dim TmpPrtWd As Integer

   'add by nickc 2005/09/09
   oStr = Replace(oStr, vbCrLf, "|")
   '先去跳行符號
   oStr = Replace(Replace(oStr, Chr(13), ""), Chr(10), "")
   BoxHeight = oDown - oUp
   BoxWidth = oRight - oLeft
   FontHeight = Printer.TextHeight(Mid(oStr, 1, 1))
   FontWidth = Printer.TextWidth(Mid(oStr, 1, 1))
   ArrStr = Split(Replace(oStr, vbCrLf, ""), "|")
   '檢查若是超過長度，自動跳行
   For oIntI = 0 To UBound(ArrStr)
      '超過
      TmpFont = ArrStr(oIntI)
       If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
           Printer.Font.Size = 9
       Else
           Printer.Font.Size = 12
       End If
       If TmpFont <> "" Then
            FontHeight = Printer.TextHeight(Mid(Trim(ArrStr(oIntI)), 1, 1))
            FontWidth = Printer.TextWidth(ArrStr(oIntI))
           If Len(TmpFont) > (BoxWidth / FontWidth) Then
              TmpAllFont = ""
              SingleFont = ""
              CalFontWidth = 0
              SingleFontWidth = 0
              TmpFont = ""
              TmpLineFont = ""
              TmpFont = ArrStr(oIntI)
              Do While Not Len(TmpFont) = 0
                 SingleFont = GetOneFont(TmpFont)
                 SingleFontWidth = Printer.TextWidth(SingleFont)
                 If Printer.TextWidth(TmpLineFont) + SingleFontWidth > BoxWidth Then
                    TmpAllFont = TmpAllFont & TmpLineFont & "|"
                    TmpLineFont = ""
                    CalFontWidth = 0
                 End If
                 TmpLineFont = TmpLineFont & SingleFont
              Loop
              If TmpLineFont <> "" Then
                 TmpAllFont = TmpAllFont & TmpLineFont
              End If
              If TmpAllFont <> "" Then
                 ArrStr(oIntI) = TmpAllFont
              End If
           End If
       End If
   Next oIntI
   oStr = Join(ArrStr, "|")
   ArrStr = Split(oStr, "|")
   If IsCenterH = True Then
      FontAllHeight = (Val(UBound(ArrStr)) + 1) * FontHeight
      If FontAllHeight > BoxHeight Then FontAllHeight = BoxHeight
      FontTop = ((BoxHeight - FontAllHeight) / 2) + oUp
   Else
      FontTop = oUp
   End If
   For oIntI = 0 To UBound(ArrStr)
      If (FontTop + (FontHeight * (oIntI + 1))) < oDown Then
         'FontTop = FontTop + (FontHeight * oIntI)
         If IsCenterW = True Then
            FontWidth = Printer.TextWidth(ArrStr(oIntI))
            FontLeft = ((BoxWidth - FontWidth) / 2) + oLeft
            'add by nickc 2007/03/08 遇到第一個是 □ 改縮小
            If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
               Printer.Font.Size = 9
   '            FontHeight = Printer.TextHeight(Mid(oStr, 1, 1))
   '            FontWidth = Printer.TextWidth(ArrStr(oIntI))
   '            FontLeft = ((BoxWidth - FontWidth) / 2) + oLeft
            End If
            Printer.CurrentX = FontLeft
            Printer.CurrentY = FontTop + (FontHeight * oIntI)
            Printer.Print ArrStr(oIntI)
            'add by nickc 2007/03/08 遇到第一個是 □ 改縮小
            If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
               Printer.Font.Size = 12
            End If
         ElseIf IsAvgFont = True Then
            'oLeft = BoxWidth \ Len(ArrStr(0))
            
            For oIntJ = 1 To Len(ArrStr(oIntI))
               TmpFont = Mid(ArrStr(oIntI), oIntJ, 1)
               FontWidth = Printer.TextWidth(TmpFont)
               TmpPrtWd = ((BoxWidth - FontWidth) \ (Len(ArrStr(oIntI)) - 1))
               FontLeft = (((BoxWidth - FontWidth) \ (Len(ArrStr(oIntI)) - 1)) * (oIntJ - 1)) + oLeft
               Printer.CurrentX = FontLeft
               Printer.CurrentY = FontTop + (FontHeight * oIntI)
               Printer.Print TmpFont
            Next oIntJ
         Else
            FontLeft = oLeft
            'add by nickc 2007/03/08 遇到第一個是 □ 改縮小
            If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
               Printer.Font.Size = 9
   '            FontHeight = Printer.TextHeight(Mid(oStr, 1, 1))
   '            FontWidth = Printer.TextWidth(ArrStr(oIntI))
   '            FontLeft = ((BoxWidth - FontWidth) / 2) + oLeft
            End If
            Printer.CurrentX = FontLeft
            Printer.CurrentY = FontTop + (FontHeight * oIntI)
            Printer.Print ArrStr(oIntI)
            'add by nickc 2007/03/08 遇到第一個是 □ 改縮小
            If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
               Printer.Font.Size = 12
            End If
         End If
   
      End If
   Next oIntI
End Sub

Function GetOneFont(ByRef oStr As String) As String
Dim i As Integer
   
   GetOneFont = ""
   If Asc(Mid(oStr, 1, 1)) < 0 Or Asc(Mid(oStr, 1, 1)) > 256 Then
      '雙位元組
      GetOneFont = Mid(oStr, 1, 1)
      oStr = Mid(oStr, 2)
      Exit Function
   Else
      Select Case Mid(oStr, 1, 1)
      '符號，或特殊字
      Case ",", " ", ":", ";", "!"
         GetOneFont = Mid(oStr, 1, 1)
         oStr = Mid(oStr, 2)
      '單位元組
      Case Else
         For i = 1 To Len(oStr)
            If Asc(Mid(oStr, i, 1)) < 0 Or Asc(Mid(oStr, i, 1)) > 256 Then
               Exit For
            Else
               Select Case Mid(oStr, i, 1)
               '符號，或特殊字
               Case ",", " ", ":", ";", "!"
                  Exit For
               Case Else
                  If Asc(Mid(oStr, i, 1)) = 13 Or Asc(Mid(oStr, i, 1)) = 10 Then
                     oStr = Mid(oStr, 2)
                     Exit For
                  Else
                     GetOneFont = GetOneFont & Mid(oStr, i, 1)
                  End If
               End Select
            End If
         Next i
         oStr = Mid(oStr, Len(GetOneFont) + 1)
      End Select
   End If
End Function

'add by nickc 2005/09/02 加入分行的符號
'edit by nickc 2007/03/03
'Function PutSplitSmb(oStr As String, oLen As Integer) As String
Function PutSplitSmb(oStr As String, oLen As Integer, Optional IsDelSpace As Boolean = True) As String
Dim oAllLen As Integer
Dim oStr2 As String
Dim oI As Integer
Dim oJ As Integer
Dim oTmpArr As Variant
Dim oTmp1Arr As Variant

   PutSplitSmb = ""
   oTmpArr = Split(oStr, vbCrLf)
   For oJ = 0 To UBound(oTmpArr)
         If IsDelSpace = True Then
             oTmpArr(oJ) = Replace(oTmpArr(oJ), " ", "")
         End If
         If oTmpArr(oJ) <> "" Then
'Modified by Morgan 2013/4/16 加考慮英文斷行
'            oAllLen = LenB(StrConv(oTmpArr(oJ), vbFromUnicode))
'            For oI = 1 To ((oAllLen \ (oLen * 2)) + IIf(oAllLen Mod (oLen * 2) <> 0, 1, 0))
'               oStr2 = Replace(StrConv(MidB(StrConv(oTmpArr(oJ), vbFromUnicode), 1, oLen * 2), vbUnicode), Chr(0), "")
'               PutSplitSmb = PutSplitSmb & oStr2 & IIf(oI = ((oAllLen \ (oLen * 2)) + IIf(oAllLen Mod (oLen * 2) <> 0, 1, 0)), "", "|")
'               oTmpArr(oJ) = StrConv(MidB(StrConv(oTmpArr(oJ), vbFromUnicode), LenB(StrConv(oStr2, vbFromUnicode)) + 1), vbUnicode)
'            Next oI
            oTmp1Arr = Split(oTmpArr(oJ), " ")
            oStr2 = ""
            For intI = 0 To UBound(oTmp1Arr)
               If oTmp1Arr(intI) <> "" Then
                  If LenB(StrConv(oStr2 & oTmp1Arr(intI), vbFromUnicode)) <= oLen * 2 Then
                     If oStr2 <> "" Then oStr2 = oStr2 & " "
                     oStr2 = oStr2 & oTmp1Arr(intI)
                  Else
                     If oStr2 <> "" Then 'Modify By Sindy 2013/5/15 +if
                        PutSplitSmb = PutSplitSmb & oStr2 & "|"
                     End If
                     oStr2 = oTmp1Arr(intI)
                  End If
               End If
            Next
            PutSplitSmb = PutSplitSmb & oStr2
'end 2013/4/16
            If oJ <> UBound(oTmpArr) Then PutSplitSmb = PutSplitSmb & "|"
         End If
   Next oJ
End Function

'Add By Sindy 2018/11/19
'產生申請書及上傳
Private Sub BegetAppFUpload()
Dim strFileName As String, stFullName As String
Dim hLocalFile As Long
Dim strSubCaseNo As String
Dim strPOAFileName As String, strPOAFileFullName As String, bolPA165 As Boolean
Dim fs As Object
Dim strChkPA165Err As String 'Add By Sindy 2025/9/11
   
On Error GoTo ErrHand
   
   '檢查/產生資料夾
   If Dir(m_strFolder, vbDirectory) = "" Then
      MkDir m_strFolder
   End If
   
   '上傳檔案
   Screen.MousePointer = vbHourglass
   For ii = 0 To lstAtt(0).ListCount - 1
      strFileName = lstAtt(0).List(ii)
      If InStrRev(strFileName, " (") > 0 Then
         strFileName = Left(strFileName, InStrRev(strFileName, " (") - 1)
      End If
      stFullName = m_strFolder & strFileName
      If stFullName <> "" Then
         If Dir(stFullName) <> "" Then
            If MsgBox("檔案[ " & strFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
               stFullName = ""
            End If
         End If
         If stFullName <> "" Then
            If GetAttachFile(strFileName, CInt(m_AttEEP02), stFullName) = False Then
               MsgBox "無法儲存檔案[ " & strFileName & " ]！"
               GoTo ErrHand
            End If
         End If
      End If
   Next ii
      
   '產生申請書
   If (cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "125") And m_PA09 = 台灣國家代號 Then
      Screen.MousePointer = vbHourglass
      'POA檔案:
      strPOAFileName = m_strCaseNo & ".POA.pdf"
      strPOAFileFullName = m_strFolder & strPOAFileName
      'Modify By Sindy 2019/5/22 玲玲:請調整"待轉檔區"產生送件資料夾時,POA檔案的帶入順序,1.先以POA資料夾為主，無資料時若系統有設定總委任書再抓總委任書
      '個案,POA存取...
      If Dir(strPOAFileFullName) = "" Then
         If Dir(m_strPOAFolder & strPOAFileName) <> "" Then
            Set fs = CreateObject("Scripting.FileSystemObject")
            fs.CopyFile m_strPOAFolder & strPOAFileName, m_strFolder
            Set fs = Nothing
            bolPA165 = True
         End If
      End If
      If bolPA165 = False And Dir(strPOAFileFullName) = "" Then
         strChkPA165Err = "" 'Add By Sindy 2025/9/11
         '檢查是否有總委任書
         If pa(26) <> "" Then
            'Modify By Sindy 2025/9/11 +pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
            '                          ,strChkPA165Err: 回傳錯誤訊息
            If PUB_ChkPA165IsY(pa(26), strSubCaseNo, , pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4), strChkPA165Err) = True Then
               bolPA165 = True
               If GetCPPFileAndDownload("POA", strSubCaseNo, strPOAFileFullName) = False Then
                  GoTo ErrHand
               End If
            End If
         End If
         If pa(27) <> "" Then
            'Modify By Sindy 2025/9/11 +pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
            '                          ,strChkPA165Err: 回傳錯誤訊息
            If PUB_ChkPA165IsY(pa(27), strSubCaseNo, , pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4), strChkPA165Err) = True Then
               bolPA165 = True
               If GetCPPFileAndDownload("POA", strSubCaseNo, strPOAFileFullName) = False Then
                  GoTo ErrHand
               End If
            End If
         End If
         If pa(28) <> "" Then
            'Modify By Sindy 2025/9/11 +pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
            '                          ,strChkPA165Err: 回傳錯誤訊息
            If PUB_ChkPA165IsY(pa(28), strSubCaseNo, , pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4), strChkPA165Err) = True Then
               bolPA165 = True
               If GetCPPFileAndDownload("POA", strSubCaseNo, strPOAFileFullName) = False Then
                  GoTo ErrHand
               End If
            End If
         End If
         If pa(29) <> "" Then
            'Modify By Sindy 2025/9/11 +pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
            '                          ,strChkPA165Err: 回傳錯誤訊息
            If PUB_ChkPA165IsY(pa(29), strSubCaseNo, , pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4), strChkPA165Err) = True Then
               bolPA165 = True
               If GetCPPFileAndDownload("POA", strSubCaseNo, strPOAFileFullName) = False Then
                  GoTo ErrHand
               End If
            End If
         End If
         If pa(30) <> "" Then
            'Modify By Sindy 2025/9/11 +pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
            '                          ,strChkPA165Err: 回傳錯誤訊息
            If PUB_ChkPA165IsY(pa(30), strSubCaseNo, , pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4), strChkPA165Err) = True Then
               bolPA165 = True
               If GetCPPFileAndDownload("POA", strSubCaseNo, strPOAFileFullName) = False Then
                  GoTo ErrHand
               End If
            End If
         End If
         'Add By Sindy 2025/9/11
         If strChkPA165Err <> "" Then
            MsgBox strChkPA165Err, vbExclamation
         End If
         '2025/9/11 END
      End If
      If bolPA165 = True Then bolHadPOAfile = True '有委任書
      
      'Add By Sindy 2019/2/22
      'POA資料夾中若有相同案號的PRI檔案(權先權證明文件)一併帶入資料夾
      If Dir(m_strPOAFolder & m_strCaseNo & ".PRI.pdf") <> "" Then
         Set fs = CreateObject("Scripting.FileSystemObject")
         fs.CopyFile m_strPOAFolder & m_strCaseNo & ".PRI.pdf", m_strFolder
         Set fs = Nothing
      End If
      '2019/2/22 END
      
      'Modify By Sindy 2019/5/16 玲玲說要改回工程師產生申請書
'      '2.申請書
'      'StartLetter2 "01", "03"
'      Pub_P_NewCaseStartLetter2 "01", "03", lblCP09, pa, cp, IIf(lblCM10.Visible = True, True, False), bolHadPOAfile, m_bolShowEng
'      NowPrint lblCP09, "01", "03", False, strUserNum, , , True, strExc(9)
'      strFileName = m_strFolder & m_strCaseNo & ".data" 'm_CPM26
'      If Dir(strFileName) <> "" Then
'         strFileName = m_strFolder & m_strCaseNo & ".data_" & Trim(lblPA08.Caption) & "專利申請書"
'      End If
''      Call PUB_MakeDoc(strExc(9), strFileName)
'
'      '1.基本資料
'      'Modify By Sindy 2018/12/5 + m_bolShowEng
'      StartLetterPA_EData "01", "14", lblCP09, pa, cp, True, True, , , m_bolShowEng
'      NowPrint lblCP09, "01", "14", False, strUserNum, , , True, strExc(10)
''      strFileName = m_strFolder & m_strCaseNo & ".contact"
'      'Chr(12):跳頁
'      Call PUB_MakeDoc(strExc(9) & Chr(12) & strExc(10), strFileName, False)
      
'      Screen.MousePointer = vbDefault
'      MsgBox "資料已產生完畢!!!"
'   Else
'      MsgBox "下載完成！"
   End If
   MsgBox "下載完成！"
   
   '開啟資料夾
   ShellExecute hLocalFile, "explore", m_strFolder, vbNullString, vbNullString, 1
   
   Screen.MousePointer = vbDefault
   
   '重新顯示附件區
   Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
   
   ChDir App.path 'Add By Sindy 2020/3/9 釋放資料夾權限
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Sub

'Add By Sindy 2018/11/20
'strFileKind:副檔名
'strCaseNo:本所案號(XXX-XXXXXX-X-XX)
Private Function GetCPPFileAndDownload(strFileKind As String, strCaseNo As String, pSavePath As String) As Boolean
Dim rsTmp As ADODB.Recordset
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
   GetCPPFileAndDownload = True
   strCP01 = SystemNumber(strCaseNo, 1)
   strCP02 = SystemNumber(strCaseNo, 2)
   strCP03 = SystemNumber(strCaseNo, 3)
   strCP04 = SystemNumber(strCaseNo, 4)
   strExc(0) = "select cpp01,cpp02 from casepaperpdf,caseprogress" & _
               " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
               " and cp09=cpp01(+)" & _
               " and upper(substr(cpp02,-8))=upper('." & strFileKind & ".PDF')" & _
               " order by CPP06 desc,CPP07 desc"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsTmp.MoveFirst
      If PUB_GetAttachFile_CPP(rsTmp.Fields("cpp01"), rsTmp.Fields("cpp02"), pSavePath, True) = False Then
         MsgBox pSavePath & "檔案下載有誤！"
         GetCPPFileAndDownload = False
      End If
   End If
   Set rsTmp = Nothing
End Function

'Add By Sindy 2018/11/21
'轉檔完成
Private Sub cmdTrans_Click()
Dim fs As Object
Dim f
Dim dblFCnt As Double
Dim bolFindFile As Boolean
Dim stFileName As String
Dim strPOAFileName As String 'Add By Sindy 2020/3/18
   
   File1.path = m_strFolder
   File1.Refresh
   If File1.ListCount = 0 Then
      MsgBox m_strFolder & " 資料夾無資料！"
      Exit Sub
   End If
   
   'Add By Sindy 2020/7/20 於待轉檔區處理(422)加速審查時，
   '必須檢查是否已收到C類來函(1204)通知實審日，
   '若尚未收到，出現訊息通知USER
   'Modified by Morgan 2024/11/18 +477再審查加速審查並改用專用模組判斷
   'If m_CP10 = "422" Then
   '   If PUB_ChkCPExist(cp, "1204") = False Then
   If m_CP10 = "422" Or m_CP10 = "447" Then
      If PUB_Chk1204(cp) = False Then
   'end 2024/11/18
         MsgBox "「本案尚未收到通知實查日來函，故暫不能發文」，並不能執行「轉檔完成」。"
         Exit Sub
      End If
   End If
   '2020/7/20 END
   
   'Added by Morgan 2021/11/2
   '指定送件日檢查
   'Modify By Sindy 2023/4/21 and cp164='1' ==> and nvl(cp164,'1')='1'
   strExc(0) = "select cp142 From caseprogress where cp09='" & lblCP09 & "' and cp141='3' and nvl(cp164,'1')='1'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp("cp142") > strSrvDate(1) Then
         If MsgBox("本案指定送件日為" & ChangeWStringToTDateString(RsTemp("cp142")) & "，請確認是否送件？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
      End If
   End If
   'end  2021/11/2
   
   '檢查檔案
   If Dir(m_strFolder & m_strCaseNo & "*.data*.pdf") = "" Then  '未做E-Set轉檔
      MsgBox "(" & m_strFolder & m_strCaseNo & ") 無 Data.pdf 檔案！"
      Exit Sub
   End If
   If Dir(m_strFolder & m_strCaseNo & "*.data.doc") = "" And _
      Dir(m_strFolder & m_strCaseNo & "*.data.docx") = "" Then
      MsgBox "(" & m_strFolder & m_strCaseNo & ") 無申請書（Data.doc）檔案！"
      Exit Sub
   End If
   '檢查檔名規則
   For dblFCnt = 0 To File1.ListCount - 1
      If UCase(Right(File1.List(dblFCnt), 4)) <> ".PDF" Then
         If PUB_ChkEmpFlowFNMRule(lblCaseNo, File1.List(dblFCnt), EMP_判發, m_CP10, , 0) = False Then
            Exit Sub
         End If
      End If
   Next dblFCnt
   
   'Add By Sindy 2019/5/28 資料列全部先變成未選取
   For dblFCnt = lstAtt(0).ListCount - 1 To 0 Step -1
      lstAtt(0).Selected(dblFCnt) = False
   Next dblFCnt
   '2019/5/28 END
   For dblFCnt = lstAtt(0).ListCount - 1 To 0 Step -1
      stFileName = Trim(GetFileName(lstAtt(0).List(dblFCnt)))
      If UCase(Right(stFileName, 9)) = ".DATA.DOC" Or _
         UCase(Right(stFileName, 10)) = ".DATA.DOCX" Then
         '重覆或多的要刪除
         'If Dir(m_strFolder & stFileName) <> "" Then
            lstAtt(0).Selected(dblFCnt) = True
            RemoveList lstAtt(0), 0, False
         'End If
      End If
   Next dblFCnt
   '重新顯示附件區
   Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), False)
   Set fs = CreateObject("Scripting.FileSystemObject")
   '檢查資料夾中電子檔是否有未存在附件區,若有,要上傳至附件區歸檔用
   'Modify By Sindy 2019/6/28 取消抓回 .PRI.PDF 或 .POA.PDF
   For dblFCnt = 0 To File1.ListCount - 1
'      If UCase(Right(Trim(File1.List(dblFCnt)), 9)) = ".DATA.DOC" Or _
'         UCase(Right(Trim(File1.List(dblFCnt)), 10)) = ".DATA.DOCX" Or _
'         UCase(Right(Trim(File1.List(dblFCnt)), 8)) = ".PRI.PDF" Or _
'         UCase(Right(Trim(File1.List(dblFCnt)), 8)) = ".POA.PDF" Then
      If UCase(Right(Trim(File1.List(dblFCnt)), 9)) = ".DATA.DOC" Or _
         UCase(Right(Trim(File1.List(dblFCnt)), 10)) = ".DATA.DOCX" Then
         stFileName = File1.List(dblFCnt)
         bolFindFile = False
         For ii = 0 To lstAtt(0).ListCount - 1
            If InStr(UCase(lstAtt(0).List(ii)), UCase(stFileName)) > 0 Then
               bolFindFile = True
               Exit For
            End If
         Next ii
         If bolFindFile = False Then
            stFileName = m_strFolder & File1.List(dblFCnt)
            Set f = fs.GetFile(stFileName)
            'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg stFileName & MsgText(9221)
               Exit Sub
            End If
            '2013/9/6 END
            If AddListX(lstAtt(0), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(0)) = True Then
               '存檔
               If SaveAttFile(CInt(m_AttEEP02), 0) = False Then
                  Exit Sub
               End If
            End If
         End If
      End If
   Next dblFCnt
   '重新顯示附件區
   Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), False)
   
   '更新電子送件轉檔完成日期
   strSql = "update caseprogress set cp160=" & strSrvDate(1) & " where cp09='" & m_EEP01 & "'"
   cnnConnection.Execute strSql
   
   'Add By Sindy 2020/3/18 玲玲說要刪除server上的個案POA檔案
   strPOAFileName = m_strCaseNo & ".POA.pdf" 'POA檔名
   If Dir(m_strPOAFolder & strPOAFileName) <> "" Then '個案,POA存取...
      fs.DeleteFile m_strPOAFolder & strPOAFileName '刪除
   End If
   '2020/3/18 END
   
   ChDir App.path 'Add By Sindy 2020/3/9 釋放資料夾權限
   
   Set fs = Nothing
   Set f = Nothing
   
   cmdTrans.BackColor = &H80C0FF '已轉檔完成
   MsgBox "轉檔上傳至附件區，已完成！", vbInformation
   'cmdExit_Click '離開結束
End Sub

'Add By Sindy 2025/10/20
'FCP:將桌面上電子檔匯入附件區
Private Function FCPSaveFileToListBox(Index As Integer) As Boolean
Dim fs As Object
Dim f
Dim dblFCnt As Double
'Dim bolFindFile As Boolean
Dim stFiName As String, stReName As String
   
   If m_strFolder = "" Then FCPSaveFileToListBox = True: Exit Function
   File1.path = m_strFolder
   File1.Refresh
   If File1.ListCount = 0 Then
      FCPSaveFileToListBox = True
      Exit Function
   End If
   
   '非PDF檔 (*.DOC、*.DOCX、*.TXT、*.XML(序列表))的電子檔就是要歸原始檔區的
   '另外，.FIG.PDF (最終版圖示)雖然是PDF檔但也要歸原始檔區。
   '      .RES.PDF (相似結果)
   '      .SEP.PDF (參考本)
   '排除不歸檔的檔名 *.zip 、*申請書*.（新申請案要留，更名為.data.）、.contact.
   For dblFCnt = 0 To File1.ListCount - 1
      If InStr(UCase(File1.List(dblFCnt)), UCase(".contact.")) = 0 Then
         If UCase(Right(File1.List(dblFCnt), 8)) = UCase(".FIG.PDF") Or _
            UCase(Right(File1.List(dblFCnt), 8)) = UCase(".RES.PDF") Or _
            UCase(Right(File1.List(dblFCnt), 8)) = UCase(".SEP.PDF") Or _
            UCase(Right(File1.List(dblFCnt), 4)) = UCase(".DOC") Or _
            UCase(Right(File1.List(dblFCnt), 5)) = UCase(".DOCX") Or _
            UCase(Right(File1.List(dblFCnt), 4)) = UCase(".TXT") Or _
            UCase(Right(File1.List(dblFCnt), 4)) = UCase(".XML") Then
            
            stFiName = m_strFolder & File1.List(dblFCnt)
            stReName = ""
            '新申請案才要留申請書，更名為.data.
            If InStr(NewCasePtyList, cp(10)) > 0 And _
               (InStr(UCase(File1.List(dblFCnt)), "申請書") > 0 Or InStr(UCase(File1.List(dblFCnt)), UCase(".data.")) > 0) Then
               stReName = PUB_CaseNo2FileName(m_CP01, m_CP02, m_CP03, m_CP04) & ".data." & Right(File1.List(dblFCnt), Len(File1.List(dblFCnt)) - InStrRev(File1.List(dblFCnt), "."))
            '其他申請書不歸
            ElseIf InStr(UCase(File1.List(dblFCnt)), "申請書") = 0 Then
               stReName = PUB_GetSimpleName(File1.List(dblFCnt)) '去掉中文
               '更名
               Call PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, stReName, stReName)
               If InStr(stReName, "." & m_CP10 & ".") > 0 Then stReName = Replace(stReName, "." & m_CP10 & ".", ".")
            End If
            If stReName <> "" Then
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFiName)
               '檔案大小為 0 KB 有誤
               If f.Size > 0 Then
                  If SaveAttFile_EEF(m_EEP01, CInt(m_AttEEP02), stFiName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS")) = False Then
                     GoTo RunExit
                  Else
                     '更新電子送件轉檔完成日期
                     strSql = "update caseprogress set cp160=" & strSrvDate(1) & " where cp09='" & m_EEP01 & "'"
                     cnnConnection.Execute strSql
                     
                     Call PUB_DelPCOrgFile(stFiName) '一併將PC上的實體檔案刪除
                  End If
               End If
            End If
         End If
      End If
   Next dblFCnt
   FCPSaveFileToListBox = True
   
RunExit:
   '重新顯示附件區
   Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), False)
   
   ChDir App.path '再改變目錄,可以釋放資料夾權限
   
   Set fs = Nothing
   Set f = Nothing
End Function
