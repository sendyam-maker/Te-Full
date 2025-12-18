VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_F_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦歷程資料查詢"
   ClientHeight    =   5772
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   8952
   Tag             =   "加班資料"
   Begin VB.TextBox txtNote 
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   3330
      TabIndex        =   37
      Text            =   "※此案屬多案歷程，請參"
      Top             =   540
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   2730
      TabIndex        =   36
      Text            =   "(共X筆)"
      Top             =   30
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdCPPAtt 
      BackColor       =   &H00C0FFC0&
      Caption         =   "多案卷宗區附件"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   6330
      MaskColor       =   &H8000000F&
      Style           =   1  '圖片外觀
      TabIndex        =   35
      Top             =   30
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1860
      TabIndex        =   28
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   5790
      Width           =   4395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7950
      TabIndex        =   0
      Top             =   30
      Width           =   765
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5050
      Left            =   30
      TabIndex        =   1
      Top             =   690
      Width           =   8930
      _ExtentX        =   15748
      _ExtentY        =   8911
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "簽辦流程"
      TabPicture(0)   =   "frm100101_F_2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(1)=   "Label3(0)"
      Tab(0).Control(2)=   "Label10(0)"
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(7)=   "lblEApp"
      Tab(0).Control(8)=   "txtEEP03_2"
      Tab(0).Control(9)=   "CboEEP05"
      Tab(0).Control(10)=   "txtEEP10_2"
      Tab(0).Control(11)=   "txtEEP08"
      Tab(0).Control(12)=   "Winsock1"
      Tab(0).Control(13)=   "CommonDialog1"
      Tab(0).Control(14)=   "GRD1"
      Tab(0).Control(15)=   "CboEEP04"
      Tab(0).Control(16)=   "txtEEP03"
      Tab(0).Control(17)=   "Frame1"
      Tab(0).Control(18)=   "txtEEP10"
      Tab(0).Control(19)=   "txtEEP02"
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "承辦單內容"
      TabPicture(1)   =   "frm100101_F_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "Label1(2)"
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(4)=   "txt1(5)"
      Tab(1).Control(5)=   "txt1(6)"
      Tab(1).Control(6)=   "txt1(0)"
      Tab(1).Control(7)=   "txt1(4)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "歷程備註"
      TabPicture(2)   =   "frm100101_F_2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtEP12"
      Tab(2).Control(1)=   "txt3(7)"
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(3)=   "Label20"
      Tab(2).Control(4)=   "Frame201"
      Tab(2).Control(5)=   "Frame945"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "存卷資料"
      TabPicture(3)   =   "frm100101_F_2.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "LblinfoNote"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame3"
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
         Left            =   -74850
         TabIndex        =   76
         Top             =   390
         Visible         =   0   'False
         Width           =   5350
         Begin MSForms.TextBox txtEED14 
            Height          =   300
            Left            =   2820
            TabIndex        =   80
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
            TabIndex        =   79
            Top             =   360
            Width           =   2700
         End
         Begin MSForms.TextBox txtEED15 
            Height          =   290
            Left            =   2820
            TabIndex        =   78
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
            TabIndex        =   77
            Top             =   690
            Width           =   2700
         End
      End
      Begin VB.Frame Frame201 
         Height          =   2320
         Left            =   -74940
         TabIndex        =   58
         Top             =   390
         Width           =   8800
         Begin VB.ComboBox CmbFL 
            Height          =   260
            Index           =   3
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   65
            Text            =   "CmbFL"
            Top             =   300
            Width           =   6060
         End
         Begin VB.CheckBox ChkEED13 
            Caption         =   "轉檔後送件（程序發文）"
            ForeColor       =   &H000000C0&
            Height          =   220
            Left            =   5640
            TabIndex        =   64
            Top             =   660
            Width           =   2740
         End
         Begin VB.Frame Frame7 
            Height          =   280
            Left            =   150
            TabIndex        =   59
            Top             =   0
            Width           =   3970
            Begin VB.Label LblEED10_N_2 
               AutoSize        =   -1  'True
               Caption         =   "LblEED10_N_2"
               Height          =   180
               Left            =   2550
               TabIndex        =   63
               Top             =   60
               Width           =   1070
            End
            Begin MSForms.TextBox txt3 
               Height          =   320
               Index           =   3
               Left            =   750
               TabIndex        =   62
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
            Begin VB.Label LblEED10 
               AutoSize        =   -1  'True
               Caption         =   "譯者："
               Height          =   180
               Left            =   210
               TabIndex        =   61
               Top             =   30
               Width           =   540
            End
            Begin VB.Label LblEED10_N 
               AutoSize        =   -1  'True
               Caption         =   "LblEED10_N"
               Height          =   180
               Left            =   1470
               TabIndex        =   60
               Top             =   60
               Width           =   910
            End
         End
         Begin VB.Label LblEED09_N 
            AutoSize        =   -1  'True
            Caption         =   "LblEED09_N"
            Height          =   180
            Left            =   4260
            TabIndex        =   75
            Top             =   690
            Width           =   910
         End
         Begin MSForms.TextBox txt3 
            Height          =   1320
            Index           =   4
            Left            =   900
            TabIndex        =   74
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "檔案名稱："
            Height          =   180
            Index           =   8
            Left            =   0
            TabIndex        =   73
            Top             =   360
            Width           =   870
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   8
            Left            =   900
            TabIndex        =   72
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
         Begin VB.Label LblEED06_N 
            AutoSize        =   -1  'True
            Caption         =   "LblEED06_N"
            Height          =   180
            Left            =   1590
            TabIndex        =   71
            Top             =   690
            Width           =   910
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   6
            Left            =   900
            TabIndex        =   70
            Top             =   630
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "打字室："
            Height          =   180
            Index           =   10
            Left            =   180
            TabIndex        =   69
            Top             =   660
            Width           =   720
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   5
            Left            =   3540
            TabIndex        =   68
            Top             =   630
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "管制人："
            Height          =   180
            Index           =   9
            Left            =   2820
            TabIndex        =   67
            Top             =   690
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "中說備註："
            Height          =   180
            Left            =   0
            TabIndex        =   66
            Top             =   1020
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3075
         Left            =   120
         TabIndex        =   48
         Top             =   330
         Width           =   8655
         Begin VB.ListBox lstAtt 
            Height          =   2460
            Index           =   1
            IntegralHeight  =   0   'False
            ItemData        =   "frm100101_F_2.frx":0070
            Left            =   60
            List            =   "frm100101_F_2.frx":0077
            MultiSelect     =   1  '簡易多重選取
            Sorted          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   150
            Width           =   8520
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   345
            Index           =   1
            Left            =   210
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "下載"
            Height          =   345
            Index           =   1
            Left            =   1710
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            Height          =   345
            Index           =   1
            Left            =   960
            TabIndex        =   49
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
         TabIndex        =   22
         Top             =   1980
         Width           =   645
      End
      Begin VB.TextBox txtEEP10 
         Height          =   270
         Left            =   -74910
         TabIndex        =   21
         Top             =   3450
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "附件區："
         Height          =   3075
         Left            =   -70320
         TabIndex        =   4
         Top             =   1950
         Width           =   4155
         Begin VB.ListBox lstAtt 
            Height          =   2220
            Index           =   0
            IntegralHeight  =   0   'False
            ItemData        =   "frm100101_F_2.frx":0083
            Left            =   60
            List            =   "frm100101_F_2.frx":008A
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   210
            Width           =   4020
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  '沒有框線
            Height          =   405
            Left            =   120
            TabIndex        =   31
            Top             =   2610
            Width           =   2505
            Begin VB.CommandButton cmdOpenAtt 
               Caption         =   "開啟"
               Height          =   345
               Index           =   0
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   0
               Width           =   675
            End
            Begin VB.CommandButton cmdSaveAtt 
               Caption         =   "下載"
               Height          =   345
               Index           =   0
               Left            =   1440
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   0
               Width           =   675
            End
            Begin VB.CommandButton cmdSelect 
               Caption         =   "全選"
               Height          =   345
               Index           =   0
               Left            =   720
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   0
               Width           =   675
            End
         End
      End
      Begin VB.TextBox txtEEP03 
         BorderStyle     =   0  '沒有框線
         Height          =   260
         Left            =   -73980
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1950
         Width           =   645
      End
      Begin VB.ComboBox CboEEP04 
         BeginProperty Font 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   -73980
         TabIndex        =   2
         Text            =   "CboEEP04"
         Top             =   2220
         Width           =   2115
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   1395
         Left            =   -74940
         TabIndex        =   5
         Top             =   540
         Width           =   8775
         _ExtentX        =   15473
         _ExtentY        =   2477
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
         Left            =   -74910
         Top             =   3780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   -74880
         Top             =   4320
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   393216
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "作業備註："
         Height          =   180
         Left            =   -74940
         TabIndex        =   57
         Top             =   2790
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "請款備註："
         Height          =   180
         Left            =   -74940
         TabIndex        =   56
         Top             =   3600
         Width           =   900
      End
      Begin MSForms.TextBox txt3 
         Height          =   680
         Index           =   7
         Left            =   -74040
         TabIndex        =   55
         Top             =   3570
         Width           =   7890
         VariousPropertyBits=   -1466941413
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "13917;1199"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEP12 
         Height          =   800
         Left            =   -74040
         TabIndex        =   54
         Top             =   2760
         Width           =   7890
         VariousPropertyBits=   -1466941415
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "13917;1411"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblinfoNote 
         Caption         =   "註：檔名則以P000000.info.pdf，多個以上資料檔則加序號（例：P000000.info2.pdf，P000000.info3.pdf）"
         ForeColor       =   &H000000C0&
         Height          =   230
         Left            =   150
         TabIndex        =   53
         Top             =   4650
         Width           =   8600
      End
      Begin MSForms.TextBox txt1 
         Height          =   2280
         Index           =   4
         Left            =   -73710
         TabIndex        =   46
         Top             =   2550
         Width           =   4680
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "8255;4022"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   750
         Index           =   0
         Left            =   -73710
         TabIndex        =   45
         Top             =   1770
         Width           =   4680
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "8255;1323"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   570
         Index           =   6
         Left            =   -73710
         TabIndex        =   44
         Top             =   1170
         Width           =   4680
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "8255;1005"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   570
         Index           =   5
         Left            =   -73710
         TabIndex        =   43
         Top             =   570
         Width           =   4680
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "8255;1005"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEEP08 
         Height          =   1815
         Left            =   -74370
         TabIndex        =   42
         Top             =   3180
         Width           =   4005
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "7064;3201"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEEP10_2 
         Height          =   285
         Left            =   -73860
         TabIndex        =   41
         Top             =   2880
         Width           =   3495
         VariousPropertyBits=   679495707
         MaxLength       =   12
         Size            =   "6165;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox CboEEP05 
         Height          =   300
         Left            =   -73980
         TabIndex        =   40
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
      Begin MSForms.TextBox txtEEP03_2 
         Height          =   260
         Left            =   -73290
         TabIndex        =   39
         Top             =   1950
         Width           =   1545
         VariousPropertyBits=   679495707
         MaxLength       =   12
         Size            =   "2725;459"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
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
         Left            =   -68310
         TabIndex        =   30
         Top             =   90
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "順序："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   -71580
         TabIndex        =   23
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Left            =   -74790
         TabIndex        =   20
         Top             =   2595
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主旨："
         Height          =   180
         Index           =   2
         Left            =   -74790
         TabIndex        =   19
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "受文者："
         Height          =   180
         Left            =   -74790
         TabIndex        =   18
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "副本收受者："
         Height          =   180
         Left            =   -74790
         TabIndex        =   17
         Top             =   1245
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "(註:雙擊選取時,下方顯示歷程資料)"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -74940
         TabIndex        =   11
         Top             =   330
         Width           =   2895
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "收  受  者："
         Height          =   180
         Left            =   -74910
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "流程狀態："
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   8
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "內容："
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   7
         Top             =   3210
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "副本收受者："
         Height          =   180
         Left            =   -74910
         TabIndex        =   6
         Top             =   2910
         Width           =   1080
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   29
      Top             =   5850
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   180
      Index           =   5
      Left            =   3570
      TabIndex        =   27
      Top             =   240
      Width           =   930
   End
   Begin VB.Label lblCP10 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4530
      TabIndex        =   26
      Top             =   240
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   180
      Index           =   4
      Left            =   0
      TabIndex        =   25
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblCP09 
      Height          =   180
      Left            =   990
      TabIndex        =   24
      Top             =   240
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   19
      Left            =   0
      TabIndex        =   16
      Top             =   30
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   18
      Left            =   0
      TabIndex        =   15
      Top             =   450
      Width           =   960
   End
   Begin VB.Label lblCaseNo 
      Height          =   180
      Left            =   990
      TabIndex        =   14
      Top             =   30
      Width           =   1710
   End
   Begin VB.Label lblPA08 
      Height          =   180
      Left            =   4530
      TabIndex        =   13
      Top             =   30
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "種類："
      Height          =   180
      Index           =   1
      Left            =   3570
      TabIndex        =   12
      Top             =   30
      Width           =   930
   End
   Begin MSForms.Label lblCaseName 
      Height          =   195
      Left            =   990
      TabIndex        =   38
      Top             =   450
      Width           =   7935
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13996;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100101_F_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/14 Form2.0已修改
'Create by Sindy 2013/5/1
Option Explicit

'變數宣告區
Public m_EEP01 As String '歷程總收文號
Public m_CP09q As String '多案總收文號
Dim m_AttEEP02 As String '序號
Dim ii As Integer
Dim dblPrevRow As Double
Dim m_PrevForm As Form '前一畫面
Dim m_PA09 As String

'附件宣告區
Dim m_AttachPath As String
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
Dim cp() As String 'Add By Sindy 2018/4/26
Dim bolPAFlow As Boolean, bolTMFlow As Boolean 'Add By Sindy 2018/5/10
Dim bolOtherFlow As Boolean 'Add By Sindy 2021/7/15
Dim bolFCPFlow As Boolean 'Add By Sindy 2023/9/12
Dim bolFMP As Boolean 'Add By Sindy 2023/10/6
Dim bolOurFMP As Boolean '是否寰華案件 Add By Sindy 2023/10/6
Dim m_EP41 As String '核稿語文 1.英2.日 Add By Sindy 2019/11/27
Dim m_EEP15 As String 'Add By Sindy 2020/10/13
'Add By Sindy 2023/9/12
Const intTab_承辦單 As Integer = 1
Const intTab_外專承辦單 As Integer = 2
Const intTab_存卷資料 As Integer = 3
'2023/9/12 END
Dim m_EP12 As String 'Add By Sindy 2024/1/23


Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("順序", "EEP03", "發送者", "EEP04", "流程狀態", "EEP05", "收受者", "送出時間", "副本收受者", "意見內容", "EEP10", "c1.CP43", "ac03", "eep15")
   arrGridHeadWidth = Array(400, 0, 950, 0, 800, 0, 700, 1300, 1000, 3300, 0, 0, 0, 0)
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
Dim m_strSys As String 'Add By Sindy 2018/5/10
Dim bolReadEEP15 As Boolean 'Add by Sindy 2020/10/13
Dim bolSpecCase As Boolean, m_LimitType As String, m_RecvNo As String, strMsgTxt As String    'Added by Lydia 2025/11/17

   QueryData = True
   '清空及預設欄位值
   GRD1.Clear
   m_PA09 = Empty
   SetGrd
   lblCaseNo.Caption = Empty
   lblPA08.Caption = Empty
   lblCaseName.Caption = Empty
   Call ClearData
   Call SetCtrlReadOnly(False)
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   'Add By Sindy 2018/4/26
   '進度檔
   cp(9) = m_EEP01
   Call PUB_ReadCaseProgressDatabase(cp(), 國外_CF)
   '2018/4/26 END
   '電子送件
   If cp(118) <> "" Then
      lblEApp.Visible = True
   Else
      lblEApp.Visible = False
   End If
   
   'Add By Sindy 2020/9/30
   If m_CP09q <> "" And m_CP09q <> m_EEP01 Then
      Me.cmdCPPAtt.Visible = True
   Else
      Me.cmdCPPAtt.Visible = False
   End If
   '2020/9/30 END
   'Added by Lydia 2025/11/17 因區塊2之部分案件性質也有走承辦歷程，故請協助將承辦歷程一併納入權限管制範圍。---- from 教威
   bolSpecCase = PUB_ChkCPPAndCPFLimits_Spec(cp(1), cp(2), cp(3), cp(4), m_LimitType, m_RecvNo, strMsgTxt)
   If bolSpecCase = True Then
      If m_LimitType = "" Then
         '隱藏所有附件區
         Frame1.Visible = False
         Frame3.Visible = False
      End If
   End If
   'end 2025/11/17
   
   '案件資料
   'Modify By Sindy 2024/1/23 +,EP12
   strSql = "Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,PA09,CP140,EP41,EP12" & _
            " From CaseProgress,EngineerProgress,Patent," & _
            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And PA09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '1'=PTM01(+) AND PA08=PTM02(+)"
   'Add By Sindy 2015/10/21 +服務
   strSql = strSql & "union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,SP09,CP140,EP41,EP12" & _
            " From CaseProgress,EngineerProgress,Servicepractice," & _
            "staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And SP09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)"
   'Add By Sindy 2018/4/17 +商標檔
   'Modify By Sindy 2018/10/11 Decode(TM10,'000',PTM03,PTM04) ==> TM09
   strSql = strSql & "union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,TM05||TM06||TM07 as 案件名稱," & _
            "NA03 as 國家,TM09 as 種類,Decode(TM10,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,TM10,CP140,EP41,EP12" & _
            " From CaseProgress,EngineerProgress,Trademark," & _
            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And TM10=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '2'=PTM01(+) AND TM08=PTM02(+)"
   'Add By Sindy 2021/7/14
   strSql = strSql & "union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,LC05||LC06||LC07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(LC15,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,LC15,CP140,EP41,EP12" & _
            " From CaseProgress,EngineerProgress,LawCase," & _
            "staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And LC15=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)"
   strSql = strSql & "union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,HC06 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode('000','000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員," & _
            "CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP18,CP27,CP57,EP06,EP09,EP07,EP08,EP35,CP118,CP09,EP01,'000',CP140,EP41,EP12" & _
            " From CaseProgress,EngineerProgress,Hirecase," & _
            "staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And '000'=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)"
   '2021/7/14 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("本所案號")) Then lblCaseNo.Caption = rsTmp.Fields("本所案號")
      If Not IsNull(rsTmp.Fields("種類")) Then lblPA08.Caption = rsTmp.Fields("種類")
      If Not IsNull(rsTmp.Fields("案件名稱")) Then lblCaseName.Caption = rsTmp.Fields("案件名稱")
      If Not IsNull(rsTmp.Fields("PA09")) Then m_PA09 = rsTmp.Fields("PA09")
      
      lblCP09 = "" & rsTmp.Fields("CP09")
      lblCP10 = "" & rsTmp.Fields("案件性質")
      m_EP41 = "" & rsTmp.Fields("EP41") 'Add By Sindy 2019/11/27 核稿語文
      m_EP12 = "" & rsTmp.Fields("EP12") 'Add By Sindy 2024/1/23
      
      'Add By Sindy 2018/5/10
      'Modify By Sindy 2021/7/15
      m_strSys = CheckSys(cp(1))
      If InStr("1,5", m_strSys) > 0 Then '專利
         bolPAFlow = True
         'Add By Sindy 2023/9/12
         bolFMP = PUB_ChkIsFMP(cp(1), cp(2), cp(3), cp(4), m_PA09)
         If bolFMP = True Then
            bolOurFMP = PUB_FMPtoCheck(1, 2, PUB_GetST05(cp(14)), cp(1), cp(2), cp(3), cp(4)) '是否寰華案件
         Else
            bolOurFMP = False
         End If
         If cp(1) = "FCP" Or _
            cp(1) = "FG" Or _
            (bolFMP = True And Left(PUB_GetST03(cp(14)), 1) = "F") Then
            bolPAFlow = False
            SSTab1.TabVisible(intTab_承辦單) = False
            bolFCPFlow = True '外專Flow
            SSTab1.TabVisible(intTab_外專承辦單) = True
            LblinfoNote.Visible = False
         Else
            SSTab1.TabVisible(intTab_外專承辦單) = False
         End If
         '2023/9/12 END
      ElseIf InStr("2,6", m_strSys) > 0 Then '商標
         bolTMFlow = True
'      ElseIf InStr("5,6", m_strSys) > 0 Then
'         If m_strSys = "6" Then '商標:服務
'            bolTMFlow = True
'         Else
'            bolPAFlow = True
'         End If
      ElseIf InStr("3,4,7,8", m_strSys) > 0 Then '其他
         bolOtherFlow = True
'         Screen.MousePointer = vbDefault
'         MsgBox "讀取系統類別有誤，請洽電腦中心！", vbExclamation
'         QueryData = False
'         rsTmp.Close
'         Set rsTmp = Nothing
'         Call cmdExit_Click
'         Exit Function
      End If
      If bolTMFlow = True Then
         SSTab1.TabVisible(intTab_承辦單) = False
         SSTab1.TabVisible(intTab_外專承辦單) = False 'Add By Sindy 2023/9/12
         Label1(1).Caption = "類別："
      ElseIf bolOtherFlow = True Then
         SSTab1.TabVisible(intTab_承辦單) = False
         SSTab1.TabVisible(intTab_外專承辦單) = False 'Add By Sindy 2023/9/12
         Label1(1).Visible = False
         lblPA08.Visible = False
      End If
      '2018/5/10 END
      '2021/7/15 END
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
   
   'Add By Sindy 2020/12/1
   txtNote.Visible = False
   If cp(163) <> "" Then
      If cp(163) <> cp(9) Then
         strSql = "Select CP01,CP02,CP03,CP04,Decode('" & m_PA09 & "','000',CPM03,CPM04) as 案件性質" & _
                  " from caseprogress,CasePropertyMap" & _
                  " where cp09='" & cp(163) & "' And CP01=CPM01(+) And CP10=CPM02(+)"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            txtNote.Text = txtNote.Text & rsTmp.Fields("cp01") & "-" & rsTmp.Fields("cp02") & IIf(rsTmp.Fields("cp03") & rsTmp.Fields("cp04") = "000", "", "-" & rsTmp.Fields("cp03") & "-" & rsTmp.Fields("cp04")) & _
                           "(" & rsTmp.Fields("案件性質") & ")"
            txtNote.Width = 5000
            txtNote.Visible = True
         End If
         rsTmp.Close
      End If
   End If
   '2020/12/1 END
   
   Call QueryEmpElectronData '承辦單 Add By Sindy 2023/11/3
   
   '承辦電子簽核資料
   strSql = "Select distinct EEP02 as 順序,EEP03,s1.ST02||eep12 as 發送者,EEP04,decode(eep04,'" & EMP_附加流程 & "',decode(c2.CP43,'',ac03,Decode('" & m_PA09 & "','000',CPM03,CPM04)),ac03) as 流程狀態,EEP05,decode(s2.ST02,null,eep05,s2.ST02) as 收受者,sqldatet(EEP06)||' ' ||sqltime(EEP07) as 送出時間,EEP10 as 副本收受者,EEP08 as 意見內容,EEP10,c1.CP43,ac03,EEP15" & _
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
         'Add by Sindy 2020/10/13
         If bolReadEEP15 = False Then
            If Val(cp(27)) > 0 Then
               If rsTmp.Fields("EEP04") = EMP_送件 Or rsTmp.Fields("EEP04") = EMP_退件重送 Or _
                  rsTmp.Fields("EEP04") = EMP_發文歸檔 Then
                  m_EEP15 = "" & rsTmp.Fields("EEP15")
                  bolReadEEP15 = True
               End If
            Else
               If Left(PUB_GetST03(rsTmp.Fields("EEP03")), 2) = "P2" And _
                  rsTmp.Fields("EEP04") <> EMP_聯絡 Then '商標處
                  m_EEP15 = "" & rsTmp.Fields("EEP15")
                  bolReadEEP15 = True
               End If
            End If
         End If
         '2020/10/13 END
         txtEEP10_2 = GRD1.TextMatrix(ii, 10)
         Call txtEEP10_2_LostFocus
         GRD1.TextMatrix(ii, 8) = txtEEP10_2
         '判斷有相關總收文號才做案件性質轉換
         If GRD1.TextMatrix(ii, 4) = "附加流程" Then
            If GRD1.TextMatrix(ii, 11) <> "" Then
               GRD1.TextMatrix(ii, 4) = Trim(lblCP10) & PUB_GetRelateCasePropertyName(m_EEP01, "1")
            End If
         'Add By Sindy 2019/11/27
         ElseIf GRD1.TextMatrix(ii, 4) = "送英核" And m_EP41 = "2" Then
            GRD1.TextMatrix(ii, 4) = "送日核"
         '2019/11/27 END
         End If
      Next ii
      Call ReadData(True)
   End If
   rsTmp.Close
   
   Call ReadAttachFile_other(m_EEP01)
   
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
      If cp(10) = "209" Or cp(10) = "235" Then
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
      
      'Add By Sindy 2024/1/23 承辦備註改顯示為作業備註
      txtEP12.Text = m_EP12
      txtEP12.Enabled = True
      txtEP12.Locked = True
      '2024/1/23 END
   '內專
   ElseIf bolPAFlow = True Then
      '讀取承辦單內容
      txt1(5) = Empty
      txt1(6) = Empty
      txt1(0) = Empty
      txt1(4) = Empty
      strSql = "Select *" & _
               " From EmpElectronData,staff" & _
               " Where EED01='" & m_EEP01 & "' and EED06=ST01(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If Not IsNull(rsTmp.Fields("EED02")) Then txt1(5).Text = rsTmp.Fields("EED02")
         If Not IsNull(rsTmp.Fields("EED03")) Then txt1(6).Text = rsTmp.Fields("EED03")
         '主旨
         If Not IsNull(rsTmp.Fields("EED04")) Then txt1(0).Text = rsTmp.Fields("EED04")
         If Not IsNull(rsTmp.Fields("EED05")) Then txt1(4).Text = rsTmp.Fields("EED05")
      End If
      rsTmp.Close
   End If
   
   Set rsTmp = Nothing
   Set Rs = Nothing
End Sub

'Add By Sindy 2025/10/8 點二下可以開啟附件檔案
Private Sub lstAtt_DblClick(Index As Integer)
   Call cmdOpenAtt_Click(Index)
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
            'Modify By Sindy 2024/6/6 mark,查詢作業不需要檢查
'            '檢查人員是否存在或離職
'            If ChkStaffST04(Trim(txt3(Index).Text)) = True Then
'               txt3(Index).SetFocus
'               Exit Sub
'            Else
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
'            End If
         End If
   End Select
End Sub

Private Sub txtEEP10_2_LostFocus()
Dim strText As String
Dim arrID
Dim strTempName As String
Dim strEEP10_Err As String
   
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
      If Trim(strEEP10_Err) <> "" Then
         MsgBox "副本收受者資料有誤！(" & strEEP10_Err & ")"
         txtEEP10_2.SetFocus
         'Call txtEEP10_2_GotFocus
         'Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2020/10/12
Private Sub SetTxtLpNote(bolQueryStar As Boolean)
   If bolQueryStar = True Then
      txtLpNote.Visible = False
      txtLpNote = ""
      If m_EEP15 <> "" And cp(163) <> "" Then
         txtLpNote.Visible = True
         txtLpNote = "(共" & UBound(Split(m_EEP15, ",")) + 1 & "筆)"
      End If
   Else
      txtLpNote.Visible = False
      txtLpNote = ""
   End If
End Sub

Private Sub ClearData()
   m_AttEEP02 = Empty
   txtEEP02 = Empty
   txtEEP03 = Empty
   txtEEP03_2 = Empty
   CboEEP04.Clear
   CboEEP05.Clear
   txtEEP10 = Empty
   txtEEP10_2 = Empty
   txtEEP08 = Empty
   lstAtt(0).Clear
'   Me.cmdOpenAtt(0).Enabled = False
'   Me.cmdSelect(0).Enabled = False
'   Me.cmdSaveAtt(0).Enabled = False
End Sub

'Add By Sindy 2020/9/30
Private Sub cmdCPPAtt_Click()
Dim rsQuery As ADODB.Recordset
Dim ii As Integer, jj As Integer
Dim strCP09 As String, strCP10 As String
Dim stFileName As String '附件檔名(含路徑,多檔以[;]號區隔)
Dim stFileDescs As String '附件檔名及說明(不含路徑,多檔以[;]號區隔)
Dim stSavePath As String, stTempFile As String, stFileList As String
Dim stSQL As String
Dim strSavePathFile As String
   
On Error GoTo ErrHnd
      
   stTempFile = cp(1) & cp(2) & IIf(cp(3) & cp(4) = "000", "", cp(3) & cp(4)) & ".pdf"
   
   stSavePath = App.path & "\CustLetter"
   If Dir(stSavePath, vbDirectory) = "" Then
      MkDir stSavePath
   Else
      PUB_KillAttach stSavePath
      If Dir(stSavePath & "\.") <> "" Then
         Kill stSavePath & "\*.*"
      End If
   End If
      
   Screen.MousePointer = vbHourglass
   If m_EEP01 <> "" Then
      stSQL = "select casepaperpdf.*,'('||Round(cpp03 / 1024, 2)||' KB) '||GETFILEDESC(cpp02,CP01,CP10,'" & m_PA09 & "') as FDesc" & _
              " from casepaperpdf,caseprogress" & _
              " where cpp01='" & m_EEP01 & "' and cp09(+)=cpp01" & _
              " and substr(upper(cpp02),-5)<>upper('.menu')"
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, stSQL)
      If intI = 1 Then
         With rsQuery
         Do While Not .EOF
            '檔名含空白無法合併
            strSavePathFile = stSavePath & "\" & Replace(.Fields("cpp02"), " ", "_")
            If Dir(strSavePathFile) = "" Then  '檔案存在時不必再下載
               If PUB_GetAttachFile_CPP(.Fields("cpp01"), .Fields("cpp02"), strSavePathFile, True) = False Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            '檔名含空白無法合併
            If UCase(Right(Trim(.Fields("cpp02")), 4)) = ".PDF" Then
               stFileName = IIf(stFileName = "", "", stFileName & ";") & strSavePathFile
               stFileList = IIf(stFileList = "", "", stFileList & ";") & Replace(.Fields("cpp02"), " ", "_")
            End If
            stFileDescs = stFileDescs & Replace(.Fields("cpp02"), " ", "_") & Chr(9) & .Fields("FDesc") & ";"
            .MoveNext
         Loop
         End With
         
         '一個也要更名，否則EMail當作附件時會有鎖住問題
         If Dir(stSavePath & "\" & stTempFile) = "" Then
            If PUB_JoinPdf(stFileList, stTempFile, stSavePath) = True Then
               stFileName = stSavePath & "\" & stTempFile
            End If
         Else
            stFileName = stSavePath & "\" & stTempFile
         End If
         
         frm100101_2_1.Caption = "多案卷宗區附件"
         frm100101_2_1.stFileName = stTempFile 'stFileName
         frm100101_2_1.stFileDescs = stFileDescs
         frm100101_2_1.stSavePath = stSavePath
         Screen.MousePointer = vbDefault
         frm100101_2_1.Show vbModal
      Else
         MsgBox "無卷宗區資料！", vbExclamation
      End If
   End If
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHnd:
   Screen.MousePointer = vbDefault
   If Err.Number = 75 Then '路徑或檔案存取錯誤:檔案刪不掉
      Resume Next
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set rsQuery = Nothing
End Sub

'結束
Private Sub cmdExit_Click()
   m_PrevForm.Show
   Unload Me
End Sub

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'全選
Private Sub cmdSelect_Click(Index As Integer)
   Dim ii As Integer, oList As Object
   
   Set oList = lstAtt(Index)
   For ii = 0 To oList.ListCount - 1
      lstAtt(Index).Selected(ii) = True
   Next
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   Me.txtEEP02.BackColor = &H8000000F
   Me.txtEEP03.BackColor = &H8000000F
   Me.txtEEP03_2.BackColor = &H8000000F
   ReDim m_FilesRemoved(0)
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   SSTab1.Tab = 0
   txtPDFPath = PUB_SetFileAssociation
   ReDim cp(TF_CP) 'Add By Sindy 2018/4/26
   
'   'Added by Sindy 2021/7/14 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
'   lstAtt(0).Height = 2400
'   lstAtt(0).Width = 4020
'   lstAtt(1).Height = 2400
'   lstAtt(1).Width = 8490

   Frame7.BorderStyle = 0 'Add By Sindy 2023/9/23
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   
   Set m_PrevForm = Nothing
   Set frm100101_F_2 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
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
   Call ClearData
   Call SetCtrlReadOnly(False)
   
   If bolFirst = True Then
      '若有資料游標停在第一筆
      GRD1.Visible = False
      GRD1.col = 0
      GRD1.row = 1
      dblPrevRow = GRD1.row
      For ii = 0 To GRD1.Cols - 1
         GRD1.col = ii
         GRD1.CellBackColor = &HFFC0C0
      Next ii
      GRD1.Visible = True
   End If
   
   m_AttEEP02 = GRD1.TextMatrix(dblPrevRow, 0)
   txtEEP02 = GRD1.TextMatrix(dblPrevRow, 0)
   txtEEP03 = GRD1.TextMatrix(dblPrevRow, 1)
   txtEEP03_2 = GRD1.TextMatrix(dblPrevRow, 2)
   CboEEP04.Text = GRD1.TextMatrix(dblPrevRow, 3) & " " & GRD1.TextMatrix(dblPrevRow, 12) '4
   CboEEP05.Text = GRD1.TextMatrix(dblPrevRow, 5) & " " & GRD1.TextMatrix(dblPrevRow, 6)
   txtEEP10_2 = GRD1.TextMatrix(dblPrevRow, 8)
   txtEEP08 = GRD1.TextMatrix(dblPrevRow, 9)
   txtEEP10 = GRD1.TextMatrix(dblPrevRow, 10)
   '讀取附件
   Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
End Sub

'查詢附件檔
Private Sub ReadAttachFile(strEEP01 As String, intEEP02 As Integer, bolQuery As Boolean)
Dim intEEFCnt As Integer
   
   KillAttach
   lstAtt(0).Clear
   Frame2.Visible = True 'Add By Sindy 2018/9/4
   intEEFCnt = 0
   'Modify By Sindy 2020/10/30 + decode(sign(instr(upper(eef03),upper('" & EMP_多案承辦單 & "'))),1,1...
   strExc(0) = "select eef03,eef04,eef09,eef10," & _
                      "decode(sign(instr(upper(eef03),upper('" & EMP_承辦單 & "'))),1,1,decode(sign(instr(upper(eef03),upper('" & EMP_多案承辦單 & "'))),1,1,decode(sign(instr(upper(eef03),upper('DWG'))),1,3,2))) as sort,eef12" & _
               " from EmpElectronFile where eef01='" & strEEP01 & "' and eef02=" & intEEP02 & _
               " order by sort desc,eef03 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         intEEFCnt = .RecordCount
         Frame2.Visible = False 'Add By Sindy 2023/11/30
         Do While Not .EOF
            lstAtt(0).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)" & " #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "#", 0
            'lstAtt(0).ItemData(0) = 1
            
            'Add By Sindy 2018/9/4
            If "" & .Fields("eef12") <> "" Then
               Frame2.Visible = True 'Modify By Sindy 2023/11/30
            End If
            '2018/9/4 END
            .MoveNext
         Loop
      End With
   Else
'      'Add By Sindy 2018/9/4
'      If bolQuery = True Then
'         Frame2.Visible = True
'   '      Me.cmdOpenAtt(0).Enabled = True
'   '      Me.cmdSelect(0).Enabled = True
'   '      Me.cmdSaveAtt(0).Enabled = True
'      Else
         Frame2.Visible = False
'      End If
   End If
   If lstAtt(0).ListCount > 0 Then SetListScroll lstAtt(0)
End Sub

'查詢存卷區
Private Sub ReadAttachFile_other(strEEP01 As String)
   KillAttach
   lstAtt(1).Clear
   strExc(0) = "select eef03,eef04,eef09,eef10 from EmpElectronFile where eef01='" & strEEP01 & "' and eef02=0 order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Do While Not .EOF
            lstAtt(1).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)" & " #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "#", 0
            'lstAtt(1).ItemData(0) = 1
            .MoveNext
         Loop
      End With
   'Add By Sindy 2023/11/8
      Me.cmdOpenAtt(1).Visible = True
      Me.cmdSelect(1).Visible = True
      Me.cmdSaveAtt(1).Visible = True
   Else
      Me.cmdOpenAtt(1).Visible = False
      Me.cmdSelect(1).Visible = False
      Me.cmdSaveAtt(1).Visible = False
      '2023/11/8 END
   End If
   If lstAtt(1).ListCount > 0 Then SetListScroll lstAtt(1)
End Sub

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

'Private Function GetAttachFile(ByRef pFileName As String, intEEP02 As Integer, Optional pSavePath As String) As Boolean
'   Dim stAttPath As String
'   Dim lngSize As Long
'   Dim iFileNo As Integer
'   Dim bytes() As Byte
'
'On Error GoTo ErrHnd
'
'   If pSavePath = "" Then
'      If Dir(m_AttachPath, vbDirectory) = "" Then
'         MkDir m_AttachPath
'      End If
'      stAttPath = m_AttachPath & "\" & pFileName
'      '檔案已存在時
'      If Dir(stAttPath) <> "" Then
'         '檢查檔案是否正在使用中
'         If PUB_ChkFileOpening(stAttPath) = True Then
'            MsgBox stAttPath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
'            Exit Function
'         End If
'         SetAttr stAttPath, vbNormal 'Add By Sindy 2020/1/17 檔案設定為正常屬性
'         Kill stAttPath
'         '不必重新下載
''         pFileName = stAttPath
''         GetAttachFile = True
''         Exit Function
'      End If
'   Else
'      stAttPath = pSavePath
'   End If
'
'   'Added by Morgan 2015/4/28
'   'Modified by Morgan 2015/5/22 FTP上線
'   'Modify By Sindy 2018/9/4 + and eef12 is not null
'      strExc(0) = "select eef12 from EmpElectronFile where eef01='" & m_EEP01 & "' and eef02=" & intEEP02 & " and eef12 is not null" & _
'               " and eef03='" & ChgSQL(pFileName) & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If Not IsNull(RsTemp(0)) Then
'            pFileName = stAttPath
'            GetAttachFile = PUB_GetFtpFile(RsTemp(0), stAttPath, "EMPELECTRONFILE", True)
'         End If
'      End If
'      Exit Function
'   'end 2015/4/28
'
'ErrHnd:
'   MsgBox Err.Description, vbCritical
'   If iFileNo > 0 Then Close #iFileNo
'End Function

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
'   strAtt = lstAtt(Index).Text
'
'   If strAtt = "" Then
'      MsgBox "請選擇欲開啟的附件！"
'   Else
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
                  'If GetAttachFile(stFileName, 0) = False Then
                  If PUB_GetAttachFile_EEF(m_EEP01, 0, stFileName, m_AttachPath) = False Then
                     'Add By Sindy 2018/9/4
                     Screen.MousePointer = vbDefault
                     MsgBox "已無附件實體檔！"
                     '2018/9/4 END
                     Exit Sub
                  End If
               Else
                  'If GetAttachFile(stFileName, CInt(m_AttEEP02)) = False Then
                  If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_AttEEP02), stFileName, m_AttachPath) = False Then
                     'Add By Sindy 2018/9/4
                     Screen.MousePointer = vbDefault
                     MsgBox "已無附件實體檔！"
                     '2018/9/4 END
                     Exit Sub
                  End If
               End If
            End If
            SetAttr stFileName, vbReadOnly 'Add By Sindy 2020/1/17 檔案設定成唯讀屬性
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Next ii
      If bolIsSelect = False Then
         MsgBox "請選擇欲開啟的附件！"
      End If
'   End If
   
   Screen.MousePointer = vbDefault
End Sub

'下載
Private Sub cmdSaveAtt_Click(Index As Integer)
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer, oList As Object
   Dim stChkFileName As String 'Add By Sindy 2018/9/4
   
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
                  'Add By Sindy 2018/9/4
                  stChkFileName = stFileName
                  If stChkFileName <> "" Then
                     If Index = 1 Then '存卷資料
                        'If GetAttachFile(stChkFileName, 0) = False Then
                        If PUB_GetAttachFile_EEF(m_EEP01, 0, stChkFileName, m_AttachPath) = False Then
                           'MsgBox "無法儲存檔案[ " & stChkFileName & " ]！"
                           MsgBox "已無附件實體檔[ " & stChkFileName & " ]！"
                           GoTo RunExit
                        End If
                     Else
                        'If GetAttachFile(stChkFileName, CInt(m_AttEEP02)) = False Then
                        If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_AttEEP02), stChkFileName, m_AttachPath) = False Then
                           'MsgBox "無法儲存檔案[ " & stChkFileName & " ]！"
                           MsgBox "已無附件實體檔[ " & stChkFileName & " ]！"
                           GoTo RunExit
                        End If
                     End If
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
                           'If GetAttachFile(stFileName, 0, stFullName) = False Then
                           If PUB_GetAttachFile_EEF(m_EEP01, 0, stFileName, stFullName, True) = False Then
                              MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                              GoTo RunExit
                           End If
                        Else
                           'If GetAttachFile(stFileName, CInt(m_AttEEP02), stFullName) = False Then
                           If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_AttEEP02), stFileName, stFullName, True) = False Then
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
         'Add By Sindy 2018/9/4
         stChkFileName = stFileName
         If stChkFileName <> "" Then
            If Index = 1 Then '存卷資料
               'If GetAttachFile(stChkFileName, 0) = False Then
               If PUB_GetAttachFile_EEF(m_EEP01, 0, stChkFileName, m_AttachPath) = False Then
                  'MsgBox "無法儲存檔案[ " & stChkFileName & " ]！"
                  'Modify By Sindy 2018/9/4
                  MsgBox "已無附件實體檔[ " & stChkFileName & " ]！"
                  '2018/9/4 END
                  GoTo RunExit
               End If
            Else
               'If GetAttachFile(stChkFileName, CInt(m_AttEEP02)) = False Then
               If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_AttEEP02), stChkFileName, m_AttachPath) = False Then
                  'MsgBox "無法儲存檔案[ " & stChkFileName & " ]！"
                  'Modify By Sindy 2018/9/4
                  MsgBox "已無附件實體檔[ " & stChkFileName & " ]！"
                  '2018/9/4 END
                  GoTo RunExit
               End If
            End If
         End If
         '2018/9/4 END
         'Modify By Sindy 2020/1/16
         'stFullName = GetSaveName(stFileName)
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
                  'If GetAttachFile(stFileName, 0, stFullName) = False Then
                  If PUB_GetAttachFile_EEF(m_EEP01, 0, stFileName, stFullName, True) = False Then
                     MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                     GoTo RunExit
                  End If
               Else
                  'If GetAttachFile(stFileName, CInt(m_AttEEP02), stFullName) = False Then
                  If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_AttEEP02), stFileName, stFullName, True) = False Then
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

Private Function GetSaveName(ByVal pFileName As String) As String
   
On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
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

Private Sub SetListScroll(oList As Object)
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

'Add By Sindy 2015/9/17
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
GRD1.ToolTipText = ""
If GRD1.MouseRow <> 0 And GRD1.MouseCol > 0 Then
   If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
      GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
   End If
End If
End Sub
