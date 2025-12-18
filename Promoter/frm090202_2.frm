VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090202_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子承辦單簽辦作業"
   ClientHeight    =   6120
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   8964
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.6
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8964
   Tag             =   "加班資料"
   Begin VB.CommandButton cmdOutlook 
      Caption         =   "匯出Outlook"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   5580
      Style           =   1  '圖片外觀
      TabIndex        =   125
      Top             =   1080
      Visible         =   0   'False
      Width           =   1010
   End
   Begin VB.CommandButton CmdCalendar 
      Caption         =   "行事曆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   6600
      Style           =   1  '圖片外觀
      TabIndex        =   148
      Top             =   1080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "相似案"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   1
      Left            =   7330
      TabIndex        =   124
      Top             =   1080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "原始檔區"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   5
      Left            =   8070
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   1080
      Width           =   885
   End
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   8040
      TabIndex        =   161
      Top             =   5970
      Visible         =   0   'False
      Width           =   1700
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "電腦中心刪除歷程(&D)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   6810
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   1380
      Width           =   2085
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3090
      Locked          =   -1  'True
      MousePointer    =   1  '箭號形狀
      TabIndex        =   43
      Text            =   "存卷資料"
      Top             =   1440
      Width           =   1305
   End
   Begin VB.CommandButton cmdManyCase 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      Picture         =   "frm090202_2.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   40
      Top             =   1110
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame6"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   0
      TabIndex        =   33
      Top             =   780
      Visible         =   0   'False
      Width           =   5205
      Begin VB.OptionButton Option1 
         Caption         =   "部分勝部分敗"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   39
         Top             =   30
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "敗 (駁)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   38
         Top             =   30
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Caption         =   "勝 (准)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   37
         Top             =   30
         Width           =   1005
      End
      Begin VB.TextBox txt2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   900
         TabIndex        =   34
         Top             =   240
         Width           =   4275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "預    估："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   30
         TabIndex        =   36
         Top             =   30
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "條款代碼："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   35
         Top             =   270
         Width           =   900
      End
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
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
      Left            =   30
      TabIndex        =   42
      Text            =   "※此案屬多案歷程，請參"
      Top             =   810
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtLpNote 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
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
      Left            =   2610
      TabIndex        =   41
      Text            =   "(共X筆)"
      Top             =   30
      Width           =   1035
   End
   Begin VB.CommandButton cmdCP118 
      Caption         =   "電子送件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   6435
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdTMQ 
      Caption         =   "查名區"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   7330
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "變更事項"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   5190
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "接洽單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   2
      Left            =   7330
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   720
      Width           =   705
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "完整卷宗"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   4
      Left            =   8070
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   720
      Width           =   885
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "商品名稱維護"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   0
      Left            =   6090
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   1
      Left            =   7100
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   0
      Left            =   6000
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   360
      Width           =   1080
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增下一流程(&A)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   5820
      TabIndex        =   3
      Top             =   0
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   6450
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      TabIndex        =   27
      Top             =   -30
      Width           =   1695
      Begin VB.CommandButton cmdChkEP08 
         Caption         =   "確認會稿完成日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   28
         Top             =   60
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "承辦進度(&E)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   3
      Left            =   7880
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   360
      Width           =   1080
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "送出(&O)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   7335
      TabIndex        =   5
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   8160
      TabIndex        =   17
      Top             =   0
      Width           =   800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4540
      Left            =   30
      TabIndex        =   45
      Top             =   1410
      Width           =   8930
      _ExtentX        =   15748
      _ExtentY        =   8022
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "簽辦流程"
      TabPicture(0)   =   "frm090202_2.frx":00FA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lbl926"
      Tab(0).Control(1)=   "CboEEP05"
      Tab(0).Control(2)=   "txtEEP03_2"
      Tab(0).Control(3)=   "txtEEP08"
      Tab(0).Control(4)=   "lblCMboth"
      Tab(0).Control(5)=   "lblCM10"
      Tab(0).Control(6)=   "lblEApp"
      Tab(0).Control(7)=   "Label2(0)"
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(11)=   "Label1(0)"
      Tab(0).Control(12)=   "Label10"
      Tab(0).Control(13)=   "Label3(0)"
      Tab(0).Control(14)=   "Label5"
      Tab(0).Control(15)=   "txtEEP10_2"
      Tab(0).Control(16)=   "cmdRemAttDB(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdAddAttDB(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CommonDialog1"
      Tab(0).Control(19)=   "GRD1"
      Tab(0).Control(20)=   "ChkEMail"
      Tab(0).Control(21)=   "cmdCaseMap"
      Tab(0).Control(22)=   "Text1"
      Tab(0).Control(23)=   "Check1"
      Tab(0).Control(24)=   "txtEEP10"
      Tab(0).Control(25)=   "Frame1"
      Tab(0).Control(26)=   "txtEEP03"
      Tab(0).Control(27)=   "CboEEP04"
      Tab(0).Control(28)=   "ChkEP11"
      Tab(0).Control(29)=   "Frame4"
      Tab(0).Control(30)=   "ChkEED08"
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "承辦單內容"
      TabPicture(1)   =   "frm090202_2.frx":0116
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(4)"
      Tab(1).Control(1)=   "txt1(0)"
      Tab(1).Control(2)=   "txt1(5)"
      Tab(1).Control(3)=   "txt1(6)"
      Tab(1).Control(4)=   "Label2(14)"
      Tab(1).Control(5)=   "Label21(1)"
      Tab(1).Control(6)=   "Label13"
      Tab(1).Control(7)=   "Label1(2)"
      Tab(1).Control(8)=   "Label9"
      Tab(1).Control(9)=   "Label7"
      Tab(1).Control(10)=   "cmdSave2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Combo2"
      Tab(1).Control(12)=   "Frame2"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "歷程備註"
      TabPicture(2)   =   "frm090202_2.frx":0132
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtEP12"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label24"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label23"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txt3(7)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label18"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame201"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame945"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdSave3"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "存卷資料"
      TabPicture(3)   =   "frm090202_2.frx":014E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LblinfoNote"
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(2)=   "Frame5"
      Tab(3).Control(3)=   "cmdAddAttDB(1)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdRemAttDB(1)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.CheckBox ChkEED08 
         BackColor       =   &H0000FFFF&
         Caption         =   "【需收文告代】"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -72420
         TabIndex        =   146
         Top             =   2250
         Visible         =   0   'False
         Width           =   2150
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame4"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72450
         TabIndex        =   69
         Top             =   2190
         Visible         =   0   'False
         Width           =   2300
         Begin VB.ComboBox CboCP10 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            ItemData        =   "frm090202_2.frx":016A
            Left            =   780
            List            =   "frm090202_2.frx":016C
            Style           =   2  '單純下拉式
            TabIndex        =   71
            Top             =   0
            Width           =   1485
         End
         Begin VB.CommandButton cmdMail 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2010
            Picture         =   "frm090202_2.frx":016E
            Style           =   1  '圖片外觀
            TabIndex        =   70
            Top             =   0
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "案件性質："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   0
            TabIndex        =   72
            Top             =   30
            Width           =   900
         End
      End
      Begin VB.CheckBox ChkEP11 
         BackColor       =   &H0000FFFF&
         Caption         =   "【不通知客戶, 不發文】"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -72510
         TabIndex        =   62
         Top             =   2250
         Visible         =   0   'False
         Width           =   2360
      End
      Begin VB.CommandButton cmdSave3 
         Caption         =   "存檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5970
         Style           =   1  '圖片外觀
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   330
         Width           =   830
      End
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
         Left            =   60
         TabIndex        =   144
         Top             =   420
         Visible         =   0   'False
         Width           =   5350
         Begin VB.Label Label1 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "委員指定送件日期【本所期限】："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   12
            Left            =   90
            TabIndex        =   147
            Top             =   690
            Width           =   2700
         End
         Begin MSForms.TextBox txtEED15 
            Height          =   290
            Left            =   2820
            TabIndex        =   47
            Top             =   630
            Width           =   1020
            VariousPropertyBits=   680542235
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
            Caption         =   "追蹤客戶指示【約定期限】："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   11
            Left            =   90
            TabIndex        =   145
            Top             =   360
            Width           =   2700
         End
         Begin MSForms.TextBox txtEED14 
            Height          =   300
            Left            =   2820
            TabIndex        =   46
            Top             =   300
            Width           =   1020
            VariousPropertyBits=   680542235
            MaxLength       =   7
            ScrollBars      =   2
            Size            =   "1799;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame Frame201 
         Height          =   2320
         Left            =   30
         TabIndex        =   126
         Top             =   390
         Width           =   8860
         Begin VB.ComboBox CmbFL 
            Height          =   280
            Index           =   3
            Left            =   930
            Locked          =   -1  'True
            TabIndex        =   133
            Text            =   "CmbFL"
            Top             =   300
            Width           =   6060
         End
         Begin VB.CheckBox ChkEED13 
            Caption         =   "轉檔後送件（程序發文）"
            ForeColor       =   &H000000C0&
            Height          =   220
            Left            =   5670
            TabIndex        =   132
            Top             =   660
            Width           =   2740
         End
         Begin VB.Frame Frame7 
            Height          =   280
            Left            =   180
            TabIndex        =   127
            Top             =   0
            Width           =   3970
            Begin VB.Label LblEED10_N_2 
               AutoSize        =   -1  'True
               Caption         =   "LblEED10_N_2"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   2550
               TabIndex        =   131
               Top             =   60
               Width           =   1070
            End
            Begin MSForms.TextBox txt3 
               Height          =   320
               Index           =   3
               Left            =   750
               TabIndex        =   130
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
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   210
               TabIndex        =   129
               Top             =   30
               Width           =   540
            End
            Begin VB.Label LblEED10_N 
               AutoSize        =   -1  'True
               Caption         =   "LblEED10_N"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   1470
               TabIndex        =   128
               Top             =   60
               Width           =   910
            End
         End
         Begin MSForms.TextBox txt3 
            Height          =   1320
            Index           =   4
            Left            =   930
            TabIndex        =   143
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
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   8
            Left            =   30
            TabIndex        =   142
            Top             =   360
            Width           =   870
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   8
            Left            =   930
            TabIndex        =   141
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
         Begin VB.Label LblEED09_N 
            AutoSize        =   -1  'True
            Caption         =   "LblEED09_N"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   4320
            TabIndex        =   140
            Top             =   690
            Width           =   910
         End
         Begin VB.Label LblEED06_N 
            AutoSize        =   -1  'True
            Caption         =   "LblEED06_N"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1620
            TabIndex        =   139
            Top             =   690
            Width           =   910
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   6
            Left            =   930
            TabIndex        =   138
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
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   10
            Left            =   210
            TabIndex        =   137
            Top             =   660
            Width           =   720
         End
         Begin MSForms.TextBox txt3 
            Height          =   320
            Index           =   5
            Left            =   3570
            TabIndex        =   136
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
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   9
            Left            =   2850
            TabIndex        =   135
            Top             =   690
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "中說備註："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   30
            TabIndex        =   134
            Top             =   1020
            Width           =   900
         End
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
         Height          =   300
         Left            =   -74010
         TabIndex        =   90
         Text            =   "CboEEP04"
         Top             =   2160
         Width           =   1485
      End
      Begin VB.TextBox txtEEP03 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74010
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   1920
         Width           =   645
      End
      Begin VB.Frame Frame1 
         Caption         =   "附件區："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   -70140
         TabIndex        =   79
         Top             =   1920
         Width           =   4005
         Begin VB.CommandButton CmdOpen 
            BackColor       =   &H00FFFFC0&
            Caption         =   "<->"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   870
            Style           =   1  '圖片外觀
            TabIndex        =   150
            Top             =   -120
            Width           =   460
         End
         Begin VB.Frame FrameFCPlink 
            Caption         =   "(註: 按下路徑可開啟資料夾)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   490
            Index           =   0
            Left            =   60
            TabIndex        =   80
            Top             =   1560
            Visible         =   0   'False
            Width           =   3910
            Begin VB.CommandButton CmdF21 
               BackColor       =   &H00FFC0C0&
               Caption         =   "上傳"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Index           =   0
               Left            =   1380
               Style           =   1  '圖片外觀
               TabIndex        =   162
               Top             =   180
               Visible         =   0   'False
               Width           =   530
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "電子送件暫存區"
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
               Left            =   60
               TabIndex        =   82
               Top             =   240
               Width           =   1260
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "\\typing2\外專送件"
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
               Left            =   2460
               TabIndex        =   81
               Top             =   240
               Width           =   1370
            End
         End
         Begin VB.ListBox lstAtt 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1308
            Index           =   0
            ItemData        =   "frm090202_2.frx":0242
            Left            =   60
            List            =   "frm090202_2.frx":0249
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   210
            Width           =   3900
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   210
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   2100
            Width           =   675
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "下載"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   1710
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   2100
            Width           =   675
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "新增"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   2460
            TabIndex        =   85
            Top             =   2100
            Width           =   675
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "刪除"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   3210
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   2100
            Width           =   675
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   960
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   2100
            Width           =   675
         End
      End
      Begin VB.TextBox txtEEP10 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74910
         TabIndex        =   78
         Top             =   3270
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74820
         TabIndex        =   74
         Top             =   390
         Width           =   8625
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "代理人："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   0
            TabIndex        =   77
            Top             =   60
            Width           =   720
         End
         Begin MSForms.Label lblFa 
            Height          =   290
            Left            =   2550
            TabIndex        =   76
            Top             =   30
            Width           =   6110
            BackColor       =   -2147483634
            VariousPropertyBits=   27
            Caption         =   "lblFa"
            Size            =   "10769;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txt1 
            Height          =   320
            Index           =   7
            Left            =   1110
            TabIndex        =   75
            Top             =   30
            Width           =   1410
            VariousPropertyBits=   680542235
            MaxLength       =   9
            ScrollBars      =   2
            Size            =   "2487;556"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68700
         TabIndex        =   73
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "一併更新英文核完日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -72480
         TabIndex        =   66
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   264
         Left            =   -74640
         TabIndex        =   65
         Text            =   "「聯絡」的附件，送件後一律刪除，欲留存者請置於「存卷資料」。"
         Top             =   1590
         Width           =   8055
      End
      Begin VB.CommandButton cmdSave2 
         Caption         =   "存檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -68730
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2730
         Width           =   855
      End
      Begin VB.CommandButton cmdCaseMap 
         BackColor       =   &H0080FF80&
         Caption         =   "多國案"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -70900
         Style           =   1  '圖片外觀
         TabIndex        =   63
         Top             =   2720
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.CheckBox ChkEMail 
         Caption         =   "E-Mail夾帶附件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -71850
         TabIndex        =   61
         Top             =   2220
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdRemAttDB 
         Caption         =   "刪除歷程附件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   -73200
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3750
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddAttDB 
         Caption         =   "新增歷程附件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   -74640
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3750
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   -74850
         TabIndex        =   51
         Top             =   330
         Width           =   8655
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   960
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "刪除"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   3210
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "新增"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   2460
            TabIndex        =   56
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "下載"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1710
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   210
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   2670
            Width           =   675
         End
         Begin VB.ListBox lstAtt 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2388
            Index           =   1
            ItemData        =   "frm090202_2.frx":0255
            Left            =   60
            List            =   "frm090202_2.frx":025C
            MultiSelect     =   1  '簡易多重選取
            Sorted          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   180
            Width           =   8490
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5520
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2610
            Width           =   855
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   1400
         Left            =   -74940
         TabIndex        =   91
         Top             =   480
         Width           =   8780
         _ExtentX        =   15473
         _ExtentY        =   2455
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "順序|發送者 |流程狀態 |收受者 | 送出時間 |  意見內容"
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
         _Band(0).Cols   =   6
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -74910
         Top             =   3570
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdAddAttDB 
         Caption         =   "新增歷程附件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -68550
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdRemAttDB 
         Caption         =   "刪除歷程附件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -67170
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1335
      End
      Begin MSForms.TextBox txtEEP10_2 
         Height          =   320
         Left            =   -73890
         TabIndex        =   106
         Top             =   2730
         Width           =   2990
         VariousPropertyBits=   671105051
         Size            =   "5274;564"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "副本收受者："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74940
         TabIndex        =   123
         Top             =   2790
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "內容："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   122
         Top             =   3030
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "流程狀態："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74940
         TabIndex        =   121
         Top             =   2220
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發  送  者："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   -74940
         TabIndex        =   120
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "收  受  者："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74940
         TabIndex        =   119
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "註：副本多人時以逗號(,)分隔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   -72480
         TabIndex        =   118
         Top             =   2550
         Width           =   2205
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "副本收受者："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74820
         TabIndex        =   117
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "受文者："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74820
         TabIndex        =   116
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主旨："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   -74820
         TabIndex        =   115
         Top             =   2250
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74820
         TabIndex        =   114
         Top             =   3030
         Width           =   540
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "代理人:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   -68610
         TabIndex        =   113
         Top             =   3870
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   14
         Left            =   -67890
         TabIndex        =   112
         Top             =   3870
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "待回的最後流程"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -72150
         TabIndex        =   111
         Top             =   1890
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "(註:雙擊選取時,下方顯示歷程資料)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   230
         Index           =   0
         Left            =   -74880
         TabIndex        =   110
         Top             =   300
         Width           =   2900
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
         Left            =   -70410
         TabIndex        =   109
         Top             =   90
         Visible         =   0   'False
         Width           =   830
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
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -69210
         TabIndex        =   108
         Top             =   90
         Visible         =   0   'False
         Width           =   830
      End
      Begin VB.Label lblCMboth 
         Caption         =   "lblCMboth"
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
         Left            =   -68280
         TabIndex        =   107
         Top             =   90
         Width           =   1310
      End
      Begin MSForms.TextBox txtEEP08 
         Height          =   1365
         Left            =   -74370
         TabIndex        =   105
         Top             =   3060
         Width           =   4245
         VariousPropertyBits=   -1466941413
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "7488;2408"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEEP03_2 
         Height          =   225
         Left            =   -73350
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1185
         VariousPropertyBits=   671105055
         Size            =   "2090;397"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   660
         Index           =   6
         Left            =   -73710
         TabIndex        =   103
         Top             =   1560
         Width           =   4830
         VariousPropertyBits=   -1466941413
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "8520;1164"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   750
         Index           =   5
         Left            =   -73710
         TabIndex        =   102
         Top             =   780
         Width           =   4830
         VariousPropertyBits=   -1466941413
         MaxLength       =   80
         ScrollBars      =   2
         Size            =   "8520;1323"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   720
         Index           =   0
         Left            =   -73710
         TabIndex        =   101
         Top             =   2250
         Width           =   4830
         VariousPropertyBits=   -1466941413
         MaxLength       =   300
         ScrollBars      =   2
         Size            =   "8520;1270"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   1290
         Index           =   4
         Left            =   -73710
         TabIndex        =   100
         Top             =   3000
         Width           =   4830
         VariousPropertyBits=   -1466941413
         MaxLength       =   400
         ScrollBars      =   2
         Size            =   "8520;2275"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox CboEEP05 
         Height          =   290
         Left            =   -74010
         TabIndex        =   99
         Top             =   2460
         Width           =   1490
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2628;512"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         Caption         =   $"frm090202_2.frx":0268
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   470
         Left            =   -71640
         TabIndex        =   98
         Top             =   3780
         Width           =   5330
      End
      Begin VB.Label LblinfoNote 
         Caption         =   "註：檔名則以P000000.info.pdf，多個以上資料檔則加序號（例：P000000.2.info.pdf，P000000.3.info.pdf）"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   230
         Left            =   -74820
         TabIndex        =   97
         Top             =   3480
         Width           =   8600
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "請款備註："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   60
         TabIndex        =   96
         Top             =   3570
         Width           =   900
      End
      Begin MSForms.TextBox txt3 
         Height          =   680
         Index           =   7
         Left            =   960
         TabIndex        =   49
         Top             =   3540
         Width           =   7920
         VariousPropertyBits=   -1466941413
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "13970;1199"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "\\typing2\外專送件\中說原始檔"
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
         Left            =   390
         TabIndex        =   95
         Top             =   4260
         Width           =   2470
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "(註: 按下路徑可開啟資料夾)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   5040
         TabIndex        =   94
         Top             =   4260
         Width           =   2440
      End
      Begin VB.Label Lbl926 
         Caption         =   "(一核 or 二核)"
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
         Left            =   -69510
         TabIndex        =   93
         Top             =   90
         Visible         =   0   'False
         Width           =   920
      End
      Begin MSForms.TextBox txtEP12 
         Height          =   800
         Left            =   960
         TabIndex        =   48
         Top             =   2730
         Width           =   7920
         VariousPropertyBits=   -1466941413
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "13970;1411"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "作業備註："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   60
         TabIndex        =   92
         Top             =   2760
         Width           =   900
      End
   End
   Begin VB.Frame Frame1Big 
      Caption         =   "附件區："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4540
      Left            =   0
      TabIndex        =   149
      Top             =   1410
      Visible         =   0   'False
      Width           =   8950
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
         TabIndex        =   160
         Top             =   -30
         Width           =   640
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全選"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   1740
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   4080
         Width           =   675
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   3990
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   4080
         Width           =   675
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "新增"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   3240
         TabIndex        =   157
         Top             =   4080
         Width           =   675
      End
      Begin VB.CommandButton cmdSaveAtt 
         Caption         =   "下載"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2490
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   4080
         Width           =   675
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   990
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   4080
         Width           =   675
      End
      Begin VB.ListBox lstAtt 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3108
         Index           =   2
         ItemData        =   "frm090202_2.frx":02F4
         Left            =   60
         List            =   "frm090202_2.frx":02FB
         MultiSelect     =   2  '進階多重選取
         Sorted          =   -1  'True
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   270
         Width           =   8790
      End
      Begin VB.Frame FrameFCPlink 
         Caption         =   "(註: 按下路徑可開啟資料夾)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   550
         Index           =   1
         Left            =   150
         TabIndex        =   151
         Top             =   3510
         Visible         =   0   'False
         Width           =   5650
         Begin VB.CommandButton CmdF21 
            BackColor       =   &H00FFC0C0&
            Caption         =   "上傳"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   1
            Left            =   1500
            Style           =   1  '圖片外觀
            TabIndex        =   163
            Top             =   180
            Visible         =   0   'False
            Width           =   530
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "\\typing2\外專送件"
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
            Index           =   1
            Left            =   3840
            TabIndex        =   153
            Top             =   240
            Width           =   1460
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "電子送件暫存區"
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
            Index           =   1
            Left            =   60
            TabIndex        =   152
            Top             =   240
            Width           =   1310
         End
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "註：歷程附件發文後一個月或取消收文後3個月刪除。（會修及會完非沿用上一道流程附件者，除外）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   3
      Left            =   0
      TabIndex        =   44
      Top             =   5970
      Width           =   8000
   End
   Begin VB.Label lblPA09 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4650
      TabIndex        =   32
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   4080
      TabIndex        =   31
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱(日)："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   0
      TabIndex        =   30
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱(英)："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   0
      TabIndex        =   29
      Top             =   780
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   0
      TabIndex        =   26
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblCP09 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   990
      TabIndex        =   25
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   2130
      TabIndex        =   24
      Top             =   240
      Width           =   930
   End
   Begin VB.Label lblCP10 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3090
      TabIndex        =   23
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   19
      Left            =   0
      TabIndex        =   22
      Top             =   30
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱(中)："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   18
      Left            =   0
      TabIndex        =   21
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label lblCaseNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   990
      TabIndex        =   20
      Top             =   30
      Width           =   1590
   End
   Begin VB.Label lblPA08 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4650
      TabIndex        =   19
      Top             =   30
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "種類："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   3690
      TabIndex        =   18
      Top             =   30
      Width           =   930
   End
   Begin MSForms.TextBox txtCaseName 
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   1110
      Width           =   5000
      VariousPropertyBits=   671105051
      Size            =   "8811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseName 
      Height          =   300
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   780
      Width           =   5000
      VariousPropertyBits=   671105051
      Size            =   "8811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseName 
      Height          =   300
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   450
      Width           =   4760
      VariousPropertyBits=   671105051
      Size            =   "8396;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm090202_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/12 Form2.0已修改
'商標歷程上線時間:107/9/17(一)
'專利處承辦電子簽核系統上線:1020902-02
'Create by Sindy 2013/4/15
Option Explicit

'變數宣告區
Dim nFrm As Form
Public m_EEP01 As String '總收文號
Dim m_EEP02 As String '序號
Public intReceiveKind As Integer '0.承辦人工作進度 1.待核判區 2.待會稿區 3.繪圖人員工作進度 99.個案聯絡
'                                 5.程序送判
'***********************************************************************
Dim m_bolSendChWrite As Boolean 'True=屬送中說 (翻譯完稿輸入\案件進度檔案維護)
'*** 員編+' '+姓名 ***
Public m_SPMan As String '智權人員
Public m_EPMan As String '承辦人
Public m_CP14_2 As String 'Add By Sindy 2017/6/15 外翻人員在所內處理的人員可能一個以上,第2個以上為副本收受者
Public m_DPMan As String '繪圖人員
Public m_EMMan As String '英文核稿人
Public m_CMMan As String '核稿主管
Public m_DCMan As String '草圖核稿人 Add By Sindy 2015/4/22
Public m_DMMan As String '繪圖主管：北所.72006張瓊玉、中所.82018李月嬌、其他.78007劉大愛。（以繪圖人員的所別決定其主管）
Public m_CSMan As String '判發主管
Public m_NPMan As String '程序人員
Public m_F21CMMan As String 'FCP工程師主管 Add By Sindy 2023/10/2
Dim m_TCT10Man As String '命名工程師 Add By Sindy 2024/1/5
'*** END
Dim PField(1 To 4) As String
'-----------------------------------------------------------------------
Public bolPAFlow As Boolean, bolTMFlow As Boolean 'Add By Sindy 2018/4/18
Public bolOtherFlow As Boolean 'Add By Sindy 2021/9/3
Public bolFCPFlow As Boolean 'Add By Sindy 2023/9/12
Public bolCFTFlow As Boolean, bolFCTFlow As Boolean 'Add By Sindy 2024/7/11
'-----------------------------------------------------------------------
Dim bolFMP As Boolean 'Add By Sindy 2023/10/6
Dim bolOurFMP As Boolean '是否寰華案件 Add By Sindy 2023/10/6
Dim pa() As String, sp() As String, tm() As String, cp() As String
Dim lC() As String, hc() As String 'Add By Sindy 2021/9/3
Dim m_strSys As String
'Add By Sindy 2023/9/12
Const intTab_承辦單 As Integer = 1
Const intTab_外專承辦單 As Integer = 2
Const intTab_存卷資料 As Integer = 3
'2023/9/12 END
Const 屬中說的案件性質 As String = "201,209,235" 'Add By Sindy 2024/1/5
'***********************************************************************

Public ShowNextData As Boolean 'Add By Sindy 2013/9/3
Public cmdState As Integer, bolQuery As Boolean '紀錄作用按鍵
Public m_FlowUserNum As String 'Add By Sindy 2013/9/12 案件流程所屬人員
Public m_CurrFlowEEP02 As Integer 'Add By Sindy 2016/3/25 傳入待會稿區,待核判區目前要處理的歷程序號

Public m_RetrunRecv As String 'Add By Sindy 2017/9/1 回傳總收文號
Dim m_RetrunRecvCnt As Integer 'Add By Sindy 2018/9/27 總收文號數量
Dim m_EEP15 As String, m_EEP11 As String 'Add by Sindy 2020/9/29
Public m_RetrunRecvSub As String 'Add by Sindy 2020/9/29
Dim bolManyCaseToMix As Boolean, m_RetrunRecvToMix As String 'Add by Sindy 2020/10/19
Public m_PrevForm As Form '前一畫面

Public m_EditMode As Integer
Dim strSubject As String, strContent As String
Dim ii As Integer, jj As Integer
Public dblPrevRow As Double
Dim m_EP41 As String  '核稿語文 1.英2.日 Add By Sindy 2015/3/16
Dim bolBCaseFlow As Boolean '是否為附加流程案件
Dim bolP020NewCase As Boolean '是否為P大陸新案 Add By Sindy 2014/9/5
Dim bol00EngCMFlow As Boolean 'Add By Sindy 2014/1/10 有聯絡送英核流程
Dim bol00EngCMFlowEmp As String 'Add By Sindy 2017/11/30 聯絡-送英核流程的發送者
Dim m_EEP11Person As String '記錄原收受者
Dim intLastEEP02 As Integer, strLastEEP03 As String, m_strLastEEP04 As String, strLastEEP11 As String, m_strLastEEP04Nm As String
Dim strLastEEP05 As String 'Add By Sindy 2013/9/12
Dim bolLastFile As Boolean '有無附件
Dim m_EP01 As String '當月目次
Dim strEEP10_Err As String, strEEP05_Err As String
Dim m_Country As String, m_SaleArea As String, m_CP10Nm As String
Dim m_PA08 As String
Dim bolSave As Boolean
Dim bolMoveFile As Boolean 'True : 將最近一道附件移至此流程,並且將相關流程中的附件一併刪除
Dim bolDeleteFile As Boolean 'True : 將上一筆流程的附件刪除
Dim m_PreviousFlow As String, m_FlowTxt As String
'Dim m_ChinaPCase As Boolean '品薇的P-大陸案
Dim m_EP08 As String, m_EP38 As String 'Add By Sindy 2013/9/11
Dim m_CPM27 As String 'Add By Sindy 2013/9/17 核稿人不可判發
Dim m_CPM26 As String 'Add By Sindy 2013/9/23 專業部電子檔副檔名
Dim m_EP06 As String '文件齊備日
Dim m_EP09 As String '完稿日 Add By Sindy 2023/9/19
Dim m_EP34 As String '是否會稿
Dim m_EP39 As String '核稿完成日
Dim m_EP42 As String '判發完成日 Add By Sindy 2018/4/27
Dim m_EP33 As String '英文核完日
Dim m_EP18 As String '墨圖完稿日 Add By Sindy 2018/6/25
Dim m_EEP12 As String 'Add By Sindy 2013/10/16 代理註明
Dim m_EEP16 As String 'Add By Sindy 2023/12/18 原收受者
Dim bolWaitReply As Boolean
Dim m_UpdEEP11 As String 'Add By Sindy 2013/11/13 系統備註
Dim m_PA48 As String 'Add By Sindy 2014/3/4 客戶案件案號
Dim m_EEP14 As String 'Add By Sindy 2018/8/29

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
Dim strSendFilePath As String

'列印宣告區
'Private Declare Function ShellExecute Lib _
'"shell32.dll" Alias "ShellExecuteA" ( _
'ByVal hwnd As Long, ByVal lpOperation As String, _
'ByVal lpFile As String, ByVal lpParameters As String, _
'ByVal lpDirectory As String, _
'ByVal nShowCmd As Long) As Long
'Private Const SW_HIDE = 0

Dim m_PER04 As String, bolHadSetProofEngReader As Boolean 'Add By Sindy 2015/3/4
Dim bolMultinationalEngOk As Boolean 'Add By Sindy 2015/3/20
Dim m_PA26 As String, m_PA27 As String, m_PA28 As String, m_PA29 As String, m_PA30 As String
Dim m_EP07 As String 'Add By Sindy 2015/12/2
Dim m_PA75 As String 'Add By Sindy 2018/4/2 FC代理人
Dim m_PA77 As String 'Add By Sindy 2018/4/2 彼所案號
Public ChkTG As Boolean 'Add By Sindy 2018/4/18
Dim m_SubjectNote As String 'Add By Sindy 2018/8/13
Dim m_ManyAppl As String 'Add By Sindy 2018/9/26 記錄多申請人
Dim m_ManyApplCP56 As String 'Add By Sindy 2023/9/5 記錄多申請人_受讓申請人
Dim m_strSaveCaseNo1 As String, m_strSaveCaseNo2 As String, m_strSaveCaseNo3 As String, m_strSaveCaseNo4 As String 'Add By Sindy 2018/10/5
Dim bolMCTFcase As Boolean 'Add By Sindy 2019/4/15
Dim m_intDataPDF As Integer 'Add By Sindy 2020/3/5
Dim m_dblCP79 As Double 'Add By Sindy 2021/12/13
Public m_strSpecState As String '特殊情況 Add By Sindy 2022/4/26 ex:尚待收款-完稿日
'Dim intlstAtt0_Height As Integer 'Add By Sindy 2024/1/4
'Add By Sindy 2023/9/12
Dim m_strColName() As String, m_strColText() As String, m_intColCnt As Integer
Dim m_PA162 As String '是否加註核准分割建議
'2023/9/12 END
'Add By Sindy 2024/1/2
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm_IR As Form
'2024/1/2 END
Dim m_EP12 As String 'Add By Sindy 2024/1/23
Dim dblErrNumber As Double, strErrText As String 'Add By Sindy 2024/5/21
Dim strPP04 As String, strPP05 As String
'Dim strCompName As String, strRepName As String 'Add By Sindy 2025/10/27
Dim bolCmdF21 As Boolean 'Add By Sindy 2025/11/4
Dim hLocalFile As Long
Dim m_EEP04New As String 'Add By Sindy 2025/11/7


'Add By Sindy 2014/1/15
Private Sub CboCP10_Click()
   'Add By Sindy 2018/8/29
   If Label11.Caption = "會稿方式：" Then
      If cmdSend.Enabled = True Then
         If Trim(Left(CboCP10.Text, 1)) = "1" Then
            cmdSend.Caption = "E-Mail"
            
            'Add By Sindy 2018/10/16 可操作多案
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
               Call cmdManyCase_Click
            End If
            '2018/10/16 END
            
            Call AskEmpIsCopyFile 'Add By Sindy 2018/10/11 詢問是否要沿用附件
         Else
            cmdSend.Caption = "送出(&O)"
         End If
      End If
   Else
   '2018/8/29 END
      If Left(CboEEP04.Text, 2) = EMP_附加流程 Then
         If bolTMFlow = True Then
            txtEEP10_2.Enabled = True
            Frame1.Enabled = True
            'Add By Sindy 2019/7/9 T非台灣案,代理人撰稿不可使用附件區,因無欲新增歷程
            If Trim(Left(CboCP10.Text, 4)) = "734" Then
               CboEEP05.Enabled = False 'Add By Sindy 2025/7/31
               CboEEP05.Text = ""
               m_EEP11Person = ""
               CboEEP04.Tag = ""
               lstAtt(0).Clear
               txtEEP10_2.Enabled = False: txtEEP10_2.Text = ""
               Frame1.Enabled = False
            End If
         'Add By Sindy 2025/7/31 +Or bolFCPFlow = True
         ElseIf bolPAFlow = True Or bolFCPFlow = True Then
            If Trim(Left(CboCP10.Text, 4)) = "936" Or _
               Trim(Left(CboCP10.Text, 4)) = "957" Or _
               Trim(Left(CboCP10.Text, 4)) = "958" Then
               CboEEP05.Enabled = False 'Add By Sindy 2025/7/31
               CboEEP05.Text = ""
               m_EEP11Person = ""
               CboEEP04.Tag = ""
            End If
         End If
      End If
   End If
End Sub

Private Sub CboEEP05_Click()
   'Add By Sindy 2013/9/4
   'Modify By Sindy 2018/10/22 S單位也不能夾帶附件了
   If Left(CboEEP04.Text, 2) = EMP_聯絡 Then
      If Trim(CboEEP05.Text) = Trim(m_SPMan) Then
'         ChkEMail.Visible = True
'         ChkEMail.Value = 1
         'Modify By Sindy 2018/9/13 修改只有S單位可以寄附件
         'Modify By Sindy 2018/10/23 已增加客戶會稿功能,且商標案不複雜因此商標案不提供此功能
         'Modify by Sindy 2019/5/13 聯絡,智權人員,專利處就可夾帶附件(摩根,禧佩,薛經理)
         'If Left(PUB_GetStaffST15(Trim(Left(m_SPMan, 6)), "1"), 1) = "S" And bolPAFlow = True Then
         If bolPAFlow = True Then
         '2019/5/13 END
            ChkEMail.Visible = True
            ChkEMail.Value = 1
         Else
            ChkEMail.Visible = False
            ChkEMail.Value = 0
         End If
         '2018/9/13 END
      Else
         ChkEMail.Visible = False
         ChkEMail.Value = 0
      End If
   End If
   '2013/9/4 END
      
   'Add By Sindy 2013/11/13 當聯絡流程,發送者為送英核主管收受者為工程師時,顯示一併更新英文核完日的勾選欄位
   Check1.Visible = False
   If bol00EngCMFlow = True Then 'Modify By Sindy 2017/1/3 應該要有工程師發的聯絡-送英(日)核才有此欄位可勾選 ex:CFP-029099
      'Modify By Sindy 2017/11/30 增加判斷”或”是否為(聯絡-送英核)的發送者
      If Left(CboEEP04.Text, 2) = EMP_聯絡 And txtEEP03 = Left(Trim(m_EMMan), 5) And _
         (Trim(Left(CboEEP05.Text, 6)) = Left(Trim(m_EPMan), 5) Or Trim(Left(CboEEP05.Text, 6)) = bol00EngCMFlowEmp) And _
         Val(m_EP33) = 0 Then
         Check1.Visible = True
      End If
   End If
   '2013/11/13 END
   
   Call SettxtEEP10_2 'Add By Sindy 2017/6/15
End Sub

Private Sub ChkEED13_Click()
   If Left(CboEEP04.Text, 2) = EMP_轉檔完成 Then
      If ChkEED13.Value = 1 Then
         CboEEP05.Text = m_NPMan '程序人員
      Else
         CboEEP05.Text = m_EPMan '承辦人
      End If
   End If
End Sub

Private Sub ChkEMail_Click()
'   If ChkEMail.Value = 1 And lstAtt(0).ListCount = 0 Then
'      MsgBox "附件區無檔案，不可勾選E-Mail夾帶附件！"
'      ChkEMail.Value = 0
'   End If
End Sub

'存承辦單內容-內專
Private Function funSaveEmpPaperData() As Boolean
Dim bolIsInsert As Boolean
   
On Error GoTo ErrHand
   
   funSaveEmpPaperData = False
   '檢查條件
   If Trim(txt1(5).Text) = "" Then
      MsgBox "承辦單受文者不可空白！", vbExclamation
      Me.txt1(5).SetFocus
      Me.SSTab1.Tab = intTab_承辦單
      Exit Function
   End If
   If Trim(txt1(0).Text) = "" Then
      MsgBox "承辦單主旨不可空白！", vbExclamation
      Me.txt1(0).SetFocus
      Me.SSTab1.Tab = intTab_承辦單
      Exit Function
   End If
   If Len(Trim(txt1(0).Text)) > txt1(0).MaxLength Then
      MsgBox "承辦單主旨長度太長！" & vbCrLf & _
             "（只可輸入中文" & txt1(0).MaxLength & "個字）", vbExclamation
      Me.txt1(0).SetFocus
      Me.SSTab1.Tab = intTab_承辦單
      Exit Function
   End If
   
   Screen.MousePointer = vbHourglass
   
   '檢查承辦單內容是否存在
   strSql = "select * From EmpElectronData where eed01='" & m_EEP01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   bolIsInsert = True
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         bolIsInsert = False
      End If
   End If
   
On Error GoTo ErrConn

   cnnConnection.BeginTrans
   If bolIsInsert = True Then
      strSql = "insert into EmpElectronData(EED01,EED02,EED03,EED04,EED05) values(" & _
               CNULL(m_EEP01) & "," & CNULL(txt1(5)) & "," & CNULL(txt1(6)) & "," & _
               CNULL(ChgSQL(txt1(0))) & "," & CNULL(Trim(txt1(4))) & ")"
   Else
      strSql = "update EmpElectronData set" & _
                     " EED02=" & CNULL(txt1(5)) & _
                     ",EED03=" & CNULL(txt1(6)) & _
                     ",EED04=" & CNULL(ChgSQL(txt1(0))) & _
                     ",EED05=" & CNULL(Trim(txt1(4))) & _
               " where EED01='" & m_EEP01 & "'"
   End If
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   Screen.MousePointer = vbDefault
   
   funSaveEmpPaperData = True
   Exit Function
   
ErrConn:
   cnnConnection.RollbackTrans
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox " 存檔失敗！" & vbCrLf & Err.Description
End Function

'Add By Sindy 2023/9/20
'存承辦單內容-外專
'Modify By Sindy 2024/1/23 改取名為作業備註
'Modify By Sindy 2025/4/22 +bolAutoSave As Boolean: 是否人員自行點選要存檔
Private Function funSaveEmpPaperData_FCP(bolAutoSave As Boolean) As Boolean
Dim bolIsInsert As Boolean
Dim bolChkOk As Boolean
Dim i As Integer
Dim Cancel As Boolean
   
On Error GoTo ErrHand
   
   funSaveEmpPaperData_FCP = False
   
   'Add By Sindy 2025/10/15
   If Frame945.Visible = True Then
      If txtEED14.Visible = True And txtEED14.Enabled = True Then
         Cancel = False
         Call txtEED14_Validate(Cancel)
         If Cancel = True Then
            SSTab1.Tab = intTab_外專承辦單
            txtEED14.SetFocus
            Exit Function
         End If
      End If
      If txtEED15.Visible = True And txtEED15.Enabled = True Then
         Cancel = False
         Call txtEED15_Validate(Cancel)
         If Cancel = True Then
            SSTab1.Tab = intTab_外專承辦單
            txtEED15.SetFocus
            Exit Function
         End If
      End If
   End If
   '2025/10/15 END
   
   '檢查條件
'   If Trim(txt3(1).Text) = "" Then
'      MsgBox "承辦單受文者不可空白！", vbExclamation
'      Me.txt3(1).SetFocus
'      Me.SSTab1.Tab = intTab_外專承辦單
'      Exit Function
'   End If
'   If Trim(txt3(2).Text) = "" Then
'      MsgBox "承辦單主旨不可空白！", vbExclamation
'      Me.txt3(2).SetFocus
'      Me.SSTab1.Tab = intTab_外專承辦單
'      Exit Function
'   End If
'   If Len(Trim(txt3(2).Text)) > txt3(2).MaxLength Then
'      MsgBox "承辦單主旨長度太長！" & vbCrLf & _
'             "（只可輸入中文" & txt3(2).MaxLength & "個字）", vbExclamation
'      Me.txt3(2).SetFocus
'      Me.SSTab1.Tab = intTab_外專承辦單
'      Exit Function
'   End If
   
   'Modify By Sindy 2024/1/23 至少要有一個欄位有值
   bolChkOk = False
   For i = 3 To 8
      If txt3(i) <> "" Then
         bolChkOk = True
         Exit For
      End If
   Next i
   'Add By Sindy 2025/4/22
   If ChkEED08.Visible = True And ChkEED08.Value = 1 Then bolChkOk = True
   If ChkEED13.Visible = True And ChkEED13.Value = 1 Then bolChkOk = True
   If Frame945.Tag = "945" Then
      If txtEED14 <> "" Or txtEED15 <> "" Then
         bolChkOk = True
         '約定期限不可大於本所期限
         If Val(Me.txtEED14.Text) >= Val(Me.txtEED15.Text) Then
            MsgBox "約定期限不可大於本所期限！", vbExclamation
            txtEED14.SetFocus
            Me.SSTab1.Tab = intTab_外專承辦單
            Exit Function
         End If
      End If
   End If
   '2025/4/22 END
   'Add By Sindy 2025/8/20
   If Frame945.Tag <> "" And txtEED14 <> "" Then
      bolChkOk = True
   End If
   '2025/8/20 END
   If bolChkOk = False Then
      'Add By Sindy 2025/4/22 各欄位均無值的話,若有資料列可以整筆刪除
      strSql = "select * From EmpElectronData where eed01='" & m_EEP01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      bolIsInsert = True
      If intI = 1 Then
         strSql = "delete From EmpElectronData where eed01='" & m_EEP01 & "'"
         cnnConnection.Execute strSql, intI
      Else
         If bolAutoSave = False Then
            MsgBox "至少要有一個欄位有值！", vbExclamation
            Me.SSTab1.Tab = intTab_外專承辦單
            Exit Function
         End If
      End If
      '2025/4/22 END
   End If
   '2024/1/23 END
   
   Screen.MousePointer = vbHourglass
   
On Error GoTo ErrConn
   
   cnnConnection.BeginTrans
   
   If bolChkOk = True Then
      '檢查承辦單內容是否存在
      strSql = "select * From EmpElectronData where eed01='" & m_EEP01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      bolIsInsert = True
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            bolIsInsert = False
         End If
      End If
      If bolIsInsert = True Then
         'Modify By Sindy 2025/4/7 +EED08,EED14,EED15
         strSql = "insert into EmpElectronData(EED01,EED05," & _
                                              "EED06,EED09,EED10,EED11," & _
                                              "EED12,EED13,EED08," & _
                                              "EED14,EED15) values(" & _
                  CNULL(m_EEP01) & "," & CNULL(ChgSQL(txt3(4))) & "," & _
                  CNULL(Trim(txt3(6))) & "," & CNULL(Trim(txt3(5))) & "," & CNULL(Trim(txt3(3))) & "," & CNULL(Trim(txt3(7))) & _
                  "," & CNULL(Trim(txt3(8))) & ",'" & IIf(ChkEED13.Value = 1, "Y", "") & "','" & IIf(ChkEED08.Value = 1, "Y", "") & "'" & _
                  ",'" & IIf(Frame945.Tag <> "" And txtEED14.Visible = True, DBDATE(txtEED14.Text), "") & "'" & _
                  ",'" & IIf(Frame945.Tag <> "" And txtEED15.Visible = True, DBDATE(txtEED15.Text), "") & "'" & _
                  ")"
      Else
         'Modify By Sindy 2025/4/7 +EED08,EED14,EED15
         strSql = "update EmpElectronData set" & _
                        " EED05=" & CNULL(ChgSQL(Trim(txt3(4)))) & _
                        ",EED06=" & CNULL(Trim(txt3(6))) & _
                        ",EED09=" & CNULL(Trim(txt3(5))) & _
                        ",EED10=" & CNULL(Trim(txt3(3))) & _
                        ",EED11=" & CNULL(Trim(txt3(7))) & _
                        ",EED12=" & CNULL(Trim(txt3(8))) & _
                        ",EED13='" & IIf(ChkEED13.Value = 1, "Y", "") & "'" & _
                        ",EED08='" & IIf(ChkEED08.Value = 1, "Y", "") & "'" & _
                        ",EED14='" & IIf(Frame945.Tag <> "" And txtEED14.Visible = True, DBDATE(txtEED14.Text), "") & "'" & _
                        ",EED15='" & IIf(Frame945.Tag <> "" And txtEED15.Visible = True, DBDATE(txtEED15.Text), "") & "'" & _
                  " where EED01='" & m_EEP01 & "'"
      End If
      cnnConnection.Execute strSql, intI
   End If
   
   'Add By Sindy 2024/1/24 更新承辦備註(作業備註)
   If txtEP12.Enabled = True And txtEP12.Locked = False And txtEP12.Tag <> txtEP12.Text Then
      strSql = "update engineerprogress set ep12='" & txtEP12.Text & "' where ep02='" & m_EEP01 & "'"
      Pub_SeekTbLog strSql 'Add By Sindy 2024/3/29
      cnnConnection.Execute strSql
      If intReceiveKind = 0 Then '0.承辦人工作進度
         m_PrevForm.txtEP12 = txtEP12.Text
      End If
   End If
   '2024/1/24 END
   
   cnnConnection.CommitTrans
   Screen.MousePointer = vbDefault
   
   funSaveEmpPaperData_FCP = True
   Exit Function
   
ErrConn:
   cnnConnection.RollbackTrans
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox " 歷程備註存檔失敗！" & vbCrLf & Err.Description
End Function

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2015/6/15 +ac03
   'Modify By Sindy 2016/3/9 +待回eep09
   'Modify By Sindy 2018/8/29 +EEP14:會稿方式
   '                        0       1        2         3        4           5        6         7           8             9           10       11         12      13      14      15       16
   arrGridHeadText = Array("順序", "EEP03", "發送者", "EEP04", "流程狀態", "EEP05", "收受者", "送出時間", "副本收受者", "意見內容", "EEP10", "c1.CP43", "AC03", "待回", "顯示", "EEP14", "EEP11")
   arrGridHeadWidth = Array(400, 0, 950, 0, 800, 0, 700, 1300, 1000, 3300, 0, 0, 0, 400, 400, 0, 0)
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
Dim strSql As String
Dim strRefEEP02 As String
Dim PrinterIndex As Integer, i As Integer 'Add By Sindy 2018/7/16
Dim IntTemp1 As Long, IntTemp2 As Long
Dim strEEP04is39_EEP11 As String, strEEP05 As String 'Add By Sindy 2023/12/18
Dim bolSpecCase As Boolean, m_LimitType As String, m_RecvNo As String, strMsgTxt As String    'Added by Lydia 2025/11/17
   
   'Add By Sindy 2023/9/19
   If InStr(UCase(App.EXEName), "FILE") > 0 Then
      Me.Caption = Me.Caption & "（排版區）"
'   ElseIf intReceiveKind = 4 Then
'      Me.Caption = Me.Caption & "（送中說）"
   '2023/9/19 END
   ElseIf intReceiveKind = 0 Then
      Me.Caption = Me.Caption & "（承辦）"
   ElseIf intReceiveKind = 1 Then
      Me.Caption = Me.Caption & "（核判區）"
   ElseIf intReceiveKind = 2 Then
      Me.Caption = Me.Caption & "（會稿區）"
   ElseIf intReceiveKind = 3 Then
      Me.Caption = Me.Caption & "（繪圖）"
   End If
   'Add By Sindy 2024/1/2
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2024/1/2 END
   
   QueryData = True
   bolCmdF21 = False 'Add By Sindy 2025/11/4
   FrameFCPlink(0).Visible = False 'Add By Sindy 2023/11/22
   '清空及預設欄位值
   GRD1.Clear
   SetGrd
   lblCaseNo.Caption = Empty
   lblPA08.Caption = Empty
   'Add By Sindy 2013/10/1
   txtCaseName(0).Text = Empty
   txtCaseName(0).Tag = Empty
   txtCaseName(1).Text = Empty
   txtCaseName(1).Tag = Empty
   txtCaseName(2).Text = Empty
   txtCaseName(2).Tag = Empty
   '2013/10/1 END
   m_PA08 = Empty
   Call ClearData
   Call SetCtrlReadOnly(False)
   m_PA26 = "": m_PA27 = "": m_PA28 = "": m_PA29 = "": m_PA30 = ""
   m_PA75 = "" 'Add By Sindy 2018/4/2 FC代理人
   m_PA77 = "" 'Add By Sindy 2018/4/2 彼所案號
   m_dblCP79 = 0
   m_EditMode = 0
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   bolSave = False
   cmdOK(5).Visible = False '原始檔區 Add By Sindy 2023/11/16
   
   'Modify By Sindy 2018/4/26
   '進度檔
   cp(9) = m_EEP01
   'Modified by Morgan 2024//7/22
   'Call PUB_ReadCaseProgressDatabase(cp(), 國外_CF)
   Call PUB_ReadCaseProgressDatabase(cp(), 國外_CF, False)
   'end 2024/7/22
   '2018/4/26 END
   'Add By Sindy 2023/9/5 記錄多申請人_受讓申請人
   m_ManyApplCP56 = ""
   If cp(56) <> "" Then
      cp(56) = ChangeCustomerL(cp(56))
      m_ManyApplCP56 = m_ManyApplCP56 & "," & cp(56)
   End If
   If cp(89) <> "" Then
      cp(89) = ChangeCustomerL(cp(89))
      m_ManyApplCP56 = m_ManyApplCP56 & "," & cp(89)
   End If
   If cp(90) <> "" Then
      cp(90) = ChangeCustomerL(cp(90))
      m_ManyApplCP56 = m_ManyApplCP56 & "," & cp(90)
   End If
   If cp(91) <> "" Then
      cp(91) = ChangeCustomerL(cp(91))
      m_ManyApplCP56 = m_ManyApplCP56 & "," & cp(91)
   End If
   If cp(92) <> "" Then
      cp(92) = ChangeCustomerL(cp(92))
      m_ManyApplCP56 = m_ManyApplCP56 & "," & cp(92)
   End If
   '2020/6/8 END
   If m_ManyApplCP56 <> "" Then m_ManyApplCP56 = Mid(m_ManyApplCP56, 2)
   '2023/9/5 END
   'Added by Lydia 2025/11/17 因區塊2之部分案件性質也有走承辦歷程，故請協助將承辦歷程一併納入權限管制範圍。---- from 教威
   bolSpecCase = PUB_ChkCPPAndCPFLimits_Spec(cp(1), cp(2), cp(3), cp(4), m_LimitType, m_RecvNo, strMsgTxt)
   If bolSpecCase = True Then
      If m_LimitType = "" Then
         '隱藏所有附件區
         Frame1.Visible = False
         Frame5.Visible = False
      End If
   End If
   'end 2025/11/17
   
   '案件資料
   'Modify By Sindy 2014/3/4 +pa48
   'Modify By Sindy 2015/1/21 +CP140
   'Modify By Sindy 2015/3/16 +EP41
   'Modify By Sindy 2024/1/23 +,EP12
   'Modify By Sindy 2025/11/7 +,GetEEPCurState(cp09) as EEP04New
   strSql = "Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,NVL(PA05,NVL(PA06,PA07)) as 案件名稱," & _
            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員,CP27,CP57,CP10,PA09," & _
            "s2.ST06 as Area,CP118,CP09,PA08,EP01,EP06,EP08,EP38,CPM26,CPM27,CP06,CP07,CP12,CP13,CP22,CP44,CP45,CP110,CP116," & _
            "PA05,PA06,PA07,CP43,pa48,cp140,EP41,PA75,PA77,EP12,GetEEPCurState(cp09) as EEP04New" & _
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
   strSql = strSql & " union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,NVL(SP05,NVL(SP06,SP07)) as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員,CP27,CP57,CP10,SP09," & _
            "s2.ST06 as Area,CP118,CP09,'' as PA08,EP01,EP06,EP08,EP38,CPM26,CPM27,CP06,CP07,CP12,CP13,CP22,CP44,CP45,CP110,CP116," & _
            "SP05,SP06,SP07,CP43,SP29,cp140,EP41,SP26,SP27,EP12,GetEEPCurState(cp09) as EEP04New" & _
            " From CaseProgress,EngineerProgress,Servicepractice," & _
            "staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And SP09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)"
   'Add By Sindy 2018/4/18 +商標
   strSql = strSql & " union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,NVL(TM05,NVL(TM06,TM07)) as 案件名稱," & _
            "NA03 as 國家,Decode(TM10,'000',PTM03,PTM04) as 種類,Decode(TM10,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員,CP27,CP57,CP10,TM10," & _
            "s2.ST06 as Area,CP118,CP09,TM08,EP01,EP06,EP08,EP38,CPM26,CPM27,CP06,CP07,CP12,CP13,CP22,CP44,CP45,CP110,CP116," & _
            "TM05,TM06,TM07,CP43,TM35,cp140,EP41,TM44,TM45,EP12,GetEEPCurState(cp09) as EEP04New" & _
            " From CaseProgress,EngineerProgress,Trademark," & _
            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And TM10=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '2'=PTM01(+) AND TM08=PTM02(+)"
   'Add By Sindy 2021/9/3
   '法務
   strSql = strSql & " union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,NVL(LC05,NVL(LC06,LC07)) as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(LC15,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員,CP27,CP57,CP10,LC15," & _
            "s2.ST06 as Area,CP118,CP09,'' as PA08,EP01,EP06,EP08,EP38,CPM26,CPM27,CP06,CP07,CP12,CP13,CP22,CP44,CP45,CP110,CP116," & _
            "LC05,LC06,LC07,CP43,LC17,cp140,EP41,LC22,LC23,EP12,GetEEPCurState(cp09) as EEP04New" & _
            " From CaseProgress,EngineerProgress,LawCase," & _
            "staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And LC15=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)"
   '顧問
   strSql = strSql & " union Select SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,HC06 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode('000','000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員,CP27,CP57,CP10,'000'," & _
            "s2.ST06 as Area,CP118,CP09,'' as PA08,EP01,EP06,EP08,EP38,CPM26,CPM27,CP06,CP07,CP12,CP13,CP22,CP44,CP45,CP110,CP116," & _
            "HC06,'','',CP43,'',cp140,EP41,'','',EP12,GetEEPCurState(cp09) as EEP04New" & _
            " From CaseProgress,EngineerProgress,HireCase," & _
            "staff s1,staff s2,nation,CasePropertyMap" & _
            " Where CP09='" & m_EEP01 & "'" & _
            " And CP09=EP02(+)" & _
            " And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And '000'=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)"
   '2021/9/3 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_EEP04New = "" & rsTmp.Fields("EEP04New") '目前歷程狀況 Add By Sindy 2025/11/7
      If Not IsNull(rsTmp.Fields("本所案號")) Then lblCaseNo.Caption = rsTmp.Fields("本所案號")
      PField(1) = SystemNumber(Me.lblCaseNo.Caption, 1)
      PField(2) = SystemNumber(Me.lblCaseNo.Caption, 2)
      PField(3) = SystemNumber(Me.lblCaseNo.Caption, 3)
      PField(4) = SystemNumber(Me.lblCaseNo.Caption, 4)
      lblCP09 = "" & rsTmp.Fields("CP09") 'Add By Sindy 2013/9/2
      m_EP12 = "" & rsTmp.Fields("EP12") 'Add By Sindy 2024/1/23
      
      'Modify By Sindy 2018/9/25 系統自動補填案號使用
      m_strSaveCaseNo1 = Trim(PField(1)) & CStr(Val(PField(2))) & IIf(PField(3) <> "0" Or PField(4) <> "00", "-" & PField(3), "") & IIf(PField(4) <> "00", "-" & PField(4), "")
      m_strSaveCaseNo2 = Trim(PField(1)) & "-" & CStr(Val(PField(2))) & IIf(PField(3) <> "0" Or PField(4) <> "00", "-" & PField(3), "") & IIf(PField(4) <> "00", "-" & PField(4), "")
      m_strSaveCaseNo3 = Trim(PField(1)) & CStr(PField(2)) & IIf(PField(3) <> "0" Or PField(4) <> "00", "-" & PField(3), "") & IIf(PField(4) <> "00", "-" & PField(4), "")
      m_strSaveCaseNo4 = Trim(PField(1)) & "-" & CStr(PField(2)) & IIf(PField(3) <> "0" Or PField(4) <> "00", "-" & PField(3), "") & IIf(PField(4) <> "00", "-" & PField(4), "")
      
      'Modify By Sindy 2018/4/18
      '基本檔
      m_strSys = CheckSys(PField(1))
      If InStr("1", m_strSys) > 0 Then '專利檔
         pa(1) = PField(1)
         pa(2) = PField(2)
         pa(3) = PField(3)
         pa(4) = PField(4)
         If ClsPDReadPatentDatabase(pa(), 國外_CF) = True Then
            bolPAFlow = True
         End If
         'Add By Sindy 2018/9/26 記錄多申請人
         'Modify By Sindy 2020/6/8 讀取長度足9碼,後面讀資料較方便
         If pa(75) <> "" Then pa(75) = ChangeCustomerL(pa(75))
         If pa(76) <> "" Then pa(76) = ChangeCustomerL(pa(76))
         If pa(86) <> "" Then pa(86) = ChangeCustomerL(pa(86))
         If pa(88) <> "" Then pa(88) = ChangeCustomerL(pa(88))
         If pa(26) <> "" Then
            pa(26) = ChangeCustomerL(pa(26))
            m_ManyAppl = m_ManyAppl & "," & pa(26)
         End If
         If pa(27) <> "" Then
            pa(27) = ChangeCustomerL(pa(27))
            m_ManyAppl = m_ManyAppl & "," & pa(27)
         End If
         If pa(28) <> "" Then
            pa(28) = ChangeCustomerL(pa(28))
            m_ManyAppl = m_ManyAppl & "," & pa(28)
         End If
         If pa(29) <> "" Then
            pa(29) = ChangeCustomerL(pa(29))
            m_ManyAppl = m_ManyAppl & "," & pa(29)
         End If
         If pa(30) <> "" Then
            pa(30) = ChangeCustomerL(pa(30))
            m_ManyAppl = m_ManyAppl & "," & pa(30)
         End If
         '2020/6/8 END
         If m_ManyAppl <> "" Then m_ManyAppl = Mid(m_ManyAppl, 2)
         '2018/9/26 END
         
         '****** 欄位值 ******
         m_PA26 = pa(26)
         m_PA27 = pa(27)
         m_PA28 = pa(28)
         m_PA29 = pa(29)
         m_PA30 = pa(30)
      ElseIf InStr("2", m_strSys) > 0 Then '商標檔
         tm(1) = PField(1)
         tm(2) = PField(2)
         tm(3) = PField(3)
         tm(4) = PField(4)
         If ClsPDReadTrademarkDatabase(tm(), 國外_CF) = True Then
            'bolTMFlow = True
            'Add By Sindy 2024/7/11
            If PField(1) = "CFT" Then
               bolCFTFlow = True
            ElseIf PField(1) = "FCT" And Left(PUB_GetST03(cp(14)), 1) = "F" Then '因FCT爭議案件是內商人員在承辦的
               bolFCTFlow = True
            Else
               bolTMFlow = True
            End If
            '2024/7/11 END
         End If
         'Add By Sindy 2018/9/26 記錄多申請人
         'Modify By Sindy 2020/6/8 讀取長度足9碼,後面讀資料較方便
         If tm(44) <> "" Then tm(44) = ChangeCustomerL(tm(44))
         If tm(54) <> "" Then tm(54) = ChangeCustomerL(tm(54))
         If tm(56) <> "" Then tm(56) = ChangeCustomerL(tm(56))
         If tm(23) <> "" Then
            tm(23) = ChangeCustomerL(tm(23))
            m_ManyAppl = m_ManyAppl & "," & tm(23)
         End If
         If tm(78) <> "" Then
            tm(78) = ChangeCustomerL(tm(78))
            m_ManyAppl = m_ManyAppl & "," & tm(78)
         End If
         If tm(79) <> "" Then
            tm(79) = ChangeCustomerL(tm(79))
            m_ManyAppl = m_ManyAppl & "," & tm(79)
         End If
         If tm(80) <> "" Then
            tm(80) = ChangeCustomerL(tm(80))
            m_ManyAppl = m_ManyAppl & "," & tm(80)
         End If
         If tm(81) <> "" Then
            tm(81) = ChangeCustomerL(tm(81))
            m_ManyAppl = m_ManyAppl & "," & tm(81)
         End If
         '2020/6/8 END
         If m_ManyAppl <> "" Then m_ManyAppl = Mid(m_ManyAppl, 2)
         '2018/9/26 END
         
         '****** 欄位值 ******
         m_PA26 = tm(23)
         m_PA27 = tm(78)
         m_PA28 = tm(79)
         m_PA29 = tm(80)
         m_PA30 = tm(81)
      ElseIf InStr("5,6", m_strSys) > 0 Then
         sp(1) = PField(1)
         sp(2) = PField(2)
         sp(3) = PField(3)
         sp(4) = PField(4)
         If ClsPDReadServicePracticeDatabase(sp(), 國外_CF) = True Then
            If m_strSys = "6" Then '商標:服務
               'bolTMFlow = True
               'Add By Sindy 2024/7/11
               If PField(1) = "CFC" Or (PField(1) = "S" And sp(9) <> "000") Then
                  bolCFTFlow = True
               ElseIf (PField(1) = "S" And sp(9) = "000") Then
                  bolFCTFlow = True
               Else
                  bolTMFlow = True
               End If
               '2024/7/11 END
            Else
               bolPAFlow = True
            End If
         End If
         'Add By Sindy 2018/9/26 記錄多申請人
         'Modify By Sindy 2020/6/8 讀取長度足9碼,後面讀資料較方便
         If sp(26) <> "" Then sp(26) = ChangeCustomerL(sp(26))
         If sp(35) <> "" Then sp(35) = ChangeCustomerL(sp(35))
         If sp(37) <> "" Then sp(37) = ChangeCustomerL(sp(37))
         If sp(8) <> "" Then
            sp(8) = ChangeCustomerL(sp(8))
            m_ManyAppl = m_ManyAppl & "," & sp(8)
         End If
         If sp(58) <> "" Then
            sp(58) = ChangeCustomerL(sp(58))
            m_ManyAppl = m_ManyAppl & "," & sp(58)
         End If
         If sp(59) <> "" Then
            sp(59) = ChangeCustomerL(sp(59))
            m_ManyAppl = m_ManyAppl & "," & sp(59)
         End If
         If sp(65) <> "" Then
            sp(65) = ChangeCustomerL(sp(65))
            m_ManyAppl = m_ManyAppl & "," & sp(65)
         End If
         If sp(66) <> "" Then
            sp(66) = ChangeCustomerL(sp(66))
            m_ManyAppl = m_ManyAppl & "," & sp(66)
         End If
         '2020/6/8 END
         If m_ManyAppl <> "" Then m_ManyAppl = Mid(m_ManyAppl, 2)
         '2018/9/26 END
         
         '****** 欄位值 ******
         m_PA26 = sp(8)
         m_PA27 = sp(58)
         m_PA28 = sp(59)
         m_PA29 = sp(65)
         m_PA30 = sp(66)
      
      'Add By Sindy 2021/9/3 + 法務
      ElseIf InStr("3,7", m_strSys) > 0 Then
         lC(1) = PField(1)
         lC(2) = PField(2)
         lC(3) = PField(3)
         lC(4) = PField(4)
         If ClsPDReadLawCaseDatabase(lC()) = True Then
            bolOtherFlow = True
         End If
         '記錄多申請人
         '讀取長度足9碼,後面讀資料較方便
         If lC(22) <> "" Then lC(22) = ChangeCustomerL(lC(22)) 'FC代理人
         If lC(12) <> "" Then lC(12) = ChangeCustomerL(lC(12)) '副本收受人
         If lC(26) <> "" Then lC(26) = ChangeCustomerL(lC(26)) '固定請款對象
         If lC(11) <> "" Then
            lC(11) = ChangeCustomerL(lC(11))
            m_ManyAppl = m_ManyAppl & "," & lC(11)
         End If
         If lC(43) <> "" Then
            lC(43) = ChangeCustomerL(lC(43))
            m_ManyAppl = m_ManyAppl & "," & lC(43)
         End If
         If lC(44) <> "" Then
            lC(44) = ChangeCustomerL(lC(44))
            m_ManyAppl = m_ManyAppl & "," & lC(44)
         End If
         If lC(45) <> "" Then
            lC(45) = ChangeCustomerL(lC(45))
            m_ManyAppl = m_ManyAppl & "," & lC(45)
         End If
         If lC(46) <> "" Then
            lC(46) = ChangeCustomerL(lC(46))
            m_ManyAppl = m_ManyAppl & "," & lC(46)
         End If
         '2020/6/8 END
         If m_ManyAppl <> "" Then m_ManyAppl = Mid(m_ManyAppl, 2)
         '2018/9/26 END
         
         '****** 欄位值 ******
         m_PA26 = lC(11)
         m_PA27 = lC(43)
         m_PA28 = lC(44)
         m_PA29 = lC(45)
         m_PA30 = lC(46)
         
      'Add By Sindy 2021/9/6 + 顧問
      ElseIf InStr("4,8", m_strSys) > 0 Then
         hc(1) = PField(1)
         hc(2) = PField(2)
         hc(3) = PField(3)
         hc(4) = PField(4)
         If ClsPDReadHireCaseDatabase(hc()) = True Then
            bolOtherFlow = True
         End If
         '記錄多申請人
         '讀取長度足9碼,後面讀資料較方便
         If hc(5) <> "" Then
            hc(5) = ChangeCustomerL(hc(5))
            m_ManyAppl = m_ManyAppl & "," & hc(5)
         End If
         If hc(24) <> "" Then
            hc(24) = ChangeCustomerL(hc(24))
            m_ManyAppl = m_ManyAppl & "," & hc(24)
         End If
         If hc(25) <> "" Then
            hc(25) = ChangeCustomerL(hc(25))
            m_ManyAppl = m_ManyAppl & "," & hc(25)
         End If
         If hc(26) <> "" Then
            hc(26) = ChangeCustomerL(hc(26))
            m_ManyAppl = m_ManyAppl & "," & hc(26)
         End If
         If hc(27) <> "" Then
            hc(27) = ChangeCustomerL(hc(27))
            m_ManyAppl = m_ManyAppl & "," & hc(27)
         End If
         '2020/6/8 END
         If m_ManyAppl <> "" Then m_ManyAppl = Mid(m_ManyAppl, 2)
         '2018/9/26 END
         
         '****** 欄位值 ******
         m_PA26 = hc(5)
         m_PA27 = hc(24)
         m_PA28 = hc(25)
         m_PA29 = hc(26)
         m_PA30 = hc(27)
         
      Else 'If InStr("3,4,7,8", m_strSys) > 0 Then '其他
         Screen.MousePointer = vbDefault
         MsgBox "讀取基本檔有誤，請洽電腦中心！", vbExclamation
         QueryData = False
         rsTmp.Close
         Set rsTmp = Nothing
         Call cmdExit_Click
         Exit Function
      End If
      '2018/4/18 END
      
      If Not IsNull(rsTmp.Fields("國家")) Then lblPA09.Caption = rsTmp.Fields("國家")
      m_PA48 = "" & rsTmp.Fields("PA48") 'Add By Sindy 2014/3/4 客戶案件案號
      m_PA75 = "" & rsTmp.Fields("PA75") 'Add By Sindy 2018/4/2 FC代理人
      m_PA77 = "" & rsTmp.Fields("PA77") 'Add By Sindy 2018/4/2 彼所案號
      If Not IsNull(rsTmp.Fields("PA08")) Then m_PA08 = rsTmp.Fields("PA08")
      m_Country = "" & rsTmp.Fields("PA09") '申請國家
      m_SaleArea = "" & rsTmp.Fields("Area")
      
      'Add By Sindy 2024/7/12 外商
      If bolCFTFlow = True Or bolFCTFlow = True Then
         SSTab1.TabVisible(intTab_承辦單) = False
         SSTab1.TabVisible(intTab_外專承辦單) = False
         txtCaseName(1).Visible = False: Label1(5).Visible = False
         txtCaseName(2).Visible = False: Label1(6).Visible = False
         If Not IsNull(tm(9)) Then lblPA08.Caption = tm(9): Label1(1).Caption = "類別："
         If bolCFTFlow = True Then
            cmd1(0).Caption = "商品名稱維護"
            If CheckUse("frm03010303_05", strExec, False) = True Then
               Call Cmd1_LostFocus(0)
               cmd1(0).Visible = True
            End If
            LblinfoNote.Caption = "註：檔名則以T000000.info.pdf，多個以上資料檔則加序號（例：T000000.info2.pdf，T000000.info3.pdf）"
         Else
            LblinfoNote.Visible = False '存卷區的加註不顯示,不鎖info
         End If
      '2024/7/12 END
      
      'Add By Sindy 2018/4/27
      '商標處
      ElseIf bolTMFlow = True Then
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
         
         'Add By Sindy 2018/9/20
         'Modify By Sindy 2020/9/9 + 99.個案聯絡,0.承辦人工作進度
         If intReceiveKind = "1" Or intReceiveKind = "99" Or intReceiveKind = "0" Then '1.待核判區
            'Modify By Sindy 2018/10/2 + 1202.核駁前先行通知
            'Modified by Lydia 2021/11/19 增加737智財協作之T案
            'If (PField(1) = "T" And (cp(10) = TMQ_T案 Or cp(10) = "1202")) Or _
               (PField(1) = "TS" And (cp(10) = TMQ_TS案 Or cp(10) = "1202")) Then
            If (PField(1) = "T" And (InStr(TMQ_T案, cp(10)) > 0 Or cp(10) = "1202")) Or _
               (PField(1) = "TS" And (InStr(TMQ_TS案, cp(10)) > 0 Or cp(10) = "1202")) Then
               cmdTMQ.Visible = True
            End If
         End If
         '2018/9/20 END
         
         'Add By Sindy 2018/7/26
         cmd1(0).Caption = "商品名稱維護"
         If CheckUse("frm03010303_05", strExec, False) = True Then
            Call Cmd1_LostFocus(0) 'Add By Sindy 2024/6/13
            cmd1(0).Visible = True
'         Else
'            cmd1(0).Visible = False
         End If
         '2018/7/26 END
         
         'Modify By Sindy 2018/10/1
         If UCase(m_PrevForm.Name) = UCase("frm090201_b") Then
            If cp(10) = "301" Or cp(10) = "102" Then '301.變更 102.延展
               Call cmdMod_LostFocus
               cmdMod.Visible = True '變更事項
            End If
         End If
         '2018/10/1 END
         
         SSTab1.TabVisible(intTab_承辦單) = False
         SSTab1.TabVisible(intTab_外專承辦單) = False 'Add By Sindy 2023/9/12
         txtCaseName(1).Visible = False: Label1(5).Visible = False
         txtCaseName(2).Visible = False: Label1(6).Visible = False
         LblinfoNote.Caption = "註：檔名則以T000000.info.pdf，多個以上資料檔則加序號（例：T000000.info2.pdf，T000000.info3.pdf）"
         'Add By Sindy 2018/9/25
         '條款
         txt2 = cp(49)
         '預估結果
         If "" & cp(23) = "1" Then
            Option1(0).Value = True
         ElseIf "" & cp(23) = "2" Then
            Option1(1).Value = True
         ElseIf "" & cp(23) = "3" Then
            Option1(2).Value = True
         End If
         '2018/9/25 END
         If Not IsNull(tm(9)) Then lblPA08.Caption = tm(9): Label1(1).Caption = "類別："
         
      Else
         If bolPAFlow = True Then '內專Flow
            bolFMP = PUB_ChkIsFMP(PField(1), PField(2), PField(3), PField(4), pa(9))
            If bolFMP = True Then
               bolOurFMP = PUB_FMPtoCheck(1, 2, PUB_GetST05(cp(14)), PField(1), PField(2), PField(3), PField(4)) '是否寰華案件
            Else
               bolOurFMP = False
            End If
            'Add By Sindy 2023/9/12 查詢外專資料
            '(cp(10) = "201" And Left(PUB_GetST03(m_EPMan), 1) = "F" And Trim(m_EPMan) <> "") :取消
            If PField(1) = "FCP" Or _
               PField(1) = "FG" Or _
               (bolFMP = True And Left(PUB_GetST03(cp(14)), 1) = "F") Then
               bolPAFlow = False
               SSTab1.TabVisible(intTab_承辦單) = False
               bolFCPFlow = True '外專Flow
               CmdCalendar.Visible = True 'Add By Sindy 2025/8/20
               cmdOK(5).Visible = True '可以看原始檔區 Add By Sindy 2023/11/16
               SSTab1.TabVisible(intTab_外專承辦單) = True
               '*****
               'Add By Sindy 2024/4/17
               If Pub_StrUserSt03 = "F22" Or Pub_StrUserSt03 = "M51" Then
                  cmdOutlook.Visible = True
               End If
               '2024/4/17 END
               'Add By Sindy 2025/10/21 排除P大陸案
               If (Pub_StrUserSt03 = "F21" Or Pub_StrUserSt03 = "M51") _
                  And cp(158) = 0 And cp(159) = 0 _
                  And InStr("203,204,205,107,433,210,242", cp(10)) > 0 _
                  And cp(118) <> "" _
                  And (PField(1) = "FCP" Or PField(1) = "FG") Then
                  If (intReceiveKind = 0 Or intReceiveKind = 99) Then '0.承辦人工作進度 99.個案聯絡
                     '上傳按鍵
                     CmdF21(0).Visible = True: CmdF21(1).Visible = True
                  End If
                  bolCmdF21 = True '要上傳
               End If
               '2025/10/21 END
               If UCase(TypeName(m_PrevForm)) = UCase("frm060107_1") Or _
                  (InStr(屬中說的案件性質, cp(10)) > 0 _
                   And UCase(TypeName(m_PrevForm)) = UCase("FRM100101_2") _
                   And GetStaffDepartment(m_FlowUserNum) = "F22") Then
                  m_bolSendChWrite = True '送中說
               End If
               '*****
               FrameFCPlink(0).Visible = True
               lstAtt(0).Height = 1300 'Add By Sindy 2024/5/30
'               'Modify By Sindy 2024/1/4 + intlstAtt0_Height
'               If intlstAtt0_Height = 0 Then
'                  lstAtt(0).Height = lstAtt(0).Height - FrameFCPlink(0).Height
'                  intlstAtt0_Height = lstAtt(0).Height
'               Else
'                  lstAtt(0).Height = intlstAtt0_Height
'               End If
'               '2024/1/4 END
               LblinfoNote.Visible = False '存卷區的加註不顯示,不鎖info
               '*****
               'Add By Sindy 2023/11/23
               '分割建議
               strExc(0) = "select pa162,DST05 from caseprogress a,patent,divsugtext" & _
                           " where dst09='" & m_EEP01 & "' and cp09=dst09" & _
                           " and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04" & _
                           " and dst01=pa01 and dst02=pa02 and dst03=pa03 and dst04=pa04 and pa162='Y'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  '中說請款修正定稿文字
                  strExc(0) = "select AMD05,nvl(CP27,0) CP27 from caseprogress a,patent,Amendedtext" & _
                              " where amd09='" & m_EEP01 & "' and cp09=amd09" & _
                              " and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04" & _
                              " and amd01=pa01 and amd02=pa02 and amd03=pa03 and amd04=pa04"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                  End If
               Else
                  m_PA162 = "" & RsTemp.Fields("pa162")
                  '待核判區 外專工程師核判主管才需要顯示此訊息
                  'Modify By Sindy 2024/3/15
                  'If intReceiveKind = 1 And Pub_StrUserSt03 = "F21" Then
                  If intReceiveKind = 1 And bolFCPFlow = True And Pub_StrUserSt03 <> "F22" Then
                     MsgBox "有輸入核准分割建議定稿文字，請審核!" & vbCrLf & vbCrLf & _
                            "(請點選【承辦進度】按鍵進入查看)", vbInformation
                  End If
               End If
               '2023/11/23 END
               '*****
               cmd1(1).Visible = True '相似案 Add By Sindy 2024/2/7
            Else
               SSTab1.TabVisible(intTab_外專承辦單) = False
            End If
            '2023/9/12 END
            cmd1(0).Caption = "專利相關案": cmd1(0).Visible = True
            If Not IsNull(rsTmp.Fields("種類")) Then lblPA08.Caption = rsTmp.Fields("種類")
         End If
         LblinfoNote.Caption = "註：檔名則以P000000.info.pdf，多個以上資料檔則加序號（例：P000000.info2.pdf，P000000.info3.pdf）"
         
         'Add By Sindy 2021/9/24
         If bolOtherFlow = True Then
            SSTab1.TabVisible(intTab_承辦單) = False
            SSTab1.TabVisible(intTab_外專承辦單) = False 'Add By Sindy 2023/9/12
            If InStr("4,8", m_strSys) > 0 Then '顧問
               txtCaseName(1).Visible = False: Label1(5).Visible = False
               txtCaseName(2).Visible = False: Label1(6).Visible = False
            End If
         End If
      End If
      '2018/4/27 END
      
      'Add By Sindy 2019/4/15 目前智權人員
      bolMCTFcase = False
      If InStr(ShowCurrCP13(PField(1), PField(2), PField(3), PField(4), m_Country), "MCTF") > 0 Then
         bolMCTFcase = True
      End If
      '2019/4/15 END
      
      'Modify By Sindy 2013/10/1
'      If Not IsNull(rsTmp.Fields("案件名稱")) Then
'         lblCaseName.Caption = rsTmp.Fields("案件名稱")
'      End If
      If Not IsNull(rsTmp.Fields("PA05")) Then
         txtCaseName(0).Text = rsTmp.Fields("PA05")
         txtCaseName(0).Tag = rsTmp.Fields("PA05")
         'Modify By Sindy 2023/1/9 薛經理:T案名稱,智權人員常會加入[商品類別],以茲區別,故開放供案件承辦人修改
         'If bolTMFlow = True Or bolOtherFlow = True Then txtCaseName(0).Locked = True
         'Modify By Sindy 2023/9/12 + Or bolFCPFlow = True
         'Modify By Sindy 2024/8/14 + Or bolFCTFlow = True
         If bolOtherFlow = True Or bolFCPFlow = True Or bolFCTFlow = True Then txtCaseName(0).Locked = True
      End If
      If Not IsNull(rsTmp.Fields("PA06")) Then
         txtCaseName(1).Text = rsTmp.Fields("PA06")
         txtCaseName(1).Tag = rsTmp.Fields("PA06")
         'Modify By Sindy 2023/1/9
         'If bolTMFlow = True Or bolOtherFlow = True Then txtCaseName(1).Locked = True
         'Modify By Sindy 2023/9/12 + Or bolFCPFlow = True
         'Modify By Sindy 2024/8/14 + Or bolFCTFlow = True
         If bolOtherFlow = True Or bolFCPFlow = True Or bolFCTFlow = True Then txtCaseName(1).Locked = True
      End If
      If Not IsNull(rsTmp.Fields("PA07")) Then
         txtCaseName(2).Text = rsTmp.Fields("PA07")
         txtCaseName(2).Tag = rsTmp.Fields("PA07")
         'Modify By Sindy 2023/1/9
         'If bolTMFlow = True Or bolOtherFlow = True Then txtCaseName(2).Locked = True
         'Modify By Sindy 2023/9/12 + Or bolFCPFlow = True
         'Modify By Sindy 2024/8/14 + Or bolFCTFlow = True
         If bolOtherFlow = True Or bolFCPFlow = True Or bolFCTFlow = True Then txtCaseName(2).Locked = True
      End If
      '2013/10/1 END
      '電子送件
      If cp(118) <> "" Then
      'Add by Sindy 2013/9/26
         'Modify By Sindy 2023/2/15 商標要增加判斷有承辦人發文日時,才要顯示電子送件
         'lblEApp.Visible = True
         If bolTMFlow = True Then
            If Val(cp(85)) > 0 Then
               lblEApp.Visible = True
            Else
               lblEApp.Visible = False
            End If
         Else
            lblEApp.Visible = True
         End If
         '2023/2/15 END
      Else
         lblEApp.Visible = False
      End If
      '2013/9/26 END
      
      '專利
      lblCMboth.Visible = False
      'Modify By Sindy 2023/9/12 + Or bolFCPFlow = True
      If bolPAFlow = True Or bolFCPFlow = True Then
         'Add By Sindy 2015/3/13
         '一案兩請
         strSql = "select cm05,cm06,cm07,cm08 from casemap where cm01='" & PField(1) & "' and cm02='" & PField(2) & "' and cm03='" & PField(3) & "' and cm04='" & PField(4) & "' and cm10='3'" & _
                  " Union select cm01,cm02,cm03,cm04 from casemap where cm05='" & PField(1) & "' and cm06='" & PField(2) & "' and cm07='" & PField(3) & "' and cm08='" & PField(4) & "' and cm10='3'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         lblCM10.Tag = ""
         If intI = 1 Then
            lblCM10.Visible = True
            lblCM10.Tag = RsTemp.Fields(0) & RsTemp.Fields(1) & IIf(RsTemp.Fields(2) & RsTemp.Fields(3) <> "000", "-" & RsTemp.Fields(2) & "-" & RsTemp.Fields(3), "")
         Else
            lblCM10.Visible = False
         End If
         '2015/3/13 END
         
         'Added by Lydia 2016/06/14 +台灣大陸案件提示
         lblCMboth.Visible = True
         lblCMboth.Caption = ""
         If (PField(1) = "P" Or PField(1) = "FCP") And m_Country = 台灣國家代號 Then
            If PUB_GetRefCaseChk(PField(1), PField(2), PField(3), PField(4), "CASEMAP", "0", "A", 大陸國家代號) Then
               lblCMboth.Caption = "有大陸案"
            End If
         ElseIf PField(1) = "P" And m_Country = 大陸國家代號 Then
            If PUB_GetRefCaseChk(PField(1), PField(2), PField(3), PField(4), "CASEMAP", "0", "A", 台灣國家代號) Then
               lblCMboth.Caption = "有台灣案"
            End If
         End If
         'end 2016/06/14
      
      'Add By Sindy 2021/9/13
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True
      ElseIf bolTMFlow = True Or bolCFTFlow = True Then
         lblCMboth.Visible = True
         lblCMboth.Caption = ""
         If cp(141) = "2" Then '註記收款後送件的案件
            'Add By Sindy 2021/12/13 國內收據才是判斷CP79；
            '國外請款單要抓acc1k0之a1k29，請參考共同查詢frm100101_2之收回
            If Left(cp(60), 1) = "X" Then
               IntTemp1 = 0: IntTemp2 = 0
               strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0,A1K25 FROM ACC1K0 WHERE A1K01='" & cp(60) & "'"
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                  If Not IsNull(adoRecordset1.Fields(0)) Then
                     IntTemp1 = IntTemp1 + adoRecordset1.Fields(0) '台幣金額
                  End If
                  If Not IsNull(adoRecordset1.Fields(4)) Then
                     IntTemp2 = IntTemp2 + adoRecordset1.Fields(4) 'decode(a1k29,'Y',a1k11,nvl(A1K30,0)) 已收金額
                  End If
                  m_dblCP79 = Val(IntTemp1) - Val(IntTemp2)
'                  If IntTemp1 = IntTemp2 Then
'                     Me.LblFee.Caption = "已收款可送件"
'                  Else
'                     Me.LblFee.Caption = "尚待收款"
'                  End If
               End If
            Else
               m_dblCP79 = Val(cp(79))
            End If
            '2021/12/13 END
            If m_dblCP79 = 0 Then '未收金額=0
               Me.lblCMboth.Caption = "已收款可送件"
            Else
               Me.lblCMboth.Caption = "尚待收款"
            End If
         End If
         '2021/9/13 END
      End If
      
      'Add By Sindy 2024/1/5
      Lbl926.Visible = False
      If bolFCPFlow = True And cp(10) = "926" Then
         If Val(cp(48)) > 0 And Val(cp(6)) > 0 Then
            Lbl926.Caption = "(二核)"
            Lbl926.Visible = True
         ElseIf Val(cp(48)) > 0 And Val(cp(6)) = 0 Then
            Lbl926.Caption = "(一核)"
            Lbl926.Visible = True
         End If
      End If
      '2024/1/5 END
      
      m_CP10Nm = "" & rsTmp.Fields("案件性質")
      'Add By Sindy 2013/11/27
      If "" & rsTmp.Fields("CP43") <> "" Then
         m_CP10Nm = m_CP10Nm & PUB_GetRelateCasePropertyName(m_EEP01, "1")
      End If
      '2013/11/27 END
      'lblCP10 = "" & rsTmp.Fields("案件性質") 'Add By Sindy 2013/9/2
      lblCP10 = m_CP10Nm 'Modify By Sindy 2018/4/26
      
'      If Not IsNull(rsTmp.Fields("收文日")) Then lblCP05.Caption = rsTmp.Fields("收文日")
'      If Not IsNull(rsTmp.Fields("智權人員")) Then lblCP13.Caption = rsTmp.Fields("智權人員")
'      If Not IsNull(rsTmp.Fields("本所期限")) Then lblCP06.Caption = rsTmp.Fields("本所期限")
'      If Not IsNull(rsTmp.Fields("承辦人")) Then lblCP14.Caption = rsTmp.Fields("承辦人")
'      If Not IsNull(rsTmp.Fields("承辦期限")) Then lblCP48.Caption = rsTmp.Fields("承辦期限")
      
      '未發文未取消收文時,新增下一流程鍵才顯示
      If Val(cp(27)) > 0 Or Val(cp(57)) > 0 Then
         Me.cmdAdd.Visible = False
         Me.cmdSend.Visible = False
         Me.cmdSave2.Visible = False 'Add By Sindy 2023/9/20
         Me.cmdSave3.Visible = False 'Add By Sindy 2023/9/20
         Me.Frame945.Enabled = False 'Add By Sindy 2025/9/4
         Me.cmdAddAtt(0).Enabled = False
         CmdF21(0).Enabled = False 'Add By Sindy 2025/10/28
         Me.cmdRemAtt(0).Enabled = False
         Me.cmdAddAtt(1).Enabled = False
         Me.cmdRemAtt(1).Enabled = False
      End If
      m_EP01 = "" & rsTmp.Fields("EP01") 'Add By Sindy 2013/9/17 目次
      m_EP06 = "" & rsTmp.Fields("EP06") '文件齊備日
      m_EP08 = "" & rsTmp.Fields("EP08") 'Add By Sindy 2013/9/17 會稿完成日
      m_EP38 = "" & rsTmp.Fields("EP38") 'Add By Sindy 2013/9/17 智權人員會稿完成日
      m_CPM27 = "" & rsTmp.Fields("CPM27") 'Add By Sindy 2013/9/17 核稿人不可判發
      m_CPM26 = "" & rsTmp.Fields("CPM26") 'Add By Sindy 2013/9/23 專業部電子檔副檔名
      'Add By Sindy 2013/9/23
      m_EP41 = "" & rsTmp.Fields("EP41") 'Add By Sindy 2015/3/16 核稿語文
      '2013/9/23 END
'      If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'         Me.Caption = Trim(Me.Caption) & "（" & rsTmp.Fields("cp09") & " - " & CP(10) & lblCP10 & "）"
'      End If
      
      'Add By Sindy 2014/9/5 是否為P大陸新案
      bolP020NewCase = False
      If PField(1) = "P" And m_Country = "020" And InStr(NewCasePtyList, cp(10)) > 0 Then
         bolP020NewCase = True
      End If
      '2014/9/5 END
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
   cmdAdd.Enabled = True
   If intReceiveKind = 99 And cp(163) <> "" Then '99.個案聯絡
      If cp(163) <> cp(9) Then
         strSql = "Select CP01,CP02,CP03,CP04,Decode('" & m_Country & "','000',CPM03,CPM04) as 案件性質" & _
                  " from caseprogress,CasePropertyMap" & _
                  " where cp09='" & cp(163) & "' And CP01=CPM01(+) And CP10=CPM02(+)"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            txtNote.Text = txtNote.Text & rsTmp.Fields("cp01") & "-" & rsTmp.Fields("cp02") & IIf(rsTmp.Fields("cp03") & rsTmp.Fields("cp04") = "000", "", "-" & rsTmp.Fields("cp03") & "-" & rsTmp.Fields("cp04")) & _
                           "(" & rsTmp.Fields("案件性質") & ")"
            txtNote.Width = 5000
            txtNote.Visible = True
            'Add By Sindy 2022/6/22 程序組要在共同查詢操作查名結果 ex:T-239802
            If UCase(m_PrevForm.Name) = "FRM100101_2" And Pub_StrUserSt03 = "P22" Then '程序組-共同查詢
               cmdAdd.Enabled = True
            Else
               cmdAdd.Enabled = False
            End If
         End If
         rsTmp.Close
      End If
   End If
   '2020/12/1 END
   
   Call QueryEmpElectronData '承辦單
   
   '***** Modify By Sindy 2016/10/6 調整先抓EMP_流程控制除外的狀態, 再抓待回覆資料 *****
   intLastEEP02 = 0
   Frame3.Visible = False 'Add By Sindy 2013/9/17
   strExc(0) = "select * From empelectronprocess where eep01='" & m_EEP01 & "'" & _
               " and eep04 not in(" & EMP_流程控制除外的狀態 & ") order by eep02 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      intLastEEP02 = rsTmp.Fields("eep02")
      m_EEP15 = "" & rsTmp.Fields("eep15") 'Add By Sindy 2020/9/29
      m_EEP11 = "" & rsTmp.Fields("eep11") 'Add By Sindy 2020/9/29
      'Add By Sindy 2013/9/17
      If intReceiveKind = 0 Then
         'Modify By Sindy 2013/11/27
'            If m_EP08 = "" And "" & rsTmp.Fields("EEP04") = EMP_會完 Then
'               'Modify By Sindy 2013/9/18 若已確認過不更新者除外
'               If PUB_ChkEmpFlowExists(lblCP09, EMP_不自動更新會完日, rsTmp.Fields("EEP02")) = False Then
'               '2013/9/18 END
'                  cmdAdd.Visible = False
'                  Frame3.Visible = True
'               End If
'            End If
         'Add By Sindy 2016/3/15
         If m_EP06 = "" Then
            If PUB_ChkEmpFlowExists(lblCP09, EMP_會圖, , strRefEEP02) = True Then
               If PUB_ChkEmpFlowExists(lblCP09, EMP_圖完, strRefEEP02, strRefEEP02) = True Then
                  If PUB_ChkEmpFlowExists(lblCP09, EMP_不自動更新齊備日, strRefEEP02) = False Then
                     cmdAdd.Visible = False
                     Frame3.Visible = True
                     'Modify By Sindy 2022/10/7 + 文
                     'cmdChkEP08.Caption = "是否會圖完成"
                     cmdChkEP08.Caption = "是否會(圖/文)完成"
                  End If
               End If
            End If
         '最後一道若為會完並且無會稿完成日時,必須確認會稿完成日後,才可繼續進行下一道流程
         ElseIf m_EP08 = "" Then
         '2016/3/15 END
         'If m_EP08 = "" Then
            If PUB_ChkEmpFlowExists(lblCP09, EMP_送會, , strRefEEP02) = True Then
               If PUB_ChkEmpFlowExists(lblCP09, EMP_會完, strRefEEP02, strRefEEP02) = True Then
                  'Modify By Sindy 2016/3/8 + PUB_ChkEmpFlowExists(lblCP09, EMP_會完重修, strRefEEP02) = False
                  If PUB_ChkEmpFlowExists(lblCP09, EMP_不自動更新會完日, strRefEEP02) = False And _
                     PUB_ChkEmpFlowExists(lblCP09, EMP_會完重修, strRefEEP02) = False Then
                     cmdAdd.Visible = False
                     Frame3.Visible = True
                  End If
               End If
            End If
         End If
         'Add By Sindy 2016/3/22 當工程師歷程最後一道為「圖修」且未做過確認
         If "" & rsTmp.Fields("EEP04") = EMP_圖修 And InStr("" & rsTmp.Fields("EEP11"), "客戶修改") = 0 Then
            '若此時「齊備日」非空白並且（齊備日＝圖修送出日期），系統請工程師確認「是否為客戶修改？」
            If m_EP06 <> "" And m_EP06 <> "" & rsTmp.Fields("EEP06") Then
               If MsgBox("是否為客戶修改？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  '若工程師點選「非客戶修改」或工程師超過4個工作小時未確認，則系統不作特別的操作
                  strSql = "update empelectronprocess set eep11='非客戶修改;'||eep11 where eep01='" & m_EEP01 & "' and eep02=" & intLastEEP02
                  cnnConnection.Execute strSql
               Else
                  '若工程師點選是「客戶修改」，則系統自動以「圖修」的日期建為新的「齊備日」
                  strSql = "update engineerprogress set ep06=" & rsTmp.Fields("EEP06") & " where ep02='" & lblCP09 & "'"
                  cnnConnection.Execute strSql
                  strSql = "update empelectronprocess set eep11='客戶修改;'||eep11 where eep01='" & m_EEP01 & "' and eep02=" & intLastEEP02
                  cnnConnection.Execute strSql
                  m_EP06 = rsTmp.Fields("EEP06")
                  If intReceiveKind = 0 Then '0.承辦人工作進度
                     m_PrevForm.txt1(2) = rsTmp.Fields("EEP06") - 19110000
                  End If
               End If
            End If
         End If
      End If
      '2013/9/13 END
   End If
   rsTmp.Close
   '2013/11/5 END
   
   'Add By Sindy 2013/9/13 增加顯示待回的最後流程
   'Add By Sindy 2013/11/5
   Label6.Caption = ""
   'Modify By Sindy 2024/11/12 eep02 => *
   strExc(0) = "select * From empelectronprocess where eep01='" & m_EEP01 & "' and eep09='Y' order by eep02 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      intLastEEP02 = rsTmp.Fields("eep02")
      'Modify By Sindy 2013/10/16 " and eep04 not in(" & EMP_流程控制除外的狀態 & ",'" & EMP_草完 & "','" & EMP_標號 & "'))"
   '   strExc(0) = "select *" & _
   '               " From empelectronprocess,allcode" & _
   '               " where eep01='" & m_EEP01 & "'" & _
   '               " and eep02=(select max(eep02) From empelectronprocess where eep01='" & m_EEP01 & "'" & _
   '               " and eep04 not in(" & EMP_流程控制除外的狀態 & ",'" & EMP_草完 & "','" & EMP_標號 & "'))" & _
   '               " and ac01='09' And eep04=ac02(+)"
      'Modify By Sindy 2013/11/5 直接在eep04 not in 中增加排除草完和標號會導至無法延用這2項的附件,因此改寫SQL
      rsTmp.Close
      strExc(0) = "select *" & _
                  " From empelectronprocess,allcode" & _
                  " where eep01='" & m_EEP01 & "'" & _
                  " and eep02=" & intLastEEP02 & _
                  " and ac01='09' And eep04=ac02(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         '檢查最後一道待回覆=Y
         If "" & rsTmp.Fields("EEP09") = "Y" Then
            m_EEP15 = "" & rsTmp.Fields("eep15") 'Add By Sindy 2020/9/29
            m_EEP11 = "" & rsTmp.Fields("eep11") 'Add By Sindy 2020/9/29
            If intReceiveKind = 0 Or intReceiveKind = 3 Or intReceiveKind = 99 Then '0.承辦人工作進度 及 3.繪圖人員工作進度
               'Add By Sindy 2015/3/16
               If rsTmp.Fields("AC03") = "送英核" And m_EP41 = "2" Then
                  Label6.Caption = "送日核中"
               Else
               '2015/3/16 END
                  Label6.Caption = rsTmp.Fields("AC03") & "中"
               End If
               'Add By Sindy 2024/11/12
               If "" & rsTmp.Fields("EEP05") = m_FlowUserNum And intReceiveKind = 99 Then '共同查詢:99=個案聯絡
                  If rsTmp.Fields("eep04") = EMP_送會 Or rsTmp.Fields("eep04") = EMP_會圖 Then
                     strExc(10) = "待會稿區"
                  ElseIf rsTmp.Fields("eep04") = EMP_送排版 Or rsTmp.Fields("eep04") = EMP_送轉檔 Then
                     strExc(10) = "待排版區"
                  Else
                     strExc(10) = "待核判區"
                  End If
                  MsgBox "此案件【" & Label6.Caption & "】請至【" & strExc(10) & "】操作適當的歷程狀態！", vbExclamation
                  QueryData = False
                  rsTmp.Close
                  Set rsTmp = Nothing
                  Call cmdExit_Click
                  Exit Function
               End If
               '2024/11/12 END
               
            ElseIf intReceiveKind = 1 Then '待核判區
               If "" & rsTmp.Fields("EEP05") = m_FlowUserNum Then
                  'Add By Sindy 2023/10/2
                  If rsTmp.Fields("AC03") = "翻譯交稿" Or rsTmp.Fields("AC03") = "送排版" Or rsTmp.Fields("AC03") = "送轉檔" Then
                     Label6.Caption = Replace(rsTmp.Fields("AC03"), "送", "") & "中"
                  ElseIf rsTmp.Fields("AC03") = "排版完成" Or rsTmp.Fields("AC03") = "送核稿分案" Then
                     Label6.Caption = rsTmp.Fields("AC03")
                  Else
                  '2023/10/2 END
                     Label6.Caption = "待核判"
                  End If
               End If
            ElseIf intReceiveKind = 2 Then '待會稿區
               If "" & rsTmp.Fields("EEP05") = m_FlowUserNum Then
                  Label6.Caption = "待會稿"
               End If
            End If
         End If
      End If
   End If
   rsTmp.Close
   '***** 2016/10/6 END
   
   '承辦電子簽核資料
   'Modify By Sindy 2013/10/16 +eep12
   'Modify By Sindy 2015/6/15 +ac03
   'Modify By Sindy 2016/3/9 +eep09
   strSql = "Select distinct EEP02 as 順序,EEP03,s1.ST02||eep12 as 發送者,EEP04,decode(eep04,'" & EMP_附加流程 & "',decode(c2.CP43,'',ac03,Decode('" & m_Country & "','000',CPM03,CPM04)),ac03) as 流程狀態,EEP05,decode(s2.ST02,null,eep05,s2.ST02) as 收受者,sqldatet(EEP06)||' ' ||sqltime(EEP07) as 送出時間,EEP10 as 副本收受者,EEP08 as 意見內容,EEP10,c1.CP43,ac03,eep09 as 待回,eep13 as 顯示,eep14,EEP11" & _
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
   pub_QL05 = ";本所案號：" & PField(1) & "-" & PField(2) & "-" & PField(3) & "-" & PField(4) & ";總收文號：" & m_EEP01 & "(承辦歷程)" 'Add By Sindy 2025/8/7
   If rsTmp.RecordCount > 0 Then
      If pub_QL04 <> "" Then InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2025/8/7
      Set GRD1.Recordset = rsTmp
      For ii = 1 To GRD1.Rows - 1
         txtEEP10_2 = GRD1.TextMatrix(ii, 10)
         Call txtEEP10_2_LostFocus
         GRD1.TextMatrix(ii, 8) = txtEEP10_2
         '判斷有相關總收文號才做案件性質轉換
         If GRD1.TextMatrix(ii, 4) = "附加流程" Then
            If GRD1.TextMatrix(ii, 11) <> "" Then
               GRD1.TextMatrix(ii, 4) = m_CP10Nm 'Modify By Sindy 2013/11/27 lblCP10 & PUB_GetRelateCasePropertyName(m_EEP01, "1")
            End If
         'Add By Sindy 2015/3/16
         ElseIf GRD1.TextMatrix(ii, 4) = "送英核" And m_EP41 = "2" Then
            GRD1.TextMatrix(ii, 4) = "送日核"
         '2015/3/16 END
         'Add By Sindy 2023/12/18 交辦時，送英核或送排版或送轉檔等，收受者要顯示原收受者
         ElseIf GRD1.TextMatrix(ii, 4) = "交辦" And _
                InStr(GRD1.TextMatrix(ii, 16), "原收受者:") > 0 And _
                InStr(GRD1.TextMatrix(ii, 16), "流程狀態:") > 0 Then
            strEEP04is39_EEP11 = Trim(GRD1.TextMatrix(ii, 16))
         ElseIf strEEP04is39_EEP11 <> "" And _
                Mid(strEEP04is39_EEP11, InStr(strEEP04is39_EEP11, "流程狀態:") + 5, 2) = GRD1.TextMatrix(ii, 3) Then
            strEEP05 = Mid(strEEP04is39_EEP11, InStr(strEEP04is39_EEP11, "原收受者:") + 5, 5)
            If GetPrjSalesNM(strEEP05) <> "" Then
               GRD1.TextMatrix(ii, 6) = GetPrjSalesNM(strEEP05)
            End If
            strEEP04is39_EEP11 = ""
         '2023/12/18 END
         End If
      Next ii
      '若有資料游標停在第一筆
      GRD1.Visible = False
      GRD1.col = 0
      GRD1.row = 1
      dblPrevRow = GRD1.row
      If rsTmp.RecordCount > 0 Then
         For ii = 0 To GRD1.Cols - 1
            GRD1.col = ii
            GRD1.CellBackColor = &HFFC0C0
         Next ii
         Call ReadData(False)
      End If
      GRD1.Visible = True
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
      GRD1.Rows = 2
      GRD1.col = 0
      GRD1.row = 1
   End If
   rsTmp.Close
   
   Call EmpFlowRole
   Call ReadAttachFile_other(m_EEP01) 'Add By Sindy 2013/9/25 查詢存卷區
   
   Call SetTxtLpNote(True) 'Add By Sindy 2020/9/29
      
   'Add By Sindy 2023/12/6
   'Modify By Sindy 2024/7/11 + bolFCTFlow, bolCFTFlow
   If Forms(0).mnuChUser.Visible = True Then
      strExc(10) = IIf(bolFCPFlow = True, " (FCPFlow)", _
                   IIf(bolPAFlow = True, " (PAFlow)", _
                   IIf(bolTMFlow = True, " (TMFlow)", _
                   IIf(bolFCTFlow = True, " (FCTFlow)", _
                   IIf(bolCFTFlow = True, " (CFTFlow)", " (OtherFlow)")))))
      If InStr(Me.Caption, strExc(10)) = 0 Then
         Me.Caption = Me.Caption & strExc(10)
      End If
   End If
   '2023/12/6 END
   
   Me.SSTab1.Tab = 0
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'讀取或預帶承辦單
Private Sub QueryEmpElectronData()
Dim objText As Object
Dim Rs As New ADODB.Recordset 'Add By Sindy 2024/1/17
   
   'Add By Sindy 2023/9/19
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
      Me.SSTab1.Tab = intTab_外專承辦單
      Frame945.Top = 420: Frame945.Left = 60
      Frame201.Top = 420: Frame201.Left = 60
      ChkEED08.Visible = False
      ChkEED08.Tag = ""
      Me.Frame945.Visible = False: Frame945.Tag = "" '預設值
      Me.Frame201.Visible = True '預設值
      If cp(10) = "945" And _
         (PUB_ChkEmpFlowExists(m_EEP01, EMP_送判) = True Or PUB_ChkEmpFlowExists(m_EEP01, EMP_發文歸檔) = True) Then
         ChkEED08.Visible = True
      '當告代掛相關收文號為電話連絡單,增加可以輸入【管制下一程序期限】
      ElseIf cp(10) = 告知代理人 And cp(43) <> "" Then
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
      'Add By Sindy 2025/11/7
      If Frame945.Tag <> "" Then
         If m_EEP04New = "已送件" Then
            Frame945.Enabled = False
         Else
            Frame945.Enabled = True
         End If
      End If
      '2025/11/7 END
      
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
         If Not IsNull(RsTemp.Fields("EED08")) Then
            ChkEED08.Value = 1
         Else
            ChkEED08.Value = 0
         End If
         ChkEED08.Tag = ChkEED08.Value
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
      Else
         If m_bolSendChWrite = True Then '送中說
            strSql = "Select *" & _
                     " From engineerprogress" & _
                     " Where EP02='" & m_EEP01 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               m_EP09 = "" & RsTemp.Fields("EP09")
            End If
            '帶預設值
            If cp(10) = "924" Then 'Claims翻譯交稿
               '產生會稿承辦單
               If Pub_PrintFCP924Form(pa(1), pa(2), pa(3), pa(4), m_EEP01, m_strColName, m_strColText, , True, m_intColCnt) = True Then
'                  txt3(0).Text = GetColValues("速別")
'                  txt3(1).Text = GetColValues("受文者")
'                  txt3(2).Text = GetColValues("主旨")
                  txt3(4).Text = Replace(Replace(GetColValues("備註"), "□", ""), "■", "")
                  If GetColValues("譯者") = lblCP10 Then
                  ElseIf GetColValues("譯者") <> "" Then
                     txt3(3).Text = GetColValues("譯者")
                     Call TXT3_LostFocus(3)
                  Else
                     Frame7.Visible = False '譯者
                  End If
                  txt3(5).Text = GetColValues("管制人")
                  Call TXT3_LostFocus(5)
               End If
            ElseIf (Val(m_EP09) > 0 And cp(10) = "201") _
               Or InStr(屬中說的案件性質, cp(10)) > 0 Then
               '產生翻譯承辦單
               If Pub_PrintFCP201Form(pa(1), pa(2), pa(3), pa(4), m_EEP01, m_strColName, m_strColText, , True, m_intColCnt) = True Then
'                  txt3(0).Text = GetColValues("速別")
'                  txt3(1).Text = GetColValues("受文者")
'                  txt3(2).Text = GetColValues("主旨")
                  txt3(4).Text = Replace(Replace(GetColValues("備註"), "□", ""), "■", "")
                  If GetColValues("譯者") = lblCP10 Then
                  ElseIf GetColValues("譯者") <> "" Then
                     txt3(3).Text = GetColValues("譯者")
                     Call TXT3_LostFocus(3)
                  Else
                     Frame7.Visible = False '譯者
                  End If
                  txt3(5).Text = GetColValues("FCP管制")
                  Call TXT3_LostFocus(5)
               End If
            End If
         End If
      End If
      
      '檔案名稱
      'Modify By Sindy 2024/1/17 Bobbie:帶客戶提供文件(中說)的檔案名稱
      'Modify By Sindy 2024/1/30 +Or cp(10) = "235" 核對中說格式
      If cp(10) = "209" Or cp(10) = "235" Then
         strSql = "Select *" & _
                  " From CustSupportDoc" & _
                  " Where csd01='" & PField(1) & "' and csd02='" & PField(2) & "' and csd03='" & PField(3) & "' and csd04='" & PField(4) & "'" & _
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
      
      txtEP12.Text = m_EP12 'Add By Sindy 2024/1/23 承辦備註改為顯示作業備註
      txtEP12.Tag = txtEP12.Text 'Add By Sindy 2024/4/10
      If Val(cp(27)) > 0 Or Val(cp(57)) > 0 Then
         For Each objText In Me.txt3
            'objText.Enabled = False
            objText.Locked = True
         Next
         txtEP12.Locked = True
      Else
         For Each objText In Me.txt3
            objText.Enabled = True
         Next
         txtEP12.Enabled = True
         txt3(6).Enabled = False '打字室僅顯示資料
         txtEP12.Enabled = True
         'Add By Sindy 2024/1/23 作業備註
         'Modify By Sindy 2024/3/29 mark: 開放其他人員也可以
'         If intReceiveKind = 0 Then '承辦人才可以修改
'            'txtEP12.Locked = False
'         Else
'            'txtEP12.Locked = True
'            txtEP12.Enabled = False
'         End If
'         '2024/1/23 END
      End If
      '2023/9/19 END
   '內專
   ElseIf bolPAFlow = True Then
      '先清除承辦單內容
      For Each objText In Me.txt1
         objText.Text = ""
         objText.Tag = ""
      Next
      Me.lblFa.Caption = ""
      If m_Country = "000" Then '台灣案不顯示代理人
         Me.Frame2.Visible = False
      Else
         Me.Frame2.Visible = True
      End If
      '讀取承辦單內容
      strSql = "Select *" & _
               " From EmpElectronData" & _
               " Where EED01='" & m_EEP01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Not IsNull(RsTemp.Fields("EED02")) Then
            txt1(5).Text = RsTemp.Fields("EED02")
            txt1(5).Tag = RsTemp.Fields("EED02")
         End If
         If Not IsNull(RsTemp.Fields("EED03")) Then
            txt1(6).Text = RsTemp.Fields("EED03")
            txt1(6).Tag = RsTemp.Fields("EED03")
         End If
         If Not IsNull(RsTemp.Fields("EED04")) Then
            txt1(0).Text = RsTemp.Fields("EED04")
            txt1(0).Tag = RsTemp.Fields("EED04")
         End If
         If Not IsNull(RsTemp.Fields("EED05")) Then
            txt1(4).Text = RsTemp.Fields("EED05")
            txt1(4).Tag = RsTemp.Fields("EED05")
         End If
      Else
         'Modify By Sindy 2013/12/24
         If intReceiveKind = 0 Or (intReceiveKind = 1 And Left(m_DMMan, 5) <> m_FlowUserNum) Then '0.承辦人工作進度 1.待核判區
         '2013/12/24 END
            '帶預設值
            Call SetEmpPaper
         End If
      End If
      If Val(cp(27)) > 0 Or Val(cp(57)) > 0 Then
         Me.Frame2.Enabled = False
         For Each objText In Me.txt1
            objText.Enabled = False
         Next
      Else
         Me.Frame2.Enabled = True
         For Each objText In Me.txt1
            objText.Enabled = True
         Next
         If Me.txt1(0) <> "" Then
            Me.txt1(0).Enabled = False 'True
         End If
      End If
   End If
   
   Set Rs = Nothing
End Sub

'Add By Sindy 2023/9/19
Private Function GetColValues(strTitN As String) As String
   GetColValues = ""
   For ii = 0 To m_intColCnt
      If m_strColName(ii) = strTitN Then
         GetColValues = m_strColText(ii)
         Exit For
      End If
   Next ii
End Function

Private Sub SetEmpPaper()
Dim arrTemp
Dim bolAsk As Boolean
'Add By Sindy 2013/9/2
Dim rsA As New ADODB.Recordset
Dim strTempName As String
'2013/9/2 END

   If bolPAFlow = True Then
      '2005/07/06 當新申請案，受文者及副本收受者不可修改，主旨不變
      If cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "105" _
         Or cp(10) = "109" Or cp(10) = "110" Or cp(10) = "112" Or cp(10) = "113" _
         Or cp(10) = "114" Or cp(10) = "115" Or cp(10) = "118" Or cp(10) = "301" _
         Or cp(10) = "302" Or cp(10) = "303" Or cp(10) = "304" Or cp(10) = "305" _
         Or cp(10) = "306" Or cp(10) = "307" Or cp(10) = "803" Then
         '受文者
         Me.txt1(5).Text = IIf(m_Country = "000", "智慧局", GetNationName(m_Country, 0) & " 代理人")
         Me.txt1(5).Enabled = False
         '副本收受者
         Me.txt1(6).Text = "北所、" & IIf(m_SaleArea = "1", "", IIf(m_SaleArea = "2", "中所、", IIf(m_SaleArea = "3", "南所、", IIf(m_SaleArea = "4", "高所、", "")))) & "客戶"
         Me.txt1(6).Enabled = False
   '      '主旨
   '      '2005/03/15 郭說不印專利種類，改印案件性質
   '      Me.txt1(0) = "為「" & Trim(lblCaseName.Caption) & "」" & GetNationName(m_country, 0) & lblCP10 & "專利案提出申請。"
         '備註
         
         'Modify By Sindy 2013/10/2 改到送會時加入EMail裡提示
   '      '2007/12/03 郭加入台灣發明新型在備註欄放入提示
   '      If (CP(10) = "101" Or CP(10) = "102") And m_country = "000" Then
   '         Me.txt1(4).Text = "請注意：本案若同時或隨後可能申請大陸專利" & vbCrLf & "　　　　，請留意是否有超頁超項問題：" & vbCrLf & "1.專利說明書(含申請專利範圍、圖式)以30頁" & vbCrLf & "　為限，每增加1頁加收新台幣500元。" & vbCrLf & "2.申請專利範圍以10項為限，每增加1項加收" & vbCrLf & "　新台幣1000元。"
   '      End If
      Else
         '受文者
         Me.txt1(5).Text = ""
         'Add By Sindy 2013/10/1
         '核駁,檢索報告
         'Modify By Sindy 2013/10/4 C類來函也掛客戶
         'Modify By Sindy 2014/5/20 +分析 Or CP(10) = "941"
         If Left(m_EEP01, 1) = "C" Or cp(10) = "1002" Or cp(10) = "1209" Or cp(10) = "941" Then
            Me.txt1(5).Text = "客戶"
         Else
         '2013/10/1 END
            'Add By Sindy 2013/9/2 顯示案件國家收費表的主管機關於畫面上
            If PField(1) = "P" Then
               If m_Country = "000" Then
                  strSql = "Select CF10 FROM CASEFEE WHERE CF01='" & PField(1) & "' AND CF02='" & m_Country & "' AND CF03='" & cp(10) & "'"
                  If rsA.State <> adStateClosed Then rsA.Close
                  rsA.CursorLocation = adUseClient
                  rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     Me.txt1(5).Text = "" & rsA.Fields(0).Value
                  End If
                  Set rsA = Nothing
               Else
                  AddAgent Combo2, PField
                  Label2(14).Caption = ""
                  strExc(1) = Combo2.Text
                  '加判斷是否為聯絡人
                  If InStr(strExc(1), "-") > 0 Then
                     If ClsPDGetContact(strExc(1), strTempName) Then
                        Combo2 = strExc(1)
                        Label2(14) = strTempName
                     End If
                  '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
                  ElseIf PUB_GetAgentName(PField(1), strExc(1), strTempName) = True Then
                     Combo2.Text = strExc(1)
                     Label2(14).Caption = strTempName
                  End If
                  If Label2(14).Caption <> "" Then
                     Me.txt1(5).Text = Label2(14).Caption
                  End If
               End If
            Else
               Me.txt1(5).Text = GetNationName(m_Country, 0) & " 代理人"
            End If
            '2013/9/2 END
         End If
         
         '副本收受者
         'Modify By Sindy 2013/9/2
         'Me.txt1(6).Text = ""
         'Add By Sindy 2013/10/1
         '核駁,檢索報告
         'Modify By Sindy 2013/10/4 C類來函也掛客戶
         'Modify By Sindy 2014/5/20 +分析 Or CP(10) = "941"
         If Left(m_EEP01, 1) = "C" Or cp(10) = "1002" Or cp(10) = "1209" Or cp(10) = "941" Then
            Me.txt1(6).Text = "北所" & IIf(m_SaleArea = "1", "", IIf(m_SaleArea = "2", "、中所", IIf(m_SaleArea = "3", "、南所", IIf(m_SaleArea = "4", "、高所", ""))))
         Else
         '2013/10/1 END
            Me.txt1(6).Text = "北所" & IIf(m_SaleArea = "1", "", IIf(m_SaleArea = "2", "、中所", IIf(m_SaleArea = "3", "、南所", IIf(m_SaleArea = "4", "、高所", "")))) & "、客戶"
         End If
         '2013/9/2 END
   '      '主旨
   '      '2005/07/22 郭說後面加專利種類   專利之 案件性質
   '      Me.txt1(0).Text = "「" & Trim(lblCaseName.Caption) & "」" & GetNationName(m_country, 0) & Trim(lblPA08.Caption) & "專利之" & lblCP10
      End If
      Call GetPaperMain '抓取主旨
      
      '2013/4/16 美專新申請案輸入完稿日列印承辦單前選擇
      '條件：1. 主張優先權日是在2013年3月16日(不含)之前的美專新申請案,2. CIP,分割案之須判斷原案之申請日或主張優先權日是在2013年3月16日(不含)之前;
      'And txt1(3) <> "" 有輸入完稿日時
      If PField(1) = "CFP" And m_Country = "101" And InStr("101,113,307", cp(10)) > 0 Then
         arrTemp = Split(Me.lblCaseNo.Caption, "-")
         bolAsk = False
         If cp(10) = "307" Then
            strExc(0) = "select nvl(pd05,pa10) from divisioncase,patent,pridate where dc01='" & arrTemp(0) & "' and dc02='" & arrTemp(1) & "' and dc03='" & arrTemp(2) & "' and dc04='" & arrTemp(3) & "'" & _
               " and pa01(+)=dc05 and pa02(+)=dc06 and pa03(+)=dc07 and pa04(+)=dc08 and pd01(+)=dc05 and pd02(+)=dc06 and pd03(+)=dc07 and pd04(+)=dc08 and nvl(pd05,pa10)<20130316 and rownum<2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               bolAsk = True
            End If
         ElseIf cp(10) = "113" Then
            strExc(0) = "select nvl(pd05,pa10) from patent,pridate where pa01='" & arrTemp(0) & "' and pa02='" & arrTemp(1) & "' and pa03='0' and pa04='" & arrTemp(3) & "'" & _
               " and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04 and nvl(pd05,pa10)<20130316 and rownum<2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               bolAsk = True
            End If
         Else
            strExc(0) = "select pd05 from patent,pridate where pa01='" & arrTemp(0) & "' and pa02='" & arrTemp(1) & "' and pa03='" & arrTemp(2) & "' and pa04='" & arrTemp(3) & "'" & _
               " and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04 and pd05<20130316 and rownum<2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               bolAsk = True
            End If
         End If
         If bolAsk = True Then
   '         strExc(0) = "This application (1) claims priority to or the benefit of an" & vbCrLf & _
   '                     " application filed before March 16, 2013 and (2) also contains," & vbCrLf & _
   '                     " or contained at any time, a claim to a claimed invention that" & vbCrLf & _
   '                     " has an effective filing date on or after March16, 2013."
            strExc(0) = "This application (1) claims priority to" & vbCrLf & _
                        "or the benefit of an application filed" & vbCrLf & _
                        "before March 16, 2013 and (2) also" & vbCrLf & _
                        "contains, or contained at any time, a" & vbCrLf & _
                        "claim to a claimed invention that has an" & vbCrLf & _
                        "effective filing date on or after" & vbCrLf & _
                        "March16, 2013."
            If MsgBox(strExc(0), vbYesNo, "美專案件FITF之控制！( 請選擇是或否 )") = vbYes Then
               'strExc(0) = Replace(strExc(0), vbCrLf, "") & vbCrLf & "Y▓ N□"
               strExc(0) = strExc(0) & vbCrLf & "Y▓ N□"
            Else
               'strExc(0) = Replace(strExc(0), vbCrLf, "") & vbCrLf & "Y□ N▓"
               strExc(0) = strExc(0) & vbCrLf & "Y□ N▓"
            End If
            Me.txt1(4).Text = strExc(0) & vbCrLf & Me.txt1(4).Text
         End If
      End If
   End If
End Sub

'Add By Sindy 2013/10/1 抓取主旨
Private Sub GetPaperMain()
Dim strCaseName As String
   
   If Trim(txtCaseName(0)) <> "" Then
      strCaseName = Trim(txtCaseName(0))
   ElseIf Trim(txtCaseName(1)) <> "" Then
      strCaseName = Trim(txtCaseName(1))
   Else
      strCaseName = Trim(txtCaseName(2))
   End If
   If cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "105" _
      Or cp(10) = "109" Or cp(10) = "110" Or cp(10) = "112" Or cp(10) = "113" _
      Or cp(10) = "114" Or cp(10) = "115" Or cp(10) = "118" Or cp(10) = "301" _
      Or cp(10) = "302" Or cp(10) = "303" Or cp(10) = "304" Or cp(10) = "305" _
      Or cp(10) = "306" Or cp(10) = "307" Or cp(10) = "803" Then
      Me.txt1(0) = "為「" & strCaseName & "」" & GetNationName(m_Country, 0) & lblCP10 & "專利案提出申請。"
   Else
      Me.txt1(0).Text = "「" & strCaseName & "」" & GetNationName(m_Country, 0) & Trim(lblPA08.Caption) & "專利之" & lblCP10
   End If
End Sub

Private Sub ClearData()
   m_EEP02 = Empty
   txtEEP03 = Empty
   txtEEP03_2 = Empty
   CboEEP04.Text = Empty 'Add By Sindy 2023/1/18
   CboEEP04.Clear
   CboEEP05.Text = Empty 'Add By Sindy 2023/1/18
   CboEEP05.Clear
   ChkEMail.Value = 0
   ChkEMail.Visible = False
   txtEEP10 = Empty
   txtEEP10_2 = Empty
   txtEEP08 = Empty
   lstAtt(0).Clear
   'lstAtt(1).Clear 'Modify By Sindy 2018/10/17 Mark:不可清除,一進入此作業,存卷有資料是都可以查看
   Me.cmdOpenAtt(0).Enabled = False
   Me.cmdSelect(0).Enabled = False
   Me.cmdSaveAtt(0).Enabled = False
   Me.cmdAddAtt(0).Enabled = False
   CmdF21(0).Enabled = False 'Add By Sindy 2025/10/28
   Me.cmdRemAtt(0).Enabled = False
   cmdDel.Enabled = False 'Add By Sindy 2013/10/3
   cmdAddAttDB(0).Enabled = False 'Add By Sindy 2013/10/24
   cmdRemAttDB(0).Enabled = False 'Add By Sindy 2013/10/24
   cmdAddAttDB(1).Enabled = False 'Add By Sindy 2013/10/24
   cmdRemAttDB(1).Enabled = False 'Add By Sindy 2013/10/24
   cmdCaseMap.Visible = False 'Add By Sindy 2017/8/30 不顯示多國案鈕
   Check1.Visible = False
   ChkEP11.Visible = False 'Add By Sindy 2018/9/20
   
   cmdManyCase.Visible = False 'Add By Sindy 2018/9/26
   cmdManyCase.Tag = "" 'Add By Sindy 2018/10/24
   m_RetrunRecv = "" 'Add By Sindy 2018/10/30
   m_RetrunRecvCnt = 0 'Add By Sindy 2018/10/30
   m_RetrunRecvSub = "" 'Add By Sindy 2020/10/12
End Sub

'Add By Sindy 2014/1/14 顯示其他國外案
Private Sub Cmd1_Click(Index As Integer)
Dim iMouse As Integer
   
   'Add By Sindy 2024/2/7
   If bolFCPFlow = True Then
      iMouse = Screen.MousePointer
      Me.Hide
      Screen.MousePointer = vbHourglass
      frm100101_h.SetParent Me
      frm100101_h.Show
      frm100101_h.cmdOK(3).Visible = False '下一筆按鍵不顯示
      frm100101_h.KeyString = Pub_RplStr(lblCaseNo.Caption)
      '區分相似案
      If Index = 1 Then
          frm100101_h.SearchKind = "相似案"
      Else
          frm100101_h.SearchKind = "本所案號"
      End If
      frm100101_h.StrMenu
      Screen.MousePointer = iMouse
   Else
   '2024/2/7 END
      '專利=專利相關案
      If bolPAFlow = True Then
         iMouse = Screen.MousePointer
         Me.Hide
         Screen.MousePointer = vbHourglass
         frm090201_2_1.SetParent Me
         frm090201_2_1.Show
         Call frm090201_2_1.StrMenu(lblCaseNo.Caption)
         Screen.MousePointer = iMouse
      'Add By Sindy 2018/7/26
      '商標=商品名稱維護
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True
      ElseIf bolTMFlow = True Or bolCFTFlow = True Then
         frm03010303_04.Hide
         Set frm03010303_04.UpForm = Me
         frm03010303_04.TGKey = PField(1) & "-" & PField(2) & "-" & PField(3) & "-" & PField(4)
         frm03010303_04.AllClass = tm(9)
         frm03010303_04.Caption = "商品及服務資料"
         frm03010303_04.Label2.Visible = True
         'frm03010303_04.cmdOK(2).Visible = True
         Me.Hide
         frm03010303_04.QueryData
         frm03010303_04.Show vbModal '強制回應表單
         Call Cmd1_LostFocus(0) 'Add By Sindy 2024/6/13
      '2018/7/26 END
      End If
   End If
End Sub
'Add By Sindy 2024/6/13
Public Sub Cmd1_LostFocus(Index As Integer)
Dim rsTmp As New ADODB.Recordset
   
   If Index = 0 And cmd1(0).Caption = "商品名稱維護" Then
      '檢查商標描述中文,英文是否有輸入
      If tm(72) <> "" Then '特殊商標
         If tm(137) = "" And tm(138) = "" Then
            cmdOK(3).BackColor = &H8080FF '紅色
         Else
            cmdOK(3).BackColor = &HC0FFC0 '淡綠色
         End If
      End If
      
      strSql = "select tg05 from Tmgoods" & _
               " where tg01='" & PField(1) & "' and tg02='" & PField(2) & "' and tg03='" & PField(3) & "' and tg04='" & PField(4) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      '無商品類別
      If rsTmp.RecordCount = 0 Then
         cmd1(0).BackColor = &H8080FF '紅色
      Else
         rsTmp.Close
         strSql = "select tg05 from Tmgoods" & _
                  " where tg01='" & PField(1) & "' and tg02='" & PField(2) & "' and tg03='" & PField(3) & "' and tg04='" & PField(4) & "'" & _
                  " and tg06||tg07||tg08||tg15||tg16||tg17 is null"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         '有商品類別,無輸入商品資料
         If rsTmp.RecordCount > 0 Then
            cmd1(0).BackColor = &H8080FF '紅色
         Else
            cmd1(0).BackColor = &HC0FFC0 '淡綠色
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Sub

'新增下一流程
Public Sub cmdAdd_Click()
Dim rsA As New ADODB.Recordset
   
   'Add By Sindy 2025/10/14
   If Frame1Big.Visible = True Then
      Call cmdClose_Click
      Exit Sub
   End If
   '2025/10/14 END
   
   'Add By Sindy 2021/10/27 王副總:專利國內部專利案之歷程，每日22:00至隔日06:00，請設定為禁止操作狀態
   'Modify By Sindy 2022/3/8 人事處:設定每日22:00至隔日06:00及週日全日為禁止操作狀態，以符合勞基法規定(套用至全所人員)
   'If bolPAFlow = True Then
      If (Left(Format(ServerTime, "000000"), 4) >= "2200" Or Left(Format(ServerTime, "000000"), 4) <= "0600") Or _
         Weekday(Format(strSrvDate(1), "####-##-##")) = 1 Then
         'MsgBox "專利國內部專利案之歷程，每日22:00至隔日06:00，禁止操作！"
         MsgBox "案件歷程，每日22:00至隔日06:00及[週日]全日，禁止操作！"
         Exit Sub
      End If
   'End If
      
   'Add By Sindy 2020/10/12 是否多案操作待回覆中
   If rsA.State <> adStateClosed Then rsA.Close
   strExc(0) = "select eep01,eep04,ac03,cp01,cp02,cp03,cp04 From empelectronprocess,caseprogress,allcode" & _
               " where instr(eep15,'" & m_EEP01 & "')>0 and eep09='Y' and eep01=cp09(+)" & _
               " and ac01='09' And eep04=ac02(+)"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      '排除操作歷程的案號,其他案才需要顯示此訊息
      If rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04") <> lblCaseNo.Caption Then
         'Add By Sindy 2022/6/22 程序組要在共同查詢操作查名結果 ex:T-239802
         If Not (txtNote.Visible = True And UCase(m_PrevForm.Name) = "FRM100101_2" And Pub_StrUserSt03 = "P22") Then '程序組-共同查詢
         '2022/6/22 END
            MsgBox "此案與 " & rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & IIf(rsA.Fields("cp03") & rsA.Fields("cp04") = "000", "", "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04")) & _
                   " 屬多案操作的歷程。" & vbCrLf & vbCrLf & _
                   "正『" & rsA.Fields("ac03") & "』中，等待回覆..."
            rsA.Close
            Exit Sub
         End If
      End If
   End If
   rsA.Close
   '2020/10/12 END
   
'   Me.cmdAdd.Visible = False
'   Me.cmdCancel.Visible = True
'   Me.cmdSend.Enabled = True
'   m_EditMode = 1 '新增
'   Frame4.Visible = False: CboCP10.Locked = False
'   cmdMail.Visible = False 'Add By Sindy 2018/8/29
   Call SetStatusCombo
End Sub

'Add By Sindy 2025/8/20
Private Sub CmdCalendar_Click()
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

'取消
'Modify By Sindy 2025/8/5 Private Sub ==> Public Sub
Public Sub cmdCancel_Click()
   'Add By Sindy 2025/10/14
   If Frame1Big.Visible = True Then
      Call cmdClose_Click
      Exit Sub
   End If
   '2025/10/14 END
   
   'If Val(dblPrevRow) = 0 And GRD1.Rows - 1 > 0 And GRD1.TextMatrix(1, 0) <> "" Then dblPrevRow = 1 'Add By Sindy 2017/9/15
   cmdSend.Caption = "送出(&O)" 'Add By Sindy 2018/8/29
   Frame4.Visible = False 'Add By Sindy 2013/9/23 案件性質/會稿方式
   Call SetTxtLpNote(True) 'Add By Sindy 2020/9/29
   'Modify By Sindy 2013/9/4
   If dblPrevRow > 0 Then
      Call ReadData(False)
   Else
   '2013/9/4 END
      Call ClearData
      Call SetCtrlReadOnly(False)
   End If
End Sub

'Add By Sindy 2015/3/15 確認是否會(圖/文)完成
Private Sub GetChkEP06()
Dim strUpdDate As String, strUpdTime As String
Dim intMaxEEP02 As Integer
Dim strCaseName As String
Dim strEEP08 As String
Dim strEEP02 As String, strEEP06 As String
Dim strTo As String 'Add By Sindy 2016/5/25
   
On Error GoTo ErrHand
   
   If Trim(txtCaseName(0)) <> "" Then
      strCaseName = Trim(txtCaseName(0))
   ElseIf Trim(txtCaseName(1)) <> "" Then
      strCaseName = Trim(txtCaseName(1))
   Else
      strCaseName = Trim(txtCaseName(2))
   End If
   
   strSql = "select eep02,eep06 From empelectronprocess where eep01='" & m_EEP01 & "' and eep04='" & EMP_圖完 & "' order by eep02 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      If RsTemp.RecordCount > 0 Then
         strEEP02 = RsTemp.Fields(0)
         strEEP06 = RsTemp.Fields(1)
      End If
   Else
      MsgBox "資料有誤，查無(圖/文)完歷程！"
      Exit Sub
   End If
   
   'Modify By Sindy 2022/10/7 + 文
   If MsgBox("是否會(圖/文)完成？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
      cnnConnection.BeginTrans
      
      '以「圖完」的日期建為新的「齊備日」
      strSql = "update engineerprogress set ep06=" & strEEP06 & " where ep02='" & m_EEP01 & "'"
      cnnConnection.Execute strSql
      
      '記錄已處理過
      strSql = "update empelectronprocess set eep11='會(圖/文)完成;'||eep11 where eep01='" & m_EEP01 & "' and eep02=" & strEEP02
      cnnConnection.Execute strSql
      
      cnnConnection.CommitTrans
      
      If intReceiveKind = 0 Then '0.承辦人工作進度
         m_PrevForm.txt1(2) = strEEP06 - 19110000
      End If
      
      cmdAdd.Visible = True
      Frame3.Visible = False
      Exit Sub
   Else
      strEEP08 = InputBox("請輸入不接受會(圖/文)完成的理由：")
      If strEEP08 = "" Then
         '取消離開
         Exit Sub
      End If
   End If
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   '記錄已處理過
   strSql = "update empelectronprocess set eep11='不承認會(圖/文)完成;'||eep11 where eep01='" & m_EEP01 & "' and eep02=" & strEEP02
   cnnConnection.Execute strSql
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   '檢查是否為工作天
   If Not ChkWorkDay(strUpdDate) Then
      strUpdDate = CompWorkDay(1, strUpdDate, 0)
   End If
   strUpdDate = ChangeWStringToTString(strUpdDate) '轉換成民國日期
   
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
   'Modify By Sindy 2023/12/18 +,eep16
   strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep12,eep16) values(" & _
            CNULL(m_EEP01) & "," & (intMaxEEP02 + 1) & "," & CNULL(m_FlowUserNum) & "," & _
            CNULL(EMP_不自動更新齊備日) & "," & CNULL(Trim(Left(m_SPMan, 6))) & "," & DBDATE(strUpdDate) & "," & _
            strUpdTime & "," & CNULL(ChgSQL(strEEP08)) & ",'" & m_EEP12 & "','" & m_EEP16 & "')"
   cnnConnection.Execute strSql
   'Modify By Sindy 2023/12/18 +,eep16
   strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep09,eep12,eep16) values(" & _
            CNULL(m_EEP01) & "," & (intMaxEEP02 + 2) & "," & CNULL(m_FlowUserNum) & "," & _
            CNULL(EMP_會圖) & "," & CNULL(Trim(Left(m_SPMan, 6))) & "," & DBDATE(strUpdDate) & "," & _
            strUpdTime & "," & CNULL("系統自動新增歷程") & ",'Y','" & m_EEP12 & "','" & m_EEP16 & "')"
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   'Modify By Sindy 2023/12/14 杜燕文協理請作,主旨加申請國家
   strSubject = Replace(lblCaseNo, "-0-00", "") & "(" & lblPA09 & ")(核會流程)-->(" & (intMaxEEP02 + 1) & "," & (intMaxEEP02 + 2) & ")工程師認為圖式有問題，需要再確認；不自動更新齊備日"
   'Modify By Sindy 2018/4/2
   'FC代理人來台
   strExc(10) = ""
   If m_PA75 <> "" And m_Country = "000" Then
      strExc(10) = "貴方卷號：" & m_PA77 & vbCrLf
   End If
   'Modify By Sindy 2023/12/15 杜燕文協理請作,內文加申請國家
   strContent = "當月目次：" & m_EP01 & vbCrLf & strExc(10) & _
                "本所案號：" & lblCaseNo & vbCrLf & _
                "案件名稱：" & strCaseName & vbCrLf & _
                "申請國家：" & lblPA09 & vbCrLf & _
                "案件性質：" & m_CP10Nm & vbCrLf & _
                "流程狀態：不自動更新齊備日、會(圖/文)" & vbCrLf & _
                "原　　因：" & strEEP06
   '發給智權人員
   strTo = PUB_ChkPersonToGetEEP03(m_EEP01, Trim(Left(m_SPMan, 6)), EMP_圖完) 'Add By Sindy 2016/5/25 檢查人員是否離職,改抓歷程的發送者
   'Modify By Sindy 2018/5/4
   If bolPAFlow = True Then
      'Modify By Sindy 2023/4/24 5/1加入游(73022)，5/11取消王副總
      'Modified by Morgan 2025/2/20 73022->pub_PMan, +99050
      'If Val(strSrvDate(1)) >= 20230501 Then
      '   PUB_SendMail strUserNum, strTo & IIf(Val(strSrvDate(1)) < 20230511, ";71011", "") & ";73022"", m_EEP01, strSubject, strContent 'Add By Sindy 2022/2/21 + m_EEP01
      'Else
      ''2023/4/24 END
      '   PUB_SendMail strUserNum, strTo & ";71011", m_EEP01, strSubject, strContent 'Add By Sindy 2022/2/21 + m_EEP01
      'End If
      pub_PMan = Pub_GetSpecMan("專利處特定編號")
      PUB_SendMail strUserNum, strTo & ";" & pub_PMan & ";99050", m_EEP01, strSubject, strContent
      'end 2025/2/20
   Else
      PUB_SendMail strUserNum, strTo, m_EEP01, strSubject, strContent 'Add By Sindy 2022/2/21 + m_EEP01
   '2018/5/4 END
   End If
   Screen.MousePointer = vbDefault
   
   cmdAdd.Visible = True
   Frame3.Visible = False
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
End Sub

'Add By Sindy 2013/9/17 確認會稿完成日
Private Sub cmdChkEP08_Click()
Dim strUpdDate As String, strUpdTime As String
Dim intMaxEEP02 As Integer
Dim strEEP08 As String
Dim strCaseName As String
Dim strTo As String 'Add By Sindy 2016/5/25
   
   'Add By Sindy 2016/3/15
   'Modify By Sindy 2022/10/7 + 文
   If cmdChkEP08.Caption = "是否會(圖/文)完成" Then
      Call GetChkEP06
      Exit Sub
   End If
   '2016/3/15 END
   
On Error GoTo ErrHand
   
   'Add By Sindy 2013/10/1
   If Trim(txtCaseName(0)) <> "" Then
      strCaseName = Trim(txtCaseName(0))
   ElseIf Trim(txtCaseName(1)) <> "" Then
      strCaseName = Trim(txtCaseName(1))
   Else
      strCaseName = Trim(txtCaseName(2))
   End If
   '2013/10/1 END
   
   If MsgBox("是否接受智權同仁輸入的會稿完成日？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
      '更新會稿完成日=智權人員會稿完成日
'      strSql = "update EngineerProgress " & _
'               "set EP08=EP38 " & _
'               "where EP02='" & Trim(lblCP09) & "'"
'      cnnConnection.Execute strSql
      If Val(m_EP38) = 0 Then
         MsgBox "無智權人員會完日，無法更新!!!"
         Exit Sub
      End If
      
      'Modify By Sindy 2024/7/11 mark,不會使用到的程式
'      If bolTMFlow = True Or bolOtherFlow = True Then
'         strSql = "update engineerprogress set ep08=" & CNULL(m_EP38, True) & " where ep02='" & m_EEP01 & "'"
'         cnnConnection.Execute strSql
'      Else
         UpdateEp08 m_EEP01, m_EP38 '更新相關會稿完成日資料
         PUB_AskUpdateRelationCase m_EEP01 'Added by Morgan 2015/5/25
'      End If
      
      PUB_SendMailCache '發郵件 (相關會稿完成日的郵件)
      If intReceiveKind = 0 Then '0.承辦人工作進度
         m_PrevForm.txt1(7) = m_EP38 - 19110000
      End If
      '2013/10/15 END
      
      cmdAdd.Visible = True
      Frame3.Visible = False
      Exit Sub
   Else
      strEEP08 = InputBox("請輸入不接受會完日的理由：")
      If strEEP08 = "" Then
         '取消離開
         Exit Sub
      End If
   End If
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   '檢查是否為工作天
   If Not ChkWorkDay(strUpdDate) Then
      strUpdDate = CompWorkDay(1, strUpdDate, 0)
   End If
   strUpdDate = ChangeWStringToTString(strUpdDate) '轉換成民國日期
   
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
   'Modify By Sindy 2023/12/18 +,eep16
   strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep12,eep16) values(" & _
            CNULL(m_EEP01) & "," & (intMaxEEP02 + 1) & "," & CNULL(m_FlowUserNum) & "," & _
            CNULL(EMP_不自動更新會完日) & "," & CNULL(Trim(Left(m_SPMan, 6))) & "," & DBDATE(strUpdDate) & "," & _
            strUpdTime & "," & CNULL(ChgSQL(strEEP08)) & ",'" & m_EEP12 & "','" & m_EEP16 & "')"
   cnnConnection.Execute strSql
      
   cnnConnection.CommitTrans
   
   'Modify By Sindy 2023/12/14 杜燕文協理請作,主旨加申請國家
   strSubject = Replace(lblCaseNo, "-0-00", "") & "(" & lblPA09 & ")(核會流程)-->(" & (intMaxEEP02 + 1) & ")不自動更新會完日"
   'Modify By Sindy 2018/4/2
   'FC代理人來台
   strExc(10) = ""
   If m_PA75 <> "" And m_Country = "000" Then
      strExc(10) = "貴方卷號：" & m_PA77 & vbCrLf
   End If
   'Modify By Sindy 2023/12/15 杜燕文協理請作,內文加申請國家
   strContent = "當月目次：" & m_EP01 & vbCrLf & strExc(10) & _
                "本所案號：" & lblCaseNo & vbCrLf & _
                "案件名稱：" & strCaseName & vbCrLf & _
                "申請國家：" & lblPA09 & vbCrLf & _
                "案件性質：" & m_CP10Nm & vbCrLf & _
                "流程狀態：不自動更新會完日" & vbCrLf & _
                "原　　因：" & strEEP08
   '發給智權人員
   strTo = PUB_ChkPersonToGetEEP03(m_EEP01, Trim(Left(m_SPMan, 6)), EMP_會完) 'Add By Sindy 2016/5/25 檢查人員是否離職,改抓歷程的發送者
   'Modify By Sindy 2018/5/4
   If bolPAFlow = True Then
      'Modify By Sindy 2023/4/24 5/1加入游(73022)，5/11取消王副總
      'Modified by Morgan 2025/2/20 73022->pub_PMan, +99050
      'If Val(strSrvDate(1)) >= 20230501 Then
      '   PUB_SendMail strUserNum, strTo & IIf(Val(strSrvDate(1)) < 20230511, ";71011", "") & ";73022", m_EEP01, strSubject, strContent 'Add By Sindy 2022/2/21 + m_EEP01
      'Else
      ''2023/4/24 END
      '   PUB_SendMail strUserNum, strTo & ";71011", m_EEP01, strSubject, strContent 'Add By Sindy 2022/2/21 + m_EEP01
      'End If
      pub_PMan = Pub_GetSpecMan("專利處特定編號")
      PUB_SendMail strUserNum, strTo & ";" & pub_PMan & ";99050", m_EEP01, strSubject, strContent
      'end 2025/2/20
   Else
      PUB_SendMail strUserNum, strTo, m_EEP01, strSubject, strContent 'Add By Sindy 2022/2/21 + m_EEP01
   '2018/5/4 END
   End If
   Screen.MousePointer = vbDefault
   
   cmdAdd.Visible = True
   Frame3.Visible = False
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
End Sub

'Add By Sindy 2013/10/3
Private Sub cmdDel_Click()
Dim strEEP04 As String, strEEP02 As String 'Add By Sindy 2014/1/14
Dim strEEP11 As String 'Add By Sindy 2018/10/26
Dim bolClearCP27 As Boolean 'Add By Sindy 2018/11/20
Dim strCP154 As String 'Add By Sindy 2018/12/17
Dim strEEP15 As String 'Add By Sindy 2020/10/15
'Add By Sindy 2024/4/12
Dim strEEP05 As String
Dim strOldEED06 As String, strPreviousEED04 As String
'2024/4/12 END
   
On Error GoTo ErrHand
   
   'Add By Sindy 2018/2/23
   'Modify By Sindy 2018/11/20
   'strSql = "select * from caseprogress where cp09 ='" & m_EEP01 & "' and cp158=0 and cp159=0"
   strSql = "select * from caseprogress where cp09 ='" & m_EEP01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then
      MsgBox "無此文號資料!!", vbExclamation, MsgText(5)
      Exit Sub
   Else
      strCP154 = "" & RsTemp.Fields("CP154") '發文室非主管機關人員
      '商標,已發文未取消收文,未經發文室,並且為發文歸檔的歷程才能刪除
      'Modify By Sindy 2021/3/4 And Val(RsTemp.Fields("cp158")) > 0 =>取消 +  Or "" & RsTemp.Fields("cp154") = "QPGMR")
      If (bolTMFlow = True And Val(RsTemp.Fields("cp159")) = 0 And _
         ("" & RsTemp.Fields("cp127") = "" Or strCP154 = "QPGMR") And Left(CboEEP04.Text, 2) = EMP_發文歸檔) Then
         '可刪除,程式往下執行(未經發文室,要改送發文室)
         bolClearCP27 = True
      ElseIf Val(RsTemp.Fields("cp159")) > 0 Then
         MsgBox "此案已取消收文，不可刪除歷程！", vbExclamation, "重要訊息！"
         Exit Sub
      ElseIf Val(RsTemp.Fields("cp158")) > 0 Then
         'Modify By Sindy 2020/3/20
         'MsgBox "此案已發文或已取消收文不可刪除歷程!!", vbExclamation, MsgText(5)
         'Add By Sindy 2020/10/15
         'Modify By Sindy 2023/11/23 +And bolTMFlow = True
         If Left(CboEEP04.Text, 2) = EMP_發文歸檔 And strCP154 <> "QPGMR" And bolTMFlow = True Then
            MsgBox "發文室已發文，請該人員先通知發文室取消發文後，再通知電腦中心刪除歷程！", vbExclamation, "重要訊息！"
            Exit Sub
         '2020/10/15 END
         Else
            If MsgBox("此案「已發文」確定要刪除歷程嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Exit Sub
            Else
               'Add By Sindy 2023/11/23 外專目前只有電話聯絡單會做發文歸檔
               If Left(CboEEP04.Text, 2) = EMP_發文歸檔 And bolFCPFlow = True Then
                  bolClearCP27 = True
               End If
               '2023/11/23 END
            End If
         End If
         '2020/3/20 END
      End If
   End If
   '2018/2/23 END
   
   'Add By Sindy 2014/1/14
   strSql = "select * from empelectronprocess where eep01 ='" & m_EEP01 & "' and eep02=" & m_EEP02
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strEEP04 = "" & RsTemp.Fields("eep04")
      strEEP05 = "" & RsTemp.Fields("eep05") 'Add By Sindy 2024/4/12
      strEEP11 = "" & RsTemp.Fields("eep11")
      'Add By Sindy 2020/10/15 多案單筆歷程時,才需要一併恢復資料
      If InStr(strEEP11, "多案單筆歷程") > 0 Then
         strEEP15 = "" & RsTemp.Fields("eep15")
      Else
         strEEP15 = ""
      End If
      '2020/10/15 END
   End If
   'Add By Sindy 2024/4/15 欲恢復 上一歷程順序 為待回覆
   If InStr(strEEP11, "上一歷程順序:") > 0 Then
      strEEP02 = Val(Mid(strEEP11, InStr(strEEP11, "上一歷程順序:") + Len("上一歷程順序:"), 3))
   End If
   '2024/4/15 END
   
   If MsgBox("確定要刪除【順序為 " & m_EEP02 & " - " & GRD1.TextMatrix(dblPrevRow, 4) & "】流程及附件檔嗎？" & vbCrLf & vbCrLf & _
             "《請確認需要人工拿掉的相關欄位日期》" & vbCrLf & vbCrLf & _
             "《請先詢問此流程是否有沿用上一流程附件，若有，告知人員先自行下載附件，再通知電腦中心刪除》" & vbCrLf & vbCrLf & _
             "《若需要恢復為待核/會/判...狀態時,最近一筆送核/會/判...流程的EEP09欄位必須更新為Y》", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
      'Modify By Sindy 2016/3/15 +EMP_圖修, EMP_圖完
      'Modify By Sindy 2023/10/30 +EMP_送件, EMP_排版完成, EMP_核稿分案, EMP_轉檔完成, EMP_程序退回
      'Modify By Sindy 2024/4/15 + Or Trim(strEEP02) <> ""
      If strEEP04 = EMP_核修 Or _
         strEEP04 = EMP_核完 Or _
         strEEP04 = EMP_會修 Or _
         strEEP04 = EMP_會完 Or _
         strEEP04 = EMP_繪圖判發 Or _
         strEEP04 = EMP_退回 Or _
         strEEP04 = EMP_判發 Or _
         strEEP04 = EMP_草修 Or _
         strEEP04 = EMP_草核完 Or _
         strEEP04 = EMP_圖修 Or _
         strEEP04 = EMP_圖完 Or _
         strEEP04 = EMP_送件 Or _
         strEEP04 = EMP_排版完成 Or _
         strEEP04 = EMP_核稿分案 Or _
         strEEP04 = EMP_轉檔完成 Or _
         strEEP04 = EMP_程序退回 Or _
         (strEEP04 = EMP_發文歸檔 And bolFCPFlow = True) Or _
         Trim(strEEP02) <> "" Then
         
         'Add By Sindy 2024/4/15
         If strEEP02 = "" Then
         '2024/4/15 END
            strEEP02 = UCase(InputBox("是否有需要恢復為待英(日)核／核／會／判等狀態？" & vbCrLf & vbCrLf & _
                                      "歷程狀態繁多，請參 【常變數：EMP_需等待回覆的狀態】" & vbCrLf & vbCrLf & _
                                      "若無，不用輸入" & vbCrLf & _
                                      "若有，請輸入要恢復（待xx中）流程的順序號碼" & vbCrLf & _
                                      "輸入【 X 】視為取消，先不刪除。"))
         End If
         If Trim(strEEP02) = "X" Then
            Exit Sub
         ElseIf Trim(strEEP02) <> "" Then
            strSql = "select * from empelectronprocess where eep01 ='" & m_EEP01 & "' and eep02=" & Val(strEEP02)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               'Modify By Sindy 2016/3/15 +EMP_會圖
               'Modify By Sindy 2023/11/17 +EMP_送排版, EMP_送核稿分案, EMP_送轉檔, EMP_程序送判
'               If RsTemp.Fields("eep04") <> EMP_會圖 And _
'                  RsTemp.Fields("eep04") <> EMP_送英核 And _
'                  RsTemp.Fields("eep04") <> EMP_送核 And _
'                  RsTemp.Fields("eep04") <> EMP_送會 And _
'                  RsTemp.Fields("eep04") <> EMP_草核 And _
'                  RsTemp.Fields("eep04") <> EMP_墨完 And _
'                  RsTemp.Fields("eep04") <> EMP_送判 And _
'                  RsTemp.Fields("eep04") <> EMP_轉回 And _
'                  RsTemp.Fields("eep04") <> EMP_送排版 And _
'                  RsTemp.Fields("eep04") <> EMP_送核稿分案 And _
'                  RsTemp.Fields("eep04") <> EMP_送轉檔 And _
'                  RsTemp.Fields("eep04") <> EMP_程序送判 Then
               If InStr(EMP_需等待回覆的狀態, RsTemp.Fields("eep04")) = 0 Then
                  MsgBox "此筆 ( " & strEEP02 & " ) 流程無須恢復為待回覆, 請重新確認!!"
                  Exit Sub
               End If
            Else
               strEEP02 = ""
            End If
         End If
      End If
      '2014/1/14 END
      
      Screen.MousePointer = vbHourglass
      cnnConnection.BeginTrans
      
      'Add By Sindy 2018/10/26 恢復判發完成日
      If strEEP04 = EMP_判發 Then
         'Modify By Sindy 2019/1/30
         strSql = "select eep01 from empelectronprocess where eep01 ='" & m_EEP01 & "' and eep02<>" & Val(m_EEP02) & " and eep04='" & EMP_判發 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
         '2019/1/30 END
            'Add By Sindy 2020/10/15
            If strEEP15 <> "" Then '多案單筆歷程
               strSql = "update engineerprogress set" & _
                              " EP42=null" & _
                        " where ep02 in('" & Replace(strEEP15, ",", "','") & "')"
            Else
            '2020/10/15 END
               strSql = "update engineerprogress set" & _
                              " EP42=null" & _
                        " where ep02='" & m_EEP01 & "'"
            End If
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
         'Add By Sindy 2019/10/25 送核直接判發,要恢復
         If InStr(strEEP11, "流程狀態:" & EMP_送核) > 0 Then
            'Add By Sindy 2020/10/15
            If strEEP15 <> "" Then '多案單筆歷程
               strSql = "update engineerprogress set" & _
                              " EP39=null" & _
                        " where ep02 in('" & Replace(strEEP15, ",", "','") & "')"
            Else
            '2020/10/15 END
               strSql = "update engineerprogress set" & _
                              " EP39=null" & _
                        " where ep02='" & m_EEP01 & "'"
            End If
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
         '2019/10/25 END
      '恢復核稿完成日
      ElseIf strEEP04 = EMP_核完 And InStr(strEEP11, "流程狀態:" & EMP_送核) > 0 Then
         'Modify By Sindy 2019/1/30
         strSql = "select eep01 from empelectronprocess where eep01 ='" & m_EEP01 & "' and eep02<>" & Val(m_EEP02) & " and eep04='" & EMP_核完 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
         '2019/1/30 END
            'Add By Sindy 2020/10/15
            If strEEP15 <> "" Then '多案單筆歷程
               strSql = "update engineerprogress set" & _
                              " EP39=null" & _
                        " where ep02 in('" & Replace(strEEP15, ",", "','") & "')"
            Else
            '2020/10/15 END
               strSql = "update engineerprogress set" & _
                              " EP39=null" & _
                        " where ep02='" & m_EEP01 & "'"
            End If
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
      '2018/10/26 END
      'Add By Sindy 2024/4/12 流程狀態:43,原收受者:73035
      ElseIf strEEP04 = EMP_交辦 And InStr(strEEP11, "流程狀態:") > 0 And InStr(strEEP11, "原收受者:") > 0 Then
         strOldEED06 = Mid(strEEP11, InStr(strEEP11, "原收受者:") + Len("原收受者:"), 5)
         strPreviousEED04 = Mid(strEEP11, InStr(strEEP11, "流程狀態:") + Len("流程狀態:"), 2)
         '承辦單子簽核流程檔
         strSql = "select eep01 from empelectronprocess" & _
                  " where eep01 ='" & m_EEP01 & "' and eep04='" & strPreviousEED04 & "'" & _
                  " and eep09='Y' and eep05='" & strEEP05 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            '電子承辦單內容
            strSql = "select eed01 from empelectrondata where eed01 ='" & m_EEP01 & "' and eed06='" & strEEP05 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               '恢復原收受者
               strSql = "update empelectronprocess set eep05='" & strOldEED06 & "'" & _
                        " where eep01 ='" & m_EEP01 & "' and eep04='" & strPreviousEED04 & "'" & _
                        " and eep09='Y' and eep05='" & strEEP05 & "'"
               cnnConnection.Execute strSql, intI
               strSql = "update empelectrondata set eed06='" & strOldEED06 & "'" & _
                        " where eed01 ='" & m_EEP01 & "'"
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
      
      If Trim(strEEP02) <> "" Then
         'Add By Sindy 2014/1/14
         'Modify By Sindy 2016/6/7 + ,eep13='Y'
         strSql = "update empelectronprocess set eep09='Y',eep13='Y' where eep01 ='" & m_EEP01 & "' and eep02=" & Val(strEEP02)
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         '2014/1/14 END
      End If
      
      '清除發文日:商標,已發文未取消收文,未經發文室,並且為發文歸檔的歷程才能刪
      If bolClearCP27 = True Then
         'Add By Sindy 2020/10/15
         If strEEP15 <> "" Then '多案單筆歷程
            strSql = "update caseprogress set cp27=null where cp09 in('" & Replace(strEEP15, ",", "','") & "')"
         Else
         '2020/10/15 END
            strSql = "update caseprogress set cp27=null where cp09='" & m_EEP01 & "'"
         End If
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         
         'Modify By Sindy 2023/11/23 +And bolTMFlow = True
         If Left(CboEEP04.Text, 2) = EMP_發文歸檔 And bolTMFlow = True Then 'And strCP154 = "QPGMR" ex:T-188909
            strSql = "select ep02,ep11 from engineerprogress where ep02 ='" & m_EEP01 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If RsTemp.Fields("ep11") = "Y" Then '要通知客戶,經發文室
                  '發文室非主管機關人員,日期,時間 (要經發文室)
                  'Add By Sindy 2020/10/15
                  If strEEP15 <> "" Then '多案單筆歷程
                     strSql = "update caseprogress set cp154=null,cp127=null,cp128=null where cp09 in('" & Replace(strEEP15, ",", "','") & "')"
                  Else
                  '2020/10/15 END
                     strSql = "update caseprogress set cp154=null,cp127=null,cp128=null where cp09='" & m_EEP01 & "'"
                  End If
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql
               End If
               'Add By Sindy 2021/9/1 清除信函進度; CB0051310 T-230060 發文歸檔刪除時,要連同不通知信函一併恢復
               If strEEP15 <> "" Then '多案單筆歷程
                  strSql = "update letterprogress set lp06=" & CNULL(Trim(Left(m_SPMan, 6))) & ",lp07=0,lp10=null,lp11=null,lp12=null where LP01 in('" & Replace(strEEP15, ",", "','") & "')"
               Else
                  strSql = "update letterprogress set lp06=" & CNULL(Trim(Left(m_SPMan, 6))) & ",lp07=0,lp10=null,lp11=null,lp12=null where LP01='" & m_EEP01 & "'"
               End If
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql
               '2021/9/1 END
            End If
         End If
      End If
      'Add By Sindy 2025/4/22
      If Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
         '以防重覆歸卷
         If DelAttFile_PDF(lblCaseNo.Caption, m_EEP01, "", "S", True) = False Then GoTo ErrHand
         If DelAttFile_File(lblCaseNo.Caption, m_EEP01, "", "S", True) = False Then GoTo ErrHand
      End If
      '2025/4/22 END
      
'      'Add By Sindy 2020/10/15
'      If strEEP15 <> "" Then '多案單筆歷程
'         If bolTMFlow = True And _
'            (Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送 Or _
'             Left(CboEEP04.Text, 2) = EMP_發文歸檔) Then
'            strSql = "update caseprogress set" & _
'                     " cp163=null" & _
'                     " where cp09 in('" & Replace(strEEP15, ",", "','") & "')"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql, intI
'         End If
'      End If
'      '2020/10/15 END
      
      strSql = "delete from empelectronprocess where eep01 ='" & m_EEP01 & "' and eep02=" & m_EEP02
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      
      'Add By Sindy 2016/6/7 刪除歷程時,要考慮到那筆歷程應該顯示出來
      'Add By Sindy 2018/8/3 EMP_不自動更新會完日 & "," & EMP_附加流程 & "," & EMP_不自動更新齊備日 ==> EMP_流程控制除外的狀態 & ",'" & EMP_發文歸檔 & "'
      'strSql = "select nvl(max(eep02),0) from empelectronprocess where eep01 ='" & m_EEP01 & "' and eep04 not in(" & EMP_不自動更新會完日 & "," & EMP_附加流程 & "," & EMP_不自動更新齊備日 & ")"
      strSql = "select nvl(max(eep02),0) from empelectronprocess where eep01 ='" & m_EEP01 & "' and eep04 not in(" & EMP_流程控制除外的狀態 & ")"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 And Val("" & RsTemp.Fields(0)) > 0 Then
         If (Trim(strEEP02) <> "" And Val(strEEP02) < Val("" & RsTemp.Fields(0))) Or _
            (strEEP02 = "" And Val(m_EEP02) - 1 >= Val("" & RsTemp.Fields(0))) Then
            strSql = "update empelectronprocess set eep13='Y' where eep01 ='" & m_EEP01 & "' and eep02=" & Val(RsTemp.Fields(0))
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
      End If
      '2016/6/7 END
      
      PUB_DelFtpFile2 m_EEP01, " and eef02=" & m_EEP02, "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
      'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
      strSql = "delete from empelectronfile where eef01 ='" & m_EEP01 & "' and eef02=" & m_EEP02
      Pub_SaveLog strUserNum, "刪除順序(" & m_EEP02 & ")全部歷程附件", CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)), m_EEP01
      cnnConnection.Execute strSql
      
      cnnConnection.CommitTrans
      Screen.MousePointer = vbDefault
      
      Call QueryData
   End If
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 刪除失敗！" & vbCrLf & Err.Description
End Sub

'結束
'Modify By Sindy 2025/8/5 Private Sub ==> Public Sub
Public Sub cmdExit_Click()
   Unload Me
End Sub

'Add By Sindy 2024/1/2
Public Sub SetParent_IR(ByRef fm As Form)
   Set m_PrevForm_IR = fm
End Sub

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Add By Sindy 2025/10/21 上傳
Private Sub CmdF21_Click(Index As Integer)
Dim sFile
Dim strSavePath As String, bolSaveOK As Boolean
   
On Error GoTo ErrHnd
   
'   strCompName = PUB_FCPCaseNo2FileName(PField(1), PField(2), PField(3), PField(4)) '經理：不用加符號
'   strRepName = PUB_CaseNo2FileName(PField(1), PField(2), PField(3), PField(4))     '經理：5碼也可以，上傳後自動換6碼
   strSavePath = SetLabel25Folder(0, True, False)
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.*"
      .Filter = "All Files *.*|(*.*)"
      '預設上一次的路徑
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
            '多選
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
            For ii = 1 To UBound(sFile)
               If UpFCPF21File(.InitDir & "\" & CStr(sFile(ii)), strSavePath) = False Then
                  Exit Sub
               Else
                  bolSaveOK = True
               End If
            Next ii
            
         Else '單選
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            If UpFCPF21File(.FileName, strSavePath) = False Then
               Exit Sub
            Else
               bolSaveOK = True
            End If
         End If
      End If
   End With
   If bolSaveOK = True Then
      'Modify By Sindy 經理覺得不要直接開啟資料夾,彈訊息較好;人員有需要查看再自己點link
      'ShellExecute hLocalFile, "explore", strSavePath, vbNullString, vbNullString, 1
      MsgBox "已上傳完成！", vbInformation
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
'   '檢查表單是否已開啟，若是，則關閉
'   For Each nFrm In Forms
'      If StrComp(nFrm.Name, "frm090905", vbTextCompare) = 0 Then
'         If frm090905.txtData(0) <> PField(1) And frm090905.txtData(1) <> PField(2) _
'            And frm090905.txtData(2) <> PField(3) And frm090905.txtData(3) <> PField(4) Then
'            Unload frm090905
'         End If
'      End If
'   Next
'   frm090905.txtData(0) = PField(1)
'   frm090905.txtData(1) = PField(2)
'   frm090905.txtData(2) = PField(3)
'   frm090905.txtData(3) = PField(4)
''   frm090905.m_strCP10Nm = Trim(lblCP10.Caption) '案件性質名稱
''   Call frm090905.cmdFind_Click
'   frm090905.Show
'   frm090905.ZOrder
'   Call Label25_Click(0)
End Sub
Private Function UpFCPF21File(stFilePathNm As String, stSavePath As String) As Boolean
Dim strFile As String
Dim fs, f
   
   UpFCPF21File = False
   
   '路徑排除&
   strExc(1) = Mid(stFilePathNm, 1, InStrRev(stFilePathNm, "\") - 1)
   If InStr(strExc(1), "#") > 0 Or InStr(strExc(1), "&") > 0 Then
      MsgBox strExc(1) & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於路徑！", vbExclamation
      Exit Function
   End If
   strFile = Mid(stFilePathNm, InStrRev(stFilePathNm, "\") + 1) '檔名
   If InStr(strFile, "#") > 0 Or InStr(strFile, "&") > 0 Then
      MsgBox strFile & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名！", vbExclamation
      Exit Function
   End If
         
   If PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, Left(CboEEP04, 2), cp(10), , 0, _
      IIf(m_FlowUserNum = Trim(Left(m_SPMan, 6)) Or bolFCPFlow = True Or bolFCTFlow = True, False, True), _
      , , , , IIf(InStr(strFile, "申請書") > 0, True, False)) = False Then
      Exit Function
   End If
   
'   '檢查檔名規則
'   If Mid(UCase(strFile), 1, Len(strCompName)) <> UCase(strCompName) And _
'      Mid(UCase(strFile), 1, Len(strRepName)) <> UCase(strRepName) Then
'      MsgBox "檔案命名不符規定，字首必須為" & strCompName
'      Exit Function
'   End If
'
'   If ChkFCPF21File(stFilePathNm, strExc(2)) = True Then
      '檢查檔案是否正在使用中
      If PUB_ChkFileOpening(stFilePathNm) = True Then
         MsgBox stFilePathNm & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
         Exit Function
      End If
         
      Set fs = CreateObject("Scripting.FileSystemObject")
      Set f = fs.GetFile(stFilePathNm)
      '檔案大小為 0 KB 有誤
      If f.Size = 0 Then
         ShowMsg strFile & MsgText(9221)
         Exit Function
      End If
      strExc(2) = stSavePath & strFile
      If Dir(strExc(2)) <> "" Then
         If MsgBox("檔案已存在，要覆蓋嗎？" & vbCrLf & strExc(2), vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            FileCopy stFilePathNm, strExc(2) '上傳
         End If
      Else
         FileCopy stFilePathNm, strExc(2) '上傳
      End If
      If Dir(strExc(2)) <> "" Then
         UpFCPF21File = True
      End If
'   End If
End Function
'Private Function ChkFCPF21File(stFilePathNm As String, ByRef strReName As String) As Boolean
'Dim strFile As String
'
'   ChkFCPF21File = False
'   strFile = Mid(stFilePathNm, InStrRev(stFilePathNm, "\") + 1) '檔名
'   If ChkAttFileName(stFilePathNm, PField(1), PField(2), False) = True Then
'      '檢查是否需要更名
'      If InStr(strFile, "申請書") = 0 Then
'         strReName = PUB_GetSimpleName(strFile) '去掉非英數字的檔名
'      Else
'         strReName = strFile
'      End If
'      ChkFCPF21File = True
'   Else
'      MsgBox "下列檔案副檔名不符規則，請參考 ***下方說明*** !" & vbCrLf & _
'      stFilePathNm & vbCrLf & vbCrLf & "【副檔名說明】" & vbCrLf & _
'         "外文提申本(*.ORI.PDF、*.FIX_*.PDF)、電子送件專用檔(*.ZIP)" & vbCrLf & _
'         "例如：FCP058901.Fix.ori.pdf" & vbCrLf & vbCrLf & _
'         "中說替換本(*.FIX.DOC、*.COR.DOC或DOCX檔、或TXT檔)" & vbCrLf & _
'         "、中說修正本(*.FIX_U.DOC、*.COR_U.DOC或DOCX檔、或TXT檔)" & vbCrLf & _
'         "、中說圖檔(*.FIG.PDF)" & vbCrLf & _
'         "設計說明書(*.DES.DOC/DOCX檔、或TXT檔)" & vbCrLf & _
'         "例如：FCP058901.fix.doc、FCP058901.fix_u.doc" & vbCrLf & vbCrLf & _
'         "最終版中說或圖式PDF (.FIX_U、.FIX.、.FIX.SEQ. 或 .FIG.PDF)" & vbCrLf & _
'         "例如：FCP053555.fix_u.doc、FCP053555.Fig.pdf" & vbCrLf & vbCrLf & _
'         "外文本Word檔(*.ORI.DOC或DOCX檔、或TXT檔)" & vbCrLf & _
'         "例如：FCP058901.fix.ori.doc、FCP058901.fix.ori.pdf" & vbCrLf, vbInformation
'   End If
'End Function
Private Sub CmdF21_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   strExc(10) = GetFCPPathVal(Label25(0).Caption, PField(1), PField(2), Trim(lblCP10.Caption))
   CmdF21(Index).ToolTipText = "可上傳檔案到網芳：" & strExc(10)
End Sub

'Add By Sindy 2018/8/31
Private Sub cmdMail_Click()
   '查詢寄件備份且可轉寄
   PUB_ShowMailForm m_EEP01, "", lblCP10, , , , , True, , , True, m_EEP02, , , , Me
End Sub

'Add By Sindy 2018/9/26 同申請人案件,可操作多案
Private Sub cmdManyCase_Click()
   'If cmdManyCase.Visible = False Then m_RetrunRecv = m_EEP01 '回傳總收文號
   If CboEEP04.Tag <> Left(CboEEP04.Text, 2) Then '判斷是否有改變歷程狀態,是否預設值
      'm_RetrunRecv = m_EEP01 '回傳總收文號
      cmdManyCase.Tag = ""
      txtLpNote.Tag = ""
   End If
   If m_RetrunRecv = "" Then m_RetrunRecv = m_EEP01 '回傳總收文號 Add By Sindy 2018/10/30
   CboEEP04.Tag = Left(CboEEP04.Text, 2) '記錄目前操作的歷程狀態
   frm090202_2_1.m_EEP01 = m_EEP01
   If Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
      If UCase(m_PrevForm.Name) = UCase("frm090202_3") Then
         If m_PrevForm.SSTab1.Tab = 0 Then '未會稿
            frm090202_2_1.m_bolCustSend = False
         Else '已會稿
            frm090202_2_1.m_bolCustSend = True
         End If
      End If
      'Modify By Sindy 2023/9/5 m_ManyAppl 改為 IIf(bolTMFlow = True And cp(10) = "501", m_ManyApplCP56, m_ManyAppl)
      'Modify By Sindy 2023/10/6 Trim(Left(m_SPMan, 6)) => IIf(m_SPMan = "", cp(13), Trim(Left(m_SPMan, 6))) ex:T-225494,T-225495
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
      If frm090202_2_1.QueryData(Left(CboEEP04.Text, 2), _
         IIf((bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True) And cp(10) = "501", m_ManyApplCP56, m_ManyAppl), _
         m_FlowUserNum, IIf(m_SPMan = "", cp(13), Trim(Left(m_SPMan, 6))), Trim(Left(m_EPMan, 6)), , , , , , cp(1)) = True Then
         txtLpNote.Tag = "" 'Add By Sindy 2020/9/28
         '顯示同申請人案件鈕
         cmdManyCase.Visible = True
         cmdManyCase.Enabled = True
         frm090202_2_1.Show vbModal
      Else
         'ShowNoData
         cmdManyCase.Visible = False
         cmdManyCase.Enabled = False
         Unload frm090202_2_1
      End If
   ElseIf Left(CboEEP04.Text, 2) = EMP_會完 Then
      'Modify By Sindy 2023/9/5 m_ManyAppl 改為 IIf(bolTMFlow = True And cp(10) = "501", m_ManyApplCP56, m_ManyAppl)
      'Modify By Sindy 2023/10/6 Trim(Left(m_SPMan, 6)) => IIf(m_SPMan = "", cp(13), Trim(Left(m_SPMan, 6))) ex:T-225494,T-225495
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
      If frm090202_2_1.QueryData(Left(CboEEP04.Text, 2), _
         IIf((bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True) And cp(10) = "501", m_ManyApplCP56, m_ManyAppl), _
         m_FlowUserNum, IIf(m_SPMan = "", cp(13), Trim(Left(m_SPMan, 6))), Trim(Left(m_EPMan, 6)), , cp(10), , , , cp(1)) = True Then
         txtLpNote.Tag = "" 'Add By Sindy 2020/9/28
         '顯示同申請人案件鈕
         cmdManyCase.Visible = True
         cmdManyCase.Enabled = True
         frm090202_2_1.Show vbModal
      Else
         'ShowNoData
         cmdManyCase.Visible = False
         cmdManyCase.Enabled = False
         Unload frm090202_2_1
      End If
   'Add By Sindy 2020/9/16
   Else
      'Modify By Sindy 2023/9/5 m_ManyAppl 改為 IIf(bolTMFlow = True And cp(10) = "501", m_ManyApplCP56, m_ManyAppl)
      'Modify By Sindy 2023/10/6 Trim(Left(m_SPMan, 6)) => IIf(m_SPMan = "", cp(13), Trim(Left(m_SPMan, 6))) ex:T-225494,T-225495
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
      If frm090202_2_1.QueryData(Left(CboEEP04.Text, 2), _
         IIf((bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True) And cp(10) = "501", m_ManyApplCP56, m_ManyAppl), _
         m_FlowUserNum, IIf(m_SPMan = "", cp(13), Trim(Left(m_SPMan, 6))), Trim(Left(m_EPMan, 6)), cp(5), cp(10), cp(44), m_strLastEEP04, m_Country, cp(1)) = True Then
         txtLpNote.Tag = "多案單筆歷程"
         '顯示同申請人案件鈕
         cmdManyCase.Visible = True
         cmdManyCase.Enabled = True
         frm090202_2_1.Show vbModal
      Else
         'ShowNoData
         cmdManyCase.Visible = False
         cmdManyCase.Enabled = False
         Unload frm090202_2_1
      End If
   End If
   '2020/9/16 END
   
   Call SetTxtLpNote(False) 'Add By Sindy 2020/9/29
End Sub
Private Sub cmdManyCase_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmdManyCase.ToolTipText = Me.m_RetrunRecv
End Sub
Private Sub ManyCaseMoveFile()
   If cmdManyCase.Visible = False Or cmdManyCase.Enabled = False Then Exit Sub
Dim rsA As New ADODB.Recordset
   
   '有回傳總收文號
   If m_RetrunRecv <> "" Then
      lstAtt(0).Clear
'      strExc(0) = "select eef01,max(eef02) eef02 From empelectronfile" & _
'                  " where (eef01,eef02) in(" & _
'                  "select eep01,max(eep02) From empelectronprocess" & _
'                  " where eep01 in('" & Replace(m_RetrunRecv, ",", "','") & "')" & _
'                  " and eep04 in('" & EMP_客戶會稿 & "','" & EMP_送會 & "')" & _
'                  " group by eep01" & _
'                  ")" & _
'                  " group by eef01 order by eef01 asc,eef02 asc"
      'Add By Sindy 2020/9/17
      If m_strLastEEP04 <> "" Then
         strExc(0) = "select eef01,max(eef02) eef02,cp01,cp02,cp03,cp04" & _
                     " From empelectronfile,empelectronprocess,caseprogress" & _
                     " where eep01 in('" & Replace(m_RetrunRecv, ",", "','") & "')" & _
                     " and eep04 in('" & m_strLastEEP04 & "')" & _
                     " and eep01=eef01 and eep02=eef02" & _
                     " and eep01=cp09(+)" & _
                     " group by cp01,cp02,cp03,cp04,eef01 order by cp01 desc,cp02 desc,cp03 desc,cp04 desc"
      Else
      '2020/9/17 END
         strExc(0) = "select eef01,max(eef02) eef02,cp01,cp02,cp03,cp04" & _
                     " From empelectronfile,empelectronprocess,caseprogress" & _
                     " where eep01 in('" & Replace(m_RetrunRecv, ",", "','") & "')" & _
                     " and eep04 in('" & EMP_客戶會稿 & "','" & EMP_送會 & "')" & _
                     " and eep01=eef01 and eep02=eef02" & _
                     " and eep01=cp09(+)" & _
                     " group by cp01,cp02,cp03,cp04,eef01 order by cp01 desc,cp02 desc,cp03 desc,cp04 desc"
      End If
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         rsA.MoveFirst
         Do While Not rsA.EOF
            Call DownloadAttFile_copy(rsA.Fields("eef01"), rsA.Fields("eef02"), _
               IIf(lblCaseNo <> rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04"), rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04"), ""))
            rsA.MoveNext
         Loop
      End If
      rsA.Close
   End If
   
   Set rsA = Nothing
End Sub
Private Function ManyCaseSaveData(strMainEEP02 As String, strUpdDate As String, strUpdTime As String)
Dim rsA As New ADODB.Recordset
Dim intMaxEEP02 As Integer
Dim strUpdEEP09 As String
Dim strEEP11 As String '系統備註
Dim strEEP05 As String
Dim strUpdEEP15 As String 'Add By Sindy 2020/10/19
   
   ManyCaseSaveData = False
   '有回傳總收文號
   If m_RetrunRecvCnt > 0 Then
      strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp10 From caseprogress" & _
                  " where cp09 in('" & Replace(m_RetrunRecv, ",", "','") & "')" & _
                  " order by 1,2,3,4 asc"
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         rsA.MoveFirst
         Do While Not rsA.EOF
            strUpdEEP15 = ""
            '檢查是否為該文號的附件檔; 此函數是為了先處理非畫面上的案件
            If rsA.Fields("cp09") <> m_EEP01 Then '*****
               'Add By Sindy 2020/10/19 檢查此文號是否為”多案單筆歷程”,若為是,要一併更新其他案號資料
               'and eep04='" & m_strLastEEP04 & "' => CNULL(EMP_送會)
               strSql = "select * From empelectronprocess" & _
                        " where eep01='" & rsA.Fields("cp09") & "' and eep04=" & CNULL(EMP_送會) & _
                        " and eep05='" & m_FlowUserNum & "' and eep09='Y' and eep15 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If InStr("" & RsTemp.Fields("eep11"), "多案單筆歷程") > 0 Then
                     strUpdEEP15 = RsTemp.Fields("eep15")
                  End If
               End If
               '2020/10/19 END
               
               If Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
                  strEEP05 = Trim(Left(CboEEP05.Text, 6))
                  '更新EP37.客戶會稿日
                  strSql = "update engineerprogress set" & _
                                 " EP37=" & DBDATE(strUpdDate) & _
                           " where ep02='" & rsA.Fields("cp09") & "' and (EP37 is null or EP37=0)"
                  cnnConnection.Execute strSql
                  
                  'Add By Sindy 2020/10/19
                  If strUpdEEP15 <> "" Then '多案單筆歷程
                     strSql = "update engineerprogress set" & _
                                    " EP37=" & DBDATE(strUpdDate) & _
                              " where ep02 in('" & Replace(strUpdEEP15, ",", "','") & "') and (EP37 is null or EP37=0)"
                     cnnConnection.Execute strSql
                  End If
                  '2020/10/19 END
                  
               ElseIf Left(CboEEP04.Text, 2) = EMP_會完 Then
                  strEEP05 = GetCP14(rsA.Fields("cp09"))
                  strSql = "update engineerprogress set" & _
                                 " EP38=" & DBDATE(strUpdDate) & _
                           " where ep02='" & rsA.Fields("cp09") & "'"
                  cnnConnection.Execute strSql
                  
                  'Add By Sindy 2020/10/19
                  If strUpdEEP15 <> "" Then '多案單筆歷程
                     strSql = "update engineerprogress set" & _
                                    " EP38=" & DBDATE(strUpdDate) & _
                              " where ep02 in('" & Replace(strUpdEEP15, ",", "','") & "')"
                     cnnConnection.Execute strSql
                  End If
                  '2020/10/19 END
                  
                  'Add By Sindy 2018/7/17 商標處會完時,更新會稿完成日
                  'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
                  If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
                     'Modify By Sindy 2018/9/20 + 是否通知客戶
                     strSql = "update engineerprogress set" & _
                                    " EP08=" & DBDATE(strUpdDate) & _
                                    IIf(ChkEP11.Visible = True, IIf(ChkEP11.Value = 1, ",EP11='N'", ",EP11='Y'"), "") & _
                              " where ep02='" & rsA.Fields("cp09") & "'"
                     Pub_SeekTbLog strSql 'Add By Sindy 2021/6/28
                     cnnConnection.Execute strSql
                     
                     'Add By Sindy 2020/10/19
                     If strUpdEEP15 <> "" Then '多案單筆歷程
                        strSql = "update engineerprogress set" & _
                                       " EP08=" & DBDATE(strUpdDate) & _
                                       IIf(ChkEP11.Visible = True, IIf(ChkEP11.Value = 1, ",EP11='N'", ",EP11='Y'"), "") & _
                                 " where ep02 in('" & Replace(strUpdEEP15, ",", "','") & "')"
                        Pub_SeekTbLog strSql 'Add By Sindy 2021/6/28
                        cnnConnection.Execute strSql
                     End If
                     '2020/10/19 END
                  End If
                  
                  '更新上一筆流程的待回覆＝null
                  strSql = "update empelectronprocess set" & _
                           " EEP09=null" & _
                           " where eep01='" & rsA.Fields("cp09") & "'" & _
                           " and EEP04=" & CNULL(EMP_送會) & _
                           " and EEP09='Y'"
                  cnnConnection.Execute strSql
               End If
               
               '******************************
               '      承辦電子簽核流程檔
               '******************************
               '取得最大序號
               intMaxEEP02 = 0
               strSql = "select eep02 From empelectronprocess where eep01='" & rsA.Fields("cp09") & "' order by eep02 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  RsTemp.MoveFirst
                  If RsTemp.RecordCount > 0 Then
                     intMaxEEP02 = RsTemp.Fields(0)
                  End If
               End If
               intMaxEEP02 = intMaxEEP02 + 1
               '是否需等待回覆
               If InStr(EMP_需等待回覆的狀態, Left(CboEEP04.Text, 2)) > 0 Then
                  strUpdEEP09 = "Y" '必須等待回覆
               Else
                  strUpdEEP09 = ""
               End If
               '記錄在那個總收文號一同產生的歷程
               strEEP11 = "一同產生歷程:" & m_EEP01 & "(" & strMainEEP02 & ")"
               'Modify By Sindy 2023/12/18 +,eep16
               strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep09,eep10,eep11,eep12,eep14,eep15,eep16) values(" & _
                        CNULL(rsA.Fields("cp09")) & "," & intMaxEEP02 & "," & CNULL(Trim(txtEEP03)) & "," & _
                        CNULL(Left(CboEEP04.Text, 2)) & "," & _
                        CNULL(strEEP05) & "," & _
                        strSrvDate(1) & "," & strUpdTime & "," & CNULL(ChgSQL(txtEEP08)) & "," & CNULL(strUpdEEP09) & "," & _
                        CNULL(txtEEP10) & ",'" & strEEP11 & "'," & CNULL(m_EEP12) & "," & CNULL(m_EEP14) & ",'" & m_RetrunRecv & "'," & CNULL(m_EEP16) & ")"
               cnnConnection.Execute strSql
               '******************************
               '      承辦電子簽核附件檔
               '******************************
               If ManyCaseSaveAttFile(rsA.Fields("cp09"), CInt(intMaxEEP02), 0, _
                  rsA.Fields("cp01"), rsA.Fields("cp02"), rsA.Fields("cp03"), rsA.Fields("cp04")) = False Then
                  Exit Function
               End If
               
               'If Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
               If UCase(cmdSend.Caption) = UCase("E-Mail") Then
                  '******************************
                  '      寄件備份
                  '******************************
                  strSql = "insert into smailbackup(smb01,smb02,smb03,smb04,smb05,smb06,smb07,smb08,smb09,smb10,smb11)" & _
                           " select " & CNULL(rsA.Fields("cp09")) & ",smb02,smb03,smb04,smb05,smb06,smb07,smb08,smb09,smb10," & intMaxEEP02 & _
                           " from smailbackup where smb01='" & m_EEP01 & "' and smb11=" & strMainEEP02
                  cnnConnection.Execute strSql
                  'Modify By Sindy 2020/2/19 電子檔名,本所案號使用函數 PUB_CaseNo2FileName
                  strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,cpp08,cpp09,cpp10)" & _
                           " values(" & CNULL(rsA.Fields("cp09")) & ",'" & PUB_CaseNo2FileName(rsA.Fields("cp01"), rsA.Fields("cp02"), rsA.Fields("cp03"), rsA.Fields("cp04")) & _
                                   "." & rsA.Fields("cp10") & "." & strUpdDate & strUpdTime & "." & EMP_Email & ".menu',0," & _
                                   strUpdDate & "," & strUpdTime & ",'Y')"
                  cnnConnection.Execute strSql
               End If
               
               If Left(CboEEP04.Text, 2) = EMP_會完 Then
                  '發通知信
                  If FlowSendMail(True, CStr(intMaxEEP02), IIf(strUpdEEP15 <> "", True, False), rsA.Fields("cp09")) = False Then
                     Exit Function
                  End If
               End If
               
            End If
            rsA.MoveNext
         Loop
      End If
      rsA.Close
   End If
   
   ManyCaseSaveData = True
   Set rsA = Nothing
End Function

Private Function GetCP14(strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCP14 = ""
StrSQLa = "Select CP14 From Caseprogress Where CP09='" & strCP09 & "' And CP14 Is Not Null "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCP14 = rsA.Fields("CP14").Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

Private Function ManyCaseSaveAttFile(strEEF01 As String, intEEF02 As Integer, Index As Integer, _
   strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As Boolean
Dim stFilePath As String
Dim UpdModifyDate As Double, UpdModifyTime As Double
Dim iFileNo As Integer
Dim lngSize As Long '檔案大小
Dim adoRst As New ADODB.Recordset
Dim strFile As String ', stReName As String, strTemp As String
'Dim bolGetFileName As Boolean
'Dim intRow As Integer '檔案數量
Dim stFtpPath As String
Dim bolDelItem As Boolean
   
On Error GoTo ErrHand
   
   ManyCaseSaveAttFile = True
   
   '從最後一筆讀取到第一筆
   For ii = lstAtt(Index).ListCount - 1 To 0 Step -1
      If lstAtt(Index).ITEMDATA(ii) = 0 Then
         stFilePath = lstAtt(Index).List(ii)
         If InStrRev(stFilePath, " (") > 0 Then
            'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
            If UCase(Mid(stFilePath, InStrRev(stFilePath, " (") + 1, Len("(X86)"))) <> "(X86)" Then
            '2021/8/6 END
               stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
            End If
         End If
         strFile = GetFileName(stFilePath)
         '檢查是否為該文號的附件檔; 此函數是為了先處理非畫面上的案件
         If InStr(UCase(strFile), UCase(strCP01)) > 0 And InStr(strFile, Val(strCP02)) > 0 Then
            UpdModifyDate = Mid(lstAtt(Index).List(ii), InStr(lstAtt(Index).List(ii), "#") + 1, 8)
            UpdModifyTime = Mid(lstAtt(Index).List(ii), InStr(lstAtt(Index).List(ii), "#") + 9, 6)
            If iFileNo > 0 Then Close #iFileNo
            iFileNo = FreeFile
            Open stFilePath For Binary Access Read As #iFileNo
            lngSize = LOF(iFileNo)
            
            If lngSize = 0 Then
               Close #iFileNo
               ManyCaseSaveAttFile = False
               ShowMsg stFilePath & MsgText(9221)
               Exit Function
            End If
            
            With adoRst
               If adoRst.State = adStateClosed Then
                  strExc(0) = "select * from EmpElectronFile where rownum<1"
                  .CursorLocation = adUseClient
                  .Open strExc(0), cnnConnection, adOpenStatic, adLockOptimistic
               End If
               
   '            If Index = 1 Then '存卷資料時,檢查是否需要ReName
   '               If InStr(UCase(strFile), UCase("." & EMP_存卷資料)) = 0 And _
   '                  InStr(UCase(strFile), UCase("." & EMP_客戶資料)) = 0 Then
   '
   '                  '取得檔名
   '                  bolGetFileName = False: intRow = 0: stReName = ""
   '                  Do While bolGetFileName = False
   '                     stReName = Trim(PField(1)) & CStr(Val(PField(2))) & IIf(PField(3) <> "0" Or PField(4) <> "00", "-" & PField(3), "") & IIf(PField(4) <> "00", "-" & PField(4), "")
   '                     If InStr(UCase(strFile), UCase("." & cp(10))) = 0 Then
   '                        stReName = stReName & "." & EMP_客戶資料 & IIf(intRow > 0, intRow, "") & Mid(strFile, InStr(strFile, ".")) '截取本所案號後的檔名
   '                     Else
   '                        strTemp = Mid(strFile, InStr(strFile, ".") + 1)
   '                        stReName = stReName & "." & cp(10) & "." & EMP_客戶資料 & IIf(intRow > 0, intRow, "") & Mid(strTemp, InStr(strTemp, ".")) '截取第2個.後面的檔名
   '                     End If
   '                     '檢查檔案是否已存在
   '                     strSql = "select eef03" & _
   '                              " From EmpElectronFile" & _
   '                              " where eef01='" & strEEF01 & "' and eef02=" & intEEF02 & _
   '                              " and upper(eef03)='" & UCase(stReName) & "'"
   '                     intI = 1
   '                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '                     If intI = 0 Then
   '                        bolGetFileName = True
   '                        Call PUB_InsEfileCaption(CStr(PField(1)), EMP_客戶資料, intRow) '檢查是否有需要新增電子檔次要副檔名說明
   '                        Exit Do
   '                     End If
   '                     intRow = intRow + 1
   '                  Loop
   '                  If stReName <> "" Then
   '                     strFile = stReName
   '                  End If
   '               End If
   '            End If
               
               .AddNew
               .Fields("eef01").Value = strEEF01
               .Fields("eef02").Value = intEEF02
               .Fields("eef03").Value = strFile 'GetFileName(stFilePath)
               .Fields("eef04").Value = lngSize
               
               .Fields("eef09").Value = UpdModifyDate
               .Fields("eef10").Value = UpdModifyTime
               Close #iFileNo
               
               '檔案改放FTP
               PUB_PutFtpFile stFilePath, strEEF01, strFile, stFtpPath, "EMPELECTRONFILE", CStr(intEEF02)
               If stFtpPath <> "" Then
                  .Fields("eef11") = strSrvDate(1)
                  .Fields("eef12") = stFtpPath
               End If
               
               .UPDATE
            End With
            
            '移除 lstAtt 附件資料
            lstAtt(Index).RemoveItem ii
            bolDelItem = True
         End If
      End If
   Next ii
   
   If bolDelItem = True Then SetListScroll lstAtt(Index)
   Exit Function
   
ErrHand:
   Close #iFileNo
   ManyCaseSaveAttFile = False
   MsgBox Err.Description, vbCritical
End Function

'Add By Sindy 2018/10/29 多案時,檢查檔名是否有輸入符合的案號
'Modify by Sindy 2021/10/14 + Optional bolNotChkFileCaseNo As Boolean = False：不檢查檔名前頭為本所案號
Private Function ManyCaseChkFileName(strPathFile As String, Optional lstAttIdx As String = "", _
   Optional bolChkInfo As Boolean = True, Optional bolNotChkFileCaseNo As Boolean = False) As Boolean
   
Dim rsA As New ADODB.Recordset
Dim strSaveCaseNo1 As String, strSaveCaseNo2 As String, strSaveCaseNo3 As String, strSaveCaseNo4 As String
Dim strCaseNo As String
Dim stAttPathFile As String
   
   'Modify By Sindy 2023/11/14 T-246330母案判發會被擋住; 'Modify By Sindy 2023/11/9 逐筆檢查檔案時 + Or txtLpNote.Tag = "多案單筆歷程"
   'If cmdManyCase.Visible = False Or cmdManyCase.Enabled = False Then Exit Function
   
   'Add By Sindy 2023/6/21 多案歷程開放.CDATA.可以放多案歷程的其他案號
   If txtLpNote.Tag = "多案單筆歷程" Then
      'Modify By Sindy 2023/11/16 mark if => T-246300會完,多案歷程附件會檢查到不符 ex:T246300.T-246302.CDATA.pdf
      'Modify By Sindy 2024/3/12 多案歷程開放.CDATA.可以放多案歷程的其他案號
      If InStr(UCase(strPathFile), ".CDATA.") = 0 Then
      '2024/3/12 END
      '2023/11/16 END
         '檢查為操作此道歷程的案號
         If lstAttIdx <> "" Then
            'Modify By Sindy 2024/11/27 + 外商FC同外專不鎖中文,因附件區不進卷宗區: IIf(lstAttIdx = "0" And (bolFCPFlow = True Or bolFCTFlow = True), True, False)
            If PUB_ChkEmpFlowFNMRule(lblCaseNo, strPathFile, Left(CboEEP04, 2), cp(10), , lstAttIdx, bolChkInfo, , , , bolNotChkFileCaseNo, IIf(lstAttIdx = "0" And (bolFCPFlow = True Or bolFCTFlow = True), True, False)) = False Then
               ManyCaseChkFileName = False
            Else
               ManyCaseChkFileName = True
            End If
         Else
            If PUB_ChkEmpFlowFNMRule(lblCaseNo, GetFileName(strPathFile), Left(CboEEP04, 2), cp(10), , , , , , , bolNotChkFileCaseNo) = False Then
               ManyCaseChkFileName = False
            Else
               ManyCaseChkFileName = True
            End If
         End If
         Exit Function
      End If
   End If
   '2023/6/21 END
   
   If lstAttIdx <> "" Then
      stAttPathFile = strPathFile
   Else
      stAttPathFile = GetFileName(strPathFile)
   End If
   '有回傳總收文號
   If m_RetrunRecv <> "" Then
      strExc(0) = "select cp01,cp02,cp03,cp04" & _
                  " From caseprogress" & _
                  " where cp09 in('" & Replace(m_RetrunRecv, ",", "','") & "')"
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
            If InStr(UCase(stAttPathFile), strSaveCaseNo1) > 0 Or _
               InStr(UCase(stAttPathFile), strSaveCaseNo2) > 0 Or _
               InStr(UCase(stAttPathFile), strSaveCaseNo3) > 0 Or _
               InStr(UCase(stAttPathFile), strSaveCaseNo4) > 0 Then
               strCaseNo = rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04")
               Exit Do
            End If
            rsA.MoveNext
         Loop
      End If
      rsA.Close
   End If
   
   If strCaseNo = "" Then strCaseNo = lblCaseNo
   If lstAttIdx <> "" Then
      'Modify By Sindy 2024/11/27 + 外商FC同外專不鎖中文,因附件區不進卷宗區: IIf(lstAttIdx = "0" And (bolFCPFlow = True Or bolFCTFlow = True), True, False)
      If PUB_ChkEmpFlowFNMRule(strCaseNo, strPathFile, Left(CboEEP04, 2), cp(10), , lstAttIdx, bolChkInfo, , , , bolNotChkFileCaseNo, IIf(lstAttIdx = "0" And (bolFCPFlow = True Or bolFCTFlow = True), True, False)) = False Then
         ManyCaseChkFileName = False
      Else
         ManyCaseChkFileName = True
      End If
   Else
      If PUB_ChkEmpFlowFNMRule(strCaseNo, GetFileName(strPathFile), Left(CboEEP04, 2), cp(10), , , , , , , bolNotChkFileCaseNo) = False Then
         ManyCaseChkFileName = False
      Else
         ManyCaseChkFileName = True
      End If
   End If
   
   Set rsA = Nothing
End Function

'Add By Sindy 2017/8/28 多國案
Private Sub cmdCaseMap_Click()
   If frm090202_2_1.QueryData("0") = True Then
      frm090202_2_1.Show vbModal
   Else
      ShowNoData
      Unload frm090202_2_1
   End If
End Sub
Private Sub cmdCaseMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmdCaseMap.ToolTipText = Me.m_RetrunRecv
End Sub

'Add By Sindy 2013/9/5
Private Sub cmdok_Click(Index As Integer)
cmdState = Index
bolQuery = True
PubShowNextData
Exit Sub
End Sub

Public Sub PubShowNextData()
Dim rsA As New ADODB.Recordset
Dim stFileName As String

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
   If cp(140) <> "" Then
      '查詢接洽記錄單
      'Modify By Sindy 2022/12/23 改用共用函數
      Call PUB_Queryfrm090801(cp(140), cp(5), Me)
'      'Modify By Sindy 2022/9/5
'      If DBDATE(cp(5)) >= 接洽單電子收文啟用日 Then
'         frm090801_Q.SetParent Me
'         frm090801_Q.m_blnCallPrint = True
'         frm090801_Q.Text5 = cp(140)
'         Call frm090801_Q.cmdOK_Click(4)
'         'frm090801_Q.ZOrder
'         frm090801_Q.Show vbModal
'      Else
'      '2022/9/5 END
'         frm090801.SetParent Me
'         frm090801.m_blnCallPrint = True 'Add By Sindy 2022/10/19
'         frm090801.Text5 = cp(140)
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
Case 3 '承辦進度
   If bolQuery = True Then
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
      If bolTMFlow = True Or bolOtherFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
         Me.Enabled = False
         '排除母層是共同查詢
         If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
            fnCloseAllFrm100
         End If
         If fnSaveParentForm(Me) = False Then
             Me.Enabled = True
             Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         bolQuery = False
         frm100101_K.Show
         frm100101_K.Process m_EEP01
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         Exit Sub
      Else
         Me.Enabled = False
         '排除母層是共同查詢
         If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
            fnCloseAllFrm100
         End If
         If fnSaveParentForm(Me) = False Then
             Me.Enabled = True
             Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         bolQuery = False
         frm100101_F.Show
         frm100101_F.Cmd(0).Visible = False
         frm100101_F.Cmd(2).Visible = False
         frm100101_F.cmd1.Visible = False
         frm100101_F.Process m_EEP01
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         Exit Sub
      End If
   End If
Case 4 '完整卷宗
   Screen.MousePointer = vbHourglass
   frm100101_L.m_strKey = lblCaseNo.Caption
   frm100101_L.SetParent Me
   If frm100101_L.QueryData = True Then
      frm100101_L.Show
      Me.Hide
   Else
      Unload frm100101_L
   End If
   Screen.MousePointer = vbDefault
'Add By Sindy 2023/11/16
Case 5 '原始檔區
   Screen.MousePointer = vbHourglass
   frm100101_M.m_strKey = lblCaseNo.Caption '總收文號
   frm100101_M.SetParent Me
   If frm100101_M.QueryData = True Then
      frm100101_M.Show
      Me.Hide
   Else
      Unload frm100101_M
   End If
   Screen.MousePointer = vbDefault
'2023/11/16 END
Case Else
End Select
End Sub

'儲存存卷資料
Private Function funSaveEEF02_0() As Boolean
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   '刪除附件
   For ii = 1 To UBound(m_FilesRemoved)
      PUB_DelFtpFile2 m_EEP01, " and eef02=0 and eef03='" & ChgSQL(m_FilesRemoved(ii)) & "'", "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
      'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
      strSql = "delete EmpElectronFile where eef01='" & m_EEP01 & "'" & _
                                       " and eef02=0 and eef03='" & ChgSQL(m_FilesRemoved(ii)) & "'"
      cnnConnection.Execute strSql
   Next
   If SaveAttFile(m_EEP01, 0, 1) = False Then
      GoTo ErrHand
   End If
   
   cnnConnection.CommitTrans
   
   cmdSave.Visible = False
   Erase m_FilesRemoved
   ReDim m_FilesRemoved(0) As String
   
   funSaveEEF02_0 = True
   
   Call ReadAttachFile_other(m_EEP01) 'Add By Sindy 2013/9/25 查詢存卷區
   Exit Function
   
ErrHand:
   funSaveEEF02_0 = False
   cnnConnection.RollbackTrans
   MsgBox " 存卷存檔失敗！" & vbCrLf & Err.Description
End Function

'Add By Sindy 2024/4/17 匯出Outlook
Private Sub cmdOutlook_Click()
   Call EmpFlowFCPOutlook(m_EEP01, PField(1), PField(2), PField(3), PField(4), txtCaseName(0), lblCP10)
End Sub

'Add By Sindy 2013/10/24
'刪除歷程附件
Private Sub cmdRemAttDB_Click(Index As Integer)
   Call RemoveList_DB(lstAtt(Index), Index)
End Sub
'Add By Sindy 2013/10/24
Private Function RemoveList_DB(oList As ListBox, Index As Integer) As Boolean
Dim ii As Integer
Dim bolDel As Boolean
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            
            If MsgBox("確定要刪除" & GetFileName(oList.List(ii)) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Function
            
            If oList.ITEMDATA(ii) > 0 Then
               intI = UBound(m_FilesRemoved) + 1
               ReDim Preserve m_FilesRemoved(intI) As String
               m_FilesRemoved(intI) = GetFileName(oList.List(ii))
            End If
            
            '直接從資料庫刪除檔案
            If Index = 1 Then '存卷資料
               bolDel = DeleteFile(GetFileName(oList.List(ii)), 0)
            Else
               bolDel = DeleteFile(GetFileName(oList.List(ii)), CInt(m_EEP02))
            End If
            If bolDel = True Then
               oList.RemoveItem ii
               SetListScroll oList
               RemoveList_DB = True
               ii = ii - 1
            End If
         End If
         ii = ii + 1
      Loop
   End If
End Function
'Add By Sindy 2013/10/24
Private Function DeleteFile(strFileName As String, intEEP02 As Integer) As Boolean
Dim stReName As String
   
On Error GoTo ErrHand
   
   DeleteFile = True
   Screen.MousePointer = vbHourglass
   
   PUB_DelFtpFile2 m_EEP01, " and eef02=" & intEEP02 & " and upper(eef03)='" & UCase(strFileName) & "'", "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
   'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
   strSql = "delete from EmpElectronFile where eef01='" & m_EEP01 & "' and eef02=" & intEEP02 & " and upper(eef03)='" & UCase(strFileName) & "'"
   cnnConnection.Execute strSql
   Pub_SaveLog strUserNum, "刪除歷程附件：順序(" & intEEP02 & ")" & strFileName, CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)), m_EEP01
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
'Add By Sindy 2013/10/24
'新增歷程附件
Private Sub cmdAddAttDB_Click(Index As Integer)
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim UpdModifyDate As Double, UpdModifyTime As Double
   Dim stFiName As String, stReName As String
   Dim strFilePath As String 'Add By Sindy 2018/9/26
   Dim bolNotChkFileCaseNo As Boolean
   
On Error GoTo ErrHnd
   
   'Add By Sindy 2018/9/26 取得開啟檔案的路徑
   If lstAtt(Index).ListCount > 0 Then
      ii = 0
      Do While ii < lstAtt(Index).ListCount
         If lstAtt(Index).Selected(ii) = True Then
            If InStr(lstAtt(Index).List(ii), "\") > 0 Then
               strFilePath = Mid(lstAtt(Index).List(ii), 1, InStrRev(lstAtt(Index).List(ii), "\") - 1)
               Exit Do
            End If
         End If
         ii = ii + 1
      Loop
   End If
   If strFilePath = "" Then
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         strFilePath = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         strFilePath = PUB_Getdesktop
      End If
      'Add By Sindy 2022/5/4
      If Dir(strFilePath, vbDirectory) = "" Then
         strFilePath = PUB_Getdesktop
      End If
      '2022/5/4 END
   End If
   '2018/9/26 END
   
   stFileName = "*.*"
   With CommonDialog1
      .CancelError = True
      'Modify By Sindy 2024/12/12
      '.FileName = stFileName
      '.Filter = "All Files (*.*)|*.*"
      Call GetAddFileKind(CommonDialog1, Index)
      '2024/12/12 END
      'Modify By Sindy 2018/9/26
'      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
'         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
'      Else
'         .InitDir = PUB_Getdesktop
'      End If
      .InitDir = strFilePath
      '2018/9/26 END
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            'Modify By Sindy 2018/9/26 不要儲存系統暫存區路徑
            If InStr(UCase(Trim(sFile(0))), UCase(App.path)) = 0 Then
            '2018/9/26 END
               'Add By Sindy 2022/6/7 路徑字元有萬國碼?不要儲存路徑
               If InStr(strConV(strConV(sFile(0), vbFromUnicode), vbUnicode), "?") = 0 Then
               '2022/6/7 END
                  SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
               End If
            End If
            For ii = 1 To UBound(sFile)
               'Add By Sindy 2013/10/9
               If InStr(CStr(sFile(ii)), "#") > 0 Or InStr(CStr(sFile(ii)), "&") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名"
                  Exit Sub
               End If
               '2013/10/9 END
               
               '檢查檔名規則
               'Modify By Sindy 2018/9/25 商標處開放下列檔名可不加本所案號,存檔時系統自動補填
               'Add By Sindy 2021/10/14 ACS案件,不限制電子檔要輸入案號
               bolNotChkFileCaseNo = False
               If NotChkFileCaseNo(CStr(sFile(ii)), Index) = True Or _
                  PField(1) = "ACS" Then
                  bolNotChkFileCaseNo = True
               End If
               '2018/9/25 END
               If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), "Y", cp(10), , Index, , , , , bolNotChkFileCaseNo) = False Then
                  Exit Sub
               End If
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
               'Add By Sindy 2014/3/11
               ElseIf f.Size > 5242880 Then
                  If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                     Exit Sub
                  End If
               '2014/3/11 END
               End If
               '2013/9/6 END
               If AddListX(lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)) = True Then
                  '存檔
                  If Index = 1 Then '存卷資料
                     If SaveAttFile(m_EEP01, 0, Index) = False Then
                        Exit Sub
                     End If
                  Else
                     If SaveAttFile(m_EEP01, CInt(m_EEP02), Index) = False Then
                        Exit Sub
                     End If
                  End If
                  Pub_SaveLog strUserNum, "新增歷程附件：順序(" & m_EEP02 & ")" & sFile(ii), CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)), m_EEP01
                  '重新顯示附件區
                  If Index = 1 Then '存卷資料
                     Call ReadAttachFile_other(m_EEP01)
                  Else
                     Call ReadAttachFile(m_EEP01, CInt(m_EEP02))
                  End If
               End If
            Next
         Else
            'stFileName = GetFileName(.FileName)
            'Modify By Sindy 2013/10/9
            'stFiName = GetFileName(.FileName) '不含路徑的檔名
            stFiName = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(stFiName, "#") > 0 Or InStr(stFiName, "&") > 0 Then
               MsgBox stFiName & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名"
               Exit Sub
            End If
            '2013/10/9 END
            
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     'Modify By Sindy 2018/9/26 不要儲存系統暫存區路徑
                     If InStr(UCase(Trim(.FileName)), UCase(App.path)) = 0 Then
                     '2018/9/26 END
                        'Add By Sindy 2022/6/7 路徑字元有萬國碼?不要儲存路徑
                        If InStr(strConV(strConV(Mid(Trim(.FileName), 1, ii - 1), vbFromUnicode), vbUnicode), "?") = 0 Then
                        '2022/6/7 END
                           SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                        End If
                        Exit For
                     End If
                  End If
               Next ii
            End If
            
            '檢查檔名規則
            'Modify By Sindy 2018/9/25 商標處開放下列檔名可不加本所案號,存檔時系統自動補填
            'Add By Sindy 2021/10/14 ACS案件,不限制電子檔要輸入案號
            bolNotChkFileCaseNo = False
            If NotChkFileCaseNo(stFiName, Index) = True Or _
               PField(1) = "ACS" Then
               bolNotChkFileCaseNo = True
            End If
            '2018/9/25 END
            If PUB_ChkEmpFlowFNMRule(lblCaseNo, stFiName, "Y", cp(10), , Index, , , , , bolNotChkFileCaseNo) = False Then
               Exit Sub
            End If
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg stFiName & MsgText(9221)
               Exit Sub
            'Add By Sindy 2014/3/11
            ElseIf f.Size > 5242880 Then
               If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                  Exit Sub
               End If
            '2014/3/11 END
            End If
            '2013/9/6 END
            If AddListX(lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)) = True Then
               '存檔
               If Index = 1 Then '存卷資料
                  If SaveAttFile(m_EEP01, 0, Index) = False Then
                     Exit Sub
                  End If
               Else
                  If SaveAttFile(m_EEP01, CInt(m_EEP02), Index) = False Then
                     Exit Sub
                  End If
               End If
               Pub_SaveLog strUserNum, "新增歷程附件：順序(" & m_EEP02 & ")" & stFiName, CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)), m_EEP01
               '重新顯示附件區
               If Index = 1 Then '存卷資料
                  Call ReadAttachFile_other(m_EEP01)
               Else
                  Call ReadAttachFile(m_EEP01, CInt(m_EEP02))
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

Private Sub CmdSave_Click()
   If funSaveEEF02_0 = False Then Exit Sub
End Sub

Private Sub cmdSave2_Click()
   If funSaveEmpPaperData = True Then
      MsgBox " 存檔完畢！", vbInformation
   End If
End Sub

'Add By Sindy 2023/9/20
Private Sub cmdSave3_Click()
   If funSaveEmpPaperData_FCP(False) = True Then
      MsgBox " 存檔完畢！", vbInformation
   End If
End Sub

'Add By Sindy 2018/9/25
Private Sub cmdMod_Click()
   frm020102_04.SetData 0, PField(1), True
   frm020102_04.SetData 1, PField(2), False
   frm020102_04.SetData 2, PField(3), False
   frm020102_04.SetData 3, PField(4), False
   frm020102_04.SetData 4, m_EEP01, False
   ' 91.09.02 modify by louis (增加案件性質參數)
   frm020102_04.SetData 5, cp(10), False
   frm020102_04.SetParent Me
   Me.Hide
   frm020102_04.Show
   frm020102_04.QueryData
End Sub
Public Sub cmdMod_LostFocus()
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_EEP01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      cmdMod.BackColor = &H8080FF '紅色
   Else
      cmdMod.BackColor = &H8000000F
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub
'2018/9/25 END

'Add By Sindy 2018/8/14
Private Sub cmdCP118_Click()
Dim bolisOpen As Boolean
   
   'Modify By Sindy 2023/12/8 檢查指定送件日
   'Modify By Sindy 2024/1/23 改為共用函數
   If PUB_ChkCP141IsSend(m_EEP01, False, "送件") = False Then
      Exit Sub
   End If
   
   '檢查表單是否已開啟，若是，則關閉
   For Each nFrm In Forms
      If StrComp(nFrm.Name, "frm090202_2_2", vbTextCompare) = 0 Then
         bolisOpen = True
      End If
   Next
   If bolisOpen = True Then
      frm090202_2_2.Show
   Else
      If frm090202_2_2.QueryData = True Then
         frm090202_2_2.Show 'vbModal
      Else
         '此段應該不會發生
         ShowNoData
         Unload frm090202_2_2
      End If
   End If
End Sub
Public Sub cmdCP118_LostFocus()
   If cmdCP118.Tag = "Y" Then
      cmdCP118.BackColor = &H80FF&
   Else
      cmdCP118.BackColor = &H8000000F
   End If
End Sub
'2018/8/14 END

'Added by Lydia 2018/09/20 呼叫查名區
Private Sub cmdTMQ_Click()
   'Added by Lydia 2022/01/06
   If PUB_CheckFormExist("frm090128") Then
       MsgBox "請先關閉〔查覆明細畫面〕畫面！"
       Exit Sub
   End If
   If PUB_CheckFormExist("frm090127") Then
       MsgBox "請先關閉〔查名/查覆區〕畫面！"
       Exit Sub
   End If
   'end 2022/01/06
   
   Set frm090127.Tmpfrm090128 = frm090128
   frm090127.SetParent Me, PField(1) & PField(2) & PField(3) & PField(4), "3" '如果是申請階段傳入2, 核駁前先行通知傳入3(已發文的查名單)
   'Modified by Lydia 2021/01/06 開放所有人員查看
   'If frm090127.IsRolePlay("待查") = True Then
   If frm090127.IsRolePlay("查覆") = True Then
      frm090127.Show
   End If
End Sub

'Add By Sindy 2025/10/14 變大的附件區
Private Sub cmdOpen_Click()
   FrameFCPlink(1).Visible = FrameFCPlink(0).Visible
   If FrameFCPlink(1).Visible = True Then lstAtt(2).Height = lstAtt(2).Height + FrameFCPlink(1).Height
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
   cmdAddAtt(2).Enabled = cmdAddAtt(0).Enabled
   cmdRemAtt(2).Enabled = cmdRemAtt(0).Enabled
   CmdF21(1).Enabled = CmdF21(0).Enabled
   lstAtt(lstAttNew).Clear
   For ii = 0 To lstAtt(lstAttOld).ListCount - 1
      lstAtt(lstAttNew).AddItem lstAtt(lstAttOld).List(ii)
      lstAtt(lstAttNew).Selected(ii) = lstAtt(lstAttOld).Selected(ii)
   Next ii
   If lstAtt(lstAttNew).ListCount > 0 Then SetListScroll lstAtt(lstAttNew)
   Frame1Big.Visible = Not bolVal
End Sub
'2025/10/14 END

Private Sub Form_Load()
   Text1.Visible = False 'Add By Sindy 2014/10/1 備註:「聯絡」的附件，送件後一律刪除，欲留存者請置於存卷資料頁籤
   
   MoveFormToCenter Me
   lstAtt(2).Tag = lstAtt(2).Height 'Add By Sindy 2025/10/14
   
   Me.Tag = Me.Caption 'Add By Sindy 2023/12/6
   Me.txtEEP03.BackColor = &H8000000F
   Me.txtEEP03_2.BackColor = &H8000000F
   Me.lblFa.BackColor = &H8000000F
   ReDim m_FilesRemoved(0)
   'Add By Sindy 2021/5/31
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath")
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   '2021/5/31 END
   'Modify By Sindy 2021\5\19
   'm_AttachPath = App.Path & "\SeminarAttach"
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   '2021\5\19 END
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   SSTab1.Tab = 0
   If m_FlowUserNum = "" Then m_FlowUserNum = strUserNum 'Add By Sindy 2013/9/12 案件流程所屬人員
   Label6.Caption = "" '待回的最後流程
   
   'Add By Sindy 2013/10/3
   If (InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 And Pub_StrUserSt03 = "M51") Or Pub_StrUserSt03 = "M51" Then
      cmdDel.Visible = True
      cmdAddAttDB(0).Visible = True 'Add By Sindy 2013/10/24
      cmdRemAttDB(0).Visible = True 'Add By Sindy 2013/10/24
      cmdAddAttDB(1).Visible = True 'Add By Sindy 2013/10/24
      cmdRemAttDB(1).Visible = True 'Add By Sindy 2013/10/24
   Else
      cmdDel.Visible = False
      cmdAddAttDB(0).Visible = False 'Add By Sindy 2013/10/24
      cmdRemAttDB(0).Visible = False 'Add By Sindy 2013/10/24
      cmdAddAttDB(1).Visible = False 'Add By Sindy 2013/10/24
      cmdRemAttDB(1).Visible = False 'Add By Sindy 2013/10/24
   End If
   '2013/10/3 END
   ReDim pa(TF_PA) 'Add By Sindy 2014/6/20
   ReDim sp(tf_SP) 'Add By Sindy 2018/4/18
   ReDim tm(TF_TM) 'Add By Sindy 2018/4/18
   ReDim cp(TF_CP) 'Add By Sindy 2018/4/18
   'Add By Sindy 2021/9/3
   ReDim lC(TF_LC)
   ReDim hc(TF_HC)
   '2021/9/3 END
   
   'Add By Sindy 2018/10/12 設定在頂層
   Text2.ZOrder '存卷資料文字框
   cmdTMQ.ZOrder '查名區
   cmdMod.ZOrder '變更事項
   cmdCP118.ZOrder '電子送件
   cmdDel.ZOrder '電腦中心刪除歷程(&D)
   '2018/10/12 END
   
   Frame7.BorderStyle = 0 'Add By Sindy 2023/9/23
   lstAtt(0).Height = 1840 'Add By Sindy 2024/5/30
   
   g_LetterDebug = True 'Modify By Sindy 2025/7/9 要記錄Log
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   strSaveConfirm = "" 'Add By Sindy 2020/1/17
   
   'Add By Sindy 2013/9/26 儲存存卷資料
   If Me.cmdSave.Visible = True Then
      If MsgBox("存卷資料有異動，是否要儲存？", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption & " 重要訊息！") = vbYes Then
         Call CmdSave_Click
      End If
   End If
   
   'Add By Sindy 2020/1/17
   If Me.cmdSend.Enabled = True Then
      If MsgBox("新歷程編輯中，確定要放棄作業嗎？", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption & " 重要訊息！") = vbNo Then
         Cancel = True
         strSaveConfirm = "Y" 'Add By Sindy 2020/1/17
         Exit Sub
      End If
   End If
   '2020/1/17 END
   
   'Add By Sindy 2021/3/4
   '檢查表單是否開啟著，若是，則繼續
   For Each nFrm In Forms
      If StrComp(nFrm.Name, TypeName(m_PrevForm), vbTextCompare) = 0 Then
         'Modify By Sindy 2025/8/1
'         If Left(CboEEP04.Text, 2) <> EMP_收文分析 Then
         '2025/8/1 END
            m_PrevForm.Show
'         End If
         'Add By Sindy 2023/9/19
         If UCase(m_PrevForm.Name) = UCase("frm060107_1") Then
            m_PrevForm.cmdBack_Click
         '2023/9/19 END
         ElseIf UCase(m_PrevForm.Name) = UCase("frm090711") Then '繪圖人員工作進度
            If bolSave = True Then
               If m_PrevForm.intBackTab = 2 Then
                  If m_PrevForm.QueryData(True) = True Then
                     m_PrevForm.SSTab1.Tab = 2
                  End If
               Else
                  m_PrevForm.SSTab1.Tab = m_PrevForm.intBackTab
               End If
               m_PrevForm.RefreshOneRecord
               m_PrevForm.MouseClick IIf(Val("" & m_PrevForm.SWPRow) < 1, 1, m_PrevForm.SWPRow)
            End If
         ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_2") Then '專利承辦人工作進度
            If bolSave = True Then
               'Add By Sindy 2014/1/15
               If Left(CboEEP04.Text, 2) = EMP_附加流程 Then
                  Call m_PrevForm.Combo1_DropButtonClick 'Combo1_Click Modify By Sindy 2022/1/20 '重新讀取資料,因附加流程有新增CP
                  Call m_PrevForm.Combo1_DropButtonClick '第二次才會生效 Add By Sindy 2025/7/31
                  If Trim(Left(CboCP10.Text, 4)) = "936" Or _
                     Trim(Left(CboCP10.Text, 4)) = "957" Or _
                     Trim(Left(CboCP10.Text, 4)) = "958" Then
                     MsgBox "附加流程的（" & Trim(Mid(CboCP10.Text, 5, Len(CboCP10.Text))) & "）簽辦流程，系統不鎖定直接送判" & vbCrLf & vbCrLf & _
                            "故回到「工作進度資料維護」中，點選此筆新增案件，自行進行所需的簽辦流程!!"
                     m_PrevForm.SSTab1.Tab = 0
                  Else
                     If m_PrevForm.intBackTab = 2 Then
                        If m_PrevForm.QueryData(True) = True Then
                           m_PrevForm.SSTab1.Tab = 2
                        End If
                     Else
                        m_PrevForm.SSTab1.Tab = 0
                     End If
                  End If
               Else
               '2014/1/15 END
         '         Call m_PrevForm.cmdOK_Click(1)
                  If m_PrevForm.intBackTab = 2 Then
                     If m_PrevForm.QueryData(True) = True Then
                        m_PrevForm.SSTab1.Tab = 2
         '            Else
         '               m_PrevForm.SSTab1.Tab = 1 '無待辦歷程，回工作進度詳細資料
                     End If
                  Else
                     m_PrevForm.SSTab1.Tab = m_PrevForm.intBackTab
                  End If
               End If
            End If
         'Add By Sindy 2023/9/25
         ElseIf UCase(m_PrevForm.Name) = UCase("frm090909") Then '外專承辦人工作進度
            If bolSave = True Then
               'Add By Sindy 2025/4/10
               If Left(CboEEP04.Text, 2) = EMP_發文歸檔 And ChkEED08.Visible = True And ChkEED08.Value = 1 Then
                  Call m_PrevForm.QueryCombo1Data '重新讀取資料,因附加流程有新增CP
                  Call m_PrevForm.QueryCombo1Data '第二次才會生效 Add By Sindy 2025/7/31
                  m_PrevForm.SSTab1.Tab = 0
               Else
               '2025/4/10 END
                  'Add By Sindy 2025/7/31
                  If Left(CboEEP04.Text, 2) = EMP_附加流程 Then
                     Call m_PrevForm.QueryCombo1Data '重新讀取資料,因附加流程有新增CP
                     Call m_PrevForm.QueryCombo1Data '第二次才會生效
                     If Trim(Left(CboCP10.Text, 4)) = "936" Or _
                        Trim(Left(CboCP10.Text, 4)) = "957" Or _
                        Trim(Left(CboCP10.Text, 4)) = "958" Then
                        MsgBox "附加流程的（" & Trim(Mid(CboCP10.Text, 5, Len(CboCP10.Text))) & "）簽辦流程，系統不鎖定直接送判" & vbCrLf & vbCrLf & _
                               "故回到「工作進度資料維護」中，點選此筆新增案件，自行進行所需的簽辦流程!!"
                        m_PrevForm.SSTab1.Tab = 0
                     End If
                  Else
                  '2025/7/31 END
                     If m_PrevForm.intBackTab = 2 Then
                        If m_PrevForm.QueryData(True) = True Then
                           m_PrevForm.SSTab1.Tab = 2
                        End If
                     Else
                        m_PrevForm.SSTab1.Tab = m_PrevForm.intBackTab
                     End If
                  End If
               End If
            End If
            '2023/9/25 END
         ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_b") Then '商標承辦人工作進度
            If bolSave = True Then
               'Add By Sindy 2025/7/31
               If Left(CboEEP04.Text, 2) = EMP_附加流程 Then
                  Call m_PrevForm.Combo1_DropButtonClick '重新讀取資料,因附加流程有新增CP
                  Call m_PrevForm.Combo1_DropButtonClick '第二次才會生效
                  m_PrevForm.SSTab1.Tab = 0
               Else
               '2025/7/31 END
                  If m_PrevForm.intBackTab = 2 Then
                     If m_PrevForm.QueryData(True) = True Then
                        m_PrevForm.SSTab1.Tab = 2
                     End If
                  Else
                     m_PrevForm.SSTab1.Tab = m_PrevForm.intBackTab
                  End If
               End If
            End If
         'Add By Sindy 2021/9/24
         ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_d") Then '法務,顧問承辦人工作進度
            If bolSave = True Then
                  If m_PrevForm.intBackTab = 2 Then
                     If m_PrevForm.QueryData(True) = True Then
                        m_PrevForm.SSTab1.Tab = 2
                     End If
                  Else
                     m_PrevForm.SSTab1.Tab = m_PrevForm.intBackTab
                  End If
            End If
         ElseIf UCase(m_PrevForm.Name) = UCase("frm100101_2") Then '案件進度查詢
            m_PrevForm.PubShowNextData
            If Me.ShowNextData = True Then
               Cancel = 1 'Add By Sindy 2020/2/25 不離開Form (Unload 陳述式被程式碼呼叫。)
               Exit Sub
            End If
         Else 'If UCase(m_PrevForm.Name) = UCase("frm090202_1") Or _
                'UCase(m_PrevForm.Name) = UCase("frm090202_3") Then
            'Modify By Sindy 2018/8/29 待會稿區
            'Memo by Lydia 2019/07/03 更名為「專利／商標會稿」
            If UCase(m_PrevForm.Name) = UCase("frm090202_3") Then
               Call m_PrevForm.QueryData(m_PrevForm.SSTab1.Tab)
               '2018/8/29 END
            'Add By Sindy 2025/4/7 直接離開
            'Else
            ElseIf UCase(m_PrevForm.Name) <> UCase("frm04010515") Then
            '2025/4/7 END
               m_PrevForm.QueryData
            End If
         End If
      End If
   Next

   Unload frm090202_2_2
   'Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   g_LetterDebug = False 'Modify By Sindy 2025/7/9 取消記錄Log
   
   'Add by Sindy 2022/12/17 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q") = True Then
      Unload frm090801_Q
   End If
   '2022/12/17 END
   
   KillAttach
   Unload frm090202_2_2
   Set frm090202_2_2 = Nothing
   Set m_PrevForm = Nothing
   Set frm090202_2 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

'簽辦流程的人員角色
Private Sub EmpFlowRole()
Dim rsA As New ADODB.Recordset
Dim strCP13 As String
Dim strCP14 As String
Dim strCP29 As String
Dim strEP03 As String
Dim strEP04 As String
Dim strEP40 As String
Dim strCP29ST06 As String
Dim strText As String
Dim arrID As Variant 'Add By Sindy 2017/6/15
Dim strST16 As String, strST70 As String
Dim strEP05 As String 'Add By Sindy 2024/7/11
Dim strTmpCP13 As String, strTmpCP14 As String
   
   m_SPMan = "": m_EPMan = "": m_DPMan = "": m_EMMan = "": m_CMMan = "": m_DMMan = "": m_CSMan = "": m_NPMan = ""
   m_DCMan = "" '草圖核稿人 Add By Sindy 2015/4/22
   m_EP41 = "" 'Add By Sindy 2015/3/16 核稿語文
   m_CP14_2 = "" 'Add By Sindy 2017/6/15
   m_F21CMMan = "" 'Add By Sindy 2023/10/2
   
   If rsA.State <> adStateClosed Then rsA.Close
   'Modify By Sindy 2015/3/16 +,EP41
   'Modify By Sindy 2024/7/11 +,EP05
   strExc(0) = "select CP12,CP13,CP14,CP29,EP03,EP04,EP40,EP41,EP05" & _
               " From caseprogress,engineerprogress" & _
               " where cp09='" & m_EEP01 & "' and cp09=ep02(+)"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      strCP13 = Trim("" & rsA.Fields("CP13"))
      strCP14 = Trim("" & rsA.Fields("CP14"))
      strCP29 = Trim("" & rsA.Fields("CP29"))
      strEP03 = Trim("" & rsA.Fields("EP03"))
      strEP04 = Trim("" & rsA.Fields("EP04"))
      strEP40 = Trim("" & rsA.Fields("EP40"))
      m_EP41 = Trim("" & rsA.Fields("EP41")) 'Add By Sindy 2015/3/16 核稿語文
      strEP05 = Trim("" & rsA.Fields("EP05")) 'Add By Sindy 2024/7/11 抓外商承辦人
   End If
   '以防資料已從工作進度畫面異動
   If UCase(m_PrevForm.Name) = UCase("frm090711") Then '繪圖人員工作進度
      strCP29 = Trim(m_PrevForm.txt1(0))
   ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_2") Then '專利處承辦人工作進度
      If Trim(m_PrevForm.Combo2.Text) <> "" Then
         strCP29 = Left(Trim(m_PrevForm.Combo2.Text), 5)
      Else
         strCP29 = ""
      End If
      If Trim(m_PrevForm.Combo4.Text) <> "" Then
         strEP03 = Left(Trim(m_PrevForm.Combo4.Text), 5)
      Else
         strEP03 = ""
      End If
      strEP04 = Trim(m_PrevForm.txt1(5).Text)
      'Modify By Sindy 2015/5/22
      'strEP40 = Trim(m_PrevForm.txt1(22).Text)
      strEP40 = Trim(Left(m_PrevForm.Combo6.Text, 6))
      '2015/5/22 END
      m_EP41 = Trim(m_PrevForm.txt1(23).Text) 'Add By Sindy 2015/3/16 核稿語文
   'Add By Sindy 2018/4/20
   ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_b") Then '商標處承辦人工作進度
      strEP04 = Trim(Left(m_PrevForm.Combo2.Text, 6)) '核稿人
      strEP40 = Trim(Left(m_PrevForm.Combo6.Text, 6)) '判發人
   '2018/4/20 END
   ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_d") Then '法務,顧問承辦人工作進度
      strEP04 = Trim(Left(m_PrevForm.Combo2.Text, 6)) '核稿人
      strEP40 = Trim(Left(m_PrevForm.Combo6.Text, 6)) '判發人
   'Add By Sindy 2023/9/23
   ElseIf UCase(m_PrevForm.Name) = UCase("frm090909") Then '外專工作進度資料維護
      '英文核稿人
      If Trim(m_PrevForm.Combo4.Text) <> "" Then
         strEP03 = Left(Trim(m_PrevForm.Combo4.Text), 5)
      Else
         strEP03 = ""
      End If
      strEP04 = Trim(Left(m_PrevForm.Combo2.Text, 6)) '核稿人
      strEP40 = Trim(Left(m_PrevForm.Combo6.Text, 6)) '判發人
   End If
   If strCP29 = "99999" Then strCP29 = "" 'Add By Sindy 2013/9/5 '不繪圖
   'Modify By Sindy 2024/9/30 內專繪圖人員休假,則以操作的人抓核判權限
   '                          ex:姍珊請假時，舒郁依自己權限處理姍珊墨圖送判。
   If strCP29 <> "" Then
      If UCase(m_PrevForm.Name) = UCase("frm090711") Then '繪圖人員工作進度
         If ChkEmpIsRest(strCP29) = True Then  '休假
            strCP29 = strUserNum
         End If
      End If
   '2024/9/30 END
      strCP29ST06 = PUB_GetST06(strCP29)
   End If
   
   '智權人員
   If strCP13 <> "" Then
      'Modify By Sindy 2016/3/24 檢查智權人員是否離職
      If ChkStaffST04(strCP13, False) = True Then
         'Modify By Sindy 2021/6/29 目前智權人員
         strCP13 = ShowCurrCP13(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)), m_Country)
'         If PField(1) = "FCP" Or PField(1) = "FG" Then
'            strCP13 = PUB_GetFCPSalesNo(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)))
'         ElseIf PField(1) = "FCL" Or PField(1) = "LIN" Then
'            strCP13 = PUB_GetFCLSalesNo(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)))
'         ElseIf PField(1) = "FCT" Then
'            strCP13 = PUB_GetFCTSalesNo(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)))
'         ElseIf PField(1) = "S" Then
'            If m_Country = "000" Then
'               strCP13 = PUB_GetFCTSalesNo(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)))
'            Else
'               strCP13 = PUB_GetAKindSalesNo(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)))
'            End If
'         Else
'            strCP13 = PUB_GetAKindSalesNo(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)))
'         End If
         '2021/6/29 END
      End If
      '2016/3/24 END
      m_SPMan = strCP13 & " " & GetPrjSalesNM(strCP13)
      'Add By Sindy 2019/7/9 員工編號小於6視為沒有抓到資料,人員自行輸入
      'Modified by Morgan 2024/8/16 30015除外
      If Left(Trim(strCP13), 1) < "6" And strCP13 <> "30015" Then
         'Add By Sindy 2024/10/30 李柏翰經理和杜協理討論後:修改設定，當智權人員為XX備用時，智權人員為該區的區主管
         '                        和秀玲討論後,抓ST14
         strText = Trim(PUB_GetST14(strCP13))
         If strText <> "" Then
            strText = Left(strText, 5)
            m_SPMan = strText & " " & GetPrjSalesNM(strText)
         Else
         '2024/10/30 END
            m_SPMan = ""
         End If
      End If
      '2019/7/9 END
   End If
   '承辦人
   'Add By Sindy 2024/7/11
   If bolFCTFlow = True Or bolCFTFlow = True Then
      m_EPMan = strEP05 & " " & GetPrjSalesNM(strEP05)
   '2024/7/11 END
   ElseIf strCP14 <> "" Then
      '若為外翻人員
      If Left(strCP14, 1) = "F" Then
         'Add By Sindy 2024/3/13 + if
         If bolPAFlow = True Then
         '2024/3/13 END
            'Modify By Sindy 2014/5/26 Mark
            'Modify By Sindy 2014/6/26 解開Mark必須先抓取ST14值,因CFP-026830蔣正偉的系統操作人員為89026.張偉成
            strText = Trim(PUB_GetST14(strCP14))
            If strText <> "" Then
               strCP14 = strText
            'Modify By Sindy 2023/9/23 + If bolPAFlow = True Then
            ElseIf bolPAFlow = True Then
               strText = Trim(Pub_GetSpecMan("H"))
               If strText <> "" Then
                  strCP14 = strText
               End If
            'Add By Sindy 2023/9/23
            Else
               strCP14 = ""
               '2023/9/23 END
            End If
         'Add By Sindy 2024/3/13
         Else
            strCP14 = ""
         '2024/3/13 END
         End If
         
         'Added by Lydia 2017/03/28 ST14改成多個編號,所以只抓第一位
         If InStr(strCP14, ",") > 0 Or InStr(strCP14, ";") > 0 Then
            strCP14 = Replace(strCP14, ";", ",")
            '檢查人員是否在職
            If InStr(strCP14, ",") > 0 Then
               arrID = Split(strCP14, ",")
               strCP14 = ""
               For intI = 0 To UBound(arrID)
                  If ChkStaffST04(CStr(arrID(intI)), False) = False Then
                     strCP14 = strCP14 & "," & arrID(intI)
                  End If
               Next intI
               strCP14 = Mid(strCP14, 2)
            End If
            m_CP14_2 = strCP14 'Add By Sindy 2017/6/15 記錄原資料
            'Added by Lydia 2017/03/28
            'strCP14 = Replace(Replace(Mid(strCP14, 1, 6), ",", ""), ";", "")
            strCP14 = Mid(strCP14, 1, 5)
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
      End If
      'Add By Sindy 2023/6/15 專利處繪圖人員協助國外部案件,承辦人改抓工程師
      If PUB_GetST03(strCP14) = "P13" And Left(PUB_GetST03(strCP13), 1) = "F" Then
'         '抓該案最後一道A,B,C類發文之工程師
'         strExc(0) = "SELECT CP14,Max(CP27||CP09) Srt FROM CASEPROGRESS,STAFF" & _
'                     " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "'" & _
'                     " AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
'                     " AND CP27 IS NOT NULL AND CP14 IS NOT NULL AND CP14=ST01(+) AND ST15='F21'" & _
'                     " AND CP10<>'927' Group By CP14 Order By Srt desc "
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            strCP14 = RsTemp.Fields("CP14")
'         End If
         strExc(10) = ""
         strExc(10) = PUB_GetSpecCP14(pa(1) & pa(2) & pa(3) & pa(4)) '工程師
         If strExc(10) <> "" Then strCP14 = strExc(10)
         'Add By Sindy 2024/6/17
         If ChkStaffST04(strCP14, False) = True Then
            '抓該案最後一道A,B,C類發文之工程師
            strExc(0) = "SELECT CP14,Max(CP27||CP09) Srt FROM CASEPROGRESS,STAFF" & _
                        " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "'" & _
                        " AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                        " AND CP27 IS NOT NULL AND CP14 IS NOT NULL AND CP14=ST01(+) AND ST15='F21' AND ST04='1'" & _
                        " Group By CP14 Order By Srt desc "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCP14 = RsTemp.Fields("CP14")
            End If
         End If
         '2024/6/17 END
      End If
      '2023/6/15 END
      m_EPMan = strCP14 & " " & GetPrjSalesNM(strCP14)
   End If
   '繪圖人員
   If strCP29 <> "" Then
      m_DPMan = strCP29 & " " & GetPrjSalesNM(strCP29)
      'Modify By Sindy 2015/4/22 草圖核稿人及繪圖主管改設定在員工檔中
      If rsA.State <> adStateClosed Then rsA.Close
      strExc(0) = "select st62,st63 From staff where st01='" & strCP29 & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         m_DCMan = Trim("" & rsA.Fields("st62")) '草圖核稿人
         m_DMMan = Trim("" & rsA.Fields("st63")) '繪圖主管
         'Modify By Sindy 2024/6/25 增加檢查核判表是否有單獨設定
         Call PUB_ChkIsSetPromoterReader(strCP29, cp(1), cp(10), strPP04, strPP05, "", m_Country)
         If strPP04 <> "" Then m_DCMan = strPP04 '草圖核稿人
         If strPP05 <> "" Then m_DMMan = strPP05 '繪圖主管
         '2024/6/25 END
         If m_DCMan <> "" Then m_DCMan = m_DCMan & " " & GetPrjSalesNM(m_DCMan)
         If m_DMMan <> "" Then m_DMMan = m_DMMan & " " & GetPrjSalesNM(m_DMMan)
      End If
      If Pub_StrUserSt03 = "P13" Then 'P13 專利處繪圖
         If m_DCMan = "" Or m_DMMan = "" Then
            MsgBox GetPrjSalesNM(strCP29) & "尚未設定繪圖核判表，不可操作此作業!!!"
            cmdAdd.Enabled = False
            cmdSend.Enabled = False
            Exit Sub
         End If
      End If
      '2015/4/22 END
'      '繪圖主管
'      If strCP29ST06 = "1" Then '北所
'         'Modify By Sindy 2015/4/7
'         'm_DMMan = "72006" & " " & GetPrjSalesNM("72006")
'         If m_FlowUserNum = "87025" Then '陳翔龍
'            m_DMMan = "91010" & " " & GetPrjSalesNM("91010")
'         ElseIf m_FlowUserNum = "91010" Then '曾維揚
'            m_DMMan = "87025" & " " & GetPrjSalesNM("87025")
'         Else
'            '72006.張瓊玉
'            m_DMMan = "72006" & " " & GetPrjSalesNM("72006")
'         End If
'         '2015/4/7 END
'      ElseIf strCP29ST06 = "2" Then '中所
'         '82018.李月嬌
'         m_DMMan = "82018" & " " & GetPrjSalesNM("82018")
'      Else '其他
'         '78007.劉大愛
'         m_DMMan = "78007" & " " & GetPrjSalesNM("78007")
'      End If
      'Modify By Sindy 2013/9/6 繪圖人員90007.賴岑飛P案的設計案件自行核判不用經過主管
      'Modify By Sindy 2014/3/31 繪圖人員90007.賴岑飛內專外專的設計案件自行核判不用經過主管
      'If strCP29 = "90007" And m_PA08 = "3" And pfield(1) = "P" Then
      'Modify By Sindy 2018/5/16 岑飛的設計案件，判發人為李目嬌(依系統設定)
'      If strCP29 = "90007" And m_PA08 = "3" Then
'      '2014/3/31 END
'         m_DMMan = m_DPMan
'      End If
      '2013/9/6 END
   End If
   '英文核稿人
   If strEP03 <> "" Then
      m_EMMan = strEP03 & " " & GetPrjSalesNM(strEP03)
   End If
   '核稿人
   If strEP04 <> "" Then
      m_CMMan = strEP04 & " " & GetPrjSalesNM(strEP04)
   End If
   '判發人
   If strEP40 <> "" Then
      m_CSMan = strEP40 & " " & GetPrjSalesNM(strEP40)
   End If
   
   'Modify By Sindy 2013/9/9 程序人員
   'Modify By Sindy 2015/1/7
   'Modify By Sindy 2020/3/12
   If bolTMFlow = True Then
      m_NPMan = Pub_GetSpecMan("內商發文人員") '內商發文人員(P2002)
   'Add By Sindy 2024/7/11
   ElseIf bolFCTFlow = True Then
      m_NPMan = GetST52SelfList(Left(Trim(m_SPMan), 5), "st57") '外商程序人員
      m_NPMan = m_NPMan & " " & GetPrjSalesNM(m_NPMan)
   ElseIf bolCFTFlow = True Then
      m_NPMan = m_EPMan '發文人員=承辦人
      'Add By Sindy 2025/1/13 案件性質「304 申請英文證明」案件，操作送件歷程時，收受者為程序人員
      'Modify By Sindy 2025/10/20 CFT案承辦人員掛FCT智權人員時, 送件歷程之收受者設定為外商程序組 ex:CFT-025597
      'Modify By Sindy 2025/11/10 增加GetNA69做檢查,因May和琬姿會跨FCT和CFT業務
      '                           ex:May說CFT-25124承辦歷程的送件不知為什麼跳到江蕎安,無法發文
      strTmpCP13 = Left(Trim(m_SPMan), 5)
      strTmpCP14 = Left(Trim(m_EPMan), 5)
      If cp(10) = "304" _
         Or PUB_GetST03(strTmpCP14) = "F12" _
         Or (PUB_GetST03(strTmpCP13) = "F11" And PUB_GetST93(strTmpCP13) <> "T24" _
             And m_EPMan = m_SPMan _
             And GetNA69(strTmpCP14, m_Country, PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)), strExc(10), cp(1), cp(2), cp(3), cp(4)) = False _
            ) Then
         strExc(9) = ""
         If PUB_GetST06(Left(m_SPMan, 5)) = "1" Then '北所
            strExc(9) = Pub_GetSpecMan("CFT程序人員-北所")
         Else
            strExc(9) = Pub_GetSpecMan("CFT程序人員-分所")
         End If
         If strExc(9) <> "" Then m_NPMan = strExc(9) & " " & GetPrjSalesNM(strExc(9))
      End If
      '2025/1/13 END
   '2024/7/11 END
   ElseIf bolPAFlow = True Then
      'modify by sonia 2020/4/1未傳本所案號抓不到程序
      'm_NPMan = GetSignOffEmp("NP", CStr(PField(1)), CStr(PField(2)), m_Country,)
      m_NPMan = GetSignOffEmp("NP", CStr(PField(1)), CStr(PField(2)), m_Country, CStr(PField(1)) & CStr(PField(2)) & CStr(PField(3)) & CStr(PField(4)))
   'Add By Sindy 2021/9/7
   ElseIf PField(1) = "ACS" Then
      m_NPMan = m_EPMan '發文人員=承辦人
   '2021/9/7 END
   'Add By Sindy 2023/9/23
   ElseIf bolFCPFlow = True Then '外專
      '非寰華案件,由專利處程序發文; 排除FMPtoFCPSendCasePtyList(FMP非寰華案,有開放FCP程序人員操作發文的案件性質)及C類
      If (bolFMP = True And bolOurFMP = False) And _
         Not (Left(m_EEP01, 1) = "C" Or InStr(FMPtoFCPSendCasePtyList, cp(10)) > 0) Then
         m_NPMan = GetSignOffEmp("NP", CStr(PField(1)), CStr(PField(2)), m_Country, CStr(PField(1)) & CStr(PField(2)) & CStr(PField(3)) & CStr(PField(4)))
      Else
         '外專程序管制人
         m_NPMan = PUB_GetFCPHandler(PField(1), PField(2), PField(3), PField(4))
      End If
      
      'Add By Sindy 2024/3/8
      If strCP14 <> "" Then
         strSql = "SELECT st01,st16,st70 FROM staff" & _
                  " WHERE st01='" & strCP14 & "'"
      Else
      '2024/3/8 END
         'Add By Sindy 2024/1/5
         strSql = "SELECT tct10,st01,st16,st70 FROM TransCaseTitle,caseprogress,staff" & _
                  " WHERE cp01='" & PField(1) & "' and cp02='" & PField(2) & "' and cp03='" & PField(3) & "' and cp04='" & PField(4) & "'" & _
                  " and TCT01=cp09 and tct10 is not null and TCT10=st01(+)"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If RsTemp.RecordCount > 0 Then
         If "" & RsTemp.Fields("st01") <> "" Then
            If strCP14 = "" Then m_TCT10Man = RsTemp.Fields("tct10") & " " & GetPrjSalesNM(RsTemp.Fields("tct10"))
            strST16 = "" & RsTemp.Fields("st16")
            strST70 = "" & RsTemp.Fields("st70")
         End If
      End If
      '2024/1/5 END
      'FCP工程師主管
      If strCP14 <> "" Then
         'Modify By Sindy 2024/3/8
         If strST16 = "3" And strST70 <> "" Then '日文組
            m_F21CMMan = Pub_GetSpecMan(strST16 & strST70)
         Else
            m_F21CMMan = Pub_GetFCPGrpMan(strST16)
         End If
         '2024/3/8 END
'         m_F21CMMan = GetDeptMan(PUB_GetST93(strCP14), 2) 'PUB_GetFCPEngSup(strCP14, True)
      Else
         If m_TCT10Man <> "" Then
            'Modify By Sindy 2024/3/8
            If strST16 = "3" And strST70 <> "" Then '日文組
               m_F21CMMan = Pub_GetSpecMan(strST16 & strST70)
            Else
               m_F21CMMan = Pub_GetFCPGrpMan(strST16)
            End If
            '2024/3/8 END
'            m_F21CMMan = GetDeptMan(PUB_GetST93(Left(Trim(m_TCT10Man), 5)), 2)
'            If m_F21CMMan = "" And strST16 <> "" And strST70 <> "" Then
'               m_F21CMMan = Pub_GetSpecMan(strST16 & strST70)
'            End If
         End If
         If m_F21CMMan = "" Then
            m_F21CMMan = Pub_GetFCPGrpMan(pa(150))
         End If
      End If
      '組員編+姓名
      If m_F21CMMan <> "" Then
         m_F21CMMan = m_F21CMMan & " " & GetPrjSalesNM(m_F21CMMan)
      End If
   '2023/9/23 END
   End If
   If m_NPMan <> "" Then
      m_NPMan = m_NPMan & " " & GetPrjSalesNM(m_NPMan)
   End If
'   If PField(1) = "CFP" Then
'      strExc(0) = "select na73,na74 from nation " & _
'                  "where na01='" & m_Country & "' "
'      intI = 1
'      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If PField(2) Mod 2 = 0 Then '雙號
'            m_NPMan = "" & rsA.Fields("na74")
'         Else
'            m_NPMan = "" & rsA.Fields("na73")
'         End If
'         m_NPMan = m_NPMan & " " & GetPrjSalesNM(m_NPMan)
'      End If
'   Else
'      If bolTMFlow = True Then
'         m_NPMan = Pub_GetSpecMan("內商發文人員") '內商發文人員
'      Else
'         If m_Country = "000" Then '台灣案
'            m_NPMan = Pub_GetSpecMan("PS1")
'         Else '非台灣案
'            m_NPMan = Pub_GetSpecMan("PS2")
'         End If
'      End If
'      m_NPMan = m_NPMan & " " & GetPrjSalesNM(m_NPMan)
'   End If
'   '2015/1/7 END
   '2020/3/12 END
   
   'Add By Sindy 2013/9/14 檢查核稿人與承辦人是否相同,若相同則清除核稿人
   '                       檢查判發人與承辦人是否相同,若相同則清除判發人
   If m_EPMan = m_CMMan Then
      m_CMMan = ""
   'Add By Sindy 2023/10/18
   'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
   ElseIf bolFCPFlow = True _
      And (cp(10) = "201" Or cp(10) = "931") And Trim(m_CMMan) <> "" Then 'FCP的新案翻譯 'And Trim(m_EPMan) = "" ex:FCP-070271
      m_EPMan = m_CMMan
      m_CMMan = ""
   '2023/10/18 END
   End If
   If m_EPMan = m_CSMan Then m_CSMan = ""
   'Modify By Sindy 2014/4/3 Mark 因志建分析案有核稿人但自行判發(ex.P-099638 AA3013462)
'   'Modify By Sindy 2013/10/28 有核稿主管時,判發主管不可空白,則判發主管=核稿主管
'   If m_CMMan <> "" And m_CSMan = "" Then
'      m_CSMan = m_CMMan
'   End If
   '2013/10/28 END
   
'   '檢查是否有自行核判的權限
'   If intReceiveKind = 0 Then '0.承辦人工作進度
'      '自行核判：承辦人=核稿人 或無核稿人
'      If m_EPMan = m_CMMan Or m_CMMan = "" Then
'         bolSelfJudgement = True
'      End If
'   ElseIf intReceiveKind = 3 Then '3.繪圖人員工作進度
'      '自行核判：繪圖人員=繪圖主管
'      If m_DPMan = m_DMMan Then
'         bolSelfJudgement = True
'      End If
'   End If
   
   'Add By Sindy 2015/3/4 是否必需送英核
   If bolPAFlow = True Then 'Add By Sindy 2023/9/23 +if
      bolHadSetProofEngReader = PUB_ChkIsSetProofEngReader(Left(Trim(m_EPMan), 5), _
                                CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)), cp(10), m_PER04, bolMultinationalEngOk)
   End If
   
   Set rsA = Nothing
End Sub

'設定流程狀態及收受者下拉式選單
Private Sub SetStatusCombo()
Dim rsA As New ADODB.Recordset
Dim strEP14 As String, strEP15 As String, strEP17 As String
Dim strEP09 As String
Dim strEP08 As String, strEP38 As String
Dim strEP37 As String 'Add By Sindy 2018/8/28
Dim strRefEEP02 As String
Dim bolDoEmpFlow32Limit As Boolean 'Add By Sindy 2016/12/21 是否有執行准許先會的權限
Dim bolHadFlow32CanSMeet As Boolean 'Add By Sindy 2016/12/21 是否已准許先會可操作送會狀態
Dim bolCP143 As Boolean, strCP143 As String 'Added by Lydia 2019/11/22 是否有查名齊備管制和查名齊備日
   
   Frame4.Visible = False: CboCP10.Locked = False
   cmdMail.Visible = False 'Add By Sindy 2018/8/29
   '清除及預設值
   Call ClearData
   Call SetCtrlReadOnly(True)
   CboEEP04.Clear '流程狀態
   Call SetCboEEP05 '收受者
   txtEEP03 = strUserNum
   txtEEP03_2 = strUserName
   'Add By Sindy 2013/10/16 代理註明
   If m_FlowUserNum <> strUserNum Then
      m_EEP12 = "(代)"
      m_EEP16 = m_FlowUserNum 'Add By Sindy 2023/12/18
   Else
      m_EEP12 = ""
      m_EEP16 = "" 'Add By Sindy 2023/12/18
   End If
   '2013/10/16 END
   
   bolWaitReply = False
   
   bolCP143 = False  'Added by Lydia 2019/11/22
   '取得最後一筆流程狀態 (必須踢除00.聯絡此狀態)
   'm_strLastEEP04 : 記錄處理的流程狀態
   intLastEEP02 = 0: strLastEEP03 = "": m_strLastEEP04 = "": m_strLastEEP04Nm = "": strLastEEP11 = "": bolLastFile = False
   
   'Modify By Sindy 2016/5/3 先檢查 m_CurrFlowEEP02 是否有傳入欲處理歷程
   If m_CurrFlowEEP02 > 0 Then
      If rsA.State <> adStateClosed Then rsA.Close
      '因有聯絡問題,判斷收受者有待回覆的歷程
      strExc(0) = "select eep02,eep04 From empelectronprocess where eep01='" & m_EEP01 & "' and EEP05='" & m_FlowUserNum & "' and eep09='Y'"
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      '有欲處理的歷程
      If rsA.RecordCount > 0 Then
         intLastEEP02 = rsA.Fields(0)
         m_CurrFlowEEP02 = rsA.Fields(0)
      End If
   End If
   '2016/5/3 END
   If intLastEEP02 = 0 Then
      'Add By Sindy 2013/11/5
      If rsA.State <> adStateClosed Then rsA.Close
      'Modify By Sindy 2016/3/9 開放正在送英核時，可以操作其他歷程 ==> + and eep04<>'" & EMP_送英核 & "'
      strExc(0) = "select eep02,eep04 From empelectronprocess where eep01='" & m_EEP01 & "' and eep09='Y' and eep04<>'" & EMP_送英核 & "' order by eep02 desc"
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      '有待回覆的歷程
      If rsA.RecordCount > 0 Then
         intLastEEP02 = rsA.Fields(0)
      Else
         If rsA.State <> adStateClosed Then rsA.Close
         'Modify By Sindy 2016/4/15 + and not (eep09='Y' and eep04='" & EMP_送英核 & "')
         strExc(0) = "select eep02,eep04 From empelectronprocess where eep01='" & m_EEP01 & "'" & _
                     " and eep04 not in(" & EMP_流程控制除外的狀態 & ") and not (eep09='Y' and eep04='" & EMP_送英核 & "') order by eep02 desc"
         rsA.CursorLocation = adUseClient
         rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            intLastEEP02 = rsA.Fields(0)
         End If
      End If
      '2013/11/5 END
   End If
   'Modify By Sindy 2013/10/16 " and eep04 not in(" & EMP_流程控制除外的狀態 & ",'" & EMP_草完 & "','" & EMP_標號 & "'))"
'   If rsA.State <> adStateClosed Then rsA.Close
'   strExc(0) = "select *" & _
'               " From empelectronprocess,allcode" & _
'               " where eep01='" & m_EEP01 & "'" & _
'               " and eep02=(select max(eep02) From empelectronprocess where eep01='" & m_EEP01 & "'" & _
'               " and eep04 not in(" & EMP_流程控制除外的狀態 & "))" & _
'               " and ac01='09' And eep04=ac02(+)"
   'Modify By Sindy 2013/11/5 直接在eep04 not in 中增加排除草完和標號會導至無法延用這2項的附件,因此改寫SQL
   If rsA.State <> adStateClosed Then rsA.Close
   strExc(0) = "select *" & _
               " From empelectronprocess,allcode" & _
               " where eep01='" & m_EEP01 & "'" & _
               " and eep02=" & intLastEEP02 & _
               " and ac01='09' And eep04=ac02(+)"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'Modify By Sindy 2013/11/5 Mark if
'      If intReceiveKind = 0 Or intReceiveKind = 3 Then '0.承辦人工作進度 及 3.繪圖人員工作進度時,才要鎖住
      '2013/11/5 END
         '檢查最後一道若待回覆=Y時,則不可新增下一流程
         'Modify By Sindy 2016/3/9 開放正在送英核時，可以操作其他歷程 ==> + And "" & rsA.Fields("EEP04") <> EMP_送英核
         'Modify By Sindy 2016/5/3 取消 And "" & rsA.Fields("EEP04") <> EMP_送英核, 改前面控管
         'If "" & rsA.Fields("EEP09") = "Y" And "" & rsA.Fields("EEP04") <> EMP_送英核 Then
         If "" & rsA.Fields("EEP09") = "Y" Then
         '2016/5/3 END
'            Me.cmdSend.Enabled = False
'            Me.cmdAdd.Visible = True
'            Me.cmdAdd.Enabled = False
'            Me.cmdCancel.Visible = False
'            MsgBox "此文尚未回覆，不可執行下一流程！"
'            Set rsA = Nothing
'            Exit Sub
            bolWaitReply = True
         End If
'      End If
      strLastEEP03 = rsA.Fields("EEP03")
      m_strLastEEP04 = rsA.Fields("EEP04")
      m_strLastEEP04Nm = rsA.Fields("AC03")
      strLastEEP05 = "" & rsA.Fields("EEP05")
      strLastEEP11 = "" & rsA.Fields("EEP11")
      
      'Modify By Sindy 2017/8/14 王副總提出歷程判發後還是可以開放聯絡
'      'Add By Sindy 2013/9/27
'      'Modify By Sindy 2013/10/22
'      'If m_strLastEEP04 = EMP_判發 Then
'      If m_strLastEEP04 = EMP_判發 Or m_strLastEEP04 = EMP_退件重送 Then
'      '2013/10/22 END
'         Me.cmdAdd.Visible = False
'         Me.cmdCancel.Visible = False
'         Me.cmdSend.Visible = False
'         Me.cmdAddAtt(0).Enabled = False
'         Me.cmdRemAtt(0).Enabled = False
'         Me.cmdAddAtt(1).Enabled = False
'         Me.cmdRemAtt(1).Enabled = False
'         MsgBox "此文已判發，不可執行歷程作業！"
'         Set rsA = Nothing
'         Exit Sub
'      End If
'      '2013/9/27 END
      
      '檢查是否有附件
      If rsA.State <> adStateClosed Then rsA.Close
      'Modify By Sindy 2018/9/4 + and eef12 is not null
      strExc(0) = "select eef03" & _
                  " From empelectronfile" & _
                  " where eef01='" & m_EEP01 & "' and eef02=" & intLastEEP02 & " and eef12 is not null"
      rsA.CursorLocation = adUseClient
      rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         bolLastFile = True
      End If
   End If
   
   '檢查是否為附加流程案件
   If rsA.State <> adStateClosed Then rsA.Close
   strExc(0) = "select eep03" & _
               " From empelectronprocess" & _
               " where eep01='" & m_EEP01 & "' and eep02=1 and eep04='" & EMP_附加流程 & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   bolBCaseFlow = False
   If rsA.RecordCount > 0 Then
      bolBCaseFlow = True
   End If
   'Add By Sindy 2014/1/10
   '檢查是否有聯絡送英核流程
   'Modify By Sindy 2015/3/16 +or instr(eep08,'[送日核]')>0
   If rsA.State <> adStateClosed Then rsA.Close
   strExc(0) = "select eep01,eep03" & _
               " From empelectronprocess" & _
               " where eep01='" & m_EEP01 & "' and eep04='" & EMP_聯絡 & "' and (instr(eep08,'[送英核]')>0 or instr(eep08,'[送日核]')>0)"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   bol00EngCMFlow = False
   If rsA.RecordCount > 0 Then
      bol00EngCMFlow = True
      bol00EngCMFlowEmp = rsA.Fields("eep03") 'Add By Sindy 2017/11/30
   End If
   '2014/1/10 END
   
   '承辦繪圖資料
   strEP14 = "": strEP15 = "": strEP17 = "": m_EP18 = ""
   m_EP01 = "": strEP09 = "": m_EP34 = "": m_EP07 = ""
   strEP08 = "": strEP38 = "": strEP37 = "": m_EP39 = "": m_EP33 = "": m_EP42 = ""
   m_EP08 = "" 'Add By Sindy 2013/9/11
   If rsA.State <> adStateClosed Then rsA.Close
   strExc(0) = "select *" & _
               " From engineerprogress,caseprogress" & _
               " where ep02='" & m_EEP01 & "' and ep02=cp09"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      '未發文未取消收文時,送出鍵才顯示
      If Val("" & rsA.Fields("CP27")) > 0 Or Val("" & rsA.Fields("CP57")) > 0 Then
         Me.cmdAdd.Visible = False
         Me.cmdCancel.Visible = False
         Me.cmdSend.Visible = False
         Me.cmdSave2.Visible = False 'Add By Sindy 2023/9/20
         Me.cmdSave3.Visible = False 'Add By Sindy 2023/9/20
         Me.Frame945.Enabled = False 'Add By Sindy 2025/9/4
         Me.cmdAddAtt(0).Enabled = False
         CmdF21(0).Enabled = False 'Add By Sindy 2025/10/28
         Me.cmdRemAtt(0).Enabled = False
         Me.cmdAddAtt(1).Enabled = False
         Me.cmdRemAtt(1).Enabled = False
         MsgBox "此文" & IIf(Val("" & rsA.Fields("CP27")) > 0, "已發文", IIf(Val("" & rsA.Fields("CP57")) > 0, "已取消收文", "")) & "，不可執行歷程作業！"
         Set rsA = Nothing
         Exit Sub
      End If
      strEP14 = "" & rsA.Fields("EP14") '草圖齊備日
      strEP15 = "" & rsA.Fields("EP15") '草圖完稿日
      strEP17 = "" & rsA.Fields("EP17") '墨圖齊備日
      m_EP18 = "" & rsA.Fields("EP18") '墨圖完稿日
      m_EP01 = "" & rsA.Fields("EP01")  '當月目次
      m_EP06 = "" & rsA.Fields("EP06") '文件齊備日
      strEP09 = "" & rsA.Fields("EP09") '完稿日
      m_EP34 = "" & rsA.Fields("EP34") '是否會稿
      m_EP07 = "" & rsA.Fields("EP07") '會稿日
      strEP08 = "" & rsA.Fields("EP08") '會稿完成日
      strEP38 = "" & rsA.Fields("EP38") '智權人員會稿完成日
      strEP37 = Val("" & rsA.Fields("EP37")) '客戶會稿日 Add By Sindy 2018/8/28
      m_EP39 = "" & rsA.Fields("EP39") '核稿完成日
      m_EP42 = "" & rsA.Fields("EP42") '判發完成日
      m_EP33 = "" & rsA.Fields("EP33") '英文核完日
      'Add By Sindy 2018/9/20
      If "" & rsA.Fields("EP11") = "N" Then '是否通知客戶
         ChkEP11.Value = 1
      Else
         ChkEP11.Value = 0
      End If
      '2018/9/20 END
   End If
   '以防資料已從工作進度畫面異動
   If UCase(m_PrevForm.Name) = UCase("frm090711") Then '繪圖人員工作進度
      'Add By Sindy 2016/5/9
      If Trim(m_PrevForm.lbl1(1)) = m_EEP01 Then
      '2016/5/9 END
         'Modify By Sindy 2018/10/1 畫面上的欄位值是空的時,不需抓取其欄位值,以上面DB資料為主
         If Trim(m_PrevForm.txt1(1)) <> "" Then strEP14 = DBDATE(Trim(m_PrevForm.txt1(1))) '草圖齊備日
         If Trim(m_PrevForm.txt1(2)) <> "" Then strEP15 = DBDATE(Trim(m_PrevForm.txt1(2))) '草圖完稿日
         If Trim(m_PrevForm.txt1(4)) <> "" Then strEP17 = DBDATE(Trim(m_PrevForm.txt1(4))) '墨圖齊備日
         If Trim(m_PrevForm.txt1(5)) <> "" Then m_EP18 = DBDATE(Trim(m_PrevForm.txt1(5))) '墨圖完稿日
      End If
   ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_2") Then '承辦人工作進度
      'Add By Sindy 2016/5/9
      If Trim(m_PrevForm.lbl1(3)) = m_EEP01 Then
      '2016/5/9 END
         'Modify By Sindy 2018/10/1 畫面上的欄位值是空的時,不需抓取其欄位值,以上面DB資料為主
         If Trim(m_PrevForm.txt1(2)) <> "" Then
            m_EP06 = DBDATE(Trim(m_PrevForm.txt1(2))) '文件齊備日
         'Modify By Sindy 2022/11/1 DB已存入齊備日,沒回前畫面,又更新為空白
         ElseIf m_EP06 <> "" Then
            m_PrevForm.txt1(2) = TAIWANDATE(m_EP06) '文件齊備日
         End If
         '2022/11/1 END
         If Trim(m_PrevForm.txt1(3)) <> "" Then strEP09 = DBDATE(Trim(m_PrevForm.txt1(3))) '完稿日
         If Trim(m_PrevForm.txt1(1)) <> "" Then m_EP34 = Trim(m_PrevForm.txt1(1))         '是否會稿
         If Trim(m_PrevForm.txt1(4)) <> "" Then m_EP07 = DBDATE(Trim(m_PrevForm.txt1(4))) '會稿日
         If Trim(m_PrevForm.txt1(7)) <> "" Then strEP08 = DBDATE(Trim(m_PrevForm.txt1(7))) '會稿完成日
         If Trim(m_PrevForm.txt1(19)) <> "" Then m_EP33 = DBDATE(Trim(m_PrevForm.txt1(19))) '英文核完日
      End If
   'Add By Sindy 2023/9/27
   ElseIf UCase(m_PrevForm.Name) = UCase("frm090909") Then '外專承辦人工作進度
      If Trim(m_PrevForm.lbl1(3)) = m_EEP01 Then
         '畫面上的欄位值是空的時,不需抓取其欄位值,以上面DB資料為主
         If Trim(m_PrevForm.txt1(2)) <> "" Then
            m_EP06 = DBDATE(Trim(m_PrevForm.txt1(2))) '文件齊備日
         'DB已存入齊備日,沒回前畫面,又更新為空白
         ElseIf m_EP06 <> "" Then
            m_PrevForm.txt1(2) = TAIWANDATE(m_EP06) '文件齊備日
         'Add By Sindy 2023/11/14 秀玲決定預設為分案日
         Else
            m_EP06 = cp(149)
            m_PrevForm.txt1(2) = TAIWANDATE(cp(149))
            '2023/11/14 END
         End If
         If Trim(m_PrevForm.txt1(3)) <> "" Then strEP09 = DBDATE(Trim(m_PrevForm.txt1(3))) '完稿日
         If Trim(m_PrevForm.txt1(7)) <> "" Then strEP08 = DBDATE(Trim(m_PrevForm.txt1(7))) '核稿期限
         If Trim(m_PrevForm.txt1(19)) <> "" Then m_EP33 = DBDATE(Trim(m_PrevForm.txt1(19))) '英文核完日
      End If
   'Add By Sindy 2018/4/20
   ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_b") Then '商標處承辦人工作進度
      If Trim(m_PrevForm.lbl1(3)) = m_EEP01 Then
         'Modify By Sindy 2018/10/1 畫面上的欄位值是空的時,不需抓取其欄位值,以上面DB資料為主
         If Trim(m_PrevForm.txt1(2)) <> "" Then
            m_EP06 = DBDATE(Trim(m_PrevForm.txt1(2))) '文件齊備日
         'Modify By Sindy 2022/11/1 通知智權人員輸入齊備日,沒回前畫面,又更新為空白
         ElseIf m_EP06 <> "" Then
            m_PrevForm.txt1(2) = TAIWANDATE(m_EP06) '文件齊備日
         'Add By Sindy 2024/8/7 同外專預設為分案日
         ElseIf bolCFTFlow = True Or bolFCTFlow = True Then
            m_EP06 = cp(149)
            m_PrevForm.txt1(2) = TAIWANDATE(cp(149))
            '2024/8/7 END
         End If
         '2022/11/1 END
         If Trim(m_PrevForm.txt1(3)) <> "" Then strEP09 = DBDATE(Trim(m_PrevForm.txt1(3))) '完稿日
         If Trim(m_PrevForm.txt1(1)) <> "" Then m_EP34 = Trim(m_PrevForm.txt1(1))         '是否會稿
         If Trim(m_PrevForm.txt1(4)) <> "" Then m_EP07 = DBDATE(Trim(m_PrevForm.txt1(4))) '會稿日
         If Trim(m_PrevForm.txt1(7)) <> "" Then strEP08 = DBDATE(Trim(m_PrevForm.txt1(7))) '會稿完成日
         If Trim(m_PrevForm.txt1(19)) <> "" Then m_EP33 = DBDATE(Trim(m_PrevForm.txt1(19))) '英文核完日 Add By Sindy 2024/12/6
         'Add By Sindy 2018/9/20
         If m_PrevForm.txt1(9) = "N" Then '是否通知客戶
            ChkEP11.Value = 1
         Else
            ChkEP11.Value = 0
         End If
         '2018/9/20 END
         If Trim(m_PrevForm.txt1(11)) <> "" Then txt2 = m_PrevForm.txt1(11) '條款 Add By Sindy 2018/9/25
         'Added by Lydia 2019/11/22 記錄前一畫面是否有查名齊備管制和查名齊備日
         If m_PrevForm.Label1(9).Tag = "Y" Then
              bolCP143 = True: strCP143 = Trim(m_PrevForm.textCP143.Text)
         End If
         'end 2019/11/22
      End If
   'Add By Sindy 2021/9/24
   ElseIf UCase(m_PrevForm.Name) = UCase("frm090201_d") Then '法務,顧問承辦人工作進度
      If Trim(m_PrevForm.lbl1(3)) = m_EEP01 Then
         'Modify By Sindy 2018/10/1 畫面上的欄位值是空的時,不需抓取其欄位值,以上面DB資料為主
         If Trim(m_PrevForm.txt1(2)) <> "" Then
            m_EP06 = DBDATE(Trim(m_PrevForm.txt1(2))) '文件齊備日
         'Modify By Sindy 2022/11/1 DB已存入齊備日,沒回前畫面,又更新為空白
         ElseIf m_EP06 <> "" Then
            m_PrevForm.txt1(2) = TAIWANDATE(m_EP06) '文件齊備日
         End If
         '2022/11/1 END
         If Trim(m_PrevForm.txt1(3)) <> "" Then strEP09 = DBDATE(Trim(m_PrevForm.txt1(3))) '完稿日
         If Trim(m_PrevForm.txt1(1)) <> "" Then m_EP34 = Trim(m_PrevForm.txt1(1))         '是否會稿
         If Trim(m_PrevForm.txt1(4)) <> "" Then m_EP07 = DBDATE(Trim(m_PrevForm.txt1(4))) '會稿日
         If Trim(m_PrevForm.txt1(7)) <> "" Then strEP08 = DBDATE(Trim(m_PrevForm.txt1(7))) '會稿完成日
      End If
   End If
   m_EP08 = strEP08 'Add By Sindy 2013/9/11
   
   'Modify By Sindy 2025/4/10 mark,無效的程式段
'   'Modify By Sindy 2018/9/21 是否通知客戶
'   If bolTMFlow = True Or bolCFTFlow = True Then
'      If PUB_ChkEmpFlowExists(m_EEP01, EMP_會完) = True And Left(lblCP09, 1) = "C" Then
'         ChkEP11.Visible = True
'         ChkEP11.Enabled = False
'      End If
'   End If
   
   If bolPAFlow = True Then
      'Add By Sindy 2016/12/21 檢查准許先會狀況
      '是否有需要准許先會的權限
      bolDoEmpFlow32Limit = False
      '是否已准許先會可操作送會狀態
      bolHadFlow32CanSMeet = False
      'Modify By Sindy 2020/6/29 雅娟反應映秀在說外翻人員沒有准許先會
      '發現外翻人員不會設英核表; 暫定if先放開, 不須鎖那麼死的條件 CFP-030418
      'If bolHadSetProofEngReader = True Then '必需送英核
      '2020/6/29 END
         '1.
         If rsA.State <> adStateClosed Then rsA.Close
         strExc(0) = "select eep01,eep02" & _
                     " From empelectronprocess" & _
                     " where eep01='" & m_EEP01 & "' and eep04='" & EMP_送英核 & "' and eep09='Y'" '送英核中
         rsA.CursorLocation = adUseClient
         rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            'Y.要會稿 且 有准許先會的權限者
            If m_EP34 = "Y" And InStr(Pub_GetSpecMan("可准許先會之主管"), strUserNum) > 0 Then
               'Modify By Sindy 2019/12/27 英核准予先會功能請控制有權限的人和承辦人若是同一人,則無法操件,讓其找主管處理
               'If strUserNum <> m_FlowUserNum Then
               If strUserNum <> Left(m_EPMan, 5) Then 'Modify Sindy 2020/3/20
               '2019/12/27 END
                  bolDoEmpFlow32Limit = True '有執行准許先會的權限
               End If
            End If
         End If
         If PUB_ChkEmpFlowExists(m_EEP01, EMP_准許先會) = True Or _
            PUB_ChkEmpFlowExists(m_EEP01, EMP_送會) = True Then
            bolDoEmpFlow32Limit = False
         End If
         '2.
         If rsA.State <> adStateClosed Then rsA.Close
         strExc(0) = "select eep01,eep02" & _
                     " From empelectronprocess" & _
                     " where eep01='" & m_EEP01 & "' and eep04='" & EMP_送英核 & "' and eep09 is null" '已送英核
         rsA.CursorLocation = adUseClient
         rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            bolHadFlow32CanSMeet = True '已准許先會可操作送會狀態
         ElseIf PUB_ChkEmpFlowExists(m_EEP01, EMP_准許先會) = True Or _
                PUB_ChkEmpFlowExists(m_EEP01, EMP_送會) = True Then
            bolHadFlow32CanSMeet = True '已准許先會可操作送會狀態
         End If
      'End If
      '2016/12/21 END
   End If
   
   'Modify By Sindy 2021/4/12 原寫在”新增下一流程(&A)”按鍵裡,調整至此
   Me.cmdAdd.Visible = False
   Me.cmdCancel.Visible = True
   Me.cmdSend.Enabled = True
   m_EditMode = 1 '新增
   '2021/4/12 END
   
   '******************************
   '新增各流程狀態
   '******************************
   'Add By Sindy 2013/9/23
   If intReceiveKind = 0 Then
      CboCP10.Clear
      'Add By Sindy 2018/9/19
      If bolTMFlow = True Then '商標處承辦人工作進度
         If PUB_ChkCPMIsExists(CStr(PField(1)), "303", m_Country) = True And cp(10) <> "303" Then
            CboCP10.AddItem "303  延期"
         End If
         'Add By Sindy 2019/7/9 T非台灣案,+734.代理人撰稿
         If PField(1) = "T" And m_Country <> "000" And cp(10) <> "734" Then
            CboCP10.AddItem "734  代理人撰稿"
         End If
      'Add By Sindy 2024/3/1
      ElseIf bolFCPFlow = True Then '外專承辦人工作進度
         If PUB_ChkCPMIsExists(CStr(PField(1)), "404", m_Country) = True And m_Country = "000" And Val(cp(7)) > 0 And cp(10) <> "404" Then
            CboCP10.AddItem "404  延期"
         End If
         'Add By Sindy 2025/7/31
         If PUB_ChkCPMIsExists(CStr(PField(1)), "936", m_Country) = True And cp(10) <> "936" Then
            CboCP10.AddItem "936  回覆委任代理人"
         End If
         If PUB_ChkCPMIsExists(CStr(PField(1)), "957", m_Country) = True And cp(10) <> "957" Then
            CboCP10.AddItem "957  詢問代理人"
         End If
         If PUB_ChkCPMIsExists(CStr(PField(1)), "958", m_Country) = True And cp(10) <> "958" Then
            CboCP10.AddItem "958  代理人撰稿"
         End If
         '2025/7/31 END
         
      '2024/3/1 END
      ElseIf bolPAFlow = True Then '專利處承辦人工作進度
      '2018/9/19 END
         'P的台灣案才可延期
         If PField(1) = "P" And m_Country = "000" And Val(cp(7)) > 0 And cp(10) <> "404" Then
            If PUB_ChkCPMIsExists(CStr(PField(1)), "404", m_Country) = True Then
               CboCP10.AddItem "404  延期"
            End If
         End If
         If PUB_ChkCPMIsExists(CStr(PField(1)), "936", m_Country) = True And cp(10) <> "936" Then
            CboCP10.AddItem "936  回覆委任代理人"
         End If
         'Add By Sindy 2018/10/9
         If PUB_ChkCPMIsExists(CStr(PField(1)), "957", m_Country) = True And cp(10) <> "957" And _
            Not (PField(1) = "P" And m_Country = "000") Then
            CboCP10.AddItem "957  詢問代理人"
         End If
         '2018/10/9 END
         'Add By Sindy 2019/5/7
         '對於不是自行核判之工程師，主程序必須要跑完核完後才可附加"代理人撰稿"流程
         If (m_CMMan = "" Or Left(m_CMMan, 5) = m_FlowUserNum) Or _
            (m_CMMan <> "" And Val(m_EP39) > 0) Then
            If PUB_ChkCPMIsExists(CStr(PField(1)), "958", m_Country) = True And cp(10) <> "958" And _
               Not (PField(1) = "P" And m_Country = "000") Then
               CboCP10.AddItem "958  代理人撰稿"
            End If
         End If
         '2019/5/7 END
      End If
      If CboCP10.ListCount > 0 Then
         CboEEP04.AddItem EMP_附加流程 & " " & "附加流程"
      End If
   End If
   '2013/9/23 END
   CboEEP04.AddItem EMP_聯絡 & " " & "聯絡"
   'Add By Sindy 2017/8/14 王副總提出歷程判發(中)後還是可以開放聯絡
   'Modify By Sindy 2018/4/30 歷程已到程序人員
   'Add By Sindy 2023/11/9 外專會在待送件區做程序送判
   If UCase(m_PrevForm.Name) = UCase("frm090202_4_1") Then
      frm090202_2.cmdCancel.Enabled = False
   End If
   If UCase(m_PrevForm.Name) = UCase("frm090202_4_1") And intReceiveKind = 5 Then
      If PUB_ChkEmpFlowExists(m_EEP01, EMP_程序送判) = False Or bolWaitReply = False Then
         CboEEP04.AddItem EMP_程序送判 & " " & "程序送判"
      End If
   Else
   '2023/11/9 END
      '外專程序才會操作 程序送判,程序退回
      'Modify By Sindy 2024/8/14 外商程序也會操作 程序送判,程序退回
      If m_strLastEEP04 = EMP_送件 Or m_strLastEEP04 = EMP_退件重送 Or m_strLastEEP04 = EMP_發文歸檔 _
         Or (m_strLastEEP04 = EMP_判發 And (bolPAFlow = True Or bolOtherFlow = True)) Or _
         ((m_strLastEEP04 = EMP_程序送判 Or m_strLastEEP04 = EMP_程序退回) And Pub_StrUserSt03 <> "F22" And Pub_StrUserSt03 <> "F12") Then
         'Modify By Sindy 2023/1/5 要鎖,不然工程師會重覆操作
'         'Add By Sindy 2023/12/6 外專排除
'         If bolFCPFlow = False Then
'         '2023/12/6 END
            Exit Sub
'         End If
      End If
   End If
   '2017/8/14 END
   
   'Add By Sindy 2019/6/27 商標處程序人員才會顯示此歷程選項
   If bolTMFlow = True And cp(10) = "101" And _
      GetStaffDepartment(m_FlowUserNum) = "P22" Then
      
      'Add By Sindy 2022/6/22 程序組要在共同查詢操作查名結果 ex:T-239802
      If txtNote.Visible = True And UCase(m_PrevForm.Name) = "FRM100101_2" And Pub_StrUserSt03 = "P22" Then '程序組-共同查詢
         CboEEP04.Clear
      End If
      '2022/6/22 END
      
      'CboEEP04.AddItem EMP_查名 & " " & "查名"
      CboEEP04.AddItem EMP_查名結果 & " " & "查名結果"
   End If
   
   'Add By Sindy 2016/3/7 已有會稿完成日未送判或未判發時，智權人員可以起「會完重修」給工程師，此時系統會把會稿完成日拿掉，工程師即可再做送會歷程
   If m_FlowUserNum = Trim(Left(m_SPMan, 6)) Then
      'Modify By Sindy 2020/4/8 ex:T-226712,T-226713:判發後又要會完重修
      'Add By Sindy 2016/9/8 會完重修不可出現在承辦人工作進度維護及待核判區
      'ex.CFP-028822:CP14=A0029,CP13=A0029
      If intReceiveKind = 2 Or intReceiveKind = 99 Then
      '2016/9/8 END
         If m_EP08 <> "" Then
            '商標
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
               If m_strLastEEP04 <> EMP_送件 And _
                  m_strLastEEP04 <> EMP_發文歸檔 And _
                  m_strLastEEP04 <> EMP_退件重送 Then
                  CboEEP04.AddItem EMP_會完重修 & " " & "會完重修"
               End If
            '專利,其他案
            ElseIf bolPAFlow = True Or bolOtherFlow = True Then
               If m_strLastEEP04 <> EMP_判發 And _
                  m_strLastEEP04 <> EMP_退件重送 Then
                  CboEEP04.AddItem EMP_會完重修 & " " & "會完重修"
               End If
            End If
         End If
      End If
      '2020/4/8 END
   End If
   '2016/3/7 END
'*********************************************************************************************
'*********************************************************************************************
   If bolWaitReply = False Then
'      'Add By Sindy 2023/11/2
'      If m_strLastEEP04 = EMP_退回 Then '翻譯交稿退回
'         If InStr(strLastEEP11, "流程狀態:" & EMP_翻譯交稿) > 0 Or InStr(strLastEEP11, "流程狀態:" & EMP_送排版) > 0 Then
'            intReceiveKind = 4
'         End If
'      End If
'      '2023/11/2 END
      
      'Add By Sindy 2023/9/27
      If m_bolSendChWrite = True Then '送中說
         If UCase(TypeName(m_PrevForm)) = UCase("frm060107_1") Then
            '已有完稿日
            If Val(strEP09) > 0 Then
               If Val(m_EP39) = 0 Then '無稿核完成日
                  CboEEP04.AddItem EMP_翻譯交稿 & " " & "翻譯交稿"
               End If
            End If
         Else
            If Val(m_EP39) = 0 Then '無稿核完成日
               If PField(1) = "P" And cp(10) = "201" Then
                  CboEEP04.AddItem EMP_送核稿分案 & " " & "送核稿分案"
               ElseIf PField(1) = "FCP" And cp(10) <> "201" And InStr("209,235", cp(10)) > 0 Then
                  'Modify By Sindy 2025/8/14 亭妙說設計案都不用進打字室排版
                  If pa(8) <> "3" Then
                  '2025/8/14 END
                     CboEEP04.AddItem EMP_送排版 & " " & "送排版"
                  End If
               End If
            End If
         End If
      End If
      '2023/9/27 END
      
      Select Case intReceiveKind
         Case 3 '繪圖人員工作進度
            'If m_DPMan <> "" Then
               'Add By Sindy 2013/10/8
               If (m_EP34 = "N" Or bolFCPFlow = True) And Val(strEP14) > 0 Then '有草齊日
                  If m_DPMan = m_DMMan Then
                     'Add By Sindy 2014/3/13 當承辦人為專利處繪圖的人員時,則可直接判發
                     If PUB_GetStaffST15(Left(m_EPMan, 5), "1") = "P13" Then
                        CboEEP04.AddItem EMP_判發 & " " & "判發"
                     Else
                     '2014/3/13 END
                        CboEEP04.AddItem EMP_繪圖判發 & " " & "繪圖判發"
                     End If
                  Else
                     CboEEP04.AddItem EMP_墨完 & " " & "墨完"
                  End If
               Else
               '2013/10/8 END
                  '草圖齊備日
                  If Val(strEP14) > 0 Then
                     'Modify By Sindy 2015/4/22
                     '新申請案且有設定草圖核稿人時,增加草核;但若相同案已有草核過時除外
                     'Modify By Sindy 2015/6/9 改判斷若相同案已有草圖完稿日時除外
                     'Modify By Sindy 2022/11/11 ex:P-130435 一案兩請(發明) + Or PUB_ChkEmpFlowExists(lblCP09, EMP_草核, , strRefEEP02) = True
                     'Modify By Sindy 2025/5/23 李柏翰經理指示取消此控制: chkSameCaseFlow(EMP_草核) = False Or
                     '                          請改為所有的CFP設計案都要經過草核
                     'Modify By Sindy 2025/9/5 CFP設計的回代跟答辯，比照CFP設計的新申請案的管控方式
                     If PField(1) <> "FCP" And _
                        (InStr(NewCasePtyList, cp(10)) > 0 Or pa(8) = "3") And _
                        m_DPMan <> m_DCMan Then
                        'Modify By Sindy 2025/6/3
                        'If (cp(1) = "CFP" And cp(10) = 設計申請) Or chkSameCaseFlow(EMP_草核) = False Then
                        If (cp(1) = "CFP" And pa(8) = "3") Or chkSameCaseFlow(EMP_草核) = False Then
                        '2025/9/5 END
                           CboEEP04.AddItem EMP_草核 & " " & "草核"
                        Else
                           'Add By Sindy 2025/9/12 有草核也要有草核完才行 ex:P-136238(發明)/P-136239(新型)
                           If chkSameCaseFlow(EMP_草核完) = False Then
                              CboEEP04.AddItem EMP_草核 & " " & "草核"
                           Else
                           '2025/9/12 END
                              CboEEP04.AddItem EMP_草完 & " " & "草完"
                           End If
                        End If
                        '2025/6/3 END
                     Else
                     '2015/4/22 END
                        CboEEP04.AddItem EMP_草完 & " " & "草完"
                     End If
                  End If
                  '草圖完稿日
                  If Val(strEP15) > 0 Then
                     If PUB_ChkEmpFlowExists(lblCP09, EMP_草核, , strRefEEP02) = True Then
                        If PUB_ChkEmpFlowExists(lblCP09, EMP_草核完, strRefEEP02) = True Then
                           CboEEP04.AddItem EMP_標號 & " " & "標號"
                           
'                        'Modify By Sindy 2022/11/11 ex:P-130435 一案兩請(發明)
'                        ElseIf chkSameCaseFlow(EMP_草核) = True Then
'                           Dim bolNotAdd As Boolean
'                           bolNotAdd = False
'                           For ii = 0 To CboEEP04.ListCount - 1
'                              If InStr(CboEEP04.List(ii), "草核") > 0 Then
'                                 bolNotAdd = True
'                              End If
'                           Next
'                           If bolNotAdd = False Then
'                              CboEEP04.AddItem EMP_標號 & " " & "標號"
'                           End If
'                        '2022/11/11 END
                        End If
                     Else
                        CboEEP04.AddItem EMP_標號 & " " & "標號"
                     End If
                  End If
                  '墨圖齊備日
                  If Val(strEP17) > 0 Then
                     If m_DPMan = m_DMMan Then '2013/6/4 自行核判者不須墨完即可繪圖判發
                        'Add By Sindy 2014/3/13 當承辦人為專利處繪圖的人員時,則可直接判發
                        If PUB_GetStaffST15(Left(m_EPMan, 5), "1") = "P13" Then
                           CboEEP04.AddItem EMP_判發 & " " & "判發"
                        Else
                        '2014/3/13 END
                           CboEEP04.AddItem EMP_繪圖判發 & " " & "繪圖判發"
                        End If
                     Else
                        CboEEP04.AddItem EMP_墨完 & " " & "墨完"
                     End If
                  End If
               End If
            'End If
            
         Case 0 '承辦人工作進度
            '有承辦人
            If m_EPMan <> "" Then
               '有齊備日
               'Modified by Lydia 2019/11/22 前一畫面是否有查名齊備管制和查名齊備日
               'If Val(strEP06) > 0 Then
               If Val(m_EP06) > 0 And (bolCP143 = False Or (bolCP143 = True And Val(strCP143) > 0)) Then
                  'Add By Sindy 2016/3/15
                  'Modify By Sindy 2023/8/29 一開始會圖是有鎖m_DPMan <> "", 但現在已是會圖文了,不檢查繪圖人員 ex:P-131714(發明申請)
                  'If m_DPMan <> "" And Val(strEP09) = 0 Then '必須無完稿日才可會圖
                  If Val(strEP09) = 0 And bolPAFlow = True Then
                  '2023/8/29 END
                     CboEEP04.AddItem EMP_會圖 & " " & "會(圖/文)" 'Modify By Sindy 2022/10/7 會圖=>會(圖/文)
                  End If
                  '2016/3/15 END
                  
                  'Add By Sindy 2016/3/10
                  '有繪圖人員
                  If m_DPMan <> "" And Left(m_DPMan, 5) <> m_FlowUserNum And bolPAFlow = True Then
                     '有草圖完稿日
                     If Val(strEP15) > 0 Then
                        CboEEP04.AddItem EMP_修改圖式 & " " & "修改圖式"
                     End If
                  End If
                  '2016/3/10 END
                  
                  'Add By Sindy 2023/12/18 分割案,工程師有時會送排版
                  If bolFCPFlow = True And (cp(10) = "307" And pa(8) <> "3") Then
                     CboEEP04.AddItem EMP_送排版 & " " & "送排版"
                  End If
                  '2023/12/18 END
                  
                  '有英文核稿人
                  If m_EMMan <> "" And Left(m_EMMan, 5) <> m_FlowUserNum Then
                     'Modify By Sindy 2016/3/9 若有待回覆的送英核，不可再新增送英核歷程
                     strExc(0) = "select eep01 from EmpElectronProcess where eep01='" & m_EEP01 & "' and eep04='" & EMP_送英核 & "' and eep09='Y'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 0 Then
                        'Modify By Sindy 2015/3/16
                        If m_EP41 = "2" Then '2.日
                           CboEEP04.AddItem EMP_送英核 & " " & "送日核"
                        Else
                        '2015/3/16 END
                           CboEEP04.AddItem EMP_送英核 & " " & "送英核"
                        End If
                     End If
                     '2016/3/9 END
                  End If
                  
                  '有核稿人
                  'Modify By Sindy 2023/10/31 排除外專新案翻譯
                  'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
                  If m_CMMan <> "" And Left(m_CMMan, 5) <> m_FlowUserNum _
                     And Not (bolFCPFlow = True And (cp(10) = "201" Or cp(10) = "931")) Then
                     CboEEP04.AddItem EMP_送核 & " " & "送核"
                  End If
                  
                  'Add By Sindy 2016/12/27 + 准許先會
                  If bolDoEmpFlow32Limit = True Then
                     CboEEP04.AddItem EMP_准許先會 & " " & "准許先會"
                  End If
                  '2016/12/27 END
                  
                  '有繪圖人員
                  If m_DPMan <> "" And Left(m_DPMan, 5) <> m_FlowUserNum And bolPAFlow = True Then
                     'Modify By Sindy 2017/7/21 Mark
                     'If m_CMMan = "" Or Left(m_CMMan, 5) = m_FlowUserNum Or (m_CMMan <> "" And Val(m_EP39) > 0) Then
                     '2017/7/21 END
                        CboEEP04.AddItem EMP_送標號 & " " & "送標號"
                     'End If
                  End If
                  
                  '要會稿時,自行核判或有核稿完成日
                  If m_EP34 = "Y" And _
                     (m_CMMan = "" Or Left(m_CMMan, 5) = m_FlowUserNum Or (m_CMMan <> "" And Val(m_EP39) > 0)) Then
                     'Add By Sindy 2016/12/22
                     If bolHadSetProofEngReader = True Then '須英核的工程師(必須先送英核才可送會)
                        If bolHadFlow32CanSMeet = True Then
                           CboEEP04.AddItem EMP_送會 & " " & "送會"
                        End If
                     Else
                     '2016/12/22 END
                        CboEEP04.AddItem EMP_送會 & " " & "送會"
                     End If
                  End If
                  
                  '有繪圖人員
                  If m_DPMan <> "" And Left(m_DPMan, 5) <> m_FlowUserNum And bolPAFlow = True Then
                     '自行核判或有核稿完成日
                     If m_CMMan = "" Or Left(m_CMMan, 5) = m_FlowUserNum Or (m_CMMan <> "" And Val(m_EP39) > 0) Then
                        '不需會稿或(要會稿有智權人員會稿完成日)
                        'Modify By Sindy 2014/1/27
                        'If m_EP34 <> "Y" Or (m_EP34 = "Y" And Val(strEP38) > 0 And Val(strEP08) > 0) Then
                        If m_EP34 <> "Y" Or (m_EP34 = "Y" And Val(strEP08) > 0) Then
                        '2014/1/27 END
                           CboEEP04.AddItem EMP_上墨 & " " & "上墨"
                        End If
                     End If
                  End If
                  
                  '有會稿完成日才能做送判或判發
                  'If Val(strEP08) > 0 Then
                  '不需會稿或(要會稿有智權人員會稿完成日)
                  'If m_strLastEEP04 = EMP_會完 Then
                  'Modify By Sindy 2013/9/3 因舊案接續要做電子化增加檢查會稿完成日
                  'Modify By Sindy 2014/1/27
                  'If m_EP34 <> "Y" Or (m_EP34 = "Y" And Val(strEP38) > 0 And Val(strEP08) > 0) Or Val(strEP08) > 0 Then
                  'Modify By Sindy 2018/4/27
                  'If m_EP34 <> "Y" Or (m_EP34 = "Y" And Val(strEP08) > 0) Or Val(strEP08) > 0 Then
                  'Modify By Sindy 2024/1/5 + 排除一核
                  If (m_EP34 <> "Y" Or Val(strEP08) > 0) _
                     And Not (Lbl926.Visible = True And InStr(Lbl926.Caption, "一核") > 0) Then
                  '2018/4/27 END
                  '2014/1/27 END
                     'Modify By Sindy 2014/3/13 Move此段程式至此處
                     'Add By Sindy 2014/1/10 有送英核但無英文核完日時,不可送判或判發
                     'Add By Sindy 2015/3/16 送日核不須控管日文核完日
                     'If m_EP41 = "1" Then '1.英
                     '2015/3/16 END
                     
'                     'Add By Sindy 2023/10/31
'                     If bolFCPFlow = True And PUB_ChkEmpFlowExists(m_EEP01, EMP_送英核) = False And m_EMMan <> "" Then
'                        GoTo RunEnd
'                     End If
'                     '2023/10/31 END
                     
                     'Modify By Sindy 2016/3/9 稍做調整彈訊息,直接不出現送判狀態
                     If m_EMMan <> "" Then 'Add By Sindy 2015/5/14 +if 增加檢查是否還有英文核稿人,若有,才要檢查是否未核完 ex.CFP-025751
                        If (bol00EngCMFlow = True Or PUB_ChkEmpFlowExists(m_EEP01, EMP_送英核) = True Or bolFCPFlow = True) _
                           And Val(m_EP33) = 0 Then
                           If m_EP41 = "2" Then
                              MsgBox "已送日核未核完，不可送判或" & IIf(bolFCPFlow = True, "送件", "判發") & "!!"
                           Else
                              MsgBox "已送英核未核完，不可送判或" & IIf(bolFCPFlow = True, "送件", "判發") & "!!"
                           End If
                           GoTo RunEnd
                        End If
                     End If
                     '自行核稿或有核稿完成日(核稿已完成)
                     If m_CMMan = "" Or _
                        Left(m_CMMan, 5) = m_FlowUserNum Or _
                        (m_CMMan <> "" And Val(m_EP39) > 0) Then
                        
                        '有判發人
                        If m_CSMan <> "" And Left(m_CSMan, 5) <> m_FlowUserNum Then
                           CboEEP04.AddItem EMP_送判 & " " & "送判"
                        End If
                        
                        'Modify By Sindy 2023/10/26 從下面程式移上來
                        If m_strLastEEP04 = EMP_退件 Then
                           CboEEP04.AddItem EMP_退件重送 & " " & "退件重送"
                           GoTo RunEnd
                        End If
                        
                        'Add By Sindy 2018/4/27
                        '自行判發或有判發完成日(判發已完成)
                        If m_CSMan = "" Or _
                           Left(m_CSMan, 5) = m_FlowUserNum Or _
                           (m_CSMan <> "" And Val(m_EP42) > 0) Then
                           If bolTMFlow = True Then
                              '商標處C類由承辦人自行上發文日,歸檔
                              If Left(lblCP09, 1) = "C" Then
                                 CboEEP04.AddItem EMP_發文歸檔 & " " & "發文歸檔"
                              Else
                                 CboEEP04.AddItem EMP_送件 & " " & "送件"
                              End If
                           'Add By Sindy 2024/7/12
                           ElseIf bolFCTFlow = True Or bolCFTFlow = True Then
                              CboEEP04.AddItem EMP_送件 & " " & "送件"
                           'Add By Sindy 2023/10/4
                           ElseIf bolFCPFlow = True Then
                              'Add By Sindy 2024/1/10
                              If PField(1) <> "P" Then
                              '2024/1/10 END
                                 '+分割 FCP-070653
                                 'Modify By Sindy 2024/1/5 因為上線前的中說 + Or InStr(屬中說的案件性質, cp(10)) > 0
                                 If (PUB_ChkEmpFlowExists(m_EEP01, EMP_排版完成) = True _
                                     And PUB_ChkEmpFlowExists(m_EEP01, EMP_核稿分案) = True _
                                    ) _
                                    Or (cp(10) = "307" And pa(8) <> "3") _
                                    Or InStr(屬中說的案件性質, cp(10)) > 0 Then
                                    CboEEP04.AddItem EMP_送轉檔 & " " & "送轉檔"
                                    If PUB_ChkEmpFlowExists(m_EEP01, EMP_送轉檔) = False Then
                                       GoTo RunEnd
                                    End If
                                 End If
                              End If
                              'Add By Sindy 2025/4/8
                              If cp(10) = "945" Then
                                 CboEEP04.AddItem EMP_發文歸檔 & " " & "發文歸檔"
                              Else
                              '2025/4/8 END
                                 CboEEP04.AddItem EMP_送件 & " " & "送件"
                              End If
                           '2023/10/4 END
                           Else
                              CboEEP04.AddItem EMP_判發 & " " & "判發"
                           End If
                        End If
                        '2018/4/27 END
                     End If
                  End If
    
'                  'Modify By Sindy 2013/10/24
'                  If m_strLastEEP04 = EMP_退件 Then
'                     CboEEP04.AddItem EMP_退件重送 & " " & "退件重送"
'                  End If
                  
               End If '有齊備日
            End If '有承辦人
         Case Else
      End Select
'*********************************************************************************************
'*********************************************************************************************
   Else
      Select Case intReceiveKind
         Case 1 '待核判區
            If strLastEEP05 = m_FlowUserNum Then 'Add By Sindy 2013/9/12 +if
               If m_strLastEEP04 = EMP_轉回 And _
                  (InStr(strLastEEP11, "轉回前流程狀態:" & EMP_送英核) > 0 Or _
                   InStr(strLastEEP11, "轉回前流程狀態:" & EMP_送核) > 0 Or _
                   InStr(strLastEEP11, "轉回前流程狀態:" & EMP_草核) > 0 Or _
                   InStr(strLastEEP11, "轉回前流程狀態:" & EMP_墨完) > 0 Or _
                   InStr(strLastEEP11, "轉回前流程狀態:" & EMP_送判) > 0) Then
                  If InStr(strLastEEP11, "原收受者") > 0 Then
                     CboEEP04.AddItem EMP_轉回 & " " & "轉回"
                  End If
               End If
               
               'Add By Sindy 2023/9/28
               If m_strLastEEP04 = EMP_翻譯交稿 Then
                  'CboEEP04.AddItem EMP_退回 & " " & "退回"
                  If PField(1) = "P" And cp(10) = "201" Then
                     CboEEP04.AddItem EMP_送核稿分案 & " " & "送核稿分案"
                  ElseIf PField(1) = "FCP" Then
                     CboEEP04.AddItem EMP_送排版 & " " & "送排版"
                  End If
                  
               ElseIf m_strLastEEP04 = EMP_送排版 Then
                  'CboEEP04.AddItem EMP_退回 & " " & "退回"
                  CboEEP04.AddItem EMP_交辦 & " " & "交辦"
                  CboEEP04.AddItem EMP_排版完成 & " " & "排版完成"
                  
               ElseIf m_strLastEEP04 = EMP_排版完成 Or m_strLastEEP04 = EMP_送核稿分案 Then
                  CboEEP04.AddItem EMP_核稿分案 & " " & "核稿分案"
                  
               ElseIf m_strLastEEP04 = EMP_送轉檔 Then
                  'CboEEP04.AddItem EMP_退回 & " " & "退回"
                  CboEEP04.AddItem EMP_交辦 & " " & "交辦"
                  CboEEP04.AddItem EMP_轉檔完成 & " " & "轉檔完成"
                  
               ElseIf m_strLastEEP04 = EMP_程序送判 Then
                  CboEEP04.AddItem EMP_程序退回 & " " & "程序退回"
                  CboEEP04.AddItem EMP_送件 & " " & "送件"
               '2023/9/28 END
                  
               ElseIf m_strLastEEP04 = EMP_送英核 Or m_strLastEEP04 = EMP_送核 Then
                  'Add By Sindy 2016/12/21 + 准許先會
                  If bolDoEmpFlow32Limit = True And m_strLastEEP04 = EMP_送英核 Then
                     CboEEP04.AddItem EMP_准許先會 & " " & "准許先會"
                  End If
                  '2016/12/21 END
                  'Add By Sindy 2023/10/2
                  If bolFCPFlow = True And m_strLastEEP04 = EMP_送英核 And InStr("F62,F71,F72", Pub_StrUserSt03) > 0 Then
                     CboEEP04.AddItem EMP_交辦 & " " & "交辦"
                  End If
                  '2023/10/2 END
                  CboEEP04.AddItem EMP_核修 & " " & "核修"
                  CboEEP04.AddItem EMP_核完 & " " & "核完"
                  'Modify By Sindy 2013/10/29 英文及日文核稿人不可直接判發
                  If m_strLastEEP04 = EMP_送核 Then
                     '專利處
                     If bolPAFlow = True Then
                        'Add By Sindy 2013/9/17 不會稿的並且核稿人可直接判發
                        If m_EP34 = "N" And m_CPM27 <> "N" Then '*m_CPM27=N:不可直接判發
                           CboEEP04.AddItem EMP_判發 & " " & "判發"
                        'Modify By Sindy 2013/9/23 一案二請,P新型可直接判發
                        'Modify By Sindy 2025/2/5 增加檢查"不需會稿"者才可直接判發 ex:P-134711
                        ElseIf (PField(1) = "P" And cp(10) = "102") And m_EP34 = "N" Then
                           If PUB_DualCaseRelationExist(pa) = True Then
                              CboEEP04.AddItem EMP_判發 & " " & "判發"
                           End If
                        End If
                        '2013/9/17 END
                     'Add By Sindy 2023/10/4
                     '外專
                     ElseIf bolFCPFlow = True Then
                        'And m_CPM27 <> "N" ex:P-XXXXXX(101)
                        'Modify By Sindy 2024/1/5 + 排除一核
                        If m_EP34 = "N" _
                           And Not (Lbl926.Visible = True And InStr(Lbl926.Caption, "一核") > 0) Then
                           If (m_CMMan = m_CSMan Or m_CSMan = "") And _
                              (m_EMMan = "" Or (m_EMMan <> "" And Val(m_EP33) > 0)) Then
                              If Not ((PUB_ChkEmpFlowExists(m_EEP01, EMP_排版完成, , strRefEEP02) = True Or InStr(屬中說的案件性質, cp(10)) > 0) _
                                      And PUB_ChkEmpFlowExists(m_EEP01, EMP_送轉檔, strRefEEP02) = False) Then
                                 '945電話聯絡單由工程師主管上發文日,歸檔
                                 If cp(10) = "945" Then
                                    CboEEP04.AddItem EMP_發文歸檔 & " " & "發文歸檔"
                                 Else
                                    CboEEP04.AddItem EMP_送件 & " " & "送件"
                                 End If
                              End If
                           End If
                        End If
                        '2023/10/4 END
                     '商標處,其他
                     Else
                        'Add By Sindy 2018/8/10 不會稿的並且核稿人=判發人或無判發人時,可直接判發
                        If m_EP34 = "N" And m_CPM27 <> "N" Then
                           If m_CMMan = m_CSMan Or m_CSMan = "" Then
                              CboEEP04.AddItem EMP_判發 & " " & "判發"
                              'Add By Sindy 2024/8/7
                              If bolCFTFlow = True Or bolFCTFlow = True Then
                                 CboEEP04.AddItem EMP_送件 & " " & "送件"
                              End If
                              '2024/8/7 END
                           End If
                        End If
                     End If
                  End If
                  
               'Add By Sindy 2015/4/22
               '有草圖完稿日
               ElseIf (m_strLastEEP04 = EMP_草核 And Val(strEP15) > 0) Then
                  CboEEP04.AddItem EMP_草修 & " " & "草修"
                  CboEEP04.AddItem EMP_草核完 & " " & "草核完"
               '2015/4/22 END
                  
               '有墨圖完稿日
               'Modify By Sindy 2016/10/24 因P-115605墨完要退回,但智權人員已做會完重修會清掉墨圖齊備及完稿日
               'ElseIf (m_strLastEEP04 = EMP_墨完 And Val(m_EP18) > 0) Then
               ElseIf m_strLastEEP04 = EMP_墨完 Then
               '2016/10/24 END
                  CboEEP04.AddItem EMP_退回 & " " & "退回"
                  'Add By Sindy 2014/3/13 當承辦人為專利處繪圖的人員時,則可直接判發
                  If PUB_GetStaffST15(Left(m_EPMan, 5), "1") = "P13" Then
                     CboEEP04.AddItem EMP_判發 & " " & "判發"
                  Else
                  '2014/3/13 END
                     CboEEP04.AddItem EMP_繪圖判發 & " " & "繪圖判發"
                  End If
                  
               ElseIf m_strLastEEP04 = EMP_送判 Then
'                  'Add By Sindy 2013/9/24
'                  If Left(m_EEP01, 1) = "B" And CP(10) = 延期 Then
'                     'B類延期不可執行退回
'                  Else
'                  '2013/9/24 END
                     CboEEP04.AddItem EMP_退回 & " " & "退回"
'                  End If
                  CboEEP04.AddItem EMP_判發 & " " & "判發"
                  'Add By Sindy 2024/8/7
                  If bolCFTFlow = True Or bolFCTFlow = True Then
                     CboEEP04.AddItem EMP_送件 & " " & "送件"
                  '2024/8/7 END
                  'Add By Sindy 2023/10/4
                  ElseIf bolFCPFlow = True Then
                     If Not ((PUB_ChkEmpFlowExists(m_EEP01, EMP_排版完成, , strRefEEP02) = True Or InStr(屬中說的案件性質, cp(10)) > 0) _
                             And PUB_ChkEmpFlowExists(m_EEP01, EMP_送轉檔, strRefEEP02) = False) Then
                        '945電話聯絡單由工程師主管上發文日,歸檔
                        If cp(10) = "945" Then
                           CboEEP04.AddItem EMP_發文歸檔 & " " & "發文歸檔"
                        Else
                           CboEEP04.AddItem EMP_送件 & " " & "送件"
                        End If
                     End If
                  End If
                  '2023/10/4 END
               End If
            End If
            
         Case 2 '待會稿區
            'Modify By Sindy 2014/2/26 因智權人員原是96027.林佳芳改為94026.林建志; 因此調整程式,不管送會收受者
            If strLastEEP05 = m_FlowUserNum Then 'Add By Sindy 2013/9/12 +if
               '有會稿日
               If m_strLastEEP04 = EMP_送會 And Val(m_EP07) > 0 Then
                  'Modify By Sindy 2019/5/7 C類不顯示客戶會稿, 其他保留(下列控制不強制客戶會稿)
                  If Left(cp(9), 1) = "A" Or Left(cp(9), 1) = "B" Then
                     CboEEP04.AddItem EMP_客戶會稿 & " " & "客戶會稿" 'Add By Sindy 2018/8/28
                  'Modify By Sindy 2025/8/1
                  'T台灣案提供C類來函（1201審查報告承辦人是商申組(T31)除外）會稿時，下一歷程狀態增加顯示[收文分析]
                  ElseIf cp(1) = "T" And m_Country = "000" And Left(cp(9), 1) = "C" _
                        And Not (cp(10) = "1201" And PUB_GetST93(cp(14)) = "T31") Then
                     CboEEP04.AddItem EMP_收文分析 & " " & "收文分析"
                  '2025/8/1 End
                  End If
                  '2019/5/7 END
                  CboEEP04.AddItem EMP_會修 & " " & "會修"
                  'Add By Sindy 2018/8/28 一定要有客戶會稿日才能做會完
                  '20180917上線前的歷程不管制
                  'Modify By Sindy 2018/9/18 C類也不管制
                  'Modify By Sindy 2018/10/1 B類也不管制
                  'Modify By Sindy 2018/10/1 A類分析也不管制
                  'If m_EP07 < "20180917" Or Left(cp(9), 1) = "C" Or Left(cp(9), 1) = "B" Or
                  'Modify By Sindy 2018/12/14 文雄提開放只要案件性質有分析二字都不鎖
                  '  Trim(lblCP10) = "分析" ==> InStr(lblCP10, "分析") > 0
                  'Modify By Sindy 2019/5/7 續展&延展開放智權人員直接會完
                  'Modify By Sindy 2024/7/11 + bolCFTFlow = True)
                  'Modify By Sindy 2025/8/6 (Left(cp(9), 1) = "A" And InStr(LblCP10, "分析") > 0 And bolPAFlow = True) Or
                  '                         改為 (Left(cp(9), 1) = "A" And InStr(LblCP10, "分析") > 0) Or
                  If m_EP07 < "20180917" Or _
                     Left(cp(9), 1) >= "B" Or _
                     (Left(cp(9), 1) = "A" And InStr(lblCP10, "分析") > 0) Or _
                     ((bolTMFlow = True Or bolCFTFlow = True) And cp(10) = "102") Then
                     CboEEP04.AddItem EMP_會完 & " " & "會完"
                  ElseIf Val(strEP37) > 0 Then
                  '2018/8/28 END
                     CboEEP04.AddItem EMP_會完 & " " & "會完"
                  End If
               'End If
               'Add By Sindy 2016/3/15
               ElseIf m_strLastEEP04 = EMP_會圖 Then
                  CboEEP04.AddItem EMP_客戶會稿 & " " & "客戶會稿" 'Add By Sindy 2018/8/28
                  CboEEP04.AddItem EMP_圖修 & " " & "(圖/文)修" 'Modify By Sindy 2022/10/7 圖修=>(圖/文)修
                  CboEEP04.AddItem EMP_圖完 & " " & "(圖/文)完" 'Modify By Sindy 2022/10/7 圖完=>(圖/文)完
               '2016/3/15 END
               End If
            End If
            
         'Add By Sindy 2013/10/16 開放草完及標號隨時都可以做
         Case 3 '繪圖人員工作進度
            If m_DPMan <> "" Then
               'Modify By Sindy 2015/4/22
'               '草圖齊備日
'               If Val(strEP14) > 0 Then
'                  CboEEP04.AddItem EMP_草完 & " " & "草完"
'               End If
'               '草圖完稿日
'               If Val(strEP15) > 0 Then
'                  CboEEP04.AddItem EMP_標號 & " " & "標號"
'               End If
               If PUB_ChkEmpFlowExists(lblCP09, EMP_草核, , strRefEEP02) = True Then
                  If PUB_ChkEmpFlowExists(lblCP09, EMP_草核完, strRefEEP02) = True Then
                     '草圖完稿日
                     If Val(strEP15) > 0 Then
                        CboEEP04.AddItem EMP_標號 & " " & "標號"
                     End If
                  End If
               Else
                  '草圖齊備日
                  If Val(strEP14) > 0 Then
                     '不需做草核的案件，此處才開放草完
                     'Modify By Sindy 2025/5/23 李柏翰經理指示取消此控制: And chkSameCaseFlow(EMP_草核) = False
                     '                          請改為所有的CFP設計案都要經過草核
                     'Modify By Sindy 2025/6/3
'                     If Not (PField(1) <> "FCP" And InStr(NewCasePtyList, cp(10)) > 0 And m_DPMan <> m_DCMan) Then
'                        CboEEP04.AddItem EMP_草完 & " " & "草完"
'                     End If
                     'Modify By Sindy 2025/9/5 CFP設計的回代跟答辯，比照CFP設計的新申請案的管控方式
                     If PField(1) <> "FCP" And _
                        (InStr(NewCasePtyList, cp(10)) > 0 Or pa(8) = "3") And _
                        m_DPMan <> m_DCMan Then
                        'If (cp(1) = "CFP" And cp(10) = 設計申請) Or chkSameCaseFlow(EMP_草核) = False Then
                        If (cp(1) = "CFP" And pa(8) = "3") Or chkSameCaseFlow(EMP_草核) = False Then
                        '2025/9/5 END
'                           CboEEP04.AddItem EMP_草核 & " " & "草核"
                        Else
                           'Add By Sindy 2025/9/12 有草核也要有草核完才行 ex:P-136238(發明)/P-136239(新型)
                           If chkSameCaseFlow(EMP_草核完) = True Then '*****
                           '2025/9/12 END
                              CboEEP04.AddItem EMP_草完 & " " & "草完"
                           End If
                        End If
                     Else
                        CboEEP04.AddItem EMP_草完 & " " & "草完"
                     End If
                     '2025/6/3 END
                  End If
                  '草圖完稿日
                  If Val(strEP15) > 0 Then
                     CboEEP04.AddItem EMP_標號 & " " & "標號"
                  End If
               End If
               '2015/4/22 END
            End If
            '2013/10/16 END
      End Select
   End If
   
RunEnd:
   'Modify By Sindy 2013/10/29 待會稿區不預設在最後一道流程,防智權人員應會修且直接會完
   If intReceiveKind <> 2 Then
   '2013/10/29 END
      '流程狀態預設為最後一道流程
'      'Add By Sindy 2014/1/10
'      If Left(CboEEP04.List(CboEEP04.ListCount - 1), 2) = EMP_送會 And Val(m_EP08) > 0 Then
'         '已有會完日不預設為送會
'      Else
'      '2014/1/10 END
         'CboEEP04.ListIndex = CboEEP04.ListCount - 1
         'Add By Sindy 2024/11/21 外專待核判區不要預設在送件,而是預設在前一個歷程判發
         If bolFCPFlow = True And intReceiveKind = 1 Then '1.待核判區
            'Modify By Sindy 2024/11/22 增加排除外專程序
            '                               排除外專電機-鄭光益(114/11/14)
            'If Left(CboEEP04.Text, 2) = EMP_送件 And Pub_StrUserSt03 <> "F22" Then
            If Pub_StrUserSt93 <> "F11" And Pub_StrUserSt93 <> "F31" Then
               If Left(CboEEP04.List(CboEEP04.ListCount - 2), 2) = EMP_判發 Then
                  CboEEP04.ListIndex = CboEEP04.ListCount - 2
               Else
                  CboEEP04.ListIndex = CboEEP04.ListCount - 1
               End If
            Else
               CboEEP04.ListIndex = CboEEP04.ListCount - 1
            End If
         Else
            CboEEP04.ListIndex = CboEEP04.ListCount - 1
         End If
         '2024/11/21 END
'      End If
   End If
   
   Set rsA = Nothing
   
   'Add By Sindy 2023/10/31
   If CboEEP04.Enabled = False Then
      Call cmdExit_Click
   End If
   '2023/10/31 END
End Sub

'Add By Sindy 2015/4/22 檢查相同案是否有該流程狀態
Private Function chkSameCaseFlow(strFlow1 As String) As Boolean
Dim rsQuery As ADODB.Recordset
Dim stVTB As String
   
   chkSameCaseFlow = False
   stVTB = PUB_GetSameCaseSQL(m_EEP01) '相同案語法(收文號)
   'Modify By Sindy 2015/6/9
   If strFlow1 = EMP_草核 Then
      strSql = "select * from engineerprogress," & _
               "(select cp09 from caseprogress," & _
               "(" & stVTB & ") V1" & _
               " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
               " and substr(V1.cno,-9,6)=cp02" & _
               " and substr(V1.cno,-3,1)=cp03" & _
               " and substr(V1.cno,-2)=cp04" & _
               " and cp10 in(" & NewCasePtyList & ")) V2" & _
               " Where V2.CP09 = ep02" & _
               " and nvl(ep15,0)>0 and ep02<>'" & m_EEP01 & "'"
   Else
   '2015/6/9 END
      strSql = "select * from empelectronprocess," & _
               "(select cp09 from caseprogress," & _
               "(" & stVTB & ") V1" & _
               " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
               " and substr(V1.cno,-9,6)=cp02" & _
               " and substr(V1.cno,-3,1)=cp03" & _
               " and substr(V1.cno,-2)=cp04" & _
               " and cp10 in(" & NewCasePtyList & ")) V2" & _
               " Where V2.CP09 = eep01" & _
               " and eep04 = '" & strFlow1 & "' and eep01<>'" & m_EEP01 & "'"
   End If
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      chkSameCaseFlow = True
   End If
   rsQuery.Close
   Set rsQuery = Nothing
End Function

Private Sub GRD1_DblClick()
GRD1.Visible = False
If GRD1.MouseRow <> 0 And GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
   'Modify By Sindy 2018/11/14 + And cmdExit.Enabled = True : 防止送出時又會按到查詢
   '  ◎FRM090201_2 新增歷程主檔有問題:Trim(txtEEP03)[81018] <> strUserNum[A5017]!! [寄件者：南所#25]
   '  新增歷程主檔有問題!
   '  insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep09,eep10,eep11,eep12,eep14,eep15) values('AA7031096',11,'81018','07','A5017',20181114,114603,'申復理由書之比較圖箭頭修改,修後可會.',NULL,NULL,'流程狀態:07',NULL,NULL,NULL)
   If cmdCancel.Visible = False And cmdExit.Enabled = True Then
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
      Call ReadData
   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

Private Sub SetCboCP10()
   If Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
      Label11.Caption = "會稿方式："
      CboCP10.Clear
      CboCP10.AddItem "1 EMail"
      CboCP10.AddItem "2 紙本"
      CboCP10.AddItem "3 微信"
      CboCP10.AddItem "4 LINE"
      CboCP10.AddItem "9 其他"
   End If
End Sub

'顯示明細資料於畫面上
'Modify By Sindy 2018/9/6 + Optional bolShowMailForm As Boolean = True
Private Sub ReadData(Optional bolShowMailForm As Boolean = True)
   Call ClearData
   Call SetCtrlReadOnly(False)
   
   m_EditMode = 4
   m_EEP02 = GRD1.TextMatrix(dblPrevRow, 0)
   txtEEP03 = GRD1.TextMatrix(dblPrevRow, 1)
   txtEEP03_2 = GRD1.TextMatrix(dblPrevRow, 2)
   CboEEP04.Text = GRD1.TextMatrix(dblPrevRow, 3) & " " & GRD1.TextMatrix(dblPrevRow, 12) '4
   ChkEED13.Enabled = False 'Add By Sindy 2023/11/17
   ChkEED08.Enabled = False 'Add By Sindy 2025/4/7
   CboEEP05.Text = GRD1.TextMatrix(dblPrevRow, 5) & " " & GRD1.TextMatrix(dblPrevRow, 6)
   txtEEP10_2 = GRD1.TextMatrix(dblPrevRow, 8)
   txtEEP08 = GRD1.TextMatrix(dblPrevRow, 9)
   txtEEP10 = GRD1.TextMatrix(dblPrevRow, 10)
   Call ReadAttachFile(m_EEP01, CInt(m_EEP02))
   
   cmdDel.Enabled = True 'Add By Sindy 2013/10/3
   cmdAddAttDB(0).Enabled = True 'Add By Sindy 2013/10/24
   cmdRemAttDB(0).Enabled = True 'Add By Sindy 2013/10/24
   cmdAddAttDB(1).Enabled = True 'Add By Sindy 2013/10/24
   cmdRemAttDB(1).Enabled = True 'Add By Sindy 2013/10/24
   cmdMail.Visible = False
   'Add By Sindy 2018/8/29 會稿方式
   If Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
      Frame4.Visible = True: CboCP10.Locked = True
      Call SetCboCP10
      If GRD1.TextMatrix(dblPrevRow, 15) <> "" Then
         For ii = 0 To CboCP10.ListCount - 1
            If Left(CboCP10.List(ii), 1) = GRD1.TextMatrix(dblPrevRow, 15) Then
               CboCP10.ListIndex = ii
               Exit For
            End If
         Next ii
      End If
      If Left(CboCP10.Text, 1) = "1" And bolShowMailForm = True Then
         'cmdMail.Visible = True 'Add By Sindy 2018/8/31
         GRD1.Visible = True
         Call cmdMail_Click
      End If
   Else
      Frame4.Visible = False: CboCP10.Locked = False
   End If
   '2018/8/29 END
End Sub

'查詢附件檔
Private Sub ReadAttachFile(strEEP01 As String, intEEP02 As Integer)
   KillAttach
   lstAtt(0).Clear 'Add By Sindy 2013/10/24
   'Modify By Sindy 2018/9/4 + and eef12 is not null
   strExc(0) = "select eef03,eef04,eef09,eef10 from EmpElectronFile where eef01='" & strEEP01 & "' and eef02=" & intEEP02 & " and eef12 is not null order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Do While Not .EOF
            'Modify By Sindy 2018/11/27 增加讀取日期時間
            'lstAtt(0).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)", 0
            lstAtt(0).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB) #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "#", 0
            lstAtt(0).ITEMDATA(0) = 1
            .MoveNext
         Loop
      End With
      Me.cmdOpenAtt(0).Enabled = True
      Me.cmdSelect(0).Enabled = True
      Me.cmdSaveAtt(0).Enabled = True
   End If
   If lstAtt(0).ListCount > 0 Then SetListScroll lstAtt(0)
End Sub

'查詢存卷區
Private Sub ReadAttachFile_other(strEEP01 As String)
   KillAttach
   lstAtt(1).Clear
   'Modify By Sindy 2014/12/17 增加顯示檔案的新增人員和日期, 依CreateDate+檔名做排序
   'strExc(0) = "select eef03,eef04 from EmpElectronFile where eef01='" & strEEP01 & "' and eef02=0 order by 1"
   'Modify By Sindy 2018/9/4 + and eef12 is not null
   strExc(0) = "select eef03,eef04,st02,sqldatet(eef07),eef09,eef10 from EmpElectronFile,staff where eef01='" & strEEP01 & "' and eef02=0 and eef12 is not null and eef06=st01(+) order by eef07,eef03 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Do While Not .EOF
            'Modify By Sindy 2014/12/17 增加顯示檔案的新增人員和日期
            'Modify By Sindy 2018/11/27 增加讀取日期時間
            'lstAtt(1).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB) " & .Fields("st02") & "-->" & .Fields(3), 0
            lstAtt(1).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)" & " #" & CStr(Format(Val(.Fields("eef09")), "00000000")) & CStr(Format(Val(.Fields("eef10")), "000000")) & "# " & .Fields("st02") & "-->" & .Fields(3), 0
            lstAtt(1).ITEMDATA(0) = 1
            .MoveNext
         Loop
      End With
      'Add By Sindy 2018/8/8
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
      If bolTMFlow = True Or bolOtherFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
         Text2.Left = 1600 '存卷資料標籤
      End If
      '2018/8/8 END
      Text2.Visible = True 'Add By Sindy 2014/10/1 有存卷附件時顯示字樣
   Else
      Text2.Visible = False 'Add By Sindy 2014/10/1
   End If
   Me.cmdSave.Visible = False
   Me.cmdAddAtt(1).Visible = False
   Me.cmdRemAtt(1).Visible = False
   '未發文未取消收文時,才可做存卷
   If Val(cp(27)) = 0 And Val(cp(57)) = 0 Then 'And m_FlowUserNum = Left(m_EPMan, 5)
      '檢查卷宗區是否已有資料，若有則表示已做歸檔動作，不可再此處異動資料
      'Modify By Sindy 2015/1/19
      'strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & m_EEP01 & "'"
      strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & m_EEP01 & "' and CPP12='S'"
      '2015/1/19 END
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         Me.cmdAddAtt(1).Visible = True
         Me.cmdRemAtt(1).Visible = True
      End If
   End If
   If lstAtt(1).ListCount > 0 Then SetListScroll lstAtt(1)
End Sub

'沿用附件檔
'Modify By Sindy 2018/10/22 + Optional strCaseNo As String
Private Function DownloadAttFile_copy(strEEP01 As String, intEEP02 As Integer, _
   Optional strCaseNo As String) As Boolean
Dim stAttPathFile As String
Dim lngSize As Long
Dim iFileNo As Integer
Dim bytes() As Byte
Dim fs, f
Dim strSMB08 As String 'Add By Sindy 2018/10/15
Dim strField(1 To 4)
Dim strSaveCaseNo1 As String, strSaveCaseNo2 As String, strSaveCaseNo3 As String, strSaveCaseNo4 As String
   
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   DownloadAttFile_copy = True
   
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   'Modify By Sindy 2018/9/27 Mark
'   Else
'      KillAttach
   End If
   '繪圖的原始檔不沿用
   'Modify By Sindy 2014/3/11 +dwg.7z
   'Modified by Morgan 2015/5/27 改用FTP
   'strExc(0) = "select eef03,eef04,eef05 from EmpElectronFile where eef01='" & strEEP01 & "' and eef02=" & intEEP02 & _
               " and substr(upper(eef03),-4)<>'.DWG' and substr(upper(eef03),-7)<>'DWG.ZIP' and substr(upper(eef03),-6)<>'DWG.7Z'" & _
               " order by 1"
   'Modify By Sindy 2018/9/4 + and eef12 is not null
   'Modify By Sindy 2018/10/15 增加檢查寄件備份的附件順序
   strExc(0) = "select smb08 from smailbackup" & _
               " where smb01='" & strEEP01 & "' and smb11=" & intEEP02
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strSMB08 = RsTemp.Fields("smb08")
   End If
   If strSMB08 <> "" Then
      strExc(0) = "select eef03,eef04,eef12,eep04 from EmpElectronFile,empelectronprocess" & _
                  " where eef01='" & strEEP01 & "' and eef02=" & intEEP02 & " and eef12 is not null" & _
                  " and substr(upper(eef03),-4)<>'.DWG' and substr(upper(eef03),-7)<>'DWG.ZIP' and substr(upper(eef03),-6)<>'DWG.7Z'" & _
                  " and eef01=eep01(+) and eef02=eep02(+)" & _
                  " order by decode(instr('" & strSMB08 & "',eef03),0,99,instr('" & strSMB08 & "',eef03)) desc"
   Else
   '2018/10/15 END
      strExc(0) = "select eef03,eef04,eef12,eep04 from EmpElectronFile,empelectronprocess" & _
                  " where eef01='" & strEEP01 & "' and eef02=" & intEEP02 & " and eef12 is not null" & _
                  " and substr(upper(eef03),-4)<>'.DWG' and substr(upper(eef03),-7)<>'DWG.ZIP' and substr(upper(eef03),-6)<>'DWG.7Z'" & _
                  " and eef01=eep01(+) and eef02=eep02(+)" & _
                  " order by 1"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Do While Not .EOF
'            lstAtt(0).AddItem .Fields("eef03") & " (" & Round(.Fields("eef04") / 1024, 2) & " KB)", 0
'            lstAtt(0).ItemData(0) = 1
            stAttPathFile = .Fields("eef03")
            
            '開始下載檔案
            'Add By Sindy 2018/9/21 送件時系統自動去掉中文
            'Modify By Sindy 2024/1/11 +And bolFCPFlow = False 排除外專
            'Modify By Sindy 2024/8/14 +And bolFCTFlow = False 排除FC外商
            If (Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送) And bolFCPFlow = False And bolFCTFlow = False Then
               'stAttPathFile = m_AttachPath & "\" & PUB_GetSimpleName(.Fields("eef03"))
               stAttPathFile = PUB_GetSimpleName(stAttPathFile)
               stAttPathFile = Replace(stAttPathFile, "..", ".")
               stAttPathFile = Replace(stAttPathFile, "-.", ".")
               stAttPathFile = Replace(stAttPathFile, "_.", ".")
            End If
            '2018/9/21 END
            
            'Add By Sindy 2018/9/20 去掉檔名前的客戶案號 ex:PAT-1685.P109085.bxa-180919.doc
            If m_PA48 <> "" Then
               If Left(CboEEP04.Text, 2) <> EMP_客戶會稿 Then
                  If InStr(stAttPathFile, m_PA48) > 0 Then
                     stAttPathFile = Replace(stAttPathFile, m_PA48 & ".", "")
                  'Modify By Sindy 2024/2/2 + PUB_FilterEFileSymbol 轉全型
                  ElseIf InStr(stAttPathFile, PUB_FilterEFileSymbol(m_PA48)) > 0 Then
                     stAttPathFile = Replace(stAttPathFile, PUB_FilterEFileSymbol(m_PA48) & ".", "")
                     '2024/2/2 END
                  End If
               End If
            End If
            '2018/9/20 END
            
            'Add By Sindy 2018/10/15 會修,會完
            If RsTemp.Fields("eep04") = EMP_客戶會稿 And Left(CboEEP04.Text, 2) <> EMP_客戶會稿 Then
               If strCaseNo <> "" Then
                  '系統自動補填案號使用
                  strField(1) = SystemNumber(strCaseNo, 1)
                  strField(2) = SystemNumber(strCaseNo, 2)
                  strField(3) = SystemNumber(strCaseNo, 3)
                  strField(4) = SystemNumber(strCaseNo, 4)
                  strSaveCaseNo1 = Trim(strField(1)) & CStr(Val(strField(2))) & IIf(strField(3) <> "0" Or strField(4) <> "00", "-" & strField(3), "") & IIf(strField(4) <> "00", "-" & strField(4), "")
                  strSaveCaseNo2 = Trim(strField(1)) & "-" & CStr(Val(strField(2))) & IIf(strField(3) <> "0" Or strField(4) <> "00", "-" & strField(3), "") & IIf(strField(4) <> "00", "-" & strField(4), "")
                  strSaveCaseNo3 = Trim(strField(1)) & CStr(strField(2)) & IIf(strField(3) <> "0" Or strField(4) <> "00", "-" & strField(3), "") & IIf(strField(4) <> "00", "-" & strField(4), "")
                  strSaveCaseNo4 = Trim(strField(1)) & "-" & CStr(strField(2)) & IIf(strField(3) <> "0" Or strField(4) <> "00", "-" & strField(3), "") & IIf(strField(4) <> "00", "-" & strField(4), "")
                  If InStr(UCase(stAttPathFile), strSaveCaseNo1) = 0 And _
                     InStr(UCase(stAttPathFile), strSaveCaseNo2) = 0 And _
                     InStr(UCase(stAttPathFile), strSaveCaseNo3) = 0 And _
                     InStr(UCase(stAttPathFile), strSaveCaseNo4) = 0 Then
                     stAttPathFile = strSaveCaseNo3 & "." & stAttPathFile
                  End If
               Else
                  'Add By Sindy 2024/11/26 多案歷程開放.CDATA.可以放多案歷程的其他案號
                  If Not (txtLpNote.Tag = "多案單筆歷程" And InStr(UCase(stAttPathFile), ".CDATA.") > 0) Then
                  '2024/11/26 END
                     If InStr(UCase(stAttPathFile), m_strSaveCaseNo1) = 0 And _
                        InStr(UCase(stAttPathFile), m_strSaveCaseNo2) = 0 And _
                        InStr(UCase(stAttPathFile), m_strSaveCaseNo3) = 0 And _
                        InStr(UCase(stAttPathFile), m_strSaveCaseNo4) = 0 Then
                        stAttPathFile = m_strSaveCaseNo3 & "." & stAttPathFile
                     End If
                  End If
               End If
            End If
            '2018/10/15 END
            
            stAttPathFile = m_AttachPath & "\" & stAttPathFile
            If Dir(stAttPathFile) <> "" Then
               SetAttr stAttPathFile, vbNormal 'Add By Sindy 2020/1/17 檔案設定為正常屬性
               Kill stAttPathFile '檔案已存在時重新下載
            End If
            lngSize = Val(.Fields("eef04").Value)
            
            'Modified by Morgan 2015/5/27 FTP上線
            'ReDim bytes(lngSize)
            'If lngSize > 0 Then bytes() = .Fields("eef05").GetChunk(lngSize)
            'iFileNo = FreeFile
            'Open stAttPathFile For Binary Access Write As #iFileNo
            'If lngSize > 0 Then Put #iFileNo, , bytes()
            'Close #iFileNo
            If Not IsNull(.Fields("eef12")) Then
               PUB_GetFtpFile .Fields("eef12"), stAttPathFile, "EMPELECTRONFILE", True
            End If
            'end 2015/5/27
            
            '加入ListBox
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stAttPathFile)
            '檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg stAttPathFile & MsgText(9221)
               Exit Function
            End If
            AddListX lstAtt(0), stAttPathFile & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(0)
            
            .MoveNext
         Loop
      End With
   End If
   If lstAtt(0).ListCount > 0 Then SetListScroll lstAtt(0)
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHnd:
   DownloadAttFile_copy = False
   Screen.MousePointer = vbDefault
   If iFileNo > 0 Then Close #iFileNo
   'Modify By Sindy 2018/10/22 ex:P-118957
   If Err.Number = 70 Then
      MsgBox "檢查檔案是否正在使用中：" & Err.Description & vbCrLf & vbCrLf & stAttPathFile, vbExclamation
   Else
      MsgBox stAttPathFile & " 檔案下載有誤！" & vbCrLf & Err.Description, vbCritical
   End If
   lstAtt(0).Clear
   'Call cmdCancel_Click 'Add By Sindy 2018/10/12 附件下載有問題,就”取消”新增下一流程
   '2018/10/22 END
End Function

'strEEP02:不可含此次要新增的流程序號
Private Function GetPreviousFlow(strEEP01 As String, intEEP02 As Integer, strEEP04 As String, _
   Optional bolShowMsg As Boolean = True) As Boolean
   m_PreviousFlow = "": m_FlowTxt = ""
   
   GetPreviousFlow = False
   
   '取得最近一道要搬檔過來的流程
   If strEEP04 = EMP_核完 Then
      m_PreviousFlow = "'" & EMP_送英核 & "','" & EMP_送核 & "'"
      m_FlowTxt = "送核"
   ElseIf strEEP04 = EMP_會完 Then
      m_PreviousFlow = "'" & EMP_送會 & "'"
      m_FlowTxt = "送會"
   'Add By Sindy 2016/3/15
   ElseIf strEEP04 = EMP_圖完 Then
      m_PreviousFlow = "'" & EMP_會圖 & "'"
      m_FlowTxt = "會(圖/文)"
   '2016/3/15 END
   'Add By Sindy 2015/4/22
   ElseIf strEEP04 = EMP_草核完 Then
      m_PreviousFlow = "'" & EMP_草核 & "'"
      m_FlowTxt = "草核"
   '2015/4/22 END
   ElseIf strEEP04 = EMP_繪圖判發 Then
      m_PreviousFlow = "'" & EMP_墨完 & "'"
      m_FlowTxt = "墨完"
   ElseIf strEEP04 = EMP_判發 Then
      'Modify By Sindy 2018/9/21 + EMP_送核:直接判發
      'Modify By Sindy 2023/10/31
'      If bolTMFlow = True Then
         m_PreviousFlow = "'" & EMP_送判 & "','" & EMP_送核 & "'"
'      Else
'      '2018/9/21 END
'         m_PreviousFlow = "'" & EMP_送判 & "'"
'      End If
      m_FlowTxt = "送判"
   'Add By Sindy 2023/10/31
   'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True
   ElseIf (bolFCPFlow = True Or bolFCTFlow = True) And strEEP04 = EMP_送件 And bolWaitReply = True Then
      'Modify By Sindy 2024/8/13 + Or PUB_GetST03(strUserNum) = "F12"
      If PUB_GetST03(strUserNum) = "F22" Or PUB_GetST03(strUserNum) = "F12" Then
         m_PreviousFlow = "'" & EMP_程序送判 & "'"
      Else
         m_PreviousFlow = "'" & EMP_送判 & "','" & EMP_送核 & "'"
      End If
      m_FlowTxt = "送件"
      '2023/10/31 END
   End If
   strSql = "select eep02,eep04 From empelectronprocess" & _
            " where eep01='" & strEEP01 & "' and eep02<=" & intEEP02 & " and eep04 in(" & m_PreviousFlow & ")" & _
            " order by eep02 desc"
   intI = 1
   m_PreviousFlow = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
         m_PreviousFlow = RsTemp.Fields("eep02")
         If m_FlowTxt = "送核" And RsTemp.Fields("eep04") = EMP_送英核 Then
            'Modify By Sindy 2015/3/16
            If m_EP41 = "2" Then '2.日
               m_FlowTxt = "送日核"
            Else
            '2015/3/16 END
               m_FlowTxt = "送英核"
            End If
         End If
      End If
   End If
   If m_PreviousFlow = "" Then
      'Modify By Sindy 2023/10/31 +And bolShowMsg = True
      If bolShowMsg = True Then
         MsgBox "沒有(" & m_FlowTxt & ")流程資料，無法沿用原附件！"
      End If
      Exit Function
   End If
   '檢查有無附件資料
   'Modify By Sindy 2018/9/4 + and eef12 is not null
   strSql = "select eef03 From empelectronFile" & _
            " where eef01='" & strEEP01 & "' and eef02='" & m_PreviousFlow & "' and eef12 is not null order by eef09,eef10 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   'If intI = 1 Then
   If intI = 0 Then
      'If RsTemp.RecordCount = 0 Then
         'Modify By Sindy 2023/10/31 +And bolShowMsg = True
         If bolShowMsg = True Then
            MsgBox "最近一道(流程順序:" & m_PreviousFlow & ")" & m_FlowTxt & "流程，沒有附件資料，無法沿用原附件！"
         End If
         m_PreviousFlow = "" 'Add By Sindy 2023/10/31
         Exit Function
      'End If
   End If
   
   GetPreviousFlow = True
End Function

'將最近一道附件移至此流程,並且將相關流程中的附件一併刪除
Private Function MoveAndDelFile(strEEP01 As String, intEEP02 As Integer, strEEP04 As String) As Boolean
Dim strDelFileFlow As String
Dim bolToChage1 As Boolean '只能改變1次
   
On Error GoTo ErrHand
   
   If strEEP04 = EMP_核完 Then
      strDelFileFlow = "'" & EMP_送英核 & "','" & EMP_送核 & "','" & EMP_核修 & "'"
   'Add By Sindy 2015/4/22
   ElseIf strEEP04 = EMP_草核完 Then
      strDelFileFlow = "'" & EMP_草核 & "','" & EMP_草修 & "'"
   '2015/4/22 END
   ElseIf strEEP04 = EMP_會完 Then
      strDelFileFlow = "'" & EMP_送會 & "','" & EMP_會修 & "'"
   'Add By Sindy 2016/3/15
   ElseIf strEEP04 = EMP_圖完 Then
      strDelFileFlow = "'" & EMP_會圖 & "','" & EMP_圖修 & "'"
   '2016/3/15 END
   ElseIf strEEP04 = EMP_繪圖判發 Then
      strDelFileFlow = "'" & EMP_墨完 & "','" & EMP_退回 & "'"
   ElseIf strEEP04 = EMP_判發 Then
      strDelFileFlow = "'" & EMP_送判 & "','" & EMP_退回 & "'"
   End If
   
   MoveAndDelFile = False
   
   '將最近一道附件移至此流程
   strSql = "update empelectronfile set eef02=" & intEEP02 & _
            " where EEF01='" & strEEP01 & "' and EEF02=" & m_PreviousFlow
   cnnConnection.Execute strSql
   '將相關流程中的附件一併刪除,會完除外
   'Modify By Sindy 2016/3/15 圖完除外
   If strEEP04 <> EMP_會完 And strEEP04 <> EMP_圖完 Then
      '逐筆反回去檢查是否為相關流程,若是則刪除附件
      strSql = "select eep01,eep02,eep04 From empelectronprocess" & _
               " where eep01='" & strEEP01 & "' and eep02<" & intEEP02 & " order by eep02 desc"
      intI = 1
      bolToChage1 = False
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               If InStr(strDelFileFlow, RsTemp.Fields("eep04")) > 0 Then
                  '要刪除附件
                  PUB_DelFtpFile2 m_EEP01, " and EEF02=" & RsTemp.Fields("eep02"), "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
                  'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
                  strSql = "delete from empelectronfile" & _
                           " where EEF01='" & strEEP01 & "'" & _
                             " and EEF02=" & RsTemp.Fields("eep02")
                  cnnConnection.Execute strSql
                  '因為送英核和送核是共用核完狀態
                  If strEEP04 = EMP_核完 Then
                     If bolToChage1 = False Then '只能改變1次
                        bolToChage1 = True
                        If RsTemp.Fields("eep04") = EMP_送英核 Then
                           strDelFileFlow = "'" & EMP_送英核 & "','" & EMP_核修 & "'"
                        Else
                           strDelFileFlow = "'" & EMP_送核 & "','" & EMP_核修 & "'"
                        End If
                     End If
                  End If
               Else
                  Exit Do
               End If
               RsTemp.MoveNext
            Loop
         End If
      End If
   End If
   
   MoveAndDelFile = True
   
ErrHand:
   Exit Function
End Function

'Add By Sindy 2018/7/16 檢查檔案是否開啟中
Private Function ChkInsFileOpening(Index As Integer) As Boolean
Dim stFileName As String
Dim fs, f
Dim strPath As String
Dim bolHadShowMsg As Boolean 'Add By Sindy 2018/9/26
   
   strPath = m_AttachPath 'Add By Sindy 2018/7/16
   ChkInsFileOpening = True
   For ii = 0 To lstAtt(Index).ListCount - 1
      stFileName = GetFileName(lstAtt(Index).List(ii), strPath)
      If Right(UCase(stFileName), 5) <> UCase(".menu") Then
         '檔案是否正在使用中
         'If PUB_ChkFileOpening(m_AttachPath & "\" & stFileName) = True Then
         If PUB_ChkFileOpening(strPath & "\" & stFileName, bolHadShowMsg) = True Then
            'Modify By Sindy 2018/9/26
            If bolHadShowMsg = False Then
            '2018/9/26 END
               'MsgBox m_AttachPath & "\" & stFileName & vbCrLf & "檔案正在使用中，請關閉後，請重新執行〔產生承辦單及歸檔〕！", vbExclamation
               MsgBox strPath & "\" & stFileName & vbCrLf & "檔案正在使用中，請關閉後，請重新執行〔產生承辦單及歸檔〕！", vbExclamation
            End If
            Me.cmdSend.Enabled = False
            ChkInsFileOpening = False
                        Screen.MousePointer = vbDefault 'Add By Sindy 2019/1/21
            Exit Function
         End If
         
         Set fs = CreateObject("Scripting.FileSystemObject")
         'Set f = fs.GetFile(m_AttachPath & "\" & stFileName)
         Set f = fs.GetFile(strPath & "\" & stFileName)
         '檔案大小為 0 KB 有誤
         If f.Size = 0 Then
            'MsgBox m_AttachPath & "\" & stFileName & vbCrLf & "檔案歸檔有誤，因檔案大小為 0 KB！請重新執行〔產生承辦單及歸檔〕！", vbExclamation
            MsgBox strPath & "\" & stFileName & vbCrLf & "檔案歸檔有誤，因檔案大小為 0 KB！請重新執行〔產生承辦單及歸檔〕！", vbExclamation
            'Me.cmdSend.Visible = False
                        Me.cmdSend.Enabled = False
            ChkInsFileOpening = False
                        Screen.MousePointer = vbDefault 'Add By Sindy 2019/1/21
            Exit Function
         End If
      End If
   Next ii
End Function

'Add By Sindy 2018/7/16 下載全部附件
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
            'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
            If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
            '2021/8/6 END
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
         End If
         'Modify By Sindy 2018/7/10
         If Right(UCase(stFileName), 5) <> UCase(".menu") Then
         '2018/7/10 END
            If InStr(stFileName, "\") = 0 Then
               If Index = 1 Then '存卷資料
                  'If GetAttachFile(stFileName, 0) = False Then
                  If PUB_GetAttachFile_EEF(m_EEP01, 0, stFileName, m_AttachPath) = False Then
                     stFileNameErr = stFileNameErr & stFileName & " 和 "
                     'Exit Function
                  End If
               Else
                  'If GetAttachFile(stFileName, intEEP02) = False Then
                  If PUB_GetAttachFile_EEF(m_EEP01, intEEP02, stFileName, m_AttachPath) = False Then
                     stFileNameErr = stFileNameErr & stFileName & " 和 "
                     'Exit Function
                  End If
               End If
               'stFileName = m_AttachPath & "\" & stFileName 'Add By Sindy 2015/9/11
            End If
         End If
         'Modify By Sindy 2015/9/10
         'Modify By Sindy 2020/9/26 + And InStr(UCase(stFileName), UCase(EMP_多案承辦單)) = 0
         If InStr(UCase(stFileName), UCase(EMP_承辦單)) = 0 And InStr(UCase(stFileName), UCase(EMP_多案承辦單)) = 0 Then
            pFiles = pFiles & ";" & stFileName 'Add By Sindy 2015/9/11
         End If
      'End If
   Next ii
   If pFiles <> "" Then pFiles = Mid(pFiles, 2) 'Add By Sindy 2015/9/11
   If stFileNameErr <> "" Then
      stFileNameErr = Left(Trim(stFileNameErr), Len(Trim(stFileNameErr)) - 1)
      MsgBox "下載附件檔有誤！(" & stFileNameErr & ")"
   End If
End Function

'Add By Sindy 2018/7/16
Private Function InsertFileData(isFileNameNoSave As String, Index As Integer) As Boolean
   Dim stFileName As String, stReName As String, stFileName2 As String
   Dim strTableName As String
   Dim UpdModifyDate As Double, UpdModifyTime As Double
   Dim bolFileSave As Boolean
   Dim strPath As String 'Add By Sindy 2018/7/16
   
On Error GoTo ErrHand
   
   strPath = m_AttachPath 'Add By Sindy 2018/7/16
   InsertFileData = True
   
   'Add By Sindy 2023/11/23 外專附件一律不歸卷，維持Backup回存信件機制；僅歸（存卷資料）區
   'Modify By Sindy 2024/8/13 外商附件一律不歸卷，維持Backup回存信件機制；僅歸（存卷資料）區
   If Index = 0 And _
      (bolFCPFlow = True Or bolCFTFlow = True Or bolFCTFlow = True) Then
      Exit Function
   End If
   
'   Screen.MousePointer = vbHourglass
'   cnnConnection.BeginTrans
   
'   If bolDel = True Then
'      strSql = "delete from CasePaperPDF where cpp01='" & m_EEP01 & "'"
'      cnnConnection.Execute strSql
'      strSql = "delete from CasePaperFile where cpf01='" & m_EEP01 & "'"
'      cnnConnection.Execute strSql
'   End If
   For ii = 0 To lstAtt(Index).ListCount - 1
      'stFileName = lstAtt(Index).List(ii)
      stFileName = GetFileName(lstAtt(Index).List(ii), strPath)
      If InStrRev(stFileName, " (") > 0 Then
         'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
         If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
         '2021/8/6 END
            stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
         End If
      End If
      'Modify By Sindy 2013/11/6
      '檢查是否有要踢除的檔案
      bolFileSave = True
      If InStr(UCase(isFileNameNoSave), UCase("DWG")) > 0 Then
         'Modify By Sindy 2014/3/11 +dwg.7z
         If Right(UCase(stFileName), 4) = UCase(".DWG") Or _
            Right(UCase(stFileName), 7) = UCase("DWG.ZIP") Or _
            Right(UCase(stFileName), 6) = UCase("DWG.7Z") Or _
            Right(UCase(stFileName), 7) = UCase("DWG.PDF") Then
            bolFileSave = False
         End If
      End If
'      'Modify By Sindy 2015/9/10
'      If InStr(UCase(stFileName), UCase(EMP_承辦單 & ".menu")) > 0 Then
'         bolFileSave = False
'      End If
'      '2015/9/10 END
      'If InStr(UCase(stFileName), UCase(isFileNameNoSave)) = 0 Then
      If bolFileSave = True Then
      '2013/11/6 END
         stReName = ""
         stFileName2 = Right(stFileName, Len(stFileName) - InStrRev(stFileName, ".") + 1)
         '更名
         Call PUB_GetEmpFlowReNameFile(PField(1), PField(2), PField(3), PField(4), cp(10), stFileName, stReName)
         
'         Set fs = CreateObject("Scripting.FileSystemObject")
'         Set f = fs.GetFile(m_AttachPath & "\" & stFileName)
'         '檔案大小為 0 KB 有誤
'         If f.Size = 0 Then
'            ShowMsg stAttPathFile & MsgText(9221)
'            Exit Function
'         End If
'         UpdModifyDate = Mid(Format(f.DateLastModified, "YYYYMMDDHHMMSS"), 1, 8)
'         UpdModifyTime = Right(Format(f.DateLastModified, "YYYYMMDDHHMMSS"), 6)

         'Add By Sindy 2021/3/24 存卷資料歸卷宗區補.info.附檔名 CFP-031067(訴願)
         If Index = 1 Then
            strExc(0) = "select * from efilecaption where efc03='存卷資料' and instr(upper('" & stReName & "'),'.'||upper(efc02)||'.')>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               stReName = Left(stReName, Len(stReName) - Len(stFileName2)) & ".INFO" & stFileName2
            End If
         End If
         '2021/3/24 END
         UpdModifyDate = Mid(lstAtt(Index).List(ii), InStr(lstAtt(Index).List(ii), "#") + 1, 8)
         UpdModifyTime = Mid(lstAtt(Index).List(ii), InStr(lstAtt(Index).List(ii), "#") + 9, 6)
         
         If UCase(stFileName2) = UCase(".PDF") Then
            strTableName = "CasePaperPDF"
            If SaveAttFile_PDF(m_EEP01, strPath & "\" & stFileName, stReName, UpdModifyDate, UpdModifyTime, False, "S") = False Then
               InsertFileData = False
               Exit Function
            End If
         'Modify By Sindy 2018/7/24
         '排除 承辦單.menu
         ElseIf UCase(stFileName2) <> UCase(".menu") Then
         '2018/7/24 END
            strTableName = "CasePaperFile"
            If SaveAttFile_Org(m_EEP01, strPath & "\" & stFileName, stReName, UpdModifyDate, UpdModifyTime, "S") = False Then
               InsertFileData = False
               Exit Function
            End If
         End If
      End If
   Next ii
'   cnnConnection.CommitTrans
   'Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
'   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   InsertFileData = False
'   Screen.MousePointer = vbDefault
'   cnnConnection.RollbackTrans
   MsgBox " 新增檔案（" & stFileName & "）至" & strTableName & "失敗！" & vbCrLf & Err.Description
End Function

'Add By Sindy 2018/7/16 歸檔
Private Function FilingFilePDF() As Boolean
Dim stFileName As String
Dim stFileTime As String
Dim arrID As Variant, intCnt As Integer
   
   FilingFilePDF = False
   Screen.MousePointer = vbHourglass
   
   '以防重覆歸卷
'   strSql = "delete from CasePaperPDF where cpp01='" & m_EEP01 & "' and cpp02='" & stFileName & "'"
'   cnnConnection.Execute strSql
   If DelAttFile_PDF(lblCaseNo.Caption, m_EEP01, "", "S", True) = False Then Exit Function
   If DelAttFile_File(lblCaseNo.Caption, m_EEP01, "", "S", True) = False Then Exit Function
   stFileTime = Right("000000" & ServerTime, 6)
   
   '新增一筆承辦單.menu至卷宗區
   'Add By Sindy 2020/10/29
   If m_RetrunRecv <> "" Then '多案
      arrID = Split(m_RetrunRecv, ",")
      For intCnt = 0 To UBound(arrID)
         If PUB_InsChkWrkSht(CStr(arrID(intCnt)), stFileTime) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
         End If
      Next intCnt
   Else
   '2020/10/29 END
      '承辦單檔案名稱
      Call PUB_ChkEmpFlowFNMRule(lblCaseNo, "", "Y", cp(10), stFileName, , False)
      stFileName = stFileName & "." & cp(10) & "." & EMP_承辦單 & ".menu"
      'Modify By Sindy 2018/11/20 + ,cpp12
      strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10,cpp12)" & _
               " values('" & m_EEP01 & "'," & _
                       "'" & stFileName & "',0,'" & strUserNum & "'," & _
                       strSrvDate(1) & "," & stFileTime & "," & _
                       strSrvDate(1) & "," & stFileTime & ",'Y','S')"
      cnnConnection.Execute strSql, intI
   End If
   
   '下載申請書附件
   'Call ReadAttachFile(m_EEP01, CInt(m_EEP02))
   'Call DownloadAllAttachFile(CInt(m_EEP02), 0)
   '檢查檔案是否開啟中
   If ChkInsFileOpening(0) = False Then
      Screen.MousePointer = vbDefault
      Exit Function
   End If
   '將電子檔分別存至卷宗區及原始檔
   If InsertFileData("DWG", 0) = True Then
      
      '下載存卷附件
      Call ReadAttachFile_other(m_EEP01)
      Call DownloadAllAttachFile(0, 1)
      '檢查檔案是否開啟中
      If ChkInsFileOpening(1) = False Then
         Screen.MousePointer = vbDefault
         Exit Function
      End If
      '將電子檔分別存至卷宗區及原始檔
      If InsertFileData("無須踢除的檔案", 1) = True Then
         'Call ReadAttachFile(m_EEP01, CInt(m_EEP02))
   'Modify By Sindy 2018/11/20
      Else
         Screen.MousePointer = vbDefault
         Exit Function
      End If
   Else
      Screen.MousePointer = vbDefault
      Exit Function
   '2018/11/20 END
   End If
   
   Screen.MousePointer = vbDefault
   FilingFilePDF = True
End Function

'Add By Sindy 2018/7/16 比對Word檔, 若無相對應檔名.PDF檔, 就自動產生一份
Private Function AutoPrintPDFfile(Index As Integer) As Boolean
   Dim stFileName As String, stFileName2 As String
   Dim strPath As String
   Dim fs, f, sFile
   Dim jj As Integer, intCnt As Integer
   Dim stPDFfileName As String, stPDFfilePathName As String
   Dim bolPDFFile As Boolean
   
On Error GoTo ErrHand
   
   strPath = m_AttachPath
   AutoPrintPDFfile = True
   '檢查檔案是否開啟中
   If ChkInsFileOpening(Index) = False Then Exit Function
   
   Screen.MousePointer = vbHourglass
'   cnnConnection.BeginTrans
   
'   If bolDel = True Then
'      strSql = "delete from CasePaperPDF where cpp01='" & m_EEP01 & "'"
'      cnnConnection.Execute strSql
'      strSql = "delete from CasePaperFile where cpf01='" & m_EEP01 & "'"
'      cnnConnection.Execute strSql
'   End If
   
   pub_OsPrinter = PUB_GetOsDefaultPrinter '取得作業系統預設印表機
'   PUB_SetOsDefaultPrinter Printers(PrinterIndex).DeviceName 'Printer.DeviceName '作業系統預設印表機指到PDFCreator
'   PUB_SetOsDefaultPrinter cboPrinter
'   PUB_SetWordActivePrinter
   For ii = 0 To lstAtt(Index).ListCount - 1
      stFileName = GetFileName(lstAtt(Index).List(ii), strPath)
      'Modify By Sindy 2018/10/4 Mark
      'T183148(1).pdf
      'T183148(1).CUS.doc
'      If InStrRev(stFileName, " (") > 0 Then
'         stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
'      End If
      stFileName2 = Right(stFileName, Len(stFileName) - InStrRev(stFileName, ".") + 1)
      If UCase(stFileName2) = ".DOC" Or UCase(stFileName2) = ".DOCX" Then
         stPDFfileName = Mid(stFileName, 1, Len(stFileName) - Len(stFileName2)) & ".pdf"
         '檢查是否已有相同檔名的PDF
         bolPDFFile = False
         For jj = 0 To lstAtt(Index).ListCount - 1
            If InStr(lstAtt(Index).List(jj), stPDFfileName) > 0 Then
               bolPDFFile = True
               Exit For
            End If
         Next jj
         
         'Added by Morgan 2025/11/11
         '半E或全E客戶案件的通知函要帶發文日期
         If InStr(LCase(stFileName), ".cus.doc") > 0 Then
            If PUB_ChkECustCase(cp(1), cp(2), cp(3), cp(4), True) = True Then
               If Not PUB_UpdCustLetterDate(strPath & "\" & stFileName) Then
                  If MsgBox("Word檔內找不到可更新的發文日，請自行確認客戶函內的發文日期是否正確，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2, "全/半E化客戶案件提醒") = vbNo Then
                     GoTo ErrHand
                  End If
                  
               End If
            End If
         End If
         'end 2025/11/11
         
         If bolPDFFile = False Then
            '開啟Word檔
            Set g_WordAp = New Word.Application
            g_WordAp.Visible = True
            g_WordAp.Documents.Open FileName:=strPath & "\" & stFileName ', ReadOnly:=True
            'Modify By Sindy 2018/9/20 用Word轉Pdf功能
            frmPDF.Show
            If pub_Word2Pdf Then
               g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strPath & "\" & stPDFfileName, ExportFormat:=17, OpenAfterExport:=False
            Else
               '轉PDF
               'frmPDF.Show
               frmPDF.StartProcess strPath, Mid(stFileName, 1, Len(stFileName) - Len(stFileName2))  'stFileName
               '切換印表機
               If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord
               g_WordAp.ActivePrinter = PUB_PdfCreatorNameInWord
               g_WordAp.ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
               frmPDF.EndtProcess
               'Unload frmPDF
            End If
            Unload frmPDF
            g_WordAp.Quit wdDoNotSaveChanges
            Set g_WordAp = Nothing
            '記錄檔案位置
            stPDFfilePathName = stPDFfilePathName & "*" & strPath & "\" & stPDFfileName
         End If
      End If
   Next ii
   PUB_SetOsDefaultPrinter pub_OsPrinter '復原作業系統預設印表機

   If stPDFfilePathName <> "" Then
      stPDFfilePathName = Mid(stPDFfilePathName, 2)
      sFile = Split(stPDFfilePathName, "*")
      For intCnt = 0 To UBound(sFile)
         'stFileName = GetFileName(sfile(intCnt), strPath)
         Set fs = CreateObject("Scripting.FileSystemObject")
         Set f = fs.GetFile(sFile(intCnt))
         AddListX lstAtt(Index), sFile(intCnt) & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)
      Next intCnt
   End If
'   cnnConnection.CommitTrans
   'Call ReadAttachFile(m_EEP01, CInt(m_AttEEP02), True)
'   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   AutoPrintPDFfile = False
   Screen.MousePointer = vbDefault
'   cnnConnection.RollbackTrans
   MsgBox " 產生檔案（" & stFileName & "）失敗！" & vbCrLf & Err.Description
End Function

'解析總收文號
Private Function AnalyzeRecv(ByRef m_RetrunRecvCnt As Integer, ByVal strRetrunRecv As String, _
   Optional strShowText As String = "") As String
Dim rsTmp As New ADODB.Recordset
Dim bolShowCPMNm As Boolean 'Add By Sindy 2020/12/1
   
   '有回傳總收文號
   AnalyzeRecv = "": m_RetrunRecvCnt = 0
   If strRetrunRecv <> "" And strRetrunRecv <> m_EEP01 Then
      strExc(0) = "select cp01,cp02,cp03,cp04,cp09,decode('" & m_Country & "','000',cpm03,cpm04) cpmNm" & _
                  " From caseprogress,casepropertymap" & _
                  " where cp09 in('" & Replace(strRetrunRecv, ",", "','") & "')" & _
                  " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
                  " order by 1,2,3,4 asc"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(0) = ""
         'Add By Sindy 2020/12/1
         '先檢查是否有案件性質不同,若有,要加註案件性質名稱
         bolShowCPMNm = False
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If lblCP10.Caption <> "" & rsTmp.Fields("cpmNm") Then
               bolShowCPMNm = True
               Exit Do
            End If
            rsTmp.MoveNext
         Loop
         '2020/12/1 END
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            m_RetrunRecvCnt = m_RetrunRecvCnt + 1 '總收文號數量
            strExc(0) = strExc(0) & "," & rsTmp.Fields("cp01") & "-" & rsTmp.Fields("cp02") & IIf(rsTmp.Fields("cp03") & rsTmp.Fields("cp04") <> "000", "-" & rsTmp.Fields("cp03"), IIf(rsTmp.Fields("cp04") <> "00", "-" & rsTmp.Fields("cp04"), "")) & _
                        IIf(bolShowCPMNm = True, rsTmp.Fields("cpmNm"), "")
            rsTmp.MoveNext
         Loop
         If strExc(0) <> "" Then
            strExc(0) = Mid(strExc(0), 2)
            'Add By Sindy 2020/6/9
            If Trim(strShowText) <> "" Then
               AnalyzeRecv = "(" & strExc(0) & "一併" & strShowText & ")"
            Else
            '2020/6/9 END
               AnalyzeRecv = "(" & strExc(0) & "一併" & Trim(Mid(CboEEP04.Text, 3)) & ")"
            End If
         End If
'            If MsgBox(strExc(0) & vbCrLf & vbCrLf & _
'               "是否一併做" & Trim(Mid(CboEEP04.Text, 3)) & "？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'               rsTmp.Close
'               cmdSend.Enabled = True
'               Exit Sub
'            End If
      End If
      rsTmp.Close
   End If
   
   Set rsTmp = Nothing
End Function

'Add By Sindy 2020/10/12
Private Sub SetTxtLpNote(bolQueryStar As Boolean)
Dim varTmp As Variant
Dim i As Integer
   
   txtLpNote = ""
   bolManyCaseToMix = False: m_RetrunRecvToMix = ""
   If bolQueryStar = True Or cmdManyCase.Tag = "" Then
      txtLpNote.Tag = ""
      m_RetrunRecv = ""
      
      '查詢時
      If bolQueryStar = True Then
         'Modify By Sindy 2020/12/14 + And InStr(m_EEP11, "多案單筆歷程") > 0
         If m_EEP15 <> "" And InStr(m_EEP11, "多案單筆歷程") > 0 Then
            txtLpNote = "(共" & UBound(Split(m_EEP15, ",")) + 1 & "筆)"
         End If
      Else
         '會完重修,應檢查會完是否為多筆
         If Left(CboEEP04.Text, 2) = EMP_會完重修 Then
            strSql = "select * From empelectronprocess" & _
                     " where eep01='" & m_EEP01 & "' and eep04='" & EMP_送會 & "'" & _
                     " order by eep02 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If InStr("" & RsTemp.Fields("eep11"), "多案單筆歷程") > 0 Then
                  If MsgBox("此為多案單筆歷程，確定要" & AnalyzeRecv(UBound(Split(RsTemp.Fields("eep15"), ",")) + 1, RsTemp.Fields("eep15")) & "嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                     m_RetrunRecv = "" & RsTemp.Fields("eep15")
                     If InStr("" & RsTemp.Fields("eep11"), "多案單筆歷程") > 0 Then
                        txtLpNote.Tag = "多案單筆歷程"
                     End If
                     txtLpNote = "(共" & UBound(Split(m_RetrunRecv, ",")) + 1 & "筆)"
                  End If
               End If
            End If
            
         '等待回覆中
         'Modify By Sindy 2025/2/24 + or 程序送判
         ElseIf m_EEP15 <> "" And _
            (bolWaitReply = True Or Left(CboEEP04.Text, 2) = EMP_程序送判) Then
            m_RetrunRecv = m_EEP15
            If InStr(m_EEP11, "多案單筆歷程") > 0 Then
               txtLpNote.Tag = "多案單筆歷程"
            End If
            txtLpNote = "(共" & UBound(Split(m_EEP15, ",")) + 1 & "筆)"
            
'         Else
'            strSql = "select * From empelectronprocess" & _
'                     " where eep01='" & m_EEP01 & "' and instr('" & Replace(EMP_待辦歷程查詢除外的狀態, "'", "") & "',eep04)=0" & _
'                     " order by eep02 desc"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               If InStr("" & RsTemp.Fields("eep11"), "多案單筆歷程") > 0 Then
'                  m_RetrunRecv = "" & RsTemp.Fields("eep15")
'                  If InStr("" & RsTemp.Fields("eep11"), "多案單筆歷程") > 0 Then
'                     txtLpNote.Tag = "多案單筆歷程"
'                  End If
'                  txtLpNote = "(共" & UBound(Split(m_RetrunRecv, ",")) + 1 & "筆)"
'               End If
'            End If
         End If
      End If
      
   '確定
   Else
      If m_RetrunRecv = m_EEP01 Or m_RetrunRecv = "" Then
         txtLpNote.Tag = ""
         m_RetrunRecv = ""
         cmdManyCase.Visible = False
         cmdManyCase.Enabled = False
      Else
         txtLpNote = "(共" & UBound(Split(m_RetrunRecv, ",")) + 1 & "筆)"
      End If
      
      '智權部的多案(多文)
      '有可能有的文號是多案單筆歷程,也有單案文號
      If intReceiveKind = 2 Then '待會稿區
         varTmp = Split(m_RetrunRecv, ",")
         For i = 0 To UBound(varTmp)
            strSql = "select * From empelectronprocess" & _
                     " where eep01='" & varTmp(i) & "' and eep04='" & m_strLastEEP04 & "'" & _
                     " and eep05='" & m_FlowUserNum & "' and eep09='Y'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If InStr("" & RsTemp.Fields("eep11"), "多案單筆歷程") > 0 Then
                  bolManyCaseToMix = True
               End If
               m_RetrunRecvToMix = m_RetrunRecvToMix & "," & IIf("" & RsTemp.Fields("eep15") = "", varTmp(i), RsTemp.Fields("eep15"))
            End If
         Next i
         If bolManyCaseToMix = True Then
            txtLpNote.Visible = False
         End If
         '檢查操作的文號本身是否為”多案單筆歷程”
         If m_EEP15 <> "" Then
            If InStr(m_EEP11, "多案單筆歷程") > 0 Then
               txtLpNote.Tag = "多案單筆歷程"
            End If
         End If
      End If
   End If
   
   If Frame6.Tag = "V" Then 'Add By Sindy 2020/10/16
      If txtLpNote.Tag = "多案單筆歷程" Then '109/9/29 承慧:可以做多案即為簡單案件
         Frame6.Visible = False
      Else
         Frame6.Visible = True
      End If
   End If
End Sub

'送出
'Modify By Sindy 2020/7/21
'Modify By Sindy 2022/4/26 Private => Public
Public Sub cmdSend_Click()
   'Add By Sindy 2025/10/14
   If Frame1Big.Visible = True Then
      Call cmdClose_Click
      Exit Sub
   End If
   '2025/10/14 END
   Call FormSave
End Sub
Private Sub FormSave()
Dim strUpdDate As String, strUpdTime As String
Dim intMaxEEP02 As Integer, strUpdEEP09 As String
Dim rsTmp As New ADODB.Recordset
Dim strCP09 As String, strCP14 As String, strCP06 As String, strCP07 As String
Dim strCP27 As String, strCP10n As String, strCP48 As String
Dim strAutoFlow As String 'Add By Sindy 2013/9/23
Dim strConSql As String
'Dim strTemp As String
Dim bolSendMail As Boolean 'Add By Sindy 2014/1/15 是否要寄Mail
'Dim strPP05 As String 'Add By Sindy 2014/9/16
'Dim strChkEEP13_EEP01 As String 'Add By Sindy 2016/8/5 記錄文號要檢查EEP13怎麼沒有更新為Y
Dim bolNotSendDept As Boolean 'Add By Sindy 2018/10/1 不經發文室送件
Dim intFileCnt As Integer 'Add By Sindy 2018/10/8 附件數
Dim strRetrunRecvText As String
Dim strGetCP13 As String, strGetCP12 As String 'Add By Sindy 2018/12/3
Dim bolRegMail As Boolean 'Add By Sindy 2020/2/18 是否掛號直寄
Dim strLP31 As String 'Add By Sindy 2020/2/18 收件人為代理人
Dim arrID As Variant, intCnt As Integer
Dim strLP11 As String
Dim strTo As String 'Add By Sindy 2021/3/2
Dim bolAddrIsNull As Boolean, strCU80 As String 'Modify By Sindy 2021/3/18
Dim strCU126 As String 'Add By Sindy 2024/11/11
Dim bolConn As Boolean 'Add By Sindy 2021/3/31
Dim strPA176 As String 'Added by Morgan 2021/7/21
Dim bolSendAppMail As Boolean 'Add By Sindy 2021/10/13
Dim bolHadCallCP163 As Boolean 'Add By Sindy 2023/3/10
Dim bolModify As Boolean, bolAdd As Boolean, objText As Object
Dim strCP113 As String 'Add By Sindy 2023/10/2
Dim s As Integer
Dim strCFP209EPP01() As String, strCFP209EPP08() As String, strCFP209Subj() As String, iRec As Integer 'Added by Morgan 2025/4/17
Dim oRunform As Form 'Add By Sindy 2025/8/1
   
   'Add By Sindy 2025/8/1
   If Left(CboEEP04.Text, 2) = EMP_收文分析 Then
      Set oRunform = Forms(0).GetForm("frm090801_New")
      '1.系統自動開啟案件接洽單畫面帶入本所案號相關資料，加入案件性質為[分析]
      '2.且此C類來函總收文號為該分析的相關總收文號，並由智權人員於備註中說明希補強之分析方向，而後進行收文，
      '3.分析案的本所及法定期限為Ｃ類來函期限。
      Screen.MousePointer = vbHourglass
      oRunform.bolExternalCall = True '記錄是外部程式呼叫使用
      oRunform.SetParent Me
      oRunform.Show
      oRunform.Tag = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
      oRunform.Option1(1).Value = True
      oRunform.Text1(6) = cp(1)
      oRunform.Text1(7) = cp(2)
      oRunform.Text1(8) = cp(3)
      oRunform.Text1(9) = cp(4)
      oRunform.m_FlowUserNum = m_FlowUserNum '案件流程所屬人員
      oRunform.cmdCRL55.Visible = False: oRunform.Label1(84) = "C類來函：" '原為本案與總號：
      oRunform.Text1(100) = m_EEP01
      oRunform.m_strGetNP01 = "727:" & m_EEP01 '要預設的案件性質+:+C類來函總收文號
      Call oRunform.Text1_LostFocus(9)
      oRunform.Text1(3) = TransDate(cp(6), 1) '本所期限=Ｃ類來函本所期限
      oRunform.Text1(1) = TransDate(cp(7), 1) '法定期限=Ｃ類來函法定期限
      '預設文件齊備、須會稿
      oRunform.OptEP06(0).Value = True '文件齊備
      oRunform.OptEP34(0).Value = True '會稿
      oRunform.bolExternalCall = False '還原預設值
      Screen.MousePointer = vbDefault
      Me.Hide
      'm_PrevForm.Hide
      'Call cmdCancel_Click '”取消”新增下一流程
      'Call cmdExit_Click
      Exit Sub
   End If
   Call GetCP43AddCC(True)
   '2025/8/1 END
   
   'Add By Sindy 2020/10/7
   Call SetTxtLpNote(False)
   If m_RetrunRecv <> "" Then '有子案
      m_RetrunRecvSub = Replace(Replace(m_RetrunRecv, m_EEP01, ""), ",,", ",")
      If Left(m_RetrunRecvSub, 1) = "," Then m_RetrunRecvSub = Mid(m_RetrunRecvSub, 2)
      If Right(m_RetrunRecvSub, 1) = "," Then m_RetrunRecvSub = Left(m_RetrunRecvSub, Len(m_RetrunRecvSub) - 1)
   End If
   '2020/10/7 END
   
   '解析總收文號
   'Add By Sindy 2018/9/27
   'Modify By Sindy 2020/10/19
   If bolManyCaseToMix = True Then
      '操作的文號本身是”多案單筆歷程”,抓出要一併更新的文號
      If txtLpNote.Tag = "多案單筆歷程" Then
         m_RetrunRecvSub = m_EEP15
         m_RetrunRecvSub = Replace(Replace(m_RetrunRecvSub, m_EEP01, ""), ",,", ",")
         If Left(m_RetrunRecvSub, 1) = "," Then m_RetrunRecvSub = Mid(m_RetrunRecvSub, 2)
         If Right(m_RetrunRecvSub, 1) = "," Then m_RetrunRecvSub = Left(m_RetrunRecvSub, Len(m_RetrunRecvSub) - 1)
      End If
      
      strRetrunRecvText = AnalyzeRecv(m_RetrunRecvCnt, m_RetrunRecvToMix)
   '2020/10/19 END
   ElseIf (cmdManyCase.Visible = True And cmdManyCase.Enabled = True) Or txtLpNote.Tag = "多案單筆歷程" Then
      strRetrunRecvText = AnalyzeRecv(m_RetrunRecvCnt, m_RetrunRecv)
   End If
   
   'Add By Sindy 2021/3/23 提醒承辦人有"多案單筆歷程"
   If (cmdManyCase.Visible = True And cmdManyCase.Enabled = True) _
      And cp(163) <> "" And strRetrunRecvText = "" Then
      If intReceiveKind = 0 Then '0.承辦人工作進度
         If MsgBox("此案有要做【多案】單筆歷程嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            Call cmdManyCase_Click '開啟多案歷程-選取作業
            Exit Sub
         'Add By Sindy 2023/3/10
         Else
            bolHadCallCP163 = True
            '2023/3/10 END
         End If
      End If
   End If
   '2021/3/23 END
   'Add By Sindy 2023/3/10 檢查是否有掉案件
   If strRetrunRecvText = "" Then
      strSql = "select CP09,CP01,CP02,CP03,CP04,CP163 From caseprogress where cp163='" & m_EEP01 & "' and cp09<>'" & m_EEP01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strExc(10) = ""
         Do While Not RsTemp.EOF
            strExc(10) = strExc(10) & RsTemp.Fields("CP09") & ":" & RsTemp.Fields("CP01") & RsTemp.Fields("CP02") & RsTemp.Fields("CP03") & RsTemp.Fields("CP04") & vbCrLf
            RsTemp.MoveNext
         Loop
'         If CheckIsPersonRest("97038", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
'            PUB_SendMail strUserNum, "97038", "", "(系統檢查)內商歷程有掉案件！ bolHadCallCP163=" & bolHadCallCP163 & "; True:放棄,做【多案】單筆歷程", _
'                        m_EEP01 & ":" & PField(1) & PField(2) & PField(3) & PField(4) & vbCrLf & vbCrLf & _
'                        strExc(10) & vbCrLf & strSql & vbCrLf, , , , , , , , , , True, False, , , False
'         End If
         If bolHadCallCP163 = False Then
            If MsgBox("此案之前有做【多案】單筆歷程，目前此歷程尚未點選其他案件，確定是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Call cmdCancel_Click
               Exit Sub
            End If
         End If
      End If
   End If
   '2023/3/10 END
   
   'Modify By Sindy 2017/8/10 有move位置
   'Add By Sindy 2013/11/26 檢查是否有專利處的核判權限
   If Left(CboEEP04.Text, 2) = EMP_判發 Then
      'Add By Sindy 2018/3/5 承辦人非程序人員時,才需檢查核判權限
      If GetStaffDepartment(strUserNum) <> "P12" Then
      '2018/3/5 END
         If m_FlowUserNum <> strUserNum Then
            'Modify By Sindy 2022/3/16 程式有特別檢查判發時,若為代理狀況,還是會檢查 strUserNum 是否有判發權限~
            '改以原判發人(m_FlowUserNum)的權限檢查。 ex:CFP-032770
            'strUserNum ==> m_FlowUserNum
            If CStr(PField(1)) = "ACS" And m_FlowUserNum = "A5024" Then
               'ACS目前核判主管只有A5024.王琇娟
            'Add By Sindy 2024/4/2
            'Modify By Sindy 2024/8/13 + Or bolCFTFlow = True Or bolFCTFlow = True
            ElseIf bolFCPFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
               '外專不檢查此權限
            Else
            '2022/7/1 end
               'Modify By Sindy 2024/6/26 +m_Country
               If PUB_ChkPromoterReader(CStr(PField(1)), cp(10), "2", m_FlowUserNum, , m_Country) = False Then
'                  'Add By Sindy 2024/1/5
'                  If bolFCPFlow = True Then
'                     strSql = "select st01 from staff" & _
'                              " where st93 in(select st93 From staff where st01='" & m_FlowUserNum & "')" & _
'                              " and st04='1'" & _
'                              " and (st52='" & strUserNum & "' or st53='" & strUserNum & "' or st54='" & strUserNum & "' or st55='" & strUserNum & "')"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 0 Then
'                        MsgBox "無代理判發的權限！" & vbCrLf & "請回至前一作業『工作進度資料維護』輸入判發人後，才可進行下一流程。"
'                        Call cmdCancel_Click
'                        Exit Sub
'                     End If
'                  Else
'                  '2024/1/5 END
                     MsgBox "無代理判發的權限！" & vbCrLf & "請回至前一作業『工作進度資料維護』輸入判發人後，才可進行下一流程。"
                     Call cmdCancel_Click
                     Exit Sub
'                  End If
               End If
            End If
         End If
      End If
   End If
   '2013/11/26 END
   
   'Added by Morgan 2021/7/21
   strPA176 = ""
   'Modified by Morgan 20201/8/6
   'If Left(CboEEP04.Text, 2) = EMP_判發 Then
   'Modify By Sindy 2023/10/2 + And bolPAFlow = True
   If Left(CboEEP04.Text, 2) = EMP_送判 And bolPAFlow = True Then
   'end 2021/8/6
      '大陸發明生醫案是否新藥專利設定
      If pa(1) = "P" And pa(9) = "020" And pa(8) = "1" And pa(158) = "3" And (cp(10) = "101" Or cp(10) = "307") Then
         intI = MsgBox("是否新藥專利？", vbYesNoCancel + vbDefaultButton3 + vbQuestion, "大陸發明生醫案是否新藥專利確認")
         If intI = vbYes Then
            strPA176 = "Y"
         ElseIf intI = vbNo Then
            strPA176 = "N"
         Else
            Exit Sub
         End If
      End If
   End If
   'end 2021/7/21
   
   'Added by Moran 2024/1/5
   '承辦於電子歷程控管「判發」流程，系統自動將「本案需於指定日方可送件」或「本案需於指定日之後方可送件」帶入歷程備註，以利提醒程序人員～
   If strSrvDate(1) >= 指定日期啟用日 And Left(CboEEP04.Text, 2) = EMP_判發 And bolPAFlow = True Then
      If cp(141) = "3" And (cp(164) = "1" Or cp(164) = "3") Then
         '因為是要提醒程序，只要能帶入內容，可不必彈訊息--郭
         strExc(0) = "系統提醒:本案需於指定日" & ChangeWStringToTDateString(cp(142)) & IIf(cp(164) = "3", "之後", "") & "方可送件。"
         If InStr(txtEEP08, strExc(0)) = 0 Then
            txtEEP08 = txtEEP08 & vbCrLf & strExc(0)
         End If
      End If
   End If
   'end 2024/1/5
   
   'Added by Morgan 2025/4/17-- 品薇
   '其他CFP案的主案(通常是美國)，每次會稿及送判時，若有相對應的日本及德國，且承辦人為外翻人員
   '，則跳出提醒視窗「本案有日本/德國案，日本/德國案於N年N月N日(檢視中說發文日)已進行中翻日/德，
   '請確認日本/德國案是否有需要修改，若有需要請連絡品薇」，工程師必須選擇「是」或「否」，若選擇是，
   '則系統自動於該相對應的日本及/或德國案以該工程師的名義發聯絡給品薇並副本給柏翰，
   '內容為「本案的相對應他國案CFPXXX(工程師會稿的案件)有修改，且本日本/德國案需一併進行修改，請與承辦工程師連絡。」
   If PField(1) = "CFP" And InStr(NewCasePtyList, cp(10)) > 0 And cp(21) = "" And (Left(CboEEP04.Text, 2) = EMP_送會 Or Left(CboEEP04.Text, 2) = EMP_送判) Then
      strSql = "select cr01||'-'||cr02||decode(cr03||cr04,'000','','-'||cr03||'-'||cr04) CaseNo" & _
         ",sqldatet(c1.cp27) DDate,c2.cp09,c2.cp14,na03" & _
         " from caserelation,patent,nation,caseprogress c1,caseprogress c2" & _
         " Where cr01='CFP'" & _
         " and cr05='" & PField(1) & "' and cr06='" & PField(2) & "' and cr07='" & PField(3) & "' and cr08='" & PField(4) & "'" & _
         " and pa01(+)=cr01 and pa02(+)=cr02 and pa03(+)=cr03 and pa04(+)=cr04 and pa09 in ('011','231') and na01(+)=pa09" & _
         " and c1.cp01(+)=pa01 and c1.cp02(+)=pa02 and c1.cp03(+)=pa03 and c1.cp04(+)=pa04 and c1.cp10='209' and c1.cp27>0" & _
         " and c2.cp01(+)=pa01 and c2.cp02(+)=pa02 and c2.cp03(+)=pa03 and c2.cp04(+)=pa04" & _
         " and c2.cp10 in(" & NewCasePtyList & ") and c2.cp14 like 'F%' and c2.cp27 is null and c2.cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         Do While Not RsTemp.EOF
            strExc(0) = "本案有" & RsTemp("na03") & "案(" & RsTemp("caseno") & ")已於 " & RsTemp("DDate") & " 進行中翻" & Left(RsTemp("na03"), 1) & "，請確認該" & RsTemp("na03") & "案是否有需要修改，若有需要請連絡品薇"
            strExc(0) = strExc(0) & vbCrLf & vbCrLf & "請選擇「是」或「否」"
            If MsgBox(strExc(0), vbYesNo + vbQuestion) = vbYes Then
               iRec = iRec + 1
               ReDim Preserve strCFP209EPP01(iRec) As String, strCFP209EPP08(iRec) As String, strCFP209Subj(iRec) As String
               strCFP209EPP01(iRec) = RsTemp("cp09")
               strCFP209EPP08(iRec) = "本案的相對應他國案CFP-" & PField(2) & IIf(PField(3) & PField(4) = "000", "", "-" & PField(3) & "-" & PField(4)) & "有修改，且本" & RsTemp("na03") & " 案需一併進行修改，請與承辦工程師連絡。"
               strCFP209Subj(iRec) = RsTemp("caseno") & RsTemp("na03") & "案的相對應他國案CFP-" & PField(2) & IIf(PField(3) & PField(4) = "000", "", "-" & PField(3) & "-" & PField(4)) & "有修改，且本" & RsTemp("na03") & "案需一併進行修改，請與承辦工程師連絡。"
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If
   'end 2025/4/17

   'Add By Sindy 2015/4/7 截取掉”內容”最後多餘的折行
   For ii = Len(txtEEP08) - 1 To 1 Step -2
      If Mid(txtEEP08, ii, 2) = vbCrLf Then
         txtEEP08 = Mid(txtEEP08, 1, ii - 1)
      Else
         Exit For
      End If
   Next ii
   '2015/4/7 END
   
   'Add By Sindy 2023/9/20
   If bolFCPFlow = True Then
      If Left(CboEEP04.Text, 2) = EMP_送轉檔 Then
         If ChkEED13.Value = 0 Then
            s = MsgBox("轉檔後是否送程序送件？(是.發文 否.送回工程師)", vbExclamation + vbYesNoCancel + vbDefaultButton1, "重要訊息！")
            If s = vbYes Then
               ChkEED13.Value = 1
            ElseIf s = vbCancel Then
               Exit Sub
            End If
         End If
      End If
'      If Left(CboEEP04.Text, 2) <> EMP_聯絡 Or _
'         (Left(CboEEP04.Text, 2) = EMP_聯絡 And txt3(1) <> "" And txt3(2) <> "") Then
         If SSTab1.TabVisible(intTab_外專承辦單) = True Then
'            '先檢查是否需要儲存承辦單資料
'            bolModify = False
'            bolAdd = True
'            For Each objText In Me.txt3
'               If objText.Text <> objText.Tag Then
'                  bolModify = True '檢查是否有異動資料
'               End If
'               If objText.Tag <> "" Then
'                  bolAdd = False '檢查是否為新增
'               End If
'            Next
'            'Modify By Sindy 2025/4/7 + ChkEED08
'            If Val(ChkEED13.Tag) <> ChkEED13.Value Or Val(ChkEED08.Tag) <> ChkEED08.Value Then
'               bolModify = True '檢查是否有異動資料
'            End If
'            If bolModify = True Then
'               '新增
'               If bolAdd = True Then
'                  If funSaveEmpPaperData_FCP = False Then
'                     cmdSend.Enabled = True
'                     Exit Sub
'                  End If
'               '修改
'               Else
'                  'Modify By Sindy 2025/4/7 + ChkEED08
'                  If Val(ChkEED13.Tag) <> ChkEED13.Value Or Val(ChkEED08.Tag) <> ChkEED08.Value Then
'                     If funSaveEmpPaperData_FCP = False Then
'                        cmdSend.Enabled = True
'                        Exit Sub
'                     End If
'                  Else
'                     If MsgBox("承辦單資料有異動，要儲存資料嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
'                        If funSaveEmpPaperData_FCP = False Then
'                           cmdSend.Enabled = True
'                           Exit Sub
'                        End If
'                     End If
'                  End If
'               End If
'            End If
            'Modify By Sindy 2025/4/22
            If funSaveEmpPaperData_FCP(True) = False Then
               cmdSend.Enabled = True
               Exit Sub
            End If
            '2025/4/22 END
         End If
'      End If
   Else
   '2023/9/20 END
      'Modify By Sindy 2013/11/18 +if
      'Modify By Sindy 2013/12/24
      If intReceiveKind = 0 Or (intReceiveKind = 1 And Left(m_DMMan, 5) <> m_FlowUserNum) Then '0.承辦人工作進度 1.待核判區
      '2013/12/24 END
      '2013/11/18 END
         If Left(CboEEP04.Text, 2) <> EMP_聯絡 Or _
            (Left(CboEEP04.Text, 2) = EMP_聯絡 And txt1(5) <> "" And txt1(0) <> "") Then 'Add By Sindy 2013/8/16 +if
            'Modify By Sindy 2023/1/9 + And SSTab1.TabVisible(intTab_承辦單) = True : 目前只有專利處有此頁籤
            If Val(m_EP06) > 0 And SSTab1.TabVisible(intTab_承辦單) = True Then 'Add By Sindy 2013/9/26 有文件齊備後才可以儲存承辦單
               '先檢查是否需要儲存承辦單資料
               bolModify = False
               bolAdd = True
               For Each objText In Me.txt1
                  If objText.Text <> objText.Tag Then
                     bolModify = True '檢查是否有異動資料
                  End If
                  If objText.Tag <> "" Then
                     bolAdd = False '檢查是否為新增
                  End If
               Next
               If bolModify = True Then
                  '新增
                  If bolAdd = True Then
                     If funSaveEmpPaperData = False Then
                        cmdSend.Enabled = True
                        Exit Sub
                     End If
                  '修改
                  Else
                     If MsgBox("承辦單資料有異動，要儲存資料嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                        If funSaveEmpPaperData = False Then
                           cmdSend.Enabled = True
                           Exit Sub
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   
   'Add By Sindy 2013/9/26 儲存存卷資料
   If Me.cmdSave.Visible = True Then
      If MsgBox("存卷資料有異動，是否要儲存？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
         'Modify By Sindy 2016/11/9
         'Call cmdSave_Click
         'Add By Sindy 2018/11/27 發文歸檔時,以防止清掉了附件區檔案
         m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum & "\otherFile"
         If funSaveEEF02_0 = False Then
            m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum 'Add By Sindy 2018/11/27 恢復暫存檔位置
            cmdSend.Enabled = True
            Exit Sub
         End If
         m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum 'Add By Sindy 2018/11/27 恢復暫存檔位置
         '2016/11/9 END
      End If
   End If
   
On Error GoTo ErrHand
   cmdSend.Tag = "" 'Add By Sindy 2021/3/31 抓程式err
   
   '檢查條件
   cmdSend.Enabled = False 'Add By Sindy 2025/6/9 純為了 TxtValidate 裡判斷用
   If TxtValidate = False Then
'      'Add By Sindy 2016/3/10 系統欲直接離開此作業
'      If cmdExit.Enabled = False Then
'         Call cmdExit_Click
'         Exit Sub
'      End If
'      '2016/3/10 END
'      cmdSend.Enabled = True
      cmdSend.Enabled = True 'Add By Sindy 2025/6/9
      Exit Sub
   End If
   cmdSend.Enabled = True 'Add By Sindy 2025/6/9 cmdSend 後面會有鎖住的時候,但不是在此處
   
   bolNotSendDept = False
   bolRegMail = False: strLP31 = ""
   'Add By Sindy 2020/2/18 商標處客戶函電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 And bolTMFlow = True Then
      '非台灣案時,AB類要做指示信寄送
      If (Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送) And _
         m_Country <> "000" And _
         (Left(lblCP09, 1) = "A" Or Left(lblCP09, 1) = "B") Then
         
         'Add By Sindy 2021/3/2 PATTA商標監視(第07,08,09,19,20類)，其他 (我方案號：TT-000161)
         strTo = PUB_GetFCeMailConText("Main_EMail", cp(1), cp(2), cp(3), cp(4), "CF", , True)
         If strTo <> "Y00000000" Then
         '2021/3/2 END
            'Modify By Sindy 2021/10/13
            bolSendAppMail = False
            If Left(CboEEP04.Text, 2) = EMP_退件重送 Then
               If MsgBox("要寄發指示信嗎？", vbYesNo + vbCritical + vbDefaultButton1, "詢問") = vbYes Then
                  bolSendAppMail = True
               End If
            Else
               bolSendAppMail = True
            End If
            If bolSendAppMail = True Then
            '2021/10/13 END
               If PUB_T_AppFormSendMail(m_EEP01, IIf(m_RetrunRecv <> "", m_RetrunRecv, m_EEP01), _
                     PField(1), PField(2), PField(3), PField(4), cp(10), Me, lstAtt(0)) = False Then
                  Exit Sub
               End If
            End If
         End If
         
      '要通知客戶,才需要詢問
      ElseIf Left(CboEEP04.Text, 2) = EMP_發文歸檔 And ChkEP11.Value = 0 Then
         'Modify By Sindy 2021/3/18 + bolAddrIsNull
         strLP11 = PUB_SetLP11(m_PA26, m_PA75, strLP31, bolAddrIsNull)
         If bolMCTFcase = True Then
            'MCTF案一般都是E-Mail送件出去,等E-Mail狀況回來後,沒寄成功時,才會再送紙本出去
            '客戶函電子化後,就直接用”寄發文件”操作寄發E-Mail
            
         'Modify By Sindy 2021/3/4 + If Left(cp(12), 1) <> "F"
         '外商收文,尚不新增信函資料
         ElseIf Left(cp(12), 1) <> "F" Then
            If strLP11 = "Y" Then '直寄,無地址不會是Y
               If GetPrjNationNumber1(m_PA26) < "010" Then '台灣申請人
                  '有期限為掛號直寄
                  'Modify By Sindy 2025/1/23 1005.延期受理非一定是掛號直寄 ex:T-251569 + And cp(10) <> "1005"
                  If Val(cp(7)) > 0 And cp(10) <> "1005" Then bolRegMail = True
               End If
               '非台灣申請人要詢問是否掛號直寄
               'ex:T-214110(台灣申請人).對方補充理由沒有期限,但是要直寄
               If bolRegMail = False Then
                  If MsgBox("通知函是否掛號直寄?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
                     bolRegMail = True
                  End If
               End If
            'Add By Sindy 2021/3/18
            ElseIf bolAddrIsNull = True Then
               strCU80 = GetPrjNationNumber1(m_PA26, "CU80")
               strCU126 = GetPrjNationNumber1(m_PA26, "CU126") 'Add By Sindy 2024/11/11
               If strCU80 <> "" Then
                  If InStr(strCU80, "業務自行處理") > 0 And strCU126 = "Y" Then 'Y=商標以Email通知
                     MsgBox "此客戶狀態為【" & strCU80 & "】且為【半E化】" & vbCrLf & vbCrLf & "不需列印客戶函，智權人員自行上【寄發文件】處理。", vbExclamation
                  Else
                     MsgBox "此客戶狀態為【" & strCU80 & "】" & vbCrLf & vbCrLf & "客戶函請交智權人員處理。", vbExclamation
                  End If
               End If
            '2021/3/18 END
            End If
         End If
      End If
   ElseIf bolTMFlow = True Then
   '2020/2/18 END
      'Add By Sindy 2018/10/1
      'MCTF案件C類歸檔時才需要詢問
      'If InStr(m_SPMan, "MCTF") > 0 And Left(m_EEP01, 1) = "C" And
      If bolMCTFcase = True And Left(m_EEP01, 1) = "C" And _
         Left(CboEEP04.Text, 2) = EMP_發文歸檔 And ChkEP11.Value = 0 Then
         If MsgBox("是否紙本送件（經發文室發文）？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            bolNotSendDept = True
         End If
      End If
   End If
   
   bolSendMail = True 'Add By Sindy 2014/1/15 一般都是要寄Mail
   cmdSend.Enabled = False
   cmdExit.Enabled = False 'Add By Sindy 2018/11/7 防止人員急著操作
   
   '********************
   '依流程狀態更新資料
   '更新相關日期
   '********************
   'Add By Sindy 2013/8/22
   '送核或送判時,若核稿人休假有調整收受者時,則記錄原收受者
   'Modify By Sindy 2013/10/2 開放更多流程檢查休假可以調整收受者
   If m_EEP11Person <> "" And Left(m_EEP11Person, 5) <> Trim(Left(CboEEP05, 6)) Then
      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "原收受者:" & Left(m_EEP11Person, 5)
   End If
   '2013/10/2 END
'   If (Left(CboEEP04, 2) = EMP_送核 And Trim(m_CMMan) <> Trim(CboEEP05.Text)) Then
'      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "原收受者:" & Left(m_CMMan, 5)
'   ElseIf (Left(CboEEP04, 2) = EMP_送判 And Trim(m_CSMan) <> Trim(CboEEP05.Text)) Then
'      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "原收受者:" & Left(m_CSMan, 5)
'   'Add By Sindy 2013/9/24
'   ElseIf (Left(CboEEP04, 2) = EMP_墨完 And Trim(m_DMMan) <> Trim(CboEEP05.Text)) Then
'      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "原收受者:" & Left(m_DMMan, 5)
'   '2013/9/24 END
''         '送核則一併更新核稿人
''         If Left(CboEEP04, 2) = EMP_送核 Then
''            m_PrevForm.txt1(5) = Left(Trim(CboEEP05.Text), 6)
''         End If
''      ElseIf Left(CboEEP04.Text, 2) = EMP_轉回 And m_strLastEEP04 = EMP_送核 Then
''         m_PrevForm.txt1(5) = Left(Trim(CboEEP05.Text), 6)
'   End If
'   '2013/8/22 END
   
   'Add By Sindy 2018/8/30
   If UCase(cmdSend.Caption) = UCase("E-Mail") Then
      strSql = "" 'Add By Sindy 2021/7/21 Find Err
      cmdSend.Tag = "EMailKeepFile In" 'Add By Sindy 2021/3/31 抓程式err
      If EMailKeepFile(strUpdDate, strUpdTime) = False Then
         'Add By Sindy 2025/7/9
         Call PUB_WriteDebugLog("【frm090202_2】m_EEP01='" & m_EEP01 & "' 呼叫 EMailKeepFile=False 離開 FormSave;")
         '2025/7/9 END
         cmdSend.Enabled = True
         cmdExit.Enabled = True 'Add By Sindy 2018/11/7
         Exit Sub
      End If
      cmdSend.Tag = "EMailKeepFile Out" 'Add By Sindy 2021/3/31 抓程式err
   End If
   strSql = "EMailKeepFile => OK" 'Add By Sindy 2022/3/11 找Bug暫放
   
   'Modify By Sindy 2023/5/26
   'If Trim(strUpdDate) = "" Then
   If Val(Trim(strUpdDate)) = 0 Then
   '2023/5/26 END
      strUpdDate = strSrvDate(1)
      strUpdTime = Right("000000" & ServerTime, 6)
      '檢查是否為工作天
      If Not ChkWorkDay(strUpdDate) Then
         strUpdDate = CompWorkDay(1, strUpdDate, 0)
      End If
      strUpdDate = ChangeWStringToTString(strUpdDate) '轉換成民國日期
   End If
   '2018/8/30 END
   
   'Add By Sindy 2018/10/23 多案件更新時,要加註內容
   If (cmdManyCase.Visible = True And cmdManyCase.Enabled = True) Or txtLpNote.Tag = "多案單筆歷程" Then
      If m_RetrunRecvCnt > 0 And strRetrunRecvText <> "" Then
         If InStr(txtEEP08, strRetrunRecvText) = 0 Then
            txtEEP08 = txtEEP08 & vbCrLf & strRetrunRecvText
         End If
      End If
   End If
   '2018/10/23 END
   
   '******************************* Start ******************************************
   '********************************************************************************
   'Add By Sindy 2013/8/20
   If intReceiveKind = 0 Then '0.承辦人工作進度
      'Modify By Sindy 2017/9/15 Mark:作業會重讀資料,會讀錯筆
      'm_PrevForm.SSTab1.Tab = intTab_承辦單 '回工作進度詳細資料才可存檔
      '2017/9/15 END
      m_PrevForm.m_Flow = Left(CboEEP04.Text, 2) 'Add By Sindy 2013/10/14 欲新增的下一流程
'   Else
'      cnnConnection.BeginTrans '**************
   End If
   'Modify By Sindy 2018/6/20
   cnnConnection.BeginTrans: bolConn = True '**************
   '2013/8/20 END
   Screen.MousePointer = vbHourglass
   
   'Added by Morgan 2021/7/21
   If strPA176 <> "" And strPA176 <> pa(176) Then
      strSql = "update patent set pa176='" & strPA176 & "' where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   'end 2021/7/21
   
   'Modify By Sindy 2025/4/21 函數程式太大了,要切為2
   '********************************************************************
   '各歷程狀態更新的資料
   '********************************************************************
   Call FormSave2(strUpdDate, strUpdTime, bolSendMail, bolRegMail, strLP31)
   '********************************************************************

   '********************************************************************************
   '呼叫前畫面存檔
   '********************************************************************************
   'Modify By Sindy 2018/6/20
   If intReceiveKind = 0 Then '0.承辦人工作進度
      Call m_PrevForm.cmdok_Click(1) 'frm090201_2:存檔完成會切換Tab重讀資料了
      If m_PrevForm.m_chkcmdok1 = False Then
         GoTo ErrHand
      Else
         'Modify By Sindy 2022/10/31 P-127705 送會後,會稿日沒寫入日期,查原因
         If Left(CboEEP04.Text, 2) = EMP_送會 Then
            strSql = "select * from engineerprogress where ep02='" & m_EEP01 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If Val("" & RsTemp.Fields("ep07")) = 0 Then '送會後,會稿日空白
                  PUB_SendMail strUserNum, "97038", "", m_EEP01 & "送會後,會稿日空白！", _
                  m_EEP01 & ":" & PField(1) & PField(2) & PField(3) & PField(4) & vbCrLf & vbCrLf & _
                  "PUB_ChkEmpFlowExists(m_EEP01, EMP_送會) = False And Val(m_EP07) > 0;" & vbCrLf & vbCrLf & _
                  "m_EP07= " & m_EP07 & vbCrLf & vbCrLf & _
                  "PUB_ChkEmpFlowExists(m_EEP01, EMP_送會)= " & PUB_ChkEmpFlowExists(m_EEP01, EMP_送會) & vbCrLf & vbCrLf & _
                  "intReceiveKind= " & intReceiveKind & vbCrLf & vbCrLf & _
                  "m_PrevForm.txt1(4)= " & m_PrevForm.txt1(4) & vbCrLf & vbCrLf & _
                  "m_PrevForm.strEP07Tag= " & m_PrevForm.strEP07Tag & vbCrLf & vbCrLf & _
                  "strUpdDate= " & strUpdDate, , , , , , , , , , True, False, , , False
               Else
                  m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "會稿日=" & RsTemp.Fields("ep07")
               End If
            End If
         End If
         '2022/10/31 END
      End If
   End If
   '2018/6/20 END
   '********************************** END *****************************************
   
   'Add By Sindy 2023/10/2 + if
   'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True
   If bolPAFlow = True Or bolTMFlow = True Or bolCFTFlow = True Then
   '2023/10/2 END
      'Add By Sindy 2013/10/1 修改基本檔
      strConSql = ""
      If (txtCaseName(0).Enabled = True And txtCaseName(0).Text <> txtCaseName(0).Tag) Or _
         (txtCaseName(1).Enabled = True And txtCaseName(1).Text <> txtCaseName(1).Tag) Or _
         (txtCaseName(2).Enabled = True And txtCaseName(2).Text <> txtCaseName(2).Tag) Then
         'If InStr("5,6,7,8", m_strSys) > 0 Then '服務檔
         If m_strSys = "5" Or m_strSys = "6" Then '專利、商標服務檔
            strSql = "update servicepractice set "
            If txtCaseName(0).Enabled = True And txtCaseName(0).Text <> txtCaseName(0).Tag Then
               strConSql = strConSql & ",SP05='" & ChgSQL(txtCaseName(0).Text) & "'"
            End If
            If txtCaseName(1).Enabled = True Then
               If txtCaseName(1).Enabled = True And txtCaseName(1).Text <> txtCaseName(1).Tag Then
                  strConSql = strConSql & ",SP06='" & ChgSQL(txtCaseName(1).Text) & "'"
               End If
            End If
            If txtCaseName(2).Enabled = True Then
               If txtCaseName(2).Enabled = True And txtCaseName(2).Text <> txtCaseName(2).Tag Then
                  strConSql = strConSql & ",SP07='" & ChgSQL(txtCaseName(2).Text) & "'"
               End If
            End If
            strConSql = Mid(strConSql, 2)
            strSql = strSql & strConSql & " where SP01='" & PField(1) & "' and SP02='" & PField(2) & "' and SP03='" & PField(3) & "' and SP04='" & PField(4) & "'"
            cnnConnection.Execute strSql
         ElseIf m_strSys = "1" Then '專利檔
            strSql = "update patent set "
            If txtCaseName(0).Enabled = True And txtCaseName(0).Text <> txtCaseName(0).Tag Then
               strConSql = strConSql & ",pa05='" & ChgSQL(txtCaseName(0).Text) & "'"
            End If
            If txtCaseName(1).Enabled = True And txtCaseName(1).Text <> txtCaseName(1).Tag Then
               strConSql = strConSql & ",pa06='" & ChgSQL(txtCaseName(1).Text) & "'"
            End If
            If txtCaseName(2).Enabled = True And txtCaseName(2).Text <> txtCaseName(2).Tag Then
               strConSql = strConSql & ",pa07='" & ChgSQL(txtCaseName(2).Text) & "'"
            End If
            strConSql = Mid(strConSql, 2)
            strSql = strSql & strConSql & " where pa01='" & PField(1) & "' and pa02='" & PField(2) & "' and pa03='" & PField(3) & "' and pa04='" & PField(4) & "'"
            cnnConnection.Execute strSql
         'Add By Sindy 2023/1/9
         ElseIf m_strSys = "2" Then '商標檔
            strSql = "update trademark set "
            If txtCaseName(0).Enabled = True And txtCaseName(0).Text <> txtCaseName(0).Tag Then
               strConSql = strConSql & ",tm05='" & ChgSQL(txtCaseName(0).Text) & "'"
            End If
            strConSql = Mid(strConSql, 2)
            strSql = strSql & strConSql & " where tm01='" & PField(1) & "' and tm02='" & PField(2) & "' and tm03='" & PField(3) & "' and tm04='" & PField(4) & "'"
            cnnConnection.Execute strSql
            '2023/1/9 END
         End If
      End If
      '2013/10/1 END
   End If
   
   'Add By Sindy 2020/9/29
   'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
   If (bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True) And _
      (Left(CboEEP04.Text, 2) = EMP_送件 Or _
       Left(CboEEP04.Text, 2) = EMP_退件重送 Or _
       Left(CboEEP04.Text, 2) = EMP_發文歸檔 Or _
       Left(CboEEP04.Text, 2) = EMP_送核 Or _
       Left(CboEEP04.Text, 2) = EMP_送會 Or _
       Left(CboEEP04.Text, 2) = EMP_送判) Then
      '先清之前資料
      strSql = "update caseprogress set" & _
               " cp163=null" & _
               " where cp163='" & m_EEP01 & "' and cp158=0 and cp159=0"
      Pub_SeekTbLog strSql 'Add By Sindy 2021/12/27
      cnnConnection.Execute strSql, intI
      If txtLpNote.Tag = "多案單筆歷程" Then
         'Update現況
         strSql = "update caseprogress set" & _
                  " cp163='" & m_EEP01 & "'" & _
                  " where cp09 in('" & Replace(m_RetrunRecv, ",", "','") & "')"
         Pub_SeekTbLog strSql 'Add By Sindy 2021/12/27
         cnnConnection.Execute strSql, intI
      End If
   End If
   
   'Add By Sindy 2025/8/20
   If bolFCPFlow = True Then
      If Me.Frame945.Tag = 告知代理人 And _
         (Left(CboEEP04.Text, 2) = EMP_判發 Or Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送) Then
         If txtEED14.Text <> "" Then
            '檢查是否已新增過行事曆
            strSql = "select * from staff_calendar" & _
                     " where sc05='" & PField(1) & "' and sc06='" & PField(2) & "' and sc07='" & PField(3) & "' and sc08='" & PField(4) & "'" & _
                     " and sc01=" & DBDATE(txtEED14.Text) & " and sc04='追蹤客戶指示-" & lblCP10 & "' and sc18 is null "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
               '系統自動產生一期限【行事曆期限】
               strExc(9) = ""
               For s = 0 To 3
                  strExc(0) = GetST52SelfList(Left(m_EPMan, 5), "st5" & CStr(Val("2") + s))
                  If strExc(0) <> "" Then
                     If Left(PUB_GetST03(strExc(0)), 1) <> "M" Then
                        strExc(9) = strExc(9) & "," & strExc(0)
                     End If
                  End If
               Next s
               If strExc(9) <> "" Then strExc(9) = Mid(strExc(9), 2)
               '提醒人員: 程序, 工程師承辦人, 工程師各階主管(到簡協理Owen)
               strExc(1) = Left(m_NPMan, 5) & "," & Left(m_EPMan, 5) & IIf(strExc(9) <> "", "," & strExc(9), "")
               Call PUB_CompRepeatReple(strExc(1), "")
               strExc(1) = Replace(strExc(1), ";", ",")
               '解除人員: 程序, 工程師各階主管(到簡協理Owen)(工程師承辦人本人不能解除)
               strExc(2) = Left(m_NPMan, 5)
               'Modify By Sindy 2025/9/22 外專英文組的解除人員，加上工程師承辦人
               If Left(PUB_GetST93(Left(m_EPMan, 5)), 1) = "F" Then
                  strExc(2) = strExc(2) & "," & Left(m_EPMan, 5)
               End If
               '2025/9/22 END
               strExc(2) = strExc(2) & IIf(strExc(9) <> "", "," & strExc(9), "")
               Call PUB_CompRepeatReple(strExc(2), "")
               strExc(2) = Replace(strExc(2), ";", ",")
               If PUB_AddFCPStaffCalendar(DBDATE(txtEED14.Text), 1, strExc(1), "追蹤客戶指示-" & lblCP10, _
                  strExc(2), "1", PField(1), PField(2), PField(3), PField(4)) Then
               End If
            End If
         End If
      End If
   End If
   '2025/8/20 END
   
   '******************************
   '      承辦電子簽核流程檔
   '******************************
   '更新上一筆流程的待回覆＝null
   'Modify By Sindy 2013/9/5 更新上一筆之前全部流程的待回覆＝null,因發生同時新增二筆送核狀況,若只更新上一筆,則就會有殘留待回覆Y的狀況
   'Modify By Sindy 2016/3/7 取消 Left(CboEEP04.Text, 2) = EMP_退件重送
   '                         +EMP_圖修,EMP_圖完
   If Left(CboEEP04.Text, 2) = EMP_核修 Or _
      Left(CboEEP04.Text, 2) = EMP_核完 Or _
      Left(CboEEP04.Text, 2) = EMP_會修 Or _
      Left(CboEEP04.Text, 2) = EMP_會完 Or _
      Left(CboEEP04.Text, 2) = EMP_退回 Or _
      Left(CboEEP04.Text, 2) = EMP_繪圖判發 Or _
      Left(CboEEP04.Text, 2) = EMP_判發 Or _
      Left(CboEEP04.Text, 2) = EMP_轉回 Or _
      Left(CboEEP04.Text, 2) = EMP_草修 Or _
      Left(CboEEP04.Text, 2) = EMP_草核完 Or _
      Left(CboEEP04.Text, 2) = EMP_圖修 Or _
      Left(CboEEP04.Text, 2) = EMP_圖完 Or _
      Left(CboEEP04.Text, 2) = EMP_送排版 Or _
      Left(CboEEP04.Text, 2) = EMP_排版完成 Or _
      Left(CboEEP04.Text, 2) = EMP_核稿分案 Or _
      Left(CboEEP04.Text, 2) = EMP_轉檔完成 Or _
      Left(CboEEP04.Text, 2) = EMP_送核稿分案 Or _
      Left(CboEEP04.Text, 2) = EMP_送件 Or _
      Left(CboEEP04.Text, 2) = EMP_程序送判 Or _
      Left(CboEEP04.Text, 2) = EMP_程序退回 Or _
      (Left(CboEEP04.Text, 2) = EMP_發文歸檔 And bolFCPFlow = True) Then
      'Modify By Sindy 2016/3/7 會有待回覆流程一筆以上,增加判斷(and EEP05='" & m_FlowUserNum & "')
      strSql = "update empelectronprocess set" & _
               " EEP09=null" & _
               " where eep01='" & m_EEP01 & "'"
      If m_CurrFlowEEP02 > 0 Then
         strSql = strSql & " and EEP02=" & m_CurrFlowEEP02
      Else
         strSql = strSql & " and EEP05='" & m_FlowUserNum & "'"
         strSql = strSql & " and eep02<=" & intLastEEP02
      End If
      cnnConnection.Execute strSql, intI
   End If
   
   '****************************************************************
   'Add By Sindy 2018/8/2 為配合專利處核判主管不收E-Mail的需求
   '****************************************************************
   '由待核判區新增非聯絡歷程時,將前面的顯示歷程改為不顯示
   If UCase(m_PrevForm.Name) = UCase("frm090202_1") And Left(CboEEP04.Text, 2) <> EMP_聯絡 Then
      'Modify By Sindy 2018/10/26 + 資料應該是更新聯絡即可 ex:P-121203(24)
      '                             需要增加判斷收受者是自己的才能改為不顯示
'      strSql = "update empelectronprocess set EEP13=null" & _
'               " where eep01='" & m_EEP01 & "' and eep13 is not null"

      'Modify By Sindy 2023/4/13 Mark: Trigger統一更新 EMPELECTRONPROCESS_AFTER
'      strSql = "update empelectronprocess set EEP13=null" & _
'               " where eep01='" & m_EEP01 & "' and eep13 is not null and EEP04='" & EMP_聯絡 & "'" & _
'               " and EEP05='" & m_FlowUserNum & "'"
'      cnnConnection.Execute strSql
   End If
   '2018/8/2 END
   '****************************************************************
   
   'Add By Sindy 2020/10/12
   If txtLpNote.Tag = "多案單筆歷程" Then
      '中途單案變多筆歷程會發生的狀況
      strSql = "update empelectronprocess set EEP13=null,EEP11=EEP11||decode(EEP11,null,'',';')||'中途單案變多筆(" & m_EEP01 & ")'" & _
               " where eep01 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and eep13 is not null" & _
               " and EEP04<>'" & EMP_聯絡 & "'" & _
               " and EEP05='" & m_FlowUserNum & "'"
      cnnConnection.Execute strSql
   End If
   '2020/10/12 END
   
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
   m_EEP02 = intMaxEEP02 + 1
   '是否需等待回覆
'   If Left(CboEEP04.Text, 2) = EMP_送英核 Or _
'      Left(CboEEP04.Text, 2) = EMP_送核 Or _
'      Left(CboEEP04.Text, 2) = EMP_送會 Or _
'      Left(CboEEP04.Text, 2) = EMP_草核 Or _
'      Left(CboEEP04.Text, 2) = EMP_墨完 Or _
'      Left(CboEEP04.Text, 2) = EMP_送判 Or _
'      Left(CboEEP04.Text, 2) = EMP_轉回 Then
   'Modify By Sindy 2023/12/18 排除外專的分割,209,235排版完成
   If InStr(EMP_需等待回覆的狀態, Left(CboEEP04.Text, 2)) > 0 _
      And Not (bolFCPFlow = True _
               And (cp(10) = "307" Or cp(10) = "209" Or cp(10) = "235") _
               And m_EPMan <> "" _
               And Left(CboEEP04.Text, 2) = EMP_排版完成) Then
      strUpdEEP09 = "Y" '必須等待回覆
   Else
      strUpdEEP09 = ""
   End If
   If Left(CboEEP04.Text, 2) = EMP_轉回 Then
      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "轉回前流程狀態:" & m_strLastEEP04
   End If
   'Add By Sindy 2017/7/31 記錄處理的流程狀態
   If m_strLastEEP04 <> "" And InStr(m_UpdEEP11, "流程狀態:" & m_strLastEEP04) = 0 Then
      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "流程狀態:" & m_strLastEEP04
   End If
   '2017/7/31 END
   'Add By Sindy 2023/10/4
   If Left(CboEEP04.Text, 2) = EMP_交辦 Then
      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "原收受者:" & strLastEEP05
   End If
   '2023/10/4 END
   'Modify By Sindy 2013/9/24 附加流程的收受者為空白
   'Modify By Sindy 2013/10/28 附加流程要存收受者 CNULL(IIf(Left(CboEEP04.Text, 2) = EMP_附加流程, "", Trim(Left(CboEEP05.Text, 6)))) & "," &
'   strChkEEP13_EEP01 = m_EEP01 'Add By Sindy 2016/8/5
   'Modify By Sindy 2018/8/29 + eep14:會稿方式
   'Modify By Sindy 2018/10/16 + eep15:多案總收文號
   'Modify By Sindy 2020/6/17 CNULL(IIf(cmdManyCase.Visible = True And cmdManyCase.Enabled = True And m_RetrunRecvCnt > 0, m_RetrunRecv, ""))
   ' => CNULL(m_RetrunRecv)
   'Add By Sindy 2020/10/8
   If txtLpNote.Tag = "多案單筆歷程" Then
      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "多案單筆歷程"
   End If
   '2020/10/8 END
   'Add By Sindy 2022/4/27
   If m_strSpecState <> "" Then
      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & m_strSpecState
   End If
   '2022/4/27 END
   'Add By Sindy 2024/4/15 記錄上一歷程順序
   If InStr(EMP_需等待回覆的狀態, m_strLastEEP04) > 0 And m_strLastEEP04 <> "" Then
      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "上一歷程順序:" & CStr(Format(intLastEEP02, "000"))
   End If
   '2024/4/15 END
   'Modify By Sindy 2023/11/17 轉檔完成此道歷程收受者直接都掛工程師(m_EPMan=承辦人); 因畫面上會因人員勾選轉檔後送件（程序發文）而顯示程序人員
   'Modify By Sindy 2023/12/18 +,eep16
   strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep09,eep10,eep11,eep12,eep14,eep15,eep16) values(" & _
            CNULL(m_EEP01) & "," & m_EEP02 & "," & CNULL(Trim(txtEEP03)) & "," & _
            CNULL(Left(CboEEP04.Text, 2)) & "," & _
            IIf(Left(CboEEP04.Text, 2) = EMP_轉檔完成, CNULL(Left(m_EPMan, 5)), CNULL(Trim(Left(CboEEP05.Text, 6)))) & "," & _
            strSrvDate(1) & "," & strUpdTime & "," & CNULL(ChgSQL(txtEEP08)) & "," & CNULL(strUpdEEP09) & "," & _
            CNULL(txtEEP10) & "," & CNULL(m_UpdEEP11) & "," & CNULL(m_EEP12) & "," & CNULL(m_EEP14) & "," & _
            CNULL(m_RetrunRecv) & "," & CNULL(m_EEP16) & ")"
   cnnConnection.Execute strSql
   '***** 抓Bug *****
   'P-121148  朱桓毅:
   '在16:22時點選判發程序後，我的信箱突然收到一封郵件，其中寄件人與收件人都是我，
   '而程序的名稱卻是繪圖判發，因此，而我在進系統確認流程時，多出一道16:22的繪圖判發程序，
   '認為可能是程式的問題，故想請您幫我確認一下。16:25的判發程序是我重新再點選判發才產生的。
   '*****************
   'Add By Sindy 2018/10/24
   If Trim(txtEEP03) <> strUserNum Then
      'Add By Sindy 2025/7/9
      Call PUB_WriteDebugLog("【frm090202_2】Trim(txtEEP03) <> strUserNum: cmdCancel.Visible=" & cmdCancel.Visible & " And cmdExit.Enabled=" & cmdExit.Enabled)
      '2025/7/9 END
      If intReceiveKind = 0 Then '0.承辦人工作進度
         PUB_SendMail strUserNum, "97038", "", UCase(m_PrevForm.Name) & " 新增歷程主檔有問題:Trim(txtEEP03)[" & txtEEP03 & "] <> strUserNum[" & strUserNum & "]!! intReceiveKind=" & intReceiveKind, _
            "(工)總收文號=" & m_PrevForm.lbl1(3).Caption & vbCrLf & _
            "(工)本所案號=" & m_PrevForm.lbl1(7).Caption & vbCrLf & _
            "(歷)總收文號=" & lblCP09.Caption & vbCrLf & _
            "(歷)本所案號=" & lblCaseNo.Caption & vbCrLf & _
            " 流 程 狀 態=" & CboEEP04.Text & vbCrLf & "新增歷程主檔有問題！" & vbCrLf & "Err Text:" & Err.Number & Err.Description & vbCrLf & strSql & _
            vbCrLf & vbCrLf & _
            "cmdCancel.Visible = " & IIf(cmdCancel.Visible = False, "False", "True") & _
            " ; cmdExit.Enabled = " & IIf(cmdExit.Enabled = True, "True", "False") & _
            " ; cmdAdd.Visible = " & IIf(cmdAdd.Visible = True, "True", "False"), , , , , , , , , , True, False, , , False
      Else
         PUB_SendMail strUserNum, "97038", "", "新增歷程主檔有問題Trim(txtEEP03)[" & txtEEP03 & "] <> strUserNum[" & strUserNum & "]!! intReceiveKind=" & intReceiveKind, _
            "(歷)總收文號=" & lblCP09.Caption & vbCrLf & _
            "(歷)本所案號=" & lblCaseNo.Caption & vbCrLf & _
            "  發  送  者=" & txtEEP03 & " " & txtEEP03_2 & vbCrLf & _
            " 流 程 狀 態=" & CboEEP04.Text & vbCrLf & "新增歷程主檔有問題！" & vbCrLf & "Err Text:" & Err.Number & Err.Description & vbCrLf & strSql, , , , , , , , , , True, False, , , False
      End If
      MsgBox "新增歷程主檔有問題，請洽電腦中心！" & vbCrLf & strSql, vbExclamation
      GoTo ErrHand
   End If
   '2018/10/24 END
   '***** 抓Bug END *****
   
   'Add By Sindy 2018/8/30 更新寄件備份的歷程序號
   If UCase(cmdSend.Caption) = UCase("E-Mail") Then
      strSql = "update smailbackup set smb11=" & m_EEP02 & _
               " where smb01='" & m_EEP01 & "' and smb02=" & DBDATE(strUpdDate) & " and smb03=" & strUpdTime
      cnnConnection.Execute strSql
      'Add By Sindy 2025/7/9
      Call PUB_WriteDebugLog("【frm090202_2】UCase(cmdSend.Caption) = UCase(E-Mail): strSql=" & strSql)
      '2025/7/9 END
   End If
   'Add By Sindy 2018/9/27
   'Modify By Sindy 2020/9/29 And txtLpNote.Tag <> "多案單筆歷程"
   If cmdManyCase.Visible = True And cmdManyCase.Enabled = True And _
      (txtLpNote.Tag <> "多案單筆歷程" Or bolManyCaseToMix = True) Then
      '有回傳總收文號
      'If m_RetrunRecv <> "" And m_RetrunRecv <> m_EEP01 Then
      If m_RetrunRecvCnt > 0 Then
         If ManyCaseSaveData(m_EEP02, strUpdDate, strUpdTime) = False Then
            GoTo ErrHand
         End If
      End If
   End If
   '2018/9/27 END
   '2018/8/30 END
   
   '******************************
   '      承辦電子簽核附件檔
   '******************************
   If Left(CboEEP04.Text, 2) <> EMP_附加流程 Then
      If bolMoveFile = True Then
         If MoveAndDelFile(m_EEP01, CInt(m_EEP02), Left(CboEEP04.Text, 2)) = False Then
            GoTo ErrHand
         End If
      Else
         If SaveAttFile(m_EEP01, CInt(m_EEP02), 0) = False Then
            GoTo ErrHand
         End If
         
         '************************************************************************
         'Add By Sindy 2018/7/13 歸檔
         'Modify By Sindy 2023/11/23 mark,外專也會使用
         'If bolTMFlow = True Then
            If Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
               If FilingFilePDF = False Then
                  GoTo ErrHand
               End If
               
               'Add By Sindy 2020/9/16
               'Modify By Sindy 2023/11/23 +And bolTMFlow = True
               If strSrvDate(1) >= T商標電子化第2階段啟用日 And bolTMFlow = True Then
                  '更新此文號電子檔齊備日
                  PUB_UpdateLP03 m_EEP01
               End If
            End If
         'End If
         '2018/4/27 END
         '************************************************************************
         
         If bolDeleteFile = True Then '將上一筆流程的附件刪除
            'Add By Sindy 2018/10/8 取得 intLastEEP02 此序號有幾筆附件
            strSql = "select eef03 From empelectronfile where EEF01='" & m_EEP01 & "'" & _
                       " and EEF02=" & intLastEEP02
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               intFileCnt = RsTemp.RecordCount '附件數
            End If
            '2018/10/8 END
            
            'Modify By Sindy 2013/9/11 檢查上一筆附件區裡是否有繪圖的原始檔,若有不可刪除
            'Modify By Sindy 2014/3/11 +dwg.7z
            strSql = "select eef03 From empelectronfile where EEF01='" & m_EEP01 & "'" & _
                       " and EEF02=" & intLastEEP02 & " and (substr(upper(eef03),-4)='.DWG' or substr(upper(eef03),-7)='DWG.ZIP' or substr(upper(eef03),-6)='DWG.7Z')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then '無繪圖檔才可刪除
            '2013/9/11 END
               PUB_DelFtpFile2 m_EEP01, " and EEF02=" & intLastEEP02, "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
               'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
               strSql = "delete from empelectronfile" & _
                        " where EEF01='" & m_EEP01 & "'" & _
                          " and EEF02=" & intLastEEP02
               cnnConnection.Execute strSql
            Else
               '只留繪圖的原始檔
               'Modify By Sindy 2014/3/11 +dwg.7z
               PUB_DelFtpFile2 m_EEP01, " and EEF02=" & intLastEEP02 & " and substr(upper(eef03),-4)<>'.DWG' and substr(upper(eef03),-7)<>'DWG.ZIP' and substr(upper(eef03),-6)<>'DWG.7Z'", "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
               'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
               strSql = "delete from empelectronfile" & _
                        " where EEF01='" & m_EEP01 & "'" & _
                          " and EEF02=" & intLastEEP02 & _
                          " and substr(upper(eef03),-4)<>'.DWG' and substr(upper(eef03),-7)<>'DWG.ZIP' and substr(upper(eef03),-6)<>'DWG.7Z'"
               cnnConnection.Execute strSql
            End If
         End If
      End If
   End If
   
   'Add By Sindy 2013/11/6 沿用附件
   'Modify By Sindy 2018/10/8 bolDeleteFile = True 增加記錄附件數
   If bolMoveFile = True Or bolDeleteFile = True Then
      strSql = "update empelectronprocess" & _
               " set eep11=eep11||decode(eep11,null,'',',')||'沿用附件" & IIf(bolMoveFile = True, "(流程順序" & m_PreviousFlow & ")", "") & IIf(bolDeleteFile = True, "(流程順序" & intLastEEP02 & "附件數" & intFileCnt & ")", "") & "CP118=" & cp(118) & "'" & _
               " where eep01='" & m_EEP01 & "'" & _
               " and eep02=" & m_EEP02
      cnnConnection.Execute strSql
   End If
   '2013/11/6 END
   
   '******************************************************************************************
   strCP06 = "": strCP07 = "": strCP27 = "": strCP48 = ""
   'Add By Sindy 2025/4/10 增加FCP-電話聯絡單需收文告代
   If Left(CboEEP04.Text, 2) = EMP_發文歸檔 And ChkEED08.Visible = True And ChkEED08.Value = 1 Then
      strCP14 = Left(Trim(m_EPMan), 5) '承辦人=工程師
      strCP10n = "901" '告知代理人
      strCP48 = Pub_GetHandleDay(PField(1), pa(9), strCP10n)
      strCP06 = PUB_GetFCPOurDeadline(DBDATE(strCP48), , , , "N") '以承辦期限＋5個工作天為本所期限
      '取得總收文號
      strCP09 = AutoNo("B", 6)
      strGetCP13 = ShowCurrCP13(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)), m_Country, strGetCP12)
      '系統自動收文
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
               "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP22,CP26,CP32,CP43,CP44,CP45,CP110,cp116," & _
               "CP27,CP48) VALUES " & _
               "('" & PField(1) & "','" & PField(2) & "','" & PField(3) & "','" & PField(4) & "'," & _
               strSrvDate(1) & "," & CNULL(strCP06) & "," & CNULL(strCP07) & _
               ",'" & strCP09 & "','" & strCP10n & "','90'," & CNULL(strGetCP12) & "," & CNULL(strGetCP13) & _
               ",'" & strCP14 & "','N'," & CNULL(cp(22)) & ",'N','N','" & m_EEP01 & _
               "'," & CNULL(cp(44)) & "," & CNULL(cp(45)) & "," & CNULL(cp(110)) & "," & CNULL(cp(116)) & _
               "," & CNULL(strCP27) & "," & CNULL(strCP48) & ")"
      cnnConnection.Execute strSql
   End If
   '2025/4/10 END
   'Add By Sindy 2013/9/23 檢查附加流程時,系統自動收文及自行判發或送判
   If Left(CboEEP04.Text, 2) = EMP_附加流程 Then
      strCP10n = Trim(Left(CboCP10.Text, 4))
      'Modify By Sindy 2025/7/31
      'If (Trim(Left(CboCP10.Text, 4)) = 延期 And bolPAFlow = True)
      If Trim(Left(CboCP10.Text, 4)) = 延期 Or _
         Trim(Mid(CboCP10.Text, 4)) = "延期" Then
      '2025/7/31 END
         'strCP14 = Left(m_NPMan, 5) '承辦人=程序
         strCP14 = m_FlowUserNum '承辦人=工程師
         strCP06 = PUB_GetWorkDay1(cp(6), True)
         strCP07 = cp(7)
      '936.回覆委任代理人 957.詢問代理人 958.代理人撰稿
      'Add By Sindy 2019/7/9 T代理人撰稿
      'Modify By Sindy 2025/7/31 +Or bolFCPFlow = True
      ElseIf ((Trim(Left(CboCP10.Text, 4)) = "936" Or Trim(Left(CboCP10.Text, 4)) = "957" Or Trim(Left(CboCP10.Text, 4)) = "958") _
               And (bolPAFlow = True Or bolFCPFlow = True)) Or _
         (Trim(Left(CboCP10.Text, 4)) = "734" And bolTMFlow = True) Then
         strCP14 = m_FlowUserNum '承辦人=工程師
         'Add By Sindy 2019/7/9 T代理人撰稿
         If (Trim(Left(CboCP10.Text, 4)) = "734" And bolTMFlow = True) Then
            strCP27 = strSrvDate(1)
         End If
         '2019/7/9 END
         'Add By Sindy 2020/3/9 936.回覆委任代理人 957.詢問代理人,本所期限預設3個工作天
         If Trim(Left(CboCP10.Text, 4)) = "936" Or _
            Trim(Left(CboCP10.Text, 4)) = "957" Then
            '本所期限預設3個工作天
            'Modify By Sindy 2022/4/6
            'strCP06 = PUB_GetWorkDay1(CompDate(2, 3, strSrvDate(1)), True)
            strCP06 = CompWorkDay(3, CompDate(2, 1, strSrvDate(1)), 0) 'CompWorkDay(4, strSrvDate(1))
            '2022/4/6 END
         End If
         '2020/3/9 END
      End If
      '取得總收文號
      strCP09 = AutoNo("B", 6)
      'Add By Sindy 2018/12/3
      strGetCP13 = ShowCurrCP13(CStr(PField(1)), CStr(PField(2)), CStr(PField(3)), CStr(PField(4)), m_Country, strGetCP12)
      '2018/12/3 END
      '系統自動收文
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
               "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP22,CP26,CP32,CP43,CP44,CP45,CP110,cp116," & _
               "CP27) VALUES " & _
               "('" & PField(1) & "','" & PField(2) & "','" & PField(3) & "','" & PField(4) & "'," & _
               strSrvDate(1) & "," & CNULL(strCP06) & "," & CNULL(strCP07) & _
               ",'" & strCP09 & "','" & strCP10n & "','90'," & CNULL(strGetCP12) & "," & CNULL(strGetCP13) & _
               ",'" & strCP14 & "','N'," & CNULL(cp(22)) & ",'N','N','" & m_EEP01 & _
               "'," & CNULL(cp(44)) & "," & CNULL(cp(45)) & "," & CNULL(cp(110)) & "," & CNULL(cp(116)) & _
               "," & CNULL(strCP27) & ")"
      cnnConnection.Execute strSql
      'Add By Sindy 2024/11/20
      '外專:依操作的案件性質檢查是否屬於有呈送主管機關(不管是否為經濟部智慧財產局)，則"電子送件"欄位，請自動上"Y"，以防人員當紙本送件
      If PUB_ChkhadCF10forEMP_46(PField(1), m_Country, Trim(Left(CboCP10.Text, 4))) = 1 _
         And m_Country = "000" And bolFCPFlow = True Then
         strSql = "update caseprogress set CP118 ='Y' where cp09='" & strCP09 & "'"
         cnnConnection.Execute strSql
      End If
      '2024/11/20 END
      '附加流程
      'Modify By Sindy 2023/12/18 +,eep16
      strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep09,eep10,eep11,eep12,eep16) values(" & _
               CNULL(strCP09) & ",1," & CNULL(Trim(txtEEP03)) & "," & _
               CNULL(Left(CboEEP04.Text, 2)) & "," & CNULL(Trim(Left(CboEEP05.Text, 6))) & "," & _
               strSrvDate(1) & "," & strUpdTime & "," & CNULL(ChgSQL(txtEEP08)) & "," & CNULL(strUpdEEP09) & "," & _
               CNULL(txtEEP10) & ",'" & m_UpdEEP11 & "','" & m_EEP12 & "','" & m_EEP16 & "')"
      cnnConnection.Execute strSql
      'Add By Sindy 2024/10/30 記錄此附加流程所產生的文號;當刪除此文號時可以一併回頭刪除此附加流程歷程
      strSql = "update empelectronprocess set eep11=eep11||decode(eep11,null,'',';')||'" & strCP09 & ";'" & _
               " where eep01='" & m_EEP01 & "' and eep02='" & m_EEP02 & "'"
      cnnConnection.Execute strSql
      '2024/10/30 END
      
      'Add By Sindy 2014/1/15 雅娟:有關936.回覆委任代理人957.詢問代理人之簽辦流程,目前是鎖定直接送判,但有仍會有需要跑其他流程的情況,故請不要鎖流程
      'Modify By Sindy 2019/5/7 + 958.代理人撰稿
      'Add By Sindy 2025/7/31 +Or bolFCPFlow = True
      If (Trim(Left(CboCP10.Text, 4)) = "936" Or _
          Trim(Left(CboCP10.Text, 4)) = "957" Or _
          Trim(Left(CboCP10.Text, 4)) = "958") And (bolPAFlow = True Or bolFCPFlow = True) Then
         bolSendMail = False '不需要寄Mail
         If lstAtt(0).ListCount > 0 Then
            If SaveAttFile(strCP09, 1, 0) = False Then
               GoTo ErrHand
            End If
         End If
         '上收卷日
         strSql = "update EngineerProgress" & _
                  " set ep27=" & strSrvDate(1) & _
                  " where ep02='" & strCP09 & "'"
         cnnConnection.Execute strSql
      Else
      '2014/1/15 END
'         'Add By Sindy 2014/9/16 抓附加流程的判發人
'         strPP05 = m_CSMan
'         strSql = "select * from promoterproofreader where pp01='" & PField(1) & "' and pp02='" & m_FlowUserNum & "' and pp03='" & Trim(Left(CboCP10.Text, 4)) & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 1 Then
'            strPP05 = RsTemp.Fields("pp05")
'         End If
'         '2014/9/16 END
'         '自行判發
'         'If m_CSMan = "" Or Left(m_CSMan, 5) = m_FlowUserNum Then
'         If strPP05 = "" Or strPP05 = m_FlowUserNum Then
'            strAutoFlow = "判發"
'            strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep12) values(" & _
'                     CNULL(strCP09) & ",2," & CNULL(Trim(txtEEP03)) & "," & _
'                     CNULL(EMP_判發) & "," & CNULL(Left(m_NPMan, 5)) & "," & strSrvDate(1) & "," & _
'                     strUpdTime & ",'" & m_EEP12 & "')"
'            cnnConnection.Execute strSql
'         '送判
'         Else
'            strAutoFlow = "送判"
'            'Modify By Sindy 2014/9/16 CNULL(Left(m_CSMan, 5))==>CNULL(strPP05)
'            strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep09,eep12) values(" & _
'                     CNULL(strCP09) & ",2," & CNULL(Trim(txtEEP03)) & "," & _
'                     CNULL(EMP_送判) & "," & CNULL(strPP05) & "," & strSrvDate(1) & "," & _
'                     strUpdTime & ",'Y','" & m_EEP12 & "')"
'            cnnConnection.Execute strSql
'            'Add By Sindy 2013/11/28 更新判發人
'            'Modify By Sindy 2014/9/16 Left(m_CSMan, 5)==>strPP05
'            strSql = "update engineerprogress set ep40='" & strPP05 & "' where ep02='" & strCP09 & "'"
'            cnnConnection.Execute strSql
'            '2013/11/28 END
'         End If
         
         'Add By Sindy 2018/11/1 竹平:FCT-042938須核稿
         If bolTMFlow = True And Trim(Left(CboCP10.Text, 4)) = "734" Then
            'Add By Sindy 2019/7/9 T代理人撰稿,無需核判直接上發文
            'If Trim(Left(CboCP10.Text, 4)) = "734" Then
               '加掛代理人撰稿的下一程序305.催審期限，期限為原進度的法定期限減3個工作日(不含法定當天)；
               '本所期限＝法定期限，管制承辦人
               '法限=系統日+1年,所限=法限
               strExc(9) = CompWorkDay(4, cp(7), 1)
               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                        "VALUES('" & strCP09 & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','305','" & strExc(9) & "','" & strExc(9) & "','" & strCP14 & "'," & GetNextProgressNo & ") "
               cnnConnection.Execute strSql, intI
            'End If
         'Modify By Sindy 2024/3/1 + Or bolFCPFlow = True
         ElseIf bolTMFlow = True Or bolFCPFlow = True Then
            '有核稿主管,並且不可自行核稿者
            If m_CMMan <> "" And Left(m_CMMan, 5) <> m_FlowUserNum And Left(m_CMMan, 5) <> Left(m_CSMan, 5) Then
               strAutoFlow = "送核"
               'Modify By Sindy 2023/12/18 +,eep16
               strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep09,eep12,eep16) values(" & _
                        CNULL(strCP09) & ",2," & CNULL(Trim(txtEEP03)) & "," & _
                        CNULL(EMP_送核) & "," & CNULL(Left(m_CMMan, 5)) & "," & strSrvDate(1) & "," & _
                        strUpdTime & ",'Y','" & m_EEP12 & "','" & m_EEP16 & "')"
               cnnConnection.Execute strSql
            '自行判發
            ElseIf m_CSMan = "" Or Left(m_CSMan, 5) = m_FlowUserNum Then
               bolSendMail = False '不需要寄Mail Add By Sindy 2018/10/1 ex:T-211776
               strAutoFlow = "送件"
               'Modify By Sindy 2023/12/18 +,eep16
               strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep12,eep16) values(" & _
                        CNULL(strCP09) & ",2," & CNULL(Trim(txtEEP03)) & "," & _
                        CNULL(EMP_送件) & "," & CNULL(Left(m_NPMan, 5)) & "," & strSrvDate(1) & "," & _
                        strUpdTime & ",'" & m_EEP12 & "','" & m_EEP16 & "')"
               cnnConnection.Execute strSql
            '送判
            Else
               strAutoFlow = "送判"
               'Modify By Sindy 2023/12/18 +,eep16
               strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep09,eep12,eep16) values(" & _
                        CNULL(strCP09) & ",2," & CNULL(Trim(txtEEP03)) & "," & _
                        CNULL(EMP_送判) & "," & CNULL(Left(m_CSMan, 5)) & "," & strSrvDate(1) & "," & _
                        strUpdTime & ",'Y','" & m_EEP12 & "','" & m_EEP16 & "')"
               cnnConnection.Execute strSql
            End If
         '2018/11/1 END
         ElseIf bolPAFlow = True Then '專利處
            '自行判發
            If m_CSMan = "" Or Left(m_CSMan, 5) = m_FlowUserNum Then
               strAutoFlow = "判發"
   '            strChkEEP13_EEP01 = strCP09 'Add By Sindy 2016/8/5
               'Modify By Sindy 2023/12/18 +,eep16
               strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep12,eep16) values(" & _
                        CNULL(strCP09) & ",2," & CNULL(Trim(txtEEP03)) & "," & _
                        CNULL(EMP_判發) & "," & CNULL(Left(m_NPMan, 5)) & "," & strSrvDate(1) & "," & _
                        strUpdTime & ",'" & m_EEP12 & "','" & m_EEP16 & "')"
               cnnConnection.Execute strSql
            '送判
            Else
               strAutoFlow = "送判"
   '            strChkEEP13_EEP01 = strCP09 'Add By Sindy 2016/8/5
               'Modify By Sindy 2023/12/18 +,eep16
               strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep09,eep12,eep16) values(" & _
                        CNULL(strCP09) & ",2," & CNULL(Trim(txtEEP03)) & "," & _
                        CNULL(EMP_送判) & "," & CNULL(Left(m_CSMan, 5)) & "," & strSrvDate(1) & "," & _
                        strUpdTime & ",'Y','" & m_EEP12 & "','" & m_EEP16 & "')"
               cnnConnection.Execute strSql
            End If
         End If
         If SaveAttFile(strCP09, 2, 0) = False Then
            GoTo ErrHand
         End If
         If strAutoFlow <> "" Then
            'Add By Sindy 2013/11/28 更新判發人
            'Modify By Sindy 2015/3/3 更新核稿人
            strSql = "update engineerprogress set ep40=" & CNULL(Left(m_CSMan, 5)) & ",ep04=" & CNULL(Left(m_CMMan, 5)) & " where ep02='" & strCP09 & "'"
            cnnConnection.Execute strSql
            '2013/11/28 END
            '上收卷日
            'Modify By Sindy 2013/9/27 書慈提要上會稿日，會稿完成日，完稿日
            'Modify By Sindy 2015/3/3 更新是否會稿，核稿完成日
            strSql = "update EngineerProgress" & _
                     " set ep27=" & strSrvDate(1) & _
                         ",ep09=" & strSrvDate(1) & _
                     " where ep02='" & strCP09 & "'"
            cnnConnection.Execute strSql
            If strAutoFlow <> "送核" Then 'Add By Sindy 2018/11/1 + if
               strSql = "update EngineerProgress" & _
                        " set ep07=" & strSrvDate(1) & _
                            ",ep08=" & strSrvDate(1) & _
                            ",ep34='N'" & _
                            ",ep39=" & strSrvDate(1) & _
                        " where ep02='" & strCP09 & "'"
               cnnConnection.Execute strSql
            End If
            'Add By Sindy 2018/9/19
            strSql = "update EngineerProgress" & _
                     " set ep06=" & strSrvDate(1) & _
                     " where ep02='" & strCP09 & "' and (EP06 is null or EP06=0)"
            cnnConnection.Execute strSql
            '2018/9/19 END
         End If
      End If
      
      'Modify By Sindy 2013/12/5 Mark 因P-107111回代需要doc
'      '附加流程判發時刪除DOC文件
'      If strAutoFlow = "判發" Then
'         strSql = "delete from empelectronfile where EEF01='" & strCP09 & "' and EEF02=2 and substr(upper(eef03),-4)='.DOC'"
'         cnnConnection.Execute strSql
'      End If
   
   'Add by Sindy 2023/10/4 轉檔後送件
   ElseIf Left(CboEEP04.Text, 2) = EMP_轉檔完成 And ChkEED13.Value = 1 Then
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
      '新增送件歷程
      'Modify By Sindy 2023/12/18 +,eep16
      'Modify By Sindy 2024/4/15 記錄上一歷程順序
      strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep11,eep12,eep16) values(" & _
               CNULL(m_EEP01) & "," & intMaxEEP02 + 1 & "," & CNULL(Trim(txtEEP03)) & "," & _
               CNULL(EMP_送件) & "," & CNULL(Left(m_NPMan, 5)) & "," & strSrvDate(1) & "," & _
               strUpdTime & ",null,'" & "上一歷程順序:" & CStr(Format(intMaxEEP02, "000")) & "','" & m_EEP12 & "','" & m_EEP16 & "')"
      cnnConnection.Execute strSql
   '2023/10/4 END
   
   '繪圖人員處理外專案件時,要輸工作時數
   ElseIf (Left(CboEEP04.Text, 2) = EMP_判發 Or Left(CboEEP04.Text, 2) = EMP_繪圖判發) _
      And (bolFCPFlow = True Or (bolFMP = True And bolOurFMP = True)) _
      And intReceiveKind = 3 Then
RunInput:
      strCP113 = InputBox("請輸入工作時數？ (請輸入數字)")
      If strCP113 = "" Then
         MsgBox "工作時數不可空白！", vbExclamation
         GoTo RunInput
      Else
         If Not IsNumeric(strCP113) Then
            MsgBox "請輸入數字！", vbExclamation
            GoTo RunInput
         End If
      End If
      strSql = "update caseprogress set" & _
               " cp113=" & CNULL(CStr(strCP113), True) & _
               " where cp09='" & m_EEP01 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2013/9/23 END
   '******************************************************************************************
      
   'Add By Sindy 2013/9/18 P案判發時,檢查是否有CFP多國計件案,若有,則自動新增一道聯絡及附件
   'Modify By Sindy 2020/1/14 P案"送判"及"退件重送"時,不論多少次,均由系統自動將檔案用"聯絡"寄給CFP工程師。
   '    + Left(CboEEP04.Text, 2) = EMP_退件重送
   'Modify By Sindy 2020/8/12 + 雅娟:惟若是只有辦CFP案,且是在不同工程師承辦的狀況,則就不會通知,目前就有一案放了兩週不知道可以承辦,故請增加CFP案也能跟P案一樣的控管
   '                            PField(1) = "P" => (PField(1) = "P" Or PField(1) = "CFP")
   If (PField(1) = "P" Or PField(1) = "CFP") And _
      InStr(NewCasePtyList, cp(10)) > 0 And _
      (Left(CboEEP04.Text, 2) = EMP_送判 Or Left(CboEEP04.Text, 2) = EMP_判發 Or Left(CboEEP04.Text, 2) = EMP_退件重送) Then
      ' (Left(CboEEP04.Text, 2) = EMP_判發 And m_CSMan <> "" And Left(m_CSMan, 5) <> Left(m_EPMan, 5))) Then 'Modify By Sindy 2014/5/9 P-108016
      If PField(1) = "P" Then
         '新申請案:NewCasePtyList
         'Modify By Sindy 2013/10/3 增加判斷承辦人不是工程師自己
         'Modified by Morgan 2025/4/17 排除日本/德國的發明/新型且承辦人為F編號者(會自動收文209檢視中說並於國內案發文時自動上齊備並管制本所期限)--品薇
         strSql = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04 as CaseNo,cp09,cp14,ep01,NVL(PA05,NVL(PA06,PA07)) as CaseName,Decode(PA09,'000',CPM03,CPM04) as CP10Nm" & _
                  " from casemap,caseprogress,engineerprogress,patent,casepropertymap" & _
                  " Where cm01='CFP' and cm10='0'" & _
                  " and cm05='" & PField(1) & "' and cm06='" & PField(2) & "' and cm07='" & PField(3) & "' and cm08='" & PField(4) & "'" & _
                  " and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+)" & _
                  " and cp10 in(" & NewCasePtyList & ")" & _
                  " and cp26 is null" & _
                  " and cp09=ep02(+)" & _
                  " and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04" & _
                  " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                  " and cp27 is null and cp57 is null" & _
                  " and cp14<>'" & m_FlowUserNum & "'" & _
                  " and not (pa09 in ('011','231') and cp10 in('101','102') and cp14 like 'F%')"
   
         'Add By Sindy 2015/10/21 +服務
         'Removed by Morgan 2025/4/17 服務業務不會有CFP案
         'strSql = strSql & " union select cp01||'-'||cp02||'-'||cp03||'-'||cp04 as CaseNo,cp09,cp14,ep01,NVL(SP05,NVL(SP06,SP07)) as CaseName,Decode(SP09,'000',CPM03,CPM04) as CP10Nm" & _
                  " from casemap,caseprogress,engineerprogress,servicepractice,casepropertymap" & _
                  " Where cm01='CFP' and cm10='0'" & _
                  " and cm05='" & PField(1) & "' and cm06='" & PField(2) & "' and cm07='" & PField(3) & "' and cm08='" & PField(4) & "'" & _
                  " and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+)" & _
                  " and cp10 in(" & NewCasePtyList & ")" & _
                  " and cp26 is null" & _
                  " and cp09=ep02(+)" & _
                  " and cp01=SP01 and cp02=SP02 and cp03=SP03 and cp04=SP04" & _
                  " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                  " and cp27 is null and cp57 is null" & _
                  " and cp14<>'" & m_FlowUserNum & "'"
      'Modify By Sindy 2020/8/12
      '系統別為CFP
      Else
         '排除CFP有國內外案件
         strSql = "select cm01||'-'||cm02||'-'||cm03||'-'||cm04" & _
                  " from casemap" & _
                  " Where cm05='P' and cm10='0'" & _
                  " and cm01='" & PField(1) & "' and cm02='" & PField(2) & "' and cm03='" & PField(3) & "' and cm04='" & PField(4) & "'"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strSql = ""
         Else
            strSql = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04 as CaseNo,cp09,cp14,ep01,NVL(PA05,NVL(PA06,PA07)) as CaseName,Decode(PA09,'000',CPM03,CPM04) as CP10Nm" & _
                     " from caserelation,caseprogress,engineerprogress,patent,casepropertymap" & _
                     " Where cr01='CFP'" & _
                     " and cr05='" & PField(1) & "' and cr06='" & PField(2) & "' and cr07='" & PField(3) & "' and cr08='" & PField(4) & "'" & _
                     " and cr01=cp01(+) and cr02=cp02(+) and cr03=cp03(+) and cr04=cp04(+)" & _
                     " and cp10 in(" & NewCasePtyList & ")" & _
                     " and cp26 is null" & _
                     " and cp09=ep02(+)" & _
                     " and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04" & _
                     " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                     " and cp27 is null and cp57 is null" & _
                     " and cp14<>'" & m_FlowUserNum & "'"
            'Removed by Morgan 2025/4/17 服務業務不會有CFP案
            'strSql = strSql & " union select cp01||'-'||cp02||'-'||cp03||'-'||cp04 as CaseNo,cp09,cp14,ep01,NVL(SP05,NVL(SP06,SP07)) as CaseName,Decode(SP09,'000',CPM03,CPM04) as CP10Nm" & _
                     " from caserelation,caseprogress,engineerprogress,servicepractice,casepropertymap" & _
                     " Where cr01='CFP'" & _
                     " and cr05='" & PField(1) & "' and cr06='" & PField(2) & "' and cr07='" & PField(3) & "' and cr08='" & PField(4) & "'" & _
                     " and cr01=cp01(+) and cr02=cp02(+) and cr03=cp03(+) and cr04=cp04(+)" & _
                     " and cp10 in(" & NewCasePtyList & ")" & _
                     " and cp26 is null" & _
                     " and cp09=ep02(+)" & _
                     " and cp01=SP01 and cp02=SP02 and cp03=SP03 and cp04=SP04" & _
                     " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                     " and cp27 is null and cp57 is null" & _
                     " and cp14<>'" & m_FlowUserNum & "'"
         End If
      End If
      '2020/8/12 END
      If strSql <> "" Then
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               'Modify By Sindy 2020/1/14 P案"送判"及"退件重送"時,不論多少次,均由系統自動將檔案用"聯絡"寄給CFP工程師。
               '    因此Mark...
'               '第二次不再傳
'               strSql = "select eep01 From empelectronprocess where eep01='" & rsTmp.Fields("cp09") & "'" & _
'                        " and eep04='" & EMP_聯絡 & "'" & _
'                        " and instr(eep11,'送件刪除')>0"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'               If intI <> 1 Then

                  'Modify By Sindy 2020/1/14 因不管次數了,所以此段程式也不管,均提供資料由工程師自己判斷
'                  'Modify By Sindy 2013/9/24 因CFP-26190,P-106351同時判發,因此增加控管若已有送判或判發流程就不增加聯絡
'                  strSql = "select eep01 From empelectronprocess where eep01='" & rsTmp.Fields("cp09") & "'" & _
'                           " and eep04 in('" & EMP_送判 & "','" & EMP_判發 & "')"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI <> 1 Then
'                  '2013/9/24 END
                     '取得最大序號
                     intMaxEEP02 = 0
                     strSql = "select eep02 From empelectronprocess where eep01='" & rsTmp.Fields("cp09") & "' order by eep02 desc"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        RsTemp.MoveFirst
                        If RsTemp.RecordCount > 0 Then
                           intMaxEEP02 = RsTemp.Fields(0)
                        End If
                     End If
                     '新增CFP聯絡歷程
                     'Modify By Sindy 2023/12/18 +,eep16
                     strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep11,eep12,eep16) values(" & _
                              CNULL(rsTmp.Fields("cp09")) & "," & intMaxEEP02 + 1 & "," & CNULL(Trim(txtEEP03)) & "," & _
                              CNULL(EMP_聯絡) & "," & CNULL(rsTmp.Fields("cp14")) & "," & strSrvDate(1) & "," & _
                              strUpdTime & ",'" & Replace(lblCaseNo, "-0-00", "") & "=>" & Replace(rsTmp.Fields("CaseNo"), "-0-00", "") & "送件刪除','" & m_EEP12 & "','" & m_EEP16 & "')"
                     cnnConnection.Execute strSql
                     '儲存附件
                     If SaveAttFile(rsTmp.Fields("cp09"), intMaxEEP02 + 1, 0) = False Then
                        GoTo ErrHand
                     End If
                     '寄Mail
                     'Modify By Sindy 2023/12/14 杜燕文協理請作,主旨加申請國家
                     strSubject = Replace(rsTmp.Fields("CaseNo"), "-0-00", "") & "(" & GetPrjNation(rsTmp.Fields("CaseNo")) & ")(核會流程)-->(" & (intMaxEEP02 + 1) & ")聯絡，請進行後續處理"
                     'Modify By Sindy 2023/8/9 + & PField(2) 方便知道
                     'Modify By Sindy 2023/12/15 杜燕文協理請作,內文加申請國家
                     strContent = "當月目次：" & rsTmp.Fields("ep01") & vbCrLf & _
                                  "本所案號：" & rsTmp.Fields("CaseNo") & vbCrLf & _
                                  "案件名稱：" & rsTmp.Fields("CaseName") & vbCrLf & _
                                  "申請國家：" & GetPrjNation(rsTmp.Fields("CaseNo")) & vbCrLf & _
                                  "案件性質：" & rsTmp.Fields("CP10Nm") & vbCrLf & _
                                  "流程狀態：聯絡" & vbCrLf & vbCrLf & vbCrLf & _
                                  "★本案" & PField(1) & "案(" & PField(1) & "-" & PField(2) & ")說明書已上傳系統，請至待辦歷程中下載。"
                     PUB_SendMail strUserNum, rsTmp.Fields("cp14"), rsTmp.Fields("cp09"), strSubject, strContent 'Add By Sindy 2022/2/21 + m_EEP01
'                  End If
'               End If
               rsTmp.MoveNext
            Loop
         End If
      End If
      End If
      rsTmp.Close
   End If
   
   'CFP案送件時檢查是否有〔聯絡〕的備註是〔送件刪除〕字樣,若有則一併刪除
   If PField(1) = "CFP" And _
      (Left(CboEEP04.Text, 2) = EMP_送判 Or Left(CboEEP04.Text, 2) = EMP_判發) Then
      strSql = "select eep02 From empelectronprocess where eep01='" & m_EEP01 & "'" & _
                 " and eep04='" & EMP_聯絡 & "' and instr(eep11,'送件刪除')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strSql = "delete from empelectronprocess where EEP01='" & m_EEP01 & "' and EEP02=" & RsTemp.Fields("EEP02")
         cnnConnection.Execute strSql
         PUB_DelFtpFile2 m_EEP01, " and EEF02=" & RsTemp.Fields("EEP02"), "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
         'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
         strSql = "delete from empelectronfile where EEF01='" & m_EEP01 & "' and EEF02=" & RsTemp.Fields("EEP02")
         cnnConnection.Execute strSql
      End If
   End If
   '2013/9/18 END
   
   'Modify By Sindy 2013/12/5 Mark 因P-107111回代需要doc
'   'Add By Sindy 2013/9/25 附加流程判發時刪除DOC文件
'   If Left(CboEEP04.Text, 2) = EMP_判發 And bolBCaseFlow = True Then
'      strSql = "delete from empelectronfile where EEF01='" & m_EEP01 & "' and EEF02=" & m_EEP02 & " and substr(upper(eef03),-4)='.DOC'"
'      cnnConnection.Execute strSql
'   End If
   
   If bolPAFlow = True Then
      'Modify By Sindy 2018/12/25 Mark,因程序人員負責做電子檔轉檔工作(及上傳E-Set)
'      'Add By Sindy 2013/10/17 判發時若為電子送件刪除相關pdf檔(說明書)及繪圖pdf檔
'      If Left(CboEEP04.Text, 2) = EMP_判發 And m_Country = "000" And cp(118) <> "" Then
'         PUB_DelFtpFile2 m_EEP01, " and EEF02=" & m_EEP02 & " and instr(upper(eef03),upper('." & m_CPM26 & "'))>0 and substr(upper(eef03),-4)='.PDF'", "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
'         'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
'         strSql = "delete from empelectronfile where EEF01='" & m_EEP01 & "' and EEF02=" & m_EEP02 & " and instr(upper(eef03),upper('." & m_CPM26 & "'))>0 and substr(upper(eef03),-4)='.PDF'"
'         cnnConnection.Execute strSql
'         'Modify By Sindy 2014/9/16 發明,新型若為新申請案時不刪除DWG.PDF
'         'Modify By Sindy 2014/9/17 秀玲:301改請發明302改請新型也不刪除DWG.PDF
'         If Not ((m_PA08 = "1" Or m_PA08 = "2") And (InStr(NewCasePtyList, cp(10)) > 0 Or cp(10) = "301" Or cp(10) = "302")) Then
'         '2014/9/16 END
'            PUB_DelFtpFile2 m_EEP01, " and EEF02=" & m_EEP02 & " and substr(upper(eef03),-7)='DWG.PDF'", "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
'            'Memo by Morgan 2015/4/28 刪除條件要和前面刪除FTP檔的同步
'            strSql = "delete from empelectronfile where EEF01='" & m_EEP01 & "' and EEF02=" & m_EEP02 & " and substr(upper(eef03),-7)='DWG.PDF'"
'            cnnConnection.Execute strSql
'         End If
'      End If
      
      'Add By Sindy 2017/9/5 多國案新增歷程
      If cmdCaseMap.Visible = True And cmdCaseMap.Enabled = True And m_RetrunRecv <> "" Then
         If ProcessCaseMap(strUpdTime) = False Then
            GoTo ErrHand
         End If
      End If
      '2017/9/5 END
   End If
   
   'Add by Sindy 2024/1/2
   If m_strIR01 <> "" And _
      (Left(CboEEP04.Text, 2) = EMP_翻譯交稿 Or cp(10) = "924") And _
      UCase(m_PrevForm.Name) = "FRM060107_1" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm060107"
   End If
   '2024/1/2 END
   
   'Add By Sindy 2024/2/20 代為操作時: 修改核稿人,英文核稿人,判發人為操作人員
   If intReceiveKind = "1" And m_EEP12 = "(代)" Then  '1.待核判區
      strExc(9) = ""
      If m_strLastEEP04 = EMP_送核 Then
         strExc(9) = " EP04='" & strUserNum & "'" '核稿人
      ElseIf m_strLastEEP04 = EMP_送英核 Then
         strExc(9) = " EP03='" & strUserNum & "'" '英文核稿人
      ElseIf m_strLastEEP04 = EMP_送判 Then
         strExc(9) = " EP40='" & strUserNum & "'" '判發人
      End If
      If strExc(9) <> "" Then
         strSql = "update engineerprogress set" & _
                        strExc(9) & _
                  " where ep02='" & m_EEP01 & "'"
         cnnConnection.Execute strSql
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                           strExc(9) & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            cnnConnection.Execute strSql
         End If
      End If
   End If
   '2024/2/20 END
   
   'Add By Sindy 2024/2/26
   If intReceiveKind = 0 Then '0.承辦人工作進度
      If bolFCPFlow = True Then
         Call m_PrevForm.UpdEngMdb '更新最新歷程狀態
      End If
   End If
   '2024/2/26 END
   
   'Added by Morgan 2025/4/17
   If iRec > 0 Then
      For iRec = 1 To UBound(strCFP209EPP01)
         '取得最大序號
         intMaxEEP02 = 0
         strSql = "select eep02 From empelectronprocess where eep01='" & strCFP209EPP01(iRec) & "' order by eep02 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            RsTemp.MoveFirst
            If RsTemp.RecordCount > 0 Then
               intMaxEEP02 = RsTemp.Fields(0)
            End If
         End If
         '新增CFP聯絡歷程
         intMaxEEP02 = intMaxEEP02 + 1
         strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep10)" & _
                  " values('" & strCFP209EPP01(iRec) & "'," & intMaxEEP02 & ",'" & strUserNum & "'" & _
                  ",'" & EMP_聯絡 & "','98012',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                  ",'" & ChgSQL(strCFP209EPP08(iRec)) & "','99050')"
         cnnConnection.Execute strSql, intI
         
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc09,mc13)" & _
            " select eep03 mc01,eep05 mc02,eep06 mc03,eep07 mc04,'" & ChgSQL(strCFP209Subj(iRec)) & "' mc07,eep10 mc09,eep01 mc13" & _
            " from empelectronprocess where eep01='" & strCFP209EPP01(iRec) & "' and eep02=" & intMaxEEP02
         cnnConnection.Execute strSql, intI
      Next
   End If
   'end 2025/4/17
   
   cnnConnection.CommitTrans: bolConn = False
   If bolTMFlow = True And cmdCP118.Tag = "Y" Then Unload frm090202_2_2 'Add By Sindy 2018/8/14
   
   Call FlowSendMail(bolSendMail, strAutoFlow, IIf(txtLpNote.Tag = "多案單筆歷程", True, False)) '發通知信 Modify By Sindy 2018/10/23 (改放函數)
   
   Screen.MousePointer = vbDefault
   bolSave = True
   
   Set rsTmp = Nothing
   KillAttach 'Add By Sindy 2013/9/13 此資料夾為系統暫存檔案時使用，以防事後有人開檔鎖住問題，在此時先一併清資料夾
   
   'Add By Sindy 2023/11/10
   If UCase(m_PrevForm.Name) = UCase("frm090202_4_1") Then
      Call m_PrevForm.cmdExit_Click
   'Add by Sindy 2024/1/2
   ElseIf UCase(m_PrevForm.Name) = "FRM060107_1" And m_strIR01 <> "" Then
      If Not m_PrevForm_IR Is Nothing Then
         Call m_PrevForm_IR.GoNext
         Set m_PrevForm_IR = Nothing
      End If
      Unload m_PrevForm
      '2024/1/2 END
   End If
   '2023/11/10 END
   
   Call cmdExit_Click
   Exit Sub
   
ErrHand:
   cmdSend.Enabled = True
   cmdExit.Enabled = True 'Add By Sindy 2018/11/7
   Screen.MousePointer = vbDefault
   'Resume Next
   If bolConn = True Then
      cnnConnection.RollbackTrans
   End If
   bolConn = False
   
   dblErrNumber = Err.Number 'Add By Sindy 2021/8/26
   strErrText = strErrText & vbCrLf & vbCrLf & Err.Description 'Add By Sindy 2021/8/26
   'MsgBox " 送出失敗！" & vbCrLf & Err.Description
   'If Err.Number <> 0 Then
   If dblErrNumber <> 0 Then 'Modify By Sindy 2021/8/26
      MsgBox "送出失敗！" & vbCrLf & dblErrNumber & vbCrLf & Err.Description, vbExclamation
      'Add By Sindy 2019/12/12 抓bug
      'Modify By Sindy 2024/11/18
      If CheckIsPersonRest("97038", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
      '2024/11/18 END
         If UCase(cmdSend.Caption) = UCase("E-Mail") Then
            PUB_SendMail strUserNum, "97038", "", "cmdSend.Caption=E-Mail [送出失敗,出現錯誤]:" & cmdSend.Tag & "; 主旨:" & strSubject, _
               strContent & vbCrLf & vbCrLf & _
               m_EEP01 & ":" & PField(1) & PField(2) & PField(3) & PField(4) & vbCrLf & vbCrLf & _
               "Err.Number: " & dblErrNumber & vbCrLf & vbCrLf & _
               "Err.Description: " & vbCrLf & strErrText & vbCrLf & vbCrLf & _
               "strsql: " & strSql, , , , , , , , , , True, False, , , False
         End If
      End If
      '2019/12/12 END
   End If
   Exit Sub
End Sub

'Add By Sindy 2025/4/21 函數程式太大了,要切為2
Private Sub FormSave2(strUpdDate As String, strUpdTime As String, bolSendMail As Boolean, bolRegMail As Boolean, _
                      strLP31 As String)
Dim intMaxEEP02 As Integer
Dim rsTmp As New ADODB.Recordset
Dim strConSql As String
Dim str_924CP09 As String 'Add By Sindy 2023/10/2
Dim strEED05 As String, strEED10 As String
Dim arrID As Variant, intCnt As Integer
Dim strCP23 As String 'Add By Sindy 2018/9/25
   
   '********************************************************************************
   If Left(CboEEP04.Text, 2) = EMP_草完 Or Left(CboEEP04.Text, 2) = EMP_草核 Then
      '更新EP15.草圖完稿日
'      If UCase(m_PrevForm.Name) = UCase("frm090711") Then
'         If Trim(m_PrevForm.txt1(2)) = "" Then
'            m_PrevForm.txt1(2) = strUpdDate
'         End If
'      Else
         strSql = "update engineerprogress set" & _
                        " EP15=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "' and (EP15 is null or EP15=0)"
         cnnConnection.Execute strSql
         If UCase(m_PrevForm.Name) = UCase("frm090711") Then
            If Trim(m_PrevForm.txt1(2)) = "" Then
               m_PrevForm.txt1(2) = strUpdDate
            End If
         End If
'      End If
   End If
   '********************************************************************************
   'Modify By Sindy 2013/9/30 +送英核
   If Left(CboEEP04.Text, 2) = EMP_送核 Or _
      Left(CboEEP04.Text, 2) = EMP_送英核 Then
      '更新EP09.完稿日
      If intReceiveKind = 0 Then '0.承辦人工作進度
         If Trim(m_PrevForm.txt1(3)) = "" Then
            m_PrevForm.txt1(3) = strUpdDate
         End If
      Else
         strSql = "update engineerprogress set" & _
                        " EP09=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "' and (EP09 is null or EP09=0)"
         cnnConnection.Execute strSql
'         If m_PrevForm.intBackTab = 1 Then
'            m_PrevForm.txt1(3) = strUpdDate
'         End If
      End If
      'Add By Sindy 2020/9/29
      If txtLpNote.Tag = "多案單筆歷程" Then
         strSql = "update engineerprogress set" & _
                        " EP09=" & DBDATE(strUpdDate) & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP09 is null or EP09=0)"
         cnnConnection.Execute strSql
         strSql = "update engineerprogress set" & _
                        " EP04='" & Left(m_CMMan, 5) & "'" & _
                        ",EP34=" & CNULL(m_EP34) & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
         cnnConnection.Execute strSql
      End If
      '2020/9/29 END
      If Left(CboEEP04.Text, 2) = EMP_送核 Then
         '清空EP39.核稿完成日
         strSql = "update engineerprogress set" & _
                        " EP39=null" & _
                  " where ep02='" & m_EEP01 & "'"
         cnnConnection.Execute strSql
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                           " EP39=null" & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
         'Add By Sindy 2018/4/30
         '清空EP42.判發完成日
         strSql = "update engineerprogress set" & _
                        " EP42=null" & _
                  " where ep02='" & m_EEP01 & "'"
         cnnConnection.Execute strSql
         '2018/4/30 END
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                           " EP42=null" & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
      End If
   End If
   '********************************************************************************
   'Modify By Sindy 2013/11/13 +Or (Check1.Visible = True And Check1.Value = 1)
   If (Check1.Visible = True And Check1.Value = 1) Then
      m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "更新英文核完日"
   End If
   If Left(CboEEP04.Text, 2) = EMP_核完 Or (Check1.Visible = True And Check1.Value = 1) Then
      If m_strLastEEP04 = EMP_送英核 Or (Check1.Visible = True And Check1.Value = 1) Then
   '2013/11/13 END
         '更新EP33.英文核完日
         If UCase(m_PrevForm.Name) = UCase("frm090201_2") Then
            If Trim(m_PrevForm.txt1(19)) = "" Then
               m_PrevForm.txt1(19) = strUpdDate
            End If
         Else
            strSql = "update engineerprogress set" & _
                           " EP33=" & DBDATE(strUpdDate) & _
                     " where ep02='" & m_EEP01 & "' and (EP33 is null or EP33=0)"
            cnnConnection.Execute strSql
'            If m_PrevForm.intBackTab = 1 Then
'               m_PrevForm.txt1(19) = strUpdDate
'            End If
         End If
         'Add By Sindy 2013/10/4 不會稿時,已核完稿才一併更新會稿日及會稿完成日
         If m_EP34 = "N" Then
            If (m_CMMan <> "" And Val(m_EP39) > 0) Or m_CMMan = "" Then
               'Modify By Sindy 2018/10/16 Mark
'               If bolTMFlow = True Then
'                  '更新EP08.會稿完成日
'                  strSql = "update engineerprogress set" & _
'                                 " EP08=" & DBDATE(strUpdDate) & _
'                           " where ep02='" & m_EEP01 & "' and (EP08 is null or EP08=0)"
'                  cnnConnection.Execute strSql
'                  '更新EP07.會稿日
'                  strSql = "update engineerprogress set" & _
'                                 " EP07=" & DBDATE(strUpdDate) & _
'                           " where ep02='" & m_EEP01 & "' and (EP07 is null or EP07=0)"
'                  cnnConnection.Execute strSql
'               Else
               'Add By Sindy 2018/11/7 + if
               If bolPAFlow = True Then
               '2018/11/7 END
                  UpdateEp08 m_EEP01, DBDATE(strUpdDate) '更新相關會稿完成日資料
                  '更新EP07.會稿日
                  strSql = "update engineerprogress set" & _
                                 " EP07=" & DBDATE(strUpdDate) & _
                           " where ep02='" & m_EEP01 & "' and (EP07 is null or EP07=0)"
                  cnnConnection.Execute strSql
               End If
            End If
         End If
         '2013/10/4 END
      End If
      If m_strLastEEP04 = EMP_送核 Then '送核-核完時,上核稿完成日
         'Add By Sindy 2024/1/5
         If Lbl926.Visible = True And InStr(Lbl926.Caption, "一核") > 0 Then
            '清除承辦期限
            strSql = "UPDATE CaseProgress SET CP48=null" & _
                     " WHERE CP09 = '" & m_EEP01 & "' and cp48>0"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
'            '清除完稿日
'            strSql = "update engineerprogress set" & _
'                           " EP09=null" & _
'                     " where ep02='" & m_EEP01 & "' and EP09>0"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
         Else
         '2024/1/5 END
            '更新EP39.核稿完成日
            strSql = "update engineerprogress set" & _
                           " EP39=" & DBDATE(strUpdDate) & _
                     " where ep02='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
            'Add By Sindy 2020/9/29
            If txtLpNote.Tag = "多案單筆歷程" Then
               strSql = "update engineerprogress set" & _
                              " EP39=" & DBDATE(strUpdDate) & _
                        " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
               cnnConnection.Execute strSql
            End If
            '2020/9/29 END
         End If
         
         'Add By Sindy 2013/10/4 不會稿時,已核完稿才一併更新會稿日及會稿完成日
         '註:如一案二請另一件為不會稿,當主管核完稿時,一併更新會完日同時系統通知上墨
         If m_EP34 = "N" Then
            If (m_EMMan <> "" And Val(m_EP33) > 0) Or m_EMMan = "" Then
               'Modify By Sindy 2018/10/16 Mark
'               If bolTMFlow = True Then
'                  '更新EP08.會稿完成日
'                  strSql = "update engineerprogress set" & _
'                                 " EP08=" & DBDATE(strUpdDate) & _
'                           " where ep02='" & m_EEP01 & "' and (EP08 is null or EP08=0)"
'                  cnnConnection.Execute strSql
'                  '更新EP07.會稿日
'                  strSql = "update engineerprogress set" & _
'                                 " EP07=" & DBDATE(strUpdDate) & _
'                           " where ep02='" & m_EEP01 & "' and (EP07 is null or EP07=0)"
'                  cnnConnection.Execute strSql
'               Else
               'Add By Sindy 2018/11/7 + if
               If bolPAFlow = True Then
               '2018/11/7 END
                  UpdateEp08 m_EEP01, DBDATE(strUpdDate) '更新相關會稿完成日資料
                  '更新EP07.會稿日
                  strSql = "update engineerprogress set" & _
                                 " EP07=" & DBDATE(strUpdDate) & _
                           " where ep02='" & m_EEP01 & "' and (EP07 is null or EP07=0)"
                  cnnConnection.Execute strSql
               End If
            End If
         End If
         '2013/10/4 END
      End If
   End If
   '********************************************************************************
   If Left(CboEEP04.Text, 2) = EMP_送會 Then
      'Modify By Sindy 2018/10/23 商標恢復要寄通知信,但不能夾帶附件
      'If bolTMFlow = True Then bolSendMail = False 'Add By Sindy 2018/9/14 暫時先不需要寄Mail
      
      'Add By Sindy 2015/12/2
      If PUB_ChkEmpFlowExists(m_EEP01, EMP_送會) = False And Val(m_EP07) > 0 Then
         '更新EP07.會稿日
         If intReceiveKind = 0 Then '0.承辦人工作進度
            m_PrevForm.txt1(4) = strUpdDate
         Else
            strSql = "update engineerprogress set" & _
                           " EP07=" & DBDATE(strUpdDate) & _
                     " where ep02='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
         End If
         '清除EP08.會稿完成日
         If intReceiveKind = 0 Then '0.承辦人工作進度
            m_PrevForm.txt1(7) = ""
         Else
            strSql = "update engineerprogress set" & _
                           " EP08=null" & _
                     " where ep02='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
         End If
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                           " EP07=" & DBDATE(strUpdDate) & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            cnnConnection.Execute strSql
            strSql = "update engineerprogress set" & _
                           " EP08=null" & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
      Else
      '2015/12/2 END
         '更新EP07.會稿日
         If intReceiveKind = 0 Then '0.承辦人工作進度
            'Modify By Sindy 2024/5/14 因為送會後,會稿日沒寫入日期,在查原因
            '                          + Or Val(m_EP07) = 0
            If Trim(m_PrevForm.txt1(4)) = "" Or Val(m_EP07) = 0 Then
               m_PrevForm.txt1(4) = strUpdDate
            End If
         Else
            strSql = "update engineerprogress set" & _
                           " EP07=" & DBDATE(strUpdDate) & _
                     " where ep02='" & m_EEP01 & "' and (EP07 is null or EP07=0)"
            cnnConnection.Execute strSql
   '         If m_PrevForm.intBackTab = 1 Then
   '            m_PrevForm.txt1(4) = strUpdDate
   '         End If
         End If
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                           " EP07=" & DBDATE(strUpdDate) & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP07 is null or EP07=0)"
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
      End If
      
      'Add By Sindy 2021/3/31 再做送會時,要清除EP37.客戶會稿日
      strSql = "update engineerprogress set" & _
                     " EP37=null" & _
               " where ep02='" & m_EEP01 & "' and EP37 is not null"
      cnnConnection.Execute strSql
      If txtLpNote.Tag = "多案單筆歷程" Then
         strSql = "update engineerprogress set" & _
                        " EP37=null" & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and EP37 is not null"
         cnnConnection.Execute strSql
      End If
      '2021/3/31 END
      
      '更新EP09.完稿日
      If intReceiveKind = 0 Then '0.承辦人工作進度
         If Trim(m_PrevForm.txt1(3)) = "" Then
            m_PrevForm.txt1(3) = strUpdDate
         End If
      Else
         strSql = "update engineerprogress set" & _
                        " EP09=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "' and (EP09 is null or EP09=0)"
         cnnConnection.Execute strSql
'         If m_PrevForm.intBackTab = 1 Then
'            m_PrevForm.txt1(3) = strUpdDate
'         End If
      End If
      'Add By Sindy 2020/9/29
      If txtLpNote.Tag = "多案單筆歷程" Then
         strSql = "update engineerprogress set" & _
                        " EP09=" & DBDATE(strUpdDate) & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP09 is null or EP09=0)"
         cnnConnection.Execute strSql
         
         strSql = "update engineerprogress set" & _
                        " EP34=" & CNULL(m_EP34) & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
         cnnConnection.Execute strSql
      End If
      '2020/9/29 END
      
      'Add By Sindy 2020/12/23
      Dim m_CP17 As String
      If bolTMFlow = True Then
         If cmdCP118.Tag = "Y" Then
            For ii = 1 To frm090202_2_2.MSHFlexGrid2.Rows - 1
               If frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 0) = "V" And _
                  frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 2) <> "" Then
                  'Modify By Sindy 2020/12/28 只更新發文規費,CP118不上Y (cp118='Y')
                  strSql = "update caseprogress set" & _
                           " cp84=" & CNULL(Val(frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 7)), True) & _
                           " where cp09='" & frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 2) & "'"
                  cnnConnection.Execute strSql
                  m_CP17 = 0
                  strSql = "select CP17 From caseprogress where cp09='" & frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 2) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     m_CP17 = "" & RsTemp.Fields("CP17")
                  End If
                  
                  '台灣案發文規費與收文規費不符時,mail給智權人員及財務處總帳人員
                  If Val(frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 7)) <> Val(m_CP17) Then
                     PUB_ChkOfficialFee frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 2), Val(frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 7))
                  End If
               End If
            Next ii
         End If
      End If
      '2020/12/23 END
   End If
   'Add By Sindy 2018/8/28
   '********************************************************************************
   If Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
      'Modify By Sindy 2019/9/5 + 客服組專利會稿工程師執行客戶會稿，同時副本通知客服組成員。
      '會稿方式不是E-Mail
      If Trim(Left(CboCP10.Text, 1)) <> "1" And _
         InStr(Pub_GetSpecMan("WSpecial"), Me.m_FlowUserNum) > 0 And _
         InStr(Pub_GetSpecMan("客服組專利會稿工程師"), strUserNum) > 0 Then '創新業務部可個人收文成員
         CboEEP05.Text = Me.m_FlowUserNum
         bolSendMail = True '寄Mail
      Else
      '2019/9/5 END
         bolSendMail = False '不需要寄Mail
      End If
      'Modify By Sindy 2018/8/29 會稿方式
      m_EEP14 = Left(Trim(CboCP10.Text), 1)
'   End If
'   If Left(CboEEP04.Text, 2) = EMP_客戶會稿 And m_strLastEEP04 = EMP_送會 Then
      '更新EP37.客戶會稿日
      strSql = "update engineerprogress set" & _
                     " EP37=" & DBDATE(strUpdDate) & _
               " where ep02='" & m_EEP01 & "' and (EP37 is null or EP37=0)"
      cnnConnection.Execute strSql
      'Add By Sindy 2020/9/29
      If txtLpNote.Tag = "多案單筆歷程" Then
         strSql = "update engineerprogress set" & _
                        " EP37=" & DBDATE(strUpdDate) & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP37 is null or EP37=0)"
         cnnConnection.Execute strSql
      End If
      '2020/9/29 END
   End If
   '2018/8/28 END
   
   '********************************************************************************
   If Left(CboEEP04.Text, 2) = EMP_會完 Then
      '更新EP38.智權人員會稿完成日
      'Modify By Sindy 2013/9/10 只要有做會完,就更新智權人員會稿完成日
'      strSql = "update engineerprogress set" & _
'                     " EP38=" & DBDATE(strUpdDate) & _
'               " where ep02='" & m_EEP01 & "' and (EP38 is null or EP38=0)"
      strSql = "update engineerprogress set" & _
                     " EP38=" & DBDATE(strUpdDate) & _
               " where ep02='" & m_EEP01 & "'"
      cnnConnection.Execute strSql
      'Add By Sindy 2020/9/29
      If txtLpNote.Tag = "多案單筆歷程" Then
         strSql = "update engineerprogress set" & _
                        " EP38=" & DBDATE(strUpdDate) & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
         cnnConnection.Execute strSql
      End If
      '2020/9/29 END
      
      'Add By Sindy 2018/7/17 商標處會完時,更新會稿完成日
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
      If bolTMFlow = True Or bolOtherFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
         'Modify By Sindy 2018/9/20 + 是否通知客戶
         strSql = "update engineerprogress set" & _
                        " EP08=" & DBDATE(strUpdDate) & _
                        IIf(ChkEP11.Visible = True, IIf(ChkEP11.Value = 1, ",EP11='N'", ",EP11='Y'"), "") & _
                  " where ep02='" & m_EEP01 & "'"
         Pub_SeekTbLog strSql 'Add By Sindy 2021/6/28
         cnnConnection.Execute strSql
         
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                           " EP08=" & DBDATE(strUpdDate) & _
                           IIf(ChkEP11.Visible = True, IIf(ChkEP11.Value = 1, ",EP11='N'", ",EP11='Y'"), "") & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/6/28
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
      End If
'      '更新EP08.會稿完成日
'      If UCase(m_PrevForm.Name) = UCase("frm090201_2") Then
'         If Trim(m_PrevForm.txt1(7)) = "" Then
'            m_PrevForm.txt1(7) = strUpdDate
'         End If
'      Else
'         strSql = "update engineerprogress set" & _
'                        " EP08=" & DBDATE(strUpdDate) & _
'                  " where ep02='" & m_EEP01 & "' and (EP08 is null or EP08=0)"
'         cnnConnection.Execute strSql
'         If m_PrevForm.intBackTab = 1 Then
'            m_PrevForm.txt1(7) = strUpdDate
'         End If
'      End If
   End If
   '********************************************************************************
   'Add By Sindy 2016/3/22 + EMP_圖修
   If Left(CboEEP04.Text, 2) = EMP_圖修 Then
      '若此時「齊備日」為空白，則系統自動以「圖修」的日期更新「齊備日」
      strSql = "update engineerprogress set" & _
                     " EP06=" & DBDATE(strUpdDate) & _
               " where ep02='" & m_EEP01 & "' and (EP06 is null or EP06=0)"
      cnnConnection.Execute strSql
   End If
   '********************************************************************************
   'Add By Sindy 2016/3/23 + EMP_圖完
   If Left(CboEEP04.Text, 2) = EMP_圖完 Then
      '有齊備日發出「圖完」時，則系統不作特別的操作，記錄狀況供後續判斷
      If m_EP06 <> "" Then
         m_UpdEEP11 = m_UpdEEP11 & IIf(m_UpdEEP11 <> "", ",", "") & "會(圖/文)完成;原齊備日:" & m_EP06
      End If
   End If
   '********************************************************************************
   'Add By Sindy 2016/3/10 +會完重修
   If Left(CboEEP04.Text, 2) = EMP_會完重修 Then
      '記錄原日期於進度備註裡
      'Modify By Sindy 2016/10/14 + 原墨圖齊備日='||ep17||'原墨圖完稿日='||ep18||'
      'Modify By Sindy 2018/10/1 ex:T-217282
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
      If bolTMFlow = True Or bolOtherFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then '商標處,其他
         strSql = "update caseprogress" & _
                  " set cp64=(select '會完重修:原會稿完成日='||ep08||'原智權人員會稿完成日='||ep38||';'||cp64 from caseprogress,engineerprogress where cp09='" & m_EEP01 & "' and cp09=ep02(+))" & _
                  " where cp09='" & m_EEP01 & "'"
         cnnConnection.Execute strSql
         
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update caseprogress" & _
                     " set cp64=(select '會完重修:原會稿完成日='||ep08||'原智權人員會稿完成日='||ep38||';'||cp64 from caseprogress,engineerprogress where cp09='" & m_EEP01 & "' and cp09=ep02(+))" & _
                     " where cp09 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
      Else
      '2018/10/1 END
         strSql = "update caseprogress" & _
                  " set cp64=(select '會完重修:原會稿完成日='||ep08||'原智權人員會稿完成日='||ep38||'原墨圖齊備日='||ep17||'原墨圖完稿日='||ep18||';'||cp64 from caseprogress,engineerprogress where cp09='" & m_EEP01 & "' and cp09=ep02(+))" & _
                  " where cp09='" & m_EEP01 & "'"
         cnnConnection.Execute strSql
      End If
      '清除EP08.會稿完成日,EP38.智權人員會稿完成日
      'Modify By Sindy 2016/10/14 + ,EP17=null,EP18=null ex:柏翰P-115618會完時,沒有再通知系統上墨
      strSql = "update engineerprogress set" & _
                     " EP08=null,EP38=null,EP17=null,EP18=null" & _
               " where ep02='" & m_EEP01 & "'"
      cnnConnection.Execute strSql
      'Add By Sindy 2020/9/29
      If txtLpNote.Tag = "多案單筆歷程" Then
         strSql = "update engineerprogress set" & _
                     " EP08=null,EP38=null,EP17=null,EP18=null" & _
               " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
         cnnConnection.Execute strSql
      End If
      '2020/9/29 END
   End If
   '2016/3/10 END
   '********************************************************************************
   'Add By Sindy 2013/10/1 工程師上墨後,系統要一併更新墨齊日
   If Left(CboEEP04.Text, 2) = EMP_上墨 Then
      '更新EP17.墨圖齊備日
      'Modify By Sindy 2014/3/31 + and (EP17 is null or EP17=0)
      strSql = "update engineerprogress set" & _
                     " EP17=" & DBDATE(strUpdDate) & _
               " where ep02='" & m_EEP01 & "' and (EP17 is null or EP17=0)"
      cnnConnection.Execute strSql
   End If
   '2013/10/1 END
   '********************************************************************************
   'Add By Sindy 2014/1/29 工程師送標號後,系統要一併更新草齊日 ex.CFP-026540
   If Left(CboEEP04.Text, 2) = EMP_送標號 Then
      '更新EP14.草圖齊備日
      strSql = "update engineerprogress set" & _
                     " EP14=" & DBDATE(strUpdDate) & _
               " where ep02='" & m_EEP01 & "' and nvl(ep14,0)=0"
      cnnConnection.Execute strSql
   End If
   '2014/1/29 END
   '********************************************************************************
   '2013/6/4 繪圖人員等於繪圖主管時,則自行核判,判發時要更新墨圖完稿日
   If Left(CboEEP04.Text, 2) = EMP_墨完 Or Left(CboEEP04.Text, 2) = EMP_繪圖判發 Then
      'Add By Sindy 2025/3/18
      If Left(CboEEP04.Text, 2) = EMP_繪圖判發 And cp(1) = "FCP" And cp(10) = "931" Then
         '更新完稿日
         strSql = "update engineerprogress set" & _
                        " EP09=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "' and (EP09 is null or EP09=0)"
         cnnConnection.Execute strSql
         '更新EP04核稿工程師
         strSql = "update engineerprogress set" & _
                        " EP04=" & CNULL(Trim(Left(CboEEP05.Text, 6))) & _
                  " where ep02='" & m_EEP01 & "'"
         cnnConnection.Execute strSql
      End If
      '2025/3/18 END
      
      '更新EP18.墨圖完稿日
'      If UCase(m_PrevForm.Name) = UCase("frm090711") Then
'         If Trim(m_PrevForm.txt1(5)) = "" Then
'            m_PrevForm.txt1(5) = strUpdDate
'         End If
'      Else
         strSql = "update engineerprogress set" & _
                        " EP18=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "' and (EP18 is null or EP18=0)"
         cnnConnection.Execute strSql
         If UCase(m_PrevForm.Name) = UCase("frm090711") Then
            If Trim(m_PrevForm.txt1(5)) = "" Then
               m_PrevForm.txt1(5) = strUpdDate
            End If
         End If
'      End If
      'Modify By Sindy 2013/10/1 檢查是否有草齊日
      strSql = "select ep02,ep14,ep15,ep17,ep18 From engineerprogress where ep02='" & m_EEP01 & "'" & _
                 " and ep14 is not null and ep14>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Modify By Sindy 2013/10/1 沒有草完日,則一併補上草完日=墨完日
         If Val("" & RsTemp.Fields("ep18")) > 0 Then
            strDate = RsTemp.Fields("ep18")
         Else
            strDate = DBDATE(strUpdDate)
         End If
         If Val("" & RsTemp.Fields("ep15")) = 0 Then
            strSql = "update engineerprogress set" & _
                           " ep15=" & strDate & _
                     " where ep02='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
            If UCase(m_PrevForm.Name) = UCase("frm090711") Then
               If Trim(m_PrevForm.txt1(2)) = "" Then
                  m_PrevForm.txt1(2) = Val(strDate) - 19110000
               End If
            End If
         End If
         'Modify By Sindy 2013/10/8 沒有墨齊日,則一併補上墨齊日=草齊日
         If Val("" & RsTemp.Fields("ep17")) = 0 Then
            strSql = "update engineerprogress set" & _
                           " ep17=" & RsTemp.Fields("ep14") & _
                     " where ep02='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
            If UCase(m_PrevForm.Name) = UCase("frm090711") Then
               If Trim(m_PrevForm.txt1(4)) = "" Then
                  m_PrevForm.txt1(4) = Val(RsTemp.Fields("ep14")) - 19110000
               End If
            End If
         End If
         '2013/10/8 END
      End If
      '2013/10/1 END
   End If
   
   'Add By Sindy 2022/4/27
   '********************************************************************************
   'm_strSpecState: 特殊情況 ex:尚待收款-完稿日
   '********************************************************************************
   If m_strSpecState = "尚待收款-完稿日" And Left(CboEEP04.Text, 2) = EMP_聯絡 Then
      '更新完稿日
      If txtLpNote.Tag = "多案單筆歷程" Then
         strSql = "update engineerprogress set" & _
                        " EP09=" & DBDATE(m_PrevForm.txt1(3)) & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP09 is null or EP09=0)"
         cnnConnection.Execute strSql
      End If
   End If
   '2022/4/27 END
   
   '********************************************************************************
   'Added by Lydia 2022/07/15 T大陸案之齊備日管控
   If bolTMFlow = True Then
     If tm(1) <> "" And tm(10) = "020" Then
        If Left(CboEEP04.Text, 2) = EMP_查名結果 Then
           strSql = "select cp10,cp06,cp14,cp149,cp48,ep06,ep34 from caseprogress,engineerprogress where cp09='" & m_EEP01 & "' and cp09=ep02(+) "
           intI = 1
           Set rsTmp = ClsLawReadRstMsg(intI, strSql)
           If intI = 1 Then
               '當承辦歷程送出查名結果，自動上查名齊備日CP143
               If "" & rsTmp.Fields("cp10") = "101" Then
                    '檢查同時文件已齊備和申請進度已有承辦人會更新承辦期限
                    If "" & rsTmp.Fields("cp14") <> "" And Val("" & rsTmp.Fields("ep06")) > 0 And Val("" & rsTmp.Fields("cp48")) = 0 Then
                       strConSql = ""
                       strConSql = Pub_GetHandleDay(tm(1), tm(10), "" & rsTmp.Fields("cp10"), "" & rsTmp.Fields("cp149"), "" & rsTmp.Fields("cp06"), m_EEP01)
                    End If
                    strSql = "UPDATE CaseProgress SET CP143 = " & strSrvDate(1) & IIf(strConSql <> "", ", CP48=" & strConSql, "") & " WHERE CP09 = '" & m_EEP01 & "' "
                    cnnConnection.Execute strSql
               End If
           End If
        End If
     End If 'If tm(1) <> "" And tm(10) = "020" Then
   End If
   'end 2022/07/15
   
   '********************************************************************************
   'Add by Sindy 2023/10/2
   If bolFCPFlow = True Then
      If Left(CboEEP04.Text, 2) = EMP_翻譯交稿 Then
         '若有924.會稿進度同時產生歷程送工程師主管
         If PUB_ChkCPExist(pa, "924", 1, strExc(1), strExc(2), "A") = True Then
            str_924CP09 = strExc(1) '會稿的總收文號
            'Add By Sindy 2024/5/2 翻譯交稿案件'輸入翻譯完稿時,若有會稿未發文，歷程會同時發一道會稿連絡給工程師主管進行分案
            '  請新增判斷，若會稿歷程已有一道聯絡內容為"Claims翻譯已交稿"則不再發聯絡
            strSql = "select * From empelectronprocess where eep01='" & str_924CP09 & "' and eep04='" & EMP_聯絡 & "'" & _
                     " and instr(eep08,'Claims翻譯已交稿，')>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
            '2024/5/2 END
               '新增會稿的:
               '產生會稿承辦單
               If Pub_PrintFCP924Form(pa(1), pa(2), pa(3), pa(4), str_924CP09, m_strColName, m_strColText, , True, m_intColCnt) = True Then
                  strEED05 = Replace(Replace(GetColValues("備註"), "□", ""), "■", "")
                  If GetColValues("譯者") = lblCP10 Then
                     strEED10 = ""
                  ElseIf GetColValues("譯者") <> "" Then
                     strEED10 = GetColValues("譯者")
                  Else
                     strEED10 = ""
                  End If
                  '電子承辦單內容
                  strSql = "select * From EmpElectronData where eed01='" & str_924CP09 & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 0 Then
                     strSql = "insert into EmpElectronData(EED01,EED02,EED04,EED05," & _
                                                          "EED09,EED10) values(" & _
                              CNULL(str_924CP09) & "," & CNULL(GetColValues("受文者")) & "," & CNULL(GetColValues("主旨")) & "," & CNULL(strEED05) & "," & _
                              CNULL(Trim(GetColValues("管制人"))) & "," & CNULL(strEED10) & ")"
                  Else
                     strSql = "update EmpElectronData set" & _
                                    " EED02=" & CNULL(GetColValues("受文者")) & _
                                    ",EED04=" & CNULL(GetColValues("主旨")) & _
                                    ",EED05=" & CNULL(strEED05) & _
                                    ",EED09=" & CNULL(Trim(GetColValues("管制人"))) & _
                                    ",EED10=" & CNULL(strEED10) & _
                              " where EED01='" & str_924CP09 & "'"
                  End If
                  cnnConnection.Execute strSql
               End If
               '取得最大序號
               intMaxEEP02 = 0
               strSql = "select eep02 From empelectronprocess where eep01='" & str_924CP09 & "' order by eep02 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  RsTemp.MoveFirst
                  If RsTemp.RecordCount > 0 Then
                     intMaxEEP02 = RsTemp.Fields(0)
                  End If
               End If
               '新增聯絡歷程
               strExc(9) = GetFCP924txtEEP08("") '另有會稿說明書承辦單，
               '淑華提副本要知會程序管制人
               'Modify By Sindy 2023/12/18 +,eep16
               strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep11,eep12,eep10,eep16) values(" & _
                        CNULL(str_924CP09) & "," & intMaxEEP02 + 1 & "," & CNULL(Trim(txtEEP03)) & "," & _
                        CNULL(EMP_聯絡) & "," & CNULL(Left(m_F21CMMan, 5)) & "," & strSrvDate(1) & "," & _
                        strUpdTime & "," & CNULL(strExc(9)) & ",null," & CNULL(m_EEP12) & "," & CNULL(Left(m_NPMan, 5)) & "," & CNULL(m_EEP16) & ")"
               cnnConnection.Execute strSql
               '寄Mail
               strSubject = PField(1) & "-" & PField(2) & _
                            IIf(PField(3) & PField(4) = "000", "", "-" & PField(3) & "-" & PField(4)) & "(" & lblPA09 & ")(核會流程)-->(" & (intMaxEEP02 + 1) & ")聯絡，請進行後續處理"
               strContent = "本所案號：" & PField(1) & "-" & PField(2) & _
                            IIf(PField(3) & PField(4) = "000", "", "-" & PField(3) & "-" & PField(4)) & vbCrLf & _
                            "案件名稱：" & txtCaseName(0) & vbCrLf & _
                            "申請國家：" & lblPA09 & vbCrLf & _
                            "案件性質：會稿" & vbCrLf & _
                            "流程狀態：聯絡" & vbCrLf & vbCrLf & vbCrLf & _
                            strExc(9) '& MailContentAddEnd("")
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " values( '" & strUserNum & "','" & Left(m_F21CMMan, 5) & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),'" & strSubject & "','" & ChgSQL(strContent) & "'," & CNULL(Left(m_NPMan, 5)) & ")"
               cnnConnection.Execute strSql, intI
            End If
         End If
      ElseIf Left(CboEEP04.Text, 2) = EMP_送排版 Then
         '更新打字室人員
         strSql = "select * from EmpElectronData where EED01='" & m_EEP01 & "'"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strSql = "UPDATE EmpElectronData SET EED06=" & CNULL(Trim(Left(CboEEP05.Text, 6))) & _
                     " where EED01='" & m_EEP01 & "'"
         Else
            strSql = "INSERT INTO EmpElectronData(EED01,EED06) VALUES(" & CNULL(m_EEP01) & "," & CNULL(Trim(Left(CboEEP05.Text, 6))) & ")"
         End If
         cnnConnection.Execute strSql
      ElseIf Left(CboEEP04.Text, 2) = EMP_送轉檔 Then
         '更新打字室人員
         strSql = "select * from EmpElectronData where EED01='" & m_EEP01 & "'"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If "" & rsTmp.Fields("EED06") = "" Then
               strSql = "UPDATE EmpElectronData SET EED06=" & CNULL(Trim(Left(CboEEP05.Text, 6))) & _
                        " where EED01='" & m_EEP01 & "'"
            End If
         Else
            strSql = "INSERT INTO EmpElectronData(EED01,EED06) VALUES(" & CNULL(m_EEP01) & "," & CNULL(Trim(Left(CboEEP05.Text, 6))) & ")"
         End If
         cnnConnection.Execute strSql
      ElseIf Left(CboEEP04.Text, 2) = EMP_交辦 Then
         'F62英文顧問、F71日文顧問、F72德文顧問
         If InStr("F62,F71,F72", Pub_StrUserSt03) > 0 Then
            strSql = "update engineerprogress set" & _
                           " EP03=" & CNULL(Trim(Left(CboEEP05.Text, 6))) & _
                     " where ep02='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
         'M13.打字室
         ElseIf Pub_StrUserSt03 = "M13" Then
            strSql = "UPDATE EmpElectronData SET EED06=" & CNULL(Trim(Left(CboEEP05.Text, 6))) & _
                     " where EED01='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
         End If
         strSql = "UPDATE EmpElectronProcess SET EEP05=" & CNULL(Trim(Left(CboEEP05.Text, 6))) & _
                  " where EEP01='" & m_EEP01 & "' and EEP04='" & m_strLastEEP04 & "' and EEP09='Y'"
         cnnConnection.Execute strSql
      ElseIf Left(CboEEP04.Text, 2) = EMP_核稿分案 Then
         If cp(10) = "201" Then '新案翻譯要更新核稿人
            strSql = "update engineerprogress set" & _
                           " EP04=" & CNULL(Trim(Left(CboEEP05.Text, 6))) & _
                     " where ep02='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
            'Add By Sindy 2024/3/13
            '請在工程師主管去掛核稿人時，若新案翻譯已請款，則系統自動寄給程序人員點數重新分配之信件
            strExc(0) = "select CP09,CP60 from caseprogress where cp09='" & m_EEP01 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Trim("" & RsTemp.Fields("CP60")) <> "" Then
                  '若已開請款單則換承辦人或核稿人時發Mail通知相關人員
                  If Trim("" & RsTemp.Fields("CP60")) > "X" Then
                     Call PUB_PointReAssignInform(PField(1) & "-" & PField(2) & IIf(PField(3) & PField(4) = "000", "", "-" & PField(3) & "-" & PField(4)), "" & RsTemp.Fields("CP60"), , , , Trim(Left(CboEEP05.Text, 6)))
                  End If
               End If
            End If
            '2024/3/13 END
         End If
      End If
   End If
   '2023/10/2 END
   
   '********************************************************************************
   If Left(CboEEP04.Text, 2) = EMP_送判 Or Left(CboEEP04.Text, 2) = EMP_判發 Or _
      Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_發文歸檔 Or _
      (bolTMFlow = True And Left(CboEEP04.Text, 2) = EMP_退件重送) Then
      'Add By Sindy 2018/4/27
      '清空EP42.判發完成日
      If Left(CboEEP04.Text, 2) = EMP_送判 Then
         strSql = "update engineerprogress set" & _
                        " EP42=null" & _
                  " where ep02='" & m_EEP01 & "'"
         cnnConnection.Execute strSql
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                        " EP40='" & Left(m_CSMan, 5) & "'" & _
                        ",EP34=" & CNULL(m_EP34) & _
                        ",EP42=null" & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
         
      '上EP42.判發完成日
      ElseIf Left(CboEEP04.Text, 2) = EMP_判發 Then
         strSql = "update engineerprogress set" & _
                        " EP42=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "'"
         cnnConnection.Execute strSql
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                        " EP42=" & DBDATE(strUpdDate) & _
                  " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
         
         'Add By Sindy 2018/9/27
         If m_strLastEEP04 = EMP_送核 Then '送核-核完時,直接判發,上核稿完成日
            '更新EP39.核稿完成日
            strSql = "update engineerprogress set" & _
                           " EP39=" & DBDATE(strUpdDate) & _
                     " where ep02='" & m_EEP01 & "' and (EP39 is null or EP39=0)"
            cnnConnection.Execute strSql
            
            'Add By Sindy 2020/9/29
            If txtLpNote.Tag = "多案單筆歷程" Then
               strSql = "update engineerprogress set" & _
                           " EP39=" & DBDATE(strUpdDate) & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP39 is null or EP39=0)"
               cnnConnection.Execute strSql
            End If
            '2020/9/29 END
         Else
            'Add By Sindy 2023/10/31 新案翻譯更新核稿完成日
            'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
            If bolFCPFlow = True And (cp(10) = "201" Or cp(10) = "931") Then
               '更新EP39.核稿完成日
               strSql = "update engineerprogress set" & _
                              " EP39=" & DBDATE(strUpdDate) & _
                        " where ep02='" & m_EEP01 & "' and (EP39 is null or EP39=0)"
               cnnConnection.Execute strSql
            End If
            '2023/10/31 END
         End If
         '2018/9/27 END
      
      'Add By Sindy 2023/10/31
      ElseIf bolFCPFlow = True Then
         If Left(CboEEP04.Text, 2) = EMP_送件 Then
            If m_strLastEEP04 = EMP_送判 Then
               '上EP42.判發完成日
               strSql = "update engineerprogress set" & _
                              " EP42=" & DBDATE(strUpdDate) & _
                        " where ep02='" & m_EEP01 & "' and (EP42 is null or EP42=0)"
               cnnConnection.Execute strSql
            ElseIf m_strLastEEP04 = EMP_送核 Then  '直接判發送件
               '更新EP39.核稿完成日
               strSql = "update engineerprogress set" & _
                              " EP39=" & DBDATE(strUpdDate) & _
                        " where ep02='" & m_EEP01 & "' and (EP39 is null or EP39=0)"
               cnnConnection.Execute strSql
            End If
            
            'Add By Sindy 2023/10/31 新案翻譯更新核稿完成日
            'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
            If (cp(10) = "201" Or cp(10) = "931") Then
               '更新EP39.核稿完成日
               strSql = "update engineerprogress set" & _
                              " EP39=" & DBDATE(strUpdDate) & _
                        " where ep02='" & m_EEP01 & "' and (EP39 is null or EP39=0)"
               cnnConnection.Execute strSql
            End If
            '2023/10/31 END
            
            If cp(10) = "1001" And m_PA162 <> "" Then
               strExc(0) = "select cp09 from caseprogress" & _
                           " where cp01='" & PField(1) & "' and cp02='" & PField(2) & "' and cp03='" & PField(3) & "' and cp04='" & PField(4) & "'" & _
                           " and cp10='1917' and cp158=0 and cp159=0 and cp43='" & m_EEP01 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strSubject = ""
                  '1.不需加註分割建議，email通知各區程序上核准發文。
                  '2.分割建議主管上完稿日，email通知各區程序上核准發文。
                  If m_PA162 = "N" Then
                     strSubject = "【工程師已確認不須分割加註】請進行告准 Our Ref: "
                  ElseIf m_PA162 = "Y" Then
                     'Memo by Morgan 2022/10/11 因為日文定稿還是要給工程師核稿,主旨保留以作識別
                     If PUB_GetLanguage(PField(1), PField(2), PField(3), PField(4)) = "3" Then
                        strSubject = "【工程師已完成分割加註(日文定稿)】請進行告准 Our Ref: "
                     Else
                        strSubject = "【工程師已完成分割加註】請進行告准 Our Ref: "
                     End If
                  End If
                  If strSubject <> "" Then
                     strContent = "請告准人員進行後續告准，感謝您。"
                     '1917=通知告准
                     strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " select '" & strUserNum & "' mc01,'" & Pub_GetSpecMan("外專告准程序") & "' mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
                        ",'" & strSubject & "'||cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) mc07,'" & strContent & "' mc08,cp14 mc09" & _
                        " from caseprogress  where cp43='" & m_EEP01 & "' and cp10='1917' and cp27 is null"
                     cnnConnection.Execute strSql, intI
                  End If
               End If
            End If
            
         ElseIf Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
            If intReceiveKind = 0 Then '承辦人工作進度
               If m_PrevForm.txt1(8) = "" Then m_PrevForm.txt1(8) = strUpdDate '發文日
            End If
            strSql = "update caseprogress set" & _
                           " CP27=" & DBDATE(strUpdDate) & _
                     " where cp09='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
         End If
         '2023/10/31 END
      End If
      '2018/4/27 END
      
      'Add By Sindy 2018/4/27 商標處流程 : 送件,退件重送,發文歸檔
      'Modify By Sindy 2024/8/19 + Or bolCFTFlow = True Or bolFCTFlow = True
      If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
         'Add By Sindy 2024/8/19
         If Left(CboEEP04.Text, 2) = EMP_送件 Then
            If m_strLastEEP04 = EMP_送判 Then
               '上EP42.判發完成日
               strSql = "update engineerprogress set" & _
                              " EP42=" & DBDATE(strUpdDate) & _
                        " where ep02='" & m_EEP01 & "' and (EP42 is null or EP42=0)"
               cnnConnection.Execute strSql
            ElseIf m_strLastEEP04 = EMP_送核 Then  '直接判發送件
               '更新EP39.核稿完成日
               strSql = "update engineerprogress set" & _
                              " EP39=" & DBDATE(strUpdDate) & _
                        " where ep02='" & m_EEP01 & "' and (EP39 is null or EP39=0)"
               cnnConnection.Execute strSql
            End If
         End If
         '2024/8/19 END
         
         If Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送 Then
            'Modify By Sindy 2018/8/14
            If cmdCP118.Tag = "Y" Then
               For ii = 1 To frm090202_2_2.MSHFlexGrid2.Rows - 1
                  If frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 0) = "V" And _
                     frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 2) <> "" Then
                     strSql = "update caseprogress set" & _
                              " cp118='Y'" & _
                              ",cp84=" & CNULL(Val(frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 7)), True) & _
                              ",cp85=" & CNULL(strSrvDate(1), True) & _
                              " where cp09='" & frm090202_2_2.MSHFlexGrid2.TextMatrix(ii, 2) & "'"
                     cnnConnection.Execute strSql
                  End If
               Next ii
            End If
            '2018/8/14 END
            
         'Add By Sindy 2018/7/13
         ElseIf Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
            'Modify By Sindy 2018/9/20
            '是否通知客戶
            If ChkEP11.Value = 1 Then  'N.不通知
'               strSql = "update caseprogress set" & _
'                        " cp27=19221111" & _
'                        " where cp09='" & m_EEP01 & "' and (cp27 is null or cp27=0)"
'               cnnConnection.Execute strSql
               m_PrevForm.txt1(9) = "N"
               If m_PrevForm.txt1(8) = "" Then m_PrevForm.txt1(8) = "111111" '發文日
               'Add By Sindy 2020/9/29
               If txtLpNote.Tag = "多案單筆歷程" Then
                  strSql = "update engineerprogress set" & _
                                 " EP11='N'" & _
                           " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
                  Pub_SeekTbLog strSql 'Add By Sindy 2021/6/28
                  cnnConnection.Execute strSql
                  strSql = "update caseprogress set" & _
                                 " CP27=19221111" & _
                           " where cp09 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
                  cnnConnection.Execute strSql
               End If
               '2020/9/29 END
            Else 'Y.通知
'               strSql = "update caseprogress set" & _
'                        " cp27=" & DBDATE(strUpdDate) & _
'                        " where cp09='" & m_EEP01 & "' and (cp27 is null or cp27=0)"
'               cnnConnection.Execute strSql
               m_PrevForm.txt1(9) = "Y"
               If m_PrevForm.txt1(8) = "" Then m_PrevForm.txt1(8) = strUpdDate '發文日
               'Add By Sindy 2020/9/29
               If txtLpNote.Tag = "多案單筆歷程" Then
                  strSql = "update engineerprogress set" & _
                                 " EP11='Y'" & _
                           " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
                  Pub_SeekTbLog strSql 'Add By Sindy 2021/6/28
                  cnnConnection.Execute strSql
                  strSql = "update caseprogress set" & _
                                 " CP27=" & DBDATE(strUpdDate) & _
                           " where cp09 in('" & Replace(m_RetrunRecvSub, ",", "','") & "')"
                  cnnConnection.Execute strSql
               End If
               '2020/9/29 END
            End If
            
            'Add By Sindy 2020/2/7
            'Modify By Sindy 2021/3/4 + And Left(cp(12), 1) <> "F"
            '外商收文,尚不新增信函資料
            If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(cp(12), 1) <> "F" Then
               'Modify By Sindy 2022/3/3 無條件重新刪除LP新增LP; + IIf(m_PrevForm.txt1(9) = "Y", True, False)
'               '檢查有沒有信函進度,若無則新增
'               strSql = "update letterprogress set lp06='" & PUB_GetAKindSalesNo(PField(1), PField(2), PField(3), PField(4)) & "'" & _
'                        " where lp01='" & m_EEP01 & "'"
'               cnnConnection.Execute strSql, intI
'               If intI = 0 Then
                  PUB_AddLetterProgress m_EEP01, 0, IIf(m_PrevForm.txt1(9) = "Y", True, False), , bolRegMail, m_PA26, cp(10), m_PA75
'               End If
               '通知客戶
               If m_PrevForm.txt1(9) = "Y" Then
                  '若有齊備日lp03時要上判發日lp05
                  strSql = "update letterprogress set lp04=null,lp05=decode(lp03,0,0," & strSrvDate(1) & ")" & _
                           ",lp08='" & strUserNum & "',lp09=" & strSrvDate(1) & ",lp10='Y'" & _
                           ",lp11='" & IIf(bolRegMail = True, "Y", "") & "',LP31='" & strLP31 & "'" & _
                           " where lp01='" & m_EEP01 & "'"
               Else
                  strExc(10) = "不通知客戶;"
                  strSql = "update letterprogress set lp04=null,lp05=decode(lp03,0,0," & strSrvDate(1) & ")" & _
                           ",lp06='" & strUserNum & "',lp07=" & strSrvDate(1) & _
                           ",lp08='" & strUserNum & "',lp09=" & strSrvDate(1) & ",lp10='N'" & _
                           ",lp11='" & IIf(bolRegMail = True, "Y", "") & "',LP31='" & strLP31 & "'" & _
                           ",lp12='" & strExc(10) & "'||replace(lp12,'" & strExc(10) & "','')" & _
                           " where lp01='" & m_EEP01 & "'"
               End If
               cnnConnection.Execute strSql, intI
               
               '其他文號合併至此文號,通知客戶
               If m_RetrunRecvSub <> "" Then
                  arrID = Split(m_RetrunRecvSub, ",")
                  For intCnt = 0 To UBound(arrID)
                     'Modify By Sindy 2022/3/3 無條件重新刪除LP新增LP; + IIf(m_PrevForm.txt1(9) = "Y", True, False)
'                     '檢查有沒有信函進度,若無則新增
'                     strSql = "update letterprogress set lp06='" & PUB_GetAKindSalesNo(PField(1), PField(2), PField(3), PField(4)) & "'" & _
'                              " where lp01='" & arrID(intCnt) & "'"
'                     cnnConnection.Execute strSql, intI
'                     If intI = 0 Then
                        PUB_AddLetterProgress CStr(arrID(intCnt)), 0, IIf(m_PrevForm.txt1(9) = "Y", True, False), , bolRegMail, m_PA26, cp(10), m_PA75
'                     End If
                     '通知客戶
                     If m_PrevForm.txt1(9) = "Y" Then
                        strExc(10) = "已併入" & lblCP10 & "通知函(" & IIf(PField(3) & PField(4) = "000", PField(1) & "-" & PField(2), PField(1) & "-" & PField(2) & "-" & PField(3) & "-" & PField(4)) & ":" & m_EEP01 & ")告知客戶;"
                        strSql = "update letterprogress set" & _
                                 " lp03=" & strSrvDate(1) & ",lp06='" & strUserNum & "',lp07=" & strSrvDate(1) & _
                                 ",lp10='N',lp11='" & IIf(bolRegMail = True, "Y", "") & "',LP31='" & strLP31 & "'" & _
                                 ",lp12='" & strExc(10) & "'||replace(lp12,'" & strExc(10) & "',''),lp42='" & m_EEP01 & "'" & _
                                 " where lp01='" & arrID(intCnt) & "'"
                     Else
                        strExc(10) = "不通知客戶;"
                        strSql = "update letterprogress set" & _
                                 " lp03=" & strSrvDate(1) & ",lp06='" & strUserNum & "',lp07=" & strSrvDate(1) & _
                                 ",lp08='" & strUserNum & "',lp09=" & strSrvDate(1) & ",lp10='N'" & _
                                 ",lp11='" & IIf(bolRegMail = True, "Y", "") & "',LP31='" & strLP31 & "'" & _
                                 ",lp12='" & strExc(10) & "'||replace(lp12,'" & strExc(10) & "',''),lp42='" & m_EEP01 & "'" & _
                                 " where lp01='" & arrID(intCnt) & "'"
                     End If
                     cnnConnection.Execute strSql, intI
                  Next intCnt
               End If
            End If
         End If
         
         'Add By Sindy 2018/9/25
         If Frame6.Visible = True Then
            '條款
            m_PrevForm.txt1(11) = txt2
            '預估結果
            strCP23 = ""
            If Option1(0).Value = True Then
               strCP23 = "1"
            ElseIf Option1(1).Value = True Then
               strCP23 = "2"
            ElseIf Option1(2).Value = True Then
               strCP23 = "3"
            End If
            strSql = "update caseprogress" & _
                     " set cp23=" & CNULL(strCP23) & ",cp49=" & CNULL(txt2) & _
                     " where cp09='" & m_EEP01 & "'"
            cnnConnection.Execute strSql
         End If
         '2018/9/25 END
      End If
      '2018/4/27 END
      
      '更新EP09.完稿日
      If intReceiveKind = 0 Then '0.承辦人工作進度
         If Trim(m_PrevForm.txt1(3)) = "" Then
            m_PrevForm.txt1(3) = strUpdDate
         End If
         'Modify By Sindy 2021/1/5 mark
'         'Add By Sindy 2018/9/27 商標處在送判,判發時還不是進程序的最終動作,所以此段程式還不可執行
'         If Not (bolTMFlow = True And (Left(CboEEP04.Text, 2) = EMP_送判 Or _
'                                       Left(CboEEP04.Text, 2) = EMP_判發)) Then
'         '2018/9/27 END
         '2021/1/5 END
            'Add By Sindy 2018/5/24
            If m_EP34 = "N" And bolFCPFlow = False Then
               'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
               If bolTMFlow = True Or bolOtherFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
                  '更新EP08.會稿完成日
                  If Trim(m_PrevForm.txt1(7)) = "" Then m_PrevForm.txt1(7) = strUpdDate
                  
                  'Add By Sindy 2020/9/29
                  If txtLpNote.Tag = "多案單筆歷程" Then
                     strSql = "update engineerprogress set" & _
                                    " EP08=" & DBDATE(strUpdDate) & _
                                    ",EP34='N'" & _
                              " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP08 is null or EP08=0)"
                     cnnConnection.Execute strSql
                     
                     'Add By Sindy 2021/1/5
                     strSql = "update engineerprogress set" & _
                                    " EP09=" & DBDATE(strUpdDate) & _
                              " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP09 is null or EP09=0)"
                     cnnConnection.Execute strSql
                     '2021/1/5 END
                  End If
                  '2020/9/29 END
               Else
                  UpdateEp08 m_EEP01, DBDATE(strUpdDate) '更新相關會稿完成日資料
               End If
               '更新EP07.會稿日
               If Trim(m_PrevForm.txt1(4)) = "" Then m_PrevForm.txt1(4) = strUpdDate
               'Add By Sindy 2020/9/29
               If txtLpNote.Tag = "多案單筆歷程" Then
                  strSql = "update engineerprogress set" & _
                                 " EP07=" & DBDATE(strUpdDate) & _
                           " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP07 is null or EP07=0)"
                  cnnConnection.Execute strSql
               End If
               '2020/9/29 END
            End If
            '2018/5/24 END
'         End If
      Else
         strSql = "update engineerprogress set" & _
                        " EP09=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "' and (EP09 is null or EP09=0)"
         cnnConnection.Execute strSql
         'Add By Sindy 2020/9/29
         If txtLpNote.Tag = "多案單筆歷程" Then
            strSql = "update engineerprogress set" & _
                           " EP09=" & DBDATE(strUpdDate) & _
                     " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP09 is null or EP09=0)"
            cnnConnection.Execute strSql
         End If
         '2020/9/29 END
         
         'Modify By Sindy 2021/1/5 mark
'         'Add By Sindy 2018/9/27 商標處在送判,判發時還不是進程序的最終動作,所以此段程式還不可執行
'         If Not (bolTMFlow = True And (Left(CboEEP04.Text, 2) = EMP_送判 Or _
'                                       Left(CboEEP04.Text, 2) = EMP_判發)) Then
'         '2018/9/27 END
         '2021/1/5 END
            'Add By Sindy 2013/10/17 不會稿時,自行判發者須一併更新會稿日及會稿完成日
            If m_EP34 = "N" And bolFCPFlow = False Then
               'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
               If bolTMFlow = True Or bolOtherFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
                  '更新EP08.會稿完成日
                  strSql = "update engineerprogress set" & _
                                 " EP08=" & DBDATE(strUpdDate) & _
                           " where ep02='" & m_EEP01 & "' and (EP08 is null or EP08=0)"
                  cnnConnection.Execute strSql
                  
                  'Add By Sindy 2020/9/29
                  If txtLpNote.Tag = "多案單筆歷程" Then
                     strSql = "update engineerprogress set" & _
                                    " EP08=" & DBDATE(strUpdDate) & _
                                    ",EP34='N'" & _
                              " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP08 is null or EP08=0)"
                     cnnConnection.Execute strSql
                  End If
                  '2020/9/29 END
               Else
                  UpdateEp08 m_EEP01, DBDATE(strUpdDate) '更新相關會稿完成日資料
               End If
               '更新EP07.會稿日
               strSql = "update engineerprogress set" & _
                              " EP07=" & DBDATE(strUpdDate) & _
                        " where ep02='" & m_EEP01 & "' and (EP07 is null or EP07=0)"
               cnnConnection.Execute strSql
         '         If m_PrevForm.intBackTab = 1 Then
         '            m_PrevForm.txt1(4) = strUpdDate
         '         End If
               'Add By Sindy 2020/9/29
               If txtLpNote.Tag = "多案單筆歷程" Then
                  strSql = "update engineerprogress set" & _
                                 " EP07=" & DBDATE(strUpdDate) & _
                           " where ep02 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and (EP07 is null or EP07=0)"
                  cnnConnection.Execute strSql
               End If
               '2020/9/29 END
            End If
            '2013/10/17 END
'         End If
      End If
      
      'Add By Sindy 2014/7/14 當承辦人為專利處繪圖的人員且直接判發時,更新草圖/墨圖完稿日
      If PUB_GetStaffST15(Left(m_EPMan, 5), "1") = "P13" And Left(CboEEP04.Text, 2) = EMP_判發 Then
         strSql = "update engineerprogress set" & _
                        " EP15=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "' and (EP15 is null or EP15=0)"
         cnnConnection.Execute strSql
         strSql = "update engineerprogress set" & _
                        " EP18=" & DBDATE(strUpdDate) & _
                  " where ep02='" & m_EEP01 & "' and (EP18 is null or EP18=0)"
         cnnConnection.Execute strSql
         If UCase(m_PrevForm.Name) = UCase("frm090711") Then
            If Trim(m_PrevForm.txt1(2)) = "" Then
               m_PrevForm.txt1(2) = strUpdDate
            End If
            If Trim(m_PrevForm.txt1(5)) = "" Then
               m_PrevForm.txt1(5) = strUpdDate
            End If
         End If
      End If
      '2014/7/14 END
   End If
   '********************************** END *****************************************
End Sub

'Modify By Sindy 2021/10/7
Private Function MailContentAddEnd(strContent As String) As String
   'Add By Sindy 2024/8/13
   If bolCFTFlow = True Or bolFCTFlow = True Then
      MailContentAddEnd = strContent & vbCrLf & vbCrLf & vbCrLf & _
                     "請至系統的下列位置進行：" & vbCrLf & _
                     "（國外部商標系統）" & vbCrLf & _
                     " 承　辦　人　員 ：歷程工作->工作進度資料維護->待辦歷程" & vbCrLf & _
                     " 核　判　人　員 ：歷程工作->待核判區" & vbCrLf
      MailContentAddEnd = MailContentAddEnd & _
                     "（承辦人系統）" & vbCrLf & _
                     " 智　權　人　員 ：智權部->專利商標作業->專利／商標會稿" & vbCrLf & vbCrLf
   'Add By Sindy 2023/10/18
   ElseIf bolFCPFlow = True Then
      MailContentAddEnd = strContent & vbCrLf & vbCrLf & vbCrLf & _
                     "請至系統的下列位置進行：" & vbCrLf & _
                     "（國外部專利及承辦人系統）" & vbCrLf & _
                     " 承　辦　人　員 ：歷程工作->工作進度資料維護->待辦歷程" & vbCrLf & _
                     " 核　判　人　員 ：歷程工作->待核判區" & vbCrLf & _
                     "（承辦人系統）" & vbCrLf & _
                     " 承　辦　人　員 ：國外部->工作進度資料維護->待辦歷程" & vbCrLf & _
                     " 核　判　人　員 ：國外部->待核判區" & vbCrLf & _
                     "（檔案室系統）" & vbCrLf & _
                     " 打 字 室 人 員 ：打字室->待排版區" & vbCrLf & vbCrLf
   Else
   '2023/10/18 END
      If PField(1) = "ACS" Then
         MailContentAddEnd = strContent & vbCrLf & vbCrLf & vbCrLf & _
                     "請至系統的下列位置進行：" & vbCrLf & vbCrLf & _
                     " 承　辦　人　員 ：法務->ＡＣＳ->承辦人工作->工作進度資料維護->待辦歷程" & vbCrLf & _
                     " 核　判　人　員 ：法務->ＡＣＳ->承辦人工作->待核判區" & vbCrLf
      Else
         MailContentAddEnd = strContent & vbCrLf & vbCrLf & vbCrLf & _
                     "請至系統的下列位置進行：" & vbCrLf & vbCrLf & _
                     " 承　辦　人　員 ：承辦人->工作進度資料維護->待辦歷程" & vbCrLf & _
                     " 核　判　人　員 ：承辦人->待核判區" & vbCrLf
      End If
      MailContentAddEnd = MailContentAddEnd & _
                     " 智　權　人　員 ：智權部->專利商標作業->專利／商標會稿" & vbCrLf
   End If
   MailContentAddEnd = MailContentAddEnd & _
               " 法　務　系　統 ：會稿判發->專利／商標會稿" & vbCrLf & _
               " 　　　　　　　　           　　　待核判區" & vbCrLf & _
               " 副本收受者,聯絡：共同查詢->案件查詢->案件資料及進度查詢->承辦歷程(聯絡)"
End Function

'Modify By Sindy 2018/10/23 發通知信
'f_bolSendMail : 是否發通知信
'f_strAutoFlow : 附加流程的流程狀態(判發,送判,送件...) / 多案件時,新增的歷程順序
'f_bolManyCaseSingleEEP : 為多案單筆歷程 Add By Sindy 2020/10/19
'f_strSubRecv : 多案件時,總收文號
'Modify By Sindy 2025/8/5 Private Function ==> Public Function
Public Function FlowSendMail(f_bolSendMail As Boolean, f_strAutoFlow As String, _
   f_bolManyCaseSingleEEP As Boolean, Optional f_strSubRecv As String) As Boolean
Dim strTo As String 'Add By Sindy 2018/7/23
Dim stVTB As String
Dim rsQuery As New ADODB.Recordset
Dim strCaseName As String, strSubCaseNo As String
Dim strSubEEP01 As String, strSubEEP02 As String
Dim strSubCP01 As String, strSubCP02 As String, strSubCP03 As String, strSubCP04 As String
Dim strSubEP01 As String
Dim strSubCP13 As String, strSubCP14 As String
Dim strSubPA75 As String, strSubPA09 As String, strSubPA77 As String, strSubPA48 As String
Dim strSubCP10 As String, strSubCP10Nm As String
Dim strAttPath As String, stFileName As String 'Add By Sindy 2019/6/28
Dim dblSize As Double '檔案大小
Dim strTemp As String
Dim strMailTextCom As String
Dim strSubPA09Nm As String 'Add By Sindy 2025/4/2
   
   FlowSendMail = False
   
   If f_strSubRecv = "" Then 'Modify By Sindy 2018/10/23 多案件時,最後一筆才一併寄出信件
      '********************************************************************************
      'Add By Sindy 2018/6/20 集中發信
      If intReceiveKind = 0 Then '0.承辦人工作進度
         If m_PrevForm.m_chkcmdok1 = True Then
            m_PrevForm.BatctMail
         End If
      End If
      '2018/6/20 END
      '********************************** END *****************************************
      PUB_SendMailCache '發郵件 (相關會稿完成日的郵件)
   End If
   strSubEEP01 = m_EEP01
   strSubEEP02 = m_EEP02
   
   'Add By Sindy 2019/6/28 有副本收受者時,要夾帶附件
   'Modify By Sindy 2023/9/28 + 排除ACS不夾帶電子檔
   If Trim(txtEEP10) <> "" And PField(1) <> "ACS" Then
      ChkEMail.Visible = True: ChkEMail.Value = 1
   End If
   '2019/6/28 END
   'E-Mail夾帶附件
   strSendFilePath = "": dblSize = 0
   If ChkEMail.Visible = True And ChkEMail.Value = 1 Then
      '多個檔案時中間加*分隔
      For ii = 0 To lstAtt(0).ListCount - 1
         'Add By Sindy 2019/9/16
         strTemp = Mid(lstAtt(0).List(ii), 1, InStr(lstAtt(0).List(ii), "KB)") - 1)
         dblSize = dblSize + Mid(strTemp, InStrRev(strTemp, "(") + 1)
         '2019/9/16 END
         strSendFilePath = strSendFilePath & Left(lstAtt(0).List(ii), InStrRev(lstAtt(0).List(ii), " (") - 1) & "*"
      Next ii
      'Add By Sindy 2019/6/28 抓存卷區電子檔
      If Left(CboEEP04.Text, 2) = EMP_查名結果 Then
         '下載存卷附件
         strAttPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum & "\otherFile"
         If Dir(strAttPath, vbDirectory) = "" Then
            MkDir strAttPath
         End If
         If Dir(strAttPath & "\.") <> "" Then
            Kill strAttPath & "\*.*"
         End If
         For ii = 0 To lstAtt(1).ListCount - 1
            stFileName = lstAtt(1).List(ii)
            'Modify By Sindy 2020/5/15 增加檢查上傳電子檔是操作當事人才需要下載
            'EX:T-227357
            If InStr(stFileName, strUserName) > 0 Then
            '2020/5/15 END
               If InStrRev(stFileName, " (") > 0 Then
                  'Add By Sindy 2019/9/16
                  strTemp = Mid(lstAtt(1).List(ii), 1, InStr(lstAtt(1).List(ii), "KB)") - 1)
                  dblSize = dblSize + Mid(strTemp, InStrRev(strTemp, "(") + 1)
                  '2019/9/16 END
                  stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
               End If
               If InStr(stFileName, "\") = 0 Then
                  'If GetAttachFile(stFileName, CInt(0), strAttPath & "\" & stFileName) = True Then
                  If PUB_GetAttachFile_EEF(m_EEP01, CInt(0), stFileName, strAttPath & "\" & stFileName, True) = True Then
                     strSendFilePath = strSendFilePath & stFileName & "*"
                  End If
               End If
            End If
         Next ii
      End If
      '2019/6/28 END
      If strSendFilePath <> "" Then strSendFilePath = Left(strSendFilePath, Len(strSendFilePath) - 1)
   End If
   
   'Add By Sindy 2018/10/23 多案件時,有傳入總收文號就要重覆讀取案件資料,反之則用畫面上原值
   If f_strSubRecv <> "" Then '多案
      strSubEEP01 = f_strSubRecv
'      If strAutoFlow <> "" Then
'         If IsNumeric(strAutoFlow) = True Then
            strSubEEP02 = f_strAutoFlow
'         End If
'      End If
      strSql = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",pa48 客戶案件案號,pa11 申請案號,pa05 案件名稱,NA03 申請國家,nvl(DECODE(Pa09,'000',cpm03,cpm04),cp10) AS 案件性質" & _
         ",pa26,pa27,pa75,pa09,pa149,cp01,cp02,cp03,cp04,'1' 案件種類,pa77,cp10,'' 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,EP01,CP13,CP14" & _
         " From caseprogress, patent, nation, casepropertymap, engineerprogress" & _
         " where cp09='" & strSubEEP01 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & _
         " and na01(+)=pa09 and cp09=ep02(+)" & _
         " and cp01=cpm01(+) and cp10=cpm02(+)"
      strSql = strSql & " union " & _
         "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",tm35 客戶案件案號,tm12 申請案號,tm05 案件名稱,NA03 申請國家,nvl(DECODE(tm10,'000',cpm03,cpm04),cp10) AS 案件性質" & _
         ",tm23,tm78,tm44,tm10,tm123,cp01,cp02,cp03,cp04,'2' 案件種類,tm45,cp10,tm09 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,EP01,CP13,CP14" & _
         " From caseprogress, trademark, nation, casepropertymap, engineerprogress" & _
         " where cp09='" & strSubEEP01 & "' and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
         " and na01(+)=tm10 and cp09=ep02(+)" & _
         " and cp01=cpm01(+) and cp10=cpm02(+)"
      strSql = strSql & " union " & _
         "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",sp29 客戶案件案號,sp11 申請案號,sp05 案件名稱,NA03 申請國家,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質" & _
         ",sp08,sp58,sp26,sp09,sp78,cp01,cp02,cp03,cp04,'5' 案件種類,sp27,cp10,'' 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,EP01,CP13,CP14" & _
         " From caseprogress, servicepractice, nation, casepropertymap, engineerprogress" & _
         " where cp09='" & strSubEEP01 & "' and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
         " and na01(+)=sp09 and cp09=ep02(+)" & _
         " and cp01=cpm01(+) and cp10=cpm02(+)"
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         With rsQuery
            strSubCP01 = rsQuery.Fields("cp01")
            strSubCP02 = rsQuery.Fields("cp02")
            strSubCP03 = rsQuery.Fields("cp03")
            strSubCP04 = rsQuery.Fields("cp04")
            strSubCaseNo = strSubCP01 & "-" & strSubCP02 & "-" & strSubCP03 & "-" & strSubCP04
            strSubEP01 = rsQuery.Fields("ep01")
            strSubCP13 = rsQuery.Fields("cp13")
            strSubCP14 = rsQuery.Fields("cp14")
            strSubPA75 = "" & rsQuery.Fields("PA75")
            strSubPA09 = "" & rsQuery.Fields("PA09")
            strSubPA09Nm = "" & rsQuery.Fields("申請國家") 'Add By Sindy 2025/4/2
            strSubPA77 = "" & rsQuery.Fields("PA77")
            strSubPA48 = "" & rsQuery.Fields("客戶案件案號")
            strSubCP10 = rsQuery.Fields("cp10")
            strSubCP10Nm = rsQuery.Fields("案件性質")
            strCaseName = rsQuery.Fields("案件名稱")
         End With
      Else
         MsgBox "找不到案件資料[FlowSendMail], 請洽電腦中心!!"
         Exit Function
      End If
      
   Else
      strSubCP01 = PField(1)
      strSubCP02 = PField(2)
      strSubCP03 = PField(3)
      strSubCP04 = PField(4)
      strSubCaseNo = lblCaseNo.Caption
      strSubEP01 = m_EP01
      strSubCP13 = Trim(Left(m_SPMan, 6))
      strSubCP14 = Trim(Left(m_EPMan, 6))
      strSubPA75 = m_PA75
      strSubPA09 = m_Country
      strSubPA09Nm = lblPA09 'Add By Sindy 2025/4/2
      strSubPA77 = m_PA77
      strSubPA48 = m_PA48
      strSubCP10 = cp(10)
      strSubCP10Nm = m_CP10Nm
      'Add By Sindy 2013/10/1
      If Trim(txtCaseName(0)) <> "" Then
         strCaseName = Trim(txtCaseName(0))
      ElseIf Trim(txtCaseName(1)) <> "" Then
         strCaseName = Trim(txtCaseName(1))
      Else
         strCaseName = Trim(txtCaseName(2))
      End If
      '2013/10/1 END
   End If
   '2018/10/23 END
   
   'Add By Sindy 2014/1/16
   'Modify By Sindy 2024/12/5 + And bolFCTFlow = False
   If Trim(Left(CboEEP05.Text, 6)) = strSubCP13 And bolFCTFlow = False Then
      'Modify By Sindy 2019/9/5 + 會稿方式 IIf(Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) = "客戶會稿", "(" & Right(CboCP10.Text, Len(CboCP10.Text) - 2) & ")", "")
      strSubject = ""
      If Trim(CboCP10.Text) <> "" Then
         strSubject = Right(CboCP10.Text, Len(CboCP10.Text) - 2)
      End If
      'Modify By Sindy 2020/9/29 + IIf(txtLpNote.Tag = "多案單筆歷程", "~" & txtLpNote.Text, "")
      strSubject = Replace(strSubCaseNo, "-0-00", "") & IIf(f_bolManyCaseSingleEEP = True, "~" & txtLpNote.Text, "") & "「" & strCaseName & "」-->" & Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) & IIf(Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) = "客戶會稿", "(" & strSubject & ")", "")
   Else
   '2014/1/16 END
      'If strAutoFlow = "" Then
      If f_strAutoFlow = "" Or f_strSubRecv <> "" Then
         'm_EEP02 ==> strSubEEP02
         'Modify By Sindy 2019/9/5 + 會稿方式 IIf(Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) = "客戶會稿", "(" & Right(CboCP10.Text, Len(CboCP10.Text) - 2) & ")", "")
         strSubject = ""
         If Trim(CboCP10.Text) <> "" Then
            strSubject = Right(CboCP10.Text, Len(CboCP10.Text) - 2)
         End If
         'Modify By Sindy 2020/9/29 + IIf(txtLpNote.Tag = "多案單筆歷程", "~" & txtLpNote.Text, "")
         'Modify By Sindy 2023/12/11 + IIf(Left(CboEEP04.Text, 2) = EMP_轉檔完成 And ChkEED13.Value = 1, "送件", "")
         strSubject = Replace(strSubCaseNo, "-0-00", "") & _
                      IIf(f_bolManyCaseSingleEEP = True, "~" & txtLpNote.Text, "") & _
                      "(" & strSubPA09Nm & ")(核會流程)-->(" & strSubEEP02 & ")" & Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) & _
                      IIf(Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) = "客戶會稿", "(" & strSubject & ")", "") & _
                      IIf(Left(CboEEP04.Text, 2) = EMP_轉檔完成 And ChkEED13.Value = 1, "送件", "") & _
                      "，請進行後續處理"
         '2019/9/5 END
      Else
         'Modify By Sindy 2020/9/29 + IIf(txtLpNote.Tag = "多案單筆歷程", "~" & txtLpNote.Text, "")
         'Modify By Sindy 2023/12/14 杜燕文協理請作,主旨加申請國家
         strSubject = Replace(strSubCaseNo, "-0-00", "") & IIf(f_bolManyCaseSingleEEP = True, "~" & txtLpNote.Text, "") & "(" & strSubPA09Nm & ")(核會流程)-->" & f_strAutoFlow & "，請進行後續處理"
      End If
   End If
'   'Modify By Sindy 2024/11/28 薛經理反應主旨中的"已送件"易混淆,因此改字樣
'   'ex:FCP-064712(台灣)(核會流程)-->(2)送件
'   If InStr(strSubject, ")送件") > 0 And Left(CboEEP04.Text, 2) = EMP_送件 Then
'      strSubject = Replace(strSubject, ")送件", ")準備送件")
'   End If
'   '2024/11/28 END
   
   'Add By Sindy 2014/1/16
   strContent = ""
   If Trim(Left(CboEEP05.Text, 6)) <> strSubCP13 Then
      strContent = "當月目次：" & strSubEP01 & vbCrLf
   End If
   '2014/1/16 END
   'Modify By Sindy 2018/4/2
   'FC代理人來台
   If strSubPA75 <> "" And strSubPA09 = "000" Then
      strContent = strContent & "貴方卷號：" & strSubPA77 & vbCrLf
   End If
   '2018/4/2 END
   strContent = strContent & "本所案號：" & strSubCaseNo & vbCrLf
   'Add By Sindy 2014/3/4 送會歷程的郵件內容, 加客戶案件案號
   'Modify By Sindy 2016/3/15 + or Left(CboEEP04.Text, 2) = EMP_會圖
   If (Left(CboEEP04.Text, 2) = EMP_送會 Or Left(CboEEP04.Text, 2) = EMP_會圖) And Trim(strSubPA48) <> "" Then
      strContent = strContent & "客戶案件案號：" & strSubPA48 & vbCrLf
   End If
   '2014/3/4 END
   'Modify By Sindy 2023/12/15 杜燕文協理請作,內文加申請國家
   strContent = strContent & _
                "案件名稱：" & strCaseName & vbCrLf & _
                "案件性質：" & strSubCP10Nm & vbCrLf & _
                "申請國家：" & strSubPA09Nm & vbCrLf & _
                "流程狀態：" & Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) & _
                IIf(Left(CboEEP04.Text, 2) = EMP_附加流程 And CboCP10.Visible = True And Trim(CboCP10.Text) <> "", " - " & Trim(Mid(CboCP10.Text, 5)), "") & vbCrLf
   If Trim(txtEEP08) <> "" Then
      strContent = strContent & "內　　容：" & Trim(txtEEP08) & vbCrLf
   End If
   'Add By Sindy 2025/8/5
   If Left(CboEEP04.Text, 2) = EMP_會完 And InStr(txtEEP08, "已收文分析") > 0 Then
      strSubject = strSubject & "【已收文分析】"
   ElseIf Left(CboEEP04.Text, 2) = EMP_送會 And GetCP43AddCC(False) = True Then
      strSubject = strSubject & "【於收文[分析]程序後，今已進行會稿程序，特通知】"
   End If
   '2025/8/5 END
   
   'Add By Sindy 2013/9/18
   'Modify By Sindy 2018/7/19 + And bolPAFlow = True
   If Left(CboEEP04.Text, 2) = EMP_會完 And bolPAFlow = True Then
      strContent = strContent & vbCrLf & vbCrLf & "★智權同仁已輸入會稿完成日，請至待辦歷程中確認該會稿完成日。" & vbCrLf
   End If
   '2013/9/18 END
   
   'Add By Sindy 2013/10/2 加到送會時EMail裡提示
   If Left(CboEEP04.Text, 2) = EMP_送會 Then
      'Modify By Sindy 2016/1/13
      'If (CP(10) = "101" Or CP(10) = "102") And m_country = "000" Then
      If bolPAFlow = True Then
      '2016/1/13 END
         'Modify By Sindy 2015/12/2
         If strSubCP10 = "101" And strSubPA09 = "000" Then
'         strContent = strContent & vbCrLf & vbCrLf & _
'                      "請注意：本案若同時或隨後可能申請大陸專利，請留意是否有超頁超項問題：" & vbCrLf & _
'                      "　　　　1.專利說明書(含申請專利範圍、圖式)以30頁為限，每增加1頁加收新台幣500元。" & vbCrLf & _
'                      "　　　　2.申請專利範圍以10項為限，每增加1項加收新台幣1000元。" & _
'                      vbCrLf
            strContent = strContent & vbCrLf & vbCrLf & _
                         "請注意：發明案請留意是否有超項(10項)，若有，則每超出1項須加收規費新台幣800元。" & vbCrLf & _
                         "　　　　若同時或隨後可能申請大陸專利，請留意是否有超頁超項問題：" & vbCrLf & _
                         "　　　　1.專利說明書(含圖式)以30頁為限，每增加1頁加收新台幣500元。" & vbCrLf & _
                         "　　　　2.申請專利範圍以10項為限，每增加1項加收新台幣1000元。" & _
                         vbCrLf
         End If
         '2015/12/2 END
         
         'Modify By Sindy 2019/5/16 玲玲說要工程師送會時含data.doc給智權人員會稿
'         'Add By Sindy 2018/9/3 P案會稿時帶出申請資訊
'         If strSubCP01 = "P" And InStr(NewCasePtyList, strSubCP10) > 0 Then
'            strContent = strContent & vbCrLf & vbCrLf & _
'                         "*煩請確認以下申請人及發明人資訊，若有任何變動，請務必通知P程序人員，以利更新系統資訊，謝謝！" & _
'                         vbCrLf
'            strContent = strContent & vbCrLf & GetApplData & vbCrLf
'         End If
'         '2018/9/3 END
      End If
      'Add By Sindy 2025/8/5
      If GetCP43AddCC(False) = True Then
         strContent = strContent & vbCrLf & _
                         "【於收文[分析]程序後，今已進行會稿程序，特通知】" & _
                         vbCrLf
      End If
      '2025/8/5 END
   End If
   '2013/10/2 END
   
   '************************************************************************************************
   'Add By Sindy 2020/9/29 組多案的子案內容
   '************************************************************************************************
   If f_bolManyCaseSingleEEP = True Then
      strSql = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",pa48 客戶案件案號,pa11 申請案號,pa05 案件名稱,NA03 申請國家,nvl(DECODE(Pa09,'000',cpm03,cpm04),cp10) AS 案件性質" & _
         ",pa26,pa27,pa75,pa09,pa149,cp01,cp02,cp03,cp04,'1' 案件種類,pa77,cp10,'' 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,EP01,CP13,CP14" & _
         " From caseprogress, patent, nation, casepropertymap, engineerprogress" & _
         " where cp09 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & _
         " and na01(+)=pa09 and cp09=ep02(+)" & _
         " and cp01=cpm01(+) and cp10=cpm02(+)"
      strSql = strSql & " union " & _
         "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",tm35 客戶案件案號,tm12 申請案號,tm05 案件名稱,NA03 申請國家,nvl(DECODE(tm10,'000',cpm03,cpm04),cp10) AS 案件性質" & _
         ",tm23,tm78,tm44,tm10,tm123,cp01,cp02,cp03,cp04,'2' 案件種類,tm45,cp10,tm09 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,EP01,CP13,CP14" & _
         " From caseprogress, trademark, nation, casepropertymap, engineerprogress" & _
         " where cp09 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
         " and na01(+)=tm10 and cp09=ep02(+)" & _
         " and cp01=cpm01(+) and cp10=cpm02(+)"
      strSql = strSql & " union " & _
         "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",sp29 客戶案件案號,sp11 申請案號,sp05 案件名稱,NA03 申請國家,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質" & _
         ",sp08,sp58,sp26,sp09,sp78,cp01,cp02,cp03,cp04,'5' 案件種類,sp27,cp10,'' 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,EP01,CP13,CP14" & _
         " From caseprogress, servicepractice, nation, casepropertymap, engineerprogress" & _
         " where cp09 in('" & Replace(m_RetrunRecvSub, ",", "','") & "') and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
         " and na01(+)=sp09 and cp09=ep02(+)" & _
         " and cp01=cpm01(+) and cp10=cpm02(+)" & _
         " order by 1 asc"
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         rsQuery.MoveFirst
         With rsQuery
         Do While Not .EOF
            strContent = strContent & vbCrLf
            If Trim(Left(CboEEP05.Text, 6)) <> .Fields("cp13") Then
               strContent = strContent & "當月目次：" & .Fields("ep01") & vbCrLf
            End If
            'FC代理人來台
            If "" & .Fields("PA75") <> "" And .Fields("PA09") = "000" Then
               strContent = strContent & "貴方卷號：" & "" & .Fields("PA77") & vbCrLf
            End If
            strContent = strContent & "本所案號：" & .Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04") & vbCrLf
            '送會\會圖歷程的郵件內容, 加客戶案件案號
            If (Left(CboEEP04.Text, 2) = EMP_送會 Or Left(CboEEP04.Text, 2) = EMP_會圖) And Trim("" & .Fields("客戶案件案號")) <> "" Then
               strContent = strContent & "客戶案件案號：" & "" & .Fields("客戶案件案號") & vbCrLf
            End If
            strContent = strContent & _
                         "案件名稱：" & .Fields("案件名稱") & vbCrLf & _
                         "案件性質：" & .Fields("案件性質") & vbCrLf
            rsQuery.MoveNext
         Loop
         End With
      End If
   End If
   '2020/9/29 END
   '************************************************************************************************
   
   'Add By Sindy 2013/9/6
   'Modify By Sindy 2014/1/16 智權人員不要顯示下列訊息:要轉寄客戶
   'Modify By Sindy 2018/11/7 商標處送會也要出現...位置進行說明
   '    Trim(Left(CboEEP05.Text, 6)) = strSubCP13 ==> (Trim(Left(CboEEP05.Text, 6)) = strSubCP13 And bolPAFlow = True)
   If (strSubCP01 = "FCP" And (Left(CboEEP04.Text, 2) = EMP_草完 Or Left(CboEEP04.Text, 2) = EMP_標號 Or _
                               Left(CboEEP04.Text, 2) = EMP_繪圖判發)) Or _
      ((bolPAFlow = True Or bolOtherFlow = True) And Left(CboEEP04.Text, 2) = EMP_判發) Or _
      (bolPAFlow = True And f_strAutoFlow = "判發") Or _
      Left(CboEEP04.Text, 2) = EMP_退件重送 Or _
      (Trim(Left(CboEEP05.Text, 6)) = strSubCP13 And (bolPAFlow = True Or bolOtherFlow = True)) Then
      'FCP尚未電子化 或判發或退件重送,不需顯示下列備註
   Else
   '2013/9/6 END
      
      'Modify By Sindy 2021/10/7
      'Modify By Sindy 2023/11/13 改到下列組合
      'strContent = MailContentAddEnd(strContent)
      strMailTextCom = MailContentAddEnd("")
'      strContent = strContent & vbCrLf & vbCrLf & vbCrLf & _
'                   "請至系統的下列位置進行：" & vbCrLf & vbCrLf & _
'                   " 承　辦　人　員 ：承辦人->工作進度資料維護->待辦歷程" & vbCrLf & _
'                   " 核　判　人　員 ：承辦人->待核判區" & vbCrLf & _
'                   " 智　權　人　員 ：智權部->專利商標作業->專利／商標會稿" & vbCrLf & _
'                   " 副本收受者,聯絡：共同查詢->案件查詢->案件資料及進度查詢->承辦歷程(聯絡)"
   End If
   
   'Add By Sindy 2014/1/15 +if
   If f_bolSendMail = True Then
   '2014/1/15 END
      'P台灣案,判發或退件重送不須寄Mail
      'If pfield(1) = "P" And m_country = "000" And (strAutoFlow = "判發" Or Left(CboEEP04.Text, 2) = EMP_判發 Or Left(CboEEP04.Text, 2) = EMP_退件重送) Then
      'Modify By Sindy 2013/11/28
      'Modify By Sindy 2013/12/27 玲玲說901.告知代理人也歸台灣送件
      'Modify By Sindy 2019/4/15 發文歸檔增加檢查收受者為空白,才為不寄信
      'Modify By Sindy 2023/10/20 外專送件也要發Mail
      '                           外商FC送件也要發Mail
      'Modify By Sindy 2024/8/13 + And bolFCTFlow = False
      '                            Or (bolCFTFlow = True And Trim(Left(CboEEP05.Text, 6)) = m_FlowUserNum)
      'Modify By Sindy 2025/4/23 雅娟說,目前CFP及P大陸案均會自動發MAIL如附檔，僅有台灣案沒有，故請協助增加台灣案也要發Mail
      '  (strSubCP01 = "P" And _
            (strSubPA09 = "000" Or _
             (strSubPA09 <> "000" And (strSubCP10 = "941" Or strSubCP10 = "901" Or Left(strSubEEP01, 1) >= "C")) _
            ) And _
          (f_strAutoFlow = "判發" Or Left(CboEEP04.Text, 2) = EMP_判發 Or Left(CboEEP04.Text, 2) = EMP_退件重送) _
         ) Or
      'Modify By Sindy 2025/5/19 Or (bolCFTFlow = True And Trim(Left(CboEEP05.Text, 6)) = m_FlowUserNum)
      '                     改為 Or (bolCFTFlow = True And strUserNum = m_FlowUserNum)
      If (((bolTMFlow = True Or (bolCFTFlow = True And strUserNum = m_FlowUserNum)) _
          And (Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送 Or Left(CboEEP04.Text, 2) = EMP_發文歸檔) _
         ) Or _
         (bolOtherFlow = True And (Left(CboEEP04.Text, 2) = EMP_判發 Or Left(CboEEP04.Text, 2) = EMP_退件重送)) _
         ) And bolFCPFlow = False And bolFCTFlow = False Then
      '2013/11/28 END
         '不寄Mail
         'Add By Sindy 2021/6/10 檢查是否有副本
         strTo = ""
         '2021/6/10 END
         '有副本收件者時,收件者改為副本收件者
         If Trim(txtEEP10.Text) <> "" Then
            strTo = Replace(Trim(txtEEP10.Text), ",", ";")
            txtEEP10 = ""
            strSubject = strSubject & " [此為副本通知]"
         End If
      Else
         'Modify By Sindy 2021/7/26 雅娟:由於目前已不需要由程序傳送接洽單給判發主管,故請取消此控制
'         'Modify By Sindy 2015/11/12 玲玲:工程師P案送判時,請預設E-MAIL副本收受本為程序
'         '                           中所工程師:A2027.潘韻丞 南,高工程師:A3014.蕭茹曣
'         'P非台灣案且承辦人非程序人員
'         strExc(0) = ""
'         If strSubCP01 = "P" And strSubPA09 <> "000" And GetStaffDepartment(strSubCP14) <> "P12" And _
'            Left(CboEEP04.Text, 2) = EMP_送判 Then
'            strExc(0) = PUB_GetST06(strSubCP14)
'            If strExc(0) = "2" Then '中所
'               strExc(0) = "A2027"
'            ElseIf strExc(0) = "3" Or strExc(0) = "4" Then '南高
'               strExc(0) = "A3014"
'            Else
'               strExc(0) = ""
'            End If
'         End If
'         If strExc(0) <> "" Then
'            If txtEEP10 <> "" Then txtEEP10 = txtEEP10 & ";"
'            txtEEP10 = txtEEP10 & strExc(0)
'         End If
         '2021/7/26 END
         
         'Add By Sindy 2018/7/23
         '取消專利處核/判主管(工程師主管、繪圖主管、送英核主管)都不需要收到E-Mail
         'Modify By Sindy 2018/10/23 多案件時,SendMail資訊先記錄下來
         'Modify By Sindy 2018/11/5 目前僅可做多案會完
         If f_strSubRecv <> "" And Left(CboEEP04.Text, 2) = EMP_會完 Then
            strTo = strSubCP14
         Else
         '2018/11/5 END
            strTo = Trim(Left(CboEEP05.Text, 6))
         End If
         '王副總:專利處核判主管等 20180803 開始不收E-Mail
         '針對該案的英文核稿人,核稿主管,草圖核稿人,繪圖主管,判發主管 不發E-Mail
         '聯絡的話也是針對上列該案的主管,修改聯絡都保留2天
         If Val(strSrvDate(1)) >= 20180803 Then
            '人員休假都要寄E-Mail:通知當事者和職代
            If ChkEmpIsRest(strTo) = False Then '沒休假
               '為專利處人員,商標處人員
               'Modify By Sindy 2024/2/7 取消內專恢復一樣要發mail; Left(GetStaffDepartment(strTo), 2) = "P1" Or
               If (Left(GetStaffDepartment(strTo), 2) = "P2") And Mid(strTo, 4, 1) <> "9" Then
                  'm_EMMan : 英文核稿人
                  'm_CMMan : 核稿主管
                  'm_DCMan : 草圖核稿人
                  'm_DMMan : 繪圖主管
                  'm_CSMan : 判發主管
                  If strTo = Left(m_EMMan, 5) Or strTo = Left(m_CMMan, 5) Or strTo = Left(m_DCMan, 5) Or _
                     strTo = Left(m_DMMan, 5) Or strTo = Left(m_CSMan, 5) Then
                     'Modify By Sindy 2018/12/24 若為這些人員,增加判斷流程狀態
                     'm_SPMan : 智權人員
                     'm_EPMan : 承辦人
                     'm_DPMan : 繪圖人員
                     If strTo <> strSubCP13 And strTo <> strSubCP14 And _
                        strTo <> Left(m_DPMan, 5) Then
                        strTo = "" '不發E-Mail
                     Else
                        If InStr(EMP_收受者為核判或繪圖主管, Left(CboEEP04.Text, 2)) > 0 Then
                           strTo = "" '不發E-Mail
                        End If
                     End If
                     If strTo = "" Then '不發E-Mail
                     '2018/12/24 END
                        '註記[保留聯絡]
                        strSql = "update empelectronprocess" & _
                                 " set eep11=eep11||'[保留聯絡]'" & _
                                 " where eep01='" & strSubEEP01 & "'" & _
                                 " and eep02=" & strSubEEP02 & " and eep04='" & EMP_聯絡 & "'"
                        cnnConnection.Execute strSql
                     End If
'                  '承辦人=智權人員,並且歷程為...
'                  ElseIf Left(m_EPMan, 5) = Trim(Left(m_SPMan, 6)) And _
'                     (Left(CboEEP04.Text, 2) = EMP_送會 Or Left(CboEEP04.Text, 2) = EMP_會圖 Or _
'                      Left(CboEEP04.Text, 2) = EMP_會修 Or Left(CboEEP04.Text, 2) = EMP_會完 Or _
'                      Left(CboEEP04.Text, 2) = EMP_圖修 Or Left(CboEEP04.Text, 2) = EMP_圖完 Or _
'                      Left(CboEEP04.Text, 2) = EMP_會完重修 Or Left(CboEEP04.Text, 2) = 不自動更新會完日 Or _
'                      Left(CboEEP04.Text, 2) = 不自動更新齊備日) Then
'                     strTo = "" '不發E-Mail
                  End If
                  '無收件者但有副本收件者時,收件者改為副本收件者
                  If strTo = "" And Trim(txtEEP10.Text) <> "" Then
                     strTo = Replace(Trim(txtEEP10.Text), ",", ";")
                     txtEEP10 = ""
                     strSubject = strSubject & " [此為副本通知]"
                  End If
               End If
            End If
         End If
      End If
      'CC副本多人時用;分隔
      'PUB_SendMail strUserNum, Trim(Left(CboEEP05.Text, 6)), "", strSubject, strContent, , strSendFilePath, , , , Replace(txtEEP10, ",", ";")
      If strTo <> "" Then
      '2018/7/23 END
         'Add By Sindy 2019/9/16 FCT-043079 (AA8033825) 做判發送出時,林經理反應寄信寄很久,畫面一直停在寄信中,是因為有輸入副本收受者(99011)
         If dblSize > 0 Then
            If (dblSize / 1024) > Val(Pub_GetSpecMan("系統寄信附件大小限制")) Then
               MsgBox "附件超過容量（" & Val(Pub_GetSpecMan("系統寄信附件大小限制")) & "ＭＢ）過大，郵件通知不含附件！", vbInformation
               '超過30MB時不再發送，但需提醒副本收受者，有超大附件，需自行至承辦歷程處查看
               strContent = strContent & vbCrLf & _
                  "★有超大附件（檔案大小超過" & Val(Pub_GetSpecMan("系統寄信附件大小限制")) & "MB以上），不夾帶附件，需自行至承辦歷程處查看。" & vbCrLf
               dblSize = 0
               strSendFilePath = ""
               ChkEMail.Value = 0
            End If
         End If
         '2019/9/16 END
         
         'Modify By Sindy 2024/11/25 雅娟提:有關CFP多國案，當不同承辦工程師時，若其中有任一CFP案有會稿修改，
         '                           則系統發mail通知要會修的工程師時,內容再增加一段文字
         If bolPAFlow = True Then
            If Left(CboEEP04.Text, 2) = EMP_會修 And strSubCP01 = "CFP" And InStr(NewCasePtyList, strSubCP10) > 0 Then
               stVTB = PUB_GetSameCaseSQL(strSubEEP01) '相同案語法(收文號)
               strSql = "select ep02,cp1.cp01 cp01,cp1.cp02 cp02,cp1.cp03 cp03,cp1.cp04 cp04,pa09,cp14,cp27,ep13,st02" & _
                        " from engineerprogress,caseprogress cp1,patent,staff," & _
                        "(select cp09 from caseprogress," & _
                        "(" & stVTB & ") V1" & _
                        " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
                        " and substr(V1.cno,-9,6)=cp02" & _
                        " and substr(V1.cno,-3,1)=cp03" & _
                        " and substr(V1.cno,-2)=cp04" & _
                        " and cp10 in(" & NewCasePtyList & ")) V2" & _
                        " Where V2.CP09 = ep02 and ep02=cp1.cp09(+)" & _
                        " and cp1.cp01=pa01(+) and cp1.cp02=pa02(+) and cp1.cp03=pa03(+) and cp1.cp04=pa04(+)" & _
                        " and cp1.cp01='CFP'" & _
                        " and cp158=0 and cp159=0" & _
                        " and cp14=st01(+) and cp14<>'" & Trim(Left(CboEEP05.Text, 6)) & "'"
               intI = 1
               Set rsQuery = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strContent = strContent & vbCrLf & _
                               "本案還有同時委辦其他國家，請確認下列案件是否要連帶進行此一會稿修改，" & vbCrLf & _
                               "若需要連帶修改請通知下列工程師:" & vbCrLf
                  rsQuery.MoveFirst
                  Do While Not rsQuery.EOF
                     strContent = strContent & _
                                 "案號: " & rsQuery.Fields("cp01") & "-" & rsQuery.Fields("cp02") & "-" & rsQuery.Fields("cp03") & "-" & rsQuery.Fields("cp04") & _
                                 " (" & GetPrjNationName(rsQuery.Fields("pa09")) & "國)" & _
                                 " 工程師: " & IIf(Left(rsQuery.Fields("cp14"), 1) = "F", Pub_GetSpecMan("H") & " " & GetPrjSalesNM(Pub_GetSpecMan("H")), rsQuery.Fields("cp14") & " " & rsQuery.Fields("st02")) & vbCrLf
                     rsQuery.MoveNext
                  Loop
               End If
               rsQuery.Close
            End If
         End If
         '2024/11/25 END
                  
         strContent = strContent & vbCrLf & strMailTextCom 'Modify By Sindy 2023/12/6
         
         'Add By Sindy 2023/12/18 專利日本部簡經理提出:代理操作歷程者,要加發原收受者
         'Modify By Sindy 2024/8/13 + Or bolCFTFlow = True Or bolFCTFlow = True
         If bolFCPFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
            If m_EEP12 <> "" And m_EEP16 <> "" Then
               If txtEEP10 <> "" Then txtEEP10 = txtEEP10 & ";"
               txtEEP10 = txtEEP10 & m_EEP16
            End If
         End If
         '2023/12/18 END
         
         'Modify By Sindy 2024/12/5 送件給FCT程序人員時, 副本要加發程序主管
         If bolFCTFlow = True And (Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送) _
            And txtEEP03 <> GetST52SelfList(Left(m_NPMan, 5)) Then
            If txtEEP10 <> "" Then txtEEP10 = txtEEP10 & ";"
            txtEEP10 = txtEEP10 & GetST52SelfList(Left(m_NPMan, 5))
         End If
         '2024/11/28 END
         
         'Trim(Left(CboEEP05.Text, 6)) ==> strTo
         'Add By Sindy 2018/8/13 商標主旨要加提醒[此案智權人員欲管控收款後才可送件]
         'Modify By Sindy 2018/10/23 多案件時,SendMail資訊先記錄下來
         If f_strSubRecv <> "" Then 'Modify By Sindy 2018/10/23 目前僅可做多案會完
            '此方法無寄附件功能
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
               ",to_char(sysdate,'hh24miss'),'" & strSubject & m_SubjectNote & "','" & ChgSQL(strContent) & "'," & CNULL(Replace(txtEEP10, ",", ";")) & ")"
            cnnConnection.Execute strSql, intI
         Else
         '2018/10/23 END
            'Modify By Sindy 2021/12/29 寄信時判斷收受者為林律師的商標案就不要轉發職代
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            If Trim(Left(CboEEP05.Text, 6)) = "98003" And (bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True) Then
               'Add By Sindy 2022/2/21 + strSubEEP01
               PUB_SendMail strUserNum, strTo, strSubEEP01, strSubject & m_SubjectNote, strContent, , strSendFilePath, , , , Replace(txtEEP10, ",", ";"), , , , True
            Else
            '2021/12/29 END
               'Modify By Sindy 2024/1/5
               If Left(CboEEP04.Text, 2) = EMP_轉檔完成 And ChkEED13.Value = 1 Then '轉檔完成程序送件
                  PUB_SendMail strUserNum, strTo, strSubEEP01, strSubject & m_SubjectNote, strContent, , strSendFilePath, , , , Left(m_EPMan, 5) & ";" & Left(m_F21CMMan, 5)
               Else
               '2024/1/5 END
                  'Modify By Sindy 2025/11/6 FCT-52372發文歸檔沒寄信(林靖傑反應的)
                  '   是因有出現"無此檔案"的訊息, 因歸檔後會將PC端的電子檔刪除, 此時在信件裡不用寄帶附件
                  If Left(CboEEP04.Text, 2) = EMP_發文歸檔 And strSendFilePath <> "" Then
                     strSendFilePath = ""
                  End If
                  '2025/11/6 END
                  'Add By Sindy 2022/2/21 + strSubEEP01
                  'Modify By Sindy 2025/6/9 偉城跟雅娟反應,不需顯示人員休假,通知職代的訊息
                  PUB_SendMail strUserNum, strTo, strSubEEP01, strSubject & m_SubjectNote, strContent, , strSendFilePath, , , , Replace(txtEEP10, ",", ";"), , , , , False
               End If
            End If
         End If
      End If
      '2015/11/12 END
      
'      '流程Damo用
'      MsgBox "收件人：" & Trim(Left(CboEEP05.Text, 6)) & " " & GetPrjSalesNM(Trim(Left(CboEEP05.Text, 6))) & vbCrLf & vbCrLf & _
'             "副本收受者：" & IIf(txtEEP10 <> "", Replace(txtEEP10, ",", ";"), "") & vbCrLf & vbCrLf & _
'             "主　旨：" & strSubject & vbCrLf & vbCrLf & _
'             "內　容：" & vbCrLf & vbCrLf & strContent & _
'             IIf(strSendFilePath <> "", vbCrLf & vbCrLf & "附　檔：" & strSendFilePath, ""), , "E-Mail內容"
      
      'Add By Sindy 2013/9/2
      'Modify By Sindy 2015/4/23 +Or Left(CboEEP04.Text, 2) = EMP_草核完
      'Modify By Sindy 2023/12/8 +Or (bolFCPFlow = True And Left(CboEEP04.Text, 2) = EMP_送件)
      'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True
      If Left(CboEEP04.Text, 2) = EMP_繪圖判發 Or _
         Left(CboEEP04.Text, 2) = EMP_草核完 Or _
         ((bolPAFlow = True Or bolOtherFlow = True) And Left(CboEEP04.Text, 2) = EMP_判發) Or _
         ((bolFCPFlow = True Or bolFCTFlow = True Or bolCFTFlow = True) And Left(CboEEP04.Text, 2) = EMP_送件) Then
         
         strSubject = Replace(strSubCaseNo, "-0-00", "") & _
                     "(" & strSubPA09Nm & ")(核會流程)-->(" & strSubEEP02 & ")已" & Right(CboEEP04.Text, Len(CboEEP04.Text) - 3)
'         'Modify By Sindy 2024/11/28 薛經理反應主旨中的"已送件"易混淆,因此改字樣
'         'ex:FCP-064712(台灣)(核會流程)-->(2)已送件
'         If InStr(strSubject, "已送件") > 0 And Left(CboEEP04.Text, 2) = EMP_送件 Then
'            strSubject = Replace(strSubject, "已送件", "已準備送件")
'         End If
'         '2024/11/28 END
         
         'Modify By Sindy 2018/4/2
         'FC代理人來台
         strExc(10) = ""
         If strSubPA75 <> "" And strSubPA09 = "000" Then
            strExc(10) = "貴方卷號：" & strSubPA77 & vbCrLf
         End If
         '2018/4/2 END
         'Modify By Sindy 2023/12/15 杜燕文協理請作,內文加申請國家
         strContent = "當月目次：" & strSubEP01 & vbCrLf & strExc(10) & _
                      "本所案號：" & strSubCaseNo & vbCrLf & _
                      "案件名稱：" & strCaseName & vbCrLf & _
                      "申請國家：" & strSubPA09Nm & vbCrLf & _
                      "案件性質：" & strSubCP10Nm & vbCrLf & _
                      "流程狀態：" & Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) & vbCrLf
         If Trim(txtEEP08) <> "" Then
            strContent = strContent & "內　　容：" & Trim(txtEEP08) & vbCrLf & vbCrLf & vbCrLf
         Else
            strContent = strContent & vbCrLf & vbCrLf
         End If
         'Modify By Sindy 2013/10/29 自行判發者本就不需再發Mail給自己,改用m_FlowUserNum判斷
         'If Left(CboEEP04.Text, 2) = EMP_繪圖判發 And Left(m_DMMan, 5) <> Left(m_DPMan, 5) Then
         'Modify By Sindy 2015/4/22 +EMP_草核完
         'Modify By Sindy 2016/3/8 姍珊:繪圖判發完成後，系統不要再傳EMAIL通知繪圖
         'If (Left(CboEEP04.Text, 2) = EMP_繪圖判發 Or Left(CboEEP04.Text, 2) = EMP_草核完) And _
            m_FlowUserNum <> Left(m_DPMan, 5) Then
         If Left(CboEEP04.Text, 2) = EMP_草核完 And m_FlowUserNum <> Left(m_DPMan, 5) Then
         '2016/3/8 END
            '發給繪圖人員
            'Add By Sindy 2022/2/21 + strSubEEP01
            PUB_SendMail strUserNum, Left(Trim(m_DPMan), 5), strSubEEP01, strSubject, strContent
         'ElseIf Left(CboEEP04.Text, 2) = EMP_判發 And Left(m_CSMan, 5) <> Left(m_EPMan, 5) Then
         ElseIf (Left(CboEEP04.Text, 2) = EMP_判發 Or Left(CboEEP04.Text, 2) = EMP_送件) _
               And m_FlowUserNum <> strSubCP14 Then
            '發給工程師
            'Add By Sindy 2022/2/21 + strSubEEP01
            'Modify By Sindy 2024/1/2 排除 承辦人為外專程序人員
            'Modify By Sindy 2024/1/3 排除 上一筆歷程m_strLastEEP04 <> EMP_程序送判
            'Modify By Sindy 2024/8/14 排除 承辦人為外商程序人員
            If PUB_GetST03(strSubCP14) <> "F22" And m_strLastEEP04 <> EMP_程序送判 And PUB_GetST03(strSubCP14) <> "F12" Then
            '2024/1/2 END
               PUB_SendMail strUserNum, strSubCP14, strSubEEP01, strSubject, strContent
            End If
         End If
         
         'Modify By Sindy 2015/4/24 雅娟提:
         If bolPAFlow = True And Left(CboEEP04.Text, 2) = EMP_判發 Then
            '針對台灣新申請案不繪圖判發時，檢查是否有大陸新申請案不繪圖且未發文的，並且承辦人是品薇時，
            '一併發E-Mail通知品薇可以處理大陸案了，無須再等待圖式。
            'Modify By Sindy 2018/10/4 98012改判斷是P12專利處程序
            'and cp14='98012' => and cp14=st01(+) and st03='P12'
            If strSubPA09 = "000" And InStr(NewCasePtyList, strSubCP10) > 0 And m_DPMan = "" Then
               stVTB = PUB_GetSameCaseSQL(strSubEEP01) '相同案語法(收文號)
               strSql = "select ep02,cp1.cp01,cp1.cp02,cp1.cp03,cp1.cp04,pa09,cp14,cp27,ep13" & _
                        " from engineerprogress,caseprogress cp1,patent,staff," & _
                        "(select cp09 from caseprogress," & _
                        "(" & stVTB & ") V1" & _
                        " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
                        " and substr(V1.cno,-9,6)=cp02" & _
                        " and substr(V1.cno,-3,1)=cp03" & _
                        " and substr(V1.cno,-2)=cp04" & _
                        " and cp10 in(" & NewCasePtyList & ")) V2" & _
                        " Where V2.CP09 = ep02 and ep02=cp1.cp09(+)" & _
                        " and cp1.cp01=pa01(+) and cp1.cp02=pa02(+) and cp1.cp03=pa03(+) and cp1.cp04=pa04(+)" & _
                        " and pa09='020'" & _
                        " and (EP13 is null or EP13='99999')" & _
                        " and cp27 is null" & _
                        " and cp14=st01(+) and st03='P12'"
               'Add By Sindy 2015/10/21 +服務
               'Modify By Sindy 2018/10/4 98012改判斷是P12專利處程序
               'and cp14='98012' => and cp14=st01(+) and st03='P12'
               strSql = strSql & " union select ep02,cp1.cp01,cp1.cp02,cp1.cp03,cp1.cp04,SP09,cp14,cp27,ep13" & _
                        " from engineerprogress,caseprogress cp1,servicepractice,staff," & _
                        "(select cp09 from caseprogress," & _
                        "(" & stVTB & ") V1" & _
                        " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
                        " and substr(V1.cno,-9,6)=cp02" & _
                        " and substr(V1.cno,-3,1)=cp03" & _
                        " and substr(V1.cno,-2)=cp04" & _
                        " and cp10 in(" & NewCasePtyList & ")) V2" & _
                        " Where V2.CP09 = ep02 and ep02=cp1.cp09(+)" & _
                        " and cp1.cp01=SP01(+) and cp1.cp02=SP02(+) and cp1.cp03=SP03(+) and cp1.cp04=SP04(+)" & _
                        " and SP09='020'" & _
                        " and (EP13 is null or EP13='99999')" & _
                        " and cp27 is null" & _
                        " and cp14=st01(+) and st03='P12'"
               intI = 1
               Set rsQuery = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  '發E-Mail給品薇
                  'strContent = strContent & vbCrLf & "★註：可以處理大陸案件(" & rsQuery.Fields(1) & "-" & rsQuery.Fields(2) & "-" & rsQuery.Fields(3) & "-" & rsQuery.Fields(4) & ")，無須再等待圖式。" & vbCrLf
                  strContent = "★註：可以處理大陸案件(" & rsQuery.Fields(1) & "-" & rsQuery.Fields(2) & "-" & rsQuery.Fields(3) & "-" & rsQuery.Fields(4) & ")，無須再等待圖式。" & _
                        vbCrLf & strContent & vbCrLf
                  'Modify By Sindy 2018/10/4 98012改判斷是P12專利處程序
                  'PUB_SendMail strUserNum, "98012", "", strSubject, strContent
                  'Add By Sindy 2022/2/21 + rsQuery.Fields("ep02")
                  PUB_SendMail strUserNum, rsQuery.Fields("cp14"), rsQuery.Fields("ep02"), strSubject, strContent
                  'Modify By Sindy 2018/12/20 加註,有加發副本收受者
                  strSql = "update empelectronprocess" & _
                           " set eep10=decode(eep10,null,'',eep10||',')||'" & rsQuery.Fields("cp14") & "'" & _
                           " where eep01='" & strSubEEP01 & "'" & _
                           " and eep02=" & strSubEEP02
                  cnnConnection.Execute strSql
               End If
               rsQuery.Close
            End If
         End If
         '2015/4/24 END
      End If
      '2013/9/2 END
      
      If bolPAFlow = True Then
         'Add By Sindy 2018/12/14 玲玲提:
         '針對台灣新申請案時，檢查是否有大陸新申請案，並且承辦人是程序人員時，
         '一併發E-Mail通知程序人員可以處理大陸案了。
         If Left(CboEEP04.Text, 2) = EMP_聯絡 _
            And PUB_GetST03(Trim(Left(CboEEP05.Text, 6))) = "P12" _
            And Left(PUB_GetStaffST15(Trim(txtEEP03), "1"), 1) = "S" Then

            If strSubPA09 = "000" And strSubCP01 = "P" And InStr(NewCasePtyList, strSubCP10) > 0 Then
               stVTB = PUB_GetSameCaseSQL(strSubEEP01) '相同案語法(收文號)
               strSql = "select ep02,cp1.cp01,cp1.cp02,cp1.cp03,cp1.cp04,pa09,cp14,cp27,ep13" & _
                        " from engineerprogress,caseprogress cp1,patent,staff," & _
                        "(select cp09 from caseprogress," & _
                        "(" & stVTB & ") V1" & _
                        " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
                        " and substr(V1.cno,-9,6)=cp02" & _
                        " and substr(V1.cno,-3,1)=cp03" & _
                        " and substr(V1.cno,-2)=cp04" & _
                        " and cp10 in(" & NewCasePtyList & ")) V2" & _
                        " Where V2.CP09 = ep02 and ep02=cp1.cp09(+)" & _
                        " and cp1.cp01=pa01(+) and cp1.cp02=pa02(+) and cp1.cp03=pa03(+) and cp1.cp04=pa04(+)" & _
                        " and pa09<>'000'" & _
                        " and cp27 is null" & _
                        " and cp14=st01(+) and st03='P12'"
               strSql = strSql & " union select ep02,cp1.cp01,cp1.cp02,cp1.cp03,cp1.cp04,SP09,cp14,cp27,ep13" & _
                        " from engineerprogress,caseprogress cp1,servicepractice,staff," & _
                        "(select cp09 from caseprogress," & _
                        "(" & stVTB & ") V1" & _
                        " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
                        " and substr(V1.cno,-9,6)=cp02" & _
                        " and substr(V1.cno,-3,1)=cp03" & _
                        " and substr(V1.cno,-2)=cp04" & _
                        " and cp10 in(" & NewCasePtyList & ")) V2" & _
                        " Where V2.CP09 = ep02 and ep02=cp1.cp09(+)" & _
                        " and cp1.cp01=SP01(+) and cp1.cp02=SP02(+) and cp1.cp03=SP03(+) and cp1.cp04=SP04(+)" & _
                        " and SP09<>'000'" & _
                        " and cp27 is null" & _
                        " and cp14=st01(+) and st03='P12'"
               intI = 1
               Set rsQuery = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  '發E-Mail給大陸承辦的程序人員
                  strContent = "★註：確認一下大陸案件(" & rsQuery.Fields(1) & "-" & rsQuery.Fields(2) & "-" & rsQuery.Fields(3) & "-" & rsQuery.Fields(4) & ")資料，P案的申請資訊可能有修改。" & _
                        vbCrLf & strContent & vbCrLf
                  'Add By Sindy 2022/2/21 + rsQuery.Fields("ep02")
                  PUB_SendMail strUserNum, rsQuery.Fields("cp14"), rsQuery.Fields("ep02"), strSubject, strContent
                  '加註,有加發副本收受者
                  strSql = "update empelectronprocess" & _
                           " set eep10=decode(eep10,null,'',eep10||',')||'" & rsQuery.Fields("cp14") & "'" & _
                           " where eep01='" & strSubEEP01 & "'" & _
                           " and eep02=" & strSubEEP02
                  cnnConnection.Execute strSql
               End If
               rsQuery.Close
            End If
         End If
      End If
      
   End If
   
   'Add By Sindy 2017/9/5 多國案新增歷程發E-Mail
   If cmdCaseMap.Visible = True And cmdCaseMap.Enabled = True And m_RetrunRecv <> "" Then
      Call ProcessCaseMapSMail
   End If
   '2017/9/5 END
   
   FlowSendMail = True
   Set rsQuery = Nothing
End Function

'Add By Sindy 2017/9/1 多國案新增歷程
Private Function ProcessCaseMap(strUpdTime As String) As Boolean
Dim arrID As Variant, intCnt As Integer
Dim strCP09 As String, intEEP02 As Integer, intMaxEEP02 As Integer
   
   arrID = Split(m_RetrunRecv, ",")
   For intCnt = 0 To UBound(arrID)
      strCP09 = arrID(intCnt)
      '******************************
      '      承辦電子簽核流程檔
      '******************************
      '取得最大序號
      intMaxEEP02 = 0
      strSql = "select eep02 From empelectronprocess where eep01='" & strCP09 & "' order by eep02 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         If RsTemp.RecordCount > 0 Then
            intMaxEEP02 = RsTemp.Fields(0)
         End If
      End If
      intEEP02 = intMaxEEP02 + 1
      
      '記錄處理的流程狀態
      m_UpdEEP11 = "多國案新增歷程:" & Replace(lblCaseNo, "-0-00", "") & "," & lblCP09
      'Modify By Sindy 2023/12/18 +,eep16
      strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep10,eep11,eep12,eep16) values(" & _
               CNULL(strCP09) & "," & intEEP02 & "," & CNULL(Trim(txtEEP03)) & "," & _
               CNULL(Left(CboEEP04.Text, 2)) & "," & _
               CNULL(Trim(Left(CboEEP05.Text, 6))) & "," & _
               strSrvDate(1) & "," & strUpdTime & "," & CNULL(ChgSQL(txtEEP08)) & "," & _
               CNULL(txtEEP10) & ",'" & m_UpdEEP11 & "','" & m_EEP12 & "','" & m_EEP16 & "')"
      cnnConnection.Execute strSql
      
      '******************************
      '      承辦電子簽核附件檔
      '******************************
      If SaveAttFile(strCP09, CInt(intEEP02), 0) = False Then
         ProcessCaseMap = False
         Exit Function
      End If
   Next intCnt
   ProcessCaseMap = True
End Function

'Add By Sindy 2017/9/4 多國案新增歷程發E-Mail
Private Sub ProcessCaseMapSMail()
Dim arrID As Variant, intCnt As Integer
Dim strCP09 As String, strCaseNo As String, strCaseName As String
Dim strEP01 As String, strCP10Nm As String
   
   arrID = Split(m_RetrunRecv, ",")
   For intCnt = 0 To UBound(arrID)
      strCP09 = arrID(intCnt)
      
      '******************************
      '      基本資料檔
      '******************************
      strSql = "select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,NVL(PA05,NVL(PA06,PA07)) as 案件名稱" & _
               ",Decode(PA09,'000',CPM03,CPM04) as 案件性質,ep01 as 目次,cp13,cp14,PA09,PA75,PA77" & _
               " From caseprogress,patent,engineerprogress,casepropertymap" & _
               " where cp09='" & strCP09 & "'" & _
               " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
               " and cp09=ep02(+)" & _
               " and cp01=cpm01(+) and cp10=cpm02(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         If RsTemp.RecordCount > 0 Then
            strCaseNo = RsTemp.Fields("本所案號")
            strCaseName = RsTemp.Fields("案件名稱")
            strEP01 = RsTemp.Fields("目次")
            strCP10Nm = RsTemp.Fields("案件性質")
         End If
      End If
      
      'Modify By Sindy 2023/12/14 杜燕文協理請作,主旨加申請國家
      If Trim(Left(CboEEP05.Text, 6)) = Trim(Left(m_SPMan, 6)) Then
         strSubject = Replace(strCaseNo, "-0-00", "") & "(" & GetPrjNation(strCaseNo) & ")「" & strCaseName & "」-->" & Right(CboEEP04.Text, Len(CboEEP04.Text) - 3)
      Else
         strSubject = Replace(strCaseNo, "-0-00", "") & "(" & GetPrjNation(strCaseNo) & ")(核會流程)-->" & Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) & "，請進行後續處理"
      End If
      strContent = ""
      If Trim(Left(CboEEP05.Text, 6)) <> Trim(Left(m_SPMan, 6)) Then
         strContent = "當月目次：" & strEP01 & vbCrLf
      End If
      'Modify By Sindy 2018/4/2
      'FC代理人來台
      If "" & RsTemp.Fields("PA75") <> "" And "" & RsTemp.Fields("PA09") = "000" Then
         strContent = strContent & "貴方卷號：" & "" & RsTemp.Fields("PA77") & vbCrLf
      End If
      '2018/4/2 END
      'Modify By Sindy 2023/12/15 杜燕文協理請作,內文加申請國家
      strContent = strContent & "本所案號：" & strCaseNo & vbCrLf
      strContent = strContent & _
                   "案件名稱：" & strCaseName & vbCrLf & _
                   "申請國家：" & GetPrjNation(strCaseNo) & vbCrLf & _
                   "案件性質：" & strCP10Nm & vbCrLf & _
                   "流程狀態：" & Right(CboEEP04.Text, Len(CboEEP04.Text) - 3) & vbCrLf
      If Trim(txtEEP08) <> "" Then
         strContent = strContent & "內　　容：" & Trim(txtEEP08) & vbCrLf
      End If
      
      'Modify By Sindy 2021/10/7
      strContent = MailContentAddEnd(strContent)
'      strContent = strContent & vbCrLf & vbCrLf & vbCrLf & _
'                   "請至系統的下列位置進行：" & vbCrLf & vbCrLf & _
'                   " 承　辦　人　員 ：承辦人->工作進度資料維護->待辦歷程" & vbCrLf & _
'                   " 核　判　人　員 ：承辦人->待核判區" & vbCrLf & _
'                   " 智　權　人　員 ：智權部->專利商標作業->專利／商標會稿" & vbCrLf & _
'                   " 副 本 收 受 者 ：共同查詢->案件查詢->案件資料及進度查詢->承辦歷程(聯絡)"
      'Add By Sindy 2022/2/21 + strCP09
      PUB_SendMail strUserNum, Trim(Left(CboEEP05.Text, 6)), strCP09, strSubject, strContent, , strSendFilePath, , , , Replace(txtEEP10, ",", ";")
   Next intCnt
End Sub

Private Function SaveAttFile(strEEF01 As String, intEEF02 As Integer, Index As Integer) As Boolean
Dim stFilePath As String
Dim iFileNo As Integer
'Dim bytes() As Byte
Dim lngSize As Long '檔案大小
Dim adoRst As New ADODB.Recordset
'Const BlockSize = 500000
'Dim Numblocks As Integer
'Dim LeftOver As Long
Dim UpdModifyDate As Double, UpdModifyTime As Double
Dim bolGetFileName As Boolean
Dim intRow As Integer '檔案數量
Dim strFile As String, stReName As String, strTemp As String
Dim stFtpPath As String 'Added by Morgan 2015/4/28
   
On Error GoTo ErrHand
   
   SaveAttFile = True
   
   For ii = 0 To lstAtt(Index).ListCount - 1
      If lstAtt(Index).ITEMDATA(ii) = 0 Then
         stFilePath = lstAtt(Index).List(ii)
         If InStrRev(stFilePath, " (") > 0 Then
            'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
            If UCase(Mid(stFilePath, InStrRev(stFilePath, " (") + 1, Len("(X86)"))) <> "(X86)" Then
            '2021/8/6 END
               stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
            End If
         End If
         UpdModifyDate = Mid(lstAtt(Index).List(ii), InStr(lstAtt(Index).List(ii), "#") + 1, 8)
         UpdModifyTime = Mid(lstAtt(Index).List(ii), InStr(lstAtt(Index).List(ii), "#") + 9, 6)
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
            
            'Add By Sindy 2015/1/19
            strFile = GetFileName(stFilePath)
            stReName = "" 'Add By Sindy 2024/3/12
            If Index = 1 Then '存卷資料
               'Add By Sindy 2019/6/27 檢查是否已有案件副檔名,若有,不用更名
               strSql = "select EFC01,EFC02 from efilecaption where EFC06='Y'" & _
                        " and instr(upper('" & ChgSQL(strFile) & "'),upper('.'||EFC02||'.'))>0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  Call PUB_GetEmpFlowReNameFile(PField(1), PField(2), PField(3), PField(4), "", strFile, stReName)
                  If stReName <> "" Then strFile = stReName
               'Modify By Sindy 2023/11/22 外專案件存卷區不用更名
               'Modify By Sindy 2024/8/13 + And bolCFTFlow = False And bolFCTFlow = False
               ElseIf bolFCPFlow = False And bolCFTFlow = False And bolFCTFlow = False Then
               '2019/6/27 END
                  If InStr(UCase(strFile), UCase("." & EMP_存卷資料)) = 0 And _
                     InStr(UCase(strFile), UCase("." & EMP_客戶資料)) = 0 And _
                     Left(UCase(strFile), 4) <> UCase(EMP_存卷資料) And _
                     Left(UCase(strFile), 4) <> UCase(EMP_客戶資料) Then
                     
                     '取得檔名
                     bolGetFileName = False: intRow = 0: stReName = ""
                     Do While bolGetFileName = False
                        stReName = Trim(PField(1)) & CStr(Val(PField(2))) & IIf(PField(3) <> "0" Or PField(4) <> "00", "-" & PField(3), "") & IIf(PField(4) <> "00", "-" & PField(4), "")
                        If InStr(UCase(strFile), UCase("." & cp(10))) = 0 Then
                           stReName = stReName & "." & EMP_客戶資料 & IIf(intRow > 0, intRow, "") & Mid(strFile, InStr(strFile, ".")) '截取本所案號後的檔名
                        Else
                           strTemp = Mid(strFile, InStr(strFile, ".") + 1)
                           stReName = stReName & "." & cp(10) & "." & EMP_客戶資料 & IIf(intRow > 0, intRow, "") & Mid(strTemp, InStr(strTemp, ".")) '截取第2個.後面的檔名
                        End If
                        '檢查檔案是否已存在
                        strSql = "select eef03" & _
                                 " From EmpElectronFile" & _
                                 " where eef01='" & strEEF01 & "' and eef02=" & intEEF02 & _
                                 " and upper(eef03)='" & UCase(stReName) & "'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 0 Then
                           bolGetFileName = True
                           Call PUB_InsEfileCaption(CStr(PField(1)), EMP_客戶資料, intRow) '檢查是否有需要新增電子檔次要副檔名說明
                           Exit Do
                        End If
                        intRow = intRow + 1
                     Loop
                     If stReName <> "" Then
                        strFile = stReName
                     End If
                  End If
               End If
            'Modify By Sindy 2023/12/12 雅娟跟薛經理反應:USPTO檔案命名規則 (本所卷宗區的副檔名可否修正為小寫.其餘不變,目前官方不接受大寫副檔名)
            Else
               'Add By Sindy 2024/3/12 多案歷程開放.CDATA.可以放多案歷程的其他案號
               If Not (txtLpNote.Tag = "多案單筆歷程" And InStr(UCase(strFile), ".CDATA.") > 0) Then
               '2024/3/12 END
                  Call PUB_GetEmpFlowReNameFile(PField(1), PField(2), PField(3), PField(4), "", strFile, stReName)
                  If stReName <> "" Then strFile = stReName
               End If
            '2023/12/12 END
            End If
            '2015/1/19 END
            
            'Modify By Sindy 2018/9/25 商標處開放下列檔名可不加本所案號,存檔時系統自動補填
            'Modify By Sindy 2018/10/5 客戶會稿在E-Mail中附加的電子檔有可能沒有本所案號,存檔時系統自動補填
            'If NotChkFileCaseNo(strFile, Index) = True Or Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
            If NotChkFileCaseNo(strFile, Index) = True Then
               If InStr(UCase(strFile), m_strSaveCaseNo1) = 0 And _
                  InStr(UCase(strFile), m_strSaveCaseNo2) = 0 And _
                  InStr(UCase(strFile), m_strSaveCaseNo3) = 0 And _
                  InStr(UCase(strFile), m_strSaveCaseNo4) = 0 Then
                  strFile = m_strSaveCaseNo3 & "." & strFile
               End If
            End If
            '2018/9/25 END
            
            .AddNew
            .Fields("eef01").Value = strEEF01
            .Fields("eef02").Value = intEEF02
            .Fields("eef03").Value = strFile 'GetFileName(stFilePath)
            .Fields("eef04").Value = lngSize
            
'Removed by Morgan 2015/5/27 不再存DB
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
'end 2015/5/27
            
            .Fields("eef09").Value = UpdModifyDate
            .Fields("eef10").Value = UpdModifyTime
            Close #iFileNo
            
            'Added by Morgan 2015/4/28 檔案改放FTP
            PUB_PutFtpFile stFilePath, strEEF01, strFile, stFtpPath, "EMPELECTRONFILE", CStr(intEEF02)
            If stFtpPath <> "" Then
               .Fields("eef11") = strSrvDate(1)
               .Fields("eef12") = stFtpPath
            End If
            'end 2015/4/28
            .UPDATE
         End With
      End If
   Next ii
   
   Exit Function
   
ErrHand:
   Close #iFileNo
   SaveAttFile = False
   MsgBox Err.Description, vbCritical
End Function

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   txtEEP03.Enabled = bEnable
   txtEEP03_2.Enabled = bEnable
   CboEEP04.Enabled = bEnable
   CboEEP05.Enabled = bEnable
   ChkEMail.Enabled = bEnable
   txtEEP10_2.Enabled = bEnable
   'txtEEP08.Enabled = bEnable
   txtEEP08.Locked = Not bEnable
   
   'Add By Sindy 2013/10/1 案件名稱
   If bEnable = True Then
      txtCaseName(0).Enabled = False
      txtCaseName(1).Enabled = False
      txtCaseName(2).Enabled = False
      'Modify By Sindy 2023/1/9 內商開放可以修改案件名稱 + Or bolTMFlow = True
      'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True
      If m_FlowUserNum = Left(m_EPMan, 5) And (bolPAFlow = True Or bolTMFlow = True Or bolCFTFlow = True) Then '工程師才可以改
         txtCaseName(0).Enabled = bEnable
         txtCaseName(1).Enabled = bEnable
         txtCaseName(2).Enabled = bEnable
      End If
   Else
      txtCaseName(0).Enabled = bEnable
      txtCaseName(1).Enabled = bEnable
      txtCaseName(2).Enabled = bEnable
   End If
   '2013/10/1 END
   
   If bEnable = False Then
      If Me.cmdSend.Visible = True Then
         Me.cmdAdd.Visible = True
      End If
      Me.cmdCancel.Visible = False
      Me.cmdSend.Enabled = False
   Else
      Me.cmdOpenAtt(0).Enabled = True
      Me.cmdSelect(0).Enabled = True
      Me.cmdAddAtt(0).Enabled = True
      CmdF21(0).Enabled = True 'Add By Sindy 2025/10/28
      Me.cmdRemAtt(0).Enabled = True
      Me.cmdOpenAtt(1).Enabled = True
      Me.cmdSelect(1).Enabled = True
      Me.cmdAddAtt(1).Enabled = True
      Me.cmdRemAtt(1).Enabled = True
   End If
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim intPDF As Integer, intDOC As Integer, intDwgPdf As Integer, intDwg As Integer
Dim intPoaPDF As Integer 'Add By Sindy 2018/7/13 , intDataPDF As Integer
Dim intDataDOC As Integer 'Add By Sindy 2019/5/16
Dim intCUSPDF As Integer 'Add By Sindy 2018/9/26
Dim stVTB As String, IsExistsDWG As Boolean
Dim strPathFile As String
Dim intQ As Integer
Dim rsA As New ADODB.Recordset
Dim bolHadShowMsg As Boolean 'Add By Sindy 2018/9/26
Dim bolNotChkFileCaseNo As Boolean
Dim intMsgQ As Integer
'Add By Sindy 2025/10/22
Dim strShowMsg As String, bolCheck As Boolean
Dim dblFCnt As Double
'2025/10/22 END
   
   TxtValidate = False
   m_UpdEEP11 = "" 'Add By Sindy 2013/11/13
   
   If Trim(CboEEP04.Text) = "" Then
      MsgBox "流程狀態不可空白！", vbExclamation
      CboEEP04.SetFocus
      SSTab1.Tab = 0
      Exit Function
   'Add By Sindy 2013/9/23
   Else
      If Left(CboEEP04.Text, 2) = EMP_附加流程 Then
         If CboCP10.ListIndex < 0 Then
            MsgBox "案件性質不可空白！", vbExclamation
            SSTab1.Tab = 0
            Exit Function
         End If
      'Add By Sindy 2018/8/29
      ElseIf Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
         If CboCP10.ListIndex < 0 Then
            MsgBox "會稿方式不可空白！", vbExclamation
            SSTab1.Tab = 0
            Exit Function
         End If
      End If
   '2013/9/23 END
   End If
   
   'Added by Morgan 2025/3/14
   If Left(CboEEP04.Text, 2) = EMP_判發 And cp(1) = "P" And cp(10) = "413" And pa(9) = "000" Then
      If PUB_ChkTW413(cp(43)) = False Then
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   'end 2025/3/14
   
   'Add By Sindy 2025/10/15
   If Frame945.Visible = True Then
      If txtEED14.Visible = True And txtEED14.Enabled = True Then
         Cancel = False
         Call txtEED14_Validate(Cancel)
         If Cancel = True Then
            SSTab1.Tab = intTab_外專承辦單
            txtEED14.SetFocus
            Exit Function
         End If
      End If
      If txtEED15.Visible = True And txtEED15.Enabled = True Then
         Cancel = False
         Call txtEED15_Validate(Cancel)
         If Cancel = True Then
            SSTab1.Tab = intTab_外專承辦單
            txtEED15.SetFocus
            Exit Function
         End If
      End If
   End If
   '2025/10/15 END
'   'Add By Sindy 2025/4/10
'   If Frame945.Tag = "945" Then
'      '若送判時 委員指定送件日期【本所期限】為空時，彈提醒:
'      '委員指定送件日期為空，請確認是否無需控管期限?
'      '1.是: 繼續下一流程判發
'      '2.否: 不可跑下一流程
'      '若工程師主管認為期限需修改\或工程師認為不需管制期限但主管認為需管制，請用退回，由承辦人修改後重新送判
'      If Me.txtEED15.Text <> "" Then
''         If MsgBox("委員指定送件日期為空白，請確認是否無需控管期限？" & vbCrLf & _
''            "是：不需控管期限,繼續..." & vbCrLf & _
''            "否：取消送出,欲輸入本所期限。", vbQuestion + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
''            txtEED15.SetFocus
''            Me.SSTab1.Tab = intTab_外專承辦單
''            Exit Function
''         End If
''      Else
'         '約定期限不可大於本所期限
'         If Me.txtEED14.Text >= Me.txtEED15.Text Then
'            MsgBox "約定期限不可大於本所期限！", vbExclamation
'            txtEED14.SetFocus
'            Me.SSTab1.Tab = intTab_外專承辦單
'            Exit Function
'         End If
'      End If
'   End If
'   '2025/4/10 END
   
   'If Left(CboEEP04.Text, 2) <> EMP_附加流程 Then
      'Modify By Sindy 2014/1/15
      'If Trim(CboEEP05.Text) = "" Then
      'Modify By Sindy 2018/7/24 + And CboEEP05.Enabled = True
      If Trim(CboEEP05.Text) = "" And _
         CboEEP05.Enabled = True Then 'And _
         Not (Left(CboEEP04.Text, 2) = EMP_附加流程 And _
              (Trim(Left(CboCP10.Text, 4)) = "936" Or _
               Trim(Left(CboCP10.Text, 4)) = "957" Or _
               Trim(Left(CboCP10.Text, 4)) = "958")) Then
      '2014/1/15 END
         MsgBox "收受者不可空白！", vbExclamation
         CboEEP05.SetFocus
         SSTab1.Tab = 0
         Exit Function
      'Add By Sindy 2024/6/17
      Else
         If Trim(Left(CboEEP05.Text, 6)) <> "" Then
            If ChkStaffST04(Trim(Left(CboEEP05.Text, 6)), False) = True Then
               MsgBox "收受者 " & Trim(Mid(CboEEP05.Text, 6)) & " 已離職！", vbExclamation
               SSTab1.Tab = 0
               Exit Function
            End If
         End If
      '2024/6/17 END
      'Modify By Sindy 2013/10/3 不可鎖,有工程師亦是智權人員,還是需要送會給客戶會稿
'      Else
'         'Add By Sindy 2013/10/2 發送者與收受者不可為同一人
'         If Trim(Left(CboEEP05.Text, 6)) = Trim(txtEEP03) Then
'            MsgBox "發送者與收受者不可為同一人！", vbExclamation
'            SSTab1.Tab = 0
'            Exit Function
'         End If
'         '2013/10/2 END
      End If
   'End If
   
   'Add By Sindy 2024/9/30 內專繪圖人員休假,則以操作的人抓核判權限
   '                       ex:姍珊請假時，舒郁依自己權限處理姍珊墨圖送判。
   '但控管送出時不要核判是自己,待核判區不會顯示自己的案子
   If UCase(m_PrevForm.Name) = UCase("frm090711") Then '繪圖人員工作進度
      If Trim(m_PrevForm.txt1(0)) <> Left(m_DPMan, 5) And _
         Trim(m_PrevForm.txt1(0)) = Trim(Left(CboEEP05.Text, 6)) And _
         InStr(EMP_需等待回覆的狀態, Left(CboEEP04.Text, 2)) > 0 Then
         '繪圖人員(96021)休假，收受者不宜是自己(96021)，請換其他收受者
         MsgBox "繪圖人員休假，核判人員不宜是自己，請換其他收受者！", vbExclamation
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   '2024/9/30 END
   
   'Add By Sindy 2018/9/25 條款
   If Me.Frame6.Visible = True Then
      'Add By Sindy 2018/10/12 竹平:沒輸入資料時,提醒訊息
      If Option1(0).Value = False And Option1(1).Value = False And _
         Option1(2).Value = False And Trim(txt2.Text) = "" Then
         If MsgBox("預估及條款代碼尚未輸入，確定是否繼續？", vbQuestion + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            txt2.SetFocus
            SSTab1.Tab = 0
            Exit Function
         End If
      End If
      '2018/10/12 END
      Cancel = False
      Call txt2_Validate(Cancel)
      If Cancel = True Then
         txt2.SetFocus
         Exit Function
      End If
   End If
   '2018/9/25 END
   
   'Add By Sindy 2018/4/27
   'Modify By Sindy 2024/7/12
   If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
      
      If bolFCTFlow <> True Then
         If (Left(CboEEP04.Text, 2) = EMP_送判 Or _
             ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發) Or _
             Left(CboEEP04.Text, 2) = EMP_送件 Or _
             Left(CboEEP04.Text, 2) = EMP_發文歸檔) Then 'And textCP118 = "Y"
            'cp141.送件方式=2.收款後送件 and cp79.未收金額大於0
            If cp(141) = "2" And m_dblCP79 > 0 Then
               If PUB_ChkPaidByCP09(m_EEP01) = False Then '出納繳款確認後就可送件
                  If cp(6) = "" Or Val(cp(6)) > strSrvDate(1) Then
                     'Modify By Sindy 2018/8/13
                     If Left(CboEEP04.Text, 2) = EMP_送判 Then
                        If MsgBox("此案智權人員欲管控收款後才可送件，確定要送出嗎？", vbQuestion + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                           Exit Function
                        End If
                     ElseIf Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
                        MsgBox "此案智權人員欲管控收款後才可送件，暫不可發文！", vbExclamation
                        SSTab1.Tab = 0
                        Exit Function
                     Else
                        m_SubjectNote = " [此案智權人員欲管控收款後才可送件]" 'Add By Sindy 2018/8/13
                     End If
                  End If
               End If
            End If
         End If
         '2018/4/27 END
      End If
      
      'Add By Sindy 2018/8/8
      If Left(CboEEP04.Text, 2) = EMP_送判 Or _
         ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發) Or _
         Left(CboEEP04.Text, 2) = EMP_送件 Then
         'Add By Sindy 2019/2/21
         'Added by Lydia 2015/11/24 管控台灣延展案102,發文日不可小於"延展期滿前6個月"
         'modify by sonia 2016/7/5 改為發文日不可小於"延展期滿前6個月+1天"  T-093656(法定1051224不可於1050624發文)
         'If m_TM10 = 台灣國家代號 And m_CP10 = "102" And TransDate(textCP27, 2) < CompDate(1, -6, m_CP07) Then
         'Modified by Lydia 2017/06/01 延展期滿日期改用模組控制
         'If m_TM10 = 台灣國家代號 And m_CP10 = "102" And TransDate(textCP27, 2) < CompDate(2, 1, CompDate(1, -6, m_CP07)) Then
         If m_Country = 台灣國家代號 And cp(10) = "102" And strSrvDate(1) < PUB_Get102DeadLine("3", cp(7)) Then
            If MsgBox("台灣延展案不得早於延展期滿前6個月+1天，確定要送件嗎？", vbQuestion + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Exit Function
            End If
         End If
         '2019/2/21 END
         
         '卷宗性質=1.申請, 非爭議案件性質
         'Modify By Sindy 2024/11/21 排除FCT非外商程序人員
         If Not (bolFCTFlow = True And Pub_StrUserSt03 <> "F12") Then 'F12=外商程序
            'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
            If tm(28) = "1" And (InStr(TMdebate, cp(10)) = 0 Or (cp(1) = "FCT" And InStr(FCT_NotTMdebate, cp(10)) > 0)) Then
               '檢查是否有輸入商品資料
               'Modify By Sindy 2024/6/12
   '            strSql = "select tg05 from Tmgoods where tg01='" & PField(1) & "' and tg02='" & PField(2) & "' and tg03='" & PField(3) & "' and tg04='" & PField(4) & "'"
   '            intI = 1
   '            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '            If intI = 0 Then
   '               If MsgBox("無商品資料，確定要送件嗎？", vbQuestion + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
   '                  Exit Function
   '               End If
   '            End If
               If cmd1(0).BackColor = &H8080FF Then '紅色
                  If MsgBox("無商品資料，確定要送件嗎？", vbQuestion + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     Exit Function
                  End If
               End If
               If cmdOK(3).BackColor = &H8080FF Then '紅色
                  If MsgBox("無商標描述資料，確定要送件嗎？", vbQuestion + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     Exit Function
                  End If
               End If
               '2024/6/12 END
               
               '檢查是否有代表圖
               'Modify By Sindy 2018/10/31 TF馬德里商標圖檔-子案的圖同母案
               '固定都以IBF01=tm01 AND IBF02=substr(tm02,1,5)||'0' AND IBF03='0' AND IBF04='00' 去抓代表圖
               If PField(1) = "TF" Then
                  strSql = "select ibf05 from imgbytefile where ibf01='" & PField(1) & "' and ibf02='" & Mid(PField(2), 1, 5) & "0" & "' And ibf03='0' And ibf04='00'"
               Else
               '2018/10/31 END
                  strSql = "select ibf05 from imgbytefile where ibf01='" & PField(1) & "' and ibf02='" & PField(2) & "' and ibf03='" & PField(3) & "' and ibf04='" & PField(4) & "'"
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
                  If MsgBox("無代表圖，確定要送件嗎？", vbQuestion + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     Exit Function
                  End If
               End If
            End If
         End If
         
         'Add By Sindy 2020/12/23 送件時,台灣案要點選電子送件,若沒有進去操作
         '檢查CP現況
         If Left(CboEEP04.Text, 2) = EMP_送件 Then
            'Modify By Sindy 2024/7/12 + And bolTMFlow = True
            If m_Country = "000" And bolTMFlow = True Then
               If cmdCP118.Tag = "" Then
                  strSql = "select cp118,cp85 From caseprogress"
                  If m_RetrunRecv <> "" Then
                     strSql = strSql & " where cp09 in('" & Replace(m_RetrunRecv, ",", "','") & "')"
                  Else
                     strSql = strSql & " where cp09 ='" & m_EEP01 & "'"
                  End If
                  strSql = strSql & " and cp118 is not null and nvl(cp85,0)=0"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     If RsTemp.RecordCount > 0 Then
                        MsgBox "請先進【電子送件】操作，再進行送件。", vbExclamation
                        Exit Function
                     Else
                        'Modify By Sindy 2023/12/8 檢查指定送件日
                        'Modify By Sindy 2024/1/23 改為共用函數
                        If PUB_ChkCP141IsSend(m_EEP01, False, "送件") = False Then
                           Exit Function
                        End If
                     End If
                  Else
                     'Modify By Sindy 2023/12/8 檢查指定送件日
                     'Modify By Sindy 2024/1/23 改為共用函數
                     If PUB_ChkCP141IsSend(m_EEP01, False, "送件") = False Then
                        Exit Function
                     End If
                  End If
               End If
            'Modify By Sindy 2024/7/12
            'Else
            ElseIf bolTMFlow = True Or bolCFTFlow = True Then
            '2024/7/12 END
               'Modify By Sindy 2023/12/8 檢查指定送件日
               'Modify By Sindy 2024/1/23 改為共用函數
               If PUB_ChkCP141IsSend(m_EEP01, False, "送件") = False Then
                  Exit Function
               End If
            End If
         End If
         '2020/12/23 END
      End If
      '2018/8/8 END
      
      'Add By Sindy 2024/7/12
      If bolTMFlow = True Then
      '2024/7/12 END
         '檢查查名單是否全部完成
         If Left(CboEEP04.Text, 2) = EMP_送判 Or _
            ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發) Then
            If UCase(TypeName(m_PrevForm)) = UCase("frm090201_b") Then
               If m_PrevForm.cmdTSMap.Tag = "Y" Then
                  intI = 1
                  strSql = "select tmq01,tmq10,st02 from trademarkquery,staff " & _
                               "where tmq11 is null and tmq10=st01(+) " & _
                               "and tmq01 in (select tqc03 from tmqcasemap where tqc02='" & m_EEP01 & "') order by 1"
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     RsTemp.MoveFirst
                     Do While Not RsTemp.EOF
                        MsgBox "委查單號:" & RsTemp.Fields("tmq01") & "，查名人:" & RsTemp.Fields("st02") & "，尚未查覆完畢不可送件!", vbExclamation
                        RsTemp.MoveNext
                     Loop
                     Exit Function
                  End If
               End If
            End If
         End If
         'Added by Lydia 2019/12/12 T案增加控管出名代理人：於電子歷程中之「送核、送會、送判、送件」流程，彈跳訊息提醒為第三人出名或不出名
         If Left(CboEEP04.Text, 2) = EMP_送核 Or _
             Left(CboEEP04.Text, 2) = EMP_送會 Or _
             Left(CboEEP04.Text, 2) = EMP_送判 Or _
             ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發) Or _
             Left(CboEEP04.Text, 2) = EMP_送件 Then
             If m_Country = 台灣國家代號 And cp(10) = "101" Then
                   strExc(0) = Pub_ChkTQD11(cp(9), strExc(1))
                   If Left(strExc(1), 1) = "2" Then '2=不出名
                         MsgBox "查名結果為近似本所案經核可後，設定為不出名代理！", vbInformation, "是否出名"
                   ElseIf Left(strExc(1), 1) = "1" Then '1=第三人
                         MsgBox "查名結果為近似本所案經核可後，設定為第三人出名！", vbInformation, "是否出名"
                   End If
             End If
         End If
         'end 2019/12/12
      End If
      
      'Add By Sindy 2018/11/21
      If Left(CboEEP04.Text, 2) = EMP_送核 Or _
         Left(CboEEP04.Text, 2) = EMP_核修 Or _
         Left(CboEEP04.Text, 2) = EMP_核完 Or _
         Left(CboEEP04.Text, 2) = EMP_送會 Or _
         Left(CboEEP04.Text, 2) = EMP_會修 Or _
         Left(CboEEP04.Text, 2) = EMP_會完 Or _
         Left(CboEEP04.Text, 2) = EMP_送判 Or _
         Left(CboEEP04.Text, 2) = EMP_退回 Or _
         Left(CboEEP04.Text, 2) = EMP_判發 Or _
         Left(CboEEP04.Text, 2) = EMP_送件 Then
         'Add By Sindy 2024/7/12
         If m_Country = 台灣國家代號 Then
         '2024/7/12 END
            '異議案逾法定期限不可發文
            If cp(10) = "601" And Val(cp(7)) > 0 Then
               If strSrvDate(1) > cp(7) Then
                  MsgBox "異議案逾法定期限不可發文!", vbInformation
                  'Add By Sindy 2018/12/3
                  If Left(CboEEP04.Text, 2) = EMP_送件 Then
                     Exit Function
                  End If
                  '2018/12/3 END
               End If
            End If
            '廢止案不可提早發文,管制期限存在CP46,發文存檔時要清除
            'Modify By Sindy 2018/12/3 + 623.部分廢止
            If (cp(10) = "605" Or cp(10) = "623") And Val(cp(46)) > 0 Then
               If strSrvDate(1) < cp(46) Then
                  MsgBox "廢止案未達管制期限 (公告" & IIf(m_Country = "000", "滿", "期滿加") & "三年 " & ChangeTStringToTDateString(ChangeWStringToTString(cp(46))) & ") 不可提早發文!", vbInformation
                  'Add By Sindy 2018/12/3
                  If Left(CboEEP04.Text, 2) = EMP_送件 Then
                     Exit Function
                  End If
                  '2018/12/3 END
               End If
            End If
         End If
      End If
      
   ElseIf bolPAFlow = True Then
      If Left(CboEEP04.Text, 2) = EMP_送核 Or _
         Left(CboEEP04.Text, 2) = EMP_送英核 Or _
         Left(CboEEP04.Text, 2) = EMP_送會 Or _
         Left(CboEEP04.Text, 2) = EMP_送判 Or _
         Left(CboEEP04.Text, 2) = EMP_判發 Then
         'Add By Sindy 2014/3/13 當承辦人為專利處繪圖的人員時,不需檢查此條件
         If PUB_GetStaffST15(Left(m_EPMan, 5), "1") <> "P13" Then
         '2014/3/13 END
            '檢查承辦單內容是否存在
            strSql = "select * From EmpElectronData where eed01='" & m_EEP01 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Or RsTemp.RecordCount = 0 Then
               MsgBox "請輸入承辦單內容，才可執行送出！", vbExclamation
               SSTab1.Tab = intTab_承辦單
               Exit Function
            End If
         End If
      End If
      
   'Add By Sindy 2023/12/14
   ElseIf bolFCPFlow = True Then
      'Add By Sindy 2024/3/8
      If Left(CboEEP04.Text, 2) = EMP_送核 Or _
         Left(CboEEP04.Text, 2) = EMP_送判 Or _
         Left(CboEEP04.Text, 2) = EMP_送件 Or _
         Left(CboEEP04.Text, 2) = EMP_發文歸檔 Or _
         Left(CboEEP04.Text, 2) = EMP_退件重送 Then
         ChkCP113CP114
      End If
      '2024/3/8 END
      
      If (Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_發文歸檔) Then
         '尚有待客戶最終指示未發文未取消收文，不可送件
         strSql = "select * From caseprogress" & _
                  " where cp01='" & PField(1) & "' and cp02='" & PField(2) & "' and cp03='" & PField(3) & "' and cp04='" & PField(4) & "'" & _
                  " and cp10='970' and cp43='" & m_EEP01 & "' and cp158=0 and cp159=0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            MsgBox "尚有待客戶最終指示未發文未取消收文，不可送件！", vbExclamation
            Exit Function
         End If
      End If
      
      'Add By Sindy 2025/10/22
      If (Left(CboEEP04.Text, 2) = EMP_送件 Or _
          Left(CboEEP04.Text, 2) = EMP_退件重送 _
         ) And bolCmdF21 = True Then
         
         'FCP設計案在發文(210)製作中說時，工程師須上傳DES
         strExc(9) = GetFCPPathVal(Label25(0).Caption, PField(1), PField(2), Trim(lblCP10.Caption), True)
         Me.Label25(0) = strExc(9)
         If pa(8) = "3" And cp(10) = "210" Then
            strShowMsg = "1.請上傳設計案之說明書和圖式至 " & strExc(9) & " 或 附件區 以利後續歸原始檔區!" & vbCrLf & _
                         "(以供承辦請款時寄給代理人)"
         Else
            strExc(10) = Pub_GetCP31toCP27(pa(1), pa(2), pa(3), pa(4)) '新申請案發文日
            '主動修正203、修正204、誤譯訂正433和申復、再審發文(有一併修正)時，工程師須上傳中說word檔最終版本
            If (InStr("107,205", cp(10)) > 0 And cp(148) = "Y") Or _
               (InStr("203,204,433", cp(10)) > 0 And cp(148) = "") Then
               If strExc(10) <> "" And _
                  ((cp(10) = "203" And strSrvDate(1) > strExc(10)) Or _
                  (cp(10) <> "203" And strSrvDate(1) >= strExc(10))) Then '判斷提申後才檢查
                  strShowMsg = "2.請上傳中說Word檔最終版本至 " & strExc(9) & " 或 附件區 以利後續歸原始檔區!"
               End If
            End If
         End If
         File1.Tag = "Y"
         If Dir(strExc(9), vbDirectory) <> "" Then
            File1.path = strExc(9)
            File1.Refresh
            If File1.ListCount = 0 Then
               File1.Tag = ""
            Else
               '檢查 電子送件暫存區 和 附件區 是否有重覆檔案
               For dblFCnt = 0 To File1.ListCount - 1
                  For intQ = 0 To lstAtt(0).ListCount - 1
                     If InStr(UCase(lstAtt(0).List(intQ)), UCase(File1.List(dblFCnt))) > 0 Then
                        MsgBox File1.List(dblFCnt) & " 檔案，重覆存放在 附件區" & vbCrLf & _
                               "及 資料夾( " & vbCrLf & strExc(9) & " )" & vbCrLf & _
                               "請擇一刪除，避免無法歸卷！", vbExclamation
                        Exit Function
                     End If
                  Next intQ
               Next dblFCnt
               File1.Refresh
               If File1.ListCount = 0 Then
                  File1.Tag = ""
               End If
            End If
         Else
            File1.Tag = ""
         End If
         If strShowMsg <> "" Then
            '1=FCP設計案在發文(210)製作中說時，工程師須上傳DES
            bolCheck = False
            If Left(strShowMsg, 1) = "1" Then
               If File1.Tag = "Y" Then
                  For dblFCnt = 0 To File1.ListCount - 1
                     If UCase(Right(Trim(File1.List(dblFCnt)), 8)) = ".DES.DOC" Or _
                        UCase(Right(Trim(File1.List(dblFCnt)), 9)) = ".DES.DOCX" Or _
                        UCase(Right(Trim(File1.List(dblFCnt)), 8)) = ".DES.PDF" Then
                        bolCheck = True
                        Exit For
                     End If
                  Next dblFCnt
               End If
               If bolCheck = False Then
                  For dblFCnt = 0 To lstAtt(0).ListCount - 1
                     If UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 8)) = ".DES.DOC" Or _
                        UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 9)) = ".DES.DOCX" Or _
                        UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 8)) = ".DES.PDF" Then
                        bolCheck = True
                        Exit For
                     End If
                  Next dblFCnt
               End If
            '2=上傳中說word檔最終版本
            Else
               If File1.Tag = "Y" Then
                  For dblFCnt = 0 To File1.ListCount - 1
                     If InStr(UCase(File1.List(dblFCnt)), ".FIX_U.") > 0 Or _
                        InStr(UCase(File1.List(dblFCnt)), ".FIX.") > 0 Or _
                        UCase(Right(Trim(File1.List(dblFCnt)), 8)) = ".FIG.PDF" Then
                        bolCheck = True
                        Exit For
                     End If
                  Next dblFCnt
               End If
               If bolCheck = False Then
                  For dblFCnt = 0 To lstAtt(0).ListCount - 1
                     If InStr(UCase(Trim(GetFileName(lstAtt(0).List(dblFCnt)))), ".FIX_U.") > 0 Or _
                        InStr(UCase(Trim(GetFileName(lstAtt(0).List(dblFCnt)))), ".FIX.") > 0 Or _
                        UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 8)) = ".FIG.PDF" Then
                        bolCheck = True
                        Exit For
                     End If
                  Next dblFCnt
               End If
            End If
            If bolCheck = False Then
               MsgBox Mid(strShowMsg, 3), vbExclamation
               Exit Function
            End If
            '至少要有一個檔案(*.DOC、*.DOCX、*.TXT、*.XML、.FIG.PDF、.RES.PDF、.SEP.PDF)
            bolCheck = False
            If File1.Tag = "Y" Then
               For dblFCnt = 0 To File1.ListCount - 1
                  If UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".DOC" Or _
                     UCase(Right(Trim(File1.List(dblFCnt)), 5)) = ".DOCX" Or _
                     UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".TXT" Or _
                     UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".XML" Or _
                     UCase(Right(Trim(File1.List(dblFCnt)), 8)) = ".FIG.PDF" Or _
                     UCase(Right(Trim(File1.List(dblFCnt)), 8)) = ".RES.PDF" Or _
                     UCase(Right(Trim(File1.List(dblFCnt)), 8)) = ".SEP.PDF" Then
                     bolCheck = True
                     Exit For
                  End If
               Next dblFCnt
            End If
            If bolCheck = False Then
               For dblFCnt = 0 To lstAtt(0).ListCount - 1
                  If UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 4)) = ".DOC" Or _
                     UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 5)) = ".DOCX" Or _
                     UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 4)) = ".TXT" Or _
                     UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 4)) = ".XML" Or _
                     UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 8)) = ".FIG.PDF" Or _
                     UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 8)) = ".RES.PDF" Or _
                     UCase(Right(Trim(GetFileName(lstAtt(0).List(dblFCnt))), 8)) = ".SEP.PDF" Then
                     bolCheck = True
                     Exit For
                  End If
               Next dblFCnt
            End If
            If bolCheck = False Then
               MsgBox "請上傳檔案至 " & strExc(9) & " 或 附件區 以利後續歸原始檔區!", vbExclamation
               Exit Function
            End If
         End If
      End If
      '2025/10/22 END
   '2023/12/14 END
   End If
   
   bolMoveFile = False
   '(((m_CMMan = "" And m_CSMan = "") Or (Left(m_CMMan, 5) = Left(m_EPMan, 5)) Or (Left(m_CSMan, 5) = Left(m_EPMan, 5))) And Left(CboEEP04.Text, 2) = EMP_判發)
   If Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_發文歸檔 Or _
      Left(CboEEP04.Text, 2) = EMP_退件重送 Or Left(CboEEP04.Text, 2) = EMP_送判 Or _
      ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發) Then
      If lstAtt(0).ListCount = 0 Then
'         'Modify By Sindy 2023/10/26 排除 PUB_GetST03(strUserNum) = "F22" 淑華說她們程序都放Server上暫存區,發文時已會歸入卷宗區
'         If PUB_GetST03(strUserNum) = "F22" Then
'            '不鎖附件
'            If Trim(txtEEP08.Text) = "" Then
'               MsgBox "無附件！" & vbCrLf & "內容，不可空白！", vbExclamation
'               SSTab1.Tab = 0
'               txtEEP08.SetFocus
'               Exit Function
'            End If
'         'Modify By Sindy 2024/3/15 + Or bolFCPFlow = True
'         ElseIf PUB_GetST03(strUserNum) = "F21" Or bolFCPFlow = True Then
         'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True
         If bolFCPFlow = True Or bolFCTFlow = True Then
            'Modify By Sindy 2023/11/16 薛經理說不用提醒
'            If MsgBox("是否要放入最終完整附件？（送件後方可歸卷）", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
'               SSTab1.Tab = 0
'               Exit Function
'            Else
            '2023/11/16 END
            If Trim(txtEEP08.Text) = "" Then
               MsgBox "無附件！" & vbCrLf & "內容，不可空白！", vbExclamation
               SSTab1.Tab = 0
               txtEEP08.SetFocus
               Exit Function
            End If
         Else
         '2023/10/26 END
            MsgBox "請放入最終完整附件！", vbExclamation
            SSTab1.Tab = 0
            Exit Function
         End If
      End If
   Else
      'Left(CboEEP04.Text, 2) = EMP_送標號 'Modify By Sindy 2013/8/22 送標號不鎖附件
      'Modify By Sindy 2013/9/23 +附加流程
      'Modify By Sindy 2013/11/13 +(Check1.Visible = True And Check1.Value = 1):一併更新英文核完日
      '上墨:惟若圖式有修改,則工程師要另外進行"上墨"流程,此時,上墨的日期則會update為繪圖的墨齊日。
      If Left(CboEEP04.Text, 2) = EMP_草完 Or _
         Left(CboEEP04.Text, 2) = EMP_草核 Or _
         Left(CboEEP04.Text, 2) = EMP_標號 Or _
         Left(CboEEP04.Text, 2) = EMP_送核 Or _
         Left(CboEEP04.Text, 2) = EMP_送英核 Or _
         Left(CboEEP04.Text, 2) = EMP_送會 Or _
         Left(CboEEP04.Text, 2) = EMP_上墨 Or _
         Left(CboEEP04.Text, 2) = EMP_墨完 Or _
         (m_DPMan = m_DMMan And Left(CboEEP04.Text, 2) = EMP_繪圖判發) Or _
         Left(CboEEP04.Text, 2) = EMP_附加流程 Or _
         (Check1.Visible = True And Check1.Value = 1) Then
         If lstAtt(0).ListCount = 0 Then
'            'Modify By Sindy 2023/10/26 排除 PUB_GetST03(strUserNum) = "F22" 淑華說她們程序都放Server上暫存區,發文時已會歸入卷宗區
'            If PUB_GetST03(strUserNum) = "F22" Then
'               '不鎖附件
'               If Trim(txtEEP08.Text) = "" Then
'                  MsgBox "無附件！" & vbCrLf & "內容，不可空白！", vbExclamation
'                  txtEEP08.SetFocus
'                  SSTab1.Tab = 0
'                  Exit Function
'               End If
'            'Modify By Sindy 2024/3/15 + Or bolFCPFlow = True
'            ElseIf PUB_GetST03(strUserNum) = "F21" Or bolFCPFlow = True Then
            'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True
            If bolFCPFlow = True Or bolFCTFlow = True Then
               'Modify By Sindy 2023/11/16 薛經理說不用提醒
'               If MsgBox("是否要放入附件？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
'                  SSTab1.Tab = 0
'                  Exit Function
'               End If
               'Add By Sindy 2025/7/31
               If Not (Left(CboEEP04.Text, 2) = EMP_附加流程 And _
                       Trim(Left(CboCP10.Text, 4)) <> 延期 And Trim(Mid(CboCP10.Text, 4)) <> "延期") Then
               '2025/7/31 END
                  If Trim(txtEEP08.Text) = "" Then
                     MsgBox "無附件！" & vbCrLf & "內容，不可空白！", vbExclamation
                     txtEEP08.SetFocus
                     SSTab1.Tab = 0
                     Exit Function
                  End If
               End If
            Else
            '2023/10/26 END
               'Add By Sindy 2014/1/15 延期是系統直接判發或送判所以要夾帶附件
               'Modify By Sindy 2019/7/10 + T 734.代理人撰稿
               'Modify By Sindy 2025/7/31
'               If Not (Left(CboEEP04.Text, 2) = EMP_附加流程 And _
'                       (((bolPAFlow = True Or bolFCPFlow = True) And (Trim(Left(CboCP10.Text, 4)) = "936" Or _
'                        Trim(Left(CboCP10.Text, 4)) = "957" Or _
'                        Trim(Left(CboCP10.Text, 4)) = "958")) Or _
'                       (bolTMFlow = True And Trim(Left(CboCP10.Text, 4)) = "734")) _
'                      ) Then
               If Not (Left(CboEEP04.Text, 2) = EMP_附加流程 And _
                       Trim(Left(CboCP10.Text, 4)) <> 延期 And Trim(Mid(CboCP10.Text, 4)) <> "延期") Then
               '2025/7/31 END
               '2014/1/15 END
                  MsgBox "請放入附件！", vbExclamation
                  SSTab1.Tab = 0
                  Exit Function
               End If
            End If
         End If
      'Modify By Sindy 2016/3/15 +EMP_圖修
      ElseIf Left(CboEEP04.Text, 2) = EMP_核修 Or _
             Left(CboEEP04.Text, 2) = EMP_會修 Or _
             Left(CboEEP04.Text, 2) = EMP_圖修 Or _
             Left(CboEEP04.Text, 2) = EMP_草修 Or _
             Left(CboEEP04.Text, 2) = EMP_退回 Then
         If lstAtt(0).ListCount = 0 Then
            'Modify By Sindy 2023/10/26 排除 PUB_GetST03(strUserNum) = "F22" 淑華說她們程序都放Server上暫存區,發文時已會歸入卷宗區
            'Modify By Sindy 2024/3/15 + 改判斷 bolFCPFlow = True
            'If PUB_GetST03(strUserNum) = "F22" Or PUB_GetST03(strUserNum) = "F21" Then
            'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True
            If bolFCPFlow = True Or bolFCTFlow = True Then
               '不鎖附件
               If Trim(txtEEP08.Text) = "" Then
                  MsgBox "無附件！" & vbCrLf & "內容，不可空白！", vbExclamation
                  txtEEP08.SetFocus
                  SSTab1.Tab = 0
                  Exit Function
               End If
            Else
            '2023/10/26 END
               If MsgBox("確定沒有附件嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
                  Exit Function
               End If
            End If
         End If
      'Modify By Sindy 2016/3/15 +EMP_圖完
      ElseIf Left(CboEEP04.Text, 2) = EMP_會完 Or Left(CboEEP04.Text, 2) = EMP_圖完 Then
         If lstAtt(0).ListCount = 0 Then
            If MsgBox("要沿用「原附件」嗎？" & vbCrLf & _
                      "按『是』：系統會將前一道附件移至此流程。" & vbCrLf & _
                      "按『否』：代表無附件！", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
               If GetPreviousFlow(m_EEP01, intLastEEP02, Left(CboEEP04.Text, 2)) = False Then
                  Exit Function
               End If
               bolMoveFile = True
            End If
         End If
      'Modify By Sindy 2023/10/31 +(bolFCPFlow = True And strEEP04 = EMP_送件 And bolWaitReply = True)
      'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True)
      ElseIf Left(CboEEP04.Text, 2) = EMP_核完 Or _
             Left(CboEEP04.Text, 2) = EMP_草核完 Or _
             Left(CboEEP04.Text, 2) = EMP_繪圖判發 Or _
             Left(CboEEP04.Text, 2) = EMP_判發 Or _
             ((bolFCPFlow = True Or bolFCTFlow = True) And Left(CboEEP04.Text, 2) = EMP_送件 And bolWaitReply = True) Then
         If lstAtt(0).ListCount = 0 Then
            'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True
            If bolFCPFlow = True Or bolFCTFlow = True Then
               If GetPreviousFlow(m_EEP01, intLastEEP02, Left(CboEEP04.Text, 2), False) = True And m_PreviousFlow <> "" Then
                  'Modify By Sindy 2024/1/11 可不沿用;增加取消
                  intMsgQ = MsgBox("要沿用「原附件」嗎？" & vbCrLf & _
                                   "按『是』：沿用檔案，系統會將前一道附件移至此流程。" & vbCrLf & _
                                   "按『否』：不沿用送出！" & vbCrLf & _
                                   "按『取消』：放棄此動作。", vbExclamation + vbYesNoCancel + vbDefaultButton3, "重要訊息！")
                  If intMsgQ = vbYes Then
                     bolMoveFile = True
                  ElseIf intMsgQ = vbNo Then
                     m_PreviousFlow = "" '*****
                     If Trim(txtEEP08.Text) = "" Then
                        MsgBox "內容及附件區至少擇一輸入資料！", vbExclamation
                        txtEEP08.SetFocus
                        SSTab1.Tab = 0
                        Exit Function
                     End If
                  Else
                     m_PreviousFlow = "" '*****
                     Exit Function
                  End If
'                  If MsgBox("要沿用「原附件」嗎？" & vbCrLf & _
'                            "按『是』：沿用檔案，系統會將前一道附件移至此流程。" & vbCrLf & _
'                            "按『否』：不沿用！請自行添加附件。", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
'                     bolMoveFile = True
'                  '*****
'                  Else
'                     m_PreviousFlow = ""
'                     'Modify By Sindy 2024/1/11 允許可以不沿用
'                     'Exit Function
'                  '*****
'                  End If
                  '2024/1/11 END
               End If
            ElseIf MsgBox("要沿用「原附件」嗎？" & vbCrLf & _
                   "按『是』：沿用檔案，系統會將前一道附件移至此流程。" & vbCrLf & _
                   "按『否』：不沿用！請自行添加附件。", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
               If GetPreviousFlow(m_EEP01, intLastEEP02, Left(CboEEP04.Text, 2)) = False Then
                  Exit Function
               End If
               bolMoveFile = True
            '*****
            Else
               Exit Function
            '*****
            End If
         End If
      'Modify By Sindy 2025/8/1 + Or Left(CboEEP04.Text, 2) = EMP_收文分析
      ElseIf Left(CboEEP04.Text, 2) = EMP_轉回 Or Left(CboEEP04.Text, 2) = EMP_客戶會稿 Or Left(CboEEP04.Text, 2) = EMP_收文分析 Then
         '內容及附件區均可空白
      Else
         If Trim(txtEEP08.Text) = "" And lstAtt(0).ListCount = 0 Then
            'Add By Sindy 2023/10/26
            If Left(CboEEP04.Text, 2) = EMP_交辦 Then
               If MsgBox("確定沒有內容及附件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Exit Function
               End If
            Else
            '2023/10/26 END
               MsgBox "內容及附件區至少擇一輸入資料！", vbExclamation
               txtEEP08.SetFocus
               SSTab1.Tab = 0
               Exit Function
            End If
         End If
      End If
   End If
   
   '************************************************************************
   'Add By Sindy 2018/7/16 比對Word檔,若無相對應檔名.PDF檔,就自動產生一份
'   Left(CboEEP04.Text, 2) = EMP_送判 Or _
'         ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發)
   If bolTMFlow = True Then
      If Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_發文歸檔 Or _
         Left(CboEEP04.Text, 2) = EMP_退件重送 Then
         Screen.MousePointer = vbHourglass
         'Modify By Sindy 2019/1/21
         'Call AutoPrintPDFfile(0)
         If AutoPrintPDFfile(0) = False Then
            Screen.MousePointer = vbDefault
            Exit Function
         End If
         '2019/1/21 END
         Screen.MousePointer = vbDefault
      End If
   End If
   '2018/4/27 END
   '************************************************************************
   
   '有繪圖人員
   IsExistsDWG = False
   If m_DPMan <> "" Then
      '檢查是否為相同案是否已有附dwg原始圖
      '新申請案:NewCasePtyList
      stVTB = PUB_GetSameCaseSQL(m_EEP01) '相同案語法(收文號)
      'Modify By Sindy 2014/3/11 +dwg.7z
      'Modify By Sindy 2018/9/4 + and eef12 is not null
      strSql = "select eef01,eef02,eef03,eep04 from empelectronfile,empelectronprocess," & _
               "(select cp09 from caseprogress," & _
               "(" & stVTB & ") V1" & _
               " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
               " and substr(V1.cno,-9,6)=cp02" & _
               " and substr(V1.cno,-3,1)=cp03" & _
               " and substr(V1.cno,-2)=cp04" & _
               " and cp10 in(" & IIf(InStr(NewCasePtyList, cp(10)) > 0, NewCasePtyList, cp(10)) & ")) V2" & _
               " Where V2.CP09 = eef01" & _
               " and (substr(upper(eef03),-4)='.DWG' or substr(upper(eef03),-7)='DWG.ZIP' or substr(upper(eef03),-6)='DWG.7Z')" & _
               " and eef01=eep01" & _
               " and eef02=eep02 and eef12 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            IsExistsDWG = True
         End If
      End If
   End If
   
   'Add By Sindy 2014/1/15
   'Modify By Sindy 2025/7/31
'   If Not (Left(CboEEP04.Text, 2) = EMP_附加流程 And _
'           (Trim(Left(CboCP10.Text, 4)) = "936" Or _
'            Trim(Left(CboCP10.Text, 4)) = "957" Or _
'            Trim(Left(CboCP10.Text, 4)) = "958") _
'          ) Then
   If Not (Left(CboEEP04.Text, 2) = EMP_附加流程 And _
           Trim(Left(CboCP10.Text, 4)) <> 延期 And Trim(Mid(CboCP10.Text, 4)) <> "延期") Then
   '2025/7/31 END
   '2014/1/15 END
      '檢查附件內容是否符合規定:
      intPDF = 0: intDOC = 0: intDwgPdf = 0: intDwg = 0
      intPoaPDF = 0 'Add By Sindy 2018/7/13
      m_intDataPDF = 0 'Modify By Sindy 2020/3/5
      intDataDOC = 0
      intCUSPDF = 0 'Add By Sindy 2018/9/26
      For ii = 0 To lstAtt(0).ListCount - 1
         If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 4) = ".PDF" Then
            intPDF = intPDF + 1
         'Modify By Sindy 2013/12/25 也可放docx檔
         'Modify By Sindy 2014/6/25 也可放txt檔,因序列表
         ElseIf Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 4) = ".DOC" Or _
                Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 5) = ".DOCX" Or _
                Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 4) = ".TXT" Then
            intDOC = intDOC + 1
         End If
         If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 7) = "DWG.PDF" Then
            intDwgPdf = intDwgPdf + 1
         End If
         'Modify By Sindy 2014/3/11 +dwg.7z
         If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 4) = ".DWG" Or _
            Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 7) = "DWG.ZIP" Or _
            Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 6) = "DWG.7Z" Then
            intDwg = intDwg + 1
         End If
         'Add By Sindy 2018/7/13
         If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 8) = ".POA.PDF" Then
            intPoaPDF = intPoaPDF + 1
         'Add By Sindy 2018/9/25
         ElseIf Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 7) = "POA.PDF" Then
            intPoaPDF = intPoaPDF + 1
         '2018/9/25 END
         End If
         If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 9) = ".DATA.PDF" Then
            m_intDataPDF = m_intDataPDF + 1
         End If
         '2018/7/13 END
         'Add By Sindy 2019/5/16
         If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 9) = ".DATA.DOC" Or _
            Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 10) = ".DATA.DOCX" Then
            intDataDOC = intDataDOC + 1
         End If
         '2019/5/16 END
         'Add By Sindy 2018/9/25
         If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 8) = ".CUS.PDF" Then '客戶函
            intCUSPDF = intCUSPDF + 1
         End If
         '2018/9/25 END
      Next ii
      
      'Add By Sindy 2018/5/3
      If bolTMFlow = True Then
         'Add By Sindy 2020/11/19 T-230257:604.評定答辯(竹平反應,應該要有DOC)
         If (Left(CboEEP04.Text, 2) = EMP_送核 Or _
             Left(CboEEP04.Text, 2) = EMP_送判) And _
            intDOC = 0 Then
            If MsgBox("確定無.DOC附件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Exit Function
            End If
         '2020/11/19 END
         ElseIf Left(CboEEP04.Text, 2) = EMP_送件 Or _
            Left(CboEEP04.Text, 2) = EMP_退件重送 Or _
            Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
'            '均至少有一個.PDF檔,非電子送件時至少有一個.DOC附件
'            If intPDF < 1 Then
'               MsgBox "至少要有一個.PDF附件！", vbExclamation
'               SSTab1.Tab = 0
'               Exit Function
'            End If
            'Modify By Sindy 2018/7/26
            '非電子送件時至少要有1個.PDF檔
            If cp(118) = "" Then
               If intPDF < 1 Then
                  MsgBox "至少要有1個.PDF附件！", vbExclamation
                  SSTab1.Tab = 0
                  If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/3
                  Exit Function
               End If
'               If intDOC < 1 Then
'                  MsgBox "至少要有一個.DOC附件！", vbExclamation
'                  SSTab1.Tab = 0
'                  Exit Function
'               End If
            '電子送件時至少要有2個.PDF檔(Data,CONTACT)
            Else
               If intPDF < 2 Then
                  MsgBox "至少要有2個.PDF附件！", vbExclamation
                  SSTab1.Tab = 0
                  If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/3
                  Exit Function
               End If
            End If
            
            'Modify By Sindy 2018/7/13 A類大陸案一律要有指示信
            If Left(lblCP09.Caption, 1) = "A" And m_Country = 大陸國家代號 And m_intDataPDF = 0 Then
               MsgBox "沒有 .Data.PDF 指示信,不可送件！", vbExclamation
               SSTab1.Tab = 0
               If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/1
               Exit Function
            End If
            
            'Add By Sindy 2021/7/19
            '有.Data.doc 檢查是否有.Data.PDF
            If intDataDOC > 0 And m_intDataPDF = 0 Then
               MsgBox "有.Data.DOC 卻沒有.DATA.PDF 不可送件！", vbExclamation
               SSTab1.Tab = 0
               Exit Function
            End If
            '2021/7/19 END
            
            'Add By Sindy 2018/7/13
            'A類沒有.POA.PDF要提醒
            If Left(lblCP09.Caption, 1) = "A" And intPoaPDF = 0 Then
               If MsgBox("沒有 .POA.PDF 確定要送件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Exit Function
               End If
            End If
            '2018/7/13 END
            
            'Add By Sindy 2018/9/26 C類一律要有客戶函
            If Left(lblCP09, 1) = "C" And intCUSPDF = 0 Then
               MsgBox "沒有 .CUS.PDF 客戶函,不可送件！" & vbCrLf & vbCrLf & "(客戶函請修正檔名為.CUS.PDF)", vbExclamation
               SSTab1.Tab = 0
               If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/1
               Exit Function
            End If
         End If
         
      'Add By Sindy 2021/9/7
      ElseIf bolOtherFlow = True Then
         If (Left(CboEEP04.Text, 2) = EMP_送核 Or _
             Left(CboEEP04.Text, 2) = EMP_送判) And _
            intDOC = 0 Then
            
'            If MsgBox("確定無.DOC附件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'               Exit Function
'            End If
            
         ElseIf Left(CboEEP04.Text, 2) = EMP_判發 Or _
            Left(CboEEP04.Text, 2) = EMP_退件重送 Then
            
'            '有.Data.doc 檢查是否有.Data.PDF
'            If intDataDOC > 0 And m_intDataPDF = 0 Then
'               MsgBox "有.Data.DOC 卻沒有.DATA.PDF 不可送件！", vbExclamation
'               SSTab1.Tab = 0
'               Exit Function
'            End If
            
'            'A類沒有.POA.PDF要提醒
'            If Left(lblCP09.Caption, 1) = "A" And intPoaPDF = 0 Then
'               If MsgBox("沒有 .POA.PDF 確定要送件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'                  Exit Function
'               End If
'            End If
            
'            'C類一律要有客戶函
'            If Left(lblCP09, 1) = "C" And intCUSPDF = 0 Then
'               MsgBox "沒有 .CUS.PDF 客戶函,不可送件！" & vbCrLf & vbCrLf & "(客戶函請修正檔名為.CUS.PDF)", vbExclamation
'               SSTab1.Tab = 0
'               If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/1
'               Exit Function
'            End If
         End If
         
      ElseIf bolPAFlow = True Then
      '2018/5/3 END
         'Modify By Sindy 2019/12/5 專利程序承辦的案件,可以不一定要有pdf檔案,但一定要夾帶檔案
         If PUB_GetST03(Left(m_EPMan, 5)) = "P12" Then
         Else
         '2019/12/5 END
            'Add By Sindy 2019/5/16 工程師需要備好電子送件申請書
            '電子送件
            If cp(118) <> "" And _
               (Left(CboEEP04.Text, 2) = EMP_送判 Or _
                Left(CboEEP04.Text, 2) = EMP_判發 Or _
                Left(CboEEP04.Text, 2) = EMP_退件重送) Then
               'Modified by Morgan 2024/11/18 +447再審查加速審查
               If InStr("101,102,103,104,105,107,125,203,204,205,206,227,239" & _
                  ",301,302,303,304,305,306,307,308,309,401,402,403,404,407" & _
                  ",421,422,425,431,434,447,807", cp(10)) > 0 And _
                  intDataDOC = 0 Then
               'Modify By Sindy 2019/6/5 有工程師和玲玲反應不需要每個歷程都要彈此訊息
   '            If Left(CboEEP04.Text, 2) = EMP_送核 Or _
   '               Left(CboEEP04.Text, 2) = EMP_送會 Or _
   '               Left(CboEEP04.Text, 2) = EMP_送判 Or _
   '               Left(CboEEP04.Text, 2) = EMP_判發 Or _
   '               Left(CboEEP04.Text, 2) = EMP_退件重送 Then
                  MsgBox "必須有DATA.DOC附件！", vbExclamation
                  SSTab1.Tab = 0
                  If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True
                  Exit Function
               End If
               'Add By Sindy 2019/7/19 關於之前週會提到 205.申復 時可能會漏送申復理由書[附送書件]誤刪
               If cp(10) = "205" Then
                  If MsgBox("請確認【附送書件】的【申復書】等欄位是否存在？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     Exit Function
                  End If
               End If
               '2019/7/19 END
            End If
            
            '送判或判發並且為自行核判時,必須至少一個doc二個pdf檔案
            '((m_CMMan = "" Or Left(m_CMMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發)
            '(((m_CMMan = "" And m_CSMan = "") Or (Left(m_CMMan, 5) = Left(m_EPMan, 5)) Or (Left(m_CSMan, 5) = Left(m_EPMan, 5))) And Left(CboEEP04.Text, 2) = EMP_判發)
            'Modify By Sindy 2013/9/23 +附加流程
            If Left(CboEEP04.Text, 2) = EMP_送判 Or _
               Left(CboEEP04.Text, 2) = EMP_退件重送 Or _
               ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發) Or _
               (lstAtt(0).ListCount <> 0 And Left(CboEEP04.Text, 2) = EMP_判發) Or _
               Left(CboEEP04.Text, 2) = EMP_附加流程 Then
               
               'Add By Sindy 2013/10/25 933.覆函
               If (cp(10) = 修正 Or cp(10) = 933) And intDOC = 0 Then
                  If MsgBox("確定無.DOC附件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     Exit Function
                  End If
               Else
               '2013/10/25 END
                  'Modify By Sindy 2013/10/16 202補文件不鎖doc檔
                  'If Left(CboEEP04.Text, 2) <> EMP_附加流程 And bolBCaseFlow = False Then
                  'Add By Sindy 2014/3/13 當承辦人為專利處繪圖的人員時,不需檢查此條件
                  If Left(CboEEP04.Text, 2) <> EMP_附加流程 And _
                     bolBCaseFlow = False And _
                     cp(10) <> 補文件 And _
                     PUB_GetStaffST15(Left(m_EPMan, 5), "1") <> "P13" Then
                  '2013/10/16 END
                     If intDOC < 1 Then
                        MsgBox "至少要有一個.DOC附件！", vbExclamation
                        SSTab1.Tab = 0
                        If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/3
                        Exit Function
                     End If
                  End If
               End If
               
               'Add By Sindy 2013/9/23
               If Left(CboEEP04.Text, 2) = EMP_附加流程 Or bolBCaseFlow = True Then
                  'Modify By Sindy 2014/9/26 P-102610(延期-申復):附加流程若為電子送件時不需控管至少要有一個.pdf附件
                  If cp(118) = "" Then
                  '2014/9/26 END
                     If intPDF < 1 Then
                        MsgBox "至少要有一個.pdf附件！", vbExclamation
                        SSTab1.Tab = 0
                        If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/3
                        Exit Function
                     End If
                  End If
               Else
               '2013/9/23 END
         '         '無繪圖並且電子送件
         '         If m_DPMan = "" And cp(118) <> "" Then
         '            '不需要繪圖及說明書的PDF檔
                  'Modify By Sindy 2013/10/17
                  '電子送件
                  If cp(118) <> "" Then
                     '不需要繪圖及說明書的PDF檔
                  '2013/10/17 END
                  
      '            '有繪圖並且非電子送件
      '            'Add By Sindy 2014/3/13 當承辦人為P13.專利處繪圖的人員時,不需檢查此條件
      '            ElseIf m_DPMan <> "" And cp(118) = "" And PUB_GetStaffST15(Left(m_EPMan, 5), "1") <> "P13" Then
      '               If intPDF < 2 Then
      '                  MsgBox "至少要有二個.pdf附件！", vbExclamation
      '                  SSTab1.Tab = 0
      '                  Exit Function
      '               End If
      '            Else
      '   '            '有繪圖 或 非電子送件
      '   '            If m_DPMan <> "" Or _
      '   '               cp(118) = "" Then
      '               'Modify By Sindy 2013/10/17
      '               '有繪圖
      '               If m_DPMan <> "" Then
      '               '2013/10/17 END
      '                  If intPDF < 1 Then
      '                     MsgBox "至少要有一個.pdf附件！", vbExclamation
      '                     SSTab1.Tab = 0
      '                     Exit Function
      '                  End If
      '               End If
      '            End If
                  
                  'Modify By Sindy 2014/9/5 重寫此段邏緝
                  '非電子送件
                  'Add By Sindy 2014/3/13 當承辦人為P13.專利處繪圖的人員時,不需檢查此條件
                  'Modify By Sindy 2014/9/5 增加P大陸新案不用控管一定要2個PDF，因有可能無文的PDF
                  'Modify By Sindy 2016/12/14 賴健桓提P大陸新案控管新型和外觀設計一定要有dwg.pdf的檔案
                  'ElseIf PUB_GetStaffST15(Left(m_EPMan, 5), "1") <> "P13" And bolP020NewCase = False Then
                  ElseIf PUB_GetStaffST15(Left(m_EPMan, 5), "1") <> "P13" Then
                     'Modify By Sindy 2017/1/20 先前賴主任有提”dwg.pdf限制為送判的要件，但是國外部的並不需要這個要件，還請協助排除國外部的新申請案排除這個應用
                     'If bolP020NewCase = True And (m_PA08 = "2" Or m_PA08 = "3") Then
                     'Modify By Sindy 2017/11/24 品薇要求放開此控管
      '               If bolP020NewCase = True And (m_PA08 = "2" Or m_PA08 = "3") And Left(CP(12), 1) <> "F" Then
      '               '2017/1/20 END
      '                  If intDwgPdf = 0 Then
      '                     MsgBox "一定要有Dwg.PDF附件！", vbExclamation
      '                     SSTab1.Tab = 0
      '                     Exit Function
      '                  End If
      '               Else
                     If bolP020NewCase = False Then
                  '2016/12/14 END
                        '有繪圖
                        If m_DPMan <> "" Then
                           If intPDF < 2 Then
                              'MsgBox "至少要有二個.pdf附件！", vbExclamation
                              MsgBox "至少要有二個（文+圖）PDF附件！", vbExclamation
                              SSTab1.Tab = 0
                              If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/3
                              Exit Function
                           End If
                        '無繪圖
                        Else
                           If intPDF < 1 Then
                              MsgBox "至少要有一個（文）PDF附件！", vbExclamation
                              SSTab1.Tab = 0
                              If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/3
                              Exit Function
                           End If
                        End If
                     End If
                  End If
                  '2014/9/5 END
                  
                  'Add By Sindy 2018/9/26 CFP案的C類一律要有客戶函
                  If PField(1) = "CFP" And Left(lblCP09, 1) = "C" And intCUSPDF = 0 Then
                     MsgBox "沒有.CUS.PDF客戶函,不可送件！" & vbCrLf & vbCrLf & "(客戶函請修正檔名為.CUS.PDF)", vbExclamation
                     SSTab1.Tab = 0
                     If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/1
                     Exit Function
                  End If
               End If
               
               'Add By Sindy 2014/10/7 附加流程不需dwg.pdf檔
               If Not (Left(CboEEP04.Text, 2) = EMP_附加流程 Or bolBCaseFlow = True) Then
               '2014/10/7 END
                  '有繪圖人員並且不是電子送件時,則必須有DWG.PDF
                  If m_DPMan <> "" And cp(118) = "" And intDwgPdf < 1 Then
                     'Modify By Sindy 2016/5/6 品薇承辦的大陸案,由於目前流程有改變,故請取消繪圖沒有判發的控制
                     'Modify By Sindy 2018/10/4 98012改判斷是P12專利處程序
                     'If Not (Left(m_EPMan, 5) = "98012" And m_Country = "020") Then
                     If Not (PUB_GetST03(Left(m_EPMan, 5)) = "P12" And m_Country = "020") Then
                     '2016/5/6 END
                        MsgBox "必須有DWG.PDF附件！", vbExclamation
                        SSTab1.Tab = 0
                        If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/1
                        Exit Function
                     End If
                  End If
               End If
            End If
         End If 'Add By Sindy 2019/12/5 +
      End If
      
      '墨完或繪圖判發並且為自行核判時,必須有dwg.pdf檔案
      If Left(CboEEP04.Text, 2) = EMP_墨完 Or _
         (m_DPMan = m_DMMan And Left(CboEEP04.Text, 2) = EMP_繪圖判發) Or _
         (lstAtt(0).ListCount <> 0 And Left(CboEEP04.Text, 2) = EMP_繪圖判發) Then
'         intPDF = 0: intDOC = 0: intDwgPdf = 0: intDwg = 0
'         For ii = 0 To lstAtt(0).ListCount - 1
'            If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 7) = "DWG.PDF" Then
'               intDwgPdf = intDwgPdf + 1
'            End If
'            'Modify By Sindy 2014/3/11 +dwg.7z
'            If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 4) = ".DWG" Or _
'               Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 7) = "DWG.ZIP" Or _
'               Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 6) = "DWG.7Z" Then
'               intDwg = intDwg + 1
'            End If
'         Next ii
         If IsExistsDWG = False Then '無關聯案
            If lstAtt(0).ListCount < 2 Then
               'Add By Sindy 2013/10/25 933.覆函
               'If (CP(10) = 修正 Or CP(10) = 933) And intDwg = 0 Then
               If intDwg = 0 Then 'Modify By Sindy 2013/11/7
                  'Modify By Sindy 2014/3/11 +dwg.7z
                  If MsgBox("確定無.DWG或DWG.ZIP或DWG.7Z附件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     Exit Function
                  End If
               Else
               '2013/10/25 END
                  'Modify By Sindy 2013/11/7
                  'MsgBox "至少要有二個附件！", vbExclamation
                  MsgBox "必須有DWG.PDF附件！", vbExclamation
                  '2013/11/7 END
                  SSTab1.Tab = 0
                  If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/1
                  Exit Function
               End If
            'Add By Sindy 2013/11/7
            Else
               If intDwg = 0 Then
                  'Modify By Sindy 2014/3/11 +dwg.7z
                  If MsgBox("確定無.DWG或DWG.ZIP或DWG.7Z附件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     Exit Function
                  End If
               End If
            '2013/11/7 END
            End If
         'Add By Sindy 2013/11/7
         Else
            If intDwg > 0 Then
               intQ = MsgBox("關聯案已存放過繪圖原始檔，無需再行存放，確定還要重覆存放嗎？", vbExclamation + vbYesNoCancel + vbDefaultButton3, "重要訊息！")
               If intQ = vbYes Then
                  '要存檔
               ElseIf intQ = vbNo Then
                  '刪繪圖原始檔不存
                  'Modify By Sindy 2014/3/11 +dwg.7z
                  For ii = lstAtt(0).ListCount - 1 To 0 Step -1
                     If Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 4) = ".DWG" Or _
                        Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 7) = "DWG.ZIP" Or _
                        Right(Trim(UCase(GetFileName(lstAtt(0).List(ii)))), 6) = "DWG.7Z" Then
                        lstAtt(0).RemoveItem ii
                     End If
                  Next ii
               Else
                  Exit Function
               End If
            End If
         '2013/11/7 END
         End If
         
         If intDwgPdf < 1 Then
            MsgBox "必須有DWG.PDF附件！", vbExclamation
            SSTab1.Tab = 0
            If lstAtt(0).ListCount > 0 Then lstAtt(0).Selected(0) = True 'Add By Sindy 2018/10/1
            Exit Function
         End If
      End If
   End If
   
   'Add By Sindy 2013/10/25
   If ((Left(CboEEP04.Text, 2) = EMP_判發 And (bolPAFlow = True Or bolOtherFlow = True)) Or _
       Left(CboEEP04.Text, 2) = EMP_送件 Or _
       Left(CboEEP04.Text, 2) = EMP_退件重送) And _
      lstAtt(0).ListCount = 0 Then
      'Add By Sindy 2023/10/30
      'Modify By Sindy 2024/8/13
      'If bolFCPFlow = False Then
      If bolPAFlow = True Or bolTMFlow = True Or bolOtherFlow = True Then
      '2023/10/30 END
         MsgBox "送件至程序發文，不可無附件！", vbExclamation
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   
   If ChkEMail.Value = 1 And lstAtt(0).ListCount = 0 Then
      MsgBox "附件區無檔案，不可勾選E-Mail夾帶附件！"
      ChkEMail.Value = 0
   End If
   
   If CboEEP04.Enabled = True Then
      Cancel = False
      Call CboEEP04_Validate(Cancel)
      If Cancel = True Then
         CboEEP04.SetFocus
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   
   If CboEEP05.Enabled = True Then
      Cancel = False
      Call CboEEP05_Validate(Cancel)
      If Cancel = True Then
         CboEEP05.SetFocus
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/12/17
   If txtEEP08.Enabled = True Then
      Cancel = False
      Call txtEEP08_Validate(Cancel)
      If Cancel = True Then
         txtEEP08.SetFocus
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   
   '副本收受者
   'Add By Sindy 2022/4/19
   If txtEEP10_2.Visible = False Then
      txtEEP10.Text = ""
   '2022/4/19 END
   ElseIf txtEEP10_2.Enabled = True Then
      Cancel = False
      Call txtEEP10_2_Validate(Cancel)
      If Cancel = True Then
         txtEEP10_2.SetFocus
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/9/13 逐筆檢查檔案
   For ii = 0 To lstAtt(0).ListCount - 1
      strPathFile = Trim(lstAtt(0).List(ii))
      If InStrRev(strPathFile, " (") > 0 Then
         'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
         If UCase(Mid(strPathFile, InStrRev(strPathFile, " (") + 1, Len("(X86)"))) <> "(X86)" Then
         '2021/8/6 END
            strPathFile = Trim(Left(strPathFile, InStrRev(strPathFile, " (") - 1))
         End If
      End If
      '檔案是否正在使用中
      If PUB_ChkFileOpening(strPathFile, bolHadShowMsg) = True Then
         'Modify By Sindy 2018/9/26
         If bolHadShowMsg = False Then
         '2018/9/26 END
            MsgBox strPathFile & vbCrLf & "檔案正在使用中，請關閉才可執行送出！", vbExclamation
         End If
         SSTab1.Tab = 0
         Screen.MousePointer = vbDefault 'Add By Sindy 2019/1/21
         Exit Function
      End If
      '檔案名稱是否符合規定
      If UCase(cmdSend.Caption) <> UCase("E-Mail") Then 'Add By Sindy 2018/9/13 + if
         'Modify By Sindy 2018/9/25 商標處開放下列檔名可不加本所案號,存檔時系統自動補填
         'Add By Sindy 2021/10/14 ACS案件,不限制電子檔要輸入案號
         bolNotChkFileCaseNo = False
         If NotChkFileCaseNo(GetFileName(strPathFile), 0) = True Or _
            PField(1) = "ACS" Then
            bolNotChkFileCaseNo = True
         End If
         '2018/9/25 END
         'Add By Sindy 2018/10/29 多案，檔名檢查
         'Modify By Sindy 2020/9/29 + And txtLpNote.Tag <> "多案單筆歷程"
         'Modify By Sindy 2023/6/21 txtLpNote.Tag <> "多案單筆歷程" => txtLpNote.Tag <> ""
         'Modify By Sindy 2023/6/26 取消 And (txtLpNote.Tag <> "" Or bolManyCaseToMix = True)
         'Modify By Sindy 2023/11/9 + Or txtLpNote.Tag = "多案單筆歷程"
         If (cmdManyCase.Visible = True And cmdManyCase.Enabled = True) Or txtLpNote.Tag = "多案單筆歷程" Then
            If ManyCaseChkFileName(GetFileName(strPathFile), , , bolNotChkFileCaseNo) = False Then
               SSTab1.Tab = 0
               Exit Function
            End If
         Else
         '2018/10/29 END
            'Modify By Sindy 2023/12/15 +, bolFCPFlow
            'Modify By Sindy 2024/8/13 + 外商FC同外專不鎖中文,因附件區不進卷宗區
            If PUB_ChkEmpFlowFNMRule(lblCaseNo, GetFileName(strPathFile), Left(CboEEP04, 2), cp(10), , , , , , , _
                  bolNotChkFileCaseNo, _
                  IIf(bolFCPFlow = True Or bolFCTFlow = True, True, False)) = False Then
               SSTab1.Tab = 0
               Exit Function
            End If
         End If
      End If
   Next ii
   
   'Add By Sindy 2013/9/5
   If Left(CboEEP04.Text, 2) = EMP_會完 Then
      If MsgBox("是否確定【會稿完成】可直接送件，不再修改？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Function
      End If
      'Add By Sindy 2018/9/20 是否通知客戶
      If ChkEP11.Visible = True And ChkEP11.Value = 0 Then
         If MsgBox("是否確定通知客戶及發文？(是.通知 否.不通知)", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
            'Add By Sindy 2020/12/11
            If Trim(txtEEP08) = "" Then
               MsgBox "請加註原因後，再送出！"
               txtEEP08.SetFocus
               Exit Function
            Else
            '2020/12/11 END
               ChkEP11.Value = 1
            End If
         End If
      End If
      '2018/9/20 END
   End If
   
   'Add By Sindy 2013/9/23
   If Left(CboEEP04.Text, 2) = EMP_附加流程 Then
      'Modify By Sindy 2018/11/1 + IIf(CboEEP04.Tag <> "" And bolTMFlow = True, "﹝" & CboEEP04.Tag & "﹞流程嗎", "")
      If MsgBox("確定是否要執行【附加流程－" & Trim(Mid(CboCP10.Text, 5)) & "】" & IIf(CboEEP04.Tag <> "" And bolTMFlow = True, "﹝" & CboEEP04.Tag & "﹞流程嗎", "") & "？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Function
      Else
         'Modify By Sindy 2025/7/31 mark,為多餘的程式段
'         'Add By Sindy 2019/6/12
'         'Add By Sindy 2025/7/31 +Or bolFCPFlow = True
'         If bolPAFlow = True Or bolFCPFlow = True Then
'            '因這些案件性質是回到「工作進度資料維護」中，點選此筆新增案件，
'            '工程師自行進行所需的簽辦流程，因此收受者不須帶入人員，帶入反而有可能會出現錯誤
'            If Trim(Left(CboCP10.Text, 4)) = "936" Or _
'               Trim(Left(CboCP10.Text, 4)) = "957" Or _
'               Trim(Left(CboCP10.Text, 4)) = "958" Then
'               CboEEP05.Enabled = False 'Add By Sindy 2025/7/31
'               CboEEP05.Text = ""
'               m_EEP11Person = ""
'               CboEEP04.Tag = ""
'            End If
'         End If
'         '2019/6/12 END
      End If
   End If
   
   'Add By Sindy 2013/11/13 當聯絡流程時,收受者為送英核主管,顯示訊息詢問是否為送英核,若是,在系統備註裡加註"流程狀態:05"
   If Left(CboEEP04.Text, 2) = EMP_聯絡 And Trim(Left(CboEEP05.Text, 6)) = Left(Trim(m_EMMan), 5) And _
      bolWaitReply = True And _
      Val(m_EP33) = 0 Then
      If MsgBox("是否送英文核稿人做核稿動作？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
         If lstAtt(0).ListCount = 0 Then
            MsgBox "請放入附件！", vbExclamation
            SSTab1.Tab = 0
            Exit Function
         End If
         m_UpdEEP11 = "流程狀態:" & EMP_送英核
         'Modify By Sindy 2015/3/16
         If m_EP41 = "2" Then '2.日
            txtEEP08.Text = "[送日核]" & txtEEP08.Text
         Else
         '2015/3/16 END
            txtEEP08.Text = "[送英核]" & txtEEP08.Text 'Modify By Sindy 2013/11/14
         End If
      End If
   End If
   '2013/11/13 END
   
   'Add By Sindy 2014/7/22 發生P-109071(AA3029564)案,判發主管做判發了,工程師又做了聯絡,導致程序人員無法送件發文
   'Add By Sindy 2015/12/18
   If rsA.State <> adStateClosed Then rsA.Close
   strExc(0) = "select cp27,cp57" & _
               " From caseprogress" & _
               " where cp09='" & m_EEP01 & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If Val("" & rsA.Fields("CP27")) > 0 Or Val("" & rsA.Fields("CP57")) > 0 Then
         MsgBox "此文" & IIf(Val("" & rsA.Fields("CP27")) > 0, "已發文", IIf(Val("" & rsA.Fields("CP57")) > 0, "已取消收文", "")) & "，不可再執行該歷程！"
         Set rsA = Nothing
         Exit Function
      End If
   'Add By Sindy 2016/3/10 送出時檢查文號是否存在
   '                       BA4033824, BA5003837已刪除但歷程有資料，檢查發現是時間點問題，作業沒reload
   Else
      MsgBox "此文號已不存在，不可再執行此作業！"
      Set rsA = Nothing
      Call cmdCancel_Click
      cmdExit.Enabled = False
      Exit Function
      '2016/3/10 END
   End If
   '2015/12/18 END
   
   'Add By Sindy 2014/9/10
   If bolPAFlow = True Then
      If Not (PField(1) = "P" And Left(cp(12), 1) = "F") Then 'Add By Sindy 2014/9/16 +if 雅娟:非FMP才需要檢查無圖式
         If InStr(NewCasePtyList, cp(10)) > 0 Then '新申請案
            If Left(CboEEP04.Text, 2) = EMP_送判 Or _
               Left(CboEEP04.Text, 2) = EMP_退件重送 Or _
               ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發) Then
               If m_DPMan = "" And cp(118) = "" And intDwgPdf < 1 Then
                  'Modify By Sindy 2014/10/24
   '               If MsgBox("請確認是否為無圖式？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
   '                  Exit Function
   '               End If
                  If MsgBox("請確認是否無圖式？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     If MsgBox("圖在文內？" & vbCrLf & _
                               "（按「否」：請插入圖檔(Dwg.PDF)）", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                        Exit Function
                     End If
                  Else
                     If InStr(txtEEP08, "【注意:無圖式】") = 0 Then
                        txtEEP08 = txtEEP08 & "【注意:無圖式】"
                     End If
                  End If
                  '2014/10/24 END
               End If
            End If
         End If
      End If
   End If
   '2014/9/10 END
   
   'Add By Sindy 2020/1/14 P台灣案主動修正對其他對應案之提醒
   'P台灣案主動修正時,檢查相關案無主動修正未發文時,就要彈提醒訊息
   If PField(1) = "P" And cp(10) = "203" And m_Country = "000" And UCase(m_PrevForm.Name) = UCase("frm090201_2") Then
      '相關案要排除FCP
      strExc(0) = "select cp09,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,cp14,ep09,ep07,ep08,ep38,cp13,cp01,cp02,cp03,cp04" & _
         " from (select cm01,cm02,cm03,cm04 from casemap where cm10='0' and cm05='" & PField(1) & "' and cm06='" & PField(2) & "' and cm07='" & PField(3) & "' and cm08='" & PField(4) & "'" & _
         " union select cm05,cm06,cm07,cm08 from casemap where cm10='0' and cm01='" & PField(1) & "' and cm02='" & PField(2) & "' and cm03='" & PField(3) & "' and cm04='" & PField(4) & "'" & _
         " union select cr01,cr02,cr03,cr04 from caserelation where cr05='" & PField(1) & "' and cr06='" & PField(2) & "' and cr07='" & PField(3) & "' and cr08='" & PField(4) & "'" & _
         "),caseprogress,engineerprogress where cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and instr('" & NewCasePtyList & "',cp10)>0" & _
         " and ep02(+)=cp09 and ep06>0 and cp01<>'FCP'" & _
         " and not exists (select * from caseprogress where cp01=cm01 and cp02=cm02 and cp03=cm03 and cp04=cm04 and cp10='203' and cp158=0 and cp159=0)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            MsgBox "本案有其他國家，請確認是否要收文主動修正。", vbExclamation
         End If
      End If
   End If
   '2020/1/14 END
   
   'Add By Sindy 2015/12/18
   'Modify By Sindy 2016/3/10 Mark
'   If rsA.State <> adStateClosed Then rsA.Close
'   strExc(0) = "select EEP01,EEP04,ac03" & _
'               " From empelectronprocess,allcode" & _
'               " where eep01='" & m_EEP01 & "'" & _
'               " and EEP09='Y'" & _
'               " and ac01='09' And eep04=ac02(+)"
'   rsA.CursorLocation = adUseClient
'   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      If InStr(EMP_需等待回覆的狀態, Left(CboEEP04.Text, 2)) > 0 Then
'         MsgBox "此文已" & rsA.Fields("ac03") & "中，不可再執行該歷程！"
'         Set rsA = Nothing
'         Exit Function
'      End If
'   End If
   
   'Add By Sindy 2017/8/14 王副總提出歷程判發(中)後還是可以開放聯絡
'   If rsA.State <> adStateClosed Then rsA.Close
'   strExc(0) = "select EEP01,EEP04" & _
'               " From empelectronprocess" & _
'               " where eep01='" & m_EEP01 & "' and eep04<>'" & EMP_聯絡 & "' order by EEP02 desc"
'   rsA.CursorLocation = adUseClient
'   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      rsA.MoveFirst
'      If rsA.Fields("EEP04") = EMP_判發 Or _
'         rsA.Fields("EEP04") = EMP_退件重送 Then
'         MsgBox "此文已判發，不可再執行該歷程！"
'         Set rsA = Nothing
'         Exit Function
'      End If
'   End If
   
   Set rsA = Nothing
   TxtValidate = True
End Function

'Add By Sindy 2014/12/16
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   GRD1.ToolTipText = ""
   If GRD1.MouseRow <> 0 And GRD1.MouseCol > 0 Then
      'Modify By Sindy 2018/10/8
      If GRD1.MouseCol = 14 Then
         GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, 16) '系統記錄
      '2018/10/8 END
      ElseIf GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
         GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
      End If
   End If
End Sub

'\\typing2\外專送件
Private Sub Label22_Click(Index As Integer)
   
On Error GoTo ErrHnd 'Add by Sindy 2024/8/8

   strExc(10) = GetFCPPathVal(Label22(Index).Caption, PField(1), PField(2), Trim(lblCP10.Caption))
   If Dir(strExc(10), vbDirectory) <> "" Then
      ShellExecute hLocalFile, "explore", strExc(10), vbNullString, vbNullString, 1
   Else
      MsgBox "無此資料夾! " & strExc(10), vbExclamation
   End If
   'Label22(Index).Tag = strExc(10)
'Add by Sindy 2024/8/8
   Exit Sub
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Number & ": " & Err.Description, , Label22(Index).Caption
   End If
'2024/8/8 END
End Sub
'Add by Sindy 2025/10/20
Private Sub Label22_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Screen.MousePointer = vbHourglass
   DoEvents
   strExc(10) = GetFCPPathVal(Label22(Index).Caption, PField(1), PField(2), Trim(lblCP10.Caption))
   Label22(Index).ToolTipText = strExc(10)
   Screen.MousePointer = vbDefault
End Sub
'2025/10/20 END
'\\typing2\外專送件\中說原始檔
Private Sub Label23_Click()
   
On Error GoTo ErrHnd 'Add by Sindy 2024/8/8
   
   strExc(10) = GetFCPPathVal(Label23.Caption, PField(1), PField(2), Trim(lblCP10.Caption))
   If Dir(strExc(10), vbDirectory) <> "" Then
      ShellExecute hLocalFile, "explore", strExc(10), vbNullString, vbNullString, 1
   Else
      MsgBox "無此資料夾! " & strExc(10), vbExclamation
   End If
   'Label23.Tag = strExc(10)
'Add by Sindy 2024/8/8
   Exit Sub
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Number & ": " & Err.Description, , Label23.Caption
   End If
'2024/8/8 END
End Sub
'Add by Sindy 2025/10/20
Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Screen.MousePointer = vbHourglass
   DoEvents
   strExc(10) = GetFCPPathVal(Label23.Caption, PField(1), PField(2), Trim(lblCP10.Caption))
   Label23.ToolTipText = strExc(10)
   Screen.MousePointer = vbDefault
End Sub
'2025/10/20 END
'\\typing2\電子送件暫存區\+案號(流水號必須足6碼)
Private Sub Label25_Click(Index As Integer)
   Call SetLabel25Folder(Index)
End Sub
Private Function SetLabel25Folder(Index As Integer, Optional ByVal bolMkDir As Boolean = False _
   , Optional ByVal bolOpen As Boolean = True) As String

On Error GoTo ErrHnd 'Add by Sindy 2024/8/8
   
   SetLabel25Folder = GetFCPPathVal(Label25(Index).Caption, PField(1), PField(2), Trim(lblCP10.Caption), True)
   If Dir(SetLabel25Folder, vbDirectory) <> "" Then
      If bolOpen = True Then
         ShellExecute hLocalFile, "explore", SetLabel25Folder, vbNullString, vbNullString, 1
      End If
   Else
      If CmdF21(0).Visible = True Then
         If bolMkDir = True Then
            MkDir SetLabel25Folder
         Else
            If MsgBox("需要建立此 " & vbCrLf & SetLabel25Folder & vbCrLf & " 子資料夾嗎？", vbYesNo + vbDefaultButton1) = vbYes Then
               MkDir SetLabel25Folder
               ShellExecute hLocalFile, "explore", SetLabel25Folder, vbNullString, vbNullString, 1
            End If
         End If
      Else
         MsgBox "無此資料夾! " & SetLabel25Folder, vbExclamation
      End If
   End If
   'Label25(Index).Tag = SetLabel25Folder
'Add by Sindy 2024/8/8
   Exit Function
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Number & ": " & Err.Description, , Label25(Index).Caption
   End If
'2024/8/8 END
End Function
'Add by Sindy 2025/10/20
Private Sub Label25_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Screen.MousePointer = vbHourglass
   DoEvents
   strExc(10) = GetFCPPathVal(Label25(Index).Caption, PField(1), PField(2), Trim(lblCP10.Caption), True)
   Label25(Index).ToolTipText = strExc(10)
   Me.Label25(Index) = "電子送件暫存區"
   Screen.MousePointer = vbDefault
End Sub
'2025/10/20 END

'Add By Sindy 2025/10/8 點二下可以開啟附件檔案
Private Sub lstAtt_DblClick(Index As Integer)
   Call cmdOpenAtt_Click(Index)
End Sub

'Add By Sindy 2014/10/1
Private Sub Text2_Click()
   SSTab1.Tab = intTab_存卷資料
End Sub

'Add By Sindy 2022/7/27
Private Sub txt1_Change(Index As Integer)
   PUB_RefreshText txt1(Index)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If Index = 7 Then
      If Len(txt1(7)) <> 0 Then
         txt1(7) = Left(txt1(7) & "0000000", 9) 'Add By Sindy 2022/1/13
         lblFa.Caption = PUB_GetFAgentName(txt1(7))
         If Trim(lblFa.Caption) <> "" Then
            txt1(5).Text = lblFa.Caption
         End If
      End If
   End If
End Sub

'條款
Private Sub txt2_GotFocus()
   InverseTextBox txt2
End Sub
Private Sub txt2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txt2_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim tmpCp49Arr As Variant
   Dim intCp49 As Integer
   Dim s As Integer
   
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(txt2) = True Then
      GoTo EXITSUB
   End If
   
   tmpCp49Arr = Split(txt2, ",")
   For intCp49 = 0 To UBound(tmpCp49Arr)
      '條款不小於3碼
      'C類來函不檢查
      If Len(Trim(tmpCp49Arr(intCp49))) < 3 And Len(Trim(tmpCp49Arr(intCp49))) <> 0 And lblCP09 < "C" Then
         s = MsgBox(tmpCp49Arr(intCp49) & "，小於 3 碼！", , "條款輸入錯誤！")
         Cancel = True
         txt2.SetFocus
         GoTo EXITSUB
      End If
      If Len(Trim(tmpCp49Arr(intCp49))) <> 0 Then
         '只抓前三碼或前四碼檢查
         strSql = "select * from law where lw01='" & Left(tmpCp49Arr(intCp49), 3) & "' or lw01='" & Left(tmpCp49Arr(intCp49), 4) & "' "
         Set rsTmp = New ADODB.Recordset
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount = 0 Then
             s = MsgBox("沒有 " & tmpCp49Arr(intCp49) & " 條款！", , "條款輸入錯誤！")
             Cancel = True
             txt2.SetFocus
             GoTo EXITSUB
         End If
         rsTmp.Close
      End If
   Next intCp49
   
EXITSUB:
   Set rsTmp = Nothing
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
               'Modify By Sindy 2024/6/6
               If Index = 5 Then '管制人
                  If cp(158) = 0 And cp(159) = 0 Then '未發文時重新帶最新管制人
                     strExc(10) = Left(PUB_GetFCPHandler(PField(1), PField(2), PField(3), PField(4)), 5)
                     If strExc(10) <> Trim(txt3(Index).Text) Then
                        txt3(Index).Text = strExc(10)
                        strText = GetPrjSalesNM(CStr(Trim(txt3(Index).Text)))
                     End If
                  End If
               Else
               '2024/6/6 END
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

'Add By Sindy 2013/10/1
Private Sub txtCaseName_GotFocus(Index As Integer)
   Select Case Index
      Case 0, 2
         OpenIme
      Case Else
         CloseIme
   End Select
   TextInverse txtCaseName(Index)
End Sub

'Add By Sindy 2013/10/1
Private Sub txtCaseName_LostFocus(Index As Integer)
   If (txtCaseName(0).Enabled = True And txtCaseName(0).Text <> txtCaseName(0).Tag) Or _
      (txtCaseName(1).Enabled = True And txtCaseName(1).Text <> txtCaseName(1).Tag) Or _
      (txtCaseName(2).Enabled = True And txtCaseName(2).Text <> txtCaseName(2).Tag) Then
      Call GetPaperMain
   End If
End Sub

'Add By Sindy 2025/4/10
'追蹤客戶指示【約定期限】
Private Sub txtEED14_GotFocus()
   txtEED14.SelStart = 0
   txtEED14.SelLength = Len(txtEED14)
End Sub
Private Sub txtEED14_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
Private Sub txtEED14_Validate(Cancel As Boolean)
   If Len(txtEED14) <> 0 Then
      'Add By Sindy 2025/11/7
      If ChkDate(txtEED14) = False Then
         Cancel = True
         Exit Sub
      '2025/11/7 END
      ElseIf Not ChkWorkDay(ChangeTStringToWString(txtEED14)) Then
         ShowDateErr
         Cancel = True
         Exit Sub
      ElseIf DBDATE(txtEED14) < DBDATE(strSrvDate(2)) Then
         'Modify By Sindy 2025/8/20
         If Me.Frame945.Tag = 告知代理人 Then
            MsgBox "管制日期需大於系統日!"
         Else
         '2025/8/20 END
            MsgBox "約定期限需大於系統日!"
         End If
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
'委員指定送件日期【本所期限】
Private Sub txtEED15_GotFocus()
   txtEED15.SelStart = 0
   txtEED15.SelLength = Len(txtEED15)
End Sub
Private Sub txtEED15_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
Private Sub txtEED15_Validate(Cancel As Boolean)
   If Len(txtEED15) <> 0 Then
      'Add By Sindy 2025/11/7
      If ChkDate(txtEED15) = False Then
         Cancel = True
         Exit Sub
      '2025/11/7 END
      ElseIf Not ChkWorkDay(ChangeTStringToWString(txtEED15)) Then
         ShowDateErr
         Cancel = True
         Exit Sub
      ElseIf DBDATE(txtEED15) < DBDATE(strSrvDate(2)) Then
         MsgBox "本所期限需大於系統日!"
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
'2025/4/10 END

'Add By Sindy 2022/7/27
Private Sub txtEEP08_Change()
   PUB_RefreshText txtEEP08
End Sub

'Add By Sindy 2022/7/15
Private Sub txtEEP08_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtEEP08
End Sub

'Add By Sindy 2013/12/17
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

Private Sub txtEEP10_2_GotFocus()
   InverseTextBox txtEEP10_2
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
      If Trim(strEEP10_Err) <> "" Then
         MsgBox "副本收受者資料有誤！(" & strEEP10_Err & ")"
         If txtEEP10_2.Visible = True Then txtEEP10_2.SetFocus
         Call txtEEP10_2_GotFocus
         'Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Modify By Sindy 2025/4/7
'Private Sub txtEEP10_2_Validate(Cancel As Boolean)
Public Sub txtEEP10_2_Validate(Cancel As Boolean)
'2025/4/7 END
Dim strMsgText As String

   If txtEEP10_2 <> "" Or txtEEP10 <> "" Then
      Call txtEEP10_2_LostFocus
      If strEEP10_Err <> "" Then
         Cancel = True
         Exit Sub
      End If
      If txtEEP10_2 <> "" Then
         If Left(txtEEP10_2, 5) = txtEEP03 Then
            MsgBox "不可為本人！", vbExclamation
            CboEEP05.SetFocus
            Call CboEEP05_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtEEP08_GotFocus()
   InverseTextBox txtEEP08
End Sub

Private Sub CboEEP04_GotFocus()
   InverseTextBox CboEEP04
End Sub

Private Sub SetCboEEP05()
Dim strTemp As String
Dim j As Integer
   
   CboEEP05.Clear
   
   'Add By Sindy 2024/10/29 CFP-34302智權人員為北一備用(m_SPMan會為空白)
   '  以防工程師不知給誰,而下拉式選單誤擇一人做操作
   '  固不產出下拉式選單 (智權人員為虛編號的狀況,在2019年開始已是由人員自行輸入了。參2019/7/9 員工編號小於6視為沒有抓到資料,人員自行輸入)
   If Left(CboEEP04.Text, 2) = EMP_送會 And m_SPMan = "" Then
      Exit Sub
   End If
   '2024/10/29 END
   
   'Modified by Lydia 2017/06/15 已離職員工在下拉清單不顯示
'   If m_DPMan <> "" And Left(m_DPMan, 5) <> m_FlowUserNum Then CboEEP05.AddItem m_DPMan
'   If m_EPMan <> "" And Left(m_EPMan, 5) <> m_FlowUserNum Then CboEEP05.AddItem m_EPMan
'   'If m_SPMan <> "" And Trim(Left(m_SPMan, 6)) <> m_FlowUserNum Then CboEEP05.AddItem m_SPMan
'   If m_SPMan <> "" And Trim(Left(m_SPMan, 6)) <> m_FlowUserNum And Trim(Left(m_SPMan, 6)) <> Left(m_CMMan, 5) Then
'      CboEEP05.AddItem m_SPMan
'   End If
'   If m_EMMan <> "" And Left(m_EMMan, 5) <> m_FlowUserNum Then CboEEP05.AddItem m_EMMan
'   If m_DCMan <> "" And Left(m_DCMan, 5) <> m_FlowUserNum Then CboEEP05.AddItem m_DCMan 'Add By Sindy 2015/4/22
'   If m_DMMan <> "" And Left(m_DMMan, 5) <> m_FlowUserNum Then CboEEP05.AddItem m_DMMan
'   If m_CMMan <> "" And Left(m_CMMan, 5) <> m_FlowUserNum Then CboEEP05.AddItem m_CMMan
'   If m_CSMan <> "" And Left(m_CSMan, 5) <> m_FlowUserNum And Left(m_CSMan, 5) <> Left(m_CMMan, 5) Then
'      CboEEP05.AddItem m_CSMan
'   End If
   
   'Modify By Sindy 2019/7/9
'   If m_DPMan <> "" And Left(m_DPMan, 5) <> m_FlowUserNum And ChkStaffST04(Trim(Left(m_DPMan, 6)), False) = False Then
'      CboEEP05.AddItem m_DPMan
'   End If
'   If m_EPMan <> "" And Left(m_EPMan, 5) <> m_FlowUserNum And ChkStaffST04(Trim(Left(m_EPMan, 6)), False) = False Then
'      CboEEP05.AddItem m_EPMan
'   End If
'   If m_SPMan <> "" And Trim(Left(m_SPMan, 6)) <> m_FlowUserNum And Trim(Left(m_SPMan, 6)) <> Left(m_CMMan, 5) And ChkStaffST04(Trim(Left(m_SPMan, 6)), False) = False Then
'      CboEEP05.AddItem m_SPMan
'   End If
'   If m_EMMan <> "" And Left(m_EMMan, 5) <> m_FlowUserNum And ChkStaffST04(Trim(Left(m_EMMan, 6)), False) = False Then
'      CboEEP05.AddItem m_EMMan
'   End If
'   If m_DCMan <> "" And Left(m_DCMan, 5) <> m_FlowUserNum And ChkStaffST04(Trim(Left(m_DCMan, 6)), False) = False Then
'      CboEEP05.AddItem m_DCMan
'   End If
'   If m_DMMan <> "" And Left(m_DMMan, 5) <> m_FlowUserNum And ChkStaffST04(Trim(Left(m_DMMan, 6)), False) = False Then
'      CboEEP05.AddItem m_DMMan
'   End If
'   If m_CMMan <> "" And Left(m_CMMan, 5) <> m_FlowUserNum And ChkStaffST04(Trim(Left(m_CMMan, 6)), False) = False Then
'      CboEEP05.AddItem m_CMMan
'   End If
'   If m_CSMan <> "" And Left(m_CSMan, 5) <> m_FlowUserNum And Left(m_CSMan, 5) <> Left(m_CMMan, 5) And ChkStaffST04(Trim(Left(m_CSMan, 6)), False) = False Then
'      CboEEP05.AddItem m_CSMan
'   End If
   For j = 1 To 9 '8
      If j = 1 Then strTemp = m_DPMan
      If j = 2 Then strTemp = m_EPMan
      If j = 3 Then strTemp = m_SPMan
      If j = 4 Then strTemp = m_EMMan
      If j = 5 Then strTemp = m_DCMan
      If j = 6 Then strTemp = m_DMMan
      If j = 7 Then strTemp = m_CMMan
      If j = 8 Then strTemp = m_CSMan
      If j = 9 Then strTemp = m_F21CMMan 'Add By Sindy 2023/10/2
      If strTemp <> "" And Left(strTemp, 5) <> m_FlowUserNum And ChkStaffST04(Trim(Left(strTemp, 6)), False) = False Then
         Call ChkCboEEP05AddItem(strTemp)
'         For i = 0 To CboEEP05.ListCount - 1
'            bolExits = False
'            If CboEEP05.List(i) = strTemp Then
'               bolExits = True
'               Exit For
'            End If
'         Next i
'         If bolExits = False Then
'            CboEEP05.AddItem strTemp
'         End If
      End If
   Next j
   '2019/7/9 END
   
   'Add By Sindy 2023/10/2
   If bolFCPFlow = True Then
      Call ChkCboEEP05AddItem(Pub_GetSpecMan("C") & " " & GetPrjSalesNM(Pub_GetSpecMan("C"))) '分案人員
   End If
   '2023/10/2 END
End Sub

'Add By Sindy 2019/7/18
Private Sub ChkCboEEP05AddItem(strFullEmp As String)
Dim i As Integer
Dim bolExits As Boolean
   
   For i = 0 To CboEEP05.ListCount - 1
      bolExits = False
      If CboEEP05.List(i) = strFullEmp Then
         bolExits = True
         Exit For
      End If
   Next i
   If bolExits = False Then
      CboEEP05.AddItem strFullEmp
   End If
End Sub

'Add By Sindy 2017/6/15 承辦人為外翻時,ST14處理人員可能設多個,一個以上的其他人員視為副本收受者
Private Sub SettxtEEP10_2()
   If m_CP14_2 <> "" And CboEEP05 <> "" Then
      txtEEP10_2 = m_CP14_2
      txtEEP10_2 = Replace(txtEEP10_2, Trim(Left(CboEEP05.Text, 6)), "")
      txtEEP10_2 = Replace(txtEEP10_2, ",,", ",")
      If Left(txtEEP10_2, 1) = "," Then txtEEP10_2 = Mid(txtEEP10, 2)
      Call txtEEP10_2_LostFocus
   End If
   
   'Add By Sindy 2021/11/10 林律師和商標處間的歷程之副本收受者均自動設定為江郁仁協理
   'Modify By Sindy 2021/12/29 + Or Trim(Left(CboEEP05.Text, 6)) = "98003")
   'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
   If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
      If (txtEEP03 = "98003" Or Trim(Left(CboEEP05.Text, 6)) = "98003") Then
         If InStr(txtEEP10_2, "98020") = 0 And InStr(txtEEP10_2, GetPrjSalesNM("98020")) = 0 Then
            If txtEEP10_2 <> "" Then txtEEP10_2 = txtEEP10_2 & ","
            txtEEP10_2 = txtEEP10_2 & "98020"
            Call txtEEP10_2_LostFocus
         End If
      ElseIf txtEEP03 <> "98003" And Trim(CboEEP05.Text) = "" And txtEEP10_2 = GetPrjSalesNM("98020") Then
         txtEEP10_2 = ""
      End If
   End If
   
   'Add By Sindy 2025/5/23 李柏翰經理指示為了確保CFP設計案的圖式的品質, 請修改繪圖人員的CFP設計的草核規則
   '   改為所有的CFP設計案都要經過草核, 因希望要有2次核稿(設定核判主管為翔龍副理),且預設副本給82018月嬌主任
   '此 strPP04 <> "" 是判斷有設定核判表, 才需檢查下列條件
   strExc(10) = "82018"
   If Left(CboEEP04.Text, 2) = EMP_草核 And strPP04 <> "" _
      And Trim(Left(CboEEP05.Text, 6)) <> strExc(10) _
      And txtEEP03 <> strExc(10) _
      And InStr(txtEEP10_2, strExc(10)) = 0 _
      And InStr(txtEEP10_2, GetPrjSalesNM(strExc(10))) = 0 Then
      'Modify By Sindy 2025/6/3
      'Modify By Sindy 2025/9/5 CFP設計的回代跟答辯，比照CFP設計的新申請案的管控方式
      'If cp(1) = "CFP" And cp(10) = 設計申請 Then
      If cp(1) = "CFP" And pa(8) = "3" Then
      '2025/9/5 END
      '2025/6/3 END
         If txtEEP10_2 <> "" Then txtEEP10_2 = txtEEP10_2 & ","
         txtEEP10_2 = txtEEP10_2 & strExc(10)
         Call txtEEP10_2_LostFocus
      End If
   End If
   '2025/5/23 END
End Sub

'Add By Sindy 2023/10/6
Private Function GetFCP924txtEEP08(strNote As String) As String
Dim strTmp As String
Dim intR As Integer
   
   intR = 0
   strExc(10) = ""
   '因, Added by Lydia 2023/10/04 FMP案待客戶最終指示相關控管：加註
   If PUB_ChkFMP970mail("2", PField(1), PField(2), PField(3), PField(4), strTmp) = True Then
      If strTmp <> "" Then
         intR = intR + 1
         strExc(10) = "【待最終指示】" & vbCrLf & _
                      intR & "." & strTmp & vbCrLf
      End If
   End If
   If strTmp <> "" Then
      intR = intR + 1
      strExc(10) = strExc(10) & intR & "."
   End If
   strExc(10) = strExc(10) & strNote & "通知工程師主管進行分案" & vbCrLf
   If strTmp <> "" And Trim(cp(64)) <> "" Then
      intR = intR + 1
      strExc(10) = strExc(10) & intR & "."
      strExc(10) = strExc(10) & "進度備註:" & Trim(cp(64)) & vbCrLf
   End If
   GetFCP924txtEEP08 = strExc(10)
End Function

'Add By Sindy 2023/11/16
Private Sub ChkCP113CP114()
Dim strCP114 As String, strCP113 As String
Dim intMaxValue As Integer
   
   '外專承辦人工作進度
   If UCase(m_PrevForm.Name) = UCase("frm090909") _
      And Not (Lbl926.Visible = True And InStr(Lbl926.Caption, "一核") > 0) Then '排除一核
      If cp(10) = "201" Then '新案翻譯
         If Trim(m_PrevForm.txt1(1).Text) = "" Then '核稿時數空白
RunInput_114:
            strCP114 = InputBox("請輸入核稿時數，不可空白! (請輸入數字)")
            If strCP114 = "" Then
               MsgBox "核稿時數不可空白！", vbExclamation
               GoTo RunInput_114
            ElseIf strCP114 <> "" Then
               If Not IsNumeric(strCP114) Then
                  MsgBox "請輸入數字！", vbExclamation
                  GoTo RunInput_114
               End If
               m_PrevForm.txt1(1).Text = strCP114
            End If
         End If
      End If
      If Trim(m_PrevForm.txt1(0).Text) = "" Then '工作時數空白
         '工作時數的檢查
         If PUB_CheckCP113(m_PrevForm.txt1(0).Text, cp(1), cp(10), cp(14), False, intMaxValue) = False Then
RunInput_113:
            strCP113 = InputBox("請輸入工作時數，不可空白! (請輸入數字)", , strCP113)
            If strCP113 = "" Then
               MsgBox "工作時數不可空白！", vbExclamation
               GoTo RunInput_113
            ElseIf strCP113 <> "" Then
               If Not IsNumeric(strCP113) Then
                  MsgBox "請輸入數字！", vbExclamation
                  GoTo RunInput_113
               End If
               If intMaxValue > 0 Then
                  If Val(strCP113) > intMaxValue Then
                     If MsgBox("工作時數已超過上限值" & intMaxValue & "，是否繼續作業？", vbYesNo + vbDefaultButton1) = vbYes Then
                        m_PrevForm.txt1(0).Text = strCP113
                        Exit Sub
                     Else
                        GoTo RunInput_113
                     End If
                  End If
               End If
               m_PrevForm.txt1(0).Text = strCP113
            End If
         End If
      End If
   End If
End Sub

'Add By Sindy 2025/8/5
Private Function GetCP43AddCC(bolAddCC As Boolean) As Boolean
   GetCP43AddCC = False
   '商一組於分析程序收文後，
   '操作送會交付於智權人員會稿時，系統同時通知商三組原承辦人員。
   If Left(CboEEP04.Text, 2) = EMP_送會 And bolTMFlow = True Then
      If cp(10) = "727" And cp(43) <> "" Then
         strSql = "select cp14,st93 From caseprogress,staff" & _
                  " where cp01='" & PField(1) & "' and cp02='" & PField(2) & "' and cp03='" & PField(3) & "' and cp04='" & PField(4) & "'" & _
                  " and cp09='" & cp(43) & "' and cp14=st01(+) and cp14 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'If RsTemp.Fields("st93") = "T31" Then
            If RsTemp.Fields("cp14") <> cp(14) Then
               GetCP43AddCC = True
               If bolAddCC = True And InStr(txtEEP10_2.Text, RsTemp.Fields("cp14")) = 0 Then
                  txtEEP10_2.Text = IIf(txtEEP10_2.Text <> "", ",", "") & RsTemp.Fields("cp14")
                  Call txtEEP10_2_LostFocus
               End If
            End If
         End If
      End If
   End If
End Function

Private Sub CboEEP04_Click()
Dim strRefCaseNo As String
Dim strRefEEP02_EP As String
Dim strRefEEP02_DP As String
Dim Cancel As Boolean
Dim strReVal As String, bolCP141 As Boolean 'Add By Sindy 2024/4/29
Dim varTemp As Variant
   
   'Add By Sindy 2024/6/14 T-249005會修尚無多案機制
   If intReceiveKind = 2 Then '待會稿區
      If CboEEP04.Tag <> Left(CboEEP04.Text, 2) Then '判斷是否有改變歷程狀態,有恢復預設值
         cmdManyCase.Tag = ""
         txtLpNote.Tag = "": txtLpNote.Text = ""
         cmdManyCase.Visible = False
         cmdManyCase.Enabled = False
         lstAtt(0).Clear
      End If
   End If
   '2024/6/14 END
   ChkEP11.Visible = False 'Add By Sindy 2019/7/10 不通知客戶, 不發文
   Text1.Visible = False 'Add By Sindy 2014/10/1 備註:「聯絡」的附件，送件後一律刪除，欲留存者請置於存卷資料頁籤
   If m_EditMode = 1 Then
      cmdCaseMap.Visible = False 'Add By Sindy 2017/8/30 不顯示多國案鈕
      
      Call SetCboEEP05 '收受者
      
      'Modify By Sindy 2013/9/3 聯絡開放可以E-Mail夾帶附件,但不預設打勾
      'If Left(CboEEP04.Text, 2) = EMP_送會 Or Left(CboEEP04.Text, 2) = EMP_聯絡 Then
      'Modify By Sindy 2013/9/6 繪圖人員在做FCP案流程時,可以夾帶附件給工程師
      'Modify By Sindy 2016/5/5 +EMP_會圖也可以夾帶附件
'      If Left(CboEEP04.Text, 2) = EMP_送會
      'Modify By Sindy 2018/10/23 已增加客戶會稿功能,且商標案不複雜因此商標案不提供此功能
      If (Left(CboEEP04.Text, 2) = EMP_送會 And bolPAFlow = True And Left(PUB_GetStaffST15(Trim(Left(m_SPMan, 6)), "1"), 1) = "S") Or _
         Left(CboEEP04.Text, 2) = EMP_會圖 Or _
         (PField(1) = "FCP" And (Left(CboEEP04.Text, 2) = EMP_草完 Or Left(CboEEP04.Text, 2) = EMP_標號 Or _
                                 Left(CboEEP04.Text, 2) = EMP_繪圖判發)) Then
         ChkEMail.Visible = True
         ChkEMail.Value = 1
      Else
         ChkEMail.Visible = False
         ChkEMail.Value = 0
      End If
      
'      Label15.Visible = True
'      CboEEP05.Visible = True
      Label4.Visible = True '註：副本多人時以逗號(,)分隔
      Label5.Visible = True '副本收受者：
      txtEEP10_2.Visible = True
      CboEEP05.Enabled = True
      Frame4.Visible = False 'Add By Sindy 2013/9/23 案件性質/會稿方式
      
      '********************
      m_EEP11Person = ""
      ChkEED13.Enabled = False 'Add By Sindy 2023/11/16
      '********************
      
'      'Add By Sindy 2024/4/29 以防修改,先取消;後面符合條件會再詢問
'      If bolFCPFlow = True And InStr(txtEEP08, "【已獲客戶最終指示】") > 0 Then
'         txtEEP08 = Replace(txtEEP08, "【已獲客戶最終指示】", "")
'      End If
'      '2024/4/29 END
      '依流程狀態帶出收受者
      Select Case Left(CboEEP04.Text, 2)
         'Add By Sindy 2023/9/27
         '*********************************************************************************
         Case EMP_翻譯交稿
            CboEEP05.Text = PUB_GetFCPHandler(PField(1), PField(2), PField(3), PField(4)) '外專程序管制人
            CboEEP05.Text = CboEEP05.Text & " " & GetPrjSalesNM(CboEEP05.Text)
            txtEEP08 = "翻譯已交稿，請進行翻譯核稿流程" & vbCrLf
         Case EMP_送排版
            CboEEP05.Text = Pub_GetSpecMan("中打室排版分案人員") & " " & GetPrjSalesNM(Pub_GetSpecMan("中打室排版分案人員"))
         Case EMP_排版完成
            '分割案
            'Modify By Sindy 2024/1/5 + Or cp(10) = "209" Or cp(10) = "235"
            If (cp(10) = "307" Or cp(10) = "209" Or cp(10) = "235") _
               And m_EPMan <> "" Then
               CboEEP05.Text = m_EPMan '承辦人
               '副本同時帶入工程師主管
               txtEEP10_2.Text = Trim(Mid(m_F21CMMan, 6))
               Call txtEEP10_2_LostFocus
            Else
               CboEEP05.Text = m_F21CMMan 'FCP工程師主管
            End If
         Case EMP_送核稿分案
            CboEEP05.Text = m_F21CMMan 'FCP工程師主管
         Case EMP_送轉檔
            ChkEED13.Enabled = True
            If Trim(txt3(6).Text) = "" Then
               CboEEP05.Text = Pub_GetSpecMan("中打室排版分案人員") & " " & GetPrjSalesNM(Pub_GetSpecMan("中打室排版分案人員"))
            Else
               CboEEP05.Text = txt3(6).Text & " " & GetPrjSalesNM(txt3(6).Text) '排版人員
            End If
            'Modify By Sindy 2024/1/9
            If PUB_ChkEmpFlowExists(m_EEP01, EMP_送排版) = False Then
               CboEEP05.Enabled = True
            End If
            '2024/1/9 END
         Case EMP_轉檔完成
            ChkEED13.Enabled = True
            Call ChkEED13_Click
'            If ChkEED13.Value = 1 Then
'               txtEEP10_2.Text = Left(m_NPMan, 5) '程序管制人
''               If MsgBox("確定要送件給程序人員嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
''                  Me.SSTab1.Tab = intTab_外專承辦單
''                  Exit Sub
''               End If
'            Else
'               txtEEP10_2.Text = ""
'            End If
'            Call txtEEP10_2_LostFocus
         Case EMP_交辦
            CboEEP05.Clear
            If Pub_StrUserSt03 = "F62" Or Pub_StrUserSt03 = "F72" Then
               strSql = "select st01,st02 from staff" & _
                        " where st03 in('F62','F72') and st01<>'99998' and st04='1' order by st01 desc"
            Else
               '該單位人員
               strSql = "select st01,st02 from staff" & _
                        " where st03='" & Pub_StrUserSt03 & "' and st01<>'99998' and st04='1' order by st01 desc"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If RsTemp.RecordCount > 0 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If Left(RsTemp.Fields("st01"), 5) <> txt3(6) _
                     And Left(RsTemp.Fields("st01"), 5) <> Left(Trim(m_EMMan), 5) Then
                     Call ChkCboEEP05AddItem(RsTemp.Fields("st01") & " " & RsTemp.Fields("st02"))
                  End If
                  RsTemp.MoveNext
               Loop
            End If
         Case EMP_核稿分案
            If cp(14) = "" Then
               MsgBox "尚未分案無承辦人，" & vbCrLf & "請通知分案人員先進行分案作業！", vbExclamation
               CboEEP04.Text = ""
               Exit Sub
            End If
            'Modify By Sindy 2025/1/24 改用共用函數
            Call Frm060101_1_SetCboCP14("", pa(150), CboEEP05)
'            CboEEP05.Clear
'            '該單位人員
'            strSql = "select st01,st02 from staff" & _
'                     " where st03='" & Pub_StrUserSt03 & "' and st04='1'" & _
'                     " and st16='" & PUB_GetStaffST16(strUserNum) & "' order by st01 desc"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If RsTemp.RecordCount > 0 Then
'               RsTemp.MoveFirst
'               Do While Not RsTemp.EOF
'                  Call ChkCboEEP05AddItem(RsTemp.Fields("st01") & " " & RsTemp.Fields("st02"))
'                  RsTemp.MoveNext
'               Loop
'            End If
            If cp(10) = "201" And (Left(m_EPMan, 5) = strUserNum Or Trim(m_EPMan) = "") Then
               CboEEP05.Text = m_TCT10Man '預帶命名工程師
            Else
               CboEEP05.Text = m_EPMan '承辦人
            End If
            '2025/1/24 END
            
         Case EMP_程序送判
            'Modify By Sindy 2025/2/24 mark
'            'Add By Sindy 2024/12/10 可操作多案
'            If bolFCTFlow = True Then
'               Call cmdManyCase_Click
'            End If
'            '2024/12/10 END
            
            'Add By Sindy 2024/1/2 帶二級主管
            'CboEEP05.Text = GetDeptMan("F22") & " " & GetPrjSalesNM(GetDeptMan("F22")) '外專程序主管
            strExc(10) = GetST52SelfList(Left(m_NPMan, 5))
            If strExc(10) <> "" Then
               'Add By Sindy 2024/10/30 外商程序主管休假,人員互判
               If Left(PUB_GetST93(strExc(10)), 1) = "T" Then
                  If ChkEmpIsRest(strExc(10)) = True Then
                     strExc(9) = Pub_GetSpecMan("每月外商延展管制表收件者")
                     varTemp = Split(strExc(9), ";")
                     For intI = 0 To UBound(varTemp)
                        If varTemp(intI) <> Left(m_NPMan, 5) And varTemp(intI) <> "" Then
                           strExc(10) = varTemp(intI)
                           Exit For
                        End If
                     Next intI
                  End If
               End If
               '2024/10/30 END
               CboEEP05.Text = strExc(10) & " " & GetPrjSalesNM(strExc(10))
               Call ChkCboEEP05AddItem(CboEEP05.Text) 'Add By Sindy 2025/2/24
            End If
            
            
         Case EMP_程序退回
            CboEEP05.Text = m_NPMan '程序人員
         '*********************************************************************************
         '2023/9/27 END
         
         'Modify By Sindy 2016/3/7 +EMP_會完重修
         'Modify By Sindy 2016/3/15 +EMP_圖修, EMP_圖完
         Case EMP_草完, EMP_標號, EMP_核修, EMP_核完, EMP_會修, EMP_會完, _
              EMP_圖修, EMP_圖完, EMP_繪圖判發, EMP_草核完, EMP_會完重修, _
              EMP_准許先會, EMP_查名, EMP_查名結果
            CboEEP05.Text = m_EPMan '承辦人
            
            'Add By Sindy 2019/6/28
            If Left(CboEEP04.Text, 2) = EMP_查名結果 Then
               txtEEP10_2.Text = Trim(Mid(m_SPMan, 6)) '智權人員
               Call txtEEP10_2_LostFocus
               'Add By Sindy 2019/7/5
               txtEEP08 = "承辦人：" & Trim(Mid(m_EPMan, 6)) & vbCrLf & _
                          "申請後續請通知承辦人，謝謝！"
            'Modify By Sindy 2019/9/5 + 副本通知客戶組
            ElseIf (Left(CboEEP04.Text, 2) = EMP_會修 Or Left(CboEEP04.Text, 2) = EMP_會完) And _
               InStr(Pub_GetSpecMan("WSpecial"), Me.m_FlowUserNum) > 0 And _
               InStr(Pub_GetSpecMan("客服組專利會稿工程師"), strUserNum) > 0 Then '創新業務部可個人收文成員
               txtEEP10_2.Text = Me.m_FlowUserNum
               Call txtEEP10_2_LostFocus
            '2019/9/5 END
            End If
            
            'Add By Sindy 2018/9/21
            '會完時商標處C類智權人員可以註記是否通知客戶
            'Modify By Sindy 2024/7/11 +  Or bolCFTFlow = True
            If Left(CboEEP04.Text, 2) = EMP_會完 And _
               (bolTMFlow = True Or bolCFTFlow = True) And _
               Left(lblCP09, 1) = "C" Then
               'Add By Sindy 2025/8/5 針對商三組之1201審查報告、1202核駁前先行通知或727分析通知會稿時，
               '管控智權人員不能點選不通知客戶，即不通知客戶之選項不出現。
               If Not ((cp(10) = "1201" Or cp(10) = "1202") And PUB_GetST93(cp(14)) = "T31") And _
                  cp(10) <> "727" Then
               '2025/8/5 END
                  ChkEP11.Visible = True 'EP11.是否通知客戶
                  ChkEP11.Enabled = True
               End If
            End If
            '2018/9/21 END
            
            'Add By Sindy 2018/10/16 可操作多案
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            If (bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True) And Left(CboEEP04.Text, 2) = EMP_會完 Then
               Call cmdManyCase_Click
            End If
            '2018/10/16 END
            
         'Modify By Sindy 2016/3/15 + EMP_會圖
         Case EMP_送會, EMP_會圖
            CboEEP05.Text = m_SPMan '智權人員
            'Modify By Sindy 2019/9/5 + 副本通知核稿人(做客戶組的客戶會稿)。
            If InStr(Pub_GetSpecMan("WSpecial"), Me.m_FlowUserNum) > 0 And _
               m_CMMan <> "" Then '創新業務部可個人收文成員
               txtEEP10_2.Text = Trim(Left(m_CMMan, 6))
               Call txtEEP10_2_LostFocus
            End If
            '2019/9/5 END
            
            'Add By Sindy 2020/9/25 可操作多案
            'And strSrvDate(1) >= T商標電子化第2階段啟用日
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            If (bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True) _
               And Left(CboEEP04.Text, 2) = EMP_送會 Then
               Call cmdManyCase_Click
            End If
            '2020/9/25 END
            
            'Add By Sindy 2020/12/23
            '有關內商申案件發文費用的更動,原設於「送件」歷程顯示「電子送件」按鈕,利於承辦同仁點選，並修改發文規費。
            '現已會經財務處 吳經理同意修改流程，提前至「送會」歷程中修改發文規費，並以此系統通知代替現行另寄Email的方式。
            If Left(CboEEP04.Text, 2) = EMP_送會 Then
               If bolTMFlow = True Then
                  If m_Country = "000" Then
                     cmdCP118.Caption = "修改規費"
                     cmdCP118.Visible = True
                  End If
                  Call GetCP43AddCC(True) 'Add By Sindy 2025/8/5
               End If
            End If
            '2020/12/23 END
            
         Case EMP_送英核
            CboEEP05.Text = m_EMMan '英文核稿人
            m_EEP11Person = m_EMMan '*****
            
         Case EMP_送核
            CboEEP05.Text = m_CMMan '核稿人
            m_EEP11Person = m_CMMan '*****
            
            'Add By Sindy 2020/9/25 可操作多案
            'Val(m_EP06) >= T商標電子化第2階段啟用日
            'And strSrvDate(1) >= T商標電子化第2階段啟用日
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
               Call cmdManyCase_Click
            '2020/9/25 END
            'Add By Sindy 2023/10/31
            'Modify By Sindy 2024/1/12 + m_CMMan = m_CSMan :FCP-60423(926)
'            ElseIf bolFCPFlow = True Then
'               If m_CSMan = "" Or _
'                  Left(m_CSMan, 5) = m_FlowUserNum Or _
'                  (m_CSMan <> "" And Val(m_EP42) > 0 Or _
'                  m_CMMan = m_CSMan) Then
'                  ChkCP113CP114 False
'               End If
            End If
            '2023/10/31 END
            
         'Add By Sindy 2015/4/22
         Case EMP_草核
            CboEEP05.Text = m_DCMan '草圖核稿人
            m_EEP11Person = m_DCMan '*****
            
         '2015/4/22 END
         Case EMP_送判
            CboEEP05.Text = m_CSMan '判發人
            m_EEP11Person = m_CSMan '*****
            
            'Add By Sindy 2020/9/25 可操作多案
            'Val(m_EP06) >= T商標電子化第2階段啟用日
            'And strSrvDate(1) >= T商標電子化第2階段啟用日
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
               Call cmdManyCase_Click
            '2020/9/25 END
            'Add By Sindy 2023/10/31
'            ElseIf bolFCPFlow = True Then
'               ChkCP113CP114 False
            End If
            '2023/10/31 END
            
         'Modify By Sindy 2016/3/9 +EMP_修改圖式
         Case EMP_送標號, EMP_上墨, EMP_草修, EMP_修改圖式
            CboEEP05.Text = m_DPMan '繪圖人員
         Case EMP_墨完
            CboEEP05.Text = m_DMMan '繪圖主管
            m_EEP11Person = m_DMMan '*****
         '2018/4/27 END
         Case EMP_判發 '判發時不必發E-Mail
'            Label15.Visible = False
            'Modify By Sindy 2013/9/9
            'CboEEP05.Text = ""
            If bolPAFlow = True Or bolOtherFlow = True Then
               'Modify By Sindy 2023/1/19
               If cp(1) = "ACS" Then
                  CboEEP05.Text = m_NPMan '程序人員
                  '收受者解除鎖定，改為下拉選單(顧服組所有成員)
                  '顯示副本收受者欄位
                  strSql = "select st01,st02 from staff" & _
                           " where st03='W20' and st04='1' and st01<>'W2001' order by st01 asc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If RsTemp.RecordCount <> 1 Then
                     RsTemp.MoveFirst
                     Do While Not RsTemp.EOF
                        Call ChkCboEEP05AddItem(RsTemp.Fields("st01") & " " & RsTemp.Fields("st02"))
                        RsTemp.MoveNext
                     Loop
                  End If
               Else
               '2023/1/19 END
                  CboEEP05.Text = m_NPMan '程序人員
                  '2013/9/9 END
      '            CboEEP05.Visible = False
                  'Modify By Sindy 2021/6/25 雅娟:有關大陸案在走歷程時,判發後收受者都寫死是品薇,
                  '但有些中間程序是由處理該道程序的人處理,並不是都給品薇,麻煩請改為預設是品薇,但是也可以選擇其他收受者
                  'Modify By Sindy 2021/7/12
                  If Left(m_NPMan, 5) = strUserNum Then
                     CboEEP05.Enabled = False
                  Else
                  '2021/7/12 END
                     CboEEP05.Enabled = True
                  End If
                  '2021/6/25 END
                  Label4.Visible = False
                  Label5.Visible = False
                  txtEEP10_2 = ""
                  txtEEP10_2.Visible = False
               End If
            'Add By Sindy 2018/4/27
            Else
               CboEEP05.Text = m_EPMan '承辦人
               'Add By Sindy 2023/10/4 +if
               If bolTMFlow = True Then
               '2023/10/4 END
                  'Add By Sindy 2021/5/20
                  '該類案件性質(401~410,601~606)的所有判發案件，
                  '在主管"判發"送件時，檢查判發主管不是林律師 或 江律師的，副本就要加上林律師和江律師
                  If (cp(10) >= "401" And cp(10) <= "410") Or _
                     (cp(10) >= "601" And cp(10) <= "606") Then
                     If m_FlowUserNum <> "98020" And m_FlowUserNum <> "98003" Then
                        'Modify By Sindy 2021/6/25 預設用員工編號,因有的人會有多編號; 用姓名轉編號是會混亂,不知要帶那一個會出錯誤訊息
                        'txtEEP10_2.Text = IIf(txtEEP10_2.Text <> "", txtEEP10_2.Text & ",", "") & GetPrjSalesNM("98020") & "," & GetPrjSalesNM("98003")
                        txtEEP10_2.Text = IIf(txtEEP10_2.Text <> "", txtEEP10_2.Text & ",", "") & "98020,98003"
                        Call txtEEP10_2_LostFocus
                        '2021/6/25 END
                     End If
                  End If
                  '2021/5/20 END
               End If
            '2018/4/27 END
            End If
'            'Add By Sindy 2013/11/26 檢查是否有專利處的核判權限
'            If Left(CboEEP04.Text, 2) = EMP_判發 Then
'               If m_FlowUserNum <> strUserNum Then
'                  If PUB_ChkPromoterReader(cstr(pfield(1)), CP(10), "2", strUserNum) = False Then
'                     MsgBox "無代理判發的權限！" & vbCrLf & "請回至前一作業『工作進度資料維護』輸入判發人後，才可進行下一流程。"
'                     Call cmdCancel_Click
'                     Exit Sub
'                  End If
'               End If
'            End If
'            '2013/11/26 END
            
'            'Add By Sindy 2020/9/25 可操作多案 - 自行判發
'            'Val(m_EP06) >= T商標電子化第2階段啟用日
'            If bolTMFlow = True And strSrvDate(1) >= T商標電子化第2階段啟用日 And _
'               (m_CSMan = "" Or Left(m_CSMan, 5) = m_FlowUserNum) And _
'               intReceiveKind = 0 Then '0.承辦人工作進度
'               Call cmdManyCase_Click
'            End If
'            '2020/9/25 END
            
         Case EMP_退回
            'Modify By Sindy 2013/10/16
            'CboEEP05.Text = strLastEEP03 & " " & GetPrjSalesNM(strLastEEP03)
            strRefEEP02_EP = ""
            strRefEEP02_DP = ""
'            'Add By Sindy 2023/10/31
'            If bolFCPFlow = True And m_strLastEEP04 <> EMP_送判 Then
'               '退回上一道發送者
'               CboEEP05.Text = strLastEEP03 & " " & GetPrjSalesNM(strLastEEP03) '發送者
'               '2023/10/31 END
'            Else
            If PUB_ChkEmpFlowExists(m_EEP01, EMP_送判, , strRefEEP02_EP) = True And _
               PUB_ChkEmpFlowExists(m_EEP01, EMP_墨完, , strRefEEP02_DP) = True Then
               If Val(strRefEEP02_EP) > Val(strRefEEP02_DP) Then
                  CboEEP05.Text = m_EPMan '承辦人
               Else
                  CboEEP05.Text = m_DPMan '繪圖人員
               End If
            Else
               If Val(strRefEEP02_EP) > 0 Then
                  CboEEP05.Text = m_EPMan '承辦人
               Else
                  CboEEP05.Text = m_DPMan '繪圖人員
               End If
            End If
            '2013/10/16 END
            
         Case EMP_轉回
            CboEEP05.Text = Mid(strLastEEP11, InStr(strLastEEP11, ":") + 1, 5)
            CboEEP05.Text = CboEEP05.Text & " " & GetPrjSalesNM(CboEEP05.Text)
         
         'Modify By Sindy 2018/4/27 + EMP_送件
         Case EMP_退件重送, EMP_送件, EMP_發文歸檔
            If bolTMFlow = True Then
               'If m_Country = "000" Then 'Modify By Sindy 2018/9/27 ex:T-213528
                  'Add By Sindy 2018/9/25 台灣商標Ｔ,FCT案若收文爭議案件性質時,開放 Frame6 欄位
                  'Modify By Sindy 2024/5/23 排除 311加速審查
                  'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
                  If (cp(1) = "T" Or cp(1) = "FCT") And _
                     cp(10) <> "311" And _
                     ((InStr(TMdebate, cp(10)) > 0 And Not (cp(1) = "FCT" And InStr(FCT_NotTMdebate, cp(10)) > 0)) _
                      Or m_PrevForm.txt1(11) <> "") Then '條款
                     Frame6.Tag = "V" 'Add By Sindy 2020/10/16
                     Frame6.Visible = True
                     If InStr(TMdebate, cp(10)) > 0 Then
                        Option1(0).Enabled = True
                        Option1(1).Enabled = True
                        Option1(2).Enabled = True
                     Else
                        Option1(0).Enabled = False
                        Option1(1).Enabled = False
                        Option1(2).Enabled = False
                     End If
                  End If
               'End If
            End If
            
            If Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
               'Modify By Sindy 2018/9/21 Mark;改在查詢函數時判斷
               If (bolTMFlow = True Or bolCFTFlow = True) Then 'Add By Sindy 2023/11/23 +if
                  ChkEP11.Visible = True 'EP11=N:不通知, 不發文 Add By Sindy 2018/9/20
                  ChkEP11.Enabled = True 'Add By Sindy 2018/9/21
                  'Add By Sindy 2025/8/5
                  strSql = "select * from EMPELECTRONPROCESS where eep01='" & m_EEP01 & "' and eep04='" & EMP_會完 & "' order by eep02 desc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If RsTemp.RecordCount > 0 Then
                     If InStr(RsTemp.Fields("eep08"), "已收文分析") > 0 Then
                        ChkEP11.Enabled = False
                     End If
                  End If
                  '2025/8/5 END
               End If
               CboEEP05.Text = ""
               CboEEP05.Enabled = False
               'Add By Sindy 2019/4/12
               '承辦人將MCTF案件發文歸檔時,系統於收受者欄位,自動帶入該案件之MCTF人員
               If bolMCTFcase = True Then
                  CboEEP05.Text = m_SPMan '智權人員
               End If
               '2019/4/12 END
            
            Else
               CboEEP05.Text = m_NPMan '程序人員
               If bolTMFlow = True Then
                  If m_Country = "000" Then
                     Label4.Visible = False
                     Label5.Visible = False
                     txtEEP10_2 = ""
                     txtEEP10_2.Visible = False
                     cmdCP118.Visible = True
                     'Call cmdCP118_Click '不自動彈出視
   '                  If cp(10) = "101" And cp(118) = "" Then textCP118 = "Y" 'Add By Sindy 2018/7/26 台灣申請案,預設為Y
                  End If
               End If
            End If
            
            'Add By Sindy 2020/9/16 可操作多案
            'Val(m_EP06) >= T商標電子化第2階段啟用日
            'And strSrvDate(1) >= T商標電子化第2階段啟用日
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            'Modify By Sindy 2025/1/22 承辦人操作
            If intReceiveKind = 0 Then '0.承辦人工作進度
            '2025/1/22 END
               If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
                  Call cmdManyCase_Click
               End If
               '2020/9/16 END
            End If
            '2025/1/22 END
            
'            'Add By Sindy 2023/12/18
'            If bolFCPFlow = True Then
'               ChkCP113CP114 False
'            End If
'            '2023/12/18 END
            
         'Add By Sindy 2018/8/29
         Case EMP_客戶會稿
            'Add By Sindy 2018/12/14 文雄提分析案不鎖客戶會稿,反而選此項狀態時要詢問訊息
            If InStr(lblCP10, "分析") > 0 Then
               If MsgBox("此為分析案件，確定需要做客戶會稿嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  CboEEP04.Text = ""
                  Exit Sub
               End If
            End If
            
            '會稿方式
            Call SetCboCP10
            Frame4.Visible = True
            CboCP10.SetFocus
            '收受者
            CboEEP05.Text = ""
            CboEEP05.Enabled = False
            txtEEP10_2.Enabled = False: txtEEP10_2.Text = ""
            'Add By Sindy 2018/9/13 P案會稿時帶出申請資訊
            txtEEP08.Tag = ""
            'Modify By Sindy 2019/5/16 玲玲說要工程師送會時含data.doc給智權人員會稿
'            If cp(1) = "P" And InStr(NewCasePtyList, cp(10)) > 0 Then
'               If PUB_ChkEmpFlowExists(lblCP09, EMP_客戶會稿) = False Then '第一次客戶會稿
'                  txtEEP08.Tag = "*煩請確認以下申請人及發明人資訊，若有任何變動，請務必通知P程序人員，以利更新系統資訊，謝謝！" & _
'                               vbCrLf
'                  txtEEP08.Tag = txtEEP08.Tag & vbCrLf & GetApplData '& vbCrLf
'               End If
'            End If
            '2018/9/13 END
         
         'Add By Sindy 2025/8/1
         Case EMP_收文分析
            CboEEP05.Text = ""
            CboEEP05.Enabled = False
            txtEEP10_2 = ""
            txtEEP10_2.Enabled = False
            
         'Add By Sindy 2013/9/23
         Case EMP_附加流程 '延期是系統直接產生判發或送判流程
            Frame4.Visible = True '案件性質
            CboEEP04.Tag = "" '記錄欲操作的歷程
            'Add By Sindy 2018/11/1 竹平:FCT-042938須核稿
            'Modify By Sindy 2024/3/4 + Or bolFCPFlow = True
            If bolTMFlow = True Or bolFCPFlow = True Then
               '有核稿主管,並且不可自行核稿者
               If m_CMMan <> "" And Left(m_CMMan, 5) <> m_FlowUserNum And Left(m_CMMan, 5) <> Left(m_CSMan, 5) Then
                  CboEEP05.Text = m_CMMan '核稿主管
                  m_EEP11Person = m_CMMan '*****
                  CboEEP04.Tag = "送核" '記錄欲操作的歷程
               '自行判發
               ElseIf m_CSMan = "" Or Left(m_CSMan, 5) = m_FlowUserNum Then
                  CboEEP05.Text = m_NPMan '程序人員
                  CboEEP04.Tag = "送件" '記錄欲操作的歷程
               '送判
               Else
                  CboEEP05.Text = m_CSMan '判發人
                  m_EEP11Person = m_CSMan '*****
                  CboEEP04.Tag = "送判" '記錄欲操作的歷程
               End If
            '2018/11/1 END
            Else '專利處
               '自行判發
               If m_CSMan = "" Or Left(m_CSMan, 5) = m_FlowUserNum Then
                  CboEEP05.Text = m_NPMan '程序人員
                  CboEEP04.Tag = "判發" '記錄欲操作的歷程
               Else
                  CboEEP05.Text = m_CSMan '判發人
                  m_EEP11Person = m_CSMan '*****
                  CboEEP04.Tag = "送判" '記錄欲操作的歷程
               End If
            End If
            'CboEEP05.Enabled = False
            
         Case Else '聯絡
            CboEEP05.Text = ""
            Text1.Visible = True 'Add By Sindy 2014/10/1 備註:「聯絡」的附件，送件後一律刪除，欲留存者請置於存卷資料頁籤
            'Add By Sindy 2018/9/13 加程序人員
            'Modify By Sindy 2023/9/23 + Or bolFCPFlow = True
            'Modify By Sindy 2024/8/22 + Or bolFCTFlow = True
            If (bolPAFlow = True Or bolFCPFlow = True Or bolFCTFlow = True) And _
               m_NPMan <> "" And Left(m_NPMan, 5) <> m_FlowUserNum And _
               ChkStaffST04(Trim(Left(m_NPMan, 6)), False) = False Then
               'Modify By Sindy 2019/7/18
               'CboEEP05.AddItem m_NPMan
               Call ChkCboEEP05AddItem(m_NPMan)
               '2019/7/18 END
               'Add By Sindy 2023/10/2 送中說:Claims翻譯交稿
               If cp(10) = "924" And m_bolSendChWrite = True Then
                  CboEEP05.Text = m_F21CMMan 'FCP工程師主管
                  txtEEP10_2.Text = Left(m_NPMan, 5) '程序管制人
                  Call txtEEP10_2_LostFocus
                  txtEEP08 = GetFCP924txtEEP08("Claims翻譯已交稿，")
               End If
               '2023/10/2 END
            End If
            '2018/9/13 END
            If bolPAFlow = True Then
               'Add By Sindy 2017/8/31 顯示多國案鈕
               cmdCaseMap.Visible = True
               m_RetrunRecv = "" '回傳總收文號
               If frm090202_2_1.QueryData("0") = True Then
                  cmdCaseMap.Enabled = True
               Else
                  cmdCaseMap.Enabled = False
               End If
               '2017/8/31 END
            End If
            
            'Add By Sindy 2020/9/25 可操作多案
            'Val(m_EP06) >= T商標電子化第2階段啟用日
            'Modify By Sindy 2022/4/26 聯絡開放可以多案操作
'            If bolTMFlow = True And strSrvDate(1) >= T商標電子化第2階段啟用日 And _
'               (m_EP06 = "") Then
            'And strSrvDate(1) >= T商標電子化第2階段啟用日
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            If bolTMFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
            '2022/4/26 END
               If intReceiveKind = 0 Then '承辦人工作進度=>承辦人操作的聯絡
                  Call cmdManyCase_Click
               End If
            End If
            '2020/9/25 END
      End Select
      If Left(CboEEP04.Text, 2) <> EMP_客戶會稿 Then cmdSend.Caption = "送出(&O)" 'Add By Sindy 2018/10/23
      '收受者
      If CboEEP05.Text <> "" Then
         'Add By Sindy 2023/10/2 +排除 EMP_聯絡, EMP_交辦 和 Not (Left(CboEEP04.Text, 2) = EMP_核稿分案 And cp(10) = "201")
         'Modify By Sindy 2023/12/6 淑華說"程序送判"有時開會,會請職代處理
         'Modify By Sindy 2024/1/9 排除歷程上線後的中途送轉檔
         If Left(CboEEP04.Text, 2) <> EMP_聯絡 And _
            Left(CboEEP04.Text, 2) <> EMP_交辦 And _
            Not (Left(CboEEP04.Text, 2) = EMP_核稿分案 And cp(10) = "201") And _
            Left(CboEEP04.Text, 2) <> EMP_程序送判 And _
            Not (Left(CboEEP04.Text, 2) = EMP_送轉檔 And PUB_ChkEmpFlowExists(m_EEP01, EMP_送排版) = False) Then
         '2023/10/2 END
            'Add By Sindy 2021/6/25 雅娟:有關大陸案在走歷程時,判發後收受者都寫死是品薇,
            '但有些中間程序是由處理該道程序的人處理,並不是都給品薇,麻煩請改為預設是品薇,但是也可以選擇其他收受者
            'Add By Sindy 2023/1/19 + Not (cp(1) = "ACS" And Left(CboEEP04.Text, 2) = EMP_判發)
            If Not (bolPAFlow = True And Left(CboEEP04.Text, 2) = EMP_判發) And _
               Not (cp(1) = "ACS" And Left(CboEEP04.Text, 2) = EMP_判發) Then
            '2021/6/25 END
               CboEEP05.Enabled = False
            Else
               'Add By Sindy 2022/4/8 下拉式選單加入程序人員
               CboEEP05.Clear
               Call ChkCboEEP05AddItem(m_NPMan)
               '2022/4/8 END
            End If
         End If
      End If
      
      'Add By Sindy 2024/4/29
      If bolFCPFlow = True Then
         'Add By Sindy 2025/4/7 945=電話聯絡單
         If cp(10) = "945" And _
            (Left(CboEEP04.Text, 2) = EMP_送判 Or Left(CboEEP04.Text, 2) = EMP_發文歸檔) Then
            ChkEED08.Visible = True
            ChkEED08.Enabled = True
         End If
         If Me.Frame945.Tag <> "" And _
            (Left(CboEEP04.Text, 2) = EMP_送判 Or (m_CSMan = "" And Left(CboEEP04.Text, 2) = EMP_送件)) Then
            If txtEED14.Text = "" Then
               '【管制下一程序期限】
               If MsgBox("是否要輸入" & _
                  IIf(Me.Frame945.Tag = "945", "委員指定送件日期", "追蹤客戶指示管制日期") & _
                  "，欲" & Frame945.Caption & "嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                  Me.SSTab1.Tab = intTab_外專承辦單
                  txtEED14.SetFocus
               End If
            End If
         End If
         '2025/4/7 END
         
         '檢查暫不送件,指定送件日
         '先清除可能的舊資訊
         txtEEP08 = Replace(txtEEP08, "【已獲客戶最終指示】", "")
         txtEEP08 = Replace(txtEEP08, "【尚待客戶最終指示】", "")
         If Left(CboEEP04.Text, 2) = EMP_送核 Or Left(CboEEP04.Text, 2) = EMP_送判 Or _
            Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送 Then
            
            bolCP141 = True: strReVal = "" '預設值
            'Add By Sindy 2024/7/9
            '當新案翻譯【暫不送】未勾選且進度備註未有"取消待客戶最終指示"2條件同時成立
            If cp(10) = "201" And cp(176) = "" And InStr(cp(64), "取消待客戶最終指示") = 0 Then
               '當【新案翻譯】有收文【924=會稿】相關收文號為新案翻譯那道
               strSql = "select cp09,cp10 from caseprogress" & _
                        " where cp43='" & m_EEP01 & "' and cp10 = '924' and cp159=0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If RsTemp.RecordCount > 0 Then
                  If MsgBox("本案有收文會稿，請確認是否已獲客戶最終指示？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                     '若按【是】，進度備註會自動備註取消日
                     strReVal = "Y"
                     strSql = "update caseprogress set cp176=null,cp64='於" & ChangeWStringToTDateString(strSrvDate(1)) & "取消待客戶最終指示;'||cp64 where cp09='" & m_EEP01 & "'"
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql
                     MsgBox "已取消暫不送件!", vbInformation
                  '若按【否】，【暫不送】打勾
                  Else
                     bolCP141 = False
                     strReVal = "N"
                     strSql = "update caseprogress set cp176='Y',cp64='" & ChangeWStringToTDateString(strSrvDate(1)) & "需待客戶最終指示;'||cp64 where cp09='" & m_EEP01 & "'"
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql
                     MsgBox "已設定暫不送件!", vbInformation
                  End If
               End If
            End If
            If strReVal = "" Then
            '2024/7/9 END
               bolCP141 = PUB_FCPChkCP141(m_EEP01, strReVal)
            End If
            If strReVal = "Y" Then
               txtEEP08 = txtEEP08 & "【已獲客戶最終指示】"
            ElseIf strReVal = "N" Then
               txtEEP08 = txtEEP08 & "【尚待客戶最終指示】"
            End If
            If (Left(CboEEP04.Text, 2) = EMP_送件 Or Left(CboEEP04.Text, 2) = EMP_退件重送) And bolCP141 = False Then
               '不可送件
               CboEEP04.Text = ""
               Exit Sub
            End If
         End If
      End If
      '2024/4/29 END
      
      'Add By Sindy 2018/4/13
      '檢查是否有回覆的權限
      If Left(CboEEP04.Text, 2) = EMP_核修 Or _
         Left(CboEEP04.Text, 2) = EMP_核完 Or _
         Left(CboEEP04.Text, 2) = EMP_會修 Or _
         Left(CboEEP04.Text, 2) = EMP_會完 Or _
         Left(CboEEP04.Text, 2) = EMP_草修 Or _
         Left(CboEEP04.Text, 2) = EMP_草核完 Or _
         Left(CboEEP04.Text, 2) = EMP_圖修 Or _
         Left(CboEEP04.Text, 2) = EMP_圖完 Then
         '2018/4/13 CFP-027475(CA7018759)/A3012,98024：應該不可執行中四區的會完
         strSql = "select EEP01,EEP09 from empelectronprocess" & _
                  " where eep01='" & m_EEP01 & "' and EEP09 = 'Y'"
         If m_CurrFlowEEP02 > 0 Then
            strSql = strSql & " and EEP02=" & m_CurrFlowEEP02
         Else
            strSql = strSql & " and EEP05='" & m_FlowUserNum & "'"
            strSql = strSql & " and eep02<=" & intLastEEP02
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If RsTemp.RecordCount <> 1 Then
            MsgBox "不可執行此回覆歷程，無權限！", vbExclamation
            CboEEP04.Text = ""
            Exit Sub
         End If
      End If
      '2018/4/13 END
      
      'Add By Sindy 2018/1/25
      If lblCM10.Visible = True And lblCM10.Tag <> "" And _
         (cp(10) = "101" Or cp(10) = "102") Then
         If Left(CboEEP04.Text, 2) = EMP_送核 Or Left(CboEEP04.Text, 2) = EMP_會修 Or _
            Left(CboEEP04.Text, 2) = EMP_會完 Or Left(CboEEP04.Text, 2) = EMP_送判 Or _
            Left(CboEEP04.Text, 2) = EMP_判發 Or _
            (Left(CboEEP04.Text, 2) = EMP_送會 And PUB_ChkEmpFlowExists(lblCP09, EMP_送會) = False) Then
            MsgBox "本案為一案兩請，另一案為" & lblCM10.Tag & "，請記得跑相同歷程。", vbInformation
         End If
      End If
      '2018/1/25 END
      
      Call ChkDutyAgent(True) '檢查收受者是否休假
      
      'Add By Sindy 2013/8/22 送核開放可以自行輸入核稿人,會有多次核稿狀況
      'Modify By Sindy 2013/9/2 送判亦也開放可自行輸入
      If Left(CboEEP04.Text, 2) = EMP_送核 Or Left(CboEEP04.Text, 2) = EMP_送判 Then
         CboEEP05.Enabled = True
      End If
      
      'Add By Sindy 2013/10/7
      If m_DMMan <> "" Then '有繪圖
         If Left(CboEEP04.Text, 2) = EMP_送判 Or Left(CboEEP04.Text, 2) = EMP_判發 Then
            'Add By Sindy 2014/3/13 當承辦人為專利處繪圖的人員時,不需檢查此條件
            'Modify By Sindy 2016/3/24 品薇承辦的非台灣案,由於目前流程有改變,故請取消繪圖沒有判發的控制
            'Modify By Sindy 2018/10/4 98012改判斷是P12專利處程序
'            If PUB_GetStaffST15(Left(m_EPMan, 5), "1") <> "P13" And _
'               Not (Left(m_EPMan, 5) = "98012" And m_Country <> "000") Then
            If PUB_GetStaffST15(Left(m_EPMan, 5), "1") <> "P13" And _
               Not (PUB_GetST03(Left(m_EPMan, 5)) = "P12" And m_Country <> "000") Then
            '2014/3/13 END
               'Modify By Sindy 2018/6/25 P120049封裝基板 - 有會完重修,系統上墨後工程師直接判發了,導致沒有墨完日
               '+ Or Val(m_EP18) = 0
               If PUB_ChkEmpFlowExists(lblCP09, EMP_繪圖判發) = False Or Val(m_EP18) = 0 Then
                  MsgBox "尚未繪圖判發，不可" & Trim(Mid(CboEEP04.Text, 3)) & "！"
                  CboEEP04.Text = ""
                  Exit Sub
               End If
            End If
         End If
      End If
      
      'Add By Sindy 2015/3/4 一定要經過英核的案件，若未完成英核動作不可判發
      If Left(CboEEP04.Text, 2) = EMP_送判 Or Left(CboEEP04.Text, 2) = EMP_判發 Then
         'If bolHadSetProofEngReader = True And m_EMMan <> "" And m_PER04 <> Left(Trim(m_EPMan), 5) Then
         If m_EMMan <> "" And bolMultinationalEngOk = False Then
            '檢查是否有送英核
            If PUB_ChkEmpFlowExists(m_EEP01, EMP_送英核) = False And bol00EngCMFlow = False Then
               If m_EP41 = "2" Then
                  MsgBox "此案件未送日核，不可" & Trim(Mid(CboEEP04.Text, 3)) & "！"
               Else
                  MsgBox "此案件未送英核，不可" & Trim(Mid(CboEEP04.Text, 3)) & "！"
               End If
               CboEEP04.Text = ""
               Exit Sub
            End If
         End If
      End If
      '2015/3/4 END
      
      If bolPAFlow = True Then
         'Add By Sindy 2013/9/27 新案,是否有保密審查未准
         If InStr(NewCasePtyList, cp(10)) > 0 And _
            (Left(CboEEP04.Text, 2) = EMP_送判 Or _
             ((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發)) Then
            If PUB_Exists430NotPassed(pa, True, strRefCaseNo) = True Then
               MsgBox "關聯案 " & strRefCaseNo & " 保密審查尚未核准，本案不可" & Trim(Mid(CboEEP04.Text, 3)) & "！"
               'Call cmdCancel_Click
               CboEEP04.Text = ""
               Exit Sub
            End If
         End If
         'Add By Sindy 2015/12/2
         If Left(CboEEP04.Text, 2) = EMP_送會 Then
            If PUB_ChkEmpFlowExists(m_EEP01, EMP_送會) = False And Val(m_EP07) > 0 Then
               If MsgBox("此流程將會刪除由原關連案帶入的會稿日及會稿完成日，以此次會稿日期為準。" & vbCrLf & _
                         "按『是』：繼續操作送會，送出時系統會更新會稿日及清除會完日。" & vbCrLf & _
                         "按『否』：不送會。", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  CboEEP04.Text = ""
                  ChkEMail.Visible = False
                  Exit Sub
               'Add By Sindy 2016/9/26
               Else
                  m_EP08 = ""
               '2016/9/26 END
               End If
            End If
         End If
         '2015/12/2 END
      End If
      
      'Add By Sindy 2013/9/11 已有會稿完成日,不可送會
      If m_EP08 <> "" And Left(CboEEP04.Text, 2) = EMP_送會 Then
         MsgBox "已有會稿完成日，不可送會！"
         'CboEEP04.ListIndex = CboEEP04.ListCount - 1 '流程狀態預設為第一道流程
         CboEEP04.Text = ""
         ChkEMail.Visible = False
         Exit Sub
      End If
      
      Call AskEmpIsCopyFile '詢問是否要沿用附件
   End If
   
   Call SettxtEEP10_2 'Add By Sindy 2017/6/15
   Call ShowRemindMsg 'Add By Sindy 2025/7/9
End Sub

'詢問是否要沿用附件
Private Sub AskEmpIsCopyFile()
Dim rsA As New ADODB.Recordset
Dim intText As Integer, intMsgQ As Integer
   
   '無附件時
   bolDeleteFile = False
   
   If lstAtt(0).ListCount = 0 Or (cmdManyCase.Visible = True And cmdManyCase.Enabled = True) Then
      'Modify By Sindy 2013/10/23
      '(m_strLastEEP04 = EMP_送核 And Left(CboEEP04.Text, 2) = EMP_判發) or
      '((m_CSMan = "" Or Left(m_CSMan, 5) = Left(m_EPMan, 5)) And Left(CboEEP04.Text, 2) = EMP_判發) Or
      ' ==> Left(CboEEP04.Text, 2) = EMP_判發
      'Modify By Sindy 2014/2/20 +EMP_核完,EMP_繪圖判發
      'Modify By Sindy 2016/3/15 +EMP_會圖
      'Modify By Sindy 2018/5/23 +EMP_送件
      'Modify By Sindy 2018/10/15 客戶會稿沿用附件另外處理
      '           (Left(CboEEP04.Text, 2) = EMP_客戶會稿 And Trim(Left(CboCP10.Text, 1)) = "1" And bolLastFile = True) Or
      '           (Left(CboEEP04.Text, 2) = EMP_會完 And bolLastFile = True) Or
      'Modify By Sindy 2023/10/26 +EMP_程序送判
      'Modify By Sindy 2024/12/17 +EMP_送核
      If Left(CboEEP04.Text, 2) = EMP_送核 Or _
         Left(CboEEP04.Text, 2) = EMP_退件重送 Or _
         Left(CboEEP04.Text, 2) = EMP_會圖 Or _
         Left(CboEEP04.Text, 2) = EMP_送會 Or _
         Left(CboEEP04.Text, 2) = EMP_送判 Or _
         Left(CboEEP04.Text, 2) = EMP_判發 Or _
         Left(CboEEP04.Text, 2) = EMP_核完 Or _
         Left(CboEEP04.Text, 2) = EMP_草核完 Or _
         Left(CboEEP04.Text, 2) = EMP_繪圖判發 Or _
         Left(CboEEP04.Text, 2) = EMP_送件 Or _
         Left(CboEEP04.Text, 2) = EMP_程序送判 Or _
         ((bolPAFlow = True Or bolOtherFlow = True) And Left(CboEEP04.Text, 2) = EMP_客戶會稿) Then
         If bolLastFile = True Then
            'Add By Sindy 2023/11/9
            If Left(CboEEP04.Text, 2) = EMP_程序送判 Or m_strLastEEP04 = EMP_程序送判 Then
               '無條件將承辦人的附件沿用過來,後面送件-歸卷使用
               If DownloadAttFile_copy(m_EEP01, intLastEEP02) = True Then
                  bolDeleteFile = True
               End If
               'Modify By Sindy 2024/11/21 mark:薛經理覺得多加了,不需要此訊息
               'MsgBox "無條件將承辦人的送件附件沿用過來(最終完整附件)！", vbInformation, "注意！"
               '2024/11/21 END
            '2023/11/9 END
            '系統會將上一筆流程的附件一併刪除。 ==> 系統會將前一道附件移至此流程。
            'Add By Sindy 2020/9/24
            'Modify By Sindy 2024/7/11 + Or bolCFTFlow = True Or bolFCTFlow = True
            ElseIf bolTMFlow = True Or bolOtherFlow = True Or bolFCPFlow = True Or bolCFTFlow = True Or bolFCTFlow = True Then
               'Modify By Sindy 2020/9/24 + 商標標沿用檔案的訊息，增加取消鍵
               intMsgQ = MsgBox("是否要沿用上一筆（" & intLastEEP02 & "." & m_strLastEEP04Nm & "）附件？" & vbCrLf & _
                                "按『是』：沿用檔案，系統會將前一道附件移至此流程。" & vbCrLf & _
                                "按『否』：要輸入【歷程序號】做附件沿用。" & vbCrLf & _
                                "按『取消』：不沿用！請自行添加附件。", vbExclamation + vbYesNoCancel + vbDefaultButton3, "重要訊息！")
               If intMsgQ = vbYes Then
                  If DownloadAttFile_copy(m_EEP01, intLastEEP02) = True Then
                     bolDeleteFile = True
                  End If
               ElseIf intMsgQ = vbNo Then
                  intText = Val(InputBox("要沿用那一道歷程附件，請輸入歷程序號？" & vbCrLf & "空白:則代表不沿用附件"))
                  If Val(intText) > 0 Then
                     If DownloadAttFile_copy(m_EEP01, intText) = True Then
                        bolDeleteFile = True
                     End If
                  End If
               End If
            Else
            '2020/9/24 END
               If MsgBox("是否要沿用上一筆（" & intLastEEP02 & "." & m_strLastEEP04Nm & "）附件？" & vbCrLf & _
                         "按『是』：沿用檔案，系統會將前一道附件移至此流程。" & vbCrLf & _
                         "按『否』：不沿用！請自行添加附件。", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                  If DownloadAttFile_copy(m_EEP01, intLastEEP02) = True Then
                     bolDeleteFile = True
                  End If
               End If
            End If
         
         'Add By Sindy 2020/10/16
         'Modify By Sindy 2023/11/9 + And Left(CboEEP04.Text, 2) <> EMP_程序送判
         ElseIf lstAtt(0).ListCount = 0 And bolPAFlow <> True _
            And ((GRD1.Rows - 1) >= 0 And GRD1.TextMatrix(1, 0) <> "") _
            And Left(CboEEP04.Text, 2) <> EMP_程序送判 Then
            'Add By Sindy 2023/10/31
            '檢查歷程目前裡面是否有放附件
            strExc(0) = "select eef01 From empelectronfile" & _
                        " where eef01='" & m_EEP01 & "' and eef12 is not null and eef02<>0"
            If rsA.State <> adStateClosed Then rsA.Close
            rsA.CursorLocation = adUseClient
            rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
            '2023/10/31 END
               intText = Val(InputBox("要沿用那一道歷程附件，請輸入歷程序號？" & vbCrLf & "空白:則代表不沿用附件"))
               If Val(intText) > 0 Then
                  If DownloadAttFile_copy(m_EEP01, intText) = True Then
                     bolDeleteFile = True
                  End If
               End If
               '2020/10/16 END
            End If
         End If
         
      'Add By Sindy 2018/5/23 +EMP_發文歸檔
      'Modify By Sindy 2018/9/4 +EMP_客戶會稿, EMP_會完
      ElseIf Left(CboEEP04.Text, 2) = EMP_發文歸檔 Or _
         (Left(CboEEP04.Text, 2) = EMP_客戶會稿 And Trim(Left(CboCP10.Text, 1)) = "1") Or _
         Left(CboEEP04.Text, 2) = EMP_會完 Then
         '下載最近的相關歷程附件
         If rsA.State <> adStateClosed Then rsA.Close
         If Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
            strExc(0) = "select eep02 From empelectronprocess" & _
                        " where eep01='" & m_EEP01 & "' and eep04='" & EMP_判發 & "'" & _
                        " order by eep02 desc"
         Else
            strExc(0) = "select eep02,eep04 From empelectronprocess,empelectronfile" & _
                        " where eep01='" & m_EEP01 & "' and eep02>=" & intLastEEP02 & _
                        " and eep04 in('" & EMP_客戶會稿 & "','" & EMP_送會 & "')" & _
                        " and eep01=eef01 and eep02=eef02" & _
                        " order by eep02 desc"
         End If
         rsA.CursorLocation = adUseClient
         rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If Left(CboEEP04.Text, 2) = EMP_發文歸檔 Then
               Call DownloadAttFile_copy(m_EEP01, rsA.Fields("eep02"))
            Else
               'Modify By Sindy 2018/10/29 客戶會稿E-Mail不彈詢問訊息,直接沿用
               If Left(CboEEP04.Text, 2) = EMP_客戶會稿 Then
                  If cmdManyCase.Visible = True And cmdManyCase.Enabled = True Then
                     '多案、單案人員按確定時無條件沿用附件，取消代表不沿用附件
                     If cmdManyCase.Tag = "確定" Then
                        Call ManyCaseMoveFile
                     Else
                        lstAtt(0).Clear
                     End If
                  Else
                     Call DownloadAttFile_copy(m_EEP01, rsA.Fields("eep02"))
                  End If
               ElseIf Left(CboEEP04.Text, 2) = EMP_會完 Then
                  If cmdManyCase.Visible = True And cmdManyCase.Enabled = True Then
                     '多案件人員按確定時無條件沿用附件，取消代表不沿用附件
                     If cmdManyCase.Tag = "確定" Then
                        '單筆時, 要詢問是否沿用
                        If InStr(Me.m_RetrunRecv, ",") = 0 Then
                           If MsgBox("是否要沿用上一筆（" & rsA.Fields("eep02") & "." & IIf(rsA.Fields("eep04") = EMP_客戶會稿, "客戶會稿", IIf(rsA.Fields("eep04") = EMP_會完, "會完", "")) & "）附件？" & vbCrLf & _
                                     "按『是』：沿用檔案，系統會將前一道附件複製至此流程。" & vbCrLf & _
                                     "按『否』：不沿用！請自行添加附件。", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                              Call DownloadAttFile_copy(m_EEP01, rsA.Fields("eep02"))
                           End If
                        Else
                           Call ManyCaseMoveFile
                        End If
                     Else
                        lstAtt(0).Clear
                     End If
                  Else
                     If MsgBox("是否要沿用上一筆（" & rsA.Fields("eep02") & "." & IIf(rsA.Fields("eep04") = EMP_客戶會稿, "客戶會稿", IIf(rsA.Fields("eep04") = EMP_會完, "會完", "")) & "）附件？" & vbCrLf & _
                               "按『是』：沿用檔案，系統會將前一道附件複製至此流程。" & vbCrLf & _
                               "按『否』：不沿用！請自行添加附件。", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                        Call DownloadAttFile_copy(m_EEP01, rsA.Fields("eep02"))
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   
   Set rsA = Nothing
End Sub

'檢查人員是否有休假,若有,則可自行輸入
'bolChange = True : 清除並且重組收受者下拉選單
Private Sub ChkDutyAgent(bolChange As Boolean)
Dim m_ABS001_1 As String
Dim m_ABS001_2 As String
Dim m_ABS001_3 As String
Dim i As Integer, j As Integer, strData As String, varTemp As Variant
Dim strRestKind As String
Dim strDutyAgent As String
   
   If m_EEP11Person <> "" Then
      If bolChange = True Then
         CboEEP05.Clear
         If Left(m_EEP11Person, 5) <> m_FlowUserNum Then
            'Modify By Sindy 2019/7/18
            'CboEEP05.AddItem m_EEP11Person
            Call ChkCboEEP05AddItem(m_EEP11Person)
            '2019/7/18 END
         End If
         CboEEP05.Text = m_EEP11Person
         'Add By Sindy 2018/11/23
         If Left(CboEEP04.Text, 2) <> EMP_附加流程 Then
         '2018/11/23 END
            CboEEP05.Enabled = True
         End If
      End If
      
      strDutyAgent = GetCaseDutyAgent(Trim(Left(CboEEP05.Text, 6)), "", False, strRestKind)
      If strDutyAgent <> "" Then
         'Add By Sindy 2018/11/23
         If Left(CboEEP04.Text, 2) = EMP_附加流程 Then
            MsgBox strRestKind & "！" & vbCrLf & vbCrLf & _
                   "若要改收受者，請至承辦進度調整" & IIf(bolTMFlow = True, "核/判", "判發") & "人員。"
         '2018/11/23 END
         Else
            MsgBox strRestKind & "！" & vbCrLf & vbCrLf & _
                   "收受者可點選或自行輸入職代。"
         End If
         
         If bolChange = True Then
            'Modify By Sindy 2014/5/6
            'Call GetABS001_1(Left(m_EEP11Person, 5), m_ABS001_1, m_ABS001_2, m_ABS001_3)
            Call GetABS001_CaseSys(Left(m_EEP11Person, 5), m_ABS001_1, m_ABS001_2, m_ABS001_3)
            '2014/5/6 END
            For j = 1 To 3 '有3組職代
               strData = ""
               If j = 1 And m_ABS001_1 <> "" Then strData = m_ABS001_1
               If j = 2 And m_ABS001_2 <> "" Then strData = m_ABS001_2
               If j = 3 And m_ABS001_3 <> "" Then strData = m_ABS001_3
               If strData <> "" Then
                  varTemp = Split(strData, ",")
                  For i = 0 To UBound(varTemp)
                     '休假者不列出
                     If ChkEmpIsRest(CStr(varTemp(i))) = False Then
                        If varTemp(i) <> m_FlowUserNum Then
                           'Modify By Sindy 2019/7/18
                           'CboEEP05.AddItem varTemp(i) & " " & GetPrjSalesNM(CStr(varTemp(i)))
                           Call ChkCboEEP05AddItem(varTemp(i) & " " & GetPrjSalesNM(CStr(varTemp(i))))
                           '2019/7/18 END
                        End If
                     End If
                  Next i
               End If
            Next j
         End If
      Else
         If bolChange = True Then
            CboEEP05.Enabled = False
         End If
      End If
   Else
      If CboEEP05.Text <> "" Then
         strDutyAgent = GetCaseDutyAgent(Trim(Left(CboEEP05.Text, 6)), "", False, strRestKind)
         If strDutyAgent <> "" Then
            MsgBox strRestKind & "！" & vbCrLf & vbCrLf & _
                   "若案件緊急請通知職代" & GetPrjSalesNM(strDutyAgent) & "，代為處理。"
         End If
      End If
   End If
End Sub

'檢查人員是否休假
Private Function ChkEmpIsRest(strUserId As String) As Boolean
   ChkEmpIsRest = False
   If CheckIsPersonRest(strUserId, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = True Then
      ChkEmpIsRest = True
   End If
End Function

Private Sub CboEEP04_Validate(Cancel As Boolean)
   If CboEEP04.Text <> "" Then
      '檢查輸入的流程代碼是否有在下拉式選單裡
      For ii = 0 To CboEEP04.ListCount - 1
         If Left(CboEEP04.Text, 2) = Left(CboEEP04.List(ii), 2) Then
            CboEEP04.Text = CboEEP04.List(ii)
            Exit For
         End If
         If ii = CboEEP04.ListCount - 1 Then
            MsgBox "輸入的流程代碼有誤！"
            Call CboEEP04_GotFocus
            Cancel = True
            Exit Sub
         End If
      Next ii
   End If
End Sub

Private Sub CboEEP05_GotFocus()
   InverseTextBox CboEEP05
End Sub

Private Sub CboEEP05_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboEEP05_LostFocus()
Dim strText As String
   CboEEP05.Text = Trim(CboEEP05.Text) 'Add By Sindy 2023/10/18
   If CboEEP05.Text > "" And (CboEEP05.Text <> CboEEP05.Tag Or strEEP05_Err <> "") Then
      strEEP05_Err = ""
      If IsNumeric(Mid(Trim(CboEEP05.Text), 2, 4)) Then
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Trim(Left(CboEEP05.Text, 6)))
         If strText <> "" Then
            CboEEP05.Text = Trim(Left(CboEEP05.Text, 6)) & " " & strText
            '檢查人員是否存在或離職
            If ChkStaffST04(Trim(Left(CboEEP05.Text, 6))) = True Then
               'CboEEP05.SetFocus
               If CboEEP05.Enabled = True Then CboEEP05.SetFocus
               Call CboEEP05_GotFocus
               strEEP05_Err = "Y"
               'Cancel = True
               Exit Sub
            End If
         Else
            strEEP05_Err = strEEP05_Err & CboEEP05.Text & ","
         End If
      Else
         '依員工姓名抓取員工編號
         strText = GetPrjSalesNM_2(CboEEP05.Text, , , , , False)
         If strText <> "" Then
            CboEEP05.Text = strText & " " & CboEEP05.Text
            '檢查人員是否存在或離職
            If ChkStaffST04(strText) = True Then
               CboEEP05.SetFocus
               Call CboEEP05_GotFocus
               strEEP05_Err = "Y"
               'Cancel = True
               Exit Sub
            End If
         Else
            strEEP05_Err = strEEP05_Err & CboEEP05.Text & ","
         End If
      End If
      CboEEP05.Tag = CboEEP05.Text
      strEEP05_Err = Left(strEEP05_Err, IIf(Len(strEEP05_Err) - 1 < 0, 0, Len(strEEP05_Err) - 1))
      If Trim(strEEP05_Err) <> "" Then
         MsgBox "收受者資料有誤！(" & strEEP05_Err & ")"
         CboEEP05.SetFocus
         Call CboEEP05_GotFocus
         'Cancel = True
         Exit Sub
      Else
         'Add By Sindy 2025/6/9 為了不要重覆出現訊息
         If cmdSend.Enabled = True Then
         '2025/6/9 END
            Call ChkDutyAgent(False) '檢查收受者是否休假
         End If
      End If
      
      Call CboEEP05_Click
   End If
End Sub

'Modify By Sindy 2025/4/7
'Private Sub CboEEP05_Validate(Cancel As Boolean)
Public Sub CboEEP05_Validate(Cancel As Boolean)
'2025/4/7 END
Dim strMsgText As String
   
   If CboEEP05 <> "" Then
      Call CboEEP05_LostFocus
      If strEEP05_Err <> "" Then
         Cancel = True
         Exit Sub
      End If
      'Add By Sindy 2013/10/16 當送核流程時,收受者不可為送英核名單的人員
      If Left(CboEEP04.Text, 2) = EMP_送核 Then
         'Modify By Sindy 2022/11/24 王協理2022/12/01即將退休，此處程式僅刪除王文安即可
         If strSrvDate(1) >= "20221130" Then
            strExc(0) = "select st01 from staff where st04='1' and st03 in ('P14','F71') and st01='" & Trim(Left(CboEEP05.Text, 6)) & "'"
         Else
         '2022/11/24 END
            strExc(0) = "select st01 from staff where st04='1' and (st01='88003' or st03 in ('P14','F71')) and st01='" & Trim(Left(CboEEP05.Text, 6)) & "'"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox Trim(Mid(CboEEP05, 6)) & "為送英(日)核人員，流程狀態必須為送英(日)核。" & vbCrLf & "或輸入其他人員！", vbExclamation
            CboEEP05.SetFocus
            Call CboEEP05_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
      '2013/10/16 END
      
      If Trim(Left(CboEEP05.Text, 6)) = txtEEP03 Then
         'Modify By Sindy 2013/10/16
         If m_FlowUserNum <> strUserNum Then
            '當事人請假代為操作流程:因此不控管此條件。
            '如.P-099461在102/10/15景惠休假,黃俊仁要代為送核,但核稿人是黃俊人本人
         '2013/10/16 END
         'Modify By Sindy 2024/1/4 核稿分案不檢查此條件,因有可能是主管自己的案件 ex:FCP-70336
         ElseIf Left(CboEEP04.Text, 2) <> EMP_核稿分案 Then
         '2024/1/4 END
            MsgBox "不可為本人！", vbExclamation
            If CboEEP05.Enabled = True And CboEEP05.Visible = True Then CboEEP05.SetFocus
            Call CboEEP05_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
      
      'Modify By Sindy 2018/10/1 商標處不鎖權限,專利處才鎖
      If bolPAFlow = True Then
      '2018/10/1 END
         'Add By Sindy 2015/4/28
         If Left(CboEEP04.Text, 2) = EMP_送核 And (Trim(Left(CboEEP05.Text, 6)) <> Left(m_CMMan, 5) Or m_CMMan = "") Then
            'Add By Sindy 2018/3/5 承辦人非程序人員時,才需檢查核判權限
            If GetStaffDepartment(Trim(Left(CboEEP05.Text, 6))) <> "P12" Then
            '2018/3/5 END
               'Modify By Sindy 2024/6/26 +m_Country
               If PUB_ChkPromoterReader(CStr(PField(1)), cp(10), "1", Trim(Left(CboEEP05.Text, 6)), , m_Country) = False Then
                  MsgBox "此人無核稿權限！"
                  CboEEP05.SetFocus
                  Call CboEEP05_GotFocus
                  Cancel = True
                  Exit Sub
               End If
            End If
         ElseIf Left(CboEEP04.Text, 2) = EMP_送判 And (Trim(Left(CboEEP05.Text, 6)) <> Left(m_CSMan, 5) Or m_CSMan = "") Then
            'Add By Sindy 2018/3/5 承辦人非程序人員時,才需檢查核判權限
            If GetStaffDepartment(Trim(Left(CboEEP05.Text, 6))) <> "P12" Then
            '2018/3/5 END
               'Modify By Sindy 2024/6/26 +m_Country
               'Modify By Sindy 2025/3/20 補加"," 少了,傳入的參數就不同了
               If PUB_ChkPromoterReader(CStr(PField(1)), cp(10), "2", Trim(Left(CboEEP05.Text, 6)), , m_Country) = False Then
                  MsgBox "此人無判發權限！"
                  CboEEP05.SetFocus
                  Call CboEEP05_GotFocus
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
         '2015/4/28 END
      End If
   End If
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
'   Dim bolHadShowMsg As Boolean 'Add By Sindy 2018/9/26
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
'         If PUB_ChkFileOpening(stAttPath, bolHadShowMsg) = True Then
'            'Modify By Sindy 2018/9/26
'            If bolHadShowMsg = False Then
'            '2018/9/26 END
'               MsgBox stAttPath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
'            End If
'            Screen.MousePointer = vbDefault 'Add By Sindy 2019/1/21
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
'      strExc(0) = "select eef12 from EmpElectronFile where eef01='" & m_EEP01 & "' and eef02=" & intEEP02 & _
'               " and eef03='" & ChgSQL(pFileName) & "' and eef12 is not null"
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
''Removed by Morgan 2015/5/22 不再存DB
''   strExc(0) = "select * from EmpElectronFile where eef01='" & m_EEP01 & "' and eef02=" & intEEP02 & _
''               " and eef03='" & ChgSQL(pFileName) & "'"
''   intI = 1
''   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''   If intI = 1 Then
''      If Dir(stAttPath) <> "" Then Kill stAttPath
''      With RsTemp
''      lngSize = Val(.Fields("eef04").Value)
''      ReDim bytes(lngSize)
''      If lngSize > 0 Then bytes() = .Fields("eef05").GetChunk(lngSize)
''      End With
''      iFileNo = FreeFile
''      Open stAttPath For Binary Access Write As #iFileNo
''      If lngSize > 0 Then Put #iFileNo, , bytes()
''      Close #iFileNo
''
''      pFileName = stAttPath
''      GetAttachFile = True
''   End If
''   Exit Function
''end 2015/5/22
'
'ErrHnd:
'   MsgBox Err.Description, vbCritical
'   If iFileNo > 0 Then Close #iFileNo
'End Function

'Private Function GetAttachFile_CPP(ByVal strCP09 As String, ByRef pFileName As String, Optional pSavePath As String) As Boolean
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
''      '檔案已存在時不必重新下載
''      If Dir(stAttPath) <> "" Then
''         'Kill stAttPath
''         pFileName = stAttPath
''         GetAttachFile_CPP = True
''         Exit Function
''      End If
'   Else
'      'Add By Sindy 2013/12/27
'      If InStr(pSavePath, m_AttachPath) > 0 Then
'         If Dir(m_AttachPath, vbDirectory) = "" Then
'            MkDir m_AttachPath
'         End If
'      End If
'      '2013/12/27 END
'      stAttPath = pSavePath
'   End If
'
'   strExc(0) = "select * from casepaperpdf where cpp01='" & strCP09 & "' and cpp02='" & ChgSQL(pFileName) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If Dir(stAttPath) <> "" Then Kill stAttPath
'      With RsTemp
'         lngSize = Val(.Fields("cpp03").Value)
'         ReDim bytes(lngSize)
'         If lngSize > 0 Then
'            bytes() = .Fields("cpp04").GetChunk(lngSize)
'         End If
'      End With
'      iFileNo = FreeFile
'      Open stAttPath For Binary Access Write As #iFileNo
'      If lngSize > 0 Then Put #iFileNo, , bytes()
'      Close #iFileNo
'
'      pFileName = stAttPath
'      GetAttachFile_CPP = True
'   End If
'   Exit Function
'
'ErrHnd:
'   If Err.NUMBER = 70 Then
'      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
'   Else
'      MsgBox Err.Description, vbCritical
'   End If
'   If iFileNo > 0 Then Close #iFileNo
'End Function

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   Dim fs, f, s
   
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
               'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
               If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
               '2021/8/6 END
                  stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
               End If
            End If
            
            If InStr(stFileName, "\") = 0 Then
               If Index = 1 Then '存卷資料
                  'If GetAttachFile(stFileName, 0) = False Then Exit Sub
                  If PUB_GetAttachFile_EEF(m_EEP01, 0, stFileName, m_AttachPath) = False Then
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
               Else
                  'If GetAttachFile(stFileName, CInt(m_EEP02)) = False Then Exit Sub
                  If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_EEP02), stFileName, m_AttachPath) = False Then
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
               End If
            End If
            'Add By Sindy 2020/2/11
            If m_EditMode <> 1 Then '非新增,才能設定為唯讀
               SetAttr stFileName, vbReadOnly 'Add By Sindy 2020/1/17 檔案設定為唯讀屬性
            End If
            '2020/2/11 END
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Next ii
      If bolIsSelect = False Then
         MsgBox "請選擇欲開啟的附件！"
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Sub

'全選
Private Sub cmdSelect_Click(Index As Integer)
   Dim ii As Integer, oList As ListBox
   
   Set oList = lstAtt(Index)
   For ii = 0 To oList.ListCount - 1
      lstAtt(Index).Selected(ii) = True
   Next
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
      'Add By Sindy 2022/6/7 路徑字元有萬國碼?不要儲存路徑
      If InStr(strConV(strConV(stFolderPath, vbFromUnicode), vbUnicode), "?") = 0 Then
      '2022/6/7 END
         SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", stFolderPath
      End If
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
                     'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
                     If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
                     '2021/8/6 END
                        stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
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
                           'If GetAttachFile(stFileName, CInt(m_EEP02), stFullName) = False Then
                           If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_EEP02), stFileName, stFullName, True) = False Then
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
         If InStrRev(stFileName, " (") > 0 Then
            'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
            If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
            '2021/8/6 END
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
         End If
         
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
                  'If GetAttachFile(stFileName, 0, stFullName) = False Then
                  If PUB_GetAttachFile_EEF(m_EEP01, 0, stFileName, stFullName, True) = False Then
                     MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                     GoTo RunExit
                  End If
               Else
                  'If GetAttachFile(stFileName, CInt(m_EEP02), stFullName) = False Then
                  If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_EEP02), stFileName, stFullName, True) = False Then
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
   Dim strFile As String
   Dim strFilePath As String 'Add By Sindy 2018/9/26
   Dim bolNotChkFileCaseNo As Boolean
   
On Error GoTo ErrHnd
   
   'Add By Sindy 2018/9/26 取得開啟檔案的路徑
   If lstAtt(Index).ListCount > 0 Then
      ii = 0
      Do While ii < lstAtt(Index).ListCount
         If lstAtt(Index).Selected(ii) = True Then
            If InStr(lstAtt(Index).List(ii), "\") > 0 Then
               strFilePath = Mid(lstAtt(Index).List(ii), 1, InStrRev(lstAtt(Index).List(ii), "\") - 1)
               Exit Do
            End If
         End If
         ii = ii + 1
      Loop
   End If
   If strFilePath = "" Then
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         strFilePath = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         strFilePath = PUB_Getdesktop
      End If
      'Add By Sindy 2022/5/4
      'Modified by Morgan 2022/6/17 修正網路資料夾會錯問題
      'If Dir(strFilePath, vbDirectory) = "" Then
      If PUB_ChkDir(strFilePath) = False Then
      'end 2022/6/17
         strFilePath = PUB_Getdesktop
      End If
      '2022/5/4 END
   End If
   '2018/9/26 END
   
   stFileName = "*.*"
   With CommonDialog1
      .CancelError = True
      'Modify By Sindy 2024/12/12
      '.FileName = stFileName
      '.Filter = "All Files (*.*)|*.*"
      Call GetAddFileKind(CommonDialog1, Index)
      '2024/12/12 END
      'Modify By Sindy 2018/9/26
'      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
'         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
'      Else
'         .InitDir = PUB_Getdesktop
'      End If
      .InitDir = strFilePath
      '2018/9/26 END
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         '選取多個檔案時
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            'Modify By Sindy 2018/9/26 不要儲存系統暫存區路徑
            If InStr(UCase(Trim(sFile(0))), UCase(App.path)) = 0 Then
            '2018/9/26 END
               'Add By Sindy 2022/6/7 路徑字元有萬國碼?不要儲存路徑
               If InStr(strConV(strConV(sFile(0), vbFromUnicode), vbUnicode), "?") = 0 Then
               '2022/6/7 END
                  SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
               End If
            End If
            For ii = 1 To UBound(sFile)
               'Add By Sindy 2013/10/9
               If InStr(CStr(sFile(ii)), "#") > 0 Or InStr(CStr(sFile(ii)), "&") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名"
                  Exit Sub
               End If
               '2013/10/9 END
               
               '檢查檔名規則
''               If Left(CboEEP04.Text, 2) <> EMP_會修 And _
''                  Left(CboEEP04.Text, 2) <> EMP_會完 And _
''                  Left(CboEEP04.Text, 2) <> EMP_聯絡 Then
'                  If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), Left(CboEEP04, 2), CP(10), , Index) = False Then
'                     Exit Sub
'                  End If
''               End If
               'Modify By Sindy 2018/9/25 商標處開放下列檔名可不加本所案號,存檔時系統自動補填
               'Add By Sindy 2021/10/14 ACS案件,不限制電子檔要輸入案號
               bolNotChkFileCaseNo = False
               If NotChkFileCaseNo(CStr(sFile(ii)), Index) = True Or _
                  PField(1) = "ACS" Then
                  bolNotChkFileCaseNo = True
               End If
               '2018/9/25 END
               'Add By Sindy 2018/10/29 多案件，檔名檢查
               'Modify By Sindy 2020/9/29 + And txtLpNote.Tag <> "多案單筆歷程"
               'Modify By Sindy 2021/1/15 + And cmdManyCase.Tag = "確定" : T-228379彈多案但選取消
               'Modify By Sindy 2023/6/21 txtLpNote.Tag <> "多案單筆歷程" => txtLpNote.Tag <> ""
               'Modify By Sindy 2023/6/26 取消 txtLpNote.Tag <> "" And
               If cmdManyCase.Visible = True And cmdManyCase.Enabled = True And _
                  cmdManyCase.Tag = "確定" Then
                  'Modify By Sindy 2015/1/19 智權人員時不控管存卷資料輸入方式
                  'Modify By Sindy 2023/11/22 外專案件不控管存卷資料輸入方式
                  'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True
                  If ManyCaseChkFileName(CStr(sFile(ii)), CStr(Index), _
                     IIf(m_FlowUserNum = Trim(Left(m_SPMan, 6)) Or bolFCPFlow = True Or bolFCTFlow = True, False, True), bolNotChkFileCaseNo) = False Then
                     Exit Sub
                  End If
               Else
               '2018/10/29 END
                  'Modify By Sindy 2015/1/19 智權人員時不控管存卷資料輸入方式
                  'Modify By Sindy 2023/11/22 外專案件不控管存卷資料輸入方式
                  'Modify By Sindy 2024/8/13 + 外商FC同外專不鎖中文,因附件區不進卷宗區
                  If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), Left(CboEEP04, 2), cp(10), , Index, _
                        IIf(m_FlowUserNum = Trim(Left(m_SPMan, 6)) Or bolFCPFlow = True Or bolFCTFlow = True, False, True), , , , _
                        bolNotChkFileCaseNo, _
                        IIf(Index = 0 And (bolFCPFlow = True Or bolFCTFlow = True), True, False)) = False Then
                     Exit Sub
                  End If
               End If
               'Modify By Sindy 2015/1/19 智權人員時不控管存卷資料輸入方式,但一定要是PDF檔
               If Index = 1 And m_FlowUserNum = Trim(Left(m_SPMan, 6)) Then '存卷資料
                  'Modify By Sindy 2024/12/9 FCT同外專案件不控管存卷資料區檔案類型
                  If bolFCTFlow = False Then
                  '2024/12/9 END
                     If UCase(Right(CStr(sFile(ii)), 4)) <> UCase(".pdf") Then
                        MsgBox "存卷資料要是PDF檔！"
                        Exit Sub
                     End If
                  End If
               End If
               
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
               'Add By Sindy 2014/3/11
               ElseIf f.Size > 5242880 Then
                  If Pub_StrUserSt15 = "P13" Or PUB_GetStaffST15(m_FlowUserNum, "1") = "P13" Then
                     If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                        Exit Sub
                     End If
                  End If
               '2014/3/11 END
               End If
               '2013/9/6 END
               
               If ChkCasePDF(CStr(sFile(ii))) = True Then Exit Sub 'Add By Sindy 2024/12/27
               AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)
               If Index = 1 Then Me.cmdSave.Visible = True
            Next
            
         '選取單檔時
         Else
            'stFileName = GetFileName(.FileName)
            'Modify By Sindy 2013/10/9
            'strFile = GetFileName(.FileName)
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Or InStr(strFile, "&") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名"
               Exit Sub
            End If
            '2013/10/9 END
            
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     'Modify By Sindy 2018/9/26 不要儲存系統暫存區路徑
                     If InStr(UCase(Trim(.FileName)), UCase(App.path)) = 0 Then
                     '2018/9/26 END
                        'Add By Sindy 2022/6/7 路徑字元有萬國碼?不要儲存路徑
                        If InStr(strConV(strConV(Mid(Trim(.FileName), 1, ii - 1), vbFromUnicode), vbUnicode), "?") = 0 Then
                        '2022/6/7 END
                           SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                        End If
                        Exit For
                     End If
                  End If
               Next ii
            End If
'            '檢查檔名規則
''            If Left(CboEEP04.Text, 2) <> EMP_會修 And _
''               Left(CboEEP04.Text, 2) <> EMP_會完 And _
''               Left(CboEEP04.Text, 2) <> EMP_聯絡 Then
'               If PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, Left(CboEEP04, 2), CP(10), , Index) = False Then
'                  Exit Sub
'               End If
''            End If
            'Modify By Sindy 2018/9/25 商標處開放下列檔名可不加本所案號,存檔時系統自動補填
            'Add By Sindy 2021/10/14 ACS案件,不限制電子檔要輸入案號
            bolNotChkFileCaseNo = False
            If NotChkFileCaseNo(strFile, Index) = True Or _
               PField(1) = "ACS" Then
               bolNotChkFileCaseNo = True
            End If
            '2018/9/25 END
            'Add By Sindy 2018/10/29 多案件，檔名檢查
            'Modify By Sindy 2020/9/29 + And txtLpNote.Tag <> "多案單筆歷程"
            'Modify By Sindy 2021/1/15 + And cmdManyCase.Tag Modify By Sindy 2023/6/21 txtLpNote.Tag <> "多案單筆歷程"= "確定" : T-228379彈多案但選取消
            'Modify By Sindy 2023/6/21 txtLpNote.Tag <> "多案單筆歷程" => txtLpNote.Tag <> ""
            'Modify By Sindy 2023/6/26 取消 txtLpNote.Tag <> "" And
            If cmdManyCase.Visible = True And cmdManyCase.Enabled = True And _
               cmdManyCase.Tag = "確定" Then
               'Modify By Sindy 2015/1/19 智權人員時不控管存卷資料輸入方式
               'Modify By Sindy 2023/11/22 外專案件不控管存卷資料輸入方式
               'Modify By Sindy 2024/8/13 + Or bolFCTFlow = True
               If ManyCaseChkFileName(strFile, CStr(Index), _
                  IIf(m_FlowUserNum = Trim(Left(m_SPMan, 6)) Or bolFCPFlow = True Or bolFCTFlow = True, False, True), bolNotChkFileCaseNo) = False Then
                  Exit Sub
               End If
            Else
            '2018/10/29 END
               'Modify By Sindy 2015/1/19 智權人員時不控管存卷資料輸入方式
               'Modify By Sindy 2023/11/22 外專案件不控管存卷資料輸入方式
               'Modify By Sindy 2024/8/13 + 外商FC同外專不鎖中文,因附件區不進卷宗區
               If PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, Left(CboEEP04, 2), cp(10), , Index, _
                     IIf(m_FlowUserNum = Trim(Left(m_SPMan, 6)) Or bolFCPFlow = True Or bolFCTFlow = True, False, True), , , , _
                     bolNotChkFileCaseNo, _
                     IIf(Index = 0 And (bolFCPFlow = True Or bolFCTFlow = True), True, False)) = False Then
                  Exit Sub
               End If
            End If
            'Modify By Sindy 2015/1/19 智權人員時不控管存卷資料輸入方式,但一定要是PDF檔
            If Index = 1 And m_FlowUserNum = Trim(Left(m_SPMan, 6)) Then '存卷資料
               'Modify By Sindy 2024/12/9 FCT同外專案件不控管存卷資料區檔案類型
               If bolFCTFlow = False Then
               '2024/12/9 END
                  If UCase(Right(strFile, 4)) <> UCase(".pdf") Then
                     MsgBox "存卷資料要是PDF檔！"
                     Exit Sub
                  End If
               End If
            End If
            
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               Exit Sub
            'Add By Sindy 2014/3/11
            ElseIf f.Size > 5242880 Then
               If Pub_StrUserSt15 = "P13" Or PUB_GetStaffST15(m_FlowUserNum, "1") = "P13" Then
                  If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                     Exit Sub
                  End If
               End If
            '2014/3/11 END
            End If
            '2013/9/6 END
            
            If ChkCasePDF(strFile) = True Then Exit Sub 'Add By Sindy 2024/12/27
            AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)
            If Index = 1 Then Me.cmdSave.Visible = True
         End If
         'Add By Sindy 2018/10/1 移除已不存在的電子檔
         If Index = 0 Then
            If lstAtt(Index).ListCount > 0 Then
               For ii = lstAtt(Index).ListCount - 1 To 0 Step -1
                  strFilePath = lstAtt(Index).List(ii)
                  If InStrRev(strFilePath, " (") > 0 Then
                     'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
                     If UCase(Mid(strFilePath, InStrRev(strFilePath, " (") + 1, Len("(X86)"))) <> "(X86)" Then
                     '2021/8/6 END
                        strFilePath = Left(strFilePath, InStrRev(strFilePath, " (") - 1)
                     End If
                  End If
                  If Dir(strFilePath) = "" Then
                     lstAtt(Index).RemoveItem ii
                  End If
               Next ii
            End If
         End If
         '2018/10/1 END
      End If
      ChDir App.path 'Add By Sindy 2020/1/13 釋放資料夾權限
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2024/12/27
Private Function ChkCasePDF(strFile As String) As Boolean
Dim stReName As String
   
   ChkCasePDF = False
   '欲取得更名後檔名
   Call PUB_GetEmpFlowReNameFile(PField(1), PField(2), PField(3), PField(4), cp(10), strFile, stReName)
   '檢查卷宗區是否有相同檔名
   'Modify By Sindy 2025/4/30 排除已刪除的暫存檔 + and substr(upper(cpp02),-4)<>'.DEL'
   strExc(0) = "select cpp01 from casepaperpdf" & _
               " where cpp01='" & lblCP09.Caption & "' and instr(cpp02,'" & stReName & "')>0 and substr(upper(cpp02),-4)<>'.DEL'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ChkCasePDF = True
      ShowMsg "此檔案：" & strFile & vbCrLf & vbCrLf & "在卷宗區已存在相同檔名（" & stReName & "），請確認！"
      Exit Function
   End If
End Function

'Add By Sindy 2024/12/12 可以新增檔案的類型
Private Function GetAddFileKind(objTmp As Object, Index As Integer) As String
   If Index = 0 Then '附件區
      objTmp.FileName = "*.*"
      objTmp.Filter = "All Files (*.*)|*.*"
   Else '存卷資料區
      If bolFCPFlow = True Or bolFCTFlow = True Then
         objTmp.FileName = "*.*"
         objTmp.Filter = "All Files (*.*)|*.*"
      Else
         objTmp.FileName = "*.pdf"
         objTmp.Filter = "All Files (*.pdf)|*.pdf"
      End If
   End If
End Function

'Add By Sindy 2018/9/25 商標處開放下列檔名可不加本所案號,存檔時系統自動補填
Private Function NotChkFileCaseNo(strFileName As String, Index As Integer) As Boolean
   NotChkFileCaseNo = False
   If bolTMFlow = True And _
      (Left(UCase(strFileName), 8) = "CONTACT." Or _
       Left(UCase(strFileName), 4) = "POA." Or _
       InStr(UCase(strFileName), ".POA.") > 0 Or _
       Left(UCase(strFileName), 4) = "ATT." Or _
       (Index = 1 And Left(UCase(strFileName), 4) = "INFO" And Right(UCase(strFileName), 4) = ".PDF") _
      ) Then
      NotChkFileCaseNo = True
   End If
End Function

'刪除
Private Sub cmdRemAtt_Click(Index As Integer)
Dim bolSel As Boolean
   
   'Add By Sindy 2018/8/8
   bolSel = False
   If lstAtt(Index).ListCount > 0 Then
      ii = 0
      Do While ii < lstAtt(Index).ListCount
         If lstAtt(Index).Selected(ii) = True Then
            bolSel = True
         End If
         ii = ii + 1
      Loop
   End If
   If bolSel = True Then
      If MsgBox("確定要刪除附件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Sub
      End If
      '2018/8/8 END
      If RemoveList(lstAtt(Index), Index) = True Then
         If Index = 1 Then Me.cmdSave.Visible = True
      End If
   End If
End Sub

'Private Function GetSaveName(ByVal pFileName As String, ByVal pFilePath As String) As String
'
'On Error GoTo ErrHnd
'
'   With CommonDialog1
'      .CancelError = True
'      .FileName = pFileName
'      .Filter = "All Files (*.*)|*.*"
'      .InitDir = pFilePath 'PUB_Getdesktop
'      .MaxFileSize = 3000
'      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
'      .ShowSave
'      If .FileName <> "" Then
'         GetSaveName = .FileName
'      End If
'   End With
'
'   Exit Function
'
'ErrHnd:
'   If Err.Number <> 32755 Then
'      MsgBox Err.Description
'   End If
'End Function

Private Function RemoveList(oList As ListBox, Index As Integer) As Boolean
   Dim ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
         
            If oList.ITEMDATA(ii) > 0 And Index = 1 Then '存卷資料才需要記錄移除的檔案
               intI = UBound(m_FilesRemoved) + 1
               ReDim Preserve m_FilesRemoved(intI) As String
               m_FilesRemoved(intI) = GetFileName(oList.List(ii))
            End If
            
            oList.RemoveItem ii
            SetListScroll oList
            RemoveList = True
            ii = ii - 1
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

'Add By Sindy 2018/8/30 E-Mail
'回傳:寄件日期/時間
Private Function EMailKeepFile(ByRef strUpdDate As String, ByRef strUpdTime As String) As Boolean
Dim bolHadFile As Boolean
Dim pbolDone As Boolean
Dim pFiles As String
Dim stFileName As String
Dim intStar As Integer
Dim intEnd As Integer
Dim rsA As New ADODB.Recordset
Dim bolContact As Boolean 'Add By Sindy 2018/10/16
Dim strEEP02 As String 'Modify By Sindy 2022/4/8
   
   EMailKeepFile = False
   '沒點選附件,就預設是全部附件
   'Modify By Sindy 2018/11/15 不管附件的選取狀況,一律全部帶入
'   bolHadFile = False
'   For ii = 0 To lstAtt(0).ListCount - 1
'      If lstAtt(0).Selected(ii) Then
'         bolHadFile = True
'         Exit For
'      End If
'   Next ii
'   If bolHadFile = False Then
      For ii = 0 To lstAtt(0).ListCount - 1
         lstAtt(0).Selected(ii) = True
      Next ii
'   End If
   
   'Add By Sindy 2020/10/21
   If rsA.State <> adStateClosed Then rsA.Close
   strExc(0) = "select * from ExtensionData where ed02='" & strUserNum & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount = 0 Then
      MsgBox "請打字室設定電話分機資料，後續資訊需要使用！", vbExclamation
      Exit Function
   End If
   rsA.Close
   '2020/10/21 END
   
   Screen.MousePointer = vbHourglass
   
   '附件
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum & "\otherFile" 'Add By Sindy 2017/1/6 以防止上面寄信時有些檔案會被咬住,後面刪檔會有權限問題
   'Modify By Sindy 2022/4/8
   'Modified by Morgan 2022/6/17
   'If Dir(m_AttachPath, vbDirectory) = "" Then
   If PUB_ChkDir(m_AttachPath) = False Then
   'end 2022/6/17
      MkDir m_AttachPath
   Else
      KillAttach 'Add By Sindy 2017/3/10
   End If
   '2022/4/8 END
   pFiles = ""
   For ii = 0 To lstAtt(0).ListCount - 1
      If lstAtt(0).Selected(ii) Then
         stFileName = lstAtt(0).List(ii)
         If InStrRev(stFileName, " (") > 0 Then
            'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
            If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
            '2021/8/6 END
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
         End If
         If InStr(stFileName, "\") = 0 Then
            'If GetAttachFile(stFileName, CInt(m_EEP02)) = False Then Exit Function
            If PUB_GetAttachFile_EEF(m_EEP01, CInt(m_EEP02), stFileName, m_AttachPath) = False Then Exit Function
         End If
'         If InStr(UCase(stFileName), ".CONTACT.") > 0 Then '只放一個CONTACT檔
'            If bolContact = False Then
'               pFiles = pFiles & ";" & stFileName
'            End If
'            bolContact = True
'         Else
            pFiles = pFiles & ";" & stFileName
'         End If
      End If
   Next ii
   If pFiles <> "" Then pFiles = Mid(pFiles, 2)
   strErrText = pFiles 'Add By Sindy 2024/5/21
   
   Screen.MousePointer = vbDefault
   
   '******
   '查詢寄件備份且可轉寄(eep14='1'會稿方式=EMail)
   strExc(0) = "select eep02 From empelectronprocess" & _
               " where eep01='" & m_EEP01 & "' and eep02>=" & intLastEEP02 & _
               " and eep04='" & EMP_客戶會稿 & "' and eep14='1'" & _
               " order by eep02 desc"
   rsA.CursorLocation = adUseClient
   rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   strErrText = strErrText & vbCrLf & "Debug:" & strExc(0) 'Add By Sindy 2021/7/21 Find Err
   If rsA.RecordCount > 0 Then
      'Modify By Sindy 2022/4/8
      strEEP02 = "" & rsA.Fields("eep02")
      rsA.Close
      '2022/4/8 END
      strErrText = strErrText & vbCrLf & "Debug:1"  'Add By Sindy 2021/7/21 Find Err
      PUB_ShowMailForm m_EEP01, pFiles, lblCP10, pbolDone, , , , _
         True, strUpdDate, strUpdTime, True, strEEP02 _
         , True, True, _
         , Me, IIf(m_RetrunRecvCnt = 0, 1, m_RetrunRecvCnt), IIf(bolManyCaseToMix = True, m_RetrunRecvToMix, m_RetrunRecv)
      strErrText = strErrText & vbCrLf & "PUB_ShowMailForm(1) => OK"  'Add By Sindy 2022/3/11 找Bug暫放
   Else
      rsA.Close 'Modify By Sindy 2022/4/8
      strErrText = strErrText & vbCrLf & "Debug:2"  'Add By Sindy 2021/7/21 Find Err
      'Modify By Sindy 2019/9/2 + ChangeWStringToTDateString(cp(6)), ChangeWStringToTDateString(cp(7))
      PUB_ShowMailForm m_EEP01, pFiles, lblCP10, pbolDone, , ChangeWStringToTDateString(cp(6)), ChangeWStringToTDateString(cp(7)), _
         True, strUpdDate, strUpdTime, , _
         , True, , txtEEP08.Tag _
         , Me, IIf(m_RetrunRecvCnt = 0, 1, m_RetrunRecvCnt), IIf(bolManyCaseToMix = True, m_RetrunRecvToMix, m_RetrunRecv)
      strErrText = strErrText & vbCrLf & "PUB_ShowMailForm(0) => OK"   'Add By Sindy 2022/3/11 找Bug暫放
   End If
'   rsA.Close
'   strErrText = strErrText & vbCrLf & "rsA.Close => 有問題嗎?" 'Add By Sindy 2022/3/11 找Bug暫放
   '******
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum 'Add By Sindy 2017/1/6 以防止上面寄信時有些檔案會被咬住,後面刪檔會有權限問題
   strErrText = strErrText & vbCrLf & "Debug: m_AttachPath=" & m_AttachPath  'Add By Sindy 2021/7/2
   If pbolDone = True Then '寄信成功
      strErrText = strErrText & vbCrLf & "Debug: pbolDone=True"  'Add By Sindy 2021/7/2
      EMailKeepFile = True
   Else
      strErrText = strErrText & vbCrLf & "Debug: pbolDone=False"  'Add By Sindy 2021/7/2
   End If
   'Screen.MousePointer = vbDefault
   'Add By Sindy 2025/7/9
   Call PUB_WriteDebugLog("【frm090202_2】m_EEP01='" & m_EEP01 & "' pbolDone(寄信是否成功)=" & pbolDone & vbCrLf & _
                      " strErrText=" & strErrText & ";")
   '2025/7/9 END
   
   Set rsA = Nothing
End Function

'Add By Sindy 2018/9/3 送會Mail上要帶出申請人,發明人資料
Private Function GetApplData() As String
Dim k As Integer
Dim strKey As String
Dim strPA26 As String
Dim strPA27 As String
Dim strPA28 As String
Dim strPA29 As String
Dim strPA30 As String
Dim strText As String
Dim strCU10 As String
Dim strCU15 As String
Dim strCU11 As String
Dim strCU04 As String
Dim strCU05 As String
Dim strCU23 As String, strCU112 As String
Dim strCU24 As String, strCU07 As String, strCU103 As String
Dim strPA79 As String, strPA80 As String, strPA82 As String
Dim strPA83 As String, strPA109 As String, strPA110 As String
Dim strPA112 As String, strPA113 As String, strPA115 As String
Dim strPA116 As String, strPA118 As String, strPA119 As String
Dim strPA121 As String, strPA122 As String, strPA124 As String
Dim strPA125 As String, strPA127 As String, strPA128 As String
Dim strPA130 As String, strPA131 As String
Dim strPerson1_C As String, strPerson1_E As String, strPerson2_C As String, strPerson2_E As String
Dim strCWord1 As String, strCWord2 As String, strEWord1 As String, strEWord2 As String 'Add By Sindy 2015/12/22
Dim varTemp As Variant 'Add By Sindy 2015/12/22
Dim strPerType As String 'Add By Sindy 2016/5/4
Dim ii As Integer, strTxt(110) As String, strTmp As String 'Add By Sindy 2018/6/22
Dim strCUX1 As String, strCUX2 As String 'Add By Sindy 2018/6/22
Dim strChaName As String, strEngName As String 'Add By Sindy 2018/6/25
Dim intRow As Integer, kk As Integer
   
   GetApplData = ""
   '讀取基本檔資料
   strExc(0) = "select * from patent where pa01=" & CNULL(PField(1)) & _
               " and pa02=" & CNULL(PField(2)) & _
               " and pa03=" & CNULL(PField(3)) & _
               " and pa04=" & CNULL(PField(4))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      '申請人:
      strPA26 = Trim("" & RsTemp.Fields("PA26"))
      strPA27 = Trim("" & RsTemp.Fields("PA27"))
      strPA28 = Trim("" & RsTemp.Fields("PA28"))
      strPA29 = Trim("" & RsTemp.Fields("PA29"))
      strPA30 = Trim("" & RsTemp.Fields("PA30"))
      '代表人:
      strPA79 = Trim("" & RsTemp.Fields("PA79"))
      strPA80 = Trim("" & RsTemp.Fields("PA80"))
      strPA82 = Trim("" & RsTemp.Fields("PA82"))
      strPA83 = Trim("" & RsTemp.Fields("PA83"))
      strPA109 = Trim("" & RsTemp.Fields("PA109"))
      strPA110 = Trim("" & RsTemp.Fields("PA110"))
      strPA112 = Trim("" & RsTemp.Fields("PA112"))
      strPA113 = Trim("" & RsTemp.Fields("PA113"))
      strPA115 = Trim("" & RsTemp.Fields("PA115"))
      strPA116 = Trim("" & RsTemp.Fields("PA116"))
      strPA118 = Trim("" & RsTemp.Fields("PA118"))
      strPA119 = Trim("" & RsTemp.Fields("PA119"))
      strPA121 = Trim("" & RsTemp.Fields("PA121"))
      strPA122 = Trim("" & RsTemp.Fields("PA122"))
      strPA124 = Trim("" & RsTemp.Fields("PA124"))
      strPA125 = Trim("" & RsTemp.Fields("PA125"))
      strPA127 = Trim("" & RsTemp.Fields("PA127"))
      strPA128 = Trim("" & RsTemp.Fields("PA128"))
      strPA130 = Trim("" & RsTemp.Fields("PA130"))
      strPA131 = Trim("" & RsTemp.Fields("PA131"))
   End If
   
   For k = 1 To 5
      strKey = ""
      strCU10 = "": strCU15 = "": strCU11 = "": strCU04 = "": strCU05 = ""
      strCU23 = "": strCU24 = "": strCU112 = ""
      strPerson1_C = "": strPerson1_E = "": strPerson2_C = "": strPerson2_E = ""
      If k = 1 Then
         strKey = strPA26
         strPerson1_C = strPA79
         strPerson1_E = strPA80
         strPerson2_C = strPA82
         strPerson2_E = strPA83
      ElseIf k = 2 Then
         strKey = strPA27
         strPerson1_C = strPA109
         strPerson1_E = strPA110
         strPerson2_C = strPA112
         strPerson2_E = strPA113
      ElseIf k = 3 Then
         strKey = strPA28
         strPerson1_C = strPA115
         strPerson1_E = strPA116
         strPerson2_C = strPA118
         strPerson2_E = strPA119
      ElseIf k = 4 Then
         strKey = strPA29
         strPerson1_C = strPA121
         strPerson1_E = strPA122
         strPerson2_C = strPA124
         strPerson2_E = strPA125
      ElseIf k = 5 Then
         strKey = strPA30
         strPerson1_C = strPA127
         strPerson1_E = strPA128
         strPerson2_C = strPA130
         strPerson2_E = strPA131
      End If
      If strKey <> "" Then
         strExc(0) = "select cu10,cu15,cu11,cu04,decode(cu05,null,'',nvl(cu05,'')||' '||nvl(cu88,'')||' '||nvl(cu89,'')||' '||nvl(cu90,'')) as cu05,cu16,cu17,cu18,cu19" & _
                     ",cu07,cu103,cu23" & _
                     ",cu39,cu40,cu41,cu42,cu43,cu44,cu45,cu46,cu47,cu48,cu49,cu50" & _
                     ",cu51,cu52,cu53,cu54,cu55,cu56" & _
                     ",decode(cu24,null,'',nvl(cu24,'')||' '||nvl(cu25,'')||' '||nvl(cu26,'')||' '||nvl(cu27,'')||' '||nvl(cu28,'')) as cu24,cu112" & _
                     ",N1.NA72 X1,N2.NA72 X2" & _
                     " from customer,NATION N1,NATION N2 where cu01='" & Left(strKey, 8) & "'" & _
                     " and cu02='" & Mid(strKey, 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strCU10 = Trim("" & RsTemp.Fields("cu10"))
            strCU15 = Trim("" & RsTemp.Fields("cu15"))
            strCU11 = Trim("" & RsTemp.Fields("cu11"))
            strCU04 = Trim("" & RsTemp.Fields("cu04")) 'ChgSQL(
            strCU05 = Trim("" & RsTemp.Fields("cu05")) 'ChgSQL(
            strCU23 = Trim("" & RsTemp.Fields("cu23")) 'ChgSQL(
            strCU24 = Trim("" & RsTemp.Fields("cu24")) 'ChgSQL(
            strCU112 = Trim("" & RsTemp.Fields("cu112"))
            strCUX1 = Trim("" & RsTemp.Fields("X1"))
            strCUX2 = Trim("" & RsTemp.Fields("X2"))
            strCU07 = Trim("" & RsTemp.Fields("cu07")) 'Add By Sindy 2019/4/12 公司負責人
            strCU103 = Trim("" & RsTemp.Fields("cu103")) 'Add By Sindy 2019/4/12 公司英文負責人
         End If
         
         GetApplData = GetApplData & "申請人" & k & "：" & vbCrLf
         If strCU15 = "0" Then
            GetApplData = GetApplData & "　　中文姓名："
         Else
            GetApplData = GetApplData & "　　中文名稱："
         End If
         '修法:106/12/01開始中文名稱要加外商國名
         If Val(strSrvDate(2)) >= 1061201 And strCU15 = "1" Then '1.公司
            GetApplData = GetApplData & GetPrjNationName(strCU10, "NA81", pa(1)) & strCU04 & vbCrLf
         Else
            GetApplData = GetApplData & strCU04 & vbCrLf
         End If
'         If strCU15 = "0" Then
'            GetApplData = GetApplData & "　　英文姓名："
'         Else
'            GetApplData = GetApplData & "　　英文名稱："
'         End If
'         If strCU05 <> "" Then
'            GetApplData = GetApplData & strCU05 & vbCrLf
'         Else
'            GetApplData = GetApplData & vbCrLf
'         End If
         If strCU10 < "011" Then
            If strCU15 = "0" And "" & strCU11 = "" Then '個人無ID時也要顯示標題
               GetApplData = GetApplData & "　　ID：" & vbCrLf
            Else
               GetApplData = GetApplData & "　　ID：" & strCU11 & vbCrLf
            End If
         End If
         GetApplData = GetApplData & "　　中文地址：" & PUB_ChgNumeralStyle(pa(30 + k)) & vbCrLf 'ChgSQL(
'         GetApplData = GetApplData & "　　英文地址：" & pa(35 + k) & vbCrLf 'ChgSQL(
         
         strChaName = "": strEngName = ""
         If strCU15 <> "0" Then '非自然人才要帶出代表人資料
            'Add By Sindy 2019/4/12
            '公司負責人
            If strCU07 <> "" Then
               If Len(strCU07) = 3 Then
                  strChaName = PUB_ConvertNameFormat(strCU07)
               Else
                  strChaName = strCU07
               End If
            End If
            strEngName = strCU103 'Add By Sindy 2019/4/12 公司英文負責人
            
'            'Modify By Sindy 2019/3/12 代表人:P案要抓客戶檔
'            intRow = 0
'            For kk = 1 To 6
'               intRow = intRow + 1
'               '代表人中文姓名-->非自然人時為必要欄位
'               strTmp = "" & RsTemp("CU" & CStr(39 + 3 * (kk - 1)))
'               If strTmp <> "" Then
'                  If Len(strTmp) = 3 Then strTmp = PUB_ConvertNameFormat(strTmp)
'                  strChaName = strChaName & " " & intRow & "." & strTmp
'               Else
'                  'Modify By Sindy 2018/1/17 只有一個代表人時不要有1.
'                  If strChaName <> "" Then
'                     strChaName = Replace(strChaName, "1.", "")
'                  End If
'                  '2018/1/17 END
'               End If
'               '代表人英文姓名-->非必要欄位
'               strTmp = "" & RsTemp("CU" & CStr(40 + 3 * (kk - 1)))
'               If strTmp <> "" Then
'                  strEngName = strEngName & " " & intRow & "." & strTmp
'               Else
'                  'Modify By Sindy 2018/1/17 只有一個代表人時不要有1.
'                  If strEngName <> "" Then
'                     strEngName = Replace(strEngName, "1.", "")
'                  End If
'                  '2018/1/17 END
'               End If
'            Next kk
''            strChaName = strChaName & " " & IIf(strPerson2_C <> "", "1.", "") & strPerson1_C
''            If strPerson2_C <> "" Then
''               strChaName = strChaName & " 2." & strPerson2_C
''            End If
''            strChaName = Trim(strChaName)
''            strEngName = strEngName & " " & IIf(strPerson2_E <> "", "1.", "") & strPerson1_E
''            If strPerson2_E <> "" Then
''               strEngName = strEngName & " 2." & strPerson2_E
''            End If
''            strEngName = Trim(strEngName)
'            '2019/3/12 END
            
            '代表人中文姓名
            GetApplData = GetApplData & "　　代表人中文姓名：" & strChaName & vbCrLf 'ChgSQL(
'            '代表人英文姓名
'            If strEngName <> "" Then
'               GetApplData = GetApplData & "　　代表人英文姓名：" & strEngName & vbCrLf 'ChgSQL(
'            End If
         End If
      End If
   Next k
   
   '發明人資料
   strExc(0) = " SELECT IN03,IN04,IN05,IN11,NA72" & _
               " FROM PatentInventor,INVENTOR,NATION" & _
               " WHERE pi01=" + CNULL(PField(1)) + " and pi02=" + CNULL(PField(2)) + " and pi03=" + CNULL(PField(3)) + " and pi04=" + CNULL(PField(4)) & _
               " AND IN01=substr(pi06,1,8) AND IN02=substr(pi06,9,2)" & _
               " AND NA01(+)=IN11" & _
               " order by pi05 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      k = 1
      Do While Not RsTemp.EOF
         GetApplData = GetApplData & "發明人" & k & "：" & vbCrLf & _
                                     "　　中文姓名：" & ChgSQL("" & RsTemp("IN04")) & vbCrLf
'         GetApplData = GetApplData & IIf("" & RsTemp("IN05") = "", "", "　　英文姓名：" & "" & RsTemp("IN05") & vbCrLf)  'ChgSQL(
         If "" & RsTemp("IN03") <> "" Then
            GetApplData = GetApplData & "　　ID：" & RsTemp("IN03") & vbCrLf
         Else
            If "" & RsTemp("IN11") < "010" Then '台灣
               GetApplData = GetApplData & "　　ID：" & RsTemp("IN03") & vbCrLf
            End If
         End If
         k = k + 1
         RsTemp.MoveNext
      Loop
   Else
      GetApplData = GetApplData & "發明人1：" & vbCrLf & _
                    "　　姓名：" & vbCrLf & _
                    "　　ID：" & vbCrLf
   End If
End Function

'Add By Sindy 2025/7/9 依歷程狀態彈提醒訊息
Private Sub ShowRemindMsg()
Dim strMsg As String
   
   strMsg = ""
   '103=設計申請
   '105=集體設計
   '125=衍生設計申請
   If PField(1) = "CFP" _
      And (cp(10) = "103" Or cp(10) = "105" Or cp(10) = "125") _
      And (Left(CboEEP04.Text, 2) = EMP_送判 Or Left(CboEEP04.Text, 2) = EMP_判發) Then
      Select Case m_Country
         Case "101" '美國
            strMsg = "請確認是否說明書及圖式均符合TE暫存區的筆記的該國規定" & vbCrLf & _
                     "尤其請注意" & vbCrLf & _
                     "1.虛線、斷線之說明，每種線的用途要分別寫出，同種線同時代表多個意思(如不主張及環境物)，也要分別指出。" & vbCrLf & _
                     "2.確認有無死角未被繪示於兩張以上視圖。特別是內凹、底面、或其他會被遮住的部位。"
         Case "012" '韓國
            strMsg = "請確認是否說明書及圖式均符合TE暫存區的筆記的該國規定" & vbCrLf & _
                     "尤其請注意" & vbCrLf & _
                     "圖式一定要有底視圖。"
         Case "239" '歐盟
            strMsg = "請確認是否說明書及圖式均符合TE暫存區的筆記的該國規定" & vbCrLf & _
                     "尤其請注意" & vbCrLf & _
                     "1.名稱請完全依照羅卡諾分類來取。" & vbCrLf & _
                     "2.選擇分類，請多方確認以使分類與產品最相符。"
         Case "042" '越南
            strMsg = "請確認是否說明書及圖式均符合TE暫存區的筆記的該國規定" & vbCrLf & _
                     "尤其請注意" & vbCrLf & _
                     "圖式的線形之粗細要一致。"
         Case "040" '印度
            strMsg = "請確認是否說明書及圖式均符合TE暫存區的筆記的該國規定" & vbCrLf & _
                     "尤其請注意" & vbCrLf & _
                     "若有主張優先權，圖式需與優先權完全相同，即使虛實線轉換也不行。"
         Case Else '其他國家
            strMsg = "請確認是否說明書及圖式均符合TE暫存區的筆記的該國規定"
      End Select
   End If
   If strMsg <> "" Then
      MsgBox strMsg, vbInformation
   End If
End Sub
