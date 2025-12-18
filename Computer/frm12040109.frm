VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040109 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件性質對照檔"
   ClientHeight    =   6010
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   7710
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6010
   ScaleWidth      =   7710
   Begin TabDlg.SSTab stb 
      Height          =   4845
      Left            =   45
      TabIndex        =   43
      Top             =   1110
      Width           =   7605
      _ExtentX        =   13406
      _ExtentY        =   8537
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "案件性質"
      TabPicture(0)   =   "frm12040109.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(9)=   "Label19"
      Tab(0).Control(10)=   "Label20"
      Tab(0).Control(11)=   "Label21"
      Tab(0).Control(12)=   "Label22"
      Tab(0).Control(13)=   "Label23"
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(15)=   "Label14(6)"
      Tab(0).Control(16)=   "Label14(7)"
      Tab(0).Control(17)=   "Label24"
      Tab(0).Control(18)=   "Label25"
      Tab(0).Control(19)=   "Label35"
      Tab(0).Control(20)=   "Label14(4)"
      Tab(0).Control(21)=   "textCPM03"
      Tab(0).Control(22)=   "textCPM04"
      Tab(0).Control(23)=   "textCPM10"
      Tab(0).Control(24)=   "textCPM13"
      Tab(0).Control(25)=   "Label26"
      Tab(0).Control(26)=   "textCPM12_2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCPM11_2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCPM12"
      Tab(0).Control(29)=   "textCPM11"
      Tab(0).Control(30)=   "textCPM09"
      Tab(0).Control(31)=   "textCPM08"
      Tab(0).Control(32)=   "textCPM07"
      Tab(0).Control(33)=   "TextCPM16"
      Tab(0).Control(34)=   "TextCPM17"
      Tab(0).Control(35)=   "TextCPM18"
      Tab(0).Control(36)=   "textCPM19"
      Tab(0).Control(37)=   "textCPM20"
      Tab(0).Control(38)=   "textCPM21"
      Tab(0).Control(39)=   "textCPM22"
      Tab(0).Control(40)=   "textCPM24"
      Tab(0).Control(41)=   "textCPM25"
      Tab(0).Control(42)=   "textCPM24_2"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCPM25_2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textCPM33"
      Tab(0).Control(45)=   "textCPM34"
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "員工群組"
      TabPicture(1)   =   "frm12040109.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvw"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "專業部"
      TabPicture(2)   =   "frm12040109.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label14(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label6"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label14(1)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label14(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label14(3)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label30"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label31"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label32"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label33"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label34"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label5"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label27"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label28"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label14(5)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label14(8)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label14(9)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "LblCPM23Code"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "textCPM26"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "textCPM23"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "textCPM15"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "textCPM14"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "textCPM06"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "textCPM27"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "textCPM28"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "textCPM29"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "textCPM31"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "textCPM32"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "textCPM35"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "textCPM36"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "textCPM30"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "textCPM05"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).ControlCount=   34
      Begin VB.TextBox textCPM05 
         Height          =   270
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   84
         Top             =   408
         Width           =   315
      End
      Begin VB.TextBox textCPM30 
         Height          =   270
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   23
         Top             =   420
         Width           =   315
      End
      Begin VB.TextBox textCPM36 
         Height          =   270
         Left            =   5700
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   32
         Top             =   3210
         Width           =   315
      End
      Begin VB.TextBox textCPM35 
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   35
         Top             =   4260
         Width           =   732
      End
      Begin VB.TextBox textCPM34 
         Height          =   270
         Left            =   -72510
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   21
         Top             =   4260
         Width           =   315
      End
      Begin VB.TextBox textCPM33 
         Height          =   270
         Left            =   -68910
         MaxLength       =   2
         TabIndex        =   19
         Top             =   3750
         Width           =   315
      End
      Begin VB.TextBox textCPM32 
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2370
         MaxLength       =   4
         TabIndex        =   34
         Top             =   3924
         Width           =   732
      End
      Begin VB.TextBox textCPM31 
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1770
         MaxLength       =   4
         TabIndex        =   33
         Top             =   3600
         Width           =   732
      End
      Begin VB.TextBox textCPM29 
         Height          =   270
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   31
         Top             =   3210
         Width           =   315
      End
      Begin VB.TextBox textCPM28 
         Height          =   270
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   30
         Top             =   2880
         Width           =   315
      End
      Begin VB.TextBox textCPM27 
         Height          =   270
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   29
         Top             =   2550
         Width           =   315
      End
      Begin VB.TextBox textCPM06 
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1530
         MaxLength       =   4
         TabIndex        =   24
         Top             =   720
         Width           =   732
      End
      Begin VB.TextBox textCPM14 
         Height          =   270
         Left            =   1530
         MaxLength       =   4
         TabIndex        =   25
         Top             =   1020
         Width           =   732
      End
      Begin VB.TextBox textCPM15 
         Height          =   270
         Left            =   1530
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1320
         Width           =   732
      End
      Begin VB.TextBox textCPM23 
         Height          =   270
         Left            =   1650
         MaxLength       =   1
         TabIndex        =   27
         Top             =   1620
         Width           =   315
      End
      Begin VB.TextBox textCPM26 
         Height          =   270
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   28
         Top             =   2220
         Width           =   690
      End
      Begin VB.TextBox textCPM25_2 
         BorderStyle     =   0  '沒有框線
         Height          =   270
         Left            =   -69150
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2370
         Width           =   1695
      End
      Begin VB.TextBox textCPM24_2 
         BorderStyle     =   0  '沒有框線
         Height          =   270
         Left            =   -69150
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1695
      End
      Begin VB.TextBox textCPM25 
         Height          =   270
         Left            =   -69915
         MaxLength       =   6
         TabIndex        =   12
         Top             =   2370
         Width           =   765
      End
      Begin VB.TextBox textCPM24 
         Height          =   270
         Left            =   -69915
         MaxLength       =   6
         TabIndex        =   10
         Top             =   2070
         Width           =   765
      End
      Begin VB.TextBox textCPM22 
         Height          =   270
         Left            =   -72510
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   20
         Top             =   3990
         Width           =   315
      End
      Begin VB.TextBox textCPM21 
         Height          =   270
         Left            =   -72510
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   18
         Top             =   3720
         Width           =   315
      End
      Begin VB.TextBox textCPM20 
         Height          =   270
         Left            =   -72510
         MaxLength       =   1
         TabIndex        =   13
         Top             =   2910
         Width           =   315
      End
      Begin VB.TextBox textCPM19 
         Height          =   270
         Left            =   -69630
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3180
         Width           =   315
      End
      Begin VB.TextBox TextCPM18 
         Height          =   270
         Left            =   -69915
         MaxLength       =   14
         TabIndex        =   14
         Top             =   2910
         Width           =   900
      End
      Begin VB.TextBox TextCPM17 
         Height          =   270
         Left            =   -72510
         MaxLength       =   1
         TabIndex        =   17
         Top             =   3450
         Width           =   315
      End
      Begin VB.TextBox TextCPM16 
         Height          =   270
         Left            =   -72510
         MaxLength       =   1
         TabIndex        =   15
         Top             =   3180
         Width           =   315
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   2955
         Left            =   -74850
         TabIndex        =   22
         Top             =   570
         Width           =   7395
         _ExtentX        =   13053
         _ExtentY        =   5203
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox textCPM07 
         Height          =   270
         Left            =   -73920
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1710
         Width           =   315
      End
      Begin VB.TextBox textCPM08 
         Height          =   270
         Left            =   -70275
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1710
         Width           =   315
      End
      Begin VB.TextBox textCPM09 
         Height          =   270
         Left            =   -68310
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1710
         Width           =   315
      End
      Begin VB.TextBox textCPM11 
         Height          =   270
         Left            =   -73965
         MaxLength       =   6
         TabIndex        =   9
         Top             =   2070
         Width           =   765
      End
      Begin VB.TextBox textCPM12 
         Height          =   270
         Left            =   -73965
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2370
         Width           =   765
      End
      Begin VB.TextBox textCPM11_2 
         BorderStyle     =   0  '沒有框線
         Height          =   270
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1695
      End
      Begin VB.TextBox textCPM12_2 
         BorderStyle     =   0  '沒有框線
         Height          =   270
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2370
         Width           =   1695
      End
      Begin VB.Label LblCPM23Code 
         Caption         =   "LblCPM23Code"
         ForeColor       =   &H00FF0000&
         Height          =   250
         Left            =   1650
         TabIndex        =   86
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數：        ( Y:算 N:不算 Null: 未預設)"
         Height          =   180
         Index           =   9
         Left            =   3552
         TabIndex        =   85
         Top             =   480
         Width           =   3720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "接洽單是否顯示：        ( N:不顯示)"
         Height          =   180
         Index           =   8
         Left            =   180
         TabIndex        =   83
         Top             =   480
         Width           =   2720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否不電子簽核-寰華案：        ( N:不電子簽核)"
         Height          =   180
         Index           =   5
         Left            =   3660
         TabIndex        =   82
         Top             =   3270
         Width           =   3680
      End
      Begin VB.Label Label28 
         Caption         =   "(1.先補文件再呈分案主管 2.程序承辦不需經由主管分案 3.可能程序或工程師承辦)"
         ForeColor       =   &H00FF0000&
         Height          =   380
         Left            =   2280
         TabIndex        =   81
         Top             =   4260
         Width           =   4310
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "內專分案狀況："
         Height          =   180
         Left            =   180
         TabIndex        =   80
         Top             =   4320
         Width           =   1260
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "(規費科目220113表示要扣律師庭費，真正規費固定用2403-代收代付款)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74850
         TabIndex        =   79
         Top             =   2700
         Width           =   5580
      End
      Begin MSForms.TextBox textCPM13 
         Height          =   300
         Left            =   -73230
         TabIndex        =   5
         Top             =   1350
         Width           =   5745
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "10134;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCPM10 
         Height          =   300
         Left            =   -73230
         TabIndex        =   4
         Top             =   1050
         Width           =   5745
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "10134;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCPM04 
         Height          =   300
         Left            =   -73230
         TabIndex        =   3
         Top             =   750
         Width           =   5745
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "10134;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCPM03 
         Height          =   300
         Left            =   -73230
         TabIndex        =   2
         Top             =   450
         Width           =   5745
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "10134;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   " ( 此欄為基礎記件值，實際會於Trigger內依案件屬性及承辦人組別計算)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   2310
         TabIndex        =   78
         Top             =   780
         Width           =   5085
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "主管機關期限：         (Y:是 N:不是 Null.未設定) "
         Height          =   180
         Index           =   4
         Left            =   -73800
         TabIndex        =   77
         Top             =   4290
         Width           =   3795
      End
      Begin VB.Label Label35 
         Caption         =   "發文未請款管制天數：       (工作天)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -70680
         TabIndex        =   76
         Top             =   3780
         Width           =   3120
      End
      Begin VB.Label Label34 
         Caption         =   "註：P案C類來函請設定為不會稿 !"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4470
         TabIndex        =   75
         Top             =   2910
         Width           =   2775
      End
      Begin VB.Label Label33 
         Caption         =   "註：新增性質時，要設定核判表"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4590
         TabIndex        =   74
         Top             =   1680
         Width           =   2565
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "FCP程序考核點數-寰華案："
         Height          =   180
         Left            =   180
         TabIndex        =   73
         Top             =   3975
         Width           =   2160
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "FCP程序考核點數："
         Height          =   180
         Left            =   180
         TabIndex        =   72
         Top             =   3645
         Width           =   1560
      End
      Begin VB.Label Label30 
         Caption         =   "(不會稿案件核稿人可否直接判發)"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   71
         Top             =   2610
         Width           =   2775
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否不電子簽核：        ( N:不電子簽核)"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   70
         Top             =   3270
         Width           =   3030
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否會稿：            (Y:會稿 N:不會稿 Null.未預設) "
         Height          =   180
         Index           =   2
         Left            =   540
         TabIndex        =   69
         Top             =   2910
         Width           =   3810
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "核稿人不可判發：        ( N:不可判發)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   68
         Top             =   2610
         Width           =   2850
      End
      Begin VB.Label Label6 
         Caption         =   "承辦人計件值："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   180
         TabIndex        =   67
         Top             =   765
         Width           =   1320
      End
      Begin VB.Label Label17 
         Caption         =   "繪圖計件值 草："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   180
         TabIndex        =   66
         Top             =   1065
         Width           =   1305
      End
      Begin VB.Label Label18 
         Caption         =   "繪圖計件值 墨："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   180
         TabIndex        =   65
         Top             =   1365
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "專業部核判分類：           (9為不列入核判表之案件性質)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   180
         TabIndex        =   64
         Top             =   1680
         Width           =   4305
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "專業部電子檔副檔名："
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   63
         Top             =   2250
         Width           =   1800
      End
      Begin VB.Label Label25 
         Caption         =   "大陸案規費科目："
         Height          =   180
         Left            =   -71355
         TabIndex        =   62
         Top             =   2415
         Width           =   1440
      End
      Begin VB.Label Label24 
         Caption         =   "大陸案收入科目："
         Height          =   180
         Left            =   -71355
         TabIndex        =   61
         Top             =   2115
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否經發文室-非主管機關：         ( Y:是 )"
         Height          =   180
         Index           =   7
         Left            =   -74760
         TabIndex        =   60
         Top             =   4050
         Width           =   3180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否經發文室-主管機關：         ( Y:是 Q:詢問 )"
         Height          =   180
         Index           =   6
         Left            =   -74580
         TabIndex        =   59
         Top             =   3780
         Width           =   3570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "來函期限是否通知智權人員：         (Y:是)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74865
         TabIndex        =   58
         Top             =   2955
         Width           =   3210
      End
      Begin VB.Label Label23 
         Caption         =   "發文未請款管制月數：       月"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -71400
         TabIndex        =   57
         Top             =   3210
         Width           =   2340
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "PS：新增案件性質時，紅色欄位及專利處頁籤，請先詢問專業部!!"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   -74910
         TabIndex        =   56
         Top             =   4530
         Width           =   5160
      End
      Begin VB.Label Label21 
         Caption         =   "來函預設請款金額"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -71400
         TabIndex        =   55
         Top             =   2955
         Width           =   1500
      End
      Begin VB.Label Label20 
         Caption         =   "是否向客戶請款：         ( N:不請  註:目前 FMP, FCP 用)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -73980
         TabIndex        =   54
         Top             =   3495
         Width           =   4305
      End
      Begin VB.Label Label19 
         Caption         =   "是否計算工作時數：         (Y:是)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74160
         TabIndex        =   53
         Top             =   3210
         Width           =   2640
      End
      Begin VB.Label Label3 
         Caption         =   "國內案件性質名稱："
         Height          =   180
         Left            =   -74865
         TabIndex        =   52
         Top             =   510
         Width           =   1620
      End
      Begin VB.Label Label4 
         Caption         =   "大陸案件性質名稱："
         Height          =   180
         Left            =   -74865
         TabIndex        =   51
         Top             =   810
         Width           =   1620
      End
      Begin VB.Label Label7 
         Caption         =   "來函期限：         (1:文到當日 2:文到次日)"
         Height          =   180
         Left            =   -74865
         TabIndex        =   50
         Top             =   1755
         Width           =   3180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "期限天數：         天"
         Height          =   180
         Left            =   -71175
         TabIndex        =   49
         Top             =   1755
         Width           =   1485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "期限月數：         月"
         Height          =   180
         Left            =   -69225
         TabIndex        =   48
         Top             =   1755
         Width           =   1485
      End
      Begin VB.Label Label10 
         Caption         =   "英文案件性質名稱："
         Height          =   180
         Left            =   -74865
         TabIndex        =   47
         Top             =   1110
         Width           =   1620
      End
      Begin VB.Label Label12 
         Caption         =   "收入科目："
         Height          =   180
         Left            =   -74865
         TabIndex        =   46
         Top             =   2115
         Width           =   900
      End
      Begin VB.Label Label13 
         Caption         =   "規費科目："
         Height          =   180
         Left            =   -74865
         TabIndex        =   45
         Top             =   2415
         Width           =   900
      End
      Begin VB.Label Label16 
         Caption         =   "日文案件性質名稱："
         Height          =   180
         Left            =   -74865
         TabIndex        =   44
         Top             =   1410
         Width           =   1620
      End
   End
   Begin VB.TextBox textCPM02 
      Height          =   270
      Left            =   5460
      MaxLength       =   4
      TabIndex        =   1
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox textCPM01 
      Height          =   270
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   732
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   480
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":0054
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":0370
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":068C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":0868
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":0B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":0EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":11BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":14D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":17F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":1B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040109.frx":1E2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
   Begin VB.Label Label2 
      Caption         =   "案件性質代號："
      Height          =   252
      Left            =   4080
      TabIndex        =   41
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   255
      Left            =   480
      TabIndex        =   40
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frm12040109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/15 改成Form2.0 ; textCPM03、textCPM04、textCPM10、textCPM13
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'Modify by Morgan 2008/9/24
'原浮動準備金欄位改為會稿加乘適用規則
Option Explicit

Dim MAX_FIELD As Integer
Dim m_FieldList() As FIELDITEM
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
' 辦識其為外商還是內商的程式
' 0 表內商
' 1 表外商
Dim m_SysKind As Integer
' 第一筆資料的本所案號
Dim m_FirstCM(2) As String
' 最後一筆資料的本所案號
Dim m_LastCM(2) As String
' 目前正在顯示的本所案號
Dim m_CurrCM(2) As String
' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean


Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT CPM01,CPM02 FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = (SELECT MIN(CPM01) FROM CASEPROPERTYMAP ) AND " & _
                  "CPM02 = (SELECT MIN(CPM02) FROM CASEPROPERTYMAP " & _
                           "WHERE CPM01 = (SELECT MIN(CPM01) FROM CASEPROPERTYMAP)) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CPM01")) = False Then: m_FirstCM(0) = rsTmp.Fields("CPM01")
      If IsNull(rsTmp.Fields("CPM02")) = False Then: m_FirstCM(1) = rsTmp.Fields("CPM02")
   End If
   rsTmp.Close

   strSql = "SELECT CPM01,CPM02 FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = (SELECT MAX(CPM01) FROM CASEPROPERTYMAP ) AND " & _
                  "CPM02 = (SELECT MAX(CPM02) FROM CASEPROPERTYMAP " & _
                           "WHERE CPM01 = (SELECT MAX(CPM01) FROM CASEPROPERTYMAP)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CPM01")) = False Then: m_LastCM(0) = rsTmp.Fields("CPM01")
      If IsNull(rsTmp.Fields("CPM02")) = False Then: m_LastCM(1) = rsTmp.Fields("CPM02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Initialize()
   strExc(0) = "select * from CASEPROPERTYMAP where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   MAX_FIELD = RsTemp.Fields.Count
   ReDim m_FieldList(MAX_FIELD) As FIELDITEM
End Sub

' Load Form
Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040109", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040109", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040109", strDel, False)
   'edit by nickc 2006/03/09 應該是當初 copy 過來沒修正的
   'm_bQuery = IsUserHasRightOfFunction("frm020501", strFind, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040109", strFind, False)
   
   textCPM11_2.BackColor = &H8000000F
   textCPM12_2.BackColor = &H8000000F
   'Add by Morgan 2010/4/26
   textCPM24_2.BackColor = &H8000000F
   textCPM25_2.BackColor = &H8000000F
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   InitialField
   
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   'Add By Cheng 2002/11/21
   '顯示可使用的員工群組
   Me.lvw.ListItems.Clear
   ShowStaffGroup
   UpdateGroupData textCPM01.Text, textCPM02.Text
   Me.stb.Tab = 0 'Add By Sindy 2013/9/30
   LblCPM23Code.Caption = "" 'Add By Sindy 2025/7/31
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CPM" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 6, 8, 9, 14, 15:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   SetFieldNewData "CPM01", textCPM01
   SetFieldNewData "CPM02", textCPM02
   SetFieldNewData "CPM03", textCPM03
   SetFieldNewData "CPM04", textCPM04
   SetFieldNewData "CPM05", textCPM05 'Modified by Lydiad 2025/06/18 改成「是否算案件數」
   SetFieldNewData "CPM06", textCPM06
   SetFieldNewData "CPM07", textCPM07
   SetFieldNewData "CPM08", textCPM08
   SetFieldNewData "CPM09", textCPM09
   SetFieldNewData "CPM10", textCPM10
   SetFieldNewData "CPM11", textCPM11
   SetFieldNewData "CPM12", textCPM12
   SetFieldNewData "CPM13", textCPM13
   'add by nickc 2006/03/09
   SetFieldNewData "CPM14", textCPM14
   SetFieldNewData "CPM15", textCPM15
   'Add by Morgan 2007/7/19
   SetFieldNewData "CPM16", TextCPM16
   SetFieldNewData "CPM17", TextCPM17
   SetFieldNewData "CPM18", TextCPM18
   'Add by Morgan 2008/7/9
   SetFieldNewData "CPM19", textCPM19
   'Add by Morgan 2008/9/24
   SetFieldNewData "CPM20", textCPM20
   'Add by Morgan 2009/3/18
   SetFieldNewData "CPM21", textCPM21
   SetFieldNewData "CPM22", textCPM22
   SetFieldNewData "CPM23", textCPM23  '2009/4/9 ADD BY SONIA
   'Add by Morgan 2010/4/26
   SetFieldNewData "CPM24", textCPM24
   SetFieldNewData "CPM25", textCPM25
   'Add by Sindy 2013/8/7
   SetFieldNewData "CPM26", textCPM26
   'Add by Sindy 2013/10/1
   SetFieldNewData "CPM27", textCPM27
   SetFieldNewData "CPM28", textCPM28
   SetFieldNewData "CPM29", textCPM29
   '2013/10/1 END
   SetFieldNewData "CPM36", textCPM36 'Add By Sindy 2023/10/18
   SetFieldNewData "CPM30", textCPM30 'Add By Sindy 2024/12/6
   'Removed by Morgan 2020/9/29
   'SetFieldNewData "CPM30", textCPM30 'Added by Morgan 2014/4/10
   'end 2020/9/29
   SetFieldNewData "CPM31", textCPM31 'Added by Morgan 2018/5/10
   SetFieldNewData "CPM32", textCPM32 'Added by Morgan 2018/5/10
   SetFieldNewData "CPM33", textCPM33 'Add By Sindy 2020/8/13
   SetFieldNewData "CPM34", textCPM34 'Add By Sindy 2021/4/28
   SetFieldNewData "CPM35", textCPM35 'Add by Amy 2022/09/19
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Private Sub QueryDB()
   RefreshRange
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   
   Dim oObj As Object
   
   For Each oObj In Me.Controls
      If TypeName(oObj) = "TextBox" Then
        oObj.Text = Empty
      End If
   Next
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
    'Add By Cheng 2002/11/21
    If Me.lvw.ListItems.Count > 0 Then
        For nIndex = 1 To Me.lvw.ListItems.Count
            Me.lvw.ListItems(nIndex).Checked = False
        Next nIndex
    End If
    
    'Removed by Morgan 2020/9/29
    'Label28 = "" 'Added by Morgan 2014/4/11
    'end 2020/9/29
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textCPM01.Locked = bEnable
   textCPM02.Locked = bEnable
   textCPM03.Locked = bEnable
   textCPM04.Locked = bEnable
   textCPM05.Locked = bEnable 'Modified by Lydiad 2025/06/18 改成「是否算案件數」
   textCPM06.Locked = bEnable
   textCPM07.Locked = bEnable
   textCPM08.Locked = bEnable
   textCPM09.Locked = bEnable
   textCPM10.Locked = bEnable
   textCPM11.Locked = bEnable
   textCPM12.Locked = bEnable
   textCPM13.Locked = bEnable
   'add by nickc 2006/03/09
   textCPM14.Locked = bEnable
   textCPM15.Locked = bEnable
   'Add by Morgan 2007/7/19
   TextCPM16.Locked = bEnable
   TextCPM17.Locked = bEnable
   TextCPM18.Locked = bEnable
   'Add by Morgan 2008/7/9
   textCPM19.Locked = bEnable
   'Add by Morgan 2008/9/24
   textCPM20.Locked = bEnable
   'Add by Morgan 2009/3/18
   textCPM21.Locked = bEnable
   textCPM22.Locked = bEnable
   textCPM23.Locked = bEnable  '2009/4/9 ADD BY SONIA
   'Add by Morgan 2010/4/26
   textCPM24.Locked = bEnable
   textCPM25.Locked = bEnable
   'Add by Sindy 2013/8/7
   textCPM26.Locked = bEnable
   'Add by Sindy 2013/10/1
   textCPM27.Locked = bEnable
   textCPM28.Locked = bEnable
   textCPM29.Locked = bEnable
   '2013/10/1 END
   textCPM36.Locked = bEnable 'Add By Sindy 2023/10/18
   textCPM30.Locked = bEnable 'Add By Sindy 2024/12/6
   'Removed by Morgan 2020/9/29
   'textCPM30.Locked = bEnable 'Added by Morgan 2014/4/10
   'end 2020/9/29
   textCPM31.Locked = bEnable 'Added by Morgan 2018/5/10
   textCPM32.Locked = bEnable 'Added by Morgan 2018/5/10
   textCPM33.Locked = bEnable 'Add By Sindy 2020/8/13
   textCPM34.Locked = bEnable 'Add By Sindy 2021/4/28
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCPM01.Locked = bEnable
   textCPM02.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCPM23Type As String 'Add By Sindy 2025/7/31
   
   strSql = "SELECT * FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = '" & m_CurrCM(0) & "' AND " & _
                  "CPM02 = '" & m_CurrCM(1) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   LblCPM23Code.Caption = "" 'Add By Sindy 2025/7/31
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("CPM01")) = False Then
         textCPM01 = rsTmp.Fields("CPM01")
      End If
      If IsNull(rsTmp.Fields("CPM02")) = False Then
         textCPM02 = rsTmp.Fields("CPM02")
      End If
      If IsNull(rsTmp.Fields("CPM03")) = False Then
         textCPM03 = rsTmp.Fields("CPM03")
      End If
      textCPM03.Tag = textCPM03.Text 'Add By Sindy 2024/11/25
      If IsNull(rsTmp.Fields("CPM04")) = False Then
         textCPM04 = rsTmp.Fields("CPM04")
      End If
      textCPM04.Tag = textCPM04.Text 'Add By Sindy 2024/11/25

      'Modified by Lydiad 2025/06/18 改成「是否算案件數」
      If IsNull(rsTmp.Fields("CPM05")) = False Then
         textCPM05 = rsTmp.Fields("CPM05")
      End If

      If IsNull(rsTmp.Fields("CPM06")) = False Then
         textCPM06 = rsTmp.Fields("CPM06")
      End If
      If IsNull(rsTmp.Fields("CPM07")) = False Then
         textCPM07 = rsTmp.Fields("CPM07")
      End If
      If IsNull(rsTmp.Fields("CPM08")) = False Then
         textCPM08 = rsTmp.Fields("CPM08")
      End If
      If IsNull(rsTmp.Fields("CPM09")) = False Then
         textCPM09 = rsTmp.Fields("CPM09")
      End If
      If IsNull(rsTmp.Fields("CPM10")) = False Then
         textCPM10 = rsTmp.Fields("CPM10")
      End If
      If IsNull(rsTmp.Fields("CPM11")) = False Then
         textCPM11 = rsTmp.Fields("CPM11")
      End If
      If IsNull(rsTmp.Fields("CPM12")) = False Then
         textCPM12 = rsTmp.Fields("CPM12")
      End If
      If IsNull(rsTmp.Fields("CPM13")) = False Then
         textCPM13 = rsTmp.Fields("CPM13")
      End If
      'add by nickc 2006/03/09
      If IsNull(rsTmp.Fields("CPM14")) = False Then
         textCPM14 = rsTmp.Fields("CPM14")
      End If
      If IsNull(rsTmp.Fields("CPM15")) = False Then
         textCPM15 = rsTmp.Fields("CPM15")
      End If
      
      'Add by Morgan 2007/7/19
      If IsNull(rsTmp.Fields("CPM16")) = False Then
         TextCPM16 = rsTmp.Fields("CPM16")
      End If
      If IsNull(rsTmp.Fields("CPM17")) = False Then
         TextCPM17 = rsTmp.Fields("CPM17")
      End If
      If IsNull(rsTmp.Fields("CPM18")) = False Then
         TextCPM18 = rsTmp.Fields("CPM18")
      End If
      'end 2007/7/19
      'Add by Morgan 2008/7/9
      textCPM19 = "" & rsTmp.Fields("CPM19")
      'Add by Morgan 2008/9/24
      textCPM20 = "" & rsTmp.Fields("CPM20")
      'Add by Morgan 2009/3/18
      textCPM21 = "" & rsTmp.Fields("CPM21")
      textCPM22 = "" & rsTmp.Fields("CPM22")
      textCPM23 = "" & rsTmp.Fields("CPM23")  '2009/4/9 ADD BY SONIA
      
      'Add By Sindy 2025/7/31
      If textCPM23 <> "" Then
         If textCPM01 = "P" Or textCPM01 = "PS" Or textCPM01 = "CFP" Or textCPM01 = "CPS" Then
            strCPM23Type = P核判分類
         ElseIf Left(textCPM01, 1) = "T" Or textCPM01 = "FCT" Then
            strCPM23Type = T核判分類
         ElseIf textCPM01 = "CFT" And textCPM01 = "S" And textCPM01 = "CFC" Then
            strCPM23Type = CF核判分類
         Else
            strCPM23Type = ""
         End If
         If strCPM23Type <> "" Then
            strExc(10) = "'" & Trim(textCPM23) & " " 'ex:'A 發明'
            If InStr(strCPM23Type, strExc(10)) > 0 Then
               strExc(9) = InStr(strCPM23Type, strExc(10)) + Len(strExc(10)) '起始字數
               strCPM23Type = Mid(strCPM23Type, strExc(9))
               strExc(8) = InStr(strCPM23Type, "'")
               If Val(strExc(8)) > 0 Then
                  LblCPM23Code.Caption = "(" & textCPM23 & "." & Mid(strCPM23Type, 1, strExc(8) - 1) & ")"
               End If
            End If
         End If
      End If
      '2025/7/31 END
      
      'Add by Morgan 2010/4/26
      textCPM24 = "" & rsTmp.Fields("CPM24")
      textCPM25 = "" & rsTmp.Fields("CPM25")
      'Add by Sindy 2013/8/7
      textCPM26 = "" & rsTmp.Fields("CPM26")
      'Add by Sindy 2013/10/1
      textCPM27 = "" & rsTmp.Fields("CPM27")
      textCPM28 = "" & rsTmp.Fields("CPM28")
      textCPM29 = "" & rsTmp.Fields("CPM29")
      '2013/10/1 END
      textCPM36 = "" & rsTmp.Fields("CPM36") 'Add By Sindy 2023/10/18
      textCPM30 = "" & rsTmp.Fields("CPM30") 'Add By Sindy 2024/12/6
      'Added by Morgan 2014/4/11
      'Removed by Morgan 2020/9/29
      'textCPM30 = "" & rsTmp.Fields("CPM30")
      'If textCPM30 <> "" Then
      '   textCPM30_Validate False
      'End If
      'end 2020/9/29
      'end 2014/4/11
      
      textCPM31 = "" & rsTmp.Fields("CPM31") 'Added by Morgan 2018/5/10
      textCPM32 = "" & rsTmp.Fields("CPM32") 'Added by Morgan 2018/5/10
      textCPM33 = "" & rsTmp.Fields("CPM33") 'Add By Sindy 2020/8/13
      textCPM34 = "" & rsTmp.Fields("CPM34") 'Add By Sindy 2021/4/28
      textCPM35 = "" & rsTmp.Fields("CPM35") 'Add by Amy 2022/09/19
      
      UpdateFieldOldData rsTmp
      
      textCPM11_Validate False
      textCPM12_Validate False
      textCPM24_Validate False
      textCPM25_Validate False
      
   End If
   rsTmp.Close
               
    'Add By Cheng 2002/11/21
    UpdateGroupData textCPM01.Text, textCPM02.Text
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strCPM01 As String, ByVal strCPM02 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strCPM01, strCPM02) = True Then
      m_CurrCM(0) = strCPM01
      m_CurrCM(1) = strCPM02
   Else
      strSql = "SELECT CPM01,CPM02 FROM CASEPROPERTYMAP " & _
               "WHERE CPM01 = '" & m_CurrCM(0) & "' AND " & _
                     "CPM02 = (SELECT MIN(CPM02) FROM CASEPROPERTYMAP " & _
                             "WHERE CPM01 = '" & m_CurrCM(0) & "' AND " & _
                                   "CPM02 > '" & m_CurrCM(1) & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CPM01")) = False Then: m_CurrCM(0) = rsTmp.Fields("CPM01")
         If IsNull(rsTmp.Fields("CPM02")) = False Then: m_CurrCM(1) = rsTmp.Fields("CPM02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT CPM01,CPM02 FROM CASEPROPERTYMAP " & _
               "WHERE CPM01 = (SELECT MIN(CPM01) FROM CASEPROPERTYMAP " & _
                              "WHERE CPM01 > '" & m_CurrCM(0) & "') AND " & _
                     "CPM02 = (SELECT MIN(CPM02) FROM CASEPROPERTYMAP " & _
                              "WHERE CPM01 = (SELECT MIN(CPM01) FROM CASEPROPERTYMAP " & _
                                             "WHERE CPM01 > '" & m_CurrCM(0) & "')) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CPM01")) = False Then: m_CurrCM(0) = rsTmp.Fields("CPM01")
         If IsNull(rsTmp.Fields("CPM02")) = False Then: m_CurrCM(1) = rsTmp.Fields("CPM02")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrCM(0) = m_FirstCM(0)
   m_CurrCM(1) = m_FirstCM(1)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrCM(0) = m_FirstCM(0) And m_CurrCM(1) = m_FirstCM(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT CPM01,CPM02 FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = '" & m_CurrCM(0) & "' AND " & _
                  "CPM02 = (SELECT MAX(CPM02) FROM CASEPROPERTYMAP " & _
                          "WHERE CPM01 = '" & m_CurrCM(0) & "' AND " & _
                                "CPM02 < '" & m_CurrCM(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CPM01")) = False Then: m_CurrCM(0) = rsTmp.Fields("CPM01")
      If IsNull(rsTmp.Fields("CPM02")) = False Then: m_CurrCM(1) = rsTmp.Fields("CPM02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT CPM01,CPM02 FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = (SELECT MAX(CPM01) FROM CASEPROPERTYMAP " & _
                           "WHERE CPM01 < '" & m_CurrCM(0) & "') AND " & _
                  "CPM02 = (SELECT MAX(CPM02) FROM CASEPROPERTYMAP " & _
                           "WHERE CPM01 = (SELECT MAX(CPM01) FROM CASEPROPERTYMAP " & _
                                          "WHERE CPM01 < '" & m_CurrCM(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CPM01")) = False Then: m_CurrCM(0) = rsTmp.Fields("CPM01")
      If IsNull(rsTmp.Fields("CPM02")) = False Then: m_CurrCM(1) = rsTmp.Fields("CPM02")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrCM(0) = m_LastCM(0) And m_CurrCM(1) = m_LastCM(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT CPM01,CPM02 FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = '" & m_CurrCM(0) & "' AND " & _
                  "CPM02 = (SELECT MIN(CPM02) FROM CASEPROPERTYMAP " & _
                          "WHERE CPM01 = '" & m_CurrCM(0) & "' AND " & _
                                "CPM02 > '" & m_CurrCM(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CPM01")) = False Then: m_CurrCM(0) = rsTmp.Fields("CPM01")
      If IsNull(rsTmp.Fields("CPM02")) = False Then: m_CurrCM(1) = rsTmp.Fields("CPM02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT CPM01,CPM02 FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = (SELECT MIN(CPM01) FROM CASEPROPERTYMAP " & _
                           "WHERE CPM01 > '" & m_CurrCM(0) & "') AND " & _
                  "CPM02 = (SELECT MIN(CPM02) FROM CASEPROPERTYMAP " & _
                           "WHERE CPM01 = (SELECT MIN(CPM01) FROM CASEPROPERTYMAP " & _
                                          "WHERE CPM01 > '" & m_CurrCM(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CPM01")) = False Then: m_CurrCM(0) = rsTmp.Fields("CPM01")
      If IsNull(rsTmp.Fields("CPM02")) = False Then: m_CurrCM(1) = rsTmp.Fields("CPM02")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrCM(0) = m_LastCM(0)
   m_CurrCM(1) = m_LastCM(1)
   
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         ' 90.07.13 modify by louis (依照權限設定其工具列的按紐狀態)
         'tlbar.Buttons(1).Enabled = True
         'tlbar.Buttons(2).Enabled = True
         'tlbar.Buttons(3).Enabled = True
         'tlbar.Buttons(4).Enabled = True
         'tlbar.Buttons(6).Enabled = True
         'tlbar.Buttons(7).Enabled = True
         'tlbar.Buttons(8).Enabled = True
         'tlbar.Buttons(9).Enabled = True
         'tlbar.Buttons(11).Enabled = False
         'tlbar.Buttons(12).Enabled = False
         'tlbar.Buttons(14).Enabled = True
         
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040109 = Nothing
End Sub

Private Sub stb_Click(PreviousTab As Integer)
    'Add By Cheng 2002/11/21
    If Me.stb.Tab = 0 Then
        If Me.Visible = True Then
        'If Me.textCPM03.Enabled Then Me.textCPM03.SetFocus
        End If
    Else
    End If
End Sub

'Removed by Morgan 2020/9/29
'Private Sub textCPM01_Change()
'   SetCPM30
'End Sub
'
'Private Sub SetCPM30()
'   If m_EditMode = 1 Or m_EditMode = 2 Then
'      If textCPM01 = "P" And Len(textCPM02) = 4 Then
'         textCPM30 = "73022"
'         textCPM30_Validate False
'      Else
'         textCPM30 = ""
'         Label28 = ""
'      End If
'   End If
'End Sub
'end 2020/9/29

Private Sub textCPM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 系統類別
Private Sub textCPM01_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPM01) = False Then
      If IsAlphabetic(textCPM01) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "系統類別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM01_GotFocus
         GoTo EXITSUB
      End If
      Select Case m_EditMode
         Case 1:
            strSql = "SELECT * FROM SYSTEMKIND " & _
                     "WHERE SK01 = '" & textCPM01 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount <= 0 Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCPM01_GotFocus
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         Case Else
      End Select
   End If
EXITSUB:
End Sub

'Removed by Morgan 2020/9/29
'Private Sub textCPM02_Change()
'   SetCPM30
'End Sub
'end 2020/9/29

' 案件性質代號
Private Sub textCPM02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPM02) = False Then
      If IsNumeric(textCPM02) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM02_GotFocus
         GoTo EXITSUB
      End If
      Select Case m_EditMode
         Case 1:
            If IsEmptyText(textCPM01) = False Then
               If IsRecordExist(textCPM01, textCPM02) = True Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "該筆記錄已經存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textCPM02_GotFocus
                  GoTo EXITSUB
               End If
            End If
         Case Else:
      End Select
   End If
EXITSUB:
End Sub

' 國內案件性質名稱
Private Sub textCPM03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCPM03, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "國內案件性質名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCPM03_GotFocus
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: textCPM03.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 大陸案件性質名稱
Private Sub textCPM04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCPM04, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "大陸案件性質名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCPM04_GotFocus
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: textCPM04.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Modified by Lydiad 2025/06/18 改成「是否算案件數」
Private Sub textCPM05_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" And Chr(KeyAscii) <> "N" Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 專業考核件數  -->2006/03/09 修改為承辦人計件值，實際使用為2005/01/01
Private Sub textCPM06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPM06) = False Then
      If IsNumeric(textCPM06) = False Then
         Cancel = True
         strTit = "檢核資料"
         'edit by nickc 206/03/09
         'strMsg = "專業考核件數請輸入數值資料"
         strMsg = "承辦人計件值請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         'add by nickc 2006/03/09
         textCPM06.SetFocus
         textCPM06_GotFocus
         'add by nickc 2006/03/09
         Exit Sub
      End If
      'add by nickc 2006/03/09
      If Val(textCPM06) >= 10 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "承辦人計件值請輸入10 以下數字資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM06.SetFocus
         textCPM06_GotFocus
         Exit Sub
      End If
   End If
End Sub

' 英文案件性質名稱
Private Sub textCPM10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCPM10, 80) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "英文案件性質名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCPM10_GotFocus
   End If
End Sub

' 日文案件性質名稱
Private Sub textCPM13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCPM13, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "日文案件性質名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCPM13_GotFocus
   End If
   'add by nickc 2007/07/11 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

' 來函期限
Private Sub textCPM07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPM07) = False Then
      Select Case textCPM07
         Case "1", "2":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "來函期限只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCPM07_GotFocus
      End Select
   End If
End Sub

' 期限天數
Private Sub textCPM08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPM08) = False Then
      If IsNumeric(textCPM08) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "期限天數只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM08_GotFocus
      End If
   End If
End Sub

' 期限月數
Private Sub textCPM09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPM09) = False Then
      If IsNumeric(textCPM09) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "期限月數只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM09_GotFocus
      End If
   End If
End Sub

' 收入會計科目
Private Sub textCPM11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCPM11_2 = Empty
   If IsEmptyText(textCPM11) = False Then
      textCPM11_2 = GetAccountingTitle(textCPM11)
      Select Case m_EditMode
         Case 1, 2:
            If IsEmptyText(textCPM11_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "收入會計科目編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCPM11_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

' 規費會計科目
Private Sub textCPM12_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCPM12_2 = Empty
   If IsEmptyText(textCPM12) = False Then
      textCPM12_2 = GetAccountingTitle(textCPM12)
      Select Case m_EditMode
         Case 1, 2:
            If IsEmptyText(textCPM12_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "規費會計科目編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCPM12_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 90.07.13 modify by louis
      ' 新增
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
      '   If m_EditMode = 0 Then
      '      OnAction KeyCode
      '      KeyCode = 0
      '   End If
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
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
'edit by nickc 2006/11/13
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub
' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   m_SubMode = 0
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            OnWork
            UpdateToolbarState
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         OnWork
         UpdateToolbarState
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub


'add by nickc 2006/03/09
Private Sub textCPM14_GotFocus()
    InverseTextBox textCPM14
End Sub

'add by nickc 2006/03/09
Private Sub textCPM14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPM14) = False Then
      If IsNumeric(textCPM14) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "繪圖計件值 草 請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM14.SetFocus
         textCPM14_GotFocus
         Exit Sub
      End If
      If Val(textCPM14) >= 10 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "繪圖計件值 草 請輸入10 以下數字資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM14.SetFocus
         textCPM14_GotFocus
         Exit Sub
      End If
   End If
End Sub

'add by nickc 2006/03/09
Private Sub textCPM15_GotFocus()
    InverseTextBox textCPM15
End Sub

'add by nickc 2006/03/09
Private Sub textCPM15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPM15) = False Then
      If IsNumeric(textCPM15) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "繪圖計件值 墨 請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM15.SetFocus
         textCPM15_GotFocus
         Exit Sub
      End If
      If Val(textCPM15) >= 10 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "繪圖計件值 墨 請輸入10 以下數字資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCPM15.SetFocus
         textCPM15_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub TextCPM16_GotFocus()
   TextInverse TextCPM16
End Sub

Private Sub TextCPM16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub TextCPM17_GotFocus()
   TextInverse TextCPM17
End Sub

Private Sub TextCPM17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub TextCPM18_GotFocus()
   TextInverse TextCPM18
End Sub
Private Sub textCPM18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(TextCPM18) = False Then
      If IsNumeric(TextCPM18) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "預設請款金額只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextCPM18_GotFocus
      End If
   End If
End Sub

Private Sub textCPM19_GotFocus()
   InverseTextBox textCPM19
End Sub

Private Sub textCPM19_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

Private Sub textCPM33_GotFocus()
   InverseTextBox textCPM33
End Sub

Private Sub textCPM33_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

Private Sub textCPM20_GotFocus()
   InverseTextBox textCPM20
End Sub

Private Sub textCPM20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCPM21_GotFocus()
   TextInverse textCPM21
End Sub

Private Sub textCPM21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("Q") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCPM22_GotFocus()
   TextInverse textCPM22
End Sub

Private Sub textCPM22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCPM23_GotFocus()
   TextInverse textCPM23
End Sub

Private Sub textCPM23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And (KeyAscii < Asc("A") Or KeyAscii > Asc("P")) And KeyAscii <> Asc("9") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCPM24_GotFocus()
   InverseTextBox textCPM24
End Sub

Private Sub textCPM24_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCPM24_2 = Empty
   If IsEmptyText(textCPM24) = False Then
      textCPM24_2 = GetAccountingTitle(textCPM24)
      Select Case m_EditMode
         Case 1, 2:
            If IsEmptyText(textCPM24_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "收入會計科目編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCPM24_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

Private Sub textCPM25_GotFocus()
   InverseTextBox textCPM25
End Sub

Private Sub textCPM25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCPM25_2 = Empty
   If IsEmptyText(textCPM25) = False Then
      textCPM25_2 = GetAccountingTitle(textCPM25)
      Select Case m_EditMode
         Case 1, 2:
            If IsEmptyText(textCPM25_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "規費會計科目編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCPM25_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

Private Sub TextCPM27_GotFocus()
   TextInverse textCPM27
End Sub
Private Sub TextCPM27_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub TextCPM28_GotFocus()
   TextInverse textCPM28
End Sub
Private Sub TextCPM28_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

''Add By Sindy 2019/3/26
'Private Sub textCPM28_Validate(Cancel As Boolean)
'   If m_EditMode = 1 Or m_EditMode = 2 Then
'      If textCPM01 = "P" And Len(textCPM02) = 4 And textCPM28 <> "N" Then
'         MsgBox "P案C類來函請設定為不會稿 !", vbExclamation
'         Cancel = True
'      End If
'   End If
'End Sub

Private Sub TextCPM29_GotFocus()
   TextInverse textCPM29
End Sub
Private Sub TextCPM29_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Sindy 2023/10/18
Private Sub TextCPM36_GotFocus()
   TextInverse textCPM36
End Sub
Private Sub TextCPM36_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'2023/10/18 END

'Add By Sindy 2024/12/6
Private Sub TextCPM30_GotFocus()
   TextInverse textCPM30
End Sub
Private Sub TextCPM30_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'2024/12/6 END

Private Sub textCPM31_GotFocus()
   InverseTextBox textCPM31
   CloseIme
End Sub

Private Sub textCPM31_Validate(Cancel As Boolean)
   If textCPM31 <> "" Then
      If IsNumeric(textCPM31) = False Then
         MsgBox "請輸入數值資料", vbExclamation
         Cancel = True
      End If
   End If
End Sub

Private Sub textCPM32_GotFocus()
   InverseTextBox textCPM32
   CloseIme
End Sub

Private Sub textCPM32_Validate(Cancel As Boolean)
   If textCPM32 <> "" Then
      If IsNumeric(textCPM32) = False Then
         MsgBox "請輸入數值資料", vbExclamation
         Cancel = True
      End If
   End If
End Sub

'Add By Sindy 2021/4/28
Private Sub TextCPM34_GotFocus()
   TextInverse textCPM34
End Sub
Private Sub TextCPM34_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add by Amy 2022/09/19
Private Sub textCPM35_GotFocus()
    TextInverse textCPM34
End Sub

Private Sub textCPM35_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2022/09/19

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
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
Private Function IsRecordExist(ByVal strCPM01 As String, ByVal strCPM02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = '" & strCPM01 & "' AND " & _
                  "CPM02 = '" & strCPM02 & "' "
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 新增記錄
Private Sub AddRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strCPM01 As String
   Dim strCPM02 As String
   
    'Add By Cheng 2002/11/21
   On Error GoTo ErrorHandler
   cnnConnection.BeginTrans
   
   strCPM01 = textCPM01
   strCPM02 = textCPM02
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strCPM01, strCPM02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO CASEPROPERTYMAP ("
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
   cnnConnection.Execute strSql
   
   If ((strCPM01 & strCPM02) < (m_FirstCM(0) & m_FirstCM(1))) Or ((strCPM01 & strCPM02) > (m_LastCM(0) & m_LastCM(1))) Then
      RefreshRange
   End If
   
    'Add By Cheng 2002/11/21
    UpdateStaffGroupData strCPM01, strCPM02
    cnnConnection.CommitTrans
   
   ShowCurrRecord strCPM01, strCPM02
'Add By Cheng 2002/11/21
Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox "新增作業失敗, 請洽電腦中心人員!!!"
EXITSUB:
End Sub

' 修改記錄
Private Sub ModRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strCPM01 As String
   Dim strCPM02 As String
   
    'Add By Cheng 2002/11/21
    On Error GoTo ErrorHandler
    cnnConnection.BeginTrans
    
   strCPM01 = m_CurrCM(0)
   strCPM02 = m_CurrCM(1)
   
   strSql = "UPDATE CASEPROPERTYMAP SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = " & ChgSQL(m_FieldList(nIndex).fiNewData)
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   
   strSql = strSql & " " & _
                  "WHERE CPM01 = '" & strCPM01 & "' AND " & _
                        "CPM02 = '" & strCPM02 & "' "
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      'Add by Morgan 2008/10/9
      If m_FieldList(17).fiName = "CPM18" And m_FieldList(17).fiNewData <> m_FieldList(17).fiOldData Then
         strSql = "Update acc1j0 set a1j17=" & CNULL(m_FieldList(17).fiNewData, True) & " where a1j01='" & strCPM01 & "' and a1j02='" & strCPM02 & "'"
         cnnConnection.Execute strSql, intI
         If intI = 0 Then
            strSql = "insert into acc1j0(a1j01,a1j02,a1j03,a1j04,a1j16,a1j17) values('" & strCPM01 & "','" & strCPM02 & "','" & textCPM03 & "','" & textCPM10 & "','" & textCPM13 & "'," & Val(TextCPM18) & ")"
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   
    'Add By Cheng 2002/11/21
    UpdateStaffGroupData strCPM01, strCPM02
    cnnConnection.CommitTrans
    
    'Add By Sindy 2024/11/25
    If bDifference = True Then
      If (textCPM03.Tag <> textCPM03.Text And textCPM03.Tag = "（無）") _
         Or (textCPM04.Tag <> textCPM04.Text And textCPM04.Tag = "（無）") Then
         If textCPM23 <> "" And textCPM23 <> "9" Then
            MsgBox "請確認是否需要設定該案件性質的核判表！", vbInformation
         End If
      End If
   End If
   '2024/11/25 END
   
    ShowCurrRecord strCPM01, strCPM02

'Add By Cheng 2002/11/21
Exit Sub

ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox "修改作業失敗, 請洽電腦中心人員!!!"
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strCPM01 As String
   Dim strCPM02 As String
   
    'Add By Cheng 2002/11/21
    On Error GoTo ErrorHandler
    cnnConnection.BeginTrans
    
   strCPM01 = m_CurrCM(0)
   strCPM02 = m_CurrCM(1)

   strSql = "DELETE FROM CASEPROPERTYMAP " & _
            "WHERE CPM01 = '" & strCPM01 & "' AND " & _
                  "CPM02 = '" & strCPM02 & "' "

   cnnConnection.Execute strSql
        
    'Add By Cheng 2002/11/21
    strSql = "DELETE FROM Staff_Group " & _
            "WHERE SG02 = '" & strCPM01 & "' AND " & _
                  "SG03 = '" & strCPM02 & "' "
    cnnConnection.Execute strSql
    
    cnnConnection.CommitTrans
    
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strCPM01 = m_LastCM(0) And strCPM02 = m_LastCM(1)) Or (strCPM01 = m_FirstCM(0) And strCPM02 = m_FirstCM(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strCPM01, strCPM02
   
Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox "刪除作業失敗, 請洽電腦中心人員!!!"
'ExitSub:
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   QueryRecord = False

   If IsRecordExist(textCPM01, textCPM02) = True Then
      m_CurrCM(0) = textCPM01
      m_CurrCM(1) = textCPM02
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            'Add By Sindy 2018/10/31
            If textCPM23 <> "" And textCPM23 <> "9" Then
               MsgBox "請確認是否需要設定該案件性質的核判表！", vbInformation
            End If
            
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
         RefreshRange
      Case 4:
         If CheckDataValid() = True Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textCPM01.SetFocus
      'Case 2: textCPM03.SetFocus
      Case 4: textCPM01.SetFocus
   End Select
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2, 4:
         ' 系統類別不可空白
         If IsEmptyText(textCPM01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入系統類別"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCPM01.SetFocus
            GoTo EXITSUB
         End If
         ' 案件性質代號不可為空白
         If IsEmptyText(textCPM02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入案件性質"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCPM02.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
      
   Select Case m_EditMode
      Case 1, 2:
         ' 國內案件性質名稱不可為空白
         If IsEmptyText(textCPM03) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入國內案件性質名稱質"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCPM03.SetFocus
            GoTo EXITSUB
         End If
         
         'Add By Sindy 2021/4/28
         If (textCPM01 = "FCP" Or textCPM01 = "FG") And Trim(textCPM34) = "" Then
            If MsgBox("未設定是否為 ""主管機關期限"" 要設定嗎?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
               textCPM34.SetFocus
               GoTo EXITSUB
            End If
         End If
      Case Else:
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCPM01_GotFocus()
   InverseTextBox textCPM01
   CloseIme
End Sub

Private Sub textCPM02_GotFocus()
   InverseTextBox textCPM02
End Sub

Private Sub textCPM03_GotFocus()
   InverseTextBox textCPM03
   'edit by nickc 2007/07/11 切換輸入法改用API
   'textCPM03.IMEMode = 1
   OpenIme
End Sub

Private Sub textCPM04_GotFocus()
   InverseTextBox textCPM04
   'edit by nickc 2007/07/11 切換輸入法改用API
   'textCPM04.IMEMode = 1
   OpenIme
End Sub

'Modified by Lydiad 2025/06/18 改成「是否算案件數」
Private Sub textCPM05_GotFocus()
   InverseTextBox textCPM05
End Sub

Private Sub textCPM06_GotFocus()
   InverseTextBox textCPM06
End Sub

Private Sub textCPM07_GotFocus()
   InverseTextBox textCPM07
End Sub

Private Sub textCPM08_GotFocus()
   InverseTextBox textCPM08
End Sub

Private Sub textCPM09_GotFocus()
   InverseTextBox textCPM09
End Sub

Private Sub textCPM10_GotFocus()
   InverseTextBox textCPM10
End Sub

Private Sub textCPM11_GotFocus()
   InverseTextBox textCPM11
End Sub

Private Sub textCPM12_GotFocus()
   InverseTextBox textCPM12
End Sub

Private Sub textCPM13_GotFocus()
   InverseTextBox textCPM13
   'edit by nickc 2007/07/11 切換輸入法改用API
   'textCPM13.IMEMode = 1
   OpenIme
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCPM01.Enabled = True Then
   Cancel = False
   textCPM01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM02.Enabled = True Then
   Cancel = False
   textCPM02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM03.Enabled = True Then
   Cancel = False
   textCPM03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM04.Enabled = True Then
   Cancel = False
   textCPM04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM06.Enabled = True Then
   Cancel = False
   textCPM06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'add by nickc 2006/03/09
If Me.textCPM14.Enabled = True Then
   Cancel = False
   textCPM14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textCPM15.Enabled = True Then
   Cancel = False
   textCPM15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM07.Enabled = True Then
   Cancel = False
   textCPM07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM08.Enabled = True Then
   Cancel = False
   textCPM08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM09.Enabled = True Then
   Cancel = False
   textCPM09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM10.Enabled = True Then
   Cancel = False
   textCPM10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM11.Enabled = True Then
   Cancel = False
   textCPM11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM12.Enabled = True Then
   Cancel = False
   textCPM12_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM13.Enabled = True Then
   Cancel = False
   textCPM13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM24.Enabled = True Then
   Cancel = False
   textCPM24_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPM25.Enabled = True Then
   Cancel = False
   textCPM25_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Morgan 2014/4/11
'Removed by Morgan 2020/9/29
'If Me.textCPM30.Enabled = True Then
'   Cancel = False
'   textCPM30_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
'end 2020/9/29
'end 2014/4/11

'Added by Morgan 2019/6/19
If m_EditMode = "1" Then
   If textCPM01 = "P" And Len(textCPM02) = 3 And textCPM03 <> "（無）" Then
      MsgBox "請詢問程序人員是否有客戶函附件副檔名要設定，若回覆有的話要手動新增到【 CustLetterRefExt 】！", vbExclamation, "P台灣案AB類客戶函附件副檔名檢查"
   End If
End If
'end 2019/6/19

'Added by Lydia 2025/06/18 改成「是否算案件數」
If InStr(",P,PS,CFP,CPS,", "," & textCPM01 & ",") > 0 And Len(Trim(textCPM02)) < 4 And Trim(textCPM05) = "" Then
   MsgBox "請輸入是否算案件數(Y/N) ！", vbExclamation
   stb.Tab = 2
   textCPM05.SetFocus
   textCPM05_GotFocus
   Exit Function
End If
'end 2025/06/18

'Added by Lydia 2021/10/15 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

TxtValidate = True
End Function

'Add By Cheng 2002/11/21
Private Sub ShowStaffGroup()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select Distinct ST11 From Staff  Where ST11 Is Not Null Order By ST11 "
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    While Not rsA.EOF
        Me.lvw.ListItems.add , "" & rsA.Fields(0).Value, "" & rsA.Fields(0).Value
        rsA.MoveNext
    Wend
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Sub

'Add By Cheng 2002/11/21
Private Sub UpdateGroupData(strSG02 As String, strSG03 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
    
On Error Resume Next
    
If Me.lvw.ListItems.Count > 0 Then
    For ii = 1 To Me.lvw.ListItems.Count
        Me.lvw.ListItems(ii).Checked = False
    Next ii
    StrSQLa = "Select * From Staff_Group  Where SG02 = '" & strSG02 & "' AND SG03 = '" & strSG03 & "' "
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            Me.lvw.ListItems.Item("" & rsA.Fields(0).Value).Checked = True
            rsA.MoveNext
        Wend
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End If
    
End Sub

'Add By Cheng 2002/11/21
Private Sub UpdateStaffGroupData(strSG02 As String, strSG03 As String)
Dim ii As Integer

    If Me.lvw.ListItems.Count > 0 Then
        strSql = "Delete From Staff_Group Where SG02 = '" & strSG02 & "' And SG03='" & strSG03 & "'"
        cnnConnection.Execute strSql
        For ii = 1 To Me.lvw.ListItems.Count
            If Me.lvw.ListItems(ii).Checked = True Then
                strSql = "Insert Into Staff_Group Values('" & Me.lvw.ListItems(ii).Text & "','" & strSG02 & "','" & strSG03 & "' ) "
                cnnConnection.Execute strSql
            End If
        Next ii
    End If
End Sub
