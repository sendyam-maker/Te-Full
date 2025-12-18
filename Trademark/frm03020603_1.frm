VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020603_1 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¦U¦¡¥Ó½Ð®Ñ-¸É¥¿,¥Ó½Ð·N¨£®Ñ"
   ClientHeight    =   5616
   ClientLeft      =   72
   ClientTop       =   996
   ClientWidth     =   8616
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5616
   ScaleWidth      =   8616
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7770
      TabIndex        =   36
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5820
      TabIndex        =   34
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¦^«eµe­±(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6660
      TabIndex        =   35
      Top             =   45
      Width           =   1080
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm03020603_1.frx":0000
      Left            =   1260
      List            =   "frm03020603_1.frx":000D
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   41
      Top             =   847
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   37
      Top             =   210
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1575
      MaxLength       =   6
      TabIndex        =   38
      Top             =   210
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2415
      MaxLength       =   1
      TabIndex        =   39
      Top             =   210
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2655
      MaxLength       =   2
      TabIndex        =   40
      Top             =   210
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1890
      MaxLength       =   1
      TabIndex        =   33
      Top             =   5310
      Width           =   300
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3225
      Left            =   180
      TabIndex        =   64
      Top             =   2100
      Width           =   8235
      _ExtentX        =   14520
      _ExtentY        =   5694
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "°ò¥»¸ê®Æ"
      TabPicture(0)   =   "frm03020603_1.frx":001D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label18(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstNameAgent"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Check1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textCP27"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Check2(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Check2(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Check2(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check1(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Check1(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame7"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Check2(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Check2(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "®Ö»é²z¥Ñ"
      TabPicture(1)   =   "frm03020603_1.frx":0039
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check4(0)"
      Tab(1).Control(1)=   "Check4(1)"
      Tab(1).Control(2)=   "Check4(2)"
      Tab(1).Control(3)=   "Check4(3)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "ªþ¥ó"
      TabPicture(2)   =   "frm03020603_1.frx":0055
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check3(6)"
      Tab(2).Control(1)=   "Check3(0)"
      Tab(2).Control(2)=   "Check3(2)"
      Tab(2).Control(3)=   "Check3(3)"
      Tab(2).Control(4)=   "Check3(4)"
      Tab(2).Control(5)=   "Check3(11)"
      Tab(2).Control(6)=   "Frame4"
      Tab(2).Control(7)=   "Frame5"
      Tab(2).Control(8)=   "Frame6"
      Tab(2).Control(9)=   "Check3(7)"
      Tab(2).Control(10)=   "Check3(5)"
      Tab(2).Control(11)=   "Text9"
      Tab(2).Control(12)=   "Check3(1)"
      Tab(2).Control(13)=   "chkAtt1(0)"
      Tab(2).ControlCount=   14
      Begin VB.CheckBox chkAtt1 
         Caption         =   "°ò¥»¸ê®Æªí"
         Height          =   255
         Index           =   0
         Left            =   -70200
         TabIndex        =   91
         Tag             =   ".contact.pdf"
         Top             =   330
         Value           =   1  '®Ö¨ú
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "«ü©w°Ó«~¡þªA°È­×¥¿²M³æ¤A¥÷¡C"
         Height          =   255
         Index           =   1
         Left            =   -74790
         TabIndex        =   27
         Top             =   570
         Width           =   2895
      End
      Begin VB.CheckBox Check2 
         Caption         =   "ªþ°Ó«~²M³æ"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   6630
         TabIndex        =   5
         Top             =   375
         Width           =   1245
      End
      Begin VB.CheckBox Check2 
         Caption         =   "¸ÉÃº³W¶O"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   5580
         TabIndex        =   4
         Top             =   375
         Width           =   1035
      End
      Begin VB.CheckBox Check4 
         Caption         =   "·~©ó ¦~ ¤ë ¤éÅÜ§ó¨ä¤¤Ä¶¦W¦b®×¡A¬G¥Ó½Ð¤H»P¾Ú¥H®Ö»é°Ó¼Ð¤§°Ó¼ÐÅv¤H¦P¤@¡A¥»®×®Ö»é²z¥Ñ§Y¤£´_¦s¦b¡AÂÔ½Ð¡@¶v§½½ç¬°®Ö­ã¤§³B¤À¡C"
         Height          =   375
         Index           =   3
         Left            =   -74790
         TabIndex        =   90
         Top             =   1800
         Width           =   7845
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame7"
         Height          =   225
         Left            =   210
         TabIndex        =   86
         Top             =   630
         Width           =   5055
         Begin VB.OptionButton Option7 
            Caption         =   "¥¿"
            Height          =   225
            Index           =   0
            Left            =   3000
            TabIndex        =   89
            Top             =   0
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.OptionButton Option7 
            Caption         =   "¼v"
            Height          =   225
            Index           =   1
            Left            =   3450
            TabIndex        =   88
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.CheckBox Check1 
            Caption         =   "¥N²z¤H©e¥ô®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   87
            Top             =   0
            Width           =   4155
         End
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   180
         Left            =   -74100
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   26
         Top             =   360
         Width           =   885
      End
      Begin VB.CheckBox Check3 
         Caption         =   "¥Ó½Ð¤H¦W±ø¤A¥÷¡C"
         Height          =   255
         Index           =   5
         Left            =   -74790
         TabIndex        =   85
         Top             =   1530
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "¤j³°¦a°Ï¤§¦ÛµM¤H©Îªk¤H¤§¨­¤ÀÃÒ©ú¤å¥ó¡C"
         Height          =   255
         Index           =   7
         Left            =   -74790
         TabIndex        =   84
         Top             =   2010
         Width           =   3795
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74790
         TabIndex        =   80
         Top             =   2250
         Width           =   4395
         Begin VB.OptionButton Option4 
            Caption         =   "¼v"
            Height          =   225
            Index           =   1
            Left            =   3450
            TabIndex        =   81
            Top             =   45
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.OptionButton Option4 
            Caption         =   "¥¿"
            Height          =   225
            Index           =   0
            Left            =   3000
            TabIndex        =   82
            Top             =   45
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CheckBox Check3 
            Caption         =   "ÅÜ§óÃÒ©ú®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74790
         TabIndex        =   76
         Top             =   2490
         Width           =   4395
         Begin VB.OptionButton Option5 
            Caption         =   "¥¿"
            Height          =   225
            Index           =   0
            Left            =   2940
            TabIndex        =   78
            Top             =   45
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.OptionButton Option5 
            Caption         =   "¼v"
            Height          =   225
            Index           =   1
            Left            =   3390
            TabIndex        =   77
            Top             =   45
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.CheckBox Check3 
            Caption         =   "²¾Âà«´¬ù®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   79
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74790
         TabIndex        =   72
         Top             =   2700
         Width           =   4395
         Begin VB.OptionButton Option6 
            Caption         =   "¥¿"
            Height          =   225
            Index           =   0
            Left            =   2910
            TabIndex        =   74
            Top             =   60
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.OptionButton Option6 
            Caption         =   "¼v"
            Height          =   225
            Index           =   1
            Left            =   3360
            TabIndex        =   73
            Top             =   60
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.CheckBox Check3 
            Caption         =   "±ÂÅv«´¬ù®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "°Ñ¦Ò¸ê®Æ¡G"
         Height          =   255
         Index           =   11
         Left            =   -74790
         TabIndex        =   32
         Top             =   2940
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Àu¥ýÅvÃÒ©ú¤å¥ó¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
         Height          =   255
         Index           =   4
         Left            =   -74790
         TabIndex        =   30
         Top             =   1290
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "©e¥ô®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
         Height          =   255
         Index           =   3
         Left            =   -74790
         TabIndex        =   29
         Top             =   1050
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "«ü©w¨Ï¥Î°Ó«~¡þªA°È¦W±ø¤A¥÷¡C"
         Height          =   255
         Index           =   2
         Left            =   -74790
         TabIndex        =   28
         Top             =   810
         Width           =   2895
      End
      Begin VB.CheckBox Check3 
         Caption         =   "³W¶O                      ¤¸¾ã¡C"
         Height          =   255
         Index           =   0
         Left            =   -74790
         TabIndex        =   25
         Top             =   330
         Width           =   3795
      End
      Begin VB.CheckBox Check3 
         Caption         =   "¦P·N®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
         Height          =   255
         Index           =   6
         Left            =   -74790
         TabIndex        =   31
         Top             =   1770
         Width           =   3735
      End
      Begin VB.CheckBox Check4 
         Caption         =   $"frm03020603_1.frx":0071
         Height          =   375
         Index           =   2
         Left            =   -74790
         TabIndex        =   24
         Top             =   1350
         Width           =   7845
      End
      Begin VB.CheckBox Check4 
         Caption         =   "¦P·N§R°£¥»¥ó°Ó¼Ð«ü©w¤§¡u¡v°Ó«~¦WºÙ¡C¸g§R°£«e­z°Ó«~«á¡A¥»®×®Ö»é²z¥Ñ§Y¤£´_¦s¦b¡AÂÔ½Ð¡@¶v§½½ç¬°®Ö­ã¤§³B¤À¡C"
         Height          =   375
         Index           =   1
         Left            =   -74790
         TabIndex        =   23
         Top             =   900
         Width           =   7485
      End
      Begin VB.CheckBox Check4 
         Caption         =   "¦P·NÁn©ú¥»¥ó°Ó¼Ð¤£´N¡u¡v¤å¦r¥D±i°Ó¼ÐÅv¡C¸g¦¹Án©ú«á¡A¥»®×®Ö»é²z¥Ñ§Y¤£´_¦s¦b¡AÂÔ½Ð¡@¶v§½½ç¬°®Ö­ã¤§³B¤À¡C"
         Height          =   375
         Index           =   0
         Left            =   -74790
         TabIndex        =   22
         Top             =   450
         Width           =   7305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "¦P·N®Ñ¡C"
         Height          =   225
         Index           =   8
         Left            =   210
         TabIndex        =   21
         Top             =   2520
         Width           =   2085
      End
      Begin VB.CheckBox Check1 
         Caption         =   "§ó¥¿¦a§}¡G"
         Height          =   225
         Index           =   7
         Left            =   210
         TabIndex        =   20
         Top             =   2280
         Width           =   2085
      End
      Begin VB.CheckBox Check2 
         Caption         =   "¥Ó½Ð·N¨£®Ñ"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   4350
         TabIndex        =   3
         Top             =   375
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "¤å¥ó"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   2910
         TabIndex        =   1
         Top             =   375
         Width           =   705
      End
      Begin VB.CheckBox Check2 
         Caption         =   "°Ó«~"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   3630
         TabIndex        =   2
         Top             =   375
         Width           =   705
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   71
         Top             =   2040
         Width           =   4395
         Begin VB.OptionButton Option3 
            Caption         =   "¼v"
            Height          =   225
            Index           =   1
            Left            =   3360
            TabIndex        =   19
            Top             =   30
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.OptionButton Option3 
            Caption         =   "¥¿"
            Height          =   225
            Index           =   0
            Left            =   2910
            TabIndex        =   18
            Top             =   30
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CheckBox Check1 
            Caption         =   "±ÂÅv«´¬ù®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   70
         Top             =   1800
         Width           =   4395
         Begin VB.OptionButton Option2 
            Caption         =   "¼v"
            Height          =   225
            Index           =   1
            Left            =   3390
            TabIndex        =   16
            Top             =   -15
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.OptionButton Option2 
            Caption         =   "¥¿"
            Height          =   225
            Index           =   0
            Left            =   2940
            TabIndex        =   15
            Top             =   -15
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CheckBox Check1 
            Caption         =   "²¾Âà«´¬ù®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   69
         Top             =   1530
         Width           =   4395
         Begin VB.OptionButton Option1 
            Caption         =   "¼v"
            Height          =   225
            Index           =   1
            Left            =   3420
            TabIndex        =   13
            Top             =   -15
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.OptionButton Option1 
            Caption         =   "¥¿"
            Height          =   225
            Index           =   0
            Left            =   2970
            TabIndex        =   12
            Top             =   -15
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CheckBox Check1 
            Caption         =   "ÅÜ§óÃÒ©ú®Ñ¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.TextBox textCP27 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   6870
         MaxLength       =   7
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2550
         MaxLength       =   7
         TabIndex        =   8
         Top             =   870
         Width           =   1725
      End
      Begin VB.CheckBox Check1 
         Caption         =   "¤j³°¦a°Ï¤§¦ÛµM¤H©Îªk¤H¤§¨­¤ÀÃÒ©ú¤å¥ó¡C"
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   10
         Top             =   1290
         Width           =   3795
      End
      Begin VB.CheckBox Check1 
         Caption         =   "¥Nªí¤H¦WºÙ"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   9
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "¥D±iÀu¥ýÅv¤§ÃÒ©ú¤å¥ó ¡Ð                                         ¥Ó½Ð®ÑÁÃ¥»¤A¥÷¡]ªþ¤¤Ä¶¤å¡^¡C"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Top             =   855
         Width           =   6975
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1260
         MaxLength       =   7
         TabIndex        =   0
         Top             =   330
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   6570
         TabIndex        =   92
         Top             =   1170
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "µo¤å¤é´Á :"
         Height          =   180
         Left            =   6030
         TabIndex        =   68
         Top             =   630
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¤º®e: "
         ForeColor       =   &H00000080&
         Height          =   180
         Index           =   2
         Left            =   2460
         TabIndex        =   67
         Top             =   375
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "¥X¦W¥N²z¤H"
         Height          =   180
         Left            =   5520
         TabIndex        =   66
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð®Ñ¤é´Á :"
         Height          =   180
         Left            =   210
         TabIndex        =   65
         Top             =   375
         Visible         =   0   'False
         Width           =   990
      End
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   10
      Left            =   4470
      TabIndex        =   63
      Top             =   1770
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   9
      Left            =   1260
      TabIndex        =   62
      Top             =   1770
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó©Ê½è:"
      Height          =   180
      Left            =   3570
      TabIndex        =   61
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "¾÷Ãö¤å¸¹:"
      Height          =   180
      Left            =   3570
      TabIndex        =   60
      Top             =   1464
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "¨Ó¨ç¦¬¤å¤é:"
      Height          =   180
      Left            =   210
      TabIndex        =   59
      Top             =   1464
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4470
      TabIndex        =   58
      Top             =   240
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "´¼Åv¤H­û:"
      Height          =   180
      Left            =   3570
      TabIndex        =   57
      Top             =   1158
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "©Ó¿ì¤H¡@:"
      Height          =   180
      Left            =   210
      TabIndex        =   56
      Top             =   1158
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹:"
      Height          =   180
      Left            =   210
      TabIndex        =   55
      Top             =   210
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð®×¸¹:"
      Height          =   180
      Left            =   210
      TabIndex        =   54
      Top             =   546
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "¼f©w¸¹¼Æ:"
      Height          =   180
      Left            =   3570
      TabIndex        =   53
      Top             =   546
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "°Ó¼Ð¦WºÙ:"
      Height          =   180
      Left            =   210
      TabIndex        =   52
      Top             =   847
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   51
      Top             =   540
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4470
      TabIndex        =   50
      Top             =   540
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1980
      TabIndex        =   49
      Top             =   840
      Width           =   6510
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "11483;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   48
      Top             =   1155
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4470
      TabIndex        =   47
      Top             =   1155
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1260
      TabIndex        =   46
      Top             =   1455
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4470
      TabIndex        =   45
      Top             =   1470
      Width           =   4020
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "7091;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_­×§ï¥Ó½Ð®Ñ¤º®e          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   44
      Top             =   5370
      Width           =   2880
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "ªk©w´Á­­:"
      Height          =   180
      Index           =   0
      Left            =   3570
      TabIndex        =   43
      Top             =   1770
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò´Á­­:"
      Height          =   180
      Left            =   210
      TabIndex        =   42
      Top             =   1770
      Width           =   765
   End
End
Attribute VB_Name = "frm03020603_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/08/04 Form2.0¤w­×§ï; Label2(index)¡BlstNameAgent
'Memo By Sindy 2012/12/4 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/16 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/29 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo By Sindy 2010/8/11 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Dim strReceiveNo As String
Dim tm() As String, m_CP110 As String, m_AgentName As String
Dim intWhere As Integer, intLastRow As Integer
Dim m_strNPReceiveNo As String 'ÂI¿ï¥¼¦¬´Á­­ªº¦¬¤å¸¹
Dim m_CP10 As String '®×¥ó©Ê½è
Dim m_CP27 As String 'µo¤å¤é´Á
Dim m_CP43 As String '¬ÛÃöÁ`¦¬¤å¸¹
Dim m_CP64 As String '¶i«×³Æµù
Dim m_strLanguage As String '©w½Z»y¤å
Dim strCaseType As String
Dim ET03_1 As String 'Memo by Lydia 2023/05/03 µo¤å®É"¸É¥¿, ©ñ±ó±M¥ÎÅv,¸ÉÀu¥ýÅvÃÒ©ú"©w½Z
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim m_CP17 As String 'Add By Sindy 2015/3/24 ¦¬¤å³W¶O
'Added by Lydia 2019/02/21
Dim bol201CP118 As Boolean '¬O§_¹q¤l°e¥ó
Dim m_CaseNo As String '¹q¤l°e¥ó-¥»©Ò®×¸¹
Dim m_F21st07 As String 'FCTµ{§Ç¤À¾÷
Dim str201Detail As String '¥Ó½Ð¤º®e

'Added by Lydia 2019/02/21
Private Sub Check1_Click(Index As Integer)
    'Added by Lydai 2022/06/14 °ò¥»¸ê®Æ­¶ÅÒ¤Ä¿ï¤§¤å¥ó¡A½Ð©óªþ¥ó­¶ÅÒ¦Û°Ê¤Ä¿ï¡C
    If bol201CP118 = True And Check1(Index).Value = 1 Then
        Select Case Index
             Case 0: '©e¥ô®Ñ
                    Check3(3).Value = 1
             Case 1: 'Àu¥ýÅvÃÒ©ú¤å¥ó
                    Check3(4).Value = 1
             Case 5: '²¾Âà«´¬ù
                    Check3(9).Value = 1
             Case 4: '§ó¦WÃÒ©ú
                    Check3(8).Value = 1
             Case 6: '±ÂÅv«´¬ù
                    Check3(10).Value = 1
             Case 8: '¦P·N®Ñ
                    Check3(6).Value = 1
        End Select
    End If
End Sub

'Add By Sindy 2018/5/15
Private Sub Check2_Click(Index As Integer)
   'Modified by Lydia 2019/02/21
   'Check3(0).Value = 0
   If bol201CP118 = False Then
       Check3(0).Value = 0
   End If
   'end 2019/02/21
   Check3(1).Value = 0
   If Check2(3).Value = 1 Then '¸ÉÃº³W¶O
      Check3(0).Value = 1   '³W¶O
   ElseIf Check2(4).Value = 1 Then 'ªþ°Ó«~²M³æ
      Check3(1).Value = 1 '«ü©w°Ó«~¡þªA°È­×¥¿²M³æ¤A¥÷¡C
   'Added by Lydia 2022/06/14
   ElseIf Check2(1).Value = 1 Then '°Ó«~
      Check3(2).Value = 1 '«ü©w¨Ï¥Î°Ó«~¡þªA°È¦W±ø¤A¥÷¡C
   'end 2022/06/14
   End If
End Sub

'Added by Lydia 2019/02/21
Private Sub Check3_Click(Index As Integer)
    'Mark by Lydia 2019/02/21 «O¯d
'    If bol201CP118 = True And Check3(Index).Value = 1 Then
'        Select Case Index
'             Case 3: '©e¥ô®Ñ
'                    Check1(0).Value = 1
'             Case 4: 'Àu¥ýÅvÃÒ©ú¤å¥ó
'                    Check1(1).Value = 1
'             Case 9: '²¾Âà«´¬ù
'                    Check1(5).Value = 1
'             Case 8: '§ó¦WÃÒ©ú
'                    Check1(4).Value = 1
'        End Select
'    End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
Dim i As Integer
'Added by Lydia 2019/02/21
Dim strFolder As String, strFileName As String
Dim mET01 As String, mET03 As String
Dim mCP09 As String '¦¬¤å¸¹(«D¬ÛÃö¦¬¤å¸¹)
Dim strContent As String 'Added by Lydia 2019/08/14
Dim strFilePath As String, strFN01 As String 'Added by Lydia 2023/05/03

   Select Case Index
      Case 0 '½T©w
         
         If InStr("201¸É¥¿, 208¸ÉÀu¥ýÅvÃÒ©ú, 202¥Ó½Ð·N¨£®Ñ", m_CP10) > 0 Then  'Added by Lydia 2020/12/31 §PÂ_"201¸É¥¿, 208¸ÉÀu¥ýÅvÃÒ©ú, 202¥Ó½Ð·N¨£®Ñ"¤~»Ý­n³]©w¤º®e
            If Check2(0).Value = 0 And Check2(1).Value = 0 And Check2(2).Value = 0 _
               And Check2(3).Value = 0 And Check2(4).Value = 0 Then
               MsgBox "½ÐÂI¿ï¤º®e !", vbCritical
               SSTab1.Tab = 0
               Check2(0).SetFocus
               Exit Sub
            End If
         End If 'Added by Lydia 2020/12/31
         
         If Check2(0).Value = 1 Then '¸É¥¿¤å¥ó
            bolChk = False
            For i = 0 To 8 '6
               If Check1(i).Value = 1 Then
                  bolChk = True
                  Exit For
               End If
            Next
            If bolChk = False Then
               MsgBox "½Ð¿ï¾Ü±ý¸É¥¿¤å¥ó !", vbCritical
               SSTab1.Tab = 0
               Exit Sub
            End If
         End If
         
         If TxtValidate = False Then Exit Sub
         If FormSave = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Exit Sub
         
         If Text7 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         
         ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
         m_strLanguage = GetLetterLanguage(Text1, Text2, Text3, Text4)
         'Add By Sindy 2013/2/4
         bolEmail = PUB_GetEMailFlag(Text1 & Text2 & Text3 & Text4, , , bolPlusPaper) 'ÀË¬d¬O§_¥HE-Mail³qª¾
         '2013/2/4 End
'         If Check2(0).Value = 1 And Check2(1).Value = 0 And Check2(2).Value = 0 Then '¸É¥¿¤å¥ó
'            strTmp = "00"
'         ElseIf Check2(0).Value = 0 And Check2(1).Value = 1 And Check2(2).Value = 0 Then '¸É¥¿°Ó«~
'            strTmp = "01"
'         Else
            strTmp = "02" '©w½Z¦X¨Ö
'         End If
         strLetterDate = Text5.Text
         mCP09 = strReceiveNo 'Added by Lydia 2019/02/21 «O¯d¦¬¤å¸¹
         
         If strTmp = "" Then
            MsgBox "¸Ó©Ê½è¨ÃµL¥Ó½Ð®Ñ¡I"
         Else
            StartLetter "90", strReceiveNo, strTmp
            If ET03_1 <> "" Then
               'Modify By Sindy 2013/2/4
               'NowPrint strReceiveNo, "01", ET03_1, False, strUserNum
               'If bolEmail Then 'Mark by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
                  '§PÂ_¬O§_EMail¦P®É±H¯È¥»
                  If Not bolPlusPaper Then
                     iCopy = 1
                  End If
                  'Modified by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
                  'NowPrint strReceiveNo, "01", ET03_1, False, strUserNum, , , , , iCopy, , True, True
                  ''Modified by Lydia 2019/02/21
                  ''MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(Text1) & " ]¡I"
                  'If bol201CP118 = False Then
                  '    MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(Text1) & " ]¡I"
                  'End If
                  ''end 2019/02/21
                  'If m_strLanguage <> "3" Then '­^¤å²Õ 'Mark by Lydia 2024/11/14 ¦]¤é¥»¥N²z¤H¯S§O­n¨D¡A»Ý±N³qª¾«H¨ç»PÄ¶¤åµ¥¤À¶}¡A¨Ã¥B²Î¤@¦WºÙ¦p¤U(¼Ò²Õ¨ú±o)
                     strFilePath = Pub_GetEFilePath_All(Text1, Text2, Text3, Text4)
                     If Pub_GetFCTeFileName(strFilePath, Text1, Text2, Text3, Text4, m_CP10, , strFN01) = False Then
                       Exit Sub
                     End If
                     NowPrint strReceiveNo, "01", ET03_1, True, strUserNum, , , , , iCopy, , True
                     If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                         Sleep 100
                     End If
                  'Mark by Lydia 2024/11/14 ¦]¤é¥»¥N²z¤H¯S§O­n¨D¡A»Ý±N³qª¾«H¨ç»PÄ¶¤åµ¥¤À¶}¡A¨Ã¥B²Î¤@¦WºÙ¦p¤U(¼Ò²Õ¨ú±o)
                  'Else  '¤é¤å²Õ:¤£§ïÅÜ¦sÀÉ¼Ò¦¡
                  '    NowPrint strReceiveNo, "01", ET03_1, False, strUserNum, , , , , iCopy, , True, True
                  'End If
                  'end 2024/11/14
                  MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(Text1) & " ]¡I"
                  'end 2023/05/03
               'Mark by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
               'Else
               '   NowPrint strReceiveNo, "01", ET03_1, False, strUserNum
               'End If
               ''2013/2/4 End
               'end 2023/05/03
            End If
            'Modified by Lydia 2019/02/21 ¯È¥»¥Ó½Ð®Ñ
            'NowPrint strReceiveNo, "90", strTmp, bolChk, strUserNum
            If bol201CP118 = False Then
                 NowPrint strReceiveNo, "90", strTmp, bolChk, strUserNum
            Else
                'Added by Lydia 2019/02/21 ¦U¦¡¥Ó½Ð-¹q¤l°e¥ó-¸É¥¿
                m_CaseNo = PUB_FCPCaseNo2FileName(tm(1), tm(2), tm(3), tm(4))
                '®à­±¤W«Ø¥ß®×¸¹¸ê®Æ§¨
                strFolder = PUB_Getdesktop
                strFolder = strFolder & "\" & m_CaseNo
                If Dir(strFolder, vbDirectory) = "" Then
                    MkDir strFolder
                End If
                mET01 = "90"
                'Modified by Lydia 2019/02/26 +Àu¥ýÅv208
                'If m_CP10 = "201" Then '¸É¥¿
                'Modified by Lydia 2019/05/09 +202¥Ó½Ð·N¨£®Ñ
                'If m_CP10 = "201" Or m_CP10 = "202" Or m_CP10 = "208" Then 'Remove by Lydia 2020/12/31 «ü©w¦¬¤å©Ê½è¥H¥~ªºA¡BBÃþ¦¬¤å¡A¬Ò¥i²£¥Í¸É¥¿¥Ó½Ð®Ñ
                       '2.¥Ó½Ð®Ñ
                       'Modified by Lydia 2019/02/26 ³B²zª¬ªp04=>10
                       'mET03 = "04"
                       mET03 = "10"
                       If StartLetter2(mET01, mET03, mCP09) = False Then Exit Sub
                       'Added by Lydia 2019/08/14 §PÂ_­n°ò¥»¸ê®Æªí,¥ý¤£¦sÀÉ
                       If chkAtt1(0).Value = 1 Then
                            NowPrint mCP09, mET01, mET03, False, strUserNum, , , True, strContent
                            strFileName = strFolder & "\" & m_CaseNo & ".¸É¥¿¥Ó½Ð®Ñ-°ÓÂ²A"
                       Else
                       'end 2019/08/14
                            NowPrint mCP09, mET01, mET03, False, strUserNum, , , True, strContent
                            strFileName = strFolder & "\" & m_CaseNo & ".¸É¥¿¥Ó½Ð®Ñ-°ÓÂ²A"
                            Call PUB_MakeDoc(strContent, strFileName)
                       End If
                'End If 'Remove by Lydia 2020/12/31
                
                'Move by Lydia 2019/08/14 ±q¥Ó½Ð®Ñ¤W¤è²¾¤U¨Ó
                '1.°ò¥»¸ê®Æ
                If chkAtt1(0).Value = 1 Then
                       'Modified by Lydia 2020/12/31 ¹q¤l°e¥ó-°ò¥»¸ê®Æªí03=>11
                       mET03 = "11"
                       If StartLetter2(mET01, mET03, mCP09) = False Then Exit Sub
                       'Modified by Lydia 2019/08/14
                       'NowPrint mCP09, mET01, mET03, False, strUserNum, , , True, strContent
                       'strFileName = strFolder & "\" & m_CaseNo & ".contact"
                       'Call PUB_MakeDoc(strContent, strFileName)
                       NowPrint mCP09, mET01, mET03, False, strUserNum, , strContent, True, strContent
                       If strFileName = "" Then strFileName = strFolder & "\" & m_CaseNo & ".contact"
                       'Modified by Lydia 2020/09/25 ¼W¥[¤À¸`³B²z­¶½X
                       'Call PUB_MakeDoc(strContent, strFileName)
                       strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(¤À¸`)#|")    '´«­¶²Å¸¹Chr(12)´À´«¬°¤À¸`²Å¸¹ "|#(¤À¸`)#|"
                       Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '¤À¸`³B²z­¶½X
                       'end 2019/08/14
                       'end 2020/09/25
                End If
            End If
            'end 2019/02/21
         End If
         
         frm030206_1.Show
         '¦^¨ì­ìµe­±­n²M°£µe­±
         frm030206_1.ClearForm
         
      Case 1 '¦^«eµe­±
         frm030206_1.Show
         
      Case 2 'µ²§ô
         Unload frm030206_1
   End Select
   Unload Me
End Sub

Private Function ReadTMData() As String
   ReadTMData = ""
   strSql = "select * from trademark where tm01='" & Text1 & "' and tm02='" & Text2 & "' and tm03='" & IIf(Text3 = "", "0", Text3) & "' and tm04='" & IIf(Text4 = "", "00", Text4) & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      '¥Nªí¤H1(¤¤)
      If Not IsNull(RsTemp.Fields("tm47").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm47").Value)
      '¥Nªí¤H1(­^)
      If Not IsNull(RsTemp.Fields("tm48").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm48").Value)
      '¥Nªí¤H2(¤¤)
      If Not IsNull(RsTemp.Fields("tm50").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm50").Value)
      '¥Nªí¤H2(­^)
      If Not IsNull(RsTemp.Fields("tm51").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm51").Value)
      '¥Nªí¤H3(¤¤)
      If Not IsNull(RsTemp.Fields("tm94").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm94").Value)
      '¥Nªí¤H3(­^)
      If Not IsNull(RsTemp.Fields("tm95").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm95").Value)
      '¥Nªí¤H4(¤¤)
      If Not IsNull(RsTemp.Fields("tm97").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm97").Value)
      '¥Nªí¤H4(­^)
      If Not IsNull(RsTemp.Fields("tm98").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm98").Value)
      '¥Nªí¤H5(¤¤)
      If Not IsNull(RsTemp.Fields("tm100").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm100").Value)
      '¥Nªí¤H5(­^)
      If Not IsNull(RsTemp.Fields("tm101").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm101").Value)
      '¥Nªí¤H6(¤¤)
      If Not IsNull(RsTemp.Fields("tm103").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm103").Value)
      '¥Nªí¤H6(­^)
      If Not IsNull(RsTemp.Fields("tm104").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm104").Value)
      '¥Nªí¤H7(¤¤)
      If Not IsNull(RsTemp.Fields("tm106").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm106").Value)
      '¥Nªí¤H7(­^)
      If Not IsNull(RsTemp.Fields("tm107").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm107").Value)
      '¥Nªí¤H8(¤¤)
      If Not IsNull(RsTemp.Fields("tm109").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm109").Value)
      '¥Nªí¤H8(­^)
      If Not IsNull(RsTemp.Fields("tm110").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm110").Value)
      '¥Nªí¤H9(¤¤)
      If Not IsNull(RsTemp.Fields("tm112").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm112").Value)
      '¥Nªí¤H9(­^)
      If Not IsNull(RsTemp.Fields("tm113").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm113").Value)
      '¥Nªí¤H10(¤¤)
      If Not IsNull(RsTemp.Fields("tm115").Value) Then ReadTMData = ReadTMData & "¡B" & Trim(RsTemp.Fields("tm115").Value)
      '¥Nªí¤H10(­^)
      If Not IsNull(RsTemp.Fields("tm116").Value) Then ReadTMData = ReadTMData & Trim(RsTemp.Fields("tm116").Value)
   End If
End Function

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, strTmp As String
Dim ii As Integer, i As Integer, j As Integer, k As Integer, t As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCaseDate As String
Dim strTemp As Variant, strPrintNote As String
Dim Type0V As Boolean, Type1V As Boolean, Type4V As Boolean, Type5V As Boolean
Dim strCP43 As String, strCP10 As String, strCP27 As String 'Add By Sindy 2010/11/19
Dim strDebitNote As String 'Add By Sindy 2017/4/13

   EndLetter ET01, ET02, ET03, strUserNum
   ii = 0: i = 0: j = 0: k = 0: t = 0: Type0V = False: Type1V = False: Type4V = False: Type5V = False
   
   'Modify By Sindy 2017/4/13¡iFCT 01 000  04 ¨çª¾¤w¸É¤å¥ó.½Ð´Ú¡j
   m_MySt(1) = tm(1): m_MySt(2) = tm(2): m_MySt(3) = tm(3): m_MySt(4) = tm(4): m_Rule = strReceiveNo
   strDebitNote = ExceptFieldData2("FCT¯S®í½Ð´Ú¤å¦r¹ï·Ó")
   '2017/4/13 END
   
   strCaseType = ""
   strCaseDate = ""
   'Modify By Sindy 2010/11/19
   '2011/5/6 MODIFY BY SONIA ªü½¬»¡¥ý§ì¸É¥¿ªº¬ÛÃöÁ`¦¬¤å¸¹,­Y¬°CÃþ«h¦A©¹«e§ì¬ÛÃöÁ`¦¬¤å¸¹ªº®×¥ó©Ê½èFCT-018061,§ï¬°¤@¦¸¥ý§ì¦n
'      StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         strCP43 = "" & rsA.Fields("CP43")
'         If strCP43 <> "" And Left(strCP43, 1) = "C" Then
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "')) "
'            rsA.CursorLocation = adUseClient
'            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount > 0 Then
'               StrCP10 = "" & rsA.Fields("CP10")
'               strCP27 = "" & rsA.Fields("CP27")
'            End If
'         Else
'            StrCP10 = "" & rsA.Fields("CP10")
'            strCP27 = "" & rsA.Fields("CP27")
'         End If
'      End If
   '2011/7/22 MODIFY BY SONIA ¥[C2.CP43
   StrSQLa = "Select C1.CP43,C2.CP10,C2.CP27,C3.CP10,C3.CP27,C2.CP43 From Caseprogress C1,Caseprogress C2,Caseprogress C3 Where C1.CP09='" & strReceiveNo & "' AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      strCP43 = "" & rsA.Fields(0)
      If strCP43 <> "" And Left(strCP43, 1) = "C" Then
         strCP10 = "" & rsA.Fields(3)
         strCP27 = "" & rsA.Fields(4)
         m_CP43 = "" & rsA.Fields(5)   '2011/7/22 ADD BY SONIA
      Else
         strCP10 = "" & rsA.Fields(1)
         strCP27 = "" & rsA.Fields(2)
         m_CP43 = "" & rsA.Fields(0)   '2011/7/22 ADD BY SONIA
      End If
   End If
   '2011/5/6 END
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   '2010/11/19 End
   
   Select Case strCP10
      Case "101" '¥Ó½Ð
         strCaseType = "µù¥U"
         strCaseDate = tm(11) '¥Ó½Ð¤é
      Case "202" '¥Ó½Ð·N¨£®Ñ
         strCaseType = "µù¥U"
      Case Else
         'Modify By Sindy 2010/01/20 ªü½¬´£¥X­×§ï
         strTemp = Split(m_CP64, "©e¥ôª¬")
         If m_CP10 = "208" Or UBound(strTemp) > 0 Then
            'Modify By Sindy 2010/01/21 ªü½¬´£¥X­×§ï:208.¸ÉÀu¥ýÅvÃÒ©ú¤Î©e¥ôª¬§¡§ì¥Ó½Ð¤é
            strCaseDate = tm(11)
         Else
            strCaseDate = ChangeWStringToTString(strCP27)
         End If
         '2010/01/20 End
         Select Case strCP10
            Case "102" '©µ®i
               strCaseType = "©µ®iµù¥U"
            Case "301" 'ÅÜ§ó
               '§PÂ_¬O§_¦³¼f©w¸¹
               If Trim(Label12(2)) = "" Then
                  strCaseType = "µù¥U«eÅÜ§ó"
               Else
                  strCaseType = "µù¥UÅÜ§ó"
               End If
            Case "501" '²¾Âà
               strCaseType = "²¾Âàµn°O"
            Case "502" '±ÂÅv
               strCaseType = "±ÂÅvµn°O"
         End Select
   End Select
   '2011/7/22 ADD by sonia FCT-022611¨Ò¥~Äæ¦ìÀÉ»P©w½ZÀÉ¦s¤£¦PÁ`¦¬¤å¸¹
   '2013/8/19 mpdify by sonia ¥[¤JText6 = "1"±ø¥ó
   If (m_CP10 = "201" Or m_CP10 = "208") And _
      m_strLanguage = "2" And (strCaseType = "µù¥U«eÅÜ§ó" Or strCaseType = "²¾Âàµn°O") And _
      (Check2(0).Value = 1 And Check2(1).Value = 0 And Check2(2).Value = 0) Then
      ET02 = m_CP43
      EndLetter ET01, ET02, ET03, strUserNum
   End If
   '2011/7/22 END
   
   If Check2(2).Value = 1 Then '¥Ó½Ð·N¨£®Ñ
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','¸É¥¿¥Ó½Ð','¥Ó½Ð')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','¥D¦®¸É¥R¤º¤å','¡A´£¥X·N¨£®Ñ¨Æ¡C')"
   'Add By Sindy 2018/5/15
   ElseIf Check2(3).Value = 1 Then '¸ÉÃº³W¶O
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','¸É¥¿¥Ó½Ð','¥Ó½Ð')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','¥D¦®¸É¥R¤º¤å','¡A¸ÉÃº¤£¨¬¤§³W¶O" & Text9.Text & "¤¸¾ã¡AÂÔ½Ð¡@¶v§½´f¤©¨Ö®×¼f²z¡C')"
   '2018/5/15 END
   Else
      TmSt = "TM01='" & Text1 & "' AND TM02='" & Text2 & "' AND TM03='" & Text3 & "' AND TM04='" & Text4 & "'"
      strTmp = ExceptFieldData("°Ó¼Ðª¬ªp")
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','¸É¥¿¥Ó½Ð','¸É¥¿" & strTmp & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','¥D¦®¸É¥R¤º¤å','¦p»¡©ú¡AÂÔ½Ð¡@¶v§½´f¤©¨Ö®×¼f¬d¡C')"
   End If
   
   If strCaseType <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','®×¥óºØÃþ','" & ChgSQL(strCaseType) & "')"
   End If
   
   'Modified by Lydia 2019/03/05 ªü½¬»¡¤£+»¡©ú
   'str201Detail = "¡@¡@»¡©ú¡G" & vbCrLf  'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
   str201Detail = ""
   'end 2019/03/05
   'Add By Sindy 2012/1/5 ³¯ª÷½¬(Emily):½ÐÀ°§Ú§ïªþ¥ó¤§¥Ó½Ð®Ñ¡]·í¦¬¤å©Ê½è¬°201¡A¦ý¬ÛÃöÁ`¦¬¤å¸¹¬°®Ö»é«e¥ý¦æ³qª¾¡^
   'Modified by Lydia 2022/09/28 +C2.CP05
   StrSQLa = "Select C1.CP43,C2.CP10,C2.CP27,C2.CP43,C2.CP05 From Caseprogress C1,Caseprogress C2 Where C1.CP09='" & strReceiveNo & "' AND C1.CP43=C2.CP09(+) "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   strCP10 = ""
   strExc(1) = "" 'Added by Lydia 2022/09/28
   If rsA.RecordCount > 0 Then
      If Trim("" & rsA.Fields(1)) > "" Then
         strCP10 = "" & rsA.Fields(1)
      End If
      strExc(1) = TransDate("" & rsA.Fields("cp05"), 1) 'Added by Lydia 2022/09/28
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If strCP10 = "1202" Then '1202®Ö»é«e¥ý¦æ³qª¾
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
          "','»¡©ú¤@','·qÂÐ¡@¶v§½" & ChgSQL(Label12(7)) & "®Ö»é²z¥Ñ¥ý¦æ³qª¾®Ñ¡C')"
      'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
      'Modified by Lydia 2019/03/05 ¥h±¼¶}ÀY¨â­Ó¥þ§ÎªÅ¥Õ(­ì¥»4­Ó)
      str201Detail = str201Detail & "¡@¡@¤@¡B·qÂÐ¡@¶v§½" & ChgSQL(Label12(7)) & "®Ö»é²z¥Ñ¥ý¦æ³qª¾®Ñ¡C" & vbCrLf
   'Added by Lydia 2022/09/28 ¨ä¹ïÀ³¤§¬ÛÃöÁ`¦¬¤å¸¹¬°¡u¹q¸Ü³qª¾¡v®É¡A¥Ó½Ð®Ñ¤§¥Ó½Ð¤º®e²Ä¤@ÂI½Ð±a¡G¤@¡B·qÂÐ  ¶v§½XX¦~XX¤ëXX¤é¤§¹q¸Ü³qª¾¡C(¤é´Á¬°¡u¹q¸Ü³qª¾¡v¤§¦¬¤å¤é)
   ElseIf strCP10 = "1727" Then
          ii = ii + 1
          strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
              "','»¡©ú¤@','·qÂÐ¡@¶v§½" & Val(Left(strExc(1), 3)) & "¦~" & Mid(strExc(1), 4, 2) & "¤ë" & Right(strExc(1), 2) & "¤é¤§¹q¸Ü³qª¾¡C')"
          str201Detail = str201Detail & "¡@¡@¤@¡B·qÂÐ¡@¶v§½" & Val(Left(strExc(1), 3)) & "¦~" & Mid(strExc(1), 4, 2) & "¤ë" & Right(strExc(1), 2) & "¤é¤§¹q¸Ü³qª¾¡C" & vbCrLf
   'end 2022/09/28
   Else
   '2012/1/5 End
      If Trim(Label12(7)) = "" Then 'µL¾÷Ãö¤å¸¹
         If strCaseDate <> "" Then
            If Len(strCaseDate) = 6 Then strCaseDate = "0" & Trim(strCaseDate)
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','»¡©ú¤@','¥»®×·~©ó" & Val(Left(strCaseDate, 3)) & "¦~" & Mid(strCaseDate, 4, 2) & "¤ë" & Right(strCaseDate, 2) & "¤é´£¥X¥Ó½Ð¦b®×¡C')"
            'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
            'Modified by Lydia 2019/03/05 ¥h±¼¶}ÀY¨â­Ó¥þ§ÎªÅ¥Õ(­ì¥»4­Ó)
            str201Detail = str201Detail & "¡@¡@¤@¡B¥»®×·~©ó" & Val(Left(strCaseDate, 3)) & "¦~" & Mid(strCaseDate, 4, 2) & "¤ë" & Right(strCaseDate, 2) & "¤é´£¥X¥Ó½Ð¦b®×¡C" & vbCrLf
         Else
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','»¡©ú¤@','¥»®×·~©ó¡@¦~¡@¤ë¡@¤é´£¥X¥Ó½Ð¦b®×¡C')"
            'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
            'Modified by Lydia 2019/03/05 ¥h±¼¶}ÀY¨â­Ó¥þ§ÎªÅ¥Õ(­ì¥»4­Ó)
            str201Detail = str201Detail & "¡@¡@¤@¡B¥»®×·~©ó¡@¦~¡@¤ë¡@¤é´£¥X¥Ó½Ð¦b®×¡C" & vbCrLf
         End If
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','»¡©ú¤@','·qÂÐ¡@¶v§½" & ChgSQL(Label12(7)) & "¨ç¡C')"
         'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
         'Modified by Lydia 2019/03/05 ¥h±¼¶}ÀY¨â­Ó¥þ§ÎªÅ¥Õ(­ì¥»4­Ó)
         str201Detail = str201Detail & "¡@¡@¤@¡B·qÂÐ¡@¶v§½" & ChgSQL(Label12(7)) & "¨ç¡C" & vbCrLf
      End If
   End If
   
   strTmp = ""
   t = 1
'   For k = 0 To 2 '»¡©ú¤G~¥|
'      If Check2(k).Value = 1 Then
         'Modify Sindy 2018/7/13 ¥i½Æ¿ï
         't = t + 1
         If Check2(0).Value = 1 Then '¤å¥ó
            t = t + 1
            strTmp = PUB_ChgNumber2Chinese(CStr(t)) & "¡B¸É¥¿¦p¤U¡G" & vbCrLf
            For i = 0 To 8 '6
               If Check1(i).Value = 1 Then
                  j = j + 1
                  If i = 0 Then
                     'Modified by Lydia 2021/11/11 ¨ú®ø¥¿¼v¥»¿ï¶µ¡A¤º®eª½±µ¥Î¢æ¢æ¢æ¤A¥÷
                     'strTmp = strTmp & "¡@¡@¡@¡@" & Replace(Replace(CStr(j) & "." & Check1(i).Caption, " ", ""), "N", IIf(Option7(0).Value = True, "¥¿", "¼v")) & vbCrLf
                     strTmp = strTmp & "¡@¡@¡@¡@" & CStr(j) & "." & Check1(i).Caption & vbCrLf
                  ElseIf i = 1 Then
                     strTmp = strTmp & "¡@¡@¡@¡@" & CStr(j) & "." & "¥D±iÀu¥ýÅv¤§ÃÒ©ú¤å¥ó ¡Ð " & Text8.Text & "¥Ó½Ð®ÑÁÃ¥»¤A¥÷¡]ªþ¤¤Ä¶¤å¡^" & vbCrLf
                  ElseIf i = 2 Then
                     strTmp = strTmp & "¡@¡@¡@¡@" & CStr(j) & "." & "¥Nªí¤H¦WºÙ¡G" & ReadTMData & vbCrLf
                  'Modify By Sindy 2015/3/18
                  ElseIf i = 4 Then
                     'Modified by Lydia 2021/11/11 ¨ú®ø¥¿¼v¥»¿ï¶µ¡A¤º®eª½±µ¥Î¢æ¢æ¢æ¤A¥÷
                     'strTmp = strTmp & "¡@¡@¡@¡@" & Replace(Replace(CStr(j) & "." & Check1(i).Caption, " ", ""), "N", IIf(Option1(0).Value = True, "¥¿", "¼v")) & vbCrLf
                     strTmp = strTmp & "¡@¡@¡@¡@" & CStr(j) & "." & Check1(i).Caption & vbCrLf
                  ElseIf i = 5 Then
                     'Modified by Lydia 2021/11/11 ¨ú®ø¥¿¼v¥»¿ï¶µ¡A¤º®eª½±µ¥Î¢æ¢æ¢æ¤A¥÷
                     'strTmp = strTmp & "¡@¡@¡@¡@" & Replace(Replace(CStr(j) & "." & Check1(i).Caption, " ", ""), "N", IIf(Option2(0).Value = True, "¥¿", "¼v")) & vbCrLf
                     strTmp = strTmp & "¡@¡@¡@¡@" & CStr(j) & "." & Check1(i).Caption & vbCrLf
                  ElseIf i = 6 Then
                     'Modified by Lydia 2021/11/11 ¨ú®ø¥¿¼v¥»¿ï¶µ¡A¤º®eª½±µ¥Î¢æ¢æ¢æ¤A¥÷
                     'strTmp = strTmp & "¡@¡@¡@¡@" & Replace(Replace(CStr(j) & "." & Check1(i).Caption, " ", ""), "N", IIf(Option3(0).Value = True, "¥¿", "¼v")) & vbCrLf
                     strTmp = strTmp & "¡@¡@¡@¡@" & CStr(j) & "." & Check1(i).Caption & vbCrLf
                  '2015/3/18 END
                  Else
                     strTmp = strTmp & "¡@¡@¡@¡@" & CStr(j) & "." & Check1(i).Caption & vbCrLf
                  End If
                  If i = 0 Then Type0V = True
                  If i = 1 Then Type1V = True
                  If i = 4 Then Type4V = True
                  If i = 5 Then Type5V = True
'                  ii = ii + 1
'                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'                       "','¸É¤å¥ó V " & Format(j) & "','" & ChgSQL(strTmp) & "')"
               End If
            Next i
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                         "','»¡©ú" & PUB_ChgNumber2Chinese(CStr(t)) & "','" & ChgSQL(strTmp) & "')"
            'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
            str201Detail = str201Detail & "¡@¡@" & strTmp & vbCrLf
         End If
         If Check2(1).Value = 1 Then '°Ó«~
            t = t + 1
            strTmp = PUB_ChgNumber2Chinese(CStr(t)) & "¡B¸É¥¿°Ó«~¡þªA°È¦p¤U¡G"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                         "','»¡©ú" & PUB_ChgNumber2Chinese(CStr(t)) & "','" & ChgSQL(strTmp) & "')"
            'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
            'Modified by Lydia 2019/03/05 ¥h±¼¶}ÀY¨â­Ó¥þ§ÎªÅ¥Õ(­ì¥»4­Ó)
            'str201Detail = str201Detail & "¡@¡@¡@¡@" & strTmp & vbCrLf
            str201Detail = str201Detail & "¡@¡@" & Replace(strTmp, vbCrLf & "¡@¡@¡@¡@", vbCrLf & "¡@¡@") & vbCrLf
         End If
         If Check2(2).Value = 1 Then '¥Ó½Ð·N¨£®Ñ
            t = t + 1
            'Modify By Sindy 2016/4/25 + Check4(3) : ·~©ó ¦~ ¤ë ¤éÅÜ§ó¨ä¤¤Ä¶¦W¦b®×¡A¬G¥Ó½Ð¤H»P¾Ú¥H®Ö»é°Ó¼Ð¤§°Ó¼ÐÅv¤H¦P¤@¡A¥»®×®Ö»é²z¥Ñ§Y¤£´_¦s¦b¡AÂÔ½Ð¡@¶v§½½ç¬°®Ö­ã¤§³B¤À¡C
            strTmp = PUB_ChgNumber2Chinese(CStr(t)) & "¡BÃö©ó®Ö»é²z¥Ñ³¡¥÷¡A¥Ó½Ð¤H" & _
                                                IIf(Check4(0).Value = 1, Check4(0).Caption, "") & _
                                                IIf(Check4(1).Value = 1, Check4(1).Caption, "") & _
                                                IIf(Check4(2).Value = 1, Check4(2).Caption, "") & _
                                                IIf(Check4(3).Value = 1, Check4(3).Caption, "")
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                         "','»¡©ú" & PUB_ChgNumber2Chinese(CStr(t)) & "','" & ChgSQL(strTmp) & "')"
            'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
            'Modified by Lydia 2019/03/05 ¥h±¼¶}ÀY¨â­Ó¥þ§ÎªÅ¥Õ(­ì¥»4­Ó)
            'str201Detail = str201Detail & "¡@¡@¡@¡@" & strTmp & vbCrLf
            str201Detail = str201Detail & "¡@¡@" & Replace(strTmp, vbCrLf & "¡@¡@¡@¡@", vbCrLf & "¡@¡@") & vbCrLf
         End If
         'Add By Sindy 2018/5/15
         If Check2(4).Value = 1 Then 'ªþ°Ó«~²M³æ
            t = t + 1
            strTmp = PUB_ChgNumber2Chinese(CStr(t)) & "¡B¸É¥¿°Ó«~¡þªA°È¦WºÙ¦p©Òªþ¤§°Ó«~¡þªA°È­×¥¿²M³æ©Ò¥Ü¡C"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                         "','»¡©ú" & PUB_ChgNumber2Chinese(CStr(t)) & "','" & ChgSQL(strTmp) & "')"
            'Added by Lydia 2019/02/21 ¹q¤l°e¥ó-¥Ó½Ð¤º®e
            'Modified by Lydia 2019/03/05 ¥h±¼¶}ÀY¨â­Ó¥þ§ÎªÅ¥Õ(­ì¥»4­Ó)
            'str201Detail = str201Detail & "¡@¡@¡@¡@" & strTmp & vbCrLf
            str201Detail = str201Detail & "¡@¡@" & Replace(strTmp, vbCrLf & "¡@¡@¡@¡@", vbCrLf & "¡@¡@") & vbCrLf
         End If
'         If strTmp <> "" Then
'         '2018/5/15 END
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'                         "','»¡©ú" & PUB_ChgNumber2Chinese(CStr(t)) & "','" & ChgSQL(strTmp) & "')"
'         End If
'      End If
'   Next k
   
   'ªþ¥ó
   If Check3(0).Value = 1 Or Check3(1).Value = 1 Or Check3(2).Value = 1 Or Check3(3).Value = 1 Or _
      Check3(4).Value = 1 Or Check3(5).Value = 1 Or Check3(6).Value = 1 Or Check3(7).Value = 1 Or _
      Check3(8).Value = 1 Or Check3(9).Value = 1 Or Check3(10).Value = 1 Or Check3(11).Value = 1 Then
      j = 0
      strTmp = "ªþ¥ó¡G" & vbCrLf
      For i = 0 To 11
         If Check3(i).Value = 1 Then
            j = j + 1
            If i = 0 Then
               strTmp = strTmp & "¡@¡@" & PUB_ChgNumber2Chinese(CStr(j)) & "¡B" & "³W¶O " & Text9.Text & " ¤¸¾ã" & vbCrLf
            ElseIf i = 8 Then
               'Modified by Lydia 2021/11/11 ¨ú®ø¥¿¼v¥»¿ï¶µ¡A¤º®eª½±µ¥Î¢æ¢æ¢æ¤A¥÷
               'strTmp = strTmp & "¡@¡@" & Replace(Replace(PUB_ChgNumber2Chinese(CStr(j)) & "¡B" & Check3(i).Caption, " ", ""), "N", IIf(Option4(0).Value = True, "¥¿", "¼v")) & vbCrLf
               strTmp = strTmp & "¡@¡@" & PUB_ChgNumber2Chinese(CStr(j)) & "¡B" & Check3(i).Caption & vbCrLf
            ElseIf i = 9 Then
               'Modified by Lydia 2021/11/11 ¨ú®ø¥¿¼v¥»¿ï¶µ¡A¤º®eª½±µ¥Î¢æ¢æ¢æ¤A¥÷
               'strTmp = strTmp & "¡@¡@" & Replace(Replace(PUB_ChgNumber2Chinese(CStr(j)) & "¡B" & Check3(i).Caption, " ", ""), "N", IIf(Option5(0).Value = True, "¥¿", "¼v")) & vbCrLf
               strTmp = strTmp & "¡@¡@" & PUB_ChgNumber2Chinese(CStr(j)) & "¡B" & Check3(i).Caption & vbCrLf
            ElseIf i = 10 Then
               'Modified by Lydia 2021/11/11 ¨ú®ø¥¿¼v¥»¿ï¶µ¡A¤º®eª½±µ¥Î¢æ¢æ¢æ¤A¥÷
               'strTmp = strTmp & "¡@¡@" & Replace(Replace(PUB_ChgNumber2Chinese(CStr(j)) & "¡B" & Check3(i).Caption, " ", ""), "N", IIf(Option6(0).Value = True, "¥¿", "¼v")) & vbCrLf
               strTmp = strTmp & "¡@¡@" & PUB_ChgNumber2Chinese(CStr(j)) & "¡B" & Check3(i).Caption & vbCrLf
            Else
               strTmp = strTmp & "¡@¡@" & PUB_ChgNumber2Chinese(CStr(j)) & "¡B" & Check3(i).Caption & vbCrLf
            End If
         End If
      Next i
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','ªþ¥ó" & "','" & ChgSQL(strTmp) & "')"
   End If
   
   'Add By Sindy 2016/5/31
   If tm(8) = "7" Then '7.ÃÒ©ú¼Ð³¹
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','ÃÒ©ú¼Ð³¹','ÃÒ©ú¼Ð³¹')"
   End If
   '2016/5/31 END
   'Added by Lydia 2019/02/21 (¸É¥¿)¹q¤l°e¥ó¥Ó½Ð®Ñ=>²MªÅ¯È¥»¥Ó½Ð®Ñ
   If bol201CP118 = True Then
        For intI = 0 To ii
            strTxt(ii) = ""
        Next intI
        ii = 0
   End If
   'end 2019/02/21
   
   If Check2(0).Value = 1 And Check2(1).Value = 0 And Check2(2).Value = 0 Then '¸É¥¿¤å¥ó
      'Add By Sindy 2012/11/26
      'bolEmail = PUB_GetEMailFlag(tm(1) & tm(2) & tm(3) & tm(4), , , bolPlusPaper) 'ÀË¬d¬O§_¥HE-Mail³qª¾
      '2012/11/26 End
      ET03_1 = ""
      Select Case m_CP10
         Case "201", "208"
            ' ©w½Z»y¤å
            Select Case m_strLanguage
               ' ­^¤å
               Case "2":
                  Select Case strCaseType
                     Case "©µ®iµù¥U"
                        ET03_1 = "06"
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "01", ET02, ET03_1, strUserNum
                        '¦C¦L³Æµù
                        strPrintNote = ""
                        If Type0V = True Then strPrintNote = "Power of Attorney"
                        If Type4V = True Then
                           If strPrintNote <> "" Then strPrintNote = strPrintNote & " and "
                           strPrintNote = strPrintNote & "documents evidencing the change of the registrant's name"
                        End If
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                     "','¦C¦L³Æµù',' " & ChgSQL(strPrintNote) & "')"
                        cnnConnection.Execute strSql
                        'Add By Sindy 2012/11/26 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                        If bolEmail = True And bolPlusPaper = False Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','Enclosed please find a scanned copy of our request for your records. " & IIf(strDebitNote = "", "Our debit note for services rendered is also attached for your kind settlement.", strDebitNote) & "')"
                           cnnConnection.Execute strSql
                        Else '¶l¥ó
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','" & IIf(strDebitNote = "", "Enclosed please find our debit note for services rendered for your kind settlement.", strDebitNote) & " Copies of our request will be mailed to you with the confirmation copy of this letter for your records.')"
                           cnnConnection.Execute strSql
                        End If
                        '2012/11/26 End
                        
                     Case "µù¥U«eÅÜ§ó", "²¾Âàµn°O"
                        ET03_1 = "04" '"07" 'Modify By Sindy 2011/5/23
                        strReceiveNo = m_CP43: ET02 = m_CP43 'Add By Sindy 2011/5/23
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "01", ET02, ET03_1, strUserNum
                        '¦C¦L³Æµù
                        strPrintNote = ""
                        If Type0V = True Then strPrintNote = "Power of Attorney"
                        If Type5V = True Then
                           If strPrintNote <> "" And Type4V = True Then strPrintNote = strPrintNote & ", "
                           If strPrintNote <> "" And Type4V = False Then strPrintNote = strPrintNote & " and "
                           strPrintNote = strPrintNote & "Deed of Assignment"
                        End If
                        If Type4V = True Then
                           If strPrintNote <> "" Then strPrintNote = strPrintNote & " and "
                           strPrintNote = strPrintNote & "documents evidencing the change of name"
                        End If
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                     "','¦C¦L³Æµù',' " & ChgSQL(strPrintNote) & "')"
                        cnnConnection.Execute strSql
                        If strCaseType = "²¾Âàµn°O" Then
                           'Add By Sindy 2012/11/26 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                           'FCT,01,501,04
                           If bolEmail = True And bolPlusPaper = False Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                       "','¨Ò¥~¤º¤å','Enclosed please find a scanned copy of our request for your records. " & IIf(strDebitNote = "", "Our debit note for services rendered is also attached for your kind settlement.", strDebitNote) & "')"
                              cnnConnection.Execute strSql
                           Else '¶l¥ó
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                       "','¨Ò¥~¤º¤å','" & IIf(strDebitNote = "", "Enclosed please find our debit note for services rendered for your kind settlement.", strDebitNote) & " Copies of our request will be mailed to you with the confirmation copy of this letter for your records.')"
                              cnnConnection.Execute strSql
                           End If
                           '2012/11/26 End
                        Else
                           'Add By Sindy 2012/11/26 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                           'FCT,01,000,04
                           If bolEmail = True And bolPlusPaper = False Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                       "','¨Ò¥~¤º¤å','Enclosed please find a scanned copy of our request for your records. " & IIf(strDebitNote = "", "Our debit note for services rendered is also attached for your kind settlement.", strDebitNote) & "')"
                              cnnConnection.Execute strSql
                           Else '¶l¥ó
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                       "','¨Ò¥~¤º¤å','" & IIf(strDebitNote = "", "Enclosed please find our debit note for services rendered for your kind settlement.", strDebitNote) & " Copies of our request will be mailed to you with the confirmation copy of this letter for your records.')"
                              cnnConnection.Execute strSql
                           End If
                           '2012/11/26 End
                        End If
                        '°Ó¼Ð¸¹¼Æ
                        If Trim(tm(15)) <> "" Then
                           strPrintNote = "Reg. No. : " & Trim(tm(15))
                        Else
                           strPrintNote = "Appl. No:" & Trim(tm(12))
                        End If
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                     "','°Ó¼Ð¸¹¼Æ','" & ChgSQL(strPrintNote) & "')"
                        cnnConnection.Execute strSql
                        'µo¤å¤é
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                     "','µo¤å¤é','" & DBDATE(m_CP27) & "')"
                        cnnConnection.Execute strSql
                        
                     Case Else
                        ET03_1 = "04"
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "01", ET02, ET03_1, strUserNum
                        '¦C¦L³Æµù
                        strPrintNote = ""
                        If Type0V = True Then strPrintNote = "Power of Attorney"
                        If Type1V = True Then
                           If strPrintNote <> "" Then strPrintNote = strPrintNote & " and "
                           strPrintNote = strPrintNote & "the priority document(s)"
                        End If
'                        ' 2009/4/17 ADD BY SONIA§PÂ_¬O§_¦P®É¦³208¸ÉÀu¥ýÅv¤å¥ó
'                        StrSQLa = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & tm(1) & "' AND CP02='" & tm(2) & "' AND CP03='" & tm(3) & "' AND CP04='" & tm(4) & "' AND CP10='208' AND CP27 IS NULL AND CP57 IS NULL "
'                        rsA.CursorLocation = adUseClient
'                        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                        If rsA.RecordCount > 0 Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                    "','¦C¦L³Æµù',' " & ChgSQL(strPrintNote) & "')"
                           cnnConnection.Execute strSql
'                        End If
'                        If rsA.State <> adStateClosed Then rsA.Close
'                        Set rsA = Nothing
'                        '2009/4/17 end
                        'Add By Sindy 2012/11/26 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                        'FCT,01,000,04
                        If bolEmail = True And bolPlusPaper = False Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','Enclosed please find a scanned copy of our request for your records. " & IIf(strDebitNote = "", "Our debit note for services rendered is also attached for your kind settlement.", strDebitNote) & "')"
                           cnnConnection.Execute strSql
                        Else '¶l¥ó
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','" & IIf(strDebitNote = "", "Enclosed please find our debit note for services rendered for your kind settlement.", strDebitNote) & " Copies of our request will be mailed to you with the confirmation copy of this letter for your records.')"
                           cnnConnection.Execute strSql
                        End If
                        '2012/11/26 End
                  End Select
                  
               ' ¤é¤å
               Case "3":
                  ET03_1 = "05"
                  ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                  EndLetter "01", ET02, ET03_1, strUserNum
                  '¦C¦L³Æµù
                  strPrintNote = ""
                  'Modified by Morgan 2023/3/15
                  'If Type0V = True Then strPrintNote = "©e¥ôûì"
                  If Type0V = True Then strPrintNote = PUB_GetUniText(Me.Name, "¦C¦L³Æµù1")
                  'end 2023/3/15
                  If Type1V = True Then 'Àu¥ýÅv
                     'Modified by Morgan 2023/3/15
                     'If strPrintNote <> "" Then strPrintNote = strPrintNote & " ¤ÎÇZ "
                     'strPrintNote = strPrintNote & "Àu¥ý“¸¥D±iÇR¥ÎÆêÇr¤é¥»¥XÄ@µý©ú®Ñ"
                     If strPrintNote <> "" Then strPrintNote = strPrintNote & PUB_GetUniText(Me.Name, "¦C¦L³Æµù2")
                     strPrintNote = strPrintNote & PUB_GetUniText(Me.Name, "¦C¦L³Æµù3")
                     'end 2023/3/15
                     'Add By Sindy 2012/9/12
                     'Modified by Morgan 2023/3/15
                     'strExc(1) = "Àu¥ý“¸µý©ú®ÑÇU¤¤üÂ»y“Õ  1³¡"
                     strExc(1) = PUB_GetUniText(Me.Name, "¦P«Êª«")
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                              "','¦P«Êª«','" & strExc(1) & "')"
                     cnnConnection.Execute strSql
                     '2012/9/12 End
                  End If
'                  ' 2009/4/23 ADD BY SONIA§PÂ_¬O§_¦P®É¦³208¸ÉÀu¥ýÅv¤å¥ó
'                  StrSQLa = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & tm(1) & "' AND CP02='" & tm(2) & "' AND CP03='" & tm(3) & "' AND CP04='" & tm(4) & "' AND CP10='208' AND CP27 IS NULL AND CP57 IS NULL "
'                  rsA.CursorLocation = adUseClient
'                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                  If rsA.RecordCount > 0 Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                              "','¦C¦L³Æµù','" & ChgSQL(strPrintNote) & "')"
                     cnnConnection.Execute strSql
'                  End If
'                  If rsA.State <> adStateClosed Then rsA.Close
'                  Set rsA = Nothing
'                  '2009/4/23 end
                  'Add By Sindy 2017/2/2 + Àu¥ýÅv©Î©e¥ôª¬
                  strPrintNote = ""
                  If Type0V = True And Type1V = False Then '¥u¦³©e¥ôª¬
                     'Modified by Morgan 2023/3/15
                     'strPrintNote = "þ÷þàÇeþêþùÇV¡B«YÇr¸É¥¿®Ñ¤ÎÇZ’U©ÒÇU½Ð¨D®ÑÇy¦P«Ê­PþêÇeþìÇUþú¡Bþç¬dƒBÇUµ{¡B©yþêþâþÝÄ@Æê¥Óþê¤WþåÇeþì¡C"
                     strPrintNote = PUB_GetUniText(Me.Name, "¦C¦L³Æµù4")
                     'end 2023/3/15
                  ElseIf Type1V = True Then 'Àu¥ýÅv
                     'Modified by Morgan 2023/3/15
                     'strPrintNote = "þ÷þàÇeþêþùÇV¡B«YÇr¸É¥¿®Ñ¤ÎÇZÀu¥ý“¸µý©ú®ÑÇU¤¤üÂ»y“Õ¡B¨ÃÇZÇR’U©ÒÇU½Ð¨D®ÑÇy¦P«Ê­PþêÇeþìÇUþú¡Bþç¬dƒBÇUµ{¡B©yþêþâþÝÄ@Æê¥Óþê¤WþåÇeþì¡C"
                     strPrintNote = PUB_GetUniText(Me.Name, "¦C¦L³Æµù5")
                     'end 2023/3/15
                  End If
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & _
                           "','Àu¥ýÅv©Î©e¥ôª¬','" & ChgSQL(strPrintNote) & "')"
                  cnnConnection.Execute strSql
                  '2017/2/2 END
            End Select
      End Select
   End If
   
   If ii <> 0 Then
      If Not ClsLawExecSQL(ii, strTxt) Then
         MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
      End If
   End If
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "¤¤"
         Label12(3) = tm(5)
      Case "­^"
         Label12(3) = tm(6)
      Case "¤é"
         Label12(3) = tm(7)
   End Select
End Sub

'Private Sub Form_Activate()
'Me.Text6.SetFocus
'End Sub

Private Sub Form_Load()
Dim tKind As String 'Added by Lydia 2019/02/21

   MoveFormToCenter Me
   intWhere = °ê¥~_FC
   With frm030206_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      tKind = .Text6.Text   'Added by Lydia 2019/03/26
      strReceiveNo = .Tag
   End With
   ReDim tm(TF_TM)
   ReadTradeMark
   '¥[¥X¦W¥N²z¤H²M³æ¨Ñ¤Ä¿ï
   lstNameAgent.Clear
   'Modified by Lydia 2021/08/04 ¶Ç¤J®×¥ó©Ê½è¡BForm 2.0
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True
   'Added by Lydia 2021/08/04 ¦pªG¤@¶}©l±NListBox©Ô¨ì»Ý­nªº¤j¤p¡A¦r«¬·|¦Û°Ê©ñ¤j¡F©Ò¥Hµe­±¹w³]¬°¤@¦C°ª«×¡AForm_Load¤~©ñ¤j¨ì»Ý­nªº¤j¤p
   lstNameAgent.Height = 1500
   lstNameAgent.Width = 1300
   
   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   If m_CP10 = "208" Then Check2(0).Value = 1
   If m_CP10 = "202" Then Check2(2).Value = 1 'Add By Sindy 2015/3/18 ¥Ó½Ð·N¨£®Ñ
   SSTab1.Tab = 0
   
   'Added by Lydia 2019/02/21 ©w½Z·|¦]¬°¤Ä¿ï¶µ¦ÓÅÜ§ó¤º®e,©Ò¥H(¸É¥¿)¹q¤l°e¥ó¥Ó½Ð®Ñ(frm03020605_1)¨Ö¤J¯È¥»ªºµe­±
   'Modified by Lydia 2019/02/26 +¸ÉÀu¥ýÅvÃÒ©ú208(©w½Z»P201(¸É¥¿)¹q¤l°e¥ó¥Ó½Ð®Ñ¬Û¦P)
   'If m_CP10 = "201" And tKind = "2" Then
   'Modified by Lydia 2019/05/09 +202¥Ó½Ð·N¨£®Ñ
   'Modified by Lydia 2020/12/31 «ü©w¦¬¤å©Ê½è¥H¥~ªºA¡BBÃþ¦¬¤å¡A¬Ò¥i²£¥Í¸É¥¿¥Ó½Ð®Ñ
   'If (m_CP10 = "201" Or m_CP10 = "202" Or m_CP10 = "208") And tKind = "2" Then
   If tKind = "2" Then
         bol201CP118 = True
         Call FormControl(m_CP10)
         'Added by Lydia 2019/03/22 ¹q¤l°e¥ó¥Ó½Ð®Ñ¹w³]Åã¥Ü³W¶O
         If bol201CP118 = True Then
             'Modified by Lydia 2019/07/05 ³W¶O¦³¤d¤À¦ì,·|³y¦¨ÂàÀÉ¿ù»~
             'Text9.Text = Format(Val(m_CP17), "#,##0")
             Text9.Text = Val(m_CP17)
             Check3(0).Value = 1
         End If
         'end 2019/03/22
   End If
   'end 2019/02/21
End Sub

'Added by Lydia 2019/02/21 ±±¨î¶µ¥Ø¤£¥iÂI¿ï
Private Sub FormControl(ByVal iType As String)
'Modified by Lydia 2020/12/31 ¹q¤l°e¥ó²Î¤@§ó¦W
'    Select Case iType
'        'Modified by Lydia 2019/02/26 +¸ÉÀu¥ýÅvÃÒ©ú208
'        'Modified by Lydia 2019/05/17 +202¥Ó½Ð·N¨£®Ñ
'
'        Case "201", "202", "208" '¸É¥¿201,¸ÉÀu¥ýÅvÃÒ©ú208
'              Me.Caption = "¦U¦¡¥Ó½Ð®Ñ-¹q¤l°e¥ó-¸É¥¿"
'              chkAtt1(0).Visible = True
'              'Mark by Lydia 2019/02/21 ªü½¬ªí¥Ü¯È¥»¶µ¥Ø¥þ³¡«O¯d,  ªþ¥ó¦WºÙ«á¸É
''              '°ò¥»¸ê®Æ­¶ÅÒ
''              Check1(2).Enabled = False
''              Check1(3).Enabled = False
''              Check1(6).Enabled = False: Option3(0).Enabled = False: Option3(1).Enabled = False
''              Check1(7).Enabled = False
''              Check1(8).Enabled = False
''              '®Ö»é²z¥Ñ­¶ÅÒ
''              SSTab1.TabVisible(1) = False
''              'ªþ¥ó­¶ÅÒ
''              Check3(0).Value = vbChecked
''              Check3(1).Enabled = False
''              Check3(2).Enabled = False
''              Check3(5).Enabled = False
''              Check3(6).Enabled = False
''              Check3(7).Enabled = False
''              Check3(10).Enabled = False:  Option6(0).Enabled = False: Option6(1).Enabled = False
''              Check3(11).Enabled = False
'              'end 2019/02/21
'    End Select
    Me.Caption = "¦U¦¡¥Ó½Ð®Ñ-¹q¤l°e¥ó-¸É¥¿"
    chkAtt1(0).Visible = True
'end 2020/12/31
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm03020603_1 = Nothing
End Sub

Private Sub ReadTradeMark()
Dim rsTemp1 As New ADODB.Recordset
'Modified by Lydia 2021/08/04
'Dim Lbl As LABEL
Dim Lbl As Object
   
   For Each Lbl In Label12
      Lbl = ""
   Next
   tm(1) = Text1
   tm(2) = Text2
   tm(3) = Text3
   tm(4) = Text4
   If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
      Text5 = tm(11)
      Label12(1) = tm(12)
      Label12(2) = tm(15)
      Label12(3) = tm(5)
   End If
   
   'Modified by Lydia 2019/02/21 FCTµ{§Ç¤À¾÷
   'strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,cp43,cp10,CP06,CP07,CP84,CP110,CP64,CP27,cp17 " & _
      "from caseprogress,casepropertymap,staff,staff staff1 " & _
      "where cp09='" & strReceiveNo & "' " & _
      "AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) " & _
      "and cp13=staff1.st01(+) "
   strExc(0) = "select cpm03,s1.st02 as st1,s2.st02 as st2,cp43,cp10,cp06,cp07,cp84,cp110,cp64,cp27,cp17,s3.st07  " & _
                    "from caseprogress,casepropertymap,staff s1 ,staff s2,staff s3 " & _
                    "where cp09='" & strReceiveNo & "' " & _
                    "and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) " & _
                    "and cp13=s2.st01(+) and s2.st57=s3.st01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      m_CP10 = "" & .Fields("CP10")
      m_CP17 = "" & .Fields("cp17") '¦¬¤å³W¶O
      If Val(m_CP17) > 0 Then Text9.Text = Format(Val(m_CP17), "#,##0")
      If Not IsNull(.Fields(0)) Then
         Label12(0) = .Fields(0) '®×¥ó©Ê½è
      End If
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1) '©Ó¿ì¤H
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2) '´¼Åv¤H­û
      m_F21st07 = "" & .Fields("st07") 'Added by Lydia 2019/02/21 FCTµ{§Ç¤À¾÷
      m_CP64 = "" & .Fields("CP64") 'Add By Sindy 2010/1/21 ¶i«×³Æµù
      'm_CP43 = "" & .Fields("cp43") 'Add By Sindy 2011/5/23 ¬ÛÃöÁ`¦¬¤å¸¹  '2011/7/22 CANCEL BY SONIA §ï¦bStartLetter§ì
      m_CP27 = "" & .Fields("CP27") 'Add By Sindy 2011/5/23 µo¤å¤é´Á
      If Not IsNull(.Fields(3)) Then
         strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields("CP05")) Then Label12(6) = TransDate(rsTemp1.Fields("CP05"), 1) '¨Ó¨ç¦¬¤å¤é
            If Not IsNull(rsTemp1.Fields("CP08")) Then Label12(7) = rsTemp1.Fields("CP08") '¾÷Ãö¤å¸¹
         End If
      End If
      If Not IsNull(.Fields(5)) Then Label12(9) = TransDate(.Fields(5), 1) '¥»©Ò´Á­­
      If Not IsNull(.Fields(6)) Then Label12(10) = TransDate(.Fields(6), 1) 'ªk©w´Á­­
   End If
   End With
   
   'Àu¥ýÅv°ê®a
   strExc(0) = "select na03 from pridate,nation where pd01='" & tm(1) & "' and pd02='" & tm(2) & "' and pd03='" & tm(3) & "' and pd04='" & tm(4) & "' and pd07=na01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      Text8 = "" & .Fields(0)
   End If
   End With
   
   'Added by Lydia 2019/03/22 FCT¦V´¼¼z§½´£¥X¤§¦U¦¡¥Ó½Ð®Ñ¤W¤§¤À¾÷¸¹½X¡A½Ð±N¤é¥»°Ï³]©w¬°011°ê®aÀÉºÞ¨î¤H¤À¾÷
   strExc(0) = "select fa10,st07 from fagent, nation, staff where fa01||fa02='" & ChangeCustomerL(tm(44)) & "' and substr(fa10,1,3)=na01(+) and na55=st01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Left("" & RsTemp.Fields("fa10"), 3) = "011" Then
         m_F21st07 = "" & RsTemp.Fields("st07")
      End If
   End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub
'Remove by Lydia 2019/02/21
'Private Sub Text6_Change()
'   If Check2(0).Value = 1 And Check2(1).Value = 0 And Check2(2).Value = 0 Then '¸É¥¿¤å¥ó
'      textCP27.Enabled = True
'   Else
'      textCP27 = ""
'      textCP27.Enabled = False
'   End If
'End Sub
'end 2019/02/21

'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub
'
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   KeyAscii = Pub_NumAscii(KeyAscii)
'   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 Then
'      KeyAscii = 0
'      Beep
'   End If
'End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
Dim strSqlText As String

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   If lstNameAgent.Visible = True Or _
      Trim(textCP27) <> "" Then
      strSql = " UPDATE CASEPROGRESS SET "
      If lstNameAgent.Visible = True Then
         If strSqlText = "" Then
            strSqlText = " cp110 = " & CNULL(m_CP110)
         Else
            strSqlText = strSqlText & " ,cp110 = " & CNULL(m_CP110)
         End If
      End If
      If Trim(textCP27) <> "" Then
         If strSqlText = "" Then
            strSqlText = " cp27 = " & ChangeTStringToWString(textCP27)
         Else
            strSqlText = strSqlText & " ,cp27 = " & ChangeTStringToWString(textCP27)
         End If
      End If
      strSql = strSql & strSqlText & " WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
      'Added by Lydia 2019/02/21 ¹w³]¬°¹q¤l°e¥ó
      If bol201CP118 = True Then
          'Modified by Morgan 2019/7/17 ¥Ø«eFCT©|¥¼¦Û°Ê¦©´Ú
          'strSql = " UPDATE CASEPROGRESS SET CP118='A' WHERE CP09='" & strReceiveNo & "' AND CP158=0 AND CP118 IS NULL"
          strSql = " UPDATE CASEPROGRESS SET CP118='Y' WHERE CP09='" & strReceiveNo & "' AND CP158=0 AND CP118 IS NULL"
          cnnConnection.Execute strSql
      End If
      'end 2019/02/21
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

'ÀË¬d¨Ã³]©wcp110¸ê®Æ
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 ­û¤u½s¸¹¤w¥i«D¼Æ¦r»Ý°µÂà´«
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Lydia 2021/08/04 §ï¼Ò²Õ
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "¡B" & lstNameAgent.List(ii)
         Cancel = False
      End If
   Next
   If Cancel = True Then
      SSTab1.Tab = 0
      MsgBox "¥X¦W¥N²z¤H¤£¥iªÅ¥Õ¡I", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2)
   End If
End Sub

'Add By Sindy 2010/4/16
Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

'Add By Sindy 2010/4/16
' µo¤å¤é
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' µo¤å¤é¤é´Á¤£¥¿½T
      If CheckIsTaiwanDate(textCP27, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªºµo¤å¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' µo¤å¤é¤é´Á¤£¥i¶W¹L¨t²Î¤é
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         'edit by nick 2004/08/31
         'strMsg = "µo¤å¤é¤£¥i¶W¹L¨t²Î¤é"
         strMsg = "µo¤å¤é¤£¥i¶W¹L¨t²Î¤é¥[¤@¤Ñ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textCP27_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

'Added by Lydia 2019/02/21 ¦U¦¡¥Ó½Ð®Ñ-¹q¤l°e¥ó¥Ó½Ð®Ñ
Private Function StartLetter2(ByVal iET01 As String, ByVal iET03 As String, ByVal iCp09 As String) As Boolean
   Dim strTxt(1 To 30) As String, strTmp As String
   Dim ii As Integer, jj As Integer
   Dim tmpArr1 As Variant, tmpArr2 As Variant 'Added by Lydia 2019/03/27
   
   EndLetter iET01, iCp09, iET03, strUserNum
   
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','¥»©Ò®×¸¹','" & m_CaseNo & "')"
   
   '¥Ó½Ð¤H¸ê®Æ
   'Modified by Lydia 2019/03/22 ²¾¨ìbasPublic
   'Call GetApplTM_EData(iET01, iET03, iCp09, tm(), False)
   'Modified by Lydia 2020/09/29 +®×¥ó©Ê½è
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, tm(), False)
   'Modified by Lydia 2023/11/08 ­ì¥»¹w³]§ì¥Ó½Ð¤H°ò¥»ÀÉ¤§¦a§};²{¦b§ï¦¨¹w³]§ì®×¥ó¥Ó½Ð¤H¸ê®Æ¤§¦a§}
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), False)
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), True)
   
   '¥X¦W¥N²z¤H
   'Modified by Lydia 2019/03/27 §ï¦¨¦@¥Î¼Ò²Õ¨ú±o¸ê®Æ
   strExc(0) = PUB_GetAgentCP110(iCp09, m_CP110, "FCT", "4")
   If strExc(0) <> "" Then
       tmpArr1 = Split(strExc(0), "|")
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
               tmpArr2 = Empty
               tmpArr2 = Split(tmpArr1(jj), ",")
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','¥N²z¤H" & jj + 1 & "-ÃÒ®Ñ¦r¸¹','" & tmpArr2(0) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','¥N²z¤H" & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','¥N²z¤H" & jj + 1 & "-¤¤¤å©m¦W','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
           End If
       Next jj
   End If
   'end 2019/03/27
   
   If iET03 = "03" Then '°ò¥»¸ê®Æªí
        ii = ii + 1
        'FCTµ{§Ç¤À¾÷
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','FCTµ{§Ç¤À¾÷','" & m_F21st07 & "')"
   End If
   
   'Modified by Lydia 2019/02/26 ³B²zª¬ªp04=>10
   'If iET03 = "04" Then '¸É¥ó¥Ó½Ð®Ñ
   If iET03 = "10" Then
        ii = ii + 1
        'Ãº¶Oª÷ÃB
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','Ãº¶Oª÷ÃB','" & Text9.Text & "')"
        
        '¥Ó½Ð¤º®e
        'Modified by Lydia 2019/02/21 ªü½¬ªí¥Ü¥Ó½Ð¤º®e¤ñ·Ó¯È¥»
'        jj = 0
'        strTmp = ""
'        If Check1(0).Value = 1 Then
'            jj = jj + 1
'            If strTmp <> "" Then strTmp = strTmp & vbCrLf
'            strTmp = strTmp & "¡@¡@" & jj & ". ©e¥ô®Ñ(§t¤¤Ä¶¤å)"
'        End If
'        If Check1(1).Value = 1 Then
'            jj = jj + 1
'            If strTmp <> "" Then strTmp = strTmp & vbCrLf
'            strTmp = strTmp & "¡@¡@" & jj & ". Àu¥ýÅvÃÒ©ú¤å¥ó(§t¤¤Ä¶¤å)"
'        End If
'        If Check1(5).Value = 1 Then
'            jj = jj + 1
'            If strTmp <> "" Then strTmp = strTmp & vbCrLf
'            strTmp = strTmp & "¡@¡@" & jj & ". ²¾Âà«´¬ù(§t¤¤Ä¶¤å)"
'        End If
'        If Check1(4).Value = 1 Then
'            jj = jj + 1
'            If strTmp <> "" Then strTmp = strTmp & vbCrLf
'            strTmp = strTmp & "¡@¡@" & jj & ". §ó¦WÃÒ©ú(§t¤¤Ä¶¤å)"
'        End If
        strTmp = str201Detail
        'end 2019/02/21
        If strTmp <> "" Then
              ii = ii + 1
              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','¥Ó½Ð¤º®e1', " & CNULL(ChgSQL(strTmp)) & ")"
        End If
        
        'ªþ°e®Ñ¥ó
        If chkAtt1(0).Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-°ò¥»¸ê®Æªí', '" & m_CaseNo & ".contact.pdf" & "')"
        'Added by Lydia 2019/04/11 ­Y¤£¤Ä¿ï°ò¥»¸ê®Æªí¡A«hªþ¥ó¦WºÙ¡u¥¼ÅÜ§ó¥»®×°ò¥»¸ê®Æ¡v¨Ã¥B¤£¥Î²£¥Í.contactÀÉ®×
        Else
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-°ò¥»¸ê®Æªí', '¥¼ÅÜ§ó¥»®×°ò¥»¸ê®Æ')"
        'end 2019/04/11
        End If
        'Added by Lydia 2022/06/14
        If Check3(1).Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-«ü©w°Ó«~ªA°È­×¥¿²M³æ', '" & m_CaseNo & ".list.pdf" & "')"
        End If
        If Check3(2).Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-«ü©w°Ó«~ªA°È¦W±ø', '" & m_CaseNo & ".gsn.pdf" & "')"
        End If
        'end 2022/06/14
        If Check3(3).Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-©e¥ô®Ñ', '" & m_CaseNo & ".poa.pdf" & "')"
        End If
        If Check3(4).Value = 1 Then
            ii = ii + 1
            'Modified by Lydia 2020/07/16 §ó¦W:¡u.priority.pdf¡v§ï¬°¡u.PRI.pdf¡v
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-Àu¥ýÅvÃÒ©ú¤å¥ó', '" & m_CaseNo & ".PRI.pdf" & "')"
        End If
        'Added by Lydia 2022/06/14
        If Check3(6).Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-¦P·N®Ñ', '" & m_CaseNo & ".consent.pdf" & "')"
        End If
        'end 2022/06/14
        If Check3(8).Value = 1 Then  'Memo by Lydia 2022/06/14 §ó¦WÃÒ©ú=>ÅÜ§óÃÒ©ú
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-§ó¦WÃÒ©ú', '" & m_CaseNo & ".change.pdf" & "')"
        End If
        If Check3(9).Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-²¾Âà«´¬ù', '" & m_CaseNo & ".assignment.pdf" & "')"
        End If
        'Added by Lydia 2022/06/14
        If Check3(10).Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','ªþ¥ó-±ÂÅv«´¬ù', '" & m_CaseNo & ".license.pdf" & "')"
        End If
        'end 2022/06/14
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function


