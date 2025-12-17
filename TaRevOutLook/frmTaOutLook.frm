VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTaOutLook 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¥x¤@¶l¥ó±µ¦¬¨t²Î"
   ClientHeight    =   7670
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   12470
   Icon            =   "frmTaOutLook.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7670
   ScaleWidth      =   12470
   Begin VB.CommandButton cmdStart 
      Caption         =   "±Ò°Ê"
      Height          =   315
      Left            =   5790
      TabIndex        =   30
      Top             =   210
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¤â°Ê¶×¤J¶l¥ó"
      Height          =   330
      Left            =   6780
      TabIndex        =   1
      Top             =   210
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8280
      TabIndex        =   0
      Top             =   210
      Width           =   800
   End
   Begin VB.Frame Frame7 
      Caption         =   "¥[³t inbound  ¤À«H³]©w:"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   2440
      Left            =   9990
      TabIndex        =   75
      Top             =   180
      Width           =   2410
      Begin VB.TextBox txtIPDeptMin 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   78
         Text            =   "5"
         Top             =   1110
         Width           =   350
      End
      Begin VB.TextBox txtIPDeptEDate 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   77
         Text            =   "1141122"
         Top             =   780
         Width           =   890
      End
      Begin VB.TextBox txtIPDeptSDate 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   76
         Text            =   "1141119"
         Top             =   450
         Width           =   890
      End
      Begin VB.Label Label20 
         Caption         =   "¶¡¹j´X¤ÀÄÁ¡G"
         Height          =   200
         Left            =   150
         TabIndex        =   81
         Top             =   1140
         Width           =   1130
      End
      Begin VB.Label Label19 
         Caption         =   "ºI¤î¤é´Á¡G"
         Height          =   200
         Left            =   150
         TabIndex        =   80
         Top             =   810
         Width           =   920
      End
      Begin VB.Label Label16 
         Caption         =   "°_©l¤é´Á¡G"
         Height          =   200
         Left            =   150
         TabIndex        =   79
         Top             =   480
         Width           =   920
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "ºÊ¬Ý±Æµ{®É¶¡"
      Height          =   3040
      Left            =   8580
      TabIndex        =   63
      Top             =   4530
      Width           =   2530
      Begin VB.Label LblMsg 
         Caption         =   " "
         Height          =   220
         Left            =   150
         TabIndex        =   74
         Top             =   2790
         Width           =   2050
      End
      Begin VB.Label Label18 
         Caption         =   "strEndTime¡G"
         Height          =   220
         Left            =   150
         TabIndex        =   73
         Top             =   2300
         Width           =   2050
      End
      Begin VB.Label LblstrEndTime 
         Caption         =   " ~ "
         Height          =   220
         Left            =   480
         TabIndex        =   72
         Top             =   2550
         Width           =   2050
      End
      Begin VB.Label Label17 
         Caption         =   "strStarTime¡G"
         Height          =   220
         Left            =   150
         TabIndex        =   71
         Top             =   1800
         Width           =   2050
      End
      Begin VB.Label LblstrStarTime 
         Caption         =   " ~ "
         Height          =   220
         Left            =   480
         TabIndex        =   70
         Top             =   2050
         Width           =   2050
      End
      Begin VB.Label LblstrChkEndTime 
         Caption         =   " ~ "
         Height          =   220
         Left            =   480
         TabIndex        =   69
         Top             =   1550
         Width           =   2050
      End
      Begin VB.Label Label15 
         Caption         =   "strChkEndTime¡G"
         Height          =   220
         Left            =   150
         TabIndex        =   68
         Top             =   1300
         Width           =   2050
      End
      Begin VB.Label LblstrChkStarTime 
         Caption         =   " ~ "
         Height          =   220
         Left            =   480
         TabIndex        =   67
         Top             =   1050
         Width           =   2050
      End
      Begin VB.Label Label14 
         Caption         =   "strChkStarTime¡G"
         Height          =   220
         Left            =   150
         TabIndex        =   66
         Top             =   800
         Width           =   2050
      End
      Begin VB.Label LblTime 
         Caption         =   " ~ "
         Height          =   220
         Left            =   480
         TabIndex        =   65
         Top             =   550
         Width           =   2050
      End
      Begin VB.Label Label13 
         Caption         =   "¾ã¤é¤À«H°_¨´®É¶¡¡G"
         Height          =   220
         Left            =   150
         TabIndex        =   64
         Top             =   300
         Width           =   2050
      End
   End
   Begin VB.Frame Frame99 
      Height          =   1000
      Left            =   8640
      TabIndex        =   51
      Top             =   2580
      Width           =   2920
      Begin VB.TextBox txtCkSDate 
         Height          =   285
         Left            =   990
         MaxLength       =   7
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtCkEDate 
         Height          =   285
         Left            =   1980
         MaxLength       =   7
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton CmdChkMail 
         BackColor       =   &H008080FF&
         Caption         =   "ÀË®Ö«H¥ó(¸ê®Æ§¨)"
         Height          =   340
         Left            =   60
         Style           =   1  '¹Ï¤ù¥~Æ[
         TabIndex        =   10
         Top             =   180
         Width           =   1510
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1680
         X2              =   2100
         Y1              =   740
         Y2              =   740
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "«H¥ó¤é´Á¡G"
         Height          =   180
         Left            =   60
         TabIndex        =   52
         Top             =   660
         Width           =   900
      End
   End
   Begin VB.Timer TmrLAbackup 
      Left            =   9570
      Top             =   5310
   End
   Begin VB.TextBox TxtIPDept 
      Height          =   285
      Left            =   60
      TabIndex        =   54
      Top             =   600
      Width           =   9890
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   9060
      MaxLength       =   7
      TabIndex        =   49
      Top             =   6780
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ListBox ListErrTxt 
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   10830
      TabIndex        =   29
      Top             =   3660
      Visible         =   0   'False
      Width           =   3350
   End
   Begin SHDocVwCtl.WebBrowser WebBrowserP 
      CausesValidation=   0   'False
      Height          =   1280
      Left            =   11010
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5070
      Width           =   2330
      ExtentX         =   4101
      ExtentY         =   2249
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer TmrTM 
      Left            =   9270
      Top             =   5310
   End
   Begin VB.Timer TmrPatent 
      Left            =   8970
      Top             =   5310
   End
   Begin VB.TextBox txtMRL02 
      Height          =   270
      Left            =   3510
      MaxLength       =   7
      TabIndex        =   15
      Top             =   4440
      Width           =   885
   End
   Begin VB.ComboBox Combo1 
      Height          =   260
      ItemData        =   "frmTaOutLook.frx":0442
      Left            =   720
      List            =   "frmTaOutLook.frx":0455
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   14
      Top             =   4410
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "¬d¸ß±µ¦¬ª¬ªp"
      Height          =   285
      Left            =   4800
      TabIndex        =   11
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Timer tmrClock 
      Left            =   0
      Top             =   120
   End
   Begin VB.Timer TmrFCPout 
      Left            =   9270
      Top             =   5010
   End
   Begin VB.Timer TmrFCPin 
      Left            =   8970
      Top             =   5010
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   450
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   310
      Left            =   0
      TabIndex        =   6
      Top             =   7360
      Width           =   12470
      _ExtentX        =   21996
      _ExtentY        =   547
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frmTaOutLook.frx":0492
      Height          =   2570
      Left            =   30
      TabIndex        =   7
      Top             =   4740
      Width           =   8510
      _ExtentX        =   15011
      _ExtentY        =   4533
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "«H½c|±µ¦¬¤é´Á|°_©l®É¶¡|ºI¤î®É¶¡|·s¼W¤H­û|±µ¦¬µ§¼Æ|¥[±Kµ§¼Æ|­Ó®×µ§¼Æ|°õ¦æª¬ªp"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Left            =   6510
      ScaleHeight     =   310
      ScaleWidth      =   910
      TabIndex        =   28
      Top             =   210
      Visible         =   0   'False
      Width           =   945
   End
   Begin SHDocVwCtl.WebBrowser WebBrowserT 
      CausesValidation=   0   'False
      Height          =   1280
      Left            =   10830
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4860
      Width           =   2330
      ExtentX         =   4101
      ExtentY         =   2249
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.FileListBox File1 
      Height          =   240
      Left            =   8820
      TabIndex        =   47
      Top             =   4410
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD2 
      Bindings        =   "frmTaOutLook.frx":04A7
      Height          =   1640
      Left            =   10950
      TabIndex        =   62
      Top             =   5310
      Width           =   2000
      _ExtentX        =   3528
      _ExtentY        =   2893
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   53
      Text            =   "frmTaOutLook.frx":04BC
      Top             =   30
      Width           =   9820
   End
   Begin VB.Frame Frame5 
      Caption         =   "¡@ªk«ß©Ò LAbackup «H½c"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   560
      Left            =   30
      TabIndex        =   55
      Tag             =   "¡@ªk«ß©Ò LAbackup «H½c"
      Top             =   3330
      Width           =   9945
      Begin VB.TextBox txtPathLAbackup 
         Height          =   270
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "C:\LAbackup"
         Top             =   240
         Width           =   3105
      End
      Begin VB.CommandButton OpenFolder 
         Caption         =   "<="
         Height          =   255
         Index           =   4
         Left            =   4290
         TabIndex        =   57
         Top             =   240
         Width           =   345
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "¤¤Â_"
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   4770
         TabIndex        =   56
         Top             =   180
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label LblLAbackup 
         Appearance      =   0  '¥­­±
         BackColor       =   &H000000C0&
         BorderStyle     =   1  '³æ½u©T©w
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   60
         TabIndex        =   60
         Top             =   30
         Width           =   150
      End
      Begin VB.Label Label12 
         Caption         =   "±H¥ó¸ê®Æ§¨¡G"
         Height          =   195
         Left            =   90
         TabIndex        =   59
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "¡@°Ó¼Ð³B tm «H½c"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   560
      Left            =   30
      TabIndex        =   37
      Tag             =   "¡@°Ó¼Ð³B tm «H½c"
      Top             =   2730
      Width           =   9945
      Begin VB.CommandButton cmdCancel 
         Caption         =   "¤¤Â_"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   4770
         TabIndex        =   40
         Top             =   180
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton OpenFolder 
         Caption         =   "<="
         Height          =   255
         Index           =   3
         Left            =   4290
         TabIndex        =   39
         Top             =   240
         Width           =   345
      End
      Begin VB.TextBox txtPathTM 
         Height          =   270
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "C:\TM"
         Top             =   240
         Width           =   3105
      End
      Begin VB.Label LblTM 
         Appearance      =   0  '¥­­±
         BackColor       =   &H000000C0&
         BorderStyle     =   1  '³æ½u©T©w
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   60
         TabIndex        =   41
         Top             =   30
         Width           =   150
      End
      Begin VB.Label Label11 
         Caption         =   "¦¬¥ó¸ê®Æ§¨¡G"
         Height          =   195
         Left            =   90
         TabIndex        =   42
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "¡@±M§Q³B patent «H½c"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   560
      Left            =   30
      TabIndex        =   31
      Tag             =   "¡@±M§Q³B patent «H½c"
      Top             =   2130
      Width           =   9945
      Begin VB.TextBox txtPathPatent 
         Height          =   270
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "C:\Patent"
         Top             =   240
         Width           =   3105
      End
      Begin VB.CommandButton OpenFolder 
         Caption         =   "<="
         Height          =   255
         Index           =   2
         Left            =   4290
         TabIndex        =   33
         Top             =   240
         Width           =   345
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "¤¤Â_"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   4770
         TabIndex        =   32
         Top             =   180
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label LblPatent 
         Appearance      =   0  '¥­­±
         BackColor       =   &H000000C0&
         BorderStyle     =   1  '³æ½u©T©w
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   60
         TabIndex        =   36
         Top             =   30
         Width           =   150
      End
      Begin VB.Label Label9 
         Caption         =   "¦¬¥ó¸ê®Æ§¨¡G"
         Height          =   195
         Left            =   90
         TabIndex        =   35
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "¡@°ê¥~³¡ backup «H½c"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   560
      Left            =   30
      TabIndex        =   17
      Tag             =   "¡@°ê¥~³¡ backup «H½c"
      Top             =   1530
      Width           =   9945
      Begin VB.CommandButton cmdCancel 
         Caption         =   "¤¤Â_"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4770
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton OpenFolder 
         Caption         =   "<="
         Height          =   255
         Index           =   1
         Left            =   4290
         TabIndex        =   19
         Top             =   240
         Width           =   345
      End
      Begin VB.TextBox txtPathIPDeptOut 
         Height          =   270
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "C:\IPDeptOut"
         Top             =   240
         Width           =   3105
      End
      Begin VB.Label LblFCPout 
         Appearance      =   0  '¥­­±
         BackColor       =   &H000000C0&
         BorderStyle     =   1  '³æ½u©T©w
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   60
         TabIndex        =   20
         Top             =   30
         Width           =   150
      End
      Begin VB.Label Label6 
         Caption         =   "±H¥ó¸ê®Æ§¨¡G"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "¡@°ê¥~³¡ inbound «H½c"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   560
      Left            =   30
      TabIndex        =   2
      Tag             =   "¡@°ê¥~³¡ inbound «H½c"
      Top             =   930
      Width           =   9945
      Begin VB.CommandButton cmdCancel 
         Caption         =   "¤¤Â_"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4740
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtPathIPDept 
         Height          =   270
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "C:\IPDept"
         Top             =   240
         Width           =   3105
      End
      Begin VB.CommandButton OpenFolder 
         Caption         =   "<="
         Height          =   255
         Index           =   0
         Left            =   4290
         TabIndex        =   3
         Top             =   240
         Width           =   345
      End
      Begin VB.Label LblFCPin 
         Appearance      =   0  '¥­­±
         BackColor       =   &H000000C0&
         BorderStyle     =   1  '³æ½u©T©w
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   60
         TabIndex        =   16
         Top             =   30
         Width           =   150
      End
      Begin VB.Label LblCntIPDept 
         Appearance      =   0  '¥­­±
         BorderStyle     =   1  '³æ½u©T©w
         ForeColor       =   &H80000008&
         Height          =   230
         Left            =   5700
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   3950
      End
      Begin VB.Label Label2 
         Caption         =   "¦¬¥ó¸ê®Æ§¨¡G"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   1125
      End
   End
   Begin MSForms.TextBox TextII17 
      Height          =   300
      Left            =   60
      TabIndex        =   50
      Top             =   4080
      Width           =   9890
      VariousPropertyBits=   746604573
      Size            =   "17436;529"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   800
      Left            =   10800
      TabIndex        =   48
      Top             =   3930
      Width           =   2330
      VariousPropertyBits=   -1400879075
      ScrollBars      =   2
      Size            =   "4101;1411"
      Value           =   "FindÂ²Åé¦r"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "±µ¦¬¤é´Á¡G"
      Height          =   200
      Left            =   2550
      TabIndex        =   13
      Top             =   4470
      Width           =   920
   End
   Begin MSForms.TextBox TextBoxT 
      Height          =   620
      Left            =   10860
      TabIndex        =   46
      Top             =   6450
      Width           =   2330
      VariousPropertyBits=   -1400879075
      ScrollBars      =   2
      Size            =   "4101;1085"
      Value           =   "FindÂ²Åé¦r"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBoxP 
      Height          =   560
      Left            =   11010
      TabIndex        =   43
      Top             =   6690
      Width           =   2330
      VariousPropertyBits=   -1400879075
      ScrollBars      =   2
      Size            =   "4101;979"
      Value           =   "FindÂ²Åé¦r"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Appearance      =   0  '¥­­±
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '³æ½u©T©w
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   4590
      TabIndex        =   26
      Top             =   300
      Width           =   150
   End
   Begin VB.Label Label5 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '³æ½u©T©w
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   3240
      TabIndex        =   25
      Top             =   300
      Width           =   150
   End
   Begin VB.Label Label4 
      Appearance      =   0  '¥­­±
      BackColor       =   &H000000C0&
      BorderStyle     =   1  '³æ½u©T©w
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   2190
      TabIndex        =   24
      Top             =   300
      Width           =   150
   End
   Begin VB.Label Label8 
      Caption         =   "ÃC¦â»¡©ú¡G   Timer°±¤î      ¥¿¦b±µ¦¬¶l¥ó      Timer±Ò°Ê¤¤"
      Height          =   200
      Left            =   1320
      TabIndex        =   27
      Top             =   300
      Width           =   4430
   End
   Begin VB.Label Label1 
      Caption         =   "«H½c¡G"
      Height          =   200
      Left            =   120
      TabIndex        =   12
      Top             =   4470
      Width           =   560
   End
   Begin VB.Menu mnuShow 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuDisplay 
         Caption         =   "Åã¥Ü"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "µ²§ô"
      End
   End
End
Attribute VB_Name = "frmTaOutLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'Memo:
'********************************************************************
'2024/2/6 §ï¬°Outlook¶³ºÝª©¤½¥Î¸ê®Æ§¨
'2024/1/18¤U¤È¤À«H¶³ºÝ¼Ò¦¡·|¥d¦í¡C­nÂ÷½u¼Ò¦¡¶}µÛoutlook
'2023/12/29(¤£½T©w¤é´Á¤F) §ï¬°Outlook¶³ºÝª©«H½c
'2023/12/25 ©Ò¤ºOutlook§ï¬°¶³ºÝª©,«H½cÁÙ¬O¦aºÝ®É
'********************************************************************
'Memo By Sindy 2021/5¤ë Form2.0¤w­×§ï
Option Explicit

Const ªk«ß©Ò¤À«H±Ò¥Î¤é As String = 20240520 'Add By Sindy 2024/5/14
Dim bolActived As Boolean
Dim dblPrevRow As Double
'°õ¦æ¹LTimerªº°_¨´®É¶¡
Dim m_RunFCPinStarTime As String, m_RunFCPinEndTime As String, bolFCPinRun As Boolean
Dim m_RunFCPoutStarTime As String, m_RunFCPoutEndTime As String, bolFCPoutRun As Boolean
Dim m_RunPatentStarTime As String, m_RunPatentEndTime As String, bolPatentRun As Boolean
Dim m_RunTMStarTime As String, m_RunTMEndTime As String, bolTMRun As Boolean
Dim m_RunLAbackupStarTime As String, m_RunLAbackupEndTime As String, bolLAbackupRun As Boolean 'Add By Sindy 2024/5/14
Dim bolCancel(0 To 4) As Boolean 'True:¤¤Â_
Dim mlngID As Long
Dim bolUserControl As Boolean '¨Ï¥ÎªÌ¤â°Ê¾Þ§@
Dim m_M51Recver As String 'Pub_GetSpecMan("¹q¸£¤¤¤ß¶l¥óÀË®Ö¤H­û")
'********** OutLook **********
'Modify By Sindy 2023/6/26 ¤@¯ë¦Ó¨¥¡A¨Ï¥Î¤Ó¦hªº¥þ°ìÅÜ¼Æ¨Ã¤£¬O¼gµ{¦¡ªº¤@­Ó¦n²ßºD¡C©Ò¥H¦pªG¥i¯àªº¸Ü¡AÀ³¸ÓºÉ¶q¨Ï¥Î¼Ò²Õ¼h¦¸©Î°Ï°ìÅÜ¼Æ¡A¦]¬°¥L­Ì¥i¥H¤@ª½ªº­«ÂÐ¨Ï¥Î¡C
''Dim olApp As outlook.Application
'Dim olApp As Object
''Dim myNamespace As outlook.NameSpace
'Dim myNamespace As Object
''Dim myFolder As outlook.Folder
'Dim myFolder As Object
''Dim myItems As outlook.Items
'Dim myItems As Object
'2023/6/26 END
Dim mail_ii As Integer
Dim strSocSubject As String
Dim strMailDate As String
Dim strMailTime As String
Dim strSender As String
'********** OutLook end **********
Dim strFileName As String, intMaxItem As Integer
Dim intKeyCnt As Integer, intRunOK As Integer, intCaseOK As Integer
Dim strErrText As String
Dim intErr2147024882 As Integer
Dim m_FormTitle As String
Dim m_strISDPath As String
Dim Cancel_idx As Integer 'Add By Sindy 2019/2/14
'Dim WithEvents eventConn As ADODB.Connection 'Add By Sindy 2023/11/29
'Dim m_SqlLogFile As String 'Add By Sindy 2023/11/29
Dim process_id As Long, m_strProcessTxt As String
'Add By Sindy 2024/5/3 Timer:1¬í(1000),³Ì¤j­È65535
Const dblTmrFCPin As Long = 10000 'FCPin ­n¥ý©ó Patent
Const dblTmrPatent As Long = 20000
Const dblTmrTM As Long = 30000
Const dblTmrLAbackup As Long = 40000 'Add By Sindy 2024/5/14
Const dblTmrFCPout As Long = 60000 '³Ì«á
Dim m_FristStar As Boolean '²Ä¤@¦¸±Ò°Ê
'2024/5/3 END
Dim strExecuTime_01 As String 'Add By Sindy 2025/5/14 IPDept¥[³t¤À«H¥i°õ¦æªº®É¶¡


Private Sub cmdCancel_Click(Index As Integer)
   If Cancel_idx = 99 Then Exit Sub 'Add By Sindy 2023/3/29
   
   bolCancel(Index) = True '¤¤Â_
   Cancel_idx = 99 'Add By Sindy 2019/2/14
   DoEvents
   Exit Sub
End Sub

Private Sub cmdCancel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2023/12/26 ÀË¬d«H¥ó¬O§_¦³¶×¤J¨t²Î¤¤
Private Sub CmdChkMail_Click()
Dim strMRL01 As String, strPath As String
Dim oFileSys As New FileSystemObject, oFolder As Object
Dim fs
Dim oFile As Object
Dim olApp As Object
Dim myItems As Object
   
   If txtCkSDate = "" Then
      MsgBox "«H¥ó°_©l¤é´Á¤£¥iªÅ¥Õ¡I", vbInformation, "¿é¤J¤é´Á¿ù»~"
      txtCkSDate.SetFocus
      Exit Sub
   Else
      If CheckIsTaiwanDate(txtCkSDate, False) = False Then
         MsgBox "½Ð¿é¤J¥Á°ê¤é´Á¤£§t/¡I", vbInformation, "¿é¤J¤é´Á¿ù»~"
         txtCkSDate.SetFocus
         Exit Sub
      End If
   End If
   If txtCkEDate = "" Then
      MsgBox "«H¥ó¨´¤î¤é´Á¤£¥iªÅ¥Õ¡I", vbInformation, "¿é¤J¤é´Á¿ù»~"
      txtCkEDate.SetFocus
      Exit Sub
   Else
      If CheckIsTaiwanDate(txtCkEDate, False) = False Then
         MsgBox "½Ð¿é¤J¥Á°ê¤é´Á¤£§t/¡I", vbInformation, "¿é¤J¤é´Á¿ù»~"
         txtCkEDate.SetFocus
         Exit Sub
      End If
   End If
   If Val(txtCkSDate) > Val(txtCkEDate) Then
      MsgBox "°_©l¤é´Á¤£¥i¤j©ó¨´¤î¤é´Á¡I", vbInformation, "¿é¤J¤é´Á¿ù»~"
      txtCkEDate.SetFocus
      Exit Sub
   End If
   
   'Add By Sindy 2024/5/16 + Or LblLAbackup.BackColor = vbBlue
   If LblFCPin.BackColor = vbBlue Or _
      LblFCPout.BackColor = vbBlue Or _
      LblPatent.BackColor = vbBlue Or _
      LblTM.BackColor = vbBlue Or _
      LblLAbackup.BackColor = vbBlue Then
      MsgBox "¦³«H½c¥¿¦b±µ¦¬«H¥ó¡A¤£¥i°õ¦æ¡I", vbExclamation
      Exit Sub
   End If
   strMRL01 = Trim(InputBox("­nÀË¬d¨º­Ó«H½cªº«H¥ó¬O§_¦³¶×¤J¨t²Î¤¤¡H¡]¥¼¿é¤J¥Nªí©ñ±ó¤£ÀË¬d¤F¡^" & vbCrLf & _
              "«H½c¥N½X:" & MRL01CName2, "­«­n°T®§¡I"))
   If strMRL01 = "" Then
      Exit Sub
   End If
   strMRL01 = Right("0" & strMRL01, 2)
   Select Case strMRL01
      Case Left(IPDept¦¬¥ó§X, 2)
         strPath = txtPathIPDept.Text
'      Case Left(Patent¦¬¥ó§X, 2)
'         strPath = txtPathPatent.Text
'      Case Left(TM¦¬¥ó§X, 2)
'         strPath = txtPathTM.Text
      Case Else
         MsgBox "©|µL¬ÛÃöµ{¦¡!!"
         Exit Sub
   End Select
      
   Set olApp = CreateObject("Outlook.Application")
   Set oFolder = oFileSys.GetFolder(strPath)
   Set fs = CreateObject("Scripting.FileSystemObject")
   If oFolder.files.Count > 0 Then
      For Each oFile In oFolder.files
         Set myItems = olApp.CreateItemFromTemplate(strPath & "\" & oFile.Name)
         Call ReadMailText_File(myItems)
         '¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶×¤J?
         '­Y¦³=§R°£¡C­Y¨S¦³=¤£³B²z,µ¥¤H­û¬d¬Ý
         strSql = "select ii01,ii03 from ipdeptinput" & _
                  " where replace(ii17,'&','') = '" & ChgSQL(Replace(TextBox3, "&", "")) & "'" & _
                  " and ii11 = '" & ChgSQL(strSender) & "'"
         If strSender <> "¥¼¶Ç»¼ªº¥D¦®" Then
            strSql = strSql & _
                  " and ii12 = " & DBDATE(strMailDate) & _
                  " and ii13 = " & Val(Replace(strMailTime, ":", ""))
         End If
         strSql = strSql & " order by ii01 desc,ii03 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.RecordCount = 1 Then
               '§R°£PCºÝÀÉ®×
               Call fs.DeleteFile(txtPathIPDept & "\" & oFile.Name)
               Sleep 1000
               DoEvents
            End If
         Else
            strSql = "select ii01,ii03,ii11,ii12,ii13,ii17 from ipdeptinput" & _
                     " where replace(ii17,'&','') = '" & ChgSQL(Replace(TextBox3, "&", "")) & "'" & _
                     " and ii11 = '" & ChgSQL(strSender) & "'" & _
                     " and ii12 >= " & DBDATE(txtCkSDate) & " and ii12 <= " & DBDATE(txtCkEDate) & _
                     " order by ii01 desc,ii03 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If RsTemp.RecordCount = 1 Then
                  '§R°£PCºÝÀÉ®×
                  Call fs.DeleteFile(txtPathIPDept & "\" & oFile.Name)
                  Sleep 1000
                  DoEvents
               End If
            Else
               strSql = "select ii01,ii03,ii11,ii12,ii13,ii17 from ipdeptinput" & _
                        " where replace(replace(ii17,'&',''),'¡i©¹¨Ó°O¿ý Saved¡j','') = '" & ChgSQL(Replace(TextBox3, "&", "")) & "'" & _
                        " and ii11 = '" & ChgSQL(strSender) & "'" & _
                        " and ii12 >= " & DBDATE(txtCkSDate) & " and ii12 <= " & DBDATE(txtCkEDate) & _
                        " order by ii01 desc,ii03 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If RsTemp.RecordCount = 1 Then
                     '§R°£PCºÝÀÉ®×
                     Call fs.DeleteFile(txtPathIPDept & "\" & oFile.Name)
                     Sleep 1000
                     DoEvents
                  End If
               Else
'               If strSender = "¥¼¶Ç»¼ªº¥D¦®" Then
'                  strSql = "select ii01,ii03,ii11,ii12,ii13,ii17 from ipdeptinput" & _
'                           " where replace(ii17,'&','') = '" & ChgSQL(Replace(TextBox3, "&", "")) & "'" & _
'                           " and ii11 = '" & strSender & "'" & _
'                           " order by ii01 desc,ii03 desc"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     If RsTemp.RecordCount = 1 Then
'                        '§R°£PCºÝÀÉ®×
'                        Call fs.DeleteFile(txtPathIPDept & "\" & oFile.Name)
'                        Sleep 1000
'                        DoEvents
'      '               Else
'      '                  MsgBox txtPathIPDept & "\" & oFile.Name
'                     End If
'                  End If
'               End If
               End If
            End If
         End If
      Next
      Set oFolder = oFileSys.GetFolder(strPath)
      If oFolder.files.Count > 0 Then
         MsgBox "ÀË¬d§¹²¦¡I"
      End If
   End If
   
   Set olApp = Nothing
   Set oFolder = Nothing
   Set fs = Nothing
End Sub

Private Sub cmdExit_Click()
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then 'Run°õ¦æÀÉ
      If MsgBox("½T©w­nÃö³¬¥x¤@¶l¥ó±µ¦¬¨t²Î¡H", vbExclamation + vbYesNo + vbDefaultButton2, "­«­n°T®§¡I") = vbNo Then
         Exit Sub
      End If
   End If
   'Add By Sindy 2025/11/4
   If strUserNum = "" Then
      End
   End If
   '2025/11/4
   Call cmdCancel_Click(0)
   Call cmdCancel_Click(1)
   Call cmdCancel_Click(2)
   Call cmdCancel_Click(3)
   Call cmdCancel_Click(4) 'Add By Sindy 2024/5/15
   cmdExit.Tag = "¥¿±`µ²§ô"
   IsClose
End Sub

Private Sub ConnectDB(bolStarTimer As Boolean)
On Error GoTo ErrHand
   
   strProvider = cOraProvider 'Added by Sindy 2021/4/12 §ï¥ÎOLEDBª«¥ó
   Forms(0).StatusBar1.Panels(1).Text = "³s½u¤¤..."
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then 'Run°õ¦æÀÉ
      'Sleep 60000 'Modify By Sindy 2024/5/7 ¤À«H¨t²ÎRun°_¨Ó«á,¥ý°±¸m1¤ÀÄÁ,¦A±Ò°Ê¤À«HªºTimer
      For intI = 1 To 30
         Sleep 1000 'Modify By Sindy 2024/5/7 ¤À«H¨t²Î°_¨Ó«á,·|°±¸m30¬í,¦A±Ò°Ê¤À«HªºTimer
         Text2.Text = "¤À«H¨t²Î°_¨Ó«á·|°±¸m30¬í,¦A±Ò°Ê¤À«HªºTimer¡C(¬í¼Æ¡G" & intI & ")"
         DoEvents
      Next intI
      Text2.Text = "¤À«H¨t²Î°_¨Ó«á·|°±¸m30¬í,¦A±Ò°Ê¤À«HªºTimer¡C"
      DoEvents
      
      'If fConnect() = False Then
      If ConnectToServer_1 = False Then
         Call OpenNeweMail(m_M51Recver, PUB_GetDbTerminal & "¥x¤@¶l¥ó±µ¦¬¨t²Î³s¤£¤W¸ê®Æ®w¡A½Ð¾¨³t¦Ü(" & UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) & ")¬d¬Ý¡I", "¦P¥D¦®")
         End
      Else
         PUB_SetSystemVar '³]©w¨t²ÎÅÜ¼Æ
         If UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then 'Run°õ¦æÀÉ¥B¬°«D¥¿¦¡¸ê®Æ®w®É,µ²§ôµ{¦¡
            MsgBox "«D¥¿¦¡¸ê®Æ®w¡A¤£¥i¶i¤J¦¹§@·~¡I", vbCritical
            End
         End If
      End If
      DoEvents
      
      'Add By Sindy 2024/8/23 ²Ä¤@¦¸±Ò°Ê
      If m_FristStar = True Then
         Call OpenNeweMail(Pub_GetSpecMan("¹q¸£¤¤¤ß¶l¥óÀË®Ö¤H­û"), PUB_GetDbTerminal & "¥x¤@¶l¥ó±µ¦¬¨t²Î¡A¤w­«·s±Ò°Ê¡I(" & UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) & ")", "¦P¥D¦®")
         m_FristStar = False
      End If
      '2024/8/23 END
   Else
      If PUB_Connect2DB() = False Then
         Call OpenNeweMail(m_M51Recver, PUB_GetDbTerminal & "¥x¤@¶l¥ó±µ¦¬¨t²Î³s¤£¤W¸ê®Æ®w¡A½Ð¾¨³t¦Ü(" & UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) & ")¬d¬Ý¡I", "¦P¥D¦®")
         End
      'Add By Sindy 2024/5/7
      Else
         bolStarTimer = True
      '2024/5/7 END
      End If
   End If
   Forms(0).StatusBar1.Panels(1).Text = "¤w³s½u..."
   strSrvDate(1) = ServerDate
   strSrvDate(2) = strSrvDate(1) - 19110000
   
   pub_HostName = PUB_ReadHostName '­n°O¿ý¹q¸£¦WºÙ§_«h±H«H·|¥¢±Ñ
   Forms(0).Caption = m_FormTitle & " " & PUB_GetDbTerminal & " (" & _
                      ChangeTStringToTDateString(strSrvDate(2)) & " " & Format(ServerTime, "##:##:##") & ")"
   
   m_M51Recver = Pub_GetSpecMan("¹q¸£¤¤¤ß¶l¥óÀË®Ö¤H­û")
   'Add By Sindy 2018/7/12
   If UCase(pub_DbTerminalName) <> UCase(¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ) Then
      m_strISDPath = PUB_Getdesktop
   Else
      m_strISDPath = Pub_GetSpecMan("°ê¥~³¡¶}©Ý¤À«H¹q¤lÀÉ¦s©ñ¸ô®|")
   End If
   '2018/7/12 END
   If ClsPDSetUserData(strUserNum, strUserName, strGroup) = False Then
      End
   End If
   g_strWriteSysLogFilePath = App.path & "\TaOutLookLog\" & pub_DbTerminalName & "TaOutLook.log" '±ý°O¿ýLogªº§¹¾ã¸ô®|¤ÎÀÉ¦W Add By Sindy 2018/5/28
   
   tmrClock.Interval = 1000
   'Add By Sindy 2017/10/30
   If bolStarTimer = True Then
   '2017/10/30 END
      Call StartMailTimer 'Modify By Sindy 2024/12/20
'      TmrFCPin.Interval = dblTmrFCPin
'      TmrFCPout.Interval = dblTmrFCPout
'      TmrPatent.Interval = dblTmrPatent
'      TmrTM.Interval = dblTmrTM
'      'Add By Sindy 2024/5/14
'      If strSrvDate(1) >= ªk«ß©Ò¤À«H±Ò¥Î¤é Then
'         TmrLAbackup.Interval = dblTmrLAbackup
'      End If
'      '2024/5/14 END
   End If
   
   'Åª¨ú¸ê®Æ§¨¹w³]¸ô®|
   If PUB_GetLastDate(Me.Name, strUserNum & "PATHFCPin") <> "" Then
      txtPathIPDept = PUB_GetLastDate(Me.Name, strUserNum & "PATHFCPin")
   End If
   If PUB_GetLastDate(Me.Name, strUserNum & "PATHFCPout") <> "" Then
      txtPathIPDeptOut = PUB_GetLastDate(Me.Name, strUserNum & "PATHFCPout")
   End If
   If PUB_GetLastDate(Me.Name, strUserNum & "PATHPatent") <> "" Then
      txtPathPatent = PUB_GetLastDate(Me.Name, strUserNum & "PATHPatent")
   End If
   If PUB_GetLastDate(Me.Name, strUserNum & "PATHTm") <> "" Then
      txtPathTM = PUB_GetLastDate(Me.Name, strUserNum & "PATHTm")
   End If
   'Add By Sindy 2024/5/15
   If PUB_GetLastDate(Me.Name, strUserNum & "PATHLAbackup") <> "" Then
      txtPathLAbackup = PUB_GetLastDate(Me.Name, strUserNum & "PATHLAbackup")
   End If
   '2024/5/15 END
   
   '±N©Ò­n©w¸qªºÄæ¦ì¼Æ¤@¦¸§ì»ô****start
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from patent where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_PA = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from trademark where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_TM = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from lawcase where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_LC = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from hirecase where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_HC = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from servicepractice where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_SP = AdoRecordSet3.Fields.Count
   CheckOC3
   '***end
   
   txtMRL02 = strSrvDate(2)
   Call cmdQuery_Click
   Exit Sub
   
ErrHand:
   If Err.Number <> 0 Then
      WLog Err.Number & " : " & Err.Description & vbCrLf
   End If
End Sub

'³]©wUser Data¦ÜSession
Private Function ClsPDSetUserData(ByRef strUserNum As String, ByRef strUserName As String, ByRef strGroup As String) As Boolean
Dim lngRt As Long, strUser As String * 100, a As String
Dim strSql As String, rsRecordset As New ADODB.Recordset

On Error GoTo ErrHand
'lngRt = WNetGetUser("", strUser, 10)
'lngRt = 0
'If lngRt = 0 Then
   strUserNum = "QPGMR"
   'strUserNum = "74001"
   strSql = "select st04,st02,st11 from staff where upper(st01)=" + CNULL(strUserNum)
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      If rsRecordset.Fields(0) = "1" Then
         strSql = "begin " + _
            "select st02,st03,st05,st11 into user_data.user_name,user_data.user_department," + _
            "user_data.user_level,user_data.user_group from staff where upper(st01)=" + CNULL(strUserNum) + ";" + _
            "user_data.user_num:=" + CNULL(strUserNum) + ";" + _
            "end;"
         cnnConnection.Execute strSql
         strUserName = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         strGroup = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
         ClsPDSetUserData = True
      Else
         ShowMsg MsgText(9165)
      End If
   Else
      ShowMsg MsgText(9166)
   End If
   rsRecordset.Close
'Else
'   ShowMsg MsgText(9167)
'End If
Exit Function
ErrHand:
   'edit by nickc 2007/02/02
   'ErrorLog
   MsgBox Err.Description
End Function

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2024/12/20
Private Sub StartMailTimer()
   TmrFCPin.Interval = dblTmrFCPin: LblFCPin.BackColor = vbGreen
   TmrFCPout.Interval = dblTmrFCPout: LblFCPout.BackColor = vbGreen
   TmrPatent.Interval = dblTmrPatent: LblPatent.BackColor = vbGreen
   TmrTM.Interval = dblTmrTM: LblTM.BackColor = vbGreen
   'Add By Sindy 2024/5/14
   If strSrvDate(1) >= ªk«ß©Ò¤À«H±Ò¥Î¤é Then
      TmrLAbackup.Interval = dblTmrLAbackup: LblLAbackup.BackColor = vbGreen
   End If
   '2024/5/14 END
End Sub
Private Sub CloseMailTimer()
   TmrFCPin.Interval = 0: LblFCPin.BackColor = vbRed
   TmrFCPout.Interval = 0: LblFCPout.BackColor = vbRed
   TmrPatent.Interval = 0: LblPatent.BackColor = vbRed
   TmrTM.Interval = 0: LblTM.BackColor = vbRed
   'Add By Sindy 2024/5/14
   If strSrvDate(1) >= ªk«ß©Ò¤À«H±Ò¥Î¤é Then
      TmrLAbackup.Interval = 0: LblLAbackup.BackColor = vbRed
   End If
   '2024/5/14 END
End Sub
'2024/12/20 END

Private Sub ClearTimer()
   tmrClock.Interval = 0
   Call CloseMailTimer 'Add By Sindy 2024/12/20
'   TmrFCPin.Interval = 0
'   TmrFCPout.Interval = 0
'   TmrPatent.Interval = 0
'   TmrTM.Interval = 0
'   TmrLAbackup.Interval = 0 'Add By Sindy 2024/5/14
End Sub

'¤H­û­n¤â°Ê±µ¦¬¶l¥ó®É¶·ÀË¬d
'¦^¶ÇTrue:¥¿¦b±µ¦¬¤¤
'   False:µL,¥iRun
Private Function ChkMailReceiving(strMRL01) As Boolean
   ChkMailReceiving = False '¹w³]¥¼°õ¦æ
   'ÀË¬d¬O§_¦³¥¿¦b°õ¦æ¤¤ªºTimer
   strSql = "select mrl01,mrl02,mrl03,mrl04,mrl05 from MailReceiveLog" & _
            " where mrl01='" & strMRL01 & "'" & _
            " and mrl02=" & strSrvDate(1) & _
            " and mrl09='Y'" & _
            " order by mrl03 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkMailReceiving = True
      MsgBox "¦¹«H½c¥¿¦b±µ¦¬¤¤¡A¤£¥i¾Þ§@¡I", vbInformation
      '¬O§_¦³¤w¹L¤@¤p®É©|¥¼µ²§ô,«h³qª¾¹q¸£¤¤¤ß¤H­û
      If Val(RsTemp.Fields("mrl03")) + 10000 <= Format(Time, "HHMMSS") Then
         PUB_SendMail strUserNum, m_M51Recver, "", "¦³¤â°Ê±µ¦¬«H½c(" & strMRL01 & ")¥¿¦b°õ¦æ¤¤,¤w¤@¤p®É©|¥¼µ²§ô,¬O§_¦³²§±`¡A½Ð¬d¬Ý¡I", _
            "mrl03=" & RsTemp.Fields("mrl03") & vbCrLf & _
            "mrl04=" & RsTemp.Fields("mrl04") & vbCrLf & _
            "mrl05=" & RsTemp.Fields("mrl05") & GetPrjSalesNM(RsTemp.Fields("mrl05")), , , , , , , , , , , False, , , False, , , False
         DoEvents
      End If
   End If
End Function

'µ¹¨Ï¥ÎªÌ¤â°Ê¶×¤J¶l¥ó
Public Function userControlFCPin(Optional mbolCancel As Boolean = False) As Boolean
   Call ClearTimer
   If mbolCancel = True Then '¤¤Â_
      bolCancel(0) = True '¤¤Â_
      DoEvents
   Else
      userControlFCPin = False
      If ChkMailReceiving(Left(IPDept¦¬¥ó§X, 2)) = True Then
         Exit Function
      End If
      If MsgBox("½T©w¬O§_­n¶×¤J" & "IPDept_" & °ê¥~³¡¦¬¥ó«H½c & "¡H", vbExclamation + vbYesNo + vbDefaultButton2, "­«­n°T®§¡I") = vbNo Then
         Exit Function
      End If
      bolFCPinRun = True
      bolUserControl = True '¨Ï¥ÎªÌ¤â°Ê¾Þ§@
      userControlFCPin = True
      'If importFCPinBound = True Then
      Call ChkExecutionTimer(Left(IPDept¦¬¥ó§X, 2))
      Unload Me
      'End If
   End If
End Function

Private Sub cmdStart_Click()
   If MsgBox("½T©w­n±Ò°Ê±µ¦¬«H½c¶l¥ó¶Ü¡H", vbExclamation + vbYesNo + vbDefaultButton2, "­«­n°T®§¡I") = vbYes Then
      Call StartMailTimer 'Modify By Sindy 2024/12/20
'      TmrFCPin.Interval = dblTmrFCPin
'      TmrFCPout.Interval = dblTmrFCPout
'      TmrPatent.Interval = dblTmrPatent
'      TmrTM.Interval = dblTmrTM
'      'Add By Sindy 2024/5/14
'      If strSrvDate(1) >= ªk«ß©Ò¤À«H±Ò¥Î¤é Then
'         TmrLAbackup.Interval = dblTmrLAbackup
'      End If
'      '2024/5/14 END
      
      'Add By Sindy 2020/10/5
      If ConnectToServer_1 = False Then
         Call OpenNeweMail(m_M51Recver, PUB_GetDbTerminal & "¥x¤@¶l¥ó±µ¦¬¨t²Î³s¤£¤W¸ê®Æ®w¡A½Ð¾¨³t¦Ü(" & UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) & ")¬d¬Ý¡I", "¦P¥D¦®")
         MsgBox PUB_GetDbTerminal & "¥x¤@¶l¥ó±µ¦¬¨t²Î³s¤£¤W¸ê®Æ®w¡A½Ð­«·s±Ò°Ê¡I", vbInformation
         cmdStart.Enabled = False
         Exit Sub
      Else '³s½u¤¤
         strSrvDate(1) = ServerDate
         strSrvDate(2) = strSrvDate(1) - 19110000
      End If
      '2020/10/5 END
   End If
End Sub

Private Sub Form_Activate()
'   Screen.MousePointer = vbHourglass
   If bolActived = False Then
      Me.Top = (Screen.Height - Me.Height) / 2
      Me.Left = (Screen.Width - Me.Width) / 2
      If cnnConnection.State = adStateClosed Then '§å¦¸§@·~,¥ý³s½u
         Call ClearTimer
         If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or UCase(pub_DbTerminalName) <> UCase(¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ) Then 'Run VB
'            If MsgBox("½T©w­n¶×¤J«H½c¶l¥ó¶Ü¡H", vbExclamation + vbYesNo + vbDefaultButton2, "­«­n°T®§¡I") = vbYes Then
'               Call ConnectDB(True)
'            Else
               Call ConnectDB(False)
'            End If
         Else
            Call ConnectDB(True)
         End If
      End If
      bolMailFailNoAlert = True 'Add by Sindy 2014/3/5 ±H«H³£¤£­n¼u¿ù»~°T®§
      'Ãö³¬¶s Âê x ÅÜ¦Ç¦â
      DisableControl frmTaOutLook
      bolActived = True
      
      '¼W¥[¥[³t¤À«H¥\¯à:­pºâ¤U¤@­Ó¥i°õ¦æªº®É¶¡
      If ((Val(strSrvDate(2)) >= Val(txtIPDeptSDate) And Val(txtIPDeptSDate) > 0) And _
          (Val(strSrvDate(2)) <= Val(txtIPDeptEDate) And Val(txtIPDeptEDate) > 0)) And _
         Val(txtIPDeptMin) > 0 Then
         strExecuTime_01 = Format(Time, "hhmmss")
      Else
         strExecuTime_01 = ""
      End If
      '2025/5/14 END
   End If
'   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Processes
   
   m_FristStar = True '²Ä¤@¦¸±Ò°Ê Add By Sindy 2024/8/23
   MoveFormToCenter Me
   m_FormTitle = Me.Caption
   
   'If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or UCase(pub_DbTerminalName) <> UCase(¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ) Then
      For i = 0 To 4 '3
         cmdCancel(i).Visible = True
      Next i
   'End If
   If mlngID = 0 Then mlngID = AddToSystemTray(Picture1.hWnd, WM_MOUSEMOVE, Me.Icon, Me.Caption)
   
'   'Add By Sindy 2019/4/18
'   If Dir(App.path & "\executeTM.txt") <> "" Then
'      WebBrowserT.Navigate App.path & "\executeTM.txt"
'      DoEvents
'      TextBoxT = Replace(Replace(WebBrowserT.Document.Body.innerhtml, "<PRE>", ""), "</PRE>", "")
'   Else
'      TextBoxT = ""
'   End If
'   '2019/4/18 END
'
'   'Add By Sindy 2017/11/23
'   If Dir(App.path & "\executePatent.txt") <> "" Then
'      WebBrowserP.Navigate App.path & "\executePatent.txt"
'      DoEvents
'      TextBoxP = Replace(Replace(WebBrowserP.Document.Body.innerhtml, "<PRE>", ""), "</PRE>", "")
'   Else
'      TextBoxP = ""
'   End If
'   '2017/11/23 END
   
   'Add By Sindy 2024/2/7
   'If PUB_CheckIsRunning("TaRevOutLook.EXE") = True Then
   Set Processes = Interaction.GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='" & App.EXEName & ".exe'")
   Me.Tag = ""
   If Processes.Count > 1 Then
      MsgBox "¥x¤@¶l¥ó±µ¦¬¨t²Î¤w¶}±Ò¤¤¡A¤£¥i­«ÂÐ¡I" & vbCrLf & vbCrLf & _
             "¡]­Y­n­«¶}¡A½Ð¥ý±N«e¤@­Óµ{¦¡Ãö³¬¡A¦A¾Þ§@¡^", vbExclamation
      Me.Tag = "­«ÂÐ"
      Unload Me
   End If
   '2024/2/7 END
   
   pub_OS = GetVersion32 'Add By Sindy 2024/4/24
End Sub

Private Sub Form_Resize()
   If Me.WindowState = "1" Then Me.Visible = False
End Sub

'Add By Sindy 2025/5/13 ¾ã§åµo³qª¾«H
Private Sub BatchSendNoticMail()
   If ((Val(strSrvDate(2)) >= Val(txtIPDeptSDate) And Val(txtIPDeptSDate) > 0) And _
       (Val(strSrvDate(2)) <= Val(txtIPDeptEDate) And Val(txtIPDeptEDate) > 0)) And _
      Val(txtIPDeptMin) > 0 Then
      Call TaRevOutLookBatchSendMail("01", True, True)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2024/2/7
   If Me.Tag <> "­«ÂÐ" Then '¥x¤@¶l¥ó±µ¦¬¨t²Î¬O§_­«ÂÐ¶}±Ò
   '2024/2/7 END
      'Add By Sindy 2025/11/4 ¼W¥[if§PÂ_¤£µM PUB_SaveLastDate ·|¿ù
      If strUserNum <> "" Then 'DB¦³³s½u¦¨¥\
      '2025/11/4 END
         Call BatchSendNoticMail 'Add By Sindy 2025/5/13 ¾ã§åµo³qª¾«H
         If bolUserControl = False Then
            'Àx¦s¸ê®Æ§¨¹w³]¸ô®|
            PUB_SaveLastDate Me.Name, strUserNum & "PATHFCPin", txtPathIPDept.Text
            PUB_SaveLastDate Me.Name, strUserNum & "PATHFCPout", txtPathIPDeptOut.Text
            PUB_SaveLastDate Me.Name, strUserNum & "PATHPatent", txtPathPatent.Text
            PUB_SaveLastDate Me.Name, strUserNum & "PATHTm", txtPathTM.Text
            PUB_SaveLastDate Me.Name, strUserNum & "PATHLAbackup", txtPathLAbackup.Text 'Add By Sindy 2024/5/15
         End If
      End If
      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then 'Run°õ¦æÀÉ
         'Modify By Sindy 2024/8/23
'         PUB_SendMail strUserNum, m_M51Recver, "", "[³qª¾] ¥x¤@¶l¥ó±µ¦¬¨t²Î¡A¤wÃö³¬¡I(" & UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) & ")" & _
'                     IIf(cmdExit.Tag <> "¥¿±`µ²§ô", "(½T»{«H¥ó¬O§_¦³§¹¾ã±µ¦¬¦Ü¨t²Î¤¤)", ""), "¦P¥D¦®" & vbCrLf & vbCrLf & _
'                     IIf(cmdExit.Tag <> "¥¿±`µ²§ô", "ª`·N¡G¡Õµ{¦¡¦³»~¡Ö­«·s¶}Ãö¨t²Î¡A¥²¶·ÀË¬d¦³°ÝÃDªº«e«á«H¥ó¡A" & vbCrLf & _
'                     "½T»{«H¥ó¬O§_¦³§¹¾ã±µ¦¬¦Ü¨t²Î¤¤¡C", ""), , , , , , , , , , , False, , , False, , , False
         Call OpenNeweMail(m_M51Recver, "[³qª¾] ¥x¤@¶l¥ó±µ¦¬¨t²Î¡A¤wÃö³¬¡I(" & UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) & ")" & _
                     IIf(cmdExit.Tag <> "¥¿±`µ²§ô", "(½T»{«H¥ó¬O§_¦³§¹¾ã±µ¦¬¦Ü¨t²Î¤¤)", ""), "¦P¥D¦®" & vbCrLf & vbCrLf & _
                     IIf(cmdExit.Tag <> "¥¿±`µ²§ô", "ª`·N¡G¡Õµ{¦¡¦³»~¡Ö­«·s¶}Ãö¨t²Î¡A¥²¶·ÀË¬d¦³°ÝÃDªº«e«á«H¥ó¡A" & vbCrLf & _
                     "½T»{«H¥ó¬O§_¦³§¹¾ã±µ¦¬¦Ü¨t²Î¤¤¡C", ""))
   '      DoEvents
      End If
      cmdExit.Tag = ""
   End If
   
   Set frmTaOutLook = Nothing
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim MSG As Long

If Me.ScaleMode = 1 Then
   MSG = x / Screen.TwipsPerPixelX
Else
  
End If
Select Case MSG
      Case WM_MOUSEMOVE '²¾°Ê·Æ¹«
          'Label1.Caption = "¥¿¦b²¾°Ê·Æ¹«"
      Case WM_LBUTTONDBLCLK '³sÂI·Æ¹«¥ªÁä
          'Label1.Caption = "³sÂI·Æ¹«¥ªÁä"
          Me.WindowState = "0"
          Me.Visible = True
      Case WM_LBUTTONDOWN '«ö¤U·Æ¹«¥ªÁä
          'Label1.Caption = "«ö¤U·Æ¹«¥ªÁä"
      Case WM_LBUTTONUP '©ñ¶}·Æ¹«¥ªÁä
          'Label1.Caption = "©ñ¶}·Æ¹«¥ªÁä"
      Case WM_RBUTTONDBLCLK '³sÂI·Æ¹«¥kÁä
          'Label1.Caption = "³sÂI·Æ¹«¥kÁä"
      Case WM_RBUTTONDOWN '«ö¤U·Æ¹«¥kÁä
          'Label1.Caption = "«ö¤U·Æ¹«¥kÁä"
          Me.PopupMenu mnuShow, vbPopupMenuLeftAlign + vbPopupMenuRightButton
      Case WM_RBUTTONUP '©ñ¶}·Æ¹«¥kÁä
          ''Label1.Caption = "©ñ¶}·Æ¹«¥kÁä"
End Select
End Sub

Private Sub OpenFolder_Click(Index As Integer)
   Dim Shl As Object, Fd As Object
   Set Shl = CreateObject("Shell.Application")
   Set Fd = Shl.BrowseForFolder(hWnd, "½Ð¿ï¨ú¸ê®Æ§¨", 0, "C:\")
   If Not Fd Is Nothing Then
      If Index = 0 Then txtPathIPDept.Text = Fd.Items.Item.path
      If Index = 1 Then txtPathIPDeptOut.Text = Fd.Items.Item.path
      If Index = 2 Then txtPathPatent.Text = Fd.Items.Item.path
      If Index = 3 Then txtPathTM.Text = Fd.Items.Item.path
      If Index = 4 Then txtPathLAbackup.Text = Fd.Items.Item.path 'Add By Sindy 2024/5/15
   End If
   Exit Sub
   
'Dim stFileName As String
'
'On Error GoTo ErrHnd
'
'   stFileName = "*.msg"
'   With CommonDialog1
'      .CancelError = True
'      .FileName = stFileName
'      .Filter = "msgÀÉ®× (*.msg)|*.msg"
'      If Index = 0 Then .InitDir = IIf(txtPathIPDept <> "", txtPathIPDept, PUB_Getdesktop)
'      If Index = 1 Then .InitDir = IIf(txtPathIPDeptOut <> "", txtPathIPDeptOut, PUB_Getdesktop)
'      .MaxFileSize = 5000
'      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
'      .ShowOpen
'      If .FileName <> "" Then
'         If Index = 0 Then txtPathIPDept.Text = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
'         If Index = 1 Then txtPathIPDeptOut.Text = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
'      End If
'   End With
'   Exit Sub
'ErrHnd:
'   If Err.Number <> 32755 Then
'      MsgBox Err.Description
'   End If
End Sub

Private Sub tmrClock_Timer()
Dim intDel As Integer
Dim strFileName As String
Dim strMailDate As String
Dim strMailTime As String
Dim bolLogMailOnlyOne As Boolean
Dim strR005006 As String
Dim rsA As New ADODB.Recordset
Dim strToCC As String 'Add By Sindy 2018/9/18
Dim strTo As String 'Add By Sindy 2019/9/10
Dim strAttachPath As String 'Add By Sindy 2020/3/31
Dim intFcnt As Integer 'Add By Sindy 2020/3/31
Dim ii As Integer
'Add By Sindy 2023/6/26
'Dim olApp As Object
'Dim myNamespace As Object
'Dim myItems As Object
'Dim myDelFolder As Object
'Dim myFolder As Object
'2023/6/26 END
   
'   'Add By Sindy 2024/5/3 ²Ä¤@¦¸±Ò°Ê, §ï¬°¤@¤ÀÄÁ«á¦A±Ò°Ê
'   If m_FristStar = False Then
'      TmrFCPin.Interval = dblTmrFCPin
'      TmrFCPout.Interval = dblTmrFCPout
'      TmrPatent.Interval = dblTmrPatent
'      TmrTM.Interval = dblTmrTM
'
'      m_FristStar = True
'   End If
'   '2024/5/3 END
   
   StatusBar1.Panels.Item(2).Text = Time
'   If Not (Weekday(Format(strSrvDate(1), "####-##-##")) >= 2 And Weekday(Format(strSrvDate(1), "####-##-##")) <= 6) Then
'      If cnnConnection.State = adStateClosed Then Exit Sub '«D¤u§@¤Ñ¤£¥Î³s½u
'   End If
   
   '±j­¢Â_½u
   'If (Format(Time, "HHMMSS") >= "010000" And Format(Time, "HHMMSS") < "090000") Then '²M±á1~9ÂIÂ_½u
   'Modified by Lydia 2019/11/08 ²M±á0~1ÂIÂ_½u(by David) =Â_½u1¤p®É+«e«á¤£¤À«H¥b¤p®É
   'If (Format(Time, "HHMMSS") >= "010000" And Format(Time, "HHMMSS") < "050000") Then '²M±á1~5ÂIÂ_½u
   'Modify By Sindy 2024/5/3 ­ì²M±á0~1ÂIÂ_½u; ¦A¤Á¥X¥b¤p®Éµ¹Outlook­«·s±Ò°Ê, §ï¬° ²M±á0~12:30ÂIÂ_½u
   'If (Format(Time, "HHMMSS") >= "000000" And Format(Time, "HHMMSS") < "010000") Then '²M±á0~1ÂIÂ_½u
   If (Format(Time, "HHMMSS") >= "000000" And Format(Time, "HHMMSS") < "003000") Then '²M±á0~00:30ÂIÂ_½u
      'Add By Sindy 2024/5/16 + And LblLAbackup.BackColor <> vbBlue
      If cnnConnection.State = adStateOpen And _
         LblFCPin.BackColor <> vbBlue And _
         LblFCPout.BackColor <> vbBlue And _
         LblPatent.BackColor <> vbBlue And _
         LblTM.BackColor <> vbBlue And _
         LblLAbackup.BackColor <> vbBlue Then
         
         Call BatchSendNoticMail 'Add By Sindy 2025/5/13 ¾ã§åµo³qª¾«H
         
         Forms(0).StatusBar1.Panels(1).Text = "±j­¢Â_½u..."
         cnnConnection.Close
         WLog Format(Time, "HHMMSS") & " : ±j­¢Â_½u..."
         g_LetterDebug = False 'Modify By Sindy 2025/11/10 ¨ú®ø°O¿ýLog
         
         'Add By Sindy 2024/4/23
         'Outlook¤£¯à°ÊµL¦^À³~ ³o¦¸§â¤À«H¨t²Î­«¶}, Outlook¨S°Ê;¤À«H®É·|¥X²{
         '  -2147418107:Automation ¿ù»~
         '  ¦b°T®§¿z¿ï¾¹¸Ì®É¤£¥i¹ï¥~©I¥s¡C
         'Ãö³¬Outlook
         process_id = Shell("taskkill /F /IM outlook.exe", vbHide)
         For ii = 1 To 10
            If PUB_CheckIsRunning("outlook.exe") = True Then
               Sleep 1000
            Else
               Exit For
            End If
         Next
'         '¶}±ÒOutlook
'         process_id = Shell("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE", vbHide)
'         For ii = 1 To 10
'            If PUB_CheckIsRunning("outlook.exe") = True Then
'               Exit For
'            Else
'               Sleep 1000
'            End If
'         Next
'         'Mark:¦]DBÂ_½u¤£¯à±H«H
'         'PUB_SendMail strUserNum, m_M51Recver, "", "¡iOutlook­«·s±Ò°Ê¡j" & Time, "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
'         m_strProcessTxt = "Outlook¤w­«·s±Ò°Ê!!!"
'         WLog m_strProcessTxt 'Add By Sindy 2024/4/27
'         '2024/4/23 END
         
         Call CloseMailTimer 'Add By Sindy 2024/12/20
'         TmrFCPin.Interval = 0 '¬õ¿OTimer¤w°±¤î
'         TmrFCPout.Interval = 0 '¬õ¿OTimer¤w°±¤î
'         TmrPatent.Interval = 0 '¬õ¿OTimer¤w°±¤î
'         TmrTM.Interval = 0 '¬õ¿OTimer¤w°±¤î
'         TmrLAbackup.Interval = 0 '¬õ¿OTimer¤w°±¤î Add By Sindy 2024/5/14
         
         Exit Sub '³o¬q®É¶¡¥ð®§,¤£¶·°õ¦æµ{¦¡
      End If
   
   'Add By Sindy 2024/5/3 ¤Á¥X¥b¤p®Éµ¹Outlook­«·s±Ò°Ê
   ElseIf (Format(Time, "HHMMSS") >= "003000" And Format(Time, "HHMMSS") < "010000") Then '²M±á00:30~1:00
      'ÀË¬d¬O§_¦³Outlook¶}±Ò¤¤, ¨S¦³­«·s±Ò°Ê
      If PUB_CheckIsRunning("outlook.exe") = False Then
         '¶}±ÒOutlook
         process_id = Shell("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE", vbHide)
         For ii = 1 To 10
            If PUB_CheckIsRunning("outlook.exe") = True Then
               Exit For
            Else
               Sleep 1000
            End If
         Next
         'Mark:¦]DBÂ_½u¤£¯à±H«H
         'PUB_SendMail strUserNum, m_M51Recver, "", "¡iOutlook­«·s±Ò°Ê¡j" & Time, "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
         m_strProcessTxt = "Outlook¤w­«·s±Ò°Ê!!!"
         WLog m_strProcessTxt 'Add By Sindy 2024/4/27
         '2024/4/23 END
      End If
      
   'µo²{Table¸ê®Æ®wµ²ºcÅÜ,¥²¶·Â_½u¦A­«·s³s½u,¤£µM·|¦³¿ù»~
   '²M±á01:00Â_½u5:00¦A³s½u
   'ElseIf (Format(Time, "HHMMSS") >= "050000" And Format(Time, "HHMMSS") < "060000") Then
   Else 'If (Format(Time, "HHMMSS") >= "090000" And Format(Time, "HHMMSS") < "093000") Then
      'Memo by Lydia 2019/11/08 ²M±á0~1ÂIÂ_½u
      If cnnConnection.State = adStateClosed Then
         Forms(0).StatusBar1.Panels(1).Text = "³s½u¸ê®Æ®w..."
         WLog Format(Time, "HHMMSS") & " : ³s½u¸ê®Æ®w..."
         '¦A³s½u
         Call ConnectDB(True)
         WLog Format(Time, "HHMMSS") & " : ¤w³s½u..."
         
         'Add By Sindy 2017/4/13 ­«·s³s½u±ý­«·s°õ¦æ·í¤é±Æµ{,¦]¦¹­n²MªÅÅÜ¼Æ­È
         m_RunFCPinStarTime = "": m_RunFCPinEndTime = ""
         m_RunFCPoutStarTime = "": m_RunFCPoutEndTime = ""
         m_RunPatentStarTime = "": m_RunPatentEndTime = ""
         m_RunTMStarTime = "": m_RunTMEndTime = ""
         m_RunLAbackupStarTime = "": m_RunLAbackupEndTime = "" 'Add By Sindy 2024/5/16
         ListErrTxt.Clear
         '2017/4/13 END
         
         'Add By Sindy 2024/4/23
         If m_strProcessTxt <> "" Then
            If PUB_CheckIsRunning("outlook.exe") = False Then
               WLog "¡iPUB_CheckIsRunning °»´ú µLOutlook Running¡j" 'Add By Sindy 2024/4/27
               PUB_SendMail strUserNum, m_M51Recver, "", "¡iPUB_CheckIsRunning °»´ú µLOutlook Running¡j" & Time, "½ÐÀË¬d¤À«H¥D¾÷ª¬ªp¬°¦ó?", , , , , , , , , , , False, , , False, , , False
            Else
               WLog "PUB_CheckIsRunning(outlook.exe) = True: °»´ú¨ì Outlook Running" 'Add By Sindy 2024/4/27
            End If
            m_strProcessTxt = ""
         End If
         '2024/4/23 END
      End If
      g_LetterDebug = True 'Modify By Sindy 2025/11/10 ¨ú®ø°O¿ýLog
      
      'Add By Sindy 2017/9/4
      '±HLog Mailµ¹¶l¥óºÞ²z¤H­û
      'Mark by Lydia 2019/11/12 ¥ýÁôÂÃ
'      If (Format(Time, "HHMMSS") >= "010000" And Format(Time, "HHMMSS") < "010030") And bolLogMailOnlyOne = False Then
'         bolLogMailOnlyOne = True '¤@¤Ñ¥u±H¤@¦¸
'         '±HLog«H¥ó
'         strFileName = App.path & "\TaOutLookLog\" & pub_DbTerminalName & "TaOutLook.log"
'         If Dir(strFileName) <> "" Then
'            'Call OpenNeweMail(m_M51Recver, PUB_GetDbTerminal & "«H¥ó¶×¤Jª¬ªp³qª¾¡F½Ð¬d¬ÝLog...", "¦P¥D¦®", strFileName)
'            PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "«H¥ó¶×¤Jª¬ªp³qª¾¡F½Ð¬d¬ÝLog...", "¦P¥D¦®", , strFileName, , , , , , , , , False, , , , , , False
'            DoEvents
'            'Kill strFileName
'         End If
'      Else
'         bolLogMailOnlyOne = False '¬°±±¨î¤@¤Ñ¥u±H¤@¦¸
'      End If
      '2017/9/4 END
   End If
   
   'If (Format(Time, "HHMMSS") > "091000" And Format(Time, "HHMMSS") < "200000") And
   'If (Format(Time, "HHMMSS") > "063000" And Format(Time, "HHMMSS") < "173000") And
   If (Format(Time, "HHMMSS") > "063000" And Format(Time, "HHMMSS") < "183000") And _
      cnnConnection.State = adStateClosed Then
      '°õ¦æ®É¬q¤¤­YÂ_½u­n³qª¾¹q¸£¤¤¤ß¬ÛÃö¤H­û
      If Me.Tag = "" Then '±±¨î¥uµo¤@¦¸Mail
         Call OpenNeweMail(m_M51Recver, "¥x¤@¶l¥ó±µ¦¬¨t²Î³s¤£¤W¸ê®Æ®w¡A½Ð¾¨³t¦Ü" & UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) & "¬d¬Ý¡I", "¦P¥D¦®")
         Me.Tag = "sendmail"
      End If
   ElseIf cnnConnection.State = adStateOpen Then
      Me.Tag = ""
   End If
   
   '*******************************************************************************************
   '±ß¤W10:00¶}©l²MªÅ[§R°£ªº¶l¥ó]
   '*******************************************************************************************
   'Modified by Lydia 2019/11/08 §ï¨ì±ß¤W11:45~11:55
   'If (Format(Time, "HHMMSS") >= "220000" And Format(Time, "HHMMSS") < "223000") Then
   If (Format(Time, "HHMMSS") >= "234500" And Format(Time, "HHMMSS") < "235500") Then
      '±HLog Mailµ¹¶l¥óºÞ²z¤H­û
      'Modified by Lydia 2019/11/08 §ï¨ì±ß¤W11:45~11:55
      'If (Format(Time, "HHMMSS") >= "220000" And Format(Time, "HHMMSS") < "220030") And bolLogMailOnlyOne = False Then
      If (Format(Time, "HHMMSS") >= "234500" And Format(Time, "HHMMSS") < "234530") And bolLogMailOnlyOne = False Then
         bolLogMailOnlyOne = True '¤@¤Ñ¥u±H¤@¦¸
         
         'Add By Sindy 2025/11/10 ·hÀÉ§ó¦W:¨C¤é°O¿ýªºLog
         strExc(8) = App.path & "\" & App.EXEName & "_Debug.log"
         strExc(9) = App.path & "\TaOutLookLog\" & App.EXEName & "_Debug_" & strSrvDate(2) & ".log"
         If Dir(strExc(8)) <> "" Then
            FileCopy strExc(8), strExc(9)
            If Dir(strExc(9)) <> "" Then
               Kill strExc(8)
            End If
         End If
         '2025/11/10 END
         
         '*******************************************************************
         'Add By Sindy 2017/7/31 ²£¥ÍLog¤å¦rÀÉ
         '*******************************************************************
         'Modify By Sindy 2019/9/5 and R005003<>'ipdept' : ¨ú®ø
         '   ID='" & strUserNum & "' => R005005='¨t²ÎLog°O¿ý,¤£¥i§R°£'
         strExc(0) = "select R005002,R005004,R005003,R005007,R005006,R005008 from R100101" & _
                     " where R005005='¨t²ÎLog°O¿ý,¤£¥i§R°£' and (instr(R005003,'¶À¬ü¬Ã')=0 and instr(R005003,'¹Q©y¬À')=0)" & _
                     " order by R005006 asc,R005008 asc,R005004 asc"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modify By Sindy 2020/3/31 «Ø¥ß¤é´Á¸ê®Æ§¨,¤è«K«á­±±H«H¨Ï¥Î
            strAttachPath = App.path & "\TaOutLookLog\" & strSrvDate(2)
            If Dir(strAttachPath, vbDirectory) = "" Then
               MkDir strAttachPath
            End If
            '2020/3/31 END
            
            rsA.MoveFirst
            strR005006 = "": strToCC = ""
            Do While Not rsA.EOF
               If strR005006 <> "" And strR005006 <> "" & rsA.Fields("R005006") Then
                  strFileName = strAttachPath & "\" & pub_DbTerminalName & "TaOutLook_" & strR005006 & "-" & strSrvDate(2) & ".log"
                  If Dir(strFileName) <> "" Then
                     'Add By Sindy 2019/9/10
                     If strR005006 = "QPGMR" Then 'Modify By Sindy 2025/10/14 §ï¨Ï¥ÎQPGMR
                        strTo = Pub_GetSpecMan("°ê¥~³¡Âà«H¥~±M¸s²Õ") & ";" & Pub_GetSpecMan("°ê¥~³¡Âà«H¥~°Ó¸s²Õ")
                        PUB_SendMail strUserNum, strTo, "", "(" & GetPrjSalesNM(strR005006) & ") ±H¥X¶l¥óµLªk¦Û°ÊÂk¤J¨÷©v°Ï¡F½Ð½T»{¬O§_¬°­Ó®×; ­Y¬O, ½Ð¥DºÞ·þ¾É²Õ­û°È¥²­n¦b¶l¥ó¥D¦®¿é¤J¥¿½T¤§¥»©Ò®×¸¹®æ¦¡¦p:  Our Ref:FCP-xxxxxx", "¦P¥D¦®", , strFileName _
                        , , , , strToCC, , , , , False, m_M51Recver, , False, , , False
                     Else
                        strTo = strR005006
                        PUB_SendMail strUserNum, strTo, "", "(" & GetPrjSalesNM(strR005006) & ") ±H¥X¶l¥óµLªk¦Û°ÊÂk¤J¨÷©v°Ï¡F½Ð½T»{¬O§_¬°­Ó®×; ­Y¬O, ½Ð¥DºÞ·þ¾É²Õ­û°È¥²­n¦b¶l¥ó¥D¦®¿é¤J¥¿½T¤§¥»©Ò®×¸¹®æ¦¡¦p:  Our Ref:FCP-xxxxxx", "¦P¥D¦®", , strFileName _
                        , , , , strToCC, , , , , False, , , False, , , False
                     End If
                     '2019/9/10 END
                  End If
               End If
               WLog_Day "==>¦¬¨ì¤é´Á:" & rsA.Fields("R005002") & " " & rsA.Fields("R005004") & vbCrLf & _
                        "==>±H¥óªÌ:" & rsA.Fields("R005003") & vbCrLf & _
                        "==>¥D¦®:" & rsA.Fields("R005007") & vbCrLf, "" & rsA.Fields("R005006") & "-", False, _
                        strAttachPath & "\"
               strR005006 = "" & rsA.Fields("R005006")
               If strR005006 <> "" & rsA.Fields("R005008") Then
                  strToCC = "" & rsA.Fields("R005008")
               Else
                  strToCC = ""
               End If
'               'Add By Sindy 2018/9/18 David­n¤@¦P±Hµ¹²Õ­û
'               If strR005006 = "77015" Then
'                  If "" & rsA.Fields("R005006") <> "" Then
'                     'Modify By Sindy 2018/10/1 David:­×§ï¬°¦¬¥óªÌ±H¤@¦¸´N¦n
'                     If InStr(strTo, rsA.Fields("R005006")) = 0 Then
'                     '2018/10/1 END
'                        strTo = strTo & ";" & rsA.Fields("R005006")
'                     End If
'                  End If
'               Else
'                  strTo = ""
'               End If
'               '2018/9/18 END
               rsA.MoveNext '*****
            Loop
            rsA.Close
            If strR005006 <> "" Then
               strFileName = strAttachPath & "\" & pub_DbTerminalName & "TaOutLook_" & strR005006 & "-" & strSrvDate(2) & ".log"
               If Dir(strFileName) <> "" Then
                  'Add By Sindy 2019/9/10
                  If strR005006 = "QPGMR" Then 'Modify By Sindy 2025/10/14 §ï¨Ï¥ÎQPGMR
                     strTo = Pub_GetSpecMan("°ê¥~³¡Âà«H¥~±M¸s²Õ") & ";" & Pub_GetSpecMan("°ê¥~³¡Âà«H¥~°Ó¸s²Õ")
                     PUB_SendMail strUserNum, strTo, "", "(" & GetPrjSalesNM(strR005006) & ") ±H¥X¶l¥óµLªk¦Û°ÊÂk¤J¨÷©v°Ï¡F½Ð½T»{¬O§_¬°­Ó®×; ­Y¬O, ½Ð¥DºÞ·þ¾É²Õ­û°È¥²­n¦b¶l¥ó¥D¦®¿é¤J¥¿½T¤§¥»©Ò®×¸¹®æ¦¡¦p:  Our Ref:FCP-xxxxxx", "¦P¥D¦®", , strFileName _
                     , , , , strToCC, , , , , False, m_M51Recver, , False, , , False
                  Else
                     strTo = strR005006
                     PUB_SendMail strUserNum, strR005006, "", "(" & GetPrjSalesNM(strR005006) & ") ±H¥X¶l¥óµLªk¦Û°ÊÂk¤J¨÷©v°Ï¡F½Ð½T»{¬O§_¬°­Ó®×; ­Y¬O, ½Ð¥DºÞ·þ¾É²Õ­û°È¥²­n¦b¶l¥ó¥D¦®¿é¤J¥¿½T¤§¥»©Ò®×¸¹®æ¦¡¦p:  Our Ref:FCP-xxxxxx", "¦P¥D¦®", , strFileName _
                     , , , , strToCC, , , , , False, , , False, , , False
                  End If
                  '2019/9/10 END
               End If
            End If
            '¨S¤ñ¹ï¨ì¥DºÞªºLog¸ê®Æ
            strFileName = strAttachPath & "\" & pub_DbTerminalName & "TaOutLook_-" & strSrvDate(2) & ".log"
            If Dir(strFileName) <> "" Then
               'Modify By Sindy 2019/9/6 ¥ý³qª¾David
               'Modify By Sindy 2019/9/9 §ï³qª¾ Pub_GetSpecMan("°ê¥~³¡Âà«H¥~±M¸s²Õ") & ";" & Pub_GetSpecMan("°ê¥~³¡Âà«H¥~°Ó¸s²Õ")
               PUB_SendMail strUserNum, Pub_GetSpecMan("°ê¥~³¡Âà«H¥~±M¸s²Õ") & ";" & Pub_GetSpecMan("°ê¥~³¡Âà«H¥~°Ó¸s²Õ"), "", "±H¥X¶l¥óµLªk¦Û°ÊÂk¤J¨÷©v°Ï¡F½Ð½T»{¬O§_¬°­Ó®×; ­Y¬O, ½Ð¥DºÞ·þ¾É²Õ­û°È¥²­n¦b¶l¥ó¥D¦®¿é¤J¥¿½T¤§¥»©Ò®×¸¹®æ¦¡¦p:  Our Ref:FCP-xxxxxx", "¦P¥D¦®", , strFileName, , , , , , , , , False, , , False, , , False
'               DoEvents
            End If
            
            'Modify By Sindy 2020/3/31 è©°Æ©Òªø­n¤@¥÷¨S¦³Âk¨÷ªº²M³æ
            File1.path = strAttachPath
            File1.Refresh
            strFileName = ""
            For intFcnt = 0 To File1.ListCount - 1
               If UCase(Right(File1.List(intFcnt), 4)) = ".LOG" And _
                  InStr(File1.List(intFcnt), strSrvDate(2)) > 0 Then
                  strFileName = strFileName & "*" & strAttachPath & "\" & File1.List(intFcnt)
               End If
            Next intFcnt
            If strFileName <> "" Then
               strFileName = Mid(strFileName, 2)
               PUB_SendMail strUserNum, "81040", "", "±H¥X¶l¥óµLªk¦Û°ÊÂk¤J¨÷©v°Ï¡F½Ð½T»{¬O§_¬°­Ó®×; ­Y¬O, ½Ð¥DºÞ·þ¾É²Õ­û°È¥²­n¦b¶l¥ó¥D¦®¿é¤J¥¿½T¤§¥»©Ò®×¸¹®æ¦¡¦p:  Our Ref:FCP-xxxxxx", "¦P¥D¦®", , strFileName, , , , , , , , , False, , , False, , , False
            End If
            '2020/3/31 END
         End If
         'Add By Sindy 2017/7/31 ²M°£°O¿ýLog
         'Modify By Sindy 2019/9/6 ID='" & strUserNum & "' => R005005='¨t²ÎLog°O¿ý,¤£¥i§R°£'
         strSql = "delete from R100101 where R005005='¨t²ÎLog°O¿ý,¤£¥i§R°£'"
         cnnConnection.Execute strSql
         '2017/7/31 END
         'Add By Sindy 2022/10/12 ²M°£ ¨t²Î¦¬¥ó°Ï ©Î ¹q¤l¦¬¤å µo³qª¾«H¥¼µo°e¥X¥hªº¸ê®Æ, ¦]¬°¤w¹L®É®Ä
         strSql = "delete from CaseUseMemo where cum05 in('02','03')"
         cnnConnection.Execute strSql
         '2022/10/12 END
      End If
      
'      '*******************************************************************
'      '²MªÅ[§R°£ªº¶l¥ó]
'      '*******************************************************************
'      Set olApp = CreateObject("Outlook.Application")
'      Set myNamespace = olApp.GetNamespace("MAPI")
'      Set myDelFolder = myNamespace.GetDefaultFolder(3) 'olFolderDeletedItems.3.[§R°£ªº¶l¥ó] ¸ê®Æ§¨
'      Set myItems = myDelFolder.Items
'      For intDel = myItems.Count To 1 Step -1
'         'If myItems.Item(intDel).MessageClass <> "IPM.Note.SMIME" Then 'IPM.Note.SMIME ¥[±K
'         'Modify By Sindy 2017/11/17 ¹J¨ì¥[±K«H¥ó¨ç¼Æ·|¿ù
'         'Modify By Sindy 2020/4/10 + IPM.Outlook.Recall
'         If InStr(UCase(myItems.Item(intDel).MessageClass), UCase("IPM.Note.SMIME")) = 0 And _
'            InStr(UCase(myItems.Item(intDel).MessageClass), UCase("IPM.Outlook.Recall")) = 0 Then
'         'If myItems.Item(intDel).Class = 43 Then
'         '2017/11/17 END
'            myItems.Item(intDel).Delete
'         End If
'      Next intDel
'      Set myItems = Nothing
'      Set myDelFolder = Nothing
'      Set myNamespace = Nothing
'      Set olApp = Nothing
   Else
      bolLogMailOnlyOne = False '¬°±±¨î¤@¤Ñ¥u±H¤@¦¸
   End If
   '*******************************************************************************************
   
   'ÃC¦â: vbBlue, vbGreen, vbRed
   'Modify By Sindy 2024/5/15
   If LblFCPin.BackColor <> vbBlue _
      And LblFCPout.BackColor <> vbBlue _
      And LblPatent.BackColor <> vbBlue _
      And LblTM.BackColor <> vbBlue _
      And LblLAbackup.BackColor <> vbBlue Then
   '2024/5/15 END
   
      If Frame1.Caption = Frame1.Tag Then '¬O§_±µ¦¬¤¤
         If TmrFCPin.Interval > 0 Then
            LblFCPin.BackColor = vbGreen 'ºñ¿OTimer±Ò°Ê¤¤
         Else
            LblFCPin.BackColor = vbRed '¬õ¿OTimer¤w°±¤î
         End If
      End If
      If Frame2.Caption = Frame2.Tag Then '¬O§_±µ¦¬¤¤
         If TmrFCPout.Interval > 0 Then
            LblFCPout.BackColor = vbGreen 'ºñ¿OTimer±Ò°Ê¤¤
         Else
            LblFCPout.BackColor = vbRed '¬õ¿OTimer¤w°±¤î
         End If
      End If
      If Frame3.Caption = Frame3.Tag Then '¬O§_±µ¦¬¤¤
         If TmrPatent.Interval > 0 Then
            LblPatent.BackColor = vbGreen 'ºñ¿OTimer±Ò°Ê¤¤
         Else
            LblPatent.BackColor = vbRed '¬õ¿OTimer¤w°±¤î
         End If
      End If
      If Frame4.Caption = Frame4.Tag Then '¬O§_±µ¦¬¤¤
         If TmrTM.Interval > 0 Then
            LblTM.BackColor = vbGreen 'ºñ¿OTimer±Ò°Ê¤¤
         Else
            LblTM.BackColor = vbRed '¬õ¿OTimer¤w°±¤î
         End If
      End If
      'Add By Sindy 2024/5/14
      If Frame5.Caption = Frame5.Tag Then '¬O§_±µ¦¬¤¤
         If TmrLAbackup.Interval > 0 Then
            LblLAbackup.BackColor = vbGreen 'ºñ¿OTimer±Ò°Ê¤¤
         Else
            LblLAbackup.BackColor = vbRed '¬õ¿OTimer¤w°±¤î
         End If
      End If
      '2024/5/14 END
   End If
   
   Set rsA = Nothing
   IsClose 'µ²§ô
End Sub

Private Sub IsClose()
   'Add By Sindy 2024/5/16 + And LblLAbackup.BackColor <> vbBlue
   If cmdExit.Tag = "¥¿±`µ²§ô" And _
      LblFCPin.BackColor <> vbBlue And _
      LblFCPout.BackColor <> vbBlue And _
      LblPatent.BackColor <> vbBlue And _
      LblTM.BackColor <> vbBlue And _
      LblLAbackup.BackColor <> vbBlue Then
'      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 And _
'         UCase(pub_DbTerminalName) = ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '°õ¦æÀÉÃö³¬¨Ã¥B¬O¥¿¦¡¸ê®Æ®w®É, ¤~±HMail³qª¾
'         Call OpenNeweMail(m_M51Recver, PUB_GetDbTerminal & "¥x¤@¶l¥ó±µ¦¬¨t²Î³QÃö±¼¤F¡A½Ð¾¨³t¦Üm51-win7¬d¬Ý¡I", "¦P¥D¦®")
'         DoEvents
'      End If
      Unload Me
   End If
End Sub

Private Sub Command1_Click()
Dim strMRL01 As String
   
   If LblFCPin.BackColor = vbBlue Or _
      LblFCPout.BackColor = vbBlue Or _
      LblPatent.BackColor = vbBlue Or _
      LblTM.BackColor = vbBlue Or _
      LblLAbackup.BackColor = vbBlue Then
      MsgBox "¦³«H½c¥¿¦b±µ¦¬«H¥ó¡A¤£¥i°õ¦æ¡I", vbExclamation
      Exit Sub
   End If
   strMRL01 = Trim(InputBox("­n¤â°Ê±µ¦¬¨º­Ó«H½c¶Ü¡H¡]¥¼¿é¤J¥Nªí©ñ±ó¡^" & vbCrLf & _
              "«H½c¥N½X: " & Replace(MRL01CName2, " ", vbCrLf), "­«­n°T®§¡I"))
   If strMRL01 = "" Then
      Exit Sub
   Else
      strMRL01 = Right("0" & strMRL01, 2)
      Command1.Tag = "¤â°Ê¶×¤J" 'Add By Sindy 2024/12/20
   End If
   Select Case strMRL01
      Case Left(IPDept¦¬¥ó§X, 2)
         If LblFCPin.BackColor = vbBlue Then 'ÂÅ¦âTimer¥¿¦bRun
            MsgBox "°ê¥~³¡ " & °ê¥~³¡¦¬¥ó«H½c & " «H½c¥¿¦b±µ¦¬«H¥ó¡I", vbExclamation
            Exit Sub
         Else
            TmrFCPin.Interval = 1000
            bolCancel(0) = False: Cancel_idx = 0 'Add By Sindy 2019/2/14
            bolFCPinRun = True
            Call TmrFCPin_Timer
         End If
      Case Left(IPDept±H¥ó§X, 2)
         If LblFCPout.BackColor = vbBlue Then 'ÂÅ¦âTimer¥¿¦bRun
            MsgBox "°ê¥~³¡ " & °ê¥~³¡±H¥ó«H½c & " «H½c¥¿¦b±µ¦¬«H¥ó¡I", vbExclamation
            Exit Sub
         Else
            TmrFCPout.Interval = 1000
            bolCancel(1) = False: Cancel_idx = 1 'Add By Sindy 2019/2/14
            bolFCPoutRun = True
            Call TmrFCPout_Timer
         End If
      Case Left(Patent¦¬¥ó§X, 2)
         If LblPatent.BackColor = vbBlue Then 'ÂÅ¦âTimer¥¿¦bRun
            MsgBox "±M§Q³B " & ±M§Q³B¦¬¥ó«H½c & " «H½c¥¿¦b±µ¦¬«H¥ó¡I", vbExclamation
            Exit Sub
         Else
            TmrPatent.Interval = 1000
            bolCancel(2) = False: Cancel_idx = 2 'Add By Sindy 2019/2/14
            bolPatentRun = True
            Call TmrPatent_Timer
         End If
      Case Left(TM¦¬¥ó§X, 2)
         If LblTM.BackColor = vbBlue Then 'ÂÅ¦âTimer¥¿¦bRun
            MsgBox "°Ó¼Ð³B " & °Ó¼Ð³B¦¬¥ó«H½c & " «H½c¥¿¦b±µ¦¬«H¥ó¡I", vbExclamation
            Exit Sub
         Else
            TmrTM.Interval = 1000
            bolCancel(3) = False: Cancel_idx = 3 'Add By Sindy 2019/2/14
            bolTMRun = True
            Call TmrTM_Timer
         End If
      'Add By Sindy 2024/5/15
      Case Left(LAbackup±H¥ó§X, 2)
         If LblLAbackup.BackColor = vbBlue Then 'ÂÅ¦âTimer¥¿¦bRun
            MsgBox "ªk«ß©Ò " & ªk«ß©Ò±H¥ó«H½c & " «H½c¥¿¦b±µ¦¬«H¥ó¡I", vbExclamation
            Exit Sub
         Else
            TmrLAbackup.Interval = 1000
            bolCancel(4) = False: Cancel_idx = 4
            bolLAbackupRun = True
            Call TmrLAbackup_Timer
         End If
         '2024/5/15 END
   End Select
End Sub

'ÀË¬d«H½c¬O§_¥i¥H°õ¦æ
'True:­n°õ¦æTimer
'­Y¦³¤H¤u•K°Ê®É¦^¶ÇPkey(strMRL01,strMRL02,strMRL03)
Private Function ExecuteSchedule(ByRef strMRL01 As String, ByRef strMRL02 As String, ByRef strMRL03 As String) As Boolean
Dim i As Integer
Dim strStarTime As String, strEndTime As String
Dim strChkStarTime As String, strChkEndTime
Dim bolHandRecv As Boolean
Dim strSubject As String, strErrText As String
Dim cntTime As String
Dim rsA As New ADODB.Recordset
'Dim intA As Integer, intB As Integer  'Added by Lydia 2019/11/08
Dim intTotCnt As Integer 'Add By Sindy 2024/8/8
Dim strRunStarTime As String, strRunEndTime As String 'Add By Sindy 2025/3/13
   
'   '®É¶¡¨S¨ì¤£°õ¦æTimer
'   If strSrvDate(1) <= 20170705 And UCase(pub_DbTerminalName) = ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then ExecuteSchedule = False: Exit Function
'   '«D¤u§@¤Ñ¤£°õ¦æTimer
'   If ChkWorkDay(strSrvDate(1)) = False Then ExecuteSchedule = False: Exit Function
   
   ExecuteSchedule = True '¹w³]­n°õ¦æTimer
   bolHandRecv = False '«D¤â°Ê
   Select Case strMRL01
      '°ê¥~³¡IPDept¦¬«H¶l¥ó / °ê¥~³¡IPDept±H«H¶l¥ó
      Case Left(IPDept¦¬¥ó§X, 2), Left(IPDept±H¥ó§X, 2)
         strRunStarTime = "013000" 'Add By Sindy 2025/3/13
         strRunEndTime = "240000" 'Add By Sindy 2025/3/13
         If strMRL01 = Left(IPDept¦¬¥ó§X, 2) Then
            strChkStarTime = m_RunFCPinStarTime
            strChkEndTime = m_RunFCPinEndTime
            If bolFCPinRun = True Then bolHandRecv = True '¤â°Ê
         Else
            strChkStarTime = m_RunFCPoutStarTime
            strChkEndTime = m_RunFCPoutEndTime
            If bolFCPoutRun = True Then bolHandRecv = True '¤â°Ê
         End If
         '°õ¦æTimerªº®É¬q
         'Modified by Lydia 2019/11/08 ²M±á0~1ÂIÂ_½u(by David)
'Modified by Lydia 2019/11/12 §ï¦¨©T©w
'         For i = 0 To 47
'            If i < 3 Then   '­â±á0~1ÂIÂ_½u, «á¥b¤p®É¤£°õ¦æ
'                strStarTime = "": strEndTime = ""
'            '±ß¤W11ÂI¤À¦¨¨â¦¸¤À«H,³Ì«á°õ¦æ²MªÅ[§R°£ªº¶l¥ó]±ß¤W23:45~23:55
'            ElseIf i = 46 Then '±ß¤W11ÂI²Ä¤@¦¸¤À«H11:00~11:19
'                strStarTime = "230000": strEndTime = "231900"
'            ElseIf i = 47 Then  '±ß¤W11ÂI²Ä¤G¦¸¤À«H11:20~11:39
'                strStarTime = "232000": strEndTime = "233900"
'            Else
'                'Memo by Lydia 2019/11/08 ¥Ø«e°õ¦æ®É¬q,½Ð¨Ï¥ÎComputer\frm000001.Command31_Click(­pºâ¤À«H®É¬q)
'                intA = i \ 2
'                intB = i Mod 2
'                strStarTime = Format(intA, "00") & IIf(intB = 1, "30", "00") & "00"
'                strEndTime = Format(intA, "00") & IIf(intB = 1, "59", "29") & "00"
'            End If
         For i = 1 To 46
            'Modify By Sindy 2022/5/27 ¦Ò¶q³o®É¤£·|¦³±H¥X«H¥ó,¥BFTP¦b¶i¦æ³Æ¥÷,®e©ö³y¦¨¡¨µLªk»PFTP Server«Ø¥ß³s½u¡I¡¨ªº¿ù»~°T®§
            If i = 1 Then strStarTime = "013000": strEndTime = "015900"
            If i = 2 Then strStarTime = "020000": strEndTime = "022900"
            If i = 3 Then strStarTime = "023000": strEndTime = "025900"
            If i = 4 Then strStarTime = "030000": strEndTime = "032900" '2024/10/21 «ì´_¤À«H
            If i = 5 Then strStarTime = "033000": strEndTime = "035900" '2024/10/21 «ì´_¤À«H
            If i = 6 Then strStarTime = "040000": strEndTime = "042900" '2024/10/21 «ì´_¤À«H
            If i = 7 Then strStarTime = "043000": strEndTime = "045900" '2024/10/21 «ì´_¤À«H
            If i = 8 Then strStarTime = "050000": strEndTime = "052900"
            If i = 9 Then strStarTime = "053000": strEndTime = "055900"
            If i = 10 Then strStarTime = "060000": strEndTime = "062900"
            If i = 11 Then strStarTime = "063000": strEndTime = "065900"
            If i = 12 Then strStarTime = "070000": strEndTime = "072900"
            If i = 13 Then strStarTime = "073000": strEndTime = "075900"
            If i = 14 Then strStarTime = "080000": strEndTime = "082900"
            If i = 15 Then strStarTime = "083000": strEndTime = "085900"
            If i = 16 Then strStarTime = "090000": strEndTime = "092900"
            If i = 17 Then strStarTime = "093000": strEndTime = "095900"
            If i = 18 Then strStarTime = "100000": strEndTime = "102900"
            If i = 19 Then strStarTime = "103000": strEndTime = "105900"
            If i = 20 Then strStarTime = "110000": strEndTime = "112900"
            If i = 21 Then strStarTime = "113000": strEndTime = "115900"
            'Modify By Sindy 2022/5/27 ¦Ò¶qFTP¦b¶i¦æ´«³Æ¥÷µwºÐ,®e©ö³y¦¨¡¨µLªk»PFTP Server«Ø¥ß³s½u¡I¡¨ªº¿ù»~°T®§
            If i = 22 Then strStarTime = "120000": strEndTime = "122900" '2024/10/21 «ì´_¤À«H
            If i = 23 Then strStarTime = "123000": strEndTime = "125900"
            If i = 24 Then strStarTime = "130000": strEndTime = "132900"
            If i = 25 Then strStarTime = "133000": strEndTime = "135900"
            If i = 26 Then strStarTime = "140000": strEndTime = "142900"
            If i = 27 Then strStarTime = "143000": strEndTime = "145900"
            If i = 28 Then strStarTime = "150000": strEndTime = "152900"
            If i = 29 Then strStarTime = "153000": strEndTime = "155900"
            If i = 30 Then strStarTime = "160000": strEndTime = "162900"
            If i = 31 Then strStarTime = "163000": strEndTime = "165900"
            If i = 32 Then strStarTime = "170000": strEndTime = "172900"
            If i = 33 Then strStarTime = "173000": strEndTime = "175900"
            If i = 34 Then strStarTime = "180000": strEndTime = "182900"
            If i = 35 Then strStarTime = "183000": strEndTime = "185900"
            If i = 36 Then strStarTime = "190000": strEndTime = "192900"
            If i = 37 Then strStarTime = "193000": strEndTime = "195900"
            If i = 38 Then strStarTime = "200000": strEndTime = "202900"
            If i = 39 Then strStarTime = "203000": strEndTime = "205900"
            If i = 40 Then strStarTime = "210000": strEndTime = "212900"
            If i = 41 Then strStarTime = "213000": strEndTime = "215900"
            If i = 42 Then strStarTime = "220000": strEndTime = "222900"
            If i = 43 Then strStarTime = "223000": strEndTime = "225900"
            '±ß¤W11ÂI¤À¦¨¨â¦¸¤À«H,³Ì«á°õ¦æ²MªÅ[§R°£ªº¶l¥ó]±ß¤W23:45~23:55
            If i = 44 Then strStarTime = "230000": strEndTime = "231900"
            If i = 45 Then strStarTime = "232000": strEndTime = "233900"
            If i = 46 Then strStarTime = "": strEndTime = ""
'--------------------------------
            'ÀË¬d¥Ø«e®É¶¡¸ÓRun Timerªº®É¬q
            If strStarTime <> "" Then
               If Format(Time, "HHMMSS") >= strStarTime And Format(Time, "HHMMSS") <= strEndTime Then
                  'Add By Sindy 2025/5/13
                  If strMRL01 = Left(IPDept¦¬¥ó§X, 2) Then
                     txtPathIPDept.Tag = "Y"
                  ElseIf strMRL01 = Left(IPDept±H¥ó§X, 2) Then
                     txtPathIPDeptOut.Tag = "Y"
                  End If
                  '2025/5/13 END
                  Exit For
               End If
            End If
         Next i
      
      '±M§Q³BPatent¦¬«H¶l¥ó / °Ó¼Ð³BTM¦¬«H¶l¥ó / ªk«ß©Ò±H¥ó«H½c
      Case Left(Patent¦¬¥ó§X, 2), Left(TM¦¬¥ó§X, 2), Left(LAbackup±H¥ó§X, 2)
         strRunStarTime = "070000" 'Add By Sindy 2025/3/13
         strRunEndTime = "191000" 'Add By Sindy 2025/3/13
         If strMRL01 = Left(Patent¦¬¥ó§X, 2) Then
            strChkStarTime = m_RunPatentStarTime
            strChkEndTime = m_RunPatentEndTime
            If bolPatentRun = True Then bolHandRecv = True '¤â°Ê
         ElseIf strMRL01 = Left(TM¦¬¥ó§X, 2) Then
            strChkStarTime = m_RunTMStarTime
            strChkEndTime = m_RunTMEndTime
            If bolTMRun = True Then bolHandRecv = True '¤â°Ê
         Else
            strChkStarTime = m_RunLAbackupStarTime
            strChkEndTime = m_RunLAbackupEndTime
            If bolLAbackupRun = True Then bolHandRecv = True '¤â°Ê
         End If
         '°õ¦æTimerªº®É¬q
         'Modify By Sindy 2024/8/8
         intTotCnt = 24
         If strMRL01 = Left(TM¦¬¥ó§X, 2) Then
            intTotCnt = intTotCnt + 1 '¦h¤@­Ó®É¬q
         End If
         '2024/8/8 END
         For i = 1 To intTotCnt '23 '22
            If i = 1 Then strStarTime = "070000": strEndTime = "072900"
            If i = 2 Then strStarTime = "073000": strEndTime = "075900"
            If i = 3 Then strStarTime = "080000": strEndTime = "082900"
            If i = 4 Then strStarTime = "083000": strEndTime = "085900"
            If i = 5 Then strStarTime = "090000": strEndTime = "092900"
            If i = 6 Then strStarTime = "093000": strEndTime = "095900"
            If i = 7 Then strStarTime = "100000": strEndTime = "102900"
            If i = 8 Then strStarTime = "103000": strEndTime = "105900"
            If i = 9 Then strStarTime = "110000": strEndTime = "112900"
            If i = 10 Then strStarTime = "113000": strEndTime = "115900"
            'Modify By Sindy 2022/5/27 ¦Ò¶qFTP¦b¶i¦æ´«³Æ¥÷µwºÐ,®e©ö³y¦¨¡¨µLªk»PFTP Server«Ø¥ß³s½u¡I¡¨ªº¿ù»~°T®§
            If i = 11 Then strStarTime = "120000": strEndTime = "122900" '2024/10/21 «ì´_¤À«H
            If i = 12 Then strStarTime = "123000": strEndTime = "125900"
            If i = 13 Then strStarTime = "130000": strEndTime = "132900"
            If i = 14 Then strStarTime = "133000": strEndTime = "135900"
            If i = 15 Then strStarTime = "140000": strEndTime = "142900"
            If i = 16 Then strStarTime = "143000": strEndTime = "145900"
            If i = 17 Then strStarTime = "150000": strEndTime = "152900"
            If i = 18 Then strStarTime = "153000": strEndTime = "155900"
            If i = 19 Then strStarTime = "160000": strEndTime = "162900"
            'Modify By Sindy 2024/8/8
            If strMRL01 = Left(TM¦¬¥ó§X, 2) Then
               If i = 20 Then strStarTime = "163000": strEndTime = "165400"
               If i = 21 Then strStarTime = "165500": strEndTime = "165900" '¦h¤@­Ó®É¬q
               If i = 22 Then strStarTime = "170000": strEndTime = "172900"
               If i = 23 Then strStarTime = "173000": strEndTime = "175900"
               If i = 24 Then strStarTime = "180000": strEndTime = "182900"
               If i = 25 Then strStarTime = "": strEndTime = ""
            Else
               If i = 20 Then strStarTime = "163000": strEndTime = "165900"
               If i = 21 Then strStarTime = "170000": strEndTime = "172900"
               If i = 22 Then strStarTime = "173000": strEndTime = "175900"
               If i = 23 Then strStarTime = "180000": strEndTime = "182900"
               If i = 24 Then strStarTime = "": strEndTime = ""
            End If
            '2024/8/8 END
            'ÀË¬d¥Ø«e®É¶¡¸ÓRun Timerªº®É¬q
            If strStarTime <> "" Then
               If Format(Time, "HHMMSS") >= strStarTime And Format(Time, "HHMMSS") <= strEndTime Then
                  'Add By Sindy 2025/5/13
                  If strMRL01 = Left(Patent¦¬¥ó§X, 2) Then
                     txtPathPatent.Tag = "Y"
                  ElseIf strMRL01 = Left(TM¦¬¥ó§X, 2) Then
                     txtPathTM.Tag = "Y"
                  ElseIf strMRL01 = Left(LAbackup±H¥ó§X, 2) Then
                     txtPathLAbackup.Tag = "Y"
                  End If
                  '2025/5/13 END
                  Exit For
               End If
            'Add By Sindy 2025/5/14
            Else
               txtPathPatent.Tag = "N"
               txtPathTM.Tag = "N"
               txtPathLAbackup.Tag = "N"
            '2025/5/14 END
            End If
         Next i
   End Select
    
   'Modify By Sindy 2025/3/13 ¼W¥[ÀË¬d«D¤@¾ã¤éªº¤À«H°_¨´®É¶¡¤º,¤£±Ò°Ê¤À«H
   '                          ¦] 114/03/11 µo¥Í¥b©]®É¶¡Ãa±¼,¤H­û¦¬¨ì´X¦Ê«Ê"¦³ª÷Æ_«H¥ó"ªº³qª¾«H
   Frame6.Caption = strMRL01 & "«H½c"
   LblTime.Caption = strRunStarTime & " ~ " & strRunEndTime
'   strChkStarTime = Format(strChkStarTime, "HHMMSS")
'   strChkEndTime = Format(strChkEndTime, "HHMMSS")
   LblstrChkStarTime.Caption = strChkStarTime
   LblstrChkEndTime.Caption = strChkEndTime
   LblstrStarTime.Caption = strStarTime
   LblstrEndTime.Caption = strEndTime
   DoEvents
   'If strStarTime = "" Then
   If strStarTime = "" Or _
      (strChkStarTime <> "" And strChkEndTime <> "" And _
        Not (Val(strChkStarTime) >= Val(strRunStarTime) And Val(strChkEndTime) <= Val(strRunEndTime)) _
      ) Then
   '2025/3/13 END
      LblMsg.Caption = "(1)ExecuteSchedule=False"
      DoEvents
      ExecuteSchedule = False: GoTo ChkHadSetA '®É¶¡¨S¨ì¤£°õ¦æTimer
   'Add By Sindy 2024/5/27
   Else
      strSql = "delete from mailreceivelog" & _
               " where mrl01='" & strMRL01 & "'" & _
               " and mrl09='A'"
      cnnConnection.Execute strSql, intI
      '2024/5/27 END
   End If
   
   'ÀË¬d¬O§_¤w¦³Run¹L¦¹®É¬qªºTimer
   If strChkStarTime <> "" And strChkEndTime <> "" And _
      (Val(strChkStarTime) >= Val(strStarTime) And Val(strChkEndTime) <= Val(strEndTime)) Then '¤À«H°Ï¬q
      LblMsg.Caption = "(2)ExecuteSchedule=False"
      DoEvents
      ExecuteSchedule = False
   Else
      'ÀË¬d¬O§_¤w¦³±µ¦¬¹L«H¥ó¸ê®Æ
      strSql = "select mrl03,mrl04 from MailReceiveLog" & _
               " where mrl01='" & strMRL01 & "'" & _
               " and mrl02=" & strSrvDate(1) & _
               " and mrl05='" & strUserNum & "'" & _
               " and mrl03 between " & strStarTime & " and " & strEndTime & _
               " and mrl09='E'"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         ExecuteSchedule = False
         LblMsg.Caption = "(3)ExecuteSchedule=False"
         DoEvents
         Select Case strMRL01
            Case Left(IPDept¦¬¥ó§X, 2)
               m_RunFCPinStarTime = "" & rsA.Fields("mrl03")
               m_RunFCPinEndTime = "" & rsA.Fields("mrl04")
            Case Left(IPDept±H¥ó§X, 2)
               m_RunFCPoutStarTime = "" & rsA.Fields("mrl03")
               m_RunFCPoutEndTime = "" & rsA.Fields("mrl04")
            Case Left(Patent¦¬¥ó§X, 2)
               m_RunPatentStarTime = "" & rsA.Fields("mrl03")
               m_RunPatentEndTime = "" & rsA.Fields("mrl04")
            Case Left(TM¦¬¥ó§X, 2)
               m_RunTMStarTime = "" & rsA.Fields("mrl03")
               m_RunTMEndTime = "" & rsA.Fields("mrl04")
            Case Left(LAbackup±H¥ó§X, 2)
               m_RunLAbackupStarTime = "" & rsA.Fields("mrl03")
               m_RunLAbackupEndTime = "" & rsA.Fields("mrl04")
         End Select
      Else
         'ÀË¬d¬O§_¦³¥¿¦b°õ¦æ¤¤ªºTimer
         strSql = "select mrl01,mrl02,mrl03,mrl04,mrl05 from MailReceiveLog" & _
                  " where mrl01='" & strMRL01 & "'" & _
                  " and mrl02=" & strSrvDate(1) & _
                  " and mrl05='" & strUserNum & "'" & _
                  " and mrl09='Y'"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            ExecuteSchedule = False
            LblMsg.Caption = "(4)ExecuteSchedule=False"
            DoEvents
            '¦pªG¤w15¤ÀÄÁ©|¥¼µ²§ô,«h³qª¾¹q¸£¤¤¤ß¤H­û
            strExc(1) = Format(rsA.Fields("mrl03"), "0#####")
            If Mid(strExc(1), 3, 2) + 15 = 59 Then
               cntTime = CStr(Left(strExc(1), 2) + 1) & "00" & CStr(Right(strExc(1), 2))
            ElseIf Mid(strExc(1), 3, 2) + 15 > 59 Then
               cntTime = CStr(Left(strExc(1), 2) + 1) & Format(CStr(Mid(strExc(1), 3, 2) + 15 - 60), "0#") & CStr(Right(strExc(1), 2))
            Else
               cntTime = CStr(Left(strExc(1), 2)) & Format(CStr(Mid(strExc(1), 3, 2) + 15), "0#") & CStr(Right(strExc(1), 2))
            End If
            If Val(cntTime) <= Val(Format(Time, "HHMMSS")) Then
               strSubject = PUB_GetDbTerminal & "¦³±µ¦¬«H½c(" & strMRL01 & ")¥¿¦b°õ¦æ¤¤,¤w15¤ÀÄÁ©|¥¼µ²§ô,¬O§_¦³²§±`¡A½Ð¬d¬Ý¡I"
               strErrText = "mrl03=" & rsA.Fields("mrl03") & vbCrLf & _
                            "mrl04=" & rsA.Fields("mrl04") & vbCrLf & _
                            "mrl05=" & rsA.Fields("mrl05") & " " & GetPrjSalesNM(rsA.Fields("mrl05")) & vbCrLf & _
                            "strStarTime=" & strStarTime & vbCrLf & _
                            "strEndTime=" & strEndTime
               If bolHandRecv = True Then '¤â°Ê
                  MsgBox strSubject & vbCrLf & strErrText, vbExclamation
                  bolHandRecv = False
               Else
                  strSql = "UPDATE MailReceiveLog SET MRL04=" & Format(Time, "HHMMSS") & ",MRL09='F'" & _
                           " where mrl01='" & strMRL01 & "'" & _
                           " and mrl02=" & strSrvDate(1) & _
                           " and mrl05='" & strUserNum & "'" & _
                           " and mrl09='Y'"
                  cnnConnection.Execute strSql
                  PUB_SendMail strUserNum, m_M51Recver, "", strSubject, strErrText, , , , , , , , , , , False, , , False, , , False
'                  DoEvents
                  ExecuteSchedule = True
                  LblMsg.Caption = "(A)ExecuteSchedule=True"
                  DoEvents
               End If
            End If
         Else
            'ÀË¬d¬O§_¦³¤â°Ê±µ¦¬«H½c¥¿¦b°õ¦æ¤¤
            strSql = "select mrl03,mrl04,mrl05 from MailReceiveLog" & _
                     " where mrl01='" & strMRL01 & "'" & _
                     " and mrl02=" & strSrvDate(1) & _
                     " and mrl05<>'" & strUserNum & "'" & _
                     " and mrl09='Y'" & _
                     " order by mrl03 desc"
            intI = 1
            Set rsA = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               ExecuteSchedule = False
               LblMsg.Caption = "(5)ExecuteSchedule=False"
               DoEvents
               '¦pªG¤w30¤ÀÄÁ©|¥¼µ²§ô,«h³qª¾¹q¸£¤¤¤ß¤H­û
               strExc(1) = Format(rsA.Fields("mrl03"), "0#####")
               If Mid(strExc(1), 3, 2) + 30 = 59 Then
                  cntTime = CStr(Left(strExc(1), 2) + 1) & "00" & CStr(Right(strExc(1), 2))
               ElseIf Mid(strExc(1), 3, 2) + 30 > 59 Then
                  cntTime = CStr(Left(strExc(1), 2) + 1) & Format(CStr(Mid(strExc(1), 3, 2) + 30 - 60), "0#") & CStr(Right(strExc(1), 2))
               Else
                  cntTime = CStr(Left(strExc(1), 2)) & Format(CStr(Mid(strExc(1), 3, 2) + 30), "0#") & CStr(Right(strExc(1), 2))
               End If
               If Val(cntTime) <= Val(Format(Time, "HHMMSS")) Then
                  strSubject = PUB_GetDbTerminal & "¦³¤â°Ê±µ¦¬«H½c(" & strMRL01 & ")¥¿¦b°õ¦æ¤¤,¤w¤@¤p®É©|¥¼µ²§ô,¬O§_¦³²§±`¡A½Ð¬d¬Ý¡I"
                  strErrText = "mrl03=" & rsA.Fields("mrl03") & vbCrLf & _
                               "mrl04=" & rsA.Fields("mrl04") & vbCrLf & _
                               "mrl05=" & rsA.Fields("mrl05") & " " & GetPrjSalesNM(rsA.Fields("mrl05")) & vbCrLf & _
                               "strStarTime=" & strStarTime & vbCrLf & _
                               "strEndTime=" & strEndTime
                  If bolHandRecv = True Then '¤â°Ê
                     MsgBox strSubject & vbCrLf & strErrText, vbExclamation
                     bolHandRecv = False
                  Else
                     PUB_SendMail strUserNum, m_M51Recver, "", strSubject, strErrText, , , , , , , , , , , False, , , False, , , False
'                     DoEvents
                  End If
               End If
            End If
         End If
      End If
   End If
   
   Set rsA = Nothing
   Exit Function
   
   'Add By Sindy 2017/11/15
ChkHadSetA:
   'ÀË¬d¬O§_¦³¤H¤u±Ò°Ê
   strSql = "select mrl02,mrl03,mrl05 from MailReceiveLog" & _
            " where mrl01='" & strMRL01 & "'" & _
            " and mrl09='A'"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ExecuteSchedule = True
      LblMsg.Caption = "(B)ExecuteSchedule=True"
      DoEvents
      strMRL02 = "" & rsA.Fields("mrl02")
      strMRL03 = Format("" & rsA.Fields("mrl03"), "00:00:00")
   End If
   '2017/11/15 END
   
   Set rsA = Nothing
End Function

'Modify By Sindy 2023/7/17
'Modify By Sindy 2024/1/31 strMailBox(«H½c)¤¤­n¸ÑªRInboxCount(²Ä´X­ÓFolder)
'inbound@taie.com.tw
'backup@taie.com.tw
'  Inbox
'  Junk Email
'patent@taie.com.tw
'tm@taie.com.tw
'  ¦¬¥ó§X
'  ©U§£¶l¥ó
Private Function OpenOutLookFolder(ByRef myNamespace As Object, ByRef myFolder As Object, _
   ByVal strMailBox As String, ByVal InboxCount As Integer) As Boolean
Dim strMailName As String
'Add By Sindy 2024/1/31
Dim strFolderName As String
Dim strTestMailName As String
Dim strTestFolderName As String
'2024/1/31 END
   
   If strMailBox = "01" Then
      strMailName = °ê¥~³¡¦¬¥ó«H½c 'inbound@taie.com.tw
      'Modify By Sindy 2024/1/31
      If InboxCount = 1 Then
         strFolderName = "Inbox"
      Else
         strFolderName = "Junk Email"
      End If
      '2024/1/31 END
   ElseIf strMailBox = "02" Then
      strMailName = °ê¥~³¡±H¥ó«H½c 'backup@taie.com.tw
      'Modify By Sindy 2024/1/31
      If InboxCount = 1 Then
         strFolderName = "Inbox"
      Else
         strFolderName = "Junk Email"
      End If
      '2024/1/31 END
   ElseIf strMailBox = "03" Then
      strMailName = ±M§Q³B¦¬¥ó«H½c 'patent@taie.com.tw
      'Modify By Sindy 2024/1/31
      If InboxCount = 1 Then
         strFolderName = "¦¬¥ó§X"
      Else
         strFolderName = "©U§£¶l¥ó"
      End If
      '2024/1/31 END
   ElseIf strMailBox = "04" Then
      strMailName = °Ó¼Ð³B¦¬¥ó«H½c 'tm@taie.com.tw
      'Modify By Sindy 2024/1/31
      If InboxCount = 1 Then
         strFolderName = "¦¬¥ó§X"
      Else
         strFolderName = "©U§£¶l¥ó"
      End If
      '2024/1/31 END
   'Add By Sindy 2024/5/14
   ElseIf strMailBox = "05" Then
      strMailName = ªk«ß©Ò±H¥ó«H½c 'LAbackup@taie.com.tw
      'Modify By Sindy 2024/1/31
      If InboxCount = 1 Then
         strFolderName = "Inbox"
      Else
         strFolderName = "Junk Email"
      End If
      '2024/5/14 END
   Else
      OpenOutLookFolder = False
      Exit Function
   End If
   
   'Add By Sindy 2024/1/31
   'If UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '´ú¸Õ¸ê®Æ®w
      'ex:PUB_ReadHostName=A97038
      If InStr(PUB_ReadHostName, "-") > 0 Then
         strExc(0) = Left(PUB_ReadHostName, Len(PUB_ReadHostName) - 1)
      End If
      strExc(0) = Right(PUB_ReadHostName, 5)
      strTestMailName = strExc(0) & "@taie.com.tw"
      If InboxCount = 1 Then
         strTestFolderName = "´ú¸Õ¤À«H"
      Else
         strTestFolderName = "©U§£¶l¥ó"
      End If
   'End If
   '2024/1/31 END
   
   'Modify By Sindy 2023/12/29
   If UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '´ú¸Õ¸ê®Æ®w
      Set myFolder = myNamespace.Folders(strTestMailName).Folders(strTestFolderName)
      strExc(0) = "¡]" & strMailName & "¡^" & vbCrLf & vbCrLf & _
                  "Folders(" & strTestMailName & ").Folders(" & strTestFolderName & ")"
   
   '¥¿¦¡¸ê®Æ®w
   Else
'      Set myFolder = myNamespace.Folders(strMailName).Folders(strFolderName)
'      strExc(0) = "¡]" & strMailName & "¡^" & vbCrLf & vbCrLf & _
'                  "Folders(" & strMailName & ").Folders(" & strFolderName & ")"
      If UCase(PUB_ReadHostName) = UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) Then
         'Modify By Sindy 2024/2/20 ¤U¤È¤S§ï¦^§Úªº³Ì·R,¦]¤½¥Î¸ê®Æ§¨·PÄ±¬O½u¤W·|¦³±Æª©°ÝÃD(¤º®eÀ½¦b¤@°_)
'         'Add By Sindy 2024/2/20
'         If strMailBox = "02" Then '°ê¥~³¡±H¥ó«H½c(backup@taie.com.tw)
            Set myFolder = myNamespace.Folders("¤½¥Î¸ê®Æ§¨ - " & Pub_GetSpecMan("¤À«H¥D¾÷¦¬¥ó§X¦WºÙ")).Folders("§Úªº³Ì·R").Folders(strMailName)
            strExc(0) = "¡]" & strMailName & "¡^" & vbCrLf & vbCrLf & _
                        "«H½c: Folders(¤½¥Î¸ê®Æ§¨ - " & Pub_GetSpecMan("¤À«H¥D¾÷¦¬¥ó§X¦WºÙ") & ").Folders(§Úªº³Ì·R).Folders(" & strMailName & ")"
'         Else
'         '2024/2/20 END
'            'Modify By Sindy 2024/2/20 Backup¤½¥Î¸ê®Æ§¨(½u¤W) ¸ê®Æ¦³¤j¶q´Ý¯d,§ï¬°¤£­n¨Ï¥Î§Úªº³Ì·R; ·í®É¸ê®Æ¦³²£¥Í­«ÂÐÂk¨÷
'            Set myFolder = myNamespace.Folders("¤½¥Î¸ê®Æ§¨ - " & Pub_GetSpecMan("¤À«H¥D¾÷¦¬¥ó§X¦WºÙ")).Folders("©Ò¦³¤½¥Î¸ê®Æ§¨").Folders(strMailName)
'            strExc(0) = "¡]" & strMailName & "¡^" & vbCrLf & vbCrLf & _
'                        "«H½c: Folders(¤½¥Î¸ê®Æ§¨ - " & Pub_GetSpecMan("¤À«H¥D¾÷¦¬¥ó§X¦WºÙ") & ").Folders(©Ò¦³¤½¥Î¸ê®Æ§¨).Folders(" & strMailName & ")"
'   '         'Modify By Sindy 2024/2/15 ©Ò¦³¤½¥Î¸ê®Æ§¨ §ï¥Î §Úªº³Ì·R(¥i³]Â÷½u)
'   '         Set myFolder = myNamespace.Folders("¤½¥Î¸ê®Æ§¨ - " & Pub_GetSpecMan("¤À«H¥D¾÷¦¬¥ó§X¦WºÙ")).Folders("§Úªº³Ì·R").Folders(strMailName)
'   '         strExc(0) = "¡]" & strMailName & "¡^" & vbCrLf & vbCrLf & _
'   '                     "«H½c: Folders(¤½¥Î¸ê®Æ§¨ - " & Pub_GetSpecMan("¤À«H¥D¾÷¦¬¥ó§X¦WºÙ") & ").Folders(§Úªº³Ì·R).Folders(" & strMailName & ")"
'         End If
      Else
         Set myFolder = myNamespace.Folders("¤½¥Î¸ê®Æ§¨ - " & strTestMailName).Folders("©Ò¦³¤½¥Î¸ê®Æ§¨").Folders(strMailName)
         strExc(0) = "¡]" & strMailName & "¡^" & vbCrLf & vbCrLf & _
                     "«H½c: Folders(¤½¥Î¸ê®Æ§¨ - " & strTestMailName & ").Folders(©Ò¦³¤½¥Î¸ê®Æ§¨).Folders(" & strMailName & ")"
      End If
   End If
   '2023/12/29 END
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or UCase(pub_DbTerminalName) <> UCase(¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ) Then 'Run VB
      If MsgBox("¬O§_½T©w­n¶×¤J¡H" & vbCrLf & vbCrLf & "¤À«H³W«h¬°: " & strExc(0) & " ¶l¥ó¶Ü¡H", vbExclamation + vbYesNo + vbDefaultButton2, "­«­n°T®§¡I") = vbNo Then
         OpenOutLookFolder = False
         Exit Function
      End If
   End If
   OpenOutLookFolder = True
End Function

Private Sub TmrFCPin_Timer()
   'Modify By Sindy 2024/5/17
   'Call importFCPinBound
   Call ChkExecutionTimer(Left(IPDept¦¬¥ó§X, 2))
   '2024/5/17 END
End Sub

''°ê¥~³¡¦¬¥ó«H½c³B²zµ{§Ç
'Private Function importFCPinBound() As Boolean
'Dim kk As Integer, jj As Integer
'Dim strTo As String, strCC As String, strTempCC As String
'Dim oFileSys As New FileSystemObject, oFolder As Object
'Dim strKind As String
''Dim myForward As outlook.MailItem
'Dim myForward As Object
''Dim myNewEmail As outlook.MailItem 'Âà±H«H¥ó
'Dim myNewEmail As Object 'Âà±H«H¥ó
'Dim ArrStr As Variant, ArrStrkk As Variant
'Dim strCaseNo As String
'Dim strIPMNoteSMIME As String '¥[±K¥D¦®
'Dim bolReStarFCPin As Boolean
'Dim strMRL01 As String, strMRL02 As String, strMRL03 As String, strMRL04 As String, strMRL05 As String
'Dim rsA As New ADODB.Recordset
'Dim strErrNumber As String 'Add By Sindy 2019/10/14
'Dim intURGENT As Integer 'Add By Sindy 2019/11/14
'Dim bolRunIPDeptISDMail As Boolean 'Add By Sindy 2020/3/9
'Dim strErrCode As String, strErrDesc As String 'Add By Sindy 2020/4/15
'Dim fs 'Add By Sindy 2022/2/22
'Dim strRecipients_1 As String, strRecipients_all As String '§ì¦¬¥óªÌ¸ê®Æ
'Dim strF1xEmp As String, strF2xEmp As String 'Add By Sindy 2023/5/23
'Dim varTmp As Variant 'Add By Sindy 2023/5/23
''Add By Sindy 2023/6/26
'Dim olApp As Object
'Dim myNamespace As Object
'Dim myFolder As Object
'Dim myItems As Object
''2023/6/26 END
'Dim oFile As Object
'Dim intFolder As Integer '­nÅª¨úªºFolder¼Æ; ex:Inbox ©M Junk Email
'
'On Error GoTo ErrNo1
'
'   If cnnConnection.State = adStateClosed Then Exit Function '±ß¤WDBÂ_½u,¤£»Ý©¹¤U°õ¦æ
'   '¥H§KTimer¦P®ÉRun°_¨Ó
'   If LblFCPout.BackColor = vbBlue Then Exit Function
'   If LblPatent.BackColor = vbBlue Then Exit Function
'   If LblFCPin.BackColor = vbBlue Then Exit Function
'   If LblTM.BackColor = vbBlue Then Exit Function
'
'   m_strMailTo = "" 'Add By Sindy 2022/5/25
'   strErrText = "" 'Add By Sindy 2020/7/22
''   If MsgBox("¬O§_­n¶×¤J" & °ê¥~³¡¦¬¥ó«H½c & "«H¥ó¡H", vbExclamation + vbYesNo + vbDefaultButton2, "­«­n°T®§¡I") = vbNo Then
''      TmrFCPin.Interval = 0
''      Exit Sub
''   End If
'
'   importFCPinBound = False
'   If txtPathIPDept = "" Then
'      MsgBox "¦¬¥ó¸ê®Æ§¨¤£¥iªÅ¥Õ¡I"
'      txtPathIPDept.SetFocus
'      Exit Function
'   End If
'   If Dir(txtPathIPDept, vbDirectory) = "" Then
'      MkDir txtPathIPDept
'   End If
'
'   strMRL01 = Left(IPDept¦¬¥ó§X, 2): strMRL02 = "": strMRL03 = ""
'strErrText = "InB-A:" 'Add By Sindy 2023/2/22 D-Bug
'   If ExecuteSchedule(strMRL01, strMRL02, strMRL03) = True Or bolFCPinRun = True Then '­n°õ¦æTimer
''      'Add By Sindy 2023/11/29
''      Set eventConn = cnnConnection
''      KillCmdLog
''      '2023/11/29 END
'
'      bolFCPinRun = False
'
'strErrText = "InB-B:" 'Add By Sindy 2023/2/22 D-Bug
'      Set olApp = CreateObject("Outlook.Application")
'strErrText = "InB-C:" 'Add By Sindy 2023/2/22 D-Bug
'      Set myNamespace = olApp.GetNamespace("MAPI")
'
'strErrText = "InB-E:" 'Add By Sindy 2023/2/22 D-Bug
'      intKeyCnt = 0: intRunOK = 0: intCaseOK = 0
'
'strErrText = "InB-C:-2" 'Add By Sindy 2023/2/22 D-Bug
'   'Add By Sindy 2024/1/31
'   For intFolder = 1 To 1 '2
'      'Modify By Sindy 2023/7/18
'      If OpenOutLookFolder(myNamespace, myFolder, Left(IPDept¦¬¥ó§X, 2), intFolder) = False Then
'         importFCPinBound = True
'         Set olApp = Nothing
'         Set myNamespace = Nothing
'         Set myFolder = Nothing
'         TmrFCPin.Interval = 0
'         LblFCPin.BackColor = vbRed
'         Exit Function
'      End If
'      '2023/7/18 END
'
'      bolReStarFCPin = False
'
'ReStarFCPin:
''      Screen.MousePointer = vbHourglass
'      Set myItems = myFolder.Items
'      strIPMNoteSMIME = "" '¥[±K¥D¦®
'      intMaxItem = myItems.Count
'
'strErrText = "InB-F:" & "intMaxItem=" & intMaxItem 'Add By Sindy 2023/2/22 D-Bug
'      '°O¿ýLogÀÉ
'      'Modify By Sindy 2024/1/31 + And intFolder = 1
'      If strMRL02 = "" And intFolder = 1 Then
'         'strMRL01 = Left(IPDept¦¬¥ó§X, 2)
'         strMRL02 = strSrvDate(1)
'         strMRL03 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         strMRL05 = strUserNum
'         strSql = "insert into MailReceiveLog(MRL01,MRL02,MRL03,MRL05,MRL09)" & _
'                  "values('" & strMRL01 & "'," & strMRL02 & "," & Format(strMRL03, "hhmmss") & ",'" & strMRL05 & "','Y')"
'         cnnConnection.Execute strSql
'      End If
'
'strErrText = "InB-G:" & "intMaxItem=" & intMaxItem 'Add By Sindy 2023/2/22 D-Bug
'      If intMaxItem > 0 Then
'         If bolUserControl = True Then
'            frmpic002.Label1.Caption = "¶l¥ó±µ¦¬¤¤...½Ðµy­Ô..."
'            frmpic002.Show
'            frmpic002.ZOrder 0
'            frmpic002.Label1.Font.Size = 12
'            frmpic002.Label1.Font.Bold = True
'         End If
'         For mail_ii = myItems.Count To 1 Step -1
'strErrText = "InB-H:" & "mail_ii=" & mail_ii & " : intMaxItem=" & intMaxItem   'Add By Sindy 2023/2/22 D-Bug
'            LblFCPin.BackColor = vbBlue 'ÂÅ¦âTimer¥¿¦bRun
'            cmdCancel(0).Enabled = True
'            DoEvents
'            If bolUserControl = True Then
'               frmpic002.Label1.Caption = "¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii & "...½Ðµy­Ô~"
'            Else
'               Frame1.Caption = Frame1.Tag & "¡@¡@¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii
'            End If
'strErrText = "InB-I:" & "Frame1.Caption=" & Frame1.Caption 'Add By Sindy 2023/2/22 D-Bug
'            DoEvents
'            strErrText = ""
'            intRunOK = intRunOK + 1 '°O¿ý±µ¦¬µ§¼Æ (2017/7/20¤~¶}©l°O¿ý¥þ³¡±µ¦¬ªºµ§¼Æ)
'            strRecipients_1 = "": strRecipients_all = "" '§ì¦¬¥óªÌ¸ê®Æ
'            Call ReadMailText(myItems, True, strRecipients_all, strRecipients_1)
'
'            'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'            strErrText = "²Ä " & mail_ii & " µ§ ¥D¦®: " & strSocSubject & vbCrLf
'            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strSender: " & strSender & vbCrLf
'            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strMailDateTime: " & strMailDate & " " & strMailTime
'            Call WLog_Day(strErrText, °ê¥~³¡¦¬¥ó«H½c)
'
'            '·í±H¥ó¤H¦³­n¨DÅª¨ú¦^±ø®É¨t²Î·|µo«H
'            '1.­nOutlook³]©w¤£¦^ÂÐÅª¨ú¦^±ø(¦ý«eÃD¬O«H¥ó¤]¥²¶·³]¬°¤w¶}±Ò)
'            '2.­n³]©w¦Û°Ê²M°£¡¨§R°£ªº¶l¥ó¡¨
'            '3.­n³]©w¥i¥H¸Ñ¶}ª÷Æ_«H¥ó:°òÂ¦ªº¦w¥þ©Ê¨t²Î§ä¤£¨ì±zªº¼Æ¦ì ID ¦WºÙ(-2146893792)
'            'IPM.Note.SMIME ¥[±K
'            'Modify By Sindy 2017/11/17
'            'Modify By Sindy 2023/7/12 + Or myItems.Item(mail_ii).Class = 45 : ·s³qª¾ => UCase(myItems.Item(mail_ii).MessageClass) = UCase("IPM.Post")
'            If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or myItems.Item(mail_ii).Class = 45 Then
'            'If myItems.Item(mail_ii).Class <> 43 Then
'            '2017/11/17 END
'               intKeyCnt = intKeyCnt + 1
'               'Add By Sindy 2017/7/18 ¥[Log°O¿ý
'               'strErrText = "²Ä " & mail_ii & " µ§ [¥[±K] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[¥[±K¶l¥ó]" & vbCrLf, °ê¥~³¡¦¬¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf '¥[±K¥D¦®
'               '2017/7/18 END
'            'Add By Sindy 2020/4/10 ¦^¦¬¶l¥ó,ª½±µ§R°£
'            ElseIf InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
'               intKeyCnt = intKeyCnt + 1
'               'strErrText = "²Ä " & mail_ii & " µ§ [¦^¦¬] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[¦^¦¬¶l¥ó]" & vbCrLf, °ê¥~³¡¦¬¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'               'myItems.Item(mail_ii).Delete '§R°£ =>µLªk§R°£,·|·í
'               'DoEvents
'            Else
'
'               strFileName = mail_ii & "." & _
'                             strSrvDate(1) & Right("000000" & ServerTime, 6) & ".msg"
'               myItems.Item(mail_ii).SaveAs txtPathIPDept & "\" & strFileName, 9 '9.Outlook Unicode¶l¥ó®æ¦¡.msg
'               'Add By Sindy 2020/2/27 SaveAs¨ç¼Æ,´N·|±Ò°Ê°»´ú¯f¬r³nÅéªº¨¾¬r¾÷¨î¤F
'               Sleep 1000
'               DoEvents
'               Call WLog_Day("²£¥Í¼È¦s¹q¤lÀÉ: " & txtPathIPDept & "\" & strFileName, °ê¥~³¡¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'               '2020/2/27 END
'
'               'Add By Sindy 2022/2/22
'               '«H¥ó¦P®É¦³±Hipdept¤Îpatent«H½c®É,¤~ÀË¬d:
'               If InStr(UCase(strRecipients_all), UCase("patent@taie.")) > 0 And _
'                  InStr(UCase(Replace(strRecipients_all, "80ipdept@taie.com.tw", "")), UCase("ipdept@taie.")) > 0 Then
'                  '¥ý¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶i¨Ó¤F¡F­Y¦³¡A§R°£¡C­Y¨S¦³¡AÄ~Äò¡C
'                  strSql = "select ii01,ii03 from ipdeptinput" & _
'                           " where ii17 = '" & ChgSQL(strSocSubject) & "'" & _
'                           " and ii11 = '" & ChgSQL(strSender) & "' and ii12 = " & DBDATE(strMailDate) & " and ii13 = " & Val(Replace(strMailTime, ":", "")) & _
'                           " order by ii01 desc,ii03 desc"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     '«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]
'                     intKeyCnt = intKeyCnt + 1
'                     Call WLog_Day("[«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]", °ê¥~³¡¦¬¥ó«H½c)
'                     strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'                     Call DeleteMyItems(myItems, °ê¥~³¡¦¬¥ó«H½c) '§R°£Outlook¸Ì­±ªº¶l¥ó
'                     '§R°£PCºÝÀÉ®×
'                     Set fs = CreateObject("Scripting.FileSystemObject")
'                     Call fs.DeleteFile(txtPathIPDept & "\" & strFileName)
'                     Sleep 1000
'                     DoEvents
'                     GoTo IsReadNext 'Run¤U¤@µ§
'                  Else
'                     'ÀË¬d±M§Q³B¬O§_¦³¦¹µ§¸ê®Æ
'                     strSql = "select pi01,pi03 from patentinput" & _
'                              " where pi17 = '" & ChgSQL(strSocSubject) & "'" & _
'                              " and pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate) & " and pi13 = " & Val(Replace(strMailTime, ":", "")) & _
'                              " order by pi01 desc,pi03 desc"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 1 Then
'                        '³oª¬ªp¬O¤£À³¸Óµo¥Íªº
'                        PUB_SendMail strUserNum, "97038", "", _
'                           "¡iIPDept-¦¹µ§¶l¥ó±M§Q³B¤w¦¬¿ý(" & RsTemp.Fields("pi01") & "-" & RsTemp.Fields("pi03") & "),°ê¥~³¡¥¼¤@¨Ö¦¬¿ý,½ÐÀË¬dª¬ªp¡H(Ä~Äò©¹¤URun,¶i¦æ¶l¥ó¦¬¿ý...)¡j", strSocSubject & vbCrLf & vbCrLf & strSql, , txtPathIPDept & "\" & strFileName, , , , , , , , True, False, , , False, , , False
'                        'Ä~Äò©¹¤URun,¶i¦æ¶l¥ó¦¬¿ý...
'                     End If
'                  End If
'               End If
'               '2022/2/22 END
'
'               If intErr2147024882 <> mail_ii Then
'                  'Add By Sindy 2018/4/12
'                  If Dir(txtPathIPDept & "\" & strFileName) = "" Then
'                     strErrText = "µL²£¥Í¹q¤lÀÉ,ºÃ¦ü¤¤¯f¬r " & "Err.Number:" & Err.Number & Err.Description & vbCrLf
'                     Call ExportEMailErr(myItems, False, °ê¥~³¡¦¬¥ó«H½c, strErrText, Err.Number, Err.Description, _
'                           strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'                  'Add By Sindy 2020/4/14 ÀË¬d¹q¤lÀÉ¬O§_¥i¥H¥¿±`¶}±Ò
'                  ElseIf ChkIsOpenEmail(txtPathIPDept & "\" & strFileName, strErrCode, strErrDesc) = False Then
'                     intKeyCnt = intKeyCnt + 1
'                     strErrText = "²Ä " & mail_ii & " µ§ [MsgµLªk¶}±Ò] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf & _
'                        txtPathIPDept & "\" & strFileName & vbCrLf & _
'                        "Err.Number:" & strErrCode & strErrDesc & vbCrLf
'                     Call WLog_Day(strErrText, °ê¥~³¡¦¬¥ó«H½c)
'                     strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'                  Else
'                  '2018/4/12 END
'                     'Add By Sindy 2018/7/10 °ê»Ú·|Ä³¶l¥ó -- (ª`·N:¥~¨Ó¶l¥ó¤@¼Ë­n¤À«H¥X¥h)
'                     bolRunIPDeptISDMail = False
'                     pub_SaveCoRec = False 'Add By Sindy 2022/6/17 °O¿ý¬O§_¦³Àx¦s©¹¨Ó°O¿ý
'                     If PUB_IPDeptISDMail(Me, "0", m_strISDPath, txtPathIPDept, strFileName, intCaseOK) = True Then
'                        Call WLog_Day("PUB_IPDeptISDMail => OK", °ê¥~³¡¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                        bolRunIPDeptISDMail = True
''                        myItems.Item(mail_ii).Delete '§R°£
''                        DoEvents
'                     End If
'                     '2018/7/10 END
'                     Sleep 100 'Add By Sindy 2019/12/13
'                     '¦s­ÓÀÉ®É¥D¦®¤£¥i¥H¦³\/:*?"<>|µ¥²Å¸¹
'                     If PUB_IPDeptTransMail_New(Me, strTo, strErrText, strKind, strFileName, strCaseNo) = True Then
'                        Call WLog_Day("PUB_IPDeptTransMail_New = True; (¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii & "); myItems.Count = " & myItems.Count, °ê¥~³¡¦¬¥ó«H½c)
'                        Call DeleteMyItems(myItems, °ê¥~³¡¦¬¥ó«H½c) '§R°£Outlook¸Ì­±ªº¶l¥ó
'
'                        'If strKind = "1" Then '­Ó®×
'                        If strCaseNo <> "" Then '¦³Âk¨÷©v°Ï´Nºâ­Ó®×¥ó¼Æ Modify By Sindy 2017/7/21
'                           intCaseOK = intCaseOK + 1
'                        End If
'
'                     Else
'                        'Add By Sindy 2020/3/9 ©¹¨Ó°O¿ý«H¥ó±H¥X, ¶Ç¦^=>¥¼¶Ç»¼ªº¥D¦®: Best wishes and update from Tai E regarding COVID-19 [Our Ref:Y53102000.B49] (EY/wc)
'                        '  ©¹¨Ó°O¿ýªº¡¨¥¼¶Ç»¼ªº¥D¦®¡¨«H¥ó=>¬Oª½±µ§R°£¶l¥ó¹q¤lÀÉ,©Ò¥H¦b¦¹­n­ç°£,¤£µM·|³Q§PÂ_¬°¯f¬rÀÉ
'                        If bolRunIPDeptISDMail = True And InStr(myItems.Item(mail_ii).Subject, "¥¼¶Ç»¼ªº¥D¦®") > 0 Then
'                           Call DeleteMyItems(myItems, °ê¥~³¡¦¬¥ó«H½c, "©¹¨Ó°O¿ýªº<¥¼¶Ç»¼ªº¥D¦®>«H¥ó => ª½±µ§R°£") '§R°£Outlook¸Ì­±ªº¶l¥ó
'
'                        Else
'                        '2020/3/9 END
'                           strErrNumber = Err.Number 'Add By Sindy 2019/10/14
'                           Call WLog_Day("¤À«H¥¢±Ñ(1): " & strErrText & ";" & Err.Number & ":" & Err.Description, °ê¥~³¡¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                           'Add By Sindy 2020/9/10
'                           If strErrText <> "" And strErrText <> "Err.Number:0;" Then
'                           Else
'                           '2020/9/10 END
'                              'Add By Sindy 2019/12/11
'                              If strErrNumber = "0" Then
'                                 strErrText = "§ä¤£¨ìÀÉ®×,ºÃ¦ü¤¤¯f¬r"
'      '                           myItems.Item(mail_ii).Delete '§R°£
'      '                           DoEvents
'                              End If
'                              '2019/12/11 END
'                           End If
'
'                           Call WLog_Day("¤À«H¥¢±Ñ(2): " & strErrText & ";" & Err.Number & ":" & Err.Description, °ê¥~³¡¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                           Call ExportEMailErr(myItems, False, °ê¥~³¡¦¬¥ó«H½c, strErrText, Err.Number, Err.Description, _
'                              strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'                           'Add By Sindy 2019/10/14
'                           'If strErrNumber = "999" Then
'                           If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then
'                              Call WLog_Day("¤À«H¥¢±Ñ(3): 999 " & strErrText & vbCrLf, °ê¥~³¡¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                              Exit For
'                           End If
'                           '2019/10/14 END
'                        End If
'                     End If
'                  End If
'               'Modify By Sindy 2020/4/15
'               Else
'                  intErr2147024882 = 0
'               '2020/4/15 END
'               End If
'            End If
'IsReadNext:
'            '¬O§_­n¤¤Â_
'            If bolCancel(0) = True Then
'               LblFCPin.BackColor = vbRed
'               DoEvents 'Add By Sindy 2024/5/7
'               GoTo IsCancel
'            End If
'         Next mail_ii
'
'IsCancel:
'         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         If bolUserControl = True Then
'            Unload frmpic002
'            Set frmpic002 = Nothing
'         End If
''         '¦³¥[±K«H¥ó¥B¬°¤u§@¤Ñ¤~­n±H«H³qª¾¤H­û³B²z
''         If intKeyCnt > 0 And ChkWorkDay(strSrvDate(1)) = True Then
''            '±HE-Mail³qª¾¦¬¥ó³B²z¤H­û
''            If UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '´ú¸Õ¸ê®Æ®w
''               strTo = m_M51Recver
''            Else
''               strTo = Pub_GetSpecMan("°ê¥~³¡«H¥ó³B²z¤H")
''            End If
'''            PUB_SendMail strUserNum, strTo, "", "inBound¦³ª÷Æ_«H¥ó¡I", °ê¥~³¡¦¬¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð³B²z¡I" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
'''                     "* ¶i¤J¨ä«H½c¸Ñ±K«áÂà±Hµ¹InBound¡A¦A±N­ì¥[±K¶l¥ó§R°£¡AÁ×§K­«ÂÐ¡]¤Á°O¡^¡A«Ý¨t²Î¤U¦¸´`Àô³B²z¡C", , , , , , , , , , , False
''            PUB_SendMail strUserNum, strTo, "", °ê¥~³¡¦¬¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð³B²z¡I", strIPMNoteSMIME & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
''                     "* ¶i¤J¨ä«H½c¸Ñ±K«áÂà±Hµ¹InBound¡A¦A±N­ì¥[±K¶l¥ó§R°£¡AÁ×§K­«ÂÐ¡]¤Á°O¡^¡A«Ý¨t²Î¤U¦¸´`Àô³B²z¡C", , , , , , , , , , , False
''            DoEvents
''         End If
'
'         '°O¿ýLogÀÉ
'         'Add By Sindy 2024/1/31
'         If intFolder = 1 Then
'         '2024/1/31 END
'            '" and MRL05='" & strMRL05 & "'"
'            strSql = "update MailReceiveLog set" & _
'                     " MRL04=" & Format(strMRL04, "hhmmss") & _
'                     ",MRL06=" & intRunOK & ",MRL07=" & intKeyCnt & ",MRL08=" & intCaseOK & _
'                     ",MRL09='" & IIf(bolCancel(0) = True, "B", "E") & "'" & _
'                     " where MRL01='" & strMRL01 & "'" & _
'                     " and MRL02=" & strMRL02 & _
'                     " and MRL03=" & Format(strMRL03, "hhmmss")
'            cnnConnection.Execute strSql
'            m_RunFCPinStarTime = strMRL03
'            m_RunFCPinEndTime = Format(strMRL04, "hh:mm:ss")
'         End If
'         If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then GoTo NotRunSec 'Add By Sindy 2023/2/18
'
'         'Add By Sindy 2017/8/8 °õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡A­Y¥u³Ñ¤U¥[±K¶l¥ó´Nµo«H³qª¾°ê¥~³¡¶l¥ó³B²z¤H­û
'         '                      ¦³«D¥[±K¶l¥ó¦A°õ¦æ¤@¦¸±µ¦¬
''         DoEvents
'         Set myItems = myFolder.Items
'         intMaxItem = myItems.Count
'         If intMaxItem > 0 Then
'            strErrText = "": intKeyCnt = 0
'            For mail_ii = myItems.Count To 1 Step -1
'               Call ReadMailText(myItems, False)
'               'Modify By Sindy 2017/11/17
'               'Modify By Sindy 2020/4/10 + IPM.Outlook.Recall
'               If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or _
'                  InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
'               'If myItems.Item(mail_ii).Class <> 43 Then
'               '2017/11/17 END
'                  If strErrText = "" Then
'                     strErrText = "***¡@(inbound) °õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡@*********************************" & vbCrLf
'                  End If
'                  intKeyCnt = intKeyCnt + 1
'                  strErrText = strErrText & "²Ä¡@" & mail_ii & "¡@µ§¡@[¥[±K]¡@¥D¦®:¡@" & strSocSubject & vbCrLf
'               Else
'                  If bolReStarFCPin = False And bolCancel(0) = False Then
'                     bolReStarFCPin = True
'                     Call WLog_Day("[­«Run²Ä¤G¦¸]" & vbCrLf, °ê¥~³¡¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                     '­«Run²Ä¤G¦¸
'                     GoTo ReStarFCPin
'                  'Add By Sindy 2022/8/5 ¤¤Â_´N¤£­n¦AÀË¬d¤F,©¹¤U°õ¦æ
'                  ElseIf bolCancel(0) = True Then
'                     Exit For
'                  '2022/8/5 END
'                  End If
'               End If
'            Next mail_ii
'
'            If strErrText <> "" Then
''               strErrText = strErrText & "*** END ************************************************************" & vbCrLf
''               Call WLog(strErrText)
'               '¦³¥[±K«H¥ó¥B¬°¤u§@¤Ñ¤~­n±H«H³qª¾¤H­û³B²z
'               If ChkWorkDay(strSrvDate(1)) = True And _
'                  (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
'                  '±HE-Mail³qª¾¦¬¥ó³B²z¤H­û
'                  If UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '´ú¸Õ¸ê®Æ®w
'                     strTo = m_M51Recver
'                  Else
'                     strTo = Pub_GetSpecMan("°ê¥~³¡«H¥ó³B²z¤H")
'                  End If
'                  PUB_SendMail strUserNum, strTo, "", °ê¥~³¡¦¬¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð³B²z¡I", strIPMNoteSMIME & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
'                        "* ¶i¤J¨ä«H½c¸Ñ±K«áÂà±Hµ¹InBound¡A¦A±N­ì¥[±K¶l¥ó§R°£¡AÁ×§K­«ÂÐ¡]¤Á°O¡^¡A«Ý¨t²Î¤U¦¸´`Àô³B²z¡C", , , , , , , , , , , False, , , False, , , False
''                  DoEvents
'               End If
'            End If
'         End If
'         '2017/8/8 END
'      End If 'Add By Sindy 2024/1/31
'   Next intFolder 'Add By Sindy 2024/1/31
'
'NotRunSec:
'      If intRunOK > 0 Then 'Add By Sindy 2024/1/31
'         Call PUB_SendMailCache 'Add By Sindy 2019/7/17
'         'Modify By Sindy 2017/12/27 ¤u§@¤Ñ¤~­n³qª¾
'         If ChkWorkDay(strSrvDate(1)) = True And _
'            (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
'            'ÀË¬d¦¬¥ó¸ê®Æ§¨¤¤¬O§_¦³´Ý¯dÀÉ®×
'            Set oFolder = oFileSys.GetFolder(txtPathIPDept.Text)
'            Set fs = CreateObject("Scripting.FileSystemObject")
'            If oFolder.files.Count > 0 Then
'               'Add By Sindy 2023/9/13
'               For Each oFile In oFolder.files
'                  Set myItems = olApp.CreateItemFromTemplate(txtPathIPDept.Text & "\" & oFile.Name)
'                  Call ReadMailText_File(myItems)
'                  '¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶×¤J?­Y¦³=§R°£¡C­Y¨S¦³=¤£³B²z,µ¥¤H­û¬d¬Ý
'                  strSql = "select ii01,ii03 from ipdeptinput" & _
'                           " where ii17 = '" & ChgSQL(strSocSubject) & "'" & _
'                           " and ii11 = '" & ChgSQL(strSender) & "' and ii12 = " & DBDATE(strMailDate) & " and ii13 = " & Val(Replace(strMailTime, ":", "")) & _
'                           " order by ii01 desc,ii03 desc"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     '§R°£PCºÝÀÉ®×
'                     Call fs.DeleteFile(txtPathIPDept & "\" & oFile.Name)
'                     Sleep 1000
'                     DoEvents
'                  End If
'               Next
'               Set oFolder = oFileSys.GetFolder(txtPathIPDept.Text)
'               If oFolder.files.Count > 0 Then
'               '2023/9/13 END
'                  PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "°ê¥~³¡¦¬¥ó¸ê®Æ§¨:" & txtPathIPDept.Text & "©|¦³´Ý¯dÀÉ®×(" & oFolder.files.Count & "­Ó),½ÐÀË¬d¡I", "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
'               End If
'            End If
'            'Add By Sindy 2017/11/16 ÀË¬d¬O§_¦³«H¥ó¥¼Âà±H
'            If UCase(pub_DbTerminalName) = ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '¥¿¦¡¸ê®Æ®w¤~µo«H
'               strExc(0) = "SELECT COUNT(*) FROM ipdeptinput WHERE ii08=0"
'               intI = 1
'               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'               If rsA.Fields(0) > 0 Then
'                  'Add By Sindy 2019/11/14 °ê¥~³¡¥D¦®¸Ì¦³ URGENT ¦r¼ËªÌ,³qª¾«H­n¥[¦³«æ¥ó! => IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") &
'                  intURGENT = 0
'                  strExc(0) = "SELECT COUNT(*) FROM ipdeptinput WHERE ii08=0 and instr(upper(ii17),'URGENT')>0"
'                  intI = 1
'                  Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'                  If rsA.Fields(0) > 0 Then
'                     intURGENT = rsA.RecordCount
'                  End If
'                  '2019/11/14 END
'                  'Modify By Sindy 2017/7/20 77015==>Pub_GetSpecMan("°ê¥~³¡«H¥ó³B²z¤H")
'                  'Modify By Sindy 2019/11/14 + IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") &
'                  PUB_SendMail strUserNum, Pub_GetSpecMan("°ê¥~³¡«H¥ó³B²z¤H"), "", IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") & "ª`·N¡G" & °ê¥~³¡¦¬¥ó«H½c & "©|¦³¥¼Âà±H«H¥ó«Ý³B²z¡I", "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
''                  DoEvents
'               End If
'            End If
'            '2017/11/16 END
'
'            'Modify By Sindy 2018/10/29 «H¥ó¦³¿ò¥¢,Âà±H¸ê°T¥¿±`,¦ý½T¹ê±H«H³Æ¥÷ºô­¶¨t²Î§ä¤£¨ì«H¥ó
'            'select ii08,ii09,ii20,ii21,ii22,ii17 from ipdeptinput where ii01='20181025' and ii03 in('F0292','F0304','F0293','F0262');
'            '/*
'            '      II08       II09 II20                       II21       II22 II17
'            '---------- ---------- -------------------- ---------- ---------- --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'            '  20181025     141308 Y                      20181025     141310 ¥¼¶Ç»¼ªº¥D¦®: Mail Delivery Failure
'            '  20181026     143250 Y                      20181026     143256 Mail Delivery Failure
'            '  20181026     143249 Y                      20181026     143255 IMPORTANT NOTICE
'            '  20181026     143249 Y                      20181026     143254 Out of Office Notice
'            '*/
'            strExc(0) = "select count(*) from ipdeptinput where ii20<>'Y' and ii20 is not null" & _
'                        " and ii01>=20181001" & _
'                        " order by ii01,ii02"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               If RsTemp.Fields(0) > 0 And ChkWorkDay(strSrvDate(1)) = True Then
''                  PUB_SendMail strUserNum, "97038", "", "¡iTaRevOutLook¡jÀË¬d«H¥ó¬O§_¦³¿ò¥¢(" & RsTemp.Fields(0) & "µ§)", strExc(0), , , , , , , , , , , False, , , False, , , False
'               End If
'            End If
'            '2018/10/29 END
'         End If
'
'         'Add By Sindy 2022/5/25
'         '±Hµo³qª¾«H
'         If m_strMailTo <> "" Then
'            '°Ï¤À³¡ªù
'            strF1xEmp = "": strF2xEmp = ""
'            varTmp = Split(m_strMailTo, ";")
'            For jj = 0 To UBound(varTmp)
'               If Left(PUB_GetST03(CStr(varTmp(jj))), 2) = "F1" Then '¥~°Ó
'                  strF1xEmp = strF1xEmp & ";" & varTmp(jj)
'               Else
'                  strF2xEmp = strF2xEmp & ";" & varTmp(jj)
'               End If
'            Next jj
'            'Call PUB_SendNotifyMail(m_strMailTo)
'            If strF1xEmp <> "" Then
'               strF1xEmp = Mid(strF1xEmp, 2)
'               Call PUB_SendNotifyMail(strF1xEmp)
'            End If
'            If strF2xEmp <> "" Then
'               strF2xEmp = Mid(strF2xEmp, 2)
'               Call PUB_SendNotifyMail(strF2xEmp)
'            End If
'         End If
'      Else
'         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         '°O¿ýLogÀÉ
'         strSql = "update MailReceiveLog set" & _
'                  " MRL04=" & Format(strMRL04, "hhmmss") & _
'                  ",MRL06=" & intRunOK & ",MRL07=" & intKeyCnt & ",MRL08=" & intCaseOK & _
'                  ",MRL09='" & IIf(bolCancel(0) = True, "B", "E") & "'" & _
'                  " where MRL01='" & strMRL01 & "'" & _
'                  " and MRL02=" & strMRL02 & _
'                  " and MRL03=" & Format(strMRL03, "hhmmss")
'         cnnConnection.Execute strSql
'         m_RunFCPinStarTime = strMRL03
'         m_RunFCPinEndTime = Format(strMRL04, "hh:mm:ss")
'      End If
''      Screen.MousePointer = vbDefault
'
'      txtMRL02 = strSrvDate(2)
'      Call cmdQuery_Click
'      Frame1.Caption = Frame1.Tag
'      DoEvents
'
''      'Add By Sindy 2023/11/29
''      Set eventConn = Nothing
''      WCmdLog "importFCPinBound µ²§ô"
''      WCmdLog ""
''      '2023/11/29 END
'   End If
'
'   cmdCancel(0).Enabled = False
'   '­n¤¤Â_
'   If bolCancel(0) = True Then
'      bolCancel(0) = False
'      TmrFCPin.Interval = 0: LblFCPin.BackColor = vbRed
'   Else
'   '¥¿±`µ²§ô
'      If TmrFCPin.Interval > 0 Then
'         TmrFCPin.Interval = dblTmrFCPin
'         LblFCPin.BackColor = vbGreen
'      Else
'         LblFCPin.BackColor = vbRed
'      End If
'   End If
'
'   importFCPinBound = True
'
'   Set olApp = Nothing
'   Set myNamespace = Nothing
'   Set myFolder = Nothing
'   Set myItems = Nothing
'   Set oFolder = Nothing
'   Set rsA = Nothing
'   Set fs = Nothing
'   Set oFile = Nothing
'
'   Exit Function
'
'ErrNo1:
'   Screen.MousePointer = vbDefault
'   'Resume
'   intErr2147024882 = ExportEMailErr(myItems, True, °ê¥~³¡¦¬¥ó«H½c, "(ErrNo1) " & strErrText & "; strSql=" & strSql, Err.Number, Err.Description, _
'                      strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'   On Error GoTo 0: Err.Clear
'   If intErr2147024882 > 0 Then
'      Call WLog_Day("intErr2147024882 > 0", °ê¥~³¡¦¬¥ó«H½c)
'      'Resume
'      'Resume Next
'      GoTo ReStarFCPin
'      Exit Function
'   End If
'
'   cmdCancel(0).Enabled = False
'   TmrFCPin.Interval = dblTmrFCPin: LblFCPin.BackColor = vbGreen
'
'   Set olApp = Nothing
'   Set myNamespace = Nothing
'   Set myFolder = Nothing
'   Set myItems = Nothing
'   Set oFolder = Nothing
'   Set rsA = Nothing
'   Set fs = Nothing
'   Set oFile = Nothing
'End Function

'Add By Sindy 2020/4/14
Private Function ChkIsOpenEmail(strFullFileName As String, ByRef strErrNumber As String, _
   ByRef strErrDesc As String) As Boolean
   
Dim objOutLook As Object
Dim objMail As Object

On Error GoTo ErrHand

   Set objOutLook = CreateObject("Outlook.Application")
   Set objMail = objOutLook.CreateItemFromTemplate(strFullFileName)
   
   ChkIsOpenEmail = True
   
   Set objMail = Nothing
   Set objOutLook = Nothing
   Exit Function

ErrHand:
   strErrNumber = Err.Number
   strErrDesc = Err.Description
   ChkIsOpenEmail = False
   
   Set objMail = Nothing
   Set objOutLook = Nothing
End Function

'¦^¶Ç bolIsEnd:¬O§_­nµ²§ô°õ¦æ
'     Integer:intErr2147024882µ§¼Æ
Private Function ExportEMailErr(ByVal f_myItems As Object, ByVal bolIsEnd As Boolean, ByVal strTimerName As String, _
   strErrText As String, strErrNumber As String, strErrDesc As String, _
   strMRL01 As String, strMRL02 As String, strMRL03 As String, strMRL04 As String, strMRL05 As String) As Integer
   
Dim strText As String 'Add By Sindy 2023/2/18
Dim ii As Integer

   ExportEMailErr = 0
   Call PUB_WriteDebugLog("strTimerName=" & strTimerName & vbCrLf & _
                          "strErrText=" & strErrText & vbCrLf & _
                          "strErrNumber=" & strErrNumber & vbCrLf & _
                          "strErrDesc=" & strErrDesc & vbCrLf & _
                          "strMRL01=" & strMRL01 & vbCrLf & _
                          "strMRL02=" & strMRL02 & vbCrLf & _
                          "strMRL03=" & strMRL03 & vbCrLf & _
                          "strMRL04=" & strMRL04 & vbCrLf & _
                          "strMRL05=" & strMRL05 & ";")    'Add By Sindy 2025/11/10
   
   'Add By Sindy 2024/4/12
   'Outlook¤£¯à°ÊµL¦^À³~ ³o¦¸§â¤À«H¨t²Î­«¶}, Outlook¨S°Ê;¤À«H®É·|¥X²{
   '  -2147418107:Automation ¿ù»~
   '  ¦b°T®§¿z¿ï¾¹¸Ì®É¤£¥i¹ï¥~©I¥s¡C
   'Modify By Sindy 2024/4/16
   '  -2147023170:Automation ¿ù»~
   '  »·ºÝµ{§Ç©I¥s¥¢±Ñ¡C
   'Modify By Sindy 2024/4/27 + (ErrNo1) ~ -2146959355:¦øªA¾¹°õ¦æ¥¢±Ñ
   If strErrNumber = "-2147418107" Or strErrNumber = "-2147023170" Or strErrNumber = "-2146959355" Then
      If strMRL01 = "01" Then
         TmrFCPin.Interval = 20000
      ElseIf strMRL01 = "02" Then
         TmrFCPout.Interval = 20000
      ElseIf strMRL01 = "03" Then
         TmrPatent.Interval = 20000
      ElseIf strMRL01 = "04" Then
         TmrTM.Interval = 20000
      'Add By Sindy 2024/5/16
      ElseIf strMRL01 = "05" Then
         TmrLAbackup.Interval = 20000
         '2024/5/16 END
      End If
      'Ãö³¬Outlook
      process_id = Shell("taskkill /F /IM outlook.exe", vbHide)
      For ii = 1 To 10
         If PUB_CheckIsRunning("outlook.exe") = True Then
            Sleep 1000
         Else
            Exit For
         End If
      Next
      Sleep 60000 '°±¸m1¤ÀÄÁ
      '¶}±ÒOutlook
      process_id = Shell("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE", vbHide)
      For ii = 1 To 10
         If PUB_CheckIsRunning("outlook.exe") = True Then
            Exit For
         Else
            Sleep 1000
         End If
      Next
      'Modify By Sindy 2024/4/27
      If Frame99.Tag = "" Then
         strExc(10) = strErrDesc & vbCrLf & "Outlook¦³­«·s±Ò°Ê, ÀË¬d¦³¥¿±`¤À«H¶Ü?"
         PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "¡i" & strErrNumber & "¡j" & strErrDesc, strExc(10) & vbCrLf & vbCrLf & strSocSubject, , , , , , , , , , , False, , , False, , , False
         Frame99.Tag = "Outlook¦³­«·s±Ò°Ê"
         
         WLog PUB_GetDbTerminal & "¡i" & strErrNumber & "¡j" & strExc(10)
         Call WLog_Day(PUB_GetDbTerminal & "¡i" & strErrNumber & "¡j" & strExc(10), strTimerName)
         'Call cmdCancel_Click(Cancel_idx)
         Sleep 60000 '°±¸m1¤ÀÄÁ
      End If
      '2024/4/27 END
      Exit Function
   End If
   '2024/4/12 END
   
   If mail_ii = 0 Then
      strText = strErrText & vbCrLf & IIf(strErrNumber <> "0", strErrNumber & ":" & strErrDesc, "") & vbCrLf
      WLog strText
      Call WLog_Day(strText, strTimerName) 'Add By Sindy 2020/11/9
      PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "¶×¤JOutLook«H½c(Err.1)(" & strTimerName & ")¦³°ÝÃD¡A½Ð¬d¬Ý¡I¡imail_ii = 0¡j", strSocSubject & vbCrLf & vbCrLf & strText, , , , , , , , , , , False, , , False, , , False
      DoEvents
   Else
      'Err.Number = "-2147352567 : ³¯¦C¯Á¤Þ¶W¥X¬É­­¡C
      'Err.Number = "-2147221233 : §@·~¥¢±Ñ¡C
      If strErrNumber <> "-2147352567" And strErrNumber <> "-2147221233" And _
         InStr(strErrText, "-2147352567") = 0 And InStr(strErrText, "-2147221233") = 0 Then
         strText = "(" & strTimerName & ")" & strMRL03 & " ~ " & strMRL04 & vbCrLf & _
                     "²Ä " & mail_ii & " µ§" & vbCrLf & _
                     "±H¥ó¤é´Á : " & strMailDate & vbCrLf & _
                     "±H¥ó®É¶¡ : " & strMailTime & vbCrLf & _
                     "±H¥óªÌ : " & strSender & vbCrLf & _
                     "¥D¦® : " & strSocSubject & vbCrLf & _
                     "strFileName : " & strFileName & vbCrLf & IIf(strErrText <> "", strErrText & vbCrLf, "")
         '***** ¥X²{¯S®íªº¿ù»~°T®§¨Ò¥~³B²z:
         'If InStr(strErrText, "-2147287038") > 0 Then 'msgÀÉ³QÀË¬d¨ì¦³¤¤¬r¯fª¬ªp
         'If strErrNumber = "-2147287038" Then
         'msgÀÉ®×³QÀË¬d¨ì¦³¤¤¬r¯fª¬ªp:
         'Modify By Sindy 2019/12/17 + or InStr(strErrText, "ºÃ¦ü¤¤¯f¬r") > 0
         If (InStr(strErrNumber, "-2147287038") > 0 And InStr(strErrDesc, "µLªk¶}±ÒÀÉ®×") > 0) Or _
            InStr(strErrText, "ºÃ¦ü¤¤¯f¬r") > 0 Then
            
            'Modify By Sindy 2019/12/17
            If InStr(strErrText, "ºÃ¦ü¤¤¯f¬r") = 0 Then
            '2019/12/17 END
               strText = strText & "@msgÀÉ®×³QÀË¬d¨ì¦³¤¤¬r¯fª¬ªp " & "strErrNumber:" & strErrNumber & " strErrDesc:" & strErrDesc '& vbCrLf
            End If
            
            If DeleteMyItems(f_myItems, strTimerName, strText) = True Then '§R°£Outlook¸Ì­±ªº¶l¥ó
               strText = strText & vbCrLf & "¡i«H¥ó¤w§R°£¡j"
            End If
            
            'DoEvents
            PUB_SendMail strUserNum, GetDeptMan("M51") & ";" & m_M51Recver, "", "¡i" & strTimerName & "¦³¯f¬r«H¡j" & strSocSubject, strText, , , , , , , , , , , False, , , False, , , False
            DoEvents
            If WLog_Day(Mid(strText, InStr(strText, "±H¥ó¤é´Á")), strTimerName) = True Then
               WLog strText
            End If
         
         'Modify By Sindy 2020/4/14
         '-2147168237:¦b¦¹¤u§@¶¥¬q¤¤µLªk±Ò°Ê§ó¦hªº²§°Ê¡C
         '-2147287035:§Ú­ÌµLªk¶}±Ò 'C:\IPDept\2.20200413110808.msg'¡C³o¥i¯à¬O¦]¬°¸ÓÀÉ®×¤w¶}±Ò¡A©Î¬O±z¨S¦³Åv­­¥i¶}±Ò¸ÓÀÉ®×
         '-2147287008:§Ú­ÌµLªk¶}±Ò 'C:\IPDept\16.20200413110103.msg'¡C³o¥i¯à¬O¦]¬°¸ÓÀÉ®×¤w¶}±Ò¡A©Î¬O±z¨S¦³Åv­­¥i¶}±Ò¸ÓÀÉ®×
         'Modify By Sindy 2020/4/16 999:C:\IPDept\17.20200416100933.msgÀÉ®×¤W¶Ç¥¢±Ñ¡I (strErrNumber = "999" And InStr(strErrDesc, "ÀÉ®×¤W¶Ç¥¢±Ñ") > 0)
         'ÄdºI°T®§«O¯d«H¥ó¤H¤u³B²z
         ElseIf strErrNumber = "-2147024882" Or strErrNumber = "-2147221233" Or _
            (InStr(strErrDesc, "§Ú­ÌµLªk¶}±Ò") > 0 And InStr(strErrDesc, "³o¥i¯à¬O¦]¬°¸ÓÀÉ®×¤w¶}±Ò¡A©Î¬O±z¨S¦³Åv­­¥i¶}±Ò¸ÓÀÉ®×") > 0) Then
            
            f_myItems.Item(mail_ii).FlagRequest = "«Ý³B²z"
            'strText = strText & "µLªk±N¶l¥ó¥t¦s¦¨MsgÀÉ,Âà¤J¥¢±Ñ(¬õ¦â¼Ð¼m:«Ý³B²z),½Ð¤H¤u¶×¤J" & vbCrLf
            strText = strText & "µ{¦¡µLªk³B²z¡A»Ý¤H¬°¤¶¤JÀË¬d­ì¦]¤Î³B²z¡C" & vbCrLf
            If WLog_Day(Mid(strText, InStr(strText, "±H¥ó¤é´Á")), strTimerName) = True Then
               WLog strText
            End If
            'intErr2147024882 = mail_ii
            'Resume Next
            PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "¶×¤JOutLook«H½c(Err.2)(" & strTimerName & ")¦³°ÝÃD¡A½Ð¬d¬Ý¡I", strText, , , , , , , , , , , False, , , False, , , False
            DoEvents
            
            ExportEMailErr = mail_ii
            Exit Function
            
         'Add By Sindy 2019/2/14
         ElseIf strErrNumber = "999" Or _
            InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then 'µLªk»PFTP Server«Ø¥ß³s½u
            
            strText = strText & IIf(strErrNumber <> "0", strErrNumber & ":" & strErrDesc, "") & vbCrLf
            WLog strText
            Call WLog_Day(strText, strTimerName) 'Add By Sindy 2020/11/9
            PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "¶×¤JOutLook«H½c(Err.3)(" & strTimerName & ")¦³°ÝÃD¡AµLªk»PFTP Server«Ø¥ß³s½u¡I", strText & vbCrLf & "½Ð¦Ü" & pub_HostName & "¹q¸£Ãö³¬¿ù»~°T®§¨Ã½T»{«H¥óª¬ªp¡C", , , , , , , , , , , False, , , False, , , False
            DoEvents
            Call cmdCancel_Click(Cancel_idx)
            'Add By Sindy 2022/9/14
            'Sleep¨Ï¥Î¤èªk:
            '³æ¦ì:²@¬í
            '1000²@¬í = 1¬í
            'Sleep 100  '100¬°©µ¿ð
            Sleep 1000 * 30 '30¬í
            '2022/9/14 END
         '2019/2/14 END
         
         Else
            strText = strText & IIf(strErrNumber <> "0", strErrNumber & ":" & strErrDesc, "") & vbCrLf & vbCrLf & _
                  "ª`·N¡GÀË¬d¦³°ÝÃDªº«e«á«H¥ó¡A½T»{«H¥ó¬O§_¦³§¹¾ã±µ¦¬¦Ü¨t²Î¤¤¡C" & vbCrLf & vbCrLf & _
                  "¡iÀË¬d«H¥ó­Y¤wÂà¤J¦¨¥\¡A§Y¥i©¿²¤¦¹¶l¥ó¡j" & vbCrLf & vbCrLf & _
                  "³Æµù¡GLog¤å¦rÀÉ¦s©ñ¦ì¸m¡G(" & pub_HostName & ") " & App.path & "\TaOutLookLog\" & vbCrLf
            WLog strText
            Call WLog_Day(strText, strTimerName) 'Add By Sindy 2020/11/9
            PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "¶×¤JOutLook«H½c(Err.4)(" & strTimerName & ")¦³°ÝÃD¡A½Ð¬d¬Ý¡I", strText, , , , , , , , , , , False, , , False, , , False
            DoEvents
         End If
         '***** END
      Else
         Call WLog_Day("ExportEMailErr:bolIsEnd=" & IIf(bolIsEnd = False, "F; ", "T; ") & strErrNumber & ":" & strErrDesc, strTimerName) 'Add By Sindy 2020/11/9
         PUB_SendMail strUserNum, m_M51Recver, "", "ExportEMailErr:bolIsEnd=" & IIf(bolIsEnd = False, "F; ", "T; ") & strErrNumber & ":" & strErrDesc, _
            "strTimerName=" & strTimerName & " strMRL01=" & strMRL01 & " strMRL02=" & strMRL02 & " strMRL03=" & strMRL03 & " strMRL04=" & strMRL04 & " strMRL05=" & strMRL05 & vbCrLf & _
            "strErrText = " & strErrText & vbCrLf & _
            "strSocSubject = " & strSocSubject & vbCrLf & _
            "strSender = " & strSender & vbCrLf & _
            "strMailDate = " & strMailDate & vbCrLf & _
            "strMailTime = " & strMailTime & vbCrLf, , , , , , , , , , , False, , , False, , , False
         DoEvents
      End If
      If bolIsEnd = False Then Exit Function
   End If
   
   If strMRL02 <> "" Then
      '°O¿ýLogÀÉ
      '" and MRL05='" & strMRL05 & "'"
      strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
      strSql = "update MailReceiveLog set" & _
               " MRL04=" & Format(strMRL04, "hhmmss") & _
               ",MRL06=" & intRunOK & ",MRL07=" & intKeyCnt & ",MRL08=" & intCaseOK & _
               ",MRL09='F'" & _
               " where MRL01='" & strMRL01 & "'" & _
               " and MRL02=" & strMRL02 & _
               " and MRL03=" & Format(strMRL03, "hhmmss")
      cnnConnection.Execute strSql
   End If
End Function

'Add By Sindy 2020/11/13 §R°£Outlook¸Ì­±ªº¶l¥ó
Private Function DeleteMyItems(ByVal f_myItems As Object, ByVal strTimerName As String, Optional strContext As String = "") As Boolean
Dim strSubject_E As String
Dim strTmp As String

   DeleteMyItems = False

   strSubject_E = f_myItems.Item(mail_ii).Subject
   Call WLog_Day("strSubject_E = " & strSubject_E, strTimerName) 'Add By Sindy 2022/2/24
   strTmp = "strSocSubject¡G" & strSocSubject & vbCrLf & "strSubject_E¡G" & strSubject_E
   Call WLog_Day("strSocSubject = " & strSocSubject, strTimerName) 'Add By Sindy 2022/2/24
   Call WLog_Day(IIf(strContext <> "", strContext, "¤À«H¦¨¥\¡A±ý§R°£¶l¥ó"), strTimerName)

   If strSocSubject <> strSubject_E Then
      PUB_SendMail strUserNum, m_M51Recver, "", "¡i" & strTimerName & " §R°£«H¥ó®É¡Aµo²{¥D¦®¤£¤@­P¡j" & strSocSubject, strTmp, , , , , , , , , , , False, , , False, , , False
      Call WLog_Day("§R°£¶l¥ó®É¡Aµo²{¥D¦®¤£¤@­P(" & mail_ii & "):" & vbCrLf & strTmp & vbCrLf, strTimerName)
   Else
      f_myItems.Item(mail_ii).Delete '§R°£
      Call WLog_Day("§R°£¶l¥ó(" & mail_ii & "):" & strSocSubject & vbCrLf, strTimerName)
      DeleteMyItems = True
   End If
   DoEvents
End Function

'°ê¥~³¡±H¥ó«H½c³B²zµ{§Ç
Private Sub TmrFCPout_Timer()
   'Modify By Sindy 2024/5/14
   Call ChkExecutionTimer(Left(IPDept±H¥ó§X, 2))
   Exit Sub
   '2024/5/14 END

'Dim strTo As String
'Dim oFileSys As New FileSystemObject, oFolder As Object
'Dim fs
'Dim strIPMNoteSMIME As String '¥[±K¥D¦®
'Dim bolForKeyWordDel As Boolean, ii As Integer
'Dim bolReStarFCPout As Boolean
'Dim strMRL01 As String, strMRL02 As String, strMRL03 As String, strMRL04 As String, strMRL05 As String
'Dim rsA As New ADODB.Recordset
'Dim kk As Integer ', strRecipients As String
'Dim strErrNumber As String 'Add By Sindy 2019/10/14
'Dim strErrCode As String, strErrDesc As String 'Add By Sindy 2020/4/15
'Dim strII01 As String, strII03 As String, strIR04 As String
''Add By Sindy 2023/6/26
'Dim olApp As Object
'Dim myNamespace As Object
'Dim myFolder As Object
'Dim myItems As Object
''2023/6/26 END
'Dim intFolder As Integer '­nÅª¨úªºFolder¼Æ; ex:Inbox ©M Junk Email
'
'On Error GoTo ErrNo1
'
'   If cnnConnection.State = adStateClosed Then Exit Sub '±ß¤WDBÂ_½u,¤£»Ý©¹¤U°õ¦æ
'   '¥H§KTimer¦P®ÉRun°_¨Ó
'   If LblFCPin.BackColor = vbBlue Then Exit Sub
'   If LblPatent.BackColor = vbBlue Then Exit Sub
'   If LblFCPout.BackColor = vbBlue Then Exit Sub
'   If LblTM.BackColor = vbBlue Then Exit Sub
'
'   strErrText = "" 'Add By Sindy 2020/7/22
''   If MsgBox("¬O§_­n¶×¤J" & °ê¥~³¡±H¥ó«H½c & "«H¥ó¡H", vbExclamation + vbYesNo + vbDefaultButton2, "­«­n°T®§¡I") = vbNo Then
''      TmrFCPout.Interval = 0
''      Exit Sub
''   End If
'
'   If txtPathIPDeptOut = "" Then
'      MsgBox "±H¥ó¸ê®Æ§¨¤£¥iªÅ¥Õ¡I"
'      txtPathIPDeptOut.SetFocus
'      Exit Sub
'   End If
'   If Dir(txtPathIPDeptOut, vbDirectory) = "" Then
'      MkDir txtPathIPDeptOut
'   End If
'
'   strMRL01 = Left(IPDept±H¥ó§X, 2): strMRL02 = "": strMRL03 = ""
'   If ExecuteSchedule(strMRL01, strMRL02, strMRL03) = True Or bolFCPoutRun = True Then '­n°õ¦æTimer
''      'Add By Sindy 2023/11/29
''      Set eventConn = cnnConnection
''      KillCmdLog
''      '2023/11/29 END
'
'      bolFCPoutRun = False
'
'      strSql = "Run:1 " 'debug
'      Set olApp = CreateObject("Outlook.Application")
'      strSql = "Run:2 " 'debug
'      Set myNamespace = olApp.GetNamespace("MAPI")
'      intKeyCnt = 0: intRunOK = 0: intCaseOK = 0
'
'strSql = "Run:3 " 'debug
'   'Add By Sindy 2024/1/31
'   For intFolder = 1 To 1 '2
'      'Modify By Sindy 2023/7/17
'      If OpenOutLookFolder(myNamespace, myFolder, Left(IPDept±H¥ó§X, 2), intFolder) = False Then
'         Set olApp = Nothing
'         Set myNamespace = Nothing
'         Set myFolder = Nothing
'         TmrFCPout.Interval = 0
'         LblFCPout.BackColor = vbRed
'         Exit Sub
'      End If
'      '2023/7/17 END
'
'      bolReStarFCPout = False
'
'      strSql = "Run:7 " 'debug
'
'ReStarFCPout:
''      Screen.MousePointer = vbHourglass
'      Set myItems = myFolder.Items
'      strSql = "Run:8 " 'debug
'      strIPMNoteSMIME = "" '¥[±K¥D¦®
'      intMaxItem = myItems.Count
'
'      '°O¿ýLogÀÉ
'      'Modify By Sindy 2024/1/31 + And intFolder = 1
'      If strMRL02 = "" And intFolder = 1 Then
'         'strMRL01 = Left(IPDept±H¥ó§X, 2)
'         strMRL02 = strSrvDate(1)
'         strMRL03 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         strMRL05 = strUserNum
'         strSql = "insert into MailReceiveLog(MRL01,MRL02,MRL03,MRL05,MRL09)" & _
'                  "values('" & strMRL01 & "'," & strMRL02 & "," & Format(strMRL03, "hhmmss") & ",'" & strMRL05 & "','Y')"
'         cnnConnection.Execute strSql
'      End If
'      strSql = "Run:9 " & intMaxItem 'debug
'      '*****
'      'intMaxItem = 0 'Add By Sindy 2024/2/20 backup¦³°ÝÃD,¼t°Ó¥¿¦b§ä°ÝÃD¤¤,¥ý¼È°±¨t²Î³B²z
'      '*****
'      If intMaxItem > 0 Then
'         Set fs = CreateObject("Scripting.FileSystemObject")
'         For mail_ii = myItems.Count To 1 Step -1
'            LblFCPout.BackColor = vbBlue 'ÂÅ¦âTimer¥¿¦bRun
'            cmdCancel(1).Enabled = True
'            DoEvents
'            Frame2.Caption = Frame2.Tag & "¡@¡@¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii
'            DoEvents
'            strErrText = ""
'            intRunOK = intRunOK + 1 '°O¿ý±µ¦¬µ§¼Æ (2017/7/20¤~¶}©l°O¿ý¥þ³¡±µ¦¬ªºµ§¼Æ)
'            Call ReadMailText(myItems, False)
'            'DATEDIFF("n", strMailTime, format(time,"HH:MM:SS")) '­pºâ®É¶¡®t´X¤ÀÄÁ
'
'            'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'            strErrText = "²Ä " & mail_ii & " µ§ ¥D¦®: " & strSocSubject & vbCrLf
'            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strSender: " & strSender & vbCrLf
'            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strMailDateTime: " & strMailDate & " " & strMailTime
'            Call WLog_Day(strErrText, °ê¥~³¡±H¥ó«H½c)
'
''            strSocSubject = myItems.Item(mail_ii).Subject
''            Text2.Text = myItems.Item(mail_ii).Subject
''            strMailSubject = Text2.Text
''            strMailDate = "": strMailTime = "": strSender = ""
'            'Modify By Sindy 2018/5/30 IPM.RECALL.REPORT.FAILURE = Message Recall Failure.µLªk¦^¦¬
'            'Modify By Sindy 2023/7/12 + Or myItems.Item(mail_ii).Class = 45 : ·s³qª¾ => UCase(myItems.Item(mail_ii).MessageClass) = UCase("IPM.Post")
'            If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.RECALL.REPORT.FAILURE")) > 0 Or myItems.Item(mail_ii).Class = 45 Then
'               intKeyCnt = intKeyCnt + 1
'               'Add By Sindy 2017/7/18 ¥[Log°O¿ý
'               'strErrText = "²Ä " & mail_ii & " µ§ [µLªk¦^¦¬] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[µLªk¦^¦¬¶l¥ó]" & vbCrLf, °ê¥~³¡±H¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'            'Add By Sindy 2019/9/23 [¥¼¶Ç»¼ªº¥D¦®] ¥D¦®: ¤wÅª¨ú: Certified AML & CFT Regulatory Compliance, Surveillance and Reporting Specialist; Taiwan
'            ElseIf myItems.Item(mail_ii).Class = 46 Then 'REPORT.IPM.Note.IPNRN
'               intKeyCnt = intKeyCnt + 1
'               'strErrText = "²Ä " & mail_ii & " µ§ [¥¼¶Ç»¼ªº¥D¦®] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'
'               Call DeleteMyItems(myItems, °ê¥~³¡±H¥ó«H½c, "[¥¼¶Ç»¼ªº¥D¦®] => §R°£") '§R°£Outlook¸Ì­±ªº¶l¥ó
'
'            'IPM.Note.SMIME ¥[±K
'            'Modify By Sindy 2017/11/17
'            ElseIf InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Then
'            'If myItems.Item(mail_ii).Class <> 43 Then
'            '2017/11/17 END
'               intKeyCnt = intKeyCnt + 1
'               'Add By Sindy 2017/7/18 ¥[Log°O¿ý
'               'strErrText = "²Ä " & mail_ii & " µ§ [¥[±K] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[¥[±K¶l¥ó]" & vbCrLf, °ê¥~³¡±H¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf '¥[±K¥D¦®
'               '2017/7/18 END
'            'Add By Sindy 2020/4/10 ¦^¦¬¶l¥ó,ª½±µ§R°£
'            ElseIf InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
'               intKeyCnt = intKeyCnt + 1
'               'strErrText = "²Ä " & mail_ii & " µ§ [¦^¦¬] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[¦^¦¬¶l¥ó]" & vbCrLf, °ê¥~³¡±H¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'               'myItems.Item(mail_ii).Delete '§R°£ =>µLªk§R°£,·|·í
'               'DoEvents
'            Else
'               'Add By Sindy 2022/6/27 ¨R¾P¦^«H
'               strExc(0) = "select ii01,ii03,ii28,ir04 from IPDeptinput,InputRecord" & _
'                           " where Ii28 is not null" & _
'                             " and Ii01=Ir01 and Ii03=Ir03 and Ir08=0" & _
'                             " and instr('" & ChgSQL(myItems.Item(mail_ii).Subject) & "',Ii28)>0" & _
'                             " and ir16='9'" '9.¦^«H
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strII01 = RsTemp.Fields("ii01")
'                  strII03 = RsTemp.Fields("ii03")
'                  strIR04 = RsTemp.Fields("ir04")
'                  '¼W¥[³¡ªù§PÂ_
'                  strExc(0) = "update InputRecord set ir08=" & strSrvDate(1) & ",ir09=" & Right("000000" & ServerTime, 6) & ",ir10='" & strUserNum & "'" & _
'                              " where ir01=" & strII01 & _
'                                " and ir03='" & strII03 & "'" & _
'                                " and upper(ir04)=upper('" & ChgSQL(strIR04) & "')" & _
'                                " and ir08=0"
'                  cnnConnection.Execute strExc(0), intI
'
'                  '­Y«H¥ó¦¬¨üªÌ¥þ³¡¤w³B²z©Î¤w§R°£,¥DÀÉ´N¥i¥H±¾¤WmsgÀÉ§R°£¤é´Á,µ¥«ÝAutoBatchDay¤@­Ó¤ë«á§R°£¹êÅéÀÉ
'                  strExc(0) = "select ir01 from InputRecord" & _
'                              " where ir01=" & strII01 & _
'                                " and ir03='" & strII03 & "'" & _
'                                " and ir08=0"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 0 Then '«H¥ó¦¬¨üªÌ¥þ³¡¤w³B²z©Î¤w§R°£
'                     strExc(0) = "update IPDeptInput set" & _
'                                 " ii16=" & strSrvDate(1) & _
'                                 " where Ii01=" & strII01 & _
'                                   " and Ii03='" & strII03 & "'" & _
'                                   " and ii16=0"
'                     cnnConnection.Execute strExc(0), intI
'                  End If
'               End If
'               '2022/6/27 END
'
'               'Modify By Sindy 2017/8/8
'               'ÀË¬d¦³³]©w¦¬¨üªÌ¬°²QµØªºÃöÁä¦r¤¤¨äºô°ì²Å¦X¦¹¶l¥ó¦¬¥óªÌ®É¡A«H¥óª½±µ§R°£¤£¶i¨t²Î
'               bolForKeyWordDel = False
'               'If InStr(ChgSQL(strSender), GetPrjSalesNM("86013")) > 0 Then
'                  For ii = myItems.Item(mail_ii).Recipients.Count To 1 Step -1
''                     strSql = "select LK01 from ipdeptkeyword" & _
''                              " where LK12='F' and LK04='86013' and LK03='2'" & _
''                              " and instr(upper('" & Replace(myItems.Item(mail_ii).Recipients(ii).address, "'", "") & "'),upper(LK01))>0"
''                     intI = 1
''                     Set rsA = ClsLawReadRstMsg(intI, strSql)
''                     If intI = 1 Then
''                        bolForKeyWordDel = True
''                        Exit For
''                     End If
'                     strSql = "select LK01 from ipdeptkeyword" & _
'                              " where LK12='F' and LK04='86013' and LK03='2'" & _
'                              " and instr(upper('" & Replace(myItems.Item(mail_ii).Recipients(ii).Name, "'", "") & "'),upper(LK01))>0"
'                     intI = 1
'                     Set rsA = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 1 Then
'                        bolForKeyWordDel = True
'                        Exit For
'                     End If
'                  Next ii
'               'End If
'               If bolForKeyWordDel = True Then
'                  Call DeleteMyItems(myItems, °ê¥~³¡±H¥ó«H½c, "[§R°£] «H¥óª½±µ§R°£¤£¶i¨t²Î") '§R°£Outlook¸Ì­±ªº¶l¥ó
'
'               Else
'               '2017/8/8 END
'                  strFileName = strSrvDate(1) & Right("000000" & ServerTime, 6) & "." & mail_ii & ".msg"
'                  myItems.Item(mail_ii).SaveAs txtPathIPDeptOut & "\" & strFileName, 9 '9.Outlook Unicode¶l¥ó®æ¦¡.msg
'                  'Add By Sindy 2020/2/27
'                  Sleep 1000
'                  DoEvents
'                  '2020/2/27 END
'                  Call WLog_Day("²£¥Í¼È¦s¹q¤lÀÉ: " & txtPathIPDeptOut & "\" & strFileName, °ê¥~³¡±H¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'
'                  If intErr2147024882 <> mail_ii Then
'                     Me.TxtIPDept = strFileName
'
'                     'Add By Sindy 2018/4/12
'                     If Dir(txtPathIPDeptOut & "\" & strFileName) = "" Then
'                        strErrText = "µL²£¥Í¹q¤lÀÉ,ºÃ¦ü¤¤¯f¬r " & "Err.Number:" & Err.Number & Err.Description & vbCrLf
'                        Call ExportEMailErr(myItems, False, °ê¥~³¡±H¥ó«H½c, strErrText, Err.Number, Err.Description, _
'                              strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'                     'Add By Sindy 2020/4/14 ÀË¬d¹q¤lÀÉ¬O§_¥i¥H¥¿±`¶}±Ò
'                     ElseIf ChkIsOpenEmail(txtPathIPDeptOut & "\" & strFileName, strErrCode, strErrDesc) = False Then
'                        intKeyCnt = intKeyCnt + 1
'                        strErrText = "²Ä " & mail_ii & " µ§ [MsgµLªk¶}±Ò] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf & _
'                           txtPathIPDeptOut & "\" & strFileName & vbCrLf & _
'                           "Err.Number:" & strErrCode & strErrDesc & vbCrLf
'                        Call WLog_Day(strErrText, °ê¥~³¡±H¥ó«H½c)
'                        strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'                     Else
'                     '2018/4/12 END
'
'                        'Add By Sindy 2018/7/10 °ê»Ú·|Ä³¶l¥ó
'                        If PUB_IPDeptISDMail(Me, "1", m_strISDPath, txtPathIPDeptOut, strFileName, intCaseOK) = True Then
'                           Call DeleteMyItems(myItems, °ê¥~³¡±H¥ó«H½c, "¤À«H¦¨¥\¡A§R°£¶l¥ó => PUB_IPDeptISDMail(©¹¨Ó°O¿ý)") '§R°£Outlook¸Ì­±ªº¶l¥ó
'
'                        Else
'                        '2018/7/10 END
'                           Sleep 100 'Add By Sindy 2019/12/13
'
'                           '*****
'                           '¦s­ÓÀÉ®É¥D¦®¤£¥i¥H¦³\/:*?"<>|µ¥²Å¸¹
'                           'If IPDeptBackupMail(Me.TextII17.Text, txtPathIPDeptOut & "\" & strFileName, strFileName, strErrText, intCaseOK, strRecipients) = True Then
'                           If IPDeptBackupMail(Me.TextII17.Text, txtPathIPDeptOut & "\" & strFileName, strFileName, strErrText, intCaseOK) = True Then
'                              Call DeleteMyItems(myItems, °ê¥~³¡±H¥ó«H½c, "IPDeptBackupMail ³B²z§¹²¦¡A§R°£¶l¥ó => IPDeptBackupMail") '§R°£Outlook¸Ì­±ªº¶l¥ó
'
'                           Else
'                              strErrNumber = Err.Number 'Add By Sindy 2019/10/14
'                              Call WLog_Day("¤À«H¥¢±Ñ(1)" & strErrText, °ê¥~³¡±H¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                              'Add By Sindy 2019/12/11
'                              If InStr(strErrText, "§ä¤£¨ìÀÉ®×") > 0 Then
'                                 strErrText = "§ä¤£¨ìÀÉ®×,ºÃ¦ü¤¤¯f¬r"
''                                 myItems.Item(mail_ii).Delete '§R°£
''                                 DoEvents
'                              End If
'                              '2019/12/11 END
'                              'Add By Sindy 2020/4/6
'                              If Me.TextII17.Text <> "" Then
'                                 If InStr(strErrText, Me.TextII17.Text) = 0 Then
'                                    strErrText = strErrText & vbCrLf & Me.TextII17.Text & vbCrLf
'                                 End If
'                              End If
'                              '2020/4/6 END
'
'                              Call WLog_Day("¤À«H¥¢±Ñ(2): " & strErrText & ";" & Err.Number & ":" & Err.Description, °ê¥~³¡±H¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                              Call ExportEMailErr(myItems, False, °ê¥~³¡±H¥ó«H½c, strErrText, Err.Number, Err.Description, _
'                                 strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'                              'Add By Sindy 2019/10/14
'                              'If strErrNumber = "999" Then
'                              If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then
'                                 Call WLog_Day("¤À«H¥¢±Ñ(3): 999 " & strErrText & vbCrLf, °ê¥~³¡±H¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                                 Exit For
'                              End If
'                              '2019/10/14 END
'                           End If
'                        End If '2018/7/10 +
'                     End If
'                  'Modify By Sindy 2020/4/15
'                  Else
'                     intErr2147024882 = 0
'                  '2020/4/15 END
'                  End If
'               End If
'            End If
'            '¬O§_­n¤¤Â_
'            If bolCancel(1) = True Then
'               LblFCPout.BackColor = vbRed
'               DoEvents 'Add By Sindy 2024/5/7
'               GoTo IsCancel
'            End If
'         Next mail_ii
'
'IsCancel:
'         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         '¦³¥[±K«H¥ó¥B¬°¤u§@¤Ñ¤~­n±H«H³qª¾¤H­û³B²z
''         If intKeyCnt > 0 And ChkWorkDay(strSrvDate(1)) = True Then
''            '±HE-Mail³qª¾¦¬¥ó³B²z¤H­û
''            If UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '´ú¸Õ¸ê®Æ®w
''               strTo = m_M51Recver
''            Else
''               strTo = Pub_GetSpecMan("°ê¥~³¡«H¥ó³B²z¤H")
''            End If
''            PUB_SendMail strUserNum, strTo, "", "backup¦³ª÷Æ_«H¥ó¡I", °ê¥~³¡±H¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð³B²z¡I" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
''                     "* ¶i¤J¨ä«H½c¸Ñ±K«áÂà±Hµ¹Backup¡A¦A±N­ì¥[±K¶l¥ó§R°£¡AÁ×§K­«ÂÐ¡]¤Á°O¡^¡A«Ý¨t²Î¤U¦¸´`Àô³B²z¡C", , , , , , , , , , , False
''            PUB_SendMail strUserNum, strTo, "", °ê¥~³¡±H¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð³B²z¡I" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
''                     "* ¶i¤J¨ä«H½c¸Ñ±K«áÂà±Hµ¹Backup¡A¦A±N­ì¥[±K¶l¥ó§R°£¡AÁ×§K­«ÂÐ¡]¤Á°O¡^¡A«Ý¨t²Î¤U¦¸´`Àô³B²z¡C", , , , , , , , , , , False
''         End If
'
'         '°O¿ýLogÀÉ
'         'Add By Sindy 2024/1/31
'         If intFolder = 1 Then
'         '2024/1/31 END
'            '" and MRL05='" & strMRL05 & "'"
'            strSql = "update MailReceiveLog set" & _
'                     " MRL04=" & Format(strMRL04, "hhmmss") & _
'                     ",MRL06=" & intRunOK & ",MRL07=" & intKeyCnt & ",MRL08=" & intCaseOK & _
'                     ",MRL09='" & IIf(bolCancel(1) = True, "B", "E") & "'" & _
'                     " where MRL01='" & strMRL01 & "'" & _
'                     " and MRL02=" & strMRL02 & _
'                     " and MRL03=" & Format(strMRL03, "hhmmss")
'            cnnConnection.Execute strSql
'            m_RunFCPoutStarTime = strMRL03
'            m_RunFCPoutEndTime = Format(strMRL04, "hh:mm:ss")
'         End If
'         If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then GoTo NotRunSec 'Add By Sindy 2023/2/18
'
'         'Add By Sindy 2017/8/8 °õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡A­Y¥u³Ñ¤U¥[±K¶l¥ó´Nµo«H³qª¾¹q¸£¤¤¤ß¶l¥óºÞ²z­û
'         '                      ¦³«D¥[±K¶l¥ó¦A°õ¦æ¤@¦¸±µ¦¬
'         DoEvents
'         Set myItems = myFolder.Items
'         intMaxItem = myItems.Count
'         If intMaxItem > 0 Then
'            strErrText = "": intKeyCnt = 0
'            For mail_ii = myItems.Count To 1 Step -1
'               Call ReadMailText(myItems, False)
'               'Modify By Sindy 2017/11/17
'               'Modify By Sindy 2020/4/10 + IPM.Outlook.Recall
'               If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or _
'                  InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
'               'If myItems.Item(mail_ii).Class <> 43 Then
'               '2017/11/17 END
'                  'Modify By Sindy 2017/9/25
'                  '¦³¥[±K«H¥ó¥B¬°¤u§@¤Ñ¤~­n±H«H³qª¾¤H­û³B²z
'                  If ChkWorkDay(strSrvDate(1)) = True Then
'                  '2017/9/25 END
'                     If strErrText = "" Then
'                        strErrText = "***¡@(backup) °õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡@*********************************" & vbCrLf
'                     End If
'                     intKeyCnt = intKeyCnt + 1
'                     strErrText = strErrText & "²Ä¡@" & mail_ii & "¡@µ§¡@[¥[±K]¡@¥D¦®:¡@" & strSocSubject & vbCrLf
'                  End If
'               Else
'                  If bolReStarFCPout = False And bolCancel(1) = False Then
'                     bolReStarFCPout = True
'                     Call WLog_Day("[­«Run²Ä¤G¦¸]" & vbCrLf, °ê¥~³¡±H¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                     '­«Run²Ä¤G¦¸
'                     GoTo ReStarFCPout
'                  'Add By Sindy 2022/8/5 ¤¤Â_´N¤£­n¦AÀË¬d¤F,©¹¤U°õ¦æ
'                  ElseIf bolCancel(1) = True Then
'                     Exit For
'                  '2022/8/5 END
'                  End If
'               End If
'            Next mail_ii
'
'            If strErrText <> "" Then
'               strErrText = strErrText & "*** END ************************************************************" & vbCrLf
'               Call WLog(strErrText)
'               'Modify By Sindy 2017/12/27 ¤u§@¤Ñ¤~­n³qª¾
'               If ChkWorkDay(strSrvDate(1)) = True And _
'                  (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
'                  PUB_SendMail strUserNum, m_M51Recver, "", °ê¥~³¡±H¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð¥ý¼Ð°O¬°¤wÅª¨ú¦A§R°£ª÷Æ_«H¥ó¡I(¹q¸£¤¤¤ßª½±µ§R°£¦¹«Ê«H¥ó,§Y¥i¡I)", strErrText & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
'                           "* Backup«H½cªº¥[±K¶l¥ó¥Ñ¹q¸£¤¤¤ß¤H­û¦Ü«H½c¤º§R°£" & vbCrLf & _
'                           "  ¡A¥~±M¤H­û·|¦Û¦æ§â¥[±K«H¥ó¸Ñ±K«á¦A±H¤@¥÷¦ÜBackup«H½cÂk¨÷¥Î¡C" & _
'                           "* ª`·N:¡]¥ý¼Ð°O¬°¤wÅª¨ú==>Á×§K¦^¶Ç¥¼Åª¨ú§Y§R°£ªº¦^±ø¡^¦A§R°£ª÷Æ_«H¥ó", , , , , , , , , , , False, , , False, , , False
'                  DoEvents
'               End If
'            End If
'         End If
'         '2017/8/8 END
'      End If 'Add By Sindy 2024/1/31
'   Next intFolder 'Add By Sindy 2024/1/31
'
'NotRunSec:
'      If intRunOK > 0 Then 'Add By Sindy 2024/1/31
'         'Modify By Sindy 2017/12/27 ¤u§@¤Ñ¤~­n³qª¾
'         If ChkWorkDay(strSrvDate(1)) = True And _
'            (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
'            'ÀË¬d±H¥ó¸ê®Æ§¨¤¤¬O§_¦³´Ý¯dÀÉ®×
'            Set oFolder = oFileSys.GetFolder(txtPathIPDeptOut.Text)
'            If oFolder.files.Count > 0 Then
'               PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "°ê¥~³¡±H¥ó¸ê®Æ§¨:" & txtPathIPDeptOut.Text & "©|¦³´Ý¯dÀÉ®×(" & oFolder.files.Count & "­Ó),½ÐÀË¬d¡I", "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
'            End If
'         End If
'
'      Else
'         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         '°O¿ýLogÀÉ
'         strSql = "update MailReceiveLog set" & _
'                  " MRL04=" & Format(strMRL04, "hhmmss") & _
'                  ",MRL06=0,MRL07=0,MRL08=0" & _
'                  ",MRL09='E'" & _
'                  " where MRL01='" & strMRL01 & "'" & _
'                  " and MRL02=" & strMRL02 & _
'                  " and MRL03=" & Format(strMRL03, "hhmmss")
'         cnnConnection.Execute strSql
'         m_RunFCPoutStarTime = strMRL03
'         m_RunFCPoutEndTime = Format(strMRL04, "hh:mm:ss")
'      End If
''      Screen.MousePointer = vbDefault
'
'      txtMRL02 = strSrvDate(2)
'      Call cmdQuery_Click
'      Frame2.Caption = Frame2.Tag
'      DoEvents
'
''      'Add By Sindy 2023/11/29
''      Set eventConn = Nothing
''      WCmdLog "TmpFCPout µ²§ô"
''      WCmdLog ""
''      '2023/11/29 END
'   End If
'
'   cmdCancel(1).Enabled = False
'   '­n¤¤Â_
'   If bolCancel(1) = True Then
'      bolCancel(1) = False
'      TmrFCPout.Interval = 0: LblFCPout.BackColor = vbRed
'   Else
'   '¥¿±`µ²§ô
'      If TmrFCPout.Interval > 0 Then
'         TmrFCPout.Interval = dblTmrFCPout
'         LblFCPout.BackColor = vbGreen
'      Else
'         LblFCPout.BackColor = vbRed
'      End If
'   End If
'
'   Set olApp = Nothing
'   Set myNamespace = Nothing
'   Set myFolder = Nothing
'   Set myItems = Nothing
'   Set fs = Nothing
'   Set oFolder = Nothing
'   Set rsA = Nothing
'
'   Exit Sub
'
'ErrNo1:
'   Screen.MousePointer = vbDefault
'   intErr2147024882 = ExportEMailErr(myItems, True, °ê¥~³¡±H¥ó«H½c, "(ErrNo1) " & strErrText & "; strSql=" & strSql, Err.Number, Err.Description, _
'                        strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'   On Error GoTo 0: Err.Clear
'   If intErr2147024882 > 0 Then
'      Call WLog_Day("intErr2147024882 > 0", °ê¥~³¡±H¥ó«H½c)
'      'Resume Next
'      GoTo ReStarFCPout
'      Exit Sub
'   End If
'
'   cmdCancel(1).Enabled = False
'   TmrFCPout.Interval = dblTmrFCPout: LblFCPout.BackColor = vbGreen
'
'   Set olApp = Nothing
'   Set myNamespace = Nothing
'   Set myFolder = Nothing
'   Set myItems = Nothing
'   Set fs = Nothing
'   Set oFolder = Nothing
'   Set rsA = Nothing
End Sub

'Add By Sindy 2017/7/20 ¸ÑªR«H¥ó¤º®e
Sub ReadMailText(ByVal f_myItems As Object, ByVal bolIsReadRecipients As Boolean, _
   Optional ByRef strRecipients_all As String, Optional ByRef strRecipients_1 As String)
   
   strSocSubject = f_myItems.Item(mail_ii).Subject
   Me.TextII17.Text = f_myItems.Item(mail_ii).Subject
   strMailDate = "": strMailTime = "": strSender = ""
   '¥t¦sÀÉ®×®É¤£­n¥H¥D¦®¦sÀÉ,¦]¬°·|¦³ÀÉ®×®æ¦¡¿ù»~°ÝÃD
   '¦]¬°¥D¦®¤º®e¦s¦b¤Ó¦h¥i¯à©Ê·|ÅýÀÉ®×®æ¦¡¿ù»~ªº²Å¸¹
   'TxtIPDept = Replace(f_myItems.Item(ii).Subject, """", "")
   
   '·í±H¥ó¤H¦³­n¨DÅª¨ú¦^±ø®É¨t²Î·|µo«H
   '1.­nOutlook³]©w¤£¦^ÂÐÅª¨ú¦^±ø(¦ý«eÃD¬O«H¥ó¤]¥²¶·³]¬°¤w¶}±Ò)
   '2.­n³]©w¦Û°Ê²M°£¡¨§R°£ªº¶l¥ó¡¨
   '3.­n³]©w¥i¥H¸Ñ¶}ª÷Æ_«H¥ó:°òÂ¦ªº¦w¥þ©Ê¨t²Î§ä¤£¨ì±zªº¼Æ¦ì ID ¦WºÙ(-2146893792)
   'IPM.Note.SMIME ¥[±K
   'f_myItems.Item(mail_ii).UnRead = False '³]¬°¤w¶}±Ò (­Y«H¦³³]Åª¨ú¦^±ø,¨S¶}±Ò®É¦b¡¨§R°£ªº¶l¥ó¡¨,§R°£®É·|¦Û°Ê¦^¶Ç¥¼Åª¨ú¤w§R°£¶l¥óµ¹±H¥óªÌ)
   'Modify By Sindy 2017/11/17
   'Modify By Sindy 2019/11/1 + ¥[±K«H¥ó,f_myItems.Item(mail_ii).Class = 43
   'Modify By Sindy 2020/4/10 + ¦^¦¬«H¥ó=IPM.Outlook.Recall,f_myItems.Item(mail_ii).Class = 43
   'Modify By Sindy 2023/7/12 + or f_myItems.Item(mail_ii).Class = 45 : ·s³qª¾
   If ((InStr(UCase(f_myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or _
        InStr(UCase(f_myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 _
       ) And f_myItems.Item(mail_ii).Class = 43) _
      Or f_myItems.Item(mail_ii).Class = 45 Then
      
   Else
   'If f_myItems.Item(mail_ii).Class = 43 Then '¤@¯ë«H¥ó
   '2017/11/17 END
      f_myItems.Item(mail_ii).UnRead = False '³]¬°¤w¶}±Ò (­Y«H¦³³]Åª¨ú¦^±ø,¨S¶}±Ò®É¦b¡¨§R°£ªº¶l¥ó¡¨,§R°£®É·|¦Û°Ê¦^¶Ç¥¼Åª¨ú¤w§R°£¶l¥óµ¹±H¥óªÌ)
      
      '·|Ä³ÁÜ½Ð
      'f_myItems.Item(ii).MessageClass = IPM.Schedule.Meeting.Request
      'f_myItems.Item(ii).Class = 53
      
      If f_myItems.Item(mail_ii).Class = 46 Then '46.olReport
         strSender = "¥¼¶Ç»¼ªº¥D¦®"
         strMailDate = ""
         strMailTime = ""
      '43.olMail
      Else
         'Modify By Sindy 2020/4/8 Mark
'         If f_myItems.Item(mail_ii).SenderEmailType = "EX" Then
'            strSender = f_myItems.Item(mail_ii).SenderName
'         Else
            If f_myItems.Item(mail_ii).SenderName = f_myItems.Item(mail_ii).senderemailaddress Then '438:ª«¥ó¤£¤ä´©¦¹ÄÝ©Ê©Î¤èªk
               strSender = f_myItems.Item(mail_ii).senderemailaddress
            'Modify By Sindy 2025/2/5 ex:"Tamas Gyomber" <no_reply@yesmywine.com>
            ElseIf f_myItems.Item(mail_ii).SenderName <> "" And f_myItems.Item(mail_ii).senderemailaddress = "" Then
               strSender = f_myItems.Item(mail_ii).SenderName
            '2025/2/5 END
            Else
               strSender = f_myItems.Item(mail_ii).SenderName & " [" & f_myItems.Item(mail_ii).senderemailaddress & "]"
            End If
'         End If
         strMailDate = Format(f_myItems.Item(mail_ii).SentOn, "YYYY/MM/DD") 'ReceivedTime
         strMailTime = Format(f_myItems.Item(mail_ii).SentOn, "HH:MM:SS")
         
         'Add By Sindy 2024/2/7
         If bolIsReadRecipients = True Then
         '2024/2/7 END
            'Modify By Sindy 2025/2/18
            'Call PUB_ReadMailText_CC(f_myItems.Item(mail_ii), strRecipients_all, strRecipients_1) 'Modify By Sindy 2024/7/30
            Call PUB_ReadMailText(f_myItems.Item(mail_ii), strRecipients_all, strRecipients_1) 'Modify By Sindy 2024/7/30
            '2025/2/18 END
'            Dim kk As Integer
'            For kk = f_myItems.Item(mail_ii).Recipients.Count To 1 Step -1
'               strExc(10) = ""
'               If InStr(UCase(f_myItems.Item(mail_ii).Recipients(kk).Name), UCase("@taie.com.tw")) > 0 Then
'                  strExc(10) = f_myItems.Item(mail_ii).Recipients(kk).Name
'               ElseIf InStr(UCase(f_myItems.Item(mail_ii).Recipients(kk).address), UCase("@taie.com.tw")) > 0 Then
'                  strExc(10) = f_myItems.Item(mail_ii).Recipients(kk).address
'               ElseIf InStr(UCase(f_myItems.Item(mail_ii).Recipients(kk).Name), UCase("ipdept")) > 0 Or _
'                      InStr(UCase(f_myItems.Item(mail_ii).Recipients(kk).Name), UCase("±M§Q³B«H½c")) > 0 Or _
'                      InStr(UCase(f_myItems.Item(mail_ii).Recipients(kk).Name), UCase("patent")) > 0 Or _
'                      InStr(UCase(f_myItems.Item(mail_ii).Recipients(kk).Name), UCase("tm")) > 0 Or _
'                      InStr(UCase(f_myItems.Item(mail_ii).Recipients(kk).Name), UCase("account")) > 0 Then
'                  strExc(10) = f_myItems.Item(mail_ii).Recipients(kk).Name
'               ElseIf f_myItems.Item(mail_ii).Recipients(kk).Name <> f_myItems.Item(mail_ii).Recipients(kk).address And _
'                  InStr(f_myItems.Item(mail_ii).Recipients(kk).address, "@") = 0 Then
'                  strRecipients_all = strRecipients_all & "," & f_myItems.Item(mail_ii).Recipients(kk).Name
'                  If f_myItems.Item(mail_ii).Recipients(kk).Type = 1 Then strRecipients_1 = strRecipients_1 & "," & f_myItems.Item(mail_ii).Recipients(kk).Name
'                  strExc(10) = Mid(f_myItems.Item(mail_ii).Recipients(kk).address, InStr(UCase(f_myItems.Item(mail_ii).Recipients(kk).address), UCase("Recipients/cn=")) + Len("Recipients/cn="))
'                  strExc(10) = Replace(strExc(10), """", "")
'                  If InStr(strRecipients_all, strExc(10)) = 0 Then
'                     strRecipients_all = strRecipients_all & "(" & strExc(10) & ")"
'                     If f_myItems.Item(mail_ii).Recipients(kk).Type = 1 Then strRecipients_1 = strRecipients_1 & "(" & strExc(10) & ")"
'                  End If
'                  strExc(10) = ""
'               End If
'               If strExc(10) <> "" Then
'                  strRecipients_all = strRecipients_all & "," & strExc(10)
'                  If f_myItems.Item(mail_ii).Recipients(kk).Type = 1 Then strRecipients_1 = strRecipients_1 & "," & strExc(10)
'               End If
'            Next kk
'            If strRecipients_all <> "" Then strRecipients_all = Mid(strRecipients_all, 2)
'            If strRecipients_1 <> "" Then strRecipients_1 = Mid(strRecipients_1, 2)
         End If
      End If
   End If
   
   If f_myItems.Item(mail_ii).Class <> 43 Then
      WLog strSocSubject & vbCrLf & "==> Class : " & f_myItems.Item(mail_ii).Class & " MessageClass : " & f_myItems.Item(mail_ii).MessageClass & vbCrLf
   End If
End Sub

'Add By Sindy 2023/9/13 ¸ÑªR«H¥ó¤º®e
Sub ReadMailText_File(ByVal f_myItems As Object)
   strSocSubject = f_myItems.Subject
   TextBox3 = f_myItems.Subject 'Add By Sindy 2023/12/26
   Me.TextII17.Text = f_myItems.Subject
   strMailDate = "": strMailTime = "": strSender = ""
   '¥t¦sÀÉ®×®É¤£­n¥H¥D¦®¦sÀÉ,¦]¬°·|¦³ÀÉ®×®æ¦¡¿ù»~°ÝÃD
   '¦]¬°¥D¦®¤º®e¦s¦b¤Ó¦h¥i¯à©Ê·|ÅýÀÉ®×®æ¦¡¿ù»~ªº²Å¸¹
   'TxtIPDept = Replace(f_myItems.Item(ii).Subject, """", "")
   
   '·í±H¥ó¤H¦³­n¨DÅª¨ú¦^±ø®É¨t²Î·|µo«H
   '1.­nOutlook³]©w¤£¦^ÂÐÅª¨ú¦^±ø(¦ý«eÃD¬O«H¥ó¤]¥²¶·³]¬°¤w¶}±Ò)
   '2.­n³]©w¦Û°Ê²M°£¡¨§R°£ªº¶l¥ó¡¨
   '3.­n³]©w¥i¥H¸Ñ¶}ª÷Æ_«H¥ó:°òÂ¦ªº¦w¥þ©Ê¨t²Î§ä¤£¨ì±zªº¼Æ¦ì ID ¦WºÙ(-2146893792)
   'IPM.Note.SMIME ¥[±K
   'f_myItems.UnRead = False '³]¬°¤w¶}±Ò (­Y«H¦³³]Åª¨ú¦^±ø,¨S¶}±Ò®É¦b¡¨§R°£ªº¶l¥ó¡¨,§R°£®É·|¦Û°Ê¦^¶Ç¥¼Åª¨ú¤w§R°£¶l¥óµ¹±H¥óªÌ)
   'Modify By Sindy 2017/11/17
   'Modify By Sindy 2019/11/1 + ¥[±K«H¥ó,f_myItems.Class = 43
   'Modify By Sindy 2020/4/10 + ¦^¦¬«H¥ó=IPM.Outlook.Recall,f_myItems.Class = 43
   'Modify By Sindy 2023/7/12 + or f_myItems.Class = 45 : ·s³qª¾
   If ((InStr(UCase(f_myItems.MessageClass), UCase("IPM.Note.SMIME")) > 0 Or _
        InStr(UCase(f_myItems.MessageClass), UCase("IPM.Outlook.Recall")) > 0 _
       ) And f_myItems.Class = 43) _
      Or f_myItems.Class = 45 Then
      
   Else
   'If f_myItems.Class = 43 Then '¤@¯ë«H¥ó
   '2017/11/17 END
      f_myItems.UnRead = False '³]¬°¤w¶}±Ò (­Y«H¦³³]Åª¨ú¦^±ø,¨S¶}±Ò®É¦b¡¨§R°£ªº¶l¥ó¡¨,§R°£®É·|¦Û°Ê¦^¶Ç¥¼Åª¨ú¤w§R°£¶l¥óµ¹±H¥óªÌ)
      
      '·|Ä³ÁÜ½Ð
      'f_myItems.Item(ii).MessageClass = IPM.Schedule.Meeting.Request
      'f_myItems.Item(ii).Class = 53
      
      If f_myItems.Class = 46 Then '46.olReport
         strSender = "¥¼¶Ç»¼ªº¥D¦®"
         strMailDate = ""
         strMailTime = ""
      '43.olMail
      Else
         'Modify By Sindy 2020/4/8 Mark
'         If f_myItems.SenderEmailType = "EX" Then
'            strSender = f_myItems.SenderName
'         Else
            If f_myItems.SenderName = f_myItems.senderemailaddress Then '438:ª«¥ó¤£¤ä´©¦¹ÄÝ©Ê©Î¤èªk
               strSender = f_myItems.senderemailaddress
            Else
               strSender = f_myItems.SenderName & " [" & f_myItems.senderemailaddress & "]"
            End If
'         End If
         strMailDate = Format(f_myItems.SentOn, "YYYY/MM/DD") 'ReceivedTime
         strMailTime = Format(f_myItems.SentOn, "HH:MM:SS")
      End If
   End If
   
   If f_myItems.Class <> 43 Then
      'WLog strSocSubject & vbCrLf & "==> Class : " & f_myItems.Class & " MessageClass : " & f_myItems.MessageClass & vbCrLf
      WLog TextBox3 & vbCrLf & "==> Class : " & f_myItems.Class & " MessageClass : " & f_myItems.MessageClass & vbCrLf
   End If
End Sub

Function WLog(oStrLog As String)
Dim ffa As Integer
Dim strNow As String
   
   If Dir(App.path & "\TaOutLookLog\", vbDirectory) = "" Then
      MkDir App.path & "\TaOutLookLog\"
   End If
   
   strNow = Trim(Now)
   '¼g¦bµe­±¤W
   'lstHistory.AddItem strNow & "  -->  " & oStrLog, 0
   '¼g¦b¤å¦rÀÉ
   ffa = FreeFile
   Open App.path & "\TaOutLookLog\" & pub_DbTerminalName & "TaOutLook.log" For Append As ffa
   Print #ffa, strNow & "  ==>  " & oStrLog
   Close ffa
End Function

Public Function WLog_Day(oStrLog As String, MailName As String, _
   Optional bolShowTime As Boolean = True, _
   Optional m_strFileName As String = "") As Boolean
Dim ffa As Integer
Dim strNow As String
Dim ii As Integer
Dim strListTxt As String
Dim strFileName As String
   
   If m_strFileName = "" Then
      strFileName = App.path & "\TaOutLookLog\"
   Else
      strFileName = m_strFileName
   End If
   
   If Dir(strFileName, vbDirectory) = "" Then
      MkDir strFileName
   End If
   
   WLog_Day = False
   If InStr(MailName & oStrLog, "strFileName : ") > 0 Then
      strListTxt = Replace(Trim(Left(MailName & oStrLog, InStr(MailName & oStrLog, "strFileName : ") - 1)), vbCrLf, "")
      For ii = 0 To ListErrTxt.ListCount - 1
         If ListErrTxt.List(ii) = strListTxt Then '¦¹¿ù»~°T®§¤w¦s¦b,¤£¶·¦A¼g¤J
            Exit Function
         End If
      Next ii
   End If
   strNow = Trim(Now)
   '¼g¦bµe­±¤W
   'lstHistory.AddItem strNow & "  -->  " & oStrLog, 0
   '¼g¦b¤å¦rÀÉ
   ffa = FreeFile
   Open strFileName & pub_DbTerminalName & "TaOutLook_" & MailName & strSrvDate(2) & ".log" For Append As ffa
   If bolShowTime = True Then
      Print #ffa, strNow & "  ==>  " & oStrLog
   Else
      Print #ffa, oStrLog
   End If
   Close ffa
   
   WLog_Day = True
   If InStr(MailName & oStrLog, "strFileName : ") > 0 Then
      ListErrTxt.AddItem strListTxt
   End If
End Function

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   '                        0       1           2           3           4           5           6           7           8
   arrGridHeadText = Array("«H½c", "±µ¦¬¤é´Á", "°_©l®É¶¡", "ºI¤î®É¶¡", "·s¼W¤H­û", "±µ¦¬µ§¼Æ", "¥[±Kµ§¼Æ", "­Ó®×µ§¼Æ", "°õ¦æª¬ªp")
   arrGridHeadWidth = Array(1400, 800, 800, 800, 800, 800, 800, 800, 800)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub cmdQuery_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
   
   strSql = ""
   If Combo1.Text <> "" Then
      strSql = strSql & " and MRL01='" & Left(Combo1.Text, 2) & "'"
   End If
   If txtMRL02.Text <> "" Then
      strSql = strSql & " and MRL02='" & DBDATE(txtMRL02.Text) & "'"
   End If
   
   GRD1.Clear
   SetGrd
   
   Screen.MousePointer = vbHourglass
   strSql = "Select " & MRL01CName & " «H½c,sqldatet(MRL02) ±µ¦¬¤é´Á,sqltime(MRL03) °_©l®É¶¡,sqltime(MRL04) ºI¤î®É¶¡,st02 ·s¼W¤H­û,MRL06 ±µ¦¬µ§¼Æ,MRL07 ¥[±Kµ§¼Æ,MRL08 ­Ó®×µ§¼Æ," & MRL09CName & " °õ¦æª¬ªp" & _
            " From MailReceiveLog,Staff" & _
            " Where MRL05=ST01(+)" & strSql & _
            " Order By MRL02||substr('000000'||MRL03,-6) desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
   Else
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   
   '­Y¦³¸ê®Æ´å¼Ð°±¦b²Ä¤@µ§
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   dblPrevRow = GRD1.row
   If rsTmp.RecordCount > 0 Then
      'GRD1.Text = "V"
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub grd1_SelChange()
Dim i As Integer
   
   GRD1.Visible = False
   If GRD1.MouseRow <> 0 Then
      '¤W¤@µ§¸ê®Æ¦C²M°£¤Ï¥Õ
      If dblPrevRow > 0 Then
         GRD1.col = 0
         GRD1.row = dblPrevRow
         'GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      End If
      '¥Ø«e¸ê®Æ¦C¤Ï¥Õ
      GRD1.col = 0
      GRD1.row = GRD1.MouseRow
      dblPrevRow = GRD1.row
   '   If grd1.Text = "V" Then
   '      grd1.Text = ""
   '      For i = 0 To grd1.Cols - 1
   '         grd1.col = i
   '         grd1.CellBackColor = QBColor(15)
   '      Next i
   '   Else
         If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
            'GRD1.Text = "V"
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = &HFFC0C0
            Next i
         End If
   '   End If
   End If
   GRD1.Visible = True
End Sub

'©I¥s·s¶l¥ó
Private Sub OpenNeweMail(strTo As String, strSubject As String, _
                         strContext As String, Optional strAttach As String)
Dim objOutLook As Object
Dim objMail As Object
Dim ArrStr As Variant
Dim jj As Integer
   
'   PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , , , , , True, False, , , False, , , False
'   DoEvents
'   Exit Sub
   
   '©I¥s·s¶l¥ó¡G
   Set objOutLook = CreateObject("Outlook.Application")
   'Set objMail = objOutLook.CreateItem(0) '·s¶l¥ó
'   If strAttach <> "" Then
'      Set objMail = objOutLook.CreateItemFromTemplate(strAttach) '­ì©l«H
'   Else
      Set objMail = objOutLook.CreateItem(0)
'   End If
   
   'objMail.PrintOut '¦C¦L¶l¥ó¤Îªþ¥ó,ªþ¥ó¥»¨­¦b¹q¸£¤¤«ö·Æ¹«¥kÁä¬O¥i¥H¦C¦Lªº
'   'ªþ¥ó
'   For jj = objMail.Attachments.Count To 1 Step -1 '­Ó¼Æ
'      objMail.Attachments.Item(jj).SaveAsFile "c:\" & objMail.Attachments.Item(jj).DisplayName '¥t¦sÀÉ®×
'   Next jj
'   '²¾°£­ì«Hªº¦¬¥ó¤H¤Î°Æ¥»;±K¥ó°Æ¥»¤£·|¯d¦bmsg¤¤
'   For jj = objMail.Recipients.Count To 1 Step -1
'      objMail.Recipients.Remove jj
'   Next jj
   
   '±H¥óªÌ (Microsoft Outlook 15.0 Object Library¤~¯à³]©w)
   'objMail.Sender.address = "qpgmr@taie.com.tw"
   'objMail.Sender = "qpgmr"
   '°Æ¥».cc
   
   '¦¬¥óªÌ.To
'   objMail.To = strTo
   ArrStr = Split(strTo, ";")
   For jj = 0 To UBound(ArrStr)
      objMail.Recipients.add ArrStr(jj)
   Next jj

   '°Æ¥»
   'objMail.To = "97038"
   '±K¥ó°Æ¥».BCC
   
   '¥D¦®.Subject
   objMail.Subject = strSubject
   
   '¥[ªþ¥ó
   If strAttach <> "" Then
      ArrStr = Split(strAttach, ";")
      For jj = 0 To UBound(ArrStr)
         objMail.Attachments.add ArrStr(jj)
      Next jj
   End If
   
   '¤º¤å.Body
   objMail.Body = strContext
   
   'objMail.Display
   objMail.Send
   
   Set objMail = Nothing
   Set objOutLook = Nothing
End Sub

Private Sub mnuDisplay_Click()
Me.WindowState = "0"
Me.Visible = True
End Sub

Private Sub mnuQuit_Click()
   Call cmdExit_Click
End Sub

'°ê¥~³¡(±H¥ó³Æ¥÷)¶l¥óÂk¨÷©v°Ï
'strTo:Âà±H¤H­û
'strII05:¤ÀÃþ
'¦^¶Ç:¬O§_¦¨¥\
Private Function IPDeptBackupMail(ByVal strSubject As String, _
   ByVal strFullFileName As String, ByVal strFileName As String, _
   Optional ByRef strErrText As String, Optional ByRef intCaseOK As Integer, _
   Optional ByVal strRecipients As String) As Boolean
Dim strText As String
Dim strUpdTime As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strCP09 As String, strCP10 As String, strII03_2 As String, stReName As String
Dim fs, f
Dim bolSaveEFile As Boolean
Dim bolConnect As Boolean
Dim intCaseKind As Integer
Dim strEmp As String, strDirector As String
Dim strII18 As String, strOurII18 As String, strYourII18 As String
Dim rsA As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim YourRefCase As String, OurRefCase As String
Dim strTemp1 As String, strTemp2 As String, strTemp3 As String, StrTemp4 As String
Dim strProc As String, intStar As Integer, intEnd As Integer, strTextSubject As String 'Add By Sindy 2018/5/16

On Error GoTo ErrHand
   
   IPDeptBackupMail = False
   strErrText = ""
   Screen.MousePointer = vbHourglass
   Set fs = CreateObject("Scripting.FileSystemObject")
   
   strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
   strCP09 = "": strCP10 = "": strII18 = ""
   YourRefCase = "": OurRefCase = ""
   
   '¸ÑªR¥»©Ò®×¸¹¨Ï¥Î
   strText = strSubject
   'Modify By Sindy 2022/8/5 ¦P¥~¨Ó¶l¥óªº³W«h°µ¥»©Ò®×¸¹
   'strII05 = PUB_IPDept_ToSortOut(strText, strII11, strII06, strCP01, strCP02, strCP03, strCP04, strII18)
   Call PUB_IPDept_ToSortOut(strText, "", "", strCP01, strCP02, strCP03, strCP04, strII18, True)
   '2022/8/5 END
'   'Modify By Sindy 2017/7/28
'   'Modify By Sindy 2018/4/17 ex: AP/lc PRC Patent Application No. 201680016298.2;Your Ref: P2000 ;Our Ref: P-118009
'   'Call PUB_IPDept_ToSortOut(strText, "", "", strCP01, strCP02, strCP03, strCP04, strII18)
'   '¥ý¸ÑªR¦³µL¥»©Ò®×¸¹
'   If PUB_IPDeptGetCaseNo(strText, "OURREF", strCP01, strCP02, strCP03, strCP04, , , , strII18) = False Then
'      If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, , , , strII18) = False Then
'      End If
'   'Modify By Sindy 2021/6/28 ­Y¬O¥Î¥Ó½Ð®×¸¹,±M§Q¸¹,©¼©Ò¸¹µ¥§ì¨ì¸ê®Æ, ¦A¸ÑªR¤@¦¸®×¸¹
'   'ex: WC/jc/bc - Taiwan Patent Application No. 106114285; Your Ref: ADVSIL-13-TW / MM; Our Ref: FCP-056692 [REPdn.205]
'   ElseIf strII18 <> "OURREF" Then
'      strTemp1 = strCP01: strTemp2 = strCP02: strTemp3 = strCP03: StrTemp4 = strCP04: strOurII18 = strII18
'      If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, , , , strII18) = False Then
'      End If
'      'YOURREF ¨S§ä¨ì ©Î §ä¨ì¤£¬O­Ó®×, ´N¥ÎOURREF§ä¨ìªº¸ê®Æ,°µ«áÄò¤ñ¹ï
'      If strII18 = "" Or strII18 <> "YOURREF" Then
'         strCP01 = strTemp1: strCP02 = strTemp2: strCP03 = strTemp3: strCP04 = StrTemp4: strII18 = strOurII18
'      End If
'   '2021/6/28 END
'   End If
'
'   strTemp1 = "": strTemp2 = "": strTemp3 = "": StrTemp4 = "": strOurII18 = ""
'   'Modify By Sindy 2021/9/29 + , IIf(InStr("¥Ó½Ð®×¸¹¡B±M§Q¸¹¼Æ¡B©¼©Ò®×¸¹", strII18) = 0 And strII18 <> "", False, True):¤w¦³§ì¨ì¥»©Ò®×¸¹
'   If PUB_IPDeptGetCaseNo(strText, "OURREF", strTemp1, strTemp2, strTemp3, StrTemp4, , , , strOurII18, IIf(InStr("¥Ó½Ð®×¸¹¡B±M§Q¸¹¼Æ¡B©¼©Ò®×¸¹", strII18) = 0 And strII18 <> "", False, True)) = True Then
''      'Modify By Sindy 2021/6/28 ­Y®×¸¹¤w§ì¨ì,´N¤£­n¦A¥Î¥Ó½Ð®×¸¹,±M§Q¸¹,©¼©Ò¸¹µ¥¸ê®Æ
''      If Not (strII18 <> "OURREF" And strCP01 & strCP02 <> "") Then
''      '2021/6/28 END
''         OurRefCase = strTemp1 & "-" & strTemp2 & "-" & strTemp3 & "-" & strTemp4
''      End If
'   End If
'
'   strTemp1 = "": strTemp2 = "": strTemp3 = "": StrTemp4 = "": strYourII18 = ""
'   If PUB_IPDeptGetCaseNo(strText, "YOURREF", strTemp1, strTemp2, strTemp3, StrTemp4, , , , strYourII18, IIf(InStr("¥Ó½Ð®×¸¹¡B±M§Q¸¹¼Æ¡B©¼©Ò®×¸¹", strII18) = 0 And strII18 <> "", False, True)) = True Then
''      'Modify By Sindy 2021/6/28 ­Y®×¸¹¤w§ì¨ì,´N¤£­n¦A¥Î¥Ó½Ð®×¸¹,±M§Q¸¹,©¼©Ò¸¹µ¥¸ê®Æ
''      If Not (strII18 <> "YOURREF" And strCP01 & strCP02 <> "") Then
''      '2021/6/28 END
''         YourRefCase = strTemp1 & "-" & strTemp2 & "-" & strTemp3 & "-" & strTemp4
''      End If
'   End If
'
'   If YourRefCase <> "" And OurRefCase <> "" And YourRefCase <> OurRefCase Then
'      strTemp1 = SystemNumber(YourRefCase, 1)
'      strTemp2 = SystemNumber(YourRefCase, 2)
'      strTemp3 = SystemNumber(YourRefCase, 3)
'      StrTemp4 = SystemNumber(YourRefCase, 4)
'      'Âk­Ó®×®É­Y¸Ó®×¥ó¶i«×©Ó¿ì¤H,·~°È­û³£¨S¦³Fxx¤H­û®É¤£Âk
'      strExc(0) = "select count(*) from caseprogress,staff s1,staff s2" & _
'                  " where cp01='" & strTemp1 & "' and cp02='" & strTemp2 & "' and cp03='" & strTemp3 & "' and cp04='" & StrTemp4 & "'" & _
'                  " and cp13=s1.st01(+) and substr(s1.st03,1,1)='F'" & _
'                  " and cp14=s2.st01(+) and substr(s2.st03,1,1)='F'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If RsTemp.Fields(0) = 0 Then
'            YourRefCase = ""
'         End If
'      End If
'      strTemp1 = SystemNumber(OurRefCase, 1)
'      strTemp2 = SystemNumber(OurRefCase, 2)
'      strTemp3 = SystemNumber(OurRefCase, 3)
'      StrTemp4 = SystemNumber(OurRefCase, 4)
'      strExc(0) = "select count(*) from caseprogress,staff s1,staff s2" & _
'                  " where cp01='" & strTemp1 & "' and cp02='" & strTemp2 & "' and cp03='" & strTemp3 & "' and cp04='" & StrTemp4 & "'" & _
'                  " and cp13=s1.st01(+) and substr(s1.st03,1,1)='F'" & _
'                  " and cp14=s2.st01(+) and substr(s2.st03,1,1)='F'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If RsTemp.Fields(0) = 0 Then
'            OurRefCase = ""
'         End If
'      End If
'      If YourRefCase <> "" And OurRefCase <> "" Then '2²Õ®×¸¹³£¦³°ê¥~³¡¤H­û
'         'strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
'         'Modify By Sindy 2018/6/28 ¦P¥~¨Ó«H¥ó³W«h
'         'Your Ref¤ÎOur Ref¦P®É¦s¦b®É,­Y¦³FCP,FCT,CFT,CFP,FG¦r¼Ë«hÀu¥ý¦Ò¼{,§_«h¥þ³¡Âk¨ä¥L
'         If SystemNumber(YourRefCase, 1) <> SystemNumber(OurRefCase, 1) Then
'            strExc(0) = "'" & SystemNumber(YourRefCase, 1) & "'"
'            strExc(1) = "'" & SystemNumber(OurRefCase, 1) & "'"
'            If InStr("'FCP','FCT','CFT','CFP','FG'", strExc(0)) > 0 Then
'               strCP01 = SystemNumber(YourRefCase, 1)
'               strCP02 = SystemNumber(YourRefCase, 2)
'               strCP03 = SystemNumber(YourRefCase, 3)
'               strCP04 = SystemNumber(YourRefCase, 4)
'               strII18 = strYourII18
'            ElseIf InStr("'FCP','FCT','CFT','CFP','FG'", strExc(1)) > 0 Then
'               strCP01 = SystemNumber(OurRefCase, 1)
'               strCP02 = SystemNumber(OurRefCase, 2)
'               strCP03 = SystemNumber(OurRefCase, 3)
'               strCP04 = SystemNumber(OurRefCase, 4)
'               strII18 = strOurII18
'            Else
'               strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = "": strII18 = ""
'            End If
'         End If
'         '2018/6/28 END
'      ElseIf YourRefCase <> "" Then
'         strCP01 = SystemNumber(YourRefCase, 1)
'         strCP02 = SystemNumber(YourRefCase, 2)
'         strCP03 = SystemNumber(YourRefCase, 3)
'         strCP04 = SystemNumber(YourRefCase, 4)
'         strII18 = strYourII18
'      ElseIf OurRefCase <> "" Then
'         strCP01 = SystemNumber(OurRefCase, 1)
'         strCP02 = SystemNumber(OurRefCase, 2)
'         strCP03 = SystemNumber(OurRefCase, 3)
'         strCP04 = SystemNumber(OurRefCase, 4)
'         strII18 = strOurII18
'      End If
'   End If
'   '2018/4/17 END
   
   If strCP01 <> "" And strCP02 <> "" Then
'               '¸Ó®×¸¹³Ì¤j¦¬¤å¤é³Ì¤pCreate¤é´Á®É¶¡ªºÁ`¦¬¤å¸¹
'               strExc(0) = "select cp09 from caseprogress" & _
'                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
'                           " and cp05=(select max(cp05) from caseprogress" & _
'                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "')" & _
'                           " order by cp66 desc,cp67 asc"
      '¸Ó®×¸¹A,B,CÃþ³Ì¤j¦¬¤å¤é³Ì¤jÁ`¦¬¤å¸¹
      'Modify By Sindy 2017/7/18 ¤£­ç°£DÃþ¶i«× : and cp09<'D'
      'Modify By Sindy 2025/5/6 ­ç°£FCPªº1920=«È¤á´£¨Ñ¤å¥ó,¦]¬°¦¹¶i«×µo¤å«á¬O·|³Q§R°£ªº
      strExc(0) = "select cp09 from caseprogress" & _
                  " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                  " and cp05=(select max(cp05) from caseprogress" & _
                  " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "' and not(substr(cp09,1,1)='D' and cp01='FCP' and cp10='1920'))" & _
                  " and not(substr(cp09,1,1)='D' and cp01='FCP' and cp10='1920')" & _
                  " order by SQLDatet2(CP05) DESC, CP66 DESC, CP67 DESC, CP09 DESC"
                  'Modify By Sindy 2018/6/27 order by cp66 desc,cp67 desc
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strCP09 = rsA.Fields(0)
         strExc(0) = "select cp10 from caseprogress where cp09='" & strCP09 & "'"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strCP10 = rsA.Fields("cp10")
         End If
      End If
      
      cnnConnection.BeginTrans
      bolConnect = True
      strUpdTime = Right("000000" & ServerTime, 6)
      
      '¦s¨÷©v°Ï
      strII03_2 = "": strProc = "": intStar = 0
      'Modify By Sindy 2018/10/5
'      '¸ÑªR¥D¦®¨Ï¥Î
'      strTextSubject = strSubject
'      strTextSubject = Replace(strTextSubject, "¡D", ".")
'      strTextSubject = Replace(strTextSubject, "..", ".")
'      strTextSubject = Replace(strTextSubject, "...", ".")
'      If UCase(strRecipients) = UCase("backup") Then '¦¬¥óªÌ¬°backup;¥Nªí«H¥ó¯Â¬°Âk¨÷©v°Ï
'         If InStr(UCase(strTextSubject), UCase("[¯È¥»±H¥X]")) > 0 Then '¯È¥»±H¥X
'            strII03_2 = Replace(strFileName, ".msg", ".paper.msg")
'         ElseIf InStr(UCase(strTextSubject), UCase("[¥­¥x¤U¸ü]")) > 0 Then '¥­¥x¤U¸ü
'            strII03_2 = Replace(strFileName, ".msg", ".dnl.msg")
'         ElseIf InStr(UCase(strTextSubject), UCase("[¥­¥x¤W¶Ç]")) > 0 Then '¥­¥x¤W¶Ç
'            strII03_2 = Replace(strFileName, ".msg", ".upl.msg")
'         End If
'      End If
'      'Add By Sindy 2018/5/16 Âk¤J¥¿½Tªº®×¥ó©Ê½è,°ÆÀÉ¦W
      'Modify By Sindy 2018/7/5 §ï¦¨¨ç¼Æ
      Call PUB_IPDept_ComparisonCP(strSubject, strFileName, strCP01, strCP02, strCP03, strCP04, strII03_2, strCP09, strCP10)
      If strII03_2 = "" Then
         strII03_2 = Replace(strFileName, ".msg", ".tx.msg")
      End If
      '2018/5/16 END
      'Modify By Sindy 2020/1/31 ¥»©Ò®×¸¹¬y¤ô¸¹­n¦s¨¬½X
'      stReName = Trim(strCP01) & Val(Trim(strCP02)) & _
'                  IIf(Val(Trim(strCP03)) = 0 And Val(Trim(strCP04)) = 0, "", "-" & strCP03) & _
'                  IIf(Val(Trim(strCP04)) = 0, "", "-" & Format(strCP04, "00")) & "." & strCP10 & "." & _
'                  strII03_2
      'Modify By Sindy 2020/2/19 ¹q¤lÀÉ¦W,¥»©Ò®×¸¹¨Ï¥Î¨ç¼Æ PUB_CaseNo2FileName
'      stReName = Trim(strCP01) & Trim(strCP02) & _
'                  IIf(Val(Trim(strCP03)) = 0 And Val(Trim(strCP04)) = 0, "", "-" & strCP03) & _
'                  IIf(Val(Trim(strCP04)) = 0, "", "-" & Format(strCP04, "00")) & "." & strCP10 & "." & _
'                  strII03_2
      stReName = PUB_CaseNo2FileName(strCP01, strCP02, strCP03, strCP04) & _
                  "." & strCP10 & "." & strII03_2
      '+ save cpp04
      'Modify By Sindy 2017/8/30 +  & IIf(strII18 <> "", " [" & strII18 & "]", "") ¦s¤ñ¹ï¨ìªºÃöÁä¦r
'      Text2 = ChgSQL(strSubject) & IIf(strII18 <> "", " [" & strII18 & "]", "") '­n¥Î¤å¦r®Ø¦s©ñ¡A¦]¤~¯à§âunicode¥h±¼
      
      Set f = fs.GetFile(strFullFileName)
      '¥u¦³¥~±MÂk¨÷©v°Ï
      '¥Ñ¨t²Î¥N¸¹¡A¨ú±o1¬°±M§Q¡A2¬°°Ó¼Ð¡A3¬°ÅU°Ý¸u¥ô¡A4¬°ªk°È
'      If ClsPDGetSystemKind(strCP01, intCaseKind) = True Then
'         If intCaseKind = ±M§Q Then
            WLog_Day "==>" & strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 & " : " & strCP09 & "(" & strCP10 & ") ¤ÀÃþ°O¿ý=[" & ChgSQL(strII18) & "] " & strFullFileName & " ==> " & stReName, °ê¥~³¡±H¥ó«H½c
            bolSaveEFile = SaveAttFile_PDF(strCP09, strFullFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True, "F", "Y", , , , Me.TextII17.Text, strErrText, False)
            If bolSaveEFile = False Then
               WLog_Day "SaveAttFile_PDF ¥¢±Ñ: " & strErrText, °ê¥~³¡±H¥ó«H½c
               'Add By Sindy 2020/4/6
               If InStr(strErrText, strSubject) = 0 Then
                  strErrText = strErrText & vbCrLf & _
                        strSubject & vbCrLf & _
                        "==>¦¬¨ì¤é´Á:" & strMailDate & " " & strMailTime & " ±H¥óªÌ:" & strSender & vbCrLf & _
                        "==>" & strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 & " : " & strCP09 & "(" & strCP10 & ")" & strFullFileName & "==>" & stReName & vbCrLf
               End If
               '2020/4/6 END
               PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & strCP01 & strCP02 & strCP03 & strCP04 & "-" & strCP09 & "­Ó®×¦s¨÷©v°Ï¥¢±Ñ¡A½Ð¬d¬Ý¡I", strErrText, , , , , , , , , , , False, , , False, , , False
               DoEvents
               'Add By Sindy 2017/7/31 °O¿ýLog
               strEmp = "": strDirector = ""
               'Modify By Sindy 2019/9/3 ¨Ì¥D¦®¸ÑªR±H¥ó¤H­û¬O½Ö
               'If BySenderToStaff(strSender, strEmp, strDirector) = True Then
               Call BySubjectToStaff(strSubject, strSender, strEmp, strDirector)
               If strEmp <> "" Then
               '2019/9/3 END
                  strSql = "insert into R100101(R005002,R005004,R005005,R005003,R005007,R005006,R005008,ID)" & _
                           " values('" & strMailDate & "','" & strMailTime & "','¨t²ÎLog°O¿ý,¤£¥i§R°£','" & ChgSQL(strSender) & "','[Âk¨÷¥¢±Ñ] " & ChgSQL(strSubject) & "'," & _
                           "'" & strEmp & "','" & strDirector & "','" & strUserNum & "')"
                  cnnConnection.Execute strSql
                  WLog_Day strSql, °ê¥~³¡±H¥ó«H½c
               End If
               '2017/7/31 END
               '§R°£PCºÝÀÉ®×
               Call fs.DeleteFile(strFullFileName)
               DoEvents
               WLog_Day "[§R°£] GoTo ErrHand" & strFullFileName, °ê¥~³¡±H¥ó«H½c
               GoTo ErrHand '¥¢±Ñµ²§ô
            Else
               intCaseOK = intCaseOK + 1 '°O¿ý­Ó®×µ§¼Æ
            End If
'         End If
'      End If
      '§R°£PCºÝÀÉ®×
      'Kill §R¤£±¼ "C:\IPdept\¡iÂàª¾¡j(1) ¸gÀÙ³¡´¼¼z°]²£§½¨Ó¨ç¡A¦Û105¦~4¤ë1¤é°_´£¥Xµo©ú±M§Q¥[³t¼f¬d¡B±M§Q¼f¬d°ª³t¤½¸ô»P¤ä´©§Q¥Î±M§Q¼f¬d°ª³t¤½¸ô¤§±M§Q¥Ó½Ð®×©|¥¼¤½¶}ªÌ¡A¤£¥²¦A¥Ó½Ð´£¦­¤½¶}¡F(2) ¸gÀÙ³¡´¼¼z°]²£§½¨Ó¨ç¡A¤½§i­×¥¿¡uµo©ú±M§Q¥[³t¼f¬d¥Ó½Ð®Ñ¤Î¨ä¥Ó½Ð¶·ª¾¡v¡B¡uµo©ú±M§QPPH¼f¬d¥Ó½Ð®Ñ¤Î¨ä¥Ó½Ð¶·ª¾¡v»P¡uµo©ú±M§QTW-SUPA¼f¬d¥Ó½Ð®Ñ¡v.msg"
      'Kill txtPathIPDept.Text & "\" & oFile.Name
      Call fs.DeleteFile(strFullFileName)
      DoEvents
      WLog_Day "[¦s¨÷¦¨¥\, §R°£]" & strFullFileName, °ê¥~³¡±H¥ó«H½c
      
      cnnConnection.CommitTrans
      bolConnect = False
      
   Else
      WLog_Day "§ä¤£¨ì¹ïÀ³®×¥ó", °ê¥~³¡±H¥ó«H½c
'      WLog_Day "§ä¤£¨ì¹ïÀ³®×¥ó : " & vbCrLf & strSubject & vbCrLf & _
'               "==>¦¬¨ì¤é´Á:" & strMailDate & " " & strMailTime & " ±H¥óªÌ:" & strSender & vbCrLf, °ê¥~³¡±H¥ó«H½c
      'Add By Sindy 2017/7/31 °O¿ýLog
      strEmp = "": strDirector = ""
      'Modify By Sindy 2019/9/3 ¨Ì¥D¦®¸ÑªR±H¥ó¤H­û¬O½Ö
      'If BySenderToStaff(strSender, strEmp, strDirector) = True Then
      Call BySubjectToStaff(strSubject, strSender, strEmp, strDirector)
      If strEmp <> "" Then
      '2019/9/3 END
         strSql = "insert into R100101(R005002,R005004,R005005,R005003,R005007,R005006,R005008,ID)" & _
                  " values('" & strMailDate & "','" & strMailTime & "','¨t²ÎLog°O¿ý,¤£¥i§R°£','" & ChgSQL(strSender) & "','" & ChgSQL(strSubject) & "'," & _
                  "'" & strEmp & "','" & strDirector & "','" & strUserNum & "')"
         cnnConnection.Execute strSql
         WLog_Day strSql, °ê¥~³¡±H¥ó«H½c
      End If
      '2017/7/31 END
      '§R°£PCºÝÀÉ®×
      Call fs.DeleteFile(strFullFileName)
      DoEvents
      WLog_Day "[µL¦s¨÷, §R°£]" & strFullFileName, °ê¥~³¡±H¥ó«H½c
   End If
   IPDeptBackupMail = True
   Screen.MousePointer = vbDefault
   Set f = Nothing
   Set fs = Nothing
   Set rsA = Nothing
   Set RsQ = Nothing
   
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConnect = True Then cnnConnection.RollbackTrans
   strErrText = strErrText & "±H¥ó³Æ¥÷¶×¤J¥¢±Ñ¡I" & vbCrLf & Err.Number & vbCrLf & Err.Description
   WLog_Day "[¥¢±Ñ IPDeptBackupMail-ErrHand]" & strErrText, °ê¥~³¡±H¥ó«H½c
   Set f = Nothing
   Set fs = Nothing
   Set rsA = Nothing
   Set RsQ = Nothing
End Function

Private Sub TmrPatent_Timer()
   'Modify By Sindy 2024/5/13
   'Call importPatentMail
   Call ChkExecutionTimer(Left(Patent¦¬¥ó§X, 2))
   '2024/5/13 END
End Sub

''±M§Q³B¦¬¥ó«H½c³B²zµ{§Ç
'Private Function importPatentMail() As Boolean
'Dim kk As Integer, jj As Integer
'Dim strTo As String, strCC As String, strTempCC As String
'Dim oFileSys As New FileSystemObject, oFolder As Object
'Dim strKind As String
'Dim myForward As Object
'Dim myNewEmail As Object 'Âà±H«H¥ó
'Dim ArrStr As Variant, ArrStrkk As Variant
'Dim strCaseNo As String
'Dim strIPMNoteSMIME As String '¥[±K¥D¦®
'Dim bolReStarPatent As Boolean
'Dim strMRL01 As String, strMRL02 As String, strMRL03 As String, strMRL04 As String, strMRL05 As String
'Dim rsA As New ADODB.Recordset
'Dim strPTo As String 'Add By Sindy 2018/2/8
'Dim strErrNumber As String 'Add By Sindy 2019/10/14
'Dim strErrCode As String, strErrDesc As String 'Add By Sindy 2020/4/15
'Dim fs 'Add By Sindy 2022/2/22
'Dim strRecipients_1 As String, strRecipients_all As String '§ì¦¬¥óªÌ¸ê®Æ
''Add By Sindy 2023/6/26
'Dim olApp As Object
'Dim myNamespace As Object
'Dim myFolder As Object
'Dim myItems As Object
''2023/6/26 END
'Dim strMailTime_Recv As String 'Add By Sindy 2023/7/13
'Dim oFile As Object
'Dim intFolder As Integer '­nÅª¨úªºFolder¼Æ; ex:Inbox ©M Junk Email
'
'On Error GoTo ErrNo1
'
'   If cnnConnection.State = adStateClosed Then Exit Function '±ß¤WDBÂ_½u,¤£»Ý©¹¤U°õ¦æ
'   '¥H§KTimer¦P®ÉRun°_¨Ó
'   If LblFCPin.BackColor = vbBlue Then Exit Function
'   If LblFCPout.BackColor = vbBlue Then Exit Function
'   If LblPatent.BackColor = vbBlue Then Exit Function
'   If LblTM.BackColor = vbBlue Then Exit Function
'
'   strErrText = "" 'Add By Sindy 2020/7/22
'   importPatentMail = False
'   If txtPathPatent = "" Then
'      MsgBox "¦¬¥ó¸ê®Æ§¨¤£¥iªÅ¥Õ¡I"
'      txtPathPatent.SetFocus
'      Exit Function
'   End If
'   If Dir(txtPathPatent, vbDirectory) = "" Then
'      MkDir txtPathPatent
'   End If
'
'   strMRL01 = Left(Patent¦¬¥ó§X, 2): strMRL02 = "": strMRL03 = ""
'   If ExecuteSchedule(strMRL01, strMRL02, strMRL03) = True Or bolPatentRun = True Then '­n°õ¦æTimer
''      'Add By Sindy 2023/11/29
''      Set eventConn = cnnConnection
''      KillCmdLog
''      '2023/11/29 END
'
'      bolPatentRun = False
'
'      'Add By Sindy 2018/2/8 ¬Â¬Â»¡¤À«H´N¦o©M¶®®S¸g²z¦b³B²z,¥ð°²®É¤£¶·ÂàÂ¾¥N,¤H­û¥ð°²®É¤£¦¬³qª¾«H
'      strPTo = Pub_GetSpecMan("±M§Q³B«H¥ó³B²z¤H")
'      ArrStr = Split(strPTo, ";")
'      strPTo = ""
'      For jj = 0 To UBound(ArrStr)
'         'ÀË¬d¬O§_¥ð°²
'         If CheckIsPersonRest(CStr(ArrStr(jj)), strSrvDate(1), Format(Left(Right("000000" & ServerTime, 6), 4), "##:##")) = False Then
'            If strPTo <> "" Then strPTo = strPTo & ";"
'            strPTo = strPTo & CStr(ArrStr(jj))
'         End If
'      Next jj
'      If strPTo = "" Then strPTo = Pub_GetSpecMan("±M§Q³B«H¥ó³B²z¤H")
'      '2018/2/8 END
'
'      strErrText = "Pa-A:" 'Add By Sindy 2023/7/11
'      Set olApp = CreateObject("Outlook.Application")
'      strErrText = "Pa-B:" 'Add By Sindy 2023/7/11
'      Set myNamespace = olApp.GetNamespace("MAPI")
'
'      intKeyCnt = 0: intRunOK = 0: intCaseOK = 0
'
'strErrText = "Pa-C:" 'Add By Sindy 2023/7/11
'   'Add By Sindy 2024/1/31
'   For intFolder = 1 To 1 '2
'      'Modify By Sindy 2023/7/17
'      If OpenOutLookFolder(myNamespace, myFolder, Left(Patent¦¬¥ó§X, 2), intFolder) = False Then
'         importPatentMail = True
'         Set olApp = Nothing
'         Set myNamespace = Nothing
'         Set myFolder = Nothing
'         TmrPatent.Interval = 0
'         LblPatent.BackColor = vbRed
'         Exit Function
'      End If
'      '2023/7/17 END
'
'      bolReStarPatent = False
'
'ReStarPatent:
'      strErrText = "Pa-D:" 'Add By Sindy 2023/7/11
'      Set myItems = myFolder.Items
'      strErrText = "Pa-E:" 'Add By Sindy 2023/7/11
'      strIPMNoteSMIME = "" '¥[±K¥D¦®
'      intMaxItem = myItems.Count
'
'      strErrText = "Pa-F:" 'Add By Sindy 2023/7/11
'      '°O¿ýLogÀÉ
'      'Modify By Sindy 2024/1/31 + And intFolder = 1
'      If strMRL02 = "" And intFolder = 1 Then
'         'strMRL01 = Left(Patent¦¬¥ó§X, 2)
'         strMRL02 = strSrvDate(1)
'         strMRL03 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         strMRL05 = strUserNum
'         strSql = "insert into MailReceiveLog(MRL01,MRL02,MRL03,MRL05,MRL09)" & _
'                  "values('" & strMRL01 & "'," & strMRL02 & "," & Format(strMRL03, "hhmmss") & ",'" & strMRL05 & "','Y')"
'         cnnConnection.Execute strSql
'      End If
'
'      strErrText = "Pa-G: intMaxItem=" & intMaxItem 'Add By Sindy 2023/7/11
'      If intMaxItem > 0 Then
'         If bolUserControl = True Then
'            frmpic002.Label1.Caption = "¶l¥ó±µ¦¬¤¤...½Ðµy­Ô..."
'            frmpic002.Show
'            frmpic002.ZOrder 0
'            frmpic002.Label1.Font.Size = 12
'            frmpic002.Label1.Font.Bold = True
'         End If
'         For mail_ii = myItems.Count To 1 Step -1
'            LblPatent.BackColor = vbBlue 'ÂÅ¦âTimer¥¿¦bRun
'            cmdCancel(2).Enabled = True
'            DoEvents
'            If bolUserControl = True Then
'               frmpic002.Label1.Caption = "¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii & "...½Ðµy­Ô~"
'            Else
'               Frame3.Caption = Frame3.Tag & "¡@¡@¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii
'            End If
'            DoEvents
'            strErrText = ""
'            intRunOK = intRunOK + 1 '°O¿ý¥þ³¡±µ¦¬ªºµ§¼Æ
'            strRecipients_1 = "": strRecipients_all = "" '§ì¦¬¥óªÌ¸ê®Æ
'            Call ReadMailText(myItems, True, strRecipients_all, strRecipients_1)
'
'            'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'            strErrText = "²Ä " & mail_ii & " µ§ ¥D¦®: " & strSocSubject & vbCrLf
'            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strSender: " & strSender & vbCrLf
'            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strMailDateTime: " & strMailDate & " " & strMailTime
'            Call WLog_Day(strErrText, ±M§Q³B¦¬¥ó«H½c)
'
'            'IPM.Note.SMIME ¥[±K
'            'Modify By Sindy 2017/11/17
'            'Modify By Sindy 2023/7/12 + Or myItems.Item(mail_ii).Class = 45 : ·s³qª¾ => UCase(myItems.Item(mail_ii).MessageClass) = UCase("IPM.Post")
'            If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or myItems.Item(mail_ii).Class = 45 Then
'            'If myItems.Item(mail_ii).Class <> 43 Then
'            '2017/11/17 END
'               intKeyCnt = intKeyCnt + 1
'               '¥[Log°O¿ý
'               'strErrText = "²Ä " & mail_ii & " µ§ [¥[±K] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[¥[±K¶l¥ó]" & vbCrLf, ±M§Q³B¦¬¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf '¥[±K¥D¦®
'            'Add By Sindy 2020/4/10 ¦^¦¬¶l¥ó,ª½±µ§R°£
'            ElseIf InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
'               intKeyCnt = intKeyCnt + 1
'               'strErrText = "²Ä " & mail_ii & " µ§ [¦^¦¬] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[¦^¦¬¶l¥ó]" & vbCrLf, ±M§Q³B¦¬¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'               'myItems.Item(mail_ii).Delete '§R°£ =>µLªk§R°£,·|·í
'               'DoEvents
'            Else
'
'               strFileName = mail_ii & "." & _
'                             strSrvDate(1) & Right("000000" & ServerTime, 6) & ".msg"
'               myItems.Item(mail_ii).SaveAs txtPathPatent & "\" & strFileName, 9 '9.Outlook Unicode¶l¥ó®æ¦¡.msg
'               'Add By Sindy 2020/2/27
'               Sleep 1000
'               DoEvents
'               Call WLog_Day("²£¥Í¼È¦s¹q¤lÀÉ: " & txtPathPatent & "\" & strFileName, ±M§Q³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'               '2020/2/27 END
'
'               'Add By Sindy 2022/2/22
'               '«H¥ó¦P®É¦³±Hipdept¤Îpatent«H½c®É,¤~ÀË¬d:
'               If InStr(UCase(strRecipients_all), UCase("patent@taie.")) > 0 And _
'                  InStr(UCase(Replace(strRecipients_all, "80ipdept@taie.com.tw", "")), UCase("ipdept@taie.")) > 0 Then
'                  strMailTime_Recv = Format(myItems.Item(mail_ii).ReceivedTime, "HHMM") '¼W¥[§PÂ_ ReceivedTime ®É¶¡
'                  '¥ý¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶i¨Ó¤F¡F­Y¦³¡A§R°£¡C­Y¨S¦³¡AÄ~Äò¡C
'                  'Modify By Sindy 2022/10/26 µo¥Í¥D¦®¬OªÅ¥Õ,¦P®É±H2­Ó«H½c
'                  If strSocSubject = "" Then
'                     'Modify By Sindy 2023/7/13 ¼W¥[§PÂ_ strMailTime_Recv
'                     strSql = "select pi01,pi03 from patentinput" & _
'                              " where pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate) & _
'                              " and (substr(lpad(pi13,6,0),1,4) = " & Format(strMailTime, "HHMM") & " or substr(lpad(pi13,6,0),1,4) = " & strMailTime_Recv & ")" & _
'                              " order by pi01 desc,pi03 desc"
'                  Else
'                  '2022/10/26 END
'                     'Modify By Sindy 2023/7/13 ¼W¥[§PÂ_ strMailTime_Recv
'                     strSql = "select pi01,pi03 from patentinput" & _
'                              " where pi17 = '" & ChgSQL(strSocSubject) & "'" & _
'                              " and pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate) & _
'                              " and (substr(lpad(pi13,6,0),1,4) = " & Format(strMailTime, "HHMM") & " or substr(lpad(pi13,6,0),1,4) = " & strMailTime_Recv & ")" & _
'                              " order by pi01 desc,pi03 desc"
'                  End If
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     '«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]
'                     intKeyCnt = intKeyCnt + 1
'                     Call WLog_Day("[«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]", ±M§Q³B¦¬¥ó«H½c)
'                     strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'                     Call DeleteMyItems(myItems, ±M§Q³B¦¬¥ó«H½c) '§R°£Outlook¸Ì­±ªº¶l¥ó
'                     '§R°£PCºÝÀÉ®×
'                     Set fs = CreateObject("Scripting.FileSystemObject")
'                     Call fs.DeleteFile(txtPathPatent & "\" & strFileName)
'                     Sleep 1000
'                     DoEvents
'                     GoTo IsReadNext 'Run¤U¤@µ§
'                  Else
'                     'ÀË¬d°ê¥~³¡¬O§_¦³¦¹µ§¸ê®Æ
'                     'Modify By Sindy 2023/7/13 ¼W¥[§PÂ_ strMailTime_Recv
'                     strSql = "select ii01,ii03 from ipdeptinput" & _
'                              " where ii17 = '" & ChgSQL(strSocSubject) & "'" & _
'                              " and ii11 = '" & ChgSQL(strSender) & "' and ii12 = " & DBDATE(strMailDate) & _
'                              " and (substr(lpad(ii13,6,0),1,4) = " & Format(strMailTime, "HHMM") & " or substr(lpad(ii13,6,0),1,4) = " & strMailTime_Recv & ")" & _
'                              " order by ii01 desc,ii03 desc"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 1 Then
'                        '³oª¬ªp¬O¤£À³¸Óµo¥Íªº
'                        PUB_SendMail strUserNum, "97038", "", _
'                           "¡iPatent-¦¹µ§¶l¥ó°ê¥~³¡¤w¦¬¿ý(" & RsTemp.Fields("ii01") & "-" & RsTemp.Fields("ii03") & "),±M§Q³B¥¼¤@¨Ö¦¬¿ý,½ÐÀË¬dª¬ªp¡H(Ä~Äò©¹¤URun,¶i¦æ¶l¥ó¦¬¿ý...)¡j", strSocSubject & vbCrLf & vbCrLf & strSql, , txtPathPatent & "\" & strFileName, , , , , , , , True, False, , , False, , , False
'                        'Ä~Äò©¹¤URun,¶i¦æ¶l¥ó¦¬¿ý...
'                     Else
'                        '*****
'                        'µ¥°ê¥~³¡«H½c¦¬¿ý¦¹µ§¬Û¦P¶l¥ó(²Î¤@¦¬¿ý)
'                        '*****
'
'                        '°»´ú¬O§_¦³²§±`ªºª¬ªp,³qª¾¹q¸£¤¤¤ß
'                        'ex:Invoice 222088 from Patentica Limited -  P-500/2RU -- CFP-025048
'                        '¦³¬í®t,©Ò¥H±M§Q«H¥ó·|´Ý¯dµÛ,­nÃöª`
'                        If DBDATE(strMailDate) < strSrvDate(1) Or _
'                           (DBDATE(strMailDate) = strSrvDate(1) And (Val(Format(Time, "HH")) - Val(Format(strMailTime, "HH"))) > 1) Then
'                           If bolReStarPatent = True Then
'                              PUB_SendMail strUserNum, "97038", "", _
'                                 "¡iPatent-¦¹µ§¶l¥ó¦P®É¦³±Hipdept¤Îpatent«H½c,ÁÙ¥¼¶i¦æ¦¬¿ý,½ÐÀË¬dª¬ªp¡H(ÀË¬d¬O§_¦³¬í®t,©Ò¥H±M§Q«H¥ó·|´Ý¯dµÛ ©Î Patent«H½c¥ý±Ò°Ê¤F)¡j" & strSocSubject, strSocSubject & vbCrLf & vbCrLf & strSql, , txtPathPatent & "\" & strFileName, , , , , , , , True, False, , , False, , , False
'                           End If
'                        End If
'
'                        'Add By Sindy 2023/7/14 patent´«¤F¤½¥Î¸ê®Æ§¨,®É¶¡©Mipdept°t¤£°_¨Ó
'                        'Print Format(myItems.Item(mail_ii).ReceivedTime, "HH:MM:SS")=16:49:28
'                        'Print Format(myItems.Item(mail_ii).SentOn, "HH:MM:SS")=16:49:28
'                        If strSocSubject <> "" Then
'                           strSql = "select pi01,pi03 from patentinput" & _
'                                    " where pi17 = '" & ChgSQL(strSocSubject) & "'" & _
'                                    " and pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate)
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                           If intI = 1 Then
'                              If RsTemp.RecordCount = 1 Then
'                                 strSql = "select ii01,ii03 from ipdeptinput" & _
'                                          " where ii17 = '" & ChgSQL(strSocSubject) & "'" & _
'                                          " and ii11 = '" & ChgSQL(strSender) & "' and ii12 = " & DBDATE(strMailDate)
'                                 intI = 1
'                                 Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                                 If intI = 1 Then
'                                    If RsTemp.RecordCount = 1 Then
''                                       PUB_SendMail strUserNum, "97038", "", _
''                                          "(¤w§RÀÉ)¡iPatent-¦¹µ§¶l¥ó¦P®É¦³±Hipdept¤Îpatent«H½c,À³¸Ó¤w¦¬¿ý,¨Ï¥Î(«H½c¤À«H¬ö¿ý¬d¸ß)ÀË¬d¬O§_¦³¦¬¶iipdept¤Îpatent«H½c¡j" & strSocSubject, strSocSubject & vbCrLf & _
''                                          "strMailTime_Recv = " & strMailTime_Recv & vbCrLf & vbCrLf & strSql, , txtPathPatent & "\" & strFileName, , , , , , , , True, False, , , False, , , False
'
'                                       '«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]
'                                       intKeyCnt = intKeyCnt + 1
'                                       Call WLog_Day("[«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]", ±M§Q³B¦¬¥ó«H½c)
'                                       strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'                                       Call DeleteMyItems(myItems, ±M§Q³B¦¬¥ó«H½c) '§R°£Outlook¸Ì­±ªº¶l¥ó
'                                    End If
'                                 End If
'                              End If
'                           End If
'                        End If
'                        '2023/7/14 END
'
'                        '§R°£PCºÝÀÉ®×
'                        Set fs = CreateObject("Scripting.FileSystemObject")
'                        Call fs.DeleteFile(txtPathPatent & "\" & strFileName)
'                        Sleep 1000
'                        DoEvents
'                        GoTo IsReadNext 'Run¤U¤@µ§
'                     End If
'                  End If
'               End If
'               '2022/2/22 END
'
'               If intErr2147024882 <> mail_ii Then
'                  'Add By Sindy 2018/4/12
'                  If Dir(txtPathPatent & "\" & strFileName) = "" Then
'                     strErrText = "µL²£¥Í¹q¤lÀÉ,ºÃ¦ü¤¤¯f¬r " & "Err.Number:" & Err.Number & Err.Description & vbCrLf
'                     Call ExportEMailErr(myItems, False, ±M§Q³B¦¬¥ó«H½c, strErrText, Err.Number, Err.Description, _
'                           strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'                  'Add By Sindy 2020/4/14 ÀË¬d¹q¤lÀÉ¬O§_¥i¥H¥¿±`¶}±Ò
'                  ElseIf ChkIsOpenEmail(txtPathPatent & "\" & strFileName, strErrCode, strErrDesc) = False Then
'                     intKeyCnt = intKeyCnt + 1
'                     strErrText = "²Ä " & mail_ii & " µ§ [MsgµLªk¶}±Ò] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf & _
'                        txtPathPatent & "\" & strFileName & vbCrLf & _
'                        "Err.Number:" & strErrCode & strErrDesc & vbCrLf
'                     Call WLog_Day(strErrText, ±M§Q³B¦¬¥ó«H½c)
'                     strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'                  Else
'                  '2018/4/12 END
'
'                     Sleep 100 'Add By Sindy 2019/12/13
'                     If PUB_PatentTransMail(Me, strTo, strErrText, strKind, strFileName, strCaseNo) = True Then
'                        Call DeleteMyItems(myItems, ±M§Q³B¦¬¥ó«H½c) '§R°£Outlook¸Ì­±ªº¶l¥ó
'
'                        If strCaseNo <> "" Then
'                           intCaseOK = intCaseOK + 1
'                        End If
'
'                     Else
'                        strErrNumber = Err.Number 'Add By Sindy 2019/10/14
'                        Call WLog_Day("¤À«H¥¢±Ñ(1): " & strErrText, ±M§Q³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                        'Add By Sindy 2020/9/10
'                        If strErrText <> "" And strErrText <> "Err.Number:0;" Then
'                        Else
'                        '2020/9/10 END
'                           'Add By Sindy 2019/12/11
'                           If strErrNumber = "0" Then
'                              strErrText = "§ä¤£¨ìÀÉ®×,ºÃ¦ü¤¤¯f¬r"
'   '                           myItems.Item(mail_ii).Delete '§R°£
'   '                           DoEvents
'                           End If
'                           '2019/12/11 END
'                        End If
'
'                        Call WLog_Day("¤À«H¥¢±Ñ(2): " & strErrText & ";" & Err.Number & ":" & Err.Description, ±M§Q³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                        Call ExportEMailErr(myItems, False, ±M§Q³B¦¬¥ó«H½c, strErrText, Err.Number, Err.Description, _
'                           strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'                        'Add By Sindy 2019/10/14
'                        'If strErrNumber = "999" Then
'                        If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then
'                           Call WLog_Day("¤À«H¥¢±Ñ(3): 999 " & strErrText & vbCrLf, ±M§Q³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                           Exit For
'                        End If
'                        '2019/10/14 END
'                     End If
'
'                  End If
'               'Modify By Sindy 2020/4/15
'               Else
'                  intErr2147024882 = 0
'               '2020/4/15 END
'               End If
'            End If
'IsReadNext:
'            '¬O§_­n¤¤Â_
'            If bolCancel(2) = True Then
'               LblPatent.BackColor = vbRed
'               DoEvents 'Add By Sindy 2024/5/7
'               GoTo IsCancel
'            End If
'         Next mail_ii
'
'IsCancel:
'         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         If bolUserControl = True Then
'            Unload frmpic002
'            Set frmpic002 = Nothing
'         End If
'
'         '°O¿ýLogÀÉ
'         'Add By Sindy 2024/1/31
'         If intFolder = 1 Then
'         '2024/1/31 END
'            '" and MRL05='" & strMRL05 & "'"
'            strSql = "update MailReceiveLog set" & _
'                     " MRL04=" & Format(strMRL04, "hhmmss") & _
'                     ",MRL06=" & intRunOK & ",MRL07=" & intKeyCnt & ",MRL08=" & intCaseOK & _
'                     ",MRL09='" & IIf(bolCancel(2) = True, "B", "E") & "'" & _
'                     " where MRL01='" & strMRL01 & "'" & _
'                     " and MRL02=" & strMRL02 & _
'                     " and MRL03=" & Format(strMRL03, "hhmmss")
'            cnnConnection.Execute strSql
'            m_RunPatentStarTime = strMRL03
'            m_RunPatentEndTime = Format(strMRL04, "hh:mm:ss")
'         End If
'         If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then GoTo NotRunSec 'Add By Sindy 2023/2/18
'
'         '°õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡A­Y¥u³Ñ¤U¥[±K¶l¥ó´Nµo«H³qª¾±M§Q³B¶l¥ó³B²z¤H­û
'         '¦³«D¥[±K¶l¥ó¦A°õ¦æ¤@¦¸±µ¦¬
''         DoEvents
'         Set myItems = myFolder.Items
'         intMaxItem = myItems.Count
'         If intMaxItem > 0 Then
'            strErrText = "": intKeyCnt = 0
'            For mail_ii = myItems.Count To 1 Step -1
'               Call ReadMailText(myItems, False)
'               'Modify By Sindy 2017/11/17
'               'Modify By Sindy 2020/4/10 + IPM.Outlook.Recall
'               If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or _
'                  InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
'               'If myItems.Item(mail_ii).Class <> 43 Then
'               '2017/11/17 END
'                  If strErrText = "" Then
'                     strErrText = "***¡@(Patent) °õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡@*********************************" & vbCrLf
'                  End If
'                  intKeyCnt = intKeyCnt + 1
'                  strErrText = strErrText & "²Ä¡@" & mail_ii & "¡@µ§¡@[¥[±K]¡@¥D¦®:¡@" & strSocSubject & vbCrLf
'               Else
'                  If bolReStarPatent = False And bolCancel(2) = False Then
'                     bolReStarPatent = True
'                     Call WLog_Day("[­«Run²Ä¤G¦¸]" & vbCrLf, ±M§Q³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                     '­«Run²Ä¤G¦¸
'                     GoTo ReStarPatent
'                  'Add By Sindy 2022/8/5 ¤¤Â_´N¤£­n¦AÀË¬d¤F,©¹¤U°õ¦æ
'                  ElseIf bolCancel(2) = True Then
'                     Exit For
'                  '2022/8/5 END
'                  End If
'               End If
'            Next mail_ii
'
'            If strErrText <> "" Then
'               '¦³¥[±K«H¥ó¥B¬°¤u§@¤Ñ¤~­n±H«H³qª¾¤H­û³B²z
'               If ChkWorkDay(strSrvDate(1)) = True And _
'                  (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
'                  '±HE-Mail³qª¾¦¬¥ó³B²z¤H­û
'                  If UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '´ú¸Õ¸ê®Æ®w
'                     strTo = m_M51Recver
'                  Else
'                     strTo = strPTo 'Pub_GetSpecMan("±M§Q³B«H¥ó³B²z¤H")
'                  End If
'                  PUB_SendMail strUserNum, strTo, "", ±M§Q³B¦¬¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð³B²z¡I", strIPMNoteSMIME & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
'                        "* ¶i¤J¨ä«H½c¸Ñ±K«áÂà±Hµ¹" & ±M§Q³B¦¬¥ó«H½c & "¡A¦A±N­ì¥[±K¶l¥ó§R°£¡AÁ×§K­«ÂÐ¡]¤Á°O¡^¡A«Ý¨t²Î¤U¦¸´`Àô³B²z¡C", , , , , , , , , , IIf(strTo = m_M51Recver, False, True), False, , , False, , , False
''                  DoEvents
'               End If
'            End If
'         End If
'      End If 'Add By Sindy 2024/1/31
'   Next intFolder 'Add By Sindy 2024/1/31
'
'NotRunSec:
'      If intRunOK > 0 Then 'Add By Sindy 2024/1/31
'         'Modify By Sindy 2017/12/27 ¤u§@¤Ñ¤~­n³qª¾
'         If ChkWorkDay(strSrvDate(1)) = True And _
'            (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
'            'ÀË¬d¦¬¥ó¸ê®Æ§¨¤¤¬O§_¦³´Ý¯dÀÉ®×
'            Set oFolder = oFileSys.GetFolder(txtPathPatent.Text)
'            Set fs = CreateObject("Scripting.FileSystemObject")
'            If oFolder.files.Count > 0 Then
'               'Add By Sindy 2023/9/13
'               For Each oFile In oFolder.files
'                  Set myItems = olApp.CreateItemFromTemplate(txtPathPatent.Text & "\" & oFile.Name)
'                  Call ReadMailText_File(myItems)
'                  '¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶×¤J?­Y¦³=§R°£¡C­Y¨S¦³=¤£³B²z,µ¥¤H­û¬d¬Ý
'                  strSql = "select pi01,pi03 from patentinput" & _
'                           " where pi17 = '" & ChgSQL(strSocSubject) & "'" & _
'                           " and pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate) & " and pi13 = " & Val(Replace(strMailTime, ":", "")) & _
'                           " order by pi01 desc,pi03 desc"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     '§R°£PCºÝÀÉ®×
'                     Call fs.DeleteFile(txtPathPatent & "\" & oFile.Name)
'                     Sleep 1000
'                     DoEvents
'                  End If
'               Next
'               Set oFolder = oFileSys.GetFolder(txtPathPatent.Text)
'               If oFolder.files.Count > 0 Then
'               '2023/9/13 END
'                  PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "±M§Q³B¦¬¥ó¸ê®Æ§¨:" & txtPathPatent.Text & "©|¦³´Ý¯dÀÉ®×(" & oFolder.files.Count & "­Ó),½ÐÀË¬d¡I", "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
'               End If
'            End If
'            'Modify By Sindy 2018/10/1 ¶®®S:¨ú®ø¦¹³qª¾
''            'Add By Sindy 2017/12/20 ÀË¬d¬O§_¦³«H¥ó¥¼Âà±H
''            If UCase(pub_DbTerminalName) = ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '¥¿¦¡¸ê®Æ®w¤~µo«H
''               strExc(0) = "SELECT COUNT(*) FROM patentinput WHERE pi08=0"
''               intI = 1
''               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
''               If rsA.Fields(0) > 0 Then
''                  'PUB_SendMail strUserNum, Pub_GetSpecMan("±M§Q³B«H¥ó³B²z¤H"), "", "ª`·N¡G" & ±M§Q³B¦¬¥ó«H½c & "©|¦³¥¼Âà±H«H¥ó«Ý³B²z¡I", "¦P¥D¦®", , , , , , , , , , True, False, , , , , , False
''                  PUB_SendMail strUserNum, strPTo, "", "ª`·N¡G" & ±M§Q³B¦¬¥ó«H½c & "©|¦³¥¼Âà±H«H¥ó«Ý³B²z¡I", "¦P¥D¦®", , , , , , , , , , True, False, , , , , , False
''                  DoEvents
''               End If
''            End If
''            '2017/12/20 END
'         End If
'
'      Else
'         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         '°O¿ýLogÀÉ
'         strSql = "update MailReceiveLog set" & _
'                  " MRL04=" & Format(strMRL04, "hhmmss") & _
'                  ",MRL06=0,MRL07=0,MRL08=0" & _
'                  ",MRL09='E'" & _
'                  " where MRL01='" & strMRL01 & "'" & _
'                  " and MRL02=" & strMRL02 & _
'                  " and MRL03=" & Format(strMRL03, "hhmmss")
'         cnnConnection.Execute strSql
'         m_RunPatentStarTime = strMRL03
'         m_RunPatentEndTime = Format(strMRL04, "hh:mm:ss")
'      End If
'
'      txtMRL02 = strSrvDate(2)
'      Call cmdQuery_Click
'      Frame3.Caption = Frame3.Tag
'      DoEvents
'
''      'Add By Sindy 2023/11/29
''      Set eventConn = Nothing
''      WCmdLog "importPatentMail µ²§ô"
''      WCmdLog ""
''      '2023/11/29 END
'   End If
'
'   cmdCancel(2).Enabled = False
'   '­n¤¤Â_
'   If bolCancel(2) = True Then
'      bolCancel(2) = False
'      TmrPatent.Interval = 0: LblPatent.BackColor = vbRed
'   Else
'   '¥¿±`µ²§ô
'      If TmrPatent.Interval > 0 Then
'         TmrPatent.Interval = dblTmrPatent
'         LblPatent.BackColor = vbGreen
'      Else
'         LblPatent.BackColor = vbRed
'      End If
'   End If
'
'   importPatentMail = True
'
'   Set olApp = Nothing
'   Set myNamespace = Nothing
'   Set myFolder = Nothing
'   Set myItems = Nothing
'   Set oFolder = Nothing
'   Set rsA = Nothing
'   Set fs = Nothing
'   Set oFile = Nothing
'
'   Exit Function
'
'ErrNo1:
'   Screen.MousePointer = vbDefault
'   intErr2147024882 = ExportEMailErr(myItems, True, ±M§Q³B¦¬¥ó«H½c, "(ErrNo1) " & strErrText & "; strSql=" & strSql, Err.Number, Err.Description, _
'                        strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'   On Error GoTo 0: Err.Clear
'   If intErr2147024882 > 0 Then
'      Call WLog_Day("intErr2147024882 > 0", ±M§Q³B¦¬¥ó«H½c)
'      'Resume Next
'      GoTo ReStarPatent
'      Exit Function
'   End If
'
'   cmdCancel(2).Enabled = False
'   TmrPatent.Interval = dblTmrPatent: LblPatent.BackColor = vbGreen
'
'   Set olApp = Nothing
'   Set myNamespace = Nothing
'   Set myFolder = Nothing
'   Set myItems = Nothing
'   Set oFolder = Nothing
'   Set rsA = Nothing
'   Set fs = Nothing
'   Set oFile = Nothing
'End Function

'Add By Sindy 2019/3/28
Private Sub TmrTM_Timer()
   'Modify By Sindy 2024/5/13
   'Call importTMMail
   Call ChkExecutionTimer(Left(TM¦¬¥ó§X, 2))
   '2024/5/13 END
End Sub

''°Ó¼Ð³B¦¬¥ó«H½c³B²zµ{§Ç
'Private Function importTMMail() As Boolean
'Dim kk As Integer, jj As Integer
'Dim strTo As String, strCC As String, strTempCC As String
'Dim oFileSys As New FileSystemObject, oFolder As Object
'Dim strKind As String
'Dim myForward As Object
'Dim myNewEmail As Object 'Âà±H«H¥ó
'Dim ArrStr As Variant, ArrStrkk As Variant
'Dim strCaseNo As String
'Dim strIPMNoteSMIME As String '¥[±K¥D¦®
'Dim bolReStarTM As Boolean
'Dim strMRL01 As String, strMRL02 As String, strMRL03 As String, strMRL04 As String, strMRL05 As String
'Dim rsA As New ADODB.Recordset
'Dim strPTo As String 'Add By Sindy 2018/2/8
'Dim strErrNumber As String 'Add By Sindy 2019/10/14
'Dim intURGENT As Integer 'Add By Sindy 2019/11/14
'Dim strErrCode As String, strErrDesc As String 'Add By Sindy 2020/4/15
''Add By Sindy 2023/6/26
'Dim olApp As Object
'Dim myNamespace As Object
'Dim myFolder As Object
'Dim myItems As Object
''2023/6/26 END
'Dim fs As Object, oFile As Object 'Add By Sindy 2023/9/13
'Dim intFolder As Integer '­nÅª¨úªºFolder¼Æ; ex:Inbox ©M Junk Email
'
'On Error GoTo ErrNo1
'
'   If cnnConnection.State = adStateClosed Then Exit Function '±ß¤WDBÂ_½u,¤£»Ý©¹¤U°õ¦æ
'   '¥H§KTimer¦P®ÉRun°_¨Ó
'   If LblFCPin.BackColor = vbBlue Then Exit Function
'   If LblFCPout.BackColor = vbBlue Then Exit Function
'   If LblPatent.BackColor = vbBlue Then Exit Function
'   If LblTM.BackColor = vbBlue Then Exit Function
'
'   strErrText = "TM-A:" 'Add By Sindy 2020/7/22
'   importTMMail = False
'   If txtPathTM = "" Then
'      MsgBox "¦¬¥ó¸ê®Æ§¨¤£¥iªÅ¥Õ¡I"
'      txtPathTM.SetFocus
'      Exit Function
'   End If
'   If Dir(txtPathTM, vbDirectory) = "" Then
'      MkDir txtPathTM
'   End If
'
'   strErrText = "TM-B:" 'Add By Sindy 2023/7/11
'   strMRL01 = Left(TM¦¬¥ó§X, 2): strMRL02 = "": strMRL03 = ""
'   If ExecuteSchedule(strMRL01, strMRL02, strMRL03) = True Or bolTMRun = True Then '­n°õ¦æTimer
''      'Add By Sindy 2023/11/29
''      Set eventConn = cnnConnection
''      KillCmdLog
''      '2023/11/29 END
'
'      bolTMRun = False
'
'      '¤À«H³B²z¤H­û:¥ð°²®É¤£¶·ÂàÂ¾¥N,¤H­û¥ð°²®É¤£¦¬³qª¾«H
'      strPTo = Pub_GetSpecMan("°Ó¼Ð³B«H¥ó³B²z¤H")
'      ArrStr = Split(strPTo, ";")
'      strPTo = ""
'      For jj = 0 To UBound(ArrStr)
'         'ÀË¬d¬O§_¥ð°²
'         If CheckIsPersonRest(CStr(ArrStr(jj)), strSrvDate(1), Format(Left(Right("000000" & ServerTime, 6), 4), "##:##")) = False Then
'            If strPTo <> "" Then strPTo = strPTo & ";"
'            strPTo = strPTo & CStr(ArrStr(jj))
'         End If
'      Next jj
'      If strPTo = "" Then strPTo = Pub_GetSpecMan("°Ó¼Ð³B«H¥ó³B²z¤H")
'
'      strErrText = "TM-C:" 'Add By Sindy 2023/7/11
'      Set olApp = CreateObject("Outlook.Application")
'      strErrText = "TM-D:" 'Add By Sindy 2023/7/11
'      Set myNamespace = olApp.GetNamespace("MAPI")
'      intKeyCnt = 0: intRunOK = 0: intCaseOK = 0
'
'strErrText = "TM-E-0:" 'Add By Sindy 2023/7/11
'   'Add By Sindy 2024/1/31
'   For intFolder = 1 To 1 '2
'      'Modify By Sindy 2023/7/17
'      If OpenOutLookFolder(myNamespace, myFolder, Left(TM¦¬¥ó§X, 2), intFolder) = False Then
'         importTMMail = True
'         Set olApp = Nothing
'         Set myNamespace = Nothing
'         Set myFolder = Nothing
'         TmrTM.Interval = 0
'         LblTM.BackColor = vbRed
'         Exit Function
'      End If
'      '2023/7/17 END
'
'      bolReStarTM = False
'
'ReStarTM:
'      strErrText = "TM-E:" 'Add By Sindy 2023/7/11
'      Set myItems = myFolder.Items
'      strErrText = "TM-F:" 'Add By Sindy 2023/7/11
'      strIPMNoteSMIME = "" '¥[±K¥D¦®
'      intMaxItem = myItems.Count
'
'      '°O¿ýLogÀÉ
'      'Modify By Sindy 2024/1/31 + And intFolder = 1
'      If strMRL02 = "" And intFolder = 1 Then
'         'strMRL01 = Left(TM¦¬¥ó§X, 2)
'         strMRL02 = strSrvDate(1)
'         strMRL03 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         strMRL05 = strUserNum
'         strSql = "insert into MailReceiveLog(MRL01,MRL02,MRL03,MRL05,MRL09)" & _
'                  "values('" & strMRL01 & "'," & strMRL02 & "," & Format(strMRL03, "hhmmss") & ",'" & strMRL05 & "','Y')"
'         cnnConnection.Execute strSql
'      End If
'
'      If intMaxItem > 0 Then
'         If bolUserControl = True Then
'            frmpic002.Label1.Caption = "¶l¥ó±µ¦¬¤¤...½Ðµy­Ô..."
'            frmpic002.Show
'            frmpic002.ZOrder 0
'            frmpic002.Label1.Font.Size = 12
'            frmpic002.Label1.Font.Bold = True
'         End If
'         For mail_ii = myItems.Count To 1 Step -1
'            LblTM.BackColor = vbBlue 'ÂÅ¦âTimer¥¿¦bRun
'            cmdCancel(3).Enabled = True
'            DoEvents
'            If bolUserControl = True Then
'               frmpic002.Label1.Caption = "¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii & "...½Ðµy­Ô~"
'            Else
'               Frame4.Caption = Frame4.Tag & "¡@¡@¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii
'            End If
'            DoEvents
'            strErrText = "TM-G:"
'            intRunOK = intRunOK + 1 '°O¿ý¥þ³¡±µ¦¬ªºµ§¼Æ
'            Call ReadMailText(myItems, False)
'
'            'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'            strErrText = strErrText & "²Ä " & mail_ii & " µ§ ¥D¦®: " & strSocSubject & vbCrLf
'            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strSender: " & strSender & vbCrLf
'            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strMailDateTime: " & strMailDate & " " & strMailTime
'            Call WLog_Day(strErrText, °Ó¼Ð³B¦¬¥ó«H½c)
'
'            'IPM.Note.SMIME ¥[±K
'            'Modify By Sindy 2017/11/17
'            'Modify By Sindy 2023/7/12 + Or myItems.Item(mail_ii).Class = 45 : ·s³qª¾ => UCase(myItems.Item(mail_ii).MessageClass) = UCase("IPM.Post")
'            If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or myItems.Item(mail_ii).Class = 45 Then
'            'If myItems.Item(mail_ii).Class <> 43 Then
'            '2017/11/17 END
'               intKeyCnt = intKeyCnt + 1
'               '¥[Log°O¿ý
'               'strErrText = "²Ä " & mail_ii & " µ§ [¥[±K] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[¥[±K¶l¥ó]" & vbCrLf, °Ó¼Ð³B¦¬¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf '¥[±K¥D¦®
'            'Add By Sindy 2020/4/10 ¦^¦¬¶l¥ó,ª½±µ§R°£
'            ElseIf InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
'               intKeyCnt = intKeyCnt + 1
'               'strErrText = "²Ä " & mail_ii & " µ§ [¦^¦¬] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
'               Call WLog_Day("[¦^¦¬¶l¥ó]" & vbCrLf, °Ó¼Ð³B¦¬¥ó«H½c)
'               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'               'myItems.Item(mail_ii).Delete '§R°£ =>µLªk§R°£,·|·í
'               'DoEvents
'            Else
'
'               strFileName = mail_ii & "." & _
'                             strSrvDate(1) & Right("000000" & ServerTime, 6) & ".msg"
'               myItems.Item(mail_ii).SaveAs txtPathTM & "\" & strFileName, 9 '9.Outlook Unicode¶l¥ó®æ¦¡.msg
'               'Add By Sindy 2020/2/27
'               Sleep 1000
'               DoEvents
'               '2020/2/27 END
'               Call WLog_Day("²£¥Í¼È¦s¹q¤lÀÉ: " & txtPathTM & "\" & strFileName, °Ó¼Ð³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'
'               If intErr2147024882 <> mail_ii Then
'                  'Add By Sindy 2018/4/12
'                  If Dir(txtPathTM & "\" & strFileName) = "" Then
'                     strErrText = "µL²£¥Í¹q¤lÀÉ,ºÃ¦ü¤¤¯f¬r " & "Err.Number:" & Err.Number & Err.Description & vbCrLf
'                     Call ExportEMailErr(myItems, False, °Ó¼Ð³B¦¬¥ó«H½c, strErrText, Err.Number, Err.Description, _
'                           strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'                  'Add By Sindy 2020/4/14 ÀË¬d¹q¤lÀÉ¬O§_¥i¥H¥¿±`¶}±Ò
'                  ElseIf ChkIsOpenEmail(txtPathTM & "\" & strFileName, strErrCode, strErrDesc) = False Then
'                     intKeyCnt = intKeyCnt + 1
'                     strErrText = "²Ä " & mail_ii & " µ§ [MsgµLªk¶}±Ò] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf & _
'                        txtPathTM & "\" & strFileName & vbCrLf & _
'                        "Err.Number:" & strErrCode & strErrDesc & vbCrLf
'                     Call WLog_Day(strErrText, °Ó¼Ð³B¦¬¥ó«H½c)
'                     strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
'                  Else
'                  '2018/4/12 END
''                     If strSrvDate(1) >= TM¤À«H¨t²Î±Ò¥Î¤é Then
'                        If PUB_TMTransMail(Me, strTo, strErrText, strKind, strFileName, strCaseNo) = True Then
'                           Call DeleteMyItems(myItems, °Ó¼Ð³B¦¬¥ó«H½c) '§R°£Outlook¸Ì­±ªº¶l¥ó
'
'                           If strCaseNo <> "" Then
'                              intCaseOK = intCaseOK + 1
'                           End If
'                        Else
'                           strErrNumber = Err.Number 'Add By Sindy 2019/10/14
'                           Call WLog_Day("¤À«H¥¢±Ñ(1): " & strErrText, °Ó¼Ð³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                           'Add By Sindy 2020/9/10
'                           If strErrText <> "" And strErrText <> "Err.Number:0;" Then
'                           Else
'                           '2020/9/10 END
'                              'Add By Sindy 2019/12/11
'                              If strErrNumber = "0" Then
'                                 strErrText = "§ä¤£¨ìÀÉ®×,ºÃ¦ü¤¤¯f¬r"
'      '                           myItems.Item(mail_ii).Delete '§R°£
'      '                           DoEvents
'                              End If
'                              '2019/12/11 END
'                           End If
'
'                           Call WLog_Day("¤À«H¥¢±Ñ(2): " & strErrText & ";" & Err.Number & ":" & Err.Description, °Ó¼Ð³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                           Call ExportEMailErr(myItems, False, °Ó¼Ð³B¦¬¥ó«H½c, strErrText, Err.Number, Err.Description, _
'                              strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'                           'Add By Sindy 2019/10/14
'                           'If strErrNumber = "999" Then
'                           If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then
'                              Call WLog_Day("¤À«H¥¢±Ñ(3): 999 " & strErrText & vbCrLf, °Ó¼Ð³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                              Exit For
'                           End If
'                           '2019/10/14 END
'                        End If
''                     Else
''                        '¥¼¤W½u,¥ý§R°£«H¥ó,¥H§K«H¥ó¶V¨Ó¶V¦h
''                        Kill txtPathTM & "\" & strFileName
''                        myItems.Item(mail_ii).Delete '§R°£
''                        Sleep 100 'Add By Sindy 2019/12/13
''                     End If
'                  End If
'               'Modify By Sindy 2020/4/15
'               Else
'                  intErr2147024882 = 0
'               '2020/4/15 END
'               End If
'            End If
'            '¬O§_­n¤¤Â_
'            If bolCancel(3) = True Then
'               LblTM.BackColor = vbRed
'               DoEvents 'Add By Sindy 2024/5/7
'               GoTo IsCancel
'            End If
'         Next mail_ii
'
'IsCancel:
'         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         If bolUserControl = True Then
'            Unload frmpic002
'            Set frmpic002 = Nothing
'         End If
'
'         '°O¿ýLogÀÉ
'         'Add By Sindy 2024/1/31
'         If intFolder = 1 Then
'         '2024/1/31 END
'            '" and MRL05='" & strMRL05 & "'"
'            strSql = "update MailReceiveLog set" & _
'                     " MRL04=" & Format(strMRL04, "hhmmss") & _
'                     ",MRL06=" & intRunOK & ",MRL07=" & intKeyCnt & ",MRL08=" & intCaseOK & _
'                     ",MRL09='" & IIf(bolCancel(3) = True, "B", "E") & "'" & _
'                     " where MRL01='" & strMRL01 & "'" & _
'                     " and MRL02=" & strMRL02 & _
'                     " and MRL03=" & Format(strMRL03, "hhmmss")
'            cnnConnection.Execute strSql
'            m_RunTMStarTime = strMRL03
'            m_RunTMEndTime = Format(strMRL04, "hh:mm:ss")
'         End If
'         If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then GoTo NotRunSec 'Add By Sindy 2023/2/18
'
'         '°õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡A­Y¥u³Ñ¤U¥[±K¶l¥ó´Nµo«H³qª¾°Ó¼Ð³B¶l¥ó³B²z¤H­û
'         '¦³«D¥[±K¶l¥ó¦A°õ¦æ¤@¦¸±µ¦¬
''         DoEvents
'         Set myItems = myFolder.Items
'         intMaxItem = myItems.Count
'         If intMaxItem > 0 Then
'            strErrText = "": intKeyCnt = 0
'            For mail_ii = myItems.Count To 1 Step -1
'               Call ReadMailText(myItems, False)
'               'Modify By Sindy 2017/11/17
'               'Modify By Sindy 2020/4/10 + IPM.Outlook.Recall
'               If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or _
'                  InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
'               'If myItems.Item(mail_ii).Class <> 43 Then
'               '2017/11/17 END
'                  If strErrText = "" Then
'                     strErrText = "***¡@(TM) °õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡@*********************************" & vbCrLf
'                  End If
'                  intKeyCnt = intKeyCnt + 1
'                  strErrText = strErrText & "²Ä¡@" & mail_ii & "¡@µ§¡@[¥[±K]¡@¥D¦®:¡@" & strSocSubject & vbCrLf
'               Else
'                  If bolReStarTM = False And bolCancel(3) = False Then
'                     bolReStarTM = True
'                     Call WLog_Day("[­«Run²Ä¤G¦¸]" & vbCrLf, °Ó¼Ð³B¦¬¥ó«H½c) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
'                     '­«Run²Ä¤G¦¸
'                     GoTo ReStarTM
'                  'Add By Sindy 2022/8/5 ¤¤Â_´N¤£­n¦AÀË¬d¤F,©¹¤U°õ¦æ
'                  ElseIf bolCancel(3) = True Then
'                     Exit For
'                  '2022/8/5 END
'                  End If
'               End If
'            Next mail_ii
'
'            If strErrText <> "" Then
'               '¦³¥[±K«H¥ó¥B¬°¤u§@¤Ñ¤~­n±H«H³qª¾¤H­û³B²z
'               If ChkWorkDay(strSrvDate(1)) = True And _
'                  (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
'                  '±HE-Mail³qª¾¦¬¥ó³B²z¤H­û
'                  If strSrvDate(1) >= TM¤À«H¨t²Î±Ò¥Î¤é Then
'                     strTo = strPTo 'Pub_GetSpecMan("°Ó¼Ð³B«H¥ó³B²z¤H")
'                  Else
'                     strTo = m_M51Recver
'                  End If
'                  PUB_SendMail strUserNum, strTo, "", °Ó¼Ð³B¦¬¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð³B²z¡I", strIPMNoteSMIME & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
'                        "* ¶i¤J¨ä«H½c¸Ñ±K«áÂà±Hµ¹" & °Ó¼Ð³B¦¬¥ó«H½c & "¡A¦A±N­ì¥[±K¶l¥ó§R°£¡AÁ×§K­«ÂÐ¡]¤Á°O¡^¡A«Ý¨t²Î¤U¦¸´`Àô³B²z¡C", , , , , , , , , , IIf(strTo = m_M51Recver, False, True), False, , , False, , , False
''                  DoEvents
'               End If
'            End If
'         End If
'      End If 'Add By Sindy 2024/1/31
'   Next intFolder 'Add By Sindy 2024/1/31
'
'NotRunSec:
'      If intRunOK > 0 Then 'Add By Sindy 2024/1/31
'         'Modify By Sindy 2017/12/27 ¤u§@¤Ñ¤~­n³qª¾
'         If ChkWorkDay(strSrvDate(1)) = True And _
'            (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
'            'ÀË¬d¦¬¥ó¸ê®Æ§¨¤¤¬O§_¦³´Ý¯dÀÉ®×
'            Set oFolder = oFileSys.GetFolder(txtPathTM.Text)
'            Set fs = CreateObject("Scripting.FileSystemObject")
'            If oFolder.files.Count > 0 Then
'               'Add By Sindy 2023/9/13
'               For Each oFile In oFolder.files
'                  Set myItems = olApp.CreateItemFromTemplate(txtPathTM.Text & "\" & oFile.Name)
'                  Call ReadMailText_File(myItems)
'                  '¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶×¤J?­Y¦³=§R°£¡C­Y¨S¦³=¤£³B²z,µ¥¤H­û¬d¬Ý
'                  strSql = "select ti01,ti03 from tminput" & _
'                           " where ti17 = '" & ChgSQL(strSocSubject) & "'" & _
'                           " and ti11 = '" & ChgSQL(strSender) & "' and ti12 = " & DBDATE(strMailDate) & " and ti13 = " & Val(Replace(strMailTime, ":", "")) & _
'                           " order by ti01 desc,ti03 desc"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     '§R°£PCºÝÀÉ®×
'                     Call fs.DeleteFile(txtPathTM & "\" & oFile.Name)
'                     Sleep 1000
'                     DoEvents
'                  End If
'               Next
'               Set oFolder = oFileSys.GetFolder(txtPathTM.Text)
'               If oFolder.files.Count > 0 Then
'               '2023/9/13 END
'                  PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "°Ó¼Ð³B¦¬¥ó¸ê®Æ§¨:" & txtPathTM.Text & "©|¦³´Ý¯dÀÉ®×(" & oFolder.files.Count & "­Ó),½ÐÀË¬d¡I", "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
'               End If
'            End If
'            'ÀË¬d¬O§_¦³«H¥ó¥¼Âà±H
'            'If UCase(pub_DbTerminalName) = ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '¥¿¦¡¸ê®Æ®w¤~µo«H
'               strExc(0) = "SELECT COUNT(*) FROM TMinput WHERE Ti08=0"
'               intI = 1
'               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'               If rsA.Fields(0) > 0 Then
'                  'Add By Sindy 2019/11/14 ¥D¦®¸Ì¦³ URGENT ¦r¼ËªÌ,³qª¾«H­n¥[¦³«æ¥ó! => IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") &
'                  intURGENT = 0
'                  strExc(0) = "SELECT COUNT(*) FROM TMinput WHERE Ti08=0 and instr(upper(Ti17),'URGENT')>0"
'                  intI = 1
'                  Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'                  If rsA.Fields(0) > 0 Then
'                     intURGENT = rsA.RecordCount
'                  End If
'                  '2019/11/14 END
'                  If strSrvDate(1) >= TM¤À«H¨t²Î±Ò¥Î¤é Then
'                     'Modify By Sindy 2019/11/14 + IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") &
'                     PUB_SendMail strUserNum, strPTo, "", IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") & "ª`·N¡G" & °Ó¼Ð³B¦¬¥ó«H½c & "©|¦³¥¼Âà±H«H¥ó«Ý³B²z¡I", "¦P¥D¦®", , , , , , , , , , True, False, , , False, , , False
'   '                  DoEvents
'                  End If
'               End If
'            'End If
'         End If
'
'      Else
'         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
'         '°O¿ýLogÀÉ
'         strSql = "update MailReceiveLog set" & _
'                  " MRL04=" & Format(strMRL04, "hhmmss") & _
'                  ",MRL06=0,MRL07=0,MRL08=0" & _
'                  ",MRL09='E'" & _
'                  " where MRL01='" & strMRL01 & "'" & _
'                  " and MRL02=" & strMRL02 & _
'                  " and MRL03=" & Format(strMRL03, "hhmmss")
'         cnnConnection.Execute strSql
'         m_RunTMStarTime = strMRL03
'         m_RunTMEndTime = Format(strMRL04, "hh:mm:ss")
'      End If
'
'      txtMRL02 = strSrvDate(2)
'      Call cmdQuery_Click
'      Frame4.Caption = Frame4.Tag
'      DoEvents
'
''      'Add By Sindy 2023/11/29
''      Set eventConn = Nothing
''      WCmdLog "importTMMail µ²§ô"
''      WCmdLog ""
''      '2023/11/29 END
'   End If
'
'   cmdCancel(3).Enabled = False
'   '­n¤¤Â_
'   If bolCancel(3) = True Then
'      bolCancel(3) = False
'      TmrTM.Interval = 0: LblTM.BackColor = vbRed
'   Else
'   '¥¿±`µ²§ô
'      If TmrTM.Interval > 0 Then
'         TmrTM.Interval = dblTmrTM
'         LblTM.BackColor = vbGreen
'      Else
'         LblTM.BackColor = vbRed
'      End If
'   End If
'
'   importTMMail = True
'
'   Set olApp = Nothing
'   Set myNamespace = Nothing
'   Set myFolder = Nothing
'   Set myItems = Nothing
'   Set oFolder = Nothing
'   Set rsA = Nothing
'   Set fs = Nothing
'   Set oFile = Nothing
'
'   Exit Function
'
'ErrNo1:
'   'Resume
'   Screen.MousePointer = vbDefault
'   intErr2147024882 = ExportEMailErr(myItems, True, °Ó¼Ð³B¦¬¥ó«H½c, "(ErrNo1) " & strErrText & "; strSql=" & strSql, Err.Number, Err.Description, _
'                        strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
'   On Error GoTo 0: Err.Clear
'   If intErr2147024882 > 0 Then
'      Call WLog_Day("intErr2147024882 > 0", °Ó¼Ð³B¦¬¥ó«H½c)
'      'Resume Next
'      GoTo ReStarTM
'      Exit Function
'   End If
'
'   cmdCancel(3).Enabled = False
'   TmrTM.Interval = dblTmrTM: LblTM.BackColor = vbGreen
'
'   Set olApp = Nothing
'   Set myNamespace = Nothing
'   Set myFolder = Nothing
'   Set myItems = Nothing
'   Set oFolder = Nothing
'   Set rsA = Nothing
'   Set fs = Nothing
'   Set oFile = Nothing
'End Function

''Add By Sindy 2023/11/29
'Private Sub eventConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
'   m_SqlLogFile = App.path & "\cmdlog_" & Me.Name & "_" & strSrvDate(1) & ".log"
'   WCmdLog pCommand.CommandText
'End Sub
'Function WCmdLog(oStrLog As String)
'On Error GoTo ErrHnd
'
'Dim ffa As Integer
'ffa = FreeFile
'Open m_SqlLogFile For Append As ffa
'Print #ffa, Trim(Now) & "  ==>  " & oStrLog
'Close ffa
'
'ErrHnd:
'End Function
'Private Sub KillCmdLog()
'On Error GoTo ErrHnd
'   '§R°£«e¤@¤éªºLogÀÉ
'   Kill App.path & "\cmdlog_" & Me.Name & "_" & CompDate(2, -1, strSrvDate(1)) & ".log"
'ErrHnd:
'End Sub
''2023/11/29 END

'Add By Sindy 2024/5/14
Private Sub TmrLAbackup_Timer()
   Call ChkExecutionTimer(Left(LAbackup±H¥ó§X, 2))
End Sub

'Add By Sindy 2024/5/14
'´¼¼z©ÒÅU°ÝªA°È¶µ¥Ø¤åÀÉ°O¿ý
'¦^¶Ç:¬O§_¦¨¥\
Private Function LAbackupMail(ByVal strSubject As String, _
   ByVal strFullFileName As String, ByVal strFileName As String, _
   Optional ByRef strErrText As String, Optional ByRef intCaseOK As Integer, _
   Optional ByVal strRecipients As String) As Boolean

Dim objOutLook As Object
Dim objMail As Object
Dim strII17 As String, strII11 As String, strII12 As String, strII13 As String

Dim strText As String
Dim strUpdTime As String
Dim strCP14 As String, strCP13 As String, strCP12 As String, strCP64 As String
Dim strCP09 As String, strCP10 As String, stReName As String, strCP10Nm As String
Dim fs, f
Dim bolSaveEFile As Boolean
Dim bolConnect As Boolean
Dim strDirector As String
Dim strContent As String, strTo As String
Dim strBCC As String 'Add By Sindy 2024/7/8

On Error GoTo ErrHand

   LAbackupMail = False
   strErrText = ""
   Screen.MousePointer = vbHourglass

   Set objOutLook = CreateObject("Outlook.Application")
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set objMail = objOutLook.CreateItemFromTemplate(strFullFileName)

   strII17 = ChgSQL(objMail.Subject)
   TextII17 = objMail.Subject 'FindÂ²Åé¦r

   If objMail.Class = 46 Then '46.olReport
      strII11 = "¥¼¶Ç»¼ªº¥D¦®"
      strII12 = "0"
      strII13 = ""
   '43.olMail
   Else
      If objMail.SenderName = objMail.senderemailaddress Then
         strII11 = objMail.senderemailaddress
      Else
         strII11 = objMail.SenderName & " [" & objMail.senderemailaddress & "]"
      End If
      strII12 = Format(objMail.SentOn, "YYYYMMDD") 'ReceivedTime
      strII13 = Format(objMail.SentOn, "HHMMSS")
   End If
   '¥Î±H«H¤H¬d¬Ý¬O©Ò¤º¨º¤@¦ì­û¤uµoªº«H
   Call BySenderToStaff(strII11, strCP14, strDirector, True)
   '§ì¦¬¥óªÌ©Î°Æ¥»²Ä¤@¦ì¬°´¼Åv¤H­û
   Call BySenderToStaff(objMail.To, strCP13, strDirector, True)
   If strCP13 = "" Then
      Call BySenderToStaff(objMail.cc, strCP13, strDirector, True)
   End If
   If strCP13 <> "" Then strCP12 = GetST15(strCP13)
   
   '¸ÑªR¥D¦®:(®×¥ó©Ê½è) + ¶i«×³Æµù
   strText = strSubject
   strText = Replace(strText, "¡]", "(")
   strText = Replace(strText, "¡^", ")")
   strCP10 = ""
   If InStr(strText, "(") > 0 And InStr(strText, ")") > 0 Then
      strCP10 = Mid(strText, InStr(strText, "(") + 1, (InStr(strText, ")") - 1) - InStr(strText, "("))
      'ÀË¬d®×¥ó©Ê½è
      strCP10Nm = GetCaseTypeName("LA", strCP10, 0)
      If IsEmptyText(strCP10Nm) = True Then
         strErrText = strCP10 & "¦¹®×¥ó©Ê½è¥N¸¹¤£¦s¦b"
         strCP10 = "" 'Add By Sindy 2025/1/7
      End If
      If strCP10 = "0" Then
         strErrText = "®×¥ó©Ê½è¥N¸¹¤£¥i¬°¡Õ0.ÅU°Ý¸u¥ô¡Ö"
         strCP10 = "" 'Add By Sindy 2025/1/7
      End If
   End If
   strCP64 = Trim(Mid(strText, InStr(strText, ")") + 1))
   
   If strCP13 = "" Or strCP14 = "" Or strCP10 = "" Or strCP64 = "" Then
      strContent = "¸ê®Æ¤£¥þ:" & vbCrLf & _
                   "´¼Åv¤H­û: " & strCP13 & IIf(strCP13 = "", " (¤£¥iªÅ¥Õ)", "") & vbCrLf & _
                   "©Ó¿ì¤H: " & strCP14 & IIf(strCP14 = "", " (¤£¥iªÅ¥Õ)", "") & vbCrLf & _
                   "®×¥ó©Ê½è: " & strCP10 & IIf(strCP10 = "", " (¤£¥iªÅ¥Õ)", "") & vbCrLf & _
                   "¶i«×³Æµù: " & strCP64 & IIf(strCP64 = "", " (¤£¥iªÅ¥Õ)", "") & vbCrLf & vbCrLf & _
                   strErrText & vbCrLf & vbCrLf & _
                   "µLªk¦¬¿ý¡A½Ð­×¥¿«á¡A­«·s±H«H!!!"
      strBCC = ""
      If strCP14 = "" Then
         'Modify By Sindy 2024/7/8 µL©Ó¿ì¤H´N±Hµ¹lawoffice@taie.com.tw
         strTo = "lawoffice@taie.com.tw"
         strBCC = m_M51Recver
         '2024/7/8 END
      Else
         strTo = strCP14
      End If
      WLog_Day "==>LA-999999-0-00 : ·s¼W¶i«×¡i´¼¼z©ÒÅU°ÝªA°È¶µ¥Ø¤åÀÉ°O¿ý¡j¤º®e¦³»~!!! " & strFullFileName & " ==> " & vbCrLf & strContent, ªk«ß©Ò±H¥ó«H½c
      PUB_SendMail strUserNum, strTo, "", _
                   "·s¼W¶i«×¡i´¼¼z©ÒÅU°ÝªA°È¶µ¥Ø¤åÀÉ°O¿ý¡j¤º®e¦³»~!!!", strContent, , strFullFileName, , , , , , , , True, False, strBCC, , False, , , False
   Else
      cnnConnection.BeginTrans
      bolConnect = True
      strUpdTime = Right("000000" & ServerTime, 6)
      
      '¦¬¿ý¦Ü¶i«×ÀÉ
      strCP09 = AutoNo("B", 6)
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09" & _
               ",CP10,CP12,CP13,CP14,CP18,CP113,CP11,CP20,CP32,CP27,CP64)" & _
               " VALUES ('LA','999999','0','00'," & strSrvDate(1) & ",'" & strCP09 & "'" & _
               ",'" & strCP10 & "','" & strCP12 & "','" & strCP13 & "','" & strCP14 & "'" & _
               ",0,0.5,'04','N','N'," & strSrvDate(1) & ",'" & ChgSQL(strCP64) & "')"
      cnnConnection.Execute strSql
      '¦s¨÷©v°Ï
      stReName = PUB_CaseNo2FileName("LA", "999999", "0", "00") & _
                  "." & strCP10 & "." & strSrvDate(1) & strUpdTime & ".tx.msg"
      Set f = fs.GetFile(strFullFileName)
      WLog_Day "==>LA-999999-0-00 : ·s¼W¶i«× " & strCP09 & "(" & strCP10 & ") " & strFullFileName & " ==> " & stReName, ªk«ß©Ò±H¥ó«H½c
      
      bolSaveEFile = SaveAttFile_PDF(strCP09, strFullFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True, "F", "Y", , , , Me.TextII17.Text, strErrText, False)
      If bolSaveEFile = False Then
         WLog_Day "SaveAttFile_PDF ¥¢±Ñ: " & strErrText, ªk«ß©Ò±H¥ó«H½c
         If InStr(strErrText, strSubject) = 0 Then
            strErrText = strErrText & vbCrLf & _
                         strSubject & vbCrLf & _
                         "==>¦¬¨ì¤é´Á:" & strMailDate & " " & strMailTime & " ±H¥óªÌ:" & strSender & vbCrLf & _
                         "==>LA-999999-0-00 : " & strCP09 & "(" & strCP10 & ")" & strFullFileName & "==>" & stReName & vbCrLf
         End If
         PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "LA999999000-" & strCP09 & "­Ó®×¦s¨÷©v°Ï¥¢±Ñ¡A½Ð¬d¬Ý¡I", strErrText, , strFullFileName, , , , , , , , , False, , , False, , , False
         DoEvents
         '§R°£PCºÝÀÉ®×
         Call fs.DeleteFile(strFullFileName)
         DoEvents
         WLog_Day "[§R°£] GoTo ErrHand" & strFullFileName, ªk«ß©Ò±H¥ó«H½c
         GoTo ErrHand '¥¢±Ñµ²§ô
      End If
      intCaseOK = intCaseOK + 1 '°O¿ý­Ó®×µ§¼Æ
   End If
   '§R°£PCºÝÀÉ®×
   Call fs.DeleteFile(strFullFileName)
   DoEvents
   WLog_Day "[³B²z§¹¦¨, §R°£]" & strFullFileName, ªk«ß©Ò±H¥ó«H½c
   
   If bolConnect = True Then cnnConnection.CommitTrans
   bolConnect = False

   LAbackupMail = True
   Screen.MousePointer = vbDefault
   Set f = Nothing
   Set fs = Nothing

   Exit Function

ErrHand:
   Screen.MousePointer = vbDefault
   If bolConnect = True Then cnnConnection.RollbackTrans
   strErrText = strErrText & "LA±H¥ó³Æ¥÷¶×¤J¥¢±Ñ¡I" & vbCrLf & Err.Number & vbCrLf & Err.Description
   WLog_Day "[¥¢±Ñ LAbackupMail-ErrHand]" & strErrText, ªk«ß©Ò±H¥ó«H½c
   Set f = Nothing
   Set fs = Nothing
End Function

'Add By Sindy 2024/5/14
Private Function ChkExecutionTimer(strMailBox As String) As Boolean
Dim bolProRun As Boolean

On Error GoTo ErrNo1
   
   If cnnConnection.State = adStateClosed Then Exit Function '±ß¤WDBÂ_½u,¤£»Ý©¹¤U°õ¦æ
   '¥H§KTimer¦P®ÉRun°_¨Ó
   If LblFCPin.BackColor = vbBlue Then Exit Function
   If LblFCPout.BackColor = vbBlue Then Exit Function
   If LblPatent.BackColor = vbBlue Then Exit Function
   If LblTM.BackColor = vbBlue Then Exit Function
   If LblLAbackup.BackColor = vbBlue Then Exit Function 'Add By Sindy 2024/5/14
   
   Select Case strMailBox
      Case "01" '°ê¥~³¡IPDept¦¬«H¶l¥ó
         bolProRun = bolFCPinRun
      Case "02" '°ê¥~³¡IPDept±H«H¶l¥ó
         bolProRun = bolFCPoutRun
      Case "03" '±M§Q³BPatent¦¬«H¶l¥ó
         bolProRun = bolPatentRun
      Case "04" '°Ó¼Ð³BTM¦¬«H¶l¥ó
         bolProRun = bolTMRun
      Case "05" 'ªk«ß©Ò±H¥ó«H½c
         bolProRun = bolLAbackupRun
   End Select
   
   'ÀË¬d¬O§_­n°õ¦æTimer
   If ExecuteSchedule(strMailBox, "", "") = True Or bolProRun = True Then
      '¶}©l³B²zµ{¦¡,¥ý°±Timer
'      If strMailBox <> "01" Then TmrFCPin.Interval = 0: LblFCPin.BackColor = vbRed
'      If strMailBox <> "02" Then TmrFCPout.Interval = 0: LblFCPout.BackColor = vbRed
'      If strMailBox <> "03" Then TmrPatent.Interval = 0: LblPatent.BackColor = vbRed
'      If strMailBox <> "04" Then TmrTM.Interval = 0: LblTM.BackColor = vbRed
'      If strMailBox <> "05" Then TmrLAbackup.Interval = 0: LblLAbackup.BackColor = vbRed
      Call CloseMailTimer 'Modify By Sindy 2025/8/27
      
      Call MainImportPro(strMailBox, True)
      
      'Modify By Sindy 2024/12/20 ¦]¬O¾Þ§@¤â°Ê¶×¤J,¸ß°Ý­n²{¦b±Ò°ÊTimer¦Û°Ê¤À«H¶Ü
      If Command1.Tag = "¤â°Ê¶×¤J" _
         And UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '´ú¸Õ¸ê®Æ®w
         Call CloseMailTimer
         If MsgBox("­n²{¦b±Ò°ÊTimer¦Û°Ê¤À«H¶Ü¡H", vbExclamation + vbYesNo + vbDefaultButton2, "­«­n°T®§¡I") = vbYes Then
            Call StartMailTimer
            Command1.Tag = "" 'Add By Sindy 2024/12/20
         End If
      Else
         Command1.Tag = "" 'Add By Sindy 2024/12/20
      '2024/12/20 END
         'µ{¦¡³B²z§¹,±Ò°Ê¤W¦C°±ªºTimer
'         If strMailBox <> "01" Then TmrFCPin.Interval = dblTmrFCPin: LblFCPin.BackColor = vbGreen
'         If strMailBox <> "02" Then TmrFCPout.Interval = dblTmrFCPout: LblFCPout.BackColor = vbGreen
'         If strMailBox <> "03" Then TmrPatent.Interval = dblTmrPatent: LblPatent.BackColor = vbGreen
'         If strMailBox <> "04" Then TmrTM.Interval = dblTmrTM: LblTM.BackColor = vbGreen
'         'Add By Sindy 2024/5/14
'         If strSrvDate(1) >= ªk«ß©Ò¤À«H±Ò¥Î¤é Then
'            If strMailBox <> "05" Then TmrLAbackup.Interval = dblTmrLAbackup: LblLAbackup.BackColor = vbGreen
'         End If
'         '2024/5/14 END
         Call StartMailTimer 'Modify By Sindy 2025/8/27
      End If
   
   'Add By Sindy 2025/5/14 ¥[³t¤À«H
   ElseIf Val(strExecuTime_01) > 0 And strMailBox = "01" And _
      (txtPathIPDept.Tag <> "" And txtPathIPDeptOut.Tag <> "" And _
       txtPathPatent.Tag <> "" And txtPathTM.Tag <> "" And _
       txtPathLAbackup.Tag <> "") Then
      If Val(Format(Time, "HHMMSS")) >= Val(strExecuTime_01) Then
         Call CloseMailTimer 'Add By Sindy 2025/8/27
         
         Call MainImportPro(strMailBox, False)
         
         'µ{¦¡³B²z§¹,±Ò°Ê¤W¦C°±ªºTimer
'         If strMailBox <> "01" Then TmrFCPin.Interval = dblTmrFCPin: LblFCPin.BackColor = vbGreen
'         If strMailBox <> "02" Then TmrFCPout.Interval = dblTmrFCPout: LblFCPout.BackColor = vbGreen
'         If strMailBox <> "03" Then TmrPatent.Interval = dblTmrPatent: LblPatent.BackColor = vbGreen
'         If strMailBox <> "04" Then TmrTM.Interval = dblTmrTM: LblTM.BackColor = vbGreen
'         If strMailBox <> "05" Then TmrLAbackup.Interval = dblTmrLAbackup: LblLAbackup.BackColor = vbGreen
         Call StartMailTimer 'Modify By Sindy 2025/8/27
         
         '¹w³]¥ý²M°£¤À«Hªº.tag
         txtPathIPDept.Tag = ""
         txtPathIPDeptOut.Tag = ""
         txtPathPatent.Tag = ""
         txtPathTM.Tag = ""
         txtPathLAbackup.Tag = ""
      End If
   ElseIf Not (((Val(strSrvDate(2)) >= Val(txtIPDeptSDate) And Val(txtIPDeptSDate) > 0) And _
                (Val(strSrvDate(2)) <= Val(txtIPDeptEDate) And Val(txtIPDeptEDate) > 0)) And _
              Val(txtIPDeptMin) > 0) Then
      strExecuTime_01 = "" 'IPDept¥[³t¤À«H¥i°õ¦æªº®É¶¡
   '2025/5/14 END
   End If
   
   Exit Function
   
'Add By Sindy 2024/5/27
ErrNo1:
   If Err.Number <> 0 Then
      WLog Err.Number & " : " & Err.Description & vbCrLf
      '¤u§@¤Ñ¤~µomail
      If ChkWorkDay(strSrvDate(1)) = True Then
         PUB_SendMail strUserNum, m_M51Recver, "", _
            Err.Number & " : " & Err.Description & vbCrLf, "ÀË¬d " & UCase(Pub_GetSpecMan("¤À«H¥D¾÷¦WºÙ")) & " ¤À«H¬O§_¥¿±`!!!" & vbCrLf, , , , , , , , , , True, False, , , False, , , False
      End If
      If Err.Number = "ORA-03114" Then 'ORA-03114: ¥¼»P ORACLE ¬Û³s--2147217900
         tmrClock.Interval = 10000
         Call StartMailTimer 'Modify By Sindy 2024/12/20
'         TmrFCPin.Interval = dblTmrFCPin
'         TmrFCPout.Interval = dblTmrFCPout
'         TmrPatent.Interval = dblTmrPatent
'         TmrTM.Interval = dblTmrTM
'         TmrLAbackup.Interval = dblTmrLAbackup
      End If
   End If
End Function
'2024/5/14 END

'Modify By Sindy 2024/5/13 ¤À«H¥Dµ{¦¡
'strMailBox: ±ý¤À«Hªº«H½c
'Modify By Sindy 2025/5/14 +, bolSendNotic As Boolean: ¬O§_­nµo³qª¾«H
Private Sub MainImportPro(strMailBox As String, bolSendNotic As Boolean)
Dim jj As Integer
Dim strTo As String ', strCC As String, strTempCC As String
Dim oFileSys As New FileSystemObject, oFolder As Object
Dim strKind As String
'Dim myForward As Object
'Dim myNewEmail As Object 'Âà±H«H¥ó
Dim strCaseNo As String
Dim strIPMNoteSMIME As String '¥[±K¥D¦®
Dim bolReStar As Boolean
Dim strMRL01 As String, strMRL02 As String, strMRL03 As String, strMRL04 As String, strMRL05 As String
Dim rsA As New ADODB.Recordset
Dim strErrNumber As String 'Add By Sindy 2019/10/14
'Dim intURGENT As Integer 'Add By Sindy 2019/11/14
Dim bolRunIPDeptISDMail As Boolean 'Add By Sindy 2020/3/9
Dim strErrCode As String, strErrDesc As String 'Add By Sindy 2020/4/15
Dim strRecipients_1 As String, strRecipients_all As String '§ì¦¬¥óªÌ¸ê®Æ
'Add By Sindy 2023/6/26
Dim olApp As Object
Dim myNamespace As Object
Dim myFolder As Object
Dim myItems As Object
'2023/6/26 END
Dim strMailTime_Recv As String 'Add By Sindy 2023/7/13
Dim fs As Object, oFile As Object 'Add By Sindy 2023/9/13
Dim intFolder As Integer '­nÅª¨úªºFolder¼Æ; ex:Inbox ©M Junk Email

Dim otxtPath As TextBox
Dim bolProRun As Boolean, dblTmrInterval As Double
Dim oTmrPro As Timer, oLblPro As Label
Dim oCmdCancel As Object, oFrame As Frame
Dim strMailName As String
Dim bolExecution As Boolean

Dim bolForKeyWordDel As Boolean, ii As Integer
Dim strII01 As String, strII03 As String, strIR04 As String
   
On Error GoTo ErrNo1
   
   strErrText = "" 'Add By Sindy 2020/7/22
   
   Select Case strMailBox
      Case "01" '°ê¥~³¡IPDept¦¬«H¶l¥ó
         Set otxtPath = txtPathIPDept
         bolProRun = bolFCPinRun
         dblTmrInterval = dblTmrFCPin
         Set oTmrPro = TmrFCPin
         Set oLblPro = LblFCPin
         Set oCmdCancel = cmdCancel(0)
         Set oFrame = Frame1
         strMailName = °ê¥~³¡¦¬¥ó«H½c
      Case "02" '°ê¥~³¡IPDept±H«H¶l¥ó
         Set otxtPath = txtPathIPDeptOut
         bolProRun = bolFCPoutRun
         dblTmrInterval = dblTmrFCPout
         Set oTmrPro = TmrFCPout
         Set oLblPro = LblFCPout
         Set oCmdCancel = cmdCancel(1)
         Set oFrame = Frame2
         strMailName = °ê¥~³¡±H¥ó«H½c
      Case "03" '±M§Q³BPatent¦¬«H¶l¥ó
         Set otxtPath = txtPathPatent
         bolProRun = bolPatentRun
         dblTmrInterval = dblTmrPatent
         Set oTmrPro = TmrPatent
         Set oLblPro = LblPatent
         Set oCmdCancel = cmdCancel(2)
         Set oFrame = Frame3
         strMailName = ±M§Q³B¦¬¥ó«H½c
      Case "04" '°Ó¼Ð³BTM¦¬«H¶l¥ó
         Set otxtPath = txtPathTM
         bolProRun = bolTMRun
         dblTmrInterval = dblTmrTM
         Set oTmrPro = TmrTM
         Set oLblPro = LblTM
         Set oCmdCancel = cmdCancel(3)
         Set oFrame = Frame4
         strMailName = °Ó¼Ð³B¦¬¥ó«H½c
      'Add By Sindy 2024/5/14
      Case "05" '°ê¥~³¡IPDept±H«H¶l¥ó
         If ªk«ß©Ò¤À«H±Ò¥Î¤é > strSrvDate(1) Then Exit Sub
         Set otxtPath = txtPathLAbackup
         bolProRun = bolLAbackupRun
         dblTmrInterval = dblTmrLAbackup
         Set oTmrPro = TmrLAbackup
         Set oLblPro = LblLAbackup
         Set oCmdCancel = cmdCancel(4)
         Set oFrame = Frame5
         strMailName = ªk«ß©Ò±H¥ó«H½c
         '2024/5/14 END
   End Select
   Call PUB_WriteDebugLog("strMailBox=" & strMailBox & ";")  'Add By Sindy 2025/11/10
   
   If otxtPath = "" Then
      MsgBox "¦¬¥ó¸ê®Æ§¨¤£¥iªÅ¥Õ¡I"
      otxtPath.SetFocus
      Exit Sub
   End If
   If Dir(otxtPath, vbDirectory) = "" Then
      MkDir otxtPath
   End If
   
   strMRL01 = strMailBox: strMRL02 = "": strMRL03 = ""
strErrText = "InB-A:" 'Add By Sindy 2023/2/22 D-Bug
'   If ExecuteSchedule(strMRL01, strMRL02, strMRL03) = True Or bolProRun = True Then '­n°õ¦æTimer
'      'Add By Sindy 2023/11/29
'      Set eventConn = cnnConnection
'      KillCmdLog
'      '2023/11/29 END
      
      bolProRun = False
      If strMailBox = "01" Then
         bolFCPinRun = bolProRun
      ElseIf strMailBox = "02" Then
         bolFCPoutRun = bolProRun
      ElseIf strMailBox = "03" Then
         bolPatentRun = bolProRun
      ElseIf strMailBox = "04" Then
         bolTMRun = bolProRun
      ElseIf strMailBox = "05" Then
         bolLAbackupRun = bolProRun
      End If
      
strErrText = "InB-B:" 'Add By Sindy 2023/2/22 D-Bug
      Set olApp = CreateObject("Outlook.Application")
      Set myNamespace = olApp.GetNamespace("MAPI")
      intKeyCnt = 0: intRunOK = 0: intCaseOK = 0
      
strErrText = "InB-C:-2" 'Add By Sindy 2023/2/22 D-Bug
   'Add By Sindy 2024/1/31
   For intFolder = 1 To 1 '2
      'Modify By Sindy 2023/7/17
      If OpenOutLookFolder(myNamespace, myFolder, strMailBox, intFolder) = False Then
         Set olApp = Nothing
         Set myNamespace = Nothing
         Set myFolder = Nothing
         oTmrPro.Interval = 0
         oLblPro.BackColor = vbRed
         Exit Sub
      End If
      '2023/7/17 END
      
      bolReStar = False
      
ReStar:
      Set myItems = myFolder.Items
      strIPMNoteSMIME = "" '¥[±K¥D¦®
      intMaxItem = myItems.Count
      mail_ii = 0 'Add By Sindy 2024/7/29
      
      'Modify By Sindy 2024/4/29
      If Frame99.Tag <> "" Then
         strExc(10) = "¤w²¤¹LOutlook²§±`¡A¦ü¥G¤w¥¿±`¤À«H¡A½ÐÀË¬d¤À«Hª¬ªp¡I"
         PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "¡i¤w²¤¹LOutlook²§±`¡j", strExc(10) & vbCrLf, , , , , , , , , , , False, , , False, , , False
         WLog PUB_GetDbTerminal & "¡F" & strExc(10)
         Frame99.Tag = "" 'Add By Sindy 2024/4/27
      End If
      '2024/4/29 END
      
strErrText = "InB-F:" & "intMaxItem=" & intMaxItem 'Add By Sindy 2023/2/22 D-Bug
      '°O¿ýLogÀÉ
      'Modify By Sindy 2024/1/31 + And intFolder = 1
      If strMRL02 = "" And intFolder = 1 Then
         'Add By Sindy 2025/5/14
         If bolSendNotic = False Then '¥[³t¤À«H
            strMRL01 = strMRL01 & "A"
         End If
         '2025/5/14 END
         strMRL02 = strSrvDate(1)
         strMRL03 = Format(Right("000000" & ServerTime, 6), "00:00:00")
         strMRL05 = strUserNum
         'Add By Sindy 2025/8/27
         If strUserNum = "" Then
            strErrText = strErrText & vbCrLf & "strUserNum ³Q²M¦¨ªÅ¥Õ¤F!!"
            GoTo ErrNo1
         End If
         '2025/8/27 END
         strSql = "insert into MailReceiveLog(MRL01,MRL02,MRL03,MRL05,MRL09)" & _
                  "values('" & strMRL01 & "'," & strMRL02 & "," & Format(strMRL03, "hhmmss") & ",'" & strMRL05 & "','Y')"
         cnnConnection.Execute strSql
      End If
         
strErrText = "InB-G:" & "intMaxItem=" & intMaxItem 'Add By Sindy 2023/2/22 D-Bug
      If intMaxItem > 0 Then
         If bolUserControl = True Then
            frmpic002.Label1.Caption = "¶l¥ó±µ¦¬¤¤...½Ðµy­Ô..."
            frmpic002.Show
            frmpic002.ZOrder 0
            frmpic002.Label1.Font.Size = 12
            frmpic002.Label1.Font.Bold = True
         End If
         For mail_ii = myItems.Count To 1 Step -1
            Call PUB_WriteDebugLog("mail_ii=" & mail_ii & " myItems.Count=" & myItems.Count & " intMaxItem=" & intMaxItem & ";")  'Add By Sindy 2025/11/10
strErrText = "InB-H:" & "mail_ii=" & mail_ii & " : intMaxItem=" & intMaxItem   'Add By Sindy 2023/2/22 D-Bug
            oLblPro.BackColor = vbBlue 'ÂÅ¦âTimer¥¿¦bRun
            oCmdCancel.Enabled = True
            DoEvents
            If bolUserControl = True Then
               frmpic002.Label1.Caption = "¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii & "...½Ðµy­Ô~"
            Else
               oFrame.Caption = oFrame.Tag & "¡@¡@¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii
            End If
strErrText = "InB-I:" & "oFrame.Caption=" & oFrame.Caption 'Add By Sindy 2023/2/22 D-Bug
            DoEvents
            strErrText = ""
            intRunOK = intRunOK + 1 '°O¿ý±µ¦¬µ§¼Æ (2017/7/20¤~¶}©l°O¿ý¥þ³¡±µ¦¬ªºµ§¼Æ)
            strRecipients_1 = "": strRecipients_all = "" '§ì¦¬¥óªÌ¸ê®Æ
            If strMailBox = "01" Or strMailBox = "03" Then
               Call ReadMailText(myItems, True, strRecipients_all, strRecipients_1)
            Else
               Call ReadMailText(myItems, False)
            End If
            
            'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
            strErrText = strErrText & "²Ä " & mail_ii & " µ§ ¥D¦®: " & strSocSubject & vbCrLf
            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strSender: " & strSender & vbCrLf
            strErrText = strErrText & "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@strMailDateTime: " & strMailDate & " " & strMailTime
            Call WLog_Day(strErrText, strMailName)
            Call PUB_WriteDebugLog("strSocSubject=" & strSocSubject & ";")  'Add By Sindy 2025/11/10
            
            '·í±H¥ó¤H¦³­n¨DÅª¨ú¦^±ø®É¨t²Î·|µo«H
            '1.­nOutlook³]©w¤£¦^ÂÐÅª¨ú¦^±ø(¦ý«eÃD¬O«H¥ó¤]¥²¶·³]¬°¤w¶}±Ò)
            '2.­n³]©w¦Û°Ê²M°£¡¨§R°£ªº¶l¥ó¡¨
            '3.­n³]©w¥i¥H¸Ñ¶}ª÷Æ_«H¥ó:°òÂ¦ªº¦w¥þ©Ê¨t²Î§ä¤£¨ì±zªº¼Æ¦ì ID ¦WºÙ(-2146893792)
            'IPM.Note.SMIME ¥[±K
            'Modify By Sindy 2017/11/17
            'Modify By Sindy 2023/7/12 + Or myItems.Item(mail_ii).Class = 45 : ·s³qª¾ => UCase(myItems.Item(mail_ii).MessageClass) = UCase("IPM.Post")
            If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 _
               Or myItems.Item(mail_ii).Class = 45 Then
            'If myItems.Item(mail_ii).Class <> 43 Then
            '2017/11/17 END
               Call PUB_WriteDebugLog("[¥[±K¶l¥ó];")  'Add By Sindy 2025/11/10
               intKeyCnt = intKeyCnt + 1
               'Add By Sindy 2017/7/18 ¥[Log°O¿ý
               'strErrText = "²Ä " & mail_ii & " µ§ [¥[±K] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
               Call WLog_Day("[¥[±K¶l¥ó]" & vbCrLf, strMailName)
               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf '¥[±K¥D¦®
               '2017/7/18 END
            'Add By Sindy 2020/4/10 ¦^¦¬¶l¥ó,ª½±µ§R°£
            ElseIf InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
               Call PUB_WriteDebugLog("[¦^¦¬¶l¥ó];")  'Add By Sindy 2025/11/10
               intKeyCnt = intKeyCnt + 1
               'strErrText = "²Ä " & mail_ii & " µ§ [¦^¦¬] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
               Call WLog_Day("[¦^¦¬¶l¥ó]" & vbCrLf, strMailName)
               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
               'myItems.Item(mail_ii).Delete '§R°£ =>µLªk§R°£,·|·í
               'DoEvents
            'Add By Sindy 2019/9/23 [¥¼¶Ç»¼ªº¥D¦®] ¥D¦®: ¤wÅª¨ú: Certified AML & CFT Regulatory Compliance, Surveillance and Reporting Specialist; Taiwan
            'For Backup
            ElseIf myItems.Item(mail_ii).Class = 46 _
               And (strMailBox = "02" Or strMailBox = "05") Then 'REPORT.IPM.Note.IPNRN
               Call PUB_WriteDebugLog("[¥¼¶Ç»¼ªº¥D¦®] => §R°£;")  'Add By Sindy 2025/11/10
               intKeyCnt = intKeyCnt + 1
               'strErrText = "²Ä " & mail_ii & " µ§ [¥¼¶Ç»¼ªº¥D¦®] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
               If strMailBox = "02" Then
                  Call DeleteMyItems(myItems, strMailName, "[¥¼¶Ç»¼ªº¥D¦®] => §R°£") '§R°£Outlook¸Ì­±ªº¶l¥ó
               Else
                  PUB_SendMail strUserNum, m_M51Recver, "", _
                           "¡iLAbackup- myItems.Item(mail_ii).Class = 46 [¥¼¶Ç»¼ªº¥D¦®] check:¦]¤£·|µo¥Íªº±¡ªp¡j", strSocSubject & vbCrLf & vbCrLf & strSql, , , , , , , , , , True, False, , , False, , , False
               End If
            'Modify By Sindy 2018/5/30 IPM.RECALL.REPORT.FAILURE = Message Recall Failure.µLªk¦^¦¬
            'For Backup
            ElseIf InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.RECALL.REPORT.FAILURE")) > 0 _
               And (strMailBox = "02" Or strMailBox = "05") Then
               Call PUB_WriteDebugLog("[µLªk¦^¦¬¶l¥ó];")  'Add By Sindy 2025/11/10
               intKeyCnt = intKeyCnt + 1
               'Add By Sindy 2017/7/18 ¥[Log°O¿ý
               'strErrText = "²Ä " & mail_ii & " µ§ [µLªk¦^¦¬] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf
               Call WLog_Day("[µLªk¦^¦¬¶l¥ó]" & vbCrLf, strMailName)
               strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
               
               PUB_SendMail strUserNum, m_M51Recver, "", _
                           "¡i02 or 05- InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase(IPM.RECALL.REPORT.FAILURE)) > 0 [µLªk¦^¦¬¶l¥ó] check:¦]¤£·|µo¥Íªº±¡ªp¡j", strSocSubject & vbCrLf & vbCrLf & strSql, , , , , , , , , , True, False, , , False, , , False
            Else
'               strFileName = mail_ii & "." & _
'                             strSrvDate(1) & Right("000000" & ServerTime, 6) & ".msg"
               strFileName = strSrvDate(1) & Right("000000" & ServerTime, 6) & "." & mail_ii & ".msg"
               myItems.Item(mail_ii).SaveAs otxtPath & "\" & strFileName, 9 '9.Outlook Unicode¶l¥ó®æ¦¡.msg
               'Add By Sindy 2020/2/27 SaveAs¨ç¼Æ,´N·|±Ò°Ê°»´ú¯f¬r³nÅéªº¨¾¬r¾÷¨î¤F
               Sleep 1000
               DoEvents
               Call WLog_Day("²£¥Í¼È¦s¹q¤lÀÉ: " & otxtPath & "\" & strFileName, strMailName) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
               '2020/2/27 END
               Call PUB_WriteDebugLog("²£¥Í¼È¦s¹q¤lÀÉ: " & otxtPath & "\" & strFileName & ";")  'Add By Sindy 2025/11/10
               
'************************************************************
'*************** ­Ó§O«H½c¥t¥~­n³B²zªºµ{¦¡ *******************
               If strMailBox = "01" Then 'Inbound
                  'Add By Sindy 2022/2/22
                  '«H¥ó¦P®É¦³±Hipdept¤Îpatent«H½c®É,¤~ÀË¬d:
                  If InStr(UCase(strRecipients_all), UCase("patent@taie.")) > 0 And _
                     InStr(UCase(Replace(strRecipients_all, "80ipdept@taie.com.tw", "")), UCase("ipdept@taie.")) > 0 Then
                     '¥ý¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶i¨Ó¤F¡F­Y¦³¡A§R°£¡C­Y¨S¦³¡AÄ~Äò¡C
                     strSql = "select ii01,ii03 from ipdeptinput" & _
                              " where ii17 = '" & ChgSQL(strSocSubject) & "'" & _
                              " and ii11 = '" & ChgSQL(strSender) & "' and ii12 = " & DBDATE(strMailDate) & " and ii13 = " & Val(Replace(strMailTime, ":", "")) & _
                              " order by ii01 desc,ii03 desc"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        '«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]
                        intKeyCnt = intKeyCnt + 1
                        Call WLog_Day("[«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]", strMailName)
                        strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
                        Call DeleteMyItems(myItems, strMailName) '§R°£Outlook¸Ì­±ªº¶l¥ó
                        '§R°£PCºÝÀÉ®×
                        Set fs = CreateObject("Scripting.FileSystemObject")
                        Call fs.DeleteFile(otxtPath & "\" & strFileName)
                        Sleep 1000
                        DoEvents
                        GoTo IsReadNext 'Run¤U¤@µ§
                     Else
                        'ÀË¬d±M§Q³B¬O§_¦³¦¹µ§¸ê®Æ
                        strSql = "select pi01,pi03 from patentinput" & _
                                 " where pi17 = '" & ChgSQL(strSocSubject) & "'" & _
                                 " and pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate) & " and pi13 = " & Val(Replace(strMailTime, ":", "")) & _
                                 " order by pi01 desc,pi03 desc"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           'Add By Sindy 2024/5/27 ¤u§@¤Ñ¤~µomail
                           If ChkWorkDay(strSrvDate(1)) = True Then
                           '2024/5/27 END
                              '³oª¬ªp¬O¤£À³¸Óµo¥Íªº
                              PUB_SendMail strUserNum, "97038", "", _
                                 "¡iIPDept-¦¹µ§¶l¥ó±M§Q³B¤w¦¬¿ý(" & RsTemp.Fields("pi01") & "-" & RsTemp.Fields("pi03") & "),°ê¥~³¡¥¼¤@¨Ö¦¬¿ý,½ÐÀË¬dª¬ªp¡H(Ä~Äò©¹¤URun,¶i¦æ¶l¥ó¦¬¿ý...)¡j", strSocSubject & vbCrLf & vbCrLf & strSql, , otxtPath & "\" & strFileName, , , , , , , , True, False, , , False, , , False
                              'Ä~Äò©¹¤URun,¶i¦æ¶l¥ó¦¬¿ý...
                           End If
                        End If
                     End If
                  End If
                  '2022/2/22 END
'************************************************************
               ElseIf strMailBox = "02" Then 'Backup
                  'Add By Sindy 2022/6/27 ¨R¾P¦^«H
                  strExc(0) = "select ii01,ii03,ii28,ir04 from IPDeptinput,InputRecord" & _
                              " where Ii28 is not null" & _
                                " and Ii01=Ir01 and Ii03=Ir03 and Ir08=0" & _
                                " and instr('" & ChgSQL(myItems.Item(mail_ii).Subject) & "',Ii28)>0" & _
                                " and ir16='9'" '9.¦^«H
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strII01 = RsTemp.Fields("ii01")
                     strII03 = RsTemp.Fields("ii03")
                     strIR04 = RsTemp.Fields("ir04")
                     '¼W¥[³¡ªù§PÂ_
                     strExc(0) = "update InputRecord set ir08=" & strSrvDate(1) & ",ir09=" & Right("000000" & ServerTime, 6) & ",ir10='" & strUserNum & "'" & _
                                 " where ir01=" & strII01 & _
                                   " and ir03='" & strII03 & "'" & _
                                   " and upper(ir04)=upper('" & ChgSQL(strIR04) & "')" & _
                                   " and ir08=0"
                     cnnConnection.Execute strExc(0), intI
                     
                     '­Y«H¥ó¦¬¨üªÌ¥þ³¡¤w³B²z©Î¤w§R°£,¥DÀÉ´N¥i¥H±¾¤WmsgÀÉ§R°£¤é´Á,µ¥«ÝAutoBatchDay¤@­Ó¤ë«á§R°£¹êÅéÀÉ
                     strExc(0) = "select ir01 from InputRecord" & _
                                 " where ir01=" & strII01 & _
                                   " and ir03='" & strII03 & "'" & _
                                   " and ir08=0"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 0 Then '«H¥ó¦¬¨üªÌ¥þ³¡¤w³B²z©Î¤w§R°£
                        strExc(0) = "update IPDeptInput set" & _
                                    " ii16=" & strSrvDate(1) & _
                                    " where Ii01=" & strII01 & _
                                      " and Ii03='" & strII03 & "'" & _
                                      " and ii16=0"
                        cnnConnection.Execute strExc(0), intI
                     End If
                  End If
                  '2022/6/27 END
                  
                  'Modify By Sindy 2017/8/8
                  'ÀË¬d¦³³]©w¦¬¨üªÌ¬°²QµØªºÃöÁä¦r¤¤¨äºô°ì²Å¦X¦¹¶l¥ó¦¬¥óªÌ®É¡A«H¥óª½±µ§R°£¤£¶i¨t²Î
                  bolForKeyWordDel = False
                  'If InStr(ChgSQL(strSender), GetPrjSalesNM("86013")) > 0 Then
                     For ii = myItems.Item(mail_ii).Recipients.Count To 1 Step -1
   '                     strSql = "select LK01 from ipdeptkeyword" & _
   '                              " where LK12='F' and LK04='86013' and LK03='2'" & _
   '                              " and instr(upper('" & Replace(myItems.Item(mail_ii).Recipients(ii).address, "'", "") & "'),upper(LK01))>0"
   '                     intI = 1
   '                     Set rsA = ClsLawReadRstMsg(intI, strSql)
   '                     If intI = 1 Then
   '                        bolForKeyWordDel = True
   '                        Exit For
   '                     End If
                        strSql = "select LK01,LK12 from ipdeptkeyword" & _
                                 " where LK12='F' and LK04='86013' and LK03='2'" & _
                                 " and instr(upper('" & Replace(myItems.Item(mail_ii).Recipients(ii).Name, "'", "") & "'),upper(LK01))>0"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           'Add By Sindy 2024/5/17 °O¿ý¨Ï¥Î¦¸¼Æ
                           cnnConnection.Execute "update ipdeptkeyword set LK16=LK16+1" & _
                                                 " where LK01='" & rsA.Fields("LK01") & "' and LK12='" & rsA.Fields("LK12") & "'" _
                                                 , intI
                           '2024/5/17 END
                           bolForKeyWordDel = True
                           Exit For
                        End If
                     Next ii
                  'End If
                  If bolForKeyWordDel = True Then
                     Call DeleteMyItems(myItems, strMailName, "[§R°£] «H¥óª½±µ§R°£¤£¶i¨t²Î") '§R°£Outlook¸Ì­±ªº¶l¥ó
                     '§R°£PCºÝÀÉ®×
                     Set fs = CreateObject("Scripting.FileSystemObject")
                     Call fs.DeleteFile(otxtPath & "\" & strFileName)
                     Sleep 1000
                     DoEvents
                     GoTo IsReadNext 'Run¤U¤@µ§
                  End If
                  '2017/8/8 END
'************************************************************
               ElseIf strMailBox = "03" Then 'Patent
                  'Add By Sindy 2022/2/22
                  '«H¥ó¦P®É¦³±Hipdept¤Îpatent«H½c®É,¤~ÀË¬d:
                  If InStr(UCase(strRecipients_all), UCase("patent@taie.")) > 0 And _
                     InStr(UCase(Replace(strRecipients_all, "80ipdept@taie.com.tw", "")), UCase("ipdept@taie.")) > 0 Then
                     strMailTime_Recv = Format(myItems.Item(mail_ii).ReceivedTime, "HHMM") '¼W¥[§PÂ_ ReceivedTime ®É¶¡
                     '¥ý¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶i¨Ó¤F¡F­Y¦³¡A§R°£¡C­Y¨S¦³¡AÄ~Äò¡C
                     'Modify By Sindy 2022/10/26 µo¥Í¥D¦®¬OªÅ¥Õ,¦P®É±H2­Ó«H½c
                     If strSocSubject = "" Then
                        'Modify By Sindy 2023/7/13 ¼W¥[§PÂ_ strMailTime_Recv
                        strSql = "select pi01,pi03 from patentinput" & _
                                 " where pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate) & _
                                 " and (substr(lpad(pi13,6,0),1,4) = " & Format(strMailTime, "HHMM") & " or substr(lpad(pi13,6,0),1,4) = " & strMailTime_Recv & ")" & _
                                 " order by pi01 desc,pi03 desc"
                     Else
                     '2022/10/26 END
                        'Modify By Sindy 2023/7/13 ¼W¥[§PÂ_ strMailTime_Recv
                        strSql = "select pi01,pi03 from patentinput" & _
                                 " where pi17 = '" & ChgSQL(strSocSubject) & "'" & _
                                 " and pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate) & _
                                 " and (substr(lpad(pi13,6,0),1,4) = " & Format(strMailTime, "HHMM") & " or substr(lpad(pi13,6,0),1,4) = " & strMailTime_Recv & ")" & _
                                 " order by pi01 desc,pi03 desc"
                     End If
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        '«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]
                        intKeyCnt = intKeyCnt + 1
                        Call WLog_Day("[«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]", strMailName)
                        strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
                        Call DeleteMyItems(myItems, strMailName) '§R°£Outlook¸Ì­±ªº¶l¥ó
                        '§R°£PCºÝÀÉ®×
                        Set fs = CreateObject("Scripting.FileSystemObject")
                        Call fs.DeleteFile(otxtPath & "\" & strFileName)
                        Sleep 1000
                        DoEvents
                        GoTo IsReadNext 'Run¤U¤@µ§
                     Else
                        'ÀË¬d°ê¥~³¡¬O§_¦³¦¹µ§¸ê®Æ
                        'Modify By Sindy 2023/7/13 ¼W¥[§PÂ_ strMailTime_Recv
                        strSql = "select ii01,ii03 from ipdeptinput" & _
                                 " where ii17 = '" & ChgSQL(strSocSubject) & "'" & _
                                 " and ii11 = '" & ChgSQL(strSender) & "' and ii12 = " & DBDATE(strMailDate) & _
                                 " and (substr(lpad(ii13,6,0),1,4) = " & Format(strMailTime, "HHMM") & " or substr(lpad(ii13,6,0),1,4) = " & strMailTime_Recv & ")" & _
                                 " order by ii01 desc,ii03 desc"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           '³oª¬ªp¬O¤£À³¸Óµo¥Íªº
                           'Add By Sindy 2024/5/27 ¤u§@¤Ñ¤~µomail
                           If ChkWorkDay(strSrvDate(1)) = True Then
                           '2024/5/27 END
                              PUB_SendMail strUserNum, "97038", "", _
                                 "¡iPatent-¦¹µ§¶l¥ó°ê¥~³¡¤w¦¬¿ý(" & RsTemp.Fields("ii01") & "-" & RsTemp.Fields("ii03") & "),±M§Q³B¥¼¤@¨Ö¦¬¿ý,½ÐÀË¬dª¬ªp¡H(Ä~Äò©¹¤URun,¶i¦æ¶l¥ó¦¬¿ý...)¡j", strSocSubject & vbCrLf & vbCrLf & strSql, , otxtPath & "\" & strFileName, , , , , , , , True, False, , , False, , , False
                              'Ä~Äò©¹¤URun,¶i¦æ¶l¥ó¦¬¿ý...
                           End If
                        Else
                           '*****
                           'µ¥°ê¥~³¡«H½c¦¬¿ý¦¹µ§¬Û¦P¶l¥ó(²Î¤@¦¬¿ý)
                           '*****
                           
                           '°»´ú¬O§_¦³²§±`ªºª¬ªp,³qª¾¹q¸£¤¤¤ß
                           'ex:Invoice 222088 from Patentica Limited -  P-500/2RU -- CFP-025048
                           '¦³¬í®t,©Ò¥H±M§Q«H¥ó·|´Ý¯dµÛ,­nÃöª`
                           If DBDATE(strMailDate) < strSrvDate(1) Or _
                              (DBDATE(strMailDate) = strSrvDate(1) And (Val(Format(Time, "HH")) - Val(Format(strMailTime, "HH"))) > 1) Then
                              If bolReStar = True Then
                                 'Add By Sindy 2024/5/27 ¤u§@¤Ñ¤~µomail
                                 If ChkWorkDay(strSrvDate(1)) = True Then
                                 '2024/5/27 END
                                    PUB_SendMail strUserNum, "97038", "", _
                                       "¡iPatent-¦¹µ§¶l¥ó¦P®É¦³±Hipdept¤Îpatent«H½c,ÁÙ¥¼¶i¦æ¦¬¿ý,½ÐÀË¬dª¬ªp¡H(ÀË¬d¬O§_¦³¬í®t,©Ò¥H±M§Q«H¥ó·|´Ý¯dµÛ ©Î Patent«H½c¥ý±Ò°Ê¤F)¡j" & strSocSubject, strSocSubject & vbCrLf & vbCrLf & strSql, , otxtPath & "\" & strFileName, , , , , , , , True, False, , , False, , , False
                                 End If
                              End If
                           End If
                           
                           'Add By Sindy 2023/7/14 patent´«¤F¤½¥Î¸ê®Æ§¨,®É¶¡©Mipdept°t¤£°_¨Ó
                           'Print Format(myItems.Item(mail_ii).ReceivedTime, "HH:MM:SS")=16:49:28
                           'Print Format(myItems.Item(mail_ii).SentOn, "HH:MM:SS")=16:49:28
                           If strSocSubject <> "" Then
                              strSql = "select pi01,pi03 from patentinput" & _
                                       " where pi17 = '" & ChgSQL(strSocSubject) & "'" & _
                                       " and pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & DBDATE(strMailDate)
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                              If intI = 1 Then
                                 If RsTemp.RecordCount = 1 Then
                                    strSql = "select ii01,ii03 from ipdeptinput" & _
                                             " where ii17 = '" & ChgSQL(strSocSubject) & "'" & _
                                             " and ii11 = '" & ChgSQL(strSender) & "' and ii12 = " & DBDATE(strMailDate)
                                    intI = 1
                                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                    If intI = 1 Then
                                       If RsTemp.RecordCount = 1 Then
   '                                       PUB_SendMail strUserNum, "97038", "", _
   '                                          "(¤w§RÀÉ)¡iPatent-¦¹µ§¶l¥ó¦P®É¦³±Hipdept¤Îpatent«H½c,À³¸Ó¤w¦¬¿ý,¨Ï¥Î(«H½c¤À«H¬ö¿ý¬d¸ß)ÀË¬d¬O§_¦³¦¬¶iipdept¤Îpatent«H½c¡j" & strSocSubject, strSocSubject & vbCrLf & _
   '                                          "strMailTime_Recv = " & strMailTime_Recv & vbCrLf & vbCrLf & strSql, , otxtPath & "\" & strFileName, , , , , , , , True, False, , , False, , , False
                                          
                                          '«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]
                                          intKeyCnt = intKeyCnt + 1
                                          Call WLog_Day("[«H¥ó¦P®É±Hµ¹patent@taie.com.tw©Mipdept@taie.com.tw«á³B²z«H½cªº²Ä2«Ê«H¥óª½±µ§R°£]", strMailName)
                                          strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
                                          Call DeleteMyItems(myItems, strMailName) '§R°£Outlook¸Ì­±ªº¶l¥ó
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                           '2023/7/14 END
                           
                           '§R°£PCºÝÀÉ®×
                           Set fs = CreateObject("Scripting.FileSystemObject")
                           Call fs.DeleteFile(otxtPath & "\" & strFileName)
                           Sleep 1000
                           DoEvents
                           GoTo IsReadNext 'Run¤U¤@µ§
                        End If
                     End If
                  End If
                  '2022/2/22 END
               End If
'*************** ­Ó§O«H½c¥t¥~­n³B²zªºµ{¦¡ END ***************
'************************************************************

               If intErr2147024882 <> mail_ii Then
                  If strMailBox = "02" Or strMailBox = "05" Then Me.TxtIPDept = strFileName
                  
                  'Add By Sindy 2018/4/12
                  If Dir(otxtPath & "\" & strFileName) = "" Then
                     strErrText = "µL²£¥Í¹q¤lÀÉ,ºÃ¦ü¤¤¯f¬r " & "Err.Number:" & Err.Number & Err.Description & vbCrLf
                     Call ExportEMailErr(myItems, False, strMailName, strErrText, Err.Number, Err.Description, _
                           strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
                  'Add By Sindy 2020/4/14 ÀË¬d¹q¤lÀÉ¬O§_¥i¥H¥¿±`¶}±Ò
                  ElseIf ChkIsOpenEmail(otxtPath & "\" & strFileName, strErrCode, strErrDesc) = False Then
                     intKeyCnt = intKeyCnt + 1
                     strErrText = "²Ä " & mail_ii & " µ§ [MsgµLªk¶}±Ò] ¥D¦®: " & myItems.Item(mail_ii).Subject & vbCrLf & _
                        otxtPath & "\" & strFileName & vbCrLf & _
                        "Err.Number:" & strErrCode & strErrDesc & vbCrLf
                     Call WLog_Day(strErrText, strMailName)
                     strIPMNoteSMIME = strIPMNoteSMIME & strErrText & vbCrLf
                  Else
                  '2018/4/12 END
                     Sleep 100 'Add By Sindy 2019/12/13
                     
'*************** ­Ó§O«H½cªº¤À«H³W«hµ{¦¡ ***************
                     Select Case strMailBox
                        Case "01" '°ê¥~³¡IPDept¦¬«H¶l¥ó
                           'Add By Sindy 2018/7/10 °ê»Ú·|Ä³¶l¥ó -- (ª`·N:¥~¨Ó¶l¥ó¤@¼Ë­n¤À«H¥X¥h)
                           bolRunIPDeptISDMail = False
                           pub_SaveCoRec = False 'Add By Sindy 2022/6/17 °O¿ý¬O§_¦³Àx¦s©¹¨Ó°O¿ý
                           Call PUB_WriteDebugLog("01 PUB_IPDeptISDMail;")  'Add By Sindy 2025/11/10
                           If PUB_IPDeptISDMail(Me, "0", m_strISDPath, otxtPath, strFileName, intCaseOK) = True Then
                              Call WLog_Day("PUB_IPDeptISDMail => OK", strMailName) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
                              bolRunIPDeptISDMail = True
      '                        myItems.Item(mail_ii).Delete '§R°£
      '                        DoEvents
                           End If
                           '2018/7/10 END
                           Sleep 100 'Add By Sindy 2019/12/13
                           '¦s­ÓÀÉ®É¥D¦®¤£¥i¥H¦³\/:*?"<>|µ¥²Å¸¹
                           Call PUB_WriteDebugLog("01 PUB_IPDeptTransMail_New;")  'Add By Sindy 2025/11/10
                           bolExecution = PUB_IPDeptTransMail_New(Me, strTo, strErrText, strKind, strFileName, strCaseNo)
                        Case "02" '°ê¥~³¡IPDept±H«H¶l¥ó
                           'Add By Sindy 2018/7/10 °ê»Ú·|Ä³¶l¥ó
                           Call PUB_WriteDebugLog("02 PUB_IPDeptISDMail;")  'Add By Sindy 2025/11/10
                           If PUB_IPDeptISDMail(Me, "1", m_strISDPath, otxtPath, strFileName, intCaseOK) = True Then
                              Call DeleteMyItems(myItems, strMailName, "¤À«H¦¨¥\¡A§R°£¶l¥ó => PUB_IPDeptISDMail(©¹¨Ó°O¿ý)") '§R°£Outlook¸Ì­±ªº¶l¥ó
                              Sleep 100
                              GoTo IsReadNext 'Run¤U¤@µ§
                           Else
                           '2018/7/10 END
                              Sleep 100 'Add By Sindy 2019/12/13
                              '*****
                              '¦s­ÓÀÉ®É¥D¦®¤£¥i¥H¦³\/:*?"<>|µ¥²Å¸¹
                              'If IPDeptBackupMail(Me.TextII17.Text, otxtPath & "\" & strFileName, strFileName, strErrText, intCaseOK, strRecipients) = True Then
                              Call PUB_WriteDebugLog("02 IPDeptBackupMail;")  'Add By Sindy 2025/11/10
                              bolExecution = IPDeptBackupMail(Me.TextII17.Text, otxtPath & "\" & strFileName, strFileName, strErrText, intCaseOK)
                           End If
                        Case "03" '±M§Q³BPatent¦¬«H¶l¥ó
                           'Add By Sindy 2025/11/18
                           Call PUB_WriteDebugLog("03 PUB_IPDeptISDMail;")
                           If PUB_IPDeptISDMail(Me, "0", m_strISDPath, otxtPath, strFileName, intCaseOK) = True Then
                              Call WLog_Day("PUB_IPDeptISDMail => OK", strMailName)
                              bolRunIPDeptISDMail = True
                           End If
                           Sleep 100
                           '2025/11/18 END
                           Call PUB_WriteDebugLog("03 PUB_PatentTransMail;")  'Add By Sindy 2025/11/10
                           bolExecution = PUB_PatentTransMail(Me, strTo, strErrText, strKind, strFileName, strCaseNo)
                        Case "04" '°Ó¼Ð³BTM¦¬«H¶l¥ó
                           Call PUB_WriteDebugLog("04 PUB_TMTransMail;")  'Add By Sindy 2025/11/10
                           bolExecution = PUB_TMTransMail(Me, strTo, strErrText, strKind, strFileName, strCaseNo)
                        Case "05" 'ªk«ß©Ò±H¥ó«H½c
                           Call PUB_WriteDebugLog("05 LAbackupMail;")  'Add By Sindy 2025/11/10
                           bolExecution = LAbackupMail(Me.TextII17.Text, otxtPath & "\" & strFileName, strFileName, strErrText, intCaseOK)
                     End Select
'*************** ­Ó§O«H½cªº¤À«H³W«hµ{¦¡ END ***************
                     If bolExecution = True Then
                        Call PUB_WriteDebugLog("bolExecution = True;")  'Add By Sindy 2025/11/10
                        strExc(10) = ""
                        If strMailBox = "02" Then
                           strExc(10) = "IPDeptBackupMail ³B²z§¹²¦¡A§R°£¶l¥ó => IPDeptBackupMail"
                        ElseIf strMailBox = "05" Then
                           strExc(10) = "LAbackupMail ³B²z§¹²¦¡A§R°£¶l¥ó => LAbackupMail"
                        Else
                           'If strKind = "1" Then '­Ó®×
                           If strCaseNo <> "" Then '¦³Âk¨÷©v°Ï´Nºâ­Ó®×¥ó¼Æ Modify By Sindy 2017/7/21
                              intCaseOK = intCaseOK + 1
                           End If
                        End If
                        Call WLog_Day("bolExecution = True; (¥þ³¡«H¥ó / ³Ñ¾l¥ó¼Æ¡G" & intMaxItem & " / " & mail_ii & "); myItems.Count = " & myItems.Count, strMailName)
                        Call DeleteMyItems(myItems, strMailName, strExc(10)) '§R°£Outlook¸Ì­±ªº¶l¥ó
                        
                     Else
                        Call PUB_WriteDebugLog("bolExecution = False;")  'Add By Sindy 2025/11/10
                        'Add By Sindy 2020/3/9 ©¹¨Ó°O¿ý«H¥ó±H¥X, ¶Ç¦^=>¥¼¶Ç»¼ªº¥D¦®: Best wishes and update from Tai E regarding COVID-19 [Our Ref:Y53102000.B49] (EY/wc)
                        '  ©¹¨Ó°O¿ýªº¡¨¥¼¶Ç»¼ªº¥D¦®¡¨«H¥ó=>¬Oª½±µ§R°£¶l¥ó¹q¤lÀÉ,©Ò¥H¦b¦¹­n­ç°£,¤£µM·|³Q§PÂ_¬°¯f¬rÀÉ
                        If bolRunIPDeptISDMail = True _
                           And InStr(myItems.Item(mail_ii).Subject, "¥¼¶Ç»¼ªº¥D¦®") > 0 Then
                           Call PUB_WriteDebugLog("bolExecution = False; bolRunIPDeptISDMail (PUB_WriteDebugLog)")  'Add By Sindy 2025/11/10
                           Call DeleteMyItems(myItems, strMailName, "©¹¨Ó°O¿ýªº<¥¼¶Ç»¼ªº¥D¦®>«H¥ó => ª½±µ§R°£") '§R°£Outlook¸Ì­±ªº¶l¥ó
                           
                        Else
                        '2020/3/9 END
                           strErrNumber = Err.Number 'Add By Sindy 2019/10/14
                           Call PUB_WriteDebugLog("strErrNumber=" & Err.Number)  'Add By Sindy 2025/11/10
                           'Add By Sindy 2019/12/11
                           If InStr(strErrText, "§ä¤£¨ìÀÉ®×") > 0 Then
                              strErrText = "§ä¤£¨ìÀÉ®×,ºÃ¦ü¤¤¯f¬r"
   '                                 myItems.Item(mail_ii).Delete '§R°£
   '                                 DoEvents
                           End If
                           '2019/12/11 END
                           If strMailBox = "02" Or strMailBox = "05" Then
                              'Add By Sindy 2020/4/6
                              If Me.TextII17.Text <> "" Then
                                 If InStr(strErrText, Me.TextII17.Text) = 0 Then
                                    strErrText = strErrText & vbCrLf & Me.TextII17.Text & vbCrLf
                                 End If
                              End If
                              '2020/4/6 END
                           End If
                           
                           'Add By Sindy 2020/9/10
                           If strErrText <> "" And strErrText <> "Err.Number:0;" Then
                           Else
                           '2020/9/10 END
                              'Add By Sindy 2019/12/11
                              If strErrNumber = "0" Then
                                 strErrText = "§ä¤£¨ìÀÉ®×,ºÃ¦ü¤¤¯f¬r"
      '                           myItems.Item(mail_ii).Delete '§R°£
      '                           DoEvents
                              End If
                              '2019/12/11 END
                           End If
                           
                           Call ExportEMailErr(myItems, False, strMailName, strErrText, Err.Number, Err.Description, _
                              strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
                           'Add By Sindy 2019/10/14
                           'If strErrNumber = "999" Then
                           If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then
                              Exit For
                           End If
                           '2019/10/14 END
                        End If
                     End If
                  End If
               'Modify By Sindy 2020/4/15
               Else
                  intErr2147024882 = 0
               '2020/4/15 END
               End If
            End If
IsReadNext:
            '¬O§_­n¤¤Â_
            If bolCancel(Val(strMailBox) - 1) = True Then
               oLblPro.BackColor = vbRed
               DoEvents 'Add By Sindy 2024/5/7
               GoTo IsCancel
            End If
            Call PUB_WriteDebugLog("mail_ii=" & mail_ii & ";") 'Add By Sindy 2025/11/10
         Next mail_ii
         
IsCancel:
         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
         If bolUserControl = True Then
            Unload frmpic002
            Set frmpic002 = Nothing
         End If
         
         '°O¿ýLogÀÉ
         'Add By Sindy 2024/1/31
         If intFolder = 1 Then
         '2024/1/31 END
            '" and MRL05='" & strMRL05 & "'"
            strSql = "update MailReceiveLog set" & _
                     " MRL04=" & Format(strMRL04, "hhmmss") & _
                     ",MRL06=" & intRunOK & ",MRL07=" & intKeyCnt & ",MRL08=" & intCaseOK & _
                     ",MRL09='" & IIf(bolCancel(Val(strMailBox) - 1) = True, "B", "E") & "'" & _
                     " where MRL01='" & strMRL01 & "'" & _
                     " and MRL02=" & strMRL02 & _
                     " and MRL03=" & Format(strMRL03, "hhmmss")
            cnnConnection.Execute strSql
            
            Select Case strMailBox
               Case "01"
                  m_RunFCPinStarTime = Format(strMRL03, "hhmmss")
                  m_RunFCPinEndTime = Format(strMRL04, "hhmmss")
               Case "02"
                  m_RunFCPoutStarTime = Format(strMRL03, "hhmmss")
                  m_RunFCPoutEndTime = Format(strMRL04, "hhmmss")
               Case "03"
                  m_RunPatentStarTime = Format(strMRL03, "hhmmss")
                  m_RunPatentEndTime = Format(strMRL04, "hhmmss")
               Case "04"
                  m_RunTMStarTime = Format(strMRL03, "hhmmss")
                  m_RunTMEndTime = Format(strMRL04, "hhmmss")
               Case "05"
                  m_RunLAbackupStarTime = Format(strMRL03, "hhmmss")
                  m_RunLAbackupEndTime = Format(strMRL04, "hhmmss")
            End Select
         End If
         'Add By Sindy 2023/2/18
         If strErrNumber = "999" Or InStr(strErrText, "µLªk»PFTP Server«Ø¥ß³s½u") > 0 Then
            Err.Clear 'Add By Sindy 2025/10/13
            GoTo NotRunSec
         End If
         '2023/2/18 END
         'Add By Sindy 2017/8/8 °õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡A­Y¥u³Ñ¤U¥[±K¶l¥ó´Nµo«H³qª¾°ê¥~³¡¶l¥ó³B²z¤H­û
         '                      ¦³«D¥[±K¶l¥ó¦A°õ¦æ¤@¦¸±µ¦¬
'         DoEvents
         Set myItems = myFolder.Items
         intMaxItem = myItems.Count
         mail_ii = 0 'Add By Sindy 2024/7/29
         If intMaxItem > 0 Then
            strErrText = "": intKeyCnt = 0
            For mail_ii = myItems.Count To 1 Step -1
               Call ReadMailText(myItems, False)
               'Modify By Sindy 2017/11/17
               'Modify By Sindy 2020/4/10 + IPM.Outlook.Recall
               If InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Note.SMIME")) > 0 Or _
                  InStr(UCase(myItems.Item(mail_ii).MessageClass), UCase("IPM.Outlook.Recall")) > 0 Then
               'If myItems.Item(mail_ii).Class <> 43 Then
               '2017/11/17 END
                  'Modify By Sindy 2017/9/25
                  '¦³¥[±K«H¥ó¥B¬°¤u§@¤Ñ¤~­n±H«H³qª¾¤H­û³B²z
                  If ChkWorkDay(strSrvDate(1)) = True Then
                  '2017/9/25 END
                     If strErrText = "" Then
                        strErrText = "***¡@(" & IIf(strMailBox = "01", "inbound", IIf(strMailBox = "02", "backup", IIf(strMailBox = "03", "Patent", IIf(strMailBox = "04", "TM", "LAbackup")))) & _
                           ") °õ¦æ§¹¦AÀË¬d¤@¦¸¦¬¥ó§¨«H¥óª¬ªp¡@*********************************" & vbCrLf
                     End If
                     intKeyCnt = intKeyCnt + 1
                     strErrText = strErrText & "²Ä¡@" & mail_ii & "¡@µ§¡@[¥[±K]¡@¥D¦®:¡@" & strSocSubject & vbCrLf
                  End If
               Else
                  If bolReStar = False And bolCancel(Val(strMailBox) - 1) = False Then
                     bolReStar = True
                     Call WLog_Day("[­«Run²Ä¤G¦¸]" & vbCrLf, strMailName) 'Add By Sindy 2020/11/9 °O¿ý°õ¦æª¬ªpªºLog
                     '­«Run²Ä¤G¦¸
                     GoTo ReStar
                  'Add By Sindy 2022/8/5 ¤¤Â_´N¤£­n¦AÀË¬d¤F,©¹¤U°õ¦æ
                  ElseIf bolCancel(Val(strMailBox) - 1) = True Then
                     Exit For
                  '2022/8/5 END
                  End If
               End If
            Next mail_ii
            
            'Add By Sindy 2025/5/14
            If bolSendNotic = True Then '­nµo³qª¾«H
            '2025/5/14 END
               If strErrText <> "" Then
                  strErrText = strErrText & "*** END ************************************************************" & vbCrLf
                  Call WLog(strErrText)
                  '¦³¥[±K«H¥ó¥B¬°¤u§@¤Ñ¤~­n±H«H³qª¾¤H­û³B²z
                  If ChkWorkDay(strSrvDate(1)) = True And _
                     (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
                     '±HE-Mail³qª¾¦¬¥ó³B²z¤H­û
                     If UCase(pub_DbTerminalName) <> ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '´ú¸Õ¸ê®Æ®w
                        strTo = m_M51Recver
                     Else
                        strTo = PUB_TaRevMailTo(strMailBox)
                     End If
                     If strMailBox = "02" Then
                        PUB_SendMail strUserNum, m_M51Recver, "", °ê¥~³¡±H¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð¥ý¼Ð°O¬°¤wÅª¨ú¦A§R°£ª÷Æ_«H¥ó¡I(¹q¸£¤¤¤ßª½±µ§R°£¦¹«Ê«H¥ó,§Y¥i¡I)", strErrText & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                              "* Backup«H½cªº¥[±K¶l¥ó¥Ñ¹q¸£¤¤¤ß¤H­û¦Ü«H½c¤º§R°£" & vbCrLf & _
                              "  ¡A¥~±M¤H­û·|¦Û¦æ§â¥[±K«H¥ó¸Ñ±K«á¦A±H¤@¥÷¦ÜBackup«H½cÂk¨÷¥Î¡C" & _
                              "* ª`·N:¡]¥ý¼Ð°O¬°¤wÅª¨ú==>Á×§K¦^¶Ç¥¼Åª¨ú§Y§R°£ªº¦^±ø¡^¦A§R°£ª÷Æ_«H¥ó", , , , , , , , , , , False, , , False, , , False
                     ElseIf strMailBox = "05" Then
                        PUB_SendMail strUserNum, m_M51Recver, "", ªk«ß©Ò±H¥ó«H½c & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§(½Ð©M¨q¬Â½T»{¦¹ª¬ªp­n¦p¦ó³B²z¡I)", "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
                     Else
                        PUB_SendMail strUserNum, strTo, "", strMailName & "¦³ª÷Æ_«H¥ó " & intKeyCnt & " µ§¡A½Ð³B²z¡I", strIPMNoteSMIME & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                              "* ¶i¤J¨ä«H½c¸Ñ±K«áÂà±Hµ¹" & strMailName & "¡A¦A±N­ì¥[±K¶l¥ó§R°£¡AÁ×§K­«ÂÐ¡]¤Á°O¡^¡A«Ý¨t²Î¤U¦¸´`Àô³B²z¡C", , , , , , , , , , IIf(strTo = m_M51Recver, False, True), False, , , False, , , False
                     End If
   '                  DoEvents
                  End If
               End If
            End If
         End If
      End If 'Add By Sindy 2024/1/31
   Next intFolder 'Add By Sindy 2024/1/31
   
NotRunSec:
      Call PUB_SendMailCache 'Add By Sindy 2019/7/17
      If intRunOK > 0 Then 'Add By Sindy 2024/1/31
         'Modify By Sindy 2017/12/27 ¤u§@¤Ñ¤~­n³qª¾
         If ChkWorkDay(strSrvDate(1)) = True And _
            (Format(Time, "HHMMSS") >= "080000" And Format(Time, "HHMMSS") < "183000") Then
            'ÀË¬d¦¬¥ó¸ê®Æ§¨¤¤¬O§_¦³´Ý¯dÀÉ®×
            Set oFolder = oFileSys.GetFolder(otxtPath.Text)
            Set fs = CreateObject("Scripting.FileSystemObject")
            If oFolder.files.Count > 0 Then
               If strMailBox = "02" Or strMailBox = "05" Then
                  PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "±H¥ó¸ê®Æ§¨:" & otxtPath.Text & "©|¦³´Ý¯dÀÉ®×(" & oFolder.files.Count & "­Ó),½ÐÀË¬d¡I", "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
               Else
                  'Add By Sindy 2023/9/13
                  For Each oFile In oFolder.files
                     Set myItems = olApp.CreateItemFromTemplate(otxtPath.Text & "\" & oFile.Name)
                     Call ReadMailText_File(myItems)
                     '¬d¬Ý¦¹«Ê«H¥ó¡A¬O§_¤w¶×¤J?­Y¦³=§R°£¡C­Y¨S¦³=¤£³B²z,µ¥¤H­û¬d¬Ý
                     Select Case strMailBox
                        Case "01" '°ê¥~³¡IPDept¦¬«H¶l¥ó
                           strSql = "select ii01,ii03 from ipdeptinput" & _
                                    " where ii17 = '" & ChgSQL(strSocSubject) & "'" & _
                                    " and ii11 = '" & ChgSQL(strSender) & "' and ii12 = " & IIf(strMailDate <> "", DBDATE(strMailDate), "0") & " and ii13 = " & Val(Replace(strMailTime, ":", "")) & _
                                    " order by ii01 desc,ii03 desc"
                        Case "03" '±M§Q³BPatent¦¬«H¶l¥ó
                           strSql = "select pi01,pi03 from patentinput" & _
                                    " where pi17 = '" & ChgSQL(strSocSubject) & "'" & _
                                    " and pi11 = '" & ChgSQL(strSender) & "' and pi12 = " & IIf(strMailDate <> "", DBDATE(strMailDate), "0") & " and pi13 = " & Val(Replace(strMailTime, ":", "")) & _
                                    " order by pi01 desc,pi03 desc"
                        Case "04" '°Ó¼Ð³BTM¦¬«H¶l¥ó
                           strSql = "select ti01,ti03 from tminput" & _
                                    " where ti17 = '" & ChgSQL(strSocSubject) & "'" & _
                                    " and ti11 = '" & ChgSQL(strSender) & "' and ti12 = " & IIf(strMailDate <> "", DBDATE(strMailDate), "0") & " and ti13 = " & Val(Replace(strMailTime, ":", "")) & _
                                    " order by ti01 desc,ti03 desc"
                     End Select
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        '§R°£PCºÝÀÉ®×
                        Call fs.DeleteFile(otxtPath & "\" & oFile.Name)
                        Sleep 1000
                        DoEvents
                     End If
                  Next
                  Set oFolder = oFileSys.GetFolder(otxtPath.Text)
                  If oFolder.files.Count > 0 Then
                  '2023/9/13 END
                     PUB_SendMail strUserNum, m_M51Recver, "", PUB_GetDbTerminal & "¦¬¥ó¸ê®Æ§¨:" & otxtPath.Text & "©|¦³´Ý¯dÀÉ®×(" & oFolder.files.Count & "­Ó),½ÐÀË¬d¡I", "¦P¥D¦®", , , , , , , , , , , False, , , False, , , False
                  End If
               End If
            End If
'            'ÀË¬d¬O§_¦³«H¥ó¥¼Âà±H
'            If strMailBox <> "02" And strMailBox <> "05" Then '±Æ°£°ê¥~³¡IPDept±H«H¶l¥ó
'               'If UCase(pub_DbTerminalName) = ¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ Then '¥¿¦¡¸ê®Æ®w¤~µo«H
'                  strExc(0) = ""
'                  Select Case strMailBox
'                     Case "01" '°ê¥~³¡IPDept¦¬«H¶l¥ó
'                        strExc(0) = "SELECT COUNT(*) FROM ipdeptinput WHERE ii08=0"
'                     Case "03" '±M§Q³BPatent¦¬«H¶l¥ó
'                        'Modify By Sindy 2018/10/1 ¶®®S:¨ú®ø¦¹³qª¾
'                        'strExc(0) = "SELECT COUNT(*) FROM patentinput WHERE pi08=0"
'                     Case "04" '°Ó¼Ð³BTM¦¬«H¶l¥ó
'                        strExc(0) = "SELECT COUNT(*) FROM TMinput WHERE Ti08=0"
'                  End Select
'                  If strExc(0) <> "" Then
'                     intI = 1
'                     Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'                     If rsA.Fields(0) > 0 Then
'                        'Add By Sindy 2019/11/14 ¥D¦®¸Ì¦³ URGENT ¦r¼ËªÌ,³qª¾«H­n¥[¦³«æ¥ó! => IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") &
'                        intURGENT = 0
'                        strExc(0) = ""
'                        Select Case strMailBox
'                           Case "01" '°ê¥~³¡IPDept¦¬«H¶l¥ó
'                              strExc(0) = "SELECT COUNT(*) FROM ipdeptinput WHERE ii08=0 and instr(upper(ii17),'URGENT')>0"
'                           Case "04" '°Ó¼Ð³BTM¦¬«H¶l¥ó
'                              strExc(0) = "SELECT COUNT(*) FROM TMinput WHERE Ti08=0 and instr(upper(Ti17),'URGENT')>0"
'                        End Select
'                        If strExc(0) <> "" Then
'                           intI = 1
'                           Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'                           If rsA.Fields(0) > 0 Then
'                              intURGENT = rsA.RecordCount
'                           End If
'                           '2019/11/14 END
'                        End If
'                        'Modify By Sindy 2019/11/14 + IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") &
'                        PUB_SendMail strUserNum, strPTo, "", IIf(intURGENT > 0, "¡]¦³«æ¥ó¡I¡^", "") & "ª`·N¡G" & strMailName & "©|¦³¥¼Âà±H«H¥ó«Ý³B²z¡I", "¦P¥D¦®", , , , , , , , , , IIf(strMailBox = "01", False, True), False, , , False, , , False
'                     End If
'                  End If
'               'End If
'
'               If strMailBox = "01" Then
'                  'Modify By Sindy 2018/10/29 «H¥ó¦³¿ò¥¢,Âà±H¸ê°T¥¿±`,¦ý½T¹ê±H«H³Æ¥÷ºô­¶¨t²Î§ä¤£¨ì«H¥ó
'                  'select ii08,ii09,ii20,ii21,ii22,ii17 from ipdeptinput where ii01='20181025' and ii03 in('F0292','F0304','F0293','F0262');
'                  '/*
'                  '      II08       II09 II20                       II21       II22 II17
'                  '---------- ---------- -------------------- ---------- ---------- --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                  '  20181025     141308 Y                      20181025     141310 ¥¼¶Ç»¼ªº¥D¦®: Mail Delivery Failure
'                  '  20181026     143250 Y                      20181026     143256 Mail Delivery Failure
'                  '  20181026     143249 Y                      20181026     143255 IMPORTANT NOTICE
'                  '  20181026     143249 Y                      20181026     143254 Out of Office Notice
'                  '*/
'                  strExc(0) = "select count(*) from ipdeptinput where ii20<>'Y' and ii20 is not null" & _
'                              " and ii01>=20181001" & _
'                              " order by ii01,ii02"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If RsTemp.Fields(0) > 0 And ChkWorkDay(strSrvDate(1)) = True Then
'      '                  PUB_SendMail strUserNum, "97038", "", "¡iTaRevOutLook¡jÀË¬d«H¥ó¬O§_¦³¿ò¥¢(" & RsTemp.Fields(0) & "µ§)", strExc(0), , , , , , , , , , , False, , , False, , , False
'                     End If
'                  End If
'                  '2018/10/29 END
'               End If
'            End If
         End If
         
'         'Add By Sindy 2022/5/25
'         '±Hµo³qª¾«H
'         If m_strMailTo <> "" Then
'            '°Ï¤À³¡ªù
'            strF1xEmp = "": strF2xEmp = ""
'            varTmp = Split(m_strMailTo, ";")
'            For jj = 0 To UBound(varTmp)
'               If Left(PUB_GetST03(CStr(varTmp(jj))), 2) = "F1" Then '¥~°Ó
'                  strF1xEmp = strF1xEmp & ";" & varTmp(jj)
'               Else
'                  strF2xEmp = strF2xEmp & ";" & varTmp(jj)
'               End If
'            Next jj
'            'Call PUB_SendNotifyMail(m_strMailTo)
'            If strF1xEmp <> "" Then
'               strF1xEmp = Mid(strF1xEmp, 2)
'               Call PUB_SendNotifyMail(strF1xEmp)
'            End If
'            If strF2xEmp <> "" Then
'               strF2xEmp = Mid(strF2xEmp, 2)
'               Call PUB_SendNotifyMail(strF2xEmp)
'            End If
'         End If
         
      Else
         strMRL04 = Format(Right("000000" & ServerTime, 6), "00:00:00")
         '°O¿ýLogÀÉ
         strSql = "update MailReceiveLog set" & _
                  " MRL04=" & Format(strMRL04, "hhmmss") & _
                  ",MRL06=0,MRL07=0,MRL08=0" & _
                  ",MRL09='E'" & _
                  " where MRL01='" & strMRL01 & "'" & _
                  " and MRL02=" & strMRL02 & _
                  " and MRL03=" & Format(strMRL03, "hhmmss")
         cnnConnection.Execute strSql
         Select Case strMailBox
            Case "01"
               m_RunFCPinStarTime = Format(strMRL03, "hhmmss")
               m_RunFCPinEndTime = Format(strMRL04, "hhmmss")
            Case "02"
               m_RunFCPoutStarTime = Format(strMRL03, "hhmmss")
               m_RunFCPoutEndTime = Format(strMRL04, "hhmmss")
            Case "03"
               m_RunPatentStarTime = Format(strMRL03, "hhmmss")
               m_RunPatentEndTime = Format(strMRL04, "hhmmss")
            Case "04"
               m_RunTMStarTime = Format(strMRL03, "hhmmss")
               m_RunTMEndTime = Format(strMRL04, "hhmmss")
            Case "05"
               m_RunLAbackupStarTime = Format(strMRL03, "hhmmss")
               m_RunLAbackupEndTime = Format(strMRL04, "hhmmss")
         End Select
      End If
      'Modify By Sindy 2025/5/14
      Call TaRevOutLookBatchSendMail(strMailBox, bolSendNotic) '¾ã§åµo³qª¾«H
      '¼W¥[¥[³t¤À«H¥\¯à:
      'strMailBox=01 IPDept¤À«H§¹²¦«á,­pºâ¤U¤@­Ó¥i°õ¦æªº®É¶¡
      If strMailBox = "01" Then
         If ((Val(strSrvDate(2)) >= Val(txtIPDeptSDate) And Val(txtIPDeptSDate) > 0) And _
             (Val(strSrvDate(2)) <= Val(txtIPDeptEDate) And Val(txtIPDeptEDate) > 0)) And _
            Val(txtIPDeptMin) > 0 Then
            strExecuTime_01 = Format(DateAdd("n", Val(5), Format(Time, "hh:mm:ss")), "hhmmss")
         Else
            strExecuTime_01 = ""
         End If
      End If
      '2025/5/14 END
      
      txtMRL02 = strSrvDate(2)
      Call cmdQuery_Click
      oFrame.Caption = oFrame.Tag
      DoEvents
      
'      'Add By Sindy 2023/11/29
'      Set eventConn = Nothing
'      WCmdLog "MainImportPro µ²§ô"
'      WCmdLog ""
'      '2023/11/29 END
'   End If
   
   oCmdCancel.Enabled = False
   '­n¤¤Â_
   If bolCancel(Val(strMailBox) - 1) = True Then
      bolCancel(Val(strMailBox) - 1) = False
      oTmrPro.Interval = 0: oLblPro.BackColor = vbRed
   Else
   '¥¿±`µ²§ô
'      If oTmrPro.Interval > 0 Then
'         oTmrPro.Interval = dblTmrInterval
'         oLblPro.BackColor = vbGreen
'      Else
'         oLblPro.BackColor = vbRed
'      End If
      oTmrPro.Interval = dblTmrInterval: oLblPro.BackColor = vbGreen
   End If
      
   Set olApp = Nothing
   Set myNamespace = Nothing
   Set myFolder = Nothing
   Set myItems = Nothing
   Set oFolder = Nothing
   Set rsA = Nothing
   Set fs = Nothing
   Set oFile = Nothing
   
   Exit Sub
   
ErrNo1:
   'Resume
   Screen.MousePointer = vbDefault
   intErr2147024882 = ExportEMailErr(myItems, True, strMailName, "(ErrNo1) " & strErrText & "; strSql=" & strSql, Err.Number, Err.Description, _
                        strMRL01, strMRL02, strMRL03, strMRL04, strMRL05)
   On Error GoTo 0: Err.Clear
   If intErr2147024882 > 0 Then
      Call WLog_Day("intErr2147024882 > 0", strMailName)
      'Resume Next
      GoTo ReStar
      Exit Sub
   End If
   
   oCmdCancel.Enabled = False
   oTmrPro.Interval = dblTmrInterval: oLblPro.BackColor = vbGreen
   
   Set olApp = Nothing
   Set myNamespace = Nothing
   Set myFolder = Nothing
   Set myItems = Nothing
   Set oFolder = Nothing
   Set rsA = Nothing
   Set fs = Nothing
   Set oFile = Nothing
End Sub
