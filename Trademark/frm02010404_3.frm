VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010404_3 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "µù¥UÃÒ¿é¤J"
   ClientHeight    =   5748
   ClientLeft      =   5580
   ClientTop       =   1740
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6268.6
   ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
   ScaleWidth      =   9144
   Begin VB.Frame FrameTM20 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Caption         =   "Frame4"
      Height          =   255
      Left            =   4653
      TabIndex        =   72
      Top             =   2850
      Visible         =   0   'False
      Width           =   2715
      Begin VB.TextBox textTM20 
         Height          =   264
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   2
         Top             =   0
         Width           =   1092
      End
      Begin VB.Label Label16 
         Caption         =   "µoÃÒ¤é :"
         Height          =   255
         Left            =   390
         TabIndex        =   73
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Frame FrameTM14 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Caption         =   "Frame4"
      Height          =   255
      Left            =   3502
      TabIndex        =   70
      Top             =   2850
      Visible         =   0   'False
      Width           =   2715
      Begin VB.TextBox textTM14 
         Height          =   264
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   1
         Top             =   0
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "µù¥U¤½§i¤é :"
         Height          =   255
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   825
      Left            =   7510
      TabIndex        =   67
      Top             =   2880
      Width           =   1302
      Begin VB.OptionButton Option5 
         Caption         =   "¯È¥»ÃÒ®Ñ"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "¹q¤lÃÒ®Ñ"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   4110
      TabIndex        =   64
      Top             =   5190
      Width           =   4215
      Begin VB.TextBox Text12 
         Height          =   252
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   22
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   840
         MaxLength       =   2
         TabIndex        =   18
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   20
         Top             =   150
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "¤å¨ì          ¤Ñ"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "        ¤ë"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      ¤é"
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   21
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1230
      TabIndex        =   63
      Top             =   5190
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "¤å¨ì·í¤é"
         Height          =   180
         Index           =   0
         Left            =   144
         TabIndex        =   15
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "¤å¨ì¦¸¤é"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "°Ó«~¤ÎªA°È¸ê®Æ¬d¸ß(&I)"
      Height          =   350
      Index           =   6
      Left            =   4230
      TabIndex        =   23
      Top             =   64
      Width           =   1935
   End
   Begin VB.TextBox textNP09 
      Height          =   264
      Left            =   5730
      MaxLength       =   7
      TabIndex        =   14
      Top             =   4890
      Width           =   2292
   End
   Begin VB.TextBox textNP08 
      Height          =   264
      Left            =   1530
      MaxLength       =   7
      TabIndex        =   13
      Top             =   4890
      Width           =   2292
   End
   Begin VB.TextBox textCP47 
      Height          =   264
      Left            =   5730
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3180
      Width           =   1092
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1830
      Width           =   2532
   End
   Begin VB.TextBox textEditPrint 
      Height          =   264
      Left            =   5730
      MaxLength       =   1
      TabIndex        =   11
      Top             =   4230
      Width           =   372
   End
   Begin VB.TextBox textTM06 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1170
      Width           =   6492
   End
   Begin VB.TextBox textTC2 
      Height          =   264
      Left            =   6540
      MaxLength       =   20
      TabIndex        =   9
      Top             =   3870
      Width           =   2415
   End
   Begin VB.TextBox textMoney 
      Height          =   264
      Left            =   5730
      MaxLength       =   8
      TabIndex        =   7
      Top             =   3510
      Width           =   1502
   End
   Begin VB.TextBox textTC1 
      Height          =   264
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   8
      Top             =   3870
      Width           =   2532
   End
   Begin VB.TextBox textDate 
      Height          =   264
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   6
      Top             =   3510
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   510
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2532
   End
   Begin VB.TextBox textTM27 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2490
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2490
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      Height          =   270
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2820
      Width           =   2530
   End
   Begin VB.TextBox textTM21 
      Height          =   264
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3180
      Width           =   1092
   End
   Begin VB.TextBox textTM22 
      Height          =   264
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3180
      Width           =   1092
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   10
      Top             =   4230
      Width           =   372
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   367
      Left            =   8244
      TabIndex        =   26
      Top             =   64
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   367
      Index           =   0
      Left            =   6192
      TabIndex        =   24
      Top             =   64
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   367
      Left            =   7020
      TabIndex        =   25
      Top             =   64
      Width           =   1200
   End
   Begin MSForms.TextBox textTM07 
      Height          =   264
      Left            =   1500
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1500
      Width           =   6492
      VariousPropertyBits=   679493663
      Size            =   "11451;466"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   264
      Left            =   1500
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   840
      Width           =   6492
      VariousPropertyBits=   679493663
      Size            =   "11451;466"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5760
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   510
      Width           =   2532
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1260
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1830
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPS 
      Height          =   285
      Left            =   1260
      TabIndex        =   12
      Top             =   4560
      Width           =   7695
      VariousPropertyBits=   -1467989989
      MaxLength       =   128
      ScrollBars      =   2
      Size            =   "13568;501"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label32 
      Caption         =   "¨Ó¨ç´Á­­:"
      Height          =   255
      Left            =   150
      TabIndex        =   66
      Top             =   5370
      Width           =   855
   End
   Begin VB.Label LabNP07 
      Height          =   255
      Left            =   8370
      TabIndex        =   65
      Top             =   5340
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¤l®×·sªk©w´Á­­ :"
      Height          =   180
      Index           =   17
      Left            =   4410
      TabIndex        =   62
      Top             =   4920
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¤l®×·s¥»©Ò´Á­­ :"
      Height          =   180
      Index           =   18
      Left            =   150
      TabIndex        =   61
      Top             =   4920
      Width           =   1350
   End
   Begin VB.Label Label15 
      Caption         =   "»â¤g©µ¦ù´£¥Ó¤é :"
      Height          =   255
      Left            =   4320
      TabIndex        =   60
      Top             =   3180
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "¥Ó½Ð®×¸¹ :"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   59
      Top             =   1830
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "(Y:­×§ï)"
      Height          =   255
      Left            =   6240
      TabIndex        =   58
      Top             =   4230
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "¬O§_­×§ï©w½Z :"
      Height          =   255
      Left            =   4530
      TabIndex        =   57
      Top             =   4230
      Width           =   1245
   End
   Begin VB.Line Line1 
      X1              =   2458.773
      X2              =   2578.616
      Y1              =   3601.064
      Y2              =   3601.064
   End
   Begin VB.Label Label10 
      Caption         =   "®×¥ó¤é¤å¦WºÙ :"
      Height          =   255
      Left            =   180
      TabIndex        =   56
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "®×¥ó­^¤å¦WºÙ :"
      Height          =   255
      Left            =   180
      TabIndex        =   55
      Top             =   1170
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "®×¥ó¤¤¤å¦WºÙ :"
      Height          =   255
      Left            =   180
      TabIndex        =   54
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "¤j³°»âÃÒ¶O :"
      Height          =   255
      Left            =   4680
      TabIndex        =   50
      Top             =   3510
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "TCµù¥U¸¹¼Æ / ÃÒ®Ñ¸¹ :"
      Height          =   255
      Left            =   4680
      TabIndex        =   49
      Top             =   3870
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "TCµn°O¸¹ :"
      Height          =   255
      Left            =   180
      TabIndex        =   48
      Top             =   3870
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Ãº¦~¶O´Á­­ :"
      Height          =   255
      Left            =   180
      TabIndex        =   47
      Top             =   3510
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "´¼Åv¤H­û :"
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   46
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "¥»©Ò®×¸¹ :"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   44
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "¥Ó½Ð¤H :"
      Height          =   255
      Left            =   180
      TabIndex        =   43
      Top             =   1830
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "°Ó¼ÐºØÃþ :"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   42
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "¥¿°Ó¼Ð¸¹¼Æ :"
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   41
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "°Ó«~Ãþ§O :"
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   40
      Top             =   2490
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "¨Ó¨ç¦¬¤å¤é :"
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   39
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "µù¥U¸¹ :"
      Height          =   255
      Left            =   180
      TabIndex        =   38
      Top             =   2820
      Width           =   1005
   End
   Begin VB.Label Label14 
      Caption         =   "±M¥Î´Á­­ :"
      Height          =   255
      Left            =   180
      TabIndex        =   37
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "¦C¦L³Æµù :"
      Height          =   255
      Left            =   180
      TabIndex        =   36
      Top             =   4590
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "¦C¦L©w½Z :"
      Height          =   255
      Left            =   180
      TabIndex        =   35
      Top             =   4230
      Width           =   975
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:¤£¦L;1:¥x->¦U°ê;2:¥~->¥x;3:­^¤å)"
      Height          =   180
      Left            =   1680
      TabIndex        =   34
      Top             =   4230
      Width           =   2745
   End
End
Attribute VB_Name = "frm02010404_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0¤w­×§ï textTM05/textTM07/textTm23/textCP13/textPS
'Memo By Sindy 2012/12/3 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/16 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/26 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo By Sindy 2010/8/5 ¤é´ÁÄæ¤w­×§ï
'2005/8/31¾ã²z
Option Explicit

' ¥»©Ò®×¸¹
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' °Ó¼ÐºØÃþ
Dim m_TM08 As String
' ¥Ó½Ð°ê®a
Dim m_TM10 As String
' ¥Ó½Ð¤é
Dim m_TM11 As String
' ¤½§i¤é
Dim m_TM14 As String
Dim m_FinalDate As String 'Add By Sindy 2020/12/14 ©w½Z¤é´Á
' ±M¥Î´Á­­°_¤é
Dim m_TM21 As String
' ±M¥Î´Á­­¤î¤é
Dim m_TM22 As String
' ¥Ó½Ð¤H¥N¸¹
Dim m_TM23 As String
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
' ¥¿°Ó¼Ð¸¹¼Æ
Dim m_TM27 As String
' §@«~ºØÃþ
Dim m_SP46 As String
' ¨Ó¨ç¦¬¤å¤é
Dim m_CP05 As String
' ¾÷Ãö¤å¸¹
Dim m_CP08 As String
' ©Ò¿ï¨úªº¦¬¤å¸¹
Dim m_CP09 As String
' ®×¥ó©Ê½è
Dim m_CP10 As String
' ´¼Åv¤H­û
Dim m_CP13 As String
Dim m_CP12 As String
' ¨Ó·½µe­±
Dim strPrevForm As String
' ·s¼Wªº¦¬¤å¸¹
Dim strCP09 As String
Dim strNP22 As String 'Modify By Sindy 2009/10/23
Dim strNP08 As String 'Modify By Sindy 2009/10/23
'Add By Cheng 2002/06/12
Dim m_SP51 As String
'Add By Cheng 2003/12/09
Dim m_blnReceiveSecond As Boolean '§PÂ_¬O§_¦¬²Ä¤G´Áµù¥U¶O
'Add By Cheng 2004/02/06
Dim m_blnNoResult As Boolean '§PÂ_´¿³QÄ³ªº®×¥ó¬O§_µLµ²ªG
'End
'2005/11/11 ADD BY SONIA
Dim m_strLanguage As String '©w½Z»y¤å
'add by nickc 2006/06/07
Dim isRuned As Boolean
Dim Is717end As Boolean
Dim Is715end As Boolean
'add by nickc 2006/08/04
Public UpForm As Form
Dim m_MonTM01 As String     '¬ö¿ý¤À³Î¥À®×®×¸¹
Dim m_MonTM02 As String
Dim m_MonTM03 As String
Dim m_MonTM04 As String
Public m_MonCP09 As String  '¶Ç¤J¤À³Î¥À®×¦¬¤å¸¹
Dim m_MonNP08 As String
Dim m_MonNP09 As String
Dim strCP05 As String
Dim ii As Integer
Dim rsTmp As New ADODB.Recordset
'92.04.16 nick ¬ö¿ý§@¥Î«öÁä
Public cmdState As Integer
'add by nick 2004/10/05 ÀË¬d¬O§_¤w¸g¦³°Ó«~¤ÎªA°È
Public ChkTG As Boolean
Dim strRvType As String 'Add By Sindy 2012/5/18
Dim m_TM13 As String 'Add By Sindy 2012/12/19 ¼f©w¨Ó¨ç¤é
'Add By Sindy 2013/5/3
Dim m_TM67 As String '©ñ±ó±M¥ÎÅv
Dim m_TM118 As String '¦P·N®Ñ°Ó¼Ð¸¹¼Æ
'2013/5/3 End
'Added by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
Dim m_TM44 As String
Dim bolA1kdataMail As Boolean 'µo¶Ê´Ú¨ç(Outlook)
Dim m_ULD02 As String   '§ó·s©w½Z¤é´Á
'Modified by Lydia 2017/04/06 ½Ð´Ú³æªº½Ð´Ú¹ï¶H,¥i¯à©M¥N²z¤£¤@­P,§ï³]ÅÜ¼Æ
'Dim m_AC2470 As String  '©w½Z¥[¦L¶Ê´Ú³æPDF
Dim m_rA1k28 As String  '½Ð´Ú³æªº½Ð´Ú¹ï¶H
Dim m_rSpec As String  '¯S©w¥N²z¤Hªºmail¤º¤å¤£¦P
'end 2017/04/06
Dim strNCP09 As String   '·s¼WªºCÃþ¦¬¤å¸¹
Dim strNcp10 As String   '·s¼WªºCÃþ¦¬¤å¸¹®×¥ó©Ê½è
Dim str1006CP64 As String 'Added by Lydia 2017/02/02 ¥x-¤j­«µoµù¥Uµý,1006³¡¤À³Ó³¡¤À±Ñªº¶i«×³Æµù(ex.T-165417)
'Added by Morgan 2017/6/14 ¹q¤l¤½¤å
Public m_DocWord As String
Public m_DocNo As String
Public m_DocPdf As String
Public m_DocPdfDate As String
Public m_DocPdfTime As String
'end 2017/6/14
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END
Dim m_NA85 As String 'Added by Lydia 2019/11/13 ­pºâ°Ó¼Ð±M¥Î´Á¬O§_´î1¤Ñ
Dim strLD18 As String 'Add By Sindy 2019/12/19 «H¨çÁ`¦¬¤å¸¹
Dim m_TM136 As String 'Added by Morgan 2025/2/18 µù¥UÃÒ§Î¦¡

'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
'add by nickc 2008/01/23 ¥[¤J¥i¥H¨ú®ø
If UpForm Is Nothing Or Me.Visible = False Then
   Select Case strPrevForm
      Case "2"
         frm02010404_2.Show
         Unload Me
      Case Else
         frm02010404_1.Show
         Unload Me
         Unload frm02010404_2
   End Select
Else
    'add by nickc 2008/01/23 ¥[¤J¥i¥H¨ú®ø
    If UpForm Is frm02010401_6 Then
        frm02010401_6.m_IsCancal = True
        Unload Me
    End If
End If
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
    '¦C¦L±µ¬¢±µ®×³æ
'move to unload by nick
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '§R°£¼È¦s¸ê®Æ
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010404_2
   Unload frm02010404_1
   Unload Me
End Sub

Public Sub cmdok_Click(Index As Integer)
'92.04.16 nick ¬ö¿ý§@¥Î«öÁä
cmdState = Index
PubShowNextData
Exit Sub
End Sub

'Add By Sindy 2009/05/14
Public Sub PubShowNextData()
Dim strTit As String
Dim strMsg As String
Dim nResponse

Select Case cmdState
Case 0
   cmdOK(0).Enabled = False  'add by sonia 2019/2/1 ¦³­«ÂÐ°õ¦æªº±¡§Î
   'Add by Morgan 2003/11/21
   Call CaseCheck(m_TM01, m_TM02, m_TM03, m_TM04, m_TM10)
   '---end
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
      If TxtValidate = False Then
         cmdOK(0).Enabled = True  'add by sonia 2019/2/1 ¦³­«ÂÐ°õ¦æªº±¡§Î
         Exit Sub
      End If
        'add by nickc 2006/08/04
        If UpForm Is Nothing Or Me.Visible = False Then
            'add by nickc 2005/04/22
            '2011/11/8 modify by sonia TF¤l®×¤£¥iµ²¾l¬G¥[¶Ç¥»©Ò®×¸¹
            'Pub_EndModCashMsg m_TM10
            Pub_EndModCashMsg m_TM10, m_TM01, m_TM02, m_TM03, m_TM04
            
          ' ³]©w·Æ¹«´å¼Ð¬°µ¥«Ýª¬ºA
          Screen.MousePointer = vbHourglass
          ' Àx¦s¸ê®Æ
            'Modify By Cheng 2002/11/07
    '      'OnSaveData
            If OnSaveData = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            'Add By Cheng 2002/11/08
            ' ¦C¦L©w½Z
            If textPrint <> "N" Then
               'add by nickc 2006/06/07
               If Is717end = True Then m_blnReceiveSecond = True
               PrintLetter
            End If
          ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
          Screen.MousePointer = vbDefault
        End If
        
        'Added by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
        '¬°¤F©µ½w¥X©w½Z,§ó·s©w½Z¤é´Á
        If m_ULD02 <> "" Then
           'Modified by Lydia 2017/04/24 §ï¦¨Function
           'Call PUB_UpdateET07LD0216("1", strNCP09, m_TM01, m_TM02, m_TM03, m_TM04, "05", m_ULD02)
           If PUB_UpdateET07LD0216("1", strNCP09, m_TM01, m_TM02, m_TM03, m_TM04, "05", m_ULD02) = False Then
           End If
           'end 2017/04/24
        End If
        'µo¶Ê´Ú¨ç
        If bolA1kdataMail = True Then
           'Modified by Lydia 2017/02/18 ¹w³]³£ªþ¶Ê´Ú,¨Ã°Ï¤À¬O§_¬°¯S©w«È¤á(±H¯È¥»)
           'Call PUB_SendA1kdataMail(Me, m_TM01, m_TM02, m_TM03, m_TM04, strNCP09, strNcp10, m_AC2470)
           'Modified by Lydia 2017/04/06 °Ï¤À½Ð´Ú¹ï¶H
           'Call PUB_SendA1kdataMail(Me, m_TM01, m_TM02, m_TM03, m_TM04, strNCP09, strNcp10, m_TM44, IIf(m_AC2470 <> "", "Y", "N"))
           'Added by Lydia 2017/11/01 ¦]¬°¶l¥ó¹w³]¦¬¥ó¤H¬°°ò¥»ÀÉ¤§¥N²z¤H,­Y¤í´Ú¤§¹ï¶H»PTM44¤£¦P®É,¼u°T®§´£¿ô§Y¥i
                                     'ex. T-156008²{¦bTM44=Y5338100,106/10/24 ®Ö­ã-©µ®iCA6066488,§PÂ_¦P®×¥ó98¦~¦³Y51318000ªº¤í´Ú(¶Ê´Ú³æªº½Ð´Ú¹ï¶H),©Ò¥H²£¥ÍDÃþ¦¬´Ú±HÃÒ©MµoMAIL; µoMAIL®M¥Î¼Ò²Õ¹w³]§ìTM44¬°¦¬¥ó¤H,µM«áµo«HY5338100³y¦¨¹ï¤èªººÃ°Ý¡C
           If m_rA1k28 <> m_TM44 Then
             MsgBox "¤í´Ú½Ð´Ú³æ¤§½Ð´Ú¹ï¶H»P²{¦bFC¥N²z¤H¤£¦P, ½Ð¦Û¦æª`·N±ý¶Ê´Ú¹ï¶H¡I¡I", vbCritical, "¦¬´Ú±HÃÒ"
           End If
           'end 2017/11/01
           Call PUB_SendA1kdataMail(Me, m_TM01, m_TM02, m_TM03, m_TM04, strNCP09, strNcp10, m_rA1k28, m_rSpec)
        End If
        'end 2016/12/22
        
      If UpForm Is Nothing Then
         Unload frm02010404_2
         'Add By Sindy 2019/5/10
         If Me.m_strIR01 <> "" Then
           Unload frm02010404_1
           If Not m_PrevForm Is Nothing Then
              Call m_PrevForm.GoNext
           End If
           
         'Modified by Morgan 2023/1/17 «D¹q¤l¤½¤å¤~¦^«eµe­±
         'Else
         ElseIf m_DocNo = "" Then
         'end 2023/1/17
         '2019/5/10 END
           'add by nick 2004/10/20
            frm02010404_1.m_TM14 = textTM14.Text
'            frm02010404_1.m_FinalDate = textFinalDate.Text 'Add By Sindy 2020/12/14
            frm02010404_1.Show
         End If
       ElseIf UpForm Is frm02010401_6 Then
          '­Y¬Oµe­±¦³¥X²{¥i¥H¿é¸ê®Æ¡A­n±N¸ê®Æ¥á¦^«e­±¦s
          If Me.Visible = True Then
            frm02010401_6.PutSeekData01 = textTM15
            frm02010401_6.PutSeekData02 = textTM14
            'frm02010401_6.PutSeekData03 = Text1
            frm02010401_6.PutSeekData04 = textTM21
            frm02010401_6.PutSeekData05 = textTM22
            frm02010401_6.PutSeekData06 = textCP47
            frm02010401_6.PutSeekData07 = textDate
            frm02010401_6.PutSeekData08 = textMoney
            frm02010401_6.PutSeekData09 = textTC1
            frm02010401_6.PutSeekData10 = textTC2
            frm02010401_6.PutSeekData11 = textPrint
            frm02010401_6.PutSeekData12 = textPS
            frm02010401_6.PutSeekData13 = textNP08
            frm02010401_6.PutSeekData14 = textNP09
          End If
       End If
              
       'Modified by Morgan 2023/1/17 ¹q¤l¤½¤å
       'Unload Me
       If m_DocNo <> "" Then
         frm02010412.m_TM14 = textTM14.Text 'Added by Morgan 2023/6/15
         Unload Me
         Unload frm02010404_1
         frm02010412.GoNext
       Else
         Unload Me
       End If
       'end 2023/1/17
       
   'add by sonia 2019/2/1 ¦³­«ÂÐ°õ¦æªº±¡§Î
   Else
      cmdOK(0).Enabled = True
   'end 2019/2/1
   End If
   
'add by nick 2004/10/05
Case 6
    'frm03010303_04.Hide 'Modify By Sindy 2009/09/17
    Set frm03010303_04.UpForm = Me
    frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 'textTMKey 'lbl1(0).Caption
    frm03010303_04.AllClass = textTM09 'txt1(0).Text
    frm03010303_04.cmdOK(0).Visible = False
    frm03010303_04.cmd.Visible = False
    frm03010303_04.cmd2.Visible = False
    frm03010303_04.txt2(0).Visible = False
    frm03010303_04.Line1.Visible = False
    frm03010303_04.txt2(1).Visible = False
    frm03010303_04.txt2(2).Visible = False
    frm03010303_04.txt2(3).Visible = False
    frm03010303_04.Caption = "°Ó«~¤ÎªA°È¸ê®Æ"
    'edit by nickc 2008/02/12 §ï¦¨¥i¥H½Æ»s
    'frm03010303_04.TXT1(0).Enabled = False
    'frm03010303_04.TXT1(1).Enabled = False
    'frm03010303_04.TXT1(2).Enabled = False
    frm03010303_04.txt1(0).Locked = True
    frm03010303_04.txt1(1).Locked = True
    frm03010303_04.txt1(2).Locked = True
    frm03010303_04.Label2.Visible = False
    'Add By Sindy 2024/12/20
    frm03010303_04.PubMsg = "¤ñ¹ï¤½³ø°Ó«~¸ê®Æ"
    frm03010303_04.m_TM08 = m_TM08
    frm03010303_04.m_TM15 = textTM15
    '2024/12/20 END
    'Me.Hide 'Modify By Sindy 2009/09/17
    frm03010303_04.QueryData
    frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 §ï¬°±j¨î¦^À³ªí³æ
End Select
End Sub

Private Sub Form_Load()
   ' ³]©w±±¨î¶µªº­I´ºÃC¦â
   textTMKey.BackColor = &H8000000F
   textTM05.BackColor = &H8000000F
   textTM06.BackColor = &H8000000F
   textTM07.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM27.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   
   MoveFormToCenter Me
   'add by nickc 2006/06/07
   isRuned = False
   Is717end = False
   Is715end = False
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010404_1.m_strIR01
   m_strIR02 = frm02010404_1.m_strIR02
   m_strIR03 = frm02010404_1.m_strIR03
   m_strIR04 = frm02010404_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "¡]«H¥ó½s¸¹:" & m_strIR01 & "-" & m_strIR03 & "¡^"
   End If
   '2019/5/10 END
   
   'Add By Sindy 2020/12/29
   FrameTM14.Left = 4650
   FrameTM14.Top = 3110
   FrameTM20.Left = 4650
   FrameTM20.Top = 3110
   If m_TM01 = "TF" Then
      FrameTM20.Visible = True
   Else
      FrameTM14.Visible = True
   End If
   '2020/12/29 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' ²M°£·j´MªºKey
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      strPrevForm = Empty
   End If
   
   Select Case nType
      ' ¥»©Ò®×¸¹ Äæ¦ì1
      Case 0: m_TM01 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì2
      Case 1: m_TM02 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì3
      Case 2: m_TM03 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì4
      Case 3: m_TM04 = strData
      ' ¨Ó¨ç¦¬¤å¤é
      Case 4: m_CP05 = strData
      ' ¨Ó·½µe­±
      Case 5: strPrevForm = strData
      'add by nick 2004/10/20
      Case 6: m_TM14 = strData: textTM14.Text = strData
'      Case 7: m_FinalDate = strData: textFinalDate.Text = strData 'Add By Sindy 2020/12/14
   End Select
End Sub

' ¨ú±o°Ó¼Ð°ò¥»ÀÉ
Private Sub QueryTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   m_blnReceiveSecond = False '2011/9/19 add by sonia
   ' ¨ú±o°Ó¼Ð°ò¥»ÀÉªº¬ÛÃö¶µ¥Ø
   'Modified by Lydia 2019/11/13 +Nation
   strSql = "SELECT x.*,y.NA85 FROM TradeMark x, Nation y " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' AND TM10=NA01(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' ¥Ó½Ð°ê®a
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         m_NA85 = "" & rsTmp.Fields("NA85") 'Added by Lydia 2019/11/13 ­pºâ°Ó¼Ð±M¥Î´Á¬O§_´î1¤Ñ
      End If
      ' ³]©w±M¥Î´Á­­ªº¤é´ÁÄæ¦ìªø«×
      'edit by nick 2004/10/06 ¥þ¬O¦è¤¸¦~
'      If m_TM10 < "010" Then
        'Add By Cheng 2003/11/19
'        Me.textTM14.MaxLength = 7
'         textTM21.MaxLength = 7
'         textTM22.MaxLength = 7
'      Else
        'Add By Cheng 2003/11/19
        Me.textTM14.MaxLength = 8
         textTM21.MaxLength = 8
         textTM22.MaxLength = 8
'      End If
      ' ¥Ó½Ð¤é
      If IsNull(rsTmp.Fields("TM11")) = False Then
         'edit by nick 2004/10/06
         'm_TM11 = TAIWANDATE(rsTmp.Fields("TM11"))
         m_TM11 = DBDATE(rsTmp.Fields("TM11"))
      End If
      ' ¥Ó½Ð®×¸¹
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      
      'Add By Sindy 2012/12/19
      ' ¼f©w¨Ó¨ç¤é
      If IsNull(rsTmp.Fields("TM13")) = False Then
         m_TM13 = rsTmp.Fields("TM13")
      Else
         m_TM13 = strSrvDate(1)
      End If
      '2012/12/19 End
      
      ' ¤½§i¤é
      If IsNull(rsTmp.Fields("TM14")) = False Then
        'edit by nick 2004/10/06 ¥þ¬O¦è¤¸¦~
'        If m_TM10 = "000" Then
'            Me.textTM14.Text = TAIWANDATE(rsTmp.Fields("TM14"))
'        Else
            Me.textTM14.Text = rsTmp.Fields("TM14")
'        End If
         'edit by nick 2004/10/06 ¥þ¬O¦è¤¸¦~
         'm_TM14 = TAIWANDATE(rsTmp.Fields("TM14"))
         m_TM14 = rsTmp.Fields("TM14")
      End If
      
      'Add By Sindy 2020/12/29
      ' µù¥U¤é
      If IsNull(rsTmp.Fields("TM20")) = False Then
         Me.textTM20.Text = rsTmp.Fields("TM20")
      End If
      '2020/12/29 END
      
      ' ¼f©w¸¹
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' °Ó¼Ð¦WºÙ(¤¤)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05 = rsTmp.Fields("TM05")
      End If
      ' °Ó¼Ð¦WºÙ(­^)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textTM06 = rsTmp.Fields("TM06")
      End If
      ' °Ó¼Ð¦WºÙ(¤é)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textTM07 = rsTmp.Fields("TM07")
      End If
      ' ¥Ó½Ð¤H
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
      End If
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
      End If
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
      End If
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
      End If
      ' °Ó¼ÐºØÃþ
      If IsNull(rsTmp.Fields("TM08")) = False Then
         m_TM08 = rsTmp.Fields("TM08")
         If m_TM10 < "010" Then
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' °Ó«~Ãþ§O
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' ¥¿°Ó¼Ð¸¹¼Æ
      If IsNull(rsTmp.Fields("TM27")) = False Then
         m_TM27 = rsTmp.Fields("TM27")
         textTM27 = rsTmp.Fields("TM27")
      End If
      'Add By Sindy 2013/5/3
      '©ñ±ó±M¥ÎÅv
      If IsNull(rsTmp.Fields("TM67")) = False Then
         m_TM67 = rsTmp.Fields("TM67")
      End If
      '¦P·N®Ñ°Ó¼Ð¸¹¼Æ
      If IsNull(rsTmp.Fields("TM118")) = False Then
         m_TM118 = rsTmp.Fields("TM118")
      End If
      '2013/5/3 End
      ' ±M¥Î´Á­­ (°_)
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_TM21 = rsTmp.Fields("TM21")
         'edit by nick 2004/10/06 ¥þ¬O¦è¤¸¦~
'         If m_TM10 < "010" Then
'            textTM21 = TAIWANDATE(rsTmp.Fields("TM21"))
'         Else
            textTM21 = DBDATE(rsTmp.Fields("TM21"))
'         End If
      End If
      ' ±M¥Î´Á­­ (¤î)
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = rsTmp.Fields("TM22")
         'edit by nick 2004/10/06 ¥þ¬O¦è¤¸¦~
'         If m_TM10 < "010" Then
'            textTM22 = TAIWANDATE(rsTmp.Fields("TM22"))
'         Else
            textTM22 = DBDATE(rsTmp.Fields("TM22"))
'         End If
      End If
      textPrint = CheckStr(rsTmp.Fields("TM77"))
      '2011/9/19 ADD BY SONIA
      If InStr("" & rsTmp.Fields("TM58"), "²Ä¤G´Á") > 0 Then
         m_blnReceiveSecond = True
      End If
      '2011/9/19 end
      
      'Added by Lydia 2016/12/22
      m_TM44 = CheckStr("" & rsTmp.Fields("TM44"))
      
      'Added by Morgan 2025/2/18
      m_TM136 = "" & rsTmp.Fields("TM136")
      ChkTM136
      'end 2025/2/18

   End If
   '2006/5/3 ADD BY SONIA °¨¼w¨½«ü©w°ê®a¬ü°ê®É¦P®É¥i¿é¤J¥Ó½Ð®×¸¹
   If m_TM10 = "101" Then
      textTM12.Locked = False
      textTM12.TabStop = True
      textTM12.BorderStyle = 1
      textTM12.BackColor = &H80000005
      textTM12.SetFocus
   Else
      textTM12.Locked = True
      textTM12.TabStop = False
      textTM12.BorderStyle = 0
      textTM12.BackColor = &H8000000F
   End If
   '2006/5/3 END
   rsTmp.Close
   Set rsTmp = Nothing
   'add by nickc 2006/06/07 ÀË¬d¦³µLµ²®×¹L²Ä¤@´Áµù¥U¶O¡A©M¥þ´Áµù¥U¶O
   If isRuned = False Then
       strSql = "select * from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07=717 and np06='N' "
       Set rsTmp = New ADODB.Recordset
       If rsTmp.State = 1 Then rsTmp.Close
       rsTmp.CursorLocation = adUseClient
       rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If rsTmp.RecordCount <> 0 Then
           Is717end = True
           'Add By Sindy 2012/4/26 ÀË¬d¸Ñ°£´Á­­«á¬O§_ÁÙ¦³¦¬²Ä¤@´Áµù¥U¶O¥B¤wµo¤å,­Y¦³,«hÁÙ¬O­n±¾²Ä¤G´Áµù¥U¶O
           strSql = "select * from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='715' and cp27 is not null and cp05>=" & rsTmp.Fields("NP11")
           Set rsTmp = New ADODB.Recordset
           If rsTmp.State = 1 Then rsTmp.Close
           rsTmp.CursorLocation = adUseClient
           rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
           If rsTmp.RecordCount <> 0 Then
               Is717end = False
           End If
           '2012/4/26 End
       Else
           strSql = "select * from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07=715 and np06='N' "
           Set rsTmp = New ADODB.Recordset
           If rsTmp.State = 1 Then rsTmp.Close
           rsTmp.CursorLocation = adUseClient
           rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
           If rsTmp.RecordCount <> 0 Then
               Is715end = True
               MsgBox "´¿¸Ñ°£¹L²Ä¤@´Áµù¥U¶O¡A½Ð½T»{¬O§_»Ý" & vbCrLf & "¸Ñ°£²Ä¤G´Áµù¥U¶O¡H", vbExclamation, "´£¿ô¡I"
           End If
       End If
       Set rsTmp = Nothing
   End If
   '2006/06/07 end
End Sub

' ¨ú±oªA°È·~°È°ò¥»ÀÉ
Private Sub QueryServicePractice()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' ¨ú±o°Ó¼Ð°ò¥»ÀÉªº¬ÛÃö¶µ¥Ø
   'Modified by Lydia 2019/11/13 +Nation
   strSql = "SELECT x.*,y.NA85 FROM ServicePractice x, Nation y " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' AND SP09=NA01(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' ¥Ó½Ð°ê®a
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         m_NA85 = "" & rsTmp.Fields("NA85") 'Added by Lydia 2019/11/13 ­pºâ°Ó¼Ð±M¥Î´Á¬O§_´î1¤Ñ
      End If
      ' ³]©w±M¥Î´Á­­ªº¤é´ÁÄæ¦ìªø«×
      If m_TM10 < "010" Then
         textTM21.MaxLength = 7
         textTM22.MaxLength = 7
      Else
         textTM21.MaxLength = 8
         textTM22.MaxLength = 8
      End If
      ' °Ó¼Ð¦WºÙ(¤¤)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textTM05 = rsTmp.Fields("SP05")
      End If
      ' °Ó¼Ð¦WºÙ(­^)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      ' °Ó¼Ð¦WºÙ(¤é)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
      ' ¥Ó½Ð¤H
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      
      'Add By Sindy 2020/12/29
      ' µù¥U¤é
      If IsNull(rsTmp.Fields("SP12")) = False Then
         Me.textTM20.Text = rsTmp.Fields("SP12")
      End If
      '2020/12/29 END
      
      'Add By Sindy 2019/12/25
      ' FC¥N²z¤H
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
      End If
      '2019/12/25 END
      
      If IsNull(rsTmp.Fields("SP58")) = False Then
         m_TM78 = rsTmp.Fields("SP58")
      End If
      If IsNull(rsTmp.Fields("SP59")) = False Then
         m_TM79 = rsTmp.Fields("SP59")
      End If
      If IsNull(rsTmp.Fields("SP65")) = False Then
         m_TM80 = rsTmp.Fields("SP65")
      End If
      If IsNull(rsTmp.Fields("SP66")) = False Then
         m_TM81 = rsTmp.Fields("SP66")
      End If
      ' ±M¥Î´Á­­ (°_)
      If IsNull(rsTmp.Fields("SP20")) = False Then
         m_TM21 = rsTmp.Fields("SP20")
         'edit by nick 2004/10/06
'         If m_TM10 < "010" Then
'            textTM21 = TAIWANDATE(rsTmp.Fields("SP20"))
'         Else
            textTM21 = DBDATE(rsTmp.Fields("SP20"))
'         End If
      End If
      ' ±M¥Î´Á­­ (¤î)
      If IsNull(rsTmp.Fields("SP21")) = False Then
         m_TM22 = rsTmp.Fields("SP21")
         'edit by nick 2004/10/06
'         If m_TM10 < "010" Then
'            textTM22 = TAIWANDATE(rsTmp.Fields("SP21"))
'         Else
            textTM22 = DBDATE(rsTmp.Fields("SP21"))
'         End If
      End If
      ' §@«~ºØÃþ
      'Add By Cheng 2002/07/17
      m_SP46 = Empty
      If IsNull(rsTmp.Fields("SP46")) = False Then
         m_SP46 = rsTmp.Fields("SP46")
      End If
      'Add By Cheng 2002/06/12
      '¥DºÞ¾÷Ãö
      m_SP51 = "" & rsTmp.Fields("SP51").Value
      'ADD BY SONIA 91.11.1
      If IsNull(rsTmp.Fields("SP13")) = False Then
         textTC1 = rsTmp.Fields("SP13")
      End If
      If IsNull(rsTmp.Fields("SP14")) = False Then
         textTC2 = rsTmp.Fields("SP14")
      End If
      'add by nickc 2006/11/20
      textPrint = CheckStr(rsTmp.Fields("SP72"))
      '91.11.1 END
   End If

   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryCaseProgress()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
'Add By Cheng 2003/12/09
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   '2012/11/1 add by sonia ¥ý§ì¥Ó½Ð101©Î¤À³Î308,­YµL101¤Î308¤~¥ý§ìAÃþ¦¬¤å (T-179141¼f©w«á¥¼¦A¦¬¤åµù¥U¶O©Î¨ä¥LAÃþ¬G·|§ì¨ì¤À³Î308)
   strSql = "SELECT * FROM CaseProgress WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' AND " & _
                  "CP10 = '101' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GoTo DisplayData
   End If
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   strSql = "SELECT * FROM CaseProgress WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' AND " & _
                  "CP10 = '308' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GoTo DisplayData
   End If
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   '2012/11/1 end
   
   ' ¨ú±o®×¥ó¶i«×ÀÉAÃþ¸ê®Æªº³Ì«á¤@µ§
    'Modify By Cheng 2003/01/10
'   strSQL = "SELECT * FROM CaseProgress " & _
'            "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                  "CP02 = '" & m_TM02 & "' AND " & _
'                  "CP03 = '" & m_TM03 & "' AND " & _
'                  "CP04 = '" & m_TM04 & "' AND " & _
'                  "CP09 LIKE 'A%' " & _
'                  "ORDER BY CP05 "
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP09 < 'B' " & _
                  "ORDER BY CP05 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
   If rsTmp.RecordCount > 0 Then
DisplayData:
      rsTmp.MoveLast
      ' ¾÷Ãö¤å¸¹
      'Add By Cheng 2002/07/17
      m_CP08 = Empty
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
      End If
      ' ¦¬¤å¸¹
      m_CP09 = Empty
      If IsNull(rsTmp.Fields("CP09")) = False Then
         m_CP09 = rsTmp.Fields("CP09")
      End If
      ' ®×¥ó©Ê½è
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
      End If
      'END 2002/07/17
      ' ´¼Åv¤H­û
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         strExc(0) = ""
         Me.textCP13.Text = GetStaffName(m_CP13, True)
      End If
      '·~°È°Ï   nick 91.08.22
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("cp12")) = False Then
          m_CP12 = rsTmp.Fields("cp12")
      End If
   'Add By Cheng 2003/01/10
   '­YµLAÃþ¸ê®Æ, ¦A§ìBÃþ¸ê®Æ
   Else
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = Nothing
        strSql = "SELECT * FROM CaseProgress " & _
                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                       "CP02 = '" & m_TM02 & "' AND " & _
                       "CP03 = '" & m_TM03 & "' AND " & _
                       "CP04 = '" & m_TM04 & "' AND " & _
                       "CP09 > 'B' AND CP09 < 'C' " & _
                       "ORDER BY CP05 "
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
        If rsTmp.RecordCount > 0 Then
            GoTo DisplayData
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   '2005/4/14 ADD BY SONIA
   If m_TM01 = "TF" And Mid(m_TM02, 6, 1) <> "0" Then
      StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='104' "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         textEditPrint = "Y"
         Label15.Visible = True
         textCP47.Visible = True
         textCP47.Locked = False
         If IsNull(rsA.Fields("CP47")) Then
            textCP47 = ""
         Else
            textCP47 = rsA.Fields("CP47")
         End If
      Else
         textEditPrint = ""
         Label15.Visible = False
         textCP47.Visible = False
         textCP47.Locked = True
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   Else
      textEditPrint = ""
      Label15.Visible = False
      textCP47.Visible = False
      textCP47.Locked = True
   End If
   '2005/4/14 END
    
    'Add By Cheng 2003/12/09
    '§PÂ_¬O§_¤w¦¬²Ä¤G´Áµù¥U¶O
    '93.10.7 MODIFY BY SONIA
    'strSQLA = "Select * From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='716' "
    StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And (CP10='716' OR CP10='717')"
    '93.10.7 END
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '2011/9/19 modify by sonia ®×¥ó³Æµùtm58­Y¤w¥[µù«h¤£ºÞ¨î
    'If rsA.RecordCount > 0 Then
    If rsA.RecordCount > 0 And m_blnReceiveSecond = False Then
        m_blnReceiveSecond = True
    Else
        m_blnReceiveSecond = False
    End If
    
    'Added by Lydia 2017/02/02 ¥x-¤j­«µoµù¥Uµý,1006³¡¤À³Ó³¡¤À±Ñªº¶i«×³Æµù
    str1006CP64 = ""
    If m_TM01 = "T" And m_TM10 = "020" And (m_CP10 = "101" Or m_CP10 = "308") Then
       ChgCaseNo textTMKey.Text, strExc
       If PUB_ChkCPExist(strExc, "1701", 2) Then '¦³µo¹Lµù¥Uµý
          'Modified by Lydia 2017/02/06 ¼W¥[³¡¥÷ºM¾P1004,­ì¥»µ{§Ç·|¦b³ÌªìªºAÃþ¦¬¤å¸É³Æµù,¥H«á­n¦bµ²ªGªºCÃþ¦¬¤å1004¸É³Æµù
          'StrSQLa = "Select CP05,CP09,CP64 From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='1006' ORDER BY CP05 DESC "
          StrSQLa = "Select CP05,CP09,CP64 From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 IN ('1006','1004') ORDER BY CP05 DESC "
          If rsA.State <> adStateClosed Then rsA.Close
          rsA.CursorLocation = adUseClient
          rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
          If rsA.RecordCount > 0 Then
             str1006CP64 = "" & rsA.Fields("CP64")
             If str1006CP64 = "" Then str1006CP64 = "TRUE"
             textEditPrint = "Y" '¹w³]¶}Word
          End If
       End If
    End If
    'end 2017/02/02
    
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
   
End Sub

' ¬d¸ß¸ê®Æ®w¨ú±o¸ê®Æ
Public Sub QueryData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim m_msg As String    'add by sonia 2019/1/30
Dim strTo As String, strFA119 As String
   
   '2005/11/11 ADD BY SONIA
   '¨ú±o©w½Z»y¤å
   m_strLanguage = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
'   If m_strLanguage = "2" Then
'      Label16.Visible = True
'      Text1.Enabled = True
'      Text1.Visible = True
'   Else
'      Label16.Visible = False
'      Text1.Enabled = False
'      Text1.Visible = False
'   End If
   '2005/11/11 END
   
   'add by nick 2004/10/20
   textTM14.Text = m_TM14
'   textFinalDate.Text = m_FinalDate 'Add By Sindy 2020/12/14
   
   m_TM10 = Empty
   m_CP13 = Empty
   m_NA85 = Empty 'Added by Lydia 2019/11/13
   
   ' ¥»©Ò®×¸¹
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' ¨Ó¨ç¦¬¤å¤é
   textCP05S = m_CP05
      
   m_SP51 = ""
   m_TM08 = Empty
   m_TM11 = Empty
   m_TM14 = Empty
   m_TM21 = Empty
   m_TM22 = Empty
   m_TM23 = Empty
   m_TM78 = Empty
   m_TM79 = Empty
   m_TM80 = Empty
   m_TM81 = Empty
   m_TM27 = Empty
   m_TM13 = Empty 'Add By Sindy 2012/12/19 ¼f©w¨Ó¨ç¤é
   m_TM67 = Empty 'Add By Sindy 2013/5/3
   m_TM118 = Empty 'Add By Sindy 2013/5/3
   
   ' Åª¨ú°ò¥»ÀÉ
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   ' Åª¨ú®×¥ó¶i«×ÀÉ
   QueryCaseProgress
   ' ³]©w±M¥Î´Á­­ªº¤é´Á
   'edit by nick 2004/10/06
'   If m_TM10 < "010" Then
'      textTM21.MaxLength = 7
'      textTM22.MaxLength = 7
'   Else
      textTM21.MaxLength = 8
      textTM22.MaxLength = 8
'   End If
   
   ' ¨ú±o±M¥Î´Á­­
   If IsEmptyText(m_TM27) = False Then
    'Modify By Cheng 2002/12/09
    '­Y°Ó¼ÐºØÃþ¬°2,3«h§ì1; 5,6«h§ì4; ¨ä¥L«h·ÓÂÂ
    If m_TM08 = "2" Or m_TM08 = "3" Then
        strSql = "SELECT * FROM TradeMark " & _
                 "WHERE TM15 = '" & m_TM27 & "' And TM08='1' "
    ElseIf m_TM08 = "5" Or m_TM08 = "6" Then
        strSql = "SELECT * FROM TradeMark " & _
                 "WHERE TM15 = '" & m_TM27 & "' And TM08='4' "
    Else
        strSql = "SELECT * FROM TradeMark " & _
                 "WHERE TM15 = '" & m_TM27 & "' "
    End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         ' ±M¥Î´Á­­°_¤é 91.8.30 MODIFY BY SONIA
         'If IsNull(rsTmp.Fields("TM21")) = False Then
         '   If rsTmp.Fields("TM21") <> "0" Then
         '      If m_TM10 < "010" Then
         '         textTM21 = TAIWANDATE(rsTmp.Fields("TM21"))
         '      Else
         '         textTM21 = DBDATE(rsTmp.Fields("TM21"))
         '      End If
         '   End If
         'End If
         '91.8.30 END
         ' ±M¥Î´Á­­¤î¤é
         If IsNull(rsTmp.Fields("TM22")) = False Then
            If rsTmp.Fields("TM22") <> "0" Then
                'edit by nick 2004/10/06
'               If m_TM10 < "010" Then
'                  textTM22 = TAIWANDATE(rsTmp.Fields("TM22"))
'               Else
                  textTM22 = DBDATE(rsTmp.Fields("TM22"))
'               End If
            End If
         End If
      End If
      rsTmp.Close
   End If
   
   'add by sonia 2019/1/29 ¤j³°®×¤~¿ï¹q¤lÃÒ®Ñ©Î¯È¥»ÃÒ®Ñ
   Frame3.Visible = False
   m_msg = ""
   ' end 2019/1/29
   
   ' ¤j³°»âÃÒ¶O
   If m_TM10 = "020" Then
      EnableTextBox textMoney, True
      If str1006CP64 = "" Then 'Added by Lydia 2017/02/02 ³¡¥÷ºM¾P­«µoµù¥UÃÒ¤£¥Î¶O¥Î
        'Add By Sindy 2009/10/23
        '******¯S§Oª`·N,¦¹³B­Y­×§ï¹w³]³ø»ù,¹q¸£¤¤¤ßªº¹wºâªíµ{¦¡¤]­n§ï
        If m_CP13 = "69010" Then
           textMoney = "5000"
        ElseIf m_CP13 = "76051" Then
           textMoney = "6000"
        Else
        '2009/10/23 End
           textMoney = "3000"
        End If
      End If 'end 2017/02/02
      
      'add by sonia 2019/1/29 ¹q¤l°e¥ó©~¦h,¥¼¿é¹L¹q¤lÃÒ®Ñ«h¹w³]¹q¤lÃÒ®Ñ
      Frame3.Visible = True
      Option5(0).Value = False: Option5(1).Value = False
      strSql = "SELECT CP09,CP64,NVL(INSTR(CP64,'¹q¤lÃÒ®Ñ'),0),NVL(INSTR(CP64,'¯È¥»ÃÒ®Ñ'),0) FROM CASEPROGRESS Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='1701' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         textMoney = 0
         With rsTmp
         .MoveFirst
         Do While Not .EOF
            If Val(rsTmp.Fields(2)) > 0 Then
               Option5(1).Value = True
               m_msg = "¹q¤lÃÒ®Ñ"
            End If
            If Val(rsTmp.Fields(3)) > 0 Then
               If m_msg = "" Then
                  Option5(0).Value = True
                  m_msg = "¯È¥»ÃÒ®Ñ"
               Else
                  m_msg = m_msg & "¡B¯È¥»ÃÒ®Ñ"
               End If
            End If
            .MoveNext
         Loop
         End With
         MsgBox "¥»®×¤w³qª¾¹L" & m_msg & "µù¥UÃÒ¡A½Ð¯d·N¡I", vbExclamation + vbOKOnly
         If Option5(0).Value = False And Option5(1).Value = False Then Option5(0).Value = True
      Else
         Option5(0).Value = True
      End If
      rsTmp.Close
      'end 2019/1/29
   Else
      EnableTextBox textMoney, False
   End If
   
   ' Ãº¦~¶O´Á­­
   If m_TM01 = "TF" And Len(m_TM02) = 6 Then
      EnableTextBox textDate, True
   Else
      EnableTextBox textDate, False
   End If
   
   ' TCµn°O¸¹¤ÎTCµù¥U¸¹¼ÆÄæ¦ì
   If m_TM01 = "TC" Then
      EnableTextBox textTC1, True
      EnableTextBox textTC2, True
   Else
      EnableTextBox textTC1, False
      EnableTextBox textTC2, False
   End If
      
   'Add By Cheng 2002/06/12
   If m_TM01 = "TC" And m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
      Me.textTM21.Enabled = False
      Me.textTM22.Enabled = False
      Me.textDate.Enabled = False
      Me.textMoney.Enabled = False
      Me.textTC1.Enabled = False
   End If
   
   Set rsTmp = Nothing
   
   '910729 Sieg
   m_TM21 = ""
   m_TM22 = ""
   
   If m_TM01 = "T" Then
        '­Y¦³¤½§i¤é
        If m_TM14 <> "" Then
            '­Y¤½§i¤é¦b920816(§t)¥H«eªÌ, ©Î¬O¤j³°®×
            'edit by nick 2004/10/06
            'If Val(m_TM14) <= 920816 Or m_TM10 = "020" Then
            'Modified Lydia 2019/12/09 ¥þ³¡§ï¥Î·sªk, ¤j³°®×±M¥Î´Á¶¡=±M¥Î´Á°_¤é¬°¤½§i¤é+3­Ó¤ë+1¤Ñ,±M¥Î´Á¤î¤é¬°¤½§i¤é+3­Ó¤ë+10¦~
'            If Val(m_TM14) <= 20030816 Or m_TM10 = "020" Then
'            'End
'                  '«D°¨¼w¨½®×±M¥Î´Á¶¡°_¤é¬°¤½§i¤é+¤T­Ó¤ë
'                  'edit by nick 2004/10/06
'                  'm_TM21 = TAIWANDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_TM14))))
'                  m_TM21 = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_TM14))))
'                  Select Case m_TM08
'                     'modify by sonia 2013/11/27 ¥[9¹ÎÅé°Ó¼Ð
'                     Case "1", "4", "7", "8", "9":
'                        '±M¥Î´Á¶¡¤î¤é¬°¤½§i¤é+¤T­Ó¤ë°_¤Q¦~´î¤@¤Ñ
'                        'edit by nick 2004/10/06
'                        'm_TM22 = TAIWANDATE(DateAdd("d", -1, DateAdd("yyyy", 10, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_TM14))))))
'                        'Modified by Lydia 2019/11/13 §ï¥Î¦@¥Î¼Ò²Õ
'                        'm_TM22 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 10, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_TM14))))))
'                        'Modified by Lydia 2019/12/05 +´î¤@¤Ñ=Y                        '
'                        m_TM22 = PUB_GetEndDate(CompDate(1, 3, DBDATE(m_TM14)), 10, "Y")
'                     Case Else
'                        strExc(0) = "SELECT TM22 FROM TRADEMARK WHERE TM15 = '" & m_TM27 & "' "
'                        intI = 1
'                        'edit by nickc 2007/02/06 ¤£¥Î dll ¤F
'                        'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                           If Not IsNull(RsTemp.Fields("TM22")) Then
'                              'edit by nick 2004/10/06
'                              'm_TM22 = TransDate(rsTemp.Fields("TM22"), 1)
'                              m_TM22 = TransDate(RsTemp.Fields("TM22"), 2)
'                           End If
'                        End If
'                  End Select
'                  '2008/5/15 add by sonia ¤j³°¤½§i¤é2007/12/1¥H«á,±M¥Î´Á°_¤é¬°¤½§i¤é+3­Ó¤ë+1¤Ñ,±M¥Î´Á¤î¤é¬°¤½§i¤é+3­Ó¤ë+10¦~
'                  If Val(m_TM14) >= 20071201 And m_TM10 = "020" Then
'                     m_TM21 = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM21))))
'                     m_TM22 = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
'                  End If
'                  '2008/5/15 end
            If m_TM10 = "020" Then
                '¤j³°®×·sªk=±M¥Î´Á°_¤é¬°¤½§i¤é+3­Ó¤ë+1¤Ñ,±M¥Î´Á¤î¤é¬°¤½§i¤é+3­Ó¤ë+10¦~
                m_TM21 = CompDate(1, 3, DBDATE(m_TM14)) '¤½§i¤é+3­Ó¤ë
                m_TM22 = PUB_GetEndDate(m_TM21, 10, m_NA85) '±M¥Î´Á¤î¤é¬°¤½§i¤é+3­Ó¤ë+10¦~
                m_TM21 = CompDate(2, 1, m_TM21)
            'end 2019/12/09
            
            '­Y¤½§i¤é¤j©ó920816¥B«D¤j³°®×
            Else  'Memo by Lydia 2019/12/09 ¥xÆW®×·sªk:±M¥Î´Á°_¤é¬°¤½§i¤é,±M¥Î´Á¶¡¤î¤é¬°¤½§i¤é¥[¤Q¦~´î¤@¤Ñ
                '«D°¨¼w¨½®×±M¥Î´Á¶¡°_¤é¬°¤½§i¤é
                m_TM21 = m_TM14
                '±M¥Î´Á¶¡¤î¤é¬°¤½§i¤é¥[¤Q¦~´î¤@¤Ñ
                'edit by nick 2004/10/06
                'm_TM22 = TAIWANDATE(DateAdd("d", -1, DateAdd("yyyy", 10, ChangeWStringToWDateString(DBDATE(m_TM14)))))
                'Modified by Lydia 2019/11/13 §ï¥Î¦@¥Î¼Ò²Õ
                'm_TM22 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 10, ChangeWStringToWDateString(DBDATE(m_TM14)))))
                'Modify By Sindy 2022/3/7 + m_TM10 : ©µ®i«á¤§±M¥Î´Á­­¦~«×­Õ¦³2¤ë29¤é®É¡A±M¥Î´Á­­¤î¤éÀ³¬°2¤ë29¤é¡A¦Ó«D¥H¥[10¦~¤§¤è¦¡­pºâ¬°2¤ë28¤é
                m_TM22 = PUB_GetEndDate(DBDATE(m_TM14), 10, m_NA85, m_TM10)
            End If
        End If
'2008/11/25 cancel by sonia TF-000570®Û­^»¡TF®×¥ó¤£·|ª¾¹D¥Ó½Ð¤é,¬G¤£ÀË¬d
'   ElseIf m_TM01 = "TF" Then
'      Dim strKey(0 To 4) As String, strTmp As String
'      strKey(0) = m_CP09
'      strKey(1) = m_TM01
'      strKey(2) = m_TM02
'      strKey(3) = m_TM03
'      strKey(4) = m_TM04
'      If TFGetMoneyDate(m_TM10, strKey, m_TM21, strTmp, m_TM22) Then
'          'edit by nick 2004/10/06
''         m_TM21 = TransDate(m_TM21, 1)
''         m_TM22 = TransDate(CompDate(2, -1, m_TM22), 1)
'         m_TM21 = TransDate(m_TM21, 2)
'         m_TM22 = TransDate(CompDate(2, -1, m_TM22), 2)
'      End If
'2008/11/25 END
   End If
   
   '91.10.24 ADD BY SONIA
   If m_TM01 = "T" Then
      'modify by sonia 2016/8/2 ®Û­^¶l¥ó´£¥X­×§ï
      'Select Case m_TM10
      'Case "000"
      '   textPS = "ªþ¥ó¡Gµù¥UÃÒ¥¿¥»¤Î¦æ¨Ï°Ó¼Ð±M¥ÎÅv¶·ª¾¤A¥÷¡C"
      'Case "020"
      '   textPS = "ªþ¥ó¡Gµù¥UÃÒ¥¿¥»¤A¯È¡C"
      'End Select
      textPS = "ªþ¥ó¡Gµù¥UÃÒ¥¿¥»¤A¯È¡C"
      'end 2016/8/2
      If Frame3.Visible = True And Option5(0).Value = True Then textPS = "ªþ¥ó¡G°Ó¼Ð¹q¤lµù¥UÃÒ¤§¯È¥»¡C"  'add by sonia 2019/2/1
   End If
   '91.10.25 END
       
    'add by nickc 2006/10/02 ¥N¥À®×¸ê®Æ
    If UpForm Is frm02010401_6 Then
       QueryMonTradeMark
    End If
   'add by nickc 2006/06/30 ±a¦C¦L©w½Z¹w³]­È
   'edit by nickc 2006/11/20
   If textPrint = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
   Call ChgType 'Add By Sindy 2012/5/18 Åª¨ú¨Ó¨ç´Á­­
   
   'Add By Sindy 2021/1/6 ¦³FC¥N²z¤H¤~­nÅã¥Ü¡i³°¥N©w½Z¥[µù¡j
   If m_TM44 <> "" Then
      strTo = "": strFA119 = ""
      strTo = PUB_GetFCeMailConText("Main_EMail", m_TM01, m_TM02, m_TM03, m_TM04, "FC", , True)
      If strTo <> "" Then
         CheckOC3
         strExc(0) = "select fa01,fa02,fa119" & _
                     " from fagent" & _
                     " where fa01='" & Left(strTo, 8) & "' and fa02='" & Mid(strTo, 9, 1) & "'"
         intI = 1
         Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strFA119 = "" & AdoRecordSet3.Fields("FA119")
         End If
         CheckOC3
      End If
      If strFA119 <> "" Then
         MsgBox "¡i³°¥N©w½Z¥[µù¡j" & vbCrLf & vbCrLf & strFA119, vbInformation
      End If
   End If
   '2021/1/6 END
End Sub

Public Function OnSaveData() As Boolean
Dim strSql As String
Dim strDateFrom As String
Dim strDateTo As String
Dim strCP10 As String
Dim strCP12 As String
Dim strCP20 As String
Dim strCP27 As String
Dim strCP32 As String
Dim strNP07 As String
Dim strNP09 As String
'93.6.11 ADD BY SONIA
Dim strCP06 As String
Dim strCP07 As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'93.6.11 END
Dim strNA38 As String    '2006/5/3 ADD BY SONIA
Dim nResponse 'Add By Sindy 2010/01/13
Dim w_CP09 As String     'add by sonia 2013/8/8
Dim ii As Integer        'add by sonia 2013/8/8
Dim strCP64 As String    'add by sonia 2019/5/7

   OnSaveData = True
   'add by nickc 2006/08/11
   If Me.Visible = True Then
       On Error GoTo ErrorHandler
       cnnConnection.BeginTrans
   End If
   ' ¦¹¬qµ{¦¡½X¦b§ó·s°Ó¼Ð°ò¥»ÀÉ©Î¬OªA°È·~°È°ò¥»ÀÉ
   ' §ó·s±M¥Î´Á­­°_¤é¤Î¤î¤é
   strDateFrom = DBDATE(textTM21)
   strDateTo = DBDATE(textTM22)
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         '2008/10/24 modify by sonia µù¥U¤À³Î¤l®×¦P®É±N¥À®×¥Ó½Ð®×¸¹§ó·s¦Ü¤l®×,TM13¼f©w¨Ó¨ç¤é¤W¨Ó¨ç¦¬¤å¤é,TM16­ã»éÄæ¤W­ã,T-137268
         'strSQL = "UPDATE TradeMark " & _
                  "SET TM17='Y',TM20 = " & DBNullDate(m_CP05) & ", " & _
                      "TM21 = " & DBNullDate(textTM21) & ", " & _
                      "TM22 = " & DBNullDate(textTM22) & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
         If m_CP10 = "308" Then
            '2011/9/20 modify by sonia ¥[¤£ºÞ¨î²Ä¤G´Á³Æµù
            'Modify By Sindy 2020/12/29 TFµe­±¼W¥[µoÃÒ¤éÄæ¦ì
            'DBNullDate(m_CP05) => IIf(FrameTM20.Visible = True, textTM20, DBNullDate(m_CP05))
            strSql = "UPDATE TradeMark " & _
                     "SET TM16='1',TM17='Y',TM20 = " & IIf(FrameTM20.Visible = True, textTM20, DBNullDate(m_CP05)) & ", " & _
                         "TM12 ='" & textTM12 & "', TM13 = " & DBNullDate(m_CP05) & ", " & _
                         "TM21 = " & DBNullDate(textTM21) & ", " & _
                         "TM22 = " & DBNullDate(textTM22) & ", " & _
                         "TM58 = " & IIf(m_blnReceiveSecond, "decode(tm58,null,'¤£ºÞ¨î²Ä¤G´Á;','¤£ºÞ¨î²Ä¤G´Á;'||tm58) ", "tm58") & " " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "'"
         Else
            'Modify By Sindy 2020/12/29 TFµe­±¼W¥[µoÃÒ¤éÄæ¦ì
            'DBNullDate(m_CP05) => IIf(FrameTM20.Visible = True, textTM20, DBNullDate(m_CP05))
            strSql = "UPDATE TradeMark " & _
                     "SET TM17='Y',TM20 = " & IIf(FrameTM20.Visible = True, textTM20, DBNullDate(m_CP05)) & ", " & _
                         "TM21 = " & DBNullDate(textTM21) & ", " & _
                         "TM22 = " & DBNullDate(textTM22) & " " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "'"
         End If
         '2008/10/24 END
         cnnConnection.Execute strSql
         'add by nickc 2006/11/20
         If textPrint <> "N" Then
            strSql = "UPDATE TradeMark " & _
                     "SET TM77='" & textPrint & "' " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "'"
            cnnConnection.Execute strSql
         End If
        'Add By Cheng 2003/11/19
'        '­Y¥Ó½Ð¤é¬°921128(§t)¥H«áªÌ
'        If m_TM01 = "T" And m_TM10 = "000" And DBDATE(Val(m_TM11)) >= 20031128 Then
        If m_TM01 = "T" And m_TM10 = "000" Then
            strSql = "Update Trademark Set TM14=" & DBDATE(Me.textTM14.Text) & ", TM15='" & Me.textTM15.Text & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute strSql
        End If
        'Add By Cheng 2004/04/12
        '§ó·sTFªºµù¥U¸¹(TF ¶}©ñµù¥U¸¹Äæ¦ì)
        If m_TM01 = "TF" Then
            '2006/5/3 MODIFY BY SONIA °¨¼w¨½¬ü°ê®×¦P®É§ó·s¥Ó½Ð®×¸¹
            'strSQL = "Update Trademark Set TM15='" & Me.textTM15.Text & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            '2012/6/14 MODIFY BY SONIA °¨¼w¨½¬ü°ê®×¦P®É§ó·sµù¥U¤½§i¤é
            'modify by sonia 2015/12/14 +TM16='1'(¤£ºÞ¬O¤£¬O¬ü°ê®×,¿é¤Jªº¨º¤@µ§¤]­n§ó·sTM16='1'
            'Modify By Sindy 2020/12/29 TFµe­±¤wµL¤½§i¤éÄæ¦ì
            'strSql = "Update Trademark Set TM16='1',TM12='" & Me.textTM12.Text & "',TM14=" & CNULL(DBDATE(Me.textTM14.Text)) & ",TM15='" & Me.textTM15.Text & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            strSql = "Update Trademark Set TM16='1',TM12='" & Me.textTM12.Text & "',TM15='" & Me.textTM15.Text & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            '2006/5/3 END
            cnnConnection.Execute strSql
        End If
        'End
      Case Else
         '91.11.3 MODIFY BY SONIA
         'strSQL = "UPDATE ServicePractice " & _
         '         "SET SP12 = " & DBNullDate(m_CP05) & ", " & _
         '             "SP21 = " & DBNullDate(textTM21) & ", " & _
         '             "SP22 = " & DBNullDate(textTM22) & " " & _
         '         "WHERE SP01 = '" & m_TM01 & "' AND " & _
         '               "SP02 = '" & m_TM02 & "' AND " & _
         '               "SP03 = '" & m_TM03 & "' AND " & _
         '               "SP04 = '" & m_TM04 & "'"
         'cnnConnection.Execute strSQL
         strSql = "UPDATE ServicePractice " & _
                  "SET SP12 = " & DBNullDate(m_CP05) & ", " & _
                      "SP13 = " & CNULL(textTC1) & ", " & _
                      "SP14 = " & CNULL(textTC2) & ", " & _
                      "SP20 = " & DBNullDate(textTM21) & ", " & _
                      "SP21 = " & DBNullDate(textTM22) & " " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "'"
         cnnConnection.Execute strSql
         'add by nickc 2006/11/20
         If textPrint <> "N" Then
            strSql = "UPDATE ServicePractice " & _
                     "SET SP72 = '" & textPrint & "' " & _
                     "WHERE SP01 = '" & m_TM01 & "' AND " & _
                           "SP02 = '" & m_TM02 & "' AND " & _
                           "SP03 = '" & m_TM03 & "' AND " & _
                           "SP04 = '" & m_TM04 & "'"
            cnnConnection.Execute strSql
         End If
         '91.11.3 END
   End Select
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' ¨t²ÎÃþ§O¬°TF®É¦P®É§ó·s©Ò¦³¥¼®Ö»éªº¤l®×
   If m_TM01 = "TF" And m_TM04 = "00" And m_TM03 = "0" Then
      'modify by sonia 2015/12/14 +TM16='1'
      'Modify By Sindy 2020/12/29 TFµe­±¼W¥[µoÃÒ¤éÄæ¦ì
      'DBDATE(m_CP05) => DBNullDate(textTM20)
      strSql = "UPDATE TradeMark " & _
               "SET TM16='1',TM17='Y',TM20 = " & DBNullDate(textTM20) & ", " & _
                     "TM21 = " & DBDATE(textTM21) & ", " & _
                     "TM22 = " & DBDATE(textTM22) & " " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "(TM16 <> '2' OR TM16 IS NULL)"
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2009/06/16
   '¤j³°®×µoÃÒ®É¡A­Y¤U¤@µ{§Ç¦³¡]³Q²§Ä³Äò®i¡^ªº´Á­­,«h¦Û°Ê§ó·s¬°¤£Äò¿ì
   If m_TM01 = "T" And m_TM10 = "020" Then
      strSql = "update nextprogress " & _
                     "set np06='N', " & _
                     "     np11=" & strSrvDate(1) & ", " & _
                     "     np12='99', " & _
                     "     np15=decode(np15,null,'µoÃÒ¤£¥²¦AºÞ¨î',np15||';'||'µoÃÒ¤£¥²¦AºÞ¨î') " & _
                     "where np06 is null and np07=109 " & _
                     "and np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' "
'                     "and np01 in (select cp09 from caseprogress " & _
'                                         "where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & _
'                                         "' and cp10 in ('1601')) "
      cnnConnection.Execute strSql
      
      'Added by Lydia 2016/09/12 ©ó¼f©w®Ö­ã¿é¤J®É,¦P®ÉºÞ¨î¶Êµù¥Uµý®É¶¡=>­Y¤U¤@µ{§Ç¦³µù¥UÃÒ(1701)´Á­­§ó·s¬°Y
      strSql = "update nextprogress set np06='Y' " & _
               "where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' " & _
               "and nvl(np06,'0')='0' and np07='1701' "
      cnnConnection.Execute strSql
      'end 2016/09/12
      
   End If
   '2009/06/16 End
   
   'add by nickc 2006/08/14
   If UpForm Is Nothing Then
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           '  ·s¼W¸ê®Æ¨ì®×¥ó¶i«×ÀÉ
           strCP09 = Empty
           strCP09 = AutoNo("C", 6)
           strNCP09 = strCP09
           strCP10 = "1701"
           'Added by Lydia 2016/12/22
           strNcp10 = strCP10
           strNCP09 = strCP09
           'end 2016/12/22
           
           ' ¬O§_¦V«È¤á¦¬´Ú
           strCP20 = "N"
           If (m_TM10 = "020") And (IsEmptyText(textMoney) = False) And (Val(textMoney) <> 0) Then
              strCP20 = ""
           End If
           strCP27 = DBDATE(SystemDate())
           ' ¬O§_¶}¹q¸£¦¬¾Ú
           strCP32 = "N"
           If (m_TM10 = "020") And (IsEmptyText(textMoney) = False) And (Val(textMoney) <> 0) Then
              strCP32 = Empty
           End If
           
           ' ·í¥Ó½Ð°ê®a¬°¤j³°®É, ¤~»Ý¿é¤J¤j³°»âÃÒ¶O, §Y¶O¥Î
           If m_TM10 = "020" And IsEmptyText(textMoney) = False Then
              '©Ó¿ì¤H¬°¨Ï¥ÎªÌ, µo¤å¤é¬°¨t²Î¤é
              '·~°È°Ï´¼Åv¤H­û¬°³Ìªñ¦¬¤åAÃþ±µ¬¢°O¿ý³æ´¼Åv¤H­ûªº·~°È°Ï¤Î´¼Åv¤H­û
        'edit by nick 2004/10/20
        '      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP16,CP18,CP20,CP26,CP27,CP32) " & _
        '               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
        '                       "'" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
        '                       "" & textMoney & "," & Val(textMoney) / 1000 & ",'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "')"
              'Modify By Sindy 2010/7/12 ©Ó¿ì¤H§ï±¾¾Þ§@¤H­û old:GetCP14BYAClass(m_TM01, m_TM02, m_TM03, m_TM04)
              '2010/9/28 MODIFY BY SONIA §º­YÄõ»¡¦]´Á­­ºÞ¨îªí©Ó¿ì¤H·|±a¦¨µ{§Ç¬G§ï¬°¤´¦^­ì±±¨î,¦ýÂ÷Â¾±¾P2001°Ó¼Ð³B,©óGetCP14BYAClass±±¨î
              'modify by sonia 2019/1/30 +CP64
              'modify by sonia 2019/5/7 ¤j³°®×¤~¦sCP64 TC-010952
              'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP16,CP18,CP20,CP26,CP27,CP32,CP64) " & _
                       "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                               "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & GetCP14BYAClass(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                               "" & textMoney & "," & Val(textMoney) / 1000 & ",'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "','" & IIf(Option5(0).Value = True, "¹q¤lÃÒ®Ñ", "¯È¥»ÃÒ®Ñ") & "')"
              If Frame3.Visible = False Then
                 strCP64 = ""
              ElseIf Option5(0).Value = True Then
                 strCP64 = "¹q¤lÃÒ®Ñ"
              Else
                 strCP64 = "¯È¥»ÃÒ®Ñ"
              End If
              strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP16,CP18,CP20,CP26,CP27,CP32,CP64) " & _
                       "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                               "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & GetCP14BYAClass(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                               "" & textMoney & "," & Val(textMoney) / 1000 & ",'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "','" & strCP64 & "')"
              'end 2019/5/7
              cnnConnection.Execute strSql
           Else
              '©Ó¿ì¤H¬°¨Ï¥ÎªÌ, µo¤å¤é¬°¨t²Î¤é
              '·~°È°Ï´¼Åv¤H­û¬°³Ìªñ¦¬¤åAÃþ±µ¬¢°O¿ý³æ´¼Åv¤H­ûªº·~°È°Ï¤Î´¼Åv¤H­û
        'edit by nick 2004/10/20
        '      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32) " & _
        '               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
        '                       "'" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
        '                       "'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "')"
              'modify by sonia 2019/1/30 +CP64
              'modify by sonia 2019/5/7 ¤j³°®×¤~¦sCP64 TC-010952
              'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP64) " & _
              '         "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
              '                 "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & GetCP14BYAClass(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
              '                 "'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "','" & IIf(Option5(0).Value = True, "¹q¤lÃÒ®Ñ", "¯È¥»ÃÒ®Ñ") & "')"
              If Frame3.Visible = False Then
                 strCP64 = ""
              ElseIf Option5(0).Value = True Then
                 strCP64 = "¹q¤lÃÒ®Ñ"
              Else
                 strCP64 = "¯È¥»ÃÒ®Ñ"
              End If
              strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP64) " & _
                       "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                               "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & GetCP14BYAClass(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                               "'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "','" & strCP64 & "')"
              'end 2019/5/7
              cnnConnection.Execute strSql
           End If
           
        'add by nickc 2007/03/06 ¥Ó½Ð°ê®a¬O¥xÆW®É¡A±N715©Î717µo¤åªº¡A¤Wcp24='1'¡Acp25=¨Ó¨ç¦¬¤å¤é¡A¨Ã±N npªº 305 np06¤W Y
        If m_TM10 = "000" Then
            strSql = "update caseprogress set cp24='1' ,cp25=" & DBDATE(m_CP05) & " where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10 in ('715','717') and cp27 is not null "
            cnnConnection.Execute strSql
            'Modify by Amy 2015/07/01 ¤º°Ó¥Ó½Ð°ê¬°¥xÆW®É¸Ó®×¸¹¤U¤@µ{§Ç¬°¶Ê¼f305 ¥Bnp06¬OnullªÌ¡u¬O§_Äò¿ì¡v¥þ³¡¤WY(Äò¿ì)
            'strSql = "update nextprogress set np06='Y' where np06 is null and np07=305 and np01 in (select cp09 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10 in ('715','717') and cp27 is not null ) "
            Dim intR  As Integer
            strSql = "update nextprogress set np06='Y' where np06 is null and np07=305 And np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' "
            cnnConnection.Execute strSql, intR
            'Add By Sindy 2013/8/5
            '¤º°ÓªºT¥xÆW®×¤Î¥~°ÓFCT, ¦sÀÉ®É­Y¸Ó®×¸¹ªº¤U¤@µ{§ÇÀÉ¦³NP06 IS NULLªº 717(µù¥U¶O)´Á­­®É, ½Ð¤@¨Ö§ó·s.
            If m_TM01 = "T" Then
               strSql = "update nextprogress set np06='N',np11=" & strSrvDate(1) & ",NP12='10' " & _
                         "where np06 is null and np07='717' " & _
                           "and NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "'"
               cnnConnection.Execute strSql
            End If
            '2013/8/5 END
        End If
           
            'add by nick 2004/11/30  §ó·scÃþªº¥N²z¤H¤Î©¼©Ò®×¸¹¡A­n¦b·s¼WcÃþ¤§«á
            Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
           
           '2011/5/6 modify by sonia TF»â¤g©µ¦ù¤]¤£±¾Äò®i´Á­­,¦]¶Ê©µ®i©w½Z·|±a¥X©Ò¦³¤l®×(§t»â¤g©µ¦ù)ªº°ê®a¬G¤£¥²­«ÂÐ±¾´Á­­,¥À®×µo¤å,»â¤g©µ¦ùªº¤l®×¤]·|¤@¨Ö³B²z
           '2012/2/24 modify by sonia TFªº¤l®×¤]¤£±¾(¬ü°êTF-000058-1-03)
           'If m_TM01 <> "TC" Then   '91.11.3 ADD BY SONIA
           'Modified by Lydia 2017/02/02 ¥x-¤j³¡¥÷ºM¾P­«µoµù¥Uµý¤£ºÞ¨î´Á­­ => And str1006CP64 = ""
           If m_TM01 <> "TC" And m_TM04 = "00" And str1006CP64 = "" And Not (m_TM01 = "TF" And Right(m_TM02, 1) <> "0") Then '91.11.3 ADD BY SONIA
              strNP07 = "102"
              ' ªk©w´Á­­¬°±M¥Î´Á­­ºI¤î¤é
        'edit by nick 2004/11/17
        '      If m_TM10 < "010" Then
        '         If IsEmptyText(textTM22) = False Then: strNP09 = ChangeTStringToWString(textTM22)
        '      Else
              If IsEmptyText(textTM22) = False Then: strNP09 = textTM22
        'edit by nick 2004/11/17
        '      End If
              ' ¥»©Ò´Á­­¬°ªk©w´Á­­-2¤Ñ
                'Modify By Cheng 2003/09/01
        '      strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)), Val(DBDAY(strNP09)) - 2)))
              'edit by nickc 2007/06/13 TF §ï¦¨¤@­Ó¤ë
              If m_TM01 = "TF" Then
                  strNP08 = DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(strNP09))))
              Else
                  'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
                  If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                     strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
                  Else
                  '2014/10/6 END
                     'modify by sonia 2023/3/7 ¤j³°®×¤]§ï¬°2­Ó¤u§@¤Ñ
                     'strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
                     strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
                  End If
              End If
              strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
              
              'Modify By Cheng 2003/04/03
              '´¼Åv¤H­û¦s³Ìªñ¦¬¤åAÃþ±µ¬¢°O¿ý³æªº´¼Åv¤H­û
              'modify by sonia 2019/1/30 ¥ýÅª¬O§_¦s¦b¦A¨M©w­×§ï©Î§R°£
              'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
              '         "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
              '                 "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
              'cnnConnection.Execute strSql
               Set rsA = New ADODB.Recordset
               StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And np07=102 And np06 is null "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  strSql = "update NextProgress set np08=" & strNP08 & ",np09=" & strNP09 & " where " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And np07=102 And np06 is null "
                  cnnConnection.Execute strSql
               Else
                  strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                           "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
                           "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
                  cnnConnection.Execute strSql
               End If
          End If  '91.11.3 ADD BY SONIA
           
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           ' ­Y¦³¿é¤JÃº¦~¶OÄæ¦ì®É, ·s¼W¸ê®Æ¨ì¤U¤@µ{§Ç¸ê®ÆÀÉ, ¨Ã¦C¦L±µ¬¢µ²®×³æ
           If (m_TM01 = "TF") And (IsEmptyText(textDate) = False) Then
              strNP07 = "708"
              ' ªk©w´Á­­¬°¿é¤J¤§Ãº¦~¶O´Á­­
              strNP09 = DBDATE(textDate)
              ' ¥»©Ò´Á­­¬°ªk©w´Á­­-2¤Ñ
                'Modify By Cheng 2003/09/01
        '      strNP08 = DBDATE(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)), Val(DBDAY(strNP09)) - 2))
              strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
              strNP22 = GetNextProgressNo()
                'Modify By Cheng 2003/04/03
                '´¼Åv¤H­û¦s³Ìªñ¦¬¤åAÃþ±µ¬¢°O¿ý³æªº´¼Åv¤H­û
              strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                    "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
                            "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
              cnnConnection.Execute strSql
              ' ©µ®i, ¨Ï¥Î«Å»}, ¥Zµn¼s§i, Ãº¦~¶O, ¶Ê¼f, ´£¥Ó, ¦¬¹F¤£¦L±µ¬¢µ²®×³æ
              '92.6.8 SONIA ¥[ ¨¥µüÅG½×, ·Ç³Æµ{§Ç
           End If
           '2005/8/31 CANCEL BY SONIA
           ' ·í¥Ó½Ð°ê®a¬°¤j³°®É, ·s¼W¸ê®Æ¨ì¤U¤@µ{§ÇÀÉ, ¨Ã¦C¦L±µ¬¢µ²®×³æ
           'If m_TM01 <> "TC" And m_TM10 = "020" Then
           '   strNP07 = "702"
           '   ' ªk©w´Á­­¬°±M¥Î´Á­­°_¤é+3¦~
           '   If m_TM10 < "010" Then
           '      If IsEmptyText(textTM22) = False Then: strNP09 = DBDATE(textTM22)
           '   Else
           '      If IsEmptyText(textTM22) = False Then: strNP09 = textTM22
           '   End If
           '   ' ¥»©Ò´Á­­¬°ªk©w´Á­­-2¤Ñ
           '   strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
           '   strNP22 = GetNextProgressNo()
           '   strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
           '         "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
           '                 "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
           '   cnnConnection.Execute strSQL
           '   ' ªk©w´Á­­¬°±M¥Î´Á­­°_¤é+6¦~
           '   strNP09 = DBDATE(DateAdd("yyyy", 6, ChangeWStringToWDateString(DBDATE(strNP09))))
           '   ' ¥»©Ò´Á­­¬°ªk©w´Á­­-2¤Ñ
           '   strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
           '   strNP22 = GetNextProgressNo()
           '   strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
           '         "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
           '                 "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
           '   cnnConnection.Execute strSQL
           '   ' ªk©w´Á­­¬°±M¥Î´Á­­°_¤é+9¦~
           '   strNP09 = DBDATE(DateAdd("yyyy", 9, ChangeWStringToWDateString(DBDATE(strNP09))))
           '   ' ¥»©Ò´Á­­¬°ªk©w´Á­­-2¤Ñ
           '   strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
           '   strNP22 = GetNextProgressNo()
           '   strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
           '         "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
           '                 "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
           '   cnnConnection.Execute strSQL
           'End If
           '2005/8/31 END
           '93.6.11 ADD BY SONIA ±¾²Ä¤G´Áµù¥U¶O´Á­­
           'edit by nickc 2006/06/07 ­Y np ªº 717 ¤wµ²®×¡A´N¤£°µ¤U­±³o¬q
           'If m_TM01 = "T" And m_blnReceiveSecond = False And m_TM10 < "010" Then
           If m_TM01 = "T" And m_blnReceiveSecond = False And m_TM10 < "010" And Is717end = False Then
              'edit by nick  2004/12/21 ¥[¥Ó½Ð¤é¦b 92/11/28 «e¡A¥B¤½§i¤é¦b 92/9/1(§t)«á¡A­Y np ¨S¦³ 716 ´N·s¼W
              'If DBDATE(textTM21) > 20031128 Then
              If (DBDATE(textTM21) >= 20031128) Or (DBDATE(m_TM11) < 20031128 And DBDATE(textTM14) >= 20030901 And Trim(textTM14) <> "") Then
                  'Add By Sindy 2012/12/19 101¦~7¤ë°Ó¼Ð·s­×ªk¼o°£¤G´Áµù¥U¶OÃº¶O¨î«× +if
                  If Val(m_TM13) < 20120701 Then
                     'add by nick 2004/08/17
                     '¥ýÀË¬d¬O§_¦³ 717
                     StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='717' and cp05 is not null and cp57 is null "
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                        'add by nickc 2006/06/07 ­Y¦³ cp717 «h¸ònpªº 717¥Bµ²®×(¤£Äò¿ì¡B¸Ñ°£¨ä­­)ªº¬Û¦P©w½Z
                        Is717end = True
                     Else
                        Set rsA = New ADODB.Recordset
                        'ªk©w´Á­­
                        strCP07 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(textTM21)))))
                        '¥»©Ò´Á­­
                        'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
                        If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                           strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
                        Else
                        '2014/10/6 END
                           strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strCP07))))
                        End If
                        strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
                        StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='716' "
                        rsA.CursorLocation = adUseClient
                        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                        '­Y¦³¦¬¤å²Ä¤G´Áµù¥U¶O, §ó·s¶i«×ÀÉ
                        If rsA.RecordCount > 0 Then
                            StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
                            cnnConnection.Execute StrSQLa
                        '­Y¥¼¦¬¤å²Ä¤G´Áµù¥U¶O, ·s¼W¤U¤@µ{§ÇÀÉ
                        Else
                            'add by nick 2004/08/17
                            ' ÀË¬d¤U¤@µ{§Ç¦³µL 716
                            Set rsA = New ADODB.Recordset
                            StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And np07=716 "
                            rsA.CursorLocation = adUseClient
                            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                            If rsA.RecordCount > 0 Then
                                strSql = "update NextProgress set np08=" & DBDATE(strCP06) & ",np09=" & DBDATE(strCP07) & " where " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And np07=716 "
                                cnnConnection.Execute strSql
                            Else
                              If m_blnReceiveSecond = False Then '2011/9/22 add by sonia ­Y®×¥ó³Æµù¤w¦³¤£ºÞ¨î«h¤£·s¼W
                                 strNP07 = "716"
                                 strNP22 = GetNextProgressNo()
                                 strNP08 = DBDATE(strCP06) 'Add By Sindy 2009/10/23
                                 strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                                 "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                                                 DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
                                 cnnConnection.Execute strSql
                              End If  '2011/9/22 end
                            End If
                        End If
                        If rsA.State <> adStateClosed Then rsA.Close
                        Set rsA = Nothing
                        'add by nick 2004/08/17
                     End If
                  End If '2012/12/19 End
               End If
           End If
           
           '2005/4/14 ADD BY SONIA §ó·s»â¤g©µ¦ù¤§´£¥Ó¤é
           If m_TM01 = "TF" And Mid(m_TM02, 6, 1) <> "0" Then
              StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='104' "
              rsA.CursorLocation = adUseClient
              rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
              If rsA.RecordCount > 0 Then
                 If Not IsNull(textCP47) Then
                    strSql = "update CASEProgress set CP47=" & DBDATE(textCP47) & " where CP09='" & rsA.Fields("CP09") & "' "
                    cnnConnection.Execute strSql
                 End If
              End If
              If rsA.State <> adStateClosed Then rsA.Close
              Set rsA = Nothing
           End If
           '2005/4/14 END
           '93.6.11 END
           'add by nickc 2006/04/26 ¤j³°µoµù¥UÃÒ®ÉÀË¬d¦³µL¦¬¤å¥¼µo¤å»âÃÒµ{§Ç¡A¦³­n¤Wµo¤å¤H¡B®É¶¡¡B¤é
           If m_TM10 = "020" Then
              cnnConnection.Execute "update caseprogress set cp27=to_number(to_char(sysdate,'YYYYMMDD')) where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " and cp10='701' and cp27 is null "
           End If
           '2006/5/3 ADD BY SONIA °¨¼w¨½¬ü°ê®×±¾²Ä6¦~105¨Ï¥Î«Å»}
           'edit by nickc 2007/11/13 §ï¦¨ ¦³³]©wªº¡A³£­n
           'If m_TM01 = "TF" And m_TM10 = "101" Then
           '2013/10/9 MODIFY BY SONIA TF-000610-1-02
           'If m_TM01 = "TF" Then
           If m_TM01 = "TF" And m_TM04 = "00" And m_TM03 = "0" Then
                'add by nickc 2007/11/13 ²Ä2½Xªº²Ä6­Ó¦r¬O0ªº¥u­n§PÂ_«e 5 ¦r¬Û¦P¡A¥B²Ä4½X<>"00"
                '­Y²Ä2½Xªº²Ä6­Ó¦r<>0ªº¥u­n§PÂ_«e 6 ¦r¬Û¦P¡A¥B²Ä4½X<>"00"
                Dim MyTFrs As New ADODB.Recordset
                Set MyTFrs = New ADODB.Recordset
                If MyTFrs.State = 1 Then MyTFrs.Close
                MyTFrs.CursorLocation = adUseClient
                '2012/6/14 modify by sonia §ì¥¼³¬¨÷®×¸¹ and tm29 is null
                MyTFrs.Open "select * from trademark where tm01='" & m_TM01 & "' and tm04<>'00' and tm29 is null " & IIf(Mid(m_TM02, 6, 1) = "0", " and substr(tm02,1,5)='" & Mid(m_TM02, 1, 5) & "' ", " and tm02='" & m_TM02 & "' "), cnnConnection, adOpenStatic, adLockReadOnly
                If MyTFrs.RecordCount <> 0 Then
                    'edit by nickc 2007/11/13 ­ì¥u¦³¬ü°ê¦³±¾¨Ï¥Î«Å»}¡A²{§ï¦¨°ê®aÀÉ¦³±¾´N­n±¾
                    MyTFrs.MoveFirst
                    Do While Not MyTFrs.EOF
                        ' ¨ú±o¨Ï¥Î«Å»}¦~«×
                        strNA38 = 0
                        Set rsA = New ADODB.Recordset
                        Set rsA = Nothing
                        StrSQLa = "SELECT * FROM Nation WHERE NA01 = '" & CheckStr(MyTFrs.Fields("tm10")) & "' AND NA38 IS NOT NULL "
                        rsA.CursorLocation = adUseClient
                        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                        'edit by nickc 2007/11/13
                        'If rsA.RecordCount > 0 Then strNA38 = rsA.Fields("NA38")
                        If rsA.RecordCount > 0 Then
                            strNA38 = rsA.Fields("NA38")
                            If rsA.State <> adStateClosed Then rsA.Close
                            'ªk©w´Á­­  '2007/11/13 µù¸Ñ  ¨q¬Â»¡¤U¤@µ{§Ç¡A·s¼W¤l®×¸ê®Æ¡A´Á­­¥Ñ¥À®×©Î¬O»â¤g©µ¦ù¥»®×¨Ó­pºâ
                            '2012/6/14 modify by sonia ¬ü°ê®×¥H¬ü°ê¤§µù¥U¤½§i¤é­pºâ
                            'strCP07 = DBDATE(DateAdd("yyyy", Val(strNA38), ChangeWStringToWDateString(DBDATE(textTM21))))
                            'Modify By Sindy 2020/12/29 ¥HµoÃÒ¤é(TM20)¬°­pºâ´Á­­¤§°òÂ¦¤é
'¿é¤JTF®×®É:
'1. µe­±¤§µù¥U¤½§i¤é(TM14)Äæ§ï¬°µoÃÒ¤é(TM20)Äæ¡A¦sÀÉ§ó·sTM20(­ì§ó·sm_CP05¨Ó¨ç¦¬¤å¤é)¡F
'2. ¦sÀÉ®É²£¥Í105¨Ï¥Î«Å»}´Á­­¡G
'¡@¥HµoÃÒ¤é(TM20)¬°­pºâ´Á­­¤§°òÂ¦¤é¡F
'¡@¾¥¦è­ô104®×­pºâ«á¦A¥[¤T­Ó¤ë¬°ªk©w´Á­­¡A(¾¥¦è­ôµoÃÒ¤é¸¨¦b2018/8/10·í¤Ñ©Î¤§«áªÌ¡AºÞ¨î¤T¦~¨Ï¥Î«Å»}´Á­­¡A§Yµù¥U¤é°_º¡¤T¦~«á¤§¤T­Ó¤ë¤ºÀ³´£¥X¨Ï¥Î«Å»})
'¡@§PÂ_·s¼W©Î§ó·s¨Ï¥Î«Å»}´Á­­®É¡A¨ú®ø¤w¦¬¤åªº§PÂ_¡A¥uºÞ¤U¤@µ{§Ç¬O§_¦³105¨Ï¥Î«Å»}´Á­­¨Ó¨M©w·s¼W©Î§ó·s¡F¦ýµá«ß»«030®×¥u¯à·s¼W¤£¯à§ó·s¡A¦]¬°µá«ß»«ÁÙ¦³¥Ó½Ð¤é+3¦~ªº´Á­­¤£¯à»\±¼¡C
'°Ñ¦Ò¥~°ÓCF¤§µù¥UÃÒ / ©µ®iÃÒ®Ñ¿é¤Jfrm03010303_03
'                            If CheckStr(MyTFrs.Fields("tm10")) <> "101" Then
'                              strCP07 = DBDATE(DateAdd("yyyy", Val(strNA38), ChangeWStringToWDateString(DBDATE(textTM21))))
'                            Else
'                              strCP07 = DBDATE(DateAdd("yyyy", Val(strNA38), ChangeWStringToWDateString(DBDATE(textTM14))))
'                            End If
                            strCP07 = DBDATE(DateAdd("yyyy", Val(strNA38), ChangeWStringToWDateString(DBDATE(textTM20))))
                            '2012/6/14 end
                            'add by Sindy 2020/12/28 ¾¥¦è­ô®Ö­ã¤é´Á(§YµoÃÒ¤é©Îµù¥U¤é)¸¨¦b2018/8/10·í¤Ñ©Î¤§«áªÌ¡AºÞ¨î¤T¦~¨Ï¥Î«Å»}´Á­­¡A§Yµù¥U¤é°_º¡¤T¦~«á¤§¤T­Ó¤ë¤ºÀ³´£¥X¨Ï¥Î«Å»}
                            'modify by sonia 2023/9/15 110®ü¦a®×¬°¤­¦~¥[¤T­Ó¤ë
                            If CheckStr(MyTFrs.Fields("tm10")) = "104" Or CheckStr(MyTFrs.Fields("tm10")) = "110" Then
                              strCP07 = CompDate(1, 3, strCP07)
                            End If
                            'end  2020/12/28
                            '¥»©Ò´Á­­
                            'MODIFY BY SONIA 2014/4/28 °t¦XCFT,·~°È»¡§ï¦¨¥»©Ò=ªk©w-2­Ó¤ë ¤£ºÞ¥ô¦ó°ê®a
                            'strCP06 = DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(strCP07))))
                            strCP06 = CompDate(1, -2, strCP07)
                            strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
                            
                            'Modify By Sindy 2020/12/29 ¨ú®ø¤w¦¬¤åªº§PÂ_¡A¥uºÞ¤U¤@µ{§Ç¬O§_¦³105¨Ï¥Î«Å»}´Á­­¨Ó¨M©w·s¼W©Î§ó·s
'                            '¥ýÀË¬d¬O§_¤w¦¬¤å 105
'                            StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & " And CP10='105' AND CP27 IS NULL AND CP57 IS NULL"
'                            rsA.CursorLocation = adUseClient
'                            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                            '­Y¦³¦¬¤å¨Ï¥Î«Å»}, §ó·s¶i«×ÀÉ
'                            If rsA.RecordCount > 0 Then
'                                StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
'                                cnnConnection.Execute StrSQLa
'                            '­Y¥¼¦¬¤å¨Ï¥Î«Å»}, ·s¼W¤U¤@µ{§ÇÀÉ
'                            Else
                                ' ÀË¬d¤U¤@µ{§Ç¦³µL¨Ï¥Î«Å»}
                                Set rsA = New ADODB.Recordset
                                StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & _
                                          " And np07=105 AND NP06 IS NULL"
                                'Modify By Sindy 2020/12/29 µá«ß»«030®×¥u¯à·s¼W¤£¯à§ó·s¡A¦]¬°µá«ß»«ÁÙ¦³¥Ó½Ð¤é+3¦~ªº´Á­­¤£¯à»\±¼¡C
                                If (CheckStr(MyTFrs.Fields("tm10")) = "030" Or CheckStr(MyTFrs.Fields("tm10")) = "112") _
                                    And CheckStr(MyTFrs.Fields("tm10")) <> "102" Then
                                    StrSQLa = StrSQLa & " and np01='" & strCP09 & "'"
                                End If
                                '2020/12/29 END
                                rsA.CursorLocation = adUseClient
                                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                                If rsA.RecordCount > 0 Then
                                    strSql = "update NextProgress set NP01='" & strCP09 & "',np08=" & DBDATE(strCP06) & ",np09=" & DBDATE(strCP07) & " where " & ChgNextProgress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & _
                                             " And np07=105 And NP06 IS NULL"
                                    'Modify By Sindy 2020/12/29 µá«ß»«030®×¥u¯à·s¼W¤£¯à§ó·s¡A¦]¬°µá«ß»«ÁÙ¦³¥Ó½Ð¤é+3¦~ªº´Á­­¤£¯à»\±¼¡C
                                    If (CheckStr(MyTFrs.Fields("tm10")) = "030" Or CheckStr(MyTFrs.Fields("tm10")) = "112") _
                                        And CheckStr(MyTFrs.Fields("tm10")) <> "102" Then
                                        strSql = strSql & " and np01='" & strCP09 & "'"
                                    End If
                                    '2020/12/29 END
                                    cnnConnection.Execute strSql
                                Else
                                    '2007/11/13 µù¸Ñ  ¨q¬Â»¡¤U¤@µ{§Ç¡A·s¼W¤l®×¸ê®Æ¡A´¼Åv¤H­û±¾¥À®×©Î¬O»â¤g©µ¦ù¥»®×¦¬¤å¸¹¤]±¾¥À®×¨º¹D
                                    strNP07 = "105"
                                    strNP22 = GetNextProgressNo()
                                    strNP08 = DBDATE(strCP06) 'Add By Sindy 2009/10/23
                                    strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                            "VALUES ('" & strCP09 & "','" & CheckStr(MyTFrs.Fields("tm01")) & "','" & CheckStr(MyTFrs.Fields("tm02")) & "','" & CheckStr(MyTFrs.Fields("tm03")) & "','" & CheckStr(MyTFrs.Fields("tm04")) & "'," & strNP07 & "," & _
                                            DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
                                    cnnConnection.Execute strSql
                                End If
                                'add by sonia 2019/10/14 ´Á­­­Y¤w¹L´Á­n´£¿ôTF-00072(¾¥¦è­ô¤l®×TF-000720-1-05)
                                If DBDATE(strCP06) < Val(strSrvDate(1)) Then
                                   MsgBox "¤l®×" & MyTFrs.Fields("tm01") & "-" & Left(MyTFrs.Fields("tm02"), 5) & "-" & Right(MyTFrs.Fields("tm02"), 1) & "-" & MyTFrs.Fields("tm03") & "-" & MyTFrs.Fields("tm04") & " ¨Ï¥Î«Å»}´Á­­¤w¹L´Á, ½Ðª`·N!!!", vbExclamation + vbOKOnly
                                End If
                                'end 2019/10/14
'                            End If
                            If rsA.State <> adStateClosed Then rsA.Close
                            Set rsA = Nothing
                        End If
                        MyTFrs.MoveNext
                    Loop
                End If
           End If
           
           '2006/5/3 END
           'add by nickc 2005/04/22
           Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04
   End If
          
    'Added by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
    bolA1kdataMail = False
    m_ULD02 = "" ': m_AC2470 = "" 'Remove by Lydia 2017/04/06
    m_rA1k28 = "": m_rSpec = ""  'Added by Lydia 2017/04/06
    'Modified by Lydia 2017/01/04 ®×¥ó©Ê½è§ï¶Çµù¥UÃÒ1701
    'bolA1kdataMail = PUB_CheckA1kdataMail(m_TM01, m_TM02, m_TM03, m_TM04, m_TM44, m_CP09, m_CP10, m_CP13, strNCP09, m_ULD02, m_AC2470)
    'Modified by Lydia 2017/03/14 §ì³Ì·sªº´¼Åv¤H­û
    'bolA1kdataMail = PUB_CheckA1kdataMail(m_TM01, m_TM02, m_TM03, m_TM04, m_TM44, m_CP09, strCP10, m_CP13, strNCP09, m_ULD02, m_AC2470)
    'Modified by Lydia 2017/04/06 °Ï¤À½Ð´Ú¹ï¶H
    'bolA1kdataMail = PUB_CheckA1kdataMail(m_TM01, m_TM02, m_TM03, m_TM04, m_TM44, m_CP09, strCP10, PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), strNCP09, m_ULD02, m_AC2470)
    If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then 'Added by Lydia 2021/05/20 ¦¬´Ú±HÃÒ-­­MCTªº®×¥ó,©Ò¥H¥²¶·¥Ó½Ð°ê®a¬O¥xÆW; ex.T-166495
        'Modifeid by Lydia 2023/04/11 +¥Ó½Ð¤H1~5 +m_TM23 & "," & m_TM78 & "," & m_TM79 & "," & m_TM80 & "," & m_TM81
        bolA1kdataMail = PUB_CheckA1kdataMail(m_TM01, m_TM02, m_TM03, m_TM04, m_TM44, m_TM23 & "," & m_TM78 & "," & m_TM79 & "," & m_TM80 & "," & m_TM81, m_CP09, strCP10, PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), strNCP09, m_ULD02, m_rA1k28, m_rSpec)
        'end 2016/12/22
    End If 'Added by Lydia 2021/05/20
    
    Dim m_MonTM11 As String
    Dim m_MonTM14 As String
    Dim m_MonTM21 As String
    'add by nickc 2006/07/24
    If m_CP10 = "308" Then
      '·s¼W¤l®×®Ö­ã¨Ó¤å
      strCP09 = AutoNo("C", 6)
      strCP10 = "1001"
      strCP05 = DBDATE(UpForm.oStrCDate)
      strCP27 = DBDATE(SystemDate())
      ' ²Õ¦¨SQL»yªk
      strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14, CP26,CP27,CP43) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "')"
      ' ·s¼W¸ê®Æ¨ì¸ê®Æ®w
      cnnConnection.Execute strSql
      
      'Added by Morgan 2017/6/14 ¹q¤l¤½¤å
      If m_DocNo <> "" Then
         '§ó·s¾÷Ãö¤å¸¹
         strSql = "update caseprogress set cp08='" & m_DocWord & "¦r²Ä" & PUB_GetEDocNo(m_DocNo) & "¸¹' where cp09='" & strCP09 & "'"
         cnnConnection.Execute strSql, intI
         '½Æ»s¥À®×¤½¤å¹q¤lÀÉ
         strExc(0) = PUB_GetEDocFileName(m_TM01, m_TM02, m_TM03, m_TM04, "1001")
         SaveAttFile_PDF strCP09, m_DocPdf, strExc(0), Format(m_DocPdfDate), Format(m_DocPdfTime), False, , , True
      End If
      'end 2017/6/14
      
      '§ó·s¤l®×®Ö­ã¤Îµ²ªG¤é
      strSql = "update caseprogress set cp24='1',cp25=" & strCP05 & " where cp09='" & m_CP09 & "' "
      cnnConnection.Execute strSql
      '2011/9/20 ADD BY SONIA ¥À®×¤Î¤l®×ªº¶Ê¼f´Á­­¤WY
      strSql = "update nextprogress set np06='Y' where np01='" & m_CP09 & "' and np07='305' and np06 is null"
      cnnConnection.Execute strSql
      strSql = "update nextprogress set np06='Y' where np02='" & m_MonTM01 & "' and np03='" & m_MonTM02 & "' and np04='" & m_MonTM03 & "' and np05='" & m_MonTM04 & "' and np01='" & frm02010401_6.oKey & "' and np07='305' and np06 is null"
      cnnConnection.Execute strSql
      '¦P®É¤l®×ºÞ¨î©µ®i´Á­­
      'Modified by Lydia 2017/02/02 ¥x-¤j³¡¥÷ºM¾P­«µoµù¥Uµý¤£ºÞ¨î´Á­­ => And str1006CP64 = ""
      If m_TM01 <> "TC" And str1006CP64 = "" And Not (m_TM01 = "TF" And Right(m_TM02, 1) <> "0") Then
         strNP07 = "102"
         If IsEmptyText(textTM22) = False Then: strNP09 = textTM22
         If m_TM01 = "TF" Then
            strNP08 = DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(strNP09))))
         Else
            'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
            If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
               strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
            Else
            '2014/10/6 END
               strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
            End If
         End If
         strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
         If rsA.State <> adStateClosed Then rsA.Close
         StrSQLa = "select * from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='102' and cp27 is null and cp57 is null "
         Set rsA = New ADODB.Recordset
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount <> 0 Then
            strSql = "update caseprogress set cp06=" & strNP08 & ",cp07=" & strNP09 & " where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='102' and cp27 is null and cp57 is null "
         Else
            If rsA.State <> adStateClosed Then rsA.Close
            StrSQLa = "select * from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07='102' and np06 is null "
            Set rsA = New ADODB.Recordset
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount <> 0 Then
               strSql = "update nextprogress set np08=" & strNP08 & ",np09=" & strNP09 & " where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07='102' and np06 is null "
            Else
               strNP22 = GetNextProgressNo()
               strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                        "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
                                "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
            End If
         End If
         cnnConnection.Execute strSql
      End If
      '2011/9/20 END
      
'2011/9/22 modify by sonia «e¤w§ì¥À®×¬O§_ºÞ¨î²Ä¤G´Á,¬G§ï¥Hm_blnReceiveSecond§PÂ_
'      '¥À®×¦³¦¬ 717 ®É¡A¤£ºÞ¡A­Y¦³ 716 ªº¤]¤£ºÞ¡A¥u¦³ 715 ªº ¤l®×­n±¾²Ä¤G´Áµù¥U¶O ¡A¦ý¶È­­´Á°_¤é+3¦~-1¤Ñ ¤j©ó ¨t²Î¤éªº¤~°µ
'      If rsA.State <> adStateClosed Then rsA.Close
'      m_MonTM11 = ""
'      m_MonTM14 = ""
'      m_MonTM21 = ""
'      StrSQLa = "select * from trademark where tm01='" & m_MonTM01 & "' and tm02='" & m_MonTM02 & "' and tm03='" & m_MonTM03 & "' and tm04='" & m_MonTM04 & "' "
'      Set rsA = New ADODB.Recordset
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         m_MonTM11 = CheckStr(rsA.Fields("tm11"))
'         m_MonTM14 = CheckStr(rsA.Fields("tm14"))
'         m_MonTM21 = CheckStr(rsA.Fields("tm21"))
'      End If
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      If (m_MonTM21 >= 20031128) Or (m_MonTM11 < 20031128 And m_MonTM14 >= 20030901 And m_MonTM14 <> "") Then
'        If ChangeWDateStringToWString(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(m_MonTM21)))) <= strSrvDate(1) Then
'            If rsA.State <> adStateClosed Then rsA.Close
'            StrSQLa = "select * from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10 in ('716','717') "
'            Set rsA = New ADODB.Recordset
'            rsA.CursorLocation = adUseClient
'            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount = 0 Then
'               If rsA.State <> adStateClosed Then rsA.Close
'               StrSQLa = "select * from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='715' "
'               Set rsA = New ADODB.Recordset
'               rsA.CursorLocation = adUseClient
'               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'               If rsA.RecordCount <> 0 Then
               'Modify By Sindy 2012/12/19 101¦~7¤ë°Ó¼Ð·s­×ªk¼o°£¤G´Áµù¥U¶OÃº¶O¨î«× +And Val(m_TM13) < 20120701
               If m_blnReceiveSecond = False And m_TM10 = "000" And Val(m_TM13) < 20120701 Then
                  '­n±¾²Ä¤G´Áªº´Á­­µ¹¤l®×
                  'ªk©w´Á­­
                  strCP07 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(m_MonTM21))))
                  '¥»©Ò´Á­­
                  'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
                  If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                     strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
                  Else
                  '2014/10/6 END
                     strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strCP07))))
                  End If
                  strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
                  strNP07 = "716"
                  strNP22 = GetNextProgressNo()
                  strNP08 = DBDATE(strCP06) 'Add By Sindy 2009/10/23
                  strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                           "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                           DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
                  cnnConnection.Execute strSql
               End If
'               If rsA.State <> adStateClosed Then rsA.Close
'            Else
               'add by nickc 2007/03/06 ¥Ó½Ð°ê®a¬O¥xÆW®É¡A±N715©Î717µo¤åªº¡A¤Wcp24='1'¡Acp25=¨Ó¨ç¦¬¤å¤é¡A¨Ã±N npªº 305 np06¤W Y
               If m_TM10 = "000" Then
                   strSql = "update caseprogress set cp24='1' ,cp25=" & strCP05 & " where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and  cp10 in ('715','717') and cp27 is not null "
                   cnnConnection.Execute strSql
                   strSql = "update nextprogress set np06='Y' where np06 is null and np07=305 and np01 in (select cp09 from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and  cp10 in ('715','717') and cp27 is not null ) "
                   cnnConnection.Execute strSql
               End If
'            End If
'        End If
'      End If

      '¦³´Á­­®É
      If textNP08.Enabled = True And textNP09.Enabled = True Then
             '­Yµe­±¦³¿é¤J·s´Á­­¥H·s´Á­­¬°¥D¡A¨S¦³ªº¸Ü±NÄ~©Ó¥À®×´Á­­
             If Trim(textNP08) <> "" And Trim(textNP09) <> "" Then
                If UpForm.IsHaveNp202 Then
                      strNP22 = GetNextProgressNo() 'Add By Sindy 2009/10/23
                      strNP08 = DBDATE(textNP08) 'Add By Sindy 2009/10/23
                      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                          "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',202," & _
                          DBDATE(textNP08) & "," & DBDATE(textNP09) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
                      cnnConnection.Execute strSql
                ElseIf UpForm.IsHaveCp202 Then
                     If Trim(textNP08) <> "" Then
                         strSql = "update caseprogress set cp06=" & DBDATE(textNP08) & ",cp07=" & DBDATE(textNP09) & ",cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     Else
                         strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     End If
                     cnnConnection.Execute strSql
                     strSql = "update caseprogress set cp43='" & m_CP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     cnnConnection.Execute strSql
                     strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     cnnConnection.Execute strSql
                     'Add by Sonia 2013/8/8 ¦P®É§ó¥¿ACC0J0ªºT-184230,¤£¥i§ó·s¤w¦¬´Ú¶Ç²¼ªº®×¸¹,¦]¬°¤À³Î»P¥Ó½Ð·N¨£®Ñªº®×¸¹¦]¤W­z»yªk¦Ó¤£¦P
                     strSql = "update acc0j0 set a0j02='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where a0j01 in (select cp09 from caseprogress where cp27 is null and cp57 is null and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='202') "
                     cnnConnection.Execute strSql
                     'end 2013/8/8
               End If
             Else
                If UpForm.IsHaveNp202 Then
                      strNP22 = GetNextProgressNo() 'Add By Sindy 2009/10/23
                      strNP08 = m_MonNP08 'Add By Sindy 2009/10/23
                      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                          "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',202," & _
                          m_MonNP08 & "," & m_MonNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
                      cnnConnection.Execute strSql
                ElseIf UpForm.IsHaveCp202 Then
                     If Trim(textNP08) <> "" Then
                         strSql = "update caseprogress set cp06=" & DBDATE(textNP08) & ",cp07=" & DBDATE(textNP09) & ",cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     Else
                         strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     End If
                     cnnConnection.Execute strSql
                     strSql = "update caseprogress set cp43='" & m_CP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     cnnConnection.Execute strSql
                     strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     cnnConnection.Execute strSql
                     'Add by Sonia 2013/8/8 ¦P®É§ó¥¿ACC0J0ªºT-184230,¤£¥i§ó·s¤w¦¬´Ú¶Ç²¼ªº®×¸¹,¦]¬°¤À³Î»P¥Ó½Ð·N¨£®Ñªº®×¸¹¦]¤W­z»yªk¦Ó¤£¦P
                     strSql = "update acc0j0 set a0j02='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where a0j01 in (select cp09 from caseprogress where cp27 is null and cp57 is null and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='202') "
                     cnnConnection.Execute strSql
                     'end 2013/8/8
                End If
             End If
             If UpForm.IsHaveNp202 Then
                  strSql = "update nextprogress set np06='N',np15=np15||'Âà¤J¤l®×¡A¤l®×®×¸¹¡G" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' where np02='" & m_MonTM01 & "' and np03='" & m_MonTM02 & "' and np04='" & m_MonTM03 & "' and np05='" & m_MonTM04 & "' and np06 is null and np07=202 "
                  cnnConnection.Execute strSql
             ElseIf UpForm.IsHaveCp202 Then
                  strSql = "update caseprogress set cp57=to_number(to_char(sysdate,'YYYYMMDD')),cp64=cp64||'Âà¤J¤l®×¡A¤l®×®×¸¹¡G" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202' and cp27 is null "
                  cnnConnection.Execute strSql
             End If
             '¥À®×¤À³Îµo¤å«áªº¦¬¤å¤Îµo¤å®×¥ó¬ÒÂà¤J¦³´Á­­ªº¤l®×
             Dim m_MonCP27 As String
             strSql = "select cp27 from caseprogress where cp09='" & m_MonCP09 & "' "
             m_MonCP27 = ""
             Set rsTmp = New ADODB.Recordset
             If rsTmp.State = 1 Then rsTmp.Close
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
             If rsTmp.RecordCount > 0 Then
                 m_MonCP27 = CheckStr(rsTmp.Fields("cp27"))
             End If
             If m_MonCP27 <> "" Then
                 strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & "' where cp05>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                 cnnConnection.Execute strSql
                 strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & "' where cp27>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                 cnnConnection.Execute strSql
                 
                 strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp05>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                 cnnConnection.Execute strSql
                 'Add by Sonia 2013/8/8 ¦P®É§ó¥¿ACC0J0ªºT-184230,¤£¥i§ó·s¤w¦¬´Ú¶Ç²¼ªº®×¸¹,¦]¬°¤À³Î»P¥Ó½Ð·N¨£®Ñªº®×¸¹¦]¤W­z»yªk¦Ó¤£¦P
                 strSql = "update acc0j0 set a0j02='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where a0j01 in (select cp09 from caseprogress where cp05>" & m_MonCP27 & " and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10<>'1001') "
                 cnnConnection.Execute strSql
                 'end 2013/8/8
                 
                 strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                 cnnConnection.Execute strSql
                 'Add by Sonia 2013/8/8 ¦P®É§ó¥¿ACC0J0ªºT-184230,¤£¥i§ó·s¤w¦¬´Ú¶Ç²¼ªº®×¸¹,¦]¬°¤À³Î»P¥Ó½Ð·N¨£®Ñªº®×¸¹¦]¤W­z»yªk¦Ó¤£¦P
                 strSql = "update acc0j0 set a0j02='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where a0j01 in (select cp09 from caseprogress where cp27>" & m_MonCP27 & " and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10<>'1001') "
                 cnnConnection.Execute strSql
                 'end 2013/8/8
             End If
      End If
      '2008/10/24 ADD BY SONIA ¤À³Î¥À®×³¬¨÷
      Set rsA = New ADODB.Recordset
      If rsA.State = 1 Then rsA.Close
      strSql = "select * from divisioncase,trademark where dc05='" & m_MonTM01 & "' and dc06='" & m_MonTM02 & "' and dc07='" & m_MonTM03 & "' and dc08='" & m_MonTM04 & "' and dc01=tm01(+) and dc02=tm02(+) and dc03=tm03(+) and dc04=tm04(+) and (tm16 is null or tm16='') "
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount = 0 Then
         strSql = "update trademark set tm29='Y',tm30=to_number(to_char(sysdate,'YYYYMMDD')),tm31='87' where tm01='" & m_MonTM01 & "' and tm02='" & m_MonTM02 & "' and tm03='" & m_MonTM03 & "' and tm04='" & m_MonTM04 & "' and (tm29 is null or tm29='') "
         cnnConnection.Execute strSql
      End If
      If rsA.State = 1 Then rsA.Close
      '2008/10/24 END
    
    'Added by Morgan 2023/1/16 ¹q¤l¤½¤å
    ElseIf m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
      
      'Added by Morgan 2025/2/18
      If m_TM136 = "" Then
         strSql = "UPDATE TradeMark set TM136='1'" & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "'"
         cnnConnection.Execute strSql, intI
      End If
      'end 2025/2/18
    'end 2023/1/16
    End If
    
   'Add By Sindy 2019/12/19 °Ó¼Ð¹q¤l¤Æ
   If strSrvDate(1) >= T°Ó¼Ð¹q¤l¤Æ²Ä2¶¥¬q±Ò¥Î¤é Then
      strLD18 = strCP09
      strExc(1) = ""
      If m_TM10 <> "000" Then '¬°¥x->¤j
         strExc(1) = Pub_GetSpecMan("¤º°Óµ{§Ç«È¤á¨çµo«á¸É¬Ý¤H­û")
      End If
      PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), , False, m_TM23, strCP10, m_TM44, , , , , strExc(1)
   End If
   '2019/12/19 END
   
   'Add By Sindy 2009/09/24
   '¦]¬°¦³¨Ç¨Ó¨ç¥Ñ¤º°Ó¿é¤J¡A¤º°Ó¦³¦Û¦æ±±ºÞ¤§©Ó¿ì´Á­­¤Îµo¤å¤é¡C§ï¬°¤º°Ó¿é¤J©Ò¦³CÃþ¨Ó¨ç¡A
   '­Y·~°È°Ï¬°F¦rÀYªÌ¡A°£ª§Ä³¨ü²z¥~¡A¦Û°Ê²£¥ÍBÃþ¦¬¤å¡A®×¥ó©Ê½è¬°¥~°Óµo¤å722¡A¤£¤Wµo¤å¤é¡A¤£¦V«È¤á½Ð´Ú
   Dim strCP48 As String, strCP09B As String
   If Left(GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)), 1) = "F" And _
      ((m_TM01 = "T" And m_TM10 = "020") Or (m_TM01 = "FCT" And m_TM10 = "000")) Then
      strCP09B = AutoNo("B", 6)
      '©Ó¿ì´Á­­¬°¨t²Î¤é¥[4­Ó¤u§@¤Ñ
      strCP48 = DBDATE(Pub_GetHandleDay(m_TM01, m_TM10, "722", strSrvDate(1), , m_CP09))
      '2011/4/28 modify by sonia ´¼Åv¤H­û­ì§ìÂI¿ï¦¬¤å¸¹¤§´¼Åv¤H­û,§ï§ì¸Ó®×³Ì«á¦¬¤å¦bÂ¾´¼Åv¤H­û
      strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp48,cp20,cp26,cp32,cp43) " & _
                     "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
                     "," & CNULL(m_TM04) & "," & CNULL(strSrvDate(1)) & "," & CNULL(strCP09B) & ",722," & _
                     CNULL(GetSalesArea(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "," & CNULL(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "," & CNULL(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "," & CNULL(strCP48) & ",'N','N','N'," & CNULL(strCP09) & ")"
      cnnConnection.Execute strSql
   End If
      
   'Add By Sindy 2010/01/13 ­Y¥¼µo,®Ö­ã¼f©w®Ñ®É,´£¿ô¬O§_ºÞ¨î²Ä¤G´Áµù¥U¶O
   '2011/9/19 modify by sonia ¥[¤J®×¥ó©Ê½è§PÂ_,§_«hµù¥U«á¤À³Î®Ö­ã¤]·|¶]¦¹¬q
   'If m_TM01 = "T" And m_TM10 = "000" Then
   'modify by sonia 101¦~7¤ë°Ó¼Ð·s­×ªk¼o°£¤G´Áµù¥U¶OÃº¶O¨î«× +And Val(m_TM13) < 20120701
   If m_TM01 = "T" And m_TM10 = "000" And m_CP10 <> "308" And Val(m_TM13) < 20120701 Then
      StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " and CP10='101' and CP09 in (Select CP43 From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " and (CP10='1001' or (CP10='1403' and cp24='1'))) "
      If rsA.State <> adStateClosed Then rsA.Close 'Add By Sindy 2019/5/28
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount = 0 Then
         nResponse = MsgBox("¦¹®×©|¥¼µo¹L¡u®Ö­ã¼f©w®Ñ¡v¬O§_­nºÞ¨î²Ä¤G´Áµù¥U¶O¡H", vbYesNo + vbCritical + vbDefaultButton2, "¸ß°Ý")
         If nResponse = vbYes Then
            m_blnReceiveSecond = False '¥¼¦¬²Ä¤G´Áµù¥U¶O
            '­n±¾²Ä¤G´Áµù¥U¶O
            'ªk©w´Á­­
            strCP07 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(textTM21)))))
            '¥»©Ò´Á­­
            'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
            If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
               strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
            Else
            '2014/10/6 END
               strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strCP07))))
            End If
            strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
            strNP07 = "716"
            strNP22 = GetNextProgressNo()
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                            "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
            cnnConnection.Execute strSql
            '¥[¦L¦^ÂÐ³æ
            Call g_PrtForm001.PrintReturnSheet(strCP09, strNP07, DBDATE(strCP07), , , , , m_TM01 & m_TM02 & m_TM03 & m_TM04)
            '¥[¦L®×¥ó±µ¬¢µ²®×³æ
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   '2010/01/13 End
   
   '2010/3/26 ADD BY SONIA T-140541µoµù¥UÃÒ®É§ó·s¥Ó½Ðªº¶Ê¼f¬°Y,¥Ó½Ð¬°®Ö­ã,¥H§KµL®Ö­ã³qª¾®É¤´¥h¶Ê¼f
   '2012/10/16 MODIFY BY SONIA TC-010630¤]­n§ó·s806µÛ§@Åvµn°O¬°®Ö­ã,¨ä¶Ê¼f¬°Y
   'cnnConnection.Execute "update caseprogress set cp24='1' where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " and cp10='101' and cp27 is NOT null AND CP24 IS NULL "
   'cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP06='Y' WHERE " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " AND NP07=305 AND NP06 IS NULL AND NP01 IN (SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " and cp10='101' and cp27 is NOT null) "
   cnnConnection.Execute "update caseprogress set cp24='1' where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " and cp10 IN ('101','806') and cp27 is NOT null AND CP24 IS NULL "
   cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP06='Y' WHERE " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " AND NP07=305 AND NP06 IS NULL AND NP01 IN (SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " and cp10 IN ('101','806') and cp27 is NOT null) "
   '2010/3/26 END
   
   'Add By Sindy 2013/9/16 ¥Ó½Ð¤H¬°X13175010¤u¬ã°|ªÌ¥B¦³±M¥Î´Á¶¡ªÌ³]©w¬°¤£¶Ê©µ®i
   If (m_TM01 = "T" Or m_TM01 = "TF") And _
      (m_TM23 = "X13175010" Or m_TM78 = "X13175010" Or m_TM79 = "X13175010" Or m_TM80 = "X13175010" Or m_TM81 = "X13175010") And _
      Val(textTM21) > 0 And _
      Val(textTM22) > 0 Then
      strSql = "update trademark set" & _
               " tm129='Y'" & _
               " where TM01='" & m_TM01 & "' and TM02='" & m_TM02 & _
                "' and TM03='" & m_TM03 & "' and TM04='" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   '2013/9/16 END
   
   'Add by Sindy 2019/5/10
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) 'ÀË¬d¬O§_¦³¹q¤lÀÉ­n¦s¤J¨÷©v°Ï
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010404_1", strCP09
   End If
   '2019/5/10 END
   
   'add by nickc 2006/08/14
   If Me.Visible = True Then
       'Add By Cheng 2002/11/07
       cnnConnection.CommitTrans
   End If
   Exit Function
ErrorHandler:
    'add by nickc 2006/08/14
    If Me.Visible = True Then
        cnnConnection.RollbackTrans
    End If
    OnSaveData = False
    'Resume Next
End Function

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 §ï¾ã§å¦L
'    'add by nickc 2006/10/02
'    If UpForm Is Nothing Then
'        PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'        '§R°£¼È¦s¸ê®Æ
'        PUB_DeleteCaseCloseSheet strUserNum
'    End If
   
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   Set frm02010404_3 = Nothing
End Sub

'add by sonia 2019/2/1
Private Sub Option5_Click(Index As Integer)
   If Me.Option5(0).Value Then
      textPS = "ªþ¥ó¡G°Ó¼Ð¹q¤lµù¥UÃÒ¤§¯È¥»¡C"
   ElseIf Me.Option5(1).Value Then
      textPS = "ªþ¥ó¡Gµù¥UÃÒ¥¿¥»¤A¯È¡C"
   End If
End Sub
'end 2019/2/1

'Private Sub Text1_GotFocus()
'   InverseTextBox Text1
'End Sub
'
''2005/11/11 ADD BY SONIA
'Private Sub Text1_Validate(Cancel As Boolean)
'   If m_strLanguage = "2" And Text1 <> "" Then
'      If CheckIsTaiwanDate(Text1) = False Then
'         Cancel = True
'         Text1_GotFocus
'      End If
'   End If
'End Sub
''2005/11/11 END

' Ãº¦~¶O´Á­­
Private Sub textDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textDate) = False Then
      If CheckIsTaiwanDate(textDate) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½TªºÃº¦~¶O´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_GotFocus
      End If
   End If
End Sub

Private Sub textEditPrint_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 89 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

''Add By Sindy 2020/12/14
'Private Sub textFinalDate_GotFocus()
'    TextInverse Me.textFinalDate
'End Sub
'Private Sub textFinalDate_Validate(Cancel As Boolean)
'Dim strTit As String
'Dim strMsg As String
'Dim nResponse
'
'   Cancel = False
'   If IsEmptyText(textFinalDate) = False Then
'      If CheckIsDate(textFinalDate, False) = False Then
'          Cancel = True
'          strTit = "¸ê®ÆÀË®Ö"
'          strMsg = "½Ð¿é¤J¦è¤¸¦~¤ë¤é"
'          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      ElseIf ChkWork(ChangeTStringToWString(Val(textFinalDate) - 19110000)) = False Then
'          Cancel = True
'      ElseIf Val(Me.textFinalDate.Text) < Val(strSrvDate(1)) Then
'          Cancel = True
'          strTit = "¸ê®ÆÀË®Ö"
'          strMsg = "©w½Z¤é´Á­n¤j©óµ¥©ó¨t²Î¤é"
'          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      End If
'   End If
'   If Cancel Then TextInverse textFinalDate
'End Sub
''2020/12/14 END

' ¤j³°»âÃÒ¶O
Private Sub textMoney_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   '2005/9/29 ADD BY SONIA
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   If IsEmptyText(textMoney) = False Then
      If IsNumeric(textMoney) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªº¤j³°»âÃÒ¶O"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textMoney_GotFocus
      End If
      '2005/9/29 ADD BY SONIA
      StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 ='701' AND CP57 IS NULL AND CP16 > 0"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "¦¹®×¤w¦¬¤å»âÃÒ, ¤£¥i¦A¿é¤J¤j³°»âÃÒ¶O"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textMoney_GotFocus
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      '2005/9/29 END
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' ¦C¦L©w½Z
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         'edit by nickc 2006/06/29
         'Case " ", "N":
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            'edit by nickc 2006/06/29
            'strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
            strMsg = "¥u¥i¿é¤J N ©Î 1-3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' ¦C¦L³Æµù
Private Sub textPS_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textPS, 128) = False Then
      Cancel = True
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "¦C¦L³ÆµùÄæ¦ì¤º®e¤Óªø"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPS_GotFocus
   End If
End Sub

Private Sub textTM14_GotFocus()
    TextInverse Me.textTM14
End Sub

Private Sub textTM14_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    If IsEmptyText(textTM14) = False Then
        ' ¥Ó½Ð°ê®a¬°»Ý¿é¤J¥Á°ê¦~, §_«h¿é¤J¦è¤¸¦~
        'edit by nick 2004/10/06
'        If m_TM10 < "010" Then
'            If CheckIsTaiwanDate(textTM14, False) = False Then
'                Cancel = True
'                strTit = "¸ê®ÆÀË®Ö"
'                strMsg = "¥Ó½Ð°ê®a¬°¥xÆW, ½Ð¿é¤J¥Á°ê¦~"
'                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            End If
'        Else
            If CheckIsDate(textTM14, False) = False Then
                Cancel = True
                strTit = "¸ê®ÆÀË®Ö"
'                strMsg = "¥Ó½Ð°ê®a«D¥xÆW, ½Ð¿é¤J¦è¤¸¦~"
                strMsg = "½Ð¿é¤J¦è¤¸¦~"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            End If
'        End If

         'Added by Lydia 2023/03/29 ¨ó§U±±ºÞ°w¹ï¥xÆWµù¥UÃÒ¿é¤J¡A¤½§i¤é´Á¥u¯à¿é¤J1¸¹©Î16¸¹
         If m_TM01 = "T" And m_TM10 = "000" And InStr("01,16,", Format(PUB_DBDAY(textTM14), "00")) = 0 Then
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¤½§i¤é´Á¥u¯à¿é¤J1¸¹©Î16¸¹"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
         'end 2023/03/29
    End If
    'Add By Cheng 2003/11/20
    '­Y¦³¿é¤J¤½§i¤é
    If Me.textTM14.Text <> "" And Cancel = False Then
        'edit by nick 2004/10/06
        'm_TM14 = TAIWANDATE(Me.textTM14.Text)
        m_TM14 = DBDATE(Me.textTM14.Text)
        If m_TM01 = "T" Then
            '­Y¤½§i¤é¦b920816(§t)¥H«eªÌ©Î¤j³°®×
            'edit by nick 2004/10/06
            'If Val(m_TM14) <= 920816 Or m_TM10 = "020" Then
            If Val(m_TM14) <= 20030816 Or m_TM10 = "020" Then
                  '«D°¨¼w¨½®×±M¥Î´Á¶¡°_¤é¬°¤½§i¤é+¤T­Ó¤ë
                  'edit by nick 2004/10/06
                  'm_TM21 = TAIWANDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_TM14))))
                  m_TM21 = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_TM14))))
                  Select Case m_TM08
                     'modify by sonia 2013/11/27 ¥[9¹ÎÅé°Ó¼Ð
                     Case "1", "4", "7", "8", "9":
                        '±M¥Î´Á¶¡¤î¤é¬°¤½§i¤é+¤T­Ó¤ë°_¤Q¦~´î¤@¤Ñ
                        'edit by nick 2004/10/06
                        'm_TM22 = TAIWANDATE(DateAdd("d", -1, DateAdd("yyyy", 10, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_TM14))))))
                        m_TM22 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 10, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_TM14))))))
                     Case Else
                        strExc(0) = "SELECT TM22 FROM TRADEMARK WHERE TM15 = '" & m_TM27 & "' "
                        intI = 1
                        'edit by nickc 2007/02/06 ¤£¥Î dll ¤F
                        'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If Not IsNull(RsTemp.Fields("TM22")) Then
                              'edit by nick 2004/10/06
                              'm_TM22 = TransDate(rsTemp.Fields("TM22"), 1)
                              m_TM22 = TransDate(RsTemp.Fields("TM22"), 2)
                           End If
                        End If
                  End Select
                  '2008/5/15 add by sonia ¤j³°¤½§i¤é2007/12/1¥H«á,±M¥Î´Á°_¤é¬°¤½§i¤é+3­Ó¤ë+1¤Ñ,±M¥Î´Á¤î¤é¬°¤½§i¤é+3­Ó¤ë+10¦~
                  If Val(m_TM14) >= 20071201 And m_TM10 = "020" Then
                     m_TM21 = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM21))))
                     m_TM22 = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
                  End If
                  '2008/5/15 end
            '­Y¤½§i¤é¤j©ó920816ªÌ¥B«D¤j³°®×
            Else
                '«D°¨¼w¨½®×±M¥Î´Á¶¡°_¤é¬°¤½§i¤é
                m_TM21 = m_TM14
                '±M¥Î´Á¶¡¤î¤é¬°¤½§i¤é¥[¤Q¦~´î¤@¤Ñ
                'edit by nick 2004/10/06
                'm_TM22 = TAIWANDATE(DateAdd("d", -1, DateAdd("yyyy", 10, ChangeWStringToWDateString(DBDATE(m_TM14)))))
                'Modified by Lydia 2019/11/13 §ï¥Î¦@¥Î¼Ò²Õ
                'm_TM22 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 10, ChangeWStringToWDateString(DBDATE(m_TM14)))))
                'Modify By Sindy 2022/3/7 + m_TM10 : ©µ®i«á¤§±M¥Î´Á­­¦~«×­Õ¦³2¤ë29¤é®É¡A±M¥Î´Á­­¤î¤éÀ³¬°2¤ë29¤é¡A¦Ó«D¥H¥[10¦~¤§¤è¦¡­pºâ¬°2¤ë28¤é
                m_TM22 = PUB_GetEndDate(DBDATE(m_TM14), 10, m_NA85, m_TM10)
            End If
        ElseIf m_TM01 = "TF" Then
           Dim strKey(0 To 4) As String, strTmp As String
           strKey(0) = m_CP09
           strKey(1) = m_TM01
           strKey(2) = m_TM02
           strKey(3) = m_TM03
           strKey(4) = m_TM04
           If TFGetMoneyDate(m_TM10, strKey, m_TM21, strTmp, m_TM22) Then
               'edit by nick 2004/10/06
'              m_TM21 = TransDate(m_TM21, 1)
'              m_TM22 = TransDate(CompDate(2, -1, m_TM22), 1)
              'Remove by Lydia 2019/12/09 ¸g¹L¾ã²z,¥HTM21_Validate°_¤éªººâªk¬°·Ç
              'm_TM21 = TransDate(m_TM21, 2)
              'm_TM22 = TransDate(CompDate(2, -1, m_TM22), 2)
              'end 2019/12/09
            End If
        End If
    End If
    If Cancel Then TextInverse textTM14
    '2006/1/24 ADD BY SONIA ¥xÆW®×¦Û°Ê±a¥X±M¥Î´Á¶¡
    If m_TM01 = "T" And m_TM10 = "000" Then
      '2009/1/14 modify by sonia ¤À³Î®×¤£¥i
      'textTM21 = m_TM21
      'textTM22 = m_TM22
      If m_CP10 <> "308" Then
         textTM21 = m_TM21
         textTM22 = m_TM22
      Else
         m_TM21 = textTM21
         m_TM22 = textTM22
      End If
      '2009/1/14 end
    End If
    '2006/1/24 END

End Sub
'2005/4/14 ADD BY SONIA
Private Sub textCP47_GotFocus()
    TextInverse Me.textCP47
End Sub

Private Sub textCP47_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP47) = False Then
      If CheckIsDate(textCP47, False) = False Then
          Cancel = True
          strTit = "¸ê®ÆÀË®Ö"
          strMsg = "½Ð¿é¤J¦è¤¸¦~¤ë¤é"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel Then TextInverse textCP47

End Sub
'2005/4/14 END
Private Sub textTM15_GotFocus()
    TextInverse Me.textTM15
End Sub

'Add By Sindy 2010/9/1
Private Sub textTM15_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textTM15) = False Then
      'ÀË¬d¼f©w¸¹©Ò¿é¤Jªºªø«×¬O§_¥¿½T
      'Add By Sindy 2017/5/17 + strRetrunText
      If PUB_ChkTm12Tm15Length("2", textTM15, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
         Cancel = True
         textTM15_GotFocus
         Exit Sub
      'Add By Sindy 2017/5/17
      Else
         textTM15 = strRetrunText
      '2017/5/17 END
      End If
   End If
End Sub

Private Sub textTM20_GotFocus()
   InverseTextBox textTM20
End Sub

' µù¥U¤é
Private Sub textTM20_Validate(Cancel As Boolean)
   Dim SysDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM20) = False Then
      ' ÀË¬d¬O§_¬°¥Á°ê¦~
      If CheckIsDate(textTM20, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¦è¤¸¦~" '"½Ð¿é¤J¥¿½Tªºµù¥U¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM20_GotFocus
      End If
      ' µù¥U¤é¤£¥i¶W¹L¨t²Î¤é
      If Val(DBDATE(textTM20)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "µù¥U¤é¤£¥i¶W¹L¨t²Î¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM20_GotFocus
      End If
   End If
   '910919 nick ÀË¬d©w¸q­Y¬OÀ³¸Ó»Pµù¥U¤é°µÀË¬d¡A«hµù¥U¤é¤£¯àªÅ¥Õ
   'If NickTmNa12 = 6 Then
        If Trim(textTM20) = "" Then
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "µù¥U¤é¤£¯àªÅ¥Õ¡I"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM20_GotFocus
        End If
   'End If
End Sub

' ±M¥Î´Á¶¡(°_)
Private Sub textTM21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM21) = False Then
      ' ¥Ó½Ð°ê®a¬°»Ý¿é¤J¥Á°ê¦~, §_«h¿é¤J¦è¤¸¦~
       'edit by nick 2004/10/06
'      If m_TM10 < "010" Then
'         If CheckIsTaiwanDate(textTM21, False) = False Then
'            Cancel = True
'            strTit = "¸ê®ÆÀË®Ö"
'            strMsg = "¥Ó½Ð°ê®a¬°¥xÆW, ½Ð¿é¤J¥Á°ê¦~"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            GoTo A0
'         End If
'      Else
         If CheckIsDate(textTM21, False) = False Then
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            '2011/12/15 MODIFY BY SONIA
            'strMsg = "¥Ó½Ð°ê®a«D¥xÆW, ½Ð¿é¤J¦è¤¸¦~"
            strMsg = "±M¥Î´Á¶¡½Ð¿é¤J¦è¤¸¦~¤ë¤é !"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo A0
         End If
'      End If
      
      Dim strTmp As String
      '2008/11/25 cancel by sonia TF-000570®Û­^»¡TF®×¥ó¤£·|ª¾¹D¥Ó½Ð¤é,¬G¤£ÀË¬d¦ý¤î¤é¬°°_¤é10¦~
      'If m_TM01 = "T" Or m_TM01 = "TF" Then
      If m_TM01 = "T" Then
          'edit by nick 2004/10/06
'         If m_TM10 < "010" Then
'            strTmp = m_TM21
'         Else
            strTmp = TransDate(m_TM21, 2)
'         End If
         If textTM21 <> strTmp Then
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            'Modified by Lydia 2019/12/09 +³Æµù
            'strMsg = "±M¥Î´Á­­°_¤éÀ³¬°<" & strTmp & ">"
            strMsg = "±M¥Î´Á­­°_¤éÀ³¬°<" & strTmp & ">¡A¬O§_Ä~Äò§@·~¡H"
            'Modify By Cheng 2002/11/08
            '­Y«ö½T©w, ¤´¥i§@·~
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            nResponse = MsgBox(strMsg, vbOKCancel, strTit)
            If nResponse = vbOK Then Cancel = False: Exit Sub
            
         End If
      '2008/10/25 ADD BY SONIA
      ElseIf m_TM01 = "TF" Then
         Dim strKey(0 To 4) As String
         strKey(0) = m_CP09
         strKey(1) = m_TM01
         strKey(2) = m_TM02
         strKey(3) = m_TM03
         strKey(4) = m_TM04
         If TFGetMoneyDate(m_TM10, strKey, m_TM21, strTmp, m_TM22) Then
            m_TM22 = CompDate(0, NickTmNa13, textTM21)
         End If
      '2008/10/25 END
      End If
   End If
A0:
   If Cancel Then TextInverse textTM21
End Sub

' ±M¥Î´Á¶¡(¨´)
Private Sub textTM22_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM22) = False Then
      ' ¥Ó½Ð°ê®a¬°»Ý¿é¤J¥Á°ê¦~, §_«h¿é¤J¦è¤¸¦~
      'edit by nick 2004/10/06
'      If m_TM10 < "010" Then
'         If CheckIsTaiwanDate(textTM22, False) = False Then
'            Cancel = True
'            strTit = "¸ê®ÆÀË®Ö"
'            strMsg = "¥Ó½Ð°ê®a¬°¥xÆW, ½Ð¿é¤J¥Á°ê¦~"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            GoTo A0
'         End If
'      Else
         If CheckIsDate(textTM22, False) = False Then
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            '2011/12/15 MODIFY BY SONIA
            'strMsg = "¥Ó½Ð°ê®a«D¥xÆW, ½Ð¿é¤J¦è¤¸¦~"
            strMsg = "±M¥Î´Á¶¡½Ð¿é¤J¦è¤¸¦~¤ë¤é !"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo A0
         End If
'      End If
   End If
A0:
   If Cancel Then TextInverse textTM22
End Sub

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/29ÀË¬dµe­±ªº TextBox ¬O§_§t¦³Unicode¤å¦r
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If

   
    'Modify By Cheng 2003/05/26
    '­Y¨t²ÎÃþ§O«DµÛ§@Åv, «h±M¥Î´Á¶¡¤@©w­n¿é¤J
    If m_TM01 <> "TC" Then
        'Add By Cheng 2003/05/23
        'ÀË¬d±M¥Î´Á¶¡
        If Me.textTM21.Text = "" Then
            MsgBox "½Ð¿é¤J±M¥Î´Á°_¤é!!!", vbExclamation + vbOKOnly
            Me.textTM21.SetFocus
            textTM21_GotFocus
            GoTo EXITSUB
        End If
        If Me.textTM22.Text = "" Then
            MsgBox "½Ð¿é¤J±M¥Î´Á¤î¤é!!!", vbExclamation + vbOKOnly
            Me.textTM22.SetFocus
            textTM22_GotFocus
            GoTo EXITSUB
        End If
    End If
   If m_TM01 <> "TC" And m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then   '91.11.3 ADD BY SONIA
      If IsEmptyText(textTM21) = True Or IsEmptyText(textTM22) = True Then
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J±M¥Î´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21.SetFocus
         GoTo EXITSUB
      Else
         If Not ChkRange(textTM21, textTM22, "±M¥Î´Á­­") Then
            textTM21.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If  '91.11.3 ADD BY SONIA
   ' ¨t²ÎÃþ§O¬°TC®É
   If m_TM01 = "TC" And m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
      ' TCµn°O¸¹¤£¥i¬°ªÅ¥Õ
      If IsEmptyText(textTC1) = True Then
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤JTCµn°O¸¹"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTC1.SetFocus
         GoTo EXITSUB
      End If
   End If
   'Add By Cheng 2002/06/12
   If m_TM01 = "TC" And m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
      'Modified by Lydia 2025/01/15 §ï¥ÎCaseFee
      'If m_SP51 = "¥xÆW¸gÀÙµo®i¬ã¨s°|" Then
      strExc(1) = ""
      strExc(0) = "SELECT Distinct(CF10) FROM CaseFee WHERE CF01='" & m_TM01 & "' AND CF02='" & ¥xÆW°ê®a¥N¸¹ & "' AND length(CF03)=3 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = "" & RsTemp.Fields(0)
      End If
      If m_SP51 = strExc(1) Then
      'end 2025/01/15
         ' TCµù¥U¸¹¼Æ¤£¥i¬°ªÅ¥Õ
         If IsEmptyText(textTC2) = True Then
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "½Ð¿é¤JTCµù¥U¸¹¼Æ"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTC2.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
'    'Add By Cheng 2003/11/19
'    '­YT¥Ó½Ð°ê®a¬°¥xÆW,¥Ó½Ð¤é¬°921128(§t)¥H«á, «hµù¥U¸¹¤Îµù¥U¤½§i¤é¤£¥iªÅ¥Õ
'    If m_TM01 = "T" And m_TM10 = "000" And DBDATE(Val(m_TM11)) >= 20031128 Then
    If m_TM01 = "T" And m_TM10 = "000" Then
        If Me.textTM15.Text = "" Then
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "½Ð¿é¤Jµù¥U¸¹"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15.SetFocus
            GoTo EXITSUB
        End If
        If Me.textTM14.Text = "" Then
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "½Ð¿é¤Jµù¥U¤½§i¤é"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM14.SetFocus
            GoTo EXITSUB
         'Modify By Sindy 2017/1/3 Mark:®Û­^­n¿é¤J20170101·|³Q¾×¦í,¦]¨Ó¨ç¤é´Á¬O20170103
'        'add by sonia 2016/11/17 ®Û­^­n¨D¤½§i¤é¥²¶·»P¨Ó¨ç¦¬¤å¤é(µoÃÒ¤é)¬Û¦P
'        ElseIf Me.textTM14.Text <> DBNullDate(m_CP05) Then
'            strTit = "¸ê®ÆÀË®Ö"
'            strMsg = "µù¥U¤½§i¤é¥²¶·»P¨Ó¨ç¦¬¤å¤é(µoÃÒ¤é)¬Û¦P !"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM14.SetFocus
'            GoTo EXITSUB
'        'end 2016/11/17
         '2017/1/3 END
        End If
'         '2005/11/11 ADD BY SONIA
'         If m_strLanguage = "2" Then
'            If Text1 <> "" Then
'               If CheckIsTaiwanDate(Text1) = False Then
'                  Text1_GotFocus
'                  GoTo EXITSUB
'               End If
'            Else
'               strTit = "¸ê®ÆÀË®Ö"
'               strMsg = "­^¤å©w½Z½Ð¿é¤JÃÒ®Ñ¤é´Á"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               Text1.SetFocus
'               GoTo EXITSUB
'            End If
'         End If
'         '2005/11/11 END
    End If
    '2005/4/14 ADD BY SONIA
    If m_TM01 = "TF" And Mid(m_TM02, 6, 1) <> "0" Then
      If Me.textCP47.Text = "" Then
          strTit = "¸ê®ÆÀË®Ö"
          strMsg = "½Ð¿é¤J»â¤g©µ¦ù´£¥Ó¤é"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textCP47.SetFocus
          GoTo EXITSUB
      End If
    End If
    '2005/4/14 END
    
   'Add By Sindy 2012/5/18
   If LabNP07.Caption <> "" Then
      'ÀË¬d¨Ó¨ç´Á­­--¤é´Á
      If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
         If Me.Option4(2).Value = True Then
            If Me.Text12.Text = "" Then
               MsgBox "½Ð¿é¤J¨Ó¨ç´Á­­!!!", vbExclamation + vbOKOnly
               Me.Text12.SetFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textDate_GotFocus()
   InverseTextBox textDate
End Sub

Private Sub textMoney_GotFocus()
   InverseTextBox textMoney
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPS_GotFocus()
   InverseTextBox textPS
End Sub

Private Sub textTC1_GotFocus()
   InverseTextBox textTC1
End Sub

Private Sub textTC2_GotFocus()
   InverseTextBox textTC2
End Sub
Private Sub textTM12_GotFocus()
   InverseTextBox textTM12
End Sub
Private Sub textTM21_GotFocus()
   InverseTextBox textTM21
End Sub

Private Sub textTM22_GotFocus()
   InverseTextBox textTM22
End Sub

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
Dim strNA1 As String
Dim strNA2 As String
Dim strTmp As String
Dim rsTmp As New ADODB.Recordset
'add by nickc 2006/06/14 ¥[¤J¤í´Ú¸ê®Æ
Dim A1kData As String
Dim arrTM09 As Variant, strGoodsKind As String 'Add By Sindy 2010/11/12
Dim str012 As String       '2013/10/9 add by sonia ¬O§_¦³«ü©wÁú°ê
Dim strET03 As String 'Add By Sindy 2014/11/28
   
   ' ¨ú¥Ó½Ð¤H°êÄy
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   strNA1 = Empty
   strNA2 = Empty
   str012 = Empty '2013/10/9 add by sonia
   
   ' ¨ú±o»â¤g©µ¥Ó«ü©w°ê®a¤Î°¨¼w¨½«ü©w°ê®a
   If m_TM01 = "TF" Then
      ' ¨ú»â¤g©µ¦ù«ü©w°ê®a
      '2006/5/3 MODIFY BY SONIA ¥u§ì¥¼®Ö»éªº¸ê®Æ
      'strSQL = "SELECT DISTINCT(TM10) FROM TradeMark " & _
      '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
      '               "TM02 = '" & m_TM02 & "' AND " & _
      '               "TM04 <> '00' "
      strSql = "SELECT DISTINCT(TM10) FROM TradeMark " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM04 <> '00' AND (TM16 IS NULL OR TM16<>'2') "
      '2006/5/3 END
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If IsNull(rsTmp.Fields("TM10")) = False Then
               strTmp = GetNationName(rsTmp.Fields("TM10"), 0)
               If IsEmptyText(strTmp) = False Then
                  If strNA1 <> Empty Then: strNA1 = strNA1 & ","
                  strNA1 = strNA1 & strTmp
               End If
               If rsTmp.Fields("TM10") = "012" Then str012 = "Y"  '2013/10/9 add by sonia
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      ' ¨ú°¨¼w¨½«ü©w°ê®a
      '2006/5/3 MODIFY BY SONIA ¥u§ì¥¼®Ö»éªº¸ê®Æ
      'strSQL = "SELECT DISTINCT(TM10) FROM TradeMark " & _
      '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
      '               "SUBSTR(TM02,1,5) = '" & Mid(m_TM02, 1, 5) & "' AND " & _
      '               "TM04 <> '00' "
      strSql = "SELECT DISTINCT(TM10) FROM TradeMark " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "SUBSTR(TM02,1,5) = '" & Mid(m_TM02, 1, 5) & "' AND " & _
                     "TM04 <> '00' AND (TM16 IS NULL OR TM16<>'2') "
      '2006/5/3 END
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If IsNull(rsTmp.Fields("TM10")) = False Then
               strTmp = GetNationName(rsTmp.Fields("TM10"), 0)
               If IsEmptyText(strTmp) = False Then
                  If strNA2 <> Empty Then: strNA2 = strNA2 & ","
                  strNA2 = strNA2 & strTmp
               End If
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   
   Select Case m_TM01
      Case "T":
         ' ¥Ó½Ð°ê®a¬°¥xÆW
         If m_TM10 < "010" Then
            ' ¥Ó½Ð¤H°êÄy¬°¥xÆW
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            'Add By Sindy 2013/5/3
            If m_strLanguage = "3" Then '¤é¤å
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "05", strCP09, "22", strUserNum
               ' Áp¦X°Ó¼Ð
               If IsEmptyText(m_TM27) = False Then
                  ' Áp¦X°Ó¼Ð
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & strCP09 & "','" & "22" & "','" & strUserNum & _
                           "','Áp¦X°Ó¼Ð','" & "¨Ì¦s ¥¿°Ó¼Ð µn¿ýµf¸¹ : (" & m_TM27 & ")" & "')"
                  cnnConnection.Execute strSql
               End If
               
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "05", strCP09, "23", strUserNum
               '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
               If m_TM67 <> "" Then
                   'Modify By Sindy 2022/10/12 ˆü¥e“¸Çy¦³ §ï¬° °Ó¼Ð“¸Çy¥D±i
                   'Modified by Morgan 2023/3/15
                   'strTmp = "°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & ChgSQL(m_TM67) & "¡vÇU°Ó¼Ð“¸Çy¥D±iþêÇQÆê¡C"
                   strTmp = PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv1") & "¡u" & ChgSQL(m_TM67) & "¡v" & PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv2")
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "05" & "','" & strCP09 & "','" & "23" & "','" & strUserNum & _
                            "','©ñ±ó±M¥ÎÅv','" & strTmp & "')"
                   cnnConnection.Execute strSql
               End If
               If m_TM118 <> "" Then
                  'Modified by Morgan 2023/3/15
                  'strTmp ="°Ó¼Ðªk²Ä30’f²Ä1¶µ²Ä10†AÇU³W©wÇR°òþøþà¡Bµn“÷°Ó¼Ð²Ä" & ChgSQL(m_TM118) & "†AÇU°Ó¼Ð“¸ªÌÇU¦P·NÇRÇoÇqµn“÷Çy³\¥iþìÇr¡C"
                  strTmp = PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ1") & ChgSQL(m_TM118) & PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ2")
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & strCP09 & "','" & "23" & "','" & strUserNum & _
                           "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & strTmp & "')"
                  cnnConnection.Execute strSql
               End If
                'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "05" & "','" & strCP09 & "','" & "22" & "','" & strUserNum & "'," & _
                         "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                cnnConnection.Execute strSql
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "05" & "','" & strCP09 & "','" & "23" & "','" & strUserNum & "'," & _
                         "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                cnnConnection.Execute strSql
                'end 2017/04/21
            Else
            '2013/5/3 End
               If textPrint = "1" Then
'                   '­Y¥Ó½Ð¤é¤p©ó921128
'                   If DBDATE(Val(m_TM11)) < 20031128 Then
'                       '­Y±M¥Î°_¤é¤p©ó921128
'                       If DBDATE(Val(Me.textTM21.Text)) < 20031128 Then
'                           'Modify By Cheng 2003/01/02
'                           '­Y°Ó¼ÐºØÃþ«D¹ÎÅé¼Ð³¹, ÃÒ©ú¼Ð³¹
'                           If m_TM08 <> "7" And m_TM08 <> "8" Then
'                               EndLetter "05", strCP09, "11", strUserNum
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "11" & "','" & strUserNum & "'," & _
'                                        "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                               cnnConnection.Execute strSql
'                               'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "11" & "','" & strUserNum & "'," & _
'                                        "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                               cnnConnection.Execute strSql
'                               'end 2017/04/21
'                           '­Y°Ó¼ÐºØÃþ¬°¹ÎÅé¼Ð³¹, ÃÒ©ú¼Ð³¹
'                           Else
'                               EndLetter "05", strCP09, "12", strUserNum
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "12" & "','" & strUserNum & "'," & _
'                                        "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                               cnnConnection.Execute strSql
'                               'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "12" & "','" & strUserNum & "'," & _
'                                        "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                               cnnConnection.Execute strSql
'                               'end 2017/04/21
'                           End If
'                       '­Y±M¥Î°_¤é¤j©óµ¥©ó921128
'                       Else
'                           'Modify By Cheng 2003/01/02
'                           '­Y°Ó¼ÐºØÃþ«D¹ÎÅé¼Ð³¹, ÃÒ©ú¼Ð³¹
'                           If m_TM08 <> "7" And m_TM08 <> "8" Then
'                               'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
'                               ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
'                               'If m_blnReceiveSecond = False Then
'                               '    EndLetter "05", strCP09, "01", strUserNum
'                               '    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               '             "VALUES ('" & "05" & "','" & strCP09 & "','" & "01" & "','" & strUserNum & "'," & _
'                               '             "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                               '    cnnConnection.Execute strSql
'                               '     'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                               '     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               '              "VALUES ('" & "05" & "','" & strCP09 & "','" & "01" & "','" & strUserNum & "'," & _
'                               '              "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                               '     cnnConnection.Execute strSql
'                               '     'end 2017/04/21
'                               ''­Y¤w¦¬²Ä¤G´Áµù¥U¶O
'                               'Else
'                               'end 2019/4/24
'                                   EndLetter "05", strCP09, "15", strUserNum
'                                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                            "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & "'," & _
'                                            "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                                   cnnConnection.Execute strSql
'                                    'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & "'," & _
'                                             "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                                    cnnConnection.Execute strSql
'                                    'end 2017/04/21
'                               'End If  'cancel by sonia 2019/4/24
'                           '­Y°Ó¼ÐºØÃþ¬°¹ÎÅé¼Ð³¹, ÃÒ©ú¼Ð³¹
'                           Else
'                               EndLetter "05", strCP09, "04", strUserNum
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "'," & _
'                                        "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                               cnnConnection.Execute strSql
'                               'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "'," & _
'                                        "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                               cnnConnection.Execute strSql
'                               'end 2017/04/21
'                           End If
'                       End If
'                   '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128
'                   Else
                       '­Y°Ó¼ÐºØÃþ¬°°Ó¼Ð
                       If m_TM08 <> "6" And m_TM08 <> "7" And m_TM08 <> "8" Then
                           'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
                           ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
                           'If m_blnReceiveSecond = False Then
                           '    EndLetter "05", strCP09, "05", strUserNum
                           '    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           '             "VALUES ('" & "05" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & "'," & _
                           '             "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                           '    cnnConnection.Execute strSql
                           '    'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                           '    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           '             "VALUES ('" & "05" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & "'," & _
                           '             "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                           '    cnnConnection.Execute strSql
                           '    'end 2017/04/21
                           ''­Y¤w¦¬²Ä¤Gµù¥U¶O
                           'Else
                           'end 2019/4/24
                              '2005/11/11 MODIFY BY SONIA ¥[¤J©w½Z»y¤å§PÂ_
                              Select Case m_strLanguage
                              Case "1"  '¤¤¤å
                                 EndLetter "05", strCP09, "08", strUserNum
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "05" & "','" & strCP09 & "','" & "08" & "','" & strUserNum & "'," & _
                                          "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                                 cnnConnection.Execute strSql
                               'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "08" & "','" & strUserNum & "'," & _
                                        "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                               cnnConnection.Execute strSql
                               'end 2017/04/21
                              Case "2"  '­^¤å
                                 EndLetter "05", strCP09, "18", strUserNum
   '                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "18" & "','" & strUserNum & "'," & _
   '                                       "'" & "ÃÒ®Ñ¤é´Á" & "','" & DBDATE(Text1) & "')"
   '                              cnnConnection.Execute strSql
                              End Select
                              '2005/11/11 END
                           'End If  'cancel by sonia 2019/4/24
                       '­Y°Ó¼ÐºØÃþ¬°¼Ð³¹
                       Else
                           'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
                           ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
                           'If m_blnReceiveSecond = False Then
                           '    EndLetter "05", strCP09, "06", strUserNum
                           '    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           '             "VALUES ('" & "05" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & "'," & _
                           '             "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                           '    cnnConnection.Execute strSql
                           '    'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                           '    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           '             "VALUES ('" & "05" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & "'," & _
                           '             "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                           '    cnnConnection.Execute strSql
                           '    'end 2017/04/21
                           ''­Y¤w¦¬²Ä¤Gµù¥U¶O
                           'Else
                           'end 2019/4/24
                               EndLetter "05", strCP09, "09", strUserNum
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "09" & "','" & strUserNum & "'," & _
                                        "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                               cnnConnection.Execute strSql
                               'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "09" & "','" & strUserNum & "'," & _
                                        "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                               cnnConnection.Execute strSql
                               'end 2017/04/21
                           'End If  'cancel by sonia 2019/4/24
                       End If
'                   End If
               
               ' ¥Ó½Ð¤H°êÄy«D¥xÆW
               'edit by nickc 2006/06/30
               'Else
               ElseIf textPrint = "2" Then
'                   '­Y¥Ó½Ð¤é¤p©ó20031128
'                   If DBDATE(Val(m_TM11)) < 20031128 Then
'                       '­Y±M¥Î°_¤é¤p©ó20031128
'                       If DBDATE(Val(Me.textTM21.Text)) < 20031128 Then
'                           m_blnNoResult = GetNoResult(m_TM01, m_TM02, m_TM03, m_TM04)
'                           If m_blnNoResult = False Then
'                               EndLetter "05", strCP09, "13", strUserNum
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & "'," & _
'                                        "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                               cnnConnection.Execute strSql
'
'                               'add by nickc 2006/06/14 ¥[¤J¤í´Ú¸ê®Æ
'                               'Remove by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
'                               'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
'                               'If A1kData <> "" Then
'                               '     A1kData = "¡@¥»©Ò¨´¤µ©|¥¼¦¬¨ì¥»¥ó°Ó¼Ð" & A1kData & "¡A·Ð½Ð¾¨³t±N¤W­z´Ú¶µÂY±H¥»©Ò¡A¬O¬è¡I" & vbCrLf '& vbCrLf
'                               '     'Modify By Sindy 2009/10/21 ¥¨¨Ê°Ó¼Ð(96030)ªº«È¤á¤£¥X´Ú¶µ
'                               '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
'                               'End If
'                               'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               '         "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & "'," & _
'                               '         "'" & "¤í´Ú¸ê®Æ" & "','" & A1kData & "')"
'                               'cnnConnection.Execute strSql
'                               'end 2016/12/22
'
'                                'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & "'," & _
'                                         "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                                cnnConnection.Execute strSql
'                                'end 2017/04/21
'                           Else
'                               EndLetter "05", strCP09, "14", strUserNum
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                        "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                               cnnConnection.Execute strSql
'
'                               'add by nickc 2006/06/14 ¥[¤J¤í´Ú¸ê®Æ
'                               'Remove by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
'                               'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
'                               'If A1kData <> "" Then
'                               '     A1kData = "¡@¥»©Ò¨´¤µ©|¥¼¦¬¨ì¥»¥ó°Ó¼Ð" & A1kData & "¡A·Ð½Ð¾¨³t±N¤W­z´Ú¶µÂY±H¥»©Ò¡A¬O¬è¡I" & vbCrLf '& vbCrLf
'                               '     'Modify By Sindy 2009/10/21 ¥¨¨Ê°Ó¼Ð(96030)ªº«È¤á¤£¥X´Ú¶µ
'                               '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
'                               'End If
'                               'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               '         "VALUES ('" & "05" & "','" & strCP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                               '         "'" & "¤í´Ú¸ê®Æ" & "','" & A1kData & "')"
'                               'cnnConnection.Execute strSql
'                               'end 2016/12/22
'
'                                'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "05" & "','" & strCP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                         "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                                cnnConnection.Execute strSql
'                                'end 2017/04/21
'                           End If
'                       '­Y±M¥Î°_¤é¤j©óµ¥©ó20031128
'                       Else
'                           'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
'                           ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
'                           'If m_blnReceiveSecond = False Then
'                           '    EndLetter "05", strCP09, "02", strUserNum
'                           '    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                           '             "VALUES ('" & "05" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
'                           '             "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                           '    cnnConnection.Execute strSql
'                           '    'add by nickc 2006/06/14 ¥[¤J¤í´Ú¸ê®Æ
'                           '    'Remove by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
'                           '    'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
'                           '    'If A1kData <> "" Then
'                           '    '     A1kData = "¡@¥»©Ò¨´¤µ©|¥¼¦¬¨ì¥»¥ó°Ó¼Ð" & A1kData & "¡A·Ð½Ð¾¨³t±N¤W­z´Ú¶µÂY±H¥»©Ò¡A¬O¬è¡I" & vbCrLf '& vbCrLf
'                           '    '     'Modify By Sindy 2009/10/21 ¥¨¨Ê°Ó¼Ð(96030)ªº«È¤á¤£¥X´Ú¶µ
'                           '    '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
'                           '    'End If
'                           '    'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                           '    '         "VALUES ('" & "05" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
'                           '    '         "'" & "¤í´Ú¸ê®Æ" & "','" & A1kData & "')"
'                           '    'cnnConnection.Execute strSql
'                           '    'end 2016/12/22
'                           '
'                           '     'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                           '     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                           '              "VALUES ('" & "05" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
'                           '              "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                           '     cnnConnection.Execute strSql
'                           '     'end 2017/04/21
'                           ''­Y¤w¦¬²Ä¤G´Áµù¥U¶O
'                           'Else
'                           'end 2019/4/24
'                               EndLetter "05", strCP09, "16", strUserNum
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "16" & "','" & strUserNum & "'," & _
'                                        "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
'                               cnnConnection.Execute strSql
'
'                               'add by nickc 2006/06/14 ¥[¤J¤í´Ú¸ê®Æ
'                               'Remove by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
'                               'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
'                               'If A1kData <> "" Then
'                               '     A1kData = "¡@¥»©Ò¨´¤µ©|¥¼¦¬¨ì¥»¥ó°Ó¼Ð" & A1kData & "¡A·Ð½Ð¾¨³t±N¤W­z´Ú¶µÂY±H¥»©Ò¡A¬O¬è¡I" & vbCrLf '& vbCrLf
'                               '     'Modify By Sindy 2009/10/21 ¥¨¨Ê°Ó¼Ð(96030)ªº«È¤á¤£¥X´Ú¶µ
'                               '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
'                               'End If
'                               'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               '         "VALUES ('" & "05" & "','" & strCP09 & "','" & "16" & "','" & strUserNum & "'," & _
'                               '         "'" & "¤í´Ú¸ê®Æ" & "','" & A1kData & "')"
'                               'cnnConnection.Execute strSql
'                               'end 2016/12/22
'
'                                'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "05" & "','" & strCP09 & "','" & "16" & "','" & strUserNum & "'," & _
'                                         "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
'                                cnnConnection.Execute strSql
'                                'end 2017/04/21
'                           'End If  'cancel by sonia 2019/4/24
'                       End If
'                   '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
'                   Else
                       'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
                       ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
                       'If m_blnReceiveSecond = False Then
                       '    EndLetter "05", strCP09, "07", strUserNum
                       '    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       '             "VALUES ('" & "05" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & "'," & _
                       '             "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                       '    cnnConnection.Execute strSql
                       '    'add by nickc 2006/06/14 ¥[¤J¤í´Ú¸ê®Æ
                       '    'Remove by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
                       '    'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                       '    'If A1kData <> "" Then
                       '    '   A1kData = "¡@¥»©Ò¨´¤µ©|¥¼¦¬¨ì¥»¥ó°Ó¼Ð" & A1kData & "¡A·Ð½Ð¾¨³t±N¤W­z´Ú¶µÂY±H¥»©Ò¡A¬O¬è¡I" & vbCrLf '& vbCrLf
                       '    '   'Modify By Sindy 2009/10/21 ¥¨¨Ê°Ó¼Ð(96030)ªº«È¤á¤£¥X´Ú¶µ
                       '    '   If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                       '    'End If
                       '    'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       '    '         "VALUES ('" & "05" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & "'," & _
                       '    '         "'" & "¤í´Ú¸ê®Æ" & "','" & A1kData & "')"
                       '    'cnnConnection.Execute strSql
                       '    'end 2016/12/22
                       '
                       '     'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                       '     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       '              "VALUES ('" & "05" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & "'," & _
                       '              "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                       '     cnnConnection.Execute strSql
                       '     'end 2017/04/21
                       ''­Y¤w¦¬²Ä¤Gµù¥U¶O
                       'Else
                       'end 2019/4/24
                           '2005/11/11 MODIFY BY SONIA ¥[¤J©w½Z»y¤å§PÂ_
                           Select Case m_strLanguage
                           Case "1"  '¤¤¤å
                              EndLetter "05", strCP09, "10", strUserNum
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "10" & "','" & strUserNum & "'," & _
                                       "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                              cnnConnection.Execute strSql
                              
                               'add by nickc 2006/06/14 ¥[¤J¤í´Ú¸ê®Æ
                               'Remove by Lydia 2016/12/22 ¤º°Ó¤§«Dª§Ä³®×®Ö­ã©Îµù¥UÃÒ¿é¤J,­ì¥»©w½Zªº¤í´Ú¸ê®Æ§ï¦¨DÃþ¦¬¤å±±¨î
                               'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                               'If A1kData <> "" Then
                               '     A1kData = "¡@¥»©Ò¨´¤µ©|¥¼¦¬¨ì¥»¥ó°Ó¼Ð" & A1kData & "¡A·Ð½Ð¾¨³t±N¤W­z´Ú¶µÂY±H¥»©Ò¡A¬O¬è¡I" & vbCrLf '& vbCrLf
                               '     'Modify By Sindy 2009/10/21 ¥¨¨Ê°Ó¼Ð(96030)ªº«È¤á¤£¥X´Ú¶µ
                               '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                               'End If
                               'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               '         "VALUES ('" & "05" & "','" & strCP09 & "','" & "10" & "','" & strUserNum & "'," & _
                               '         "'" & "¤í´Ú¸ê®Æ" & "','" & A1kData & "')"
                               'cnnConnection.Execute strSql
                               'end 2016/12/22
                               
                                'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "05" & "','" & strCP09 & "','" & "10" & "','" & strUserNum & "'," & _
                                         "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                                cnnConnection.Execute strSql
                                'end 2017/04/21
                           Case "2"  '­^¤å
                              EndLetter "05", strCP09, "18", strUserNum
   '                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "18" & "','" & strUserNum & "'," & _
   '                                    "'" & "ÃÒ®Ñ¤é´Á" & "','" & DBDATE(Text1) & "')"
   '                           cnnConnection.Execute strSql
                           End Select
                           '2005/11/11 END
                       'End If  'cancel by sonia 2019/4/24
'                   End If

               'Added by Lydia 2017/04/21 ¼W¥[­^¤å©w½ZªºÄæ¦ì
               ElseIf textPrint = "3" Then
                    EndLetter "05", strCP09, "17", strUserNum
                    '¼W¥[©w½Zµo¨ç¤é´Á
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "17" & "','" & strUserNum & "'," & _
                             "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                    cnnConnection.Execute strSql
                    
                    EndLetter "05", strCP09, "18", strUserNum
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "18" & "','" & strUserNum & "'," & _
                             "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                    cnnConnection.Execute strSql
               'end 2017/04/21
               End If
            End If
            
         ' ¥Ó½Ð°ê®a¬°¤j³°
         ElseIf m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                If Val(Trim(Me.textMoney.Text)) <> 0 Then
                  EndLetter "05", strCP09, "20", strUserNum
                  'Modify By Sindy 2009/10/23 §ï¬°³ø»ù³qª¾
                  strNP22 = "0" 'Added by Morgan 2015/6/16 ¦Û°ÊµoÃÒªº»âÃÒ³ø»ù²Î¤@³] 0
                  'modify by sonia 2019/2/1
                  'PUB_AddLetterCache strCP09, strNP22, strCP09, "05", "20"
                  If Option5(0).Value = True Then '¹q¤lÃÒ®Ñ
                     'Modify By Sindy 2020/2/19 + «H¨ç¦¬¤å¸¹
                     PUB_AddLetterCache strCP09, strNP22, strCP09, "05", "27", , IIf(strSrvDate(1) >= T°Ó¼Ð¹q¤l¤Æ²Ä2¶¥¬q±Ò¥Î¤é, strLD18, "")
                  Else '¯È¥»ÃÒ®Ñ
                     'Modify By Sindy 2020/2/19 + «H¨ç¦¬¤å¸¹
                     PUB_AddLetterCache strCP09, strNP22, strCP09, "05", "20", , IIf(strSrvDate(1) >= T°Ó¼Ð¹q¤l¤Æ²Ä2¶¥¬q±Ò¥Î¤é, strLD18, "")
                  End If
                  'end 2019/2/1
                  '********************************
                  InsExpField1 strCP09, strNP22, "20"
                  strExc(0) = CompWorkDay(5, strSrvDate(1))
                  strExc(1) = DBDATE(strNP08)
                  '********************************
                  '­Y[¨t²Î¤é+5­Ó¤u§@¤Ñ>=©Ò­­]®É¡A¤£¥²Åý´¼Åv¤H­û½T»{¡Aª½±µ¦C¦L
                  If Val(strExc(1)) <= Val(strExc(0)) Then
                     PUB_Cache2Letter strCP09, strNP22, False, False
                  End If
                  '2009/10/23 End
                Else
                  'Added by Lydia 2017/02/02 ¥x-¤j°Ï¤À¤@¯ëµù¥UÃÒ©M³¡¥÷ºM¾P­«µoµù¥UÃÒ;¦]¬°²Ä¤@¦¸»âÃÒ¥UÃÒ¤w¥I¶O,©Ò¥H³¡¥÷ºM¾P¤£»Ý»âÃÒ¶O
                  If str1006CP64 = "" Then
                     'add by sonia 2019/2/1 ¦A¤À¹q¤lÃÒ®Ñ,¯È¥»ÃÒ®Ñ,¥H¤UET03§ï¶ÇÅÜ¼Æ
                     If Option5(0).Value = True Then
                        strET03 = "25"
                     Else
                        strET03 = "03"
                     End If
                     'end 2019/2/1 ¥H¤UET03§ï¶ÇÅÜ¼Æ
                     '­ìµ{¦¡->¤@¯ëµù¥UÃÒ
                     EndLetter "05", strCP09, strET03, strUserNum
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                             "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                     cnnConnection.Execute strSql
                    'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                             "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                    cnnConnection.Execute strSql
                    'end 2017/04/21
                  Else
                     'add by sonia 2019/2/1 ¦A¤À¹q¤lÃÒ®Ñ,¯È¥»ÃÒ®Ñ,¥H¤UET03§ï¶ÇÅÜ¼Æ
                     If Option5(0).Value = True Then
                        strET03 = "26"
                     Else
                        strET03 = "24"
                     End If
                     'end 2019/2/1 ¥H¤UET03§ï¶ÇÅÜ¼Æ
                     EndLetter "05", strCP09, strET03, strUserNum
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                             "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                     cnnConnection.Execute strSql
                     strExc(1) = IIf(str1006CP64 = "TRUE", " ", str1006CP64)
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                              "'" & "³¡¥÷ºM¾Pµù¥UÃÒ" & "','" & strExc(1) & "')"
                     cnnConnection.Execute strSql
                    'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                             "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
                    cnnConnection.Execute strSql
                    'end 2017/04/21
                  End If
                  'end 2017/02/02
                End If
            'add by nickc 2007/07/24 ¥[¤J­^¤å
            ElseIf textPrint = "3" Then
               EndLetter "05", strCP09, "19", strUserNum
               'Add By Sindy 2013/6/5 ÀË¬d¤§«e¬O§_¤w¦³¦¬¹L»âÃÒ
               strSql = "SELECT cp10 FROM caseprogress " & _
                        "WHERE cp01='" & m_TM01 & "' AND cp02='" & m_TM02 & "' AND cp03='" & m_TM03 & "' " & _
                        "AND cp04 = '" & m_TM04 & "' AND cp10 = '701' " & _
                        "AND cp27<=" & strSrvDate(1)
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount <= 0 Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & strCP09 & "','" & "19" & "','" & strUserNum & "'," & _
                           "'" & "©|¥¼¦¬¤å»âÃÒ¤º¤å" & "','Our relevant debit note is also enclosed for your kind settlement.')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & strCP09 & "','" & "19" & "','" & strUserNum & "'," & _
                           "'" & "©|¥¼¦¬¤å»âÃÒªþ¥ó" & "','2.Debit Note.')"
                  cnnConnection.Execute strSql
               End If
               rsTmp.Close
               '2013/6/5 End
               
               'Added by Lydia 2017/04/21 ¼W¥[©w½Zµo¨ç¤é´Á
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "19" & "','" & strUserNum & "'," & _
                        "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
               cnnConnection.Execute strSql
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "21" & "','" & strUserNum & "'," & _
                        "'" & "©w½Zµo¨ç¤é´Á" & "','" & strSrvDate(1) & "')"
               cnnConnection.Execute strSql
               'end 2017/04/21
            End If
         End If
      Case "TF":
            'add by nickc 2006/06/30
            If textPrint = "1" And m_TM04 = "00" And m_TM03 = "0" Then
               ' ¥»©Ò®×¸¹²Ä¤E½X
               If Mid(m_TM02, 6, 1) = "0" Then
                  EndLetter "05", strCP09, "04", strUserNum
                  '2013/10/9 ADD BY SONIA
                  If str012 = "Y" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & "05" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "','«ü©wÁú°ê','¡ð')"
                     cnnConnection.Execute strSql
                  End If
                  '2013/10/9 END
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "'," & _
                           "'" & "¦C¦L³Æµù" & "','" & textPS & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "'," & _
                           "'" & "°¨¼w¨½«ü©w°ê®a" & "','" & strNA2 & "')"
                  cnnConnection.Execute strSql
                  'Add By Sindy 2010/11/12
                  '1-34°Ó«~ 35-45ªA°È
                  strGoodsKind = "°Ó«~"
                  If Trim(textTM09.Text) > "" Then
                    arrTM09 = Split(textTM09.Text, ",")
                    If Val(arrTM09(0)) >= 35 And Val(arrTM09(0)) <= 45 Then
                       strGoodsKind = "ªA°È"
                    End If
                  End If
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "'," & _
                       "'°Ó«~©ÎªA°È','" & strGoodsKind & "')"
                  cnnConnection.Execute strSql
                  '2010/11/12 End
               Else
                  EndLetter "05", strCP09, "05", strUserNum
                  '2013/10/9 ADD BY SONIA
                  If str012 = "Y" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & "05" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & "','«ü©wÁú°ê','¡ð')"
                  End If
                  '2013/10/9 END
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & "'," & _
                           "'" & "»â¤g©µ¦ù«ü©w°ê®a" & "','" & strNA1 & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & "'," & _
                           "'" & "¨ä¥L¤é´Á" & "','" & DBDATE(textCP47) & "')"
                  cnnConnection.Execute strSql
               End If
            End If
      Case "TC":
         ' ¥Ó½Ð°ê®a¬°¤j³°
         If m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                '93.12.9 MODIFY BY SONIA ­ì¥¼°Ï¤À§@«~ºØÃþ, ¥[¤J¬ü³NµÛ§@(08)¤§©w½Z
                Select Case Trim(m_SP46)
                   Case "¬ü³NµÛ§@":
                        strET03 = "08" 'Add By Sindy 2014/11/28
                        EndLetter "05", strCP09, "08", strUserNum
                        'add by nickc 2007/04/27 °Ï¤À§@«~ºØÃþ®É¡A§Ñ°O¥[¡A¸É¤W
'                        If Me.textMoney.Text <> "" Then
'                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                     "VALUES ('" & "05" & "','" & strCP09 & "','" & "08" & "','" & strUserNum & "'," & _
'                                     "'" & "¤j³°»âÃÒ¶O" & "','" & textMoney & "')"
'                            cnnConnection.Execute strSql
'                        End If
                   Case "­pºâ¾÷³n¥ó":
                        strET03 = "06" 'Add By Sindy 2014/11/28
                        EndLetter "05", strCP09, "06", strUserNum
                        'add by nickc 2007/04/27 ¥Ñ¤U­±·h¤W¨Ó
'                        If Me.textMoney.Text <> "" Then
'                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                     "VALUES ('" & "05" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & "'," & _
'                                     "'" & "¤j³°»âÃÒ¶O" & "','" & textMoney & "')"
'                            cnnConnection.Execute strSql
'                        End If
                   '2010/9/8 ADD BY SONIA
                   Case Else
                        strET03 = "09" 'Add By Sindy 2014/11/28
                        EndLetter "05", strCP09, "09", strUserNum
'                        If Me.textMoney.Text <> "" Then
'                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                     "VALUES ('" & "05" & "','" & strCP09 & "','" & "09" & "','" & strUserNum & "'," & _
'                                     "'" & "¤j³°»âÃÒ¶O" & "','" & textMoney & "')"
'                            cnnConnection.Execute strSql
'                        End If
                End Select
                'Modify By Sindy 2014/11/28 §ï¬°³ø»ù³qª¾
                If Val(Trim(Me.textMoney.Text)) <> 0 Then
                   If Val(strNP22) = 0 Then strNP22 = 0 'Add By Sindy 2014/11/28
                   'Modify By Sindy 2020/2/19 + «H¨ç¦¬¤å¸¹
                   PUB_AddLetterCache strCP09, strNP22, strCP09, "05", strET03, , IIf(strSrvDate(1) >= T°Ó¼Ð¹q¤l¤Æ²Ä2¶¥¬q±Ò¥Î¤é, strLD18, "")
                   InsExpField1 strCP09, strNP22, strET03
'                   strExc(0) = CompWorkDay(5, strSrvDate(1))
'                   strExc(1) = DBDATE(strNP08)
'                   '­Y[¨t²Î¤é+5­Ó¤u§@¤Ñ>=©Ò­­]®É¡A¤£¥²Åý´¼Åv¤H­û½T»{¡Aª½±µ¦C¦L
'                   If Val(strExc(1)) <= Val(strExc(0)) Then
'                      PUB_Cache2Letter strCP09, strNP22, False, False
'                   End If
                End If
                '2014/11/28 END
                
                '93.12.9 end
'edit by nickc 2007/04/27 ©¹¤W·h
'                If Me.textMoney.Text <> "" Then
'                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & "'," & _
'                             "'" & "¤j³°»âÃÒ¶O" & "','" & textMoney & "')"
'                    cnnConnection.Execute strSQL
'                End If
            End If
         ' ¥Ó½Ð°ê®a¬°¥xÆW
         ElseIf m_TM10 < "010" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
               EndLetter "05", strCP09, "07", strUserNum
            'Add By Sindy 2010/01/20 ¤j->¥x
            ElseIf textPrint = "2" And m_CP10 = "806" Then
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "05", m_CP09, "01", strUserNum
            '2010/01/20 End
            End If
         End If
   End Select
End Sub

'Add By Sindy 2009/10/23
'¼g¨Ò¥~Äæ¦ì¨ì¼È¦sÀÉ
Private Sub InsExpField1(NP01 As String, NP22 As String, Optional ET03 As String)
   strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                   "VALUES ('" & NP01 & "'," & NP22 & ",'»âÃÒ¶O','" & Me.textMoney.Text & "','Y')"
   cnnConnection.Execute strSql
   strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                   "VALUES ('" & NP01 & "'," & NP22 & ",'»âÃÒ¶OÂI¼Æ','" & (Val(Me.textMoney.Text) / 1000) & "','')"
   cnnConnection.Execute strSql
   'modify by sonia 2019/2/1 +27¹q¤lÃÒ®Ñ
   If ET03 = "20" Or ET03 = "27" Then
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                      "VALUES ('" & NP01 & "'," & NP22 & ",'¦C¦L³Æµù','" & Me.textPS.Text & "','')"
      cnnConnection.Execute strSql
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ¦C¦L©w½Z
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean, ET03_1 As String
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   InsExpField
   
   'Add By Sindy 2012/1/13
   ET01 = "05"
   ET02 = strCP09
   bolEdit = IIf(Me.textEditPrint.Text = "Y", True, False)
   '2012/1/13 End
   
   Select Case m_TM01
      Case "T":
         ' ¥Ó½Ð°ê®a¬°¥xÆW
         If m_TM10 < "010" Then
            ' ¥Ó½Ð¤H°êÄy¬°¥xÆW
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            'Add By Sindy 2013/5/3
            If m_strLanguage = "3" Then '¤é¤å
               ET03 = "22"
               ET03_1 = "23" 'Ä¶¤å
            Else
            '2013/5/3 End
               If textPrint = "1" Then
'                   '­Y¥Ó½Ð¤é¤p©ó20031128
'                   If DBDATE(Val(m_TM11)) < 20031128 Then
'                       '­Y±M¥Î°_¤é¤p©ó920901
'                       If DBDATE(Val(Me.textTM21.Text)) < 20031128 Then
'                            'Modify By Cheng 2003/01/02
'                            '­Y°Ó¼ÐºØÃþ«D¹ÎÅé¼Ð³¹, ÃÒ©ú¼Ð³¹
'                            If m_TM08 <> "7" And m_TM08 <> "8" Then
'   '                             NowPrint strCP09, "05", "11", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                                 ET03 = "11" 'Modify By Sindy 2012/1/13
'                            '­Y°Ó¼ÐºØÃþ¬°¹ÎÅé¼Ð³¹, ÃÒ©ú¼Ð³¹
'                            Else
'   '                             NowPrint strCP09, "05", "12", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                                 ET03 = "12" 'Modify By Sindy 2012/1/13
'                            End If
'                       '­Y±M¥Î°_¤é¤j©óµ¥©ó921128
'                       Else
'                            'Modify By Cheng 2003/01/02
'                            '­Y°Ó¼ÐºØÃþ«D¹ÎÅé¼Ð³¹, ÃÒ©ú¼Ð³¹
'                            If m_TM08 <> "7" And m_TM08 <> "8" Then
'                               'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
'                               ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
'                               'If m_blnReceiveSecond = False Then
'   '                           '     NowPrint strCP09, "05", "01", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                               '     ET03 = "01" 'Modify By Sindy 2012/1/13
'                               ''­Y¤w¦¬²Ä¤G´Áµù¥U¶O
'                               'Else
'                               'end 2019/4/24
'   '                                NowPrint strCP09, "05", "15", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                                    ET03 = "15" 'Modify By Sindy 2012/1/13
'                               'End If  'cancel by sonia 2019/4/24
'                            '­Y°Ó¼ÐºØÃþ¬°¹ÎÅé¼Ð³¹, ÃÒ©ú¼Ð³¹
'                            Else
'   '                             NowPrint strCP09, "05", "04", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                                 ET03 = "04" 'Modify By Sindy 2012/1/13
'                            End If
'                       End If
'                   '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
'                   Else
                       '­Y°Ó¼ÐºØÃþ¬°°Ó¼Ð
                       If m_TM08 <> "6" And m_TM08 <> "7" And m_TM08 <> "8" Then
                           'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
                           ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
                           'If m_blnReceiveSecond = False Then
   '                       '    NowPrint strCP09, "05", "05", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                           '   ET03 = "05" 'Modify By Sindy 2012/1/13
                           ''­Y¤w¦¬²Ä¤G´Áµù¥U¶O
                           'Else
                           'end 2019/4/24
   '                           NowPrint strCP09, "05", "08", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                              ET03 = "08" 'Modify By Sindy 2012/1/13
                           'End If  'cancel by sonia 2019/4/24
                       '­Y°Ó¼ÐºØÃþ¬°¼Ð³¹
                       Else
                           'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
                           ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
                           'If m_blnReceiveSecond = False Then
   '                       '     NowPrint strCP09, "05", "06", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                           '   ET03 = "06" 'Modify By Sindy 2012/1/13
                           ''­Y¤w¦¬²Ä¤G´Áµù¥U¶O
                           'Else
                           'end 2019/4/24
   '                            NowPrint strCP09, "05", "09", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                              ET03 = "09" 'Modify By Sindy 2012/1/13
                           'End If 'cancel by sonia 2019/4/24
                       End If
'                   End If
               ' ¥Ó½Ð¤H°êÄy«D¥xÆW
               'edit by nickc 2006/06/30
               'Else
               ElseIf textPrint = "2" Then
'                   '¥Ó½Ð¤é¤p©ó20031128
'                   If DBDATE(Val(m_TM11)) < 20031128 Then
'                       '­Y±M¥Î°_¤é¤p©ó20031128
'                       If DBDATE(Val(Me.textTM21.Text)) < 20031128 Then
'                           If m_blnNoResult = False Then
'   '                            NowPrint strCP09, "05", "13", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                              ET03 = "13" 'Modify By Sindy 2012/1/13
'                           Else
'   '                            NowPrint strCP09, "05", "14", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                              ET03 = "14" 'Modify By Sindy 2012/1/13
'                           End If
'                       '­Y±M¥Î°_¤é¤j©óµ¥©ó20031128
'                       Else
'                           'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
'                           ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
'                           'If m_blnReceiveSecond = False Then
'   '                       '    NowPrint strCP09, "05", "02", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                           '   ET03 = "02" 'Modify By Sindy 2012/1/13
'                           ''­Y¤w¦¬²Ä¤G´Áµù¥U¶O
'                           'Else
'                           'end 2019/4/24
'   '                           NowPrint strCP09, "05", "16", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
'                              ET03 = "16" 'Modify By Sindy 2012/1/13
'                           'End If  'cancel by sonia 2019/4/24
'                       End If
'                   '¥Ó½Ð¤é¤j©óµ¥©ó20031128
'                   Else
                       'cancel by sonia 2019/4/24 T-217534«È¤á¦Û¦æÃºµù¥U¶O
                       ''­Y¥¼¦¬²Ä¤G´Áµù¥U¶O
                       'If m_blnReceiveSecond = False Then
   '                   '     NowPrint strCP09, "05", "07", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                       '    ET03 = "07" 'Modify By Sindy 2012/1/13
                       ''­Y¤w¦¬²Ä¤G´Áµù¥U¶O
                       'Else
                       'end 2019/4/24
   '                        NowPrint strCP09, "05", "10", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                           ET03 = "10" 'Modify By Sindy 2012/1/13
                       'End If 'cancel by sonia 2019/4/24
'                   End If
               '2010/4/8 modify by sonia ­^¤å©w½Z²¾¤U¨Ó
               ElseIf textPrint = "3" Then
                  '2005/11/11 add BY SONIA ¥[¤J©w½Z»y¤å§PÂ_
   '               NowPrint strCP09, "05", "17", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                  ET03 = "17" 'Modify By Sindy 2012/1/13
                  'Ä¶¤å
   '               NowPrint strCP09, "05", "18", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                  ET03_1 = "18" 'Modify By Sindy 2012/1/13
               '2010/4/8 end
               End If
            End If
         ' ¥Ó½Ð°ê®a¬°¤j³°
         ElseIf m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
               'Add By Sindy 2009/10/23
               If Val(Trim(Me.textMoney.Text)) = 0 Then
               '2009/10/23 End
'                  NowPrint strCP09, "05", "03", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                  If str1006CP64 = "" Then 'Added by Lydia 2017/02/02 ¥x-¤j°Ï¤À¤@¯ëµù¥UÃÒ©M³¡¥÷ºM¾Pµù¥UÃÒ
                     'modify by sonia 2019/1/30 ¦A¤À¹q¤lÃÒ®Ñ,¯È¥»ÃÒ®Ñ
                     'ET03 = "03" 'Modify By Sindy 2012/1/13
                     If Option5(0).Value = True Then
                        ET03 = "25"
                     Else
                        ET03 = "03"
                     End If
                     'end 2019/1/30
                  Else
                     'modify by sonia 2019/1/30 ¦A¤À¹q¤lÃÒ®Ñ,¯È¥»ÃÒ®Ñ
                     'ET03 = "24"
                     If Option5(0).Value = True Then
                        ET03 = "26"
                     Else
                        ET03 = "24"
                     End If
                     'end 2019/1/30
                  End If 'end 2017/02/02
               Else
                  '¼g¦b³ø»ù©w½Z¸Ì(PUB_Cache2Letter)
               End If
            'add by nickc 2007/07/24 ¥[¤J­^¤å
            ElseIf textPrint = "3" Then
'               NowPrint strCP09, "05", "19", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
               ET03 = "19" 'Modify By Sindy 2012/1/13
               'Ä¶¤å
'               NowPrint strCP09, "05", "21", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
               ET03_1 = "21" 'Modify By Sindy 2012/1/13
            End If
         End If
      Case "TF":
            'add by nickc 2006/06/30
            If textPrint = "1" And m_TM04 = "00" And m_TM03 = "0" Then
                ' ¥»©Ò®×¸¹²Ä¤E½X
                If Mid(m_TM02, 6, 1) = "0" Then
'                   NowPrint strCP09, "05", "04", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                  ET03 = "04" 'Modify By Sindy 2012/1/13
                Else
'                   NowPrint strCP09, "05", "05", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                  ET03 = "05" 'Modify By Sindy 2012/1/13
                End If
            End If
      Case "TC":
         ' ¥Ó½Ð°ê®a¬°¤j³°
         If m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
               'Add By Sindy 2014/11/28
               If Val(Trim(Me.textMoney.Text)) = 0 Then
               '2014/11/28 End
                  '93.12.9 MODIFY BY SONIA ­ì¥¼°Ï¤À§@«~ºØÃþ, ¥[¤J¬ü³NµÛ§@(08)¤§©w½Z
                  Select Case Trim(m_SP46)
                      Case "¬ü³NµÛ§@":
   '                           NowPrint strCP09, "05", "08", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                           ET03 = "08" 'Modify By Sindy 2012/1/13
                      Case "­pºâ¾÷³n¥ó":
   '                           NowPrint strCP09, "05", "06", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                           ET03 = "06" 'Modify By Sindy 2012/1/13
                      '2010/9/8 ADD BY SONIA
                      Case Else
   '                           NowPrint strCP09, "05", "09", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
                           ET03 = "09" 'Modify By Sindy 2012/1/13
                  End Select
               Else
                  '¼g¦b³ø»ù©w½Z¸Ì(PUB_Cache2Letter)
               End If
            End If
            '93.12.9 end
         ' ¥Ó½Ð°ê®a¬°¥xÆW
         ElseIf m_TM10 < "010" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
'                  NowPrint strCP09, "05", "07", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
               ET03 = "07" 'Modify By Sindy 2012/1/13
            'Add By Sindy 2010/01/20 ¤j->¥x
            ElseIf textPrint = "2" And m_CP10 = "806" Then
   '            NowPrint m_CP09, "05", "01", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
               ET03 = "01" 'Modify By Sindy 2012/1/13
            '2010/01/20 End
            End If
         End If
   End Select
   
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         '§PÂ_¬O§_EMail¦P®É±H¯È¥»
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + «H¨çÁ`¦¬¤å¸¹
         If strSrvDate(1) >= T°Ó¼Ð¹q¤l¤Æ²Ä2¶¥¬q±Ò¥Î¤é Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
            If ET03_1 <> "" Then
               NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
            End If
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            If ET03_1 <> "" Then
               NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            End If
            MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_TM01) & " ]¡I"
         End If
      Else
         'Add By Sindy 2019/12/19 + strLD18.«H¨çÁ`¦¬¤å¸¹
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         If ET03_1 <> "" Then
            'Add By Sindy 2019/12/19 + strLD18.«H¨çÁ`¦¬¤å¸¹
            NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         End If
      End If
      
   'Added by Lydia 2016/12/22 ¤£¥X©w½Z,¨ú®øDÃþ¦¬¤å±±¨î
   Else
      'Add By Sindy 2021/1/5 ¨S¦³¨t²Î²£¥Xªº©w½Z
      'Add By Sindy 2021/2/1 ¸ß°Ý¦³¨S¦³«È¤á¨ç
      If strLD18 <> "" Then
         If Val(Trim(Me.textMoney.Text)) = 0 Then 'Add By Sindy 2024/9/18 ±Æ°£¦³³ø»ù©w½Z
            Call PUB_TCaseAskIsPost_C(strLD18)
         End If
      End If
      '2021/1/5 EMD
   
      m_ULD02 = ""
      bolA1kdataMail = False
      'Modified by Lydia 2017/04/06
      'm_AC2470 = ""
      m_rA1k28 = ""
      m_rSpec = ""
      'end 2017/04/06
   'end 2016/12/22
   
   End If
   '2012/1/13 End
End Sub
   
'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse

TxtValidate = False

'Add By Sindy 2010/12/24
If Me.textTM15.Enabled = True Then
   Cancel = False
   textTM15_Validate Cancel
   If Cancel = True Then
      textTM15.SetFocus
      Exit Function
   End If
End If

If Me.textDate.Enabled = True Then
   Cancel = False
   textDate_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textMoney.Enabled = True Then
   Cancel = False
   textMoney_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPS.Enabled = True Then
   Cancel = False
   textPS_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Modify By Sindy 2020/12/29
If FrameTM20.Visible = True Then
   If Me.textTM20.Enabled = True Then
      Cancel = False
      textTM20_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Else
   If Me.textTM14.Enabled = True Then
      Cancel = False
      textTM14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
End If
'2020/12/29 END

If Me.textTM21.Enabled = True Then
   Cancel = False
   textTM21_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM22.Enabled = True Then
   Cancel = False
   textTM22_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'2005/4/14 ADD BY SONIA
If Me.textCP47.Enabled = True Then
   Cancel = False
   textCP47_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'2005/4/14 END

If m_TM01 = "T" Or m_TM01 = "TF" Then
   If m_TM10 < "010" Then
      strTmp = m_TM22
   Else
      strTmp = TransDate(m_TM22, 2)
   End If
   If textTM22 <> strTmp Then
      strTit = "¸ê®ÆÀË®Ö"
      'Modified by Lydia 2019/12/09 +³Æµù
      'strMsg = "±M¥Î´Á­­¤î¤éÀ³¬°<" & strTmp & ">"
      strMsg = "±M¥Î´Á­­¤î¤éÀ³¬°<" & strTmp & ">¡A¬O§_Ä~Äò§@·~¡H"
      nResponse = MsgBox(strMsg, vbOKCancel, strTit)
      If nResponse = vbCancel Then Cancel = True: Exit Function
   End If
End If

If ChkTM136(True) = False Then Exit Function 'Added by Morgan 2025/2/18

''Add By Sindy 2020/12/14 T¥xÆW®×­n¿é¤J©w½Z¤é´Á
'If m_TM01 = "T" And m_TM10 < "010" Then
'   If Me.textFinalDate.Text = "" Then
'      MsgBox "½Ð¿é¤J©w½Z¤é´Á!!!", vbExclamation + vbOKOnly
'      Me.textFinalDate.SetFocus
'      Exit Function
'   Else
'      Cancel = False
'      textFinalDate_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
'End If
''2020/12/14 END

TxtValidate = True
End Function

'Add By Cheng 2004/02/06
'§PÂ_´¿³QÄ³²§¬O§_µLµ²ªG
Private Function GetNoResult(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'¹w³]¦³µ²ªG
GetNoResult = False
'§ì²§Ä³µªÅG(602)ªº¸ê®Æ
StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(strCP01 & strCP02 & strCP03 & strCP04) & " And CP10 ='602' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    '¹w³]µLµ²ªG
    GetNoResult = True
    Do While Not rsA.EOF
        '­Y¹ê»Úµ²ªG¦³­È, ³]¦¨¦³µ²ªG
        If "" & rsA("CP24").Value <> "" Then GetNoResult = False: Exit Do
        rsA.MoveNext
    Loop
Else
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '§ì³Q²§Ä³(1601), ³Q²§Ä³²z¥Ñ(1602)ªº¸ê®Æ
    StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(strCP01 & strCP02 & strCP03 & strCP04) & " And CP10 In ('1601','1602') "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        '¹w³]µLµ²ªG
        GetNoResult = True
        Do While Not rsA.EOF
            '­Y¹ê»Úµ²ªG¦³­È, ³]¦¨¦³µ²ªG
            If "" & rsA("CP24").Value <> "" Then GetNoResult = False: Exit Do
            rsA.MoveNext
        Loop
    Else
        '¹w³]¬°¦³µ²ªG
        GetNoResult = False
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Function GetCP14BYAClass(oCP01 As String, oCP02 As String, oCP03 As String, oCP04 As String) As String
   GetCP14BYAClass = ""
   '2010/9/28 ADD BY SONIA §PÂ_¸Ó©Ó¿ì¤H­YÂ÷Â¾§ï§ìP2001
   'strSql = "select cp14  From caseprogress where cp09 in (select min(cp09) from caseprogress where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' ) "
   strSql = "select cp14,ST04  From caseprogress,STAFF where cp09 in (select min(cp09) from caseprogress where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' ) AND CP14=ST01(+) "
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      '­Y¦³¸ê®Æ
      If .RecordCount > 0 Then
         GetCP14BYAClass = CheckStr(.Fields("cp14").Value)
         '2010/9/28 ADD BY SONIA §PÂ_­ì©Ó¿ì¤H­YÂ÷Â¾§ï§ìP2001
         If "" & .Fields("ST04") = "2" Then GetCP14BYAClass = "P2001"
         '2010/9/28 END
      End If
   End With
   CheckOC3
End Function

' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
Private Sub QueryMonTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   
   m_blnReceiveSecond = False '2011/9/22 add by sonia
   ' ¨ú±o°Ó¼Ð°ò¥»ÀÉªº¬ÛÃö¶µ¥Ø
   strSql = "SELECT * FROM TradeMark,divisioncase " & _
            "WHERE dc01 = '" & m_TM01 & "' AND " & _
                  "dc02 = '" & m_TM02 & "' AND " & _
                  "dc03 = '" & m_TM03 & "' AND " & _
                  "dc04 = '" & m_TM04 & "' and dc05=tm01(+) and dc06=tm02(+) and dc07=tm03(+) and dc08=tm04(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTM12 = CheckStr(rsTmp.Fields("TM12"))         '2008/10/24 ADD BY SONIA ¤À³Î¤l®×¥Ó½Ð®×¸¹¹w³]¥À®×¥Ó½Ð®×¸¹
      textTM14 = (CheckStr(rsTmp.Fields("TM14")))
      textTM21 = (CheckStr(rsTmp.Fields("TM21")))
      m_TM21 = textTM21
      textTM22 = (CheckStr(rsTmp.Fields("TM22")))
      m_TM22 = textTM22
      m_MonTM01 = CheckStr(rsTmp.Fields("tm01"))
      m_MonTM02 = CheckStr(rsTmp.Fields("tm02"))
      m_MonTM03 = CheckStr(rsTmp.Fields("tm03"))
      m_MonTM04 = CheckStr(rsTmp.Fields("tm04"))
      '2011/9/22 ADD BY SONIA ¥À®×­Y¤£ºÞ¨î²Ä¤G´Á,¤À³Î®×¤]¤£ºÞ¨î
      If InStr("" & rsTmp.Fields("TM58"), "²Ä¤G´Á") > 0 Then
         m_blnReceiveSecond = True
      End If
      '2011/9/19 end
      If textNP08.Enabled = True And textNP09.Enabled = True Then
           strSql = "SELECT * FROM nextprogress " & _
                    "WHERE np02 = '" & m_MonTM01 & "' AND " & _
                         " np03 = '" & m_MonTM02 & "' AND " & _
                         " np04 = '" & m_MonTM03 & "' AND " & _
                         " np05 = '" & m_MonTM04 & "' and np06 is null and np07=202 "
          rsTmp.Close
          rsTmp.CursorLocation = adUseClient
          rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If rsTmp.RecordCount > 0 Then
              m_MonNP08 = CheckStr(rsTmp.Fields("np08"))
              m_MonNP09 = CheckStr(rsTmp.Fields("np09"))
          End If
      End If
   End If
   
   '2011/9/22 add by sonia ¥À®×¬O§_¤w¦¬²Ä¤G´Á
   If m_blnReceiveSecond = False Then
      strSql = "Select * From Caseprogress Where " & ChgCaseprogress(m_MonTM01 & m_MonTM02 & m_MonTM03 & m_MonTM04) & " And (CP10='716' OR CP10='717')"
      rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then m_blnReceiveSecond = True
   End If
   '2011/9/22 end
   
   '2011/9/19 add by sonia §ì»P¥À®×ÂI¿ï¦¬¤å¸¹¤§¬Û¦P®×¥ó©Ê½èªº¤l®×¦¬¤å¸¹T-175229(§_«h¤l®×T-175230·|§ì¨ì²§Ä³µªÅG602)
   strSql = "SELECT c1.cp09,c1.cp10,c2.cp09 FROM CaseProgress c1,caseprogress c2 WHERE c1.CP09= '" & frm02010401_6.oKey & "' " & _
            "and c2.cp01='" & m_TM01 & "' and c2.cp02='" & m_TM02 & "' and c2.cp03='" & m_TM03 & "' and c2.cp04='" & m_TM04 & "' and c1.cp10=c2.cp10 "
   rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(2)) = False Then
         m_CP09 = rsTmp.Fields(2)
      End If
      If IsNull(rsTmp.Fields(1)) = False Then
         m_CP10 = rsTmp.Fields(1)
      End If
   End If
   '2011/9/19 END
   
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2012/5/18
Private Sub Option1_Click(Index As Integer)
   If Me.Option4(0).Value Then
      Text10_Validate False
   ElseIf Me.Option4(1).Value Then
      Text11_Validate False
   ElseIf Me.Option4(2).Value Then
      Text12_Validate False
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_LostFocus()
   '«D¥xÆW"¤Ñ"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
      If textNP08.Enabled = True Then textNP08.SetFocus
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
   CloseIme
End Sub

Private Sub Text11_LostFocus()
   '«D¥xÆW"¤ë"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   'If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
   '   If textNP08.Enabled = True Then textNP08.SetFocus
   'End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   '«D¥xÆW"¤é"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
      If textNP08.Enabled = True Then textNP08.SetFocus
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
   Else
      If ChkDate(Text12) Then
         If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "¨Ó¨ç´Á­­¤£¥i¤p©ó¨t²Î¤é !", vbCritical
               Cancel = True
            Else
               textNP09 = Text12
               'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
               If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                  textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
               Else
               '2014/10/6 END
                  textNP08 = TransDate(CompDate(2, -2, TransDate(textNP09, 2)), 1)
               End If
               textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub GetTime()
   Dim i As Integer
   Dim strFromDate As String '´Á­­°_ºâ¤é
   
   'Add By Sindy 2012/8/30
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   '2012/8/30 End
   
   strFromDate = DBDATE(textCP05S)
   
   If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
      '¤å¨ì¤Ñ¼Æ
      If Option4(0).Value = True Then
         textNP09 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         If Option1(0).Value = True Then textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      '¤å¨ì¤ë¼Æ
      ElseIf Option4(1).Value = True Then
         textNP09 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
         If Option1(0).Value = True Then textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
      If textNP09 <> "" Then
         'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
         If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
            textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
         Else
         '2014/10/6 END
            textNP08 = TransDate(CompDate(2, i, TransDate(textNP09, 2)), 1)
         End If
      End If
      textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
   End If
End Sub

'Åª¨ú¨Ó¨ç´Á­­
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '´Á­­°_ºâ¤é
   
   strFromDate = DBDATE(textCP05S)
   
   ChgType = False
   If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   ' ®×¥ó©Ê½è
   strRvType = LabNP07.Caption '202.¥Ó½Ð·N¨£®Ñ
   If strRvType = "" Then Exit Function
   
   If ClsPDGetCaseProperty(m_TM01, strRvType, strTempName, bolTmp) Then
      textNP08 = ""
      textNP09 = ""
      
      If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & strRvType & "'"
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  If Not IsNull(.Fields(1)) Then
                     '¤å¨ì¤Ñ¼Æ
                     Option4(0).Value = True
                     Text10 = .Fields(1)
                     textNP09 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
                  ElseIf Not IsNull(.Fields(2)) Then
                     '¤å¨ì¤ë¼Æ
                     Option4(1).Value = True
                     Text11 = .Fields(2)
                     textNP09 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
                  Else
                     '¤å¨ì¤Ñ¼Æ
                     Option4(0).Value = True
                     Text10 = ""
                     Text11 = ""
                  End If
                  If textNP09 <> "" And Not IsNull(.Fields(0)) Then
                     '¤å¨ì·í¤é
                     If .Fields(0) = "1" Then
                        Option1(0).Value = True
                        textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
                     '¤å¨ì¦¸¤é
                     Else
                        Option1(1).Value = True
                     End If
                  End If
                  '¤å¨ì¤Ñ¼Æ
                  If Text10 <> "" Then
                     If Val(Text10) >= 60 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  '¤å¨ì¤ë¼Æ
                  ElseIf Not IsNull(.Fields(2)) Then
                     If Val(.Fields(2)) >= 2 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  End If
                  If textNP09 <> "" Then
                     'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
                     If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                        textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
                     Else
                     '2014/10/6 END
                        textNP08 = TransDate(CompDate(2, i, TransDate(textNP09, 2)), 1)
                     End If
                  End If
                  textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function

'Added by Lydia 2020/07/07
Private Sub textNP08_GotFocus()
   InverseTextBox textNP08
End Sub
'¥»©Ò´Á­­
Private Sub textNP08_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Cancel = False
   If IsEmptyText(textNP08) = False Then
      ' ÀË¬d¬O§_¬°¥Á°ê¦~
      If CheckIsTaiwanDate(textNP08, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªº¥»©Ò´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP08.SetFocus
         textNP08_GotFocus
      'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
      Else
          textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08, True), 1)
      'end 2020/07/07
      End If
   End If
End Sub
Private Sub textNP09_GotFocus()
   InverseTextBox textNP09
End Sub
' ªk©w´Á­­
Private Sub textNP09_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Cancel = False
   If IsEmptyText(textNP09) = False Then
      ' ÀË¬d¬O§_¬°¥Á°ê¦~
      If CheckIsTaiwanDate(textNP09, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªºªk©w´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP09_GotFocus
      End If
   End If
End Sub
'end 2020/07/07

'Added by Morgan 2025/2/18
Private Function ChkTM136(Optional pReset As Boolean) As Boolean
   If m_TM10 = "000" And m_CP10 <> "308" Then
      '¥i¯à¶i¥»µe­±«á¤~¥h°ò¥»ÀÉ§ï³]©w¡A¬G§ì°ò¥»ÀÉ³Ì·s³]©w
      If pReset Then
         strExc(0) = "select tm136 from trademark " & _
             "WHERE TM01 = '" & m_TM01 & "' AND " & _
                              "TM02 = '" & m_TM02 & "' AND " & _
                              "TM03 = '" & m_TM03 & "' AND " & _
                              "TM04 = '" & m_TM04 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_TM136 = "" & RsTemp(0)
         End If
      End If
      If (m_DocNo <> "" And m_TM136 = "2") Or (m_DocNo = "" And m_TM136 = "1") Then
         MsgBox "¥»¦¸¿é¤JªºÃÒ®Ñ«¬¦¡¡i" & IIf(m_DocNo = "", "¯È¥»", "¹q¤l") & "¡j»P°ò¥»ÀÉ³]©w¡i" & IIf(m_TM136 = "2", "¯È¥»", "¹q¤l") & "¡j¤£¦P¡A½Ð½T»{¡I", vbExclamation
         Exit Function
      End If
   End If
   ChkTM136 = True
End Function
