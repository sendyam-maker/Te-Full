VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_07 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "µo¤å(ÅÜ§ó, §ó§ï)"
   ClientHeight    =   5892
   ClientLeft      =   4740
   ClientTop       =   4128
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5892
   ScaleWidth      =   9156
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2959
      Width           =   2532
   End
   Begin VB.TextBox textCP118 
      Height          =   285
      Left            =   4290
      MaxLength       =   1
      TabIndex        =   7
      Top             =   5010
      Width           =   375
   End
   Begin VB.TextBox textCP113 
      Height          =   285
      Left            =   8055
      MaxLength       =   4
      TabIndex        =   2
      Top             =   3997
      Width           =   600
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   6165
      MaxLength       =   1
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5010
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox textCP84 
      Alignment       =   1  '¾a¥k¹ï»ô
      Height          =   285
      Left            =   5700
      TabIndex        =   1
      Top             =   3997
      Width           =   765
   End
   Begin VB.TextBox textAdd 
      Height          =   285
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   5
      Top             =   4671
      Width           =   852
   End
   Begin VB.TextBox textDN 
      Height          =   285
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   3
      Top             =   4334
      Width           =   492
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1274
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1611
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1611
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   937
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1274
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   937
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textTM27 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5820
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3300
      Width           =   2412
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4671
      Width           =   780
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   6
      Top             =   5010
      Width           =   372
   End
   Begin VB.TextBox textUargeDate 
      Height          =   285
      Left            =   5700
      MaxLength       =   7
      TabIndex        =   4
      Top             =   4334
      Width           =   1092
   End
   Begin VB.TextBox textCP27 
      Height          =   285
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   3997
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¦P®Éµo¤å(&N)"
      Height          =   400
      Index           =   1
      Left            =   2700
      TabIndex        =   9
      Top             =   60
      Width           =   1092
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "¬ÛÃö¨÷¸¹(&F)"
      Height          =   400
      Left            =   3840
      TabIndex        =   10
      Top             =   60
      Width           =   1092
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "ÅÜ§ó¨Æ¶µ(&R)"
      Height          =   400
      Left            =   4992
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   11
      Top             =   60
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   13
      Top             =   60
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6120
      TabIndex        =   12
      Top             =   60
      Width           =   852
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8160
      TabIndex        =   14
      Top             =   60
      Width           =   852
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   1200
      TabIndex        =   69
      Top             =   3300
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1200
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1948
      Width           =   2535
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5700
      TabIndex        =   67
      Top             =   1948
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7410
      TabIndex        =   66
      Top             =   4410
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
   Begin MSForms.TextBox textCP64 
      Height          =   495
      Left            =   1170
      TabIndex        =   8
      Top             =   5340
      Width           =   7815
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13785;873"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   64
      Top             =   2285
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   285
      Left            =   5700
      TabIndex        =   63
      Top             =   2285
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   285
      Left            =   1200
      TabIndex        =   62
      Top             =   2622
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   285
      Left            =   5700
      TabIndex        =   61
      Top             =   2622
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM81 
      Height          =   285
      Left            =   1200
      TabIndex        =   60
      Top             =   2959
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1170
      TabIndex        =   59
      Top             =   3660
      Width           =   7875
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13891;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_¹q¤l°e¥ó:          (Y: ¬O)"
      Height          =   285
      Left            =   3120
      TabIndex        =   58
      Top             =   5010
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¤u§@®É¼Æ:"
      Height          =   285
      Index           =   12
      Left            =   7290
      TabIndex        =   57
      Top             =   3997
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H5 :"
      Height          =   285
      Left            =   120
      TabIndex        =   56
      Top             =   2966
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H4 :"
      Height          =   285
      Left            =   4740
      TabIndex        =   55
      Top             =   2634
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H3 :"
      Height          =   285
      Left            =   120
      TabIndex        =   54
      Top             =   2634
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H2 :"
      Height          =   285
      Left            =   4740
      TabIndex        =   53
      Top             =   2302
      Width           =   720
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "¥X¦W¥N²z¤H"
      Height          =   285
      Left            =   6480
      TabIndex        =   52
      Top             =   5010
      Width           =   900
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "µo¤å³W¶O¡G"
      Height          =   285
      Left            =   4740
      TabIndex        =   50
      Top             =   3997
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(1:©e¥ôª¬ 2:ÅÜ§óÃÒ©ú)"
      Height          =   285
      Left            =   2640
      TabIndex        =   49
      Top             =   4671
      Width           =   1695
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_¸É¥ó(¥i½Æ¿ï) :"
      Height          =   285
      Left            =   120
      TabIndex        =   48
      Top             =   4671
      Width           =   1470
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_¿é¤JD/N :"
      Height          =   285
      Left            =   90
      TabIndex        =   47
      Top             =   4334
      Width           =   1095
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "(Y:¿é¤J)"
      Height          =   285
      Left            =   2010
      TabIndex        =   46
      Top             =   4334
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¼f©w¸¹¼Æ :"
      Height          =   285
      Left            =   120
      TabIndex        =   45
      Top             =   1306
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "´¼Åv¤H­û :"
      Height          =   285
      Index           =   11
      Left            =   4740
      TabIndex        =   44
      Top             =   1970
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "©¼©Ò®×¸¹ :"
      Height          =   285
      Index           =   9
      Left            =   4740
      TabIndex        =   43
      Top             =   1638
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó©Ê½è :"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   42
      Top             =   1638
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¦¬¤å¸¹ :"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   41
      Top             =   642
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹ :"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   974
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð®×¸¹ :"
      Height          =   285
      Left            =   4740
      TabIndex        =   39
      Top             =   1306
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "µoÃÒ¤é :"
      Height          =   285
      Index           =   3
      Left            =   4740
      TabIndex        =   38
      Top             =   974
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "·~°È°Ï§O :"
      Height          =   285
      Index           =   2
      Left            =   4740
      TabIndex        =   37
      Top             =   642
      Width           =   810
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "©Ó¿ì¤H :"
      Height          =   285
      Left            =   120
      TabIndex        =   36
      Top             =   1970
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥¿°Ó¼Ð¸¹¼Æ:"
      Height          =   285
      Index           =   8
      Left            =   4740
      TabIndex        =   35
      Top             =   3300
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "°Ó¼ÐºØÃþ :"
      Height          =   285
      Index           =   4
      Left            =   4740
      TabIndex        =   34
      Top             =   2966
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H1 :"
      Height          =   285
      Left            =   120
      TabIndex        =   33
      Top             =   2302
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó¦WºÙ :"
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   3712
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÂI¡@¡@¼Æ :"
      Height          =   285
      Index           =   10
      Left            =   4830
      TabIndex        =   31
      Top             =   4671
      Width           =   810
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:¤£¦L)"
      Height          =   285
      Left            =   1680
      TabIndex        =   30
      Top             =   5010
      Width           =   645
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "¦C¦L©w½Z :"
      Height          =   285
      Left            =   120
      TabIndex        =   29
      Top             =   5010
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "¥N²z¤H :"
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   3300
      Width           =   630
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "¶Ê¼f´Á­­ :"
      Height          =   285
      Left            =   4830
      TabIndex        =   27
      Top             =   4334
      Width           =   810
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "µo¤å¤é :"
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   3997
      Width           =   630
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "¶i«×³Æµù :"
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   5340
      Width           =   810
   End
End
Attribute VB_Name = "frm030202_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/02 §ï¦¨Form2.0 ; cmbTM05¡BtextTM23¡BtextTM78~81¡BtextCP13¡BtextCP14¡BtextCP64¡BtextTM44¡BlstNameAgent
'Memo By Sindy 2012/12/4 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/16 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/29 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo By Sindy 2010/8/11 ¤é´ÁÄæ¤w­×§ï
Option Explicit

' ¥»©Ò®×¸¹
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' ¦¬¤å¸¹
Dim m_CP09 As String
Dim m_CP43 As String '¬ÛÃöÁ`¦¬¤å¸¹
' ¥Ó½Ð°ê®a
Dim m_TM10 As String
' ®×¥ó©Ê½è¥N¸¹
Dim m_CP10 As String
' ©Ó¿ì¤H Add By Sindy 98/03/11
Dim m_CP14 As String
Dim m_CP82 As String 'Added by Lydia 2018/08/10 µo¤å®É¶¡

' «Å§iÄæ¦ì¤º®eµ²ºc
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' Àx¦s°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÀÉ®×Äæ¦ìªº¦ê¦C
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' Àx¦s®×¥ó¶i«×ÀÉÀÉ®×Äæ¦ìªº¦ê¦C
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '¬O§_«ö¤UÅÜ§ó¨Æ¶µ«ö¶s
'add by nick 2004/08/13
Dim m_CP84 As String       'µo¤å³W¶O
'add by nickc 2006/01/26
Dim m_CP110 As String
'add by nickc 2008/02/22
Dim m_CP44 As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
'Add By Sindy 2009/06/03
Dim m_TM23 As String
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
'2009/06/03 End
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 ¦¬¤å¸¹,¬O§_ºâµo¤å«Ç®×¥ó
Dim m_CP130s As String 'Add by Sindy 2009/4/24 µo¤å-¥DºÞ¾÷Ãö
Dim m_strLanguage As String 'Add By Sindy 2012/10/12
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim m_IsSend As Boolean 'Add By Sindy 2018/11/23 ¬O§_¸gµo¤å«Çµo¤å
Dim m_CP148 As String '¬O§_¬°¤@¥Ó½Ð®Ñ¦h¥ó Add By Sindy 2019/3/26
Dim m_CP28 As String, m_CP27 As String 'Add By Sindy 2019/3/26
Dim bolCase102_1001 As Boolean 'Add By Sindy 2019/10/8 ¬O§_¬°©µ®i®Ö­ã


Private Sub cmdCancel_Click()
   frm030202_01.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Unload frm030202_01
   Unload Me
   'frm030202_01.Show
End Sub

Private Sub cmdMod_Click()
   frm030202_05.SetData 0, m_TM01, True
   frm030202_05.SetData 1, m_TM02, False
   frm030202_05.SetData 2, m_TM03, False
   frm030202_05.SetData 3, m_TM04, False
   frm030202_05.SetData 4, m_CP09, False
   'Add By Sindy 2009/06/03
   frm030202_05.SetData 5, m_TM23, False
   frm030202_05.SetData 6, m_TM78, False
   frm030202_05.SetData 7, m_TM79, False
   frm030202_05.SetData 8, m_TM80, False
   frm030202_05.SetData 9, m_TM81, False
   If textCP27.Text = "" Then
      frm030202_05.SetData 10, strSrvDate(1), False
   Else
      frm030202_05.SetData 10, DBDATE(Trim(textCP27.Text)), False
   End If
   '2009/06/03 End
   
   'frm030202_05.SetParent Me
   frm030202_05.SetParent "frm030202_07"
   'Me.Hide
   frm030202_05.Show
   frm030202_05.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim bolIsCaseNum As Boolean 'Add By Sindy 2018/11/23
   
   'Modify By Sindy 2010/11/19 §â¡u½T©w¡v¤Î¡u¦P®Éµo¤å¡v«ö¶sµ{¦¡½X¦X¨Ö
   Select Case Index
      Case 0, 1
         If CheckDataValid = True Then
            'Add By Sindy 2022/2/7 T©MFCTªº´îÁY°Ó«~313µo¤å¡G
            '¦Û°Ê·s¼WÅÜ§ó¨Æ¶µÀÉ¨Ã¼g¤JCE01(Á`¦¬¤å¸¹)¡ACE45(¡§´îÁY°Ó«~½Ð°Ñ¨÷©v°Ïªþ¥ó¡¨)¡A­Y¤w¸Ó¦¬¤å¸¹¤w¦s¦b©óÅÜ§ó¨Æ¶µÀÉ«h§ó·s¡C
            If m_CP10 = "313" Then
               If IsChangeEventExist(m_CP09) = True Then
                  strSql = "Update CHANGEEVENT Set ce45='´îÁY°Ó«~½Ð°Ñ¨÷©v°Ïªþ¥ó' Where ce01='" & m_CP09 & "'"
                  cnnConnection.Execute strSql
               Else
                  strSql = "insert into CHANGEEVENT(ce01,ce45) values ('" & m_CP09 & "','´îÁY°Ó«~½Ð°Ñ¨÷©v°Ïªþ¥ó')"
                  cnnConnection.Execute strSql
               End If
            End If
            '2022/2/7 END
            
            'Add By Cheng 2002/05/23
            '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
            If TxtValidate = False Then Exit Sub
            ' 90.08.29 ÀË¬dÅÜ§ó¨Æ¶µÀÉ
            '2009/4/2 modify by sonia 302§ó¥¿¥B¶i«×³Æµù¬° §ó§ïµù¥UÃÒ ©Î §ó§ï®Ö­ã¨ç ®É¤£ÀË¬d
            '2012/4/12 modify by sonia 302§ó¥¿³£¤£ÀË¬d-³¯ª÷½¬
            'If textCP10 = "§ó¥¿" And (InStr(1, textCP64, "§ó§ïµù¥UÃÒ", 1) > 0 Or InStr(1, textCP64, "§ó§ï®Ö­ã¨ç", 1) > 0) Then
            If textCP10 = "§ó¥¿" Then
            '2009/4/2 end
            ElseIf IsChangeEventExist(m_CP09) = True Then
            Else
               MsgBox "½Ð¥ý¿é¤JÅÜ§ó¨Æ¶µ¸ê®Æ!", vbCritical + vbOKOnly, "ÀË®Ö¸ê®Æ"
               Exit Sub
            End If
               
            'Add by Sindy 98/3/24 ³]©w¬O§_ºâµo¤å«Ç®×¥ó
            If m_TM10 = "000" Then
               'Modify By Sindy 2012/12/20 ­Y¬°¹q¤l°e¥ó«h¤£¸gµo¤å«Ç
               'Modify By Sindy 2023/8/1 ¹q¤l°e¥óÄæ¦ì­È¤£¬OªÅ¥ÕªÌ,§Y¬°¹q¤l°e¥ó
               If (textCP118.Visible = True And textCP118 <> "") Then
                  'Added by Morgan 2016/5/16 ¹q¤l°e¥ó¤]­n°O¿ý¥DºÞ¾÷Ãö
                  'Modify By Sindy 2018/11/23 + bolIsCaseNum:¬O§_ºâµo¤å«Ç¥ó¼Æ
                  If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True, bolIsCaseNum) = False Then
                     Exit Sub
                  End If
                  'end 2016/5/16
                  
                  'Add By Sindy 2018/11/23 ¦]¦³¤@¤å¦h®×ªº°ÝÃD¡A¦ý­Y¹q¤l°e¥ó§¡¤£¸gµo¤å«Ç
                  '                        ´NµLªk§PÂ_¬O§_­n¦L©w½Z, §ï§PÂ_¬O§_ºâµo¤å«Ç¥ó¼Æ
                  'Modify Sindy 2019/1/17 Mark,¹q¤l°e¥ó¤@©w­n¥X©w½Z,¤£¥Î¯S§O§PÂ_
'                  If bolIsCaseNum = True Then
                     m_IsSend = True
'                  Else
'                     m_IsSend = False
'                  End If
'                  'ÅÜ§óµù¥U(«á)¤é¤å©w½Z¤~­n¸ß°Ý
'                  If bolIsCaseNum = True And Trim(textPrint.Text) = "" And _
'                     m_CP10 = "301" And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3" And _
'                     Trim(textTM15.Text) <> "" Then
'                     If MsgBox("¬O§_»Ý­n¦C¦L©w½Z¡H", vbExclamation + vbYesNo) = vbNo Then
'                        textPrint.Text = "N"
'                     End If
'                  End If
'                  '2018/11/23 End
               
                  'add by sonia 2016/3/31
                  strExc(0) = Trim(InputBox("½Ð¿é¤J´¼¼z§½¦¬¤å¤å¸¹!!"))
                  If strExc(0) = "" Then
                     Exit Sub
                  Else
                     textCP64 = "´¼¼z§½¦¬¤å¤å¸¹:" & strExc(0) & ";" & Trim(textCP64)
                  End If
                  'end 2016/3/31
               Else
                  'Add by Sindy 2009/4/24
                  If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
                     Exit Sub
                  Else
                     'Add By Sindy 2018/11/23 ¦]¦³¤@¤å¦h®×ªº°ÝÃD¡A©Ò¥H­Y¸gµo¤å«Ç¥B§@·~µe­±¤W¬°¦C¦L©w½Z®É,¸ß°Ý¨Ï¥ÎªÌ
                     '§PÂ_¬O§_ºâµo¤å«Ç¥ó¼Æ
                     If m_CP123s = "Y" Then
                        m_IsSend = True
                     Else
                        m_IsSend = False
                     End If
                     'Modify By Sindy 2021/9/11 ­n¸ß°Ý,¤£µM¤@¤å¦h®×ªº²Ä¤@¥ó³£·|²£¥Í©w½Z,¥¦¬OµLªk¿ë»{¥X¬O­n¥X¦h¥óªº
                     'Modify By Sindy 2019/4/1 Mark
                     'ªü½¬»¡­n¼W¥[§PÂ_¬O¤é¤å©w½Z¤~»Ý­n¸ß°Ý
                     If m_CP123s = "Y" And Trim(textPrint.Text) = "" And _
                        m_CP10 = "301" And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3" And _
                        Trim(textTM15.Text) <> "" Then
                        If MsgBox("¬O§_»Ý­n¦C¦L©w½Z¡H" & vbCrLf & "(¤@¤å¦h®×»Ý¨ì<©w½Z¸ê®ÆºûÅ@>²£¥X©w½Z)", vbExclamation + vbYesNo) = vbNo Then
                           textPrint.Text = "N"
                        End If
                     End If
                     '2018/11/23 End
                                          
                     If m_CP123s = "Y" Then
                        'modify by sonia 2014/6/23 ¥[¶Çµo¤å³W¶O, P-108903
                        If ModifyDispatch(textCP09, m_CP09s, m_CP123s, textCP84, textCP27) = False Then
                            Exit Sub
                        End If
                     End If
                  End If
               End If '2012/12/20 End
            End If
            
            ' ³]©w·Æ¹«´å¼Ð¬°µ¥«Ýª¬ºA
            Screen.MousePointer = vbHourglass
            ' §ó·sÄæ¦ì¿é¤Jªº¤º®e
            OnUpdateField
            ' ¦sÀÉ
            'edit by  nick 2004/11/03
            'OnSaveData
            If OnSaveData = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            
            'Add By Sindy 2010/4/14
            ' ¦C¦L©w½Z
            If textPrint <> "N" Then
               If m_IsSend = True Then 'Modify By Sindy 2018/11/23 +if ¦]¦³¤@¤å¦h®×ª¬ªp¡A©Ò¥H¼W¥[§PÂ_¸gµo¤å«Ç®É¤~»Ý¥X©w½Z
                  If m_CP27 = "" Then m_CP27 = DBDATE(Me.textCP27) 'Added by Lydia 2024/11/14 debug:°õ¦æPUB_GetFCTAppendix_JP·|¥X¿ù
                  PrintLetter
               End If
            End If
            
            ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
            Screen.MousePointer = vbDefault
            
            'Add By Sindy 2012/4/5 CFT,FCT©Ò¦³®×¥ó©Ê½èµo¤å®É,ÀË¬d¥Nªí¹Ï¬O§_¦s¦b
            'Mark by Amy 2018/07/31 ¦]ChkIsExistImg¤£¨Ï¥Î,»PSindy½T»{FCT¤£¼uMsg¬G®³±¼
            'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)
            
            'Added by Lyddia 2018/08/10 ¼W¥[­«·sµo¤å§PÂ_
            strExc(1) = m_CP82
            If Val(m_CP82) > 0 Then
                 If MsgBox("­«·sµo¤å¬O§_¤W¶ÇÀÉ®×¨ì¨÷©v°Ï¡H", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     strExc(1) = ""
                 End If
            End If
            If Val(strExc(1)) = 0 Then
            'end 2018/08/10
                'Added by Lydia 2018/07/19 FCTµo¤å¦Û°Ê±N¤U¸üªºPDFÀÉ,¤W¶Ç¨ì¨÷©v°Ï
                If Pub_AutoSavePdf_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10) = False Then
                End If
                'end 2018/07/19
            End If 'end 2018/08/10
      
            '************   90.11.23 nick   ²Mµe­±
            'frm030202_01.radio(0).Value = True
            'frm030202_01.textCP09.Enabled = True
            'frm030202_01.textCP09.Text = ""
            'frm030202_01.textTM01.Enabled = False
            'frm030202_01.textTM01.Text = ""
            'frm030202_01.textTM02.Enabled = False
            'frm030202_01.textTM02.Text = ""
            'frm030202_01.textTM02_2.Enabled = False
            'frm030202_01.textTM02_2.Text = ""
            'frm030202_01.textTM03.Enabled = False
            'frm030202_01.textTM03.Text = ""
            'frm030202_01.textTM04.Enabled = False
            'frm030202_01.textTM04.Text = ""
            'frm030202_01.grdList.Clear
            'frm030202_01.grdList.Rows = 2
            'frm030202_01.QueryData
            'frm030202_01.Show
            '*************************************
               
            'Ken 91.04.09 -- Start
            If textDN = "Y" Then
               'Add By Cheng 2003/03/19
               '·s¼W¦a§}±ø¦Cªí¸ê®Æ
      'edit by nick 2004/11/17  ¦]¬°½Ð´Ú¤w¸g¦³²£¥Í¤F
      '            pub_AddressListSN = pub_AddressListSN + 1
      '            PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
               Screen.MousePointer = vbHourglass
               Frmacc21h0.Show
               mdiMain.ToolShow
               mdiMain.tool1_enabled
               Screen.MousePointer = vbDefault
               Set Frmacc21h0.frmlink = frm030202_01
               'add by nick 2004/11/24
               Frmacc21h0.IsPrintAddress = False
'            Else
'               'Add By Cheng 2002/04/30
'               '­Y¦³¥¼µo¤å¸ê®ÆÅã¥ÜÄµ§i
'               PUB_GetCPunIssueDatas "" & Me.textTMKey.Text
'
'               frm030202_01.Show
'               ' 90.12.07 modify by louis
'               frm030202_01.Clear1
            End If
            'Ken 91.04.09 -- End
            
            Call PUB_FCTSendRecvMail(m_CP09) 'Add By Sindy 2024/10/30 ¥~°Óµo¤å®É,¼W¥[µoMail³qª¾©Ó¿ì¤H¤Î°Æ¥»µ¹§Pµo¥DºÞ
            'Add By Sindy 2024/8/19
            If frm030202_01.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2024/8/19 End
            If Index = 0 Then '½T©wÁä
               'Ken 91.04.09 -- Start
               If textDN <> "Y" Then
                  'Add By Cheng 2002/04/30
                  '­Y¦³¥¼µo¤å¸ê®ÆÅã¥ÜÄµ§i
                  If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = True Then
                     frm030202_01.Show
                     ' 90.12.07 modify by louis
                     frm030202_01.Clear1
                  Else
                     'Add By Sindy 2024/8/19
                     If frm030202_01.bolIsEMPFlow = True Then
                        Unload frm030202_01
                        frm090202_4.Show
                     Else
                     '2024/8/19 End
                        frm030202_01.Show
                        frm030202_01.Clear1
                     End If
                  End If
               End If
               'Ken 91.04.09 -- End
               Unload Me
            ElseIf Index = 1 Then '¦P®Éµo¤åÁä
               If textDN <> "Y" Then
                  ' ©I¥s²Ä¤@­Óµe­±
                  frm030202_01.SetData 0, m_TM01, True
                  frm030202_01.SetData 1, m_TM02, False
                  frm030202_01.SetData 2, m_TM03, False
                  frm030202_01.SetData 3, m_TM04, False
                  frm030202_01.SetQueryFromTM
                  Unload Me
                  frm030202_01.Show
                  frm030202_01.radio(1).Value = True
                  frm030202_01.radio_Click 1
                  frm030202_01.QueryData
               Else
                  Unload Me
               End If
            End If
         End If
      Case Else
   End Select
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

'Private Sub cmdTogether_Click()
'   If CheckDataValid = True Then
'      'Add By Cheng 2002/07/15
'      '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
'      If TxtValidate = False Then Exit Sub
'
'      ' 90.08.29 ÀË¬dÅÜ§ó¨Æ¶µÀÉ
'      '2009/4/2 modify by sonia 302§ó¥¿¥B¶i«×³Æµù¬° §ó§ïµù¥UÃÒ ©Î §ó§ï®Ö­ã¨ç ®É¤£ÀË¬d
'      If textCP10 = "§ó¥¿" And (InStr(1, textCP64, "§ó§ïµù¥UÃÒ", 1) > 0 Or InStr(1, textCP64, "§ó§ï®Ö­ã¨ç", 1) > 0) Then
'      '2009/4/2 end
'      ElseIf IsChangeEventExist(m_CP09) = True Then
'      Else
'         MsgBox "½Ð¥ý¿é¤JÅÜ§ó¨Æ¶µ¸ê®Æ!", vbCritical + vbOKOnly, "ÀË®Ö¸ê®Æ"
'         Exit Sub
'      End If
'
'      'Add by Sindy 98/3/24 ³]©w¬O§_ºâµo¤å«Ç®×¥ó
'      If m_TM10 = "000" Then
'         'Add by Sindy 2009/4/24
'         If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
'            Exit Sub
'         Else
'            If m_CP123s = "Y" Then
'               If ModifyDispatch(textCP09, m_CP09s, m_CP123s, textCP27) = False Then
'                   Exit Sub
'               End If
'            End If
'         End If
'      End If
'
'      ' ³]©w·Æ¹«´å¼Ð¬°µ¥«Ýª¬ºA
'      Screen.MousePointer = vbHourglass
'      ' §ó·sÄæ¦ì¿é¤Jªº¤º®e
'      OnUpdateField
'      ' ¦sÀÉ
'      'edit by nick 2004/11/03
'      'OnSaveData
'      If OnSaveData = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'
'      ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
'      Screen.MousePointer = vbDefault
'
'      ' ©I¥s²Ä¤@­Óµe­±
'      frm030202_01.SetData 0, m_TM01, True
'      frm030202_01.SetData 1, m_TM02, False
'      frm030202_01.SetData 2, m_TM03, False
'      frm030202_01.SetData 3, m_TM04, False
'      frm030202_01.SetQueryFromTM
'      Unload Me
'      frm030202_01.Show
'      frm030202_01.radio(1).Value = True
'      frm030202_01.radio_Click 1
'      frm030202_01.QueryData
'   End If
'
'End Sub

'Private Sub Form_Activate()
'    'Add By Cheng 2003/10/06
'    '­Y¦³«ö¤UÅÜ§ó¨Æ¶µ«ö¶s, «h­«·sÅª¨ú¸ê®Æ
'    'edit by nickc 2005/08/23
'    'If m_blnClkChgButton = True Then
'    If m_blnClkChgButton = True Or (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'        pub_ModifyCaseNum = ""
'        QueryData
''        m_blnClkChgButton = False
'    End If
'End Sub

Private Sub Form_Load()
   ' ³]©w±±¨î¶µªº­I´ºÃC¦â
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'add by nickc 2007/01/29
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   textTM27.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   
   MoveFormToCenter Me
'    m_blnClkChgButton = False
   'Add by nickc 2006/01/26
   '¥xÆW¥[¥X¦W¥N²z¤H²M³æ¨Ñ¤Ä¿ï,­ì¬O§_¥X¦WÄæ¦ì¤£Åã¥Ü
   Text7.Visible = False
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Added by Lydia 2021/09/02 ¦pªG¤@¶}©l±NListBox©Ô¨ì»Ý­nªº¤j¤p¡A¦r«¬·|¦Û°Ê©ñ¤j¡F©Ò¥Hµe­±¹w³]¬°¤@¦C°ª«×¡AForm_Load¤~©ñ¤j¨ì»Ý­nªº¤j¤p
   lstNameAgent.Height = 825
   lstNameAgent.Width = 1300
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/4/17
   
   ' ²M°£·j´MªºKey
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' ¦¬¤å¸¹
      Case 0: m_CP09 = strData
         'Add By Sindy 2012/4/17
         strSql = "SELECT * FROM ChangeEvent " & _
                  "WHERE CE01 = '" & m_CP09 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            m_blnClkChgButton = True
         Else
            m_blnClkChgButton = False
         End If
         rsTmp.Close
   End Select
End Sub

' ²M°£°Ó¼Ð°ò¥»ÀÉÀÉ®×Äæ¦ì¦ê¦C
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' ³]©w°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' ³]©w°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' ²M°£®×¥ó¶i«×ÀÉÀÉ®×Äæ¦ì¦ê¦C
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' ³]©w®×¥ó¶i«×ÀÉÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' ³]©w®×¥ó¶i«×ÀÉÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' ¨ú±o°Ó¼Ð°ò¥»ÀÉªºÄæ¦ì¤º®e
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSubSQL As String
   Dim rsSubTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_TM44 = CheckStr(rsTmp.Fields("TM44"))
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      ' ¼f©w¸¹¼Æ
      If IsNull(rsTmp.Fields("TM15")) = False Then: textTM15 = rsTmp.Fields("TM15")
      ' ¥Ó½Ð®×¸¹
      If IsNull(rsTmp.Fields("TM12")) = False Then: textTM12 = rsTmp.Fields("TM12")
      ' µoÃÒ¤é
      If IsNull(rsTmp.Fields("TM20")) = False Then: textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      ' ®×¥ó¤¤¤å¦WºÙ
      If IsNull(rsTmp.Fields("TM05")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM05")
      ' ®×¥ó­^¤å¦WºÙ
      If IsNull(rsTmp.Fields("TM06")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM06")
      ' ®×¥ó¤é¤å¦WºÙ
      If IsNull(rsTmp.Fields("TM07")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM07")
      ' Åã¥Ü®×¥ó¦WºÙ
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' °Ó¼ÐºØÃþ
      textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      ' ¥Ó½Ð°ê®a
      If IsNull(rsTmp.Fields("TM10")) = False Then: m_TM10 = rsTmp.Fields("TM10")
      ' ¥Ó½Ð¤H
      If IsNull(rsTmp.Fields("TM23")) = False Then: textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      'add by nickc 2007/01/29
      If IsNull(rsTmp.Fields("TM78")) = False Then: textTM78 = GetCustomerName("" & rsTmp.Fields("TM78"), 0)
      If IsNull(rsTmp.Fields("TM79")) = False Then: textTM79 = GetCustomerName("" & rsTmp.Fields("TM79"), 0)
      If IsNull(rsTmp.Fields("TM80")) = False Then: textTM80 = GetCustomerName("" & rsTmp.Fields("TM80"), 0)
      If IsNull(rsTmp.Fields("TM81")) = False Then: textTM81 = GetCustomerName("" & rsTmp.Fields("TM81"), 0)
      'Add By Sindy 2009/06/03
      If IsNull(rsTmp.Fields("TM23")) = False Then: m_TM23 = rsTmp.Fields("TM23")
      If IsNull(rsTmp.Fields("TM78")) = False Then: m_TM78 = rsTmp.Fields("TM78")
      If IsNull(rsTmp.Fields("TM79")) = False Then: m_TM79 = rsTmp.Fields("TM79")
      If IsNull(rsTmp.Fields("TM80")) = False Then: m_TM80 = rsTmp.Fields("TM80")
      If IsNull(rsTmp.Fields("TM81")) = False Then: m_TM81 = rsTmp.Fields("TM81")
      
      ' ¥¿°Ó¼Ð¸¹¼Æ
      textTM27 = Empty
      If IsNull(rsTmp.Fields("TM27")) = False Then: textTM27 = rsTmp.Fields("TM27")
      ' FC¥N²z¤H
      textTM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then: textTM44 = GetFAgentName(rsTmp.Fields("TM44"))
      ' ©¼©Ò®×¸¹
      If IsNull(rsTmp.Fields("TM45")) = False Then: textTM45 = rsTmp.Fields("TM45")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub
' ¨ú±o®×¥ó¶i«×ÀÉªºÄæ¦ì¤º®e
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSubTmp As New ADODB.Recordset
   Dim strDate As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   ' ¨t²Î¤é
   strDate = DBDATE(SystemDate())
   ' ¦¬¤å¸¹
   textCP09 = m_CP09
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      m_CP44 = CheckStr(rsTmp.Fields("CP44"))
      m_CP82 = "" & rsTmp.Fields("CP82")  'Added by Lydia 2018/08/10 µo¤å®É¶¡
      ' ®×¥ó©Ê½è
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' ·~°È°Ï§O
      '910718 Sieg
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' ´¼Åv¤H­û
      If IsNull(rsTmp.Fields("CP13")) = False Then: textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      
      'Add By Sindy 98/03/11
      '¤u§@®É¼Æ
      textCP113 = "" & rsTmp.Fields("CP113")
      SetCPFieldOldData "CP113", textCP113, 1
      ' ©Ó¿ì¤H
      m_CP14 = "" & rsTmp.Fields("CP14")
      '98/03/11 End
      
      ' ©Ó¿ì¤H­û
      If IsNull(rsTmp.Fields("CP14")) = False Then: textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      ' µo¤å¤é(¹w³]¬°¨t²Î¤é)
      m_CP27 = Empty
      'Modify By Cheng 2004/02/04
      '­Yµo¤å¤éÄæ¦ìµL­È®É, ¹w³]¨t²Î¤é
'      textCP27 = TAIWANDATE(strDate)
      If Me.textCP27.Text = "" Then
          Me.textCP27.Text = strSrvDate(2)
      End If
      'End
      If IsNull(rsTmp.Fields("CP27")) = False Then: m_CP27 = rsTmp.Fields("CP27")
      SetCPFieldOldData "CP27", m_CP27, 1
      ' ÂI¼Æ
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      ' ¶i«×³Æµù
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      SetCPFieldOldData "CP64", textCP64, 0
      
      'Add By Sindy 2012/12/20
      ' ¬O§_¹q¤l°e¥ó
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      
      'add by nick 2004/08/13 µo¤å³W¶O
      If IsNull(rsTmp.Fields("CP17")) = False And textCP84.Enabled = True Then
          m_CP84 = CheckStr(rsTmp.Fields("CP17"))
      End If
      'Add By Sindy 2012/12/20 ¹q¤l°e¥óµo¤å³W¶O¹w³]¬°©Ó¿ì¤H¤w¿é¤Jªºª÷ÃB
      If rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/12/20
      
      'Add By Sindy 2010/4/14 ¬ÛÃöÁ`¦¬¤å¸¹
      m_CP43 = ""
      If IsNull(rsTmp.Fields("CP43")) = False Then
          m_CP43 = rsTmp.Fields("CP43")
      End If
      
      'Add By Sindy 2019/3/26
      'µo¤å¦r¸¹
      m_CP28 = "" & rsTmp.Fields("CP28")
      '¬O§_¬°¤@¥Ó½Ð®Ñ¦h¥ó
      m_CP148 = Empty
      If IsNull(rsTmp.Fields("CP148")) = False Then
         m_CP148 = rsTmp.Fields("CP148")
      End If
      '¼W¥[ÀË¬d¦Pµo¤å¦r¸¹¬O§_¦³¦h¥ó
      If m_CP148 = "Y" Then
         If PUB_ChkIsOneAppMuchCase(m_CP28) = False Then
            m_CP148 = Empty
         End If
      End If
      '2019/3/26 End
      
      'add by nickc 2006/02/10
      Text7 = CheckStr(rsTmp.Fields("CP22"))
      SetCPFieldOldData "CP22", Text7, 0
   End If
   'add by nickc 2006/01/26
   'SetCPFieldOldData "CP110", m_CP110, 0
   'Modify By Sindy 2010/9/20
   If m_CP110 = "" Then m_CP110 = CheckStr(rsTmp.Fields("cp110"))
   SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
   '2010/9/20 End
   rsTmp.Close
   
   'Add By Sindy 2019/10/8 §ó¥¿ªº¬ÛÃöÁ`¦¬¤å¸¹¬O§_¬°©µ®i®Ö­ã
   bolCase102_1001 = False
   If m_CP10 = "302" And Trim(m_CP43) <> "" Then
      strSql = "SELECT CP10 FROM CaseProgress" & _
               " WHERE CP09 in (SELECT CP43 FROM CaseProgress WHERE CP09 = '" & m_CP43 & "' and CP10='1001')" & _
               " AND CP10='102'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         bolCase102_1001 = True
      End If
      rsTmp.Close
   End If
   '2019/10/8 END
   
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' Åª¨ú¸ê®Æ®w
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   'add by nickc 2006/01/26
   Dim tm(1 To 4) As String
   ' ¥ý²M°£°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÄæ¦ì¦ê¦C
   ClearTMSPFieldList
   ' ¥ý²M°£®×¥ó¶i«×ÀÉÄæ¦ì¦ê¦C
   ClearCPFieldList
   
   ' ¥ý¨ú±o¥»©Ò®×¸¹
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' ¥»©Ò®×¸¹
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
   
   ' ¥»©Ò®×¸¹
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)
   
   'add by nickc 2006/01/26
   tm(1) = m_TM01
   tm(2) = m_TM02
   tm(3) = m_TM03
   tm(4) = m_TM04
   'Modify By Sindy 2010/9/20 ¹w³]¥X¦W¥N²z¤H,²¾¨ì¤U­±Åª§¹CP¦A°µ
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   '2010/9/20 End
   
   ' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
   QueryTradeMark
      
   ' ¨ú±o®×¥ó¶i«×ÀÉªºÄæ¦ì
   QueryCaseProgress
   'Modified by Lydia 2021/09/02 + Form 2.0 = True
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True 'Modify By Sindy 2010/9/20
   
   ' ¨ú±o¶Ê¼f´Á­­ªº¤é´Á
   textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
   Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
   
   'Add By Sindy 2012/12/20 ¥~°Ó000¥xÆW®×©Ò¦³®×¥ó©Ê½è¥[¹q¤l°e¥ó¥\¯à
   If m_TM01 = "FCT" And m_TM10 = "000" Then
      Label43.Visible = True
      textCP118.Visible = True
   Else
      Label43.Visible = False
      textCP118.Visible = False
   End If
   '2012/12/20 End
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030202_07 = Nothing
End Sub

'add by nickc 2006/01/26
'ÀË¬d¨Ã³]©wcp110¸ê®Æ
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/5 ­û¤u½s¸¹¤w¥i«D¼Æ¦r»Ý°µÂà´«
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Lydia 2021/09/02 §ï¼Ò²Õ
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      Text7 = ""
   Else
      Text7 = "N"
      MsgBox "¥¼¤Ä¿ï¥N²z¤H!", vbInformation, "¥²­nÄæ¦ì¡I"
      Cancel = True
   End If
End Sub
' ¬O§_¸É¥ó
Private Sub textAdd_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Cancel = False
   
   ' µL¸ê®Æ®É¤£°µ¥ô¦óÀË¬d
   If IsEmptyText(textAdd) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textAdd)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textAdd, nIndex)
      Select Case strTemp
         Case "1", "2":
         Case Else
            Cancel = True
            strTit = "ÀË®Ö¸ê®Æ"
            strMsg = "¬O§_¸É¥ó¶µ¥Ø<" & strTemp & ">¤£¥¿½T"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textAdd_GotFocus
            GoTo EXITSUB
      End Select
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textAdd, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textAdd, nCount) Then
               Cancel = True
               strTit = "ÀË®Ö¸ê®Æ"
               strMsg = "¬O§_¸É¥ó¶µ¥Ø<" & strTemp & ">¤£¥i­«ÂÐ"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textAdd_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
End Sub

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
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' µo¤å¤é¤é´Á¤£¥i¶W¹L¨t²Î¤é
      'edit by nick 2004/08/31 ¨t²Î¤é¥[¤@¤Ñ
      'If Val(DBDATE(textCP27)) > Val(DBDATE(SystemDate())) Then
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         'edit by nick 2004/08/31
         'strMsg = "µo¤å¤é¤£¥i¶W¹L¨t²Î¤é"
         strMsg = "µo¤å¤é¤£¥i¶W¹L¨t²Î¤é¥[¤@¤Ñ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' ¨ú±o¶Ê¼f´Á­­ªº¤é´Á
      If Me.textCP27.Tag <> Me.textCP27.Text Then 'Added by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
          textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
      End If
      Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
   End If
EXITSUB:
End Sub

'add by nick 2004/08/13
Private Sub textCP84_GotFocus()
   InverseTextBox textCP84
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
If IsEmptyText(textCP84) = False Then
    If IsNumeric(textCP84) = False Then
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "½Ð¿é¤J¼Æ¦r"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP84_GotFocus
    Else
        textCP84.Text = Trim(Val(textCP84.Text))
    End If
End If
End Sub

Private Sub textDN_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¬O§_¿é¤JD/N
Private Sub textDN_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textDN) = False Then
      Select Case textDN
         Case " ", "Y":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎY"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDN_GotFocus
      End Select
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¦C¦L©w½Z
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' §ó·sÄæ¦ìªº¤º®e
Private Sub OnUpdateField()
   ' µo¤å¤é
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' ¶i«×³Æµù
   '910801 Sieg 602
'edit by nick 2006/01/26
'   If textCP64_2 <> "" Then
'      If textCP64 = "" Then
'         textCP64 = textCP64_2
'      Else
'         textCP64 = textCP64 & "," & textCP64_2
'      End If
'   End If
   SetCPFieldNewData "CP64", textCP64
   
   'add by nickc 2006/01/26
   SetCPFieldNewData "CP110", m_CP110
   'add by nickc 2006/02/10
   SetCPFieldNewData "CP22", Text7
   ' Add By Sindy 98/03/11
   SetCPFieldNewData "CP113", textCP113
   
   'Add By Sindy 2012/12/20
   ' ¬O§_¹q¤l°e¥ó
   SetCPFieldNewData "CP118", textCP118
End Sub

'edit by nickc 2006/01/26
'Private Sub textCP64_2_GotFocus()
'   TextInverse textCP64_2
'End Sub

' §ó·s°Ó¼Ð°ò¥»ÀÉªº¬ÛÃöÄæ¦ì
Private Sub OnUpdateCaseProperty()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
   ' §ó·s®×¥ó¶i«×ÀÉ
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (³æ¤Þ¸¹)
            'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
            strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
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
   ' ³]©wSQL»yªk§ó·sªº±ø¥ó
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' °õ¦æSQL«ü¥O
   If bDifference = True Then: cnnConnection.Execute strSql
   
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strTmp As String
   Dim strSql As String
   Dim strNP08 As String
   Dim strNP07 As String
   Dim strNP22 As String
      
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' §ó·s®×¥ó¶i«×ÀÉ
   OnUpdateCaseProperty
      
   ' ­Y¦³¿é¤J¶Ê¼f´Á­­®É, ·s¼W¤@µ§¶Ê¼fªº°O¿ý¨ì¤U¤@µ{§ÇÀÉ
   If IsEmptyText(textUargeDate) = False Then
      'Add By Sindy 2023/5/5 FCT­«·sµo¤å¡A­Y¤U¤@µ{§Ç¤w¦³¸Ó¦¬¤å¸¹¥¼Äò¿ì¤§¶Ê¼f´Á­­¡A«h§ó·s´Á­­§Y¥i¡A¤£­n¥t·s¼W´Á­­
      strExc(0) = "SELECT NP01,NP22 from NextProgress" & _
                  " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "UPDATE NextProgress SET NP08=" & PUB_GetWorkDay1(textUargeDate, True) & ",NP09=" & DBDATE(textUargeDate) & _
                  " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null"
         cnnConnection.Execute strSql
      Else
      '2023/5/5 END
         strNP07 = "305"
         strNP22 = GetNextProgressNo()
           'Modify By Cheng 2003/09/05
   '      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
   '               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
   '                        DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
         'Modified by Lydia 2020/07/09 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                           DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                           PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
      End If
      
      ' ©µ®i, ¨Ï¥Î«Å»}, ¥Zµn¼s§i, Ãº¦~¶O, ¶Ê¼f, ´£¥Ó, ¦¬¹F¤£¦L±µ¬¢µ²®×³æ
      Select Case strNP07
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
            'Modify By Cheng 2002/01/15
            '¨ú®ø¥~°ÓFCT¦C¦L±µ¬¢µ²®×³æ
'            ' ¦C¦L°ê¤º®×¥ó±µ¬¢¤Îµ²®×°O¿ý³æ
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
      End Select
   End If
   
   'add by nick 2004/08/13 §ó·s¹ê»Úµo¤å³W¶O
   If textCP84.Enabled = True Then
      strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2012/12/20 ­Y¬°¹q¤l°e¥ó«h¦Û°Ê³]©w¬°¤£¸gµo¤å«Ç
   '¥H¨¾°Ê§@¬°­«·sµo¤å, ©Ò¥H¤@¨Ö§âµo¤å«Ç¬ÛÃöÄæ¦ì²MªÅ
   If textCP118.Visible = True And textCP118 = "Y" Then
      strSql = "Update CaseProgress Set CP123=null" & _
                                                          ",CP124=null" & _
                                                          ",CP125=null" & _
                                                          ",CP28=null" & _
                                                          ",CP131=null" & _
                                                          ",CP132=null" & _
                   " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
    'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
    
    'Add By Sindy 2010/7/8 ÀË¬d°Ó«~¸ê®Æ»P°ò¥»ÀÉ°Ó«~Ãþ§O¬O§_¤@­P
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2012/9/26 ÀË¬d¬O§_¬°¤@¥Ó½Ð®Ñ¦h¥ó¨Ã§ó·s¸ê®Æ
   'ÅÜ§ó®×
   'Modify By Sindy 2013/4/9 ©w½Z»y¤å¬O­^¤å®É¤~°µ¤@¥Ó½Ð®Ñ¦h¥ó
   'Modify By Sindy 2014/6/24 mark : ¤£ºÞ©w½Z»y¤å
   'If m_CP10 = "301" And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" Then
   If m_CP10 = "301" Then
   '2013/4/9 End
   '2014/6/24 END
      Call PUB_UpdateCP148(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, textCP27)
   End If
   
    '911107 nick transation
    cnnConnection.CommitTrans
    
     'Add by nickc 2008/02/22 ÀË¬d¥N²z¤HEmail(»Ý¦Ò¼{¥i¯à¬°FF®×¥ó)
    PUB_CheckEMail m_CP44, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
    Exit Function
CheckingErr:
    MsgBox (Err.Description)
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   'Add By Sindy 2012/4/17
   If m_CP10 = "301" Then
      ' ÀË¬dÅÜ§ó¨Æ¶µÀÉ¬O§_¦³¸ê®Æ
      If IsChangeEventExist(m_CP09) = False Or m_blnClkChgButton = False Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "½Ð¿é¤JÅÜ§ó¨Æ¶µ!!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         cmdMod.SetFocus
         GoTo EXITSUB
      End If
   Else
      'Modify By Sindy 2022/2/7
      'T©MFCTªº´îÁY°Ó«~313µo¤å¡G¤£¥²¦AÀË¬d¤@©w­n«öÅÜ§ó¨Æ¶µªº«ö¶s¡C
      If m_CP10 <> "313" Then
      '2022/2/7 END
         If m_blnClkChgButton = False Then
            MsgBox "½Ð¿é¤JÅÜ§ó¨Æ¶µ!!!", vbExclamation + vbOKOnly
            Me.cmdMod.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' µo¤å¤é
   If IsEmptyText(textCP27) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤Jµo¤å¤é"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2011/01/06
   '¥~°Ó(S)¥Ó½Ð¤H1©ÎFC¥N²z¤H¦Ü¤Ö­n¿é¤J¤@­Ó
   '¨ä¥Lªº¤@©w­n¿é¤J¥Ó½Ð¤H1
   If m_TM01 = "S" Then
        If textTM23 = "" And m_TM44 = "" Then
            MsgBox "¥Ó½Ð¤H1©ÎFC¥N²z¤H¦Ü¤Ö­n¿é¤J¤@­Ó!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   Else
        If textTM23 = "" Then
            MsgBox "¥Ó½Ð¤H1¤£¥iªÅ¥Õ!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   End If
   
    'Added by Lydia 2021/09/02 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         GoTo EXITSUB
    End If
    
   CheckDataValid = True
EXITSUB:
End Function


' ¶Ê¼f´Á­­
Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsTaiwanDate(textUargeDate, False) = False Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¶Ê¼f´Á­­¤é´Á¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textUargeDate_GotFocus
      End If
   End If
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textAdd_GotFocus()
   InverseTextBox textAdd
End Sub

Private Sub textDN_GotFocus()
   InverseTextBox textDN
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

'add by nick 2004/08/13 µo¤å³W¶O¡A¥Ó½Ð°ê®a¥xÆW¤~ÀË¬d
If Me.textCP84.Enabled = True Then
   Cancel = False
   textCP84_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textCP84.Enabled = True And m_TM10 = "000" Then
    If Val(textCP84.Text) <> Val(m_CP84) Then
        MsgBox "µo¤å³W¶O[" & Trim(Val(m_CP84)) & "] »P¹ê»Úµo¤å³W¶O[" & Trim(Val(textCP84.Text)) & "]¤£¦P", , "Äµ§i¡I"
        textCP84_GotFocus
        Exit Function
    End If
End If

If Me.textAdd.Enabled = True Then
   Cancel = False
   textAdd_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP27.Enabled = True Then
   Cancel = False
   textCP27_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 98/03/11
If Me.textCP113.Enabled = True Then
   Cancel = False
   textCP113_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'98/03/11 End

If Me.textDN.Enabled = True Then
   Cancel = False
   textDN_Validate Cancel
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

If Me.textUargeDate.Enabled = True Then
   Cancel = False
   textUargeDate_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'add by nickc 2006/01/27
'edit by nickc 2006/02/07
If m_TM01 = "FCT" Then
    If Me.lstNameAgent.Enabled = True Then
        Cancel = False
        lstNameAgent_Validate Cancel
        If Cancel = True Then
            lstNameAgent.SetFocus
            Exit Function
        End If
    End If
End If
TxtValidate = True
End Function

'Add By Sindy 98/03/11
Private Sub textCP113_GotFocus()
   TextInverse textCP113
End Sub

Private Sub textCP113_Validate(Cancel As Boolean)
   If textCP113 <> "" Then
      If Not IsNumeric(textCP113) Then
         MsgBox "½Ð¿é¤J¼Æ¦r¡I", vbExclamation
         textCP113.SetFocus
         textCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   If GetPrjNation1(textTMKey) = "000" Then
      '2009/4/2 modify by sonia 302§ó¥¿¥B¶i«×³Æµù¬° §ó§ïµù¥UÃÒ ©Î §ó§ï®Ö­ã¨ç ®É¤£ÀË¬d
      'Cancel = Not PUB_CheckCP113(textCP113, m_TM01, m_CP10, m_CP14)
      '2012/4/12 modify by sonia 302§ó¥¿³£¤£ÀË¬d-³¯ª÷½¬
      'If textCP10 = "§ó¥¿" And (InStr(1, textCP64, "§ó§ïµù¥UÃÒ", 1) > 0 Or InStr(1, textCP64, "§ó§ï®Ö­ã¨ç", 1) > 0) Then
      If textCP10 = "§ó¥¿" Then
      Else
         Cancel = Not PUB_CheckCP113(textCP113, m_TM01, m_CP10, m_CP14)
      End If
   End If
End Sub
'98/03/11 End

'Add By Sindy 2010/4/14
Private Sub PrintLetter()
   Dim ET03 As String, ET03_1 As String, stContent As String
   Dim strFilePath As String, strFN01 As String, strFN02 As String  'Added by Lydia 2023/05/03
   
   m_strLanguage = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) 'Add By Sindy 2012/10/12
   
   'Add By Sindy 2012/11/23 ±q¤U­±µ{¦¡©¹¤WMove¦Ü¦¹
   bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, m_CP10 = "102", , bolPlusPaper) 'ÀË¬d¬O§_¥HE-Mail³qª¾
   '2012/11/23 End
   
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   InsExpField
   
   ' ®×¥ó©Ê½è
   Select Case m_CP10
      'Add By Sindy 2012/10/12
      'ÅÜ§ó
      Case "301":
         ' ©w½Z»y¤å
         Select Case m_strLanguage
            ' ¤é¤å
            Case "3":
               If Trim(textTM15.Text) = "" Then 'µù¥U«eÅÜ§ó
                  ET03 = "01"
                  ET03_1 = "02" 'Ä¶¤å
               Else 'µù¥U«áÅÜ§ó
                  ET03 = "03"
                  'Add By Sindy 2019/3/26
                  If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                     ET03_1 = "05"
                  Else
                  '2019/3/26 END
                     ET03_1 = "04" 'Ä¶¤å
                  End If
               End If
         End Select
      '2012/10/12 End
      ' §ó¥¿
      Case "302":
         ' ©w½Z»y¤å
         Select Case m_strLanguage
            ' ¤é¤å
            Case "3":
               If InStr(1, textCP64, "§ó§ïµù¥UÃÒ", 1) > 0 Then
                  'Modify By Sindy 2012/10/12
                  'NowPrint m_CP09, "01", "01", False, strUserNum, 0
                  ET03 = "01"
                  '2012/10/12 End
               'Add By Sindy 2019/10/8 §ó¥¿ªº¬ÛÃöÁ`¦¬¤å¸¹¬°©µ®i®Ö­ã
               ElseIf bolCase102_1001 = True Then
                  ET03 = "02"
               End If
         End Select
      Case Else:
   End Select
   
   'Add By Sindy 2012/10/12
   If ET03 <> "" Then
'      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, m_CP10 = "102", , bolPlusPaper)
      'If bolEmail Then 'Mark by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
         '§PÂ_¬O§_EMail¦P®É±H¯È¥»
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Added by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW: ­^¤å²Õ¤À¦¨«H¨ç©MÂ½Ä¶¨â­ÓÀÉ®×
         If m_strLanguage <> "3" Then
            strFilePath = Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04)
            If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, , strFN01, strFN02) = False Then
              Exit Sub
            End If
            NowPrint m_CP09, "01", ET03, True, strUserNum, , , , , iCopy, , True
            If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                Sleep 100
            End If
            If ET03_1 <> "" Then
               NowPrint m_CP09, "01", ET03_1, True, strUserNum, , , , , iCopy, , True
               If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                   Sleep 100
               End If
            End If
         Else '¤é¤å²Õ:¤£§ïÅÜ¦sÀÉ¼Ò¦¡
         'end 2023/05/03
            'Added by Lydia 2024/11/14 ¦]¤é¥»¥N²z¤H¯S§O­n¨D¡A»Ý±N³qª¾«H¨ç»PÄ¶¤åµ¥¤À¶}¡A¨Ã¥B²Î¤@¦WºÙ¦p¤U(¼Ò²Õ¨ú±o)¡F­ì¥»ªºÀÉ®×(®×¸¹_¤é´Á=³qª¾¨ç+Ä¶¤å)¤´­n²£¥Í¡A¥H§K¤é«á¤S¦³¥N²z¤H­n¨D¦X¨Ö
            strFilePath = Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04)
            If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, , strFN01, strFN02) = False Then
              Exit Sub
            End If
            NowPrint m_CP09, "01", ET03, True, strUserNum, , , , , iCopy, , True
            If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                Sleep 100
            End If
            If ET03_1 <> "" Then
               NowPrint m_CP09, "01", ET03_1, True, strUserNum, , , , , iCopy, , True
               If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                   Sleep 100
               End If
            End If
            'end 2024/11/14
            If ET03_1 <> "" Then
               NowPrint m_CP09, "01", ET03, False, strUserNum, , , , , iCopy
               NowPrint m_CP09, "01", ET03_1, False, strUserNum, , , , , iCopy
               NowPrint m_CP09, "01", ET03, False, strUserNum, , , True, stContent, , , , True
               NowPrint m_CP09, "01", ET03_1, False, strUserNum, , stContent, , , , , True, True
            Else
               NowPrint m_CP09, "01", ET03, False, strUserNum, , , , , iCopy, , True, True
            End If
         End If 'Added by Lydia 2023/05/03
         MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_TM01) & " ]¡I"
      'Mark by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
      'Else
      '   NowPrint m_CP09, "01", ET03, False, strUserNum, 0
      '   If ET03_1 <> "" Then
      '      NowPrint m_CP09, "01", ET03_1, False, strUserNum, 0
      '   End If
      'End If
      'end 2023/05/03
   End If
End Sub

'Add By Sindy 2010/4/14
' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
Dim strTemp09 As String, strTemp38 As String, strTemp As String
Dim intCnt As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ET03_1 As String
   
   ' ®×¥ó©Ê½è
   Select Case m_CP10
      'Add By Sindy 2012/10/12
      'ÅÜ§ó
      Case "301":
         ' ©w½Z»y¤å
         Select Case m_strLanguage
            ' ¤é¤å
            Case "3":
               'ÀË¬dÅÜ§ó¨Æ¶µ
               strSql = "select * from changeevent where ce01='" & m_CP09 & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If "" & RsTemp.Fields("ce04") <> "" Or _
                     "" & RsTemp.Fields("ce05") <> "" Or _
                     "" & RsTemp.Fields("ce06") <> "" Or _
                     "" & RsTemp.Fields("ce07") <> "" Or _
                     "" & RsTemp.Fields("ce08") <> "" Then
                     strTemp09 = "Y"
                  End If
                  If "" & RsTemp.Fields("ce23") <> "" Or _
                     "" & RsTemp.Fields("ce24") <> "" Or _
                     "" & RsTemp.Fields("ce25") <> "" Or _
                     "" & RsTemp.Fields("ce26") <> "" Or _
                     "" & RsTemp.Fields("ce27") <> "" Or _
                     "" & RsTemp.Fields("ce28") <> "" Or _
                     "" & RsTemp.Fields("ce29") <> "" Or _
                     "" & RsTemp.Fields("ce30") <> "" Or _
                     "" & RsTemp.Fields("ce31") <> "" Or _
                     "" & RsTemp.Fields("ce32") <> "" Or _
                     "" & RsTemp.Fields("ce33") <> "" Or _
                     "" & RsTemp.Fields("ce34") <> "" Or _
                     "" & RsTemp.Fields("ce35") <> "" Or _
                     "" & RsTemp.Fields("ce36") <> "" Or _
                     "" & RsTemp.Fields("ce37") <> "" Then
                     strTemp38 = "Y"
                  End If
                  If strTemp09 = "Y" And strTemp38 = "Y" Then
                     'Modified by Morgan 2024/4/2
                     'strTemp = "¡]¥XÄ@¤Hªí¥Ü¤ÎÇZ¦í©ÒŒi§ó¡^"
                     strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó°Ó¼Ð¥Ó½Ð¤H¤Î¦a§}")
                  ElseIf strTemp09 = "Y" Then
                     'Modified by Morgan 2024/4/2
                     'strTemp = "¡]¥XÄ@¤Hªí¥ÜŒi§ó¡^"
                     strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó°Ó¼Ð¥Ó½Ð¤H")
                  ElseIf strTemp38 = "Y" Then
                     'Modified by Morgan 2024/4/2
                     'strTemp = "¡]¥XÄ@¤H¦í©ÒŒi§ó¡^"
                     strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó°Ó¼Ð¥Ó½Ð¦a§}")
                  End If
               End If
               If Trim(textTM15.Text) = "" Then 'µù¥U«eÅÜ§ó
                  EndLetter "01", m_CP09, "01", strUserNum
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('01','" & m_CP09 & "','01','" & strUserNum & _
                           "','ÅÜ§ó¨Æ¶µ','" & strTemp & "')"
                  cnnConnection.Execute strSql
                  'Ä¶¤å
                  EndLetter "01", m_CP09, "02", strUserNum
               Else 'µù¥U«áÅÜ§ó
                  EndLetter "01", m_CP09, "03", strUserNum
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('01','" & m_CP09 & "','03','" & strUserNum & _
                           "','ÅÜ§ó¨Æ¶µ','" & strTemp & "')"
                  cnnConnection.Execute strSql
                  
                  'Add By Sindy 2019/3/26
                  If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                     'Ä¶¤å
                     ET03_1 = "05"
                     EndLetter "01", m_CP09, "05", strUserNum
                     '¤@¤å¦h®×²M³æšd
                     strTemp = PUB_GetFCTAppendix_JP(m_TM01, m_TM02, m_TM03, m_TM04, "301", m_CP27, "01", m_CP28, m_CP09, "05", intCnt)
                     ' ¤@®×¦h¥ó¥ó¼Æ
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & _
                              "','¤@®×¦h¥ó¥ó¼Æ','" & intCnt & "')"
                     cnnConnection.Execute strSql
                  Else
                     'Ä¶¤å
                     ET03_1 = "04"
                     EndLetter "01", m_CP09, "04", strUserNum
                  End If
                  'Åª¨úÅÜ§óÀÉ
                  StrSQLa = "Select ce01,ce04,ce05,ce06,ce07,ce08,ce09" & _
                            ",ce23,ce24,ce25,ce26,ce27,ce28,ce29,ce30,ce31,ce32,ce33,ce34,ce35,ce36,ce37,ce38" & _
                            ",ce10,ce11,ce12,ce13,ce14,ce15,ce16" & _
                            ",ce68,ce69,ce70,ce71,ce72,ce73,ce74,ce75,ce76,ce77,ce78,ce79,ce80,ce81,ce82,ce83,ce84,ce85,ce86,ce87,ce88,ce89,ce90,ce91" & _
                            ",ce55,ce56 From changeevent Where ce01='" & m_CP09 & "'"
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.Fields(0).Value > 0 Then
                     If "" & rsA.Fields("ce04") <> "" Or _
                        "" & rsA.Fields("ce05") <> "" Or _
                        "" & rsA.Fields("ce06") <> "" Or _
                        "" & rsA.Fields("ce07") <> "" Or _
                        "" & rsA.Fields("ce08") <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & ET03_1 & "','" & strUserNum & _
                                 "','ÅÜ§ó¥Ó½Ð¤H¦WºÙ','¡ð')"
                        cnnConnection.Execute strSql
                     End If
                     If "" & rsA.Fields("ce23") <> "" Or _
                        "" & rsA.Fields("ce24") <> "" Or _
                        "" & rsA.Fields("ce25") <> "" Or _
                        "" & rsA.Fields("ce26") <> "" Or _
                        "" & rsA.Fields("ce27") <> "" Or _
                        "" & rsA.Fields("ce28") <> "" Or _
                        "" & rsA.Fields("ce29") <> "" Or _
                        "" & rsA.Fields("ce30") <> "" Or _
                        "" & rsA.Fields("ce31") <> "" Or _
                        "" & rsA.Fields("ce32") <> "" Or _
                        "" & rsA.Fields("ce33") <> "" Or _
                        "" & rsA.Fields("ce34") <> "" Or _
                        "" & rsA.Fields("ce35") <> "" Or _
                        "" & rsA.Fields("ce36") <> "" Or _
                        "" & rsA.Fields("ce37") <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & ET03_1 & "','" & strUserNum & _
                                 "','ÅÜ§ó¥Ó½Ð¤H¦í©Ò','¡ð')"
                        cnnConnection.Execute strSql
                     End If
                     If "" & rsA.Fields("ce10") <> "" Or "" & rsA.Fields("ce11") <> "" Or "" & rsA.Fields("ce12") <> "" Or _
                        "" & rsA.Fields("ce13") <> "" Or "" & rsA.Fields("ce14") <> "" Or "" & rsA.Fields("ce15") <> "" Or _
                        "" & rsA.Fields("ce68") <> "" Or "" & rsA.Fields("ce69") <> "" Or "" & rsA.Fields("ce70") <> "" Or _
                        "" & rsA.Fields("ce71") <> "" Or "" & rsA.Fields("ce72") <> "" Or "" & rsA.Fields("ce73") <> "" Or _
                        "" & rsA.Fields("ce74") <> "" Or "" & rsA.Fields("ce75") <> "" Or "" & rsA.Fields("ce76") <> "" Or _
                        "" & rsA.Fields("ce77") <> "" Or "" & rsA.Fields("ce78") <> "" Or "" & rsA.Fields("ce79") <> "" Or _
                        "" & rsA.Fields("ce80") <> "" Or "" & rsA.Fields("ce81") <> "" Or "" & rsA.Fields("ce82") <> "" Or _
                        "" & rsA.Fields("ce83") <> "" Or "" & rsA.Fields("ce84") <> "" Or "" & rsA.Fields("ce85") <> "" Or _
                        "" & rsA.Fields("ce86") <> "" Or "" & rsA.Fields("ce87") <> "" Or "" & rsA.Fields("ce88") <> "" Or _
                        "" & rsA.Fields("ce89") <> "" Or "" & rsA.Fields("ce90") <> "" Or "" & rsA.Fields("ce91") <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & ET03_1 & "','" & strUserNum & _
                                 "','ÅÜ§ó¥Nªí¤H','¡ð')"
                        cnnConnection.Execute strSql
                     End If
                     If "" & rsA.Fields("ce55") <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & ET03_1 & "','" & strUserNum & _
                                 "','ÅÜ§ó¥X¦W¥N²z¤H','¡ð')"
                        cnnConnection.Execute strSql
                     End If
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  '2019/3/26 END
               End If
         End Select
      '2012/10/12 End
      ' §ó¥¿
      Case "302":
         ' ©w½Z»y¤å
         Select Case m_strLanguage
            ' ¤é¤å
            Case "3":
               If InStr(1, textCP64, "§ó§ïµù¥UÃÒ", 1) > 0 Then
                  EndLetter "01", m_CP09, "01", strUserNum
               'Add By Sindy 2019/10/8 §ó¥¿ªº¬ÛÃöÁ`¦¬¤å¸¹¬°©µ®i®Ö­ã
               ElseIf bolCase102_1001 = True Then
                  EndLetter "01", m_CP09, "02", strUserNum
               End If
         End Select
      Case Else:
   End Select
End Sub

'Add By Sindy 2012/12/20
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2012/12/20
Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
