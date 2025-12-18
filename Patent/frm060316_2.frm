VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060316_2 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "®Ö­ã¨ç/ÃÒ®Ñ¨ç"
   ClientHeight    =   5592
   ClientLeft      =   0
   ClientTop       =   948
   ClientWidth     =   8988
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5592
   ScaleWidth      =   8988
   Begin VB.CommandButton cmdPrev 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   400
      Left            =   7860
      TabIndex        =   2
      Top             =   36
      Width           =   1092
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6852
      TabIndex        =   3
      Top             =   36
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5844
      TabIndex        =   1
      Top             =   36
      Width           =   972
   End
   Begin VB.TextBox textPA04 
      BackColor       =   &H00FFFFFF&
      Height          =   264
      Left            =   2568
      MaxLength       =   2
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   372
   End
   Begin VB.TextBox textPA03 
      BackColor       =   &H00FFFFFF&
      Height          =   264
      Left            =   2328
      MaxLength       =   1
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   252
   End
   Begin VB.TextBox textPA02 
      BackColor       =   &H00FFFFFF&
      Height          =   264
      Left            =   1608
      MaxLength       =   6
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   732
   End
   Begin VB.TextBox textPA01 
      BackColor       =   &H00FFFFFF&
      Height          =   264
      Left            =   1128
      MaxLength       =   3
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   492
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4830
      Left            =   45
      TabIndex        =   4
      Top             =   720
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   8530
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "°ò¥»¸ê®Æ"
      TabPicture(0)   =   "frm060316_2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label12"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label7(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label28"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textPA101_2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textPA05"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textPA102"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textPA56"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textPA54"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textPA53"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textPA51"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textPA07"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textPA139"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textPA103"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textPA101"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textPA55"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textPA52"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textPA06"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textPA48"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textPA89"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textPA104"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "µo©ú¤H"
      TabPicture(1)   =   "frm060316_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label17(0)"
      Tab(1).Control(1)=   "textPA60_3"
      Tab(1).Control(2)=   "textPA60_1"
      Tab(1).Control(3)=   "textPA60_2"
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(5)=   "cmbPA60"
      Tab(1).Control(6)=   "cmdAddRow"
      Tab(1).Control(7)=   "cmdDelRow"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdDelRow 
         Caption         =   "§R°£"
         Height          =   285
         Left            =   -73110
         TabIndex        =   51
         Top             =   810
         Width           =   735
      End
      Begin VB.CommandButton cmdAddRow 
         Caption         =   "¥[¤J"
         Height          =   285
         Left            =   -73935
         TabIndex        =   50
         Top             =   810
         Width           =   735
      End
      Begin VB.TextBox textPA104 
         Height          =   264
         Left            =   1890
         MaxLength       =   140
         TabIndex        =   25
         Top             =   4140
         Width           =   6855
      End
      Begin VB.TextBox textPA89 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   264
         Left            =   6723
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1692
      End
      Begin VB.TextBox textPA48 
         Height          =   264
         Left            =   1890
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox textPA06 
         Height          =   264
         Left            =   1890
         MaxLength       =   250
         TabIndex        =   10
         Top             =   615
         Width           =   6855
      End
      Begin VB.TextBox textPA52 
         Height          =   264
         Left            =   1890
         TabIndex        =   16
         Top             =   1680
         Width           =   6855
      End
      Begin VB.TextBox textPA55 
         Height          =   264
         Left            =   1890
         TabIndex        =   19
         Top             =   2505
         Width           =   6855
      End
      Begin VB.TextBox textPA101 
         Height          =   264
         Left            =   1890
         MaxLength       =   9
         TabIndex        =   22
         Top             =   3315
         Width           =   1695
      End
      Begin VB.TextBox textPA103 
         Height          =   264
         Left            =   1890
         MaxLength       =   140
         TabIndex        =   24
         Top             =   3870
         Width           =   6855
      End
      Begin VB.ComboBox cmbPA60 
         Height          =   300
         Left            =   -73920
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   0
         Top             =   408
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   3525
         Left            =   -74910
         TabIndex        =   52
         Top             =   1140
         Width           =   8685
         _ExtentX        =   15325
         _ExtentY        =   6223
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "V|µo©ú¤H½s¸¹|¤¤¤å¦WºÙ|­^¤å¦WºÙ|¤é¤å¦WºÙ"
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
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.TextBox textPA139 
         Height          =   285
         Left            =   1890
         TabIndex        =   21
         Top             =   3030
         Width           =   6855
         VariousPropertyBits=   671105051
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   315
         Left            =   1890
         TabIndex        =   26
         Top             =   4410
         Width           =   6855
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "12091;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA07 
         Height          =   285
         Left            =   1890
         TabIndex        =   12
         Top             =   870
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA51 
         Height          =   285
         Left            =   1890
         TabIndex        =   15
         Top             =   1410
         Width           =   6855
         VariousPropertyBits=   671105051
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA53 
         Height          =   285
         Left            =   1890
         TabIndex        =   17
         Top             =   1935
         Width           =   6855
         VariousPropertyBits=   671105051
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA54 
         Height          =   285
         Left            =   1890
         TabIndex        =   18
         Top             =   2220
         Width           =   6855
         VariousPropertyBits=   671105051
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA56 
         Height          =   285
         Left            =   1890
         TabIndex        =   20
         Top             =   2760
         Width           =   6855
         VariousPropertyBits=   671105051
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA102 
         Height          =   285
         Left            =   1890
         TabIndex        =   23
         Top             =   3585
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA05 
         Height          =   285
         Left            =   1890
         TabIndex        =   9
         Top             =   345
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA101_2 
         Height          =   285
         Left            =   3600
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3300
         Width           =   5130
         VariousPropertyBits=   671105055
         Size            =   "9049;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA60_2 
         Height          =   285
         Left            =   -72330
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   2925
         VariousPropertyBits=   671105055
         Size            =   "5159;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA60_1 
         Height          =   285
         Left            =   -72330
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   405
         Width           =   2925
         VariousPropertyBits=   671105055
         Size            =   "5159;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA60_3 
         Height          =   285
         Left            =   -69360
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   405
         Width           =   2925
         VariousPropertyBits=   671105055
         Size            =   "5159;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H³¡ªù (¤é)¡G"
         Height          =   180
         Left            =   120
         TabIndex        =   49
         Top             =   3060
         Width           =   1425
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "¦C¦L³Æµù¡G"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   4410
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "®×¥ó¦WºÙ(¤¤)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   345
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "«áÄò­ã»éÂ²³æ³ø§i¡G"
         Height          =   252
         Index           =   2
         Left            =   4908
         TabIndex        =   45
         Top             =   1176
         Width           =   1750
      End
      Begin VB.Label Label7 
         Caption         =   "«È¤á®×¥ó®×¸¹¡G"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "®×¥ó¦WºÙ(­^)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "®×¥ó¦WºÙ(¥~)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Ápµ¸¤H1 (¤¤)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1410
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Ápµ¸¤H1 (­^)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Ápµ¸¤H1 (¤é)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1935
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Ápµ¸¤H2 (¤¤)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Ápµ¸¤H2 (­^)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Ápµ¸¤H2 (¤é)¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2790
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "¹êÅé°Æ¥»¦¬¨ü¤H¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3315
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "¹êÅé°Æ¥»³sµ¸¤H¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3585
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "¹êÅé°Æ¥»©¼©Ò®×¸¹1¡G"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3870
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "¹êÅé°Æ¥»©¼©Ò®×¸¹2¡G"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   4140
         Width           =   1815
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "µo©ú¤H¡G"
         Height          =   180
         Index           =   0
         Left            =   -74730
         TabIndex        =   27
         Top             =   450
         Width           =   720
      End
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "µLµo©ú¤H¸ê®Æ"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4095
      TabIndex        =   48
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "¥»©Ò®×¸¹¡G"
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   11
      Top             =   390
      Width           =   975
   End
End
Attribute VB_Name = "frm060316_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/7/4 ¤é¤å¤w§ï§ìTable
'Memo By Sindy 2022/3/2 Form2.0¤w­×§ï
'Memo By Morgan 2012/12/10 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo by Morgan2010/12/27 ¥Ó½Ð®×¸¹Äæ¤w­×§ï
'2010/12/6 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan2010/8/16 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Dim m_PA01 As String
Dim m_PA02 As String
Dim m_PA03 As String
Dim m_PA04 As String
Dim m_PA22 As String 'Add by Morgan 2005/6/23
Dim m_CP09 As String
Dim m_CP43 As String
Dim m_PrevForm As String
' «Å§iµo©ú¤H
Private Type INVENTOR
   iN01 As String
   iN02 As String
   iN04 As String
   IN05 As String
   IN06 As String
End Type

Dim m_InventorList() As INVENTOR
Dim m_InventorListCount As Integer

Dim m_LetterLanguage As String
Dim m_LetterKind As Integer
'Add By Cheng 2003/01/02
Dim m_blnPriData As Boolean '¬O§_¦³Àu¥ýÅv¸ê®Æ
Dim m_bln3PriData As Boolean '¬O§_¦³¤T­Ó¥H¤WÀu¥ýÅv¸ê®Æ
'Add by Morgan 2004/7/20
Dim m_PA08 As String
Dim m_PA14 As String
'Add by Morgan 2004/7/27
Dim m_CP05 As String '¨Ó¨ç¦¬¤å¤é
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean, m_iCopy As Integer
Dim strSpecNO As Boolean   'add by sonia 2014/4/28
'Added by Morgan 2014/6/3
Dim m_bolDNEmail As Boolean, m_bolDNPlusPaper As Boolean
Dim pPrevRow As Integer 'Add By Sindy 2014/11/6
Dim m_PA178 As String 'Added by Morgan 2023/2/18

Private Sub cmbPA60_Click()
   Dim strIN01 As String
   Dim strIN02 As String
   Dim nPos As Integer
   
   strIN01 = Mid(cmbPA60.List(cmbPA60.ListIndex), 1, 8)
   strIN02 = Mid(cmbPA60.List(cmbPA60.ListIndex), 9, 2)
   For nPos = 0 To m_InventorListCount - 1
      If strIN01 = m_InventorList(nPos).iN01 And strIN02 = m_InventorList(nPos).iN02 Then
         textPA60_1.Text = m_InventorList(nPos).iN04
         textPA60_2.Text = m_InventorList(nPos).IN05
         textPA60_3.Text = m_InventorList(nPos).IN06
         Exit For
      End If
   Next nPos
End Sub

'Add By Sindy 2014/11/10
Private Sub cmdAddRow_Click()
Dim bolChk As Boolean
Dim ii As Integer
   
   'ÀË¬dµo©ú¤H
   bolChk = True
   strExc(1) = Trim(cmbPA60.Text)
   If strExc(1) = "" Then Exit Sub
   For ii = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(ii, 1) = strExc(1) Then
         bolChk = False
         Exit For
      End If
   Next ii
   If Not bolChk Then
      MsgBox "µo©ú¤H¤£¥i­«ÂÐ !", vbCritical
      cmbPA60.SetFocus
      Exit Sub
   End If
   If Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) <> "" Then
      GRD1.AddItem ""
   End If
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) = strExc(1)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 2) = textPA60_1
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 3) = textPA60_2
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 4) = textPA60_3
   cmdAddRow.Tag = "I" '°O¿ý¦³²§°Ê¸ê®Æ
   '²MªÅÄæ¦ì
   cmbPA60.ListIndex = 0
   textPA60_1 = ""
   textPA60_2 = ""
   textPA60_3 = ""
End Sub

'Add By Sindy 2014/11/10
Private Sub cmdDelRow_Click()
   If pPrevRow <= 0 Then Exit Sub
   GRD1.col = 0
   GRD1.row = pPrevRow
   If GRD1.CellBackColor <> &HFFC0C0 Then Exit Sub
   If pPrevRow = 1 And GRD1.Rows = 2 Then
      GRD1.TextMatrix(pPrevRow, 0) = ""
      GRD1.TextMatrix(pPrevRow, 1) = ""
      GRD1.TextMatrix(pPrevRow, 2) = ""
      GRD1.TextMatrix(pPrevRow, 3) = ""
      GRD1.TextMatrix(pPrevRow, 4) = ""
   Else
      If pPrevRow > 0 Then
         Call GRD1.RemoveItem(pPrevRow)
      Else
         Exit Sub
      End If
   End If
   pPrevRow = pPrevRow - 1
   cmdDelRow.Tag = "D" '°O¿ý¦³²§°Ê¸ê®Æ
   '²MªÅÄæ¦ì
   cmbPA60.ListIndex = 0
   textPA60_1 = ""
   textPA60_2 = ""
   textPA60_3 = ""
End Sub

Private Sub cmdExit_Click()
   Unload Me
   Select Case m_PrevForm
'Removed by Morgan 2022/11/23 2015¤w¨ú®ø
'      Case "frm060316_1":
'         Unload frm060316_1
'end 2022/11/23
      Case "frm060317_1":
         Unload frm060317_1
   End Select
End Sub

Private Sub CmdPrev_Click()
   Select Case m_PrevForm
'Removed by Morgan 2022/11/23 2015¤w¨ú®ø
'      Case "frm060316_1":
'         frm060316_1.Show
'end 2022/11/23
      Case "frm060317_1":
         frm060317_1.Show
   End Select
   Unload Me
End Sub

Private Sub Combo1_Click()
   'Modify by Morgan 2006/8/7 ¥[¤é¤å³Æµù
   Select Case Me.Combo1.Text
      Case "¿ù»~°h¦^"
         If m_LetterLanguage = "3" Then
            'Modified by Morgan 2017/6/29
            'Me.Combo1.Text = "¥»¥óµý®ÑÇU°O¸ü¨Æ¶µÇRþ÷þàÇeþêþù¡B’½°OÇyúú¨£þêÇeþêþòÇUþú¡B¤@¥¹´¼¼z°]²£§½ÇRªðÁÙþê¡B­q¥¿ÇU¤â“dþàÇy¤â°tþèþîþù³»þàÇeþì¡C­q¥¿ÇU¤â“dþàþß§¹¤F¦¸²Ä¡B§ïÇhþù°e¥I¥ÓþêÆèþåÇeþì¡C"
            'Modified by Morgan 2022/7/4 §ï§ìTable
            'Me.Combo1.Text = "¥»¥óµý®ÑÇU°O¸ü¨Æ¶µÇRþ÷þàÇeþêþù¡B’½°OÇyúú¨£þêÇeþêþòÇUþú¡B¤@¥¹þðÇU­ì¥»Çy´¼¼z°]²£§½ÇRªðÁÙþê¡B­q¥¿¤â“dþàÇy¤â°tþèþîþù³»þàÇeþì¡C­q¥¿«áÇUµý®ÑÇy¨ü¨úÇqÇeþêþòÇp¡B§ïÇhþùþÝ°eÇq­PþêÇeþì¡C"
            'Added by Morgan 2023/2/18 +§PÂ_¹q¤lÃÒ®Ñ±a¤£¦P¤º®e
            If m_PA178 = "1" Then
               Me.Combo1.Text = PUB_GetUniText(Me.Name, "¿ù»~°h¦^(¹q¤lÃÒ®Ñ)")
            Else
            'end 2023/2/18
               Me.Combo1.Text = PUB_GetUniText(Me.Name, "¿ù»~°h¦^")
            End If
            'end 2022/7/4
            'end 2017/6/29
         Else
            'Modify by Morgan 2007/1/31 -- David
            'Me.Combo1.Text = "As some errors have occurred in the Letters Patent, we have returned the same to the Patent Office for correction. We shall send the Letters Patent to you upon receipt."
            'Modify by Morgan 2007/1/31 -- David
            'Me.Combo1.Text = "As some errors have occurred in the Patent Certificate, we have returned the same to the Patent Office for correction. We shall send the Patent Certificate to you upon receipt."
            'Modified by Morgan 2023/2/18
            'Me.Combo1.Text = "As some errors have occurred in the Patent Certificate, we will return the same to the Patent Office for correction. We shall send the Patent Certificate to you upon receipt."
            Me.Combo1.Text = "As some errors have occurred in the Patent Certificate, we will request for correction with the Patent Office. We shall send the corrected Patent Certificate to you upon receipt."
         End If
      Case "Åý»P"
         If m_LetterLanguage = "3" Then
            'Modified by Morgan 2022/7/4 §ï§ìTable
            'Me.Combo1.Text = "’ð´ç¤â“dþàÇR¥²­nÇOÇQÇrþòÇh¡B¥»¥óµý®ÑÇU­ì¥»ÇV¨úÇqÆèÆîþí¹ú©ÒÇRþùþÝ¹wþÞÇq­PþêÇeþì¡C"
            Me.Combo1.Text = PUB_GetUniText(Me.Name, "Åý»P")
            'end 2022/7/4
         Else
            Me.Combo1.Text = "We will keep the original Letters Patent for Assignment application."
         End If
      Case "­×¥¿"
         If m_LetterLanguage = "3" Then
            'Modified by Morgan 2022/7/4 §ï§ìTable
            'Me.Combo1.Text = "­q¥¿¤â“dþàÇR¥²­nÇOÇQÇrþòÇh¡B¥»¥óµý®ÑÇU­ì¥»ÇV¨úÇqÆèÆîþí¹ú©ÒÇRþùþÝ¹wþÞÇq­PþêÇeþì¡C"
            Me.Combo1.Text = PUB_GetUniText(Me.Name, "­×¥¿")
            'end 2022/7/4
         Else
            Me.Combo1.Text = "We will keep the original Letters Patent for Amendment application."
         End If
      Case "¦X¨Ö"
         If m_LetterLanguage = "3" Then
            'Modified by Morgan 2022/7/4 §ï§ìTable
            'Me.Combo1.Text = "¦X¨Ö¨Æ¶µÇUµn“÷ÇR¥²­nÇOÇQÇrþòÇh¡B¥»¥óµý®ÑÇU­ì¥»ÇV¨úÇqÆèÆîþí¹ú©ÒÇRþùþÝ¹wþÞÇq­PþêÇeþì¡C"
            Me.Combo1.Text = PUB_GetUniText(Me.Name, "¦X¨Ö")
            'end 2022/7/4
         Else
            Me.Combo1.Text = "We will keep the original Letters Patent for Merger application."
         End If
      Case "±ÂÅv"
         If m_LetterLanguage = "3" Then
            'Modified by Morgan 2022/7/4 §ï§ìTable
            'Me.Combo1.Text = "±Â“¸¨Æ¶µÇUµn“÷ÇR¥²­nÇOÇQÇrþòÇh¡B¥»¥óµý®ÑÇU­ì¥»ÇV¨úÇqÆèÆîþí¹ú©ÒÇRþùþÝ¹wþÞÇq­PþêÇeþì¡C"
            Me.Combo1.Text = PUB_GetUniText(Me.Name, "±ÂÅv")
            'end 2022/7/4
         Else
            Me.Combo1.Text = "We will keep the original Letters Patent for License application."
         End If
      Case "ÅÜ§ó"
         If m_LetterLanguage = "3" Then
            'Modified by Morgan 2022/7/4 §ï§ìTable
            'Me.Combo1.Text = "“Ä§ó¤â“dþàÇR¥²­nÇOÇQÇrþòÇh¡B¥»¥óµý®ÑÇU­ì¥»ÇV¨úÇqÆèÆîþí¹ú©ÒÇRþùþÝ¹wþÞÇq­PþêÇeþì¡C"
            Me.Combo1.Text = PUB_GetUniText(Me.Name, "ÅÜ§ó")
            'end 2022/7/4
         Else
            Me.Combo1.Text = "We will keep the original Letters Patent for Change application."
         End If
      'Modified by Morgan 2017/6/29
      'Case "§¹¦¨±HÃÒ®Ñ"
      '   If m_LetterLanguage = "3" Then
      '      Me.Combo1.Text = "¥ý¯ëÇU¹ú©Ò³ø§iÇR¤Þþà“dþàÇeþêþù¡B¥»¥óµý®ÑÇU­ì¥»Çy¦P«ÊþÝ°eÇq­PþêÇeþìÇUþú¡Bþç¬d’Ú¤UþèÆê¡C"
      Case "¿ù»~°h¦^«áµý®Ñ§ó§ï§¹¦¨"
         If m_LetterLanguage = "3" Then
            'Modified by Morgan 2022/7/4 §ï§ìTable
            'Me.Combo1.Text = "¥ý¯ëÇU¹ú©Ò³ø§iÇR¤Þþà“dþàÇeþêþù¡B´¼¼z°]²£§½ÇoÇq­q¥¿«áÇUµý®ÑÇy¨ü¨úÇqÇeþêþòÇUþú¡B­ì¥»ÇyþÝ°eÇq­PþêÇeþì¡Cþç¬d’Ú¤UþèÆê¡C"
            Me.Combo1.Text = PUB_GetUniText(Me.Name, "¿ù»~°h¦^«áµý®Ñ§ó§ï§¹¦¨")
            'end 2022/7/4
      'end 2017/6/29
         Else
            Me.Combo1.Text = "Further to our previous reports, we are pleased to enclose the completed Letters Patent."
         End If
   End Select
End Sub

Private Sub Form_Load()
   Dim nIndex As Integer
   MoveFormToCenter Me
   
   SSTab1.Tab = 0
   
   EnableTextBox textPA01, False
   EnableTextBox textPA02, False
   EnableTextBox textPA03, False
   EnableTextBox textPA04, False
   
   textPA89.BackColor = &H8000000F
   textPA101_2.BackColor = &H8000000F
   textPA60_1.BackColor = &H8000000F
   textPA60_2.BackColor = &H8000000F
   textPA60_3.BackColor = &H8000000F
   'Add By Cheng 2003/01/19
   '§PÂ_¥Ñ¦ó³B¶i¤J
   Select Case m_PrevForm
   Case "frm060317_1" 'ÃÒ®Ñ¨ç
      Me.Combo1.Clear
      Me.Combo1.AddItem "¿ù»~°h¦^"
      Me.Combo1.AddItem "Åý»P"
      Me.Combo1.AddItem "­×¥¿"
      Me.Combo1.AddItem "¦X¨Ö"
      Me.Combo1.AddItem "±ÂÅv"
      Me.Combo1.AddItem "ÅÜ§ó"
      'Modified by Morgan 2017/6/29
      'Me.Combo1.AddItem "§¹¦¨±HÃÒ®Ñ"
      Me.Combo1.AddItem "¿ù»~°h¦^«áµý®Ñ§ó§ï§¹¦¨"
      'end 2017/6/29
   'Modify By Sindy 2022/3/2
   Case Else
      Me.Combo1.Clear
      Me.Combo1.AddItem "This case has been allowed. If your client(s) want(s) to maintain this case, please notify us immediately."
   '2022/3/2 END
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' §R°£¦ê¦Cµ²ºc
    If m_InventorListCount > 0 Then
       Erase m_InventorList
    End If
    m_InventorListCount = 0
    'Add By Cheng 2002/07/18
    Set frm060316_2 = Nothing
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' ²M°£·j´MªºKey
   If bClear = True Then
      m_PA01 = Empty
      m_PA02 = Empty
      m_PA03 = Empty
      m_PA04 = Empty
      m_CP09 = Empty
      m_CP43 = Empty
      m_PrevForm = Empty
      'Add by Morgan 2004/7/27
      m_CP05 = Empty
   End If
   
   Select Case nType
      ' ¥»©Ò®×¸¹ Äæ¦ì1
      Case 0: m_PA01 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì2
      Case 1: m_PA02 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì3
      Case 2: m_PA03 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì4
      Case 3: m_PA04 = strData
      ' Á`¦¬¤å¸¹
      Case 4: m_CP09 = strData
      ' «eµe­±ªº¦WºÙ
      Case 5: m_PrevForm = strData
      'Add by Morgan 2004/7/27
      '¨Ó¨ç¦¬¤å¤é
      Case 6: m_CP05 = strData
   End Select
End Sub

Private Sub cmdok_Click()
   If CheckDataValid() = True Then
      OnSaveData

      ' ¦^«eµe­±
      Select Case m_PrevForm
'Removed by Morgan 2022/11/23 2015¤w¨ú®ø
'         Case "frm060316_1":
'            frm060316_1.Show
'            frm060316_1.Clear
'            frm060316_1.SetInputFocus
'end 2022/11/23
         Case "frm060317_1":
            frm060317_1.Show
            frm060317_1.Clear
      End Select
   
      Unload Me
   End If
End Sub

' ¼W¥[µo©ú¤H
Private Sub AddInventor(ByVal strInventor As String)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strIN01 As String, strIN02 As String
   
   If IsEmptyText(strInventor) Then
      ReDim Preserve m_InventorList(m_InventorListCount + 1)
      m_InventorList(m_InventorListCount).iN01 = Empty
      m_InventorList(m_InventorListCount).iN02 = Empty
      m_InventorList(m_InventorListCount).iN04 = Empty
      m_InventorList(m_InventorListCount).IN05 = Empty
      m_InventorList(m_InventorListCount).IN06 = Empty
      m_InventorListCount = m_InventorListCount + 1
      Exit Sub
   End If
   
   ' ¦r¦ê¸Éº¡¤K½X©Î¥u¨ú¤K½X
   If Len(strInventor) > 8 Then
      strIN01 = Mid(strInventor, 1, 8)
   Else
      strIN01 = strInventor & String(8 - Len(strInventor), "0")
   End If
   
   strSql = "SELECT * FROM INVENTOR " & _
            "WHERE IN01 = '" & strIN01 & "' "
   'Added by Morgan 2014/10/7
   If Len(strInventor) = 10 Then
      strSql = strSql & " and IN02='" & Right(strInventor, 2) & "'"
   End If
   'end 2014/10/7
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Do While rsTmp.EOF = False
         ReDim Preserve m_InventorList(m_InventorListCount + 1)
         If IsNull(rsTmp.Fields("IN01")) = False Then
            m_InventorList(m_InventorListCount).iN01 = rsTmp.Fields("IN01") '«È¤á½s¸¹(8½X)
         End If
         If IsNull(rsTmp.Fields("IN02")) = False Then
            m_InventorList(m_InventorListCount).iN02 = rsTmp.Fields("IN02") 'µo©ú¤H¥N¸¹
         End If
         If IsNull(rsTmp.Fields("IN04")) = False Then
            m_InventorList(m_InventorListCount).iN04 = rsTmp.Fields("IN04") '(µo©ú¤H)¤¤¤å¦WºÙ
         End If
         If IsNull(rsTmp.Fields("IN05")) = False Then
            m_InventorList(m_InventorListCount).IN05 = rsTmp.Fields("IN05") '(µo©ú¤H)­^¤å¦WºÙ
         End If
         If IsNull(rsTmp.Fields("IN06")) = False Then
            m_InventorList(m_InventorListCount).IN06 = rsTmp.Fields("IN06") '(µo©ú¤H)¤é¤å¦WºÙ
         End If
         m_InventorListCount = m_InventorListCount + 1
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2014/11/10
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "µo©ú¤H½s¸¹", "¤¤¤å¦WºÙ", "­^¤å¦WºÙ", "¤é¤å¦WºÙ")
   arrGridHeadWidth = Array(200, 1100, 2200, 2200, 2200)
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

'Add By Sindy 2014/11/10
Private Sub Grd1_Click()
Dim nCol As Integer, nRow As Integer
Dim iCol As Integer
   
   With GRD1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
      nCol = .col
      If pPrevRow > 0 Then
         If pPrevRow <> nRow Then
            .row = pPrevRow
            .TextMatrix(pPrevRow, 0) = ""
            If .FixedCols > 0 Then
               .col = .FixedCols - 1
               .CellBackColor = .BackColorFixed
               .CellForeColor = .ForeColor
            End If
            For iCol = .FixedCols To .Cols - 1
               .col = iCol
               .CellBackColor = .BackColor
            Next
         End If
      End If
   
      If nRow > 0 Then
         .row = nRow
         .TextMatrix(nRow, 0) = "V"
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorSel
            .CellForeColor = .ForeColorSel
         End If
         For iCol = .FixedCols To .Cols - 1
           .col = iCol
           .CellBackColor = &HFFC0C0
         Next
      End If
      .col = nCol
      pPrevRow = nRow
      Call SetInventerData(.TextMatrix(nRow, 1))
   End If
   .Visible = True
   End With
End Sub

Public Function QueryData() As Boolean
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim bFind As Boolean
   Dim ii As Integer
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   textPA01 = m_PA01
   textPA02 = m_PA02
   textPA03 = m_PA03
   textPA04 = m_PA04
   
   'Add By Sindy 2014/11/10
   cmdAddRow.Tag = ""
   cmdDelRow.Tag = ""
   GRD1.Clear
   SetGrd
   '2014/11/10 END
   
   pub_QL05 = pub_QL05 & ";¥»©Ò®×¸¹¡G" & textPA01 & "-" & textPA02 & "-" & textPA03 & "-" & textPA04 'Add By Sindy 2010/12/7
   
   bFind = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & m_PA01 & "' AND " & _
                  "PA02 = '" & m_PA02 & "' AND " & _
                  "PA03 = '" & m_PA03 & "' AND " & _
                  "PA04 = '" & m_PA04 & "' "
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/12/7
      bFind = True
      'Add by Morgan 2004/7/20
      m_PA08 = "" & rsTmp.Fields("PA08")
      m_PA14 = "" & rsTmp.Fields("PA14")
      
      ' ®×¥ó¦WºÙ(¤¤)
      If IsNull(rsTmp.Fields("PA05")) = False Then
         textPA05 = rsTmp.Fields("PA05")
      End If
      ' ®×¥ó¦WºÙ(­^)
      If IsNull(rsTmp.Fields("PA06")) = False Then
         textPA06 = rsTmp.Fields("PA06")
      End If
      ' ®×¥ó¦WºÙ(¤é)
      If IsNull(rsTmp.Fields("PA07")) = False Then
         textPA07 = rsTmp.Fields("PA07")
      End If
      'Memo by Amy 2025/08/06  ¤£Äò¿ì¦ý­ã³qª¾ §ï¬° «áÄò­ã»éÂ²³æ³ø§i
      If IsNull(rsTmp.Fields("PA89")) = False Then
         textPA89 = rsTmp.Fields("PA89")
      End If
      'Memo by Amy 2025/08/06  ¤£Äò¿ì¦ý­ã³qª¾ §ï¬° «áÄò­ã»éÂ²³æ³ø§i
      'Add By Cheng 2003/01/02
      '­Y¤£Äò¿ì¦ý­ã³qª¾¬°"Y"
      If Me.textPA89.Text = "Y" Then
          Me.Combo1.ListIndex = 0
      End If
      ' «È¤á®×¥ó®×¸¹
      If IsNull(rsTmp.Fields("PA48")) = False Then
         textPA48 = rsTmp.Fields("PA48")
      End If
      ' Ápµ¸¤H 1 (¤¤)
      If IsNull(rsTmp.Fields("PA51")) = False Then
         textPA51 = rsTmp.Fields("PA51")
      End If
      ' Ápµ¸¤H 1 (­^)
      If IsNull(rsTmp.Fields("PA52")) = False Then
         textPA52 = rsTmp.Fields("PA52")
      End If
      ' Ápµ¸¤H 1 (¤é)
      If IsNull(rsTmp.Fields("PA53")) = False Then
         textPA53 = rsTmp.Fields("PA53")
      End If
      ' Ápµ¸¤H 2 (¤¤)
      If IsNull(rsTmp.Fields("PA54")) = False Then
         textPA54 = rsTmp.Fields("PA54")
      End If
      ' Ápµ¸¤H 2 (­^)
      If IsNull(rsTmp.Fields("PA55")) = False Then
         textPA55 = rsTmp.Fields("PA55")
      End If
      ' Ápµ¸¤H 2 (¤é)
      If IsNull(rsTmp.Fields("PA56")) = False Then
         textPA56 = rsTmp.Fields("PA56")
      End If
      
      textPA139 = "" & rsTmp.Fields("PA139") 'Add by Morgan 2006/10/20
      
      ' ¹êÅé°Æ¥»¦¬¨ü¤H
      If IsNull(rsTmp.Fields("PA101")) = False Then
         textPA101 = rsTmp.Fields("PA101")
      End If
      ' ¹êÅé°Æ¥»Ápµ¸¤H
      If IsNull(rsTmp.Fields("PA102")) = False Then
         textPA102 = rsTmp.Fields("PA102")
      End If
      ' ¹êÅé°Æ¥»©¼©Ò®×¸¹1
      If IsNull(rsTmp.Fields("PA103")) = False Then
         textPA103 = rsTmp.Fields("PA103")
      End If
      ' ¹êÅé°Æ¥»©¼©Ò®×¸¹2
      If IsNull(rsTmp.Fields("PA104")) = False Then
         textPA104 = rsTmp.Fields("PA104")
      End If
      
      m_PA22 = "" & rsTmp.Fields("PA22")
      m_PA178 = "" & rsTmp.Fields("PA178") 'Added by Morgan 2023/2/18
      ' ¥ý¥[¤J¤@µ§ªÅ¥Õªº¸ê®Æ
      AddInventor Empty
      
      ' ¥Ó½Ð¤H1
      If IsNull(rsTmp.Fields("PA26")) = False Then
         If IsEmptyText(rsTmp.Fields("PA26")) = False Then
            AddInventor rsTmp.Fields("PA26")
         End If
      End If
      ' ¥Ó½Ð¤H2
      If IsNull(rsTmp.Fields("PA27")) = False Then
         If IsEmptyText(rsTmp.Fields("PA27")) = False Then
            AddInventor rsTmp.Fields("PA27")
         End If
      End If
      ' ¥Ó½Ð¤H3
      If IsNull(rsTmp.Fields("PA28")) = False Then
         If IsEmptyText(rsTmp.Fields("PA28")) = False Then
            AddInventor rsTmp.Fields("PA28")
         End If
      End If
      ' ¥Ó½Ð¤H4
      If IsNull(rsTmp.Fields("PA29")) = False Then
         If IsEmptyText(rsTmp.Fields("PA29")) = False Then
            AddInventor rsTmp.Fields("PA29")
         End If
      End If
      ' ¥Ó½Ð¤H5
      If IsNull(rsTmp.Fields("PA30")) = False Then
         If IsEmptyText(rsTmp.Fields("PA30")) = False Then
            AddInventor rsTmp.Fields("PA30")
         End If
      End If
      
      '910702 Sieg 1234¼È¶}©ñ­¶
'      If IsNull(rsTmp.Fields("PA60")) Then
'         SSTab1.TabVisible(1) = False
'         Label27.Visible = True
'      Else
'         SSTab1.TabVisible(2) = False
'         Label27.Visible = False
'      End If
      'Add By Sindy 2014/11/10 Åª¨úµo©ú¤H¸ê®Æ
      StrSQLa = "SELECT '' as V,pi06 as µo©ú¤H½s¸¹,in04 as ¤¤¤å¦WºÙ,in05 as ­^¤å¦WºÙ,in06 as ¤é¤å¦WºÙ from PatentInventor,Inventor where pi01=" + CNULL(m_PA01) + " and pi02=" + CNULL(m_PA02) + " and pi03=" + CNULL(m_PA03) + " and pi04=" + CNULL(m_PA04) & _
                " and substr(pi06,1,8)=in01(+) and substr(pi06,9,2)=in02(+)" & _
                " order by pi05 asc"
      If rsA.State <> adStateClosed Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Set GRD1.Recordset = rsA
         Label27.Visible = False '¦³µo©ú¤H¸ê®Æ
      Else
         Label27.Visible = True 'µLµo©ú¤H¸ê®Æ
      End If
      '2014/11/10 END
      
      'Added by Morgan 2014/10/7
      '­Y¦³Åý»P®É¥Ó½Ð¤H¥i¯à·|»Pµo©ú¤H¤£¦P Ex.FCP-36663
      For ii = 1 To GRD1.Rows - 1
         If Trim(GRD1.TextMatrix(ii, 1)) <> "" Then
            If InStr(rsTmp.Fields("PA26") & rsTmp.Fields("PA27") & rsTmp.Fields("PA28") & rsTmp.Fields("PA29") & rsTmp.Fields("PA30"), Left(Trim(GRD1.TextMatrix(ii, 1)), 8)) = 0 Then
               AddInventor Trim(GRD1.TextMatrix(ii, 1))
            End If
         End If
      Next ii
      'end 2014/10/7
      
      ' ±Nµo©ú¤H§ó·s¨ìComboBox¤¤
      If bFind Then
         Dim nIndex As Integer
         Dim nPos As Integer
         For nIndex = 0 To m_InventorListCount - 1
            cmbPA60.AddItem m_InventorList(nIndex).iN01 & m_InventorList(nIndex).iN02
         Next nIndex
      End If
'Memo by Lydia 2021/08/17 §R°£ÂÂµ{¦¡½X¡G±M§Qµo©ú¤H¦b±M§Q°ò¥»ÀÉ60~69

   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/7
   End If
   
   'ADD BY SONIA 2014/4/28 ÃÒ®Ñ¨ç°£¯S©w«È¤á/¥N²z¤H¥~¤£ºÞ¬O§_E¤Æ³£­n¦L¦W±ø
   strSpecNO = False
   If m_PrevForm = "frm060317_1" Then
      'MODIFY BY SONIA 2014/5/9 ¨ú®øY52218¦A¥[¤JX47833,X47833020,X17901010
      'If "" & rsTmp.Fields("PA75") = "Y52218000" Or "" & rsTmp.Fields("PA75") = "Y20085000" Or "" & rsTmp.Fields("PA26") = "X34291000" Or "" & rsTmp.Fields("PA26") = "X21382010" Then
      Select Case "" & rsTmp.Fields("PA75")
         Case "Y20085000", "X34291000", "X21382010", "X47833000", "X47833020", "X17901010"
            strSpecNO = True
      End Select
   End If
   '2014/4/28 END
   
   rsTmp.Close
   
   ' ¨ú±o¬ÛÃöÁ`¦¬¤å¸¹
   If IsEmptyText(m_CP09) = False Then
      strSql = "SELECT CP43 FROM CASEPROGRESS " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CP43")) = False Then
            m_CP43 = rsTmp.Fields("CP43")
         End If
      End If
   End If
   Set rsTmp = Nothing
   
   'Add by Morgan 2006/8/7 ³Æµù­n§PÂ_»y¤å¥Î
   m_LetterLanguage = PUB_GetLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
End Function

Private Sub SetInventerData(ByVal strInventor As String)
   Dim strIN01 As String
   Dim strIN02 As String
   Dim nPos As Integer
   
'   If Len(strInventor) = 11 Then
'      strIN01 = Mid(strInventor, 1, 8)
'      strIN02 = Mid(strInventor, 10, 2)
'   Else
   If Len(strInventor) > 8 Then
      strIN01 = Mid(strInventor, 1, 8)
      strIN02 = Mid(strInventor, 9, 2)
   Else
      strIN01 = Mid(strInventor, 1, 8)
      strIN02 = "00"
   End If
   For nPos = 0 To cmbPA60.ListCount - 1
      If strIN01 = Mid(cmbPA60.List(nPos), 1, 8) And strIN02 = Mid(cmbPA60.List(nPos), 9, 2) Then
         cmbPA60.ListIndex = nPos
         cmbPA60_Click
      End If
   Next nPos
End Sub

Private Function OnSaveData()
   Dim strPA05 As String
   Dim strPA06 As String
   Dim strPA07 As String
   Dim strPA48 As String
   Dim strPA51 As String
   Dim strPA52 As String
   Dim strPA53 As String
   Dim strPA54 As String
   Dim strPA55 As String
   Dim strPA56 As String
   Dim strPA101 As String
   Dim strPA102 As String
   Dim strPA103 As String
   Dim strPA104 As String
   Dim strPA139 As String 'Add by Morgan 2006/10/20
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
   Dim ii As Integer
   
   strPA05 = textPA05
   strPA06 = textPA06
   strPA07 = textPA07
   strPA48 = textPA48
   strPA51 = textPA51
   strPA52 = textPA52
   strPA53 = textPA53
   strPA54 = textPA54
   strPA55 = textPA55
   strPA56 = textPA56
   strPA139 = textPA139 'Add by Morgan 2006/10/20
   
'Memo by Lydia 2021/08/17 §R°£ÂÂµ{¦¡½X¡G±M§Qµo©ú¤H¦b±M§Q°ò¥»ÀÉ60~69

      'Add By Sindy 2014/11/10
      If cmdAddRow.Tag = "I" Or cmdDelRow.Tag = "D" Then '¦³²§°Êµo©ú¤H¸ê®Æ
         '¥þ³¡§R°£,­«·s·s¼W
         strSql = "delete from patentInventor where pi01=" + CNULL(m_PA01) + " and pi02=" + CNULL(m_PA02) + " and pi03=" + CNULL(m_PA03) + " and pi04=" + CNULL(m_PA04)
         Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
         cnnConnection.Execute strSql
         For ii = 1 To GRD1.Rows - 1
            strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                     CNULL(m_PA01) & "," & CNULL(m_PA02) & "," & CNULL(m_PA03) & "," & CNULL(m_PA04) & "," & ii & ",'" & GRD1.TextMatrix(ii, 1) & "')"
            Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
            cnnConnection.Execute strSql
         Next ii
      End If
      '2014/11/10 END
'   End If
   
   If textPA101 <> "" Then strPA101 = textPA101 & String(9 - Len(textPA101), "0")
   strPA102 = textPA102
   strPA103 = textPA103
   strPA104 = textPA104
   'Modify by Morgan 2006/10/20 ¥[ PA139
   'Memo by Lydia 2021/08/17 §R°£ÂÂµ{¦¡½X¡G±M§Qµo©ú¤H¦b±M§Q°ò¥»ÀÉ60~69
   'Modify By Sindy 2014/11/10
   strSql = "UPDATE PATENT SET PA05=" & DBNullString(strPA05) & "," & "PA06=" & DBNullString(strPA06) & "," & "PA07=" & DBNullString(strPA07) & "," & _
                              "PA48=" & DBNullString(strPA48) & "," & "PA51=" & DBNullString(strPA51) & "," & "PA52=" & DBNullString(strPA52) & "," & _
                              "PA53=" & DBNullString(strPA53) & "," & "PA54=" & DBNullString(strPA54) & "," & "PA55=" & DBNullString(strPA55) & "," & _
                              "PA56=" & DBNullString(strPA56) & "," & _
                              "PA101=" & DBNullString(strPA101) & "," & _
                              "PA102=" & DBNullString(strPA102) & "," & "PA103=" & DBNullString(strPA103) & "," & "PA104=" & DBNullString(strPA104) & "," & _
                              "PA139=" & CNULL(ChgSQL(strPA139)) & _
            " WHERE PA01 = '" & m_PA01 & "' AND " & _
                  "PA02 = '" & m_PA02 & "' AND " & _
                  "PA03 = '" & m_PA03 & "' AND " & _
                  "PA04 = '" & m_PA04 & "' "
   '2014/11/10 END
   cnnConnection.Execute strSql
   'Remove by Morgan 2006/8/7 QueryData
'   'Modify by Morgan 2006/6/2
'   'm_LetterLanguage = GetLetterLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
'   m_LetterLanguage = PUB_GetLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
   
   'Modify By Sindy 2015/9/14
   If UCase(m_PrevForm) = UCase("frm060317_1") Then
      Call frm060317_1.QueryLetterData(Me.Combo1.Text)
      Exit Function
   End If
   '2015/9/14 END
   
   '­^¤å©w½Z
   'Modify by Morgan 2004/10/13 ¥[¤é¤å
   'If m_LetterLanguage = 2 Then
   If m_LetterLanguage = 2 Or m_LetterLanguage = 3 Then
      'ÃÒ®Ñ¨ç­^¤é¤å©w½Z
      If m_PrevForm = "frm060317_1" Then
         'ÃÒ®ÑºØÃþ
         'Modify by Morgan 2006/7/26
         'm_LetterKind = GetLetterKind(m_PA01, m_PA02, m_PA03, m_PA04)
         m_LetterKind = frm060317_1.GetLetterKind(m_PA01, m_PA02, m_PA03, m_PA04)
         
'Removed by Morgan 2022/11/23 2015¤w¨ú®ø
'      'Add By Cheng 2003/01/02
'      '®Ö­ã¨ç­^¤é¤å©w½Z
'      ElseIf m_PrevForm = "frm060316_1" Then
'         StrSQLa = "Select COUNT(*) From PriDate Where PD01='" & m_PA01 & "' AND PD02='" & m_PA02 & "' AND PD03='" & m_PA03 & "' AND PD04='" & m_PA04 & "' "
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'             If rsA.Fields(0).Value > 0 Then
'                 m_blnPriData = True
'                 '­Y¦³¤T­Ó¥H¤WÀu¥ýÅv¸ê®Æ
'                 If rsA.Fields(0).Value >= 3 Then
'                     m_bln3PriData = True
'                 Else
'                     m_bln3PriData = False
'                 End If
'             Else
'                 m_blnPriData = False
'                 m_bln3PriData = False
'             End If
'         Else
'            m_blnPriData = False
'            m_bln3PriData = False
'         End If
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'end 2022/11/23

      End If
   End If
   m_bolEmail = False
   PrintLetter

'MODIFY BY SONIA 2014/4/28 ÃÒ®Ñ¨ç­Y¯S©w«È/¥N¥~³£­n¥[¦L¦a§}±ø
'   If Not m_bolEmail Or m_bolPlusPaper Then
'       '·s¼W¦a§}±ø¦Cªí¸ê®Æ
'       pub_AddressListSN = pub_AddressListSN + 1
'       PUB_AddNewAddressList strUserNum, textPA01.Text, textPA02.Text, Left(textPA03.Text & "0", 1), Left(textPA04.Text & "00", 2), "" & pub_AddressListSN, "0"
'   End If

   If m_PrevForm = "frm060317_1" Then
      If Not strSpecNO Then
         '·s¼W¦a§}±ø¦Cªí¸ê®Æ
         pub_AddressListSN = pub_AddressListSN + 1
         PUB_AddNewAddressList strUserNum, textPA01.Text, textPA02.Text, Left(textPA03.Text & "0", 1), Left(textPA04.Text & "00", 2), "" & pub_AddressListSN, "0"
      End If
      
'Removed by Morgan 2022/11/23 2015¤w¨ú®ø
'   ElseIf m_PrevForm = "frm060316_1" Then
'      If Not m_bolEmail Or m_bolPlusPaper Then
'          '·s¼W¦a§}±ø¦Cªí¸ê®Æ
'          pub_AddressListSN = pub_AddressListSN + 1
'          PUB_AddNewAddressList strUserNum, textPA01.Text, textPA02.Text, Left(textPA03.Text & "0", 1), Left(textPA04.Text & "00", 2), "" & pub_AddressListSN, "0"
'      End If
'end 2022/11/23
   
   End If
'2014/4/28 END
   
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bFind As Boolean
   Dim nIndex As Integer
   
   CheckDataValid = False
   
   ' ®×¥ó¤¤­^¤é
   If IsEmptyText(textPA05) And IsEmptyText(textPA06) And IsEmptyText(textPA07) Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "®×¥ó¤¤­^¤é¤å¦WºÙ¤£¥i¦P®ÉªÅ¥Õ"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      textPA05.SetFocus
      GoTo EXITSUB
   End If
   
'   If SSTab1.TabVisible(1) Then
'      ' µo©ú¤H¤£¥i¥þªÅ¥Õ
'      bFind = False
'      nIndex = 0
'      For nIndex = 0 To 9
'         If IsEmptyText(cmbPA60(nIndex).Text) = False Then
'            bFind = True
'         End If
'      Next nIndex
'      If bFind = False Then
'         strTit = "ÀË®Ö¸ê®Æ"
'         strMsg = "½Ð¿é¤J¦Ü¤Ö¤@­Óµo©ú¤H"
'         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
'         GoTo EXITSUB
'      End If
'      ' ÀË¬dµo©ú¤H¸ê®Æªº¿é¤J¬O§_³sÄò
'      bFind = False
'      For nIndex = 0 To 9
'         If IsEmptyText(cmbPA60(nIndex).Text) = True Then
'            Dim nPos As Integer
'            For nPos = nIndex To 9
'               If IsEmptyText(cmbPA60(nPos).Text) = False Then
'                  bFind = True
'                  Exit For
'               End If
'            Next nPos
'         End If
'         If bFind Then
'            Exit For
'         End If
'      Next nIndex
'      If bFind Then
'         strTit = "ÀË®Ö¸ê®Æ"
'         strMsg = "µo©ú¤H¸ê®Æ¥²¶·«ö·Ó¶¶§Ç³sÄò¿é¤J!"
'         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
'         GoTo EXITSUB
'      End If
      'Modify By Sindy 2014/11/10
      If Label27.Visible = False Then '¦³µo©ú¤H¸ê®Æ 'Add By Sindy 2014/12/16
         'µo©ú¤H¤£¥i¥þªÅ¥Õ
         If GRD1.Rows = 2 And GRD1.TextMatrix(1, 1) = "" Then
            strTit = "ÀË®Ö¸ê®Æ"
            strMsg = "½Ð¿é¤J¦Ü¤Ö¤@­Óµo©ú¤H"
            nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
            GoTo EXITSUB
         End If
      End If
      '2014/11/10 END
'   End If
   
   'Added by Sindy 2022/3/2 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
   If PUB_ChkUniText(Me, , True) = False Then
      Exit Function
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' ®×¥ó¤¤¤å¦WºÙ
Private Sub textPA05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA05) = False Then
      If StrLength(textPA05) > 160 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "®×¥ó¤¤¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'If Cancel = False Then textPA05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' ®×¥ó­^¤å¦WºÙ
Private Sub textPA06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA06) = False Then
      If StrLength(textPA06) > 180 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "®×¥ó­^¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

' ®×¥ó¤é¤å¦WºÙ
Private Sub textPA07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA07) = False Then
      If StrLength(textPA07) > 160 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "®×¥ó¤é¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'If Cancel = False Then textPA07.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textPA101_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPA139_GotFocus()
   InverseTextBox textPA139
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'textPA139.IMEMode = 1
   OpenIme
End Sub

Private Sub textPA139_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA139) = False Then
      'Modified by Lydia 2017/06/14
      'If StrLength(textPA139) > textPA139.MaxLength Then
      If StrLength(textPA139) > 60 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "Ápµ¸¤H³¡ªù(¤é)¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'If Cancel = False Then textPA139.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' «È¤á®×¥ó®×¸¹
Private Sub textPA48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA48) = False Then
      If StrLength(textPA48) > 30 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "«È¤á®×¥ó®×¸¹¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

' Ápµ¸¤H1 (¤¤)
Private Sub textPA51_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA51) = False Then
      'Modified by Lydia 2017/06/14 Ápµ¸¤H(¤¤)§ï¬°30¦r
      'If StrLength(textPA51) > 10 Then
      If StrLength(textPA51) > 30 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "Ápµ¸¤H1¤¤¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'If Cancel = False Then textPA51.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' Ápµ¸¤H1 (­^)
Private Sub textPA52_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA52) = False Then
      'Modified by Lydia 2017/06/14
      'If StrLength(textPA52) > textPA52.MaxLength Then
      If StrLength(textPA52) > 35 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "Ápµ¸¤H1­^¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

' Ápµ¸¤H1 (¤é)
Private Sub textPA53_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA53) = False Then
      'Modified by Lydia 2017/06/14
      'If StrLength(textPA53) > 20 Then
      If StrLength(textPA53) > 60 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "Ápµ¸¤H1¤é¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'If Cancel = False Then textPA53.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' Ápµ¸¤H2 (¤¤)
Private Sub textPA54_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA54) = False Then
      'Modified by Lydia 2017/06/14 Ápµ¸¤H(¤¤)§ï¬°30¦r
      'If StrLength(textPA54) > 10 Then
      If StrLength(textPA54) > 30 Then
        Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "Ápµ¸¤H2¤¤¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'If Cancel = False Then textPA54.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' Ápµ¸¤H2 (­^)
Private Sub textPA55_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA55) = False Then
      'Modified by Lydia 2017/06/14
      'If StrLength(textPA55) > textPA55.MaxLength Then
      If StrLength(textPA55) > 35 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "Ápµ¸¤H2­^¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

' Ápµ¸¤H2 (¤é)
Private Sub textPA56_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA56) = False Then
      'Modified by Lydia 2017/06/14
      'If StrLength(textPA56) > 20 Then
      If StrLength(textPA56) > 60 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "Ápµ¸¤H2¤é¤å¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'If Cancel = False Then textPA56.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' ¹êÅé°Æ¥»¦¬¨ü¤H
Private Sub textPA101_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textPA101_2 = Empty
   If IsEmptyText(textPA101) = False Then
      Select Case Mid(textPA101, 1, 1)
         Case "X":
            textPA101_2 = GetCustomerName(textPA101, 0)
         Case "Y":
            textPA101_2 = GetFAgentName(textPA101)
         Case Else:
      End Select
      If IsEmptyText(textPA101_2) = True Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¹êÅé°Æ¥»¦¬¨ü¤H¥N½X¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

' ¹êÅé°Æ¥»Ápµ¸¤H
Private Sub textPA102_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA102) = False Then
      If StrLength(textPA102) > textPA102.MaxLength Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¹êÅé°Æ¥»Ápµ¸¤H¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

' ¹êÅé°Æ¥»¦¬¨ü¤H©¼©Ò®×¸¹1
Private Sub textPA103_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA103) = False Then
      If StrLength(textPA103) > 140 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¹êÅé°Æ¥»¦¬¨ü¤H©¼©Ò®×¸¹1¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

' ¹êÅé°Æ¥»¦¬¨ü¤H©¼©Ò®×¸¹2
Private Sub textPA104_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPA104) = False Then
      If StrLength(textPA104) > 140 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¹êÅé°Æ¥»¦¬¨ü¤H©¼©Ò®×¸¹2¦WºÙ¤º®e¤Óªø"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

Private Sub textPA05_GotFocus()
   InverseTextBox textPA05
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'textPA05.IMEMode = 1
   OpenIme
End Sub

Private Sub textPA06_GotFocus()
   InverseTextBox textPA06
End Sub

Private Sub textPA07_GotFocus()
   InverseTextBox textPA07
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'textPA07.IMEMode = 1
   OpenIme
End Sub

Private Sub textPA48_GotFocus()
   InverseTextBox textPA48
End Sub

Private Sub textPA51_GotFocus()
   InverseTextBox textPA51
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'textPA51.IMEMode = 1
   OpenIme
End Sub

Private Sub textPA52_GotFocus()
   InverseTextBox textPA52
End Sub

Private Sub textPA53_GotFocus()
   InverseTextBox textPA53
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'textPA53.IMEMode = 1
   OpenIme
End Sub

Private Sub textPA54_GotFocus()
   InverseTextBox textPA54
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'textPA54.IMEMode = 1
   OpenIme
End Sub

Private Sub textPA55_GotFocus()
   InverseTextBox textPA55
End Sub

Private Sub textPA56_GotFocus()
   InverseTextBox textPA56
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'textPA56.IMEMode = 1
   OpenIme
End Sub

Private Sub textPA101_GotFocus()
   InverseTextBox textPA101
End Sub

Private Sub textPA102_GotFocus()
   InverseTextBox textPA102
End Sub

Private Sub textPA103_GotFocus()
   InverseTextBox textPA103
End Sub

Private Sub textPA104_GotFocus()
   InverseTextBox textPA104
End Sub

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
   Dim strSql As String
   Dim strTemp As String
   Dim strKey As String
   Dim ET01 As String, ET02 As String, ET03 As String
   
   Select Case m_PrevForm
   
'Removed by Morgan 2022/11/23 2015¤w¨ú®ø
'      Case "frm060316_1"
'         ET01 = "04"
'         ET02 = m_CP43
'         ' ©w½Z»y¤å
'         Select Case m_LetterLanguage
'            ' ¤¤¤å
'            Case "1":
'               ET03 = "01"
'            ' ­^¤å
'            Case "2":
''               ET03 = "02"
'               'Add By Cheng 2003/01/03
'               '­Y¦³¤T­ÓÀu¥ýÅv¸ê®Æ
'               If m_bln3PriData = True Then
'                    ET03 = "05"
'                    'Add By Cheng 2003/02/17
'                    'ªþ¥ó
'                    EndLetter ET01, ET02, "08", strUserNum
'                '­Y¦³Àu¥ýÅv¸ê®Æ
'                ElseIf m_blnPriData = True Then
'                    ET03 = "02"
'                    'Add By Cheng 2003/02/17
'                    'ªþ¥ó
'                    EndLetter ET01, ET02, "06", strUserNum
'                '­YµLÀu¥ýÅv¸ê®Æ
'                Else
'                    ET03 = "04"
'                    'Add By Cheng 2003/02/17
'                    'ªþ¥ó
'                    EndLetter ET01, ET02, "07", strUserNum
'                End If
'            ' ¤é¤å
'            Case "3":
'               ET03 = "03"
'            Case Else:
'         End Select
'end 2022/11/23

      Case "frm060317_1"
         ET01 = "07"
         ET02 = m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603"
         ' ©w½Z»y¤å
         Select Case m_LetterLanguage
            ' ¤¤¤å
            Case "1":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               ET03 = "01"
            ' ­^¤å
            Case "2":
               Select Case m_LetterKind
                  Case 1:
                     'Modify by Morgan 2004/7/20
                     '±±¨î93.7.1¥H«á¥Î·s©w½Z
                     If Val(m_PA14) > 0 And Val(m_PA14) < 20040701 Then
                        ET03 = "02"
                     '·s«¬
                     ElseIf m_PA08 = "2" Then
                        ET03 = "14"
                     Else
                        ET03 = "13"
                     End If
                     
                  Case 2:
                     'Modify by Morgan 2004/7/20
                     '±±¨î93.7.1¥H«á¥Î·s©w½Z
                     If Val(m_PA14) > 0 And Val(m_PA14) < 20040701 Then
                        ET03 = "03"
                     '·s«¬
                     ElseIf m_PA08 = "2" Then
                        ET03 = "18"
                     Else
                        ET03 = "17"
                     End If
'Modify by Morgan 2005/1/26 ¤£ºÞ¬O§_¦³¤U¦¸Ãº¶O¤é²Î¤@¥X·s©w½Z--David
                  Case 3, 5
                        If m_PA08 = "3" Then
                           ET03 = "19"
                        Else
                           ET03 = "21"
                        End If
                  Case 4:
                     ET03 = "05"
               End Select
            
            'Add by Morgan 2006/7/26
            Case "3" '¤é¤å
               '¦Û°Ê¥NÃº
               If m_LetterKind = "2" Then
                  ET03 = "09"
               Else
                  ET03 = "06"
               End If
            Case Else:
         End Select
      Case Else
   End Select
   
   EndLetter ET01, ET02, ET03, strUserNum
   
   
   Dim i As Integer, j As Integer
   Dim strTxt(1 To 30) As String
   Dim stET03(1 To 2) As String
   
   j = 1
   'ÃÒ®Ñ¨ç¨Ò¥~Äæ¦ì
   If m_PrevForm = "frm060317_1" Then
      
      'Added by Morgan 2013/8/5
      '¤@®×¨â½Ð´£¿ô
      If m_PA08 = "2" Then
         strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa11,pa77,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
            " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and " & ChgCaseMap(m_PA01 & m_PA02 & m_PA03 & m_PA04, , 0) & _
            " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and " & ChgCaseMap(m_PA01 & m_PA02 & m_PA03 & m_PA04, , 1) & ") X" & _
            ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('07','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¤@®×¨â½Ð·s«¬®×­n¦L','¡ð')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('07','" & ET02 & "','" & ET03 & "','" & strUserNum & "','µo©ú®×¥Ó½Ð¸¹','" & RsTemp("pa11") & "')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','µo©ú®×©¼©Ò®×¸¹','" & IIf(IsNull(RsTemp("pa77")), "", "" & RsTemp("pa77")) & "')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('07','" & ET02 & "','" & ET03 & "','" & strUserNum & "','µo©ú®×¥»©Ò®×¸¹','" & RsTemp("CNo") & "')"
            cnnConnection.Execute strSql
         End If
      End If
      'end 2013/8/5
      
      ' ©w½Z»y¤å
      Select Case m_LetterLanguage
         ' ¤¤¤å
         Case "1":
              '¦Ü¤U¤@µ{§ÇÀÉ¤¤§ä¤U¤@µ{§Ç¥N¸¹¬OÃº¦~¶O¤Î¬O§_Äò¿ì¬°ªÅ¡A¬O«h¤@¯ë¡A­YªÅªº«h¬O³Ì«á¤@¦¸¦~¶O
              strExc(0) = "SELECT np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(m_PA01 & m_PA02 & m_PA03 & m_PA04) & _
                 " AND NP07=" & ¦~¶O & " AND NP06 IS NULL"
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
              If intI = 1 Then
                 If RsTemp.Fields(0) <> "" Then
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                       "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & CNULL(DBDATE(RsTemp.Fields(0))) & ")"
                      cnnConnection.Execute strSql
                 End If
              End If
            
         ' ­^¤å
         Case "2":
            Select Case m_LetterKind
               Case 1:
                   'Add by Morgan 2005/6/23
                   strTemp = ""
                   Select Case Left(m_PA22, 1)
                     Case "I"
                        strTemp = "Please note that the patent number includes ""I"" for ""Invention"" patent."
                     Case "M"
                        strTemp = "Please note that the patent number includes ""M"" for ""Utility Model"" patent."
                     Case "D"
                        strTemp = "Please note that the patent number includes ""D"" for ""Design"" patent."
                  End Select
                  If strTemp <> "" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','±M§Q¸¹»¡©ú','" & strTemp & "')"
                     cnnConnection.Execute strSql
                  End If
                   '2005/6/23 end
          
                  'Add By Cheng 2003/01/19
                  '¶Ç¤J¨Ò¥~Äæ¦ì­È
                  '¤U¦¸Ãº¦~¶O¤é
                  'Modify by Morgan 2006/7/26
                  'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é','" & GetEngMMDD(CompDate(2, -1, GetPA14(m_PA01 & m_PA02 & m_PA03 & m_PA04))) & "')"
                  strExc(1) = CompDate(2, -1, GetPA14(m_PA01 & m_PA02 & m_PA03 & m_PA04))
                  If Right(strExc(1), 4) = "0229" Then
                     strExc(1) = Left(strExc(1), 4) & "0228"
                  End If
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é','" & GetEngMMDD(strExc(1)) & "')"
                  'end 2006/7/26
                  cnnConnection.Execute strSql
                  '¦Ü¤U¤@µ{§ÇÀÉ¤¤§ä¤U¤@µ{§Ç¥N¸¹¬OÃº¦~¶O¤Î¬O§_Äò¿ì¬°ªÅ¡A¬O«h¤@¯ë¡A­YªÅªº«h¬O³Ì«á¤@¦¸¦~¶O
                  strExc(0) = "SELECT np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(m_PA01 & m_PA02 & m_PA03 & m_PA04) & _
                     " AND NP07=" & ¦~¶O & " AND NP06 IS NULL"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp.Fields(0) <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & CNULL(DBDATE(RsTemp.Fields(0))) & ")"
                          cnnConnection.Execute strSql
                     End If
                  End If
                  
               Case 2:
                  EndLetter "07", ET02, ET03, strUserNum
                  
                   'Add by Morgan 2005/6/23
                   strTemp = ""
                   Select Case Left(m_PA22, 1)
                     Case "I"
                        strTemp = "Please note that the patent number includes ""I"" for ""Invention"" patent."
                     Case "M"
                        strTemp = "Please note that the patent number includes ""M"" for ""Utility Model"" patent."
                     Case "D"
                        strTemp = "Please note that the patent number includes ""D"" for ""Design"" patent."
                  End Select
                  If strTemp <> "" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','±M§Q¸¹»¡©ú','" & strTemp & "')"
                     cnnConnection.Execute strSql
                  End If
                   '2005/6/23 end
                   
                  'Add By Cheng 2003/01/19
                  '¶Ç¤J¨Ò¥~Äæ¦ì­È
                  '¤U¦¸Ãº¦~¶O¤é
                  'Modify by Morgan 2006/7/26
                  'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é','" & GetEngMMDD(CompDate(2, -1, GetPA14(m_PA01 & m_PA02 & m_PA03 & m_PA04))) & "')"
                  strExc(1) = CompDate(2, -1, GetPA14(m_PA01 & m_PA02 & m_PA03 & m_PA04))
                  If Right(strExc(1), 4) = "0229" Then
                     strExc(1) = Left(strExc(1), 4) & "0228"
                  End If
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é','" & GetEngMMDD(strExc(1)) & "')"
                  'end 2006/7/26
                  cnnConnection.Execute strSql
                  '¦Ü¤U¤@µ{§ÇÀÉ¤¤§ä¤U¤@µ{§Ç¥N¸¹¬OÃº¦~¶O¤Î¬O§_Äò¿ì¬°ªÅ¡A¬O«h¤@¯ë¡A­YªÅªº«h¬O³Ì«á¤@¦¸¦~¶O
                  strExc(0) = "SELECT np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(m_PA01 & m_PA02 & m_PA03 & m_PA04) & _
                     " AND NP07=" & ¦~¶O & " AND NP06 IS NULL"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp.Fields(0) <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & CNULL(DBDATE(RsTemp.Fields(0))) & ")"
                          cnnConnection.Execute strSql
                     End If
                  End If
               'Modify by Morgan 2005/1/26 ¤£ºÞ¬O§_¦³¤U¦¸Ãº¶O¤é²Î¤@¥X·s©w½Z--David
               Case 3, 5
                  EndLetter "07", ET02, ET03, strUserNum
                  
                   'Add by Morgan 2005/6/23
                   strTemp = ""
                   Select Case Left(m_PA22, 1)
                     Case "I"
                        strTemp = "Please note that the patent number includes ""I"" for ""Invention"" patent."
                     Case "M"
                        strTemp = "Please note that the patent number includes ""M"" for ""Utility Model"" patent."
                     Case "D"
                        strTemp = "Please note that the patent number includes ""D"" for ""Design"" patent."
                  End Select
                  If strTemp <> "" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','±M§Q¸¹»¡©ú','" & strTemp & "')"
                     cnnConnection.Execute strSql
                  End If
                   '2005/6/23 end
                   
                  'Add By Cheng 2003/01/19
                  '¶Ç¤J¨Ò¥~Äæ¦ì­È
                  '¤U¦¸Ãº¦~¶O¤é
                  'Modify by Morgan 2006/7/26
                  'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é','" & GetEngMMDD(CompDate(2, -1, GetPA14(m_PA01 & m_PA02 & m_PA03 & m_PA04))) & "')"
                  strExc(1) = CompDate(2, -1, GetPA14(m_PA01 & m_PA02 & m_PA03 & m_PA04))
                  If Right(strExc(1), 4) = "0229" Then
                     strExc(1) = Left(strExc(1), 4) & "0228"
                  End If
                  
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é','" & GetEngMMDD(strExc(1)) & "')"
                  'end 2006/7/26
                  cnnConnection.Execute strSql
                  '¦Ü¤U¤@µ{§ÇÀÉ¤¤§ä¤U¤@µ{§Ç¥N¸¹¬OÃº¦~¶O¤Î¬O§_Äò¿ì¬°ªÅ¡A¬O«h¤@¯ë¡A­YªÅªº«h¬O³Ì«á¤@¦¸¦~¶O
                  '§ì¥À®×¤§¤U¦¸Ãº¶O¤é
                  'Modify by Morgan 2010/12/27 ¥Ó½Ð®×¸¹§ï½X¼Æ
                  strExc(0) = "SELECT PA01||PA02||PA03||PA04 FROM PATENT WHERE PA01='" & m_PA01 & "' AND PA11 = ( " & _
                              "SELECT SUBSTR(PA11,1,9) FROM PATENT WHERE PA01='" & m_PA01 & "' AND " & _
                              "PA02='" & m_PA02 & "' AND PA03='" & m_PA03 & "' AND PA04='" & m_PA04 & "') "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp.Fields(0) <> "" Then
                        strTemp = RsTemp.Fields(0)
                        strExc(0) = "SELECT np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(strTemp) & _
                           " AND NP07=" & ¦~¶O & " AND NP06 IS NULL"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If RsTemp.Fields(0) <> "" Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & CNULL(DBDATE(RsTemp.Fields(0))) & ")"
                                cnnConnection.Execute strSql
                           End If
                        End If
                     End If
                  End If
                   'Add By Cheng 2003/01/19
                    ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                    EndLetter "07", ET02, "09", strUserNum
              'Add By Cheng 2003/07/25
               Case 4: 'µL¤U¦¸Ãº¶O¤é(¤@¯ë©Î»âÃÒ¦Û°Ê¥NÃº)
                  ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                  EndLetter "07", ET02, "05", strUserNum
                  
                   'Add by Morgan 2005/6/23
                   strTemp = ""
                   Select Case Left(m_PA22, 1)
                     Case "I"
                        strTemp = "Please note that the patent number includes ""I"" for ""Invention"" patent."
                     Case "M"
                        strTemp = "Please note that the patent number includes ""M"" for ""Utility Model"" patent."
                     Case "D"
                        strTemp = "Please note that the patent number includes ""D"" for ""Design"" patent."
                  End Select
                  If strTemp <> "" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & "07" & "','" & ET02 & "','05','" & strUserNum & "','±M§Q¸¹»¡©ú','" & strTemp & "')"
                     cnnConnection.Execute strSql
                  End If
                   '2005/6/23 end
            End Select
         
         Case "3" ' ¤é¤å
            
            EndLetter "07", ET02, "07", strUserNum
            
            strExc(0) = "SELECT np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(m_PA01 & m_PA02 & m_PA03 & m_PA04) & _
               " AND NP07=" & ¦~¶O & " AND NP06 IS NULL"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) <> "" Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & CNULL(DBDATE(RsTemp.Fields(0))) & ")"
                  cnnConnection.Execute strSql
               End If
            End If
            '¤U¦¸Ãº¦~¶O¤é
            strExc(1) = CompDate(2, -1, GetPA14(m_PA01 & m_PA02 & m_PA03 & m_PA04))
            If Right(strExc(1), 4) = "0229" Then
               strExc(1) = Left(strExc(1), 4) & "0228"
            End If
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é'," & strExc(1) & ")"
            cnnConnection.Execute strSql
                              
            If m_PA08 = "2" Then
            
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','·s«¬§Þ³N³ø§i´£¥Ü','Çeþò¡BüÁ¥Î·s®×ÇU“¸§QªÌÇV¡BüÁ¥Î·s®×§Þ³Nµû’þ®ÑÇy´£¥ÜþêþùÄµ§iÇyþêþò«áþúÇQþäÇsÇW¡BþðÇU“¸§QÇy¦æ¨ÏþìÇrþæÇOþßþúþàÇeþîÇz¡C')"
               cnnConnection.Execute strSql
               
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & "07" & "','" & ET02 & "','" & "07" & "','" & strUserNum & "','·s«¬§Þ³N³ø§iªk±ø','²Ä104†A¡@üÁ¥Î·s®×“¸ªÌÇV¡BüÁ¥Î·s®×§Þ³Nµû’þ®ÑÇy´£¥ÜþêþùÄµ§iÇyþêþò«áþú" & vbCrLf & "¡@¡@¡@¡@¡@ÇQþäÇsÇW¡BþðÇU“¸§QÇy¦æ¨ÏþìÇrþæÇOþßþúþàÇQÆê¡C')"
               cnnConnection.Execute strSql
            End If
         Case Else:
      End Select
     
      'Added by Morgan 2014/11/5
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " select '" & "07" & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','®Ö¹ï¤w­ã±M§Q¤¤­n¦L','¡ð'" & _
         " from caseprogress where CP01='" & m_PA01 & "' AND CP02='" & m_PA02 & "' AND CP03='" & m_PA03 & "' AND CP04='" & m_PA04 & "'" & _
         " and cp10='926' and cp27||cp57 is null and rownum=1"
      cnnConnection.Execute strSql
      'end 2014/11/5
   
      'Add by Morgan 2006/7/26 ¤é¤åÄ¶¤å
      If m_LetterLanguage = "3" Then
         stET03(1) = "07"
      Else
      
         'Modify by Morgan 2005/1/26 ¤£ºÞ¬O§_¦³¤U¦¸Ãº¶O¤é²Î¤@¥X·s©w½Z--David
         stET03(1) = "20"
         If m_PA08 = "2" Then
            stET03(2) = "16"
         Else
            stET03(2) = "15"
         End If
         EndLetter ET01, ET02, stET03(1), strUserNum
         EndLetter ET01, ET02, stET03(2), strUserNum
      End If
   End If
      
   '¦C¦L³Æµù
   If Me.Combo1.Text <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
            "'¦C¦L³Æµù','" & IIf(m_LetterLanguage = "3", "°l¦ù¡G", "P.S.  ") & ChgSQL(Me.Combo1.Text) & "')"
      j = j + 1
   End If
   'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
   'If Not objLawDll.ExecSQL(j - 1, strTxt) Then
   If Not ClsLawExecSQL(j - 1, strTxt) Then
      MsgBox "¨Ò¥~Äæ¦ìÀx¦s¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
   End If
End Sub

Private Sub PrintLetter()
   'Add by Morgan 2004/7/27
   Dim stET03 As String
   Dim stContent As String
   Dim strKey As String
   Dim strBillNo As String '«Ý¦L½Ð´Ú³æ¸¹ Add by Morgan 2011/6/24
   Dim iBillPageCount As Integer '½Ð´Ú³æ­¶¼Æ Add by Morgan 2011/6/27
   
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   'InsExpField 'Removed by Morgan 2014/12/10
   Select Case m_PrevForm
'Removed by Morgan 2022/11/23 2015¤w¨ú®ø
'      Case "frm060316_1"
'         InsExpField 'Added by Morgan 2014/12/10
'         'Add by Morgan 2011/7/8
'         pub_OsPrinter = PUB_GetOsDefaultPrinter
'         PUB_SetOsDefaultPrinter frm060316_1.Combo2.Text
'         PUB_SetWordActivePrinter
'         'end 2011/7/8
'
'         'Add by Morgan 2008/3/24 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
'         m_bolEmail = PUB_GetEMailFlag(m_PA01 & m_PA02 & m_PA03 & m_PA04, , , m_bolPlusPaper)
'
'         'Added by Morgan 2014/6/3
'         If m_bolEmail = False Then
'            m_bolDNEmail = PUB_GetEMailFlag(m_PA01 & m_PA02 & m_PA03 & m_PA04, , , m_bolDNPlusPaper, , True)
'         Else
'            m_bolDNEmail = m_bolEmail
'            m_bolDNPlusPaper = m_bolPlusPaper
'         End If
'         'end 2014/6/3
'
'         'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
'         If m_bolPlusPaper Then
'            m_iCopy = 0
'         Else
'            m_iCopy = 1
'         End If
'         'end 2009/10/20
'
'         'Add by Morgan 2011/6/27
'         frm060316_1.m_bPrintBill = PUB_GetUnPaidBill(m_PA01, m_PA02, m_PA03, m_PA04, strBillNo)
'         '¦C¦L½Ð´Ú³æ
'         If strBillNo <> "" Then
'            'Modified by Morgan 2014/6/3
'            'PUB_PrintBill strBillNo, frm060316_1.Combo2.Text, m_bolEmail, m_bolPlusPaper, Me.Name, iBillPageCount, 2
'            PUB_PrintBill strBillNo, frm060316_1.Combo2.Text, m_bolDNEmail, m_bolDNPlusPaper, Me.Name, iBillPageCount, 2
'            'end 2014/6/3
'            frm060316_1.m_iBillPageCount = iBillPageCount
'         End If
'         'end 2011/6/27
'
'         ' ©w½Z»y¤å
'         Select Case m_LetterLanguage
'            ' ¤¤¤å
'            Case "1":
'               ' ¦C¦L©w½Z
'               NowPrint m_CP43, "04", "01", False, strUserNum, 0, , , , , , , , , True
'            ' ­^¤å
'            Case "2":
'               'Add by Morgan 2004/7/27
'               '93.7.1¥H«á¥Î¤G¦X¤@©w½Z
'               '·sªk
'               If Val(m_CP05) >= 930701 Then
'
'                  '«ü¥Ü«H 09~15
'                  stET03 = frm060316_1.GetET03(m_PA01 & m_PA02 & m_PA03 & m_PA04)
'                  frm060316_1.StartLetter "04", m_CP43, stET03, m_PA01 & m_PA02 & m_PA03 & m_PA04, ChgSQL(Me.Combo1.Text), "98"
'                  'Add by Morgan 2008/3/24 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
'                  If m_bolEmail Then
'                     NowPrint m_CP43, "04", stET03, False, strUserNum, , , , , m_iCopy
'                     '¦]¬°­n¦LEMail¸ê®Æ,©Ò¥HSave2File°Ñ¼Æ¤]­n³]
'                     NowPrint m_CP43, "04", stET03, False, strUserNum, , , True, stContent, , , , True
'                  Else
'                  'End 2008/3/24
'                     'Add by Morgan 2006/2/13
'                     '­^¤å©w½Z¥[¶Ç¯u«Ê­±
'                     NowPrint m_CP43, "04", "98", False, strUserNum, 0, , , , 1
'                     NowPrint m_CP43, "04", stET03, False, strUserNum, 0
'                  End If
'
'                  'ªþ¥ó
'                  '­Y¦³¤T­ÓÀu¥ýÅv¸ê®Æ
'                  If m_bln3PriData = True Then
'                     stET03 = "15"
'                  '­Y¦³Àu¥ýÅv¸ê®Æ
'                  ElseIf m_blnPriData = True Then
'                     stET03 = "13"
'                  '­YµLÀu¥ýÅv¸ê®Æ
'                  Else
'                     stET03 = "14"
'                  End If
'                  'Add by Morgan 2008/3/24
'                  If m_bolEmail Then
'                     NowPrint m_CP43, "04", stET03, False, strUserNum, , , , , m_iCopy
'                     NowPrint m_CP43, "04", stET03, False, strUserNum, , stContent, , , , , True, True
'                  Else
'                  'end 2008/3/24
'                     NowPrint m_CP43, "04", stET03, False, strUserNum, 0
'                  End If
'
'               'ÂÂªk
'               Else
'                  '­Y¦³¤T­ÓÀu¥ýÅv¸ê®Æ
'                  If m_bln3PriData = True Then
'                     'Add by Morgan 2008/3/24 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
'                     If m_bolEmail Then
'                        NowPrint m_CP43, "04", "05", False, strUserNum, , , , , m_iCopy
'                        NowPrint m_CP43, "04", "05", False, strUserNum, , , True, stContent, , , , True
'                        NowPrint m_CP43, "04", "08", False, strUserNum, , , , , m_iCopy
'                        NowPrint m_CP43, "04", "08", False, strUserNum, , stContent, , , , , True, True
'                     Else
'                     'End 2008/3/24
'                        NowPrint m_CP43, "04", "05", False, strUserNum, 0
'                        'ªþ¥ó
'                        NowPrint m_CP43, "04", "08", False, strUserNum, 0
'                     End If
'                   '­Y¦³Àu¥ýÅv¸ê®Æ
'                   ElseIf m_blnPriData = True Then
'                     'Add by Morgan 2008/3/24 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
'                     If m_bolEmail Then
'                        NowPrint m_CP43, "04", "02", False, strUserNum, , , , , m_iCopy
'                        NowPrint m_CP43, "04", "02", False, strUserNum, , , True, stContent, , , , True
'                        NowPrint m_CP43, "04", "06", False, strUserNum, , , , , m_iCopy
'                        NowPrint m_CP43, "04", "06", False, strUserNum, , stContent, , , , , True, True
'                     Else
'                     'End 2008/3/24
'                        NowPrint m_CP43, "04", "02", False, strUserNum, 0
'                        'ªþ¥ó
'                        NowPrint m_CP43, "04", "06", False, strUserNum, 0
'                     End If
'                   '­YµLÀu¥ýÅv¸ê®Æ
'                   Else
'                     'Add by Morgan 2008/3/24 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
'                     If m_bolEmail Then
'                        NowPrint m_CP43, "04", "04", False, strUserNum, , , , , m_iCopy
'                        NowPrint m_CP43, "04", "04", False, strUserNum, , , True, stContent, , , , True
'                        NowPrint m_CP43, "04", "07", False, strUserNum, , , , , m_iCopy
'                        NowPrint m_CP43, "04", "07", False, strUserNum, , stContent, , , , , True, True
'                     Else
'                     'End 2008/3/24
'                        NowPrint m_CP43, "04", "04", False, strUserNum, 0
'                        'ªþ¥ó
'                        NowPrint m_CP43, "04", "07", False, strUserNum, 0
'                     End If
'                   End If
'               End If
'            ' ¤é¤å
'            Case "3":
'               stET03 = "03"
'               'Add by Morgan 2004/12/22
'               '·Ó­^¤å§ìªk§PÂ_¬O§_¦Û°Ê¥NÃº
'               stET03 = frm060316_1.GetET03(m_PA01 & m_PA02 & m_PA03 & m_PA04, "3")
'               frm060316_1.StartLetter "04", m_CP43, stET03, m_PA01 & m_PA02 & m_PA03 & m_PA04, ChgSQL(Me.Combo1.Text), "98"
'               '2004/12/22 end
'
'               'Add by Morgan 2008/3/24 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
'               If m_bolEmail Then
'                  NowPrint m_CP43, "04", stET03, False, strUserNum, , , , , m_iCopy
'                  NowPrint m_CP43, "04", stET03, False, strUserNum, , , True, stContent, , , , True
'                  'Add by Morgan 2004/10/13 Ä¶¤å
'                  If m_blnPriData = True Then
'                     stET03 = "18"
'                  Else
'                     stET03 = "17"
'                  End If
'                  NowPrint m_CP43, "04", stET03, False, strUserNum, , , , , m_iCopy
'                  NowPrint m_CP43, "04", stET03, False, strUserNum, , stContent, , , , , True, True
'               Else
'               'End 2008/3/24
'                  'Add by Morgan 2006/3/15
'                  '¥[­^¤å¶Ç¯u«Ê­±
'                  NowPrint m_CP43, "04", "98", False, strUserNum, 0, , , , 1
'                  '2006/3/15 end
'                  NowPrint m_CP43, "04", stET03, False, strUserNum, 0
'                  'Add by Morgan 2004/10/13 Ä¶¤å
'                  If m_blnPriData = True Then
'                     stET03 = "18"
'                  Else
'                     stET03 = "17"
'                  End If
'                  NowPrint m_CP43, "04", stET03, False, strUserNum, 0
'               End If
'
'            Case Else:
'         End Select
'
'         If m_bolEmail Then
'            MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_PA01) & " ]¡I"
'         End If
'
'         'Add by Morgan 2011/6/24
'         '¦C¦L³qª¾¨ç
'         PUB_PrintLetter m_CP43
'         PUB_SetOsDefaultPrinter pub_OsPrinter
'         'end 2011/6/24
'end 2022/11/23
         
      Case "frm060317_1"
'Modified by Morgan 2014/12/10
         frm060317_1.PrintLetter m_PA01, m_PA02, m_PA03, m_PA04, m_PA08, m_PA14, m_PA22, Me.Combo1.Text
'         strKey = m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603"
'         'add by sonia 2014/4/11
'         m_bolEmail = PUB_GetEMailFlag(m_PA01 & m_PA02 & m_PA03 & m_PA04, , , m_bolPlusPaper)
'         stET03 = ""
'         '2014/4/11 END
'         ' ©w½Z»y¤å
'         Select Case m_LetterLanguage
'            ' ¤¤¤å
'            Case "1":
'               'add by sonia 2014/4/24 «De¤Æ¥[¶Ç¯u«Ê­±
'               If Not m_bolEmail Then
'                  frm060317_1.StartLetter "07", strKey, "01"  '§ì¶Ç¯u­¶¼Æ
'                  NowPrint strKey, "07", "98", False, strUserNum, , , , , 1
'               End If
'               '2014/4/24 end
'
'               ' ¦C¦L©w½Z
'               NowPrint strKey, "07", "01", False, strUserNum, 0
'               ' ´Á­­ªí
'               NowPrint strKey, "07", "11", False, strUserNum, 0
'            ' ­^¤å
'            Case "2":
'               Select Case m_LetterKind
'                  Case 1:
'                     'Removed by Morgan 2013/7/19 ¤wµL¾A¥Î®×¥ó,ÂÂ©w½Z§R°£
'                     ''Modify by Morgan 2004/7/20
'                     ''±±¨î93.7.1¥H«á¥Î·s©w½Z
'                     'If Val(m_PA14) > 0 And Val(m_PA14) < 20040701 Then
'                     '   ' ¦C¦L©w½Z
'                     '   NowPrint strKey, "07", "02", False, strUserNum, 0
'                     '   ' Ä¶¤å
'                     '   NowPrint strKey, "07", "10", False, strUserNum, 0
'                     'end 2013/7/19
'
'                     'add by sonia 2014/4/11 «De¤Æ¥[¶Ç¯u«Ê­±
'                     If Not m_bolEmail Then
'                        frm060317_1.StartLetter "07", strKey, "13"  '§ì¶Ç¯u­¶¼Æ
'                        NowPrint strKey, "07", "98", False, strUserNum, , , , , 1
'                     End If
'                     '2014/4/11 end
'
'                     '·s«¬
'                     If m_PA08 = "2" Then
'                        NowPrint strKey, "07", "14", False, strUserNum, 0
'                        NowPrint strKey, "07", "16", False, strUserNum, 0
'                        stET03 = "14"
'                     Else
'                        NowPrint strKey, "07", "13", False, strUserNum, 0
'                        NowPrint strKey, "07", "15", False, strUserNum, 0
'                        stET03 = "13"
'                     End If
'                     ' ´Á­­ªí
'                     NowPrint strKey, "07", "12", False, strUserNum, 0
'
'                 Case 2:
'                     'Removed by Morgan 2013/7/19 ¤wµL¾A¥Î®×¥ó,ÂÂ©w½Z§R°£
'                     ''Modify by Morgan 2004/7/20
'                     '±±¨î93.7.1¥H«á¥Î·s©w½Z
'                     ''If Val(m_PA14) > 0 And Val(m_PA14) < 20040701 Then
'                     '   ' ¦C¦L©w½Z
'                     '   NowPrint strKey, "07", "03", False, strUserNum, 0
'                     '   ' Ä¶¤å
'                     '   NowPrint strKey, "07", "10", False, strUserNum, 0
'                     'end 2013/7/19
'
'                     'add by sonia 2014/4/11 «De¤Æ¥[¶Ç¯u«Ê­±
'                     If Not m_bolEmail Then
'                        frm060317_1.StartLetter "07", strKey, "17"  '§ì¶Ç¯u­¶¼Æ
'                        NowPrint strKey, "07", "98", False, strUserNum, , , , , 1
'                     End If
'                     '2014/4/11 end
'
'                     '·s«¬
'                     If m_PA08 = "2" Then
'                        NowPrint strKey, "07", "18", False, strUserNum, 0
'                        NowPrint strKey, "07", "16", False, strUserNum, 0
'                        stET03 = "18"
'                     Else
'                        NowPrint strKey, "07", "17", False, strUserNum, 0
'                        NowPrint strKey, "07", "15", False, strUserNum, 0
'                        stET03 = "17"
'                     End If
'                     ' ´Á­­ªí
'                     NowPrint strKey, "07", "12", False, strUserNum, 0
'
''Modify by Morgan 2005/1/26 ¤£ºÞ¬O§_¦³¤U¦¸Ãº¶O¤é²Î¤@¥X·s©w½Z--David
'                  Case 3, 5
'                        If m_PA08 = "3" Then
'                           'add by sonia 2014/4/11 «De¤Æ¥[¶Ç¯u«Ê­±
'                           If Not m_bolEmail Then
'                              frm060317_1.StartLetter "07", strKey, "19"  '§ì¶Ç¯u­¶¼Æ
'                              NowPrint strKey, "07", "98", False, strUserNum, , , , , 1
'                           End If
'                           '2014/4/11 end
'
'                           NowPrint strKey, "07", "19", False, strUserNum, 0
'                           stET03 = "19"
'
'                        'Removed by Morgan 2013/7/19 ¤wµL°l¥[®×,©w½Z§R°£
'                        'Else
'                        '   NowPrint strKey, "07", "21", False, strUserNum, 0
'                        'end 2013/7/19
'
'                        End If
'                        ' Ä¶¤å
'                        NowPrint strKey, "07", "20", False, strUserNum, 0
''                     End If
''2005/1/26 end
'
'                 Case 4: 'µL¤U¦¸Ãº¶O¤é(¤@¯ë©Î»âÃÒ¦Û°Ê¥NÃº)
'                     'add by sonia 2014/4/11 «De¤Æ¥[¶Ç¯u«Ê­±
'                     If Not m_bolEmail Then
'                        frm060317_1.StartLetter "07", strKey, "05"  '§ì¶Ç¯u­¶¼Æ
'                        NowPrint strKey, "07", "98", False, strUserNum, , , , , 1
'                     End If
'                     '2014/4/11 end
'
'                     ' ¦C¦L©w½Z
'                     NowPrint strKey, "07", "05", False, strUserNum, 0
'                     stET03 = "05"
'
'                     'Removed by Morgan 2013/7/19 ¤wµL¾A¥Î®×¥ó,ÂÂ©w½Z§R°£
'                     ''Modify by Morgan 2004/7/20
'                     ''±±¨î93.7.1¥H«á¥Î·s©w½Z
'                     'If Val(m_PA14) > 0 And Val(m_PA14) < 20040701 Then
'                     '   ' Ä¶¤å
'                     '   NowPrint strKey, "07", "10", False, strUserNum, 0
'                     'end 2013/7/19
'
'                     '·s«¬
'                     If m_PA08 = "2" Then
'                        ' Ä¶¤å
'                        NowPrint strKey, "07", "16", False, strUserNum, 0
'                     Else
'                        ' Ä¶¤å
'                        NowPrint strKey, "07", "15", False, strUserNum, 0
'                     End If
'
'                  'Added by Morgan 2012/1/4
'                  Case 6: '¿nÅé¹q¸ô
'                     'add by sonia 2014/4/11 «De¤Æ¥[¶Ç¯u«Ê­±
'                     If Not m_bolEmail Then
'                        frm060317_1.StartLetter "07", strKey, "22  '§ì¶Ç¯u­¶¼Æ"
'                        NowPrint strKey, "07", "98", False, strUserNum, , , , , 1
'                     End If
'                     '2014/4/11 end
'
'                     NowPrint strKey, "07", "22", False, strUserNum, 0
'                     NowPrint strKey, "07", "23", False, strUserNum, 0
'                     stET03 = "22"
'
'               End Select
'
'            'Add by Morgan 2006/7/26
'            Case "3" '¤é¤å
'               'add by sonia 2014/4/24 «De¤Æ¥[¶Ç¯u«Ê­±
'               If Not m_bolEmail Then
'                  frm060317_1.StartLetter "07", strKey, "06"  '§ì¶Ç¯u­¶¼Æ
'                  NowPrint strKey, "07", "98", False, strUserNum, , , , , 1
'               End If
'               '2014/4/24 end
'
'               If m_LetterKind = "2" Then
'                  stET03 = "09"
'               Else
'                  stET03 = "06"
'               End If
'               NowPrint strKey, "07", stET03, False, strUserNum, 0
'               ' Ä¶¤å
'               NowPrint strKey, "07", "07", False, strUserNum, 0
'               ' ´Á­­ªí
'               NowPrint strKey, "07", "08", False, strUserNum, 0
'            Case Else:
'         End Select
'end 2014/12/10
         
      Case Else
   End Select
   
   
End Sub

'Add By Cheng 2003/01/19
'¨ú±o¤½§i¤é
Private Function GetPA14(strPA0104 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetPA14 = ""
If m_LetterKind = "3" Then
   'Modify by Morgan 2010/12/26 ¥Ó½Ð¸¹§ï½X¼Æ
   StrSQLa = "SELECT * FROM PATENT WHERE PA01='" & m_PA01 & "' AND PA11 = ( " & _
            "SELECT SUBSTR(PA11,1,9) FROM PATENT WHERE PA01='" & m_PA01 & "' AND " & _
            "PA02='" & m_PA02 & "' AND PA03='" & m_PA03 & "' AND PA04='" & m_PA04 & "')"
Else
   StrSQLa = "Select * From Patent Where " & ChgPatent(strPA0104)
End If
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetPA14 = "" & rsA("PA14").Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'92.1.17 add by sonia
Private Function GetEngMMDD(ByVal strValue As String) As String
Dim strTmp As String
Dim ii As Integer
Dim arrTmp
   
GetEngMMDD = ""
'­Y¦³¶Ç¤J­È
If strValue <> "" Then
    arrTmp = Split(strValue, "; ")
    For ii = 0 To UBound(arrTmp)
        Select Case Mid(arrTmp(ii), 5, 2)
           Case "01": strTmp = "January "
           Case "02": strTmp = "February "
           Case "03": strTmp = "March "
           Case "04": strTmp = "April "
           Case "05": strTmp = "May "
           Case "06": strTmp = "June "
           Case "07": strTmp = "July "
           Case "08": strTmp = "August "
           Case "09": strTmp = "September "
           Case "10": strTmp = "October "
           Case "11": strTmp = "November "
           Case "12": strTmp = "December "
        End Select
        GetEngMMDD = GetEngMMDD & strTmp & Right(strValue, 2)
    Next ii
Else
   GetEngMMDD = ""
End If
End Function


