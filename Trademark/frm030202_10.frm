VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_10 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "µo¤å(¸É¥¿, ©ñ±ó±M¥ÎÅv)"
   ClientHeight    =   6750
   ClientLeft      =   4880
   ClientTop       =   2270
   ClientWidth     =   9140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9140
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "ÅÜ§ó¨Æ¶µ(&R)"
      Height          =   400
      Left            =   4740
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   21
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   400
      Left            =   6900
      TabIndex        =   23
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5940
      TabIndex        =   22
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   24
      Top             =   0
      Width           =   912
   End
   Begin VB.TextBox textCP08 
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2700
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2415
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5790
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   420
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5790
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   705
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   990
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   705
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   420
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1275
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5790
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1275
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   990
      Width           =   2532
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3405
      Left            =   120
      TabIndex        =   50
      Top             =   3300
      Width           =   8955
      _ExtentX        =   15804
      _ExtentY        =   5997
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "°ò¥»¸ê®Æ"
      TabPicture(0)   =   "frm030202_10.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label22"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label25"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label28"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label36"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label37"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label29"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label39"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblNameAgent"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(12)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label43"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lstNameAgent"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCP64"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textTM58"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "grdList"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textPrint"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCP27"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textDN"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCP18"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textPetition"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP84"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text7"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP113"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP118"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "©ñ±ó±M¥ÎÅv/¥Nªí¤H"
      TabPicture(1)   =   "frm030202_10.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5(12)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5(11)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5(10)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5(9)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label5(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label14(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label18(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label18(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label14(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label5(3)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label5(4)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label5(5)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label5(6)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label5(7)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label5(8)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textTM67"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Combo2(0)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "textTM47"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "textTM48"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "textTM49"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Combo2(2)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "textTM94"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "textTM95"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "textTM96"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Combo2(1)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "textTM50"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "textTM51"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "textTM52"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Combo2(3)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "textTM97"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "textTM98"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "textTM99"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "¥Nªí¤H-2"
      TabPicture(2)   =   "frm030202_10.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5(24)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5(23)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label5(22)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label5(21)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label5(20)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label5(19)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label14(3)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label18(3)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label5(18)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label5(17)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label5(16)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label5(15)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label5(14)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label5(13)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label14(2)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label18(1)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Combo2(4)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "textTM100"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "textTM101"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "textTM102"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Combo2(5)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "textTM103"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "textTM104"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "textTM105"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "TextTM106"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "TextTM107"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "TextTM109"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "TextTM110"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Combo2(7)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Combo2(6)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "TextTM108"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "TextTM111"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).ControlCount=   32
      TabCaption(3)   =   "¥Nªí¤H-3"
      TabPicture(3)   =   "frm030202_10.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TextTM117"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "TextTM114"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Combo2(8)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Combo2(9)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "TextTM116"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "TextTM115"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "TextTM113"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "TextTM112"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label5(30)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label5(29)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label5(28)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label5(27)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label5(26)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label5(25)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label14(4)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Label18(4)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).ControlCount=   16
      Begin VB.TextBox textCP118 
         Height          =   285
         Left            =   4080
         MaxLength       =   1
         TabIndex        =   7
         Top             =   870
         Width           =   375
      End
      Begin VB.TextBox textCP113 
         Height          =   285
         Left            =   5775
         MaxLength       =   4
         TabIndex        =   2
         Top             =   315
         Width           =   600
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   6990
         MaxLength       =   1
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   285
         Left            =   3420
         TabIndex        =   1
         Top             =   315
         Width           =   1425
      End
      Begin VB.TextBox textPetition 
         Height          =   285
         Left            =   3900
         MaxLength       =   1
         TabIndex        =   5
         Top             =   593
         Width           =   492
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   285
         Left            =   5790
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   593
         Width           =   1155
      End
      Begin VB.TextBox textDN 
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   4
         Top             =   593
         Width           =   492
      End
      Begin VB.TextBox textCP27 
         Height          =   285
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   0
         Top             =   315
         Width           =   1092
      End
      Begin VB.TextBox textPrint 
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   6
         Top             =   870
         Width           =   492
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   972
         Left            =   1200
         TabIndex        =   152
         Top             =   1224
         Width           =   7572
         _ExtentX        =   13353
         _ExtentY        =   1711
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin MSForms.TextBox TextTM117 
         Height          =   285
         Left            =   -69570
         TabIndex        =   151
         Top             =   1290
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM114 
         Height          =   285
         Left            =   -73980
         TabIndex        =   150
         Top             =   1290
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   8
         Left            =   -73980
         TabIndex        =   18
         Top             =   390
         Width           =   3195
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5644;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   9
         Left            =   -69570
         TabIndex        =   19
         Top             =   390
         Width           =   3195
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5636;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM116 
         Height          =   285
         Left            =   -69570
         TabIndex        =   149
         Top             =   990
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM115 
         Height          =   285
         Left            =   -69570
         TabIndex        =   148
         Top             =   705
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM113 
         Height          =   285
         Left            =   -73980
         TabIndex        =   147
         Top             =   990
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM112 
         Height          =   285
         Left            =   -73980
         TabIndex        =   146
         Top             =   705
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   30
         Left            =   -70080
         TabIndex        =   145
         Top             =   1335
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   29
         Left            =   -70080
         TabIndex        =   144
         Top             =   1045
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   28
         Left            =   -70080
         TabIndex        =   143
         Top             =   755
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   27
         Left            =   -74490
         TabIndex        =   142
         Top             =   1342
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   26
         Left            =   -74490
         TabIndex        =   141
         Top             =   1049
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   25
         Left            =   -74490
         TabIndex        =   140
         Top             =   757
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H9"
         Height          =   180
         Index           =   4
         Left            =   -74820
         TabIndex        =   139
         Top             =   465
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H10"
         Height          =   180
         Index           =   4
         Left            =   -70410
         TabIndex        =   138
         Top             =   465
         Width           =   720
      End
      Begin MSForms.TextBox TextTM111 
         Height          =   285
         Left            =   -69660
         TabIndex        =   137
         Top             =   2430
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM108 
         Height          =   285
         Left            =   -74070
         TabIndex        =   136
         Top             =   2430
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   6
         Left            =   -74070
         TabIndex        =   16
         Top             =   1525
         Width           =   3195
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5636;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   7
         Left            =   -69660
         TabIndex        =   17
         Top             =   1525
         Width           =   3195
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5636;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM110 
         Height          =   285
         Left            =   -69660
         TabIndex        =   135
         Top             =   2130
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM109 
         Height          =   285
         Left            =   -69660
         TabIndex        =   134
         Top             =   1835
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM107 
         Height          =   285
         Left            =   -74070
         TabIndex        =   133
         Top             =   2130
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM106 
         Height          =   285
         Left            =   -74070
         TabIndex        =   132
         Top             =   1835
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM105 
         Height          =   285
         Left            =   -69660
         TabIndex        =   131
         Top             =   1230
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM104 
         Height          =   285
         Left            =   -69660
         TabIndex        =   130
         Top             =   935
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM103 
         Height          =   285
         Left            =   -69660
         TabIndex        =   129
         Top             =   640
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   5
         Left            =   -69660
         TabIndex        =   15
         Top             =   330
         Width           =   3200
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5644;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM102 
         Height          =   285
         Left            =   -74070
         TabIndex        =   128
         Top             =   1230
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM101 
         Height          =   285
         Left            =   -74070
         TabIndex        =   127
         Top             =   935
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM100 
         Height          =   285
         Left            =   -74070
         TabIndex        =   126
         Top             =   640
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   4
         Left            =   -74070
         TabIndex        =   14
         Top             =   330
         Width           =   3200
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5644;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM99 
         Height          =   285
         Left            =   -69690
         TabIndex        =   125
         Top             =   2760
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM98 
         Height          =   285
         Left            =   -69690
         TabIndex        =   124
         Top             =   2475
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM97 
         Height          =   285
         Left            =   -69690
         TabIndex        =   123
         Top             =   2190
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   -69690
         TabIndex        =   13
         Top             =   1890
         Width           =   3195
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5636;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM52 
         Height          =   285
         Left            =   -69690
         TabIndex        =   122
         Top             =   1605
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM51 
         Height          =   285
         Left            =   -69690
         TabIndex        =   121
         Top             =   1305
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM50 
         Height          =   285
         Left            =   -69690
         TabIndex        =   120
         Top             =   1020
         Width           =   3195
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5636;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   -69690
         TabIndex        =   119
         Top             =   720
         Width           =   3195
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5636;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM96 
         Height          =   285
         Left            =   -73770
         TabIndex        =   118
         Top             =   2760
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM95 
         Height          =   285
         Left            =   -73770
         TabIndex        =   117
         Top             =   2475
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM94 
         Height          =   285
         Left            =   -73770
         TabIndex        =   116
         Top             =   2190
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   -73770
         TabIndex        =   12
         Top             =   1890
         Width           =   3200
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5644;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM49 
         Height          =   285
         Left            =   -73770
         TabIndex        =   115
         Top             =   1605
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM48 
         Height          =   285
         Left            =   -73770
         TabIndex        =   114
         Top             =   1305
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM47 
         Height          =   285
         Left            =   -73770
         TabIndex        =   113
         Top             =   1020
         Width           =   3200
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5644;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   -73770
         TabIndex        =   11
         Top             =   720
         Width           =   3200
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5644;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM67 
         Height          =   285
         Left            =   -73770
         TabIndex        =   10
         Top             =   360
         Width           =   7395
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "13044;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   525
         Left            =   1200
         TabIndex        =   9
         Top             =   2790
         Width           =   7572
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13356;926"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   525
         Left            =   1200
         TabIndex        =   8
         Top             =   2250
         Width           =   7572
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13356;926"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   7350
         TabIndex        =   3
         Top             =   280
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
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¹q¤l°e¥ó:          (Y: ¬O)"
         Height          =   180
         Left            =   2910
         TabIndex        =   103
         Top             =   922
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤u§@®É¼Æ:"
         Height          =   180
         Index           =   12
         Left            =   5010
         TabIndex        =   102
         Top             =   367
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H6"
         Height          =   180
         Index           =   1
         Left            =   -70470
         TabIndex        =   101
         Top             =   390
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H5"
         Height          =   180
         Index           =   2
         Left            =   -74790
         TabIndex        =   100
         Top             =   390
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   13
         Left            =   -74460
         TabIndex        =   99
         Top             =   688
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   14
         Left            =   -74460
         TabIndex        =   98
         Top             =   986
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   15
         Left            =   -74460
         TabIndex        =   97
         Top             =   1284
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   16
         Left            =   -70140
         TabIndex        =   96
         Top             =   687
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   17
         Left            =   -70140
         TabIndex        =   95
         Top             =   984
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   18
         Left            =   -70140
         TabIndex        =   94
         Top             =   1281
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H8"
         Height          =   180
         Index           =   3
         Left            =   -70470
         TabIndex        =   93
         Top             =   1578
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H7"
         Height          =   180
         Index           =   3
         Left            =   -74790
         TabIndex        =   92
         Top             =   1582
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   19
         Left            =   -74460
         TabIndex        =   91
         Top             =   1880
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   20
         Left            =   -74460
         TabIndex        =   90
         Top             =   2178
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   21
         Left            =   -74460
         TabIndex        =   89
         Top             =   2482
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   22
         Left            =   -70140
         TabIndex        =   88
         Top             =   1875
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   23
         Left            =   -70140
         TabIndex        =   87
         Top             =   2172
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   24
         Left            =   -70140
         TabIndex        =   86
         Top             =   2475
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   8
         Left            =   -70110
         TabIndex        =   85
         Top             =   1657
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   7
         Left            =   -70110
         TabIndex        =   84
         Top             =   1357
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   6
         Left            =   -70110
         TabIndex        =   83
         Top             =   1072
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   5
         Left            =   -74190
         TabIndex        =   82
         Top             =   1657
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   4
         Left            =   -74190
         TabIndex        =   81
         Top             =   1357
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   3
         Left            =   -74190
         TabIndex        =   80
         Top             =   1072
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H1"
         Height          =   180
         Index           =   1
         Left            =   -74520
         TabIndex        =   79
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H2"
         Height          =   180
         Index           =   2
         Left            =   -70440
         TabIndex        =   78
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H4"
         Height          =   180
         Index           =   0
         Left            =   -70440
         TabIndex        =   77
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H3"
         Height          =   180
         Index           =   0
         Left            =   -74520
         TabIndex        =   76
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   1
         Left            =   -74190
         TabIndex        =   75
         Top             =   2242
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   2
         Left            =   -74190
         TabIndex        =   74
         Top             =   2527
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   9
         Left            =   -74190
         TabIndex        =   73
         Top             =   2812
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   10
         Left            =   -70110
         TabIndex        =   72
         Top             =   2242
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   11
         Left            =   -70110
         TabIndex        =   71
         Top             =   2527
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   12
         Left            =   -70110
         TabIndex        =   70
         Top             =   2812
         Width           =   345
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "¥X¦W¥N²z¤H"
         Height          =   180
         Left            =   6450
         TabIndex        =   65
         Top             =   367
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "µo¤å³W¶O¡G"
         Height          =   180
         Left            =   2490
         TabIndex        =   64
         Top             =   367
         Width           =   900
      End
      Begin VB.Label Label21 
         Caption         =   "©ñ±ó±M¥ÎÅv :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   63
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "¥»®×´Á­­ :"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1170
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "®×¥ó³Æµù :"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "(Y:¦L)"
         Height          =   255
         Left            =   4410
         TabIndex        =   60
         Top             =   608
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "¬O§_¦C¦L¥Ó½Ð®Ñ :"
         Height          =   255
         Left            =   2460
         TabIndex        =   59
         Top             =   608
         Width           =   1455
      End
      Begin VB.Label Label37 
         Caption         =   "(Y:¿é¤J)"
         Height          =   255
         Left            =   1740
         TabIndex        =   58
         Top             =   608
         Width           =   855
      End
      Begin VB.Label Label36 
         Caption         =   "¬O§_¿é¤JD/N:"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   608
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "¶i«×³Æµù :"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   2250
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "µo¤å¤é :"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "¦C¦L©w½Z :"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "(N:¤£¦L)"
         Height          =   255
         Left            =   1740
         TabIndex        =   53
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "ÂI¡@¼Æ :"
         Height          =   255
         Index           =   10
         Left            =   5010
         TabIndex        =   52
         Top             =   608
         Width           =   900
      End
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5790
      TabIndex        =   112
      Top             =   1560
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
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1230
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   1560
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
   Begin MSForms.TextBox textTM78 
      Height          =   285
      Left            =   5790
      TabIndex        =   110
      Top             =   1845
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1230
      TabIndex        =   109
      Top             =   1845
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
      Left            =   1230
      TabIndex        =   108
      Top             =   2415
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
      Left            =   5790
      TabIndex        =   107
      Top             =   2130
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
      Left            =   1230
      TabIndex        =   106
      Top             =   2130
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   1230
      TabIndex        =   105
      Top             =   2700
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
      Left            =   1200
      TabIndex        =   104
      Top             =   2970
      Width           =   7815
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13785;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H5 :"
      Height          =   180
      Left            =   150
      TabIndex        =   69
      Top             =   2467
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H4 :"
      Height          =   180
      Left            =   4770
      TabIndex        =   68
      Top             =   2182
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H3 :"
      Height          =   180
      Left            =   150
      TabIndex        =   67
      Top             =   2182
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H2 :"
      Height          =   180
      Left            =   4770
      TabIndex        =   66
      Top             =   1897
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¾÷Ãö¤å¸¹ :"
      Height          =   180
      Index           =   5
      Left            =   4770
      TabIndex        =   48
      Top             =   2752
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "¥N²z¤H :"
      Height          =   180
      Left            =   150
      TabIndex        =   47
      Top             =   2752
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H1 :"
      Height          =   180
      Left            =   150
      TabIndex        =   46
      Top             =   1897
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "°Ó¼ÐºØÃþ :"
      Height          =   180
      Index           =   4
      Left            =   4770
      TabIndex        =   45
      Top             =   2467
      Width           =   810
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "©Ó¿ì¤H :"
      Height          =   180
      Left            =   150
      TabIndex        =   44
      Top             =   1612
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "·~°È°Ï§O :"
      Height          =   180
      Index           =   2
      Left            =   4770
      TabIndex        =   43
      Top             =   472
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "µoÃÒ¤é :"
      Height          =   180
      Index           =   3
      Left            =   4770
      TabIndex        =   42
      Top             =   757
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð®×¸¹ :"
      Height          =   180
      Left            =   4770
      TabIndex        =   41
      Top             =   1042
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹ :"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   40
      Top             =   757
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¦¬¤å¸¹ :"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   39
      Top             =   472
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó©Ê½è :"
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   38
      Top             =   1327
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "©¼©Ò®×¸¹ :"
      Height          =   180
      Index           =   9
      Left            =   4770
      TabIndex        =   37
      Top             =   1327
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "´¼Åv¤H­û :"
      Height          =   180
      Index           =   11
      Left            =   4770
      TabIndex        =   36
      Top             =   1612
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¼f©w¸¹¼Æ :"
      Height          =   180
      Left            =   150
      TabIndex        =   35
      Top             =   1042
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó¦WºÙ :"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   34
      Top             =   3022
      Width           =   810
   End
End
Attribute VB_Name = "frm030202_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0¤w­×§ï ; grdList±qMSFlexGrid§ï¬°MSHFlexGrid
'Memo by Lydia 2021/09/09 §ï¦¨Form2.0 ;cmbTM05¡BtextCP13¡BtextCP14¡BtextCP64¡BtextTM58¡BtextTM44¡BtextTM23¡BtextTM78~81¡BtextTM67¡BtextCP50~52¡BlstNameAgent¡BCombo2(index)¡BtextTM47~52¡BtextTM94~117¡FgrdList§ï¦r«¬=·s²Ó©úÅé-ExtB
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
' ¥Ó½Ð°ê®a
Dim m_TM10 As String
' ®×¥ó©Ê½è¥N¸¹
Dim m_CP10 As String
'©Ó¿ì¤H Add By Sindy 98/03/11
Dim m_CP14 As String
Dim m_CP82 As String 'Added by Lydia 2018/08/10 µo¤å®É¶¡

Dim m_CurrSel As Integer
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
Dim m_strLanguage As String '©w½Z»y¤å
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
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 ¦¬¤å¸¹,¬O§_ºâµo¤å«Ç®×¥ó
Dim m_CP130s As String 'Add by Sindy 2009/4/24 µo¤å-¥DºÞ¾÷Ãö
'Add By Sindy 2009/06/03
Dim m_TM23 As String
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
'2009/06/03 End
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim m_CP43 As String 'Added by Lydia 2016/05/09
Dim m_bolMsg208 As Boolean  'Added by Lydia 2020/11/10 ¬O§_¤@¨Öµo¤å¸ÉÀu¥ýÅv208
Dim m_str208 As String 'Added by Lydia 2021/08/13 Àu¥ýÅv¤§¦¬¤å¸¹

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
   frm030202_05.SetParent "frm030202_10"
   Me.Hide
   frm030202_05.Show
   frm030202_05.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
      If TxtValidate = False Then Exit Sub
      
      'Add by Sindy 98/3/24 ³]©w¬O§_ºâµo¤å«Ç®×¥ó
      If m_TM10 = "000" Then
         'Modify By Sindy 2012/12/20 ­Y¬°¹q¤l°e¥ó«h¤£¸gµo¤å«Ç
         'Modify By Sindy 2023/8/1 ¹q¤l°e¥óÄæ¦ì­È¤£¬OªÅ¥ÕªÌ,§Y¬°¹q¤l°e¥ó
         If (textCP118.Visible = True And textCP118 <> "") Then
            'Added by Morgan 2016/5/16 ¹q¤l°e¥ó¤]­n°O¿ý¥DºÞ¾÷Ãö
            If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
               Exit Sub
            End If
            'end 2016/5/16
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
        'Added by Lydia 2021/08/13 ¦P®Éµo¤å¸ÉÀu¥ýÅv208¡A¦P®É¤W¶ÇÀÉ®×¨ì¨÷©v°Ï
        If m_str208 <> "" Then
            If Pub_AutoSavePdf_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_str208, "208") = False Then
            End If
        End If
        'end 2021/08/13
        
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
      
      Call PUB_FCTSendRecvMail(m_CP09) 'Add By Sindy 2024/10/30 ¥~°Óµo¤å®É,¼W¥[µoMail³qª¾©Ó¿ì¤H¤Î°Æ¥»µ¹§Pµo¥DºÞ
      'Add By Sindy 2024/8/19
      If frm030202_01.bolIsEMPFlow = True Then
         frm090202_4.QueryData
      End If
      '2024/8/19 End
      'Ken 91.04.09 -- Start
      If textDN = "Y" Then
        'Add By Cheng 2003/03/19
        '·s¼W¦a§}±ø¦Cªí¸ê®Æ
'edit by nick 2004/11/17  ¦]¬°½Ð´Ú¤w¸g¦³²£¥Í¤F
'        pub_AddressListSN = pub_AddressListSN + 1
'        PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
         Screen.MousePointer = vbHourglass
         Frmacc21h0.Show
         mdiMain.ToolShow
         mdiMain.tool1_enabled
         Screen.MousePointer = vbDefault
         Set Frmacc21h0.frmlink = frm030202_01
         'add by nick 2004/11/24
         Frmacc21h0.IsPrintAddress = False
      Else
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
   End If
End Sub

'add by nickc 2008/03/12 ³¯ª÷³s½Ð§@³æ
Private Sub Combo2_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   
   If (Combo2(Index).Text = "") Then
      For i = 0 To 2
         Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
      Next i
      Exit Sub
   End If
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For i = 0 To 2
         If Not IsNull(RsTemp.Fields(i)) Then
            Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = RsTemp.Fields(i)
         'add by nickc 2008/04/08 ­×¥¿¿ù»~
         Else
            Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
         End If
      Next
   End If
End Sub

'Private Sub Form_Activate()
'    'Add By Cheng 2003/10/06
'    '­Y¦³«ö¤UÅÜ§ó¨Æ¶µ«ö¶s, «h­«·sÅª¨ú¸ê®Æ
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
   
   textTM44.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP08.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   
   'Added by Lydia 2016/09/10 ³]©w¥Nªí¤H¤¤¤å¦WºÙ©M­^¤å¦WºÙªø«×
    textTM47.MaxLength = Pub_MaxCEL10
    textTM48.MaxLength = Pub_MaxCEL11
    textTM50.MaxLength = Pub_MaxCEL10
    textTM51.MaxLength = Pub_MaxCEL11
    textTM94.MaxLength = Pub_MaxCEL10
    textTM95.MaxLength = Pub_MaxCEL11
    textTM97.MaxLength = Pub_MaxCEL10
    textTM98.MaxLength = Pub_MaxCEL11
    textTM100.MaxLength = Pub_MaxCEL10
    textTM101.MaxLength = Pub_MaxCEL11
    textTM103.MaxLength = Pub_MaxCEL10
    textTM104.MaxLength = Pub_MaxCEL11
    TextTM106.MaxLength = Pub_MaxCEL10
    TextTM107.MaxLength = Pub_MaxCEL11
    TextTM109.MaxLength = Pub_MaxCEL10
    TextTM110.MaxLength = Pub_MaxCEL11
    TextTM112.MaxLength = Pub_MaxCEL10
    TextTM113.MaxLength = Pub_MaxCEL11
    TextTM115.MaxLength = Pub_MaxCEL10
    TextTM116.MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   MoveFormToCenter Me
'    m_blnClkChgButton = False
   'Add by nickc 2006/01/26
   '¥xÆW¥[¥X¦W¥N²z¤H²M³æ¨Ñ¤Ä¿ï,­ì¬O§_¥X¦WÄæ¦ì¤£Åã¥Ü
   Text7.Visible = False
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Added by Lydia 2021/09/09 ¦pªG¤@¶}©l±NListBox©Ô¨ì»Ý­nªº¤j¤p¡A¦r«¬·|¦Û°Ê©ñ¤j¡F©Ò¥Hµe­±¹w³]¬°¤@¦C°ª«×¡AForm_Load¤~©ñ¤j¨ì»Ý­nªº¤j¤p
   lstNameAgent.Height = 855
   lstNameAgent.Width = 1500
   Me.SSTab1.Tab = 0
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
      
      ' FC¥N²z¤H
      If IsNull(rsTmp.Fields("TM44")) = False Then: textTM44 = GetFAgentName(rsTmp.Fields("TM44"))
      ' ©¼©Ò®×¸¹
      If IsNull(rsTmp.Fields("TM45")) = False Then: textTM45 = rsTmp.Fields("TM45")
      ' ¥Nªí¤H1(¤¤)
      If IsNull(rsTmp.Fields("TM47")) = False Then: textTM47 = rsTmp.Fields("TM47")
      SetTMSPFieldOldData "TM47", textTM47, 0
      ' ¥Nªí¤H1(­^)
      If IsNull(rsTmp.Fields("TM48")) = False Then: textTM48 = rsTmp.Fields("TM48")
      SetTMSPFieldOldData "TM48", textTM48, 0
      ' ¥Nªí¤H1(¤é)
      If IsNull(rsTmp.Fields("TM49")) = False Then: textTM49 = rsTmp.Fields("TM49")
      SetTMSPFieldOldData "TM49", textTM49, 0
      ' ¥Nªí¤H2(¤¤)
      If IsNull(rsTmp.Fields("TM50")) = False Then: textTM50 = rsTmp.Fields("TM50")
      SetTMSPFieldOldData "TM50", textTM50, 0
      ' ¥Nªí¤H2(­^)
      If IsNull(rsTmp.Fields("TM51")) = False Then: textTM51 = rsTmp.Fields("TM51")
      SetTMSPFieldOldData "TM51", textTM51, 0
      ' ¥Nªí¤H2(¤é)
      If IsNull(rsTmp.Fields("TM52")) = False Then: textTM52 = rsTmp.Fields("TM52")
      SetTMSPFieldOldData "TM52", textTM52, 0
      ' ®×¥ó³Æµù
      If IsNull(rsTmp.Fields("TM58")) = False Then: textTM58 = rsTmp.Fields("TM58")
      SetTMSPFieldOldData "TM58", textTM58, 0
      ' ©ñ±ó±M¥ÎÅv
      If IsNull(rsTmp.Fields("TM67")) = False Then: textTM67 = rsTmp.Fields("TM67")
      SetTMSPFieldOldData "TM67", textTM67, 0
      'add by nickc 2007/02/13 ¥[¥Nªí¤H 3-10
      If IsNull(rsTmp.Fields("TM94")) = False Then: textTM94 = rsTmp.Fields("TM94")
      SetTMSPFieldOldData "TM94", textTM94, 0
      If IsNull(rsTmp.Fields("TM95")) = False Then: textTM95 = rsTmp.Fields("TM95")
      SetTMSPFieldOldData "TM95", textTM95, 0
      If IsNull(rsTmp.Fields("TM96")) = False Then: textTM96 = rsTmp.Fields("TM96")
      SetTMSPFieldOldData "TM96", textTM96, 0
      If IsNull(rsTmp.Fields("TM97")) = False Then: textTM97 = rsTmp.Fields("TM97")
      SetTMSPFieldOldData "TM97", textTM97, 0
      If IsNull(rsTmp.Fields("TM98")) = False Then: textTM98 = rsTmp.Fields("TM98")
      SetTMSPFieldOldData "TM98", textTM98, 0
      If IsNull(rsTmp.Fields("TM99")) = False Then: textTM99 = rsTmp.Fields("TM99")
      SetTMSPFieldOldData "TM99", textTM99, 0
      If IsNull(rsTmp.Fields("TM100")) = False Then: textTM100 = rsTmp.Fields("TM100")
      SetTMSPFieldOldData "TM100", textTM100, 0
      If IsNull(rsTmp.Fields("TM101")) = False Then: textTM101 = rsTmp.Fields("TM101")
      SetTMSPFieldOldData "TM101", textTM101, 0
      If IsNull(rsTmp.Fields("TM102")) = False Then: textTM102 = rsTmp.Fields("TM102")
      SetTMSPFieldOldData "TM102", textTM102, 0
      If IsNull(rsTmp.Fields("TM103")) = False Then: textTM103 = rsTmp.Fields("TM103")
      SetTMSPFieldOldData "TM103", textTM103, 0
      If IsNull(rsTmp.Fields("TM104")) = False Then: textTM104 = rsTmp.Fields("TM104")
      SetTMSPFieldOldData "TM104", textTM104, 0
      If IsNull(rsTmp.Fields("TM105")) = False Then: textTM105 = rsTmp.Fields("TM105")
      SetTMSPFieldOldData "TM105", textTM105, 0
      If IsNull(rsTmp.Fields("TM106")) = False Then: TextTM106 = rsTmp.Fields("TM106")
      SetTMSPFieldOldData "TM106", TextTM106, 0
      If IsNull(rsTmp.Fields("TM107")) = False Then: TextTM107 = rsTmp.Fields("TM107")
      SetTMSPFieldOldData "TM107", TextTM107, 0
      If IsNull(rsTmp.Fields("TM108")) = False Then: TextTM108 = rsTmp.Fields("TM108")
      SetTMSPFieldOldData "TM108", TextTM108, 0
      If IsNull(rsTmp.Fields("TM109")) = False Then: TextTM109 = rsTmp.Fields("TM109")
      SetTMSPFieldOldData "TM109", TextTM109, 0
      If IsNull(rsTmp.Fields("TM110")) = False Then: TextTM110 = rsTmp.Fields("TM110")
      SetTMSPFieldOldData "TM110", TextTM110, 0
      If IsNull(rsTmp.Fields("TM111")) = False Then: TextTM111 = rsTmp.Fields("TM111")
      SetTMSPFieldOldData "TM111", TextTM111, 0
      If IsNull(rsTmp.Fields("TM112")) = False Then: TextTM112 = rsTmp.Fields("TM112")
      SetTMSPFieldOldData "TM112", TextTM112, 0
      If IsNull(rsTmp.Fields("TM113")) = False Then: TextTM113 = rsTmp.Fields("TM113")
      SetTMSPFieldOldData "TM113", TextTM113, 0
      If IsNull(rsTmp.Fields("TM114")) = False Then: TextTM114 = rsTmp.Fields("TM114")
      SetTMSPFieldOldData "TM114", TextTM114, 0
      If IsNull(rsTmp.Fields("TM115")) = False Then: TextTM115 = rsTmp.Fields("TM115")
      SetTMSPFieldOldData "TM115", TextTM115, 0
      If IsNull(rsTmp.Fields("TM116")) = False Then: TextTM116 = rsTmp.Fields("TM116")
      SetTMSPFieldOldData "TM116", TextTM116, 0
      If IsNull(rsTmp.Fields("TM117")) = False Then: TextTM117 = rsTmp.Fields("TM117")
      SetTMSPFieldOldData "TM117", TextTM117, 0
      
      'add by nickc 2008/03/12 ³¯ª÷½¬ ½Ð§@³æ ¥Nªí¤H
      Dim i As Integer, j As Integer
      For i = 0 To 9
         Combo2(i).AddItem ""
      Next
      
      If rsTmp.Fields("TM23").Value <> "" Then
         'edit by nickc 2008/03/24 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM23").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM23").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(0).AddItem rsTmp.Fields("TM23").Value & "-" & j & strExc(0)
               Combo2(1).AddItem rsTmp.Fields("TM23").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("TM78").Value <> "" Then
         'edit by nickc 2008/03/24 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM78").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM78").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(2).AddItem rsTmp.Fields("TM78").Value & "-" & j & strExc(0)
               Combo2(3).AddItem rsTmp.Fields("TM78").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("TM79").Value <> "" Then
         'edit by nickc 2008/03/24 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM79").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM79").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(4).AddItem rsTmp.Fields("TM79").Value & "-" & j & strExc(0)
               Combo2(5).AddItem rsTmp.Fields("TM79").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("TM80").Value <> "" Then
         'edit by nickc 2008/03/24 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM80").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM80").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(6).AddItem rsTmp.Fields("TM80").Value & "-" & j & strExc(0)
               Combo2(7).AddItem rsTmp.Fields("TM80").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("TM81").Value <> "" Then
         'edit by nickc 2008/03/24 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM81").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM81").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(8).AddItem rsTmp.Fields("TM81").Value & "-" & j & strExc(0)
               Combo2(9).AddItem rsTmp.Fields("TM81").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' ¨ú±o®×¥ó¶i«×ÀÉªºÄæ¦ì¤º®e
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSubTmp As New ADODB.Recordset
   Dim strDate As String
   Dim strCP27 As String
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
      ' ¾÷Ãö¤å¸¹
      If IsNull(rsTmp.Fields("CP08")) = False Then
         textCP08 = rsTmp.Fields("CP08")
      End If
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
      '©Ó¿ì¤H
      m_CP14 = "" & rsTmp.Fields("CP14")
      '98/03/11 End
      
      ' ©Ó¿ì¤H­û
      If IsNull(rsTmp.Fields("CP14")) = False Then: textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      ' µo¤å¤é(¹w³]¬°¨t²Î¤é)
      strCP27 = Empty
      'Modify By Sindy 2010/01/22 §PÂ_µo¤å¤é¬°ªÅ¥Õ¤~¹w³]¬°¨t²Î¤é
      If textCP27 = "" Then
         textCP27 = TAIWANDATE(strDate)
      End If
      If IsNull(rsTmp.Fields("CP27")) = False Then: strCP27 = rsTmp.Fields("CP27")
      SetCPFieldOldData "CP27", strCP27, 1
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
   m_CP43 = "" & rsTmp.Fields("cp43") 'Added by Lydia 2016/05/09
   
   rsTmp.Close
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
   
   'Åª¨ú°Ó¼Ð°ò¥»ÀÉ
   QueryTradeMark
   
   ' Åª¨ú®×¥ó¶i«×ÀÉ
   QueryCaseProgress
   'Modified by Lydia 2021/09/09 + Form 2.0 = True
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True 'Modify By Sindy 2010/9/20
   
   ' ¥»®×´Á­­
   InitialGrdList
   ' ¨ú±o¤U¤@µ{§ÇÀÉ®×¤¤ªº¸ê®Æ¦Cªí¦b Grid List ¤¤
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' ¬O§_Äò¿ìÄæ¦ì¥²¶·¬°ªÅ¥Õ
         If IsNull(rsTmp.Fields("NP06")) = False Then
            If IsEmptyText(rsTmp.Fields("NP06")) = False Then
               GoTo NextRecord
            End If
         End If
         
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         
         ' ¦¬¤å¸¹
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("NP01")
         End If
         ' ¤U¤@µ{§Ç
         If IsNull(rsTmp.Fields("NP07")) = False Then
            grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"), 0)
            grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("NP07")
         End If
         ' ¥»©Ò´Á­­
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(grdList.row, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' ªk©w´Á­­
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(grdList.row, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' ¾÷Ãö¤å¸¹
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("NP13")
         End If
         ' ¬ÛÃö¤H
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("NP14")
         End If
         ' ³Æµù
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("NP15")
         End If
         ' §Ç¸¹
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("NP22")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/17
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/17
   End If
   rsTmp.Close
   
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

' ªì©l¤Æ GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 10
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "¤U¤@µ{§Ç"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "¥»©Ò´Á­­"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "ªk©w´Á­­"
   grdList.ColWidth(3) = 1000
   grdList.col = 4
   grdList.Text = "¾÷Ãö¤å¸¹"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "¬ÛÃö¤H"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "³Æµù"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "¦¬¤å¸¹"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "¤U¤@µ{§Ç¥N¸¹"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "§Ç¸¹"
   grdList.ColWidth(9) = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030202_10 = Nothing
End Sub

Private Sub grdList_Click()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            grdList.TextMatrix(grdList.row, 0) = Empty
         Else
            grdList.TextMatrix(grdList.row, 0) = "V"
         End If
      End If
   End If
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' ±NGridList©Ò¿ï¨úªº¦C¤Ï¥Õ, ¨Ã±N¥¼¿ï¨úªº¦C³]¦¨¤@¯ëÃC¦â
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' »P«e¤@¿ï¾Üªº¦C¦ì¸m¬Û¦P«h¤£³B²z
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
      GoTo EXITSUB
   End If
   
   ' ±N­ì¥ý¿ï¨úªº¦C¦^´_¨ì¥¿±`ªºÃC¦â
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' ³]©w¦¨©Ò¿ï¨úªº¦C
   m_CurrSel = nCurrSel
   ' ±N©Ò¿ï¨úªº¦C¤Ï¥Õ
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
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
         'Modified by Lydia 2021/09/09 §ï¼Ò²Õ
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
   End If
EXITSUB:
End Sub

'edit by nickc 2006/10/26
'Private Sub textCP64_2_GotFocus()
'   TextInverse textCP64_2
'End Sub

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

Private Sub textPetition_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¬O§_¦C¦L¥Ó½Ð®Ñ
Private Sub textPetition_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPetition) = False Then
      Select Case textPetition
         Case " ", "Y":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎY"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPetition_GotFocus
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
   ' ¥Nªí¤H1(¤¤)
   SetTMSPFieldNewData "TM47", textTM47
   ' ¥Nªí¤H1(­^)
   SetTMSPFieldNewData "TM48", textTM48
   ' ¥Nªí¤H1(¤é)
   SetTMSPFieldNewData "TM49", textTM49
   ' ¥Nªí¤H2(¤¤)
   SetTMSPFieldNewData "TM50", textTM50
   ' ¥Nªí¤H2(­^)
   SetTMSPFieldNewData "TM51", textTM51
   ' ¥Nªí¤H2(¤é)
   SetTMSPFieldNewData "TM52", textTM52
   ' ®×¥ó³Æµù
   SetTMSPFieldNewData "TM58", textTM58
   ' ©ñ±ó±M¥ÎÅv
   SetTMSPFieldNewData "TM67", textTM67
   'add by nickc 2007/02/13
   SetTMSPFieldNewData "TM94", textTM94
   SetTMSPFieldNewData "TM95", textTM95
   SetTMSPFieldNewData "TM96", textTM96
   SetTMSPFieldNewData "TM97", textTM97
   SetTMSPFieldNewData "TM98", textTM98
   SetTMSPFieldNewData "TM99", textTM99
   SetTMSPFieldNewData "TM100", textTM100
   SetTMSPFieldNewData "TM101", textTM101
   SetTMSPFieldNewData "TM102", textTM102
   SetTMSPFieldNewData "TM103", textTM103
   SetTMSPFieldNewData "TM104", textTM104
   SetTMSPFieldNewData "TM105", textTM105
   SetTMSPFieldNewData "TM106", TextTM106
   SetTMSPFieldNewData "TM107", TextTM107
   SetTMSPFieldNewData "TM108", TextTM108
   SetTMSPFieldNewData "TM109", TextTM109
   SetTMSPFieldNewData "TM110", TextTM110
   SetTMSPFieldNewData "TM111", TextTM111
   SetTMSPFieldNewData "TM112", TextTM112
   SetTMSPFieldNewData "TM113", TextTM113
   SetTMSPFieldNewData "TM114", TextTM114
   SetTMSPFieldNewData "TM115", TextTM115
   SetTMSPFieldNewData "TM116", TextTM116
   SetTMSPFieldNewData "TM117", TextTM117
   
   ' Add By Sindy 98/03/11
   SetCPFieldNewData "CP113", textCP113
   ' µo¤å¤é
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' ¶i«×³Æµù
   '910801 Sieg 602
'edit by nickc 2006/01/26
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
   
   'Add By Sindy 2012/12/20
   ' ¬O§_¹q¤l°e¥ó
   SetCPFieldNewData "CP118", textCP118
End Sub

' §ó·s°Ó¼Ð°ò¥»ÀÉªº¬ÛÃöÄæ¦ì
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' §ó·s®×¥ó¶i«×ÀÉ
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (³æ¤Þ¸¹)
            'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
            strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' °õ¦æSQL«ü¥O
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

' §ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÄæ¦ì
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
Dim nIndex As Integer
Dim strSql As String
Dim strNP07 As String
Dim strNP08 As String
Dim strNP22 As String
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/9/10

'911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' §ó·s°Ó¼Ð°ò¥»ÀÉ
   OnUpdateTradeMark
   ' §ó·s®×¥ó¶i«×ÀÉ
   OnUpdateCaseProperty
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' §ó·s¨Ï¥ÎªÌ©Ò¿ï¨úªº¥»®×´Á­­¸ê®Æ
   For nIndex = 1 To grdList.Rows - 1
      ' §PÂ_¸Ó¦C¬O§_¦³³Q¿ï¨ú
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = grdList.TextMatrix(nIndex, 8)
         strNP22 = grdList.TextMatrix(nIndex, 9)

         strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND " & _
                        "NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND " & _
                        "NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         cnnConnection.Execute strSql
      End If
   Next nIndex
   
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
        
   'Added by Lydia 2020/11/10 ·í¡u¸É¥¿¡]201¡^¡vµo¤å®É¡A­Y¦³¡u¸ÉÀu¥ýÅv¡]208¡^¡v©|¥¼µo¤å¡A½Ð¼u¡u¸ÉÀu¥ýÅv©|¥¼µo¤å¡A½T©w§_¦P®Éµo¤å?¡v;¿ï¡u¬O¡v«h¤@¨Ö±N¸ÉÀu¥ýÅv¦¬¤å¤Wµo¤å¤é
   If m_bolMsg208 = True Then
        'Modified by Lydia 2021/08/13 §ï§ì¦¬¤å¸¹
        'strSql = "Update CaseProgress Set CP27=" & DBDATE(textCP27) & " Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='208' and cp158=0 and cp159=0 "
        strSql = "Update CaseProgress Set CP27=" & DBDATE(textCP27) & " Where cp09='" & m_str208 & "' "
        cnnConnection.Execute strSql
   End If
   'end 2020/11/10
   
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Add By Sindy 2010/7/8 ÀË¬d°Ó«~¸ê®Æ»P°ò¥»ÀÉ°Ó«~Ãþ§O¬O§_¤@­P
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2012/9/10
   ' ­Y¦³¼f¬d¤Ñ¼Æ, ·s¼W¤@µ§¶Ê¼f´Á­­ªº°O¿ý¨ì¤U¤@µ{§ÇÀÉ
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF05")) = False Then
         strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
         'Add By Sindy 2023/5/5 FCT­«·sµo¤å¡A­Y¤U¤@µ{§Ç¤w¦³¸Ó¦¬¤å¸¹¥¼Äò¿ì¤§¶Ê¼f´Á­­¡A«h§ó·s´Á­­§Y¥i¡A¤£­n¥t·s¼W´Á­­
         strExc(0) = "SELECT NP01,NP22 from NextProgress" & _
                     " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "UPDATE NextProgress SET NP08=" & PUB_GetWorkDay1(strNP08, True) & ",NP09=" & strNP08 & _
                     " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null "
            cnnConnection.Execute strSql
         Else
         '2023/5/5 END
            strNP07 = "305"
            strNP22 = GetNextProgressNo()
            'Modified by Lydia 2020/07/09 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ; +PUB_GetWorkDay1()
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                               PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
            cnnConnection.Execute strSql
         End If
      End If
   End If
   rsTmp.Close
   '2012/9/10 End
    'Added by Lydia 2016/05/09 §ó·sFCT¤§¡u¥Ó½Ð¡v¡B¡uÅÜ§ó¡v¡B¡u²¾Âà¡v¡B¡u±ÂÅv¡vµ¥®×¥ó©Ê½è¤§¶Ê¼f´Á­­
    If m_CP10 = "201" And m_CP43 <> "" Then
       strExc(2) = ""
       strExc(1) = CompDate(1, 3, DBDATE(textCP27))
       If Left(m_CP43, 1) >= "C" Then '¨Ì¾÷Ãö¨Ó¨ç
          strSql = "select c2.cp09,c2.cp10 from caseprogress c1,caseprogress c2 where c1.cp09='" & m_CP43 & "' and c1.cp43=c2.cp09(+) and nvl(c2.cp27,0) > 0  "
       Else
          'µo¤å®×¥ó©Ê½è¬°¸É¥¿201®É¡A¬ÛÃöÁ`¦¬¤å¸¹¬°ÅÜ§ó¡B²¾Âà¡B±ÂÅv¡B©µ®i®×¥B¤wµo¤å®É¡A¥H¸É¥¿ªºµo¤å¤é+3­Ó¤ë§ó·sÅÜ§ó¡B²¾Âà¡B±ÂÅv¡B©µ®i®×ªº¶Ê¼f´Á­­¡F¨Ò¡GFCT-009891¤§AA4010746
          strSql = "select cp09,cp10 from caseprogress where cp09='" & m_CP43 & "' and nvl(cp27,0) > 0 "
       End If
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
       If intI = 1 Then
          'Modified by Lydia 2016/05/24 +¥Ó½Ð®×
          If InStr("101,102,301,501,502", "" & RsTemp.Fields("cp10")) > 0 And "" & RsTemp.Fields("cp09") <> "" Then
             strExc(2) = RsTemp.Fields("cp09")
          End If
          'Added by Lydia 2016/06/08 ¥Ó½Ð®×¥Hµo¤å¤é+6­Ó¤ë§ó·s¥Ó½Ð¤§¶Ê¼f´Á­­
          If "" & RsTemp.Fields("cp10") = "101" And "" & RsTemp.Fields("cp09") <> "" Then
             strExc(1) = CompDate(1, 6, DBDATE(textCP27))
          End If
       End If
       If strExc(2) <> "" Then
           'Modified by Lydia 2020/07/09 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ; +PUB_GetWorkDay1()
           'Modified by Lydia 2023/11/13 ­Y¨Ì­ì³]©w³W«h§ó·s¤§¶Ê¼f´Á­­¤p©ó­ì¶Ê¼f´Á­­¡A«h¤£§ó·s­ì¶Ê¼f´Á­­¡C=>AND NP09<
           strSql = "UPDATE NEXTPROGRESS SET NP08=" & PUB_GetWorkDay1(strExc(1), True) & ",NP09=" & strExc(1) & _
                    " WHERE NP01='" & strExc(2) & "' AND NP07='305' AND NP06 IS NULL AND NP09 < " & strExc(1)
           cnnConnection.Execute strSql, intI
       End If
    End If
    'end 2016/05/09
    
'911107 nick transation
cnnConnection.CommitTrans

     'Add by nickc 2008/02/22 ÀË¬d¥N²z¤HEmail(»Ý¦Ò¼{¥i¯à¬°FF®×¥ó)
    PUB_CheckEMail m_CP44, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22

    ' ¦C¦L©w½Z
    If textPrint <> "N" Then
        PrintLetter
    End If
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
   If m_blnClkChgButton = False Then
      MsgBox "½Ð¿é¤JÅÜ§ó¨Æ¶µ!!!", vbExclamation + vbOKOnly
      Me.cmdMod.SetFocus
      GoTo EXITSUB
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
   
   'Added by Lydia 2020/11/10 ·í¡u¸É¥¿¡]201¡^¡vµo¤å®É¡A­Y¦³¡u¸ÉÀu¥ýÅv¡]208¡^¡v©|¥¼µo¤å¡A½Ð¼u¡u¸ÉÀu¥ýÅv©|¥¼µo¤å¡A½T©w¬O§_¦P®Éµo¤å?¡v;¿ï¡u¬O¡v«h¤@¨Ö±N¸ÉÀu¥ýÅv¦¬¤å¤Wµo¤å¤é
   m_bolMsg208 = False
   If m_TM01 = "FCT" And m_TM10 = "000" And m_CP10 = "201" And textCP118 = "Y" Then '¦]¬°¯È¥»­n¸gµo¤å«Ç¡A©Ò¥H±Æ°£¯È¥»
       'Modified by Lydia 2021/08/13 §ï§ì¦¬¤å¸¹ count(*) cnt=> CP09
       strExc(0) = "select Cp09 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='208' and cp158=0 and cp159=0 "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
          'Modified by Lydia 2021/08/13 §ï§ì¦¬¤å¸¹
          'If Val("" & RsTemp.Fields("cnt")) > 0 Then
          If "" & RsTemp.Fields("CP09") <> "" Then
             If MsgBox("¸ÉÀu¥ýÅv©|¥¼µo¤å¡A½T©w¬O§_¦P®Éµo¤å???", vbExclamation + vbOKCancel, "¦P®Éµo¤å") = vbOK Then
                m_bolMsg208 = True
                m_str208 = "" & RsTemp.Fields("CP09") 'Added by Lydia 2021/08/13  »Pªü½¬½T»{¹L¡G¸ÉÀu¥ýÅv¦P®É¥u¦³¤@¹D¡F¥t¥~¡A¥Ó½Ð®É¤@¨Öµo¤å¥D±iÀu¥ýÅv108¬O±NÀu¥ýÅv¤å¥óÂk¤J¥Ó½Ð¡A©Ò¥H¤£¥Î¥t¥~¤W¶ÇÀÉ®×¨ì¨÷©v°Ï¡C
             End If
          End If
       End If
   End If
   'end 2020/11/10
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textDN_GotFocus()
   InverseTextBox textDN
End Sub

Private Sub textPetition_GotFocus()
   InverseTextBox textPetition
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textTM58_GotFocus()
   InverseTextBox textTM58
End Sub

Private Sub textTM67_GotFocus()
   InverseTextBox textTM67
End Sub

Private Sub textTM47_GotFocus()
   InverseTextBox textTM47
End Sub

Private Sub textTM48_GotFocus()
   InverseTextBox textTM48
End Sub

Private Sub textTM49_GotFocus()
   InverseTextBox textTM49
End Sub

Private Sub textTM50_GotFocus()
   InverseTextBox textTM50
End Sub

Private Sub textTM51_GotFocus()
   InverseTextBox textTM51
End Sub

Private Sub textTM52_GotFocus()
   InverseTextBox textTM52
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

If Me.textPetition.Enabled = True Then
   Cancel = False
   textPetition_Validate Cancel
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

'Added by Lydia 2021/09/09 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If
    
TxtValidate = True
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ¦C¦L©w½Z
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   'Add by Morgan 2008/6/12
   Dim ET03 As String
   
   'Add By Sindy 2012/11/23 ±q¤U­±µ{¦¡©¹¤WMove¦Ü¦¹
   bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper) 'ÀË¬d¬O§_¥HE-Mail³qª¾
   '2012/11/23 End
   
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   m_strLanguage = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
   'Modify By Sindy 2010/6/30 µo¤å®É¸É¥¿¤£¥X©w½Z, §ï¦Ü¥Ó½Ð®Ñ®É¤@¨Ö¥X
'   InsExpField
'   Select Case m_CP10
'     Case "201": '¸É¥¿
'         ' ©w½Z»y¤å
'         Select Case m_strLanguage
'            ' ­^¤å
'            Case "2":
'                ET03 = "04"
'            ' ¤é¤å
'            Case "3":
'                ET03 = "05"
'         End Select
'   End Select
    
   If ET03 <> "" Then
      'Add by Morgan 2008/6/12
'      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'end 2009/10/20
         NowPrint m_CP09, "01", ET03, False, strUserNum, , , , , iCopy, , True, True
         MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_TM01) & " ]¡I"
      Else
      'end 2008/6/12
         NowPrint m_CP09, "01", ET03, False, strUserNum
      End If
   End If
   
End Sub

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
      
    Select Case m_CP10
    Case "201": '¸É¥¿
        ' ©w½Z»y¤å
        Select Case m_strLanguage
        ' ­^¤å
        Case "2":
            ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
            EndLetter "01", m_CP09, "04", strUserNum
            ' 2009/4/17 ADD BY SONIA§PÂ_¬O§_¦P®É¦³208¸ÉÀu¥ýÅv¤å¥ó
            StrSQLa = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' AND CP10='208' AND CP27 IS NULL AND CP57 IS NULL "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & _
                        "','¦C¦L³Æµù',' and the priority document(s)')"
               cnnConnection.Execute strSql
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2009/4/17 end
        ' ¤é¤å
        Case "3":
            ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
            EndLetter "01", m_CP09, "05", strUserNum
            ' 2009/4/23 ADD BY SONIA§PÂ_¬O§_¦P®É¦³208¸ÉÀu¥ýÅv¤å¥ó
            StrSQLa = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' AND CP10='208' AND CP27 IS NULL AND CP57 IS NULL "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & _
                        "','¦C¦L³Æµù','¤ÎÇZÀu¥ý“¸¥D±iÇR¥ÎÆêÇr¤é¥»¥XÄ@µý©ú®Ñ')"
               cnnConnection.Execute strSql
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2009/4/23 end
        End Select
    End Select
End Sub

'add by nickc 2007/02/13 ¥[¥Nªí¤H
Private Sub textTM94_GotFocus()
InverseTextBox textTM94
End Sub
Private Sub textTM95_GotFocus()
InverseTextBox textTM95
End Sub
Private Sub textTM96_GotFocus()
InverseTextBox textTM96
End Sub
Private Sub textTM97_GotFocus()
InverseTextBox textTM97
End Sub
Private Sub textTM98_GotFocus()
InverseTextBox textTM98
End Sub
Private Sub textTM99_GotFocus()
InverseTextBox textTM99
End Sub
Private Sub textTM100_GotFocus()
InverseTextBox textTM100
End Sub
Private Sub textTM101_GotFocus()
InverseTextBox textTM101
End Sub
Private Sub textTM102_GotFocus()
InverseTextBox textTM102
End Sub
Private Sub textTM103_GotFocus()
InverseTextBox textTM103
End Sub
Private Sub textTM104_GotFocus()
InverseTextBox textTM104
End Sub
Private Sub textTM105_GotFocus()
InverseTextBox textTM105
End Sub
Private Sub textTM106_GotFocus()
InverseTextBox TextTM106
End Sub
Private Sub textTM107_GotFocus()
InverseTextBox TextTM107
End Sub
Private Sub textTM108_GotFocus()
InverseTextBox TextTM108
End Sub
Private Sub textTM109_GotFocus()
InverseTextBox TextTM109
End Sub
Private Sub textTM110_GotFocus()
InverseTextBox TextTM110
End Sub
Private Sub textTM111_GotFocus()
InverseTextBox TextTM111
End Sub
Private Sub textTM112_GotFocus()
InverseTextBox TextTM112
End Sub
Private Sub textTM113_GotFocus()
InverseTextBox TextTM113
End Sub
Private Sub textTM114_GotFocus()
InverseTextBox TextTM114
End Sub
Private Sub textTM115_GotFocus()
InverseTextBox TextTM115
End Sub
Private Sub textTM116_GotFocus()
InverseTextBox TextTM116
End Sub
Private Sub textTM117_GotFocus()
InverseTextBox TextTM117
End Sub

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
      Cancel = Not PUB_CheckCP113(textCP113, m_TM01, m_CP10, m_CP14)
   End If
End Sub
'98/03/11 End

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
