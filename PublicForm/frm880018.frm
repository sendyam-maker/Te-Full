VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880018 
   BorderStyle     =   1  '單線固定
   Caption         =   "相同國家有舊案申請地址與客戶目前申請地址不同者"
   ClientHeight    =   5745
   ClientLeft      =   195
   ClientTop       =   2520
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdOK 
      Caption         =   "離開(&X)"
      Height          =   345
      Index           =   0
      Left            =   7860
      TabIndex        =   0
      Top             =   30
      Width           =   945
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5280
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   9313
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "申請人1"
      TabPicture(0)   =   "frm880018.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabAppl1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LabAppl1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LabAppl1(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LabAppl1(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "grdDataList(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "申請人2"
      TabPicture(1)   =   "frm880018.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LabAppl2(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "LabAppl2(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LabAppl2(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LabAppl2(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "grdDataList(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "申請人3"
      TabPicture(2)   =   "frm880018.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "LabAppl3(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label11"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "LabAppl3(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label12"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label13"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "LabAppl3(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "LabAppl3(3)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "grdDataList(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "申請人4"
      TabPicture(3)   =   "frm880018.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label14"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "LabAppl4(0)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label15"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "LabAppl4(1)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label16"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label17"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "LabAppl4(2)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "LabAppl4(3)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "grdDataList(3)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "申請人5"
      TabPicture(4)   =   "frm880018.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label18"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "LabAppl5(0)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label19"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "LabAppl5(1)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label20"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label21"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "LabAppl5(2)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "LabAppl5(3)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "grdDataList(4)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).ControlCount=   9
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   3765
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   1440
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   6641
         _Version        =   393216
         BackColor       =   -2147483624
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "本所案號| 案件申請地址"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   3765
         Index           =   1
         Left            =   -74910
         TabIndex        =   19
         Top             =   1440
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   6641
         _Version        =   393216
         BackColor       =   -2147483624
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "本所案號| 案件申請地址"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   3765
         Index           =   2
         Left            =   -74910
         TabIndex        =   28
         Top             =   1440
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   6641
         _Version        =   393216
         BackColor       =   -2147483624
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "本所案號| 案件申請地址"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   3765
         Index           =   3
         Left            =   -74910
         TabIndex        =   37
         Top             =   1440
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   6641
         _Version        =   393216
         BackColor       =   -2147483624
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "本所案號| 案件申請地址"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   3765
         Index           =   4
         Left            =   -74910
         TabIndex        =   46
         Top             =   1440
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   6641
         _Version        =   393216
         BackColor       =   -2147483624
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "本所案號| 案件申請地址"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.Label LabAppl5 
         Height          =   255
         Index           =   3
         Left            =   -73380
         TabIndex        =   45
         Top             =   1170
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl5(3)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabAppl5 
         Height          =   255
         Index           =   2
         Left            =   -73380
         TabIndex        =   44
         Top             =   900
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl5(2)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         Caption         =   "　　　　　(日)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   43
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label Label20 
         Caption         =   "　　　　　(英)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   42
         Top             =   900
         Width           =   1395
      End
      Begin MSForms.Label LabAppl5 
         Height          =   255
         Index           =   1
         Left            =   -73380
         TabIndex        =   41
         Top             =   630
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl5(1)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label19 
         Caption         =   "申請人地址(中)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   40
         Top             =   630
         Width           =   1395
      End
      Begin MSForms.Label LabAppl5 
         Height          =   255
         Index           =   0
         Left            =   -73380
         TabIndex        =   39
         Top             =   360
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl5(0)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         Caption         =   "申請人名稱："
         Height          =   255
         Left            =   -74520
         TabIndex        =   38
         Top             =   360
         Width           =   1095
      End
      Begin MSForms.Label LabAppl4 
         Height          =   255
         Index           =   3
         Left            =   -73380
         TabIndex        =   36
         Top             =   1170
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl4(3)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabAppl4 
         Height          =   255
         Index           =   2
         Left            =   -73380
         TabIndex        =   35
         Top             =   900
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl4(2)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label17 
         Caption         =   "　　　　　(日)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   34
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "　　　　　(英)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   33
         Top             =   900
         Width           =   1395
      End
      Begin MSForms.Label LabAppl4 
         Height          =   255
         Index           =   1
         Left            =   -73380
         TabIndex        =   32
         Top             =   630
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl4(1)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label15 
         Caption         =   "申請人地址(中)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   31
         Top             =   630
         Width           =   1395
      End
      Begin MSForms.Label LabAppl4 
         Height          =   255
         Index           =   0
         Left            =   -73380
         TabIndex        =   30
         Top             =   360
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl4(0)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         Caption         =   "申請人名稱："
         Height          =   255
         Left            =   -74520
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin MSForms.Label LabAppl3 
         Height          =   255
         Index           =   3
         Left            =   -73380
         TabIndex        =   27
         Top             =   1170
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl3(3)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabAppl3 
         Height          =   255
         Index           =   2
         Left            =   -73380
         TabIndex        =   26
         Top             =   900
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl3(2)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         Caption         =   "　　　　　(日)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   25
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "　　　　　(英)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   24
         Top             =   900
         Width           =   1395
      End
      Begin MSForms.Label LabAppl3 
         Height          =   255
         Index           =   1
         Left            =   -73380
         TabIndex        =   23
         Top             =   630
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl3(1)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         Caption         =   "申請人地址(中)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   22
         Top             =   630
         Width           =   1395
      End
      Begin MSForms.Label LabAppl3 
         Height          =   255
         Index           =   0
         Left            =   -73380
         TabIndex        =   21
         Top             =   360
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl3(0)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         Caption         =   "申請人名稱："
         Height          =   255
         Left            =   -74520
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin MSForms.Label LabAppl2 
         Height          =   255
         Index           =   3
         Left            =   -73380
         TabIndex        =   18
         Top             =   1170
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl2(3)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabAppl2 
         Height          =   255
         Index           =   2
         Left            =   -73380
         TabIndex        =   17
         Top             =   900
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl2(2)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         Caption         =   "　　　　　(日)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   16
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label Label8 
         Caption         =   "　　　　　(英)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   15
         Top             =   900
         Width           =   1395
      End
      Begin MSForms.Label LabAppl2 
         Height          =   255
         Index           =   1
         Left            =   -73380
         TabIndex        =   14
         Top             =   630
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl2(1)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "申請人地址(中)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   13
         Top             =   630
         Width           =   1395
      End
      Begin MSForms.Label LabAppl2 
         Height          =   255
         Index           =   0
         Left            =   -73380
         TabIndex        =   12
         Top             =   360
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl2(0)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "申請人名稱："
         Height          =   255
         Left            =   -74520
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin MSForms.Label LabAppl1 
         Height          =   255
         Index           =   3
         Left            =   1620
         TabIndex        =   9
         Top             =   1170
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl1(3)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabAppl1 
         Height          =   255
         Index           =   2
         Left            =   1620
         TabIndex        =   8
         Top             =   900
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl1(2)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         Caption         =   "　　　　　(日)："
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "　　　　　(英)："
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   1395
      End
      Begin MSForms.Label LabAppl1 
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   5
         Top             =   630
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl1(1)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "申請人地址(中)："
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   630
         Width           =   1395
      End
      Begin MSForms.Label LabAppl1 
         Height          =   255
         Index           =   0
         Left            =   1620
         TabIndex        =   3
         Top             =   360
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LabAppl1(0)"
         Size            =   "12568;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "申請人名稱："
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label26 
      Caption         =   "申請人5"
      Height          =   195
      Left            =   4230
      TabIndex        =   56
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label25 
      Caption         =   "申請人4"
      Height          =   195
      Left            =   3240
      TabIndex        =   55
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label24 
      Caption         =   "申請人3"
      Height          =   195
      Left            =   2250
      TabIndex        =   54
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label23 
      Caption         =   "申請人2"
      Height          =   195
      Left            =   1260
      TabIndex        =   53
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label22 
      Caption         =   "申請人1"
      Height          =   195
      Left            =   300
      TabIndex        =   52
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   4080
      TabIndex        =   51
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   3090
      TabIndex        =   50
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   2100
      TabIndex        =   49
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   1140
      TabIndex        =   48
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   150
      TabIndex        =   47
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frm880018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/01/13 Form2.0已修改 labApplX(0)/labApplX(1)/labApplX(3)/grdDataList()
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Dim m_iSelRow As Integer
Public fmParent As Form
Public RsTemp As New ADODB.Recordset
Public m_Appl1 As String, m_Appl2 As String, m_Appl3 As String, m_Appl4 As String, m_Appl5 As String
Dim bSStab0 As Boolean, bSStab1 As Boolean, bSStab2 As Boolean, bSStab3 As Boolean, bSStab4 As Boolean


Private Sub cmdOK_Click(Index As Integer)
   Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
   MoveFormToCenter Me
   bSStab0 = False: Label2(0).BackColor = &H8000000F
   bSStab1 = False: Label2(1).BackColor = &H8000000F
   bSStab2 = False: Label2(2).BackColor = &H8000000F
   bSStab3 = False: Label2(3).BackColor = &H8000000F
   bSStab4 = False: Label2(4).BackColor = &H8000000F
   For i = 0 To 3
      LabAppl1(i) = ""
   Next i
   For i = 0 To 3
      LabAppl2(i) = ""
   Next i
   For i = 0 To 3
      LabAppl3(i) = ""
   Next i
   For i = 0 To 3
      LabAppl4(i) = ""
   Next i
   For i = 0 To 3
      LabAppl5(i) = ""
   Next i
   QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm880018 = Nothing
End Sub

Private Sub grdSelected(Index As Integer, p_iRow As Integer)
   Dim lColor As Long, ii As Integer
   With grdDataList(Index)
      .row = p_iRow
      .col = 0
      If .CellBackColor = &H80000018 Then
         m_iSelRow = .row
         lColor = &HFFC0C0
      Else
         m_iSelRow = -1
         lColor = &H80000018
      End If
      For ii = 0 To .Cols - 1
         .col = ii
         .CellBackColor = lColor
      Next
   End With
End Sub

Private Sub GrdDataList_Click(Index As Integer)
   Dim iRow As Integer
   With grdDataList(Index)
      If .MouseRow > 0 And .MouseRow < .Rows Then
         .Visible = False
         iRow = .MouseRow
         If m_iSelRow > 0 Then
            grdSelected Index, m_iSelRow
         End If
         If m_iSelRow <> iRow Then
            grdSelected Index, iRow
         End If
         .Visible = True
      End If
   End With
End Sub

' 初始化 grdDataList
Private Sub InitialgrdDataList(Index As Integer)
   grdDataList(Index).Clear
'   grdDataList(Index).Rows = 1
'   grdDataList(Index).Cols = 2
   grdDataList(Index).row = 0
   grdDataList(Index).col = 0
   grdDataList(Index).Text = "本所案號"
   grdDataList(Index).ColWidth(0) = 1500
   grdDataList(Index).CellAlignment = flexAlignLeftCenter
   grdDataList(Index).col = 1
   grdDataList(Index).Text = " 案件申請地址"
   grdDataList(Index).ColWidth(1) = 7000
   grdDataList(Index).CellAlignment = flexAlignLeftCenter
End Sub

Private Sub QueryData()
Dim i As Integer, strName As String, nIndex As Integer, strText As String, strCurI As String
   
   strText = "": strCurI = ""
   With RsTemp
      .MoveFirst
      Do While .EOF = False
         If Trim(.Fields("id")) = m_Appl1 Then i = 0
         If Trim(.Fields("id")) = m_Appl2 Then i = 1
         If Trim(.Fields("id")) = m_Appl3 Then i = 2
         If Trim(.Fields("id")) = m_Appl4 Then i = 3
         If Trim(.Fields("id")) = m_Appl5 Then i = 4
         '申請人名稱
         strName = ""
         If "" & Trim(.Fields("cname")) <> "" Then strName = "" & Trim(.Fields("cname"))
         If "" & Trim(.Fields("ename")) <> "" Then
            If strName <> "" Then strName = strName & "／"
            strName = strName & "" & Trim(.Fields("ename"))
         End If
         If "" & Trim(.Fields("jname")) <> "" Then
            If strName <> "" Then strName = strName & "／"
            strName = strName & "" & Trim(.Fields("jname"))
         End If
         If i = 0 And bSStab0 = False Then
            Call InitialgrdDataList(i): bSStab0 = True: Label2(0).BackColor = &HFF&
            LabAppl1(0).Caption = "" & Trim(.Fields("id")) & "　" & strName
            LabAppl1(1).Caption = "" & Trim(.Fields("caddr"))
            LabAppl1(2).Caption = "" & Trim(.Fields("eaddr"))
            LabAppl1(3).Caption = "" & Trim(.Fields("jaddr"))
         End If
         If i = 1 And bSStab1 = False Then
            Call InitialgrdDataList(i): bSStab1 = True: Label2(1).BackColor = &HFF&
            LabAppl2(0).Caption = "" & Trim(.Fields("id")) & "　" & strName
            LabAppl2(1).Caption = "" & Trim(.Fields("caddr"))
            LabAppl2(2).Caption = "" & Trim(.Fields("eaddr"))
            LabAppl2(3).Caption = "" & Trim(.Fields("jaddr"))
         End If
         If i = 2 And bSStab2 = False Then
            Call InitialgrdDataList(i): bSStab2 = True: Label2(2).BackColor = &HFF&
            LabAppl3(0).Caption = "" & Trim(.Fields("id")) & "　" & strName
            LabAppl3(1).Caption = "" & Trim(.Fields("caddr"))
            LabAppl3(2).Caption = "" & Trim(.Fields("eaddr"))
            LabAppl3(3).Caption = "" & Trim(.Fields("jaddr"))
         End If
         If i = 3 And bSStab3 = False Then
            Call InitialgrdDataList(i): bSStab3 = True: Label2(3).BackColor = &HFF&
            LabAppl4(0).Caption = "" & Trim(.Fields("id")) & "　" & strName
            LabAppl4(1).Caption = "" & Trim(.Fields("caddr"))
            LabAppl4(2).Caption = "" & Trim(.Fields("eaddr"))
            LabAppl4(3).Caption = "" & Trim(.Fields("jaddr"))
         End If
         If i = 4 And bSStab4 = False Then
            Call InitialgrdDataList(i): bSStab4 = True: Label2(4).BackColor = &HFF&
            LabAppl5(0).Caption = "" & Trim(.Fields("id")) & "　" & strName
            LabAppl5(1).Caption = "" & Trim(.Fields("caddr"))
            LabAppl5(2).Caption = "" & Trim(.Fields("eaddr"))
            LabAppl5(3).Caption = "" & Trim(.Fields("jaddr"))
         End If
         '新增資料
         If grdDataList(i).TextMatrix((grdDataList(i).Rows - 1), 1) <> "" Then
            grdDataList(i).Rows = grdDataList(i).Rows + 1
         End If
         nIndex = grdDataList(i).Rows - 1
         If strCurI <> CStr(i) Or strText <> (.Fields(0) & "-" & .Fields(1) & "-" & .Fields(2) & "-" & .Fields(3)) Then
            grdDataList(i).TextMatrix(nIndex, 0) = .Fields(0) & "-" & .Fields(1) & "-" & .Fields(2) & "-" & .Fields(3)
         End If
         grdDataList(i).TextMatrix(nIndex, 1) = " " & Trim(.Fields("caseAddr"))
         strText = .Fields(0) & "-" & .Fields(1) & "-" & .Fields(2) & "-" & .Fields(3)
         strCurI = i
         .MoveNext
      Loop
   End With
   '顯示Focus的項目
   If bSStab0 = True Then
      i = 0
   ElseIf bSStab1 = True Then
      i = 1
   ElseIf bSStab2 = True Then
      i = 2
   ElseIf bSStab3 = True Then
      i = 3
   ElseIf bSStab4 = True Then
      i = 4
   End If
   SSTab1.Tab = i
   'grdDataList(i).row = grdDataList(i).Rows - 1
End Sub

