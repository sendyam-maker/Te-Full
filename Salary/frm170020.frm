VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170020 
   BorderStyle     =   1  '單線固定
   Caption         =   "年終考績資料"
   ClientHeight    =   5976
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8040
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5976
   ScaleWidth      =   8040
   Begin TabDlg.SSTab SSTab1 
      Height          =   5310
      Left            =   0
      TabIndex        =   187
      Top             =   630
      Width           =   7995
      _ExtentX        =   14097
      _ExtentY        =   9356
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170020.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textYM01"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "picU2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "VU"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "picC2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "picD2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "VC"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "VD"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170020.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line5"
      Tab(1).Control(1)=   "Line4"
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(5)=   "txt1(0)"
      Tab(1).Control(6)=   "txt1(1)"
      Tab(1).Control(7)=   "txt1(2)"
      Tab(1).Control(8)=   "txt1(3)"
      Tab(1).Control(9)=   "cmdok"
      Tab(1).ControlCount=   10
      Begin VB.VScrollBar VD 
         Height          =   825
         Left            =   7680
         TabIndex        =   380
         Top             =   4410
         Width           =   255
      End
      Begin VB.VScrollBar VC 
         Height          =   855
         Left            =   7680
         TabIndex        =   379
         Top             =   3210
         Width           =   255
      End
      Begin VB.PictureBox picD2 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   300
         ScaleHeight     =   804
         ScaleWidth      =   7356
         TabIndex        =   317
         Top             =   4410
         Width           =   7380
         Begin VB.PictureBox picD1 
            Appearance      =   0  '平面
            BackColor       =   &H80000004&
            BorderStyle     =   0  '沒有框線
            ForeColor       =   &H80000008&
            Height          =   5415
            Left            =   0
            ScaleHeight     =   5412
            ScaleWidth      =   7380
            TabIndex        =   318
            Top             =   0
            Width           =   7380
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   0
               Left            =   0
               MaxLength       =   6
               TabIndex        =   121
               Top             =   0
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   1
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   122
               Top             =   0
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   2
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   123
               Top             =   0
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   3
               Left            =   0
               MaxLength       =   6
               TabIndex        =   124
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   4
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   125
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   5
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   126
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   6
               Left            =   0
               MaxLength       =   6
               TabIndex        =   127
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   7
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   128
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   8
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   129
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   9
               Left            =   0
               MaxLength       =   6
               TabIndex        =   130
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   10
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   131
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   11
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   132
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   12
               Left            =   0
               MaxLength       =   6
               TabIndex        =   133
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   13
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   134
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   14
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   135
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   15
               Left            =   0
               MaxLength       =   6
               TabIndex        =   136
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   16
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   137
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   17
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   138
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   18
               Left            =   0
               MaxLength       =   6
               TabIndex        =   139
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   19
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   140
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   20
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   141
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   21
               Left            =   0
               MaxLength       =   6
               TabIndex        =   142
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   22
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   143
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   23
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   144
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   24
               Left            =   0
               MaxLength       =   6
               TabIndex        =   145
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   25
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   146
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   26
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   147
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   27
               Left            =   0
               MaxLength       =   6
               TabIndex        =   148
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   28
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   149
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   29
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   150
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   30
               Left            =   0
               MaxLength       =   6
               TabIndex        =   151
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   31
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   152
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   32
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   153
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   33
               Left            =   0
               MaxLength       =   6
               TabIndex        =   154
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   34
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   155
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   35
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   156
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   36
               Left            =   0
               MaxLength       =   6
               TabIndex        =   157
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   37
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   158
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   38
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   159
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   39
               Left            =   0
               MaxLength       =   6
               TabIndex        =   160
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   40
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   161
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   41
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   162
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   42
               Left            =   0
               MaxLength       =   6
               TabIndex        =   163
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   43
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   164
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   44
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   165
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   45
               Left            =   0
               MaxLength       =   6
               TabIndex        =   166
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   46
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   167
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   47
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   168
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   48
               Left            =   0
               MaxLength       =   6
               TabIndex        =   169
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   49
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   170
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   50
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   171
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   51
               Left            =   0
               MaxLength       =   6
               TabIndex        =   172
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   52
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   173
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   53
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   174
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   54
               Left            =   0
               MaxLength       =   6
               TabIndex        =   175
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   55
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   176
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   56
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   177
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   57
               Left            =   0
               MaxLength       =   6
               TabIndex        =   178
               Top             =   5130
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   58
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   179
               Top             =   5130
               Width           =   645
            End
            Begin VB.TextBox txtD 
               Height          =   270
               Index           =   59
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   180
               Top             =   5130
               Width           =   645
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   0
               Left            =   645
               TabIndex        =   378
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   1
               Left            =   3105
               TabIndex        =   377
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   2
               Left            =   5565
               TabIndex        =   376
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   3
               Left            =   645
               TabIndex        =   375
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   4
               Left            =   3105
               TabIndex        =   374
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   5
               Left            =   5565
               TabIndex        =   373
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   6
               Left            =   645
               TabIndex        =   372
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   7
               Left            =   3105
               TabIndex        =   371
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   8
               Left            =   5565
               TabIndex        =   370
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   9
               Left            =   645
               TabIndex        =   369
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   10
               Left            =   3105
               TabIndex        =   368
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   11
               Left            =   5565
               TabIndex        =   367
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   12
               Left            =   645
               TabIndex        =   366
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   13
               Left            =   3105
               TabIndex        =   365
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   14
               Left            =   5565
               TabIndex        =   364
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   15
               Left            =   645
               TabIndex        =   363
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   16
               Left            =   3105
               TabIndex        =   362
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   17
               Left            =   5565
               TabIndex        =   361
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   18
               Left            =   645
               TabIndex        =   360
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   19
               Left            =   3105
               TabIndex        =   359
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   20
               Left            =   5565
               TabIndex        =   358
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   21
               Left            =   645
               TabIndex        =   357
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   22
               Left            =   3105
               TabIndex        =   356
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   23
               Left            =   5565
               TabIndex        =   355
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   24
               Left            =   645
               TabIndex        =   354
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   25
               Left            =   3105
               TabIndex        =   353
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   26
               Left            =   5565
               TabIndex        =   352
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   27
               Left            =   645
               TabIndex        =   351
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   28
               Left            =   3105
               TabIndex        =   350
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   29
               Left            =   5565
               TabIndex        =   349
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   30
               Left            =   645
               TabIndex        =   348
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   31
               Left            =   3105
               TabIndex        =   347
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   32
               Left            =   5565
               TabIndex        =   346
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   33
               Left            =   645
               TabIndex        =   345
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   34
               Left            =   3105
               TabIndex        =   344
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   35
               Left            =   5565
               TabIndex        =   343
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   36
               Left            =   645
               TabIndex        =   342
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   37
               Left            =   3105
               TabIndex        =   341
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   38
               Left            =   5565
               TabIndex        =   340
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   39
               Left            =   645
               TabIndex        =   339
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   40
               Left            =   3105
               TabIndex        =   338
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   41
               Left            =   5565
               TabIndex        =   337
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   42
               Left            =   645
               TabIndex        =   336
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   43
               Left            =   3105
               TabIndex        =   335
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   44
               Left            =   5565
               TabIndex        =   334
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   45
               Left            =   645
               TabIndex        =   333
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   46
               Left            =   3105
               TabIndex        =   332
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   47
               Left            =   5565
               TabIndex        =   331
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   48
               Left            =   645
               TabIndex        =   330
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   49
               Left            =   3105
               TabIndex        =   329
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   50
               Left            =   5565
               TabIndex        =   328
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   51
               Left            =   645
               TabIndex        =   327
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   52
               Left            =   3105
               TabIndex        =   326
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   53
               Left            =   5565
               TabIndex        =   325
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   54
               Left            =   645
               TabIndex        =   324
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   55
               Left            =   3105
               TabIndex        =   323
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   56
               Left            =   5565
               TabIndex        =   322
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   57
               Left            =   645
               TabIndex        =   321
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   58
               Left            =   3105
               TabIndex        =   320
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblD 
               Height          =   270
               Index           =   59
               Left            =   5565
               TabIndex        =   319
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
         End
      End
      Begin VB.PictureBox picC2 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   300
         ScaleHeight     =   828
         ScaleWidth      =   7356
         TabIndex        =   255
         Top             =   3210
         Width           =   7380
         Begin VB.PictureBox picC1 
            Appearance      =   0  '平面
            BackColor       =   &H80000004&
            BorderStyle     =   0  '沒有框線
            ForeColor       =   &H80000008&
            Height          =   5415
            Left            =   0
            ScaleHeight     =   5412
            ScaleWidth      =   7380
            TabIndex        =   256
            Top             =   0
            Width           =   7380
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   0
               Left            =   0
               MaxLength       =   6
               TabIndex        =   61
               Top             =   0
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   1
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   62
               Top             =   0
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   2
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   63
               Top             =   0
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   3
               Left            =   0
               MaxLength       =   6
               TabIndex        =   64
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   4
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   65
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   5
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   66
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   6
               Left            =   0
               MaxLength       =   6
               TabIndex        =   67
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   7
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   68
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   8
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   69
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   9
               Left            =   0
               MaxLength       =   6
               TabIndex        =   70
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   10
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   71
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   11
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   72
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   12
               Left            =   0
               MaxLength       =   6
               TabIndex        =   73
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   13
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   74
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   14
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   75
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   15
               Left            =   0
               MaxLength       =   6
               TabIndex        =   76
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   16
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   77
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   17
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   78
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   18
               Left            =   0
               MaxLength       =   6
               TabIndex        =   79
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   19
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   80
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   20
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   81
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   21
               Left            =   0
               MaxLength       =   6
               TabIndex        =   82
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   22
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   83
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   23
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   84
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   24
               Left            =   0
               MaxLength       =   6
               TabIndex        =   85
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   25
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   86
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   26
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   87
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   27
               Left            =   0
               MaxLength       =   6
               TabIndex        =   88
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   28
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   89
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   29
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   90
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   30
               Left            =   0
               MaxLength       =   6
               TabIndex        =   91
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   31
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   92
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   32
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   93
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   33
               Left            =   0
               MaxLength       =   6
               TabIndex        =   94
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   34
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   95
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   35
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   96
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   36
               Left            =   0
               MaxLength       =   6
               TabIndex        =   97
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   37
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   98
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   38
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   99
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   39
               Left            =   0
               MaxLength       =   6
               TabIndex        =   100
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   40
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   101
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   41
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   102
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   42
               Left            =   0
               MaxLength       =   6
               TabIndex        =   103
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   43
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   104
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   44
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   105
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   45
               Left            =   0
               MaxLength       =   6
               TabIndex        =   106
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   46
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   107
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   47
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   108
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   48
               Left            =   0
               MaxLength       =   6
               TabIndex        =   109
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   49
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   110
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   50
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   111
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   51
               Left            =   0
               MaxLength       =   6
               TabIndex        =   112
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   52
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   113
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   53
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   114
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   54
               Left            =   0
               MaxLength       =   6
               TabIndex        =   115
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   55
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   116
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   56
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   117
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   57
               Left            =   0
               MaxLength       =   6
               TabIndex        =   118
               Top             =   5130
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   58
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   119
               Top             =   5130
               Width           =   645
            End
            Begin VB.TextBox txtC 
               Height          =   270
               Index           =   59
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   120
               Top             =   5130
               Width           =   645
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   0
               Left            =   645
               TabIndex        =   316
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   1
               Left            =   3105
               TabIndex        =   315
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   2
               Left            =   5565
               TabIndex        =   314
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   3
               Left            =   645
               TabIndex        =   313
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   4
               Left            =   3105
               TabIndex        =   312
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   5
               Left            =   5565
               TabIndex        =   311
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   6
               Left            =   645
               TabIndex        =   310
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   7
               Left            =   3105
               TabIndex        =   309
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   8
               Left            =   5565
               TabIndex        =   308
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   9
               Left            =   645
               TabIndex        =   307
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   10
               Left            =   3105
               TabIndex        =   306
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   11
               Left            =   5565
               TabIndex        =   305
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   12
               Left            =   645
               TabIndex        =   304
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   13
               Left            =   3105
               TabIndex        =   303
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   14
               Left            =   5565
               TabIndex        =   302
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   104
               Left            =   645
               TabIndex        =   301
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   16
               Left            =   3105
               TabIndex        =   300
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   17
               Left            =   5565
               TabIndex        =   299
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   18
               Left            =   645
               TabIndex        =   298
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   19
               Left            =   3105
               TabIndex        =   297
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   20
               Left            =   5565
               TabIndex        =   296
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   21
               Left            =   645
               TabIndex        =   295
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   22
               Left            =   3105
               TabIndex        =   294
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   23
               Left            =   5565
               TabIndex        =   293
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   24
               Left            =   645
               TabIndex        =   292
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   25
               Left            =   3105
               TabIndex        =   291
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   26
               Left            =   5565
               TabIndex        =   290
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   27
               Left            =   645
               TabIndex        =   289
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   28
               Left            =   3105
               TabIndex        =   288
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   29
               Left            =   5565
               TabIndex        =   287
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   30
               Left            =   645
               TabIndex        =   286
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   31
               Left            =   3105
               TabIndex        =   285
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   32
               Left            =   5565
               TabIndex        =   284
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   33
               Left            =   645
               TabIndex        =   283
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   34
               Left            =   3105
               TabIndex        =   282
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   35
               Left            =   5565
               TabIndex        =   281
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   36
               Left            =   645
               TabIndex        =   280
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   37
               Left            =   3105
               TabIndex        =   279
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   38
               Left            =   5565
               TabIndex        =   278
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   39
               Left            =   645
               TabIndex        =   277
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   40
               Left            =   3105
               TabIndex        =   276
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   41
               Left            =   5565
               TabIndex        =   275
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   42
               Left            =   645
               TabIndex        =   274
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   43
               Left            =   3105
               TabIndex        =   273
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   44
               Left            =   5565
               TabIndex        =   272
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   45
               Left            =   645
               TabIndex        =   271
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   46
               Left            =   3105
               TabIndex        =   270
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   47
               Left            =   5565
               TabIndex        =   269
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   48
               Left            =   645
               TabIndex        =   268
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   49
               Left            =   3105
               TabIndex        =   267
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   50
               Left            =   5565
               TabIndex        =   266
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   51
               Left            =   645
               TabIndex        =   265
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   52
               Left            =   3105
               TabIndex        =   264
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   53
               Left            =   5565
               TabIndex        =   263
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   54
               Left            =   645
               TabIndex        =   262
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   55
               Left            =   3105
               TabIndex        =   261
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   56
               Left            =   5565
               TabIndex        =   260
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   57
               Left            =   645
               TabIndex        =   259
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   58
               Left            =   3105
               TabIndex        =   258
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblC 
               Height          =   270
               Index           =   59
               Left            =   5565
               TabIndex        =   257
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
         End
      End
      Begin VB.VScrollBar VU 
         Height          =   1935
         Left            =   7680
         TabIndex        =   194
         Top             =   990
         Width           =   255
      End
      Begin VB.PictureBox picU2 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   300
         ScaleHeight     =   1908
         ScaleWidth      =   7356
         TabIndex        =   192
         Top             =   990
         Width           =   7380
         Begin VB.PictureBox picU1 
            Appearance      =   0  '平面
            BackColor       =   &H80000004&
            BorderStyle     =   0  '沒有框線
            ForeColor       =   &H80000008&
            Height          =   5415
            Left            =   0
            ScaleHeight     =   5412
            ScaleWidth      =   7380
            TabIndex        =   193
            Top             =   0
            Width           =   7380
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   59
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   60
               Top             =   5130
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   58
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   59
               Top             =   5130
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   57
               Left            =   0
               MaxLength       =   6
               TabIndex        =   58
               Top             =   5130
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   56
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   57
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   55
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   56
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   54
               Left            =   0
               MaxLength       =   6
               TabIndex        =   55
               Top             =   4860
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   53
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   54
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   52
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   53
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   51
               Left            =   0
               MaxLength       =   6
               TabIndex        =   52
               Top             =   4590
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   50
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   51
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   49
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   50
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   48
               Left            =   0
               MaxLength       =   6
               TabIndex        =   49
               Top             =   4320
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   47
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   48
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   46
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   47
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   45
               Left            =   0
               MaxLength       =   6
               TabIndex        =   46
               Top             =   4050
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   44
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   45
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   43
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   44
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   42
               Left            =   0
               MaxLength       =   6
               TabIndex        =   43
               Top             =   3780
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   41
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   42
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   40
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   41
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   39
               Left            =   0
               MaxLength       =   6
               TabIndex        =   40
               Top             =   3510
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   38
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   39
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   37
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   38
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   36
               Left            =   0
               MaxLength       =   6
               TabIndex        =   37
               Top             =   3240
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   35
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   36
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   34
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   35
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   33
               Left            =   0
               MaxLength       =   6
               TabIndex        =   34
               Top             =   2970
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   32
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   33
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   31
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   32
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   30
               Left            =   0
               MaxLength       =   6
               TabIndex        =   31
               Top             =   2700
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   29
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   30
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   28
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   29
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   27
               Left            =   0
               MaxLength       =   6
               TabIndex        =   28
               Top             =   2430
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   26
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   27
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   25
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   26
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   24
               Left            =   0
               MaxLength       =   6
               TabIndex        =   25
               Top             =   2160
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   23
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   24
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   22
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   23
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   21
               Left            =   0
               MaxLength       =   6
               TabIndex        =   22
               Top             =   1890
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   20
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   21
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   19
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   20
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   18
               Left            =   0
               MaxLength       =   6
               TabIndex        =   19
               Top             =   1620
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   17
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   18
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   16
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   17
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   15
               Left            =   0
               MaxLength       =   6
               TabIndex        =   16
               Top             =   1350
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   14
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   15
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   13
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   14
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   12
               Left            =   0
               MaxLength       =   6
               TabIndex        =   13
               Top             =   1080
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   11
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   12
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   10
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   11
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   9
               Left            =   0
               MaxLength       =   6
               TabIndex        =   10
               Top             =   810
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   8
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   9
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   7
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   8
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   6
               Left            =   0
               MaxLength       =   6
               TabIndex        =   7
               Top             =   540
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   5
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   6
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   4
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   5
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   3
               Left            =   0
               MaxLength       =   6
               TabIndex        =   4
               Top             =   270
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   2
               Left            =   4920
               MaxLength       =   6
               TabIndex        =   3
               Top             =   0
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   1
               Left            =   2460
               MaxLength       =   6
               TabIndex        =   2
               Top             =   0
               Width           =   645
            End
            Begin VB.TextBox txtU 
               Height          =   270
               Index           =   0
               Left            =   0
               MaxLength       =   6
               TabIndex        =   1
               Top             =   0
               Width           =   645
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   59
               Left            =   5565
               TabIndex        =   254
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   58
               Left            =   3105
               TabIndex        =   253
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   57
               Left            =   645
               TabIndex        =   252
               Top             =   5160
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   56
               Left            =   5565
               TabIndex        =   251
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   55
               Left            =   3105
               TabIndex        =   250
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   54
               Left            =   645
               TabIndex        =   249
               Top             =   4890
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   53
               Left            =   5565
               TabIndex        =   248
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   52
               Left            =   3105
               TabIndex        =   247
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   51
               Left            =   645
               TabIndex        =   246
               Top             =   4620
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   50
               Left            =   5565
               TabIndex        =   245
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   49
               Left            =   3105
               TabIndex        =   244
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   48
               Left            =   645
               TabIndex        =   243
               Top             =   4350
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   47
               Left            =   5565
               TabIndex        =   242
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   46
               Left            =   3105
               TabIndex        =   241
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   45
               Left            =   645
               TabIndex        =   240
               Top             =   4080
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   44
               Left            =   5565
               TabIndex        =   239
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   43
               Left            =   3105
               TabIndex        =   238
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   42
               Left            =   645
               TabIndex        =   237
               Top             =   3810
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   41
               Left            =   5565
               TabIndex        =   236
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   40
               Left            =   3105
               TabIndex        =   235
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   39
               Left            =   645
               TabIndex        =   234
               Top             =   3540
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   38
               Left            =   5565
               TabIndex        =   233
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   37
               Left            =   3105
               TabIndex        =   232
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   36
               Left            =   645
               TabIndex        =   231
               Top             =   3270
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   35
               Left            =   5565
               TabIndex        =   230
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   34
               Left            =   3105
               TabIndex        =   229
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   33
               Left            =   645
               TabIndex        =   228
               Top             =   3000
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   32
               Left            =   5565
               TabIndex        =   227
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   31
               Left            =   3105
               TabIndex        =   226
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   30
               Left            =   645
               TabIndex        =   225
               Top             =   2730
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   29
               Left            =   5565
               TabIndex        =   224
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   28
               Left            =   3105
               TabIndex        =   223
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   27
               Left            =   645
               TabIndex        =   222
               Top             =   2460
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   26
               Left            =   5565
               TabIndex        =   221
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   25
               Left            =   3105
               TabIndex        =   220
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   24
               Left            =   645
               TabIndex        =   219
               Top             =   2190
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   23
               Left            =   5565
               TabIndex        =   218
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   22
               Left            =   3105
               TabIndex        =   217
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   21
               Left            =   645
               TabIndex        =   216
               Top             =   1920
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   20
               Left            =   5565
               TabIndex        =   215
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   19
               Left            =   3105
               TabIndex        =   214
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   18
               Left            =   645
               TabIndex        =   213
               Top             =   1650
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   17
               Left            =   5565
               TabIndex        =   212
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   16
               Left            =   3105
               TabIndex        =   211
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   15
               Left            =   645
               TabIndex        =   210
               Top             =   1380
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   14
               Left            =   5565
               TabIndex        =   209
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   13
               Left            =   3105
               TabIndex        =   208
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   12
               Left            =   645
               TabIndex        =   207
               Top             =   1110
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   11
               Left            =   5565
               TabIndex        =   206
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   10
               Left            =   3105
               TabIndex        =   205
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   9
               Left            =   645
               TabIndex        =   204
               Top             =   840
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   8
               Left            =   5565
               TabIndex        =   203
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   7
               Left            =   3105
               TabIndex        =   202
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   6
               Left            =   645
               TabIndex        =   201
               Top             =   570
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   5
               Left            =   5565
               TabIndex        =   200
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   4
               Left            =   3105
               TabIndex        =   199
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   3
               Left            =   645
               TabIndex        =   198
               Top             =   300
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   2
               Left            =   5565
               TabIndex        =   197
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   1
               Left            =   3105
               TabIndex        =   196
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.Label lblU 
               Height          =   270
               Index           =   0
               Left            =   645
               TabIndex        =   195
               Top             =   30
               Width           =   1000
               VariousPropertyBits=   27
               Caption         =   "LblFM2"
               Size            =   "5741;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
         End
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   255
         Left            =   -68700
         TabIndex        =   185
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70080
         MaxLength       =   4
         TabIndex        =   184
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71070
         MaxLength       =   4
         TabIndex        =   183
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72960
         MaxLength       =   6
         TabIndex        =   182
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -74010
         MaxLength       =   6
         TabIndex        =   181
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox textYM01 
         Height          =   285
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   0
         Top             =   390
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   4665
         Left            =   -74970
         TabIndex        =   186
         Top             =   660
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   8234
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
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
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "丙等："
         Height          =   180
         Left            =   330
         TabIndex        =   383
         Top             =   4140
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "乙等："
         Height          =   180
         Left            =   330
         TabIndex        =   382
         Top             =   2970
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "優等："
         Height          =   180
         Left            =   330
         TabIndex        =   381
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "年度："
         Height          =   180
         Left            =   -71790
         TabIndex        =   191
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Left            =   -74940
         TabIndex        =   190
         Top             =   360
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -73320
         X2              =   -72630
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line5 
         X1              =   -70350
         X2              =   -69750
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   2190
         TabIndex        =   189
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年度：                     (ex:96)"
         Height          =   180
         Left            =   330
         TabIndex        =   188
         Top             =   480
         Width           =   1995
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":0670
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":084C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":0E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170020.frx":1E10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   384
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
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
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm170020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/23 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by nickc 2006/12/03 copy from frm140401
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的本所案號
Dim m_FirstKEY As String
' 最後一筆資料的本所案號
Dim m_LastKEY As String
' 目前正在顯示的本所案號
Dim m_CurrKEY As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim MyKind As String
Const MaxTxtU As Integer = 60  '60     最多 60 目前只 30 個 'Modified by Morgan 2012/1/10 改 60
Const MaxTxtC As Integer = 9    '60     最多 60 目前只 9 個
Const MaxTxtD As Integer = 9    '60     最多 60 目前只 9 個
Dim NowMaxU As Integer
Dim NowMaxC As Integer
Dim NowMaxD As Integer
Dim i As Integer


Private Sub txtC_GotFocus(Index As Integer)
   If m_EditMode <> 0 Then
       InverseTextBox txtC(Index)
       CloseIme
       If txtC(Index).Top + txtC(Index).Height + picC1.Top > picC2.Height Then
           If VC.Value > VC.max Then
               If VC.Value - ((((txtC(Index).Top + txtC(Index).Height + picC1.Top) - picC2.Height) \ 270) + (IIf(((txtC(Index).Top + txtC(Index).Height + picC1.Top) - picC2.Height) Mod 270 <> 0, 1, 0))) >= VC.max Then
                   VC.Value = VC.Value - ((((txtC(Index).Top + txtC(Index).Height + picC1.Top) - picC2.Height) \ 270) + (IIf(((txtC(Index).Top + txtC(Index).Height + picC1.Top) - picC2.Height) Mod 270 <> 0, 1, 0)))
               Else
                   VC.Value = VC.max
               End If
           Else
               VC.Value = VC.max
           End If
       ElseIf txtC(Index).Top + picC1.Top < 0 Then
           If VC.Value >= VC.max Then
               If VC.Value + (Abs(txtC(Index).Top + picC1.Top) / 270) >= VC.max Then
                   VC.Value = VC.Value + (Abs(txtC(Index).Top + picC1.Top) / 270)
               Else
                   VC.Value = 0
               End If
           Else
               VC.Value = 0
           End If
       End If
   End If
End Sub

Private Sub txtC_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtC_Validate(Index As Integer, Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If txtC(Index) = "" Then lblC(Index).Caption = ""
   If m_EditMode <> 0 And txtC(Index) <> "" Then
      
      'Modify by Morgan 2009/6/24 改電腦中心補資料時判斷人事異動，否則不可輸入離職員工GetStaffNameSpec
      If Pub_StrUserSt03 = "M51" Then
         lblC(Index).Caption = GetValidStaffName(txtC(Index), DBDATE(textYM01 & "1231"))
      Else
         'modify by sonia 2016/1/6 可以輸留職停薪人員
         'lblC(Index).Caption = GetStaffName(txtC(Index), IIf(m_EditMode = 1 Or m_EditMode = 2, False, True))
         lblC(Index).Caption = GetStaffNameSpec(txtC(Index), IIf(m_EditMode = 1 Or m_EditMode = 2, False, True))
      End If
      'end 2009/6/24
      If lblC(Index).Caption = "" Then
         MsgBox "員工代號錯誤！查無此員工！", vbInformation
         Cancel = True
         Exit Sub
      End If
      'add by sonia 2016/1/6 不可為當年度不得參加考績YM02='*'的人員
      strSql = "SELECT * FROM YearMerit WHERE YM01 = '" & Val(textYM01) + 1911 & "' AND YM02='*' AND YM03='" & txtC(Index) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         MsgBox "此員工 " & Val(textYM01) & "年不得參加年終考績！", vbInformation
         Cancel = True
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Sub
      End If
      rsTmp.Close
      Set rsTmp = Nothing
      'end 2016/1/6
   End If
End Sub

Private Sub txtD_GotFocus(Index As Integer)
   If m_EditMode <> 0 Then
       InverseTextBox txtD(Index)
       CloseIme
       If txtD(Index).Top + txtD(Index).Height + picD1.Top > picD2.Height Then
           If VD.Value > VD.max Then
               If VD.Value - ((((txtD(Index).Top + txtD(Index).Height + picD1.Top) - picD2.Height) \ 270) + (IIf(((txtD(Index).Top + txtD(Index).Height + picD1.Top) - picD2.Height) Mod 270 <> 0, 1, 0))) >= VD.max Then
                   VD.Value = VD.Value - ((((txtD(Index).Top + txtD(Index).Height + picD1.Top) - picD2.Height) \ 270) + (IIf(((txtD(Index).Top + txtD(Index).Height + picD1.Top) - picD2.Height) Mod 270 <> 0, 1, 0)))
               Else
                   VD.Value = VD.max
               End If
           Else
               VD.Value = VD.max
           End If
       ElseIf txtD(Index).Top + picD1.Top < 0 Then
           If VD.Value >= VD.max Then
               If VD.Value + (Abs(txtD(Index).Top + picD1.Top) / 270) >= VD.max Then
                   VD.Value = VD.Value + (Abs(txtD(Index).Top + picD1.Top) / 270)
               Else
                   VD.Value = 0
               End If
           Else
               VD.Value = 0
           End If
       End If
   End If
End Sub

Private Sub txtD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtD_Validate(Index As Integer, Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If txtD(Index) = "" Then lblD(Index).Caption = ""
   If m_EditMode <> 0 And txtD(Index) <> "" Then
      'Modify by Morgan 2009/6/24 改電腦中心補資料時判斷人事異動，否則不可輸入離職員工
      If Pub_StrUserSt03 = "M51" Then
         lblD(Index).Caption = GetValidStaffName(txtD(Index), DBDATE(textYM01 & "1231"))
      Else
         'modify by sonia 2016/1/6 可以輸留職停薪人員
         'lblD(Index).Caption = GetStaffName(txtD(Index), IIf(m_EditMode = 1 Or m_EditMode = 2, False, True))
         lblD(Index).Caption = GetStaffNameSpec(txtD(Index), IIf(m_EditMode = 1 Or m_EditMode = 2, False, True))
      End If
      'end 2009/6/24
      If lblD(Index).Caption = "" Then
         MsgBox "員工代號錯誤！查無此員工！", vbInformation
         Cancel = True
         Exit Sub
      End If
      'add by sonia 2016/1/6 不可為當年度不得參加考績YM02='*'的人員
      strSql = "SELECT * FROM YearMerit WHERE YM01 = '" & Val(textYM01) + 1911 & "' AND YM02='*' AND YM03='" & txtD(Index) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         MsgBox "此員工 " & Val(textYM01) & "年不得參加年終考績！", vbInformation
         Cancel = True
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Sub
      End If
      rsTmp.Close
      Set rsTmp = Nothing
      'end 2016/1/6
   End If
End Sub

Private Sub txtU_GotFocus(Index As Integer)
   If m_EditMode <> 0 Then
       InverseTextBox txtU(Index)
       CloseIme
       If txtU(Index).Top + txtU(Index).Height + picU1.Top > picU2.Height Then
           If VU.Value > VU.max Then
               If VU.Value - ((((txtU(Index).Top + txtU(Index).Height + picU1.Top) - picU2.Height) \ 270) + (IIf(((txtU(Index).Top + txtU(Index).Height + picU1.Top) - picU2.Height) Mod 270 <> 0, 1, 0))) >= VU.max Then
                   VU.Value = VU.Value - ((((txtU(Index).Top + txtU(Index).Height + picU1.Top) - picU2.Height) \ 270) + (IIf(((txtU(Index).Top + txtU(Index).Height + picU1.Top) - picU2.Height) Mod 270 <> 0, 1, 0)))
               Else
                   VU.Value = VU.max
               End If
           Else
               VU.Value = VU.max
           End If
       ElseIf txtU(Index).Top + picU1.Top < 0 Then
           If VU.Value >= VU.max Then
               If VU.Value + (Abs(txtU(Index).Top + picU1.Top) / 270) >= VU.max Then
                   VU.Value = VU.Value + (Abs(txtU(Index).Top + picU1.Top) / 270)
               Else
                   VU.Value = 0
               End If
           Else
               VU.Value = 0
           End If
       End If
   End If
End Sub

Private Sub txtU_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtU_Validate(Index As Integer, Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If txtU(Index) = "" Then lblU(Index).Caption = ""
   If m_EditMode <> 0 And txtU(Index) <> "" Then
      'Modify by Morgan 2009/6/24 改電腦中心補資料時判斷人事異動，否則不可輸入離職員工
      If Pub_StrUserSt03 = "M51" Then
         lblU(Index).Caption = GetValidStaffName(txtU(Index), DBDATE(textYM01 & "1231"))
      Else
         'modify by sonia 2016/1/6 可以輸留職停薪人員
         'lblU(Index).Caption = GetStaffName(txtU(Index), IIf(m_EditMode = 1 Or m_EditMode = 2, False, True))
         lblU(Index).Caption = GetStaffNameSpec(txtU(Index), IIf(m_EditMode = 1 Or m_EditMode = 2, False, True))
      End If
      'end 2009/6/24
      If lblU(Index).Caption = "" Then
         MsgBox "員工代號錯誤！查無此員工！", vbInformation
         Cancel = True
         Exit Sub
      End If
      'add by sonia 2016/1/6 不可為當年度不得參加考績YM02='*'的人員
      strSql = "SELECT * FROM YearMerit WHERE YM01 = '" & Val(textYM01) + 1911 & "' AND YM02='*' AND YM03='" & txtU(Index) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         MsgBox "此員工 " & Val(textYM01) & "年不得參加年終考績！", vbInformation
         Cancel = True
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Sub
      End If
      rsTmp.Close
      Set rsTmp = Nothing
      'end 2016/1/6
   End If
End Sub

Private Sub VU_Change()
   picU1.Move 0, VU.Value * 270, picU1.Width, picU1.Height
End Sub
Private Sub VU_Scroll()
   picU1.Move 0, VU.Value * 270, picU1.Width, picU1.Height
End Sub
Private Sub VC_Change()
picC1.Move 0, VC.Value * 270, picC1.Width, picC1.Height
   End Sub
Private Sub VC_Scroll()
   picC1.Move 0, VC.Value * 270, picC1.Width, picC1.Height
End Sub
Private Sub VD_Change()
   picD1.Move 0, VD.Value * 270, picD1.Width, picD1.Height
End Sub
Private Sub VD_Scroll()
   picD1.Move 0, VD.Value * 270, picD1.Width, picD1.Height
End Sub

Sub ReSizePicU(oCount As Integer)
Dim Mytxt As TextBox

   For Each Mytxt In txtU
       Mytxt.Visible = False
   Next
   '將 pic 物件定義高度
   picU1.Height = (((oCount \ 3) + IIf(oCount Mod 3 >= 1, 1, 0)) * 270)
   '將捲軸定義
   If oCount > 21 Then
      VU.max = (picU2.Height - picU1.Height) / 270
      VU.Min = 0
      VU.Value = 0
      VU.Enabled = True
   Else
      VU.Enabled = False
   End If
End Sub

Sub ReSizePicC(oCount As Integer)
Dim Mytxt As TextBox

   For Each Mytxt In txtC
       Mytxt.Visible = False
   Next
   '將 pic 物件定義高度
   picC1.Height = (((oCount \ 3) + IIf(oCount Mod 3 >= 1, 1, 0)) * 270)
   '將捲軸定義
   If oCount > 9 Then
      VC.max = (picC2.Height - picC1.Height) / 270
      VC.Min = 0
      VC.Value = 0
      VC.Enabled = True
   Else
      VC.Enabled = False
   End If
End Sub

Sub ReSizePicD(oCount As Integer)
Dim Mytxt As TextBox

   For Each Mytxt In txtD
       Mytxt.Visible = False
   Next
   '將 pic 物件定義高度
   picD1.Height = (((oCount \ 3) + IIf(oCount Mod 3 >= 1, 1, 0)) * 270)
   '將捲軸定義
   If oCount > 9 Then
      VD.max = (picD2.Height - picD1.Height) / 270
      VD.Min = 0
      VD.Value = 0
      VD.Enabled = True
   Else
      VD.Enabled = False
   End If
End Sub

Private Sub cmdok_Click()
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
       If RunNick(txt1(0), txt1(1)) Then
           txt1(0).SetFocus
           Exit Sub
       End If
       If RunNick2(txt1(2), txt1(3)) Then
           txt1(2).SetFocus
           Exit Sub
       End If
       GetData
   Else
       MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
   End If
End Sub

Private Sub Form_Initialize()
   'Set rsA = New ADODB.Recordset
   'If rsA.State = 1 Then rsA.Close
   'rsA.CursorLocation = adUseClient
   'rsA.Open "select * from YearMerit where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   'tf_WF = rsA.Fields.Count
   SetGrd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
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
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textYM01.BackColor = &H8000000F
   
   MoveFormToCenter Me
   '將 pic 物件定義高度
   ReSizePicU MaxTxtU
   ReSizePicC MaxTxtC
   ReSizePicD MaxTxtD
   '將txt 物件定位
   For i = 0 To MaxTxtU - 1
      txtU(i).Top = 0 + ((i \ 3) * 270)
      txtU(i).Left = 0 + ((i Mod 3) * 2460)
      txtU(i).TabIndex = i + 1
      txtU(i).Text = ""
      txtU(i).Visible = True
      lblU(i).Top = 30 + ((i \ 3) * 270)
      lblU(i).Left = 645 + ((i Mod 3) * 2460)
      lblU(i).Caption = ""
   Next i
   For i = 0 To MaxTxtC - 1
      txtC(i).Top = 0 + ((i \ 3) * 270)
      txtC(i).Left = 0 + ((i Mod 3) * 2460)
      txtC(i).TabIndex = i + 61
      txtC(i).Text = ""
      txtC(i).Visible = True
      lblC(i).Top = 30 + ((i \ 3) * 270)
      lblC(i).Left = 645 + ((i Mod 3) * 2460)
      lblC(i).Caption = ""
   Next i
   For i = 0 To MaxTxtD - 1
      txtD(i).Top = 0 + ((i \ 3) * 270)
      txtD(i).Left = 0 + ((i Mod 3) * 2460)
      txtD(i).TabIndex = i + 121
      txtD(i).Text = ""
      txtD(i).Visible = True
      lblD(i).Top = 30 + ((i \ 3) * 270)
      lblD(i).Left = 645 + ((i Mod 3) * 2460)
      lblD(i).Caption = ""
   Next i

   '將捲軸定義
   VU.Enabled = False
   VC.Enabled = False
   VD.Enabled = False
   
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170020 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j

   GRD1.Visible = False
   tmpMouseRow = GRD1.row
   GRD1.Visible = True
   If tmpMouseRow <> 0 Then
       GRD1.row = tmpMouseRow
       GRD1.col = 0
       If GRD1.CellBackColor <> &HFFC0C0 Then
                     GRD1.Visible = False
            For j = 1 To GRD1.Rows - 1
                GRD1.row = j
                For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                Next i
           Next j
           GRD1.row = tmpMouseRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            '2008/12/12 ADD BY SONIA
            textYM01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
            QueryRecord
            '2008/12/12 END
            GRD1.Visible = True
       End If
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim iii As Integer, jjj As Integer
Dim Cancel As Boolean

   TxtValidate = False
   If Me.textYM01.Enabled = True Then
      Cancel = False
      textym01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textYM01.Text = "" Then
       MsgBox "年度不可以空白！", vbExclamation
       textYM01.SetFocus
       Exit Function
   End If
   
   For iii = 0 To MaxTxtU - 1
        If txtU(iii) <> "" Then
            For jjj = 0 To MaxTxtC - 1
                If txtU(iii) = txtC(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    Exit Function
                End If
            Next jjj
            For jjj = 0 To MaxTxtD - 1
                If txtU(iii) = txtD(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    Exit Function
                End If
            Next jjj
        End If
   Next iii
   For iii = 0 To MaxTxtC - 1
        If txtC(iii) <> "" Then
            For jjj = 0 To MaxTxtU - 1
                If txtC(iii) = txtU(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    Exit Function
                End If
            Next jjj
            For jjj = 0 To MaxTxtD - 1
                If txtC(iii) = txtD(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    Exit Function
                End If
            Next jjj
        End If
   Next iii
   For iii = 0 To MaxTxtD - 1
        If txtD(iii) <> "" Then
            For jjj = 0 To MaxTxtU - 1
                If txtD(iii) = txtU(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    Exit Function
                End If
            Next jjj
            For jjj = 0 To MaxTxtC - 1
                If txtD(iii) = txtC(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    Exit Function
                End If
            Next jjj
        End If
   Next iii

   TxtValidate = True
End Function

' 新增記錄
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strYM01 As String
Dim iii As Integer
   
   AddRecord = False
   
   '2008/12/27 modify by sonia 改存西元年
   'strYM01 = textYM01
   strYM01 = Val(textYM01) + 1911

   ' 檢查記錄是否已存在
   If IsRecordExist(strYM01) = True Then
      strTit = "新增資料"
      strMsg = "該年度記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   For iii = 0 To MaxTxtU - 1
        If txtU(iii) <> "" Then
            strSql = "insert into YearMerit (ym01,ym02,ym03) values (" & strYM01 & ",'1','" & txtU(iii) & "')"
            cnnConnection.Execute strSql
        End If
   Next iii
   For iii = 0 To MaxTxtC - 1
        If txtC(iii) <> "" Then
            strSql = "insert into YearMerit (ym01,ym02,ym03) values (" & strYM01 & ",'3','" & txtC(iii) & "')"
            cnnConnection.Execute strSql
        End If
   Next iii
   For iii = 0 To MaxTxtD - 1
        If txtD(iii) <> "" Then
            strSql = "insert into YearMerit (ym01,ym02,ym03) values (" & strYM01 & ",'4','" & txtD(iii) & "')"
            cnnConnection.Execute strSql
        End If
   Next iii
   
   'Add by Morgan 2009/6/18
   '其他的也要寫甲等
   AddRestRecord strYM01
   
   If ((strYM01) < (m_FirstKEY)) Or ((strYM01) > (m_LastKEY)) Then
      RefreshRange
   End If
      
   cnnConnection.CommitTrans
   
   ShowCurrRecord strYM01
   AddRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
    
End Function

'Add by Morgan 2009/6/18
'新增甲等考績資料
Private Sub AddRestRecord(pYear As String)
   
   'modify by sonia 2016/1/6 留職停薪人員也要新增
   'strSql = "insert into YearMerit ( ym01,ym02,ym03 ) " & _
      " select " & pYear & ",'2',sc01 from ( select sc01" & _
      ",max(decode(sign(instr('01,02',sc03)),1,sc02)) dt1" & _
      ",min(decode(sign(instr('03,04,08,09,10',sc03)),1,sc02)) dt2" & _
      " from staff_change where sc02<=" & pYear & "1231 group by sc01" & _
      " ) x,salarydata,yearmerit" & _
      " where dt1>0 and (dt2 is null or dt1>dt2)" & _
      " and sd01(+)=sc01 and sd02 in ('T','R')" & _
      " and ym03(+)=sc01 and ym01(+)=" & pYear & " and ym02 is null"
   'Modified by Morgan 2025/1/22 排除第4碼為9的
   strSql = "insert into YearMerit ( ym01,ym02,ym03 ) " & _
      " select " & pYear & ",'2',sc01 from ( select sc01,dt1,dt2,sd02,st51 from ( select sc01" & _
      ",max(decode(sign(instr('01,02',sc03)),1,sc02)) dt1" & _
      ",min(decode(sign(instr('03,08,09,10',sc03)),1,sc02)) dt2" & _
      " from staff_change where sc02<=" & pYear & "1231 and substr(sc01,4,1)<>'9' group by sc01" & _
      " ) x,salarydata,yearmerit,staff" & _
      " where dt1>0 and (dt2 is null or dt1>dt2)" & _
      " and sd01(+)=sc01 and sd02 in ('T','R','S')" & _
      " and ym03(+)=sc01 and ym01(+)=" & pYear & " and ym02 is null and sc01=st01(+)) where (nvl(st51,0)=0 or sd02='S')"
   cnnConnection.Execute strSql, intI
   'end 2009/6/18
End Sub

' 修改記錄
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strYM01 As String
Dim iii As Integer
       
   ModRecord = False
   
   strYM01 = m_CurrKEY

On Error GoTo ErrHand
   cnnConnection.BeginTrans
        
   'modify by sonia 2016/1/6
   'cnnConnection.Execute "delete from yearmerit where ym01=" & strYM01 & " "
   cnnConnection.Execute "delete from yearmerit where ym01=" & strYM01 & " and ym02<>'*'"
   'end 2016/1/6
   
   For iii = 0 To MaxTxtU - 1
       If txtU(iii) <> "" Then
           strSql = "insert into YearMerit (ym01,ym02,ym03) values (" & strYM01 & ",'1','" & txtU(iii) & "')"
           cnnConnection.Execute strSql
       End If
   Next iii
   For iii = 0 To MaxTxtC - 1
       If txtC(iii) <> "" Then
           strSql = "insert into YearMerit (ym01,ym02,ym03) values (" & strYM01 & ",'3','" & txtC(iii) & "')"
           cnnConnection.Execute strSql
       End If
   Next iii
   For iii = 0 To MaxTxtD - 1
       If txtD(iii) <> "" Then
           strSql = "insert into YearMerit (ym01,ym02,ym03) values (" & strYM01 & ",'4','" & txtD(iii) & "')"
           cnnConnection.Execute strSql
       End If
   Next iii
     
   'Add by Morgan 2009/6/18
   '其他的也要寫甲等
   AddRestRecord strYM01

   cnnConnection.CommitTrans

   ShowCurrRecord strYM01
   
   ModRecord = True
   Exit Function

ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strYM01 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strYM01 = m_CurrKEY

   'modify by sonia 2016/1/6
   'strSql = "DELETE FROM YearMerit WHERE YM01 = '" & strYM01 & "' "
   strSql = "DELETE FROM YearMerit WHERE YM01 = '" & strYM01 & "' and ym02<>'*'"
   'end 2016/1/6

   cnnConnection.Execute strSql

   If (strYM01 = m_LastKEY) Or (strYM01 = m_FirstKEY) Then
      RefreshRange
   End If
   ShowCurrRecord strYM01
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strYM01 As String
   
   QueryRecord = False
   '2008/12/27 modify by sonia
   'strYM01 = textYM01
   strYM01 = Val(textYM01) + 1911
   
   If IsRecordExist(strYM01) = True Then
      m_CurrKEY = strYM01
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textYM01 <> "" Then
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
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then textYM01.SetFocus
      Case 2: If Me.Visible = True Then txtU(0).SetFocus
      Case 4: If Me.Visible = True Then textYM01.SetFocus
   End Select
End Sub
' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   'modify by sonia 2016/1/6
   'strSql = "SELECT * FROM YearMerit " & _
            "WHERE YM01 = '" & strKEY01 & "' "
   strSql = "SELECT * FROM YearMerit WHERE YM01 = '" & strKEY01 & "' and YM02<>'*'"
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY = strKEY01
   Else
      strSql = "SELECT YM01 FROM YearMerit " & _
               "WHERE YM01 = '" & m_CurrKEY & "'  "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("YM01")) = False Then: m_CurrKEY = rsTmp.Fields("YM01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT YM01 FROM YearMerit " & _
               "WHERE YM01=(select min(YM01) from YearMerit) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("YM01")) = False Then: m_CurrKEY = rsTmp.Fields("YM01")
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
   m_CurrKEY = m_FirstKEY
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY = m_FirstKEY Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT YM01 FROM YearMerit " & _
            "WHERE YM01 = (SELECT MAX(YM01) FROM YearMerit " & _
                           "WHERE YM01 < '" & m_CurrKEY & "')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YM01")) = False Then: m_CurrKEY = rsTmp.Fields("YM01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT YM01 FROM YearMerit " & _
            "WHERE YM01 = (SELECT min(YM01) FROM YearMerit) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YM01")) = False Then: m_CurrKEY = rsTmp.Fields("YM01")
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
   
   If m_CurrKEY = m_LastKEY Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT YM01 FROM YearMerit " & _
            "WHERE YM01 = (SELECT MIN(YM01) FROM YearMerit " & _
                           "WHERE YM01 > '" & m_CurrKEY & "')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YM01")) = False Then: m_CurrKEY = rsTmp.Fields("YM01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT YM01 FROM YearMerit " & _
            "WHERE YM01 = (SELECT max(YM01) FROM YearMerit ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YM01")) = False Then: m_CurrKEY = rsTmp.Fields("YM01")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY = m_LastKEY
   UpdateCtrlData
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
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
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
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
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
         If OnWork = True Then
            Me.SSTab1.TabEnabled(1) = True
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  Me.SSTab1.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.SSTab1.TabEnabled(1) = True
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT YM01 FROM YearMerit " & _
            "WHERE YM01 = (SELECT MIN(YM01) FROM YearMerit) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YM01")) = False Then: m_FirstKEY = rsTmp.Fields("YM01")
   End If
   rsTmp.Close

   strSql = "SELECT YM01 FROM YearMerit " & _
            "WHERE YM01 = (SELECT MAX(YM01) FROM YearMerit) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YM01")) = False Then: m_LastKEY = rsTmp.Fields("YM01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim rsTmp2 As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
      
   ClearField
   strSql = "SELECT * FROM YearMerit " & _
            "WHERE YM01='" & m_CurrKEY & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      '2008/12/27 modify by sonia
      'If IsNull(rsTmp.Fields("YM01")) = False Then: textYM01 = rsTmp.Fields("YM01")
      If IsNull(rsTmp.Fields("YM01")) = False Then: textYM01 = Val(rsTmp.Fields("YM01")) - 1911
   End If

   '優等
   strSql = "select * from yearmerit where ym01='" & m_CurrKEY & "' and ym02='1' order by ym03 "
   If rsTmp2.State = 1 Then rsTmp2.Close
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   NowMaxU = 0
   picU1.Top = 0
   VU.Value = 0
   If rsTmp2.RecordCount > 0 Then
         NowMaxU = rsTmp2.RecordCount
         If NowMaxU > MaxTxtU Then NowMaxU = MaxTxtU
         ReSizePicU NowMaxU
         rsTmp2.MoveFirst
         Do While Not rsTmp2.EOF
             txtU(rsTmp2.AbsolutePosition - 1).Visible = True
             txtU(rsTmp2.AbsolutePosition - 1).Text = CheckStr(rsTmp2.Fields("ym03"))
             lblU(rsTmp2.AbsolutePosition - 1).Caption = GetStaffName(txtU(rsTmp2.AbsolutePosition - 1), True)
             rsTmp2.MoveNext
         Loop
   Else
         '將 pic 物件定義高度
         ReSizePicU MaxTxtU
         '將捲軸定義
         VU.Enabled = False
   End If
   '乙等
   strSql = "select * from yearmerit where ym01='" & m_CurrKEY & "' and ym02='3' order by ym03 "
   If rsTmp2.State = 1 Then rsTmp2.Close
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   NowMaxC = 0
   picC1.Top = 0
   VC.Value = 0
   If rsTmp2.RecordCount > 0 Then
         NowMaxC = rsTmp2.RecordCount
         If NowMaxC > MaxTxtC Then NowMaxC = MaxTxtC
         ReSizePicC NowMaxC
         rsTmp2.MoveFirst
         Do While Not rsTmp2.EOF
             txtC(rsTmp2.AbsolutePosition - 1).Visible = True
             txtC(rsTmp2.AbsolutePosition - 1).Text = CheckStr(rsTmp2.Fields("ym03"))
             lblC(rsTmp2.AbsolutePosition - 1).Caption = GetStaffName(txtC(rsTmp2.AbsolutePosition - 1), True)
             rsTmp2.MoveNext
         Loop
   Else
         '將 pic 物件定義高度
         ReSizePicC MaxTxtC
         '將捲軸定義
         VC.Enabled = False
   End If
   '丙等
   strSql = "select * from yearmerit where ym01='" & m_CurrKEY & "' and ym02='4' order by ym03 "
   If rsTmp2.State = 1 Then rsTmp2.Close
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   NowMaxD = 0
   picD1.Top = 0
   VD.Value = 0
   If rsTmp2.RecordCount > 0 Then
         NowMaxD = rsTmp2.RecordCount
         If NowMaxD > MaxTxtD Then NowMaxD = MaxTxtD
         ReSizePicD NowMaxD
         rsTmp2.MoveFirst
         Do While Not rsTmp2.EOF
             txtD(rsTmp2.AbsolutePosition - 1).Visible = True
             txtD(rsTmp2.AbsolutePosition - 1).Text = CheckStr(rsTmp2.Fields("ym03"))
             lblD(rsTmp2.AbsolutePosition - 1).Caption = GetStaffName(txtD(rsTmp2.AbsolutePosition - 1), True)
             rsTmp2.MoveNext
         Loop
   Else
         '將 pic 物件定義高度
         ReSizePicD MaxTxtD
         '將捲軸定義
         VD.Enabled = False
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
   
   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and ym03>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and ym03<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       '2008/12/27 modify by sonia
       'strSQL = strSQL & " and ym01>=" & txt1(2) & " "
       strSql = strSql & " and ym01>=" & Val(txt1(2)) + 1911 & " "
   End If
   If txt1(3) <> "" Then
       '2008/12/27 modify by sonia
       'strSQL = strSQL & " and ym01<=" & txt1(3) & " "
       strSql = strSql & " and ym01<=" & Val(txt1(3)) + 1911 & " "
   End If
   '抓取資料
   '2008/12/27 modify by sonia
   'strSQL = "SELECT ym01,ym02||' '||decode(ym02,'1','優','3','乙','4','丙',''),ym03,st02 FROM yearmerit,staff where ym03=st01(+) " & strSQL & _
            " order by ym01,ym02,ym03 "
   'Modify by Morgan 2009/6/19 已改甲等存所以要排除
   'modify by sonia 2016/1/6
   'strSql = "SELECT ym01-1911,ym02||' '||decode(ym02,'1','優','3','乙','4','丙',''),ym03,st02 FROM yearmerit,staff where ym03=st01(+) and ym02<>'2' " & strSql & _
           " order by ym01,ym02,ym03 "
   strSql = "SELECT ym01-1911,ym02||' '||decode(ym02,'1','優','3','乙','4','丙','*','不得參加',''),ym03,st02 FROM yearmerit,staff where ym03=st01(+) and ym02<>'2' " & strSql & _
           " order by ym01,ym02,ym03 "
   'end 2016/1/6
   '2008/12/27 end
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   SetGrd
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
   
End Sub

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTmp  As String
Dim iii As Integer, jjj As Integer
   
   CheckDataValid = False
   
   nResponse = False
   textym01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   For iii = 0 To MaxTxtU - 1
        nResponse = False
        txtU_Validate iii, nResponse
        If nResponse = True Then GoTo EXITSUB
        If txtU(iii) <> "" Then
            For jjj = 0 To MaxTxtU - 1
                If txtU(iii) = txtU(jjj) And iii <> jjj Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
            For jjj = 0 To MaxTxtC - 1
                If txtU(iii) = txtC(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
            For jjj = 0 To MaxTxtD - 1
                If txtU(iii) = txtD(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
        End If
   Next iii
   For iii = 0 To MaxTxtC - 1
        nResponse = False
        txtC_Validate iii, nResponse
        If nResponse = True Then GoTo EXITSUB
        If txtC(iii) <> "" Then
            For jjj = 0 To MaxTxtU - 1
                If txtC(iii) = txtU(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
            For jjj = 0 To MaxTxtC - 1
                If txtC(iii) = txtC(jjj) And iii <> jjj Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
            For jjj = 0 To MaxTxtD - 1
                If txtC(iii) = txtD(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
        End If
   Next iii
   For iii = 0 To MaxTxtD - 1
        nResponse = False
        txtD_Validate iii, nResponse
        If nResponse = True Then GoTo EXITSUB
        If txtD(iii) <> "" Then
            For jjj = 0 To MaxTxtU - 1
                If txtD(iii) = txtU(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
            For jjj = 0 To MaxTxtC - 1
                If txtD(iii) = txtC(jjj) Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
            For jjj = 0 To MaxTxtD - 1
                If txtD(iii) = txtD(jjj) And iii <> jjj Then
                    MsgBox "輸入的員工代號重複，請詳細檢查！", vbCritical, "操作錯誤！"
                    GoTo EXITSUB
                End If
            Next jjj
        End If
   Next iii
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textYM01.Locked = bEnable
   If bEnable Then textYM01.BackColor = &H8000000F Else textYM01.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
Dim Mytxt As TextBox
   textYM01.Locked = bEnable
   If bEnable Then textYM01.BackColor = &H8000000F Else textYM01.BackColor = &H80000005
   If bEnable = False Then
       If m_EditMode = 1 Or m_EditMode = 2 Then
           ReSizePicU MaxTxtU
           ReSizePicC MaxTxtC
           ReSizePicD MaxTxtD
       Else
           ReSizePicU NowMaxU
           ReSizePicC NowMaxC
           ReSizePicD NowMaxD
       End If
   End If
       
   For Each Mytxt In txtU
       Mytxt.Locked = bEnable
       If bEnable = False Then
           If m_EditMode = 1 Or m_EditMode = 2 Then
               If Mytxt.Index < MaxTxtU Then
                   Mytxt.Visible = True
               End If
           Else
               If Mytxt.Index < NowMaxU Then
                   Mytxt.Visible = False
               End If
           End If
       Else
           If Mytxt.Text = "" Then
               If Mytxt.Index < NowMaxU Then
                   Mytxt.Visible = False
               End If
           End If
       End If
   Next
   For Each Mytxt In txtC
       Mytxt.Locked = bEnable
       If bEnable = False Then
           If m_EditMode = 1 Or m_EditMode = 2 Then
               If Mytxt.Index < MaxTxtC Then
                   Mytxt.Visible = True
               End If
           Else
               If Mytxt.Index < NowMaxC Then
                   Mytxt.Visible = False
               End If
           End If
       Else
           If Mytxt.Text = "" Then
               If Mytxt.Index < NowMaxC Then
                   Mytxt.Visible = False
               End If
           End If
       End If
   Next
   For Each Mytxt In txtD
       Mytxt.Locked = bEnable
       If bEnable = False Then
           If m_EditMode = 1 Or m_EditMode = 2 Then
               If Mytxt.Index < MaxTxtD Then
                   Mytxt.Visible = True
               End If
           Else
               If Mytxt.Index < NowMaxD Then
                   Mytxt.Visible = False
               End If
           End If
       Else
           If Mytxt.Text = "" Then
               If Mytxt.Index < NowMaxD Then
                   Mytxt.Visible = False
               End If
           End If
       End If
   Next
End Sub

Private Sub ClearField()
Dim Mytxt As TextBox

   textYM01 = Empty
   For Each Mytxt In txtU
       If Mytxt.Index < MaxTxtU Then
           Mytxt = Empty
           lblU(Mytxt.Index).Caption = ""
       End If
   Next
   For Each Mytxt In txtC
       If Mytxt.Index < MaxTxtC Then
           Mytxt = Empty
           lblC(Mytxt.Index).Caption = ""
       End If
   Next
   For Each Mytxt In txtD
       If Mytxt.Index < MaxTxtD Then
           Mytxt = Empty
           lblD(Mytxt.Index).Caption = ""
       End If
   Next

   SetGrd
End Sub

'帶預設資料
Private Sub InitialData()
Dim ii As Integer
   'For ii = 0 To MaxTxtU - 1
   '    txtU(ii).Visible = True
   'Next ii
   'For ii = 0 To MaxTxtC - 1
   '    txtC(ii).Visible = True
   'Next ii
   'For ii = 0 To MaxTxtD - 1
   '    txtD(ii).Visible = True
   'Next ii
   'ReSizePicU MaxTxtU
   'ReSizePicC MaxTxtC
   'ReSizePicD MaxTxtD
   SetGrd
End Sub

Private Sub textym01_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textYM01
   End If
End Sub

Private Sub textym01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textym01_Validate(Cancel As Boolean)

   If m_EditMode = 1 And textYM01 <> "" Then
       '2008/12/27 modify by sonia
       'If IsRecordExist(textYM01) = True And textYM01.Enabled = True And textYM01.Locked = False Then
       If IsRecordExist(Val(textYM01) + 1911) = True And textYM01.Enabled = True And textYM01.Locked = False Then
           MsgBox "當年度已有資料，請修改！", vbInformation
           Cancel = True
           Exit Sub
       End If
       If CheckIsTaiwanDate(textYM01 & "0101", False) = False Then
           Cancel = True
           MsgBox "請輸入民國年度！", vbInformation, "輸入年度錯誤"
           Exit Sub
       End If
   
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("年度", "考績", "編號", "姓名")
   arrGridHeadWidth = Array(600, 1000, 800, 1200)
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

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = UpperCase(KeyAscii)
      Case 2, 3
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)

   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 1
           If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
               Exit Sub
           End If
      Case 2, 3
           If CheckIsTaiwanDate(txt1(Index) & "0101", False) = False Then
               Cancel = True
               MsgBox "請輸入民國年度！", vbInformation, "輸入年度錯誤"
               Exit Sub
           End If
           If Index = 3 Then
               If RunNick2(txt1(Index - 1), txt1(Index)) Then
                   Cancel = True
                   Exit Sub
               End If
           End If
      Case Else
   End Select
End Sub

'Add by Morgan 2009/6/19
Private Function GetValidStaffName(pST01 As String, Optional ByVal pDate As String) As String
Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
   
   If pDate = "" Then
      pDate = strSrvDate(1)
   End If
   
   'MODIFY BY SONIA 2016/1/6 取消SC03='04'留職停薪
   stSQL = " select st02 from ( select sc01" & _
      ",max(decode(sign(instr('01,02',sc03)),1,sc02)) dt1" & _
      ",min(decode(sign(instr('03,08,09,10',sc03)),1,sc02)) dt2" & _
      " from staff_change where sc01='" & pST01 & "' and sc02<=" & pDate & " group by sc01" & _
      " ) x,staff" & _
      " where dt1>0 and (dt2 is null or dt1>dt2)" & _
      " and st01(+)=sc01"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetValidStaffName = "" & adoRst(0)
   End If
   Set adoRst = Nothing
   
End Function

'add by sonia 2016/1/6 含留職停薪人員
Public Function GetStaffNameSpec(ByVal strStuff As String, Optional ByVal bAll As Boolean = False) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetStaffNameSpec = Empty
   
   strSql = "SELECT ST02,DECODE(SD02,'S','1',ST04) ST04 FROM Staff,salarydata " & _
            "WHERE ST01 = '" & strStuff & "' and st01=sd01(+) and (st04='1' or (st04<>'1' and sd02='S')) and sd02 not in ('P','F')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("ST02")) = False Then
         GetStaffNameSpec = rsTmp.Fields("ST02")
      End If
      If bAll = False Then
         If IsNull(rsTmp.Fields("ST04")) = False Then
            If rsTmp.Fields("ST04") = "2" Then
               GetStaffNameSpec = Empty
            End If
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
'end 2016/1/6
