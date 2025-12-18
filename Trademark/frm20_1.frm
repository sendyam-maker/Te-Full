VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm20_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "商品名稱查詢--整批新增"
   ClientHeight    =   4980
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7380
   Begin TabDlg.SSTab stb 
      Height          =   4284
      Left            =   144
      TabIndex        =   3
      Top             =   600
      Width           =   7116
      _ExtentX        =   12541
      _ExtentY        =   7567
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      TabCaption(0)   =   "國際分類"
      TabPicture(0)   =   "frm20_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "處理狀況"
      TabPicture(1)   =   "frm20_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lst"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.ListBox lst 
         Height          =   2940
         ItemData        =   "frm20_1.frx":0038
         Left            =   -74784
         List            =   "frm20_1.frx":003A
         TabIndex        =   48
         Top             =   744
         Width           =   6708
      End
      Begin VB.Frame fra 
         Height          =   3516
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   528
         Width           =   6684
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   45
            Left            =   5370
            TabIndex        =   52
            Top             =   1692
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class45.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   44
            Left            =   5370
            TabIndex        =   51
            Top             =   1404
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class44.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   43
            Left            =   5370
            TabIndex        =   50
            Top             =   1116
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class43.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   0
            Left            =   216
            TabIndex        =   47
            Top             =   264
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   1
            Left            =   216
            TabIndex        =   46
            Top             =   548
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class01.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   2
            Left            =   216
            TabIndex        =   45
            Top             =   832
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class02.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   3
            Left            =   216
            TabIndex        =   44
            Top             =   1116
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class03.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   4
            Left            =   216
            TabIndex        =   43
            Top             =   1400
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class04.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   5
            Left            =   216
            TabIndex        =   42
            Top             =   1684
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class05.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   6
            Left            =   216
            TabIndex        =   41
            Top             =   1968
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class06.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   7
            Left            =   216
            TabIndex        =   40
            Top             =   2252
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class07.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   8
            Left            =   216
            TabIndex        =   39
            Top             =   2536
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class08.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   9
            Left            =   216
            TabIndex        =   38
            Top             =   2820
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class09.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   10
            Left            =   216
            TabIndex        =   37
            Top             =   3096
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class10.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   11
            Left            =   1512
            TabIndex        =   36
            Top             =   552
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class11.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   12
            Left            =   1512
            TabIndex        =   35
            Top             =   840
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class12.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   13
            Left            =   1512
            TabIndex        =   34
            Top             =   1116
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class13.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   14
            Left            =   1512
            TabIndex        =   33
            Top             =   1404
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class14.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   15
            Left            =   1512
            TabIndex        =   32
            Top             =   1692
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class15.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   16
            Left            =   1512
            TabIndex        =   31
            Top             =   1968
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class16.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   17
            Left            =   1512
            TabIndex        =   30
            Top             =   2256
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class17.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   18
            Left            =   1512
            TabIndex        =   29
            Top             =   2544
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class18.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   19
            Left            =   1512
            TabIndex        =   28
            Top             =   2820
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class19.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   20
            Left            =   1512
            TabIndex        =   27
            Top             =   3096
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class20.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   21
            Left            =   2808
            TabIndex        =   26
            Top             =   552
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class21.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   22
            Left            =   2808
            TabIndex        =   25
            Top             =   840
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class22.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   23
            Left            =   2808
            TabIndex        =   24
            Top             =   1116
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class23.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   24
            Left            =   2808
            TabIndex        =   23
            Top             =   1404
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class24.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   25
            Left            =   2808
            TabIndex        =   22
            Top             =   1692
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class25.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   26
            Left            =   2808
            TabIndex        =   21
            Top             =   1968
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class26.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   27
            Left            =   2808
            TabIndex        =   20
            Top             =   2256
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class27.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   28
            Left            =   2808
            TabIndex        =   19
            Top             =   2544
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class28.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   29
            Left            =   2808
            TabIndex        =   18
            Top             =   2820
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class29.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   30
            Left            =   2808
            TabIndex        =   17
            Top             =   3096
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class30.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   31
            Left            =   4128
            TabIndex        =   16
            Top             =   552
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class31.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   32
            Left            =   4128
            TabIndex        =   15
            Top             =   840
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class32.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   33
            Left            =   4128
            TabIndex        =   14
            Top             =   1116
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class33.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   34
            Left            =   4128
            TabIndex        =   13
            Top             =   1404
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class34.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   35
            Left            =   4128
            TabIndex        =   12
            Top             =   1692
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class35.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   36
            Left            =   4128
            TabIndex        =   11
            Top             =   1968
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class36.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   37
            Left            =   4128
            TabIndex        =   10
            Top             =   2256
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class37.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   38
            Left            =   4128
            TabIndex        =   9
            Top             =   2544
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class38.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   39
            Left            =   4128
            TabIndex        =   8
            Top             =   2820
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class39.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   40
            Left            =   4128
            TabIndex        =   7
            Top             =   3096
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class40.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   41
            Left            =   5376
            TabIndex        =   6
            Top             =   552
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class41.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox chk 
            Height          =   285
            Index           =   42
            Left            =   5376
            TabIndex        =   5
            Top             =   840
            Width           =   1140
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2037;503"
            Value           =   "0"
            Caption         =   "Class42.txt"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "處理過程："
         Height          =   180
         Index           =   0
         Left            =   -74736
         TabIndex        =   49
         Top             =   408
         Width           =   900
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "全選／取消(&C)"
      Height          =   348
      Index           =   2
      Left            =   3048
      TabIndex        =   2
      Top             =   48
      Width           =   1356
   End
   Begin VB.CommandButton cmd 
      Caption         =   "回前畫面(&X)"
      Height          =   348
      Index           =   0
      Left            =   5424
      TabIndex        =   1
      Top             =   48
      Width           =   1110
   End
   Begin VB.CommandButton cmd 
      Caption         =   "確定(&B)"
      Default         =   -1  'True
      Height          =   348
      Index           =   1
      Left            =   4416
      TabIndex        =   0
      Top             =   48
      Width           =   996
   End
End
Attribute VB_Name = "frm20_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; chk(index)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim m_blnChkStatus As Boolean

Private Sub cmd_Click(Index As Integer)
    Dim ii As Integer
    Dim blnChk As Boolean
    Dim StrSQLa As String
    Dim jj As Double
    Dim blnBeginTrans As Boolean
   
    On Error GoTo ErrorHandler
    blnBeginTrans = False
    Select Case Index
    Case 0 '離開
       Unload Me
    Case 1 '確定
        '整批新增
        If Me.Tag = "1" Then
            blnChk = False
            For ii = 0 To Me.chk.Count - 1
                'Modified by Lydia 2021/09/24 改成Form 2.0
                'If Me.chk(ii).Value = vbChecked Then
                If Me.chk(ii).Value = True Then
                    blnChk = True
                    Exit For
                End If
            Next ii
            If blnChk = True Then
                Me.stb.Tab = 1
                '整批新增動作開始
                Me.Enabled = False
                Screen.MousePointer = vbHourglass
                BatchProcess
                Screen.MousePointer = vbDefault
                Me.Enabled = True
            Else
                MsgBox "請至少勾選一項欲新增國際分類資料!!!", vbExclamation + vbOKOnly
            End If
        '整批刪除
        ElseIf Me.Tag = "8" Then
            blnChk = False
            For ii = 0 To Me.chk.Count - 1
                'Modified by Lydia 2021/09/24 改成Form 2.0
                'If Me.chk(ii).Value = vbChecked Then
                If Me.chk(ii).Value = True Then
                    blnChk = True
                    Exit For
                End If
            Next ii
            If blnChk = True Then
                '整批刪除動作開始
                Me.Enabled = False
                Screen.MousePointer = vbHourglass
                For ii = 0 To Me.chk.Count - 1
                    'Modified by Lydia 2021/09/24 改成Form 2.0
                    'If Me.chk(ii).Value = vbChecked Then
                    If Me.chk(ii).Value = True Then
                        If MsgBox("是否刪除第 " & ii & " 類的資料!!!", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                            cnnConnection.BeginTrans
                            blnBeginTrans = True
                            If ii = 0 Then
                                StrSQLa = "Delete From TrademarkMerchandiseName Where TMN01 Is Null "
                            Else
                                StrSQLa = "Delete From TrademarkMerchandiseName Where TMN01='" & Format(ii, "00") & "' "
                            End If
                            cnnConnection.Execute StrSQLa, jj
                            cnnConnection.CommitTrans
                            blnBeginTrans = False
                            MsgBox "第 " & ii & " 類的資料計 " & Format(jj, "#,##0") & " 筆已刪除完畢!!!", vbExclamation + vbOKOnly
                        End If
                    End If
                Next ii
                Screen.MousePointer = vbDefault
                Me.Enabled = True
            Else
                MsgBox "請至少勾選一項欲刪除的國際分類資料!!!", vbExclamation + vbOKOnly
            End If
        End If
    Case 2 '全選／取消
        Me.stb.Tab = 0
        If m_blnChkStatus = False Then
            m_blnChkStatus = True
            For ii = 0 To Me.chk.Count - 1
                'Modified by Lydia 2021/09/24 改成Form 2.0
                'Me.chk(ii).Value = vbChecked
                Me.chk(ii).Value = True
            Next ii
        Else
            m_blnChkStatus = False
            For ii = 0 To Me.chk.Count - 1
                Me.chk(ii).Value = vbUnchecked
            Next ii
        End If
    End Select
    Exit Sub
ErrorHandler:
    If blnBeginTrans = True Then cnnConnection.RollbackTrans
    blnBeginTrans = False
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox Err.Description
End Sub

Private Sub BatchProcess()
Dim ii As Integer
Dim jj As Integer
Dim kk As Double
Dim ll As Double
Dim mm As Integer
Dim fs As New FileSystemObject
Dim txtf As TextStream
Dim s As String
Dim col
Dim strClass As String '國際分類
Dim strChnName As String '商品中文名稱
Dim strEngName As String '商品英文名稱
'add by nick 2004/10/12
Dim strJpnName As String '商品日文名稱
Dim blnProcessAdd As Boolean

On Error GoTo ErrorHandler

   Me.lst.Clear
   For ii = 0 To Me.chk.Count - 1
      'Modified by Lydia 2021/09/24 改成Form 2.0
      'If Me.chk(ii).Value = vbChecked Then
      If Me.chk(ii).Value = True Then
         kk = 0
         ll = 0
         strClass = IIf(ii = 0, "", Format(ii, "00"))
         '刪除錯誤記錄檔
         DeleteErrLogFile "ERR" & strClass & ".txt"
         '若國際分類的文字檔存在才處理
         If fs.FileExists(App.path & "\" & Me.chk(ii).Caption) Then
            AddLogOnList "目前準備處理" & Me.chk(ii).Caption & "...": DoEvents
            '開啟國際分類文字檔
            Set txtf = fs.OpenTextFile(App.path & "\" & Me.chk(ii).Caption)
            While Not txtf.AtEndOfStream
               kk = kk + 1
               s = Trim(txtf.ReadLine)
               If s <> "" Then
                  col = Split(s, ":")
                  strChnName = ""
                  strEngName = ""
                  'add by nick 2004/10/12
                  strJpnName = ""
                  strEngName = "" & col(0)
                  '若資料無分隔號(:)時的處理
                  If LBound(col) = UBound(col) Then
                     AddLogOnList "***ERR***" & "資料無分隔號(:)" & "...": DoEvents
                     AddLogOnList Left(strClass & s, 50) & "...": DoEvents
                     AddLogOnList ""
                     '將有問題的資料寫入Log檔
                     WriteErrLogFile "ERR" & strClass & ".txt", s
                     GoTo NextLine
                  End If
                  strChnName = "" & col(1)
                  'add by nick 2004/10/12
                  If UBound(col) = 2 Then
                        strJpnName = "" & col(2)
                  End If
                  blnProcessAdd = False
                  If CheckDataRepeat(strClass, strChnName, strEngName, strJpnName) = False Then
                     blnProcessAdd = True
                     'edit by nick 2004/10/12
                     'strSQL = "INSERT INTO TRADEMARKMERCHANDISENAME (tmn01,tmn02,tmn03,tmn04) VALUES('" & strClass & "','" & ChgSQL(strChnName) & "','" & ChgSQL(strEngName) & "')"
                     'Modified Lydia 2022/03/15 +流水號tmn05
                     'strSql = "INSERT INTO TRADEMARKMERCHANDISENAME (tmn01,tmn02,tmn03,tmn04) VALUES('" & strClass & "','" & ChgSQL(strChnName) & "','" & ChgSQL(strEngName) & "','" & ChgSQL(strJpnName) & "')"
                     strSql = "INSERT INTO TRADEMARKMERCHANDISENAME (tmn01,tmn02,tmn03,tmn04,tmn05) VALUES('" & strClass & "','" & ChgSQL(strChnName) & "','" & ChgSQL(strEngName) & "','" & ChgSQL(strJpnName) & "', " & GetMaxNo & " )"
                     cnnConnection.Execute strSql, mm
                     ll = ll + mm
                  End If
NextLine:
                  Me.Label1(0).Caption = "處理狀況：讀取" & Me.chk(ii).Caption & "筆數 " & Format(kk, "#,##0") & " 筆, 新增筆數 " & Format(ll, "#,##0") & " 筆!!!": DoEvents
               End If
            Wend
            AddLogOnList Me.chk(ii).Caption & "筆數 " & Format(kk, "#,##0") & " 筆, 新增筆數 " & Format(ll, "#,##0") & " 筆!!!": DoEvents
         Else
            Me.Label1(0).Caption = "處理狀況：": DoEvents
            AddLogOnList "目錄中並無 " & Me.chk(ii).Caption & " 商品名稱文字檔!!!": DoEvents
         End If
      End If
   Next ii
   Me.Label1(0).Caption = "處理狀況：作業完成!!!"
   AddLogOnList "整批新增作業完成!!!": DoEvents
   Exit Sub

ErrorHandler:
   AddLogOnList Left("***ERR***" & Err.Description, 50) & "...": DoEvents
   AddLogOnList Left(strClass & s, 50) & "...": DoEvents
   AddLogOnList ""
   '將有問題的資料寫入Log檔
   WriteErrLogFile "ERR" & strClass & ".txt", s
   Resume Next
End Sub

Private Sub WriteErrLogFile(strFileName As String, str As String)
On Error Resume Next
   
   Open App.path & "\" & strFileName For Append As #10
   Print #10, str
   Close #10

End Sub

Private Sub DeleteErrLogFile(strFileName As String)
On Error Resume Next

   Kill App.path & "\" & strFileName

End Sub

Private Sub AddLogOnList(strLog As String)
   Me.lst.AddItem strLog
   Me.lst.Selected(Me.lst.ListCount - 1) = True
End Sub

'edit by nick 2004/10/12
'Private Function CheckDataRepeat(strTMN01 As String, strTMN02 As String, strTMN03 As String) As Boolean
Private Function CheckDataRepeat(strTMN01 As String, strTMN02 As String, strTMN03 As String, strTMN04 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

On Error GoTo ErrorHandler

StrSQLa = IIf(strTMN01 = "", " TMN01 IS NULL ", " TMN01 = '" & strTMN01 & "' ")
StrSQLa = StrSQLa & IIf(strTMN02 = "", " AND TMN02 IS NULL ", " AND TMN02 = '" & ChgSQL(strTMN02) & "' ")
StrSQLa = StrSQLa & IIf(strTMN03 = "", " AND TMN03 IS NULL ", " AND TMN03 = '" & ChgSQL(strTMN03) & "' ")
'add by nick 2004/10/12
StrSQLa = StrSQLa & IIf(strTMN04 = "", " AND TMN04 IS NULL ", " AND TMN04 = '" & ChgSQL(strTMN04) & "' ")
StrSQLa = "Select * From TRADEMARKMERCHANDISENAME WHERE " & StrSQLa
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   CheckDataRepeat = True
Else
   CheckDataRepeat = False
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
Exit Function

ErrorHandler:
   CheckDataRepeat = False
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   m_blnChkStatus = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frm20.Show
   Set frm20_1 = Nothing
End Sub

'Added by Lydia 2022/03/15
Private Function GetMaxNo() As String
Dim strQ As String, intQ As Integer
Dim RsQ As New ADODB.Recordset

    strQ = "select nvl(max(tmn05),0)+1 as mno from trademarkmerchandisename "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetMaxNo = "" & RsQ.Fields("mno")
    End If
    Set RsQ = Nothing
End Function

