VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090128 
   BorderStyle     =   1  '單線固定
   Caption         =   "查覆明細作業"
   ClientHeight    =   7464
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8988
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7464
   ScaleWidth      =   8988
   Begin VB.CommandButton cmdRoute 
      BackColor       =   &H00C0FFFF&
      Caption         =   "輸入"
      Height          =   300
      Left            =   4368
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   1992
      Width           =   720
   End
   Begin VB.CommandButton cmdSendMail 
      BackColor       =   &H00C0FFFF&
      Caption         =   "通知送件"
      Height          =   345
      Left            =   4306
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   39
      Top             =   40
      Width           =   900
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   23
      Left            =   8064
      MaxLength       =   9
      TabIndex        =   20
      Text            =   "23"
      Top             =   1968
      Width           =   860
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   19
      Left            =   6240
      MaxLength       =   9
      TabIndex        =   19
      Text            =   "19"
      Top             =   1968
      Width           =   860
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   2
      Left            =   7272
      MaxLength       =   6
      TabIndex        =   5
      Text            =   "2"
      Top             =   710
      Width           =   675
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   5
      Left            =   5352
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "5"
      Top             =   710
      Width           =   900
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   4
      Left            =   7272
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "4"
      Top             =   448
      Width           =   900
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   3
      Left            =   5352
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "3"
      Top             =   448
      Width           =   900
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   0
      Left            =   2952
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "0"
      Top             =   448
      Width           =   675
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   11
      Left            =   1032
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "11"
      Top             =   1968
      Width           =   3300
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一張單(&P)"
      Height          =   345
      Left            =   2040
      Style           =   1  '圖片外觀
      TabIndex        =   37
      Top             =   40
      Width           =   1100
   End
   Begin VB.PictureBox tmpPic 
      Height          =   4215
      Left            =   9500
      ScaleHeight     =   347
      ScaleMode       =   3  '像素
      ScaleWidth      =   295
      TabIndex        =   99
      Top             =   960
      Visible         =   0   'False
      Width           =   3585
      Begin VB.Image tmpImg 
         Height          =   1770
         Left            =   1425
         Stretch         =   -1  'True
         Top             =   1095
         Width           =   1890
      End
      Begin VB.Image tmpInsPDF 
         Height          =   372
         Left            =   0
         Picture         =   "frm090128.frx":0000
         Top             =   2880
         Visible         =   0   'False
         Width           =   1296
      End
   End
   Begin VB.PictureBox G_SeekPicColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2415
      Left            =   9600
      ScaleHeight     =   197
      ScaleMode       =   3  '像素
      ScaleWidth      =   246
      TabIndex        =   92
      Top             =   240
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox txtField 
      Height          =   280
      Index           =   9
      Left            =   8112
      TabIndex        =   9
      Text            =   "9"
      Top             =   1013
      Width           =   435
   End
   Begin VB.TextBox txtField 
      Height          =   280
      Index           =   8
      Left            =   6792
      TabIndex        =   8
      Text            =   "8"
      Top             =   1013
      Width           =   435
   End
   Begin VB.TextBox txtField 
      Height          =   280
      Index           =   7
      Left            =   5352
      TabIndex        =   7
      Text            =   "7"
      Top             =   1013
      Width           =   435
   End
   Begin VB.TextBox txtField 
      Height          =   280
      Index           =   6
      Left            =   1032
      MaxLength       =   30
      TabIndex        =   6
      Text            =   "6"
      Top             =   1013
      Width           =   2595
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   1
      Left            =   2952
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "1"
      Top             =   710
      Width           =   675
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查覆完畢(&O)"
      Height          =   345
      Left            =   6360
      Style           =   1  '圖片外觀
      TabIndex        =   35
      Top             =   40
      Width           =   1200
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一張單(&N)"
      Height          =   345
      Left            =   3173
      Style           =   1  '圖片外觀
      TabIndex        =   38
      Top             =   40
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4788
      Left            =   24
      TabIndex        =   79
      Top             =   2664
      Width           =   8928
      _ExtentX        =   15748
      _ExtentY        =   8446
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "文字1"
      TabPicture(0)   =   "frm090128.frx":082E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtDT(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl3(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "GRD1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tmpKeyPic1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FR11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FR12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdKey(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdKD(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FR31"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "文字2"
      TabPicture(1)   =   "frm090128.frx":084A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(1)"
      Tab(1).Control(1)=   "txtDT(4)"
      Tab(1).Control(2)=   "lbl3(2)"
      Tab(1).Control(3)=   "GRD2"
      Tab(1).Control(4)=   "tmpKeyPic2"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "FR21"
      Tab(1).Control(7)=   "FR22"
      Tab(1).Control(8)=   "cmdKD(1)"
      Tab(1).Control(9)=   "cmdKey(1)"
      Tab(1).Control(10)=   "FR32"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "資料維護"
      TabPicture(2)   =   "frm090128.frx":0866
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LBL1(18)"
      Tab(2).Control(1)=   "LBL1(19)"
      Tab(2).Control(2)=   "LBL1(20)"
      Tab(2).Control(3)=   "LBL1(17)"
      Tab(2).Control(4)=   "LBL1(21)"
      Tab(2).Control(5)=   "Label7"
      Tab(2).Control(6)=   "LBL1(22)"
      Tab(2).Control(7)=   "Label4"
      Tab(2).Control(8)=   "Label5"
      Tab(2).Control(9)=   "LBL1(23)"
      Tab(2).Control(10)=   "LBL1(24)"
      Tab(2).Control(11)=   "Line1"
      Tab(2).Control(12)=   "LBL1(25)"
      Tab(2).Control(13)=   "LBL1(26)"
      Tab(2).Control(14)=   "Label6"
      Tab(2).Control(15)=   "txtField(18)"
      Tab(2).Control(16)=   "txtField(20)"
      Tab(2).Control(17)=   "txtField(17)"
      Tab(2).Control(18)=   "chk1"
      Tab(2).Control(19)=   "txtField(21)"
      Tab(2).Control(20)=   "txt1(2)"
      Tab(2).Control(21)=   "txt1(1)"
      Tab(2).Control(22)=   "txtField(22)"
      Tab(2).Control(23)=   "txtField(24)"
      Tab(2).Control(24)=   "txtChange(0)"
      Tab(2).Control(25)=   "cmdChange"
      Tab(2).Control(26)=   "txtChange(1)"
      Tab(2).Control(27)=   "txtChange(2)"
      Tab(2).Control(28)=   "txtChange(3)"
      Tab(2).Control(29)=   "txtChange(4)"
      Tab(2).ControlCount=   30
      Begin VB.TextBox txtChange 
         Height          =   270
         Index           =   4
         Left            =   -72120
         MaxLength       =   2
         TabIndex        =   33
         Top             =   3112
         Width           =   465
      End
      Begin VB.TextBox txtChange 
         Height          =   270
         Index           =   3
         Left            =   -72480
         MaxLength       =   1
         TabIndex        =   32
         Top             =   3112
         Width           =   345
      End
      Begin VB.TextBox txtChange 
         Height          =   270
         Index           =   2
         Left            =   -73320
         MaxLength       =   6
         TabIndex        =   31
         Top             =   3112
         Width           =   825
      End
      Begin VB.TextBox txtChange 
         Height          =   270
         Index           =   1
         Left            =   -73800
         MaxLength       =   3
         TabIndex        =   30
         Top             =   3112
         Width           =   465
      End
      Begin VB.CommandButton cmdChange 
         BackColor       =   &H00FFC0FF&
         Caption         =   "歸入"
         Height          =   345
         Left            =   -71400
         MaskColor       =   &H00FFFFFF&
         Style           =   1  '圖片外觀
         TabIndex        =   34
         Top             =   3075
         Width           =   900
      End
      Begin VB.TextBox txtChange 
         Height          =   270
         Index           =   0
         Left            =   -73800
         TabIndex        =   29
         Top             =   2760
         Width           =   6345
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   24
         Left            =   -73440
         MaxLength       =   9
         TabIndex        =   25
         Text            =   "24"
         Top             =   1320
         Width           =   1035
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   22
         Left            =   -69720
         MaxLength       =   1
         TabIndex        =   28
         Top             =   1725
         Width           =   435
      End
      Begin VB.Frame FR32 
         BorderStyle     =   0  '沒有框線
         Height          =   324
         Left            =   -70680
         TabIndex        =   131
         Top             =   1728
         Width           =   1815
         Begin VB.CommandButton cmdAFUpd 
            BackColor       =   &H00C0FFFF&
            Caption         =   "變更"
            Height          =   255
            Index           =   1
            Left            =   0
            Style           =   1  '圖片外觀
            TabIndex        =   133
            Top             =   0
            Width           =   800
         End
         Begin VB.CommandButton cmdAFDel 
            BackColor       =   &H00C0FFFF&
            Caption         =   "刪除"
            Height          =   255
            Index           =   1
            Left            =   960
            Style           =   1  '圖片外觀
            TabIndex        =   132
            Top             =   0
            Width           =   800
         End
      End
      Begin VB.Frame FR31 
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   348
         Left            =   4344
         TabIndex        =   128
         Top             =   1680
         Width           =   1815
         Begin VB.CommandButton cmdAFDel 
            BackColor       =   &H00C0FFFF&
            Caption         =   "刪除"
            Height          =   255
            Index           =   0
            Left            =   960
            Style           =   1  '圖片外觀
            TabIndex        =   130
            Top             =   48
            Width           =   800
         End
         Begin VB.CommandButton cmdAFUpd 
            BackColor       =   &H00C0FFFF&
            Caption         =   "變更"
            Height          =   255
            Index           =   0
            Left            =   0
            Style           =   1  '圖片外觀
            TabIndex        =   129
            Top             =   48
            Width           =   800
         End
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -68160
         TabIndex        =   124
         Top             =   1440
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -68160
         TabIndex        =   123
         Top             =   1800
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   21
         Left            =   -69720
         MaxLength       =   7
         TabIndex        =   23
         Text            =   "21"
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chk1 
         Caption         =   "查名單輸入時，已收文"
         Height          =   255
         Left            =   -69720
         TabIndex        =   27
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   17
         Left            =   -73440
         MaxLength       =   7
         TabIndex        =   22
         Text            =   "17"
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   20
         Left            =   -73440
         MaxLength       =   9
         TabIndex        =   26
         Text            =   "20"
         Top             =   1725
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   18
         Left            =   -73440
         MaxLength       =   1
         TabIndex        =   24
         Top             =   960
         Width           =   435
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00C0FFFF&
         Caption         =   "另開視窗"
         Height          =   345
         Index           =   1
         Left            =   -68688
         Style           =   1  '圖片外觀
         TabIndex        =   90
         Top             =   2544
         Width           =   1155
      End
      Begin VB.CommandButton cmdKD 
         BackColor       =   &H00C0FFFF&
         Caption         =   "下載"
         Height          =   345
         Index           =   1
         Left            =   -67128
         Style           =   1  '圖片外觀
         TabIndex        =   91
         Top             =   2544
         Width           =   795
      End
      Begin VB.CommandButton cmdKD 
         BackColor       =   &H00C0FFFF&
         Caption         =   "下載"
         Height          =   345
         Index           =   0
         Left            =   7896
         Style           =   1  '圖片外觀
         TabIndex        =   87
         Top             =   2520
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00C0FFFF&
         Caption         =   "另開視窗"
         Height          =   345
         Index           =   0
         Left            =   6336
         Style           =   1  '圖片外觀
         TabIndex        =   85
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Frame FR22 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   1656
         Left            =   -74928
         TabIndex        =   110
         Top             =   3072
         Width           =   4575
         Begin VB.CheckBox Chk2 
            Caption         =   "第三人出名"
            Height          =   225
            Index           =   2
            Left            =   1020
            TabIndex        =   69
            Top             =   1344
            Width           =   1395
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不出名代理"
            Height          =   225
            Index           =   3
            Left            =   2520
            TabIndex        =   70
            Top             =   1344
            Width           =   1395
         End
         Begin VB.CommandButton cmdSaveD 
            Caption         =   "存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   3
            Left            =   2760
            Style           =   1  '圖片外觀
            TabIndex        =   67
            Top             =   0
            Width           =   960
         End
         Begin VB.ComboBox Cbo4 
            Height          =   276
            IntegralHeight  =   0   'False
            ItemData        =   "frm090128.frx":0882
            Left            =   900
            List            =   "frm090128.frx":0884
            TabIndex        =   66
            Text            =   "Cbo4"
            Top             =   0
            Width           =   1485
         End
         Begin MSForms.TextBox txtDT 
            Height          =   300
            Index           =   7
            Left            =   3912
            TabIndex        =   153
            Top             =   1296
            Visible         =   0   'False
            Width           =   492
            VariousPropertyBits=   -1466941413
            MaxLength       =   30
            ScrollBars      =   2
            Size            =   "873;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtDT 
            Height          =   984
            Index           =   5
            Left            =   720
            TabIndex        =   68
            Top             =   300
            Width           =   3800
            VariousPropertyBits=   -1466941413
            MaxLength       =   300
            ScrollBars      =   2
            Size            =   "6703;1736"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "是否出名："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   3
            Left            =   24
            TabIndex        =   151
            Top             =   1380
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "意見："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   10
            Left            =   0
            TabIndex        =   112
            Top             =   385
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "覆核結果："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   11
            Left            =   0
            TabIndex        =   111
            Top             =   60
            Width           =   900
         End
      End
      Begin VB.Frame FR21 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   996
         Left            =   -74928
         TabIndex        =   107
         Top             =   1704
         Width           =   6060
         Begin VB.CommandButton cmdSaveD 
            Caption         =   "存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   2
            Left            =   2760
            Style           =   1  '圖片外觀
            TabIndex        =   57
            Top             =   10
            Width           =   960
         End
         Begin VB.ComboBox Cbo3 
            Height          =   276
            IntegralHeight  =   0   'False
            ItemData        =   "frm090128.frx":0886
            Left            =   900
            List            =   "frm090128.frx":0888
            TabIndex        =   56
            Text            =   "Cbo3"
            Top             =   0
            Width           =   1485
         End
         Begin MSForms.TextBox txtDT 
            Height          =   620
            Index           =   3
            Left            =   720
            TabIndex        =   58
            Top             =   330
            Width           =   5244
            VariousPropertyBits=   -1466941413
            MaxLength       =   300
            ScrollBars      =   2
            Size            =   "9250;1094"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "意見："
            Height          =   180
            Index           =   7
            Left            =   0
            TabIndex        =   109
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "查覆結果："
            Height          =   180
            Index           =   6
            Left            =   0
            TabIndex        =   108
            Top             =   60
            Width           =   900
         End
      End
      Begin VB.Frame FR12 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   1656
         Left            =   48
         TabIndex        =   104
         Top             =   3072
         Width           =   4575
         Begin VB.CheckBox Chk2 
            Caption         =   "不出名代理"
            Height          =   225
            Index           =   1
            Left            =   2352
            TabIndex        =   55
            Top             =   1344
            Width           =   1395
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "第三人出名"
            Height          =   225
            Index           =   0
            Left            =   936
            TabIndex        =   54
            Top             =   1344
            Width           =   1395
         End
         Begin VB.CommandButton cmdSaveD 
            Caption         =   "存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   2760
            Style           =   1  '圖片外觀
            TabIndex        =   52
            Top             =   0
            Width           =   960
         End
         Begin VB.ComboBox Cbo2 
            Height          =   300
            IntegralHeight  =   0   'False
            ItemData        =   "frm090128.frx":088A
            Left            =   900
            List            =   "frm090128.frx":088C
            TabIndex        =   51
            Text            =   "Cbo2"
            Top             =   0
            Width           =   1485
         End
         Begin MSForms.TextBox txtDT 
            Height          =   300
            Index           =   6
            Left            =   3744
            TabIndex        =   152
            Top             =   1320
            Visible         =   0   'False
            Width           =   492
            VariousPropertyBits=   -1466941413
            MaxLength       =   30
            ScrollBars      =   2
            Size            =   "873;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtDT 
            Height          =   984
            Index           =   2
            Left            =   720
            TabIndex        =   53
            Top             =   300
            Width           =   3792
            VariousPropertyBits=   -1466941413
            MaxLength       =   300
            ScrollBars      =   2
            Size            =   "6689;1736"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "是否出名："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   2
            Left            =   0
            TabIndex        =   150
            Top             =   1368
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "意見："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   9
            Left            =   0
            TabIndex        =   106
            Top             =   385
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "覆核結果："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   8
            Left            =   0
            TabIndex        =   105
            Top             =   60
            Width           =   900
         End
      End
      Begin VB.Frame FR11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   996
         Left            =   72
         TabIndex        =   101
         Top             =   1704
         Width           =   6060
         Begin VB.ComboBox Cbo1 
            Height          =   300
            IntegralHeight  =   0   'False
            ItemData        =   "frm090128.frx":088E
            Left            =   900
            List            =   "frm090128.frx":0890
            TabIndex        =   42
            Text            =   "Cbo1"
            Top             =   0
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveD 
            Caption         =   "存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   2760
            Style           =   1  '圖片外觀
            TabIndex        =   43
            Top             =   0
            Width           =   960
         End
         Begin MSForms.TextBox txtDT 
            Height          =   620
            Index           =   0
            Left            =   720
            TabIndex        =   44
            Top             =   300
            Width           =   5244
            VariousPropertyBits=   -1466941413
            MaxLength       =   840
            ScrollBars      =   2
            Size            =   "9250;1094"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "查覆結果："
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   103
            Top             =   60
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "意見："
            Height          =   180
            Index           =   1
            Left            =   0
            TabIndex        =   102
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "附件區："
         Height          =   1695
         Left            =   -70272
         TabIndex        =   97
         Top             =   3072
         Width           =   4125
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            Height          =   345
            Index           =   1
            Left            =   960
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1224
            Width           =   675
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "刪除"
            Height          =   345
            Index           =   1
            Left            =   3210
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1224
            Width           =   675
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "新增"
            Height          =   345
            Index           =   1
            Left            =   2460
            TabIndex        =   63
            Top             =   1224
            Width           =   675
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "下載"
            Height          =   345
            Index           =   1
            Left            =   1710
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1224
            Width           =   675
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   345
            Index           =   1
            Left            =   210
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1224
            Width           =   675
         End
         Begin VB.ListBox lstAtt 
            Height          =   852
            Index           =   1
            IntegralHeight  =   0   'False
            ItemData        =   "frm090128.frx":0892
            Left            =   72
            List            =   "frm090128.frx":0899
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   240
            Width           =   4020
         End
      End
      Begin VB.PictureBox tmpKeyPic2 
         Height          =   2000
         Left            =   -68700
         ScaleHeight     =   163
         ScaleMode       =   3  '像素
         ScaleWidth      =   204
         TabIndex        =   96
         Top             =   420
         Width           =   2500
         Begin VB.Image tmpKeyImg2 
            Height          =   1770
            Left            =   0
            Top             =   0
            Width           =   1890
         End
      End
      Begin VB.PictureBox tmpKeyPic1 
         Height          =   2000
         Left            =   6300
         ScaleHeight     =   163
         ScaleMode       =   3  '像素
         ScaleWidth      =   204
         TabIndex        =   93
         Top             =   420
         Width           =   2500
         Begin VB.Image tmpKeyImg1 
            Height          =   1770
            Left            =   0
            Top             =   0
            Width           =   1890
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   1200
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   2117
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|查名組群|查覆結果|查覆意見|覆核結果|覆核意見"
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
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD2 
         Height          =   1200
         Left            =   -74880
         TabIndex        =   95
         Top             =   480
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   2117
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|查名組群|查覆結果|查覆意見|覆核結果|覆核意見"
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
         _Band(0).Cols   =   6
      End
      Begin VB.Frame Frame1 
         Caption         =   "附件區："
         Height          =   1668
         Left            =   4728
         TabIndex        =   81
         Top             =   3072
         Width           =   4125
         Begin VB.ListBox lstAtt 
            Height          =   852
            Index           =   0
            IntegralHeight  =   0   'False
            ItemData        =   "frm090128.frx":08A5
            Left            =   60
            List            =   "frm090128.frx":08AC
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   264
            Width           =   4020
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   345
            Index           =   0
            Left            =   192
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1224
            Width           =   675
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "下載"
            Height          =   345
            Index           =   0
            Left            =   1692
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1224
            Width           =   675
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "新增"
            Height          =   345
            Index           =   0
            Left            =   2436
            TabIndex        =   49
            Top             =   1224
            Width           =   675
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "刪除"
            Height          =   345
            Index           =   0
            Left            =   3192
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1224
            Width           =   675
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            Height          =   345
            Index           =   0
            Left            =   936
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1224
            Width           =   675
         End
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "審定號/申請號: "
         Height          =   180
         Index           =   2
         Left            =   -74952
         TabIndex        =   155
         Top             =   2760
         Width           =   1224
      End
      Begin MSForms.TextBox txtDT 
         Height          =   300
         Index           =   4
         Left            =   -73700
         TabIndex        =   59
         Top             =   2712
         Width           =   3792
         VariousPropertyBits=   -1466941413
         MaxLength       =   30
         ScrollBars      =   2
         Size            =   "6703;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "審定號/申請號:"
         Height          =   180
         Index           =   1
         Left            =   48
         TabIndex        =   154
         Top             =   2760
         Width           =   1176
      End
      Begin MSForms.TextBox txtDT 
         Height          =   300
         Index           =   1
         Left            =   1300
         TabIndex        =   45
         Top             =   2712
         Width           =   3792
         VariousPropertyBits=   -1466941413
         MaxLength       =   30
         ScrollBars      =   2
         Size            =   "6703;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         Caption         =   "(請以"","" 區隔)"
         Height          =   255
         Left            =   -67320
         TabIndex        =   149
         Top             =   2768
         Width           =   1095
      End
      Begin VB.Label LBL1 
         Caption         =   " P.S 與畫面上的查覆明細無關"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   26
         Left            =   -72120
         TabIndex        =   148
         Top             =   2460
         Width           =   2535
      End
      Begin VB.Label LBL1 
         Caption         =   "整批委查單號歸入卷宗區"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   25
         Left            =   -74760
         TabIndex        =   147
         Top             =   2460
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C000C0&
         X1              =   -74760
         X2              =   -66480
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label LBL1 
         Caption         =   "委查單號："
         Height          =   255
         Index           =   24
         Left            =   -74760
         TabIndex        =   146
         Top             =   2790
         Width           =   1515
      End
      Begin VB.Label LBL1 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   23
         Left            =   -74760
         TabIndex        =   145
         Top             =   3120
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "(Y)"
         Height          =   255
         Left            =   -69240
         TabIndex        =   139
         Top             =   1740
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "(Y)"
         Height          =   255
         Left            =   -72960
         TabIndex        =   138
         Top             =   975
         Width           =   375
      End
      Begin VB.Label LBL1 
         Caption         =   "是否撤回："
         Height          =   255
         Index           =   22
         Left            =   -70680
         TabIndex        =   137
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "(當申請編號的所有委查單完成會上日期)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -69720
         TabIndex        =   122
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label LBL1 
         Caption         =   "查覆完成日期："
         Height          =   255
         Index           =   21
         Left            =   -71040
         TabIndex        =   121
         Top             =   608
         Width           =   1335
      End
      Begin VB.Label LBL1 
         Caption         =   "覆核日期："
         Height          =   255
         Index           =   17
         Left            =   -74760
         TabIndex        =   120
         Top             =   615
         Width           =   975
      End
      Begin VB.Label LBL1 
         Caption         =   "櫃台收文號："
         Height          =   255
         Index           =   20
         Left            =   -74760
         TabIndex        =   119
         Top             =   1733
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LBL1 
         Caption         =   "收件分發日期："
         Height          =   255
         Index           =   19
         Left            =   -74760
         TabIndex        =   118
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label LBL1 
         Caption         =   "查覆結果已讀："
         Height          =   255
         Index           =   18
         Left            =   -74760
         TabIndex        =   117
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "(註:選取時,下方顯示查覆資料)"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   -72000
         TabIndex        =   94
         Top             =   280
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "(註:選取時,下方顯示查覆資料)"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   0
         Left            =   100
         TabIndex        =   80
         Top             =   320
         Width           =   2500
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   7920
      TabIndex        =   36
      Top             =   40
      Width           =   900
   End
   Begin VB.CommandButton cmdTo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "收文"
      Height          =   345
      Left            =   5240
      Style           =   1  '圖片外觀
      TabIndex        =   40
      Top             =   40
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   -24
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame FraCase 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame3"
      Height          =   300
      Left            =   4752
      TabIndex        =   115
      Top             =   1332
      Width           =   3975
      Begin VB.TextBox txtField 
         Height          =   300
         Index           =   15
         Left            =   2820
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "15"
         Top             =   0
         Width           =   480
      End
      Begin VB.TextBox txtField 
         Height          =   300
         Index           =   14
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   13
         Text            =   "1"
         Top             =   0
         Width           =   360
      End
      Begin VB.TextBox txtField 
         Height          =   300
         Index           =   13
         Left            =   1700
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "13"
         Top             =   0
         Width           =   740
      End
      Begin VB.TextBox txtField 
         Height          =   300
         Index           =   12
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "12"
         Top             =   0
         Width           =   500
      End
      Begin VB.Label LBL1 
         Caption         =   "已收文案件："
         Height          =   180
         Index           =   12
         Left            =   0
         TabIndex        =   116
         Top             =   45
         Width           =   1170
      End
   End
   Begin MSForms.TextBox textService 
      Height          =   300
      Left            =   1368
      TabIndex        =   21
      Top             =   2316
      Width           =   7392
      VariousPropertyBits=   671105051
      Size            =   "13044;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCName 
      Height          =   300
      Left            =   1032
      TabIndex        =   10
      Top             =   1332
      Width           =   3600
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "6350;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LBL1 
      Alignment       =   1  '靠右對齊
      Caption         =   "(保留)"
      Height          =   180
      Index           =   15
      Left            =   720
      TabIndex        =   144
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LBL1 
      Caption         =   "發 文 日："
      Height          =   180
      Index           =   14
      Left            =   7224
      TabIndex        =   143
      Top             =   2028
      Width           =   852
   End
   Begin VB.Label LBL1 
      Caption         =   "通知送件日："
      Height          =   180
      Index           =   13
      Left            =   5208
      TabIndex        =   142
      Top             =   2028
      Width           =   1092
   End
   Begin MSForms.Label lblAppNo 
      Height          =   252
      Index           =   5
      Left            =   1032
      TabIndex        =   141
      Top             =   204
      Width           =   1020
      VariousPropertyBits=   27
      Caption         =   "lblA(5)"
      Size            =   "1799;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   40
      TabIndex        =   140
      Top             =   237
      Width           =   1000
   End
   Begin VB.Label LBL1 
      Caption         =   "圖形"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   9
      Left            =   7608
      TabIndex        =   136
      Top             =   1068
      Width           =   408
   End
   Begin VB.Label LBL1 
      Caption         =   "英文"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   8
      Left            =   6312
      TabIndex        =   135
      Top             =   1068
      Width           =   408
   End
   Begin VB.Label LBL1 
      Caption         =   "中文"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   7
      Left            =   4956
      TabIndex        =   134
      Top             =   1068
      Width           =   408
   End
   Begin MSForms.Label lblAppNo 
      Height          =   252
      Index           =   4
      Left            =   7992
      TabIndex        =   127
      Top             =   732
      Width           =   732
      VariousPropertyBits=   27
      Caption         =   "lblA(4)"
      Size            =   "1296;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LBL1 
      Caption         =   "覆核主管："
      Height          =   180
      Index           =   2
      Left            =   6432
      TabIndex        =   126
      Top             =   768
      Width           =   936
   End
   Begin VB.Label LBL1 
      Alignment       =   1  '靠右對齊
      Caption         =   "查覆日期："
      Height          =   180
      Index           =   5
      Left            =   4512
      TabIndex        =   125
      Top             =   768
      Width           =   900
   End
   Begin VB.Label LBL1 
      Caption         =   "查名路徑："
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   11
      Left            =   72
      TabIndex        =   114
      Top             =   2028
      Width           =   960
   End
   Begin VB.Label LBL1 
      Caption         =   "客戶名稱："
      Height          =   180
      Index           =   10
      Left            =   72
      TabIndex        =   113
      Top             =   1392
      Width           =   960
   End
   Begin MSForms.Label lblAppNo 
      Height          =   252
      Index           =   2
      Left            =   3672
      TabIndex        =   100
      Top             =   468
      Width           =   756
      VariousPropertyBits=   27
      Caption         =   "lblA(2)"
      Size            =   "1323;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtUnicode 
      Height          =   300
      Index           =   2
      Left            =   5352
      TabIndex        =   16
      Top             =   1644
      Width           =   3408
      VariousPropertyBits=   679495707
      MaxLength       =   50
      Size            =   "6006;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtUnicode 
      Height          =   300
      Index           =   1
      Left            =   1032
      TabIndex        =   15
      Top             =   1644
      Width           =   3600
      VariousPropertyBits=   679495707
      MaxLength       =   50
      Size            =   "6350;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAppNo 
      Height          =   252
      Index           =   3
      Left            =   3672
      TabIndex        =   89
      Top             =   732
      Width           =   732
      VariousPropertyBits=   27
      Caption         =   "lblA(3)"
      Size            =   "1296;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAppNo 
      Height          =   252
      Index           =   1
      Left            =   1032
      TabIndex        =   88
      Top             =   732
      Width           =   1116
      VariousPropertyBits=   27
      Caption         =   "lblAppNo(1)"
      Size            =   "1958;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "文字2："
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   2
      Left            =   4752
      TabIndex        =   86
      Top             =   1704
      Width           =   696
   End
   Begin VB.Label Label1 
      Caption         =   "文字："
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   1
      Left            =   432
      TabIndex        =   84
      Top             =   1704
      Width           =   648
   End
   Begin VB.Label LBL1 
      Caption         =   "指定商品/服務:"
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   16
      Left            =   72
      TabIndex        =   83
      Top             =   2340
      Width           =   1212
   End
   Begin VB.Label LBL1 
      Caption         =   "委查組群："
      Height          =   180
      Index           =   6
      Left            =   72
      TabIndex        =   78
      Top             =   1068
      Width           =   936
   End
   Begin VB.Label LBL1 
      Caption         =   "申請日期："
      Height          =   180
      Index           =   3
      Left            =   4512
      TabIndex        =   77
      Top             =   504
      Width           =   936
   End
   Begin VB.Label LBL1 
      Caption         =   "委查人："
      Height          =   180
      Index           =   0
      Left            =   2208
      TabIndex        =   76
      Top             =   504
      Width           =   936
   End
   Begin MSForms.Label lblAppNo 
      Height          =   252
      Index           =   0
      Left            =   1032
      TabIndex        =   75
      Top             =   468
      Width           =   1116
      VariousPropertyBits=   27
      Caption         =   "lblAppNo(0)"
      Size            =   "1958;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請編號："
      Height          =   180
      Index           =   19
      Left            =   72
      TabIndex        =   74
      Top             =   504
      Width           =   960
   End
   Begin VB.Label LBL1 
      Caption         =   "查名人："
      Height          =   180
      Index           =   1
      Left            =   2208
      TabIndex        =   73
      Top             =   768
      Width           =   816
   End
   Begin VB.Label Label1 
      Caption         =   "委查單號："
      Height          =   180
      Index           =   4
      Left            =   72
      TabIndex        =   72
      Top             =   768
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "委查筆數：　　　　　  筆　　　　 　　 筆　　　　 　　 筆"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   5
      Left            =   4032
      TabIndex        =   71
      Top             =   1068
      Width           =   4872
   End
   Begin VB.Label LBL1 
      Alignment       =   1  '靠右對齊
      Caption         =   "期限日期："
      Height          =   180
      Index           =   4
      Left            =   6432
      TabIndex        =   65
      Top             =   504
      Width           =   900
   End
End
Attribute VB_Name = "frm090128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/01 改成Form2.0 ; GRD1,GRD改字型=新細明體-ExtB、txtDT(index)、lblAppNo(index)、txtField(10)=>textCName、txtField(16)=>textService
'Create by Lydia 2015/08/14 查覆明細作業
Option Explicit

Public m_NoList As String '傳入單號(多筆)
Public m_NoIdx As Integer '操作的單號
Public iStiu As Integer  '狀態 :0查詢 1編輯
'使用者角色 U:查名人 Q:委查人(包含非分配的查名人) M:覆核主管 A:查名單維護(限電腦中心)
Public R_type As String '依條件判斷權限記錄在cmdSend.tag
Public mbolCall As Boolean  '外部呼叫
Private Const fileMax As Integer = 99  '預設最大附件數
'Private Const TMQ_無 = "6"  'Added by Lydia 2016/04/25 'Remove by Lydia 2016/07/06
'設定可使用表單
Public Tmpfrm090129 As Form
Dim addRow(0 To 1) As Integer '目前附件最大流水號

Dim m_PrevForm As Form '前一畫面
Dim m_TMQApp As String '查名單申請號

Dim jj As Integer
Dim dblPrevRow As Integer
Dim dblPrevRow2 As Integer
Dim tmpArr As Variant
Dim adoTMQ As New ADODB.Recordset  '查名單
Dim adoD1 As New ADODB.Recordset  '文字1/圖形
Dim adoD2 As New ADODB.Recordset  '文字2
Dim oText As Control

'附件宣告區
Dim m_AttachPath As String
'Public strLoadPath As String '讀取前次設定路徑 'Remove by Lydia 2016/05/26
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Dim colAns1 As Integer  '查覆結果-位置
Dim colAns2 As Integer  '覆核結果-位置
Dim colAcase As Integer '近似本所案件的申請號
Dim colTQD11 As Integer 'Added by Lydia 2019/12/12 是否出名
Dim mTQD01 As String '申請編號
Dim mTQD02 As String '委查單號(查名單號)
Dim mTQD03 As String '0=圖形, 1=文字
Dim mTQF03 As String '類別: 0=圖形, 1=文字1, 2=文字2
Dim mTQF04 As String '流水號
Dim mTMQ19 As String '是否已讀
Dim mTQA03 As String '申請編號的全部組群

'Public mTMQ20 As String 'Modified by Lydia 2016/04/19 查名代號,以Table對應
Public ShowCP09 As String '目前進度的總收文號(可能是多次申請)
'Added by Lydia 2016/07/07 目前進度的案號資料
Dim ShowCP(1 To 4) As String
Dim ShowCP14 As String
Dim ShowCP57 As String
'end 2016/07/07
Dim FirstCP09 As String '總收文號(第一次收文=TMQ21)
Dim FirstCP(1 To 4) As String '本所案號
Dim FirstCPP02t As String '符合規則的卷宗區檔名開頭
Dim FirstCP14 As String 'Added by Lydia 2016/04/29 承辦人
Dim FirstCP57 As String 'Added by Lydia 2016/04/29 取消收文日
Dim mPrevTM1215 As String '保留上次輸入的審定號/申請號
Dim bolSave As String '是否已存檔
Dim mPreStabs As Integer '記錄頁籤
Dim mbolSend As Boolean '是否查覆完畢
Dim mTQA15 As String '已收文案件(查名單輸入)
Dim mTQA20 As String '委查人自請撤回
Dim bolModify As Boolean '查覆完畢再次修改(先有mbolSend=true,之後有修改內容)
Dim bolModCheck As Boolean '筆數已變更
Dim strMod(0 To 2)  As String '再次修改的記錄mail(0=主旨,1=修改歷程,2=收件人)
Dim strAgree As String 'Added by Lydia 2016/06/27 內商核可人員
Dim strPreAgree As String 'Added by Lydia 2022/05/25 內商查名覆核人員
Dim bolChgTM22 As Boolean 'Added by Lydia 2016/09/12 覆核結果更改
'Added by Lydia 2017/10/20 用在OpenDocument
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32
Private Const ERROR_BAD_FORMAT As Long = 11

Dim bolMsgRight As Boolean 'Added by Lydia 2018/11/22 Form 2.0表單是否彈過提示滑鼠右鍵無效
Dim SyxMsg As String 'Added by Lydia 2018/11/22 Form 2.0表單是否彈過提示滑鼠右鍵無效(記錄前一位置)
'Added by Lydia 2019/01/30 文件和查名同時齊備=>更新承辦期限
'Remove by Lydia 2019/05/21 依照各案件的現況
'Dim m_CP13 As String, m_EP06 As String '收文號資料
'Dim m_CP06 As String, m_CP122 As String '本所期限, 是否急件

Dim stIdList As String 'Added by Lydia 2019/08/12 創新業務組成員可操作清單(WXX部門的人可以操作自已部門所有人的資料,
                                                                                                                        '例W10所有人都可操作W1001，W20所有人都可操作W2001。
'Added by Lydia 2019/12/25 開放特殊設定權限
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim strGrpTmp1 As String, strGrpTmp2 As String 'Added by Lydia 2020/12/04
Dim nFrm As Form 'Added by Lydia 2024/07/17
Dim mTMQ20 As String 'Added by Lydia 2024/12/20

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Function IsSaveData() As Boolean
Dim tmpBol As Boolean
   IsSaveData = False
   '檢查明細是否已存檔
   If Me.SSTab1.Tab = 0 Then
      Call ChangeDetailAns(0, Cbo1, txtDt(0), tmpBol, txtDt(1))
      'Modified by Lydia 2019/12/12 +是否出名
      'Call ChangeDetailAns(1, Cbo2, txtDT(2))
      Call ChangeDetailAns(1, Cbo2, txtDt(2), , , txtDt(6))
   Else
      Call ChangeDetailAns(2, Cbo3, txtDt(3), tmpBol, txtDt(4))
      'Modified by Lydia 2019/12/12 +是否出名
      'Call ChangeDetailAns(3, Cbo4, txtDT(5))
      Call ChangeDetailAns(3, Cbo4, txtDt(5), , , txtDt(7))
   End If
   
   If tmpBol = False Then IsSaveData = True
End Function
'查名單維護-檢查查名資料是否已存檔
Private Function IsSaveAppNo() As Boolean
Dim strTit As String

   IsSaveAppNo = False
   '檢查申請資料是否已存檔
   For Each oText In txtField
       If oText.Text <> oText.Tag Then
          strTit = LBL1(oText.Index).Caption
          If oText.Index < 7 And oText.Index > 9 Then
             strTit = Mid(strTit, 1, Len(strTit) - 1)
          End If
          GoTo JumpMSG
       End If
   Next

   If txtUnicode(1).Text <> txtUnicode(1).Tag Then
      strTit = Label1(1).Caption
      strTit = Mid(strTit, 1, Len(strTit) - 1)
      GoTo JumpMSG
   End If
   If txtUnicode(2).Text <> txtUnicode(2).Tag Then
      strTit = Label1(2).Caption
      strTit = Mid(strTit, 1, Len(strTit) - 1)
      GoTo JumpMSG
   End If
   
   IsSaveAppNo = True
   Exit Function
   
JumpMSG:
    If MsgBox(strTit & "已變更,是否修改?", vbYesNo + vbCritical) = vbYes Then
       Call cmdSend_Click
       IsSaveAppNo = True
    Else
       Exit Function
    End If
End Function
'依權限設定欄位
Private Sub FormEnabled()
Dim tmpBol As Boolean

'查名單維護
If R_type = "A" Then
    cmdSend.Enabled = True
    cmdTo.Visible = False
    cmdSendMail.Visible = False 'Added by Lydia 2016/04/29
    FR31.Visible = True
    FR32.Visible = True
    If mTQD03 = TMQ_AkindPic Then
       cmdAFDel(0).Visible = False
       cmdAFDel(1).Visible = False
    Else
       cmdAFDel(0).Visible = True
       cmdAFDel(1).Visible = True
    End If
Else
    FR31.Visible = False
    FR32.Visible = False
    If iStiu = 1 Then
       cmdSend.Enabled = True
    Else
       cmdSend.Enabled = False
    End If

    For intI = 0 To 3
        cmdSaveD(intI).Visible = False
    Next intI
    'Modified by Lydia 2016/04/25 改成鎖定欄位=>Lock可以鎖輸入,卻不可鎖滑鼠右鍵的貼上
    'FR11.Enabled = False: FR21.Enabled = False
    'FR12.Enabled = False: FR22.Enabled = False
    Cbo1.Locked = True: Cbo2.Locked = True
    Cbo3.Locked = True: Cbo4.Locked = True
    'Added by Lydia 2019/12/12
    Label3(2).Visible = False: Label3(3).Visible = False
    chk2(0).Visible = False: chk2(1).Visible = False
    chk2(2).Visible = False: chk2(3).Visible = False
                  
    '外部呼叫隱藏按鈕
    If mbolCall = True Then
       cmdSend.Visible = False: cmdTo.Visible = False
       cmdSendMail.Visible = False 'Added by Lydia 2016/04/29
       If Len(m_NoList) <= 10 Then
          cmdPrevious.Visible = False: cmdNext.Visible = False
       End If
    Else
       Select Case R_type
           Case "U" '查名人
               'Added by Lydia 2016/04/21 +狀態
               If mTQA20 = "" And iStiu = 1 Then
                  'Modified by Lydia 2016/04/25
                  'FR11.Enabled = True: FR21.Enabled = True
                  Cbo1.Locked = False: txtDt(0).Locked = False: txtDt(1).Locked = False
                  Cbo3.Locked = False: txtDt(3).Locked = False: txtDt(4).Locked = False
                  cmdSaveD(0).Visible = True: cmdSaveD(2).Visible = True
               End If
               cmdTo.Visible = False
               cmdSendMail.Visible = False 'Added by Lydia 2016/04/29
           Case "M" '覆核主管
               If mTQA20 = "" Then
                  'Modified by Lydia 2016/04/25
                  'FR12.Enabled = True: FR22.Enabled = True
                  Cbo2.Locked = False: txtDt(2).Locked = False
                  Cbo4.Locked = False: txtDt(5).Locked = False
                  txtDt(1).Locked = False: txtDt(4).Locked = False   'Added by Lydia 2023/07/06 開放覆核主管可修改「審定號/申請號」
                  cmdSaveD(1).Visible = True: cmdSaveD(3).Visible = True
                  'Added by Lydia 2019/12/12 核可人員 => 設定" 是否出名 "
                  If InStr("M,A", R_type) > 0 And (InStr(strAgree, strUserNum) > 0 Or Pub_StrUserSt03 = "M51") Then
                     Label3(2).Visible = True: Label3(3).Visible = True
                     chk2(0).Visible = True: chk2(1).Visible = True
                     chk2(2).Visible = True: chk2(3).Visible = True
                  End If
                  
               'Added by Lydia 2017/12/13
               Else
                   MsgBox "委查人已撤回委查單 !", vbCritical
                   cmdSend.Enabled = False
               'end 2017/12/13
               End If
               cmdTo.Visible = False
               cmdSendMail.Visible = False 'Added by Lydia 2016/04/29
           Case "Q" '委查人
               'Modified by Lydia 2017/11/23 主管或職代增加撤回權限。
               'If Trim(txtField(0).Text) = strUserNum And iStiu = 1 Then
               If mbolCall = False And iStiu = 1 Then
                  cmdSend.Visible = True
               Else
                  cmdSend.Visible = False
               End If
       End Select
    End If
   'Modified by Lydia 2016/04/06 +已收文判斷(若已收文,明細的收文按鈕不能按,但是查覆區可以按)
   'If (Pub_StrUserSt03 = "M51" Or Trim(txtField(0).Text) = strUserNum) And mbolSend = True And mTQA15 = "" And (TMQ_CtrRead = False Or mTMQ19 = "Y") Then
   'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
   'If (Pub_StrUserSt03 = "M51" Or Trim(txtField(0).Text) = strUserNum) And mbolSend = True And lblAppNo(5).Caption = "" And (TMQ_CtrRead = False Or mTMQ19 = "Y") And mTQA20 = "" Then
   'Modified by Lydia 2019/12/25 開放特殊設定權限
   'If (Pub_StrUserSt03 = "M51" Or InStr(stIdList, txtField(0).Text) > 0) And mbolSend = True And lblAppNo(5).Caption = "" And (TMQ_CtrRead = False Or mTMQ19 = "Y") And mTQA20 = "" Then
   If (Pub_StrUserSt03 = "M51" Or InStr(stIdList, txtField(0).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, txtField(0).Text) > 0)) _
             And mbolSend = True And lblAppNo(5).Caption = "" And (TMQ_CtrRead = False Or mTMQ19 = "Y") And mTQA20 = "" Then
       cmdTo.Enabled = True
       cmdSendMail.Enabled = False 'Added by Lydia 2016/04/29
   Else
       cmdTo.Enabled = False
        'Added by Lydia 2016/04/29
       'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
       'If txtField(19) & txtField(23) = "" And lblAppNo(5).Caption <> "" And Trim(txtField(0).Text) = strUserNum Then
       'Modified by Lydia 2019/12/25 開放特殊設定權限
       'If txtField(19) & txtField(23) = "" And lblAppNo(5).Caption <> "" And InStr(stIdList, txtField(0).Text) > 0 Then
       If txtField(19) & txtField(23) = "" And lblAppNo(5).Caption <> "" And (InStr(stIdList, txtField(0).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, txtField(0).Text) > 0)) Then
          cmdSendMail.Enabled = True
       Else
          cmdSendMail.Enabled = False
       End If
   End If
   '委查人或非分派到的查名人員不可變更
   If iStiu = 0 Or R_type = "Q" Then
      tmpBol = True 'Locked
   Else
      tmpBol = False
   End If

   For Each oText In txtField
       Select Case oText.Index
           'Modified by Lydia 2024/07/17 鎖定圖形路徑的人工輸入
           'Case 7, 8, 9, 11
           Case 7, 8, 9
                '暫時不鎖編輯,若有組群不查則重新計算筆數
               oText.Locked = tmpBol
           Case Else
               oText.Locked = True
       End Select
   Next
   'Added by Lydia 2021/10/01 txtField(10)=>textCName、txtField(16)=>textService
   textService.Locked = True
   textCName.Locked = True
   'end 2021/10/01
   
   '文字不變更
   txtUnicode(1).Locked = True
   txtUnicode(2).Locked = True
   '已收文案件
   'Remove by Lydia 2016/03/28 預設由委查人輸入(可空白)
   'If (Pub_StrUserSt03 = "M51" Or R_type = "U") And iStiu = "1" Then
   '   For intI = 12 To 15
   '       txtField(intI).Enabled = True
   '   Next intI
   'Else
      For intI = 12 To 15
          txtField(intI).Enabled = False
      Next intI
   'End If
   'end 2016/03/28

End If

'顯示附件(新增,刪除)
If iStiu = 0 Or R_type = "Q" Then
   cmdAddAtt(0).Visible = False
   cmdRemAtt(0).Visible = False
   cmdAddAtt(1).Visible = False
   cmdRemAtt(1).Visible = False
Else
   'Added by Lydia 2021/11/02 判斷覆核主管不可刪除附件; ex.T-234290的HB0040687、HB0040686查名附件不存在，推測是在覆核或核可階段刪除
                                            '另外增加刪除附件的操作者非建立人員，另外寫log(在basUpdate.PUB_TMQAFileDel)。
   If R_type = "M" Then
       cmdRemAtt(0).Visible = False
       cmdRemAtt(1).Visible = False
   Else '以下包含U查名人員，A電腦中心維護
   'end 2021/11/02
       cmdAddAtt(0).Visible = True
       cmdRemAtt(0).Visible = True
       cmdAddAtt(1).Visible = True
       cmdRemAtt(1).Visible = True
   End If 'Added by Lydia 2021/11/02
End If

'查覆/覆核完畢後
'Modified by Lydia 2019/12/12 +Tab
If dblPrevRow > 0 Then Call SetComboAns(0, GRD1, dblPrevRow, Cbo1, Cbo2, txtDt(0), txtDt(1), txtDt(2))
If dblPrevRow2 > 0 Then Call SetComboAns(1, grd2, dblPrevRow2, Cbo3, Cbo4, txtDt(3), txtDt(4), txtDt(5))
End Sub
'清除欄位
Private Sub FormReset()
Dim oLabel As Control

   For Each oText In txtField
      oText.Text = ""
      oText.Tag = ""
      oText.Locked = False
   Next
   For Each oText In txtDt
      oText.Text = ""
      oText.Tag = ""
      oText.Locked = False
   Next
   'Added by Lydia 2017/08/31
   For Each oText In txtChange
      oText.Text = ""
   Next
   For Each oLabel In lblAppNo
      oLabel.Caption = ""
   Next
   
   'Added by Lydia 2021/10/01 txtField(10)=>textCName、txtField(16)=>textService
   textService.Text = "": textService.Tag = ""
   textCName.Text = "": textCName.Tag = ""
   'end 2021/10/01
   
   Call SetCombo(Cbo1)
   Call SetCombo(Cbo3)
   'Modified by Lydia 2016/06/01 區分覆核結果
   Call SetCombo(Cbo2, "A")
   Call SetCombo(Cbo4, "A")
   
   'Added by Lydia 2016/05/04
   Cbo1.Text = "":   Cbo2.Text = ""
   Cbo3.Text = "":   Cbo4.Text = ""
   'end 2016/05/04
   Cbo1.Tag = "":   Cbo2.Tag = ""
   Cbo3.Tag = "":   Cbo4.Tag = ""
   lstAtt(0).Clear:    lstAtt(1).Clear
   txtUnicode(1).Text = "": txtUnicode(1).Tag = ""
   txtUnicode(2).Text = "": txtUnicode(2).Tag = ""
   dblPrevRow = 0: dblPrevRow2 = 0
   For jj = 0 To 1
       cmdKey(jj).Visible = False
       cmdKey(jj).Tag = ""
       cmdKD(jj).Visible = False
       cmdKD(jj).Tag = ""
   Next jj
   'Added by Lydia 2019/12/12
   For jj = 0 To 3
       chk2(jj).Value = vbUnchecked
   Next jj
   
   mPrevTM1215 = ""
   cmdSend.Caption = "查覆完畢(&O)"
   cmdSend.Tag = ""
End Sub
'設定查覆結果
'Modified by Lydia 2016/06/01 +nKind
Private Sub SetCombo(ByRef cmbN As ComboBox, Optional ByRef nKind As String = "")
Dim iX As Integer
    
   tmpArr = Empty
   'Modified by Lydia 2016/06/02 查名結果去掉核可
   'TmpArr = Split(TMQ_結果清單, ",")
   tmpArr = Split(PUB_GetTMQans("2", IIf(nKind = "A", True, False)), ",")
   'TmpArr = Split(IIf(nKind = "A", TMQ_結果清單, Replace(TMQ_結果清單, "0 核可,", "")), ",")
   cmbN.Clear
  '增加空白
   cmbN.AddItem ""
   For iX = 0 To UBound(tmpArr)
       cmbN.AddItem Mid(tmpArr(iX), 3, Len(tmpArr(iX)) - 2)
   Next iX
End Sub
'檢查查覆結果
'Modified by Lydia 2016/06/01 +nKind
Private Function CheckCombo(ByRef cmbN As ComboBox, Optional ByRef nKind As String = "") As Boolean
Dim iX As Integer
    
   CheckCombo = False
   tmpArr = Empty
   'Modified by Lydia 2016/06/04 核可限商標主管(和職代)使用
   'Remove by Lydia 2016/07/07
   'If strSrvDate(1) >= TMQFileFTP Then
        If nKind = "A" And cmbN.Text = "核可" And cmbN.Text <> cmbN.Tag Then
           If InStr("M,A", R_type) > 0 And (InStr(strAgree, strUserNum) > 0 Or Pub_StrUserSt03 = "M51") Then
              CheckCombo = True: Exit Function
           Else
              MsgBox "無核可權限", vbInformation
              cmbN.Text = cmbN.Tag
              Exit Function
           End If
        ElseIf cmbN.Text = "核可" And cmbN.Text = cmbN.Tag Then
           CheckCombo = True: Exit Function
        End If
   'End If
   
   'Modified by Lydia 2016/06/02 查名結果改成模組
   'TmpArr = Split(TMQ_結果清單, ",")
   tmpArr = Split(PUB_GetTMQans("2", False), ",")
   For iX = 0 To UBound(tmpArr)
       If Mid(tmpArr(iX), 3, Len(tmpArr(iX)) - 2) = cmbN.Text Then
          CheckCombo = True: Exit Function
       End If
   Next iX
   
   '可以空白
   If cmbN.Text <> "" Then
      MsgBox "請選取正確的查覆結果", vbInformation
   Else
      CheckCombo = True
   End If
End Function

'讀取查覆結果
'Modified by Lydia 2019/12/12 +頁籤 iTabs
'Modified by Lydia 2021/10/01 TextBox => Control
Private Sub SetComboAns(ByVal iTabs As Integer, ByRef nGrid As MSHFlexGrid, ByRef nRow As Integer, _
                    ByRef Cmb1 As ComboBox, ByRef Cmb2 As ComboBox, ByRef Tbox1 As Control, ByRef Tbox2 As Control, ByRef Tbox3 As Control)
   '查覆
   If "" & nGrid.TextMatrix(nRow, colAns1) = "0" Or Trim("" & nGrid.TextMatrix(nRow, colAns1)) = "" Then
      Cmb1.Text = ""
   Else
      Cmb1.Text = nGrid.TextMatrix(nRow, colAns1 + 1)
   End If
   If "" & nGrid.TextMatrix(nRow, colAns1) = "" Then
      Tbox1.Text = ""
   Else
      Tbox1.Text = nGrid.TextMatrix(nRow, colAns1 + 2)
   End If
   '審定號/申請號
   If "" & nGrid.TextMatrix(nRow, colAcase) = "" Then
      Tbox2.Text = ""
   Else
      Tbox2.Text = nGrid.TextMatrix(nRow, colAcase)
   End If
   '覆核
   If "" & nGrid.TextMatrix(nRow, colAns2) = "0" Or Trim("" & nGrid.TextMatrix(nRow, colAns2)) = "" Then
      Cmb2.Text = ""
   Else
      Cmb2.Text = nGrid.TextMatrix(nRow, colAns2 + 1)
   End If
   If "" & nGrid.TextMatrix(nRow, colAns2) = "" Then
      Tbox3.Text = ""
   Else
      Tbox3.Text = nGrid.TextMatrix(nRow, colAns2 + 2)
   End If
   'Added by Lydia 2019/12/12 是否出名
   If iTabs = 0 Then
      If Val("" & nGrid.TextMatrix(nRow, colTQD11)) = 0 Then
          chk2(0).Value = vbUnchecked: chk2(1).Value = vbUnchecked
      Else
          chk2(Val("" & nGrid.TextMatrix(nRow, colTQD11)) - 1).Value = vbChecked
      End If
      txtDt(6).Text = "" & nGrid.TextMatrix(nRow, colTQD11)
      txtDt(6).Tag = txtDt(6).Text
   Else
      If Val("" & nGrid.TextMatrix(nRow, colTQD11)) = 0 Then
          chk2(2).Value = vbUnchecked: chk2(3).Value = vbUnchecked
      Else
          chk2(Val("" & nGrid.TextMatrix(nRow, colTQD11)) + 1).Value = vbChecked
      End If
      txtDt(7).Text = "" & nGrid.TextMatrix(nRow, colTQD11)
      txtDt(7).Tag = txtDt(7).Text
   End If
   'end 2019/12/12
   
   Cmb1.Tag = Cmb1.Text: Cmb2.Tag = Cmb2.Text
   Tbox1.Tag = Tbox1.Text: Tbox2.Tag = Tbox2.Text: Tbox3.Tag = Tbox3.Text
End Sub

Private Sub SetGrd(ByRef nGRD As MSHFlexGrid)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modified by Lydia 2019/12/12 +是否出名,TQD11
   arrGridHeadText = Array("V", "TQD01", "TQD02", "TQD03", "TQD04", "查名組群", "TQD06", "查覆結果", "查覆意見" _
                        , "TQD08", "TQD09", "覆核結果", "覆核意見", "KEYTYPE", "KEYLEN", "是否出名", "TQD11")
   arrGridHeadWidth = Array(200, 0, 0, 0, 0, 860, 0, 860, 860 _
                        , 0, 0, 860, 860, 0, 0, 860, 0)
   nGRD.Visible = False
   nGRD.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To nGRD.Cols - 1
      nGRD.row = 0
      nGRD.col = iRow
      nGRD.Text = arrGridHeadText(iRow)
      nGRD.ColWidth(iRow) = arrGridHeadWidth(iRow)
      nGRD.CellAlignment = flexAlignCenterCenter
   Next
   If colAns1 = 0 Then
      colAns1 = PUB_MGridGetId("TQD06", nGRD) '查覆結果
      colAns2 = PUB_MGridGetId("TQD09", nGRD) '覆核結果
      colAcase = PUB_MGridGetId("TQD08", nGRD) '近似本所案件的審定號/申請號
      colTQD11 = PUB_MGridGetId("TQD11", nGRD) 'Added by Lydia 2019/12/12 是否出名
   End If
   nGRD.Visible = True
End Sub

Public Function QueryData() As Boolean
Dim rStr As String
Dim strTmp1 As String, strTmp2 As String 'Added by Lydia 2020/12/04

On Error GoTo ErrQuery:
QueryData = False

    If m_NoList = "" Or m_NoIdx < 0 Then
       Exit Function
    Else
       tmpArr = Empty
       tmpArr = Split(m_NoList, ",")
       If tmpArr(m_NoIdx) = "" Then
          Exit Function
       Else
                   
          If adoTMQ.State <> adStateClosed Then adoTMQ.Close
          'Modified by Lydia 2016/04/06 抓收文號資料
          'Modified by Lydia 2016/04/29 +CP27,CP14,CP57
          'Modified by Lydia 2016/07/06 +TQC07
          'Modified by Lydia 2019/01/30 +CP06,CP13,CP122,EP06
          'rStr = "SELECT A.*,B.*,S1.ST02 QNAME,S2.ST02 ANAME,TD06,TD10,CP01,CP02,CP03,CP04,CP10,CP27,CP14,CP57,TQC07 " & _
                 "FROM trademarkquery A ,TMQAPP B ,STAFF S1 ,STAFF S2 ,CASEPROGRESS,TMQCASEMAP, " & _
                 "(SELECT TQD01,COUNT(TQD06) TD06,COUNT(TQD09) TD10 FROM TMQDETAIL WHERE TQD06 <> " & CNULL(TMQ_不查) & " OR TQD09 <> " & CNULL(TMQ_不查) & " GROUP BY TQD01) D " & _
                 "WHERE TMQ01='" & tmpArr(m_NoIdx) & "' AND TMQ18=TQA01(+) AND TMQ10=S1.ST01(+) AND TQA02=S2.ST01(+) AND TQA01=D.TQD01(+) AND TMQ21=CP09(+)" & _
                 "AND TQC02(+)=TMQ21 AND TQC03(+)=TMQ01 "
          'Remove by Lydia 2019/05/21 依照各案件的現況 ; 去掉CP06,CP13,CP122,EP06
          'rStr = "SELECT A.*,B.*,S1.ST02 QNAME,S2.ST02 ANAME,TD06,TD10,CP01,CP02,CP03,CP04,CP10,CP27,CP14,CP57,TQC07,CP06,CP13,CP122,EP06 " & _
                 "FROM trademarkquery A ,TMQAPP B ,STAFF S1 ,STAFF S2 ,CASEPROGRESS,TMQCASEMAP, ENGINEERPROGRESS, " & _
                 "(SELECT TQD01,COUNT(TQD06) TD06,COUNT(TQD09) TD10 FROM TMQDETAIL WHERE TQD06 <> " & CNULL(TMQ_不查) & " OR TQD09 <> " & CNULL(TMQ_不查) & " GROUP BY TQD01) D " & _
                 "WHERE TMQ01='" & tmpArr(m_NoIdx) & "' AND TMQ18=TQA01(+) AND TMQ10=S1.ST01(+) AND TQA02=S2.ST01(+) AND TQA01=D.TQD01(+) AND TMQ21=CP09(+)" & _
                 "AND TQC02(+)=TMQ21 AND CP09=EP02(+) AND TQC03(+)=TMQ01 "
          rStr = "SELECT A.*,B.*,S1.ST02 QNAME,S2.ST02 ANAME,TD06,TD10,CP01,CP02,CP03,CP04,CP10,CP27,CP14,CP57,TQC07 " & _
                 "FROM trademarkquery A ,TMQAPP B ,STAFF S1 ,STAFF S2 ,CASEPROGRESS,TMQCASEMAP, " & _
                 "(SELECT TQD01,COUNT(TQD06) TD06,COUNT(TQD09) TD10 FROM TMQDETAIL WHERE TQD06 <> " & CNULL(TMQ_不查) & " OR TQD09 <> " & CNULL(TMQ_不查) & " GROUP BY TQD01) D " & _
                 "WHERE TMQ01='" & tmpArr(m_NoIdx) & "' AND TMQ18=TQA01(+) AND TMQ10=S1.ST01(+) AND TQA02=S2.ST01(+) AND TQA01=D.TQD01(+) AND TMQ21=CP09(+)" & _
                 "AND TQC02(+)=TMQ21 AND TQC03(+)=TMQ01 "
          adoTMQ.CursorLocation = adUseClient
          adoTMQ.Open rStr, cnnConnection, adOpenStatic, adLockReadOnly
          If adoTMQ.RecordCount > 0 Then

             iStiu = 0  '預設不可編輯
             bolModify = False: bolModCheck = False
             strMod(0) = "": strMod(1) = "": strMod(2) = ""
             If IsNull(adoTMQ.Fields("TMQ11")) Then
                mbolSend = False
             Else
                mbolSend = True '已查覆完畢
             End If
             '查覆日期
             txtField(5).Text = TransDate("" & adoTMQ.Fields("TMQ11"), 1)
             cmdSend.BackColor = &HC0C0FF
             '開放在查覆完畢後並且智權人員未收文的情況,可以再修改 ->排除已撤回
             'Modified by Lydia 2016/05/03 排除已發文
             If R_type = "U" And (Pub_StrUserSt03 = "M51" Or strUserNum = adoTMQ.Fields("TMQ10")) And IsNull(adoTMQ.Fields("TQA20")) And Trim("" & adoTMQ.Fields("CP27")) = "" Then
                'Modified by Lydia 2016/03/28 不限階段
                If Not IsNull(adoTMQ.Fields("TMQ11")) Then
                    iStiu = 1: cmdSend.Tag = "U"
                    cmdSend.Caption = "修改完畢(&O)"
                    cmdSend.BackColor = &HFF00& '查覆完畢後再修改，按鈕變綠色
                Else
                    cmdSend.Caption = "查覆完畢(&O)"
                    If IsNull(adoTMQ.Fields("TMQ11")) Then
                       iStiu = 1: cmdSend.Tag = "U"
                    End If
                End If
             ElseIf R_type = "M" Then
                    cmdSend.Caption = "覆核完畢(&O)"
                    'Modified by Lydia 2016/03/28 記錄最後覆核人員
                    'If Len("" & adoTMQ.Fields("TMQ22")) = 0 Then
                       iStiu = 1
                       cmdSend.Tag = "M"
                    'End If
             ElseIf R_type = "Q" Then
                    cmdSend.Caption = "撤　回" '申請的所有委查單尚未輸入結果，可自請撤回申請
                    If Len("" & adoTMQ.Fields("TQA20") & adoTMQ.Fields("TMQ11") & adoTMQ.Fields("TMQ22")) = 0 And Val("" & adoTMQ.Fields("TD06")) = 0 And Val("" & adoTMQ.Fields("TD10")) = 0 Then
                       iStiu = 1
                       cmdSend.Tag = "Q"
                    End If
             ElseIf R_type = "A" Then
                    cmdSend.Caption = "維護完畢(&O)"
                    iStiu = 1
                    cmdSend.Tag = "A"
             End If
             '設定欄位值
             '申請編號
             lblAppNo(0).Caption = "" & adoTMQ.Fields("TQA01")
             '委查單號
             lblAppNo(1).Caption = "" & adoTMQ.Fields("TMQ01")
             '委查人
             txtField(0).Text = "" & adoTMQ.Fields("TQA02")
             lblAppNo(2).Caption = "" & adoTMQ.Fields("ANAME")
             'Added by Lydia 2019/08/12 創新業務組成員可操作清單
             If R_type = "Q" And stIdList = "" Then
                 'Modified by Lydia 2020/12/04 debug-影響到全域變數
                 'stIdList = PUB_GetSalesList(txtField(0).Text, , , , , Pub_StrUserSt15, Pub_StrUserSt15)
                 'If InStr(stIdList, "W") = 0 Or Left(Pub_StrUserSt15, 1) <> "W" Then
                 stIdList = PUB_GetSalesList(txtField(0).Text, , , , , strGrpTmp1, strGrpTmp2)
                 If InStr(stIdList, "W") = 0 Or Left(strGrpTmp1, 1) <> "W" Then
                 'end 2020/12/04
                     stIdList = CNULL(strUserNum) '非創新業務組
                 End If
             End If
             '查名人
             txtField(1).Text = "" & adoTMQ.Fields("TMQ10")
             lblAppNo(3).Caption = "" & adoTMQ.Fields("QNAME")
             '覆核主管
             txtField(2).Text = "" & adoTMQ.Fields("TMQ22")
             If cmdSend.Tag = "M" And txtField(2).Text = "" Then txtField(2).Text = strUserNum
             lblAppNo(4).Caption = GetStaffName(txtField(2).Text)
             '覆核日期
             txtField(17).Text = TransDate("" & adoTMQ.Fields("TMQ23"), 1)
             '委查組群
             txtField(6).Text = "" & adoTMQ.Fields("TMQ03")
             mTQA03 = "" & adoTMQ.Fields("TQA03") '申請編號的全部組群
             '中文筆數
             txtField(7).Text = "" & adoTMQ.Fields("TMQ07")
             '英文筆數
             txtField(8).Text = "" & adoTMQ.Fields("TMQ08")
             '圖形筆數
             txtField(9).Text = "" & adoTMQ.Fields("TMQ09")
             '客戶名稱
             'Modified by Lydia 2021/10/01 txtField(10) => textCName
             textCName.Text = "" & adoTMQ.Fields("TQA04")
             '申請日期
             txtField(3).Text = TransDate("" & adoTMQ.Fields("TQA11"), 1)
             '期限日期
             txtField(4).Text = TransDate("" & adoTMQ.Fields("TMQ06"), 1)
             '指定商品/服務
             'Added by Lydia 2024/07/17 3519組群輸入啟用日
             If adoTMQ.Fields("TQA11") >= "20240718" Then
                '不用顯示指定商品/服務
                LBL1(16).Visible = False
                textService.Visible = False
                Me.SSTab1.Top = 2400
                Me.Height = 7632
             Else
                LBL1(16).Visible = True
                textService.Visible = True
                Me.SSTab1.Top = 2664
                Me.Height = 7884
             'end 2024/07/17
                'Modified by Lydia 2021/10/01 txtField(16) => textService
                textService.Text = "" & adoTMQ.Fields("TQA05")
             End If
             '查名路徑
             txtField(11).Text = "" & adoTMQ.Fields("TMQ24")
             '文字1
             If Not IsNull(adoTMQ.Fields("TQA07")) Then
                txtUnicode(1).Text = adoTMQ.Fields("TQA07") 'Unicode
             ElseIf Not IsNull(adoTMQ.Fields("TQA13")) Then
                txtUnicode(1).Text = adoTMQ.Fields("TQA13") 'Big5
             End If
             '文字2
             If Not IsNull(adoTMQ.Fields("TQA08")) Then
                txtUnicode(2).Text = adoTMQ.Fields("TQA08")
             ElseIf Not IsNull(adoTMQ.Fields("TQA14")) Then
                txtUnicode(2).Text = adoTMQ.Fields("TQA14")
             End If
             '查名單輸入時,已收文
             mTQA15 = "" & adoTMQ.Fields("TQA15")
             '收文－本所案號
             txtField(12).Text = "" & adoTMQ.Fields("TQA16")
             txtField(13).Text = "" & adoTMQ.Fields("TQA17")
             txtField(14).Text = "" & adoTMQ.Fields("TQA18")
             txtField(15).Text = "" & adoTMQ.Fields("TQA19")
             If mTQA15 = "Y" And txtField(12) <> "" Then
                FraCase.Visible = True
                Chk1.Value = 1
             Else
                FraCase.Visible = False
                Chk1.Value = 0
             End If
             
             mTQA20 = "" & adoTMQ.Fields("TQA20")
             mTQD01 = lblAppNo(0).Caption
             mTQD02 = lblAppNo(1).Caption
             '查覆結果已讀(Y=全部開啟過)
             mTMQ19 = "" & adoTMQ.Fields("TMQ19")
             txtField(18).Text = mTMQ19
             mTMQ20 = "" & adoTMQ.Fields("TMQ20") 'Added by Lydia 2024/12/20
             
             '業務收文組群
             'Modified by Lydia 2016/04/19 查名代號已改成Table對應,原欄位保留
             'mTMQ20 = "" & adoTMQ.Fields("TMQ20")
             'txtField(19).Text = mTMQ20
             'Modified by Lydia 2016/04/28 TMQ20改成通知送件日期
             'Remvoe by Lydia 2016/07/06 通知送件日期改成TQC07
             'txtField(19).Text = ChangeWStringToTString("" & adoTMQ.Fields("TMQ20"))
             'Added by Lydia 2016/04/29 +發文日
             txtField(23).Text = ChangeWStringToTString("" & adoTMQ.Fields("CP27"))
             'Added by Lydia 2017/01/25 +收件分發日期
             txtField(24).Text = ChangeWStringToTString("" & adoTMQ.Fields("TMQ05"))
             
             '櫃台收文號
             FirstCP09 = "" & adoTMQ.Fields("TMQ21")
             FirstCP(1) = "" & adoTMQ.Fields("CP01")
             FirstCP(2) = "" & adoTMQ.Fields("CP02")
             FirstCP(3) = "" & adoTMQ.Fields("CP03")
             FirstCP(4) = "" & adoTMQ.Fields("CP04")
             txtField(20).Text = FirstCP09
             If FirstCP09 <> "" Then
                FirstCPP02t = Trim(FirstCP(1)) & CStr(Val(FirstCP(2))) & IIf(FirstCP(3) <> "0" Or FirstCP(4) <> "00", "-" & FirstCP(3), "") & IIf(FirstCP(4) <> "00", "-" & FirstCP(4), "")
                FirstCPP02t = FirstCPP02t & "." & adoTMQ.Fields("CP10") & "." '& mTQD02
             End If
             FirstCP14 = "" & adoTMQ.Fields("CP14") 'Added by Lydia 2016/04/29
             FirstCP57 = "" & adoTMQ.Fields("CP57") 'Added by Lydia 2016/04/29
             'Added by Lydia 2019/01/30 收文號資料
             'Remove by Lydia 2019/05/21 依照各案件的現況
             'm_CP06 = "" & adoTMQ.Fields("CP06")
             'm_CP13 = "" & adoTMQ.Fields("CP13")
             'm_CP122 = "" & adoTMQ.Fields("CP122")
             'm_EP06 = "" & adoTMQ.Fields("EP06")
             
             '目前進度的總收文號
             'Modified by Lydia 2016/06/02 抓目前收文號的發文日
             'lblAppNo(5).Caption = IIf(ShowCP09 <> "", ShowCP09, FirstCP09)
             If ShowCP09 <> "" And FirstCP09 <> ShowCP09 Then
                lblAppNo(5).Caption = ShowCP09
                'Modified by Lydia 2016/07/06 +TQC07
                'strExc(0) = "select CP27 from caseprogress where cp09='" & ShowCP09 & "' "
                'Modified by Lydia 2019/01/30 +CP06,CP13,CP122,EP06
                'strExc(0) = "select CP01,CP02,CP03,CP04,CP27,CP14,CP57,TQC07 from caseprogress,TMQCASEMAP where cp09='" & ShowCP09 & "' AND CP09=TQC02(+) AND TQC03='" & mTQD02 & "' "
                'Modified by Lydia 2019/05/21 依照各案件的現況
                'strExc(0) = "select CP01,CP02,CP03,CP04,CP27,CP14,CP57,TQC07,CP06,CP13,CP122,EP06 from caseprogress,TMQCASEMAP,engineerprogress " & _
                                  "where cp09='" & ShowCP09 & "' AND CP09=TQC02 AND TQC03='" & mTQD02 & "' AND CP09=EP02(+) "
                strExc(0) = "select CP01,CP02,CP03,CP04,CP27,CP14,CP57,TQC07 from caseprogress,TMQCASEMAP " & _
                                  "where cp09='" & ShowCP09 & "' AND CP09=TQC02 AND TQC03='" & mTQD02 & "' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   txtField(23).Text = ChangeWStringToTString("" & RsTemp.Fields("CP27"))
                   'Added by Lydia 2016/07/06
                   txtField(19).Text = ChangeWStringToTString("" & RsTemp.Fields("TQC07"))
                   ShowCP(1) = "" & RsTemp.Fields("CP01")
                   ShowCP(2) = "" & RsTemp.Fields("CP02")
                   ShowCP(3) = "" & RsTemp.Fields("CP03")
                   ShowCP(4) = "" & RsTemp.Fields("CP04")
                   ShowCP14 = "" & RsTemp.Fields("CP14")
                   ShowCP57 = "" & RsTemp.Fields("CP57")
                   'Added by  Lydia 2019/01/30 收文號資料
                   'Remove by Lydia 2019/05/21 依照各案件的現況
                   'm_CP06 = "" & RsTemp.Fields("CP06")
                   'm_CP13 = "" & RsTemp.Fields("CP13")
                   'm_CP122 = "" & RsTemp.Fields("CP122")
                   'm_EP06 = "" & RsTemp.Fields("EP06")
                End If
             Else
                lblAppNo(5).Caption = FirstCP09
                txtField(19).Text = ChangeWStringToTString("" & adoTMQ.Fields("TQC07")) 'Added by Lydia 2016/07/06
             End If
             
             If Len(lblAppNo(5)) > 0 Then
                Label1(0).Visible = True
             Else
                Label1(0).Visible = False
             End If
             '查覆完成日期
             txtField(21).Text = TransDate("" & adoTMQ.Fields("TQA09"), 1)
             '是否撤回
             txtField(22).Text = mTQA20
             
             For Each oText In txtField
                oText.Tag = oText.Text
             Next
             'Added by Lydia 2021/10/01 txtField(10)=>textCName、txtField(16)=>textService
             textService.Tag = textService.Text
             textCName.Tag = textCName.Text
             'end 2021/10/01
             
             For Each oText In txtDt
                oText.Tag = oText.Text
                If R_type <> "A" Then oText.Locked = True 'Added by Lydia 2016/04/25 鎖定欄位
             Next
             FormEnabled '欄位編輯控制
                          
             txtUnicode(1).Tag = txtUnicode(1).Text
             txtUnicode(2).Tag = txtUnicode(2).Text
             
             '抓明細1
             If adoD1.State <> adStateClosed Then adoD1.Close
             'Modified by Lydia 2016/06/02 TMQ_結果查詢改成模組PUB_GetTMQans
             'Modified by Lydia 2019/12/12 +TQD11t, TQD11
             rStr = "SELECT '' V,TQD01,TQD02,TQD03,TQD04,TQD05,TQD06, DECODE(TQD06," & PUB_GetTMQans("3", True) & "),TQD07," & _
                    "TQD08,TQD09, DECODE(TQD09," & PUB_GetTMQans("3", True) & ") ,TQD10,(F1.TQF05) KEYTYPE,(F1.TQF06) KEYLEN,DECODE(TQD11,'1','第三人','2','不出名','') TQD11T,TQD11 " & _
                    "FROM TMQDETAIL, TMQFILE F1, TMQFILE F2 WHERE TQD01=F1.TQF01(+) AND F1.TQF02(+)='" & TMQ_附件F02 & "' AND TQD03=F1.TQF03(+) AND F1.TQF04(+)='" & TMQ_附件F04 & "' AND TQD01=F2.TQF01(+) AND TQD02=F2.TQF02(+) AND TQD03=F2.TQF03(+) AND TQD04=F2.TQF04(+) "
             
             strExc(1) = rStr & "and tqd01='" & "" & adoTMQ.Fields("TQA01") & "' and tqd02='" & "" & adoTMQ.Fields("TMQ01") & "' and tqd03 in ('" & TMQ_AkindPic & "','" & TMQ_AkindWord1 & "') order by 2,3,4,5 "
             adoD1.CursorLocation = adUseClient
             adoD1.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
             If adoD1.RecordCount > 0 Then
                mTQD03 = adoD1.Fields("TQD03")
                 '文字查名不可輸入查名路徑
                If mTQD03 <> TMQ_AkindPic Then txtField(11).Locked = True
                
                'Added by Lydia 2024/07/17
                cmdRoute.Caption = "顯示"
                cmdRoute.Visible = False
                If mTQD03 = TMQ_AkindPic And cmdRoute.Tag <> "" Then
                   txtField(11).Width = 3300
                   If cmdRoute.Tag = "M" Then
                      cmdRoute.Caption = "輸入"
                      cmdRoute.Visible = True
                   ElseIf DBDATE(txtField(3)) >= "20240718" Then
                      cmdRoute.Visible = True
                   End If
                Else
                   txtField(11).Width = 3600
                End If
                'end 2024/07/17
                
                Set GRD1.Recordset = adoD1
                Call SetGrd(GRD1)
                If Not IsNull(adoD1.Fields("KEYTYPE")) Then
                   tmpKeyPic1.Visible = True

                   If UCase("" & adoD1.Fields("KEYTYPE")) <> "" Then
                        If mbolCall = False Then cmdKey(0).Visible = True
                        cmdKey(0).Tag = adoD1.Fields("KEYTYPE") '記錄查名內容附件的類型
                        cmdKD(0).Visible = True
                        cmdKD(0).Tag = adoD1.Fields("KEYTYPE")
                        If adoD1.Fields("TQD03") = TMQ_AkindPic Then
                           Me.SSTab1.TabCaption(0) = "圖形"
                           'Modified by Lydia 2016/06/23
                           'If KeyFileGet(mTQD01, TMQ_AkindPic) = False Then
                           If KeyFileGet(mTQD01, TMQ_AkindPic, , , cmdKey(0).Tag) = False Then
                           End If
                        Else
                           'Modified by Lydia 2016/03/28 只顯示文字1
                           'Me.SSTab1.TabCaption(0) = "文字1"
                           Me.SSTab1.TabCaption(0) = "文字"
                           'Modified by Lydia 2016/06/23
                           'If KeyFileGet(mTQD01, TMQ_AkindWord1) = False Then
                           If KeyFileGet(mTQD01, TMQ_AkindWord1, , , cmdKey(0).Tag) = False Then
                           End If
                        End If
                        If InStr(UCase("" & adoD1.Fields("KEYTYPE")), "PDF") > 0 Then
                           Set tmpKeyImg1.Picture = tmpInsPDF.Picture
                        End If
                   End If
                Else
                   tmpKeyPic1.Visible = False
                   'Added by Lydia 2023/06/08 改頁籤抬頭
                   If "" & adoD1.Fields("TQD03") = TMQ_AkindPic Then
                      Me.SSTab1.TabCaption(0) = "圖形"
                   Else
                      Me.SSTab1.TabCaption(0) = "文字1"
                   End If
                   'end 2023/06/08
                End If
             End If
             
             If adoD2.State <> adStateClosed Then adoD2.Close
             '抓明細2
             strExc(2) = rStr & "and tqd01='" & "" & adoTMQ.Fields("TQA01") & "' and tqd02='" & "" & adoTMQ.Fields("TMQ01") & "' and tqd03='" & TMQ_AkindWord2 & "' order by 2,3,4,5 "
             adoD2.CursorLocation = adUseClient
             adoD2.Open strExc(2), cnnConnection, adOpenStatic, adLockReadOnly
             If adoD2.RecordCount > 0 Then
                Set grd2.Recordset = adoD2
                Call SetGrd(grd2)
                 Me.SSTab1.TabVisible(1) = True
                 
                If Not IsNull(adoD2.Fields("KEYTYPE")) Then
                    If mbolCall = False Then cmdKey(1).Visible = True
                    cmdKey(1).Tag = adoD2.Fields("KEYTYPE")
                    cmdKD(1).Visible = True
                    cmdKD(1).Tag = adoD2.Fields("KEYTYPE")
                    tmpKeyPic2.Visible = True
                    If InStr(UCase("" & adoD2.Fields("KEYTYPE")), "PDF") = 0 Then
                       'Modified by Lydia 2016/06/23
                       'If KeyFileGet(mTQD01, TMQ_AkindWord2) = False Then
                       If KeyFileGet(mTQD01, TMQ_AkindWord2, , , cmdKey(1).Tag) = False Then
                       End If
                    Else
                       Set tmpKeyImg2.Picture = tmpInsPDF.Picture
                    End If
                Else
                    cmdKey(1).Visible = False
                    tmpKeyPic2.Visible = False
                End If
             Else
                Me.SSTab1.TabVisible(1) = False
             End If
             
             If adoD1.RecordCount = 0 And adoD2.RecordCount = 0 Then
                MsgBox "查無此委查單號資料!", vbCritical, "查覆明細作業"
                GoTo ExitClose
             End If
          Else
             MsgBox "查無此委查單號資料!", vbCritical, "查覆明細作業"
             GoTo ExitClose
          End If
               
          QueryData = True
          Me.SSTab1.Tab = 0
            '若有資料游標停在第一筆
            GRD1.Visible = False
            GRD1.col = 0
            GRD1.row = 1
            dblPrevRow = GRD1.row
            If adoD1.RecordCount > 0 Then
               GRD1.Text = "V"
               For jj = 0 To GRD1.Cols - 1
                  GRD1.col = jj
                  GRD1.CellBackColor = &HFFC0C0
               Next jj
               'Modified by Lydia 2019/12/12 +Tab
               Call SetComboAns(0, GRD1, dblPrevRow, Cbo1, Cbo2, txtDt(0), txtDt(1), txtDt(2))
            End If
            GRD1.Visible = True
          If Me.SSTab1.TabVisible(1) = True Then
            grd2.Visible = False
            grd2.col = 0
            grd2.row = 1
            dblPrevRow2 = grd2.row
            If adoD2.RecordCount > 0 Then
               grd2.Text = "V"
               For jj = 0 To grd2.Cols - 1
                  grd2.col = jj
                  grd2.CellBackColor = &HFFC0C0
               Next jj
               'Modified by Lydia 2019/12/12 +Tab
               Call SetComboAns(1, grd2, dblPrevRow2, Cbo3, Cbo4, txtDt(3), txtDt(4), txtDt(5))
            End If
            grd2.Visible = True
            '---不鎖定明細列
            Call AttachFileRead(1, mTQD01, mTQD02, "2")
            Call ShowFieldTQD09(1)
          End If
            '預設畫面-文字1
            Call ShowFieldTQD09(0)
            '---不鎖定明細列
            Call AttachFileRead(0, mTQD01, mTQD02, mTQD03)
            If Cbo1.Enabled = True Then Cbo1.SetFocus

          GoTo ExitClose
       End If 'If TmpArr(m_NoIdx) = ""
    End If 'If m_NoList = ""

ErrQuery:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
ExitClose:
    Set adoTMQ = Nothing
    Set adoD1 = Nothing
    Set adoD2 = Nothing
End Function
'Added by Lydia 2016/06/23 +mTQF05檔案類型
Private Function KeyFileGet(mTQF01 As String, mTQFkind As String, Optional mLoad As Boolean = True, Optional mPath As String, Optional mTQF05 As String) As Boolean
Dim adoRst As New ADODB.Recordset
Dim outType As String
Dim stTempFile As String
Dim fileN As Integer
Dim bytes() As Byte

On Error GoTo ErrHnd
   
    KeyFileGet = False
    '開啟時,無法刪除,預設下次開啟表單執行刪檔
    If Dir(m_AttachPath & "\HM*.jpg") <> "" Then
        Kill m_AttachPath & "\HM*.jpg"
    End If
    If Dir(m_AttachPath & "\HM*.pdf") <> "" Then
        Kill m_AttachPath & "\HM*.pdf"
    End If
    If Dir(m_AttachPath & "\HM*.JPG") <> "" Then
        Kill m_AttachPath & "\HM*.JPG"
    End If
    If Dir(m_AttachPath & "\HM*.PDF") <> "" Then
        Kill m_AttachPath & "\HM*.PDF"
    End If
    'Modified by Lydia 2016/06/23 改放在FTP
    'If adoRst.State <> adStateClosed Then adoRst.Close
    'Set adoRst = Nothing
    'adoRst.CursorLocation = adUseClient
    'adoRst.Open "select * from TMQFile where TQF01='" & mTQF01 & "' AND TQF02='" & TMQ_附件F02 & "' AND TQF03='" & mTQFkind & "' AND TQF04='" & TMQ_附件F04 & "'", cnnConnection, adOpenStatic, adLockOptimistic
    'If adoRst.RecordCount > 0 Then

    '   outType = "" & adoRst.Fields("TQF05")
    '   stTempFile = m_AttachPath & "\" & mTQF01 & "_" & mTQFkind & "." & LCase(Trim(outType))
    '   mPath = stTempFile
       
    '    ReDim bytes(Val(adoRst.Fields("TQF06").Value))
    '    bytes() = adoRst.Fields("TQF07").GetChunk(Val(adoRst.Fields("TQF06").Value))
    '    fileN = FreeFile
    '    Open stTempFile For Binary Access Write As #fileN
    '    Put #fileN, , bytes()
    '    Close #fileN
    outType = UCase(Trim(mTQF05))
    stTempFile = mTQF01 & TMQ_附件F02 & mTQFkind & TMQ_附件F04 & "." & outType
    
    If PUB_TMQGetAFile(m_AttachPath, stTempFile, mTQD01, TMQ_附件F02, mTQFkind, TMQ_附件F04, outType) = False Then
       MsgBox "無法儲存檔案[ " & stTempFile & " ]！"
       Exit Function
    Else
       If InStr(UCase(outType), "PDF") = 0 Then
          Set G_SeekPicColor.Picture = pvGetStdPicture(Trim(stTempFile))
          Select Case mTQFkind
             Case TMQ_AkindPic, TMQ_AkindWord1
                 '固定PictureBox中的image,載入圖片後調整圖片大小
                 Call Pub_PicToObj(Trim(stTempFile), G_SeekPicColor, tmpKeyPic1, tmpKeyImg1)
             Case TMQ_AkindWord2
                 Call Pub_PicToObj(Trim(stTempFile), G_SeekPicColor, tmpKeyPic2, tmpKeyImg2)
          End Select
       End If
    End If
    'Else
    '   Exit Function
    'End If
    'end 2016/06/23
    
    KeyFileGet = True
    Exit Function

ErrHnd:
   MsgBox Err.Description, vbCritical
   
   If fileN > 0 Then Close #fileN
End Function

Private Sub Cbo1_LostFocus()
   If iStiu = 1 Then
      If (Cbo1.Text = TMQ_近似T1 Or Cbo1.Text = TMQ_近似T2) And Cbo1.Text <> Cbo1.Tag And txtDt(1).Text = "" Then
         txtDt(1).Text = mPrevTM1215
      End If
   End If
End Sub

Private Sub Cbo1_Validate(Cancel As Boolean)
    Call ShowFieldTQD09(0)
End Sub

Private Sub Cbo3_LostFocus()
   If iStiu = 1 Then
      If (Cbo3.Text = TMQ_近似T1 Or Cbo3.Text = TMQ_近似T2) And Cbo3.Text <> Cbo3.Tag And txtDt(4).Text = "" Then
         txtDt(4).Text = mPrevTM1215
      End If
   End If
End Sub
Private Sub Cbo3_Validate(Cancel As Boolean)
    Call ShowFieldTQD09(1)
End Sub

'查名單維護-變更查名附件
Private Sub cmdAFDel_Click(Index As Integer)

On Error GoTo ErrHand

   bolSave = IsSaveData '判斷是否已存檔
   If IsSaveAppNo = False Then
      Exit Sub
   End If
   
   If txtUnicode(Index + 1).Text <> txtUnicode(Index + 1).Tag Then
       MsgBox "文字" & Index + 1 & "已變更,請先按" & Mid(cmdSend.Caption, 1, 4) & "!", vbInformation
       Exit Sub
   End If
   If cmdKey(Index).Tag = "" Then
      MsgBox "無資料,可供刪除!"
      Exit Sub
   End If
   
   Call GetTQF0304(Index, mTQF03, mTQF04)
   
   If PUB_TMQFileIsExist(mTQD01, TMQ_附件F02, mTQF03, TMQ_附件F04) Then
      If MsgBox("已有查名內容附件,是否刪除?", vbCritical + vbYesNo, "查名內容") = vbYes Then
         If txtUnicode(Index + 1).Text = "" Then
            MsgBox "若刪除附件將無查名內容,請檢查文字或附件!", vbCritical
            Exit Sub
         Else
            cnnConnection.BeginTrans
               'Modified by Lydia 2016/06/23 改放在FTP
                ' strExc(0) = "delete from TMQFILE where TQF01='" & mTQD01 & "' AND TQF02='" & TMQ_附件F02 & "' AND TQF03='" & mTQF03 & "' AND TQF04='" & TMQ_附件F04 & "'"
                ' cnnConnection.Execute strExc(0), intI
                If PUB_TMQAFileDel(mTQD01, TMQ_附件F02, mTQF03, TMQ_附件F04) Then
                    cmdKey(Index).Tag = ""
                    cmdKey(Index).Visible = False
                    cmdKD(Index).Visible = False
                    MsgBox "存檔完成,請重新呼叫委查單!", vbInformation
                End If
            cnnConnection.CommitTrans
            'end 2016/06/23
         End If
      End If
   End If
   
   Exit Sub
   
ErrHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub
'查名單維護-變更查名附件
Private Sub cmdAFUpd_Click(Index As Integer)

Dim stFileName As String, stFolderPath As String, stFullName As String
Dim tmpBol As Boolean
   bolSave = IsSaveData '判斷是否已存檔

   If IsSaveAppNo = False Then
      Exit Sub
   End If
   
   Call GetTQF0304(Index, mTQF03, mTQF04)
   If PUB_TMQFileIsExist(mTQD01, TMQ_附件F02, mTQF03, TMQ_附件F04) Then
      If MsgBox("已有查名內容附件,是否覆蓋原有內容?", vbCritical + vbYesNo, "查名內容") = vbNo Then
         Exit Sub
      End If
   End If

    cmdAFUpd(Index).Enabled = False
    stFullName = GetSaveName(stFileName)
    If stFullName <> "" Then
      If stFullName <> "" Then
        If InStr(CStr(stFullName), "#") > 0 Then
           MsgBox CStr(stFullName) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
           Exit Sub
        End If
        If InStr(UCase(".pdf.jpg"), UCase(Right(CStr(stFullName), 4))) = 0 Then
           MsgBox "只能插入PDF或JPG檔！"
           Exit Sub
        End If
        If PUB_TMQAFileSave(mTQD01, TMQ_附件F02, mTQF03, TMQ_附件F04, UCase(Right(CStr(stFullName), 3)), stFullName) Then
           cmdKey(Index).Visible = True
           cmdKD(Index).Visible = True
           cmdKey(Index).Tag = UCase(Right(CStr(stFullName), 3))
           MsgBox "存檔完成,請重新呼叫委查單!", vbInformation
           If Val(mTQD03) = 0 Then
              strExc(9) = "圖形查詢:變更內容"
           Else
              strExc(9) = "文字" & mTQF03 & ":變更內容"
           End If
           'Modified by Lydia 2021/10/01 txtField(10) => textCName
           Call ChangeDataMail(mTQD01, strExc(9), textCName)
        Else
           MsgBox "存檔失敗!", vbCritical
        End If
      End If
    End If
    cmdAFUpd(Index).Enabled = True

End Sub

'結束
Private Sub cmdExit_Click()
   bolSave = IsSaveData '判斷是否已存檔
   If cmdSend.Tag = "A" Then
      If IsSaveAppNo = False Then
      End If
   '判斷不查，有修改筆數(結果為不查)
   ElseIf bolModify = True And bolModCheck = False And strMod(1) <> "" Then
       If InStr(strMod(1), "不查") > 0 Then
          Call cmdSend_Click
       End If
   End If
   
   'Added by Lydia 2016/09/12 結束前檢查是否有按覆核完畢(發mail通知)
   If bolChgTM22 = True Then
      If MsgBox("覆核結果有更改，是否要繼續執行覆核完畢?", vbYesNo + vbDefaultButton1) = vbYes Then
         Call cmdSend_Click
      End If
   End If
   'end 2016/09/12
   
   Unload Me
End Sub

Private Sub cmdKD_Click(Index As Integer)
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim ii As Integer, oList As ListBox

   bolSave = IsSaveData '判斷是否已存檔
   
   Screen.MousePointer = vbHourglass

   Call GetTQF0304(Index, mTQF03)
    
   stFileName = mTQD01 & mTQF03 & "." & cmdKD(Index).Tag
      
    stFullName = GetSaveName(stFileName)
    If stFullName <> "" Then
      If Dir(stFullName) <> "" Then
        If MsgBox("檔案[ " & stFullName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
           stFullName = ""
        End If
      End If
      If stFullName <> "" Then
          If PUB_TMQGetAFile("", stFullName, mTQD01, TMQ_附件F02, mTQF03, TMQ_附件F04, cmdKD(Index).Tag) = False Then
             MsgBox "無法儲存檔案[ " & stFullName & " ]！"
             GoTo RunExit
          End If
      End If
    End If
    
    If stFullName <> "" Then
       MsgBox Me.SSTab1.TabCaption(Index) & " 下載完成！", vbInformation, "申請附件"
    End If

RunExit:
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmdKey_Click(Index As Integer)
   bolSave = IsSaveData '判斷是否已存檔

   Call GetTQF0304(Index, mTQF03)
   Tmpfrm090129.SetParent Me, mTQD01, Val(mTQF03), cmdKey(Index).Tag
   Tmpfrm090129.Show
   If Tmpfrm090129.iStiu = 1 Then

   Else
      Unload Tmpfrm090129
   End If
End Sub

Private Sub cmdNext_Click()
  bolSave = IsSaveData '判斷是否已存檔
  FormReset
  m_NoIdx = m_NoIdx + 1
  If QueryData() = False Then
     If m_NoIdx > 0 Then m_NoIdx = m_NoIdx - 1
     Call cmdExit_Click
  End If
End Sub

Private Sub cmdPrevious_Click()
  bolSave = IsSaveData '判斷是否已存檔
  
  FormReset
  If m_NoIdx > 0 Then
     m_NoIdx = m_NoIdx - 1
    If QueryData() = False Then
       If m_NoIdx > 0 Then m_NoIdx = m_NoIdx + 1
       Call cmdExit_Click
    End If
  Else
    Call cmdExit_Click
  End If
End Sub

'明細-存檔
Private Sub cmdSaveD_Click(Index As Integer)
Dim tmpAns As String
Dim tmpBol As Boolean

   Select Case Index
   Case 0, 1
      If dblPrevRow <= 0 Then
         MsgBox "請選取查名組群", vbInformation
         Exit Sub
      End If
      If Index = 0 Then
         If CheckCombo(Cbo1) = False Then Exit Sub

         strExc(1) = "查覆"
         If (Cbo1.Text = TMQ_近似T1 Or Cbo1.Text = TMQ_近似T2) Then
            If txtDt(1).Text = "" Then
               MsgBox "請輸入" & Left(Cbo1.Text, 2) & "的本所案件之申請號或審定號!", vbCritical
               Call ShowFieldTQD09(0)
               GRD1.row = dblPrevRow
               txtDt(1).SetFocus: Exit Sub
            Else
               Call txtDt_Validate(1, tmpBol)
               If tmpBol = True Then
                  GRD1.row = dblPrevRow
                  txtDt(1).SetFocus: Exit Sub
               End If
            End If
         ElseIf txtDt(1).Text <> "" Then
               MsgBox "非與本所案件" & Left(Cbo1.Text, 2) & "，請勿輸入申請號或審定號!", vbCritical
               GRD1.row = dblPrevRow
               Exit Sub
         End If
      ElseIf Cbo1.Text <> TMQ_近似T1 And Cbo1.Text <> TMQ_近似T2 Then
            MsgBox "請選擇結果為" & TMQ_近似T1 & " / " & TMQ_近似T2 & "的組群!", vbCritical
            GRD1.row = dblPrevRow: Exit Sub
      Else
         'Modified by Lydia 2016/06/01
         'If CheckCombo(Cbo2) = False Then Exit Sub
         If CheckCombo(Cbo2, "A") = False Then Exit Sub
         strExc(1) = "覆核"
         'Added by Lydia 2019/12/12 提醒
         If Cbo2.Text <> Cbo2.Tag And Cbo2.Text <> "核可" And (chk2(0).Value = vbChecked Or chk2(1).Value = vbChecked) Then
             If MsgBox("非核可案卻勾選" & IIf(chk2(0).Value = vbChecked, "第三人出名", "不出名") & "，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                 Exit Sub
             End If
         End If
      End If
            
      'Modified by Lydia 2019/12/12 +txtDT(6)
      tmpAns = SaveDetailAns(strExc(1), GRD1, dblPrevRow, Cbo1, Cbo2, txtDt(0), txtDt(1), txtDt(2), txtDt(6), Index)
   Case 2, 3
      If dblPrevRow2 <= 0 Then
         MsgBox "請選取查名組群", vbInformation
         Exit Sub
      End If
      If Index = 2 Then
         If CheckCombo(Cbo3) = False Then Exit Sub
         
         strExc(1) = "查覆"
         If (Cbo3.Text = TMQ_近似T1 Or Cbo3.Text = TMQ_近似T2) Then
            If txtDt(4).Text = "" Then
               MsgBox "請輸入" & Left(Cbo3.Text, 2) & "的本所案件之申請號或審定號!", vbCritical
               Call ShowFieldTQD09(1)
               grd2.row = dblPrevRow2
               txtDt(4).SetFocus: Exit Sub
            Else
               Call txtDt_Validate(4, tmpBol)
               If tmpBol = True Then
                  grd2.row = dblPrevRow2
                  txtDt(4).SetFocus: Exit Sub
               End If
            End If
         ElseIf txtDt(4).Text <> "" Then
               MsgBox "非與本所案件" & Left(Cbo1.Text, 2) & "，請勿輸入申請號或審定號!", vbCritical
               grd2.row = dblPrevRow2
               Exit Sub
         End If
      ElseIf Cbo3.Text <> TMQ_近似T1 And Cbo3.Text <> TMQ_近似T2 Then
            MsgBox "請選擇結果為" & TMQ_近似T1 & " / " & TMQ_近似T2 & "的組群!", vbCritical
            grd2.row = dblPrevRow2: Exit Sub
      Else
         'Modified by Lydia 2016/06/01
         'If CheckCombo(Cbo4) = False Then Exit Sub
         If CheckCombo(Cbo4, "A") = False Then Exit Sub
         strExc(1) = "覆核"
         'Added by Lydia 2019/12/12 提醒
         If Cbo4.Text <> Cbo4.Tag And Cbo4.Text <> "核可" And (chk2(2).Value = vbChecked Or chk2(3).Value = vbChecked) Then
             If MsgBox("非核可案卻勾選" & IIf(chk2(2).Value = vbChecked, "第三人出名", "不出名") & "，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                 Exit Sub
             End If
         End If
      End If
      'Modified by Lydia 2019/12/12 +txtDT(7)
      tmpAns = SaveDetailAns(strExc(1), grd2, dblPrevRow2, Cbo3, Cbo4, txtDt(3), txtDt(4), txtDt(5), txtDt(7), Index)
   End Select
   '回傳N=無資料寫入,T=存檔成功,F=存檔失敗
End Sub

'Modified by Lydia 2019/12/12 + mTXT4
'Modified by Lydia 2021/10/01 TextBox => Control
Private Function SaveDetailAns(ByVal iTyp As String, ByRef mGRID As MSHFlexGrid, ByRef mRow As Integer, ByRef mCMB As ComboBox, ByRef mCMB2 As ComboBox, _
                                  ByRef mTXT As Control, ByRef mTXT2 As Control, ByRef mTXT3 As Control, ByRef mTXT4 As Control, Optional ByVal rKind As String) As String

Dim tmpX As Integer
Dim tmpE As Integer, tmpC As Integer
Dim strUpd As String

On Error GoTo ErrSaveDetail
   
   'Modified by Lydia 2019/12/12 + mTXT4
   If mCMB.Text = "" And mCMB2.Text = "" And mTXT.Text = "" And mTXT2.Text = "" And mTXT3.Text = "" _
      And mCMB.Tag = mCMB.Text And mCMB2.Tag = mCMB2.Text And mTXT.Tag = mTXT.Text _
      And mTXT2.Tag = mTXT2.Text And mTXT3.Tag = mTXT3.Text And mTXT4.Tag = mTXT4.Text Then    '沒有輸入
      SaveDetailAns = "N"
   Else '有輸入
      tmpX = 0
      tmpArr = Empty
      'Modified by Lydia 2016/06/02
      'TmpArr = Split(TMQ_結果清單, ",")
      tmpArr = Split(PUB_GetTMQans("2", True), ",")
      For jj = 0 To UBound(tmpArr)
        If Mid(tmpArr(jj), 3, Len(tmpArr(jj)) - 2) = IIf(iTyp = "查覆", mCMB.Text, mCMB2.Text) Then
           tmpX = Val(Mid(tmpArr(jj), 1, 1))
           Exit For
        End If
      Next jj

      If iTyp = "查覆" Then 'TQD06,TQD07,TQD08
        If tmpX = 0 Then
           strExc(1) = ",TQD06=null"
           mGRID.TextMatrix(mRow, colAns1) = ""
        Else
           strExc(1) = ",TQD06=" & CNULL(Trim(tmpX))
           mGRID.TextMatrix(mRow, colAns1) = tmpX
        End If
        mGRID.TextMatrix(mRow, colAns1 + 1) = mCMB.Text
        '意見
        strExc(1) = strExc(1) & ",TQD07='" & Trim(mTXT.Text) & "'"
        mGRID.TextMatrix(mRow, colAns1 + 2) = Trim(mTXT.Text)
        '審定號/申請號
        strExc(1) = strExc(1) & ",TQD08='" & Trim(mTXT2.Text) & "'"
        mGRID.TextMatrix(mRow, colAcase) = Trim(mTXT2.Text)
        
        If mbolSend = True And R_type <> "A" Then
           bolModify = True
           '查名組群
           If rKind < 2 Then
              strExc(5) = Me.SSTab1.TabCaption(0)
           Else
              strExc(5) = Me.SSTab1.TabCaption(1)
           End If
           'Modified by Lydia 2016/04/21 +已收文案件提示(第一次收文案件)
           'Modified by Lydia 2021/10/01 txtField(10) => textCName
           strMod(0) = "「" & textCName.Text & "」" & GetStrTitle & " 委查單: " & lblAppNo(1).Caption & _
                      IIf(FirstCP09 <> "", "，已收文案件" & FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) <> "000", "-" & FirstCP(3) & "-" & FirstCP(4), ""), "") & _
                      "，查名結果有變更!!"
           strMod(2) = IIf(strMod(2) = "", txtField(0).Text, strMod(2))
           If mCMB.Text <> mCMB.Tag Then
              If InStr(mCMB.Text & "," & mCMB.Tag, TMQ_近似T1) > 0 Or InStr(mCMB.Text & "," & mCMB.Tag, TMQ_近似T2) > 0 Then
                If rKind < 2 Then
                   Call txtDt_Validate(1, False)
                Else
                   Call txtDt_Validate(4, False)
                End If
              End If
              strMod(1) = strMod(1) & strExc(5) & " 組群" & PUB_MGridGetValue(mRow, "查名組群", mGRID) & " 查名結果由" & CNULL(mCMB.Tag) & "->" & CNULL(mCMB.Text) & " ;" & vbCrLf
                '不查->重計筆數
                tmpE = 0: tmpC = 0: strUpd = ""
                If mCMB.Text = "不查" Or mCMB.Tag = "不查" Then
                   If mTQD03 = TMQ_AkindPic Then
                       strUpd = "UPDATE trademarkquery SET tmq09=tmq09-1 where tmq01='" & mTQD02 & "' "
                       If mCMB.Text = "不查" Then
                          txtField(9) = Val(txtField(9)) - 1
                       Else
                          txtField(9) = Val(txtField(9)) + 1
                       End If
                   Else
                     If rKind < 2 Then
                        Call PUB_CountTxtNEC(tmpE, tmpC, txtUnicode(1).Text)
                        If cmdKey(0).Visible = True And txtUnicode(1).Text = "" Then tmpC = tmpC + 1
                     Else
                        Call PUB_CountTxtNEC(tmpE, tmpC, txtUnicode(2).Text)
                        If cmdKey(1).Visible = True And txtUnicode(2).Text = "" Then tmpC = tmpC + 1
                     End If
                     strUpd = "UPDATE trademarkquery SET tmq07=tmq07-" & tmpC & ",tmq08=tmq08-" & tmpE & " where tmq01='" & mTQD02 & "' "
                        If mCMB.Text = "不查" Then
                            txtField(7) = Val(txtField(7)) - tmpC
                            txtField(8) = Val(txtField(8)) - tmpE
                        Else
                            txtField(7) = Val(txtField(7)) + tmpC
                            txtField(8) = Val(txtField(8)) + tmpE
                        End If
                   End If
                End If
           End If
           
           If mTXT.Text <> mTXT.Tag Then
              strMod(1) = strMod(1) & strExc(5) & " 組群" & PUB_MGridGetValue(mRow, "查名組群", mGRID) & " 查名意見由" & Replace(CNULL(mTXT.Tag), "NULL", "空白") & "->" & CNULL(mTXT.Text) & " ;" & vbCrLf
           End If
           If mTXT2.Text <> mTXT2.Tag Then
              strMod(1) = strMod(1) & strExc(5) & " 組群" & PUB_MGridGetValue(mRow, "查名組群", mGRID) & " 審定號/申請號由" & Replace(CNULL(mTXT2.Tag), "NULL", "空白") & "->" & CNULL(mTXT2.Text) & " ;" & vbCrLf
           End If
        End If
      Else '覆核 TQD09,TQD10
        If tmpX = 0 Then
           strExc(1) = ",TQD09=null"
           mGRID.TextMatrix(mRow, colAns2) = ""
        Else
           strExc(1) = ",TQD09=" & CNULL(Trim(tmpX))
           mGRID.TextMatrix(mRow, colAns2) = tmpX
        End If
        mGRID.TextMatrix(mRow, colAns2 + 1) = mCMB2.Text
        '意見
        strExc(1) = strExc(1) & ",TQD10='" & Trim(mTXT3.Text) & "'"
        mGRID.TextMatrix(mRow, colAns2 + 2) = Trim(mTXT3.Text)
        'Added by Lydia 2023/07/06 開放覆核主管可修改「審定號/申請號」
        '審定號/申請號
        strExc(1) = strExc(1) & ",TQD08='" & Trim(mTXT2.Text) & "'"
        mGRID.TextMatrix(mRow, colAcase) = Trim(mTXT2.Text)
        'end 2023/07/06
        'Added by Lydia 2019/12/12 是否出名
        strExc(1) = strExc(1) & ",TQD11='" & Trim(mTXT4.Text) & "'"
        strExc(2) = ""
        If mTXT4.Text = "1" Then
            strExc(2) = "第三人"
        ElseIf mTXT4.Text = "2" Then
            strExc(2) = "不出名"
        End If
        mGRID.TextMatrix(mRow, colTQD11) = Trim(mTXT4.Text)
        mGRID.TextMatrix(mRow, colTQD11 - 1) = strExc(2)
        'end 2019/12/12
        
        bolChgTM22 = True 'Added by Lydia 2016/09/12
      End If
      
      strSql = "UPDATE TMQDETAIL SET " & Mid(strExc(1), 2, Len(strExc(1)) - 1) & " WHERE TQD01='" & mGRID.TextMatrix(mRow, 1) & "' " & _
               "AND TQD02='" & mGRID.TextMatrix(mRow, 2) & "' AND TQD03='" & mGRID.TextMatrix(mRow, 3) & "' AND TQD04='" & mGRID.TextMatrix(mRow, 4) & "' "
               
      cnnConnection.BeginTrans
        cnnConnection.Execute strSql, intI
        If bolModify = True Or R_type = "A" Then Pub_SeekTbLog strSql
        
        If strUpd <> "" Then
           cnnConnection.Execute strUpd, intI
           txtField(7).Tag = txtField(7).Text
           txtField(8).Tag = txtField(8).Text
           txtField(9).Tag = txtField(9).Text
           bolModCheck = True
        End If
        SaveDetailAns = "T"
      cnnConnection.CommitTrans
      mCMB.Tag = mCMB.Text
      mTXT.Tag = mTXT.Text
      mCMB2.Tag = mCMB2.Text
      mTXT2.Tag = mTXT2.Text
      mTXT3.Tag = mTXT3.Text
      mTXT4.Tag = mTXT4.Text 'Added by Lydia 2019/12/12
   End If

ErrSaveDetail:
   If Err.Number <> 0 Then
      SaveDetailAns = "F"
      MsgBox "存檔失敗:" & Err.Description & vbCrLf & "請重新進入系統", vbCritical, "查覆結果"
   End If
End Function
'Added by Lydia 2016/04/29
Private Sub cmdSendMail_Click()

On Error GoTo ErrHand01
   
   'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
   'If txtField(0) <> strUserNum And cmdSendMail.Visible = True Then
   'Modified by Lydia 2019/12/25 開放特殊設定權限
   'If InStr(stIdList, txtField(0)) = 0 And cmdSendMail.Visible = True Then
   If cmdSendMail.Visible = True Then
      strExc(1) = "N"
      If InStr(stIdList, txtField(0)) > 0 Then
          strExc(1) = ""
      '代理-總經理、A7
      ElseIf bolSpecMan = True And InStr(strSpecCode, txtField(0)) > 0 Then
          strExc(1) = ""
      End If
      If strExc(1) = "N" Then
      'end 2019/12/25
         MsgBox "無權限!!", vbCritical
         Exit Sub
      End If
   End If
   If FirstCP09 = "" Then
      MsgBox "委查單: " & mTQD01 & " 未收文", vbCritical
      Exit Sub
   End If
   'Added by Lydia 2016/07/07
   If FirstCP09 <> ShowCP09 And ShowCP09 <> "" Then
      strExc(2) = "本所案號: " & ShowCP(1) & "-" & ShowCP(2) & IIf(ShowCP(3) & ShowCP(4) = "000", "", "-" & ShowCP(3) & "-" & ShowCP(4))
   Else
      strExc(2) = "本所案號: " & FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) = "000", "", "-" & FirstCP(3) & "-" & FirstCP(4))
   End If
   'Modified by Lydia 2016/07/07
   'If FirstCP57 <> "" Then
   If (lblAppNo(5).Caption = FirstCP09 And FirstCP57 <> "") Or (lblAppNo(5).Caption = ShowCP09 And ShowCP57 <> "") Then
      MsgBox strExc(2) & " 已取消收文", vbCritical
      Exit Sub
   End If
   'If FirstCP14 = "" Then
   If (lblAppNo(5).Caption = FirstCP09 And FirstCP14 = "") Or (lblAppNo(5).Caption = ShowCP09 And ShowCP14 = "") Then
      MsgBox strExc(2) & " 未分案", vbCritical
      Exit Sub
   Else
      'Added by Lydia 2016/07/07 判斷委查單是否查覆完畢
      If PUB_TMQCheckOver(lblAppNo(5).Caption) = False Then
         Exit Sub
      End If
      'end 2016/07/07
      'Modified by Lydia 2016/07/07
      'PUB_SendMail strUserNum, FirstCP14, "", strExc(2) & "案，經智權人員確認，請送件!", vbCrLf & "如主旨"
      'Modifie dby Lydia 2021/11/19 + Or ShowCP09 = ""
      PUB_SendMail strUserNum, IIf(FirstCP09 = ShowCP09 Or ShowCP09 = "", FirstCP14, ShowCP14), "", strExc(2) & "案，經智權人員確認，請送件!", vbCrLf & "如主旨"
   End If
   
    '同一收文號，只通知一次
    'Memo by Lydia 2016/07/07 若有追加查名結果，可再通知
    cnnConnection.BeginTrans
       'Modified by Lydia 2016/07/06 改TQC07
       'strSql = "UPDATE TRADEMARKQUERY SET TMQ20=" & strSrvDate(1) & " WHERE TMQ21 = " & CNULL(FirstCP09)
       strSql = "UPDATE TMQCASEMAP SET TQC07=" & strSrvDate(1) & " WHERE TQC02=" & CNULL(lblAppNo(5).Caption) & " AND TQC07 IS NULL "
       cnnConnection.Execute strSql, intI
    cnnConnection.CommitTrans
   txtField(19).Text = strSrvDate(2)
   cmdSendMail.Enabled = False
   
   Exit Sub
ErrHand01:
   
   MsgBox Err.Description, vbCritical
   cnnConnection.RollbackTrans
End Sub

Private Sub cmdTo_Click()
   bolSave = IsSaveData '判斷是否已存檔
   
   If cmdSend.Enabled = False Then
      'Modified by Lydia 2016/04/06 控制是否重複收文
      If TMQ_ReApp = False And lblAppNo(5).Caption <> "" Then
         MsgBox "委查單已收文,不可重複收文!!", vbCritical
         Exit Sub
      End If
      'Modified by Lydia 2016/03/28 控制是否已讀
      If TMQ_CtrRead Then
        strExc(0) = "SELECT TQA01 申請編號,COUNT(TMQ01) 分單量,COUNT(TMQ11) 已查單,COUNT(TMQ19) 已讀單 FROM TMQAPP,trademarkquery WHERE TQA01='" & mTQD01 & "' AND TQA01=TMQ18(+) GROUP BY TQA01"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            If RsTemp.Fields("分單量") - RsTemp.Fields("已查單") <> 0 Then
               MsgBox "申請編號所屬的委查單尚有未查覆,請洽查名人員資詢!", vbInformation, "收文"
               Exit Sub
            End If
            If RsTemp.Fields("分單量") - RsTemp.Fields("已讀單") <> 0 Then
               MsgBox "申請編號所屬的委查單尚有未讀查覆結果,請檢查!", vbInformation, "收文"
               Exit Sub
            End If
        End If
      End If
      'Added by Lydia 2016/03/28
      m_TMQApp = ""
      PubShowNextData
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   Call PUB_GetTMQans("1", True) 'Added by Lydia 2016/06/02 求近似本所案
   'Modified by Lydia 2016/06/27
   If R_type = "M" And strSrvDate(1) >= TMQFileFTP Then
        strAgree = Pub_GetSpecMan("內商查名核可人員")
        'Added by Lydia 2021/11/12 林嘉雯請假時職代處理
        'Memo by Lydia 2021/11/15 已通過電話與嘉雯確認，同時擁有覆核和核可權限，但是覆核和核可作業還是會分開確認(確認作業會發email)；若嘉雯請假則職代有相同權限。
        strExc(1) = GetDutyList(strAgree)
        If strExc(1) <> "" Then strAgree = strAgree & ";" & strExc(1)
        'end 2021/11/12
   End If
   
   'Added by Lydia 2023/07/06
   FR11.BackColor = &H8000000F
   FR12.BackColor = &H8000000F
   FR21.BackColor = &H8000000F
   FR22.BackColor = &H8000000F
   'end 2023/07/06
   
   strPreAgree = Pub_GetSpecMan("內商查名覆核人員") 'Added by Lydia 2022/05/25
   
   'Remove by Lydia 2016/05/26
    '讀取前次設定路徑
   ' strLoadPath = GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "")
   ' If strLoadPath = "" Then
   '    strLoadPath = PUB_Getdesktop
   ' End If
    
   mPreStabs = 0
   SSTab1.Tab = 0
   '因為表單上方空間不足,其他欄位放在第三頁
   If R_type <> "A" Then
      SSTab1.TabVisible(2) = False
      cmdAFUpd(0).Visible = False
      cmdAFUpd(1).Visible = False
   End If
   
   FormReset
   
   'Added by Lydia 2019/12/25 開放特殊設定權限
    If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
       bolSpecMan = True
       strSpecCode = Pub_GetSpecMan("總經理員工編號")
   '開放專利處部份智權同仁資料給彥葶代為處理
   ElseIf CheckLevel(strUserNum, "A8") = True Then
        bolSpecMan = True
        strSpecCode = Pub_GetSpecMan("A7")
   End If
   'end 2019/12/25
   
   If TypeName(Tmpfrm090129) = "Nothing" Then
      cmdKey(0).Visible = False:        cmdKey(1).Visible = False
   End If
   'Added by Lydia 2024/07/17
   cmdRoute.Tag = ""
   Set nFrm = Forms(0).GetForm("frm090131")
   If Not nFrm Is Nothing Then
      If R_type = "U" Or R_type = "A" Then
         cmdRoute.Tag = "M"
      Else
         cmdRoute.Tag = "Q"
      End If
   End If
   'end 2024/07/17
   
End Sub

'Added by Lydia 2019/10/23
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Added by Lydia 2019/10/23 查覆完成後，再進入查名單明細(修改)，刪除附件按結束沒有檢查附件是否存在。
   'ex.T-221140和T-221141的查名結果附件被刪: 非發證和核駁閉卷,所以非批次刪除(TMQ20未上註記); 推測可能查覆完成後，查名人員刪除附件。
   'Modified by Lydia 2024/12/20 排除歷史資料,已刪除附件 And InStr(mTMQ20 & ",", "刪除") = 0
   If Val(txtField(5)) > 0 And R_type <> "Q" And cmdSend.Visible = True And cmdSend.Enabled = True And InStr(mTMQ20 & ",", "刪除") = 0 Then
        'Modified by Lydia 2023/09/08 檢核查名單的「文字一」及「文字二」，若查覆結果各非「無」或「不查」時，必需各別有放置結果附件才可以上查覆。
        'strSql = "SELECT TQD02,SUM(DECODE(TQD06,'" & TMQ_無 & "',0,'" & TMQ_不查 & "',0,1)) TQDCNT,TQFCNT " & _
                     "FROM TMQDETAIL,(SELECT TQF02,SUM(DECODE(UPPER(TQF05),'" & UCase(TMQ_查名作業 & ".pdf") & "',1,0)) AS TQFCNT FROM TMQFILE WHERE TQF02='" & mTQD02 & "' GROUP BY TQF02) " & _
                     "WHERE TQD02='" & mTQD02 & "' AND TQD02=TQF02(+) GROUP BY TQD02,TQFCNT "
        strSql = "SELECT * FROM ( " & _
                 "SELECT TQD02,TQD03,SUM(DECODE(TQD06,'" & TMQ_無 & "',0,'" & TMQ_不查 & "',0,1)) TQDCNT,NVL(TQFCNT,0) TQFCNT FROM TMQDETAIL," & _
                 "(SELECT TQF02,TQF03,SUM(DECODE(UPPER(TQF05),'" & UCase(TMQ_查名作業 & ".pdf") & "',1,0)) AS TQFCNT FROM TMQFILE WHERE TQF02='" & mTQD02 & "' GROUP BY TQF02,TQF03) " & _
                 "WHERE TQD02='" & mTQD02 & "' AND TQD02=TQF02(+) AND TQD03=TQF03(+) GROUP BY TQD02,TQD03,TQFCNT " & _
                  ") WHERE TQDCNT > 0 AND TQFCNT = 0 "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            If Val("" & RsTemp.Fields("tqdcnt")) > 0 And Val("" & RsTemp.Fields("tqfcnt")) = 0 Then
                'Modified by Lydia 2023/09/08
                'MsgBox "尚未新增查覆附件,請確認資料的正確性!", vbCritical
                MsgBox IIf("" & RsTemp.Fields("tqd03") = "0", "圖形", IIf("" & RsTemp.Fields("tqd03") = "1", "文字1", "文字2")) & "尚未新增查覆附件,請確認資料的正確性!", vbCritical
                Cancel = True
            End If
        End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rsB As New ADODB.Recordset
Dim strB As String
   
   '查覆完畢後,再次修改需寄信給相關人員(委查人,近似本所案的智權人員)
   If bolModify = True And strMod(1) <> "" Then
       'Added by Lydia 2016/04/06 已收文加寄給承辦人(所有申請案的承辦人
       If FirstCP09 <> "" Or ShowCP09 <> "" Then
          'Modified by Lydia 2016/07/06 改TQC07
          'strB = "select distinct(nvl(cp14,'')) from caseprogress where cp09 in (select tqc02 from tmqcasemap where tqc03='" & mTQD02 & "' and not(tqc02 is null) and tqc07 is null) and cp57 is null"
          strB = "select distinct(nvl(cp14,'')) from caseprogress where cp09 in (select tqc02 from tmqcasemap where tqc03='" & mTQD02 & "' and not(tqc02 is null)) and cp57 is null"
          intI = 1
          Set rsB = ClsLawReadRstMsg(intI, strB)
          If intI = 1 Then
             rsB.MoveFirst
             Do While Not rsB.EOF
                If rsB(0) <> "" And InStr(strMod(2), "" & rsB(0)) = 0 Then
                   strMod(2) = strMod(2) & ";" & rsB(0)
                End If
                rsB.MoveNext
             Loop
             
          End If
       End If
       strExc(6) = vbCrLf & vbCrLf & "變更內容如下列：" & vbCrLf & vbCrLf & strMod(1)
       PUB_SendMail strUserNum, strMod(2), "", strMod(0), strExc(6)
   End If
   
   DestroyToolTip 'Added by Lydia 2021/10/01 清除物件
   
   Set frm090128 = Nothing
   If TypeName(m_PrevForm) <> "Nothing" Then
        m_PrevForm.Show
        'Modified by Lydia 2025/04/30 +Or m_PrevForm.Name = "frm090127_1"
        If m_PrevForm.Name = "frm090127" Or m_PrevForm.Name = "frm090127_1" Then
           Call m_PrevForm.QueryData
        End If
   End If

   m_NoList = ""
   mbolCall = False
   ShowCP09 = ""
   
   Set nFrm = Nothing 'Added by Lydia 2024/11/08
   
   Set m_PrevForm = Nothing
End Sub
'查覆完畢,並檢查所屬申請編號的查名單是否全部查覆完
Private Sub cmdSend_Click()
Dim tmpC As Integer, tmpE As Integer
Dim rs1 As New ADODB.Recordset
Dim tmpBol As Boolean
Dim bCount As Boolean, bUpdCase As Boolean
Dim mType As String '發信型態
'寫入二進位檔
Dim rsWrite As New ADODB.Recordset
Dim file_num1 As Integer, file_num2 As Integer
Dim btHead(1) As Byte, btHead2(1) As Byte
Dim btTemp() As Byte, btTemp2() As Byte
Dim lLength As Long
Dim strFileName As String
Dim Sub1 As String
Dim strUpd As String
'Added by Lydia 2019/05/21
Dim rsB1 As New ADODB.Recordset
Dim bolEmail As Boolean
Dim inB As Integer

   '檢查明細是否已存檔
   If IsSaveData = False Then
      'Modified by Lydia 2016/05/03 取消最後一筆結果修改
      If mbolSend = True And InStr(strMod(1), "不查") = 0 Then
         GoTo JumpNextStp
      Else
      'end 2016/05/03
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   Select Case cmdSend.Tag
   Case "U" '查覆完畢,修改完畢(未收文前查名人員可修改)
        Call txtField_Validate(11, tmpBol)
        If mTQD03 <> TMQ_AkindPic And txtField(11).Text <> "" Then
            MsgBox "文字查詢不用輸入查名路徑!!", vbCritical
            txtField(11).SetFocus
            Exit Sub
        End If
        
       '合併MAIL
       'Added by Lydia 2018/12/10 T案收文管控齊備:全部匯入的查名單完成查覆才發通知信
       'Remove by Lydia 2019/05/20 在查覆完畢時，針對所有案件進行通知；所以先判斷查名單同批。
       'If strSrvDate(1) >= T案收文齊備啟用日 And lblAppNo(5).Caption <> "" Then
       '   strExc(0) = " SELECT TQC02 申請編號,TMQ01 委查單號," & CNULL(txtField(21)) & " 查覆完成日期,TMQ11 查覆日期,COUNT(TQD04) 明細筆數,COUNT(TQD06) 已查覆筆數," & _
                           " TC2 附件筆數,SUM(DECODE(TQD03||TQD06,'09',1,0)) 不查0,SUM(DECODE(TQD03||TQD06,'19',1,0)) 不查1,SUM(DECODE(TQD03||TQD06,'29',1,0)) 不查2," & _
                            " SC2 近似,SC3 近似Q,SC4 近似2,SC5 近似Q2,VC2,VC3 ," & _
                            " SUM(DECODE(TQD03||TQD06,'07',1,0)) 無0,SUM(DECODE(TQD03||TQD06,'17',1,0)) 無1,SUM(DECODE(TQD03||TQD06,'27',1,0)) 無2" & _
                            " FROM TRADEMARKQUERY,TMQDETAIL,TMQCASEMAP," & _
                            " (SELECT TQC02 VC1,COUNT(TMQ01) VC2,COUNT(TMQ11) VC3 FROM TRADEMARKQUERY,TMQCASEMAP WHERE TQC03=TMQ01(+) GROUP BY TQC02) VT ," & _
                            " (SELECT TQC02 SC1,SUM(DECODE(TQD06,'2',1,0)) SC2,SUM(DECODE(TQD06,'2',DECODE(TQD08,NULL,1,0),0)) SC3,SUM(DECODE(TQD06,'3',1,0)) SC4,SUM(DECODE(TQD06,'3',DECODE(TQD08,NULL,1,0),0)) SC5  FROM TMQCASEMAP,TMQDETAIL WHERE TQC03=TQD02(+) GROUP BY TQC02) VT2 ," & _
                            " (SELECT TMQ01 TC1,COUNT(*) TC2,COUNT(TQF11) TC3 FROM TRADEMARKQUERY,TMQFILE WHERE TMQ01=TQF02(+) AND TQF04<>'00' GROUP BY TMQ01) VT3" & _
                            " WHERE TMQ01='" & mTQD02 & "' AND TQD02=TQC03(+) AND TQC02=" & CNULL(lblAppNo(5).Caption) & _
                            " AND TMQ18=TQD01(+) AND TQD02=TMQ01 AND TQC02=VC1(+) AND TQC02=SC1(+) AND TQD02=TC1(+)" & _
                            " GROUP BY TQC02,TMQ01,TMQ11,VC2,VC3,SC2,SC3,TC2,SC4,SC5"
       'Else '原程式: 如果未收文,依照舊方法
        'Modified by Lydia 2016/04/25 +抓結果為無的筆數
        'Modified by Lydia 2023/09/08 檢核查名單的「文字一」及「文字二」，若查覆結果各非「無」或「不查」時，必需各別有放置結果附件才可以上查覆。
        '  strExc(0) = "SELECT TQA01 申請編號,TMQ01 委查單號,TQA09 查覆完成日期,TMQ11 查覆日期,COUNT(TQD04) 明細筆數,COUNT(TQD06) 已查覆筆數,TC2 附件筆數," & _
                     "SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindPic & TMQ_不查 & "',1,0)) 不查0,SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindWord1 & TMQ_不查 & "',1,0)) 不查1,SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindWord2 & TMQ_不查 & "',1,0)) 不查2" & _
                     ",SC2 近似,SC3 近似Q,SC4 近似2,SC5 近似Q2,VC2,VC3 " & _
                     ",SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindPic & TMQ_無 & "',1,0)) 無0,SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindWord1 & TMQ_無 & "',1,0)) 無1,SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindWord2 & TMQ_無 & "',1,0)) 無2 " & _
                     "FROM TMQAPP,trademarkquery,TMQDETAIL,(SELECT TQA01 VC1,COUNT(TMQ01) VC2,COUNT(TMQ11) VC3 FROM TMQAPP,trademarkquery WHERE TQA01=TMQ18(+) GROUP BY TQA01) VT " & _
                     ",(SELECT TQA01 SC1,SUM(DECODE(TQD06,'" & TMQ_近似1 & "',1,0)) SC2,SUM(DECODE(TQD06,'" & TMQ_近似1 & "',DECODE(TQD08,NULL,1,0),0)) SC3,SUM(DECODE(TQD06,'" & TMQ_近似2 & "',1,0)) SC4,SUM(DECODE(TQD06,'" & TMQ_近似2 & "',DECODE(TQD08,NULL,1,0),0)) SC5 FROM TMQAPP,TMQDETAIL WHERE TQA01=TQD01(+) GROUP BY TQA01) VT2 " & _
                     ",(SELECT TMQ01 TC1,COUNT(*) TC2,COUNT(TQF11) TC3 FROM trademarkquery,TMQFILE WHERE TMQ01=TQF02(+) AND TQF04<>'" & TMQ_附件F04 & "' GROUP BY TMQ01) VT3 " & _
                     "WHERE TQA01='" & mTQD01 & "' AND TQD02='" & mTQD02 & "' AND TQA01=TMQ18(+) AND TQA01=TQD01(+) AND TQD02=TMQ01 AND TQA01=VC1(+) AND TQA01=SC1(+) " & _
                     "AND TQD02=TC1(+) GROUP BY TQA01,TMQ01,TQA09,TMQ11,VC2,VC3,SC2,SC3,TC2,SC4,SC5 "
       'End If 'end 2018/12/10
        strExc(0) = "SELECT TQA01 申請編號,TMQ01 委查單號,TQA09 查覆完成日期,TMQ11 查覆日期,COUNT(TQD04) 明細筆數,COUNT(TQD06) 已查覆筆數," & _
                     "SUM(DECODE(TQD03,'0',1,0)) 明細0,SUM(DECODE(TQD03,'1',1,0)) 明細1,SUM(DECODE(TQD03,'2',1,0)) 明細2,NVL(V3TC2,0) 附件0,NVL(V4TC2,0) 附件1,NVL(V5TC2,0) 附件2," & _
                     "SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindPic & TMQ_不查 & "',1,0)) 不查0,SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindWord1 & TMQ_不查 & "',1,0)) 不查1,SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindWord2 & TMQ_不查 & "',1,0)) 不查2," & _
                     "SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindPic & TMQ_無 & "',1,0)) 無0,SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindWord1 & TMQ_無 & "',1,0)) 無1,SUM(DECODE(TQD03||TQD06,'" & TMQ_AkindWord2 & TMQ_無 & "',1,0)) 無2," & _
                     "SC2 近似,SC3 近似Q,SC4 近似2,SC5 近似Q2,VC2,VC3 " & _
                     "FROM TMQAPP,TRADEMARKQUERY,TMQDETAIL,(SELECT TQA01 VC1,COUNT(TMQ01) VC2,COUNT(TMQ11) VC3 FROM TMQAPP,TRADEMARKQUERY WHERE TQA01=TMQ18(+) GROUP BY TQA01) VT ," & _
                     "(SELECT TQA01 SC1,SUM(DECODE(TQD06,'" & TMQ_近似1 & "',1,0)) SC2,SUM(DECODE(TQD06,'" & TMQ_近似1 & "',DECODE(TQD08,NULL,1,0),0)) SC3,SUM(DECODE(TQD06,'" & TMQ_近似2 & "',1,0)) SC4,SUM(DECODE(TQD06,'" & TMQ_近似2 & "',DECODE(TQD08,NULL,1,0),0)) SC5 FROM TMQAPP,TMQDETAIL WHERE TQA01=TQD01(+) GROUP BY TQA01) VT2, " & _
                     "(SELECT TMQ01 V3TC1,COUNT(*) V3TC2,COUNT(TQF11) V3TC3 FROM TRADEMARKQUERY,TMQFILE WHERE TMQ01=TQF02(+) AND TQF04<>'00' AND TQF03='0' GROUP BY TMQ01) VT3, " & _
                     "(SELECT TMQ01 V4TC1,COUNT(*) V4TC2,COUNT(TQF11) V4TC3 FROM TRADEMARKQUERY,TMQFILE WHERE TMQ01=TQF02(+) AND TQF04<>'00' AND TQF03='1' GROUP BY TMQ01) VT4, " & _
                     "(SELECT TMQ01 V5TC1,COUNT(*) V5TC2,COUNT(TQF11) V5TC3 FROM TRADEMARKQUERY,TMQFILE WHERE TMQ01=TQF02(+) AND TQF04<>'00' AND TQF03='2' GROUP BY TMQ01) VT5 " & _
                     "WHERE TQA01='" & mTQD01 & "' AND TQD02='" & mTQD02 & "' AND TQA01=TMQ18(+) AND TQA01=TQD01(+) AND TQD02=TMQ01 AND TQA01=VC1(+) AND TQA01=SC1(+) AND TQD02=V3TC1(+) AND TQD02=V4TC1(+) AND TQD02=V5TC1(+) " & _
                     "GROUP BY TQA01,TMQ01,TQA09,TMQ11,VC2,VC3,SC2,SC3,V3TC2,V4TC2,V5TC2,SC4,SC5 "
        intI = 1
        Set rs1 = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           If mTQD03 = TMQ_AkindPic And rs1.Fields("不查0") = 0 And txtField(11).Text = "" Then
              MsgBox "圖形查詢請輸入查名路徑!!", vbCritical
              txtField(11).SetFocus
              Exit Sub
           End If
           '暫時不鎖編輯
           If txtField(9).Tag <> txtField(9).Text Or txtField(7).Tag <> txtField(7).Text Or txtField(8).Tag <> txtField(8).Text Then
              MsgBox "查詢筆數不可變更!!", vbCritical
              txtField(9).Text = txtField(9).Tag:    txtField(7).Text = txtField(7).Tag:    txtField(8).Text = txtField(8).Tag
              Exit Sub
           End If
           If rs1.Fields("近似Q") > 0 Or rs1.Fields("近似Q2") > 0 Then
              MsgBox "有" & TMQ_近似T1 & " / " & TMQ_近似T2 & "未輸入申請號或審定號,請確認資料的正確性!", vbCritical
              Exit Sub
           End If
           'Modified by Lydia 2016/04/25 結果為不查或無，不需新增查覆附件
           ''Modified by Lydia 2023/09/08 檢核查名單的「文字一」及「文字二」，若查覆結果各非「無」或「不查」時，必需各別有放置結果附件才可以上查覆。
           'If Val("" & rs1.Fields("附件筆數")) = 0 And ((mTQD03 = TMQ_AkindPic And rs1.Fields("明細筆數") <> rs1.Fields("不查0") + rs1.Fields("無0")) _
              Or (mTQD03 <> TMQ_AkindPic And (rs1.Fields("明細筆數") <> rs1.Fields("不查1") + rs1.Fields("無1") + rs1.Fields("不查2") + rs1.Fields("無2")))) Then
              'MsgBox "尚未新增查覆附件,請確認資料的正確性!", vbCritical
           strExc(1) = ""
           If Val("" & rs1.Fields("明細0")) - Val("" & rs1.Fields("不查0")) - Val("" & rs1.Fields("無0")) > 0 And Val("" & rs1.Fields("附件0")) = 0 Then
              strExc(1) = "圖形"
           ElseIf Val("" & rs1.Fields("明細1")) - Val("" & rs1.Fields("不查1")) - Val("" & rs1.Fields("無1")) > 0 And Val("" & rs1.Fields("附件1")) = 0 Then
              strExc(1) = "文字1"
           ElseIf Val("" & rs1.Fields("明細2")) - Val("" & rs1.Fields("不查2")) - Val("" & rs1.Fields("無2")) > 0 And Val("" & rs1.Fields("附件2")) = 0 Then
              strExc(1) = "文字2"
           End If
           If strExc(1) <> "" Then
              MsgBox strExc(1) & "尚未新增查覆附件,請確認資料的正確性!", vbCritical
           'end 2023/09/08
              Exit Sub
           End If
           If Not IsNull(rs1.Fields("查覆完成日期")) Or Not IsNull(rs1.Fields("查覆日期")) Then
              bolModify = True
           End If

            '若有組群不查則重新計算筆數
            If rs1.Fields("不查0") > 0 Or (mTQD03 = TMQ_AkindPic And bolModify = True And bolModCheck = False) Then
                 txtField(9).Text = rs1.Fields("明細筆數") - rs1.Fields("不查0")
                 '修改已查覆的結果時,就重新計算筆數,不再詢問
                 If mTQD03 = TMQ_AkindPic And bolModify = True Then
                     bCount = True
                 Else
                     If MsgBox("原本圖形筆數:" & txtField(9).Tag & " ,現在:" & txtField(9).Text & " ", vbInformation + vbYesNo, "委查筆數是否變更") = vbNo Then
                        txtField(9).Text = txtField(9).Tag:    txtField(7).Text = txtField(7).Tag:    txtField(8).Text = txtField(8).Tag
                        Exit Sub
                     Else
                        bCount = True
                     End If
                 End If
            ElseIf rs1.Fields("不查1") > 0 Or rs1.Fields("不查2") > 0 Or (mTQD03 = "1" And bolModify = True And bolModCheck = False) Then
               txtField(7).Text = "0": txtField(8).Text = "0"
               For jj = 1 To 2
                  tmpE = 0: tmpC = 0
                  If cmdKey(jj - 1).Visible = True Then
                      If txtUnicode(jj).Text = "" Then tmpC = tmpC + 1
                  End If
                  '文字可能以替代字輸入,所以筆數依文字欄位來判斷
                  Call PUB_CountTxtNEC(tmpE, tmpC, txtUnicode(jj).Text)
                  If jj = 1 Then
                     txtField(7).Text = Format(Val(txtField(7).Text) + tmpC * (GRD1.Rows - 1 - rs1.Fields("不查1")), "0")
                     txtField(8).Text = Format(Val(txtField(8).Text) + tmpE * (GRD1.Rows - 1 - rs1.Fields("不查1")), "0")
                  Else
                     txtField(7).Text = Format(Val(txtField(7).Text) + tmpC * (grd2.Rows - 1 - rs1.Fields("不查2")), "0")
                     txtField(8).Text = Format(Val(txtField(8).Text) + tmpE * (grd2.Rows - 1 - rs1.Fields("不查2")), "0")
                  End If
               Next jj
               If mTQD03 = "1" And bolModify = True Then
                     bCount = True
               Else
                  If MsgBox("原本中文筆數:" & txtField(7).Tag & " ,現在:" & txtField(7).Text & vbCrLf & "原本英文筆數:" & txtField(8).Tag & " ,現在:" & txtField(8).Text & " ", vbInformation + vbYesNo, "委查筆數是否變更") = vbNo Then
                     txtField(9).Text = txtField(9).Tag:    txtField(7).Text = txtField(7).Tag:    txtField(8).Text = txtField(8).Tag
                     Exit Sub
                  Else
                     bCount = True
                  End If
               End If
            End If
              
              If rs1.Fields("明細筆數") = rs1.Fields("已查覆筆數") Then
                 cnnConnection.BeginTrans
                    '只上查覆日期
                    strSql = ""
                    If bCount = True Then
                         If rs1.Fields("不查0") > 0 Then
                               strSql = strSql & ",tmq09=" & Val(txtField(9).Text)
                         ElseIf rs1.Fields("不查1") > 0 Or rs1.Fields("不查2") > 0 Then
                               strSql = strSql & ",tmq07=" & Val(txtField(7).Text) & ",tmq08=" & Val(txtField(8).Text)
                         End If
                         '變更筆數=UPDATE
                         If strSql <> "" Then strSql = strSql & ",tmq15='" & strUserNum & "',tmq16=" & strSrvDate(1) & ",tmq17=" & Left(Format(ServerTime, "000000"), 4)
                    End If
                    'Modified by Lydia 2016/05/03 判斷查覆完畢後的再次修改
                    'strExc(1) = "update trademarkquery set tmq11=" & strSrvDate(1) & strSql & " ,tmq24='" & Trim(txtField(11)) & "' where tmq01='" & rs1.Fields("委查單號") & "' "
                    strExc(1) = "update trademarkquery set tmq24=" & CNULL(Trim(txtField(11))) & IIf(mbolSend = False, ", tmq11=" & strSrvDate(1), "") & strSql & " where tmq01='" & rs1.Fields("委查單號") & "' "
                    cnnConnection.Execute strExc(1), intI

                    'end 2016/05/30
                    '+修改記錄
                    If bCount = True Then Pub_SeekTbLog strExc(1)
                    
                    'Memo by Lydia 2019/05/21 T-221600和T-221601先填接洽單後勾查名單
                    '1.T-221601在5/15補勾HA8050536~544
                    '2.T-221600在5/15補勾HA8050536~544；之後T-221601取消勾選HA8050538~544(bug1: 直接清空TMQ21)
                    '3.T-221601又在5/20勾HA8050538和HA8050541
                    '4. HA8050538和HA8050541在查覆完畢時，沒有針對所有案件進行通知。
                    
                    'Added by Lydia 2019/05/21 在查覆完畢時，針對所有案件進行通知。
                    'Modified by Lydia 2025/06/26 判斷同一收文不同申請編號的近似案
                    'strSql = "select cp01,cp02,cp03,cp04,tqc02,cp06,cp13,cp14,cp122,ep06,count(tqc03) as 已勾 ,sum(decode(tmq11,null,0,1)) as 已完成 " & _
                                 "from tmqcasemap,trademarkquery,caseprogress,engineerprogress where tqc02 in (select tqc02 from tmqcasemap where nvl(tqc02,'N') <>'N' and tqc03='" & mTQD02 & "' ) " & _
                                 "and tqc03=tmq01(+) and tqc02=cp09(+) and cp09=ep02(+) group by cp01,cp02,cp03,cp04,tqc02,cp06,cp13,cp14,cp122,ep06 "
                    strSql = "select cp01,cp02,cp03,cp04,tqc02,cp06,cp13,cp14,cp122,ep06,count(tqc03) as 已勾 ,sum(decode(tmq11,null,0,1)) as 已完成,SUM(DECODE(TQD06,'3',1,0)) as 近似,SUM(DECODE(TQD06,'2',1,0)) as 近似2 " & _
                                 "from tmqcasemap,trademarkquery,caseprogress,engineerprogress,tmqdetail where tqc02 in (select tqc02 from tmqcasemap where nvl(tqc02,'N') <>'N' and tqc03='" & mTQD02 & "' ) " & _
                                 "and tqc03=tmq01(+) and tqc02=cp09(+) and cp09=ep02(+) AND tmq01=tqd02(+) group by cp01,cp02,cp03,cp04,tqc02,cp06,cp13,cp14,cp122,ep06 "
                    inB = 1
                    Set rsB1 = ClsLawReadRstMsg(inB, strSql)
                    'Move by Lydia 2019/05/21 從”T案收文齊備啟用日”下面移上來
                    '查名單申請上完成日期
                    strExc(1) = "update tmqapp set tqa09=" & strSrvDate(1) & " where tqa09 is null and tqa01=(" & _
                                      " select tmq18 from (select tmq18,count(distinct(tmq01)) 查名單量,sum(decode(tmq11,null,0,1)) 已查覆 from trademarkquery where tmq18='" & mTQD01 & "' group by tmq18) where 查名單量=已查覆 ) "
                    cnnConnection.Execute strExc(1), intI
                        
                    'VC2=申請編號有幾張查名單 ,VC3=已查覆的單量
                    'Added by Lydia 2018/12/10 T案收文管控齊備:全部匯入的查名單完成查覆才發通知信
                    'Modified by Lydia 2019/05/21 在查覆完畢時，針對所有案件進行通知。
                    'If strSrvDate(1) >= T案收文齊備啟用日 And lblAppNo(5).Caption <> "" Then
                    '    If rs1.Fields("VC2") = rs1.Fields("VC3") + 1 Then
                    '       strExc(1) = "update caseprogress set cp143=" & strSrvDate(1) & " where cp09='" & rs1.Fields("申請編號") & "' " 'SQL收文號=申請編號
                    If strSrvDate(1) >= T案收文齊備啟用日 And inB = 1 Then
                        rsB1.MoveFirst
                        mType = ""
                        Do While Not rsB1.EOF
                            If Val("" & rsB1.Fields("已勾")) = Val("" & rsB1.Fields("已完成")) Then
                              '查名已齊備
                               strExc(1) = "update caseprogress set cp143=" & strSrvDate(1) & " where cp09='" & rsB1.Fields("tqc02") & "' " '所有案件的收文號
                    'end 2019/05/21
                               cnnConnection.Execute strExc(1), intI
                               '發信給申請者
                               mType = mType & "查覆" & ";"
                               'Added by Lydia 2019/01/30 (有承辦人)文件和查名同時齊備=>更新承辦期限
                               'Modified by Lydia 2019/05/21 依照各案件的現況
                               'If (FirstCP14 <> "" Or ShowCP14 <> "") And Val(m_EP06) > 0 Then
                               '     strExc(0) = PUB_TMdebateCountCP48(m_CP06, m_CP122, m_EP06, IIf(ShowCP09 <> "", ShowCP09, FirstCP09), m_CP13)
                               '     If strExc(0) <> "" Then
                               '         strExc(1) = "UPDATE CaseProgress SET CP48 = " & strExc(0) & " " & _
                               '                  "WHERE CP09 = '" & IIf(ShowCP09 <> "", ShowCP09, FirstCP09) & "' "
                               '         cnnConnection.Execute strExc(1), intI
                               '     End If
                               'End If
                               'end 2019/01/30
                               If "" & rsB1.Fields("cp14") <> "" And Val("" & rsB1.Fields("ep06")) > 0 Then
                                    strExc(0) = PUB_TMdebateCountCP48("" & rsB1.Fields("cp06"), "" & rsB1.Fields("cp122"), "" & rsB1.Fields("ep06"), "" & rsB1.Fields("tqc02"), "" & rsB1.Fields("cp13"))
                                    If strExc(0) <> "" Then
                                        strExc(1) = "UPDATE CaseProgress SET CP48 = " & strExc(0) & " " & _
                                                 "WHERE CP09 = '" & "" & rsB1.Fields("tqc02") & "' "
                                        cnnConnection.Execute strExc(1), intI
                                    End If
                               End If
                               If mType <> "" Then
                                   'Modified by Lydia 2025/06/26 判斷同一收文不同申請編號的近似案
                                   'If "" & rs1.Fields("近似") > 0 Then mType = mType & "近似本所案" & ";"
                                   'If "" & rs1.Fields("近似2") > 0 Then mType = mType & "相同本所案" & ";"
                                   If Val("" & rs1.Fields("近似")) + Val("" & rsB1.Fields("近似")) > 0 Then mType = mType & "近似本所案" & ";"
                                   If Val("" & rs1.Fields("近似2")) + Val("" & rsB1.Fields("近似2")) > 0 Then mType = mType & "相同本所案" & ";"
                                   'end 2025/06/26
                                   CloseMail IIf("" & rsB1.Fields("cp03") & rsB1.Fields("cp04") <> "000", rsB1.Fields("cp01") & "-" & rsB1.Fields("cp02") & "-" & rsB1.Fields("cp03") & "-" & rsB1.Fields("cp04"), rsB1.Fields("cp01") & "-" & rsB1.Fields("cp02")), mType, rsB1.Fields("已勾"), Trim(txtField(0).Text) & ";" & rsB1.Fields("cp14"), True
                                   bolEmail = True
                               End If
                               'end 2019/05/21
                            End If
                        'Added by Lydia 2019/05/21
                            rsB1.MoveNext
                        Loop
                        'end 2019/05/21
                    Else
                        If rs1.Fields("VC2") = rs1.Fields("VC3") + 1 Then
                           strExc(1) = "update tmqapp set tqa09=" & strSrvDate(1) & " where tqa01='" & rs1.Fields("申請編號") & "' "
                           cnnConnection.Execute strExc(1), intI
                           '發信給申請者
                           mType = mType & "查覆" & ";"
                        End If
                    End If
                 cnnConnection.CommitTrans
                 
                 'Added by Lydia 2016/08/25 查名單筆數歸零,再重新修正拿單量
                 If Val(txtField(7).Text) + Val(txtField(8).Text) + Val(txtField(9).Text) = 0 Then
                    '更新當日拿單量
                     Call PUB_TMQtake("2", txtField(1).Text, , , , , "0")
                    '更新前2日統計量
                     Call PUB_TMQtake("2", txtField(1).Text, , , , , "1")
                 End If
                 'end 2016/08/25
                 
                 iStiu = 0
                 FormEnabled
                 'Modified by Lydia 2019/05/21 排除已發過的情況
                 'If mType <> "" And bolModify = False Then
                 If bolEmail = True Then
                     PUB_SendMailCache
                 End If
                 If mType <> "" And bolModify = False And bolEmail = False Then
                 'end 2019/05/21
                    If rs1.Fields("近似") > 0 Then mType = mType & "近似本所案" & ";"
                    If rs1.Fields("近似2") > 0 Then mType = mType & "相同本所案" & ";"
                    'Modified by Lydia 2016/03/28 通知預設覆核人員
                    'Remove by Lydia 2016/03/31 因為目前由智權人員推動案件的進行,所以若要繼續申請則由智權人員通知覆核人員
                    'CloseMail rs1.Fields("申請編號"), mType, rs1.Fields("VC2"), Trim(txtField(0).Text) & IIf(strExc(10) <> "", ";" & strExc(10) & ";84027;69008", "")
                    'Modified by Lydia 2016/04/29 傳承辦人
                    CloseMail rs1.Fields("申請編號"), mType, rs1.Fields("VC2"), Trim(txtField(0).Text) & IIf(FirstCP14 <> "", ";" & FirstCP14, "")
                 End If
JumpNextStp:
                 Call cmdNext_Click
              Else
                 MsgBox "尚有組群未輸入查覆結果,請檢查!", vbCritical
                 Exit Sub
              End If
           
        End If

   Case "M" '覆核完畢
        'Added by Lydia 2019/12/12 先檢查  核可人員 => 設定" 是否出名 "
        If InStr("M,A", R_type) > 0 And (InStr(strAgree, strUserNum) > 0 Or Pub_StrUserSt03 = "M51") Then
            strExc(1) = Pub_ChkTQD11(mTQD02, strExc(2))
            If Len(strExc(2)) > 1 Then
                MsgBox "核可案設定""是否出名""的結果有不一致，請修改！", vbCritical, "核可案－是否出名"
                Exit Sub
            End If
        End If
        
        'Modified by Lydia 2016/06/27 第2次覆核完畢檢查,TQD09=1(核可案)
        'strExc(0) = "SELECT TQA01 申請編號,TMQ01 委查單號,TQA09 查覆完成日期,TMQ22,TMQ23 覆核日期,COUNT(TQD04) 明細筆數,COUNT(TQD09) 已覆核筆數,VC2,VC3 " & _
                    "FROM TMQAPP,trademarkquery,TMQDETAIL,(SELECT TQD01 VC1,COUNT(TMQ22) VC2,COUNT(TMQ01) VC3 FROM TMQDETAIL,trademarkquery WHERE TQD01=TMQ18(+) AND TQD02=TMQ01(+) AND TQD06 in (" & TMQ_近似1 & "," & TMQ_近似2 & ") GROUP BY TQD01) VT " & _
                    "WHERE TQA01='" & mTQD01 & "' AND TQD02='" & mTQD02 & "' AND TQD06 in (" & TMQ_近似1 & "," & TMQ_近似2 & ") AND TQA01=TMQ18(+) AND TQA01=TQD01(+) AND TQD02=TMQ01 AND TQA01=VC1(+) " & _
                    "GROUP BY TQA01,TMQ01,TQA09,TMQ22,TMQ23,VC2,VC3 "
        strExc(0) = "SELECT TQA01 申請編號,TMQ01 委查單號,TQA09 查覆完成日期,TMQ22,TMQ23 覆核日期,COUNT(TQD04) 明細筆數,COUNT(TQD09) 已覆核筆數,VC2,VC3,VC4,VC5,VC6 " & _
                    "FROM TMQAPP,trademarkquery,TMQDETAIL,(SELECT TQD01 VC1,COUNT(TMQ22) VC2,COUNT(TMQ01) VC3,COUNT(TQD06) VC4,SUM(DECODE(TQD09,'1',1,0)) VC5," & _
                    "SUM(DECODE(TQD09,'" & TMQ_近似1 & "',1,'" & TMQ_近似2 & "',1,0)) VC6 FROM TMQDETAIL,trademarkquery " & _
                    "WHERE TQD01=TMQ18(+) AND TQD02=TMQ01(+) AND TQD06 in (" & TMQ_近似1 & "," & TMQ_近似2 & ") GROUP BY TQD01) VT " & _
                    "WHERE TQA01='" & mTQD01 & "' AND TQD02='" & mTQD02 & "' AND TQD06 in (" & TMQ_近似1 & "," & TMQ_近似2 & ") " & _
                    "AND TQA01=TMQ18(+) AND TQA01=TQD01(+) AND TQD02=TMQ01 AND TQA01=VC1(+) " & _
                    "GROUP BY TQA01,TMQ01,TQA09,TMQ22,TMQ23,VC2,VC3,VC4,VC5,VC6 "
       
        intI = 1
        Set rs1 = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           'Modified by Lydia 2016/03/28 記錄最後覆核主管/日期
           'If Not IsNull(rs1.Fields("TMQ22")) Then
           '   MsgBox "委查單號已覆核完畢,請確認資料的正確性!", vbCritical
           '   Exit Sub
           'ElseIf rs1.Fields("明細筆數") <> rs1.Fields("已覆核筆數") Then
           '   MsgBox "有" & TMQ_近似T1 & " / " & TMQ_近似T2 & "尚未覆核,請確認資料的正確性!", vbCritical
           '   Exit Sub
           'Else
           If rs1.Fields("明細筆數") <> rs1.Fields("已覆核筆數") Then
              'Modified by Lydia 2022/06/14 預設按鈕=否 => + vbDefaultButton2
              If MsgBox("有" & TMQ_近似T1 & " / " & TMQ_近似T2 & "尚未覆核,確定是否要覆核完畢?", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                 Exit Sub
              End If
           End If
                'Added by Lydia 2017/05/08 申請編號須要覆核的查名單
                strExc(0) = "select COUNT(DISTINCT(TMQ01)) from trademarkquery,tmqdetail where tmq01=tqd02(+) AND TMQ22 IS NULL and tmq18='" & mTQD01 & "' and tqd08 is not null "
                intI = 1
                strExc(2) = "0"
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   strExc(2) = "" & RsTemp(0)
                End If
                'END 2017/05/08
              cnnConnection.BeginTrans
                    strExc(1) = "update trademarkquery set tmq22='" & strUserNum & "',tmq23=" & strSrvDate(1) & " where tmq01='" & rs1.Fields("委查單號") & "' "
                    cnnConnection.Execute strExc(1), intI
                    'Added by Lydia 2019/12/12 核可人員 => 設定" 是否出名 "=2不出名，更新進度檔
                    strExc(3) = Pub_ChkTQD11(rs1.Fields("委查單號"), strExc(4))
                    If strExc(3) <> "" And Left(strExc(4), 1) = "2" Then
                        strExc(1) = "update caseprogress set cp22='N',cp110=null where cp22 is null and cp09 in (" & GetAddStr(strExc(3)) & ") "
                        cnnConnection.Execute strExc(1), intI
                    End If
                    'end 2019/12/12
              cnnConnection.CommitTrans
                iStiu = 0
                FormEnabled
               
                 strExc(9) = "": strExc(10) = ""
                'VC2=已覆核的查名單, VC3=申請編號有幾張查名單
                'VC4=申請編號的查名單需覆核的筆數,VC5=核可案的筆數,VC6=覆核後仍為近似本所案的筆數
                'Modified by Lydia 2017/05/08 申請編號須要覆核的查名單-現在的查名單=0
                'If rs1.Fields("VC2") + 1 = rs1.Fields("VC3") Then
                'If Val(strExc(2)) - 1 = 0 Then 'Mark by Lydia 2024/10/30 第1次以後也要發和第一次覆核的通知Email
                   '發信給申請者(當全部分單覆核完畢後,只發一次)
                   'Modified by Lydia 2016/04/29 傳承辦人
                   mType = "覆核"
                   'Added by Lydia 2016/06/27
                   'Modified by Lydia 2022/06/06 覆核增加確認機制; 當查名結果從近似本所案△改為非近似本所案是查名人員的看法，尚需要智權人員確認結果，所以修改通知email提醒兩方要再次確認(杜經理的需求by 嘉雯)
                   'If rs1.Fields("VC6") > 0 And strSrvDate(1) >= TMQFileFTP Then mType = mType & ";近似本所案;相同本所案;"
                   If rs1.Fields("VC6") > 0 Then
                      mType = mType & ";近似本所案;相同本所案;"
                   Else
                      mType = mType & ";增加確認;"
                   End If
                   'end 2022/06/06
                   'Added by Lydia 2017/01/20 相同△、近似△之查名結果，覆核完畢後，系統除通知委查人及案件承辦人(已收文)外，增加通知查名人。
                   strSql = "select tmq10 from trademarkquery,tmqdetail where tmq18='" & mTQD01 & "' and tmq01=tqd02(+) and tqd06 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') group by tmq10 "
                   intI = 1: strExc(0) = ""
                   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                   If intI = 1 Then
                      strExc(0) = ";" & RsTemp.GetString(adClipString, , , ";")
                   End If
                   
                   'Added by Lydia 2022/06/06
                   If InStr(mType, "增加確認") > 0 Then
                      '只傳入查名結果有近似本所案的查名單之查名人員
                      CloseMail IIf(FirstCP(1) <> "", FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) <> "000", "-" & FirstCP(3) & "-" & FirstCP(4), ""), rs1.Fields("申請編號")), mType, rs1.Fields("VC3"), strExc(0)
                   Else
                   'end 2022/06/06
                      'Modified by Lydia 2017/01/20 +查名結果有近似本所案的查名單之查名人員 -> strexc(0)
                      CloseMail rs1.Fields("申請編號"), mType, rs1.Fields("VC3"), Trim(txtField(0).Text) & IIf(FirstCP14 <> "", ";" & FirstCP14, "") & strExc(0)
                   End If 'Added by Lydia 2022/06/06
                       
                'Modified by Lydia 2016/06/27 第2次覆核完畢檢查,TQD09=1(核可案)
                'Modified by Lydia 2016/09/12 與嘉雯確認,當第1次覆核完畢後,若再次修改覆核結果,則每張查名單就發mail通知
                'ElseIf strSrvDate(1) >= TMQFileFTP And rs1.Fields("VC2") = rs1.Fields("VC3") And rs1.Fields("VC4") = rs1.Fields("VC5") And (InStr(strAgree, strUserNum) > 0 Or Pub_StrUserSt03 = "M51") Then
                'Mark by Lydia 2024/10/30 第1次以後也要發和第一次覆核的通知Email
                'ElseIf rs1.Fields("VC2") = rs1.Fields("VC3") Then
                '   If Pub_StrUserSt03 = "M51" Then
                '      If MsgBox("電腦中心人員請問你現在是覆核人員?", vbYesNo + vbDefaultButton2) = vbNo Then
                '         strAgree = strAgree & "," & strUserNum
                '      End If
                '   End If
                '   '更改覆核結果後按覆核完畢,就發mail通知
                '   If InStr(strAgree, strUserNum) = 0 And bolChgTM22 = True And txtField(2).Tag <> "" And txtField(17).Tag <> "" Then
                '        mType = "更改"
                '        If rs1.Fields("VC6") > 0 Then mType = mType & ";近似本所案;相同本所案;"
                '        CloseMail rs1.Fields("申請編號"), mType, rs1.Fields("VC3"), Trim(txtField(0).Text) & IIf(FirstCP14 <> "", ";" & FirstCP14, "")
                '
                '   ElseIf rs1.Fields("VC4") = rs1.Fields("VC5") Then '核可案
                ''end 2016/09/12
                '       CloseMail rs1.Fields("申請編號"), "覆核核可案", rs1.Fields("VC3"), Trim(txtField(0).Text) & IIf(FirstCP14 <> "", ";" & FirstCP14, "")
                '   End If
                ''end 2016/06/27
                '
                'End If
                'end ---- Mark by Lydia 2024/10/30 第1次以後也要發和第一次覆核的通知Email
                
                bolChgTM22 = False 'Added by Lydia 2016/09/12
                Call cmdNext_Click
           'End If
           'end 2016/03/28
        End If
   Case "Q" '撤回
        If MsgBox("確定撤回查名作業嗎?", vbInformation + vbYesNo) = vbYes Then
           'Modified by Lydia 2016/04/26
           'strExc(0) = "SELECT COUNT(*) FROM TMQDETAIL WHERE TQD01='" & mTQD01 & "' AND ( TQD06 <> " & CNULL(TMQ_不查) & " OR TQD09 <> " & CNULL(TMQ_不查) & ") "
           strExc(0) = "SELECT NVL(SUM(CNT1),0) FROM (SELECT TQD01,COUNT(*) CNT1 FROM TMQDETAIL WHERE TQD01='" & mTQD01 & "' AND ( TQD06 <> " & CNULL(TMQ_不查) & " OR TQD09 <> " & CNULL(TMQ_不查) & ") GROUP BY TQD01" & _
                       " UNION SELECT TQF01,COUNT(*) CNT1 FROM TMQFILE WHERE TQF01='" & mTQD01 & "' AND TQF04 <> '" & TMQ_附件F04 & "' GROUP BY TQF01) "

           intI = 1
           Set rs1 = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
              If rs1(0) > 0 Then
                 MsgBox "查名作業已經有輸入查覆結果或附件，不可撤回!!", vbCritical
                 Exit Sub
              Else
                  strExc(0) = "SELECT TMQ01,TMQ10 FROM trademarkquery WHERE TMQ18='" & mTQD01 & "' "
                   intI = 1
                   strUpd = ""
                   Set rs1 = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                      
                      '保留資料,結果上不查
                      'Modified by Lydia 2017/11/23 加註
                      'strUpd = strUpd & "UPDATE TMQDETAIL SET TQD06='" & TMQ_不查 & "',TQD07='" & ChangeTStringToTDateString(strSrvDate(2)) & "已由委查人自行撤回" & "' WHERE TQD01='" & mTQD01 & "' ; "
                      strUpd = strUpd & "UPDATE TMQDETAIL SET TQD06='" & TMQ_不查 & "',TQD07='" & ChangeTStringToTDateString(strSrvDate(2)) & _
                                        "已由" & IIf(strUserNum <> txtField(1), strUserName & "代替", "") & "委查人自行撤回" & "' WHERE TQD01='" & mTQD01 & "' ; "
                      '上查覆日期=已完成,清空收文號
                      'Modified by Lydia 2016/06/15 委查人撤回,更新update欄位
                      'strUpd = strUpd & "UPDATE trademarkquery SET TMQ11=" & strSrvDate(1) & ",TMQ21=NULL ,TMQ07=0 ,TMQ08=0 ,TMQ09=0 WHERE TMQ18='" & mTQD01 & "' ; "
                      strUpd = strUpd & "UPDATE trademarkquery SET TMQ11=" & strSrvDate(1) & ",TMQ21=NULL ,TMQ07=0 ,TMQ08=0 ,TMQ09=0,tmq15='" & strUserNum & "',tmq16=" & strSrvDate(1) & ",tmq17=" & Left(Format(ServerTime, "000000"), 4) & " WHERE TMQ18='" & mTQD01 & "' ; "
                      strUpd = strUpd & "UPDATE TMQAPP SET TQA20='Y' WHERE TQA01='" & mTQD01 & "' AND TQA20 IS NULL ; "
                      Sub1 = GetStrTitle
                      rs1.MoveFirst
                      Do While Not rs1.EOF
                         'Added by Lydia 2016/04/26 已收文案件的處理
                         If FirstCP09 <> "" Then
                             'Modified by Lydia 2016/07/06 刪除記錄
                             'strUpd = strUpd & "update tmqcasemap set tqc07='N' where tqc02='" & FirstCP09 & "' and tqc03='" & rs1.Fields("TMQ01") & "' AND TQC07 IS NULL ; "
                             'strUpd = strUpd & "delete from casepaperpdf where cpp01='" & FirstCP09 & "' and instr(cpp02,'" & FirstCPP02t & rs1.Fields("TMQ01") & "." & TMQ_查名作業 & ".menu" & "') > 0 ;"
                             strUpd = strUpd & "delete from tmqcasemap where tqc03='" & rs1.Fields("TMQ01") & "' ; "
                             strUpd = strUpd & "delete from casepaperpdf where instr(cpp02,'" & rs1.Fields("TMQ01") & "." & TMQ_查名作業 & ".menu" & "') > 0 ;"
                         End If
                         'end 2016/04/26
                         'Added by Lydia 2024/11/12 查名單(網中)：平行測試
                         If strSrvDate(1) >= 查名單網中系統平行測試 Then
                            strUpd = strUpd & "Update tmqappform set TMA14=to_char(sysdate,'yyyymmdd'), TMA34=NULL, TMA35=NULL, TMA36=0, TMA37=0, TMA38=0,TMA13='Y',TMA15=TMA15||'" & ChangeTStringToTDateString(strSrvDate(2)) & IIf(strUserNum <> txtField(1), strUserName & "代替", "") & "委查人自行撤回' where TMA71='" & rs1.Fields("tmq01") & "'; "
                         End If
                         'end 2024/11/12
                         rs1.MoveNext
                      Loop
                      'Move by Lydia 2016/08/25 先將查名單筆數歸零,再重新修正拿單量
                      If strUpd <> "" Then
                           tmpArr = Empty
                           tmpArr = Split(strUpd, ";")
                           cnnConnection.BeginTrans
                           For intI = 0 To UBound(tmpArr)
                              If Trim(tmpArr(intI)) <> "" Then
                                 cnnConnection.Execute Trim(tmpArr(intI)), jj
                              End If
                           Next intI
                           cnnConnection.CommitTrans
                      End If
                      
                      rs1.MoveFirst
                      Do While Not rs1.EOF
                          If "" & rs1.Fields("TMQ10") <> "" Then 'Addded by Lydia 2018/05/25 判斷有查名人員
                              '更新當日拿單量
                              Call PUB_TMQtake("2", rs1.Fields("TMQ10"), , , , , "0")
                              '更新前2日統計量
                              Call PUB_TMQtake("2", rs1.Fields("TMQ10"), , , , , "1")
                              '發Mail通知查名人員
                              'Modified by Lydia 2017/11/23 加註
                              'strExc(5) = "「" & txtField(10).Text & "」" & Trim(Sub1) & " 委查單: " & rs1.Fields("TMQ01") & " ,已由委查人自行撤回, 請不用繼續查名!!"
                              'Modified by Lydia 2021/10/01 txtField(10) => textCName
                              strExc(5) = "「" & textCName.Text & "」" & Trim(Sub1) & " 委查單: " & rs1.Fields("TMQ01") & " ,已由" & IIf(strUserNum <> txtField(1), strUserName & "代替", "") & "委查人自行撤回, 請不用繼續查名!!"
                              strExc(6) = vbCrLf & vbCrLf & "若有疑問，可向委查人詢問。"
                        
                              PUB_SendMail strUserNum, rs1.Fields("TMQ10"), "", strExc(5), strExc(6)
                         End If 'end 2018/05/25
                         rs1.MoveNext
                      Loop
                      'end 2016/08/25
                   End If
                 
                 mTQA20 = "Y"
                 MsgBox "查名單已撤回!", vbInformation
                 Call cmdExit_Click
              End If
           End If
        End If
   Case "A" '資料維護(限電腦中心)
        If txtUnicode(1).Text <> "" And txtUnicode(1).Tag = "" And cmdKey(0).Tag = "" Then
           MsgBox "原申請就無文字1,不可輸入!!", vbCritical
           txtUnicode(1).SetFocus
           Exit Sub
        ElseIf txtUnicode(2).Text <> "" And txtUnicode(2).Tag = "" And cmdKey(1).Tag = "" Then
           MsgBox "原申請就無文字2,不可輸入!!", vbCritical
           txtUnicode(2).SetFocus
           Exit Sub
        ElseIf txtUnicode(1).Text = "" And txtUnicode(1).Tag <> "" And cmdKey(0).Tag = "" Then
           MsgBox "文字1無查詢內容!!", vbCritical
           txtUnicode(1).SetFocus
           Exit Sub
        ElseIf txtUnicode(2).Text = "" And txtUnicode(2).Tag <> "" And cmdKey(1).Tag = "" Then
           MsgBox "文字2無查詢內容!!", vbCritical
           txtUnicode(2).SetFocus
           Exit Sub
        ElseIf mTQD03 > 0 And txtField(11).Text <> "" Then
           MsgBox "文字查詢不可輸入查名路徑!!", vbCritical
           txtField(11).SetFocus
           Exit Sub
        ElseIf mTQD03 = 0 And txtUnicode(1).Text & txtUnicode(2).Text <> "" Then
           MsgBox "圖形查詢不可輸入文字!!", vbCritical
           If txtUnicode(1).Text <> "" Then
              txtUnicode(1).SetFocus
           Else
              txtUnicode(2).SetFocus
           End If
           Exit Sub
        ElseIf Len(txtField(6).Tag) <> Len(txtField(6).Text) Then
                MsgBox "委查組群不可減少!!", vbCritical
                txtField(6).SetFocus
                Exit Sub
        'Modified by Lydia 2021/10/01 txtField(16) => textService
        'Modified by Lydia 2024/07/17 3519組群輸入啟用日 And DBDATE(txtField(3)) < "20240718"
        ElseIf InStr(txtField(6).Text, "3519") > 0 And textService.Text = "" And DBDATE(txtField(3)) < "20240718" Then
               MsgBox "指定商品/服務不可以空白！", vbCritical
               Exit Sub
        'Modified by Lydia 2024/07/17 3519組群輸入啟用日 And DBDATE(txtField(3)) < "20240718"
        ElseIf InStr(txtField(6).Text, "3519") = 0 And textService.Text <> "" And DBDATE(txtField(3)) < "20240718" Then
               MsgBox "指定商品/服務只有組群3519可輸入！", vbCritical
               Exit Sub
        'ElseIf chk1.Value = 0 And FraCase.Visible = True And Trim(txtField(12).Text & txtField(13).Text & txtField(14).Text & txtField(15).Text) <> "" Then
        '       'MsgBox "取消已收文時,案件請清空!!", vbCritical
        'ElseIf Trim(txtField(12).Text & txtField(13).Text & txtField(14).Text & txtField(15).Text) <> Trim(txtField(12).Tag & txtField(13).Tag & txtField(14).Tag & txtField(15).Tag) Then
        '       MsgBox "暫不提供修改已收文!", vbCritical
        '       Exit Sub
        'ElseIf txtField(20).Text <> txtField(20).Tag Then
         '      MsgBox "暫不提供修改已收文!", vbCritical
        '       Exit Sub
        'Added by Lydia 2016/04/29
        ElseIf txtField(23).Text <> txtField(23).Tag Then
               MsgBox "不提供修改發文日期!", vbCritical
               Exit Sub
        End If

        For Each oText In txtField
            txtField_Validate oText.Index, tmpBol
            If tmpBol = True Then
               Exit Sub
            End If
        Next

        jj = 0
        
        strExc(5) = "" 'Added by Lydia 2016/07/06 通知送件日期
        strExc(7) = "" '更新委查單資料
        strExc(8) = "" '更新查名單
        strExc(9) = "" '通知查名人員有關變更
        For Each oText In txtField
            If oText.Tag <> oText.Text Then
                Select Case jj
                    '委查人
                    Case 0:
                         strExc(7) = strExc(7) & " ,TMQ02=" & CNULL(oText.Text)
                         strExc(8) = strExc(8) & " ,TQA02=" & CNULL(oText.Text)
                         strExc(9) = strExc(9) & "委查人:" & Trim(lblAppNo(2).Caption) & vbCrLf
                    '查名人
                    Case 1: strExc(7) = strExc(7) & " ,TMQ10=" & CNULL(oText.Text)
                    '覆核主管
                    Case 2: strExc(7) = strExc(7) & " ,TMQ22=" & CNULL(oText.Text)
                    '覆核日期
                    Case 17
                            strExc(7) = strExc(7) & " ,TMQ23=" & CNULL(ChangeTStringToWString(oText.Text), True)
                    '申請日期
                    Case 3
                          'Modified by Lydia 2017/01/25
                          'strExc(7) = strExc(7) & " ,TMQ04=" & ChangeTStringToWString(oText.Text) & " ,TMQ05=" & ChangeTStringToWString(oText.Text)
                          strExc(7) = strExc(7) & " ,TMQ04=" & ChangeTStringToWString(oText.Text)
                          strExc(8) = strExc(8) & " ,TQA11=" & ChangeTStringToWString(oText.Text)
                          strExc(9) = strExc(9) & "申請日期:" & ChangeTStringToWString(oText.Text) & vbCrLf
                    '期限日期
                    Case 4: strExc(7) = strExc(7) & " ,TMQ06=" & CNULL(ChangeTStringToWString(oText.Text), True)
                    '查覆日期
                    Case 5
                         If oText.Text = "" Then
                            strExc(7) = strExc(7) & " ,TMQ11=NULL"
                            mbolSend = False
                         Else
                            strExc(7) = strExc(7) & " ,TMQ11=" & ChangeTStringToWString(oText.Text)
                            mbolSend = True
                         End If
                    '委查組群(這張單的組群) TMQ03->TQD05->TQA03 '考慮到資料的一致性,必需有相同組群數,不在結果明細中修改
                    Case 6
                         strExc(7) = strExc(7) & " ,TMQ03='" & Trim(oText.Text) & "'"
                         strExc(8) = strExc(8) & " ,TQA03=REPLACE(TQA03,'" & oText.Tag & "'," & CNULL(oText.Text) & ")"
                    '中文筆數
                    Case 7: strExc(7) = strExc(7) & " ,TMQ07=" & CNULL(oText.Text)
                    '英文筆數
                    Case 8: strExc(7) = strExc(7) & " ,TMQ08=" & CNULL(oText.Text)
                    '圖形筆數
                    Case 9: strExc(7) = strExc(7) & " ,TMQ09=" & CNULL(oText.Text)
                    '客戶名稱
                    'Remove by Lydia 2021/10/01 txtField(10)=>textCName、txtField(16)=>textService
                    'Case 10
                    '     strExc(8) = strExc(8) & " ,TQA04=" & CNULL(oText.Text)
                    '     strExc(9) = strExc(9) & "客戶名稱:" & Trim(oText.Text) & vbCrLf
                    'end 2021/10/01
                    '查名路徑
                    Case 11: strExc(7) = strExc(7) & " ,TMQ24=" & CNULL(oText.Text)
                    '案件
                    Case 12, 13, 14, 15
                          If jj = 12 Then strExc(8) = strExc(8) & " ,TQA16=" & CNULL(oText.Text)
                          If jj = 13 Then strExc(8) = strExc(8) & " ,TQA17=" & CNULL(oText.Text)
                          If jj = 14 Then strExc(8) = strExc(8) & " ,TQA18=" & CNULL(oText.Text)
                          If jj = 15 Then strExc(8) = strExc(8) & " ,TQA19=" & CNULL(oText.Text)
                          If InStr(strExc(9), "案件") = 0 Then
                             strExc(9) = strExc(9) & "案件:" & Trim(txtField(12)) & Trim(txtField(13)) & IIf(Trim(txtField(14)) & Trim(txtField(15)) = "000", "", Trim(txtField(14)) & Trim(txtField(15))) & vbCrLf
                          End If
                    '指定商品/服務
                    'Remove by Lydia 2021/10/01 txtField(10)=>textCName、txtField(16)=>textService
                    'Case 16
                    '     strExc(8) = strExc(8) & " ,TQA05=" & CNULL(oText.Text)
                    '     strExc(9) = strExc(9) & "指定商品/服務:" & Trim(oText.Text) & vbCrLf
                    'end 2021/10/01
                    '查覆結果已讀
                    Case 18
                         strExc(7) = strExc(7) & " ,TMQ19=" & CNULL(oText.Text)
                         mTMQ19 = Trim(oText.Text)
'                    '業務收文組群
'                    Case 19
'                         strExc(7) = strExc(7) & " ,TMQ20=" & CNULL(oText.Text)
'                         mTMQ20 = Trim(oText.Text)
                    '櫃台收文號
'                    Case 20
'                         strExc(7) = strExc(7) & " ,TMQ21=" & CNULL(oText.Text)
'                         FirstCP09 = Trim(oText.Text)
                    'Added by Lydia 2016/04/29
                    Case 19  '通知送件日期
                         'Modified by Lydia 2016/07/06
                         'strExc(7) = strExc(7) & " ,TMQ20=" & CNULL(ChangeTStringToWString(oText.Text), True)
                         strExc(5) = "UPDATE TMQCASEMAP SET TQC07=" & CNULL(ChangeTStringToWString(oText.Text), True) & " WHERE TQC02='" & lblAppNo(5).Caption & "' AND TQC03='" & mTQD02 & "' "
                    '查覆完成日期
                    Case 21
                         strExc(8) = strExc(8) & " ,TQA09=" & CNULL(ChangeTStringToWString(oText.Text), True)
                         strExc(9) = strExc(9) & "查覆完成日期:" & Trim(oText.Text) & vbCrLf
                    '是否撤回
                    Case 22
                         strExc(8) = strExc(8) & " ,TQA20=" & CNULL(oText.Text)
                         strExc(9) = strExc(9) & "是否撤回:" & Trim(oText.Text) & vbCrLf
                    'Added by Lydia 2017/01/25
                    Case 24 '收件分發日期
                        strExc(7) = strExc(7) & " ,TMQ05=" & CNULL(ChangeTStringToWString(oText.Text), True)
                End Select
            End If
            jj = jj + 1
        Next
        
        'Added by Lydia 2021/10/01 txtField(10)=>textCName、txtField(16)=>textService
        '客戶名稱
        If textCName.Tag <> textCName.Text Then
            strExc(8) = strExc(8) & " ,TQA04=" & CNULL(ChgSQL(textCName.Text))
            strExc(9) = strExc(9) & "客戶名稱:" & Trim(textCName.Text) & vbCrLf
        End If
        '指定商品/服務
        'Modified by Lydia 2024/07/17 3519組群輸入啟用日 And DBDATE(txtField(3)) < "20240718"
        If textService.Tag <> textService.Text And DBDATE(txtField(3)) < "20240718" Then
            strExc(8) = strExc(8) & " ,TQA05=" & CNULL(ChgSQL(textService.Text))
            strExc(9) = strExc(9) & "指定商品/服務:" & Trim(textService.Text) & vbCrLf
        End If
        'end 2021/10/01
        
        If (mTQA15 = "Y" And Chk1.Value = 0) Or (mTQA15 = "" And Chk1.Value = 1) Then
           If Chk1.Value = 1 Then
              strExc(8) = strExc(8) & " ,TQA15='Y'"
              mTQA15 = "Y"
              strExc(9) = strExc(9) & "查名單輸入時，已收文: Y" & vbCrLf
           Else
              strExc(8) = strExc(8) & " ,TQA15=NULL"
              mTQA15 = ""
              strExc(9) = strExc(9) & "查名單輸入時，已收文: N" & vbCrLf
           End If
        End If
        '文字查詢1~2
        If txtUnicode(1).Text & txtUnicode(2).Text & txtUnicode(1).Tag & txtUnicode(2).Tag <> "" Then
            'Added by Lydia 2021/10/01 用日期控制不用經過二進位處理存入TQA07-TQA08，直接存入TQA13-TQA14
            If strSrvDate(1) >= Form20上線日 Then
                For jj = 1 To 2
                    If txtUnicode(jj).Text <> txtUnicode(jj).Tag Then
                       strExc(9) = strExc(9) & "文字" & Trim(jj) & ":" & Trim(txtUnicode(jj).Text) & vbCrLf
                       If jj = 1 Then
                            strExc(8) = strExc(8) & " ,TQA07=NULL ,TQA13=" & CNULL(txtUnicode(jj).Text)
                       Else
                            strExc(8) = strExc(8) & " ,TQA08=NULL ,TQA14=" & CNULL(txtUnicode(jj).Text)
                       End If
                    End If
                Next jj
            Else
            'end 2021/10/01
                '先將Unicode存成二進位檔
                 UnicodeSave
                For jj = 1 To 2
                    txt1(jj).Text = txtUnicode(jj).Text
                    If txtUnicode(jj).Text <> txtUnicode(jj).Tag Then
                       strExc(9) = strExc(9) & "文字" & Trim(jj) & ":" & Trim(txt1(jj).Text) & vbCrLf
                       If txt1(jj).Text = txtUnicode(jj).Text Then
                          If jj = 1 Then
                             strExc(8) = strExc(8) & " ,TQA07=NULL ,TQA13=" & CNULL(txt1(jj).Text)
                          Else
                             strExc(8) = strExc(8) & " ,TQA08=NULL ,TQA14=" & CNULL(txt1(jj).Text)
                          End If
                       Else
                         '寫入二進位檔
                          If jj = 1 Then
                             strExc(8) = strExc(8) & " ,TQA13=NULL"
                             file_num1 = FreeFile
                             strFileName = m_AttachPath & "\unicode1.txt"
                             lLength = FileLen(strFileName) - 2
                             ReDim btTemp(lLength) As Byte
                            
                             Open strFileName For Binary As #file_num1
                             Get #file_num1, , btHead
                             Get #file_num1, , btTemp
                             Close #file_num1
                          Else
                             strExc(8) = strExc(8) & " ,TQA14=NULL"
                             file_num2 = FreeFile
                             strFileName = m_AttachPath & "\unicode2.txt"
                             lLength = FileLen(strFileName) - 2
                             ReDim btTemp2(lLength) As Byte
                            
                             Open strFileName For Binary As #file_num2
                             Get #file_num2, , btHead2
                             Get #file_num2, , btTemp2
                             Close #file_num2
                          End If
    
                       End If
                    End If
                Next jj
            End If  'Added 2021/10/01
        End If
        
        If strExc(7) & strExc(8) <> "" Then
           cnnConnection.BeginTrans
              If strExc(7) <> "" Then
                 strExc(7) = Trim(strExc(7))
                 strSql = "update trademarkquery set " & IIf(Left(strExc(7), 1) = ",", Mid(strExc(7), 2), strExc(7)) & " where tmq01='" & mTQD02 & "' "
                 Pub_SeekTbLog strSql
                 cnnConnection.Execute strSql, intI
                 'Added by Lydia 2017/02/07 一併變更同一申請編號的委查單
                 If InStr(strExc(9), "委查人:") > 0 Or InStr(strExc(9), "申請日期:") > 0 Then
                    strSql = "update trademarkquery set tmq02='" & txtField(0).Text & "',tmq04=" & ChangeTStringToWString(txtField(3).Text) & " where tmq18='" & mTQD01 & "' "
                    cnnConnection.Execute strSql, intI
                 End If
                 'end 2017/02/07
              End If
              If strExc(8) <> "" Then
                 strExc(8) = Trim(strExc(8))
                 strSql = "update tmqapp set " & IIf(Left(strExc(8), 1) = ",", Mid(strExc(8), 2), strExc(8)) & " where tqa01='" & mTQD01 & "' "
                 Pub_SeekTbLog strSql
                 cnnConnection.Execute strSql, intI
              End If
              '變更組群
              If txtField(6).Tag <> txtField(6).Text Then
                 If UPD_ClassDetail(txtField(6)) = False Then
                 End If
              End If
              '變更文字查詢(Unicode)
              If InStr(strExc(8), "TQA13=NULL") > 0 Or InStr(strExc(8), "TQA14=NULL") > 0 Then
                 If rsWrite.State <> adStateClosed Then rsWrite.Close
                 Set rsWrite = Nothing
                 rsWrite.CursorLocation = adUseClient
                 rsWrite.Open "select * from TMQApp where tqa01='" & mTQD01 & "' ", cnnConnection, adOpenStatic, adLockOptimistic
                 
                 If InStr(strExc(8), "TQA13=NULL") > 0 And txtUnicode(1).Text <> "" Then
                    rsWrite.Fields(6).Value = btTemp
                 End If
                 If InStr(strExc(8), "TQA14=NULL") > 0 And txtUnicode(2).Text <> "" Then
                    rsWrite.Fields(7).Value = btTemp2
                 End If
                 rsWrite.UPDATE
                 
              End If
              '查覆結果已讀
              If txtField(18).Tag <> txtField(18).Text Then
                 If txtField(18).Text = "Y" Then
                    strSql = "UPDATE TMQFILE SET TQF11='Y' WHERE TQF02='" & mTQD02 & "' AND TQF04<>'" & TMQ_附件F04 & "' "
                 Else
                    strSql = "UPDATE TMQFILE SET TQF11=NULL WHERE TQF02='" & mTQD02 & "' AND TQF04<>'" & TMQ_附件F04 & "' "
                 End If
                 cnnConnection.Execute strSql, intI
              End If
              
              '是否撤回
              If txtField(22).Tag <> txtField(22).Text Then
                 If txtField(22).Tag = "Y" Then
                     strSql = "UPDATE TMQDETAIL SET TQD06=null,TQD07=null WHERE TQD01='" & mTQD01 & "' "
                     cnnConnection.Execute strSql, intI
                     '圖形筆數:1筆
                     If mTQD03 = TMQ_AkindPic Then
                        strSql = "UPDATE trademarkquery SET TMQ11=null,TMQ07=null ,TMQ08=null ,TMQ09=1 WHERE TMQ18='" & mTQD01 & "' "
                        cnnConnection.Execute strSql, intI
                     Else
                        strExc(3) = "0": strExc(4) = "0"
                        For jj = 1 To 2
                           tmpE = 0: tmpC = 0
                           If cmdKey(jj - 1).Visible = True Then
                              If txtUnicode(jj).Text = "" Then tmpC = tmpC + 1
                           End If
                           '文字可能以替代字輸入,所以筆數依文字欄位來判斷
                           Call PUB_CountTxtNEC(tmpE, tmpC, txtUnicode(jj).Text)
                           strExc(3) = Val(strExc(3)) + tmpC
                           strExc(4) = Val(strExc(4)) + tmpE
                        Next jj
                        strSql = "select tmq01,sum(counting(tmq03)) cnt from trademarkquery WHERE TMQ18='" & mTQD01 & "' group by tmq01"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           RsTemp.MoveFirst
                           Do While Not RsTemp.EOF
                              strSql = "UPDATE trademarkquery SET TMQ11=null,TMQ07=" & Val(strExc(3)) * RsTemp.Fields("cnt") & " ,TMQ08=" & Val(strExc(4)) * RsTemp.Fields("cnt") & " ,TMQ09=null WHERE TMQ01='" & RsTemp.Fields("tmq01") & "' "
                              cnnConnection.Execute strSql, jj
                              RsTemp.MoveNext
                           Loop
                        End If
                     End If
                     
                     strSql = "UPDATE TMQAPP SET TQA20=null WHERE TQA01='" & mTQD01 & "' AND TQA20='Y'"
                     cnnConnection.Execute strSql, intI
                 Else
                    '保留資料,結果上不查
                     strSql = "UPDATE TMQDETAIL SET TQD06='" & TMQ_不查 & "',TQD07='" & ChangeTStringToTDateString(strSrvDate(2)) & "已由委查人自行撤回" & "' WHERE TQD01='" & mTQD01 & "' "
                     cnnConnection.Execute strSql, intI
                     strSql = "UPDATE trademarkquery SET TMQ11=" & strSrvDate(1) & ",TMQ07=0 ,TMQ08=0 ,TMQ09=0 WHERE TMQ18='" & mTQD01 & "' "
                     cnnConnection.Execute strSql, intI
                     strSql = "UPDATE TMQAPP SET TQA20='Y' WHERE TQA01='" & mTQD01 & "' AND TQA20 IS NULL"
                     cnnConnection.Execute strSql, intI
                 End If
              End If
              
              'Added by Lydia 2016/07/06 更新TQC07
              If strExc(5) <> "" Then
                 cnnConnection.Execute strExc(5), intI
              End If
              
           cnnConnection.CommitTrans
           
           For Each oText In txtField
               oText.Tag = oText.Text
           Next
           'Added by Lydia 2021/10/01 txtField(10)=>textCName、txtField(16)=>textService
           If DBDATE(txtField(3)) < "20240718" Then 'Added by Lydia 2024/07/17 3519組群輸入啟用日
              textService.Tag = textService.Text
           End If
           textCName.Tag = textCName.Text
           'end 2021/10/01
           
           txtUnicode(1).Tag = txtUnicode(1).Text
           txtUnicode(2).Tag = txtUnicode(2).Text
           
           MsgBox "存檔完成,請重新呼叫委查單!", vbInformation
            '變更查名人->通知原查名人員
            Sub1 = GetStrTitle
            If txtField(1).Tag <> txtField(1).Text Then
                'Modified by Lydia 2021/10/01 txtField(10) => textCName
                strExc(5) = "「" & textCName.Text & "」" & Trim(Sub1) & " 委查單: " & mTQD02 & " ,查名人變更為" & lblAppNo(3).Caption & "!!"
                strExc(6) = vbCrLf & vbCrLf & "如旨" & vbCrLf
        
                PUB_SendMail "QPGMR", txtField(1).Tag & ";" & txtField(1).Text, "", strExc(5), strExc(6)
            End If
           '通知查名人員有關變更
           If strExc(9) <> "" Then
              strExc(10) = ""
              'Modified by Lydia 2021/10/01 txtField(10) => textCName
              Call ChangeDataMail(mTQD01, strExc(9), textCName)
           End If
        End If
   Case Else
        MsgBox "error code!!"
   End Select
   
   'Added by Lydia 2019/05/21
   Set rs1 = Nothing
   Set rsB1 = Nothing
   
   Exit Sub
   'end 2019/05/21
ErrHand:
   If Err.Number <> 0 Then
      Screen.MousePointer = vbDefault
      MsgBox " 送出失敗！" & vbCrLf & Err.Description
   End If
End Sub
'查名單維護->變更查名內容通知信
'Modified by Lydia 2021/10/01 TextBox => Control
Private Sub ChangeDataMail(ByRef aNo As String, dStr As String, custTX As Control)
Dim rsAD As New ADODB.Recordset
Dim Sub1 As String
      strExc(0) = "SELECT TMQ01,TMQ10 FROM trademarkquery WHERE TMQ18='" & aNo & "' "
      intI = 1
      Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Sub1 = GetStrTitle
         rsAD.MoveFirst
         Do While Not rsAD.EOF
            '發Mail通知查名人員
            strExc(5) = "「" & custTX.Tag & "」" & Trim(Sub1) & " 委查單: " & rsAD.Fields("TMQ01") & " ,查名內容有所變更, 請檢查明細!!"
            strExc(6) = vbCrLf & vbCrLf & "變更內容如下：" & vbCrLf & dStr
    
            PUB_SendMail "QPGMR", rsAD.Fields("TMQ10"), "", strExc(5), strExc(6)
            rsAD.MoveNext
         Loop
      End If
End Sub
'回傳Mail主旨的單據類別
Private Function GetStrTitle() As String

   GetStrTitle = txtUnicode(1).Text & IIf(txtUnicode(2).Text <> "", "," & txtUnicode(2).Text, "")
   If Trim(GetStrTitle) <> "" Then
       GetStrTitle = "(" & GetStrTitle & ")"
   ElseIf mTQD03 = TMQ_AkindPic Then
       GetStrTitle = "(圖形查詢)"
   Else
       GetStrTitle = "(文字查詢)"
   End If
End Function

'申請編號所屬的委查單號完成,發mail通知委查人
'Modified by Lydia 2019/05/21 是否存mailcache
'Private Sub CloseMail(ByVal rNo As String, rTYP As String, rCnt As Integer, rTo As String)
Private Sub CloseMail(ByVal rNo As String, rTYP As String, ByVal rCnt As Integer, ByVal rTo As String, Optional ByVal bolCache As Boolean = False)
'rNo : 傳入查名單的申請編號／本所案號
Dim mSno As String, iX As Integer
Dim rsA As New ADODB.Recordset
Dim Sub1 As String
Dim sub2 As String
Dim subD As String
Dim rCC As String
Dim Str01 As String, Str02 As String
Dim strTempCC As String 'Added by Lydia 2016/06/27 指定職代
'Added by Lydia 2017/08/28
Dim strSalesP As String '指定智權人員
Dim tmpArr As Variant
Dim bolS2X As Boolean 'Added by Lydia 2022/06/06 是否有中所人員
Dim strMailKind As String 'Added by Lydia 2024/10/14 覆核增加確認機制; 依覆核結果區別通知內容:1-經覆核結果為「相同TQD09=4」、「近似TQD09=5」時，通知內容為：覆核意見：（帶入覆核意見）,
                                                                 '2-經覆核結果為「稍近似TQD09=6」或「無TQD09=7」時，通知內容為：覆核意見：（帶入覆核意見）（如有覆核意見,請先列覆核意見，再接續原通知內容，如無則僅列通知內容）
Dim strTQD10List As String, strCP13name As String 'Added by Lydia 2024/10/14

If rTYP <> "撤回" Then
   Sub1 = GetStrTitle
   strExc(1) = IIf(InStr(rTYP, "查覆") > 0, "查覆", "")
   strExc(1) = IIf(InStr(rTYP, "覆核") > 0, "覆核", strExc(1))
   
   'Modified by Lydia 2016/04/21 加已收文案件的提示
   'Modified by Lydia 2019/05/21 傳入本所案號
   'Sub1 = "「" & txtField(10).Text & "」" & Trim(Sub1) & " 查名作業已完成" & strExc(1) & "(共" & rCnt & "張)" & _
          IIf(FirstCP09 <> "", "，已收文案件" & FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) <> "000", "-" & FirstCP(3) & "-" & FirstCP(4), ""), "")
   'Modified by Lydia 2021/10/01 txtField(10) => textCName
   Sub1 = "「" & textCName.Text & "」" & Trim(Sub1) & " 查名作業已完成" & strExc(1) & "(共" & rCnt & "張)" & _
          IIf(rNo <> "" And Left(rNo, 1) <> "H", "，已收文案件" & rNo, "")
          
   'Added by Lydia 2016/09/12 覆核結果更改
   If InStr(rTYP, "更改") > 0 Then
      'Modified by Lydia 2019/05/21 傳入本所案號
      'Sub1 = "「" & txtField(10).Text & "」" & Trim(GetStrTitle) & " 查名作業的覆核結果有更改(委查單號: " & mTQD02 & ")" & _
             IIf(FirstCP09 <> "", "，已收文案件" & FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) <> "000", "-" & FirstCP(3) & "-" & FirstCP(4), ""), "")
      'Modified by Lydia 2021/10/01 txtField(10) => textCName
      Sub1 = "「" & textCName.Text & "」" & Trim(GetStrTitle) & " 查名作業的覆核結果有更改(委查單號: " & mTQD02 & ")" & _
            IIf(rNo <> "" And Left(rNo, 1) <> "H", "，已收文案件" & rNo, "")
    End If
   'end 2016/09/12
   
   'Modified by Lydia 2016/05/03 拿掉已收文
   'sub2 = vbCrLf & vbCrLf & "請進入承辦人系統->智權部->日常作業->查覆區" & IIf(FirstCP09 <> "", "的已收文", "") _
          & "，點選記錄進入查覆明細作業查閱結果和附件。" & vbCrLf & vbCrLf
   'Modified by Lydia 2019/07/03 更名
   'sub2 = vbCrLf & vbCrLf & "請進入承辦人系統->智權部->日常作業->查覆區" & _
           "，點選記錄進入查覆明細作業查閱結果和附件。" & vbCrLf & vbCrLf
   sub2 = vbCrLf & vbCrLf & "請進入承辦人系統->智權部->專利商標作業->商標查名／查覆區" & _
           "，點選記錄進入查覆明細作業查閱結果和附件。" & vbCrLf & vbCrLf
   'Modified by Lydia 2019/05/21
   'If FirstCP09 <> "" Then
   If rNo <> "" And Left(rNo, 1) <> "H" Then
      sub2 = sub2 & "已收文案件日後可到共同查詢的卷宗區點選查名結果。" & vbCrLf & vbCrLf
   End If
   
   'Added by Lydia 2016/06/27 核可案
   If InStr(rTYP, "核可案") > 0 Then
      'Modified by Lydia 2019/05/21 傳入本所案號
      'Sub1 = IIf(FirstCP09 <> "", "已收文案件" & FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) <> "000", "-" & FirstCP(3) & "-" & FirstCP(4), ""), "")
      Sub1 = IIf(rNo <> "" And Left(rNo, 1) <> "H", "，已收文案件" & rNo, "")
      'Modified by Lydia 2021/10/01 txtField(10) => textCName
      Sub1 = IIf(Sub1 <> "", Sub1 & "，", "") & "「" & textCName.Text & "」" & Trim(GetStrTitle) & " 查名作業已完成" & strExc(1) & "(共" & rCnt & "張)，並且已經過相關主管核可，可以辦理申請案!"
      GoTo JumpMailTo
   End If
   'end 2016/06/27
   
   'Added by Lydia 2022/06/06 覆核增加確認機制; 當查名結果從近似本所案△改為非近似本所案是查名人員的看法，尚需要智權人員確認結果，所以修改通知email提醒兩方要再次確認(杜經理的需求by 嘉雯)
   If InStr(rTYP, "增加確認") > 0 Then
      'Email內文：因　Ａ智權人員　之客戶委查商標與　Ｂ智權人員　的客戶　Ｂ智權人員之本所客戶名稱　商標相同或近似，但經覆核後，商標部認為在商標整體上有可辦理機會，但為避免往後客戶間的衝突事件，請　Ａ智權人員　在向客戶說明之前，先與Ｂ智權人員連絡進行協商，若Ｂ智權人員表示同意，請於收文時於接洽單備註處載明，且於收文後，將相關資訊存入該案之卷宗區中。
      '收件者：委查人(A智權)、A智權的區主管、相關衝突智權(B智權)、B智權的區主管、杜協理、中所林協理、商標承辦人員、商標部覆核人員、江協理。
      Sub1 = Sub1 & "，智權人員請確認覆核結果"
      'Modified by Lydia 2024/10/14 + TQD10
      strExc(0) = " select tqd01,tqd02,tqd03,tqd04,tqd09 as tqd06,tqd08,tqd10 from tmqdetail where TQD01='" & mTQD01 & "' and tqd06 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') order by 5,6,2,3,4 "
      intI = 1:    strExc(1) = "":        strExc(7) = ""
      subD = "" 'Added by Lydia 2023/09/22
      strMailKind = "": strTQD10List = ""  'Added by Lydia 2024/10/14
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
           RsTemp.MoveFirst
           strExc(4) = ""
           '抓全部的審定號
           Do While Not RsTemp.EOF
              'Added by Lydia 2023/09/22
              If InStr(subD & ",", "" & RsTemp.Fields("tqd02")) = 0 Then
                 subD = subD & "," & RsTemp.Fields("tqd02")
              End If
              'end 2023/09/22
              '不抓相同結果
              If strExc(4) <> RsTemp.Fields("tqd06") & RsTemp.Fields("tqd08") Then
                tmpArr = Empty
                tmpArr = Split(Trim("" & RsTemp.Fields("tqd08")), ",")
                For intI = 0 To UBound(tmpArr)
                   If Trim(tmpArr(intI)) <> "" Then
                      If InStr(strExc(7) & ",", Trim(tmpArr(intI))) = 0 Then
                          strExc(7) = strExc(7) & Trim(tmpArr(intI)) & ","
                          'Added by Lydia 2024/10/14 覆核增加確認機制; 依覆核結果(tqd09 as tqd06) 區別通知內容
                          If "" & RsTemp.Fields("tqd06") = "4" Or "" & RsTemp.Fields("tqd06") = "5" Then
                             strMailKind = "1"
                          End If
                          If "" & RsTemp.Fields("tqd10") <> "" Then  '覆核意見
                             If InStr("|" & strTQD10List, "|" & RsTemp.Fields("tqd10")) = 0 Then
                                strTQD10List = strTQD10List & "|" & RsTemp.Fields("tqd10")
                             End If
                          End If
                          'end 2024/10/14
                      End If
                   End If
                Next intI
              End If
              strExc(4) = RsTemp.Fields("tqd06") & RsTemp.Fields("tqd08")
              RsTemp.MoveNext
           Loop
           
           strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = "": strExc(6) = ""
           If strExc(7) <> "" Then
              If strMailKind = "" Then strMailKind = "2" 'Added by Lydia 2024/10/14 覆核增加確認機制; 依覆核結果(tqd09 as tqd06) 區別通知內容

              tmpArr = Empty
              tmpArr = Split(strExc(7), ",")
              For intI = 0 To UBound(tmpArr)
                  If Trim(tmpArr(intI)) <> "" Then
                      'Modifiedd by Lydia 2024/10//09 +CU80
                      strExc(0) = "select tm23, nvl(cu04,nvl(cu05,cu06)) as cname,tm01,tm02,tm03,tm04,nvl(cu80,'N') as CU80 " & _
                                       "from trademark,customer where (tm12=" & CNULL(IIf(Left(Trim(tmpArr(intI)), 1) = "P", Mid(Trim(tmpArr(intI)), 2), Trim(tmpArr(intI)))) & " or tm15=" & CNULL(IIf(Left(Trim(tmpArr(intI)), 1) = "P", Mid(Trim(tmpArr(intI)), 2), Trim(tmpArr(intI)))) & " ) and substr(tm23,1,8)=cu01(+) and  substr(tm23,9,1)=cu02(+) "
                      'Added by Lydai 2022/07/12 區分本所案或統一案/統一公司案件(P審定號/申請號，系統發通知要求應徵詢的智權同仁為「蘇威廷」);
                      If Left(Trim(tmpArr(intI)), 1) = "P" Then
                         '統一公司非本所案用T-230642為代表
                         'Modifiedd by Lydia 2024/10//09 +CU80
                         strExc(0) = strExc(0) & " union select tm23, nvl(cu04,nvl(cu05,cu06)) as cname,'T' as tm01,'999999' as tm02, '0' as tm03,'00' as tm04,nvl(cu80,'N') as CU80 " & _
                                          "from trademark,customer where tm01='T' and tm02='230642' and tm03='0' and tm04='00' and substr(tm23,1,8)=cu01(+) and  substr(tm23,9,1)=cu02(+) "
                      End If
                      strExc(0) = strExc(0) & " order by tm01,tm02,tm03,tm04"
                      'end 2022/07/12
                      iX = 1
                      Set rsA = ClsLawReadRstMsg(iX, strExc(0))
                      If iX = 1 Then
                        If Left(Trim(tmpArr(intI)), 1) = "P" Then
                            '區分本所案或統一案/統一公司案件(P審定號/申請號，系統發通知要求應徵詢的智權同仁為「蘇威廷」);
                            strExc(2) = "A2026"
                        Else
                            strExc(2) = PUB_GetAKindSalesNo(rsA.Fields("tm01"), rsA.Fields("tm02"), rsA.Fields("tm03"), rsA.Fields("tm04"))
                        End If
                        'Added by Lydia 2024/10/09 覆核增加確認機制;覆核後為稍近似的案件,若前案為無效客戶時,無須發送通知
                        'Modified by Lydia 2024/10/14 覆核增加確認機制; 依覆核結果區別通知內容 +  And strMailKind = "2"
                        If InStr("解散,廢止,撤銷,死亡", "" & rsA.Fields("cu80")) > 0 And strMailKind = "2" Then
                           strExc(2) = ""
                        End If
                        'end 2024/10/09
                        If strExc(2) <> "" Then
                            'Added by Lydia 2024/10/14 判斷智權人員名稱
                            strCP13name = ""
                            If Left(strExc(2), 4) = "MCTF" Then
                               strExc(0) = Pub_GetSpecMan(strExc(2))
                               If strExc(0) <> "" Then
                                  strCP13name = PUB_ReadUserData(strExc(0))
                               End If
                            End If
                            If strCP13name = "" Then
                               strCP13name = GetStaffName(strExc(2), True)
                            End If
                            'end 2024/10/14
                            If InStr(strExc(4) & ",", "(" & rsA.Fields("tm01") & "-" & rsA.Fields("tm02")) = 0 Then
                               'Added by Lydia 2022/07/12 統一公司非本所案
                               If Left(Trim(tmpArr(intI)), 1) = "P" And "" & rsA.Fields("tm02") = "999999" Then
                                   'Modified by Lydia 2024/10/14 改用變數; GetStaffName(strExc(2), True)>> strCP13name
                                   strExc(4) = strExc(4) & "、 " & strCP13name & "的" & rsA.Fields("tm23") & rsA.Fields("cname") & Mid(Trim(tmpArr(intI)), 2)
                               Else
                               'end 2022/07/12
                                   'Modified by Lydia 2024/10/14 改用變數; GetStaffName(strExc(2), True)>> strCP13name
                                   strExc(4) = strExc(4) & "、 " & strCP13name & "的" & rsA.Fields("tm23") & rsA.Fields("cname") & "(" & rsA.Fields("tm01") & "-" & rsA.Fields("tm02") & IIf(rsA.Fields("tm03") <> "0", "-" & rsA.Fields("tm03"), "") & IIf(rsA.Fields("tm04") <> "00", "-" & rsA.Fields("tm04"), "") & ")"
                               End If 'Added by Lydia 2022/07/12
                               
                               '只要兩方智權人員有一人為中所人員，才通知林協理
                               'Mark by Lydia 2025/07/21 不用通知林協理
                               'If bolS2X = False And PUB_GetST06(strExc(2)) = "2" Then
                               '     bolS2X = True
                               'End If
                               'end 2025/07/21
                               If InStr(";" & strExc(5), strExc(2)) = 0 Then
                                   strExc(5) = strExc(5) & ";" & strExc(2) '相關衝突智權(B智權)
                                   'Modified by Lydia 2022/09/21 外商日文組的此類案件協調不寄給承辦人，而直接只寄給主管
                                   'If InStr(";" & strExc(6), strExc(2)) = 0 Then
                                   '    strExc(6) = strExc(6) & ";" & strExc(2)
                                   'End If
                                   strExc(1) = strExc(2)
                                   If rsA.Fields("tm01") = "FCT" Or rsA.Fields("tm01") = "CFT" Then
                                       strExc(1) = PUB_GetF11ToMan(strExc(2))
                                   End If
                                   '收件者
                                   If InStr(";" & strExc(6), strExc(1)) = 0 Then
                                       strExc(6) = strExc(6) & ";" & strExc(1)
                                   End If
                                   'end 2022/09/21
                                   strExc(3) = GetDeptMan(GetST15(strExc(2)))
                                   'Modified by Lydia 2022/09/21 排除外商日文組
                                   'If InStr(";" & strExc(6), strExc(3)) = 0 Then
                                   If InStr(";" & strExc(6), strExc(3)) = 0 And strExc(1) = strExc(2) Then
                                       strExc(6) = strExc(6) & ";" & strExc(3) '相關衝突智權(B智權)的區主管
                                   End If
                               End If
                            End If
                        End If
                      End If
                  End If
              Next intI
              If strExc(5) & strExc(6) <> "" Then 'Added by Lydia 2024/10/09 覆核增加確認機制;覆核後為稍近似的案件,若前案為無效客戶時,無須發送通知
                  'Added by Lydia 2024/10/14
                  If strTQD10List <> "" Then strTQD10List = Mid(strTQD10List, 2)
                  If strMailKind = "1" Then
                     sub2 = "覆核意見：" & Replace(strTQD10List, "|", vbCrLf)
                  Else
                  'end 2024/10/14
                     strExc(5) = Mid(strExc(5), 2)
                     sub2 = "因 " & lblAppNo(2) & " 之客戶委查商標與" & Mid(strExc(4), 2) & " 商標相同或近似，" & _
                               "但經覆核後，商標部認為在商標整體上有可辦理機會，但為避免往後客戶間的衝突事件，" & _
                               "請 " & lblAppNo(2) & " 在向客戶說明之前，先與 " & PUB_ReadUserData(strExc(5)) & " 連絡進行協商，" & _
                               "若 " & PUB_ReadUserData(strExc(5)) & " 表示同意，請於收文時於接洽單備註處載明，且於收文後，將相關資訊存入該案之卷宗區中。"
                     If strTQD10List <> "" Then sub2 = "覆核意見：" & Replace(strTQD10List, "|", vbCrLf) & vbCrLf & vbCrLf & sub2    'Added by Lydia 2024/10/14
                  End If 'Added by Lydia 2024/10/14
                  '系統通知發送：委查人(A智權)、A智權的區主管、相關衝突智權(B智權)、B智權的區主管、杜協理(全所智權部主管)、中所林協理(中所智權部主管)、商標承辦人員(FirstCP14)、商標部覆核人員(strUserNme)、江協理(V2)。
                                            'rTo 傳入=查名結果給近似△的查名人員
                  strExc(2) = GetDeptMan(GetST15(txtField(0)))
                  If InStr(strExc(6), txtField(0)) = 0 Then
                      strExc(6) = ";" & txtField(0) & strExc(6)
                  End If
                  If InStr(strExc(6), strExc(2)) = 0 Then
                      strExc(6) = ";" & strExc(2) & strExc(6)
                  End If
                  '只要兩方智權人員有一人為中所人員，才通知林協理
                  'Modfied by Lydia 2025/07/21  不用通知林協理 => 拿掉 & IIf(bolS2X = True, ";" & Pub_GetSpecMan("中所智權部主管"), "")
                  rTo = Mid(strExc(6), 2) & ";" & Pub_GetSpecMan("全所智權部主管") & _
                          IIf(FirstCP14 <> "", ";" & FirstCP14, "") & ";" & strUserNum & ";" & Pub_GetSpecMan("V2") & ";" & rTo
                  rTo = Replace(rTo, ";;", ";")
                  'Added by Lydia 2023/09/22 經智權同仁反映,若該客戶同時就同一文字委查多個組群查名單時,系統通知信內容無法使其於查名/查覆區快速找到該筆查名單;
                                        '建議在主旨加入「委查單號」,即：委查單號:XXX「AAA-BBB」(XXXX,YYY) 查名作業已完成覆核(共n張)，智權人員請確認覆核結果
                  If subD <> "" Then
                     Sub1 = "委查單號:" & Mid(subD, 2) & Sub1
                  End If
                  'end 2023/09/22
                  If bolCache = True Then
                       strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                          " values( '" & strUserNum & "','" & rTo & "',to_char(sysdate,'yyyymmdd')" & _
                          ",to_char(sysdate,'hh24miss'),'" & ChgSQL(Sub1) & "','" & ChgSQL(sub2) & "',null)"
                       cnnConnection.Execute strSql
                  Else
                       PUB_SendMail strUserNum, rTo, "", Sub1, sub2
                  End If
              End If 'Added by Lydia 2024/10/09 覆核增加確認機制;覆核後為稍近似的案件,若前案為無效客戶時,無須發送通知
           End If 'If strExc(7) <> "" Then
         Exit Sub
      End If
   End If 'If InStr(rTYP, "增加確認") > 0 Then
   'end 2022/05/024
   
   strExc(2) = IIf(InStr(rTYP, "近似本所案") > 0 Or InStr(rTYP, "相同本所案") > 0, "近似/相同本所案", "")
   
     If strExc(2) <> "" Then
        'Modified by Lydia 2016/06/27 近似本所案都需覆核,先通知覆核人員
        If InStr(rTYP, "查覆") > 0 And strSrvDate(1) >= TMQFileFTP Then
           Sub1 = Replace(Sub1, "張)", "張)，有" & strExc(2) & "的結果需覆核")
           sub2 = vbCrLf & vbCrLf & "請進入覆核區進行覆核作業"
           rTo = Pub_GetSpecMan("內商查名覆核通知")
           If Len(rTo) <= 6 Then 'Added by Lydia 2025/07/21 排除職代直接納入收件人
              'Modified by Lydia 2021/11/12 林嘉雯請假時職代處理=>取得人員請假的職代
              'strTempCC = GetCaseDutyAgent(rTo, "", False, subD)
              'If strTempCC <> "" Then
              '   strTempCC = "69008" '因為嘉雯的假單職代是淑鈴,沒有覆核權限,改發給林純貞
              '   sub2 = "因收件人" & subD & "，請副本收件人處理此郵件。" & vbCrLf & vbCrLf & sub2
              'End If
              strTempCC = GetDutyList(rTo)
              If strTempCC <> "" Then
                  sub2 = "因收件人" & subD & "，請副本收件人處理此郵件。" & vbCrLf & vbCrLf & sub2
              End If
              'end 2021/11/12
           End If 'Added by Lydia 2025/07/21
           GoTo JumpMailTo
        End If
        '覆核結果取代查名結果
        'Remove by Lydia 2016/07/07
        'If strSrvDate(1) < TMQFileFTP Then
        '    strExc(0) = " select tqd01,tqd02,tqd03,tqd04,tqd06,tqd08 from tmqdetail where TQD01='" & mTQD01 & "' and tqd06 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') order by 5,6,2,3,4 "
        'Else
            strExc(0) = " select tqd01,tqd02,tqd03,tqd04,tqd09 as tqd06,tqd08 from tmqdetail where TQD01='" & mTQD01 & "' and tqd09 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') order by 5,6,2,3,4 "
        'End If
        ''end 2016/06/27
        intI = 1: strExc(1) = ""
        strExc(7) = ""  'Added by Lydia 2017/08/28
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           RsTemp.MoveFirst
           strExc(4) = "" 'Added by Lydia 2016/06/27
           '抓全部的審定號
           Do While Not RsTemp.EOF
              'Modified by Lydia 2016/06/27 不抓相同結果
              If strExc(4) <> RsTemp.Fields("tqd06") & RsTemp.Fields("tqd08") Then
                'Modified by Lydia 2017/08/28 區分本所案或統一案;
                '查名若發現與統一公司商標近似情形時 , 仍列為客戶間利益衝突案件:
                '1.於審定號/申請號欄位，在號數前加上P時，查名結果可為「相同△、近似△」，並且不檢查是否為本所案；
                '2.遇到P審定號/申請號，系統發通知要求應徵詢的智權同仁為「蘇威廷」。
                'strExc(1) = strExc(1) & RsTemp.Fields("tqd08") & ","
                tmpArr = Empty
                tmpArr = Split(Trim("" & RsTemp.Fields("tqd08")), ",")
                For intI = 0 To UBound(tmpArr)
                   If Trim(tmpArr(intI)) <> "" Then
                      If Left(Trim(tmpArr(intI)), 1) = "P" Then
                         strExc(7) = strExc(7) & Trim(tmpArr(intI)) & ","
                      Else
                         strExc(1) = strExc(1) & Trim(tmpArr(intI)) & ","
                      End If
                   End If
                Next intI
                'end 2017/08/28
                
                If "" & RsTemp("tqd06") = TMQ_近似1 Then
                    Str01 = Str01 & RsTemp.Fields("tqd08") & ","
                ElseIf "" & RsTemp("tqd06") = TMQ_近似2 Then
                    Str02 = Str02 & RsTemp.Fields("tqd08") & ","
                End If
              End If
              strExc(4) = RsTemp.Fields("tqd06") & RsTemp.Fields("tqd08") 'Added by Lydia 2016/06/27
              RsTemp.MoveNext
           Loop
             '發mail時,unicode字元會造成mail內文中斷
            'Modified by Lydia 2016/06/27
            'Remove by Lydia 2016/07/07
            'If strSrvDate(1) < TMQFileFTP Then
            '   sub2 = sub2 & "已查覆但與本所"
            'Else
               sub2 = sub2 & "已覆核但與本所"
            'End If
            ''end 2016/06/27
            
            'Added by Lydia 2017/08/28 統一公司案件
            If strExc(7) <> "" Then
                If strSalesP = "" Then
                    'Added by Lydia 2023/12/26
                    If DBDATE(txtField(3)) >= 新部門啟用日 Then
                       strExc(0) = " select decode(s1.st04,'1',s1.st01,nvl(a0924,a0908)) sno,decode(s1.st04,'1',s1.st02,getstaffnamelist(nvl(a0924,a0908))) sname " & _
                            "from staff s1, acc090,acc090new where s1.st01='A2026' and s1.st03=a0901(+) and s1.st93=a0921(+) "
                    Else
                    'end 2023/12/26
                       strExc(0) = " select decode(s1.st04,'1',s1.st01,s2.st01) sno,decode(s1.st04,'1',s1.st02,s2.st02) sname " & _
                            "from staff s1,staff s2,acc090 where s1.st01='A2026' and s1.st03=a0901(+) and a0908=s2.st01(+)  "
                    End If
                    intI = 1
                    Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then
                       strSalesP = "" & rsA.Fields("sname")
                       rCC = rCC & strSalesP & ","
                    End If
                End If
                If InStr(Str01, strExc(7)) > 0 And InStr(Str02, strExc(7)) > 0 Then
                    Str02 = Replace(Str02, strExc(7) & ",", "")
                End If
            End If
            
            '本所案
            If strExc(1) <> "" Then
            'end 2017/08/28
                strExc(1) = GetAddStr(strExc(1))
                'Modified by Lydia 2017/03/15 拿掉閉卷
                'strExc(0) = "SELECT '1' ORD,TM01,TM02,TM03,TM04,TM15 AS TM1215 FROM TRADEMARK WHERE TM10='000' AND TM29||TM57||TM73 IS NULL AND TM15 IN (" & strExc(1) & ") " & _
                            "Union SELECT '2' ORD,TM01,TM02,TM03,TM04,TM12 AS TM1215 FROM TRADEMARK WHERE TM10='000' AND TM29||TM57||TM73 IS NULL AND TM12 IN (" & strExc(1) & ") ORDER BY 1,2,3 "
                strExc(0) = "SELECT '1' ORD,TM01,TM02,TM03,TM04,TM15 AS TM1215 FROM TRADEMARK WHERE TM10='000' AND TM15 IN (" & strExc(1) & ") " & _
                            "Union SELECT '2' ORD,TM01,TM02,TM03,TM04,TM12 AS TM1215 FROM TRADEMARK WHERE TM10='000' AND TM12 IN (" & strExc(1) & ") ORDER BY 1,2,3 "
    
                intI = 1
                Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   rsA.MoveFirst
                   Do While Not rsA.EOF
                      '若同時有一樣的審定號,只顯示在相同本所案
                      If InStr(Str01, rsA.Fields("tm1215")) > 0 And InStr(Str02, rsA.Fields("tm1215")) > 0 Then
                         Str02 = Replace(Str02, rsA.Fields("tm1215") & ",", "")
                      End If
                      'Remove by Lydia 2016/04/26 不要寄副本給近似本所案的智權人員
                      'strExc(3) = PUB_GetFCTSalesNo(rsA.Fields("tm01"), rsA.Fields("tm02"), rsA.Fields("tm03"), rsA.Fields("tm04"))
                      'If strExc(3) <> "" Then rCC = rCC & strExc(3) & ";"
                      'Added by Lydia 2016/04/28 內文代出近似本所案的智權人員
                      strExc(3) = PUB_GetAKindSalesNo(rsA.Fields("tm01"), rsA.Fields("tm02"), rsA.Fields("tm03"), rsA.Fields("tm04"))
                      If strExc(3) <> "" Then rCC = rCC & GetStaffName(strExc(3)) & ","
                      rsA.MoveNext
                   Loop
                End If
            End If 'end 2017/08/28
            
            If Str01 <> "" Then
               sub2 = sub2 & Mid(Str01, 1, Len(Str01) - 1) & Mid(TMQ_近似T1, 1, 2) & ",不得申請" & vbCrLf
            End If
            If Str02 <> "" Then
               sub2 = sub2 & Mid(Str02, 1, Len(Str02) - 1) & Mid(TMQ_近似T2, 1, 2) & ",不得申請" & vbCrLf
            End If
            
            'Added by Lydia 2017/08/28 加統一案件備註
            If strSalesP <> "" Then
               sub2 = sub2 & "P.S.統一公司案件會在審定號/申請號數前加上P" & vbCrLf
            End If
            'end 2017/08/28
            sub2 = sub2 & vbCrLf
        End If
       
        strExc(2) = ""
        'Modified by Lydia 2016/04/28
        'sub2 = sub2 & "商標處主管由承辦人系統->商標處->商標委查作業->覆核區,進行覆核作業。"
        'Modified by Lydia 2016/06/27
        'sub2 = sub2 & "若欲申請，請與相關智權人員（" & Mid(rCC, 1, Len(rCC) - 1) & "）協商或委請商標處主管進行覆核。"
        sub2 = sub2 & "若欲申請，請與相關智權人員（" & Mid(rCC, 1, Len(rCC) - 1) & "）協商。"
                
     End If
    'Modified by Lydia 2016/04/26
    'PUB_SendMail strUserNum, rTo, "", Sub1, sub2, , , , , , rCC
    
JumpMailTo: 'Added by Lydia 2016/06/27
    'Added by Lydia 2019/05/21 依照各案件的現況通知
    If bolCache = True Then
          strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
             " values( '" & strUserNum & "','" & rTo & "',to_char(sysdate,'yyyymmdd')" & _
             ",to_char(sysdate,'hh24miss'),'" & ChgSQL(Sub1) & "','" & sub2 & "'," & CNULL(strTempCC) & ")"
          cnnConnection.Execute strSql
    Else
    'end 2019/05/21
        PUB_SendMail strUserNum, rTo, "", Sub1, sub2, , , , , , strTempCC
    End If
Else

End If

End Sub

'移動資料列前,先確認是否要存檔
'Modified by Lydia 2019/12/12 + rTxt3
'Modified by Lydia 2021/10/01 TextBox=>Control
Private Sub ChangeDetailAns(ByVal Id As Integer, ByRef rCbo As ComboBox, ByRef rTxt As Control, Optional bolC As Boolean, Optional ByRef rTxt2 As Control, Optional ByRef rTxt3 As Control)
Dim bolA As Boolean
bolC = False
bolA = False
If Id = 0 Or Id = 2 Then
   'Added by Lydia 2023/07/06 開放覆核主管可修改「審定號/申請號」
   If R_type = "M" Then
      If rCbo.Tag <> rCbo.Text Or rTxt.Tag <> rTxt.Text Then bolA = True
   Else
   'end 2023/07/06
      If rCbo.Tag <> rCbo.Text Or rTxt.Tag <> rTxt.Text Or rTxt2.Tag <> rTxt2.Text Then bolA = True
   End If 'Added by Lydia 2023/07/06
Else '覆核主管
   'Modified by Lydia 2019/12/12 + rTxt3
   If rCbo.Tag <> rCbo.Text Or rTxt.Tag <> rTxt.Text Or rTxt3.Tag <> rTxt3.Text Then bolA = True
End If
If bolA = True Then
   'Added by Lydia 2016/04/25 +判斷是否有存檔權限
   If rCbo.Locked = True Then
      MsgBox "無權限修改!", vbCritical
      If Id = 0 Or Id = 2 Then
        rCbo.Text = rCbo.Tag: rTxt.Text = rTxt.Tag: rTxt2.Text = rTxt2.Tag
      Else
        rCbo.Text = rCbo.Tag: rTxt.Text = rTxt.Tag
      End If
      bolA = False
   Else '存檔
      '未查覆完畢前,自動存檔
      If mbolSend = False Then
         Call cmdSaveD_Click(Id)
      Else
        'Added by Lydia 2016/05/03 查覆完畢後,要詢問是否存檔
        If MsgBox("查覆結果或意見有變更,是否要存檔?", vbYesNo + vbDefaultButton1) = vbYes Then
           Call cmdSaveD_Click(Id)
        Else
            If Id = 0 Or Id = 2 Then
                rCbo.Text = rCbo.Tag: rTxt.Text = rTxt.Tag: rTxt2.Text = rTxt2.Tag
            Else
                rCbo.Text = rCbo.Tag: rTxt.Text = rTxt.Tag
            End If
            bolC = True
        End If
      End If
   End If
End If

Exit Sub

End Sub



Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
If nCol >= 0 Then GRD1.col = nCol
If nRow >= 0 Then GRD1.row = nRow
End Sub
' 更新各控制項的狀態
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   GRD1.ToolTipText = ""
   If GRD1.MouseRow <> 0 And GRD1.MouseCol > 0 Then
      If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
         GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
      End If
   End If
End Sub

Private Sub grd1_SelChange()
Dim TmpRow As Integer
TmpRow = GRD1.MouseRow

If iStiu = 1 Then
   If R_type = "M" Then
        'Modified by Lydia 2019/12/12 + 是否出名
        'Call ChangeDetailAns(1, Cbo2, txtDT(2))
        Call ChangeDetailAns(1, Cbo2, txtDt(2), , , txtDt(6))
   Else
        Call ChangeDetailAns(0, Cbo1, txtDt(0), , txtDt(1))
        '不同頁籤不記錄審定號/申請號
        If txtDt(1).Text <> "" Then
           mPrevTM1215 = txtDt(1).Text
        ElseIf (Cbo1.Text = TMQ_近似T1 Or Cbo1.Text = TMQ_近似T2) And txtDt(1).Text = "" Then
               txtDt(1).Text = mPrevTM1215
        End If
        
   End If
End If

GRD1.Visible = False
If TmpRow > 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 0
      GRD1.row = dblPrevRow
      GRD1.Text = ""
      For jj = 0 To GRD1.Cols - 1
         GRD1.col = jj
         GRD1.CellBackColor = QBColor(15)
      Next jj
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = TmpRow
   dblPrevRow = GRD1.row
    If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
       GRD1.Text = "V"
       For jj = 0 To GRD1.Cols - 1
          GRD1.col = jj
          GRD1.CellBackColor = &HFFC0C0
       Next jj
       'Modified by Lydia 2019/12/12 +Tab
       Call SetComboAns(0, GRD1, dblPrevRow, Cbo1, Cbo2, txtDt(0), txtDt(1), txtDt(2))
       Call ShowFieldTQD09(0)
    End If
End If
GRD1.Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab <> mPreStabs Then
      bolSave = IsSaveData '判斷是否已存檔
      mPreStabs = PreviousTab
   End If
End Sub

'Added by Lydia 2023/07/06
Private Sub txtDT_Change(Index As Integer)
  Select Case Index
     Case 0, 2, 3, 5
        PUB_RefreshText txtDt(Index)
  End Select
End Sub

Private Sub txtDT_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   txtDt(Index).ToolTipText = txtDt(Index).Text
End Sub

'Added by Lydia 2021/10/01 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtDT_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then Forms(0).PopupMenu2 txtDt(Index)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtField(Index).SelStart = 0
   txtField(Index).SelLength = Len(txtField(Index))
   'Mark by Lydia 2016/10/28 受win7輸入法影響,不切換輸入法
   'CloseIme
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
       Case 0, 1, 2, 11, 12, 18, 19, 20, 6, 22
          '11查名路徑預設大寫
          KeyAscii = UpperCase(KeyAscii)
       'Remove by Lydia 2021/10/01 拿掉
       'Case 10, 16
       Case Else
          KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub txtField_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   'Remove by Lydia 2021/10/01 txtField(16)=>textService
   'If Index = 16 Then
   '   txtField(Index).ToolTipText = txtField(Index).Text
   'End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
Dim tmpGrp As String

   Select Case Index
       Case 0 '委查人
            If txtField(Index).Text <> "" Then
               If txtField(Index).Tag <> txtField(Index).Text Then 'Added by Lydia 2019/06/25 有異動才查詢(因為遇到人員離職,但是尚未查覆完畢)
                    'Modified by Lydia 2019/06/25
                    'If ClsPDGetStaff(txtField(Index).Text, strExc(1)) Then
                    strExc(1) = GetStaffName(txtField(Index).Text, True)
                    If strExc(1) <> "" Then
                    'end 2019/06/25
                        lblAppNo(2).Caption = strExc(1)
                    Else
                        lblAppNo(2).Caption = ""
                        GoTo JumpCancel
                    End If
               End If
            Else
               lblAppNo(2).Caption = ""
               MsgBox "委查人不可空白!", vbCritical
               GoTo JumpCancel
            End If
       Case 1 '查名人
            If txtField(Index).Text <> "" Then
               If ClsPDGetStaff(txtField(Index).Text, strExc(1)) Then
                   lblAppNo(3).Caption = strExc(1)
               Else
                   lblAppNo(3).Caption = ""
                   GoTo JumpCancel
               End If
            'Modified by Lydia 2017/01/20 查名單維護才提示　+ if r_type = "A" then
            ElseIf R_type = "A" Then
               lblAppNo(3).Caption = ""
               MsgBox "查名人不可空白!", vbCritical
               GoTo JumpCancel
            End If
       Case 2 '覆核主管
            If txtField(Index).Text <> "" Then
                'Modified by Lydia 2020/05/05 葉特助:退休
                'If InStr("67002,69008", txtField(Index).Text) = 0 Then
                'Modified by Lydia 2022/05/25 改用「內商查名覆核人員」
                'If InStr("69008", txtField(Index).Text) = 0 Then
                '    MsgBox "請輸入商標處主管的員工編號!", vbCritical
                If InStr(strPreAgree, txtField(Index).Text) = 0 Then
                    MsgBox "請輸入內商查名覆核人員的員工編號!", vbCritical
                'end 2022/05/25
                    GoTo JumpCancel
                Else
                    If ClsPDGetStaff(txtField(Index).Text, strExc(1)) Then
                       lblAppNo(4).Caption = strExc(1)
                    Else
                       lblAppNo(4).Caption = ""
                       GoTo JumpCancel
                    End If
                End If
            Else
                lblAppNo(4).Caption = ""
            End If
       Case 3, 4, 5, 17, 21 '申請日期、期限日期、查覆日期、覆核日期、查覆完成日期(TQA09)
            If txtField(Index).Text <> "" Then
               If CheckIsTaiwanDate(txtField(Index).Text) = False Then
                  MsgBox "請輸入民國年月日!", vbCritical
                  GoTo JumpCancel
               ElseIf ChkWorkDay(ChangeTStringToWString(txtField(Index).Text)) = False Then
                  MsgBox "請輸入工作日!", vbCritical
                  GoTo JumpCancel
               End If
            End If
       Case 6
            If cmdSend.Tag = "A" Then
                txtField(Index).Text = Replace(txtField(Index).Text, ".", ",")
                '檢查組群
                If Check_ClassDouble(txtField(Index)) = True Then
                   GoTo JumpCancel
                End If
                '檢查申請編號的全部組群
                txt1(1).Text = Replace(mTQA03, txtField(Index).Tag, txtField(Index).Text)
                If Check_ClassDouble(txt1(1), "A") = True Then
                   GoTo JumpCancel
                End If
            End If
       Case 10 '客戶名稱
            If txtField(Index).Text = "" Then
               MsgBox "客戶名稱不可空白!", vbCritical
               GoTo JumpCancel
            End If
       Case 11 '查名路徑-取代間隔符號
            txtField(Index).Text = Replace(txtField(Index).Text, ".", ",")
            strExc(4) = PUB_RepToOneSpace(PUB_StringFilter(txtField(Index).Text))   '清除字串中的enter & 清除連續空白
       Case 13, 14, 15
            If txtField(Index) <> Empty Then
               If Len(Trim(txtField(12).Text & txtField(13).Text)) < 7 Then
                  MsgBox "請輸入本所案號!!", vbCritical
                  GoTo JumpCancel
               Else
                  txtField(14).Text = Mid(txtField(14).Text & "0", 1, 1)
                  txtField(15).Text = Mid(txtField(15).Text & "00", 1, 2)
               End If
               If ClsPDCheckCaseCodeIsExist(txtField(12).Text, txtField(13).Text, txtField(14).Text, txtField(15).Text) = False Then
                  GoTo JumpCancel
               End If
            End If
       Case 18, 22 '查覆結果是否已讀/是否撤回
            If txtField(Index).Text <> "" And txtField(Index).Text <> "Y" Then
                MsgBox "請輸入Y!", vbCritical
                GoTo JumpCancel
            End If
'       Case 19  '業務收文組群
'            If txtField(Index).Text <> "" Then
'                strSql = "select tmq01 from trademarkquery where tmq01=" & CNULL(txtField(Index).Text)
'                intI = 1
'                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                If intI = 0 Then
'                   MsgBox "請輸入正確的收文組群!", vbCritical
'                   GoTo JumpCancel
'                End If
'            End If
       Case 20  '櫃台收文號
            If txtField(Index).Text <> "" Then
                strSql = "select CP01,CP02,CP03,CP04,CP10,CP57 from CASEPROGRESS where CP09=" & CNULL(txtField(Index).Text)
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 0 Then
                   MsgBox "無此收文號!", vbCritical
                   GoTo JumpCancel
                ElseIf "" & RsTemp.Fields("CP10") <> 申請 Then
                   MsgBox "收文號案件性質不屬於申請!", vbCritical
                   GoTo JumpCancel
                ElseIf Not IsNull(RsTemp.Fields("CP57")) Then
                   MsgBox "收文號已取消收文!", vbCritical
                   GoTo JumpCancel
                End If
            End If

            '櫃台已收文的情況
            If txtField(Index).Text <> txtField(Index).Tag And txtField(Index).Tag <> "" Then
                strSql = "SELECT cpp01 FROM casepaperpdf WHERE cpp01 ='" & txtField(Index).Tag & "' and instr(upper(cpp02),'" & txtField(Index).Tag & "') > 0 "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then
                    MsgBox "櫃台已收文,若要修改請洽該案商標承辦人員,從工作進度維護的查名結果進行修改!", vbCritical
                    GoTo JumpCancel
                End If
            End If
   End Select
   Exit Sub
   
JumpCancel:
    txtField(Index).SetFocus
    Cancel = True
End Sub


Private Sub txtUnicode_GotFocus(Index As Integer)
   txtUnicode(Index).SelStart = 0
   txtUnicode(Index).SelLength = Len(txtUnicode(Index))
End Sub

Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd2, x, y, nCol, nRow
grd2.col = nCol
grd2.row = nRow
End Sub
' 更新各控制項的狀態
Private Sub grd2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   grd2.ToolTipText = ""
   If grd2.MouseRow <> 0 And grd2.MouseCol > 0 Then
      If grd2.TextMatrix(grd2.MouseRow, grd2.MouseCol) <> "" Then
         grd2.ToolTipText = grd2.TextMatrix(grd2.MouseRow, grd2.MouseCol)
      End If
   End If
End Sub

Private Sub GRD2_SelChange()
Dim TmpRow As Integer
TmpRow = grd2.MouseRow

If iStiu = 1 Then
    If R_type = "M" Then
        'Modified by Lydia 2019/12/12 + 是否出名
        'Call ChangeDetailAns(3, Cbo4, txtDT(5))
        Call ChangeDetailAns(3, Cbo4, txtDt(5), , , txtDt(7))
    Else
        Call ChangeDetailAns(2, Cbo3, txtDt(3), , txtDt(4))
        If txtDt(4).Text <> "" Then
           mPrevTM1215 = txtDt(4).Text
        ElseIf (Cbo3.Text = TMQ_近似T1 Or Cbo3.Text = TMQ_近似T2) And txtDt(4).Text = "" Then
               txtDt(4).Text = mPrevTM1215
        End If
    End If
End If
   
grd2.Visible = False
If TmpRow > 0 Then
   '上一筆資料列清除反白
   If dblPrevRow2 > 0 Then
      grd2.col = 0
      grd2.row = dblPrevRow2
      grd2.Text = ""
      For jj = 0 To grd2.Cols - 1
         grd2.col = jj
         grd2.CellBackColor = QBColor(15)
      Next jj
   End If
   '目前資料列反白
   grd2.col = 0
   grd2.row = TmpRow
   dblPrevRow2 = grd2.row
    If grd2.TextMatrix(grd2.row, 1) <> "" Then
       grd2.Text = "V"
       For jj = 0 To grd2.Cols - 1
          grd2.col = jj
          grd2.CellBackColor = &HFFC0C0
       Next jj
       'Modified by Lydia 2019/12/12 +Tab
       Call SetComboAns(1, grd2, dblPrevRow2, Cbo3, Cbo4, txtDt(3), txtDt(4), txtDt(5))
       Call ShowFieldTQD09(1)
    End If
End If
grd2.Visible = True
End Sub

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

'開啟結果附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   Dim bolIsCPP As Boolean
   bolSave = IsSaveData '判斷是否已存檔
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   strAtt = lstAtt(Index).Text
   
   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      For jj = 0 To lstAtt(Index).ListCount - 1
         If lstAtt(Index).Selected(jj) Then
            bolIsSelect = True
            stFileName = lstAtt(Index).List(jj)
            '讀取檔案名稱
            'If InStrRev(stFileName, " (") > 0 Then
            If InStr(stFileName, " (") > 0 Then
               stFileName = Left(stFileName, InStr(stFileName, " (") - 1)
            End If
            
            Call GetTQF0304(Index, mTQF03, mTQF04, stFileName)
            'Modified by Lydia 2016/06/23 改放在FTP
            '已收文-檔案歸到卷宗區
            'Modified by Lydia 2016/04/25 +TS案
            'If FirstCP(1) = "T" Or FirstCP(1) = "TS" Then
            '    stFileName = FirstCPP02t & stFileName
            '    bolIsCPP = PUB_GetAttachFile_CPP(FirstCP09, stFileName, m_AttachPath & "\" & stFileName, True)
            '    'Modified by Lydia 2016/05/25 修正檔案路徑:可能發生在上附件時,正在收文;兩者平行作業產生db有檔案,卷宗區有menu,卻沒新增到查名附件
            '    If bolIsCPP = False And InStr(stFileName, m_AttachPath) = 0 Then
            '       bolIsCPP = GetTQFtoCPP(stFileName)
            '       stFileName = m_AttachPath & "\" & stFileName
            '    End If
            '    If bolIsCPP Then
            '    'end 2016/05/25
            '       ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            '    End If
            'Else '未收文
                strExc(3) = UCase(TMQ_查名作業 & ".pdf")
                If PUB_TMQGetAFile(m_AttachPath, stFileName, mTQD01, mTQD02, mTQF03, mTQF04, strExc(3)) = False Then Exit Sub
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            'End If
            'end 2016/06/23
             
            '不限查覆完畢後,委查人開啟附件更新->已讀
    On Error GoTo ErrOpenAF
            'If mTMQ19 = "" And mbolSend = True And txtField(0).Text = strUserNum And R_type = "Q" Then
            'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
            'If mTMQ19 = "" And txtField(0).Text = strUserNum And R_type = "Q" Then
            'Modified by Lydia 2019/12/25 開放特殊設定權限
            'If mTMQ19 = "" And InStr(stIdList, txtField(0).Text) > 0 And R_type = "Q" Then
            If mTMQ19 = "" And R_type = "Q" And (InStr(stIdList, txtField(0).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, txtField(0).Text) > 0)) Then
               cnnConnection.BeginTrans
                 strSql = "UPDATE TMQFILE SET TQF11='Y' WHERE TQF01='" & mTQD01 & "' AND TQF02='" & mTQD02 & "' AND TQF03='" & mTQF03 & "' AND TQF04='" & mTQF04 & "' "
                 cnnConnection.Execute strSql, intI
                  intI = 1
                  strExc(0) = "select count(*),count(tqf11) from TMQFILE where tqf01='" & mTQD01 & "' and tqf02='" & mTQD02 & "' and TQF04<>'" & TMQ_附件F04 & "' "
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp(0) = RsTemp(1) Then
                        strSql = "UPDATE trademarkquery SET TMQ19='Y' WHERE TMQ01='" & mTQD02 & "' "
                        cnnConnection.Execute strSql, intI
                        mTMQ19 = "Y"
                        'cmdTo.Enabled = True 'Remove by Lydia 2016/04/06
                     End If
                  End If
               cnnConnection.CommitTrans
            End If
            '清除列表中的未讀提示(本次作業中的未讀,實際上由申請者開啟,則狀態不變)
             strExc(0) = lstAtt(Index).List(jj)
             lstAtt(Index).List(jj) = Replace(strExc(0), " (未讀)", "")
         End If
      Next jj
      If bolIsSelect = False Then
         MsgBox "請選擇欲開啟的附件！"
      End If
   End If
   
   Screen.MousePointer = vbDefault
ErrOpenAF:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'全選
Private Sub cmdSelect_Click(Index As Integer)
   Dim ii As Integer, oList As ListBox

   bolSave = IsSaveData '判斷是否已存檔
   
   Set oList = lstAtt(Index)
   For ii = 0 To oList.ListCount - 1
      lstAtt(Index).Selected(ii) = True
   Next
End Sub

'下載
Private Sub cmdSaveAtt_Click(Index As Integer)
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer, oList As ListBox
   Dim bolIsCPP As Boolean
   'Added by Lydia 2025/08/27
   Dim pIdx As Integer
   pIdx = -1
   'end 2025/08/27
   
   bolSave = IsSaveData '判斷是否已存檔

   Screen.MousePointer = vbHourglass
   
   Set oList = lstAtt(Index)
   
   stFileName = ""
   bMultiFile = False
   For ii = 0 To oList.ListCount - 1
      If oList.Selected(ii) Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = oList.List(ii)
            pIdx = ii 'Addded by Lydia 2025/08/27
         End If
      End If
   Next
   
   If stFileName = "" Then
      MsgBox "請選擇欲存檔的附件！"
   Else
      '多選
      If bMultiFile Then
         stFolderPath = BrowseForFolder()
         If stFolderPath <> "" Then
            For ii = 0 To oList.ListCount - 1
               If oList.Selected(ii) Then
                  stFileName = oList.List(ii)
                  If InStr(stFileName, " (") > 0 Then
                     stFileName = Left(stFileName, InStr(stFileName, " (") - 1)
                  End If
                  stFullName = stFolderPath & stFileName
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        Call GetTQF0304(Index, mTQF03, mTQF04, stFullName)
                        'Modified by Lydia 2016/06/23 改放在FTP
                        ''已收文-檔案歸到卷宗區
                        ''Modified by Lydia 2016/04/25 +TS案
                        'If FirstCP(1) = "T" Or FirstCP(1) = "TS" Then
                        '    stFileName = FirstCPP02t & stFileName
                        '    If PUB_GetAttachFile_CPP(FirstCP09, stFileName, stFullName, True) = False Then
                        '       MsgBox "無法儲存檔案[ " & stFullName & " ]！"
                        '       GoTo RunExit
                        '    End If
                        'Else '未收文
                             strExc(3) = UCase(TMQ_查名作業 & ".pdf")
                            If PUB_TMQGetAFile("", stFullName, mTQD01, mTQD02, mTQF03, mTQF04, strExc(3)) = False Then
                               MsgBox "無法儲存檔案[ " & stFullName & " ]！"
                               GoTo RunExit
                            'Added by Lydia 2025/08/27 因為智權常用下載PDF來看，所以下載=附件已讀 by 杜協理, 嘉雯
                            Else
                              If mTMQ19 = "" And R_type = "Q" And (InStr(stIdList, txtField(0).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, txtField(0).Text) > 0)) Then
                                 cnnConnection.BeginTrans
                                   strSql = "UPDATE TMQFILE SET TQF11='Y' WHERE TQF01='" & mTQD01 & "' AND TQF02='" & mTQD02 & "' AND TQF03='" & mTQF03 & "' AND TQF04='" & mTQF04 & "' "
                                   cnnConnection.Execute strSql, intI
                                    intI = 1
                                    strExc(0) = "select count(*),count(tqf11) from TMQFILE where tqf01='" & mTQD01 & "' and tqf02='" & mTQD02 & "' and TQF04<>'" & TMQ_附件F04 & "' "
                                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                                    If intI = 1 Then
                                       If RsTemp(0) = RsTemp(1) Then
                                          strSql = "UPDATE trademarkquery SET TMQ19='Y' WHERE TMQ01='" & mTQD02 & "' "
                                          cnnConnection.Execute strSql, intI
                                          mTMQ19 = "Y"
                                       End If
                                    End If
                                 cnnConnection.CommitTrans
                              End If
                              '清除列表中的未讀提示(本次作業中的未讀,實際上由申請者開啟,則狀態不變)
                               strExc(0) = lstAtt(Index).List(ii)
                               lstAtt(Index).List(ii) = Replace(strExc(0), " (未讀)", "")
                            'end 2025/08/27
                            End If
                        'End If
                        'end 2016/06/23
                     End If
                  End If
               End If
            Next
         End If
      
      Else
            If InStr(stFileName, " (") > 0 Then
                stFileName = Left(stFileName, InStr(stFileName, " (") - 1)
            End If
            
            stFullName = GetSaveName(stFileName)
            If stFullName <> "" Then
              If Dir(stFullName) <> "" Then
                If MsgBox("檔案[ " & stFullName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                   stFullName = ""
                End If
              End If
              If stFullName <> "" Then
                  'Modified by Lydia 2017/02/21 單檔下載可能會更名
                  'Call GetTQF0304(Index, mTQF03, mTQF04, stFullName)
                  Call GetTQF0304(Index, mTQF03, mTQF04, stFileName)
                  'end 2017/02/21
                  'Modified by Lydia 2016/06/23 改放在FTP
                  ''已收文-檔案歸到卷宗區
                  ''Modified by Lydia 2016/04/25 +TS案
                  'If FirstCP(1) = "T" Or FirstCP(1) = "TS" Then
                  '    stFileName = FirstCPP02t & stFileName
                  '    If PUB_GetAttachFile_CPP(FirstCP09, stFileName, stFullName, True) = False Then
                  '       MsgBox "無法儲存檔案[ " & stFullName & " ]！"
                  '       GoTo RunExit
                  '    End If
                  'Else '未收文
                       strExc(3) = UCase(TMQ_查名作業 & ".pdf")
                      If PUB_TMQGetAFile("", stFullName, mTQD01, mTQD02, mTQF03, mTQF04, strExc(3)) = False Then
                         MsgBox "無法儲存檔案[ " & stFullName & " ]！"
                         GoTo RunExit
                      'Added by Lydia 2025/08/27 因為智權常用下載PDF來看，所以下載=附件已讀 by 杜協理, 嘉雯
                      Else
                         If mTMQ19 = "" And R_type = "Q" And (InStr(stIdList, txtField(0).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, txtField(0).Text) > 0)) Then
                            cnnConnection.BeginTrans
                              strSql = "UPDATE TMQFILE SET TQF11='Y' WHERE TQF01='" & mTQD01 & "' AND TQF02='" & mTQD02 & "' AND TQF03='" & mTQF03 & "' AND TQF04='" & mTQF04 & "' "
                              cnnConnection.Execute strSql, intI
                              intI = 1
                              strExc(0) = "select count(*),count(tqf11) from TMQFILE where tqf01='" & mTQD01 & "' and tqf02='" & mTQD02 & "' and TQF04<>'" & TMQ_附件F04 & "' "
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 If RsTemp(0) = RsTemp(1) Then
                                    strSql = "UPDATE trademarkquery SET TMQ19='Y' WHERE TMQ01='" & mTQD02 & "' "
                                    cnnConnection.Execute strSql, intI
                                    mTMQ19 = "Y"
                                 End If
                              End If
                            cnnConnection.CommitTrans
                         End If
                         '清除列表中的未讀提示(本次作業中的未讀,實際上由申請者開啟,則狀態不變)
                         strExc(0) = lstAtt(Index).List(pIdx)
                         lstAtt(Index).List(pIdx) = Replace(strExc(0), " (未讀)", "")
                     'end 2025/08/27
                      End If
                  'End If
                  'end 2016/06/23
              End If
            End If
      End If
      
      If stFullName <> "" Then
         MsgBox "下載完成！"
      End If
   End If
RunExit:
   Screen.MousePointer = vbDefault
End Sub
Private Function BrowseForFolder(Optional sCaption As String = "請選擇欲儲存的位置", Optional sDefault As String) As String
    Const BIF_RETURNONLYFSDIRS = 1
    Const MAX_PATH = 260
    Dim lPos As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, tBrowse As BrowseInfo

    With tBrowse
        'Set the owner window
        .hwndOwner = GetActiveWindow        'Me.hWnd in VB
        .lpszTitle = sCaption
        .ulFlags = BIF_RETURNONLYFSDIRS     'Return only if the user selected a directory
    End With

    'Show the dialog
    lpIDList = SHBrowseForFolder(tBrowse)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        lPos = InStr(sPath, vbNullChar)
        If lPos Then
            BrowseForFolder = Left$(sPath, lPos - 1)
            If Right$(BrowseForFolder, 1) <> "\" Then
                BrowseForFolder = BrowseForFolder & "\"
            End If
        End If
    Else
        'User cancelled, return default path
        BrowseForFolder = sDefault
    End If
End Function

'取得類別和流水號
Private Sub GetTQF0304(ByRef inD As Integer, oTQF03 As String, Optional oTQF04 As String, Optional sFname As String)
Dim tmpInx As Integer
    If sFname <> "" Then
      If InStr(UCase(sFname), "." & UCase(TMQ_查名作業 & ".pdf")) > 0 Then oTQF04 = Mid(sFname, InStr(UCase(sFname), "." & UCase(TMQ_查名作業 & ".pdf")) - 2, 2)
    Else
      oTQF04 = addRow(inD)
    End If
    oTQF03 = mTQD03
    tmpInx = Val(oTQF04)
    If inD = 1 Then
       oTQF03 = TMQ_AkindWord2
    End If
End Sub
'新增
Private Sub cmdAddAtt_Click(Index As Integer)
Dim stFileName As String
Dim sFile
Dim ii As Integer
Dim fs, f, s
Dim strFile As String
Dim intS As Integer
Dim inX As Integer
Dim StrStr1 As String
Dim bolUpd As Boolean 'Added by Lydia 2016/05/26

    bolSave = IsSaveData '判斷是否已存檔
    Call CheckTMQ21isExists 'Added by Lydia 2016/05/25 判斷是否已收文
    
    '判斷附件數量
    Call GetTQF0304(Index, mTQF03, mTQF04)
    strExc(3) = UCase(TMQ_查名作業 & ".pdf")

    intS = Val(mTQF04) + 1
    If intS > fileMax Then
        MsgBox "附件的數量不可超過" & fileMax & "個!", vbCritical
        Exit Sub
    End If
On Error GoTo ErrHnd
   stFileName = "*.PDF"
   bolUpd = False
   
    With CommonDialog1
       .CancelError = True
       .FileName = stFileName
       .Filter = "All Files (*.PDF)|*.PDF"
       'Modified by Lydia 2016/05/26
       '.InitDir = strLoadPath
        If GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "") <> "" Then
           .InitDir = GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "")
        Else
           .InitDir = PUB_Getdesktop
        End If
       .MaxFileSize = 3000
       '允許多選
       .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
       .ShowOpen
      If .FileName <> "" Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            'Modified by Lydia 2016/05/26 記錄路徑只到資料夾位置
            'SaveSetting "TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", sFile(0)
            'strLoadPath = sFile(0)
            SaveSetting "TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", Left(sFile(0), InStrRev(sFile(0), "\") - 1)
            '多選
            If UBound(sFile) > 1 Then
               inX = 1
            '單選
            Else
               inX = 0
            End If
            For ii = inX To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                  Exit Sub
               End If
               If UCase(Right(CStr(sFile(ii)), 4)) <> UCase(".pdf") Then
                  MsgBox "只能新增PDF檔！"
                  Exit Sub
               End If
               If intS > fileMax Then
                  MsgBox "附件的數量不可超過" & fileMax & "個!", vbCritical
                  Exit Sub
               End If
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  Exit Sub
               'Modified by Lydia 2024/12/18 從5MB放大到10MB
               ElseIf f.Size > 5242880 * 2 Then
                     If MsgBox("檔案過大（容量超過10MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                        Exit Sub
                     End If
               End If
               '允許多選檔案
               If AddListX(lstAtt(Index), mTQD01, mTQD02, mTQF03, Format(intS, TMQ_附件F04), strExc(3), FirstCP09, Format(f.Size, "0")) = True Then
                  If PUB_TMQAFileSave(mTQD01, mTQD02, mTQF03, Format(intS, TMQ_附件F04), UCase(TMQ_查名作業 & ".pdf"), stFileName, "N") = True Then
                    'Modified by Lydia 2016/04/06 補已收文資料
                    'Remove by Lydia 2016/07/07
                    '   If FirstCP09 <> "" Then
                    '      strExc(1) = "": strExc(2) = ""
                    '      StrStr1 = "select cpp01,cpp02 from casepaperpdf where cpp01='" & FirstCP09 & "' and instr(upper(cpp02),upper('" & mTQD02 & "." & TMQ_查名作業 & ".menu')) > 0 "
                    '      ii = 1
                    '      Set RsTemp = ClsLawReadRstMsg(ii, StrStr1)
                    '      If ii = 0 Then
                    '         StrStr1 = FirstCPP02t & mTQD02 & "." & TMQ_查名作業 & ".menu"
                    '         '新增TS.menu 至卷宗區
                    '         strExc(1) = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
                    '                    " values('" & FirstCP09 & "','" & StrStr1 & "',0,'" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",'Y')"
                    '         'Added by Lydia 2016/06/23
                    '         If strSrvDate(1) >= TMQFileFTP Then
                    '             cnnConnection.Execute strExc(1)
                    '         End If
                    '      End If
                          'Modified by Lydia 2016/06/23 改放在FTP
                          'If strSrvDate(1) < TMQFileFTP Then
                          '    '結果附件存至卷宗區
                          '    StrStr1 = FirstCPP02t & mTQD02 & mTQF03 & Format(intS, TMQ_附件F04) & "." & UCase(TMQ_查名作業 & ".pdf")
                          '    If SaveAttFile_PDF(FirstCP09, stFileName, StrStr1, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False, "A") = False Then
                          '        Exit Sub
                          '    End If
                          '    bolUpd = True 'Added by Lydia 2016/05/26
                          '    '移除DB資料,寫入TQF12=CPP02
                          '    strExc(2) = "update tmqfile set tqf07=null,tqf12='" & StrStr1 & "' where tqf01='" & mTQD01 & "' and tqf02='" & mTQD02 & "' and tqf03='" & mTQF03 & "' and tqf04='" & Format(intS, TMQ_附件F04) & "' "
                          '    cnnConnection.BeginTrans
                          '       If strExc(1) <> "" Then cnnConnection.Execute strExc(1)
                          '       If strExc(2) <> "" Then cnnConnection.Execute strExc(2)
                          '    cnnConnection.CommitTrans
                          'End If
                     '  End If
                     'end Remove by Lydia 2016/07/07
                     
                     addRow(Index) = addRow(Index) + 1
                     
                     If Pub_StrUserSt03 <> "M51" Then
                        'Modified by Lydia 2016/04/26 桂紹禎反應無法刪檔
                        'Call PUB_DelPCOrgFile(stFileName) '一併將PC上的實體檔案刪除
                        SetAttr stFileName, vbNormal '改檔案性質為一般
                        Kill stFileName
                     End If
                     
                     If mbolSend = True And R_type <> "A" Then
                         bolModify = True
                         'Modified by Lydia 2016/04/21 +已收文案件提示
                         'Modified by Lydia 2021/10/01 txtField(10) => textCName
                         strMod(0) = "「" & textCName.Text & "」" & GetStrTitle & " 委查單: " & lblAppNo(1).Caption & _
                                  IIf(FirstCP09 <> "", "，已收文案件" & FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) <> "000", "-" & FirstCP(3) & "-" & FirstCP(4), ""), "") & _
                                  "，查名結果有變更!!"
                         strMod(2) = IIf(strMod(2) = "", txtField(0).Text, strMod(2))
                         strMod(1) = strMod(1) & SSTab1.TabCaption(Index) & " 結果附件 " & mTQD02 & mTQF03 & Format(intS, TMQ_附件F04) & UCase(TMQ_查名作業 & ".pdf") & "有所變更" & " ;" & vbCrLf
                     End If
                     
                     If addRow(Index) > intS Then addRow(Index) = intS
                     intS = intS + 1
                  End If
               End If
            Next
      'Added by Lydia 2016/05/26
      Else
          Exit Sub
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
      'Added by Lydia 2016/05/26
      If bolUpd Then
         cnnConnection.RollbackTrans
      End If
   End If

End Sub

'刪除
Private Sub cmdRemAtt_Click(Index As Integer)
  Dim tmpInx As Integer
  bolSave = IsSaveData '判斷是否已存檔
  Call CheckTMQ21isExists 'Added by Lydia 2016/05/25 判斷是否已收文
  
   If AttachFileDel(lstAtt(Index), Index) = True Then
   End If
End Sub

'刪除已下載的結果附件檔
Private Sub AttachFileKill(Optional fName As String)
Dim mStr As String
On Error Resume Next
    mStr = Dir(m_AttachPath & "\" & fName)
      If mStr <> "" Then
         If PUB_ChkFileOpening(m_AttachPath & "\" & fName) = True Then
            MsgBox m_AttachPath & "\" & fName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
            Exit Sub
         End If
         Kill m_AttachPath & "\" & fName
      End If
End Sub

Private Sub AttachFileRead(inX As Integer, iL01 As String, iL02 As String, iL03 As String, Optional iL04 As Integer)
Dim tmpTit As String

   lstAtt(inX).Clear
   addRow(inX) = 0
   strExc(0) = "select TQF01,TQF02,TQF03,TQF04,TQF05,TQF06,TQF11 from TMQFile where TQF01='" & iL01 & "' and TQF02='" & iL02 & "' and TQF03='" & iL03 & "' order by TQF04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         jj = 0
         Do While Not .EOF
            'AttachFileKill GetAFName(iL02, iL03, Format(.Fields("TQF04"), TMQ_附件F04), UCase(TMQ_查名作業 & ".pdf"), FirstCP09)
            'lstAtt(inX).AddItem GetAFName(.Fields("TQF02"), .Fields("TQF03"), Format(.Fields("TQF04"), TMQ_附件F04), .Fields("TQF05"), FirstCP09, "" & .Fields("TQF06"), "" & .Fields("TQF11")), jj
            tmpTit = GetAFName(iL02, iL03, Format(.Fields("TQF04"), TMQ_附件F04), UCase(TMQ_查名作業 & ".pdf"), FirstCP09, "" & .Fields("TQF06"), "" & .Fields("TQF11"))
            AttachFileKill tmpTit
            lstAtt(inX).AddItem tmpTit, jj
            lstAtt(inX).ItemData(0) = 0
            addRow(inX) = Val(.Fields("TQF04"))
            .MoveNext
            jj = jj + 1
         Loop
      End With
      
      Me.cmdOpenAtt(inX).Enabled = True
      Me.cmdSelect(inX).Enabled = True
      Me.cmdSaveAtt(inX).Enabled = True
   End If
   If lstAtt(inX).ListCount > 0 Then SetListScroll lstAtt(inX)
    '申請的所有委查單尚未輸入結果，可自請撤回申請
   If R_type = "Q" And cmdSend.Caption = "撤　回" And iStiu = 1 Then
       If addRow(inX) > 0 Then cmdSend.Visible = False
   End If
End Sub
'刪除結果附件
Private Function AttachFileDel(oList As ListBox, Index As Integer) As Boolean
Dim iX As Integer
Dim rsA1 As New ADODB.Recordset
Dim iR As Integer
Dim strR As String

On Error Resume Next
   If oList.ListCount > 0 Then
      iX = 0
      Do While iX < oList.ListCount
         If oList.Selected(iX) = True Then
            strExc(1) = Trim(oList.List(iX))
            'If InStrRev(strExc(1), " (") > 0 Then
            If InStr(strExc(1), " (") > 0 Then
               strExc(1) = Left(strExc(1), InStr(strExc(1), " (") - 1)
            End If
            strExc(2) = strExc(1)
            If MsgBox("確定要刪除" & strExc(1) & "？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Function
                                 
            Call GetTQF0304(Index, mTQF03, mTQF04, strExc(1))
            '直接從資料庫刪除檔案
            'Modified by Lydia 2016/06/23 改放在FTP
            'Remove by Lydia 2016/07/07
            'If strSrvDate(1) < TMQFileFTP Then
            '    cnnConnection.BeginTrans
            '      strSql = "delete from TMQFILE where TQF01='" & mTQD01 & "' AND TQF02='" & mTQD02 & "' AND TQF03='" & mTQF03 & "' AND TQF04='" & mTQF04 & "' "
            '      cnnConnection.Execute strSql, jj
            '    cnnConnection.CommitTrans
            '    '已收文,異動附件(檔案歸到卷宗區)
            '    'Modified by Lydia 2016/04/25 +TS案
            '    'If FirstCP(1) = "T" And FirstCP09 <> "" Then
            '    If (FirstCP(1) = "T" Or FirstCP(1) = "TS") And FirstCP09 <> "" Then
            '        strExc(2) = FirstCPP02t & strExc(2)
            '        If DelAttFile_PDF(FirstCP(1) & "-" & FirstCP(2) & "-" & FirstCP(3) & "-" & FirstCP(4), FirstCP09, strExc(2)) = False Then
            '        End If
            '    End If
            'Else
                '刪除FTP檔案
                cnnConnection.BeginTrans
                    If PUB_TMQAFileDel(mTQD01, mTQD02, mTQF03, mTQF04) Then
                        jj = Val(mTQF04)
                    End If
                cnnConnection.CommitTrans
            'End If
            ''end 2016/06/23
            
            If Val(mTQF04) >= addRow(Index) Then
               addRow(Index) = addRow(Index) - 1
            End If
            If jj > 0 Then
               oList.RemoveItem iX
               SetListScroll oList
               AttachFileDel = True
               iX = iX - 1
            End If
            'end 2016/06/23
            
         End If
         iX = iX + 1
      Loop
   End If

End Function
Private Function AddListX(oList As ListBox, iFD01 As String, iFD02 As String, iFD03 As String, iFD04 As String, iFD05 As String, Optional iFD06 As String, Optional iLen As String, Optional iFD11 As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
      
    If oList.ListCount > 0 Then
        For idx = 0 To oList.ListCount - 1
           stFileName = oList.List(idx)
           If stFileName <> "" Then
               If iFD04 = Mid(stFileName, InStr(UCase(stFileName), "." & UCase(TMQ_查名作業 & ".pdf")) - 2, 2) Then
                   MsgBox "附件 " & stFileName & " 已存在！"
                   AddListX = False
                   bFound = True
                   Exit For
               End If
           End If
        Next
    End If
    
    strExc(6) = GetAFName(iFD02, iFD03, iFD04, iFD05, iFD06, iLen, iFD11)
    idx = oList.ListCount
    If bFound = False And strExc(6) <> "" Then
       oList.AddItem strExc(6), idx
       SetListScroll oList
       AddListX = True
    End If

End Function

Private Function GetAFName(iR01 As String, iR02 As String, iR03 As String, Optional iR04 As String, Optional iR05 As String, Optional iLen As String, Optional iR11 As String) As String
'iR01:委查單號, iR02:類別, iR03:流水號,  iR04:檔案類型, iR05:收文號 , iR11:已讀
    If iR01 <> "" And iR02 <> "" Then
       GetAFName = iR01 & iR02 & iR03 & "." & UCase(iR04)
       If iLen <> "" Then GetAFName = GetAFName & " (" & Round(iLen / 1024, 2) & " KB)" & IIf(iR11 <> "Y", " (未讀)", "")
    End If
End Function
Private Function GetSaveName(ByVal pFileName As String) As String
Dim sFile

On Error GoTo ErrHnd
         
   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      'Modified by Lydia 2016/05/26
      '.InitDir = strLoadPath
      If GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
   End With
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        sFile = Split(CommonDialog1.FileName, ChrW$(0))
        '記錄路徑
        'Modified by Lydia 2016/05/26 記錄路徑只到資料夾位置
        'SaveSetting "TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", sFile(0)
        'strLoadPath = sFile(0)
        SaveSetting "TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", Left(sFile(0), InStrRev(sFile(0), "\") - 1)
        GetSaveName = CommonDialog1.FileName
    End If
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim tmpS  '存Unicode
Dim mLoad As Boolean
Dim sPath As String
Dim APKind As String
Dim oRunform As Form 'Add By Sindy 2022/9/16
   
   'Add By Sindy 2022/9/16
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      Set oRunform = frm090801_New
   Else
      Set oRunform = frm090801
   End If
   '2022/9/16 END
   
If m_TMQApp <> "" Then
'    If mTMQ20 <> "" Then
'       cmdTo.Enabled = False
'    End If
   'Added by Lydia 2016/05/10 從接洽單回來
   If mbolCall = True Then
      Unload Me
   End If
Else
    Me.Enabled = False: mLoad = False
    Screen.MousePointer = vbHourglass
    APKind = mTQD01 '先代入申請號
    'Modify By Sindy 2022/9/16 frm090801 改用 oRunform
    oRunform.bolExternalCall = True '記錄是外部程式呼叫使用
    oRunform.SetParent Me
    oRunform.Show
    oRunform.Option1(0).Value = True '新案
    oRunform.Text1(6) = "T" '商標案
    Call oRunform.Text1_LostFocus(9)

    If mTQD03 = 0 Then '圖形
       If InStr(UCase(cmdKey(0).Tag), "PDF") = 0 Then
          mLoad = True
          'Modified by Lydia 2016/06/23
          'APKind = APKind & "_0"
          APKind = APKind & TMQ_附件F02 & TMQ_AkindPic & TMQ_附件F04
       End If
    Else
       '文字
       If cmdKey(0).Tag = "" And txtUnicode(1).Text <> "" Then tmpS = tmpS & Trim(txtUnicode(1).Text) & " "
       If cmdKey(1).Tag = "" And txtUnicode(2).Text <> "" Then tmpS = tmpS & Trim(txtUnicode(2).Text) & " "
       
       If InStr(UCase(cmdKey(0).Tag), "PDF") = 0 And cmdKey(0).Tag <> "" Then
             mLoad = True
             'Modified by Lydia 2016/06/23
             'APKind = APKind & "_1"
             APKind = APKind & TMQ_附件F02 & TMQ_AkindWord1 & TMQ_附件F04
       ElseIf InStr(UCase(cmdKey(1).Tag), "PDF") = 0 And cmdKey(1).Tag <> "" Then
             mLoad = True
             'Modified by Lydia 2016/06/23
             'APKind = APKind & "_2"
             APKind = APKind & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04
       End If
    End If
    m_TMQApp = mTQD01

    If tmpS <> "" Then
      oRunform.opt1(0).Value = True
      'oRunform.PicText = tmpS 'Mark by Lydia 2024/10/07 商標文字欄位中，勿直接帶入文字，以留空方式讓智權人員填寫---杜協理
    ElseIf mLoad = True Then
      sPath = Dir(m_AttachPath & "\" & APKind & "*.*")
      If sPath = "" Then
         'Modified by Lydia 2016/06/23
         'mLoad = KeyFileGet(mTQD01, Right(APKind, 1), False, sPath)
         mLoad = KeyFileGet(mTQD01, Right(APKind, 1), False, sPath, cmdKey(0).Tag)
      Else
         sPath = m_AttachPath & "\" & sPath
      End If
      If mLoad = True Then
         oRunform.opt1(1).Value = True
         oRunform.optColor(0).Value = True
         Call oRunform.PicToObj(sPath)
      End If
    End If
    oRunform.cmdTMQ.Tag = mTQD01
    oRunform.Combo1(0).Text = "000" & " " & GetPrjNationName("000")
    '設定案件性質
    Call oRunform.Text1_LostFocus(6)
    Call oRunform.QueryTMQ
    'Added by Lydia 2016/07/12 TS案無商標種類
    If oRunform.Text1(6) = "TS" Then
    
    ElseIf oRunform.Text1(6) = "T" Then
       oRunform.Combo6.ListIndex = 0 'Added by Lydia 2016/05/30 接洽單的商標種類
    End If
    
    oRunform.bolExternalCall = False '還原預設值
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Me.Hide
End If
End Sub
'顯示/隱藏覆核欄位
Private Sub ShowFieldTQD09(sSID As Integer)
Dim bolEna As Boolean

    bolEna = False
    Select Case sSID
        Case 0
            If (Cbo1.Text = TMQ_近似T1 Or Cbo1.Text = TMQ_近似T2) Then
               bolEna = True
            ElseIf R_type = "U" And iStiu = 1 Then
               txtDt(1).Text = "" '預設-清空
            End If
            'Modified by Lydia 2023/07/06
            'lbl3(0).Visible = bolEna: lbl3(1).Visible = bolEna
            lbl3(1).Visible = bolEna
            txtDt(1).Visible = bolEna
            FR12.Visible = bolEna
        Case 1
            If (Cbo3.Text = TMQ_近似T1 Or Cbo3.Text = TMQ_近似T2) Then
                bolEna = True
            ElseIf R_type = "U" And iStiu = 1 Then
               txtDt(4).Text = "" '預設-清空
            End If
            'Modified by Lydia 2023/07/06
            'lbl3(2).Visible = bolEna: lbl3(3).Visible = bolEna
            lbl3(2).Visible = bolEna
            txtDt(4).Visible = bolEna
            FR22.Visible = bolEna
    End Select
    If R_type <> "M" Then
       cmdSaveD(1).Visible = False
       cmdSaveD(3).Visible = False
    Else
       cmdSaveD(1).Visible = True
       cmdSaveD(3).Visible = True
    End If
End Sub
'寫入Unicode文字(暫存本機二進位檔)
Private Sub UnicodeSave()
Dim btHead(1) As Byte
Dim btTemp() As Byte
Dim p As String

    '先刪檔
    If Dir(m_AttachPath & "\unicode1.txt") <> "" Then Kill m_AttachPath & "\unicode1.txt"
    If Dir(m_AttachPath & "\unicode2.txt") <> "" Then Kill m_AttachPath & "\unicode2.txt"

    
    btHead(0) = 255
    btHead(1) = 254
    
    If txtUnicode(1).Text <> "" Then
        btTemp = txtUnicode(1).Text
        Open m_AttachPath & "\unicode1.txt" For Binary As #1
        Put #1, , btHead
        Put #1, , btTemp
        Close #1
    End If
    
    If txtUnicode(2).Text <> "" Then
        btTemp = txtUnicode(2).Text
        Open m_AttachPath & "\unicode2.txt" For Binary As #2
        Put #2, , btHead
        Put #2, , btTemp
        Close #2
    End If
End Sub

'檢查委查組群是否有重覆
Private Function Check_ClassDouble(ByRef textGrp As TextBox, Optional aKind As String) As Boolean
Dim StrArray As Variant
Dim i As Integer
Dim j As Integer
Dim strGrp As String
Dim rsMe As New ADODB.Recordset

   StrArray = ""
   If Len(textGrp) <> 0 Then
      StrArray = Split(textGrp, ",")
      strGrp = "-"
      For i = 0 To UBound(StrArray)
         If (Len(StrArray(i)) < 1 Or Len(StrArray(i)) > 4) Or IsNumeric(StrArray(i)) = False Then
            Check_ClassDouble = True
            MsgBox IIf(aKind = "A", "申請編號的", "") & "委查組群格式輸入錯誤!!!", vbCritical
            Exit Function
         End If
         For j = i + 1 To UBound(StrArray)
            If StrArray(i) = StrArray(j) Then
               Check_ClassDouble = True
               MsgBox IIf(aKind = "A", "申請編號的", "") & "委查組群重覆輸入" & StrArray(i) & "，請查明再輸!", vbCritical
               Exit Function
            End If
         Next j
         If strGrp = "-" Then
            strGrp = Mid(StrArray(i), 1, 2)
         End If
         If mTQF03 > TMQ_AkindPic And strGrp <> Mid(StrArray(i), 1, 2) Then
               Check_ClassDouble = True
               MsgBox IIf(aKind = "A", "申請編號的", "") & "委查組群必須同一類，請查明再輸!", vbCritical
               Exit Function
         End If
         '檢查不可存在於組群刪除資料檔
         If aKind = "" Then
            If rsMe.State <> adStateClosed Then rsMe.Close
            Set rsMe = Nothing
            rsMe.CursorLocation = adUseClient
            rsMe.Open "Select * From ClassDelete Where CD01='" & StrArray(i) & "'", cnnConnection, adOpenStatic, adLockReadOnly
            If rsMe.RecordCount > 0 Then
               Check_ClassDouble = True
               MsgBox StrArray(i) & "為已刪除的組群，輸入錯誤!!!", vbExclamation + vbOKOnly
               Exit Function
            End If
            If rsMe.State <> adStateClosed Then rsMe.Close
            Set rsMe = Nothing
         End If
      Next i
      If aKind = "" Then
            If UBound(StrArray) = 0 Then
               If (Len(StrArray(0)) < 1 Or Len(StrArray(0)) > 4) Or IsNumeric(StrArray(0)) = False Then
                  Check_ClassDouble = True
                  MsgBox "委查組群格式輸入錯誤!!!", vbCritical
                  Exit Function
               End If
            End If
      End If
   End If
End Function
'變更組群
Private Function UPD_ClassDetail(ByRef textGrp As TextBox) As Boolean
Dim StrArray As Variant
Dim i As Integer
Dim strUpd As String

    StrArray = Split(textGrp, ",")
    For i = 0 To UBound(StrArray)
       If StrArray(i) <> "" Then
          strUpd = "update tmqdetail set TQD05='" & StrArray(i) & "' where tqd01='" & mTQD01 & "' and tqd02='" & mTQD02 & "' and tqd04='" & i + 1 & "'"
          cnnConnection.Execute strUpd, intI
       End If
    Next i
    
    UPD_ClassDetail = True
    Exit Function

End Function

Private Sub txtDt_GotFocus(Index As Integer)
   txtDt(Index).SelStart = 0
   txtDt(Index).SelLength = Len(txtDt(Index))
   'Mark by Lydia 2016/10/28 受win7輸入法影響,不切換輸入法
   'If Index = 1 Or Index = 4 Then
   '   CloseIme
   'End If
End Sub

'Modified by Lydia 2021/10/01 改成Form 2.0
'Private Sub txtDT_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtDT_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
       Case 1, 4
          KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txtDt_Validate(Index As Integer, Cancel As Boolean)
Dim tmpGrp As String
Dim tmpSales As String
   
   txtDt(Index).Text = PUB_RepToOneSpace(PUB_StringFilter(txtDt(Index).Text))    '清除字串中的enter & 清除連續空白
   Select Case Index
       Case 1, 4
            txtDt(Index).Text = Replace(txtDt(Index).Text, ".", ",")
            strExc(4) = txtDt(Index).Text
            If strExc(4) <> "" Then
                tmpArr = Split(strExc(4), ",")
                tmpGrp = tmpArr(0)
    
                '檢查審定號/申請號
                For jj = 0 To UBound(tmpArr)
                  'Modified by Lydia 2017/08/28 查名若發現與統一公司商標近似情形時，仍列為客戶間利益衝突案件，於審定號/申請號前加上P
                  'Memo by Lydia 2022/07/12 排除P開頭: 因為近似商標權人若為統一企業，即使代理人非為本所，也視為本所代理案件。
                  If Trim(tmpArr(jj)) <> "" And Left(Trim(tmpArr(jj)), 1) <> "P" Then 'Added by Lydia 2017/03/15
                    intI = 1
                    'Modified by Lydia 2016/05/05 閉卷有可能復活,拿掉tm29
                    'Modified by Lydia 2017/03/15 閉卷改成彈訊息和增加意見(備註)
                    'strExc(0) = "select tm01,tm02,tm03,tm04 from trademark where tm10='000' and tm57||tm73 is null and (tm12='" & tmpArr(jj) & "' or tm15='" & tmpArr(jj) & "') "
                    strExc(0) = "select tm01,tm02,tm03,tm04,tm29,tm57 from trademark where tm10='000' and (tm12='" & Trim(tmpArr(jj)) & "' or tm15='" & Trim(tmpArr(jj)) & "') "
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 0 Then
                       'Modified by Lydia 2017/08/28
                       'MsgBox tmpArr(jj) & " 查無相關本所案號，請確認!!!", vbExclamation + vbOKOnly
                       MsgBox tmpArr(jj) & " 查無相關本所案號，請確認!!!" & vbCrLf & "統一公司案件，請在審定號/申請號數前加上P", vbExclamation + vbOKOnly
                       'txtDT(Index).SetFocus  'Mark by Lydia 2022/07/12 會造成無法跳出
                       Cancel = True
                    Else
                       'Remove by Lydia 2016/06/01 不通知
                       ''抓相關案的智權人員
                       'If bolModify Then
                       '   tmpSales = PUB_GetFCTSalesNo(RsTemp.Fields("tm01"), RsTemp.Fields("tm02"), RsTemp.Fields("tm03"), RsTemp.Fields("tm04"))
                       '   If tmpSales <> "" And InStr(strMod(2), tmpSales) = 0 Then
                       '      strMod(2) = strMod(2) & ";" & tmpSales
                       '   End If
                       'End If
                       'Added by Lydia 2017/03/15 閉卷改成彈訊息和增加意見(備註)
                       If Trim("" & RsTemp.Fields("tm57")) <> "" Or Trim("" & RsTemp.Fields("tm29")) <> "" Then
                           If InStr(txtDt(Index - 1), tmpArr(jj)) = 0 Then
                              strExc(1) = "審定號/申請號:" & tmpArr(jj) & IIf(Trim("" & RsTemp.Fields("tm57")) <> "", " 已銷卷", " 已閉卷")
                              MsgBox strExc(1) & "!", vbExclamation + vbOKOnly
                              txtDt(Index - 1).Text = txtDt(Index - 1).Text & IIf(txtDt(Index - 1).Text <> "", ";", "") & strExc(1)
                           End If
                       End If
                       'end 2017/03/15
                       
                    End If
                  End If  'end by Lydia 2017/03/15
                Next jj
            End If
   End Select
   'Added by Lydia 2016/04/22 檢查長度
   If Not CheckLengthIsOK(txtDt(Index), txtDt(Index).MaxLength) Then
      txtDt(Index).SetFocus
      Cancel = True
   End If
End Sub

'Added by Lydia 2016/05/25 檢查是否已收文
'因為查名人進入畫面時尚未收文,當要上傳附件時可能要檢查
Private Sub CheckTMQ21isExists()
Dim rsRd As New ADODB.Recordset
Dim strRead As String
Dim intB As Integer

   If FirstCP09 = "" Then
      strRead = "SELECT TMQ01,TMQ21,CP01,CP02,CP03,CP04,CP10,CP27,CP14,CP57 FROM TRADEMARKQUERY,CASEPROGRESS WHERE TMQ01='" & mTQD02 & "' AND TMQ21=CP09(+) "
      intB = 1
      Set rsRd = ClsLawReadRstMsg(intB, strRead)
      If intB = 1 Then
         If "" & rsRd.Fields("TMQ21") <> "" Then
             txtField(23).Text = ChangeWStringToTString("" & rsRd.Fields("CP27"))
             '櫃台收文號
             FirstCP09 = "" & rsRd.Fields("TMQ21")
             FirstCP(1) = "" & rsRd.Fields("CP01")
             FirstCP(2) = "" & rsRd.Fields("CP02")
             FirstCP(3) = "" & rsRd.Fields("CP03")
             FirstCP(4) = "" & rsRd.Fields("CP04")
             txtField(20).Text = FirstCP09
             If FirstCP09 <> "" Then
                FirstCPP02t = Trim(FirstCP(1)) & CStr(Val(FirstCP(2))) & IIf(FirstCP(3) <> "0" Or FirstCP(4) <> "00", "-" & FirstCP(3), "") & IIf(FirstCP(4) <> "00", "-" & FirstCP(4), "")
                FirstCPP02t = FirstCPP02t & "." & rsRd.Fields("CP10") & "."
             End If
             FirstCP14 = "" & rsRd.Fields("CP14")
             FirstCP57 = "" & rsRd.Fields("CP57")
             '目前進度的總收文號
             lblAppNo(5).Caption = FirstCP09
             Label1(0).Visible = True
         End If
      End If
   End If
End Sub

'Added by Lydia 2017/08/31 歸入案號
Private Sub cmdChange_Click()
Dim strCase As String, strList As String
Dim strCP09 As String
Dim strMid As String
Dim rsA2 As New ADODB.Recordset
Dim intJ As Integer
Dim tmpB As Boolean
   
   If txtChange(0) = "" Then
      MsgBox "請輸入委查單號(多筆請用逗號區隔)　!!"
      txtChange(0).SetFocus
   Else
      For intJ = 1 To 4
          txtChange_Validate intJ, tmpB
          If tmpB = True Then
              Exit Sub
          End If
      Next intJ
      
      txtChange(0) = Trim(PUB_RepToOneSpace(PUB_StringFilter(txtChange(0))))
      'Modified by Lydia 2021/11/19 增加737智財協作之T案
      'strSql = "select cp09,cp10,cp159 from caseprogress " & _
               " where cp01='" & txtChange(1) & "' and cp02='" & txtChange(2) & "' and cp03='" & txtChange(3) & "' and cp04='" & txtChange(4) & "'" & _
               " and cp10='" & IIf(txtChange(1) = "T", TMQ_T案, TMQ_TS案) & "' group by cp09,cp10,cp159 "
      strSql = "select cp09,cp10,cp159 from caseprogress " & _
               " where cp01='" & txtChange(1) & "' and cp02='" & txtChange(2) & "' and cp03='" & txtChange(3) & "' and cp04='" & txtChange(4) & "'" & _
               " and instr('" & IIf(txtChange(1) = "T", TMQ_T案, TMQ_TS案) & "', cp10) > 0 group by cp09,cp10,cp159 "
      intJ = 0
      Set rsA2 = ClsLawReadRstMsg(intJ, strSql)
      
      If intJ = 0 Then
         MsgBox "本所案號無" & IIf(txtChange(1) = "T", "申請" & TMQ_T案, "查名" & TMQ_TS案) & "的案件進度 !!"
         GoTo Exit001
      Else
         If Val(rsA2.Fields("cp159")) > 0 Then
            If MsgBox("收文號已取消收文，確定繼續歸入?", vbYesNo + vbDefaultButton2) = vbNo Then
               GoTo Exit001
            End If
         End If
           
            strCase = txtChange(1) & "-" & txtChange(2) & "-" & txtChange(3) & "-" & txtChange(4)
            strCP09 = "" & rsA2.Fields("cp09")
            strMid = GetAddStr(txtChange(0))
            
            strSql = "select tmq01,tqc01 from trademarkquery, tmqcasemap " & _
                     "where tmq01 in (" & strMid & ") and tqc02(+)='" & strCP09 & "' and tmq01=tqc03(+) order by tqc01 "
            intJ = 1
            Set rsA2 = ClsLawReadRstMsg(intJ, strSql)
            If intJ = 1 Then
               With rsA2
                  .MoveFirst
                  Do While Not .EOF
                     '沒有對照檔，才需要新增
                     If "" & .Fields("tqc01") = "" Then
                        If strList = "" Then
                           strList = "+" & .Fields("tmq01") & ","
                        ElseIf InStr(strList, .Fields("tmq01")) = 0 Then
                           strList = strList & "+" & .Fields("tmq01") & ","
                        End If
                     End If
                     .MoveNext
                  Loop
               End With
               
               If strList <> "" Then
                  If MsgBox("委查單號：" & Replace(strList, "+", "") & "尚未歸入到本所案號：" & strCase & "的卷宗區，是否要歸入？", vbYesNo + vbDefaultButton1) = vbNo Then
                     GoTo Exit001
                  End If
                  'Modified by Lydia 2024/03/14 +True
                  If PUB_TMQtoCP(True, "", strCP09, strList, "", , True) Then
                     MsgBox "委查單號：" & Replace(strList, "+", "") & "已歸入到收文號：" & strCP09 & "(本所案號：" & strCase & ")　!!"
                  End If
               Else
                  MsgBox "無委查單可歸入到入到收文號：" & strCP09 & "(本所案號：" & strCase & ")，" & vbCrLf & "請確認輸入的資料是否已經歸入收文號的卷宗區！"
               End If
            End If
      End If
   End If
      
   Exit Sub
Exit001:
   Set rsA2 = Nothing
   
End Sub

Private Sub txtChange_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtChange_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
       Case 1
          If txtChange(Index) <> "T" And txtChange(Index) <> "TS" Then
             MsgBox "請輸入T案或TS案 !!"
             txtChange(Index).SetFocus
             Cancel = True
          End If
       Case 2
          If Len(txtChange(Index)) <> 6 Then
             MsgBox "請輸入完整案號 !!"
             txtChange(Index).SetFocus
             Cancel = True
          Else
             txtChange(3) = "0": txtChange(4) = "00"
          End If
   End Select
End Sub

Private Sub txtChange_GotFocus(Index As Integer)
   TextInverse txtChange(Index)
End Sub

'Added by Lydia 2017/10/19 ShellExecute程式無法直接開啟PDF檔的問題
'在網路上找的資料，應該是.pdf檔副檔名在本機註冊機碼中，預設開啟的程式路徑有問題。
Private Sub OpenDocument(sFile As String, sPath As String)
Dim sResult As String
Dim lSuccess As Long, lPos As Long
sResult = Space$(MAX_PATH)

 lSuccess = FindExecutable(sFile, sPath, sResult)
Select Case lSuccess
    Case ERROR_FILE_NO_ASSOCIATION
        If Right$(sFile, 3) = "pdf" Then
           MsgBox "You must have a PDF viewer such as Acrobat Reader to view pdf files."
        Else
           MsgBox "There is no registered program to open the selected file." & vbCrLf & sFile
        End If
    Case ERROR_FILE_NOT_FOUND: MsgBox "File not found: " & sFile
    Case ERROR_PATH_NOT_FOUND: MsgBox "Path not found: " & sPath
    Case ERROR_BAD_FORMAT: MsgBox "Bad format."
    Case Is >= ERROR_FILE_SUCCESS:
        lPos = InStr(sResult, Chr$(0))
        If lPos Then sResult = Left$(sResult, lPos - 1)
        MsgBox "PDF預設開啟程式:" & sResult
        
        SHELL sResult & " " & sFile, vbMaximizedFocus
End Select

'如何查看某個文件是和誰相關聯呢？例如：.txt是由哪個程式開啟，
'1.查[HKEY_CLASSES_ROOT\.txt]
'取預設值，如本人電腦預設值為 "txtfile"
'2.查[HKEY_CLASSES_ROOT\txtfile\shell\open\command]
'取預設值，如本人電腦預設值為 "C:\WINDOWS\NOTEPAD.EXE %1"
'如此可知.txt 是內定由NotePad.exe所執行。
'註：若step 1.取得的預設值是 "xxxx"，則step 2.便是查
'[HKEY_CLASSES_ROOT\xxxx\shell\open\command] 的預設值
End Sub

'Added by Lydia 2018/11/22
Private Sub txtUnicode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "txtUnicode_" & Format(Index, "00") Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "txtUnicode_" & Format(Index, "00")
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

'Added by Lydia 2019/12/12
Private Sub Chk2_Click(Index As Integer)
   If chk2(Index).Value = vbChecked Then
       If Index < 2 Then  '文字1/圖形
            If Index = 0 Then
               chk2(1).Value = vbUnchecked
            Else
               chk2(0).Value = vbUnchecked
            End If
            txtDt(6).Text = Index + 1
       Else  '文字2
           If Index = 2 Then
               chk2(3).Value = vbUnchecked
           Else
               chk2(2).Value = vbUnchecked
           End If
           txtDt(7).Text = Index - 1
       End If
   Else '取消
       If Index < 2 Then  '文字1/圖形
            If txtDt(6).Text = Trim(Index + 1) Then txtDt(6).Text = ""
       Else  '文字2
            If txtDt(7).Text = Trim(Index - 1) Then txtDt(7).Text = ""
       End If
   End If
End Sub

'Added by Lydia 2021/10/01
Private Sub textService_GotFocus()
    TextInverse textService
End Sub

'Added by Lydia 2021/10/01
Private Sub textCName_GotFocus()
    TextInverse textCName
End Sub

'Added by Lydia 2021/10/01
Private Sub textService_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If textService.Text <> "" Then
        CreateToolTip GetHWndForToolTip(textService), textService.Text
    End If

End Sub

'Added by Lydia 2021/10/01
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件ToolTip
End Sub

'Added by Lydia 2021/11/12 取得人員請假的職代
Private Function GetDutyList(ByVal stIdList As String) As String
Dim tmpArr As Variant
Dim inX As Integer
Dim stTmp1 As String
      
    GetDutyList = ""
    If stIdList <> "" Then
        tmpArr = Split(stIdList, ";")
        For inX = 0 To UBound(tmpArr)
            If Trim(tmpArr(inX)) <> "" Then
                stTmp1 = GetCaseDutyAgent(tmpArr(inX), "", False, , True, "A") 'A.指定抓全部職代
                If stTmp1 <> "" Then
                    GetDutyList = GetDutyList & ";" & stTmp1
                End If
            End If
        Next inX
    End If
    If GetDutyList <> "" Then GetDutyList = Mid(GetDutyList, 2)
End Function

'Added by Lydia 2024/07/17
Private Sub cmdRoute_Click()

   If Not nFrm Is Nothing Then
      nFrm.SetParent IIf(cmdRoute.Caption = "顯示", "Q", "M"), Me, mTQD02, txtField(11)
      nFrm.Show vbModal
   End If
   
End Sub

'Added by Lydia 2024/07/17
Public Sub SetData(ByVal pInputVal As String)
   txtField(11).Text = pInputVal
End Sub
