VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm000001 
   BorderStyle     =   1  '單線固定
   Caption         =   "維護作業1"
   ClientHeight    =   6720
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   10250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10250
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   2
      Left            =   9030
      TabIndex        =   0
      Top             =   0
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5952
      Left            =   48
      TabIndex        =   1
      Top             =   432
      Width           =   10128
      _ExtentX        =   17868
      _ExtentY        =   10495
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   882
      TabMaxWidth     =   3528
      TabCaption(0)   =   "維護作業"
      TabPicture(0)   =   "frm000001.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtCPP13"
      Tab(0).Control(1)=   "Command62"
      Tab(0).Control(2)=   "Command61"
      Tab(0).Control(3)=   "cmdExpWord(1)"
      Tab(0).Control(4)=   "txtTot"
      Tab(0).Control(5)=   "txtNow"
      Tab(0).Control(6)=   "Command60"
      Tab(0).Control(7)=   "Frame13"
      Tab(0).Control(8)=   "Command55"
      Tab(0).Control(9)=   "txtSQL"
      Tab(0).Control(10)=   "txtFTM(7)"
      Tab(0).Control(11)=   "cmdExpWord(0)"
      Tab(0).Control(12)=   "txtFTM(1)"
      Tab(0).Control(13)=   "txtFTM(2)"
      Tab(0).Control(14)=   "txtFTM(3)"
      Tab(0).Control(15)=   "txtFTM(4)"
      Tab(0).Control(16)=   "Command49"
      Tab(0).Control(17)=   "Command48"
      Tab(0).Control(18)=   "Command1(1)"
      Tab(0).Control(19)=   "Command1(0)"
      Tab(0).Control(20)=   "Text1(3)"
      Tab(0).Control(21)=   "Text1(2)"
      Tab(0).Control(22)=   "Text1(1)"
      Tab(0).Control(23)=   "Text1(0)"
      Tab(0).Control(24)=   "MSHFlexGrid1"
      Tab(0).Control(25)=   "ProgressBar1"
      Tab(0).Control(26)=   "lblElapse"
      Tab(0).Control(27)=   "Label16"
      Tab(0).Control(28)=   "Label29(12)"
      Tab(0).Control(29)=   "Label15"
      Tab(0).Control(30)=   "Label14"
      Tab(0).Control(31)=   "Label13"
      Tab(0).Control(32)=   "Label3(1)"
      Tab(0).Control(33)=   "Label3(2)"
      Tab(0).Control(34)=   "Label21"
      Tab(0).Control(35)=   "Label1(8)"
      Tab(0).Control(36)=   "Label1(3)"
      Tab(0).Control(37)=   "Label1(2)"
      Tab(0).Control(38)=   "Label1(1)"
      Tab(0).Control(39)=   "Label1(0)"
      Tab(0).ControlCount=   40
      TabCaption(1)   =   "其他"
      TabPicture(1)   =   "frm000001.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "TxtCaseNo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command44"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command43"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command42"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command41"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command40"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command36"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text5"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdCha"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command32"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Command30"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Command24"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Command23"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Command28"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtSpecW"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Command22"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Command21"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Command20"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Command2"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Command3"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Command4"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Command7"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Command9"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Command11"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Command12"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Command13"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Command14"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Command15"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Command16"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cmdSendMail"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Command46"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Command47"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Combo2"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Command51"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Command52"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Command53"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "過濾研討會 ＆ 大宗Email退件處理"
      TabPicture(2)   =   "frm000001.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "Frame6"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "其他2"
      TabPicture(3)   =   "frm000001.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command59"
      Tab(3).Control(1)=   "Command57"
      Tab(3).Control(2)=   "Frame12"
      Tab(3).Control(3)=   "Command54"
      Tab(3).Control(4)=   "Frame10"
      Tab(3).Control(5)=   "Frame3"
      Tab(3).Control(6)=   "Frame4"
      Tab(3).Control(7)=   "Command25"
      Tab(3).Control(8)=   "Frame5"
      Tab(3).Control(9)=   "Command31"
      Tab(3).Control(10)=   "Frame7"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "HP檔案 ＆ 利益衝突權限檢查"
      TabPicture(4)   =   "frm000001.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).Control(1)=   "Frame8"
      Tab(4).ControlCount=   2
      Begin VB.TextBox txtCPP13 
         Height          =   270
         Left            =   -66468
         TabIndex        =   197
         Text            =   "20240501"
         Top             =   3936
         Width           =   1344
      End
      Begin VB.CommandButton Command62 
         Caption         =   "暫停"
         Height          =   372
         Left            =   -67128
         TabIndex        =   196
         Top             =   3504
         Width           =   2004
      End
      Begin VB.CommandButton Command61 
         Caption         =   "封存->FTP"
         Height          =   348
         Left            =   -66072
         TabIndex        =   195
         Top             =   5520
         Width           =   1020
      End
      Begin VB.CommandButton cmdExpWord 
         Caption         =   "產生定稿(名稱帶入案件性質)"
         Height          =   564
         Index           =   1
         Left            =   -68040
         TabIndex        =   194
         Top             =   2784
         Width           =   3060
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  '靠右對齊
         Height          =   324
         Left            =   -67128
         Locked          =   -1  'True
         TabIndex        =   193
         Text            =   "0"
         Top             =   5136
         Width           =   1932
      End
      Begin VB.TextBox txtNow 
         Alignment       =   1  '靠右對齊
         Height          =   324
         Left            =   -67128
         Locked          =   -1  'True
         TabIndex        =   192
         Text            =   "0"
         Top             =   4536
         Width           =   1932
      End
      Begin VB.CommandButton Command60 
         Caption         =   "FTP->封存"
         Height          =   348
         Left            =   -67152
         TabIndex        =   187
         Top             =   5520
         Width           =   1020
      End
      Begin VB.CommandButton Command59 
         Caption         =   "CFT商品類別補足2碼"
         Height          =   444
         Left            =   -66912
         TabIndex        =   186
         Top             =   4848
         Width           =   1236
      End
      Begin VB.Frame Frame13 
         Caption         =   "SMTP測試"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1212
         Left            =   -74856
         TabIndex        =   176
         Top             =   552
         Width           =   6588
         Begin VB.TextBox txtSubj 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   2976
            TabIndex        =   185
            Text            =   "Mail Test"
            Top             =   264
            Width           =   2076
         End
         Begin VB.TextBox txtCC 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   1104
            TabIndex        =   183
            Top             =   864
            Width           =   3972
         End
         Begin VB.TextBox txtRcvr 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   1104
            TabIndex        =   182
            Text            =   "taie.taie@msa.hinet.net"
            Top             =   576
            Width           =   3948
         End
         Begin VB.CommandButton Command58 
            Caption         =   "測試"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   876
            Left            =   5208
            TabIndex        =   181
            Top             =   216
            Width           =   1212
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   1104
            TabIndex        =   178
            Text            =   "192.168.1.4"
            Top             =   240
            Width           =   1188
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "主旨:"
            Height          =   180
            Index           =   3
            Left            =   2496
            TabIndex        =   184
            Top             =   312
            Width           =   408
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "副本 :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Index           =   2
            Left            =   192
            TabIndex        =   180
            Top             =   912
            Width           =   504
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "收件者 :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Index           =   1
            Left            =   192
            TabIndex        =   179
            Top             =   624
            Width           =   708
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "SMTP IP :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Index           =   0
            Left            =   192
            TabIndex        =   177
            Top             =   312
            Width           =   816
         End
      End
      Begin VB.CommandButton Command57 
         Caption         =   "查名附件超過20M"
         Height          =   444
         Left            =   -66960
         TabIndex        =   175
         Top             =   4080
         Width           =   1476
      End
      Begin VB.Frame Frame12 
         Caption         =   "網中-查名單資料"
         Height          =   1692
         Left            =   -67176
         TabIndex        =   167
         Top             =   1896
         Width           =   2172
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   174
            Top             =   912
            Width           =   492
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   984
            MaxLength       =   7
            TabIndex        =   173
            Top             =   600
            Width           =   996
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   984
            MaxLength       =   7
            TabIndex        =   172
            Top             =   288
            Width           =   996
         End
         Begin VB.CommandButton Command56 
            Caption         =   "執行"
            Height          =   372
            Left            =   144
            TabIndex        =   171
            Top             =   1224
            Width           =   900
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "是否下載附件："
            Height          =   180
            Index           =   11
            Left            =   72
            TabIndex        =   170
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "終止日期："
            Height          =   180
            Index           =   10
            Left            =   72
            TabIndex        =   169
            Top             =   660
            Width           =   900
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "起始日期："
            Height          =   180
            Index           =   9
            Left            =   72
            TabIndex        =   168
            Top             =   336
            Width           =   900
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "利益衝突權限檢查"
         Height          =   1716
         Left            =   -74856
         TabIndex        =   151
         Top             =   3960
         Width           =   9516
         Begin VB.TextBox Text6 
            Height          =   300
            Index           =   4
            Left            =   4896
            TabIndex        =   164
            Top             =   264
            Width           =   780
         End
         Begin VB.CommandButton cmdChkCufa 
            Caption         =   "檢查"
            Height          =   348
            Left            =   7656
            TabIndex        =   162
            Top             =   264
            Width           =   1164
         End
         Begin VB.TextBox Text6 
            Height          =   300
            Index           =   3
            Left            =   1032
            MaxLength       =   12
            TabIndex        =   159
            Top             =   1296
            Width           =   1476
         End
         Begin VB.TextBox Text6 
            Height          =   300
            Index           =   2
            Left            =   888
            TabIndex        =   158
            Top             =   960
            Width           =   1050
         End
         Begin VB.TextBox Text6 
            Height          =   300
            Index           =   1
            Left            =   888
            TabIndex        =   157
            Top             =   624
            Width           =   1050
         End
         Begin VB.TextBox Text6 
            Height          =   300
            Index           =   0
            Left            =   888
            TabIndex        =   156
            Text            =   "ALL"
            Top             =   288
            Width           =   1884
         End
         Begin VB.Label Label10 
            Caption         =   "P.S. X/Y編號和本所案號可以輸入多項，若只輸入本所案號將以該案X+Y編號+案號為條件。"
            ForeColor       =   &H000000FF&
            Height          =   348
            Index           =   4
            Left            =   2568
            TabIndex        =   166
            Top             =   1320
            Width           =   6564
         End
         Begin MSForms.Label Label11 
            Height          =   276
            Index           =   4
            Left            =   5712
            TabIndex        =   165
            Top             =   264
            Width           =   1212
            BackColor       =   16777215
            Size            =   "2138;487"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label10 
            Caption         =   "員工編號："
            Height          =   228
            Index           =   6
            Left            =   3984
            TabIndex        =   163
            Top             =   288
            Width           =   948
         End
         Begin MSForms.Label Label11 
            Height          =   276
            Index           =   2
            Left            =   2000
            TabIndex        =   161
            Top             =   960
            Width           =   3500
            BackColor       =   16777215
            Size            =   "6174;487"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label Label11 
            Height          =   276
            Index           =   1
            Left            =   2000
            TabIndex        =   160
            Top             =   648
            Width           =   3500
            BackColor       =   16777215
            Size            =   "6174;487"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label10 
            Caption         =   "本所案號："
            Height          =   228
            Index           =   3
            Left            =   120
            TabIndex        =   155
            Top             =   1344
            Width           =   948
         End
         Begin VB.Label Label10 
            Caption         =   "系統別："
            Height          =   228
            Index           =   2
            Left            =   120
            TabIndex        =   154
            Top             =   336
            Width           =   948
         End
         Begin VB.Label Label10 
            Caption         =   "申請人："
            Height          =   228
            Index           =   1
            Left            =   120
            TabIndex        =   153
            Top             =   1008
            Width           =   948
         End
         Begin VB.Label Label10 
            Caption         =   "代理人："
            Height          =   228
            Index           =   0
            Left            =   120
            TabIndex        =   152
            Top             =   672
            Width           =   948
         End
      End
      Begin VB.CommandButton Command55 
         Caption         =   "更新FCT核准定稿"
         Height          =   492
         Left            =   -70200
         TabIndex        =   150
         Top             =   1824
         Width           =   924
      End
      Begin VB.CommandButton Command54 
         Caption         =   "外專備註設定還原6碼"
         Height          =   345
         Left            =   -74790
         TabIndex        =   149
         Top             =   1260
         Width           =   1845
      End
      Begin VB.TextBox txtSQL 
         Height          =   1272
         Left            =   -67200
         TabIndex        =   142
         Top             =   1476
         Width           =   2190
      End
      Begin VB.TextBox txtFTM 
         Height          =   345
         Index           =   7
         Left            =   -67215
         MaxLength       =   2
         TabIndex        =   141
         Top             =   1080
         Width           =   465
      End
      Begin VB.CommandButton cmdExpWord 
         Caption         =   "產生定稿"
         Height          =   390
         Index           =   0
         Left            =   -66570
         TabIndex        =   143
         Top             =   1020
         Width           =   1545
      End
      Begin VB.TextBox txtFTM 
         Height          =   345
         Index           =   1
         Left            =   -67215
         MaxLength       =   3
         TabIndex        =   137
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox txtFTM 
         Height          =   345
         Index           =   2
         Left            =   -66585
         MaxLength       =   2
         TabIndex        =   138
         Top             =   600
         Width           =   465
      End
      Begin VB.TextBox txtFTM 
         Height          =   345
         Index           =   3
         Left            =   -66090
         MaxLength       =   4
         TabIndex        =   139
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox txtFTM 
         Height          =   345
         Index           =   4
         Left            =   -65460
         MaxLength       =   2
         TabIndex        =   140
         Top             =   600
         Width           =   465
      End
      Begin VB.CommandButton Command53 
         Caption         =   "批次下載卷宗區檔案"
         Height          =   465
         Left            =   6420
         TabIndex        =   136
         Top             =   4080
         Width           =   1785
      End
      Begin VB.CommandButton Command52 
         Caption         =   "比對電子檔並刪除(信件)"
         Height          =   435
         Left            =   8550
         TabIndex        =   135
         Top             =   4500
         Width           =   1425
      End
      Begin VB.CommandButton Command51 
         Caption         =   "更新卷宗區電子檔名案號流水號足6碼"
         Height          =   435
         Left            =   3900
         TabIndex        =   130
         Top             =   5430
         Width           =   1755
      End
      Begin VB.Frame Frame10 
         Caption         =   "整批查詢造字EditEudcSearch，請注意查詢一筆約一分鐘"
         Height          =   675
         Left            =   -74700
         TabIndex        =   121
         Top             =   4890
         Width           =   7305
         Begin VB.CommandButton Command50 
            Caption         =   "查詢"
            Height          =   315
            Left            =   6240
            TabIndex        =   125
            Top             =   255
            Width           =   735
         End
         Begin VB.TextBox txtEES 
            Height          =   285
            Index           =   2
            Left            =   4710
            TabIndex        =   124
            Top             =   270
            Width           =   705
         End
         Begin VB.TextBox txtEES 
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   123
            Top             =   270
            Width           =   705
         End
         Begin VB.TextBox txtEES 
            Height          =   285
            Index           =   0
            Left            =   990
            TabIndex        =   122
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "結束號碼"
            Height          =   180
            Index           =   8
            Left            =   3930
            TabIndex        =   128
            Top             =   322
            Width           =   720
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "起始號碼"
            Height          =   180
            Index           =   7
            Left            =   2220
            TabIndex        =   127
            Top             =   322
            Width           =   720
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "匯入日期"
            Height          =   180
            Index           =   6
            Left            =   180
            TabIndex        =   126
            Top             =   322
            Width           =   720
         End
      End
      Begin VB.CommandButton Command49 
         Caption         =   "單排地址條-Excel套印"
         Height          =   435
         Left            =   -71910
         TabIndex        =   120
         Top             =   1788
         Width           =   1155
      End
      Begin VB.CommandButton Command48 
         Caption         =   "抓C類備註設定"
         Height          =   495
         Left            =   -74610
         TabIndex        =   119
         Top             =   1944
         Width           =   1365
      End
      Begin VB.ComboBox Combo2 
         Height          =   260
         Left            =   9030
         Style           =   2  '單純下拉式
         TabIndex        =   110
         Top             =   5100
         Width           =   780
      End
      Begin VB.CommandButton Command47 
         Caption         =   "重下載卷宗區電子檔"
         Height          =   465
         Left            =   2010
         TabIndex        =   109
         Top             =   5400
         Width           =   1785
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Email符合清單 - Lydia"
         Height          =   525
         Left            =   3000
         TabIndex        =   108
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "確定(&O)"
         Height          =   405
         Index           =   1
         Left            =   -70200
         TabIndex        =   101
         Top             =   2808
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&S)"
         Default         =   -1  'True
         Height          =   405
         Index           =   0
         Left            =   -70200
         TabIndex        =   100
         Top             =   2364
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   -72150
         MaxLength       =   12
         TabIndex        =   99
         Top             =   2892
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   -73950
         MaxLength       =   12
         TabIndex        =   98
         Top             =   2892
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   -71790
         MaxLength       =   9
         TabIndex        =   97
         Top             =   2532
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   -73950
         MaxLength       =   9
         TabIndex        =   96
         Top             =   2532
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         Caption         =   "過濾研討會資料"
         Height          =   1305
         Left            =   -74880
         TabIndex        =   89
         Top             =   870
         Width           =   7125
         Begin VB.CommandButton Command10 
            Caption         =   "測試IE"
            Height          =   345
            Left            =   3900
            TabIndex        =   94
            Top             =   780
            Width           =   1005
         End
         Begin VB.CommandButton Command8 
            Caption         =   "產生文字檔"
            Height          =   345
            Left            =   1980
            TabIndex        =   93
            Top             =   780
            Width           =   1665
         End
         Begin VB.CommandButton Command6 
            Caption         =   "<="
            Height          =   345
            Left            =   6690
            TabIndex        =   92
            Top             =   330
            Width           =   345
         End
         Begin VB.CommandButton Command5 
            Caption         =   "過濾研討會Excel檔"
            Height          =   345
            Left            =   5160
            TabIndex        =   91
            Top             =   780
            Width           =   1875
         End
         Begin VB.TextBox txtFileName 
            Height          =   264
            Left            =   1560
            TabIndex        =   90
            Top             =   360
            Width           =   5085
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   90
            Top             =   750
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            Caption         =   "研討會檔案："
            Height          =   210
            Index           =   4
            Left            =   390
            TabIndex        =   95
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmdSendMail 
         BackColor       =   &H000080FF&
         Caption         =   "寄郵件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2880
         Style           =   1  '圖片外觀
         TabIndex        =   88
         Top             =   2295
         Width           =   2025
      End
      Begin VB.CommandButton Command16 
         Caption         =   "刪除電子檔"
         Height          =   435
         Left            =   2880
         TabIndex        =   87
         Top             =   3015
         Width           =   1425
      End
      Begin VB.CommandButton Command15 
         Caption         =   "往來記錄附件轉檔"
         Height          =   435
         Left            =   4860
         TabIndex        =   86
         Top             =   4095
         Width           =   1425
      End
      Begin VB.CommandButton Command14 
         Caption         =   "專利EPC英國有效及EU案抓關聯案申請日105/07/06-Sonia"
         Height          =   700
         Left            =   120
         TabIndex        =   85
         Top             =   1695
         Width           =   1965
      End
      Begin VB.CommandButton Command13 
         Caption         =   "刪除已發文聯絡附件"
         Height          =   345
         Left            =   120
         TabIndex        =   84
         Top             =   1215
         Width           =   2115
      End
      Begin VB.CommandButton Command12 
         Caption         =   "專利公報申請人更新客戶編號及代理人名稱"
         Height          =   630
         Left            =   5520
         TabIndex        =   83
         Top             =   3255
         Width           =   1965
      End
      Begin VB.CommandButton Command11 
         Caption         =   "轉郵遞區號資料"
         Height          =   315
         Left            =   360
         TabIndex        =   82
         Top             =   5175
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "ACC1K0轉檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         TabIndex        =   81
         Top             =   1695
         Width           =   1665
      End
      Begin VB.CommandButton Command7 
         Caption         =   "補掛期限"
         Height          =   435
         Left            =   5550
         TabIndex        =   80
         Top             =   2580
         Width           =   1725
      End
      Begin VB.CommandButton Command4 
         Caption         =   "五都修改地址通知函"
         Height          =   345
         Left            =   2820
         TabIndex        =   79
         Top             =   1215
         Width           =   2115
      End
      Begin VB.CommandButton Command3 
         Caption         =   "CFT商品類別比對"
         Height          =   345
         Left            =   5550
         TabIndex        =   78
         Top             =   2085
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CommandButton Command2 
         Caption         =   "TM,TG商品類別比對"
         Height          =   345
         Left            =   5550
         TabIndex        =   77
         Top             =   1725
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Frame Frame2 
         Caption         =   "大宗Email退件處理"
         Height          =   1665
         Left            =   -74820
         TabIndex        =   66
         Top             =   2430
         Width           =   7125
         Begin VB.TextBox Text2 
            Height          =   300
            Left            =   1560
            TabIndex        =   72
            Top             =   360
            Width           =   5085
         End
         Begin VB.CommandButton Command17 
            Caption         =   "<="
            Height          =   345
            Left            =   6680
            TabIndex        =   71
            Top             =   360
            Width           =   345
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   70
            Top             =   780
            Width           =   735
         End
         Begin VB.CommandButton Command18 
            Caption         =   "大宗Email退件處理"
            Height          =   345
            Left            =   1560
            TabIndex        =   69
            Top             =   1200
            Width           =   1845
         End
         Begin VB.TextBox Text4 
            Height          =   300
            Left            =   5400
            MaxLength       =   8
            TabIndex        =   68
            Top             =   780
            Width           =   975
         End
         Begin VB.CommandButton Command19 
            Caption         =   "補LOG"
            Height          =   375
            Left            =   4200
            TabIndex        =   67
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "來源檔案："
            Height          =   180
            Index           =   5
            Left            =   480
            TabIndex        =   76
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "提供人員："
            Height          =   180
            Index           =   6
            Left            =   480
            TabIndex        =   75
            Top             =   840
            Width           =   900
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            Height          =   180
            Left            =   2400
            TabIndex        =   74
            Top             =   840
            Width           =   45
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "退件郵件整理日期："
            Height          =   180
            Index           =   7
            Left            =   3720
            TabIndex        =   73
            Top             =   840
            Width           =   1620
         End
      End
      Begin VB.CommandButton Command20 
         Caption         =   "107/3/5 後 FCP中說進度發文 批次刪*.msg Lydia"
         Height          =   735
         Left            =   90
         TabIndex        =   65
         Top             =   2460
         Width           =   1935
      End
      Begin VB.CommandButton Command21 
         Caption         =   "刪除FC撰寫信函上傳到卷宗區的檔案 - Lydia"
         Height          =   735
         Left            =   8130
         TabIndex        =   64
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command22 
         Caption         =   "更新命名901,203     承辦期限-  Lydia"
         Height          =   555
         Left            =   8190
         TabIndex        =   63
         Top             =   2160
         Width           =   1785
      End
      Begin VB.TextBox txtSpecW 
         Height          =   300
         Left            =   240
         TabIndex        =   62
         Top             =   5175
         Visible         =   0   'False
         Width           =   192
      End
      Begin VB.Frame Frame3 
         Caption         =   "GDPR回函處理"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   56
         Top             =   1800
         Width           =   7455
         Begin VB.TextBox txtGDPR 
            Height          =   300
            Index           =   0
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   "\\Typing2\國外部\業務拓展處\GDPR回覆\I CONSENT"
            Top             =   337
            Width           =   5055
         End
         Begin VB.TextBox txtGDPR 
            Height          =   300
            Index           =   1
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   58
            Text            =   "\\Typing2\國外部\業務拓展處\GDPR回覆\I DO NOT CONSENT"
            Top             =   697
            Width           =   5055
         End
         Begin VB.CommandButton Command27 
            Caption         =   "執行"
            Height          =   330
            Left            =   6480
            TabIndex        =   57
            Top             =   480
            Width           =   700
         End
         Begin VB.Label Label6 
            Caption         =   "同意-路徑"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "不同意-路徑"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "改定稿日期"
         Height          =   705
         Left            =   -74760
         TabIndex        =   47
         Top             =   3270
         Width           =   7485
         Begin VB.TextBox Text29 
            Height          =   270
            Index           =   2
            Left            =   5085
            TabIndex        =   51
            Top             =   270
            Width           =   960
         End
         Begin VB.TextBox Text29 
            Height          =   270
            Index           =   1
            Left            =   3330
            TabIndex        =   50
            Top             =   270
            Width           =   960
         End
         Begin VB.CommandButton Command29 
            Caption         =   "執行"
            Height          =   315
            Left            =   6600
            TabIndex        =   49
            Top             =   270
            Width           =   735
         End
         Begin VB.TextBox Text29 
            Height          =   270
            Index           =   0
            Left            =   810
            TabIndex        =   48
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "XXX"
            Height          =   180
            Index           =   3
            Left            =   1845
            TabIndex        =   55
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "新日期"
            Height          =   180
            Index           =   2
            Left            =   4455
            TabIndex        =   54
            Top             =   315
            Width           =   540
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "原日期"
            Height          =   180
            Index           =   1
            Left            =   2700
            TabIndex        =   53
            Top             =   315
            Width           =   540
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "建立人"
            Height          =   180
            Index           =   0
            Left            =   180
            TabIndex        =   52
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.CommandButton Command28 
         Caption         =   "CFT補催審期限 - Lydia"
         Height          =   495
         Left            =   8310
         TabIndex        =   46
         Top             =   2820
         Width           =   1575
      End
      Begin VB.CommandButton Command25 
         Caption         =   "國外帳單檢查原文字數"
         Height          =   405
         Left            =   -74760
         TabIndex        =   45
         Top             =   750
         Width           =   2115
      End
      Begin VB.CommandButton Command23 
         Caption         =   "北所南所刪址客戶處理Lydia"
         Height          =   525
         Left            =   120
         TabIndex        =   44
         Top             =   3270
         Width           =   1935
      End
      Begin VB.CommandButton Command24 
         Caption         =   "財務寄發郵件 Sindy"
         Height          =   465
         Left            =   120
         TabIndex        =   43
         Top             =   3900
         Width           =   2085
      End
      Begin VB.Frame Frame5 
         Caption         =   "電子公文設定手動下載"
         Height          =   705
         Left            =   -74730
         TabIndex        =   37
         Top             =   4080
         Width           =   7485
         Begin VB.TextBox txtWD07 
            Height          =   270
            Left            =   3840
            TabIndex        =   40
            Top             =   270
            Width           =   480
         End
         Begin VB.CommandButton Command26 
            Caption         =   "存檔"
            Height          =   315
            Left            =   6615
            TabIndex        =   39
            Top             =   270
            Width           =   735
         End
         Begin VB.TextBox txtWD01 
            Height          =   270
            Left            =   810
            MaxLength       =   7
            TabIndex        =   38
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "是否手動下載"
            Height          =   180
            Index           =   5
            Left            =   2700
            TabIndex        =   42
            Top             =   315
            Width           =   1080
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "工作日"
            Height          =   180
            Index           =   4
            Left            =   180
            TabIndex        =   41
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.CommandButton Command30 
         Caption         =   "匯入xls比對下一程序 Sindy"
         Height          =   495
         Left            =   150
         TabIndex        =   36
         Top             =   4500
         Width           =   2355
      End
      Begin VB.CommandButton Command31 
         Caption         =   "計算分信時段"
         Height          =   405
         Left            =   -72540
         TabIndex        =   35
         Top             =   750
         Width           =   1545
      End
      Begin VB.CommandButton Command32 
         Caption         =   "匯入xls比對Email查資料 Sindy"
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   630
         Width           =   2355
      End
      Begin VB.Frame Frame6 
         Caption         =   "外專利益衝突管制"
         Height          =   1545
         Left            =   -74790
         TabIndex        =   26
         Top             =   4200
         Width           =   7095
         Begin VB.CommandButton Command45 
            Caption         =   "更換"
            Height          =   375
            Left            =   5640
            TabIndex        =   133
            Top             =   1080
            Width           =   945
         End
         Begin VB.TextBox TxtFile 
            Height          =   285
            Left            =   210
            TabIndex        =   31
            Text            =   "TxtFile"
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton Command34 
            Caption         =   "更換Unicode的檔名"
            Height          =   345
            Left            =   180
            TabIndex        =   30
            Top             =   210
            Width           =   1665
         End
         Begin VB.CommandButton Command33 
            Caption         =   "English_Vers案件清單"
            Height          =   255
            Left            =   4920
            TabIndex        =   29
            Top             =   120
            Width           =   2145
         End
         Begin VB.CommandButton Command35 
            Caption         =   "檢查D類"
            Height          =   345
            Left            =   1980
            TabIndex        =   28
            Top             =   150
            Width           =   1035
         End
         Begin VB.CommandButton Command39 
            Caption         =   "刪除D類"
            Height          =   345
            Left            =   3570
            TabIndex        =   27
            Top             =   150
            Width           =   1035
         End
         Begin VB.Label Label9 
            Caption         =   "P.S.全中文要逐一檢查定稿，最好找到有含日文字 "
            ForeColor       =   &H00FF00FF&
            Height          =   285
            Left            =   2250
            TabIndex        =   134
            Top             =   810
            Width           =   4635
         End
         Begin VB.Label Label8 
            Caption         =   "更換後："
            Height          =   315
            Index           =   1
            Left            =   3210
            TabIndex        =   132
            Top             =   1110
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "更換日文定稿："
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   131
            Top             =   1110
            Width           =   1335
         End
         Begin MSForms.TextBox txtFM2 
            Height          =   345
            Index           =   1
            Left            =   4050
            TabIndex        =   33
            Top             =   1095
            Width           =   1500
            VariousPropertyBits=   671107099
            Size            =   "2646;609"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtFM2 
            Height          =   345
            Index           =   0
            Left            =   1560
            TabIndex        =   32
            Top             =   1095
            Width           =   1500
            VariousPropertyBits=   671107099
            Size            =   "2646;609"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.CommandButton cmdCha 
         Caption         =   "檢查中文字"
         Height          =   315
         Left            =   6690
         TabIndex        =   25
         Top             =   1170
         Width           =   1155
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   2850
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   24
         Top             =   630
         Width           =   5085
      End
      Begin VB.CommandButton Command36 
         Caption         =   "更新a0w16"
         Height          =   375
         Left            =   5130
         TabIndex        =   23
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Frame Frame7 
         Caption         =   "刪除卷宗區CPP電子檔"
         Height          =   1245
         Left            =   -70860
         TabIndex        =   17
         Top             =   600
         Width           =   3765
         Begin VB.CommandButton cmdDelCPP 
            Caption         =   "刪除電子檔"
            Height          =   315
            Left            =   2370
            TabIndex        =   20
            Top             =   120
            Width           =   1305
         End
         Begin VB.ListBox lstImport 
            Height          =   220
            ItemData        =   "frm000001.frx":008C
            Left            =   180
            List            =   "frm000001.frx":008E
            TabIndex        =   19
            Top             =   450
            Width           =   2595
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "匯入..."
            Height          =   315
            Left            =   2820
            TabIndex        =   18
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label3 
            Caption         =   "輸入總收文號："
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   22
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label lblImpCount 
            AutoSize        =   -1  'True
            Caption         =   "000"
            Height          =   180
            Left            =   2910
            TabIndex        =   21
            Top             =   1020
            Width           =   270
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "HP檔案壓縮加密"
         Height          =   2760
         Left            =   -74820
         TabIndex        =   7
         Top             =   660
         Width           =   9585
         Begin VB.CommandButton Command37 
            Caption         =   "開始"
            Height          =   375
            Left            =   8250
            TabIndex        =   12
            Top             =   270
            Width           =   1185
         End
         Begin VB.ListBox lstHistory 
            Height          =   940
            Left            =   90
            TabIndex        =   11
            Top             =   1290
            Width           =   9405
         End
         Begin VB.Timer tmrClock 
            Left            =   7590
            Top             =   810
         End
         Begin VB.OptionButton Option1 
            Caption         =   "casepaperpdf"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   10
            Top             =   990
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton Option1 
            Caption         =   "casepaperfile"
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   9
            Top             =   990
            Width           =   1725
         End
         Begin VB.CommandButton Command38 
            Caption         =   "暫停"
            Height          =   375
            Left            =   8250
            TabIndex        =   8
            Top             =   780
            Width           =   1185
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   270
            Width           =   6090
            _ExtentX        =   10724
            _ExtentY        =   512
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label4 
            Alignment       =   2  '置中對齊
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1320
            TabIndex        =   16
            Top             =   600
            Width           =   1140
         End
         Begin VB.Label Label5 
            Caption         =   "已過時間："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.5
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   15
            Top             =   630
            Width           =   1140
         End
         Begin VB.Label Label7 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6450
            TabIndex        =   14
            Top             =   300
            Width           =   1725
         End
      End
      Begin VB.CommandButton Command40 
         Caption         =   "比對清單之對造 - Lydia"
         Height          =   495
         Left            =   8340
         TabIndex        =   6
         Top             =   750
         Width           =   1605
      End
      Begin VB.CommandButton Command41 
         Caption         =   "檢查-更換造字欄位表 - Lydia"
         Height          =   495
         Left            =   8340
         TabIndex        =   5
         Top             =   3660
         Width           =   1485
      End
      Begin VB.CommandButton Command42 
         Caption         =   "客戶及代理人清單 Lydia"
         Height          =   465
         Left            =   3000
         TabIndex        =   4
         Top             =   3660
         Width           =   1305
      End
      Begin VB.CommandButton Command43 
         Caption         =   "輸入收文號查詢本所案號"
         Height          =   435
         Left            =   7530
         TabIndex        =   3
         Top             =   5040
         Width           =   1425
      End
      Begin VB.CommandButton Command44 
         Caption         =   "重新下載CFT,CFP之UK檔"
         Height          =   465
         Left            =   2970
         TabIndex        =   2
         Top             =   4170
         Width           =   1425
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1752
         Left            =   -74880
         TabIndex        =   102
         Top             =   3516
         Width           =   7548
         _ExtentX        =   13317
         _ExtentY        =   3087
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame9"
         Height          =   1035
         Left            =   5970
         TabIndex        =   111
         Top             =   4770
         Width           =   2025
         Begin VB.TextBox Text10 
            Height          =   270
            Index           =   12
            Left            =   1050
            TabIndex        =   118
            Text            =   "Text6"
            Top             =   630
            Width           =   315
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Index           =   28
            Left            =   570
            TabIndex        =   117
            Text            =   "Text6"
            Top             =   630
            Width           =   315
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Index           =   23
            Left            =   120
            TabIndex        =   116
            Text            =   "Text6"
            Top             =   630
            Width           =   315
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Index           =   5
            Left            =   1530
            TabIndex        =   115
            Text            =   "Text6"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Index           =   4
            Left            =   990
            TabIndex        =   114
            Text            =   "Text6"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Index           =   0
            Left            =   570
            TabIndex        =   113
            Text            =   "Text6"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   112
            Text            =   "Text6"
            Top             =   240
            Width           =   315
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   288
         Left            =   -74880
         TabIndex        =   188
         Top             =   5544
         Width           =   5784
         _ExtentX        =   10195
         _ExtentY        =   512
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblElapse 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   -69024
         TabIndex        =   200
         Top             =   5304
         Width           =   1668
      End
      Begin VB.Label Label16 
         Alignment       =   1  '靠右對齊
         Caption         =   "已過時間："
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   10
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   -70512
         TabIndex        =   199
         Top             =   5304
         Width           =   1428
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "CPP13<"
         Height          =   180
         Index           =   12
         Left            =   -67104
         TabIndex        =   198
         Top             =   3984
         Width           =   564
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "累計大小(K)："
         Height          =   180
         Left            =   -67128
         TabIndex        =   191
         Top             =   4920
         Width           =   1140
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "目前大小(K)："
         Height          =   180
         Left            =   -67128
         TabIndex        =   190
         Top             =   4296
         Width           =   1140
      End
      Begin VB.Label Label13 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   -69024
         TabIndex        =   189
         Top             =   5580
         Width           =   1728
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "自訂條件："
         Height          =   180
         Index           =   1
         Left            =   -68160
         TabIndex        =   147
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "( SQL語法 )"
         Height          =   180
         Index           =   2
         Left            =   -68160
         TabIndex        =   146
         Top             =   1710
         Width           =   885
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文："
         Height          =   180
         Left            =   -68160
         TabIndex        =   145
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿編號："
         Height          =   180
         Index           =   8
         Left            =   -68160
         TabIndex        =   144
         Top             =   690
         Width           =   900
      End
      Begin MSForms.TextBox TxtCaseNo 
         Height          =   315
         Left            =   8040
         TabIndex        =   129
         Top             =   5520
         Width           =   1995
         VariousPropertyBits=   746604571
         Size            =   "3519;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "---"
         Height          =   180
         Index           =   3
         Left            =   -72444
         TabIndex        =   107
         Top             =   2928
         Width           =   180
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號 : "
         Height          =   228
         Index           =   2
         Left            =   -74880
         TabIndex        =   106
         Top             =   2928
         Width           =   912
      End
      Begin VB.Label Label1 
         Caption         =   "新編號 : "
         Height          =   228
         Index           =   1
         Left            =   -72720
         TabIndex        =   105
         Top             =   2568
         Width           =   768
      End
      Begin VB.Label Label1 
         Caption         =   "舊編號 : "
         Height          =   228
         Index           =   0
         Left            =   -74880
         TabIndex        =   104
         Top             =   2568
         Width           =   768
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   103
         Top             =   5655
         Width           =   4185
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   315
      Left            =   0
      TabIndex        =   148
      Top             =   6420
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   547
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm000001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/5 改成Form2.0 (無;Printer列印未改;都為測試或一次性程式,要用時再判斷是否需調整)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'sonia 2010/8/17 日期欄已修改
Option Explicit

'開啟檔案對話框
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
  "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
  
Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Dim m_bolStop As Boolean 'Added by Morgan 2020/7/9
Dim m_oFileSys As New FileSystemObject 'Added by Morgan 2025/5/26

'Add By Sindy 2020/3/26 檢查中文字
Private Sub cmdCha_Click()
Dim strText As String
Dim jj As Integer

'   StrConv(string, conversion, LCID)
'   半形轉全形: vbWide
'   全形轉半形: vbNarrow
'   TT = StrConv("abcd123", vbNarrow)
'   ==>TT="ａｂｃｄ１２３"
   
'   Dim strChr As String
'   'chr(88)=X
'   strChr = Chr(Asc("q")) '【株式?社????】?連絡?更??願? FCP  FW: ◎Fax data from 3f@taie.com.tw
'   If LenB(strChr) = LenB(StrConv(strChr, vbFromUnicode)) Then
'      MsgBox "全形字元"
'   Else
'      MsgBox "半形字元"
'   End If
'   Exit Sub
   
   If Text5.Text = "" Then
      strText = InputBox("請輸入要檢查的字串：", "檢查中文字")
      Text5.Text = strText
   Else
      strText = Text5.Text
   End If
   If strText <> "" Then
      '檢查字串是否有中文或全形字
      For jj = 1 To Len(strText)
         If Asc(Mid(strText, jj, 1)) <= 0 Then
            If InStr(strText, "?") > 0 Then
               MsgBox strText & vbCrLf & vbCrLf & "有含unicode的中文字!!!"
            Else
               MsgBox strText & vbCrLf & vbCrLf & "有中文字!!!"
            End If
            Exit Sub
         End If
      Next jj
      MsgBox "無"
   End If
End Sub

'Add By Sindy 2020/6/2 可以整批刪除卷宗區電子檔; 如.商標處整批催延展要重新輸入
Private Sub cmdDelCPP_Click()
   
End Sub

'Modified by Lydia 2025/04/08 +ByVal pIdx As Integer
Private Sub MkFile(Optional bolPrint As Boolean, Optional ByVal pIdx As Integer)
   Dim stSQL As String, stCon As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim strFN As String 'Added by Lydia 2025/04/08
   
   stCon = ""
   If txtFTM(1) <> "" Then
      stCon = stCon & " and ftm01='" & txtFTM(1) & "'"
   End If
   If txtFTM(2) <> "" Then
      stCon = stCon & " and ftm02='" & txtFTM(2) & "'"
   End If
   If txtFTM(3) <> "" Then
      stCon = stCon & " and ftm03='" & txtFTM(3) & "'"
   End If
   If txtFTM(4) <> "" Then
      stCon = stCon & " and ftm04='" & txtFTM(4) & "'"
   End If
   If txtFTM(7) <> "" Then
      stCon = stCon & " and ftm07='" & txtFTM(7) & "'"
   End If
   If txtSQL.Text <> "" Then
      stCon = stCon & txtSQL.Text
   End If
   
   'Added by Lydia 2025/04/08 名稱帶入案件性質
   If pIdx = 1 Then
      stSQL = "select decode(ftm03,'000','通用',decode(nvl(cpm03,cpm04),'（無）',nvl(cpm04,cpm03),nvl(cpm03,cpm04))) as cpm0304,f1.*,t1.* " & _
              "from finaltextmap f1,texttype t1,casepropertymap c1 where typ01(+)=ftm02 and typ02(+)=ftm01 and ftm01=cpm01(+) and ftm03=cpm02(+) " & stCon
   Else
   'end 2025/04/08
      'stcon = stcon & " AND ftm01='CFP' and instr(ftm06,'PCT')>0　AND FTM02='01' AND FTM04<'85'"
      stSQL = "select * from finaltextmap,TEXTTYPE where typ01(+)=ftm02 and typ02(+)=ftm01 " & stCon
   End If
   StatusBar1.Panels(1).Text = "讀取定稿..."
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      StatusBar1.Panels(1).Text = "建立定稿Doc..."
      If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
      intQ = 0
      Do While Not .EOF
         'Added by Lydia 2025/04/08 名稱帶入案件性質
         If pIdx = 1 Then
            strFN = .Fields("FTM01") & "-" & .Fields("FTM02") & "-" & .Fields("FTM03") & "-" & .Fields("FTM04") & ChgFileName("(" & .Fields("cpm0304") & ")" & .Fields("FTM06"))
         Else
            strFN = .Fields("FTM01") & "-" & .Fields("FTM02") & "-" & .Fields("FTM03") & "-" & .Fields("FTM04") & ChgFileName("" & .Fields("FTM06"))
         End If
         'end 2025/04/08
         'If MakeDoc(.Fields("FTM05"), .Fields("FTM01") & "-" & .Fields("FTM02") & "-" & .Fields("FTM03") & "-" & .Fields("FTM04") & .Fields("TYP03"), "" & .Fields("FTM06")) = False Then
         'Modified by Lydia 2025/04/08
         'If MakeDoc(.Fields("FTM05") & .Fields("FTM08"), .Fields("FTM01") & "-" & .Fields("FTM02") & "-" & .Fields("FTM03") & "-" & .Fields("FTM04") & ChgFileName("" & .Fields("FTM06")), "", bolPrint) = False Then
         If MakeDoc(.Fields("FTM05") & .Fields("FTM08"), strFN, "", bolPrint) = False Then
            Exit Do
         Else
            intQ = intQ + 1
            If bolPrint Then Exit Do
         End If
         .MoveNext
      Loop
      StatusBar1.Panels(1).Text = "完成..."
      If g_WordAp.Documents.Count = 0 Then g_WordAp.Quit
      End With
      If intQ > 0 Then MsgBox "定稿已產生至我的文件共 " & intQ & " 個檔案。"
   Else
      MsgBox "無符合條件定稿！", vbInformation
   End If

   Set g_WordAp = Nothing
   Set rsQuery = Nothing
End Sub

Private Function ChgFileName(p_OldName As String) As String
   Dim stNewName As String
   stNewName = Replace(p_OldName, ">", "＞")
   stNewName = Replace(stNewName, "/", "／")
   stNewName = Replace(stNewName, vbCrLf, "")
   ChgFileName = stNewName
End Function

Private Function MakeDoc(ByVal p_Text As String, ByVal p_Name As String, ByVal p_Add As String, Optional p_PrintOut As Boolean) As Boolean

   Dim b2Time As Boolean
   
On Error GoTo ErrHandle
   
   With g_WordAp
      .Documents.add
      '切換為整頁模式
      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
         .ActiveWindow.ActivePane.View.Type = wdPageView
      Else
         .ActiveWindow.View.Type = wdPageView
      End If
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .ActiveWindow.Selection.TypeText p_Text
      .ActiveWindow.View.SeekView = wdSeekCurrentPageHeader
      .Selection.TypeText p_Name & p_Add
      If p_PrintOut Then
         .PrintOut Background:=False, Copies:=1, Collate:=True
      Else
         .Documents(1).SaveAs p_Name & ".doc"
      End If
      .ActiveDocument.Close wdDoNotSaveChanges
   End With
   
   MakeDoc = True
   
ErrHandle:
   If Err.Number = 462 And b2Time = False Then
      Set g_WordAp = New Word.Application
      b2Time = True
      Resume
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
      Resume
   End If
   
End Function

'Modified by Lydia 2025/04/08 增加按鈕 +index
'Private Sub cmdExpWord_Click()
 ' MkFile
Private Sub cmdExpWord_Click(Index As Integer)
   MkFile , Index
'end 2025/04/08
End Sub

'Add By Sindy 2020/6/2 匯入...
Private Sub cmdImPort_Click()
   Dim stFileName As String
   Dim OpenFile As OPENFILENAME
   Dim lReturn As Long
   Dim sFilter As String
   Dim SNo As String
   Dim stCaption As String
   Dim intA As Integer
   Dim rsAD As New ADODB.Recordset
   Dim fso As New FileSystemObject
   Dim ts As TextStream
   Dim ii As Integer
   
On Error GoTo ErrHnd
   
   OpenFile.lStructSize = Len(OpenFile)
   OpenFile.hwndOwner = Me.hWnd
   OpenFile.hInstance = App.hInstance
      sFilter = "文字檔(*.TXT)" & Chr(0) & "*.txt" & Chr(0) & "所有檔案" & Chr(0) & "*.*" & Chr(0)
   OpenFile.lpstrFilter = sFilter
   OpenFile.nFilterIndex = 1
   OpenFile.lpstrFile = String(257, 0)
   OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
   OpenFile.lpstrFileTitle = OpenFile.lpstrFile
   OpenFile.nMaxFileTitle = OpenFile.nMaxFile
   OpenFile.lpstrInitialDir = PUB_Getdesktop
   stCaption = "匯入清單"
   OpenFile.lpstrTitle = stCaption
   OpenFile.Flags = 0
   lReturn = GetOpenFileName(OpenFile)
   If lReturn <> 0 Then
      stFileName = Trim(OpenFile.lpstrFile)
      lstImport.Clear
      lblImpCount.Caption = ""
      If fso.FileExists(stFileName) Then
         Set ts = fso.OpenTextFile(stFileName)
         Do While Not ts.AtEndOfStream
            SNo = Replace(RTrim(ts.ReadLine), " ", "")
            If SNo <> "" Then
               lstImport.AddItem SNo
            End If
         Loop
         ts.Close
      End If
      lblImpCount.Caption = lstImport.ListCount
   End If
   If Val(lblImpCount.Caption) > 0 Then
      If MsgBox("確定是否要刪除卷宗區電子檔？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Sub
      End If
      For ii = lstImport.ListCount - 1 To 0 Step -1
         '檢查文號是否不存在進度檔,不存在,才刪卷宗區
         strExc(1) = "select cp09 from caseprogress where cp09='" & lstImport.List(ii) & "'"
         intA = 1
         Set rsAD = ClsLawReadRstMsg(intA, strExc(1))
         If intA = 0 Then
            '卷宗區是否有資料,有,執行刪除
            strExc(1) = "select cpp01 from casepaperpdf where cpp01='" & lstImport.List(ii) & "'"
            intA = 1
            Set rsAD = ClsLawReadRstMsg(intA, strExc(1))
            If intA = 1 Then
               ', " and CPP02=" & CNULL("" & RsTemp.Fields("cpp02"))
               If PUB_DelFtpFile2(rsAD.Fields("cpp01")) = True Then
                  strSql = "delete from casepaperpdf where cpp01='" & rsAD.Fields("cpp01") & "'" 'and cpp02=" & CNULL(RsTemp.Fields("cpp02"))
                  cnnConnection.Execute strSql, intI
                  '有刪除成功,才把畫面上文號移除
                  lstImport.RemoveItem ii
               End If
            End If
         End If
      Next ii
   Else
      MsgBox "無匯入資料!"
      Exit Sub
   End If
   
   MsgBox "電子檔刪除完畢!"
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2017/7/21 寄郵件
Private Sub cmdSendMail_Click()
Dim str01002 As String
Dim str01005 As String
Dim strName As String, strMail As String
Dim strSubject As String, strContext As String

'   strSql = "select r01002,r01005,r01013,r01004,ID,r01001 from r100104 where ID='M31' and r01001 is null"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         If "" & adoRecordset.Fields("r01002") <> "" Then
'            str01002 = adoRecordset.Fields("r01002")
'            str01005 = "" & adoRecordset.Fields("r01005")
'         Else
'            str01002 = ""
'            str01005 = adoRecordset.Fields("r01005")
'         End If
'         strName = adoRecordset.Fields("r01013")
'         strMail = adoRecordset.Fields("r01004")
'
'         strSubject = "為建立客戶及客戶會計師事務所email資料(" & adoRecordset.Fields("r01013") & IIf(str01002 <> "", str01002, "") & ")"
'         strContext = "您好,收信愉快," & vbCrLf & vbCrLf & _
'"本所為提供更完善的服務,欲建立客戶及會計師事務所email資料" & vbCrLf & _
'"以期在每年年底能系統化的將「年度扣繳明細」 eMAIL或傳真給客戶或事務所參考" & vbCrLf & _
'"煩請:" & vbCrLf & _
'"1.若此信箱是屬會計師事務所信箱,請直接以本MAIL回覆以下 貴事務所資料" & vbCrLf & _
'"2.若此信箱是為公司內部信箱而您的扣繳事項是由會計師事務所處理,請以下會計師事務所資料供本所建檔" & vbCrLf & _
'"3.所需要的會計師事務所資料如下:" & vbCrLf & _
'"　事務所名稱:" & vbCrLf & _
'"　電話" & vbCrLf & _
'"　傳真" & vbCrLf & _
'"　Email" & vbCrLf & _
'"　地址" & vbCrLf & vbCrLf & _
'"以上謝謝您的支持與合作!" & vbCrLf & _
'"如有任何問題 請隨時聯絡!" & vbCrLf & vbCrLf & _
'"財務處　楊瑞婷  分機 545" & vbCrLf & _
'"台一國際專利法律事務所" & vbCrLf & _
'"台北市長安東路2段112號9樓" & vbCrLf & _
'"電話：０２－２５０６１０２３" & vbCrLf & _
'"傳真：０２－２５０６８１４７" & vbCrLf
'
'         PUB_SendMail "71006", strMail, "", strSubject, _
'                      strContext, , , , , , , "71006", , , , False
'         DoEvents
'         If bolMailSendOk = False Then
'            '失敗結束
'            strSql = "update r100104 set r01001='N'" & _
'                     " where ID='M31' and " & IIf(str01002 <> "", "r01002='" & str01002 & "'", "r01005='" & str01005 & "'")
'            cnnConnection.Execute strSql
'         Else
'            strSql = "update r100104 set r01001='Y'" & _
'                     " where ID='M31' and " & IIf(str01002 <> "", "r01002='" & str01002 & "'", "r01005='" & str01005 & "'")
'            cnnConnection.Execute strSql
'         End If
'         adoRecordset.MoveNext
'      Loop
'   End If
'   MsgBox "寄發郵件,已結束!"
End Sub

Private Sub Command1_Click(Index As Integer)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim ii As Integer
Dim jj As Integer

Select Case Index
Case 0 '尋找
    InitialGrid
    If Me.Text1(0).Text = "" Then MsgBox "請輸入舊編號!!!": Exit Sub
    If Me.Text1(0).Text <> "" Then Me.Text1(0).Text = Left(Me.Text1(0).Text & "000000000", 9)
    If Me.Text1(1).Text <> "" Then Me.Text1(1).Text = Left(Me.Text1(1).Text & "000000000", 9)
    Screen.MousePointer = vbHourglass
                                      StrSQLa = "Select '' as V ,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as 本所案號 FROM PATENT WHERE PA01||PA02||PA03||PA04>='" & Me.Text1(2).Text & "' AND PA01||PA02||PA03||PA04<='" & Me.Text1(3).Text & "' AND PA26='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as 本所案號 FROM PATENT WHERE PA01||PA02||PA03||PA04>='" & Me.Text1(2).Text & "' AND PA01||PA02||PA03||PA04<='" & Me.Text1(3).Text & "' AND PA27='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as 本所案號 FROM PATENT WHERE PA01||PA02||PA03||PA04>='" & Me.Text1(2).Text & "' AND PA01||PA02||PA03||PA04<='" & Me.Text1(3).Text & "' AND PA28='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as 本所案號 FROM PATENT WHERE PA01||PA02||PA03||PA04>='" & Me.Text1(2).Text & "' AND PA01||PA02||PA03||PA04<='" & Me.Text1(3).Text & "' AND PA29='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as 本所案號 FROM PATENT WHERE PA01||PA02||PA03||PA04>='" & Me.Text1(2).Text & "' AND PA01||PA02||PA03||PA04<='" & Me.Text1(3).Text & "' AND PA30='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as 本所案號 FROM PATENT WHERE PA01||PA02||PA03||PA04>='" & Me.Text1(2).Text & "' AND PA01||PA02||PA03||PA04<='" & Me.Text1(3).Text & "' AND PA75='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as 本所案號 FROM TRADEMARK WHERE TM01||TM02||TM03||TM04>='" & Me.Text1(2).Text & "' AND TM01||TM02||TM03||TM04<='" & Me.Text1(3).Text & "' AND TM23='" & Me.Text1(0).Text & "' "
    'Add By Sindy 2011/2/21
    StrSQLa = StrSQLa & " UNION Select '' as V ,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as 本所案號 FROM TRADEMARK WHERE TM01||TM02||TM03||TM04>='" & Me.Text1(2).Text & "' AND TM01||TM02||TM03||TM04<='" & Me.Text1(3).Text & "' AND TM78='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as 本所案號 FROM TRADEMARK WHERE TM01||TM02||TM03||TM04>='" & Me.Text1(2).Text & "' AND TM01||TM02||TM03||TM04<='" & Me.Text1(3).Text & "' AND TM79='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as 本所案號 FROM TRADEMARK WHERE TM01||TM02||TM03||TM04>='" & Me.Text1(2).Text & "' AND TM01||TM02||TM03||TM04<='" & Me.Text1(3).Text & "' AND TM80='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as 本所案號 FROM TRADEMARK WHERE TM01||TM02||TM03||TM04>='" & Me.Text1(2).Text & "' AND TM01||TM02||TM03||TM04<='" & Me.Text1(3).Text & "' AND TM81='" & Me.Text1(0).Text & "' "
    '2011/2/21 End
    StrSQLa = StrSQLa & " UNION Select '' as V ,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as 本所案號 FROM TRADEMARK WHERE TM01||TM02||TM03||TM04>='" & Me.Text1(2).Text & "' AND TM01||TM02||TM03||TM04<='" & Me.Text1(3).Text & "' AND TM44='" & Me.Text1(0).Text & "' "
    'Add By Sindy 2011/2/21
    StrSQLa = StrSQLa & " UNION Select '' as V ,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as 本所案號 FROM SERVICEPRACTICE WHERE SP01||SP02||SP03||SP04>='" & Me.Text1(2).Text & "' AND SP01||SP02||SP03||SP04<='" & Me.Text1(3).Text & "' AND SP08='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as 本所案號 FROM SERVICEPRACTICE WHERE SP01||SP02||SP03||SP04>='" & Me.Text1(2).Text & "' AND SP01||SP02||SP03||SP04<='" & Me.Text1(3).Text & "' AND SP58='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as 本所案號 FROM SERVICEPRACTICE WHERE SP01||SP02||SP03||SP04>='" & Me.Text1(2).Text & "' AND SP01||SP02||SP03||SP04<='" & Me.Text1(3).Text & "' AND SP59='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as 本所案號 FROM SERVICEPRACTICE WHERE SP01||SP02||SP03||SP04>='" & Me.Text1(2).Text & "' AND SP01||SP02||SP03||SP04<='" & Me.Text1(3).Text & "' AND SP65='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as 本所案號 FROM SERVICEPRACTICE WHERE SP01||SP02||SP03||SP04>='" & Me.Text1(2).Text & "' AND SP01||SP02||SP03||SP04<='" & Me.Text1(3).Text & "' AND SP66='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as 本所案號 FROM SERVICEPRACTICE WHERE SP01||SP02||SP03||SP04>='" & Me.Text1(2).Text & "' AND SP01||SP02||SP03||SP04<='" & Me.Text1(3).Text & "' AND SP26='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as 本所案號 FROM LAWCASE WHERE LC01||LC02||LC03||LC04>='" & Me.Text1(2).Text & "' AND LC01||LC02||LC03||LC04<='" & Me.Text1(3).Text & "' AND LC11='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as 本所案號 FROM LAWCASE WHERE LC01||LC02||LC03||LC04>='" & Me.Text1(2).Text & "' AND LC01||LC02||LC03||LC04<='" & Me.Text1(3).Text & "' AND LC43='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as 本所案號 FROM LAWCASE WHERE LC01||LC02||LC03||LC04>='" & Me.Text1(2).Text & "' AND LC01||LC02||LC03||LC04<='" & Me.Text1(3).Text & "' AND LC44='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as 本所案號 FROM LAWCASE WHERE LC01||LC02||LC03||LC04>='" & Me.Text1(2).Text & "' AND LC01||LC02||LC03||LC04<='" & Me.Text1(3).Text & "' AND LC45='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as 本所案號 FROM LAWCASE WHERE LC01||LC02||LC03||LC04>='" & Me.Text1(2).Text & "' AND LC01||LC02||LC03||LC04<='" & Me.Text1(3).Text & "' AND LC46='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as 本所案號 FROM LAWCASE WHERE LC01||LC02||LC03||LC04>='" & Me.Text1(2).Text & "' AND LC01||LC02||LC03||LC04<='" & Me.Text1(3).Text & "' AND LC22='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,HC01||'-'||HC02||'-'||HC03||'-'||HC04 as 本所案號 FROM HIRECASE WHERE HC01||HC02||HC03||HC04>='" & Me.Text1(2).Text & "' AND HC01||HC02||HC03||HC04<='" & Me.Text1(3).Text & "' AND HC05='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,HC01||'-'||HC02||'-'||HC03||'-'||HC04 as 本所案號 FROM HIRECASE WHERE HC01||HC02||HC03||HC04>='" & Me.Text1(2).Text & "' AND HC01||HC02||HC03||HC04<='" & Me.Text1(3).Text & "' AND HC24='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,HC01||'-'||HC02||'-'||HC03||'-'||HC04 as 本所案號 FROM HIRECASE WHERE HC01||HC02||HC03||HC04>='" & Me.Text1(2).Text & "' AND HC01||HC02||HC03||HC04<='" & Me.Text1(3).Text & "' AND HC25='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,HC01||'-'||HC02||'-'||HC03||'-'||HC04 as 本所案號 FROM HIRECASE WHERE HC01||HC02||HC03||HC04>='" & Me.Text1(2).Text & "' AND HC01||HC02||HC03||HC04<='" & Me.Text1(3).Text & "' AND HC26='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,HC01||'-'||HC02||'-'||HC03||'-'||HC04 as 本所案號 FROM HIRECASE WHERE HC01||HC02||HC03||HC04>='" & Me.Text1(2).Text & "' AND HC01||HC02||HC03||HC04<='" & Me.Text1(3).Text & "' AND HC27='" & Me.Text1(0).Text & "' "
    '2011/2/21 End
    StrSQLa = StrSQLa & " UNION Select '' as V ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號 FROM CASEPROGRESS WHERE CP01||CP02||CP03||CP04>='" & Me.Text1(2).Text & "' AND CP01||CP02||CP03||CP04<='" & Me.Text1(3).Text & "' AND CP55='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號 FROM CASEPROGRESS WHERE CP01||CP02||CP03||CP04>='" & Me.Text1(2).Text & "' AND CP01||CP02||CP03||CP04<='" & Me.Text1(3).Text & "' AND CP56='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號 FROM CASEPROGRESS WHERE CP01||CP02||CP03||CP04>='" & Me.Text1(2).Text & "' AND CP01||CP02||CP03||CP04<='" & Me.Text1(3).Text & "' AND CP72='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " UNION Select '' as V ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號 FROM CASEPROGRESS WHERE CP01||CP02||CP03||CP04>='" & Me.Text1(2).Text & "' AND CP01||CP02||CP03||CP04<='" & Me.Text1(3).Text & "' AND CP44='" & Me.Text1(0).Text & "' "
    StrSQLa = StrSQLa & " Order By 1 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    Set Me.MSHFlexGrid1.Recordset = rsA
    If rsA.RecordCount <= 0 Then
        InitialGrid
        MsgBox "查無資料!!!"
    End If
    Screen.MousePointer = vbDefault
Case 1 '確定
    If Me.Text1(0).Text = "" Then MsgBox "請輸入舊編號!!!": Exit Sub
    If Me.Text1(0).Text <> "" Then Me.Text1(0).Text = Left(Me.Text1(0).Text & "000000000", 9)
    If Me.Text1(1).Text = "" Then MsgBox "請輸入新編號!!!": Exit Sub
    If Me.Text1(1).Text <> "" Then Me.Text1(1).Text = Left(Me.Text1(1).Text & "000000000", 9)
    If Me.MSHFlexGrid1.Rows <= 1 Or Me.MSHFlexGrid1.TextMatrix(1, 1) = "" Then
        MsgBox "無任何資料可更新!!!"
    End If
    Screen.MousePointer = vbHourglass
    UpdateData
    Command1_Click 0
    Screen.MousePointer = vbDefault
Case 2 '結束
'    'Add By Cheng 2004/05/11
'    '新增第二期註冊費下一程序期限
'    Dim strNP01 As String, strNP09 As String, strNP08 As String, strNP10 As String, strNP22 As String
''    Pub_DeleteLogFile
'    strSQLA = "Select * From Trademark Where TM01 In ('FCT', 'T', ' ') And TM10='000' And TM11<20031128 And TM14>=20030901 Order By TM01, TM02, TM03, TM04 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
''        Open App.Path & "\Log_" & App.EXEName & ".doc" For Append As #10
'        While Not rsA.EOF
'            If "" & rsA("TM29").Value = "Y" Then
'                strSQL = "Delete From Nextprogress Where " & ChgNextProgress(rsA("TM01").Value & rsA("TM02").Value & rsA("TM03").Value & rsA("TM04").Value) & " And NP07=716 And NP16='90019' "
'                cnnConnection.Execute strSQL
'            End If
''            strSQLB = "Select Count(*) From Caseprogress Where " & ChgCaseprogress(rsA("TM01").Value & rsA("TM02").Value & rsA("TM03").Value & rsA("TM04").Value) & " And CP10='716' And CP57 Is Null "
''            rsB.CursorLocation = adUseClient
''            rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
''            If Val("" & rsB.Fields(0).Value) <= 0 Then
''                If rsB.State <> adStateClosed Then rsB.Close
''                Set rsB = Nothing
''                strSQLB = "Select Count(*) From Nextprogress Where " & ChgNextProgress(rsA("TM01").Value & rsA("TM02").Value & rsA("TM03").Value & rsA("TM04").Value) & " And NP07=716 "
''                rsB.CursorLocation = adUseClient
''                rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
''                If Val("" & rsB.Fields(0).Value) <= 0 Then
''                    strNP01 = GetNP01(rsA("TM01").Value & rsA("TM02").Value & rsA("TM03").Value & rsA("TM04").Value)
''                    strNP09 = ChangeWDateStringToWString(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(Val("" & rsA("TM14").Value)))))
''                    strNP08 = ChangeWDateStringToWString(DateAdd("d", -2, ChangeWStringToWDateString(Val("" & strNP09))))
''                    strNP10 = PUB_GetAKindSalesNo(rsA("TM01").Value, rsA("TM02").Value, rsA("TM03").Value, rsA("TM04").Value)
''                    strNP22 = GetNextProgressNo()
''                    strSQL = "Insert Into Nextprogress Values ('" & strNP01 & "','" & rsA("TM01") & "','" & rsA("TM02").Value & "','" & rsA("TM03").Value & "','" & rsA("TM04").Value & "', Null, 716," & Val(strNP08) & "," & Val(strNP09) & ",'" & strNP10 & "'," & _
''                                    "Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null," & Val(strNP22) & " ) "
''                    cnnConnection.Execute strSQL
'''                    Print #10, rsA.Fields("TM01").Value & "-" & rsA.Fields("TM02").Value & "-" & rsA.Fields("TM03").Value & "-" & rsA.Fields("TM04").Value & "," & _
''                                    ChangeWStringToTString(rsA("TM11").Value) & "," & ChangeWStringToTString(rsA("TM14").Value) & "," & Left(rsA("TM05").Value, 10) & "," & _
''                                    Left(PUB_GetCustName(rsA("TM01").Value & rsA("TM02").Value & rsA("TM03").Value & rsA("TM04").Value), 10) & "," & _
''                                    ChangeWStringToTString(GetCKindCP05(rsA("TM01").Value & rsA("TM02").Value & rsA("TM03").Value & rsA("TM04").Value))
''                End If
''            End If
''            If rsB.State <> adStateClosed Then rsB.Close
''            Set rsB = Nothing
'            Debug.Print rsA.AbsolutePosition
'            rsA.MoveNext
'        Wend
''        Close #10
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'End
'    'Add By Cheng 2004/03/12
'    '抓CFT巴拿馬下一程序延展NP06 Is Null的資料
'    strSQLA = "Select * From Nextprogress, Trademark where np02=tm01 and np03=tm02 and np04=tm03 " & _
'                    " and np05=tm04 and tm10='103' and tm01='CFT' and np06 is null and np07=102 " & _
'                    " order by 2, 3, 4, 5 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
'        While Not rsA.EOF
'            If "" & rsA("NP09").Value <> "" Then
'                strSQLA = "Update Nextprogress Set NP08=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(rsA("NP09").Value))) & " Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'                cnnConnection.Execute strSQLA
'            End If
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'End
    'Add By Cheng 2004/02/24
'    On Error GoTo Error1
'    strSQLA = "Select SP01, SP02, SP03, SP04, SP05, SP06, SP07 From Servicepractice Where SP01 In ('TS','S') And SP06 Is Not Null Or SP07 Is Not Null Order By SP01, SP02, SP03, SP04 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        rsA.MoveLast
'        rsA.MoveFirst
'        While Not rsA.EOF
'            rsA("SP05").Value = Trim("" & rsA("SP05").Value & IIf("" & rsA("SP06").Value <> "", " ", "") & rsA("SP06").Value & IIf("" & rsA("SP07").Value <> "", " ", "") & rsA("SP07").Value)
'            rsA("SP06").Value = Null
'            rsA("SP07").Value = Null
'            rsA.UPDATE
'            Debug.Print rsA.AbsolutePosition & " / " & rsA.RecordCount
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'
'    On Error GoTo Error2
'    strSQLA = "Select CP01, CP02, CP03, CP04, CP09, CP37, CP38, CP39 From Caseprogress Where CP01 In ('TS','S') And (CP38 Is Not Null Or CP39 Is Not Null) Order By CP01, CP02, CP03, CP04, CP09 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        rsA.MoveLast
'        rsA.MoveFirst
'        While Not rsA.EOF
'            rsA("CP37").Value = Trim("" & rsA("CP37").Value & IIf("" & rsA("CP38").Value <> "", " ", "") & rsA("CP38").Value & IIf("" & rsA("CP39").Value <> "", " ", "") & rsA("CP39").Value)
'            rsA("CP38").Value = Null
'            rsA("CP39").Value = Null
'            rsA.UPDATE
'            Debug.Print rsA.AbsolutePosition & " / " & rsA.RecordCount
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'
'    On Error GoTo Error2
'    strSQLA = "Select CP01, CP02, CP03, CP04, CP09, CE41, CE42, CE43 From Caseprogress, ChangeEvent Where CP09=CE01 And CP01 In ('TS','S') And (CE42 Is Not Null Or CE43 Is Not Null) Order By CP01, CP02, CP03, CP04, CP09 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        rsA.MoveLast
'        rsA.MoveFirst
'        While Not rsA.EOF
'            rsA("CE41").Value = Trim("" & rsA("CE41").Value & IIf("" & rsA("CE42").Value <> "", " ", "") & rsA("CE42").Value & IIf("" & rsA("CE43").Value <> "", " ", "") & rsA("CE43").Value)
'            rsA("CE42").Value = Null
'            rsA("CE43").Value = Null
'            rsA.UPDATE
'            Debug.Print rsA.AbsolutePosition & " / " & rsA.RecordCount
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing

''    'Add By Cheng 2004/02/16
'    strSQLA = "Select * From Caseprogress, Engineerprogress, Casepropertymap, Staff Where CP09=EP02 And EP05=ST01 And CP01=CPM01 And CP10=CPM02 And EP06 Is Not Null And CP48 Is Null And CP01 In ('CFP', 'P') And CP09<'C' And CP27 Is Null Order By EP05 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
'        rsA.MoveFirst
'        Open App.Path & "\Log_" & App.EXEName & ".doc" For Append As #10
'        While Not rsA.EOF
'            strSQLB = "Select * From Casemap, CaseProgress, Casepropertymap Where CM05=CP01 And CM06=CP02 And CM07=CP03 And CM08=CP04 And CP01=CPM01 And CP10=CPM02 And CM01='" & rsA("CP01").Value & "' And CM02='" & rsA("CP02").Value & "' And CM03='" & rsA("CP03").Value & "' And CM04='" & rsA("CP04").Value & "' And CP01 In ('P') And CP09<'C' Order By CM05, CM06, CM07, CM08 "
'            rsB.CursorLocation = adUseClient
'            rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsB.RecordCount > 0 Then
'                While Not rsB.EOF
'                    Print #10, "" & rsA.Fields("CP01").Value & "-" & rsA.Fields("CP02").Value & "-" & rsA.Fields("CP03").Value & "-" & rsA.Fields("CP04").Value & "　" & rsA("CP09").Value & "　" & rsA("CP10").Value & rsA("CPM03").Value & "　" & rsA("EP06").Value & "　" & rsA("EP05").Value & rsA("ST02").Value & "　" & rsB("CP01").Value & "-" & rsB("CP02").Value & "-" & rsB("CP03").Value & "-" & rsB("CP04").Value & "　" & rsB("CP10").Value & rsB("CPM03").Value & "　" & rsB("CP27").Value
'                    rsB.MoveNext
'                Wend
'            Else
'                Print #10, "" & rsA.Fields("CP01").Value & "-" & rsA.Fields("CP02").Value & "-" & rsA.Fields("CP03").Value & "-" & rsA.Fields("CP04").Value & "　" & rsA("CP09").Value & "　" & rsA("CP10").Value & rsA("CPM03").Value & "　" & rsA("EP06").Value & "　" & rsA("EP05").Value & rsA("ST02").Value
'            End If
'            If rsB.State <> adStateClosed Then rsB.Close
'            Set rsB = Nothing
'            rsA.MoveNext
'        Wend
'        Close #10
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'End
'    'Add By 200402/13
'    strSQLA = "Select CP01, CP02, CP03, CP04 From Caseprogress Where CP01 In ('CFT','FCT','T','TF') Group By CP01, CP02, CP03, CP04 Order By 1, 2, 3, 4 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
'        Open App.Path & "\Log_" & App.EXEName & ".txt" For Append As #10
'        While Not rsA.EOF
'            If "" & rsA.Fields(0).Value <> "" And "" & rsA.Fields(1).Value <> "" And "" & rsA.Fields(2).Value <> "" And "" & rsA.Fields(3).Value <> "" Then
'                Select Case UCase("" & rsA.Fields(0).Value)
'                Case "P", "CFP", "FCP"
'                    strSQLB = "Select PA01 From Patent Where " & ChgPatent("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    rsB.CursorLocation = adUseClient
'                    rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'                    If rsB.RecordCount <= 0 Then
'                        Print #10, "" & rsA.Fields(0).Value & "-" & rsA.Fields(1).Value & "-" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value
'                    End If
'                    If rsB.State <> adStateClosed Then rsB.Close
'                    Set rsB = Nothing
'                Case "T", "FCT", "CFT", "TF"
'                    strSQLB = "Select TM01 From Trademark Where " & ChgTradeMark("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    rsB.CursorLocation = adUseClient
'                    rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'                    If rsB.RecordCount <= 0 Then
'                        Print #10, "" & rsA.Fields(0).Value & "-" & rsA.Fields(1).Value & "-" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value
'                    End If
'                    If rsB.State <> adStateClosed Then rsB.Close
'                    Set rsB = Nothing
'                Case "L", "CFL", "FCL"
'                    strSQLB = "Select LC01 From Lawcase Where " & ChgLawcase("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    rsB.CursorLocation = adUseClient
'                    rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'                    If rsB.RecordCount <= 0 Then
'                        Print #10, "" & rsA.Fields(0).Value & "-" & rsA.Fields(1).Value & "-" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value
'                    End If
'                    If rsB.State <> adStateClosed Then rsB.Close
'                    Set rsB = Nothing
'                Case "LA"
'                    strSQLB = "Select HC01 From Hirecase Where " & ChgHirecase("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    rsB.CursorLocation = adUseClient
'                    rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'                    If rsB.RecordCount <= 0 Then
'                        Print #10, "" & rsA.Fields(0).Value & "-" & rsA.Fields(1).Value & "-" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value
'                    End If
'                    If rsB.State <> adStateClosed Then rsB.Close
'                    Set rsB = Nothing
'                Case Else
'                    strSQLB = "Select SP01 From Servicepractice Where " & ChgService("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    rsB.CursorLocation = adUseClient
'                    rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'                    If rsB.RecordCount <= 0 Then
'                        Print #10, "" & rsA.Fields(0).Value & "-" & rsA.Fields(1).Value & "-" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value
'                    End If
'                    If rsB.State <> adStateClosed Then rsB.Close
'                    Set rsB = Nothing
'                End Select
'            Else
'                Print #10, "" & rsA.Fields(0).Value & "-" & rsA.Fields(1).Value & "-" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value
'            End If
'            rsA.MoveNext
'        Wend
'        Close #10
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'Add By Cheng 2003/11/28
'    strSQLA = "Select * From Finaltextmap where FTM05 LIKE '%|@%'"
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
'        While Not rsA.EOF
'            strSQL = ""
'            For ii = 1 To Len("" & rsA("FTM05").Value)
'                If Mid("" & rsA("FTM05").Value, ii, 2) <> "|@" Then
'                    strSQL = strSQL & Mid("" & rsA("FTM05").Value, ii, 1)
'                Else
'                    strSQLA = "Update FinalTextMap Set FTM05='" & strSQL & "' Where FTM01='" & rsA("FTM01").Value & "' And FTM02='" & rsA("FTM02").Value & "' And FTM03='" & rsA("FTM03").Value & "' And FTM04='" & rsA("FTM04").Value & "'"
'                    cnnConnection.Execute strSQLA
'                    Exit For
'                End If
'            Next ii
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'Add By Cheng 2003/11/24
'    strSQLA = "Select NP01, NP07, NP22, NP02, NP03, NP04, NP05 From NextProgress Where NP06 Is Null And NP08>=" & strSrvDate(1)
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
''    cnnConnection.BeginTrans
'    While Not rsA.EOF
'        strSQLA = ""
'        Select Case "" & rsA("NP02").Value
'        Case "FCP", "FG"
'            Select Case "" & rsA("NP07").Value
'            Case "997", "998", "411"
'                strSQLA = ""
'            Case Else
'                strSQLA = "Update NextProgress Set NP10='" & PUB_GetFCPSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'            End Select
'        Case "FCT"
'            Select Case "" & rsA("NP07").Value
'            Case "997", "998", "305"
'                strSQLA = ""
'            Case Else
'                strSQLA = "Update NextProgress Set NP10='" & PUB_GetFCTSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'            End Select
'        Case Else
'            Select Case "" & rsA("NP02").Value
'            Case "CFT", "CFC", "S", "T", "TB", "TC", "TD", "TF", "TM", "TR", "TS", "TT"
'                Select Case "" & rsA("NP07").Value
'                Case "997", "998", "305"
'                    strSQLA = ""
'                Case Else
'                    strSQLA = "Update NextProgress Set NP10='" & PUB_GetAKindSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'                End Select
'            Case Else
'                strSQLA = "Update NextProgress Set NP10='" & PUB_GetAKindSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'            End Select
'        End Select
'        If strSQLA <> "" Then cnnConnection.Execute strSQLA
'        rsA.MoveNext
'    Wend
''    cnnConnection.CommitTrans
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'Add By Cheng 2003/11/11
'    On Error GoTo Error1
'    strSQLA = "Select TM01, TM02, TM03, TM04, TM05, TM06, TM07 From Trademark Where TM06 Is Not Null Or TM07 Is Not Null Order By TM01, TM02, TM03, TM04 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        rsA.MoveLast
'        rsA.MoveFirst
'        While Not rsA.EOF
'            rsA("TM05").Value = Trim("" & rsA("TM05").Value & IIf("" & rsA("TM06").Value <> "", " ", "") & rsA("TM06").Value & IIf("" & rsA("TM07").Value <> "", " ", "") & rsA("TM07").Value)
'            rsA("TM06").Value = Null
'            rsA("TM07").Value = Null
'            rsA.Update
'            Debug.Print rsA.AbsolutePosition & " / " & rsA.RecordCount
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'
'    On Error GoTo Error2
'    strSQLA = "Select CP01, CP02, CP03, CP04, CP09, CP37, CP38, CP39 From Caseprogress Where CP01 In ('T','FCT','CFT','TF') And (CP38 Is Not Null Or CP39 Is Not Null) Order By CP01, CP02, CP03, CP04, CP09 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        rsA.MoveLast
'        rsA.MoveFirst
'        While Not rsA.EOF
'            rsA("CP37").Value = Trim("" & rsA("CP37").Value & IIf("" & rsA("CP38").Value <> "", " ", "") & rsA("CP38").Value & IIf("" & rsA("CP39").Value <> "", " ", "") & rsA("CP39").Value)
'            rsA("CP38").Value = Null
'            rsA("CP39").Value = Null
'            rsA.Update
'            Debug.Print rsA.AbsolutePosition & " / " & rsA.RecordCount
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'
'    On Error GoTo Error2
'    strSQLA = "Select CP01, CP02, CP03, CP04, CP09, CE41, CE42, CE43 From Caseprogress, ChangeEvent Where CP09=CE01 And CP01 In ('T','FCT','CFT','TF') And (CE42 Is Not Null Or CE43 Is Not Null) Order By CP01, CP02, CP03, CP04, CP09 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        rsA.MoveLast
'        rsA.MoveFirst
'        While Not rsA.EOF
'            rsA("CE41").Value = Trim("" & rsA("CE41").Value & IIf("" & rsA("CE42").Value <> "", " ", "") & rsA("CE42").Value & IIf("" & rsA("CE43").Value <> "", " ", "") & rsA("CE43").Value)
'            rsA("CE42").Value = Null
'            rsA("CE43").Value = Null
'            rsA.Update
'            Debug.Print rsA.AbsolutePosition & " / " & rsA.RecordCount
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'Add By Cheng 2004/04/16
'    '轉南所之分所案號
'    Dim DAOWS As DAO.Workspace
'    Dim DAODB As DAO.Database
'    Dim CP1 As DAO.Recordset
'    Dim strExc(1) As String
'    Dim strCaseNo As String '本所案號
'    Dim strCaseNo1 As String '本所案號
'    Dim arrCaseNo
'    Dim strSystemKind As String '系統類別
'    Dim strN_NO As String
'    Dim strS_NO As String
'    strN_NO = ""
'    strS_NO = ""
'    Pub_DeleteLogFile
'    Set DAOWS = DBEngine.Workspaces(0)
'    Set DAODB = DAOWS.OpenDatabase("C:\", False, False, "Dbase 5.0")
'    strExc(1) = "Select N_NO, S_NO From S_No Order By N_NO, S_NO "
'    Set CP1 = DAODB.OpenRecordset(strExc(1))
'    Do While Not CP1.EOF
'        '若有本所案號及分所案號
'        If "" & CP1("N_NO").Value <> "" And "" & CP1("S_NO").Value <> "" Then
'            '若本所案號不同
'            If strN_NO <> "" & CP1("N_NO").Value Then
'                strN_NO = "" & CP1("N_NO").Value
'                '檢查本所案號
'                If Left("" & CP1("N_NO").Value, 1) <> "A" And Left("" & CP1("N_NO").Value, 1) <> "B" And Left("" & CP1("N_NO").Value, 1) <> "C" And Left("" & CP1("N_NO").Value, 1) <> "D" And _
'                    Left("" & CP1("N_NO").Value, 1) <> "L" And Left("" & CP1("N_NO").Value, 1) <> "P" And Left("" & CP1("N_NO").Value, 1) <> "T" And Left("" & CP1("N_NO").Value, 1) <> "W" And _
'                    Left("" & CP1("N_NO").Value, 2) <> "TC" And Left("" & CP1("N_NO").Value, 3) <> "CFC" And Left("" & CP1("N_NO").Value, 3) <> "CFL" And Left("" & CP1("N_NO").Value, 3) <> "CFP" And _
'                    Left("" & CP1("N_NO").Value, 3) <> "CFT" Then
'                    Pub_WriteSysLog "本所案號系統類別錯誤,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("S_NO").Value
'                    GoTo NextRec
'                Else
'                    strCaseNo = ""
'                    strCaseNo1 = ""
'                    For ii = 1 To Len("" & CP1("N_NO").Value)
'                        If Mid("" & CP1("N_NO").Value, ii, 1) < "A" Then
'                            Exit For
'                        Else
'                            strCaseNo1 = strCaseNo1 & Mid("" & CP1("N_NO").Value, ii, 1)
'                        End If
'                    Next ii
'                    strCaseNo1 = strCaseNo1 & "-" & Right("" & CP1("N_NO").Value, Len("" & CP1("N_NO").Value) - Len(strCaseNo1))
'                    arrCaseNo = Split(strCaseNo1, "-")
'                    For ii = LBound(arrCaseNo) To UBound(arrCaseNo)
'                        If ii = 1 Then
'                            arrCaseNo(ii) = Right("000000" & arrCaseNo(ii), 6)
'                        ElseIf ii = 2 Then
'                            arrCaseNo(ii) = Right("0" & arrCaseNo(ii), 1)
'                        ElseIf ii = 3 Then
'                            arrCaseNo(ii) = Right("00" & arrCaseNo(ii), 2)
'                        End If
'                    Next ii
'                    strSystemKind = ""
'                    For jj = LBound(arrCaseNo) To UBound(arrCaseNo)
'                        If jj = 0 Then
'                            Select Case arrCaseNo(0)
'                            Case "A"
'                                arrCaseNo(0) = "LA"
'                            Case "B"
'                                arrCaseNo(0) = "TB"
'                            Case "C"
'                                arrCaseNo(0) = "TC"
'                            Case "D"
'                                arrCaseNo(0) = "TD"
'                            Case "W"
'                                arrCaseNo(0) = "L"
'                            End Select
'                            For ii = 1 To Len(arrCaseNo(jj))
'                                If Mid(arrCaseNo(jj), ii, 1) < "A" Then
'                                    If Len(Mid(arrCaseNo(jj), ii, Len(arrCaseNo(jj)) - ii + 1)) > 6 Then
'                                        Pub_WriteSysLog "本所案號系統類別錯誤,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("S_NO").Value
'                                        GoTo NextRec
'                                    Else
'                                        strCaseNo = strCaseNo & Format(Mid(arrCaseNo(jj), ii, Len(arrCaseNo(jj)) - ii + 1), "000000")
'                                        Exit For
'                                    End If
'                                Else
'                                    strCaseNo = strCaseNo & Mid(arrCaseNo(jj), ii, 1)
'                                    strSystemKind = strSystemKind & Mid(arrCaseNo(jj), ii, 1)
'                                End If
'                            Next ii
'                        ElseIf jj = 1 Then
'                            strCaseNo = strCaseNo & arrCaseNo(jj)
'                        ElseIf jj = 2 Then
'                            '2005/6/13 MODIFY BY SONIA T97258-1會變成T097258010
'                            'strCaseNo = strCaseNo & Format(arrCaseNo(jj), "00")
'                            strCaseNo = strCaseNo & Format(arrCaseNo(jj), "0")
'                            '2005/6/13 END
'                        End If
'                    Next jj
'                End If
'
'                strSQLA = "Select PA01, PA02, PA03, PA04, PA47 From Patent Where " & ChgPatent(strCaseNo)
'                strSQLA = strSQLA & " Union Select TM01, TM02, TM03, TM04, TM34 From Trademark Where " & ChgTradeMark(strCaseNo)
'                strSQLA = strSQLA & " Union Select LC01, LC02, LC03, LC04, LC16 From Lawcase Where " & ChgLawcase(strCaseNo)
'                strSQLA = strSQLA & " Union Select HC01, HC02, HC03, HC04, HC07 From Hirecase Where " & ChgHirecase(strCaseNo)
'                strSQLA = strSQLA & " Union Select SP01, SP02, SP03, SP04, SP28 From Servicepractice Where " & ChgService(strCaseNo)
'                rsA.CursorLocation = adUseClient
'                rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'                If rsA.RecordCount > 0 Then
'                    '若已有分所案號,且分所案號不同
'                    If "" & rsA.Fields(4).Value <> "" And "" & rsA.Fields(4).Value <> "" & CP1("S_NO").Value Then
'                        Pub_WriteSysLog "分所案號已存在,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("S_NO").Value
'                        GoTo NextRec
'                    '若已有分所案號,且分所案號相同(不需更新)
'                    ElseIf "" & rsA.Fields(4).Value <> "" And "" & rsA.Fields(4).Value = "" & CP1("S_NO").Value Then
'                        GoTo NextRec
'                    End If
'                Else
'                    Pub_WriteSysLog "無此本所案號,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("S_NO").Value
'                    GoTo NextRec
'                End If
'                Select Case strSystemKind
'                Case "P", "CFP", "FCP"
'                    strSQLA = "Update Patent Set PA47='" & CP1("S_NO").Value & "' Where " & ChgPatent("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "T", "TF", "CFT", "FCT"
'                    strSQLA = "Update Trademark Set TM34='" & CP1("S_NO").Value & "' Where " & ChgTradeMark("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "CFL", "FCL", "L"
'                    strSQLA = "Update Lawcase Set LC16='" & CP1("S_NO").Value & "' Where " & ChgLawcase("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "LA"
'                    strSQLA = "Update Hirecase Set HC07='" & CP1("S_NO").Value & "' Where " & ChgHirecase("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case Else
'                    strSQLA = "Update Servicepractice Set SP28='" & CP1("S_NO").Value & "' Where " & ChgService("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                End Select
'            '若本所案號相同
'            Else
'                Pub_WriteSysLog "本所案號重覆,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("S_NO").Value
'                GoTo NextRec
'            End If
'        End If
'NextRec:
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
'        Debug.Print CP1.AbsolutePosition
'        CP1.MoveNext
'    Loop
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'Add By Cheng 2003/09/12
'    '轉台中之分所案號
'    Dim DAOWS As DAO.Workspace
'    Dim DAODB As DAO.Database
'    Dim CP1 As DAO.Recordset
'    Dim strExc(1) As String
'    strSQLA = "Select PA01, PA02, PA03, PA04, PA47 From Patent Where PA47 Is Not Null "
'    strSQLA = strSQLA & " Union Select TM01, TM02, TM03, TM04, TM34 From Trademark Where TM34 Is Not Null  "
'    strSQLA = strSQLA & " Union Select LC01, LC02, LC03, LC04, LC16 From Lawcase Where LC16 Is Not Null  "
'    strSQLA = strSQLA & " Union Select HC01, HC02, HC03, HC04, HC07 From Hirecase Where HC07 Is Not Null  "
'    strSQLA = strSQLA & " Union Select SP01, SP02, SP03, SP04, SP28 From Servicepractice Where SP28 Is Not Null  "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
'        While Not rsA.EOF
'            If Left("" & rsA.Fields(4).Value, 1) >= "0" And Left("" & rsA.Fields(4).Value, 1) <= "9" Then
'                Select Case "" & rsA.Fields(0).Value
'                Case "P", "CFP", "FCP"
'                    strSQLA = "Update Patent Set PA47=PA01||PA47 Where " & ChgPatent("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "T", "TF", "CFT", "FCT"
'                    strSQLA = "Update Trademark Set TM34=TM01||TM34 Where " & ChgTradeMark("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "CFL", "FCL", "L"
'                    strSQLA = "Update Lawcase Set LC16=LC01||LC16 Where " & ChgLawcase("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "LA"
'                    strSQLA = "Update Hirecase Set HC07=HC01||HC07 Where " & ChgHirecase("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case Else
'                    strSQLA = "Update Servicepractice Set SP28=SP01||SP28 Where " & ChgService("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                End Select
'            End If
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    Dim strCaseNo As String '本所案號
'    Dim arrCaseNO
'    Dim strSystemKind As String '系統類別
'    Dim strN_NO As String
'    Dim strM_NO As String
'    strN_NO = ""
'    strM_NO = ""
'    Pub_DeleteLogFile
'    Set DAOWS = DBEngine.Workspaces(0)
'    Set DAODB = DAOWS.OpenDatabase("C:\中所No.dbf", False, False, "Dbase 5.0")
'    strExc(1) = "Select N_NO, M_NO From No Order By N_NO, M_NO "
'    Set CP1 = DAODB.OpenRecordset(strExc(1))
'    Do While Not CP1.EOF
'        '若有本所案號及分所案號
'        If "" & CP1("N_NO").Value <> "" And "" & CP1("M_NO").Value <> "" Then
'            '若本所案號不同
'            If strN_NO <> "" & CP1("N_NO").Value Then
'                strN_NO = "" & CP1("N_NO").Value
'                '檢查本所案號
'                If "" & CP1("N_NO").Value < "A" Then
'                    Pub_WriteSysLog "本所案號系統類別錯誤,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("M_NO").Value
'                    GoTo NextRec
'                Else
'                    strCaseNo = ""
'                    arrCaseNO = Split("" & CP1("N_NO").Value, "-")
'                    strSystemKind = ""
'                    For jj = LBound(arrCaseNO) To UBound(arrCaseNO)
'                        If jj = 0 Then
'                            For ii = 1 To Len(arrCaseNO(jj))
'                                If Mid(arrCaseNO(jj), ii, 1) < "A" Then
'                                    If Len(Mid(arrCaseNO(jj), ii, Len(arrCaseNO(jj)) - ii + 1)) > 6 Then
'                                        Pub_WriteSysLog "本所案號系統類別錯誤,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("M_NO").Value
'                                        GoTo NextRec
'                                    Else
'                                        strCaseNo = strCaseNo & Format(Mid(arrCaseNO(jj), ii, Len(arrCaseNO(jj)) - ii + 1), "000000")
'                                        Exit For
'                                    End If
'                                Else
'                                    strCaseNo = strCaseNo & Mid(arrCaseNO(jj), ii, 1)
'                                    strSystemKind = strSystemKind & Mid(arrCaseNO(jj), ii, 1)
'                                End If
'                            Next ii
'                        ElseIf jj = 1 Then
'                            strCaseNo = strCaseNo & arrCaseNO(jj)
'                        ElseIf jj = 2 Then
'                            strCaseNo = strCaseNo & Format(arrCaseNO(jj), "00")
'                        End If
'                    Next jj
'                End If
'
'                strSQLA = "Select PA01, PA02, PA03, PA04, PA47 From Patent Where " & ChgPatent(strCaseNo)
'                strSQLA = strSQLA & " Union Select TM01, TM02, TM03, TM04, TM34 From Trademark Where " & ChgTradeMark(strCaseNo)
'                strSQLA = strSQLA & " Union Select LC01, LC02, LC03, LC04, LC16 From Lawcase Where " & ChgLawcase(strCaseNo)
'                strSQLA = strSQLA & " Union Select HC01, HC02, HC03, HC04, HC07 From Hirecase Where " & ChgHirecase(strCaseNo)
'                strSQLA = strSQLA & " Union Select SP01, SP02, SP03, SP04, SP28 From Servicepractice Where " & ChgService(strCaseNo)
'                rsA.CursorLocation = adUseClient
'                rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'                If rsA.RecordCount > 0 Then
'                    '若已有分所案號,且分所案號不同
'                    If "" & rsA.Fields(4).Value <> "" And "" & rsA.Fields(4).Value <> "" & CP1("M_NO").Value Then
'                        Pub_WriteSysLog "分所案號已存在,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("M_NO").Value
'                        GoTo NextRec
'                    '若已有分所案號,且分所案號相同(不需更新)
'                    ElseIf "" & rsA.Fields(4).Value <> "" And "" & rsA.Fields(4).Value = "" & CP1("M_NO").Value Then
'                        GoTo NextRec
'                    End If
'                Else
'                    Pub_WriteSysLog "無此本所案號,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("M_NO").Value
'                    GoTo NextRec
'                End If
'                Select Case strSystemKind
'                Case "P", "CFP", "FCP"
'                    strSQLA = "Update Patent Set PA47='" & CP1("M_NO").Value & "' Where " & ChgPatent("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "T", "TF", "CFT", "FCT"
'                    strSQLA = "Update Trademark Set TM34='" & CP1("M_NO").Value & "' Where " & ChgTradeMark("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "CFL", "FCL", "L"
'                    strSQLA = "Update Lawcase Set LC16='" & CP1("M_NO").Value & "' Where " & ChgLawcase("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case "LA"
'                    strSQLA = "Update Hirecase Set HC07='" & CP1("M_NO").Value & "' Where " & ChgHirecase("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                Case Else
'                    strSQLA = "Update Servicepractice Set SP28='" & CP1("M_NO").Value & "' Where " & ChgService("" & rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value)
'                    cnnConnection.Execute strSQLA
'                End Select
'            '若本所案號相同
'            Else
'                Pub_WriteSysLog "本所案號重覆,本所案號," & CP1("N_NO").Value & ",分所案號," & CP1("M_NO").Value
'                GoTo NextRec
'            End If
'        End If
'NextRec:
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
'        Debug.Print CP1.AbsolutePosition
'        CP1.MoveNext
'    Loop
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'Add By Cheng 2003/09/12
'    Dim strPromoteDate As String
'    strSQLA = "Select * From EngineerProgress, CaseProgress Where EP02=CP09 And EP06 Is Not Null And CP48 Is Null And CP27 Is Null "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic
'    If rsA.RecordCount > 0 Then
'        While Not rsA.EOF
'            strSQLB = "Select NVL(CF04,0) From CaseProgress, Patent, Casefee Where CP09='" & rsA("CP09").Value & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and PA01=cf01(+) and pa09=cf02(+) and cp10=cf03 "
'            rsB.CursorLocation = adUseClient
'            rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsB.RecordCount > 0 Then
'                If Val("" & rsB.Fields(0).Value) > 0 Then
'                    strPromoteDate = CompWorkDay(rsB.Fields(0).Value, rsA("EP06").Value, 0)
'                    '若有計算出承辦期限且有本所期限
'                    If strPromoteDate <> "" And "" & rsA("CP06").Value <> "" Then
'                        '若承辦期限大於本所期限
                         '2010/8/17 modify by sonia
                         'If strPromoteDate > "" & rsA("CP06").Value Then
'                        If val(strPromoteDate) > "" & val(rsA("CP06").Value) Then
'                            strPromoteDate = "" & rsA("CP06").Value
'                        End If
'                    End If
'                    strSQLA = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09='" & rsA("CP09").Value & "' "
'                    cnnConnection.Execute strSQLA
'                End If
'            End If
'            If rsB.State <> adStateClosed Then rsB.Close
'            Set rsB = Nothing
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'Add By Cheng 2003/09/04
'    strSQLA = "Select NP01, NP07, NP22, NP02, NP03, NP04, NP05 From NextProgress Where NP06 Is Null And NP08>=" & strSrvDate(1)
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
''    cnnConnection.BeginTrans
'    While Not rsA.EOF
'        Select Case "" & rsA("NP02").Value
'        Case "FCP", "FG"
'            strSQLA = "Update NextProgress Set NP10='" & PUB_GetFCPSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'        Case "FCT"
'            strSQLA = "Update NextProgress Set NP10='" & PUB_GetFCTSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'        Case Else
'            strSQLA = "Update NextProgress Set NP10='" & PUB_GetAKindSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'        End Select
'        cnnConnection.Execute strSQLA
'        rsA.MoveNext
'    Wend
''    cnnConnection.CommitTrans
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'Add By Cheng 2003/08/12
'    strSQLA = "Select NP01, NP07, NP22, NP02, NP03, NP04, NP05 From NextProgress Where NP06 Is Null And NP09>=20020701 And NP16 Is Null "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
''    cnnConnection.BeginTrans
'    While Not rsA.EOF
'        Select Case "" & rsA("NP02").Value
'        Case "FCP"
'            strSQLA = "Update NextProgress Set NP10='" & PUB_GetFCPSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'        Case Else
'            strSQLA = "Update NextProgress Set NP10='" & PUB_GetAKindSalesNo(rsA("NP02").Value, rsA("NP03").Value, rsA("NP04").Value, rsA("NP05").Value) & "' Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'        End Select
'        cnnConnection.Execute strSQLA
'        rsA.MoveNext
'    Wend
''    cnnConnection.CommitTrans
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    '2003/06/10
'    strSQLA = "Select * From Acc190 Where a1902>='U090' And a1917 Is Null "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    While Not rsA.EOF
'        rsA("a1917").Value = GetCompany("" & rsA("a1902").Value)
'        rsA.Update
'        Debug.Print rsA.AbsolutePosition
'        rsA.MoveNext
'    Wend
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
    '2003/05/08
'    Dim dblEP35 As Double
'    strSQLA = "Select * From EngineerProgress "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    While Not rsA.EOF
'        '計算承辦天數
'        dblEP35 = 0
'        If "" & rsA("EP07").Value <> "" And "" & rsA("EP06").Value <> "" Then
'             dblEP35 = Trim(str(GetWorkDay(rsA("EP07").Value, rsA("EP06").Value)))
'        Else
'            If "" & rsA("EP09").Value <> "" And "" & rsA("EP06").Value <> "" Then
'                dblEP35 = Trim(str(GetWorkDay(rsA("EP09").Value, rsA("EP06").Value)))
'            Else
'                dblEP35 = 0
'            End If
'        End If
'        strSQLA = "Update Engineerprogress Set EP35=" & IIf(dblEP35 = 0, "Null", dblEP35) & " Where EP02='" & rsA("EP02").Value & "' "
'        cnnConnection.Execute strSQLA
'        Debug.Print rsA.AbsolutePosition
'        rsA.MoveNext
'    Wend
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
    
    '2003/04/18
'    Dim strAXF15 As String
''    strSQLA = "select * from acc150, acc151 where a1501=axf01 and axf15=0 and axf01>='U092' "
'    strSQLA = "select * from acc150, acc151 where a1501=axf01 and axf03>='P' and axf03<='P9' and axf01>='U092' and axf14 is null "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.RecordCount > 0 Then
'        rsA.MoveFirst
'        While Not rsA.EOF
'            strSQLB = "Select A2103 From ACC210 Where A2101 <= " & rsA("A1502").Value & " AND A2102 = '" & rsA("A1505").Value & "' Order By A2101 Desc "
'            rsB.CursorLocation = adUseClient
'            rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsB.RecordCount > 0 Then
'                strAXF15 = Format(("" & rsA("AXF04").Value) * Val("" & rsB("A2103").Value), "##0.00")
'            Else
'                strAXF15 = Format(("" & rsA("AXF04").Value) * 1, "##0.00")
'            End If
'            If rsB.State <> adStateClosed Then rsB.Close
'            Set rsB = Nothing
'            rsA("AXF15").Value = strAXF15
'            rsA.Update
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
    '2003/03/28
'    strSQLA = "Select * From TradeMark Where substr(TM05,length(TM05),1)=' ' or substr(TM05,length(TM05),1)='　' or substr(TM06,length(TM06),1)=' ' or substr(TM06,length(TM06),1)='　' or substr(TM07,length(TM07),1)=' ' or substr(TM07,length(TM07),1)='　'  "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        While Not rsA.EOF
'            rsA("TM05").Value = Trim("" & rsA("TM05").Value)
'            rsA("TM06").Value = Trim("" & rsA("TM06").Value)
'            rsA("TM07").Value = Trim("" & rsA("TM07").Value)
'            rsA.Update
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    strSQLA = "Select * From Patent Where substr(PA05,length(PA05),1)=' ' or substr(PA05,length(PA05),1)='　' or substr(PA06,length(PA06),1)=' ' or substr(PA06,length(PA06),1)='　' or substr(PA07,length(PA07),1)=' ' or substr(PA07,length(PA07),1)='　'  "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        While Not rsA.EOF
'            rsA("PA05").Value = Trim("" & rsA("PA05").Value)
'            rsA("PA06").Value = Trim("" & rsA("PA06").Value)
'            rsA("PA07").Value = Trim("" & rsA("PA07").Value)
'            rsA.Update
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
    
    '2003/03/24
'    'CFT 中v 英x 日v
'    strSQLA = "Select * From Trademark Where TM01='CFT' And TM05 is not null And TM06 is null And TM07 is not null "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        While Not rsA.EOF
'            If Right("" & rsA("TM05").Value, 2) = "及圖" Then
'                rsA("TM07").Value = "" & rsA("TM07").Value & "及圖"
'                rsA("TM05").Value = Replace("" & rsA("TM05").Value, "及圖", "")
'                rsA.Update
'            End If
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'CFT 中v 英v
'    strSQLA = "Select * From Trademark Where TM01='CFT' And TM05 is not null And TM06 is not null "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        While Not rsA.EOF
'            If Right("" & rsA("TM05").Value, 2) = "及圖" Then
'                rsA("TM06").Value = "" & rsA("TM06") & " & DEVICE"
'                rsA("TM05").Value = Replace("" & rsA("TM05").Value, "及圖", "")
'                rsA.Update
'            End If
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'FCT 中v 英x 日v
'    strSQLA = "Select * From Trademark Where TM01='FCT' And TM05 is not null And TM06 is null And TM07 is not null "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        While Not rsA.EOF
'            If Right("" & rsA("TM05").Value, 2) = "及圖" Then
'                rsA("TM07").Value = "" & rsA("TM07").Value & "及圖"
'                rsA("TM05").Value = Replace("" & rsA("TM05").Value, "及圖", "")
'                rsA.Update
'            End If
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    'FCT 中v 英x
'    strSQLA = "Select * From Trademark Where TM01='FCT' And TM05 is not null And TM06 is not null "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenDynamic, adLockOptimistic
'    If rsA.EOF = False Then
'        While Not rsA.EOF
'            If Right("" & rsA("TM05").Value, 2) = "及圖" Then
'                rsA("TM06").Value = "" & rsA("TM06") & " & DEVICE"
'                rsA("TM05").Value = Replace("" & rsA("TM05").Value, "及圖", "")
'                rsA.Update
'            End If
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
    
'   '2006/10/30 ADD BY SONIA
'   '加掛內商下一程序催審期限
'   Dim strNP01 As String, strNP09 As String, strNP08 As String, strNP10 As String, strNP22 As String
'   StrSQLa = "SELECT A.CP09 AS CP09,A.CP01 AS CP01,A.CP02 AS CP02,A.CP03 AS CP03,A.CP04 AS CP04,TO_CHAR(TO_DATE(A.CP27,'YYYYMMDD') + CF05,'YYYYMMDD') AS NP09,A.CP14 AS CP14 " & _
'             "FROM CASEPROGRESS A,TRADEMARK,CASEFEE,NEXTPROGRESS,CASEPROGRESS B " & _
'             "WHERE A.CP01='T' AND A.CP27>=20030201 AND A.CP27<=20061023 " & _
'             "AND A.CP09=B.CP43(+) AND '306'=B.CP10(+) AND B.CP09 IS NULL " & _
'             "AND A.CP24 IS NULL AND A.CP10 IN ('101','102','103','304','702') AND A.CP57 IS NULL " & _
'             "AND A.CP01=TM01 AND A.CP02=TM02 AND A.CP03=TM03 AND A.CP04=TM04 " & _
'             "AND TM29 IS NULL AND TM22 IS NULL AND A.CP01=CF01 AND TM10=CF02 AND A.CP10=CF03 AND CF05 IS NOT NULL " & _
'             "AND A.CP09=NP01(+) AND 305=NP07(+) AND NP01 IS NULL "
'    rsA.CursorLocation = adUseClient
'    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
'       While Not rsA.EOF
'          strNP22 = GetNextProgressNo()
'          strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                   "VALUES ('" & rsA.Fields("CP09") & "','" & rsA.Fields("CP01") & "','" & rsA.Fields("CP02") & "','" & rsA.Fields("CP03") & "','" & rsA.Fields("CP04") & "','305', " & _
'                    rsA.Fields("NP09") & "," & rsA.Fields("NP09") & ",'" & rsA.Fields("CP14") & "'," & strNP22 & ")"
'          cnnConnection.Execute strSQL
'          rsA.MoveNext
'       Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'    '2006/10/30 End
'
    Unload Me
End Select
Exit Sub
Error1:
    'Debug.Print Err.Description
    'Debug.Print rsA.Fields(0).Value & "-" & rsA.Fields(1).Value & "-" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value
    Resume Next
Error2:
    'Debug.Print Err.Description
    'Debug.Print rsA.Fields(0).Value & "-" & rsA.Fields(1).Value & "-" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value & "-" & rsA.Fields(4).Value
    Resume Next
End Sub

Private Sub Command10_Click()
   'Call ShellExecute(Me.hwnd, "open", "http://www.yahoo.com.tw", "", "", 5) 'SW_SHOW=5
'   WebBrowser1.Navigate "https://accounts.google.com/ServiceLogin?service=mail&passive=true&rm=false&continue=https://mail.google.com/mail/&ss=1&scc=1&ltmpl=default&ltmplcache=2&hl=zh-TW"
'   Do While WebBrowser1.Busy
'      DoEvents
'   Loop
'   With WebBrowser1
'      .Document.All("login").Value = "" ' 帳號
'      .Document.All("passwd").Value = "" ' 密碼
'      .Document.All("submit").Click ' 登入
'      MsgBox .Document("login").Value = "" ' 帳號
'   End With
   Dim myweb As Object
   Set myweb = CreateObject("InternetExplorer.Application")
   With myweb
      .ToolBar = 0
      .Visible = True ' 顯示IE
      '.Navigate "https://www.lativ.com.tw/Home/Login" ' 瀏覽網址
      .Navigate "http://www.sipo.gov.cn/zljs/"
      ' 等待網頁載入完成
      Do While .Busy
         DoEvents
      Loop
      .Document.All("textfield4").Value = "2013.01.16"
      .Document.All("textfield8").Value = "台"
      '.Document.All("login").Value = "sindygirllu" ' 帳號
      '.Document.All("email").Value = "sindygirllu" ' 帳號 login
      '.Document.All("pw").Value = "sindybb1109" ' 密碼 passwd
      '.Document.All("signIn").Click ' 登入 signIn
      '.Document.All("submit").Click     ' 登入
      .Document.All("Submit").Click
      ' 等待網頁載入完成
      Do While .Busy
         DoEvents
      Loop
      MsgBox "完成!!!"
   End With
   
   Set myweb = Nothing ' 釋放IE 物件
'   myweb.Toolbar = 0
'   myweb.Visible = True
'   myweb.Navigate "http://www.yahoo.com.tw"
'   MsgBox "暫停觀察sessionid"
'   myweb.Quit
'   DoEvents
End Sub

'Modfify by Amy 2018/05/04 郵局郵遞區號轉檔(匯入郵局Excel)
Private Sub Command11_Click()
    Dim xlsSalesPoint As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim strFileName As String, strIns As String
    Dim i As Integer, iRow As Double
    Dim strPZD01 As String, strPZD02 As String, strPZD03 As String, strPZD05 As String
    Dim strPZD06 As String, strPZD10 As String, strPZD11 As String
    Dim strPZD04 As String, strPZD04_SP As String
    Dim strContent As String, strTp As String
    Dim strNo As String 'Add by Amy 2019/11/05 增加編號,利於修改資料
   
On Error GoTo flgErr
   
    If txtFileName = "" Then
        MsgBox "檔案不可空白！"
        txtFileName.SetFocus
        Exit Sub
    End If
   
    strFileName = txtFileName
   
   '開檔
    Screen.MousePointer = vbHourglass
    xlsSalesPoint.Workbooks.Open strFileName
    Set wksrpt = xlsSalesPoint.Worksheets(1)
    
    iRow = 2
    strNo = "1" 'Add by Amy 2019/11/05
    cnnConnection.BeginTrans
    cnnConnection.Execute "Delete From Postzipdata_New"
    
    Do While wksrpt.Range("A" & iRow).Value <> MsgText(601)
        strPZD05 = "": strPZD06 = "": strPZD10 = "": strPZD11 = ""
        strPZD01 = Trim(wksrpt.Range("A" & iRow).Value) '郵遞區號
        strPZD02 = Trim(wksrpt.Range("B" & iRow).Value) '縣市
        strPZD03 = Trim(wksrpt.Range("C" & iRow).Value) '鄉鎮
        txtSpecW = Trim(wksrpt.Range("D" & iRow).Value) '街道(用變數抓雖有? InStr(變數,'?')會是0)
        strPZD04 = txtSpecW
        strTp = Trim(wksrpt.Range("E" & iRow).Value) '單雙號碼
        '判斷是否有造字
        If InStr(strPZD04, "?") > 0 Then
            strPZD04_SP = GetZIPSpecWord(strPZD01, strPZD02 & strPZD03 & strPZD04)
            '有未造字資料
            If strPZD04_SP = "" Then
                strContent = strContent & strPZD01 & "　" & strPZD02 & "　" & strPZD03 & "　" & strPZD04 & vbCrLf
            '已存在郵遞區號特殊字對照檔,但經理未造字,不處理
            ElseIf InStr(strPZD04_SP, "?") > 0 Then
            '有造字對應,取代成造字
            Else
                strPZD04 = strPZD04_SP
            End If
        End If
        '設定所別/國別(秀玲2018/05/04 3:04 mail 主旨postzipdata之所別及國籍代號)
        Select Case strPZD02
            Case "基隆市", "新北市", "臺北市", "宜蘭縣", "連江縣", "釣魚臺"
                strPZD10 = "1" '所別
                strPZD11 = "001" '國別
            Case "桃園市", "新竹市", "新竹縣"
                strPZD10 = "1" '北
                strPZD11 = "002"
            'Modify by Amy 2019/11/08 原:strPZD10 = "2" /strPZD11 = "004"
            Case "苗栗縣"
                strPZD10 = "2" '中
                strPZD11 = "002"
            Case "臺中市", "南投縣"
                strPZD10 = "2" '中
                strPZD11 = "004"
            Case "彰化縣"
                strPZD10 = "2"
                strPZD11 = "005"
            'Modify by Amy 2019/11/08 原:strPZD10 = "2"/strPZD11 = "005"
            Case "雲林縣"
                strPZD10 = "2"
                strPZD11 = "006"
            Case "嘉義市", "嘉義縣"
                strPZD10 = "3" '南
                strPZD11 = "006"
            Case "臺南市"
                strPZD10 = "3"
                strPZD11 = "007"
            Case "屏東縣", "高雄市", "花蓮縣", "臺東縣", "金門縣", "南海島", "澎湖縣"
                strPZD10 = "4" '高
                strPZD11 = "008"
        End Select
        
        If strTp <> MsgText(601) Then
            If Mid(strTp, 1, 2) = "單全" Or Mid(strTp, 1, 2) = "雙全" Then
                strPZD05 = Mid(strTp, 1, 2)
                strPZD06 = LTrim(Mid(strTp, 3))
            ElseIf Left(strTp, 1) = "單" Or Left(strTp, 1) = "雙" Or Left(strTp, 1) = "全" Or Left(strTp, 1) = "連" Then
                strPZD05 = Left(strTp, 1)
                strPZD06 = LTrim(Mid(strTp, 2))
            Else
                strPZD06 = strTp
            End If
        End If
      
        
        '新增
        'Modify by Amy 2019/11/05 +序號pzd14
        strIns = "Insert Into PostZipData_New (pzd01,pzd02,pzd03,pzd04,pzd05,pzd06,pzd10,pzd11,pzd14) Values(" & _
                    CNULL(strPZD01) & "," & CNULL(strPZD02) & "," & CNULL(strPZD03) & "," & _
                    CNULL(strPZD04) & "," & CNULL(strPZD05) & "," & CNULL(strPZD06) & "," & _
                    CNULL(strPZD10) & "," & CNULL(strPZD11) & ",'" & strNo & "')"

        cnnConnection.Execute strIns
        Label2.Caption = "第 " & iRow & " 列": DoEvents
        iRow = iRow + 1
        strNo = Val(strNo) + 1 'Add by Amy 2019/11/05
    Loop
    cnnConnection.CommitTrans
    
    '新資料未造字,發mail通知
    If strContent <> MsgText(601) Then
        PUB_SendMail "QPGMR", strUserNum, "", "郵遞區號轉檔未造字通知！", strContent
    End If
    strContent = "請增加下列語法：" & vbCrLf & _
                        "Drop Table postzipdata_Old" & vbCrLf & _
                        "Alter table Postzipdata rename to postzipdata_Old" & vbCrLf & _
                        "Alter table Postzipdata_New rename to postzipdata"
    PUB_SendMail "QPGMR", strUserNum, "", "請於每日批次執行語法增加ReName資料表（For 郵遞區號轉檔）！", strContent
    
    '關閉
    xlsSalesPoint.Workbooks.Close
    '離開
    xlsSalesPoint.Quit
    Set wksrpt = Nothing
    Set xlsSalesPoint = Nothing
    Screen.MousePointer = vbDefault
   
    MsgBox "資料新增完畢！"
    Exit Sub
   
flgErr:
    cnnConnection.RollbackTrans
    Screen.MousePointer = vbDefault
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set wksrpt = Nothing
    Set xlsSalesPoint = Nothing
    If Err.Number <> 0 Then
        MsgBox iRow & " 筆 : " & Err.Description
    End If
End Sub

'Add By Sindy 2013/3/1
Private Sub Command11_Click_Old()
'Dim strFileName As String
'Dim xlsSalesPoint As New Excel.Application
'Dim wksaccrpt114 As New Worksheet
'Dim bolReadExit As Boolean
'Dim iRow As Double
'Dim strUpdateZip As String, strZip As String
'Dim rs As ADODB.Recordset
'Dim strMZ06 As String
'Dim strPZD02 As String, strPZD03 As String, strPZD04 As String, strPZD05 As String, strPZD06 As String
'Dim i As Integer
'
'   On Error GoTo flgErr
'
'   If txtFileName = "" Then
'      MsgBox "檔案不可空白！"
'      txtFileName.SetFocus
'      Exit Sub
'   End If
'
'   strFileName = txtFileName
'
'   '開檔
'   Screen.MousePointer = vbHourglass
'   xlsSalesPoint.Workbooks.Open strFileName
'   Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
'
'   '過濾客戶資料
'   strUpdateZip = "": iRow = 36705
'   cnnConnection.BeginTrans
'   Dim strName As String, strAddr As String, strText As String
''   strSql = "delete from referencenames"
''   strText = ""
''   cnnConnection.Execute strSql
'   Do While wksaccrpt114.Range("A" & iRow).Value <> "結束"
'''[轉參考名條]
''      strName = Trim(wksaccrpt114.Range("A" & iRow).Value)
''      strZip = Left(Trim(wksaccrpt114.Range("B" & iRow).Value), 3)
''      strZip = Replace(strZip, "0", "０")
''      strZip = Replace(strZip, "1", "１")
''      strZip = Replace(strZip, "2", "２")
''      strZip = Replace(strZip, "3", "３")
''      strZip = Replace(strZip, "4", "４")
''      strZip = Replace(strZip, "5", "５")
''      strZip = Replace(strZip, "6", "６")
''      strZip = Replace(strZip, "7", "７")
''      strZip = Replace(strZip, "8", "８")
''      strZip = Replace(strZip, "9", "９")
''      strAddr = Mid(Trim(wksaccrpt114.Range("B" & iRow).Value), 4)
''
''      strExc(0) = "select * from referencenames where rn01='" & strName & "'"
''      intI = 1
''      Set rs = ClsLawReadRstMsg(intI, strExc(0))
''      If intI = 1 Then
''         strText = strText & ";" & strName
''      Else
''         strSql = "insert into referencenames values('" & strName & "','" & strZip & "','" & strAddr & "')"
''         cnnConnection.Execute strSql
''      End If
'
''[轉郵遞區號]
'      strZip = Trim(wksaccrpt114.Range("A" & iRow).Value)
''      strZip = Replace(strZip, "0", "０")
''      strZip = Replace(strZip, "1", "１")
''      strZip = Replace(strZip, "2", "２")
''      strZip = Replace(strZip, "3", "３")
''      strZip = Replace(strZip, "4", "４")
''      strZip = Replace(strZip, "5", "５")
''      strZip = Replace(strZip, "6", "６")
''      strZip = Replace(strZip, "7", "７")
''      strZip = Replace(strZip, "8", "８")
''      strZip = Replace(strZip, "9", "９")
'
''      If strUpdateZip <> strZip Then
''         strMZ06 = ""
''         If Trim(wksaccrpt114.Range("B" & iRow).Value) = "台北市" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "新北市" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "基隆市" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "宜蘭縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "花蓮縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "新竹縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "新竹市" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "桃園縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "苗栗縣" Then
''            strMZ06 = "1" '北
''         ElseIf Trim(wksaccrpt114.Range("B" & iRow).Value) = "台中市" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "南投縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "彰化縣" Then
''            strMZ06 = "2" '中
''         ElseIf Trim(wksaccrpt114.Range("B" & iRow).Value) = "嘉義巿" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "嘉義縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "雲林縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "台南市" Then
''            strMZ06 = "3" '南
''         ElseIf Trim(wksaccrpt114.Range("B" & iRow).Value) = "高雄市" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "澎湖縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "屏東縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "台東縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "金門縣" Or _
''            Trim(wksaccrpt114.Range("B" & iRow).Value) = "連江縣" Then
''            strMZ06 = "4" '高
''         End If
'         strPZD02 = Trim(wksaccrpt114.Range("B" & iRow).Value)
'         strPZD03 = Trim(wksaccrpt114.Range("C" & iRow).Value)
'         strPZD04 = Trim(wksaccrpt114.Range("D" & iRow).Value)
'         strPZD05 = ""
'         strPZD06 = ""
'         If Trim(wksaccrpt114.Range("E" & iRow).Value) <> "" Then
'            For i = 1 To 5
'               If Mid(Trim(wksaccrpt114.Range("E" & iRow).Value), i, 1) <> "單" And _
'                  Mid(Trim(wksaccrpt114.Range("E" & iRow).Value), i, 1) <> "雙" And _
'                  Mid(Trim(wksaccrpt114.Range("E" & iRow).Value), i, 1) <> "全" And _
'                  Mid(Trim(wksaccrpt114.Range("E" & iRow).Value), i, 1) <> "連" Then
'                  strPZD06 = Trim(Mid(Trim(wksaccrpt114.Range("E" & iRow).Value), i, Len(Trim(wksaccrpt114.Range("E" & iRow).Value))))
'                  Exit For
'               Else
'                  strPZD05 = strPZD05 & Mid(Trim(wksaccrpt114.Range("E" & iRow).Value), i, 1)
'               End If
'            Next i
'            If i >= 5 Then
'               MsgBox "i>=5"
'            End If
'         End If
''         strExc(0) = "select * from mailzip where mz01='" & strZip & "' and (mz04 is null or mz06 is null) "
''         intI = 1
''         Set rs = ClsLawReadRstMsg(intI, strExc(0))
''         If intI = 1 Then
''            strSql = "update mailzip set " & _
''                     "mz04='" & Trim(wksaccrpt114.Range("B" & iRow).Value) & "'," & _
''                     "mz05='" & Trim(wksaccrpt114.Range("C" & iRow).Value) & "'," & _
''                     "mz06='" & strMZ06 & "' " & _
''                     "where mz01='" & strZip & "'"
''            cnnConnection.Execute strSql
''            strUpdateZip = strZip
''         Else
''            strExc(0) = "select count(*) from mailzip where mz01='" & strZip & "'"
''            intI = 1
''            Set rs = ClsLawReadRstMsg(intI, strExc(0))
''            If intI = 1 Then
''               If rs.Fields(0) <= 0 Then
''                  MsgBox strZip & "無資料"
'                  strSql = "insert into postzipdata (pzd01,pzd02,pzd03,pzd04,pzd05,pzd06) values(" & _
'                  CNULL(strZip) & "," & _
'                  CNULL(strPZD02) & "," & _
'                  CNULL(strPZD03) & "," & _
'                  CNULL(strPZD04) & "," & _
'                  CNULL(strPZD05) & "," & _
'                  CNULL(strPZD06) & _
'                  ")"
'                  cnnConnection.Execute strSql
'                  Label2.Caption = "第 " & iRow & " 筆": DoEvents
''               End If
''            End If
''            strUpdateZip = strZip
''         End If
''      End If
'
'      iRow = iRow + 1
'   Loop
'   cnnConnection.CommitTrans
'   rs.Close
'   Set rs = Nothing
'
'   '關閉
'   xlsSalesPoint.Workbooks.Close
'   '離開
'   xlsSalesPoint.Quit
'   Set wksaccrpt114 = Nothing
'   Set xlsSalesPoint = Nothing
'   Screen.MousePointer = vbDefault
'
'   MsgBox "資料過濾完畢！"
'
'   Exit Sub
'
'flgErr:
'   'cnnConnection.RollbackTrans
'   cnnConnection.CommitTrans
'   Screen.MousePointer = vbDefault
'   If Err.Number <> 0 Then
'      MsgBox iRow & " 筆 : " & Err.Description
'   End If
End Sub

'2013/4/19 ADD BY SONIA (只做004中區資料)專利公報申請人tpbulletin_cust更新客戶編號tpc03,及非本所案件所有公報人代理人名稱tpc04
Private Sub Command12_Click()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strCUNo As String
Dim strFAName As String
   
   Screen.MousePointer = vbHourglass
   
   '更新客戶編號tpc03
   cnnConnection.Execute "Update tpbulletin_cust Set tpc03=null Where tpc02='004'"
   StrSQLa = "select distinct tpc01,tpc02 from tpbulletin_cust where tpc02='004' order by tpc01"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      MsgBox "查無資料!!!"
      GoTo Nextstep
   Else
      rsA.MoveFirst
      Do While Not rsA.EOF
         StrSqlB = "select cu01||cu02 from customer where cu04='" & rsA.Fields("TPC01") & "' order by cu01||cu02"
         rsB.CursorLocation = adUseClient
         rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
         If rsB.RecordCount > 0 Then
            strCUNo = ""
            rsB.MoveFirst
            Do While Not rsB.EOF
               If "" & rsB.Fields(0) <> "" Then strCUNo = rsB.Fields(0) & "，" & strCUNo
               rsB.MoveNext
            Loop
            cnnConnection.Execute "Update tpbulletin_cust Set tpc03='" & strCUNo & "' Where tpc01='" & rsA.Fields("TPC01") & "' and tpc02='" & rsA.Fields("TPC02") & "'"
         End If
         If rsB.State <> adStateClosed Then rsB.Close
         Set rsB = Nothing
         rsA.MoveNext
      Loop
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   '更新非本所案件所有公報人代理人名稱tpc04
   cnnConnection.Execute "Update tpbulletin_cust Set tpc04=null Where tpc02='004'"
   StrSQLa = "select tpc01,tpc02 from tpbulletin_cust where tpc02='004' and tpc03 is not null order by tpc01"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
   Else
      rsA.MoveFirst
      Do While Not rsA.EOF
         StrSqlB = "select distinct substr(tpb08,1,4) from tpbulletin where instr(tpb14||tpb15||tpb16||tpb17||tpb18||tpb19||tpb20||tpb21||tpb22||tpb23,'" & rsA.Fields("TPC01") & "')>0 and tpb06='" & rsA.Fields("TPC02") & "' and (tpb07 is null or tpb07<>'01')"
         rsB.CursorLocation = adUseClient
         rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
         If rsB.RecordCount > 0 Then
            strFAName = ""
            rsB.MoveFirst
            Do While Not rsB.EOF
               If "" & rsB.Fields(0) <> "" Then strFAName = rsB.Fields(0) & "，" & strFAName
               rsB.MoveNext
            Loop
            cnnConnection.Execute "Update tpbulletin_cust Set tpc04='" & strFAName & "' Where tpc01='" & rsA.Fields("TPC01") & "' and tpc02='" & rsA.Fields("TPC02") & "'"
         End If
         If rsB.State <> adStateClosed Then rsB.Close
         Set rsB = Nothing
         rsA.MoveNext
      Loop
   End If
   
Nextstep:
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   MsgBox "更新完畢"
   Screen.MousePointer = vbDefault
   Exit Sub
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Sub

'Add By Sindy 2014/10/8 刪除已發文聯絡附件
Private Sub Command13_Click()
Dim m_EEP01 As String, m_EEP02 As String
   
On Error GoTo CheckingErr
   
   strSql = "select eep01,eep02,eep04,cp27,cp57" & _
            " from empelectronfile,caseprogress,empelectronprocess" & _
            " where eep01=cp09(+)" & _
            " and cp27 is not null and cp27<=20141002" & _
            " and eep04='00'" & _
            " and eep01=eef01(+) and eep02=eef02(+)" & _
            " and eef03 is not null" & _
            " Union" & _
            " select eep01,eep02,eep04,cp27,cp57" & _
            " from empelectronfile,caseprogress,empelectronprocess" & _
            " where eep01=cp09(+)" & _
            " and cp57 is not null and cp57<=20141002" & _
            " and eep04='00'" & _
            " and eep01=eef01(+) and eep02=eef02(+)" & _
            " and eef03 is not null"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         If "" & adoRecordset.Fields("eep04") = EMP_聯絡 Then
            '已發文或已取消收文
            If (Val("" & adoRecordset.Fields("cp27")) <= 20141002 And Val("" & adoRecordset.Fields("cp27")) > 0) Or _
               (Val("" & adoRecordset.Fields("cp57")) <= 20141002 And Val("" & adoRecordset.Fields("cp57")) > 0) Then
               cnnConnection.BeginTrans
               m_EEP01 = "" & adoRecordset.Fields("eep01")
               m_EEP02 = "" & adoRecordset.Fields("eep02")
               strSql = "delete empelectronfile" & _
                        " where eef01='" & m_EEP01 & "' and eef02=" & m_EEP02
               cnnConnection.Execute strSql
               cnnConnection.CommitTrans
            End If
         End If
         adoRecordset.MoveNext
      Loop
   End If
   
   MsgBox "刪除完畢!!!"
   Exit Sub
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Sub

'ADD BY SONIA 2016/7/6
'專利EPC英國有效及EU案抓關聯案申請日(PATENT_EU_SONIA20160707檔案資料來源在C:\83002\案件系統文件\雜文\專利處\CFP\英國脫歐CFP抓資料語法.txt)
Private Sub Command14_Click()
Dim strSQL1 As String
On Error GoTo CheckingErr

   strSql = "select PA01,PA02,PA03,PA04 from PATENT_EU_SONIA20160707 "
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         cnnConnection.Execute "delete from r100101_h where id='" & strUserNum & "' "
         cnnConnection.Execute "insert into r100101_h select '" & adoRecordset.Fields("pa01") & "','" & adoRecordset.Fields("pa02") & "','" & adoRecordset.Fields("pa03") & "','" & adoRecordset.Fields("pa04") & "',0,'1','" & strUserNum & "' from dual "
         cnnConnection.Execute "insert into r100101_h select '" & adoRecordset.Fields("pa01") & "','" & adoRecordset.Fields("pa02") & "','" & adoRecordset.Fields("pa03") & "','" & adoRecordset.Fields("pa04") & "',0,'2','" & strUserNum & "' from dual "
         cnnConnection.Execute "begin   db_r100101_h('" & strUserNum & "'); end;"
      
         strSQL1 = "insert into PATENT_EU_SONIA20160707 " & _
                   "(select '" & adoRecordset.Fields("pa01") & "','" & adoRecordset.Fields("pa02") & "','" & adoRecordset.Fields("pa03") & "','" & adoRecordset.Fields("pa04") & "',R001001,R001002,R001003,R001004,pa10 from r100101_h,patent " & _
                   "where id='83002' and R001001=pa01(+) and R001002=pa02(+) and R001003=pa03(+) and R001004=pa04(+) " & _
                   "and R001001||R001002||R001003||R001004<>'" & adoRecordset.Fields("pa01") & adoRecordset.Fields("pa02") & adoRecordset.Fields("pa03") & adoRecordset.Fields("pa04") & "')"
         cnnConnection.Execute strSQL1
         adoRecordset.MoveNext
      Loop
   End If
   MsgBox "更新完畢"
   Exit Sub
   
CheckingErr:
   MsgBox (Err.Description)

End Sub
'END 2016/7/6

'Add By Sindy 2017/5/22 往來記錄附件轉檔
'DB轉電子檔
Private Sub Command15_Click()
Dim strReName As String
Dim strFullFileName As String
Dim strFtpPath As String
Dim lngSize As Double
Dim iFileNo As Integer
Dim bytes() As Byte
Dim strCR09 As String, strCR20 As String, strCF02 As String, strCF01 As String
Dim varTemp1
Dim varTemp2
Dim j As Integer
Dim intStar As Integer
Dim intEnd As Integer
Dim strFileName As String, strSize As String
Dim strNotInID As String
Dim rsA1 As New ADODB.Recordset
   
On Error GoTo CheckingErr
   
   Screen.MousePointer = vbHourglass
   strSql = "select * from MAILSCHEDULEDETAIL where msd01='1418'"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      'cnnConnection.BeginTrans
      Do While Not adoRecordset.EOF
         strCF01 = UCase("" & adoRecordset.Fields("msd02"))
         strExc(0) = "select msd06 from MAILSCHEDULEDETAIL where msd01<>'1418' and upper(msd02)='" & strCF01 & "' and msd06<>'None' and (substr(msd06,1,1)='X' or substr(msd06,1,1)='Y' or substr(msd06,1,1)='R') order by msd03 desc,msd04 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "update MAILSCHEDULEDETAIL set msd06='" & RsTemp.Fields("msd06") & "'" & _
                     " where msd01='1418'" & _
                     " and upper(msd02)='" & strCF01 & "'"
            cnnConnection.Execute strSql, intI
         End If
         adoRecordset.MoveNext
      Loop
      'cnnConnection.CommitTrans
   End If
   
   Screen.MousePointer = vbDefault
   MsgBox "匯入完畢!!!"
   
   Exit Sub
   
'   strSql = "delete from CPMAXCP05"
'   cnnConnection.Execute strSql
'
'   'Add By Sindy 110/4/16 Run暫存資料
'   'modify by sonia 2023/2/23 再加其他,不得代理專利,不得代理商標,解除對造,國內同業
'   strSql = "SELECT cu01||cu02 From customer" & _
'            " where cu02='0' and (cu80 is null or cu80='業務自行處理' or cu80='其他' or cu80='不得代理專利' or cu80='不得代理商標' or cu80='解除對造' or cu80='國內同業')" & _
'            " AND (SUBSTR(CU12,1,1)<>'F' OR (CU12 IS NULL AND CU10>='001' AND CU10<='008'))" & _
'            " AND ((instr(cu20,'@')=0 AND instr(cu116,'@')=0 AND instr(cu117,'@')=0 AND instr(cu118,'@')=0) OR cu20||cu116||cu117||cu118 IS NULL)"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      adoRecordset.MoveFirst
'      'cnnConnection.BeginTrans
'      Do While Not adoRecordset.EOF
'         strSql = "SELECT MAX(cp05) FROM caseprogress" & _
'                  " where (cp01,cp02,cp03,cp04) in(" & _
'                  " SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK" & _
'                  " WHERE TM23='" & adoRecordset.Fields(0) & "' or TM78='" & adoRecordset.Fields(0) & "' or TM79='" & adoRecordset.Fields(0) & "' or TM80='" & adoRecordset.Fields(0) & "' or TM81='" & adoRecordset.Fields(0) & "'" & _
'                  " union all select PA01,PA02,PA03,PA04 FROM PATENT" & _
'                  " WHERE PA26='" & adoRecordset.Fields(0) & "' or PA27='" & adoRecordset.Fields(0) & "' or PA28='" & adoRecordset.Fields(0) & "' or PA29='" & adoRecordset.Fields(0) & "' or PA30='" & adoRecordset.Fields(0) & "'" & _
'                  " union all select SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE" & _
'                  " WHERE SP08='" & adoRecordset.Fields(0) & "' or SP58='" & adoRecordset.Fields(0) & "' or SP59='" & adoRecordset.Fields(0) & "' or SP65='" & adoRecordset.Fields(0) & "' or SP66='" & adoRecordset.Fields(0) & "'" & _
'                  " union all select LC01,LC02,LC03,LC04 FROM LAWCASE" & _
'                  " WHERE LC11='" & adoRecordset.Fields(0) & "' or LC43='" & adoRecordset.Fields(0) & "' or LC44='" & adoRecordset.Fields(0) & "' or LC45='" & adoRecordset.Fields(0) & "' or LC46='" & adoRecordset.Fields(0) & "'" & _
'                  " union all select HC01,HC02,HC03,HC04 FROM HIRECASE" & _
'                  " WHERE HC05='" & adoRecordset.Fields(0) & "' or HC24='" & adoRecordset.Fields(0) & "' or HC25='" & adoRecordset.Fields(0) & "' or HC26='" & adoRecordset.Fields(0) & "' or HC27='" & adoRecordset.Fields(0) & "'" & _
'                  ")"
'         intI = 1
'         Set rsA1 = ClsLawReadRstMsg(intI, strSql)
'         If intI = 1 Then
'            If Val("" & rsA1.Fields(0)) = 0 Then
'               'MsgBox "無收文" & adoRecordset.Fields(0)
'            Else
'               strSql = "insert into CPMAXCP05(CUNO,CP05)" & _
'                        "values('" & adoRecordset.Fields(0) & "'," & rsA1.Fields(0) & ")"
'               cnnConnection.Execute strSql
'            End If
'         Else
'            MsgBox "無資料" & adoRecordset.Fields(0)
'         End If
'         adoRecordset.MoveNext
'      Loop
'      'cnnConnection.CommitTrans
'   End If
'   MsgBox "完成"
'   Exit Sub
   
   '往來記錄附件 - 切出檔案大小
   strNotInID = "'KA1000131','KA7000181','KA7000255','KA7000197','KA7000212','KA7000222','KA7000226','KA7000244','KA7000305','KA7000315','KA7000433','KA7000594','KA7000562','KA7000587','KA7000633','KA7000676','KA7000627','KA7001282'"

   Screen.MousePointer = vbHourglass
   strSql = "select * from CONTACTfile where cf01 not in(" & strNotInID & ")"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      cnnConnection.BeginTrans
      Do While Not adoRecordset.EOF
         strCF01 = UCase("" & adoRecordset.Fields("cf01"))
         strCF02 = UCase("" & adoRecordset.Fields("cf02"))
         If InStr(strCF02, "KB)") > 0 And InStr(strNotInID, strCF01) = 0 Then
            intEnd = InStrRev(strCF02, "KB)")
            intStar = InStrRev(strCF02, "(")
            strFileName = Trim(Mid(adoRecordset.Fields("cf02"), 1, intStar - 1))
            strSize = Val(Mid(strCF02, intStar + 1, Len(strCF02) - intEnd + 1))
            strSql = "update CONTACTFILE set cf02='" & ChgSQL(strFileName) & "',cf07='" & strSize & "'" & _
                     " where CF01='" & adoRecordset.Fields("cF01") & "'" & _
                     " and CF02='" & ChgSQL(adoRecordset.Fields("cF02")) & "'"
            cnnConnection.Execute strSql
         End If
         adoRecordset.MoveNext
      Loop
      cnnConnection.CommitTrans
   End If
   
'   '往來記錄附件轉檔
'   Screen.MousePointer = vbHourglass
'   strSql = "select cr01,cr09,cr20 from CONTACTRECORD where cr01>='KA8'"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      adoRecordset.MoveFirst
'      'cnnConnection.BeginTrans
'      Do While Not adoRecordset.EOF
'         strCR09 = "" & adoRecordset.Fields("cr09")
'         strCR20 = "" & adoRecordset.Fields("cr20")
'         If strCR09 <> "" And strCR20 <> "" Then
'            varTemp1 = Split(strCR09, ",")
'            varTemp2 = Split(strCR20, ",")
'            If UBound(varTemp1) = UBound(varTemp2) Then
'               For j = 0 To UBound(varTemp1)
'                  strExc(0) = "select * from CONTACTFILE where cf01='" & adoRecordset.Fields("cr01") & "'" & _
'                              " and upper(cf02)='" & ChgSQL(UCase(varTemp1(j))) & "'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 0 Then
'                     MsgBox adoRecordset.Fields("cr01") & "附件不存在" & ChgSQL(UCase(varTemp1(j))) & "!"
'                  End If
'               Next j
''               For j = 0 To UBound(varTemp1)
''                  strSql = "insert into CONTACTFILE(cf01,cf02,cf06)" & _
''                           "values('" & adoRecordset.Fields("cr01") & "','" & ChgSQL(varTemp1(j)) & "','" & ChgSQL(varTemp2(j)) & "')"
''                  cnnConnection.Execute strSql
''               Next j
'            Else
'               MsgBox adoRecordset.Fields("cr01") & "附件資料檔名和路徑個數有誤!"
'            End If
'         ElseIf strCR09 <> "" Or strCR20 <> "" Then
'            MsgBox adoRecordset.Fields("cr01") & "附件資料有誤!"
'         End If
'         adoRecordset.MoveNext
'      Loop
'      'cnnConnection.CommitTrans
'   End If
   
'   strSql = "select plf01,plf02 from pricelistfile"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         strExc(0) = "select * from pricelistfile where plf01='" & adoRecordset.Fields("plf01") & "' and plf02='" & adoRecordset.Fields("plf02") & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            With RsTemp
'            strReName = .Fields("plf02").Value & "." & .Fields("plf03").Value & ".pdf"
'            strFullFileName = App.path & "\pricelistfile\" & .Fields("plf01").Value & "." & .Fields("plf02").Value & "." & .Fields("plf03").Value & ".pdf"
'
'            lngSize = Val(.Fields("plf03").Value)
'            ReDim bytes(lngSize)
'            If lngSize > 0 Then bytes() = .Fields("plf04").GetChunk(lngSize)
'            End With
'            iFileNo = FreeFile
'            Open strFullFileName For Binary Access Write As #iFileNo
'            If lngSize > 0 Then Put #iFileNo, , bytes()
'            Close #iFileNo
'
'            If Dir(strFullFileName) <> "" Then
'               cnnConnection.BeginTrans
'
'               ' 檔案改放FTP
'               PUB_PutFtpFile strFullFileName, adoRecordset.Fields("plf01").Value, strReName, strFtpPath, UCase("pricelistfile")
'               If strFtpPath <> "" Then
'                  '更新資料庫資料
'                  strSql = "update pricelistfile set " & _
'                           "plf11='" & strFtpPath & "'" & _
'                           " where plf01='" & adoRecordset.Fields("plf01").Value & "' and plf02='" & adoRecordset.Fields("plf02").Value & "'"
'                  cnnConnection.Execute strSql
'               Else
'                  GoTo CheckingErr
'               End If
'               'Call PUB_DelPCOrgFile(strFullFileName) '一併將PC上的實體檔案刪除
'
'               cnnConnection.CommitTrans
'            End If
'         End If
'         adoRecordset.MoveNext
'      Loop
'   End If

'   strSql = "select * from imgbytefile where ibf01='CFP' and ibf15 is null and rownum<=2000"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
''         strExc(0) = "select * from consultrecimagef where crif01='" & adoRecordset.Fields("crif01") & "'"
''         intI = 1
''         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''         If intI = 1 Then
'            With adoRecordset
'               strReName = .Fields("ibf01").Value & "-" & .Fields("ibf02").Value & "-" & .Fields("ibf03").Value & "-" & .Fields("ibf04").Value & "-" & .Fields("ibf05").Value
'               strFullFileName = App.path & "\imgbytefile\" & strReName
'
'               lngSize = Val(.Fields("ibf13").Value)
'               ReDim bytes(lngSize)
'               If lngSize > 0 Then bytes() = .Fields("ibf14").GetChunk(lngSize)
'               iFileNo = FreeFile
'               Open strFullFileName For Binary Access Write As #iFileNo
'               If lngSize > 0 Then Put #iFileNo, , bytes()
'               Close #iFileNo
'
'               If Dir(strFullFileName) <> "" Then
'                  cnnConnection.BeginTrans
'
'                  ' 檔案改放FTP
'                  PUB_PutFtpFile strFullFileName, strReName, strReName, strFtpPath, UCase("imgbytefile")
'                  If strFtpPath <> "" Then
'                     '更新資料庫資料
'                     strSql = "update imgbytefile set " & _
'                              "ibf15='" & strFtpPath & "'" & _
'                              " where ibf01='" & .Fields("ibf01").Value & "'" & _
'                              " and ibf02='" & .Fields("ibf02").Value & "'" & _
'                              " and ibf03='" & .Fields("ibf03").Value & "'" & _
'                              " and ibf04='" & .Fields("ibf04").Value & "'" & _
'                              " and ibf05='" & .Fields("ibf05").Value & "'"
'                     cnnConnection.Execute strSql
'                  Else
'                     GoTo CheckingErr
'                  End If
'                  Call PUB_DelPCOrgFile(strFullFileName) '一併將PC上的實體檔案刪除
'
'                  cnnConnection.CommitTrans
'               End If
'            End With
''         End If
'         adoRecordset.MoveNext
'      Loop
'   End If
   
'   strSql = "select cf01,cf02 from custwebfile"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         strExc(0) = "select * from custwebfile where cf01='" & adoRecordset.Fields("cf01") & "' and cf02='" & ChgSQL(adoRecordset.Fields("cf02")) & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            With RsTemp
'            strReName = .Fields("cf03").Value & "." & .Fields("cf02").Value
'            strFullFileName = App.path & "\SeminarAttach\" & strUserNum & "\" & strReName
'
'            lngSize = Val(.Fields("cf03").Value)
'            ReDim bytes(lngSize)
'            If lngSize > 0 Then bytes() = .Fields("cf04").GetChunk(lngSize)
'            End With
'            iFileNo = FreeFile
'            Open strFullFileName For Binary Access Write As #iFileNo
'            If lngSize > 0 Then Put #iFileNo, , bytes()
'            Close #iFileNo
'
'            If Dir(strFullFileName) <> "" Then
'               cnnConnection.BeginTrans
'
'               ' 檔案改放FTP
'               PUB_PutFtpFile strFullFileName, adoRecordset.Fields("cf01").Value, strReName, strFtpPath, UCase("custwebfile")
'               If strFtpPath <> "" Then
'                  '更新資料庫資料
'                  strSql = "update custwebfile set " & _
'                           "cf08='" & strFtpPath & "'" & _
'                           " where cf01='" & adoRecordset.Fields("cf01").Value & "' and cf02='" & adoRecordset.Fields("cf02").Value & "'"
'                  cnnConnection.Execute strSql
'               Else
'                  GoTo CheckingErr
'               End If
'               'Call PUB_DelPCOrgFile(strFullFileName) '一併將PC上的實體檔案刪除
'
'               cnnConnection.CommitTrans
'            End If
'         End If
'         adoRecordset.MoveNext
'      Loop
'   End If
   
'   strSql = "select sa01,sa02 from seminarattachment where sa01<=60 and sa07 is null"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         strExc(0) = "select * from seminarattachment where sa01='" & adoRecordset.Fields("sa01") & "' and sa02='" & ChgSQL(adoRecordset.Fields("sa02")) & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            With RsTemp
'            strReName = .Fields("sa01").Value & "." & .Fields("sa05").Value & "." & .Fields("sa03").Value & "." & .Fields("sa02").Value
'            strFullFileName = App.path & "\TransFile\" & strReName
'
'            lngSize = Val(.Fields("sa03").Value)
'            ReDim bytes(lngSize)
'            If lngSize > 0 Then bytes() = .Fields("sa04").GetChunk(lngSize)
'            End With
'            iFileNo = FreeFile
'            Open strFullFileName For Binary Access Write As #iFileNo
'            If lngSize > 0 Then Put #iFileNo, , bytes()
'            Close #iFileNo
'
'            If Dir(strFullFileName) <> "" Then
'               cnnConnection.BeginTrans
'
'               ' 檔案改放FTP
'               PUB_PutFtpFile strFullFileName, adoRecordset.Fields("sa01").Value, strReName, strFtpPath, "SEMINARATTACHMENT"
'               If strFtpPath <> "" Then
'                  '更新資料庫資料
'                  strSql = "update seminarattachment set " & _
'                           "sa06=0,sa07='" & strFtpPath & "'" & _
'                           " where sa01='" & adoRecordset.Fields("sa01").Value & "' and sa02='" & adoRecordset.Fields("sa02").Value & "'"
'                  cnnConnection.Execute strSql
'               Else
'                  GoTo CheckingErr
'               End If
'               'Call PUB_DelPCOrgFile(strFullFileName) '一併將PC上的實體檔案刪除
'
'               cnnConnection.CommitTrans
'            End If
'         End If
'         adoRecordset.MoveNext
'      Loop
'   End If
   
   Screen.MousePointer = vbDefault
   MsgBox "匯入完畢!!!"
   Exit Sub
   
CheckingErr:
   
   If Err.Number = -2147217873 Then
      strNotInID = strNotInID & ",'" & strCF01 & "'"
      Resume Next
   Else
      Screen.MousePointer = vbDefault
      cnnConnection.RollbackTrans
      MsgBox (Err.Description)
   End If
End Sub

Private Sub Command16_Click()
   
On Error GoTo CheckingErr
   
   If MsgBox("確定要刪除電子檔嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
      Exit Sub
   End If
   
   strSql = "select mst01,mst03 from mailscheduletemplet where mst03 is not null"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         strSql = "update mailscheduletemplet set " & _
                  "mst03=null" & _
                  " where mst01='" & adoRecordset.Fields("mst01").Value & "'"
         cnnConnection.Execute strSql
         adoRecordset.MoveNext
      Loop
   End If
   
   MsgBox "刪除完畢!!!"
   Exit Sub
   
CheckingErr:
   'Resume
   'cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Sub

'財務寄發郵件
'Add By Sindy 2019/9/11 郵件通知:客戶公司有代表號無財務信箱和會計師信箱
Private Sub Command24_Click()
Dim strSubject As String, strContent As String, strEmp As String, strEMP_Tel As String, strTo As String
Dim PrintRpt As Boolean
Dim ff1 As Integer, i As Integer
Dim strFileName As String
Dim strTemp(1 To 7) As String
Dim stAccPerson As String, stTxtPerson As String 'Add by Amy 2024/05/17
   
   'Modify by Amy 2024/05/17 財務2個特殊設定拆成3個
   If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
       stAccPerson = Pub_GetSpecMan("財務處應收處理人員")
   Else
      stAccPerson = Pub_GetSpecMan("財務處總帳人員")
   End If
   stTxtPerson = stAccPerson '取第一個人
   If InStr(stTxtPerson, ";") > 0 Then stTxtPerson = Mid(stTxtPerson, 1, Val(InStr(stTxtPerson, ";")) - 1)
   strExc(0) = "select st02,ed01" & _
               " from staff,ExtensionData" & _
               " where ST01=ED02(+)" & _
               " and st01='" & stTxtPerson & "'"
   'end 2024/05/17
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strEmp = RsTemp.Fields("st02")
      strEMP_Tel = "" & RsTemp.Fields("ed01")
   End If
   
   '資料抓107-108有產生過收據E單號者
   'modify by sonia 2023/2/23 再加其他,不得代理專利,不得代理商標,解除對造,國內同業
   strSql = "SELECT T1,CU04,CU80,CU20,T2,CU12,CU13" & _
            " from (" & _
            " SELECT CU01||CU02 T1,CU04,CU80,CU20,A2.a4901 T2,CU12,CU13" & _
            " FROM customer,(SELECT a4901 FROM acc490 WHERE substr(a4901,1,1)='X' AND a4905 IS NULL) A2" & _
            " WHERE CU15='1'" & _
            " AND CU20 IS NOT NULL AND CU115 IS NULL AND CU02='0'" & _
            " AND NOT EXISTS (SELECT * FROM acc490 WHERE a4901=cu01||cu02 AND a4905 IS NOT NULL)" & _
            " AND (CU80 IS NULL OR CU80='業務自行處理' or cu80='其他' or cu80='不得代理專利' or cu80='不得代理商標' or cu80='解除對造' or cu80='國內同業')" & _
            " AND CU01||CU02=A2.a4901(+)" & _
            " AND CU158 IS NULL" & _
            " AND substr(CU12,1,1)<>'F') A3," & _
            " (SELECT a0k03,a0k04,count(*) FROM acc0k0 WHERE a0k02>=1070101 GROUP BY a0k03,a0k04) A1" & _
            " Where a1.a0k03 = A3.t1" & _
            " OR A1.a0k04=A3.cu04" & _
            " GROUP BY T1,CU04,CU80,CU20,T2,CU12,CU13" & _
            " order BY A3.CU80 asc,A3.T1 asc"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   Screen.MousePointer = vbHourglass
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         strTo = adoRecordset.Fields("CU20")
         strSubject = adoRecordset.Fields("T1") & " " & Left(adoRecordset.Fields("cu04"), 6) & "【重要提醒】請提供會計聯絡信箱,以便核對扣繳事項"
         'Modify by Amy 2024/05/17 原:71006@taie.com.tw
         strContent = "E-Mail：" & strTo & vbCrLf & vbCrLf & _
                     "敬啟者：" & vbCrLf & vbCrLf & _
                     "為<B>便於</B>年底扣繳憑單開立前的核對作業，" & vbCrLf & _
                     "目前本所尚無　貴司提供的（會計E-Mail或 會計師E-Mail）" & vbCrLf & vbCrLf & _
                     "若您方便" & vbCrLf & _
                     "請您將會會計E-Mail或 會計師E-Mail 寄至<U>" & stTxtPerson & "@taie.com.tw</U>" & vbCrLf & vbCrLf & vbCrLf & _
                     "<B>財務信箱：</B>" & vbCrLf & vbCrLf & _
                     "<B>會計師信箱：</B>" & vbCrLf & _
                     "<B>會計師電話：</B>" & vbCrLf & vbCrLf & _
                     "本所於每年12月20左右，會將　貴司當年往來的扣繳明細資料，寄到您" & vbCrLf & _
                     "所提供的信箱以供參考。" & vbCrLf & vbCrLf & _
                     "謝謝您的合作！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                     "財務處　" & strEmp & vbCrLf & _
                     "台一國際專利法律事務所" & vbCrLf & _
                     "台北市長安東路２段１１２號９樓" & vbCrLf & _
                     "電話：０２－２５０６１０２３" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
                     "傳真：０２－２５０１１６６６"
         ''strTo = "97038"
         PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , True, , , , , , True, , , , False
         If bolMailSendOk = False Then
            If PrintRpt = False Then
               PrintRpt = True
               If ff1 > 0 Then Close #ff1
               ff1 = FreeFile
               strFileName = "寄郵件失敗資料檢核表" & strSrvDate(2) & ".txt"
               Open PUB_Getdesktop & "\" & strFileName For Output As ff1
               Print #ff1, "備註：改字型Fixedsys標準11號字以橫式上下左右各10MM列印"
               Print #ff1, "客戶編號   公司名稱             客戶狀態        E-Mail                               有建立會計師資料"
               Print #ff1, "========== ==================== =============== ==================================== ================"
            End If
            For i = 1 To 5
               strTemp(i) = ""
            Next i
            strTemp(1) = Trim("" & adoRecordset.Fields(0))
            strTemp(2) = Trim("" & adoRecordset.Fields(1))
            strTemp(3) = Trim("" & adoRecordset.Fields(2))
            strTemp(4) = Trim("" & adoRecordset.Fields(3))
            strTemp(5) = Trim("" & adoRecordset.Fields(4))
            
            strTemp(1) = convForm(CheckStr(strTemp(1)), 10)
            strTemp(2) = convForm(CheckStr(strTemp(2)), 20)
            strTemp(3) = convForm(CheckStr(strTemp(3)), 15)
            strTemp(4) = convForm(CheckStr(strTemp(4)), 30)
            strTemp(5) = CheckStr(strTemp(5))
            Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5)
         End If
         adoRecordset.MoveNext
      Loop
      If PrintRpt = True Then Close ff1
   End If
   
   Screen.MousePointer = vbDefault
   MsgBox "郵件寄出,完畢!!!"
   Exit Sub
End Sub
'Added by Morgan 2019/9/20
'設定電子公文手動下載
Private Sub Command26_Click()
   If txtWD01 = "" Then
      MsgBox "請輸入工作日！", vbCritical
      txtWD01.SetFocus
   ElseIf ChkWorkDay(DBDATE(txtWD01)) = False Then
      MsgBox "請輸入工作日！", vbCritical
      txtWD01.SetFocus
   ElseIf txtWD07 <> "Y" And txtWD07 <> "" Then
      MsgBox "只可輸入Y！", vbCritical
      txtWD07.SetFocus
   Else
      strSql = "update workday set wd07='" & txtWD07 & "' where wd01=" & DBDATE(txtWD01)
      cnnConnection.Execute strSql, intI
      MsgBox "設定完成！", vbOKOnly
      txtWD01 = "": txtWD07 = ""
      txtWD01.SetFocus
   End If
End Sub

'Added by Morgan 2018/9/17 改定稿日期
Private Sub Command29_Click()
   Dim stSQL As String
   Dim Int1 As Integer, Int2 As Integer
   Dim strDate1 As String, StrDate2 As String
   Dim strOld1 As String, strOld2 As String
   Dim strNew1 As String, strNew2 As String
   
   
On Error GoTo ErrHand29 'Added by Lydia 2018/09/27

   If Text29(0) = "" Then MsgBox "請輸入" & Label29(0) & "！", vbExclamation: Text29(0).SetFocus: Exit Sub
   If Text29(1) = "" Then
      MsgBox "請輸入" & Label29(1) & "！", vbExclamation: Text29(1).SetFocus: Exit Sub
   ElseIf ChkDate(Text29(1)) = False Then
      Text29(1).SetFocus
      Text29_GotFocus 1
      Exit Sub
   End If
   If Text29(2) = "" Then
      MsgBox "請輸入" & Label29(2) & "！", vbExclamation: Text29(2).SetFocus: Exit Sub
   ElseIf ChkDate(Text29(2)) = False Then
      Text29(2).SetFocus
      Text29_GotFocus 2
      Exit Sub
   End If
   
   strDate1 = DBDATE(Text29(1))
   StrDate2 = DBDATE(Text29(2))
   'Modified by Morgan 2019/11/22 改民國年(西元年)
   'strOld1 = (Left(strDate1, 4) - 1911) & "年" & Mid(strDate1, 5, 2) & "月" & Mid(strDate1, 7)
   'strNew1 = (Left(strDate2, 4) - 1911) & "年" & Mid(strDate2, 5, 2) & "月" & Mid(strDate2, 7)
   strOld1 = (Left(strDate1, 4) - 1911) & "(" & Left(strDate1, 4) & ")年" & Mid(strDate1, 5, 2) & "月" & Mid(strDate1, 7)
   strNew1 = (Left(StrDate2, 4) - 1911) & "(" & Left(StrDate2, 4) & ")年" & Mid(StrDate2, 5, 2) & "月" & Mid(StrDate2, 7)
   'end 2019/11/22
   strOld2 = Left(strDate1, 4) & "年" & Mid(strDate1, 5, 2) & "月" & Mid(strDate1, 7)
   strNew2 = Left(StrDate2, 4) & "年" & Mid(StrDate2, 5, 2) & "月" & Mid(StrDate2, 7)
   
   cnnConnection.BeginTrans
   
   stSQL = "UPDATE LETTERDEMAND SET LD15=REPLACE(LD15,'中華民國" & strOld1 & "日','中華民國" & strNew1 & "日')" & _
      " WHERE LD01='" & Text29(0) & "' AND LD16 IS NULL AND INSTR(LD15,'中華民國" & strOld1 & "日')>0"
   
   cnnConnection.Execute stSQL, Int1

   stSQL = "UPDATE LETTERDEMAND SET LD15=REPLACE(LD15,'發函日期：" & strOld2 & "日','發函日期：" & strNew2 & "日')" & _
      " WHERE LD01='" & Text29(0) & "' anD LD16 IS NULL AND INSTR(LD15,'發函日期：" & strOld2 & "日')>0"
   cnnConnection.Execute stSQL, Int2
   
   If MsgBox("台對台: " & Int1 & " 筆" & vbCrLf & "大對台: " & Int2 & " 筆" & vbCrLf & vbCrLf & "是否確定？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
      cnnConnection.CommitTrans
      MsgBox "完成! " 'Added by Lydia 2018/09/27
   Else
      cnnConnection.RollbackTrans
   End If
   
   Exit Sub
   
'Added by Lydia 2018/09/27
ErrHand29:
   If Err.Number <> 0 Then
       MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2019/10/4 匯入xls比對下一程序
Private Sub Command30_Click()
Dim strFileName As String
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim bolReadExit As Boolean
Dim iRow As Integer
Dim strCaseNo As String, strNP08 As String, strNP07n As String, strNP07 As String
Dim strNote As String, Str01 As String, Str02 As String, Str03 As String, Str04 As String
Dim strNA01 As String, strNP09 As String
   
   On Error GoTo flgErr
   
   strFileName = "C:\Users\97038\Desktop\逾本所期限未處理案件\合併全部.xls"
   
   '開檔
   Screen.MousePointer = vbHourglass
   xlsSalesPoint.Workbooks.Open strFileName
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
   
'   '新增A欄
'   wksaccrpt114.Columns("A:A").Select
'   xlsSalesPoint.Selection.Insert Shift:=xlToRight
'   wksaccrpt114.Range("A1").Select
'   wksaccrpt114.Range("A1").Value = "本所客戶"
'   wksaccrpt114.Columns("A:A").ColumnWidth = 14
   
   '逐筆檢查案件資料
   bolReadExit = False: iRow = 2
   Do While bolReadExit = False
      If wksaccrpt114.Range("A" & iRow).Value = "" Then
         bolReadExit = True
      Else
         strCaseNo = Trim(wksaccrpt114.Range("E" & iRow).Value) '本所案號
         If UBound(Split(strCaseNo, "-")) = 2 And Right(strCaseNo, 3) = "-00" Then
            '寫回Excel檔
            wksaccrpt114.Range("L" & iRow).Value = "案號有誤"
            GoTo ReadNext
         ElseIf UBound(Split(strCaseNo, "-")) = 1 Then
            strCaseNo = strCaseNo & "-0-00"
         End If
         strNP08 = DBDATE(Trim(wksaccrpt114.Range("C" & iRow).Value)) '本所期限
         strNP09 = DBDATE(Trim(wksaccrpt114.Range("D" & iRow).Value)) '法定期限
         strNP07n = Trim(wksaccrpt114.Range("H" & iRow).Value) '案件性質
         strNote = Trim(wksaccrpt114.Range("K" & iRow).Value) '備註
         
         If UBound(Split(strCaseNo, "-")) = 3 Then '正規案號
            Str01 = SystemNumber(strCaseNo, 1)
            Str02 = SystemNumber(strCaseNo, 2)
            Str03 = SystemNumber(strCaseNo, 3)
            Str04 = SystemNumber(strCaseNo, 4)
            
            '主檔
            strSql = "select pa09 from patent,casepropertymap where pa01='" & Str01 & "' and pa02='" & Str02 & "' and pa03='" & Str03 & "' and pa04='" & Str04 & "'"
            strSql = strSql & " Union Select tm10 From Trademark,casepropertymap Where tm01='" & Str01 & "' and tm02='" & Str02 & "' and tm03='" & Str03 & "' and tm04='" & Str04 & "'"
            strSql = strSql & " Union Select LC15 From Lawcase,casepropertymap Where LC01='" & Str01 & "' and LC02='" & Str02 & "' and LC03='" & Str03 & "' and LC04='" & Str04 & "'"
            strSql = strSql & " Union Select '000' From Hirecase,casepropertymap Where HC01='" & Str01 & "' and HC02='" & Str02 & "' and HC03='" & Str03 & "' and HC04='" & Str04 & "'"
            strSql = strSql & " Union Select SP09 From ServicePractice,casepropertymap Where SP01='" & Str01 & "' and SP02='" & Str02 & "' and SP03='" & Str03 & "' and SP04='" & Str04 & "'"
            CheckOC
            With adoRecordset
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If .RecordCount > 0 Then
                  strNA01 = .Fields("pa09")
                  '案件性質
                  strExc(0) = "select cpm02" & _
                              " From casepropertymap" & _
                              " where cpm01='" & Str01 & "'" & _
                              IIf(strNA01 = "000", " and cpm03='" & Trim(strNP07n) & "'", " and cpm04='" & Trim(strNP07n) & "'")
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strNP07 = RsTemp.Fields("cpm02")
                  Else
                     '寫回Excel檔
                     wksaccrpt114.Range("L" & iRow).Value = "無對應的案件性質(" & strNA01 & ")"
                     GoTo ReadNext
                  End If
               Else
                  '寫回Excel檔
                  wksaccrpt114.Range("L" & iRow).Value = "無案號"
                  GoTo ReadNext
               End If
            End With
            
            '過濾下一程序資料:
            '請將附件的客戶資料匯回系統，若案件原已結案則略過；
            '備註欄中有內容者全部做結案處理：參考CFP-030175之CA7017131
            '將下一程序之NP06上N，NP11上系統日，NP12上99，
            '在原NP15之後加註 '108/x/x整批結案;'
            
            '備註有值才更新
            If Trim(strNote) <> "" Then
               strSql = "select np01,np07,np22 from nextprogress" & _
                        " where np02='" & Str01 & "' and np03='" & Str02 & "' and np04='" & Str03 & "' and np05='" & Str04 & "'" & _
                        " and np07='" & strNP07 & "' and (np06 is null or np06='')"
               If Val(strNP08) > 0 Then
                  strSql = strSql & " and np08=" & strNP08
               Else
                  strSql = strSql & " and (np08 is null or np08=0)"
               End If
               If Val(strNP09) > 0 Then
                  strSql = strSql & " and np09=" & strNP09
               Else
                  strSql = strSql & " and (np09 is null or np09=0)"
               End If
               If adoRecordset.State = adStateOpen Then
                  adoRecordset.Close
               End If
               adoRecordset.CursorLocation = adUseClient
               adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
               If adoRecordset.RecordCount = 1 Then
                  adoRecordset.MoveFirst
                  Do While Not adoRecordset.EOF
                     strSql = "update nextprogress set" & _
                              " np06='N',np11=" & strSrvDate(1) & _
                              ",np12='99',NP15=decode(NP15,null,'',NP15||';')||'" & ChangeWStringToTDateString(strSrvDate(1)) & "整批結案;" & Trim(strNote) & ";'" & _
                              " where np01='" & adoRecordset.Fields("np01") & "'" & _
                              " and np07='" & adoRecordset.Fields("np07") & "'" & _
                              " and np22='" & adoRecordset.Fields("np22") & "'"
                     cnnConnection.Execute strSql
                     adoRecordset.MoveNext
                  Loop
               ElseIf adoRecordset.RecordCount > 1 Then
                  '寫回Excel檔
                  wksaccrpt114.Range("L" & iRow).Value = "多筆資料"
                  GoTo ReadNext
               Else
                  '寫回Excel檔
                  wksaccrpt114.Range("L" & iRow).Value = "無資料可更新"
                  GoTo ReadNext
               End If
            End If
         Else
            '寫回Excel檔
            wksaccrpt114.Range("L" & iRow).Value = "非正規案號"
            GoTo ReadNext
         End If
         
ReadNext:
         iRow = iRow + 1
      End If
   Loop
   adoRecordset.Close
   '存檔
   xlsSalesPoint.Workbooks(1).SaveAs FileName:=Left(strFileName, Len(strFileName) - 4) & "_" & strSrvDate(2) & ServerTime & ".xls"
   
   '關閉
   xlsSalesPoint.Workbooks.Close
   '離開
   xlsSalesPoint.Quit
   Set wksaccrpt114 = Nothing
   Set xlsSalesPoint = Nothing
   Screen.MousePointer = vbDefault
   
   MsgBox "資料過濾完畢！"
   
   Exit Sub
   
flgErr:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub
'Addded by Lydia 2019/11/08 計算分信時段
Private Sub Command31_Click()
Dim i As Integer
Dim intA As Integer, intB As Integer
Dim strStarTime As String, strEndTime As String
Dim strMsg As String

         For i = 0 To 47
            If i < 3 Then   '凌晨0~1點斷線, 後半小時不執行
                strStarTime = "": strEndTime = ""
            '晚上11點分成兩次分信,最後執行清空[刪除的郵件]晚上23:45~23:55
            ElseIf i = 46 Then '晚上11點第一次分信11:00~11:19
                strStarTime = "230000": strEndTime = "231900"
            ElseIf i = 47 Then  '晚上11點第二次分信11:20~11:39
                strStarTime = "232000": strEndTime = "233900"
            Else
                intA = i \ 2
                intB = i Mod 2
                strStarTime = Format(intA, "00") & IIf(intB = 1, "30", "00") & "00"
                strEndTime = Format(intA, "00") & IIf(intB = 1, "59", "29") & "00"
            End If
            Debug.Print "NO." & Format(i, "00") & "  Start:" & Format(strStarTime, "000000") & "  End:" & Format(strEndTime, "000000")
         Next i
         Debug.Print vbCrLf
         MsgBox "請看即時運算視窗"
End Sub

'Add By Sindy 2019/11/27 匯入xls比對Email查資料 Sindy
Private Sub Command32_Click()
Dim strFileName As String, strFileName2 As String, strData As String
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim bolReadExit As Boolean
Dim iRow As Integer
Dim strEmail As String, strCompanyName As String
Dim strEmail2 As String
Dim strCheckWay As String, strTp(3) As String
   
On Error GoTo flgErr
   
   strFileName2 = ".xlsx"
   strFileName = PUB_Getdesktop & "\Reception follow-up_1129" & strFileName2
   
'   If Dir(strFileName) <> MsgText(601) Then
'      Kill strFileName
'   End If
   
   '開檔
   Dim strWkName As String
   Screen.MousePointer = vbHourglass
   xlsSalesPoint.Workbooks.Open strFileName
   '工作表名稱改為中文
   If strWkName = MsgText(601) Then strWkName = xlsSalesPoint.Worksheets(3).Name
   'Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(strWkName)
   
'   '新增B欄
'   wksaccrpt114.Columns("B:B").Select
'   xlsSalesPoint.Selection.Insert Shift:=xlToRight
'   wksaccrpt114.Range("B1").Select
'   wksaccrpt114.Range("B1").Value = "*用網域查資料"
'   wksaccrpt114.Columns("B:B").ColumnWidth = 10
'   wksaccrpt114.Columns("C:C").Select
'   xlsSalesPoint.Selection.Insert Shift:=xlToRight
'   wksaccrpt114.Range("C1").Select
'   wksaccrpt114.Range("C1").Value = "客戶資料"
'   wksaccrpt114.Columns("C:C").ColumnWidth = 50
   
   '過濾客戶資料
   bolReadExit = False: iRow = 2
   Do While bolReadExit = False
      'E-Mail,公司名稱
      If wksaccrpt114.Range("A" & iRow).Value = "" And _
         wksaccrpt114.Range("F" & iRow).Value = "" Then
         bolReadExit = True
      Else
         '過濾資料
         strData = ""
         If Trim(wksaccrpt114.Range("C" & iRow).Value) <> "" Then GoTo ReadNext
         '名稱
         strCheckWay = ">0"
         strTp(3) = ChgSQL(UCase(Trim(wksaccrpt114.Range("F" & iRow).Value))) '完整名稱
         If InStr(strTp(3), " ") = 0 Then GoTo ReadNext
         strTp(3) = Mid(strTp(3), 1, InStr(strTp(3), " ") - 1) '第1個單字
         If strTp(3) <> "" Then
            '查customer 客戶 檔
            strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,CU04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU04,'" & strTp(3) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+) and CU02='0'"
            strSql = strSql & " union all SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註 From CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & strTp(3) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+) and CU02='0'"
            strSql = strSql & " union all SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,CU06 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註 From CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU06,'" & strTp(3) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+) and CU02='0'"
            
            '查Fagent 代理人 檔
            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,fa04 as 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註 From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa04,'" & strTp(3) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 and FA02='0'"
            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,fa05||' '||fa63||' '||fa64||' '||fa65 as 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註 From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(upper(fa05||' '||fa63||' '||fa64||' '||fa65),'" & strTp(3) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 and FA02='0'"
            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,fa06 as 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註 From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa06,'" & strTp(3) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 and FA02='0'"
            
'            'Modify by Amy 2015/04/15 客戶端平台帳號資料
'            strSql = strSql & " union all Select ' ' as V,'平台'||CW01 AS 編號, CW12 AS 名稱,'平台' AS 國籍,' ' AS 智權人員,Nvl(CW19,'') AS 狀態,'' AS 備註 From CustWeb Where InStr(Upper(CW12),'" & strTp(3) & "') " & strCheckWay
            
            '查potcustomer 國外潛在客戶 檔
            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,PCU08 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註 From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(pcu08,'" & strTp(3) & "')" & strCheckWay & ") A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+) and PCU02='0'"
            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註 From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & strTp(3) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+) and PCU02='0'"
            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,PCU07 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註 From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(pCU07,'" & strTp(3) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+) and PCU02='0'"
            
            '查potcustomer1 國內潛在客戶 檔
            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,POC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,POC15 AS 備註 From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(poc03,'" & strTp(3) & "')" & strCheckWay & ") A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+) and POC02='0'"
            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,RTRIM(POC23||' '||POC24||' '||POC25||' '||POC26) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,POC15 AS 備註 From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & strTp(3) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+) and POC02='0'"
            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,POC27 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,POC15 AS 備註 From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(POC27,'" & strTp(3) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+) and POC02='0'"
            
'            '查NotAgent 不得代理案件之客戶或代理人 檔
'            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT02 as 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(nt02,'" & strTp(3) & "')" & strCheckWay & ") A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
'            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT03||' '||NT04||' '||NT05||' '||NT06 as 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(upper(nt03||' '||nt04||' '||nt05||' '||nt06),'" & strTp(3) & "')" & strCheckWay & " ) A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
'            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT07 as 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(nt07,'" & strTp(3) & "')" & strCheckWay & " ) A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
            
            '查聯絡人(中文)
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(pcc05,'" & strTp(3) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(pcc05,'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(pcc05,'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註 From (Select * From potcustcont Where instr(pcc05,'" & strTp(3) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
            
            '查聯絡人(英文)
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(upper(pcc03),'" & strTp(3) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(upper(pcc03),'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註 From (Select * From potcustcont Where instr(upper(pcc03),'" & strTp(3) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(upper(pcc03),'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
            
            '查聯絡人(日文)
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(PCC04,'" & strTp(3) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(PCC04,'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註 From (Select * From potcustcont Where instr(PCC04,'" & strTp(3) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註 From (Select * From potcustcont Where instr(PCC04,'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
'         End If
'
'         'E-Mail
'         strEmail = Trim(wksaccrpt114.Range("A" & iRow).Value) '完整E-Mail
'         If strEmail <> "" Then
'            strEmail2 = Mid(strEmail, InStr(strEmail, "@")) '網域
'            If UCase(strEmail2) = UCase("@gmail.com") Then
'               strEmail2 = ""
'            End If
'            '完整E-Mail
'            strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION,staff Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(strEmail))) & "')>0 or instr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(strEmail))) & "')>0 or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(strEmail))) & "')>0 or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(strEmail))) & "')>0 or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(strEmail))) & "') > 0 ) and CU10=NA01(+) AND CU13=ST01(+) and CU02='0'"
'            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,Decode(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註 FROM potcustomer,nation,staff Where (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(Trim(strEmail))) & "') >0 ) and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) and PCU02='0'"
'            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,NVL(PoC03,Decode(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,POC15 AS 備註 FROM potcustomer1,nation,staff Where (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(Trim(strEmail))) & "') >0 ) and poc04=na01(+) and poc13=ST01(+) and POC02='0'"
'            strSql = strSql & " union all SELECT ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,NVL(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註 FROM fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(strEmail))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(strEmail))) & "')> 0 or instr(NLS_Upper(fa105),'" & UCase(ChgSQL(Trim(strEmail))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(strEmail))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(strEmail))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(strEmail))) & "') > 0 ) and fa10=na01(+) and FA02='0'"
'            strSql = strSql & " union all SELECT ' ' AS V,PCC01||'0-'||PCC02 AS 編號,NVL(PCC05,NVL(PCC03,PCC04)) AS 名稱,' ' AS 國籍,' ' AS 智權人員,' ' AS 狀態,PCC13 AS 備註 FROM PotCustCont Where (instr(NLS_Upper(PCC08),'" & UCase(ChgSQL(Trim(strEmail))) & "') > 0) and PCC02='0'"
            If adoRecordset.State = adStateOpen Then
               adoRecordset.Close
            End If
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While Not adoRecordset.EOF
                  If strData <> "" Then
                     strData = strData & vbCrLf
                  End If
                  strData = strData & adoRecordset.Fields(1) & adoRecordset.Fields(2) '& "-" & adoRecordset.Fields(4)
                  adoRecordset.MoveNext
               Loop
'            ElseIf strEmail2 <> "" Then
'               wksaccrpt114.Range("B" & iRow).Value = "*"
'               '用網域查資料
'               strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION,staff Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(strEmail2))) & "')>0 or instr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(strEmail2))) & "')>0 or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(strEmail2))) & "')>0 or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(strEmail2))) & "')>0 or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(strEmail2))) & "') > 0 ) and CU10=NA01(+) AND CU13=ST01(+) and CU02='0'"
'               strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,Decode(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註 FROM potcustomer,nation,staff Where (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(Trim(strEmail2))) & "') >0 ) and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) and PCU02='0'"
'               strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,NVL(PoC03,Decode(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,POC15 AS 備註 FROM potcustomer1,nation,staff Where (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(Trim(strEmail2))) & "') >0 ) and poc04=na01(+) and poc13=ST01(+) and POC02='0'"
'               strSql = strSql & " union all SELECT ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,NVL(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註 FROM fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(strEmail2))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(strEmail2))) & "')> 0 or instr(NLS_Upper(fa105),'" & UCase(ChgSQL(Trim(strEmail2))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(strEmail2))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(strEmail2))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(strEmail2))) & "') > 0 ) and fa10=na01(+) and FA02='0'"
'               strSql = strSql & " union all SELECT ' ' AS V,PCC01||'0-'||PCC02 AS 編號,NVL(PCC05,NVL(PCC03,PCC04)) AS 名稱,' ' AS 國籍,' ' AS 智權人員,' ' AS 狀態,PCC13 AS 備註 FROM PotCustCont Where (instr(NLS_Upper(PCC08),'" & UCase(ChgSQL(Trim(strEmail2))) & "') > 0) and PCC02='0'"
'               If adoRecordset.State = adStateOpen Then
'                  adoRecordset.Close
'               End If
'               adoRecordset.CursorLocation = adUseClient
'               adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'               If adoRecordset.RecordCount > 0 Then
'                  adoRecordset.MoveFirst
'                  Do While Not adoRecordset.EOF
'                     If strData <> "" Then
'                        strData = strData & vbCrLf
'                     End If
'                     strData = strData & adoRecordset.Fields(1) & adoRecordset.Fields(2) '& "-" & adoRecordset.Fields(4)
'                     adoRecordset.MoveNext
'                  Loop
'               End If
            End If
            '寫回Excel檔
            If strData <> "" Then
               'strData = Right(strData, Len(strData) - 1)
               wksaccrpt114.Range("C" & iRow).Value = strData
            End If
         End If
ReadNext:
         iRow = iRow + 1
      End If
   Loop
   adoRecordset.Close
   '存檔
   xlsSalesPoint.Workbooks(1).SaveAs FileName:=Left(strFileName, Len(strFileName) - Len(strFileName2)) & "_new" & strFileName2
   
   '關閉
   xlsSalesPoint.Workbooks.Close
   '離開
   xlsSalesPoint.Quit
   Set wksaccrpt114 = Nothing
   Set xlsSalesPoint = Nothing
   Screen.MousePointer = vbDefault
   
   MsgBox "資料過濾完畢！"
   
   Exit Sub
   
flgErr:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Added by Lydia 2019/12/24 English_Vers案件清單:
Private Sub Command33_Click()
Dim xlsReport
Dim wksReport
Dim lngRow As Long
Dim strType As String
Dim strTemp(1 To 3) As String

'產生清單
'1.用cmd產生清單(12/20):因為VB用比對資料夾或檔名型態遇見Unicode會程式出錯
'2.人工處理清單:1-P案 -大陸案:以前刪除
'216\21694: 以前刪除
'0-KOIKE: 以前刪除
'0-Ushiki線上下載: 刪除
'KSI侵害鑑定資料: 刪除
'TRACKING_NO: 刪除
'確認FMP的Tracking:
'p -119709
'可以匯入
'100074
'107644
'107645
'107647
'120285 以後
'3.匯入成Lydia_A003，逐筆讀取R003經過VB有產生問號則列出
   
   strType = "1" '  1- 分成3欄, 2-只顯示全檔名1欄
   strSql = "select R003 from lydia_a003 where r001=20191220 and instr(r003,'.') > 0 order by r003 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
       With RsTemp
           .MoveFirst
           Do While Not .EOF
               strExc(0) = "" & .Fields("r003")
               If strExc(0) <> "" And InStr(strExc(0), "?") > 0 Then '經過VB有產生問號則列出
                   If lngRow = 0 Then
                        Set xlsReport = CreateObject("Excel.Application")
                        xlsReport.SheetsInNewWorkbook = 1
                        xlsReport.Workbooks.add
                        xlsReport.Visible = True
                        
                        lngRow = 1
                        Set wksReport = xlsReport.Worksheets(1)
                        wksReport.Cells.NumberFormatLocal = "@"
                        If strType = "2" Then
                            wksReport.Range("A:A").ColumnWidth = 80
                        Else
                            wksReport.Range("A:A").ColumnWidth = 30
                            wksReport.Range("B:B").ColumnWidth = 30
                            wksReport.Range("C:C").ColumnWidth = 80
                        End If
                        '抬頭
                        If strType = "2" Then
                            wksReport.Range("A" & lngRow).Value = "完整檔名路徑"
                        Else
                            strTemp(1) = "資料夾路徑"
                            strTemp(2) = "檔案名稱"
                            strTemp(3) = "完整檔名路徑"
                            wksReport.Range("A" & lngRow & ":" & "C" & lngRow).Value = strTemp
                        End If
                        lngRow = lngRow + 1
                        wksReport.Range("B2").Select
                        xlsReport.ActiveWindow.FreezePanes = True '凍結窗格
                        wksReport.Range("A1").Select
                   End If
                   If strType = "2" Then
                        wksReport.Range("A" & lngRow).Value = strExc(0)
                   Else
                        intI = InStrRev(strExc(0), "\")
                        strExc(1) = Mid(strExc(0), 1, intI - 1)
                        strExc(2) = Mid(strExc(0), intI + 1)
                        strTemp(1) = strExc(1)
                        strTemp(2) = strExc(2)
                        strTemp(3) = strExc(0)
                        wksReport.Range("A" & lngRow & ":" & "C" & lngRow).Value = strTemp
                   End If
                   lngRow = lngRow + 1
               End If
               .MoveNext
           Loop
       End With
       If lngRow > 0 Then
          xlsReport.Workbooks(1).SaveAs PUB_Getdesktop & "\" & Command33.Caption & "_" & strSrvDate(1) & Format(ServerTime, "000000")
          xlsReport.Workbooks.Close
          xlsReport.Quit
       End If
   End If
   
   MsgBox "OK!"
End Sub

'Added by Lydia 2020/01/07 更換Unicode的檔名
Private Sub Command34_Click()
Dim intP As Integer, intK As Integer, intU As Integer
Dim strNewName As String
Dim strDefPath As String
Dim strPass As String
Dim fs, fso, fl
Dim rsAD As New ADODB.Recordset
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim strKey As String
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, m_CP10 As String
Dim nCP09 As String 'D類收文號
Dim m_TempDir As String 'Key之前的路徑
Dim m_TempName As String  '統一名稱
Dim nMax As Integer
Dim nPos As String

'-----Test 2020/03/02
    'Table 只記錄第一層資料夾路徑, 之後直接抓目前檔案
    strExc(0) = "select * from lydia_a002 where r003=20200302 and r004 is null order by r001"
    intI = 1
    Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        With rsAD
            .MoveFirst
            Do While Not .EOF
                m_TempDir = ""
                strDefPath = "" & .Fields("r002")
                nPos = "" & .Fields("r001")
                If Right(strDefPath, 1) <> "\" Then strDefPath = strDefPath & "\" '與抓子資料夾有關
                
'English_vers分析：本所案號、上傳類型
                strKey = "\ENGLISH_VERS\"
                intK = InStr(UCase(strDefPath), strKey)
                If intK > 0 Then
                    m_CP10 = cntEnglish_Vers
                    strExc(1) = Mid(strDefPath, intK + Len(strKey))
                    tmpArr1 = Split(strExc(1), "\") '用\區隔路徑層級
                    If Len(tmpArr1(0)) = 6 Then  'FMP案
                        m_CP01 = "P"
                        m_CP02 = tmpArr1(0)
                    ElseIf Len(tmpArr1(0)) = 3 Then 'FCP案
                        m_CP01 = "FCP"
                        m_CP02 = tmpArr1(0)
                    End If
                    m_CP03 = "0": m_CP04 = "00"
                
                    If m_CP01 = "FCP" And Len(m_CP01 & m_CP02 & m_CP03 & m_CP04) < 12 Then
                        '先將所有案號資料夾記錄在字串
                        strExc(9) = ""
                        strPass = Dir(strDefPath, vbDirectory)
                        Do While strPass <> ""
                             If strPass <> "." And strPass <> ".." Then
                                 If GetAttr(strDefPath & strPass) = vbDirectory Then
                                      strExc(9) = strExc(9) & "," & strPass
                                 End If
                             End If
                             strPass = Dir()
                        Loop
                        If strExc(9) <> "" Then
                            tmpArr1 = Empty
                            tmpArr1 = Split(Mid(strExc(9), 2), ",")
                            nMax = UBound(tmpArr1) + 1
                            'm_TempDir = m_CP01 & Format(tmpArr1(0), "000000") & m_CP03 & m_CP04
                            m_CP02 = Format(tmpArr1(0), "000000")
                            m_TempDir = strDefPath & tmpArr1(0) & "\"
                        End If
                    ElseIf m_CP01 = "P" And Len(m_CP01 & m_CP02 & m_CP03 & m_CP04) = 10 Then
                         nMax = 1
                         m_TempDir = strDefPath
                    End If
                    
                    '逐案號資料夾上傳
                    For intI = 1 To nMax
                        If intI > 1 Then
                            m_CP02 = Format(tmpArr1(intI - 1), "000000")
                            m_TempDir = strDefPath & tmpArr1(intI - 1) & "\"
                        End If
                        nCP09 = ""
                        intP = 0
                        '1.先拿掉Unicode字
                         Set fs = CreateObject("Scripting.FileSystemObject")
                         Set fso = fs.GetFolder(m_TempDir)
                         For Each fl In fso.files
                            TxtFile.Text = fl.Name
                            If TxtFile.Text <> fl.Name Then
                                strNewName = Replace(TxtFile.Text, "?", "x")
                                '指定更換
                                fl.Name = strNewName
                            End If
                         Next
                         '2.上傳檔案，若有子資料夾則壓縮為.zip
JumpToReDir:
                         strPass = Dir(m_TempDir, vbDirectory)
                         intP = 0: intU = 0
                         Do While strPass <> ""
                            If strPass <> "." And strPass <> ".." Then
                                TxtFile.Text = strPass
                                If InStr(TxtFile.Text, ".") = 0 And InStr(TxtFile.Text, "?") > 0 Then '子資料夾有Unicode
                                    Debug.Print "有Unicode: " & m_TempDir & TxtFile.Text
                                    intU = intU + 1
                                    GoTo JumpToPass
                                End If
                                strNewName = ""
'                                If GetAttr(m_TempDir & strPass) = vbDirectory Then
'                                    strNewName = m_TempDir & Mid(strPass, 1, 20) & "." & Format(FileDateTime(m_TempDir & strPass), "YYYYMMDDHHMMSS")
'                                    If ZipFolder(m_TempDir & strPass, strNewName) = True Then
'                                          Debug.Print "壓縮檔成功: " & strNewName
'                                          Call PUB_KillTempFolder(strPass, m_TempDir)
'                                          intP = 0: intU = 0 '因為重讀資料夾,計數歸0
'                                          GoTo JumpToReDir  '因為刪除子資料夾所以要重新查詢，不然子資料夾後的檔案無法讀到; 可是清單會重跑一次
'                                    Else
'                                          Debug.Print "壓縮檔失敗: " & strNewName
'                                    End If
'                                Else
'                                    Debug.Print "檔案: " & m_TempDir & strPass
'                                End If
                                If GetAttr(m_TempDir & strPass) = vbDirectory Then
                                    strNewName = m_TempDir & Mid(strPass, 1, 20) & "." & Format(FileDateTime(m_TempDir & strPass), "YYYYMMDDHHMMSS")
                                    If ZipFolder(m_TempDir & strPass, strNewName) = True Then
                                          'Debug.Print "壓縮檔成功: " & strNewName
                                          strNewName = strNewName & ".zip"
                                          Call PUB_KillTempFolder(strPass, m_TempDir)
                                          intP = 0: intU = 0 '因為重讀資料夾,計數歸0
                                          GoTo JumpToReDir  '因為刪除子資料夾所以要重新查詢，不然子資料夾後的檔案無法讀到
                                    Else
                                          'Debug.Print "壓縮檔失敗: " & strNewName
                                          strNewName = ""
                                    End If
                                Else
                                    If UCase(strPass) = UCase("Thumbs.db") Then '刪除-瀏覽縮圖暫存檔
                                       Debug.Print "刪除:" & m_TempDir & strPass
                                       Kill m_TempDir & strPass
                                       strNewName = ""
                                    Else
                                       'Debug.Print "檔案: " & m_tempdir & strPass
                                       strNewName = m_TempDir & strPass
                                    End If
                                End If
                                If strNewName <> "" Then '上傳檔案
                                    strExc(6) = "": strExc(2) = ""
                                    If PUB_UploadCPFfile("0", strNewName, m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, nCP09, , , , strExc(6), strExc(2)) = True Then
                                         Debug.Print "上傳成功：" & nCP09 & " " & strExc(2)
                                    Else
                                         Debug.Print "上傳失敗：" & nCP09 & " " & strExc(6)
                                    End If
                                End If
                                 intP = intP + 1
                            End If
JumpToPass:                         '計數：intP
                            strPass = Dir()
                         Loop
                         '上傳完後,直接刪除案號資料夾
                         If intU = 0 Then
                             intP = 0
                             strExc(2) = Dir(m_TempDir, vbDirectory)
                             Do While strExc(2) <> ""
                                  If strExc(2) <> "." And strExc(2) <> ".." Then
                                      If GetAttr(m_TempDir & strExc(2)) = vbDirectory Then
                                           intP = intP + 1
                                      Else
                                           intP = intP + 1
                                      End If
                                  End If
                                  strExc(2) = Dir()
                             Loop
                             If intP = 0 Then
                                 'Debug.Print "刪除資料夾:" & m_TempDir
                                 m_TempDir = Mid(m_TempDir, 1, Len(m_TempDir) - 1)
                                 Call PUB_KillTempFolder(Val(m_CP02), Mid(m_TempDir, 1, InStrRev(m_TempDir, "\") - 1))
                             End If
                         End If
                         '隔日,人工刪除第一層資料夾
                    Next intI
                End If
                
'專利案件分析：本所案號
                strKey = "\專利案件\"
                intK = InStr(UCase(strDefPath), strKey)
                If intK > 0 Then
                    m_CP01 = "FCP"
                    m_CP10 = cnt專利案件
                    m_CP03 = "0": m_CP04 = "00"
                    strExc(9) = ""
                    strExc(1) = Mid(strDefPath, intK + Len(strKey))
                    tmpArr1 = Split(strExc(1), "\") '用\區隔路徑層級
                    If Len(tmpArr1(0)) = 3 Then '前3碼相同:放同一層
                        strExc(9) = strExc(9) & "," & strDefPath
                    ElseIf Len(tmpArr1(0)) = 7 Then 'ex. 前3碼200_299: 底下再分前3碼子資料夾
                        '讀取:前3碼子資料夾
                        strPass = Dir(strDefPath, vbDirectory)
                        Do While strPass <> ""
                             If strPass <> "." And strPass <> ".." Then
                                 If GetAttr(strDefPath & strPass) = vbDirectory Then
                                      strExc(9) = strExc(9) & "," & strDefPath & strPass
                                 End If
                             End If
                             strPass = Dir()
                        Loop
                        'If strExc(9) = "" Then strExc(9) = "," & strDefPath
                    End If
                    '逐案號資料夾上傳
                    If strExc(9) <> "" Then
                        tmpArr1 = Empty
                        tmpArr1 = Split(Mid(strExc(9), 2), ",")
                        nMax = UBound(tmpArr1)
                        
                        For intI = 0 To nMax
                            m_TempDir = Trim(tmpArr1(intI))
                            If Right(m_TempDir, 1) <> "\" Then m_TempDir = m_TempDir & "\" '與抓子資料夾有關
                            
                            nCP09 = ""
                            intP = 0
                            '1.先拿掉Unicode字
                             Set fs = CreateObject("Scripting.FileSystemObject")
                             Set fso = fs.GetFolder(m_TempDir)
                             For Each fl In fso.files
                                TxtFile.Text = fl.Name
                                If TxtFile.Text <> fl.Name Then
                                    strNewName = Replace(TxtFile.Text, "?", "x")
                                    '指定更換
                                    fl.Name = strNewName
                                End If
                             Next
                             '2.上傳檔案，若有子資料夾則壓縮為.zip
JumpToReDir2:
                             strPass = Dir(m_TempDir, vbDirectory)
                             intP = 0: intU = 0
                             Do While strPass <> ""
                                If strPass <> "." And strPass <> ".." Then
                                    TxtFile.Text = strPass
                                    If InStr(TxtFile.Text, ".") = 0 And InStr(TxtFile.Text, "?") > 0 Then '子資料夾有Unicode
                                        Debug.Print "有Unicode: " & m_TempDir & TxtFile.Text
                                        intU = intU + 1
                                        GoTo JumpToPass2
                                    End If
                                    strNewName = ""
'                                    If GetAttr(m_TempDir & strPass) = vbDirectory Then
'                                        strNewName = m_TempDir & Mid(strPass, 1, 20) & "." & Format(FileDateTime(m_TempDir & strPass), "YYYYMMDDHHMMSS")
'                                        If ZipFolder(m_TempDir & strPass, strNewName) = True Then
'                                              Debug.Print "壓縮檔成功: " & strNewName
'                                              Call PUB_KillTempFolder(strPass, m_TempDir)
'                                              intP = 0: intU = 0 '因為重讀資料夾,計數歸0
'                                              GoTo JumpToReDir2  '因為刪除子資料夾所以要重新查詢，不然子資料夾後的檔案無法讀到; 可是清單會重跑一次
'                                        Else
'                                              Debug.Print "壓縮檔失敗: " & strNewName
'                                        End If
'                                    Else
'                                        Debug.Print "檔案: " & m_TempDir & strPass
'                                    End If
                                    If GetAttr(m_TempDir & strPass) = vbDirectory Then
                                        strNewName = m_TempDir & Mid(strPass, 1, 20) & "." & Format(FileDateTime(m_TempDir & strPass), "YYYYMMDDHHMMSS")
                                        If ZipFolder(m_TempDir & strPass, strNewName) = True Then
                                              'Debug.Print "壓縮檔成功: " & strNewName
                                              strNewName = strNewName & ".zip"
                                              Call PUB_KillTempFolder(strPass, m_TempDir)
                                              intP = 0: intU = 0 '因為重讀資料夾,計數歸0
                                              GoTo JumpToReDir2  '因為刪除子資料夾所以要重新查詢，不然子資料夾後的檔案無法讀到
                                        Else
                                              'Debug.Print "壓縮檔失敗: " & strNewName
                                              strNewName = ""
                                        End If
                                    Else
                                        If UCase(strPass) = UCase("Thumbs.db") Then '刪除-瀏覽縮圖暫存檔
                                           Debug.Print "刪除:" & m_TempDir & strPass
                                           Kill m_TempDir & strPass
                                           strNewName = ""
                                        Else
                                           'Debug.Print "檔案: " & m_tempdir & strPass
                                           strNewName = m_TempDir & strPass
                                        End If
                                    End If
                                    If strNewName <> "" Then '上傳檔案
                                        strExc(1) = Mid(strNewName, InStrRev(strNewName, "\") + 1)
                                        If InStr(strExc(1), "FCP0") = 1 Then 'FCP開頭6碼
                                           m_CP02 = Mid(strExc(1), 4, 6)
                                        ElseIf InStr(strExc(1), "FCP") = 1 Then 'FCP開頭5碼
                                           m_CP02 = Format(Val(Mid(strExc(1), 4, 5)), "000000")
                                        Else
                                           m_CP02 = Format(Val(Mid(strExc(1), 1, 5)), "000000")
                                        End If
                                        strExc(6) = "": strExc(2) = ""
                                        nCP09 = "" '因為專利案件不是一個案號一個資料夾,所以預設空白來抓案件是否有收文專利案件991
                                        If PUB_UploadCPFfile("0", strNewName, m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, nCP09, , , , strExc(6), strExc(2)) = True Then
                                             Debug.Print "上傳成功：" & nCP09 & " " & strExc(2)
                                        Else
                                             Debug.Print "上傳失敗：" & nCP09 & " " & strExc(6)
                                        End If
                                    End If
                                     intP = intP + 1
                                End If
JumpToPass2:                             '計數：intP
                                strPass = Dir()
                             Loop
                            '上傳完後,直接刪除案號資料夾
                            If intU = 0 Then
                                intP = 0
                                strExc(2) = Dir(m_TempDir, vbDirectory)
                                Do While strExc(2) <> ""
                                     If strExc(2) <> "." And strExc(2) <> ".." Then
                                         If GetAttr(m_TempDir & strExc(2)) = vbDirectory Then
                                              intP = intP + 1
                                         Else
                                              intP = intP + 1
                                         End If
                                     End If
                                     strExc(2) = Dir()
                                Loop
                                If intP = 0 Then
                                    'Debug.Print "刪除資料夾:" & m_TempDir
                                    m_TempDir = Mid(m_TempDir, 1, Len(m_TempDir) - 1)
                                    Call PUB_KillTempFolder(Mid(m_TempDir, InStrRev(m_TempDir, "\") + 1), Mid(m_TempDir, 1, InStrRev(m_TempDir, "\") - 1))
                                End If
                            End If
                            '隔日,人工刪除第一層資料夾
                        Next intI
                    End If
                End If
                cnnConnection.Execute " update lydia_a002 set r004=" & strSrvDate(1) & " where r001='" & nPos & "' " '記錄已處理
                .MoveNext
            Loop
        End With
    End If
     
     MsgBox "OK"
Exit Sub
'------Test 2020/02/24
    'strDefPath = "C:\Users\A3034\Desktop\English_Vers\210\21032"  'FCP021032=FCP021031+FCP056602集結Unicode, 子資料夾, 過長檔名
    strDefPath = "\\typing2\English_Vers\210\21032" '嘗試不複製到本機端
    If Right(strDefPath, 1) <> "\" Then strDefPath = strDefPath & "\" '與抓子資料夾有關
    
    '分析：本所案號、上傳類型
    strKey = "\ENGLISH_VERS\"
    intK = InStr(UCase(strDefPath), strKey)
    If intK > 0 Then
        m_CP10 = cntEnglish_Vers
        m_TempDir = Mid(strDefPath, 1, intK + Len(strKey) - 1)
        strExc(1) = Mid(strDefPath, intK + Len(strKey))
        tmpArr1 = Split(strExc(1), "\") '用\區隔路徑層級
        nMax = UBound(tmpArr1)
        strExc(2) = tmpArr1(nMax)
        If Len(tmpArr1(0)) = 6 Then  'FMP案
            m_CP01 = "P"
            m_CP02 = tmpArr1(0)
        ElseIf Len(tmpArr1(1)) = 5 Then 'FCP案
            m_CP01 = "FCP"
            m_CP02 = Format(tmpArr1(1), "000000")
        End If
        If m_CP01 <> "" Then
            m_CP03 = "0": m_CP04 = "00"
        Else
            MsgBox "分析案號出錯!!"
            Exit Sub
        End If
    End If
    
    nCP09 = ""
    intP = 0
    '1.先拿掉Unicode字
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set fso = fs.GetFolder(strDefPath)
     For Each fl In fso.files
        TxtFile.Text = fl.Name
        If TxtFile.Text <> fl.Name Then
            strNewName = Replace(TxtFile.Text, "?", "x")
            '指定更換
            fl.Name = strNewName
        End If
     Next
     '2.上傳檔案，若有子資料夾則壓縮為.zip
JumpToReDir1:
     strPass = Dir(strDefPath, vbDirectory)
     intP = 0
     Do While strPass <> ""
        If strPass <> "." And strPass <> ".." Then
            TxtFile.Text = strPass
            If InStr(TxtFile.Text, ".") = 0 And InStr(TxtFile.Text, "?") > 0 Then '子資料夾有Unicode
                Debug.Print "有Unicode: " & strDefPath & TxtFile.Text
                intP = intP + 1
                GoTo JumpToPass1
            End If
            strNewName = ""
'            If GetAttr(strDefPath & strPass) = vbDirectory Then
'                strNewName = strDefPath & Mid(strPass, 1, 20) & "." & Format(FileDateTime(strDefPath & strPass), "YYYYMMDDHHMMSS")
'                If ZipFolder(strDefPath & strPass, strNewName) = True Then
'                      Debug.Print "壓縮檔成功: " & strNewName
'                      Call PUB_KillTempFolder(strPass, strDefPath)
'                      GoTo JumpToReDir  '因為刪除子資料夾所以要重新查詢，不然子資料夾後的檔案無法讀到; 可是清單會重跑一次
'                Else
'                      Debug.Print "壓縮檔失敗: " & strNewName
'                End If
'            Else
'                Debug.Print "檔案: " & strDefPath & strPass
'            End If
            If GetAttr(strDefPath & strPass) = vbDirectory Then
                strNewName = strDefPath & Mid(strPass, 1, 20) & "." & Format(FileDateTime(strDefPath & strPass), "YYYYMMDDHHMMSS")
                If ZipFolder(strDefPath & strPass, strNewName) = True Then
                      'Debug.Print "壓縮檔成功: " & strNewName
                      strNewName = strNewName & ".zip"
                      Call PUB_KillTempFolder(strPass, strDefPath)
                      GoTo JumpToReDir1  '因為刪除子資料夾所以要重新查詢，不然子資料夾後的檔案無法讀到
                Else
                      'Debug.Print "壓縮檔失敗: " & strNewName
                      strNewName = ""
                End If
            Else
                If UCase(strPass) = UCase("Thumbs.db") Then '刪除-瀏覽縮圖暫存檔
                   Kill strDefPath & strPass
                   strNewName = ""
                Else
                   'Debug.Print "檔案: " & strDefPath & strPass
                   strNewName = strDefPath & strPass
                End If
            End If
            If strNewName <> "" Then '上傳檔案
                strExc(6) = "": strExc(2) = ""
                If PUB_UploadCPFfile("0", strNewName, m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, nCP09, , , , strExc(6), strExc(2)) = True Then
                     Debug.Print "上傳成功：" & nCP09 & " " & strExc(2)
                Else
                     Debug.Print "上傳失敗：" & nCP09 & " " & strExc(6)
                End If
            End If
        End If
JumpToPass1: '計數：intP
        strPass = Dir()
     Loop
     
     If intP = 0 Then
         strDefPath = Mid(strDefPath, 1, Len(strDefPath) - 1)
         Call PUB_KillTempFolder(Val(m_CP02), Mid(strDefPath, 1, InStrRev(strDefPath, "\") - 1))
     End If
     
     MsgBox "OK"
     Exit Sub
'----------------------
    'strDefPath = "\\typing2\English_Vers\566\56602\" '嘗試不複製到本機端
    strPass = Dir(strDefPath, vbDirectory)
    Do While strPass <> ""
        If strPass <> "." And strPass <> ".." Then
            If GetAttr(strDefPath & strPass) = vbDirectory Then
                Debug.Print "資料夾: " & strDefPath & strPass
                'Debug.Print "建立時間: " & Format(FileDateTime(strDefPath & strPass), "YYYYMMDDHHMMSS")
                strNewName = strDefPath & Mid(strPass, 1, 20) & "." & Format(FileDateTime(strDefPath & strPass), "YYYYMMDDHHMMSS") & ".zip"
                Debug.Print "壓縮檔: " & strNewName
            Else
                Debug.Print "檔案: " & strDefPath & strPass
            End If
        End If
        strPass = Dir()
    Loop
MsgBox "OK"
Exit Sub
'---------
  '測試壓縮檔OK
'  If ZipFolder("C:\Users\A3034\Desktop\English_Vers\566\56602\Disclosure 16-0659-TW-NP - Request for New Application-Mar 31 2017-12 31 PM", "C:\Users\A3034\Desktop\English_Vers\566\56602\" & strSrvDate(1) & "_" & Format(ServerTime, "000000")) = True Then
'      MsgBox "OK"
'  Else
'      MsgBox "Error"
'  End If
'  Exit Sub
  
'---------分析為資料夾路徑或檔案路徑
  
  '來源: dir清單內含資料夾、子資料夾和檔案路徑
  strSql = "select * from lydia_a20200204 " 'where r003 >=168 and r003<=184 " ' where r004 like 'C:\Users\A3034\Desktop\English_Vers\566\56602%' "
  strSql = strSql & "order by r003 "
  'strDefPath = "C:\Users\A3034\Desktop\English_Vers\566\56602"
  m_CP03 = "0": m_CP04 = "00"
  
  intI = 1
  Set rsAD = ClsLawReadRstMsg(intI, strSql)
  If intI = 1 Then
     rsAD.MoveFirst
     Do While Not rsAD.EOF
         strDefPath = "" & rsAD.Fields("r004")
         '找English_Vers的位置
         strKey = "\ENGLISH_VERS\"
         intK = InStr(UCase(strDefPath), strKey)
         If intK > 0 Then
            m_CP10 = cntEnglish_Vers
            m_TempDir = Mid(strDefPath, 1, intK + Len(strKey) - 1)
            strExc(1) = Mid(strDefPath, intK + Len(strKey))
            tmpArr1 = Empty
            tmpArr1 = Split(strExc(1), "\") '用\區隔路徑層級
            nMax = UBound(tmpArr1)
            strExc(2) = tmpArr1(nMax)
            'PUB_GetFileListOrderby 讀取檔名排序
            If Len(tmpArr1(0)) = 6 Then  'FMP案
               m_CP01 = "P"
               m_CP02 = tmpArr1(0)
               'If UBound(tmpArr1) = 0 And InStr(tmpArr1(0), ".") = 0 Then
               If InStr(strExc(2), ".") = 0 Then
                   'Debug.Print "資料夾路徑: " & strDefPath
                    '指定資料夾, 即時讀取檔案
                    If nMax > 1 Then  '案號有子資料夾
                        GoTo JumpToNext
                    Else
                        Debug.Print "資料夾路徑: " & strDefPath
                    End If
               Else
                   'Debug.Print "檔案路徑: " & strDefPath
               End If
            ElseIf Len(tmpArr1(0)) = 3 Then 'FCP案
               m_CP01 = "FCP"
               If InStr(strExc(2), ".") = 0 Then
                    '指定資料夾, 即時讀取檔案
                    If nMax < 1 Or nMax > 1 Then '前3碼, 案號有子資料夾
                        If nMax < 1 Then Debug.Print "前3碼: " & strDefPath
                        GoTo JumpToNext
                    Else
                        m_CP01 = tmpArr1(1)
                        Debug.Print "資料夾路徑: " & strDefPath
                    End If
               Else
                    GoTo JumpToNext '只抓案號資料夾
                    m_CP02 = tmpArr1(1)
                    '將子資料夾的檔案搬到上一層,直到案號資料
                    If nMax > 2 Then
                        For intP = (nMax - 1) To 2 Step -1
                           strExc(2) = tmpArr1(intP) & "-" & strExc(2)
                        Next intP
                    End If

                    'Debug.Print "Old: " & strDefPath
                    strExc(3) = PUB_GetReNameMax(strExc(2), m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, 75)
                    'Debug.Print "New: " & m_TempDir & tmpArr1(0) & "\" & tmpArr1(1) & "\" & strExc(3)
                    m_TempName = ""
                    '上傳自動+本所案號 PUB_UploadCPFfile
                    'If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, strExc(3), m_TempName, True, 0) = False Then
                    '     Debug.Print "except: " & strExc(3)
                    'Else
                         'Debug.Print "New: " & m_TempDir & tmpArr1(0) & "\" & tmpArr1(1) & "\" & m_TempName
                         cnnConnection.Execute "Update Lydia_a20200204 set R005='" & ChgSQL(m_TempDir & tmpArr1(0) & "\" & tmpArr1(1) & "\" & strExc(3)) & "' where R003=" & rsAD.Fields("R003")
                    'End If

               End If
            End If
         End If
         '找專利案件的位置
         strKey = "\專利案件\"
         intK = InStr(UCase(strDefPath), strKey)
         If intK > 0 Then
            m_CP10 = cnt專利案件
            strExc(1) = Mid(strDefPath, intK + Len(strKey))
            tmpArr1 = Empty
            tmpArr1 = Split(strExc(1), "\") '用\區隔路徑層級
            'PUB_GetFileListOrderby 讀取檔名排序
            If Len(tmpArr1(0)) = 3 Then
               If UBound(tmpArr1) > 0 Then
                   Debug.Print "檔案路徑: " & strDefPath
               Else
                   Debug.Print "資料夾路徑: " & strDefPath
               End If
            End If
         End If
JumpToNext:
         rsAD.MoveNext
     Loop
  End If
  MsgBox "OK"
'---------分析為資料夾路徑或檔案路徑

'---找目錄
'    strDefPath = "C:\Users\A3034\Desktop\English_Vers\"
'    strExc(1) = Dir(strDefPath, vbDirectory)
'    Do While strExc(1) <> ""
'         intP = intP + 1
'         If strExc(1) <> "." And strExc(1) <> ".." Then
'             Debug.Print Format(intP, "00") & ". " & strExc(1)
'         End If
'         strExc(1) = Dir
'    Loop
'MsgBox "OK"
'Exit Sub

'指定更換檔名
'     strDefPath = "C:\Users\A3034\Desktop\English_Vers\565\56506"
'     Set fs = CreateObject("Scripting.FileSystemObject")
'     Set fso = fs.GetFolder(strDefPath)
'     For Each fl In fso.files
'           TxtFile.Text = fl.Name
'           If TxtFile.Text <> fl.Name Then
'                'strNewName = Replace(TxtFile.Text, "?", "x")
'                '指定更換
'                strNewName = fl.Name
'                intI = 1
'                strSql = "select * from lydia_a001 order by seqid"
'                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                If intI = 1 Then
'                    RsTemp.MoveFirst
'                    Do While Not RsTemp.EOF
'                        '改Provider, 丟到FM20.TextBox
'                        txtFM2(0).Text = "" & RsTemp.Fields("R001")
'                        txtFM2(1).Text = "" & RsTemp.Fields("R002")
'                        strNewName = Replace(strNewName, txtFM2(0).Text, txtFM2(1).Text)
'                        RsTemp.MoveNext
'                    Loop
'                End If
'                fl.Name = strNewName
'                Debug.Print "Before: " & TxtFile.Text & vbCrLf & "After: " & strNewName
'                intP = intP + 1
'           End If
'     Next
'
'     MsgBox "OK", vbInformation
     '經理: 去掉Thumb.
End Sub

'Added by Lydia 2020/02/11 將資料夾壓縮為.zip ; 參考frm170104
Private Function ZipFolder(pFolder As String, pNewFolder As String) As Boolean
   Dim program_name As String, program_path As String
   Dim process_id As Long
   Dim process_handle As Long

   program_name = "C:\Program Files\7-Zip\7z.exe"
   '檢查執行檔
   If Dir(program_name) = "" Then
      MsgBox "未安裝 7-Zip 程式，壓縮檔產生失敗！。"
      Exit Function
   End If
    
On Error GoTo ShellError
        
   '刪除舊檔 '2020/02/25 不刪除舊檔,因為子資料夾有可能是壓縮檔解開來的,兩者皆要保留
   'If Dir(pFolder & ".zip") <> "" Then
   '   Kill pFolder & ".zip"
   'End If
   If Dir(pFolder & ".zip") <> "" Then
       pNewFolder = pNewFolder & ".New"
   End If
   
   '-y 指有相同檔案存在時, 直接覆蓋. 不給的話會需要在Console 給 yes/no. 適用於Automation
   '-p 解壓縮密碼
   'process_id = Shell("""" & program_name & """ a -pCTCB """ & pNewFolder & ".zip"" """ & pFolder & "\*""", vbNormalNoFocus)
   process_id = Shell("""" & program_name & """ a """ & pNewFolder & ".zip"" """ & pFolder & "\*""", vbNormalNoFocus)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
        ZipFolder = True
    End If
    Exit Function

ShellError:
    ZipFolder = False
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Function

Private Sub Command36_Click()
Dim strVal As String
Dim ii As Integer

   strSql = "select * from acc0w0 where a0w06 is not null and nvl(a0w16,0)=0"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         If IsNumeric(adoRecordset.Fields("a0w06").Value) Then
            strVal = adoRecordset.Fields("a0w06").Value
            strSql = "update acc0w0 set a0w16=" & strVal & _
            " where a0w01=" & adoRecordset.Fields("a0w01").Value & " and a0w02='" & adoRecordset.Fields("a0w02").Value & "'"
            cnnConnection.Execute strSql
         Else
            strVal = ""
            For ii = 1 To Len(adoRecordset.Fields("a0w06").Value)
               If IsNumeric(Mid(adoRecordset.Fields("a0w06").Value, ii, 1)) Then
                  strVal = strVal & Mid(adoRecordset.Fields("a0w06").Value, ii, 1)
               Else
                  Exit For
               End If
            Next ii
            If strVal <> "" And IsNumeric(strVal) Then
               strSql = "update acc0w0 set a0w16=" & strVal & _
               " where a0w01=" & adoRecordset.Fields("a0w01").Value & " and a0w02='" & adoRecordset.Fields("a0w02").Value & "'"
               cnnConnection.Execute strSql
            End If
         End If
         adoRecordset.MoveNext
      Loop
   End If
   MsgBox "更新acc0w0-a0w16結束!!"
End Sub

'Added by Morgan 2020/5/5
'用 7z.exe 檔案解壓
Public Function PUB_UnZipFile(pZipFile As String, pToPath As String, Optional pPWD As String) As Boolean
   Dim program_name As String, program_path As String
   Dim process_id As Long
   Dim process_handle As Long

   program_name = "C:\Program Files\7-Zip\7z.exe"
   '檢查執行檔
   If Dir(program_name) = "" Then
      MsgBox "未安裝 7-Zip 程式，解壓縮檔產生失敗！。"
      Exit Function
   End If
    
On Error GoTo ShellError
              
   process_id = Shell("""" & program_name & """ x """ & pZipFile & """ -aoa -o""" & pToPath & """" & IIf(pPWD <> "", " -p" & pPWD, ""), vbNormalNoFocus)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    PUB_UnZipFile = True
    Exit Function

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Function

'Added by Morgan 2020/7/9
Public Function PUB_ZipFile(pFromPath As String, pZipName As String, Optional pPWD As String) As Boolean
   Dim program_name As String, program_path As String
   Dim process_id As Long
   Dim process_handle As Long

   program_name = "C:\Program Files\7-Zip\7z.exe"
   '檢查執行檔
   If Dir(program_name) = "" Then
      MsgBox "未安裝 7-Zip 程式，壓縮檔產生失敗！。"
      Exit Function
   End If
    
On Error GoTo ShellError
        
   '刪除舊檔
   If Dir(pZipName) <> "" Then
      Kill pZipName
   End If
   
   process_id = Shell("""" & program_name & """ a" & IIf(pPWD <> "", " -p" & pPWD, "") & " """ & pZipName & """ """ & pFromPath & """", vbHide)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    PUB_ZipFile = True
    Exit Function

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Function

Private Sub ShowTime()
   StatusBar1.Panels.Item(2).Text = Now
End Sub
Private Sub Command37_Click()
   Const stTempFolder As String = "C:\ZipConvert"
   Const stZipPwd As String = "HPCOMPANY"
   
   Dim bolRlt As Boolean
   Dim stDir As String, jj As Integer, kk As Integer, iMax2 As Integer
   Dim oTime As Variant
   Dim lngSize As Long, strMsg As String, stTxtFile As String
   Dim oCheck As CheckBox
   Dim stOrgFile As String, stOrgFilePath As String, stZipFile As String, stZipFilePath As String
   Dim fs, f, s
   
On Error GoTo ErrHnd
   
   strMsg = "讀取待壓縮資料..."
   lstHistory.AddItem Now & "--> " & strMsg, 0
   Write2File lstHistory.List(0), stTxtFile
   StatusBar1.Panels.Item(1).Text = strMsg
      
   '卷宗區
   If Option1(0).Value = True Then
      strExc(0) = "select cpp01 fNo,cpp02 fName, cpp14 fPath" & _
         " from patent,caseprogress,casepaperpdf c" & _
         " where pa26='X74164000'" & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
         " and cpp01(+)=cp09  and cpp10<>'D' and (lower(substr(cpp02,-4))<>'.zip' or instr(lower(cpp02),'.encrypt.zip')=0)"
         
   '原始檔
   Else
      strExc(0) = "select cpf01 fNo, cpf02 fName, cpf13 fPath" & _
         " from patent,caseprogress,casepaperfile c" & _
         " where pa26='X74164000'" & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
         " and cpf01(+)=cp09  and cpf10<>'D' and (lower(substr(cpf02,-4))<>'.zip' or instr(lower(cpf02),'.encrypt.zip')=0)"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   
      If Dir(stTempFolder, vbDirectory) = "" Then MkDir stTempFolder
      
      stTxtFile = stTempFolder & "\Log" & Format(Now, "yyyymmdd") & ".txt"
      
      m_bolStop = False
      
      oTime = time
      Label4.Caption = 0
         
      ProgressBar2.Value = 0
      Label7.Caption = 0
   
      Set fs = CreateObject("Scripting.FileSystemObject")
      
      With RsTemp
      kk = (.RecordCount \ 32000) + 1
      iMax2 = .RecordCount \ kk
      ProgressBar2.max = iMax2
      ProgressBar2.Value = 0
      Label7.Caption = .AbsolutePosition & "/" & .RecordCount
      Do While Not .EOF
         strMsg = .Fields("fName") & " ...下載中"
         lstHistory.AddItem Now & "--> " & strMsg, 0
         Write2File lstHistory.List(0), stTxtFile
         StatusBar1.Panels.Item(1).Text = strMsg
         DoEvents
   
         stOrgFilePath = stTempFolder & "\" & .Fields("fName")
         '卷宗區
         If Option1(0).Value = True Then
            'bolRlt = PUB_GetAttachFile_CPP(.Fields("fNo"), .Fields("fName"), stTempFolder)
            bolRlt = PUB_GetFtpFile(.Fields("fPath"), stOrgFilePath, "CASEPAPERPDF", True)
         '原始檔
         Else
            bolRlt = PUB_GetFtpFile(.Fields("fPath"), stOrgFilePath, "CASEPAPERFILE", True)
         End If
         If bolRlt = False Then
            strMsg = .Fields("fName") & " ...下載失敗"
            lstHistory.AddItem Now & "--> " & strMsg, 0
            Write2File lstHistory.List(0), stTxtFile
            StatusBar1.Panels.Item(1).Text = strMsg
            Exit Sub
         End If
         
         strMsg = .Fields("fName") & " ...已下載"
         lstHistory.AddItem Now & "--> " & strMsg, 0
         Write2File lstHistory.List(0), stTxtFile
         StatusBar1.Panels.Item(1).Text = strMsg
         DoEvents
            
         stZipFile = .Fields("fName") & ".encrypt.zip"
         stZipFilePath = stTempFolder & "\" & stZipFile
            
         If PUB_ZipFile(stOrgFilePath, stZipFilePath, stZipPwd) = False Then
            strMsg = .Fields("fName") & " ...壓縮失敗"
            lstHistory.AddItem Now & "--> " & strMsg, 0
            Write2File lstHistory.List(0), stTxtFile
            StatusBar1.Panels.Item(1).Text = strMsg
            Exit Sub
         End If
         
         strMsg = .Fields("fName") & " ...已壓縮"
         lstHistory.AddItem Now & "--> " & strMsg, 0
         Write2File lstHistory.List(0), stTxtFile
         StatusBar1.Panels.Item(1).Text = strMsg
         DoEvents
               
         Set f = fs.GetFile(stZipFilePath)
         '卷宗區
         If Option1(0).Value = True Then
            bolRlt = SaveAttFile_PDF(.Fields("fNo"), stZipFilePath, stZipFile, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True)
         '原始檔
         Else
            bolRlt = SaveAttFile_PDF(.Fields("fNo"), stZipFilePath, stZipFile, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True)
         End If
         If bolRlt = False Then
            strMsg = stZipFile & " ...上傳失敗"
            lstHistory.AddItem Now & "--> " & strMsg, 0
            Write2File lstHistory.List(0), stTxtFile
            StatusBar1.Panels.Item(1).Text = strMsg
            Exit Sub
         End If
         
         strMsg = stZipFile & " ...已上傳"
         lstHistory.AddItem Now & "--> " & strMsg, 0
         Write2File lstHistory.List(0), stTxtFile
         StatusBar1.Panels.Item(1).Text = strMsg
         
         '卷宗區
         If Option1(0).Value = True Then
            strSql = "update casepaperpdf set cpp02=cpp02||'.del',cpp10='D' where cpp01='" & .Fields("fNo") & "' and cpp02='" & ChgSQL(.Fields("fName")) & "'"
         '原始檔
         Else
            strSql = "update casepaperpdf set cpp02=cpp02||'.del',cpp10='D' where cpp01='" & .Fields("fNo") & "' and cpp02='" & ChgSQL(.Fields("fName")) & "'"
         End If
         
         cnnConnection.Execute strSql, intI
         If intI = 1 Then
            strMsg = .Fields("fName") & " ...已上刪除註記"
            lstHistory.AddItem Now & "--> " & strMsg, 0
            Write2File lstHistory.List(0), stTxtFile
            StatusBar1.Panels.Item(1).Text = strMsg
            DoEvents
         Else
            strMsg = .Fields("fName") & " ...刪除註記更新失敗"
            lstHistory.AddItem Now & "--> " & strMsg, 0
            Write2File lstHistory.List(0), stTxtFile
            StatusBar1.Panels.Item(1).Text = strMsg
            Exit Sub
         End If
            
         
         jj = jj + 1
         If jj >= kk Then
            ProgressBar2.Value = ProgressBar2.Value + 1
            jj = 0
         End If
         Label7.Caption = .AbsolutePosition & "/" & .RecordCount
         lngSize = DateDiff("s", oTime, time)
         Label4 = Format(lngSize \ 3600, "00") & ":" & Format((lngSize Mod 3600) \ 60, "00") & ":" & Format(lngSize Mod 60, "00")
         
         ShowTime
         DoEvents
         
         If m_bolStop = True Then
            If MsgBox("是否要繼續？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
               Exit Sub
            Else
               m_bolStop = False
            End If
         End If
               
         .MoveNext
      Loop
      End With
   End If
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   Write2File Err.Description, stTxtFile
   
End Sub

Private Sub Command38_Click()
   m_bolStop = True
End Sub

'Added by Lydia 2020/09/14 刪除D類收文(參考專利基本檔維護)
Private Sub Command39_Click()
Dim intJ As Integer
Dim rsRd As New ADODB.Recordset
Dim bolUpd As Boolean
Dim strContent As String

    strExc(0) = "Select Cp01,Cp02,Cp03,Cp04,Cp05,Cp09,Cp10 From Caseprogress,Patent " & _
                     "Where Cp01 In ('FCP','P') And Cp05>=20200330 And Cp31='Y' And Cp10 Not In (" & NewCasePtyList & ") " & _
                     "and cp12 like 'F%' and pa24||pa25 is not null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
    intI = 1
    Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
    
    If intI = 1 Then
On Error GoTo ErrHandle
        strContent = convForm("本所案號", 15) & convForm("收文號", 10) & "案件性質" & vbCrLf
        rsRd.MoveFirst
        Do While Not rsRd.EOF
            strExc(1) = "select cp01,cp02,cp03,cp04,cp09,cp10 from caseprogress where cp01='" & rsRd.Fields("cp01") & "' and cp02='" & rsRd.Fields("cp02") & "' and cp03='" & rsRd.Fields("cp03") & "' and cp04='" & rsRd.Fields("cp04") & "' " & _
                             "and substr(cp09,1,1)='A' and cp159=0 and cp10 in (" & NewCasePtyList & ") "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
            If intI = 0 Then
               strExc(2) = "select cp09,cp10,nvl(count(cpf02),0) cnt1,nvl(count(cpp02),0) cnt2 From caseprogress, casepaperfile,casepaperpdf " & _
                                "where cp01='" & rsRd.Fields("cp01") & "' and cp02='" & rsRd.Fields("cp02") & "' and cp03='" & rsRd.Fields("cp03") & "' and cp04='" & rsRd.Fields("cp04") & "' " & _
                                "and substr(cp09,1,1)='D' and cp10 in (" & cnt專利案件 & "," & cntEnglish_Vers & " ) and cp09=cpf01(+) and cp09=cpp01(+) group by cp09,cp10 "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(2))
               If intI = 1 Then
                    RsTemp.MoveFirst
                    Do While Not RsTemp.EOF
                         If Val("" & RsTemp.Fields("cnt1")) + Val("" & RsTemp.Fields("cnt2")) = 0 Then
                              If bolUpd = False Then
                                  cnnConnection.BeginTrans
                                  bolUpd = True
                              End If
                              strSql = "delete from caseprogress where cp09='" & RsTemp.Fields("cp09") & "' "
                              Pub_SeekTbLog strSql
                              cnnConnection.Execute strSql
                              strContent = strContent & convForm(rsRd.Fields("cp01") & "-" & rsRd.Fields("CP02") & IIf(rsRd.Fields("cp03") & rsRd.Fields("cp04") <> "000", "-" & rsRd.Fields("CP03") & "-" & rsRd.Fields("CP04"), ""), 15) & RsTemp.Fields("cp09") & "  " & RsTemp.Fields("cp10") & vbCrLf
                              intJ = intJ + 1
                         End If
                         RsTemp.MoveNext
                    Loop
               End If
            End If
            rsRd.MoveNext
        Loop
        If bolUpd = True Then cnnConnection.CommitTrans
    End If
    
    If intJ > 0 Then
      PUB_SendMail strUserNum, strUserNum, "", "整批", strContent
    End If
    
    MsgBox "OK!" & IIf(bolUpd = True, vbCrLf & "共" & intJ & "筆", ""), vbInformation
    
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
        If bolUpd = True Then cnnConnection.RollbackTrans
        MsgBox Err.Description
    End If
End Sub

'Added by Lydia 2020/10/26 客戶及代理人清單
Private Sub Command42_Click()
Dim rsAD As ADODB.Recordset
Dim strTemp(0 To 10) As String
Dim iCount As Long
Dim TempFileName As String
Dim strTitle As Variant, strTitleL As Variant, strTitleW As Variant

Dim stDate1 As String
Dim tmpArr As Variant

Dim xRows As Long '目前列位置
Dim xlsPoint As New Excel.Application
Dim wksA92 As New Worksheet
Dim inJ As Integer
Dim strPath As String
Dim strGrp As String
Dim intGrp As Integer

'On Error GoTo DebugErr

    TempFileName = "客戶及代理人清單"
    strPath = "C:\xls\" & TempFileName & ".xls"
    If Dir(strPath) <> "" Then
       Kill strPath
    End If
    

strTitle = Split("R/X/Y編號,建立日期,建立部門,國籍,英文名稱,中文名稱,R/X/Y狀態,原R編號,開發日期,開發人員,備註/客源/性質", ",")
strTitleL = Split("A,B,C,D,E,F,G,H,I,J,K", ",")
strTitleW = Split("10,9.5,9.5,9.5,28,28,10,9.5,9.5,9.5,13", ",")

   'stDate1 = CompWorkDay(2, strSrvDate(1), 1)
   stDate1 = "20200101"
   
'-----------------日期範圍
'   '客戶檔
'   strSql = "select '02' as ord1,'F' as atype, cu01||cu02 as cno,cu82 as cdate,na01,na03,cu05||' '||cu88||' '||cu89||' '||cu90 as ename," & _
'                "cu04 As cname, cu79 As cmemo, cu129 As ano, cu14 As adate, cu12 as cdept,csm02 as Remark, cu80 as ctype " & _
'                "from customer b1,nation,staff,casesourcemap where cu13=st01(+) and cu10=na01(+) and substr(cu12,1,1)='F' and cu02='0' and cu82>=" & stDate1 & " and cu82<" & strSrvDate(1) & _
'                " and cu09=csm01(+) "
'   '代理人檔
'   strSql = strSql & " Union All select '03' as ord1,'F' as atype, fa01||fa02 as cno,fa47 as cdate,na01,na03,fa05||' '||fa63||' '||fa64||' '||fa65 as ename," & _
'                "fa04 As cname, fa29 As cmemo, fa94 As ano, fa11 As adate, st03 as cdept, decode(FA76,'A','律師事務所','B','公司直接委辦','C','其他',FA76) as Remark, fa69 as ctype " & _
'                "from fagent b1,nation,staff where fa46=st01(+) and fa10=na01(+) and substr(st03,1,1)='F' and fa02='0' and fa47>=" & stDate1 & " and fa47<" & strSrvDate(1)
'   '潛在客戶檔
'   strSql = strSql & " Union All select '01' as ord1, 'F' as atype, pcu01||pcu02 as cno,pcu42 as cdate,na01,na03,pcu03||' '||pcu04||' '||pcu05||' '||pcu06 as ename," & _
'                "pcu08 as cname,pcu40 as cmemo,pcu38 as ano,pcu37 as adate, '' as cdept, '' as Remark, pcu39 as ctype " & _
'                "from potcustomer b1,nation where pcu09=na01(+) and pcu02='0' and pcu42>=" & stDate1 & " and pcu42<" & strSrvDate(1)
'   'Added by Lydia 2020/02/15 增加非國外部新增之X、Y、R編號
'   strSql = strSql & "Union All select '02' as ord1, 'S' as atype, cu01||cu02 as cno,cu82 as cdate,na01,na03,cu05||' '||cu88||' '||cu89||' '||cu90 as ename," & _
'                "cu04 As cname, cu79 As cmemo, cu129 As ano, cu14 As adate, a0902 as cdept,csm02 as Remark, cu80 as ctype " & _
'                "from customer b1,nation,staff,acc090,casesourcemap where cu13=st01(+) and cu10=na01(+) and substr(nvl(cu12,'N'),1,1)<>'F' and cu02='0' and cu82>=" & stDate1 & " and cu82<" & strSrvDate(1) & _
'                " and cu12=a0901(+) and cu09=csm01(+) "
'   '代理人檔
'   strSql = strSql & " Union All select '03' as ord1, 'S' as atype, fa01||fa02 as cno,fa47 as cdate,na01,na03,fa05||' '||fa63||' '||fa64||' '||fa65 as ename," & _
'                "fa04 As cname, fa29 As cmemo, fa94 As ano, fa11 As adate, a0902 as cdept, decode(FA76,'A','律師事務所','B','公司直接委辦','C','其他',FA76) as Remark, fa69 as ctype " & _
'                "from fagent b1,nation,staff,acc090 where fa46=st01(+) and fa10=na01(+) and substr(st03,1,1)<>'F' and fa02='0' and fa47>=" & stDate1 & " and fa47<" & strSrvDate(1) & _
'                " and st03=a0901(+) "
'   '國內潛在客戶檔
'   strSql = strSql & " Union All Select '01' as ord1, 'S' As Atype,Poc01||Poc02 As Cno,Poc18 As Cdate,Na01,Na03,Poc23||' '||Poc24||' '||Poc25||' '||Poc26||' '||Poc27 As Ename," & _
'                            " Nvl(Poc03,Poc28) As Cname,Poc15 As Cmemo,Poc13 As Ano,Poc12 As Adate,a0902 As Cdept, '' as Remark, poc14 as ctype " & _
'                            " From Potcustomer1 B1,Nation,Staff,Acc090 Where Poc04=Na01(+) And Poc02='0' And Poc18>=" & stDate1 & " and poc18<" & strSrvDate(1) & _
'                            " and poc13=st01(+) and st15=a0901(+) "
'-------------沒有日期範圍
   '客戶檔
   strSql = "select '02' as ord1,'F' as atype, cu01||cu02 as cno,cu82 as cdate,na01,na03,cu05||' '||cu88||' '||cu89||' '||cu90 as ename," & _
                "cu04 As cname, cu79 As cmemo, cu129 As ano, cu14 As adate, cu12 as cdept,csm02 as Remark, cu80 as ctype " & _
                "from customer b1,nation,staff,casesourcemap where cu13=st01(+) and cu10=na01(+) and substr(cu12,1,1)='F' and cu02='0' and cu09=csm01(+) "
   '代理人檔
   strSql = strSql & " Union All select '03' as ord1,'F' as atype, fa01||fa02 as cno,fa47 as cdate,na01,na03,fa05||' '||fa63||' '||fa64||' '||fa65 as ename," & _
                "fa04 As cname, fa29 As cmemo, fa94 As ano, fa11 As adate, st03 as cdept, decode(FA76,'A','律師事務所','B','公司直接委辦','C','其他',FA76) as Remark, fa69 as ctype " & _
                "from fagent b1,nation,staff where fa46=st01(+) and fa10=na01(+) and substr(st03,1,1)='F' and fa02='0' "
   '潛在客戶檔
   strSql = strSql & " Union All select '01' as ord1, 'F' as atype, pcu01||pcu02 as cno,pcu42 as cdate,na01,na03,pcu03||' '||pcu04||' '||pcu05||' '||pcu06 as ename," & _
                "pcu08 as cname,pcu40 as cmemo,pcu38 as ano,pcu37 as adate, '' as cdept, '' as Remark, pcu39 as ctype " & _
                "from potcustomer b1,nation where pcu09=na01(+) and pcu02='0' "
   'Added by Lydia 2020/02/15 增加非國外部新增之X、Y、R編號
   strSql = strSql & "Union All select '02' as ord1, 'S' as atype, cu01||cu02 as cno,cu82 as cdate,na01,na03,cu05||' '||cu88||' '||cu89||' '||cu90 as ename," & _
                "cu04 As cname, cu79 As cmemo, cu129 As ano, cu14 As adate, a0902 as cdept,csm02 as Remark, cu80 as ctype " & _
                "from customer b1,nation,staff,acc090,casesourcemap where cu13=st01(+) and cu10=na01(+) and substr(nvl(cu12,'N'),1,1)<>'F' and cu02='0'  and cu12=a0901(+) and cu09=csm01(+) "
   '代理人檔
   strSql = strSql & " Union All select '03' as ord1, 'S' as atype, fa01||fa02 as cno,fa47 as cdate,na01,na03,fa05||' '||fa63||' '||fa64||' '||fa65 as ename," & _
                "fa04 As cname, fa29 As cmemo, fa94 As ano, fa11 As adate, a0902 as cdept, decode(FA76,'A','律師事務所','B','公司直接委辦','C','其他',FA76) as Remark, fa69 as ctype " & _
                "from fagent b1,nation,staff,acc090 where fa46=st01(+) and fa10=na01(+) and substr(st03,1,1)<>'F' and fa02='0' and st03=a0901(+) "
   '國內潛在客戶檔
   strSql = strSql & " Union All Select '01' as ord1, 'S' As Atype,Poc01||Poc02 As Cno,Poc18 As Cdate,Na01,Na03,Poc23||' '||Poc24||' '||Poc25||' '||Poc26||' '||Poc27 As Ename," & _
                            " Nvl(Poc03,Poc28) As Cname,Poc15 As Cmemo,Poc13 As Ano,Poc12 As Adate,a0902 As Cdept, '' as Remark, poc14 as ctype " & _
                            " From Potcustomer1 B1,Nation,Staff,Acc090 Where Poc04=Na01(+) And Poc02='0' and poc13=st01(+) and st15=a0901(+) "
'---------------------------
   strSql = strSql & "  order by ord1,cno,na01"
   Set rsAD = New ADODB.Recordset
   rsAD.CursorLocation = adUseClient
   rsAD.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   With rsAD
       If .RecordCount > 0 And .RecordCount <> 0 Then
           .MoveFirst
           Do While Not .EOF
               '檢查是否為國外部人員開發
               If "" & .Fields("aType") = "F" And Right("" & .Fields("cno"), 1) = "R" Then
                    If "" & .Fields("ano") = "" Then GoTo JumpNextRec
                    strSql = "select count(*) cnt from staff where st01 in (" & GetAddStr("" & .Fields("ano")) & ") and substr(st03,1,1)='F' "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    If intI = 1 Then
                        If Val("" & RsTemp.Fields("cnt")) = 0 Then
                            GoTo JumpNextRec
                        End If
                    End If
               End If
               '申請人超過6萬筆,再分工作表; 因為2003版excel只能存65536筆
               If xRows > 60004 Or strGrp = "" Or (strGrp <> Left("" & .Fields("cno"), 1)) Then
                    If intGrp = 0 Then
                       xlsPoint.SheetsInNewWorkbook = 5 '預設工作表數目
                       xlsPoint.Workbooks.add
                    End If
                    intGrp = intGrp + 1
                    Set wksA92 = xlsPoint.Worksheets(intGrp)
                    wksA92.PageSetup.Orientation = xlLandscape '橫印
                    wksA92.Range("E1").Value = TempFileName
                    wksA92.Range("A2").Value = "列印日期："
                    wksA92.Range("B2").Value = ChangeTStringToTDateString(strSrvDate(2))
                   
                    '抬頭
                    xRows = 4
                    For inJ = 0 To UBound(strTitle)
                       wksA92.Range(Trim(strTitleL(inJ)) & xRows).Value = Trim(strTitle(inJ)) '欄位名稱
                       wksA92.Columns(Trim(strTitleL(inJ)) & ":" & Trim(strTitleL(inJ))).ColumnWidth = Val(strTitleW(inJ))     '欄寬
                       wksA92.Columns(Trim(strTitleL(inJ)) & ":" & Trim(strTitleL(inJ))).NumberFormatLocal = "@"   '欄位設定為文字型態
                    Next inJ
                    xRows = 5
                    strGrp = Left("" & .Fields("cno"), 1)
               End If
               strTemp(0) = "" & convForm("" & .Fields("cno"), 9)
               strTemp(1) = "" & convForm(ChangeTStringToTDateString(TransDate("" & .Fields("cdate"), 1)), 9)
               strTemp(2) = "" & convForm("" & .Fields("na03"), 8)
               strTemp(3) = "" & .Fields("ename")
               strTemp(4) = "" & .Fields("cname")
               If "" & .Fields("cmemo") <> "" And InStr("" & .Fields("cmemo"), ";原潛在客戶編號:") > 0 Then
                    strExc(1) = Mid("" & .Fields("cmemo"), InStr("" & .Fields("cmemo"), ";原潛在客戶編號:") + 1)
                    strExc(1) = Mid(strExc(1), Len(";原潛在客戶編號:"), 9)
                    strTemp(5) = "" & convForm(strExc(1), 9)
               Else
                    strTemp(5) = String(9, " ")
               End If

               strTemp(8) = "" & .Fields("cdept") 'Added by Lydia 2018/05/24 建立部門
               '備註/客源/性質; R編號/X編號/Y編號
               strTemp(9) = "" & .Fields("Remark")
               If Left(strTemp(0), 1) = "R" Then
                   strTemp(9) = "" & .Fields("cmemo")
               End If
               strTemp(10) = "" & .Fields("ctype")
               
               '開發日期
               'X08805的開發日期19110920
               strTemp(6) = TransDate("" & .Fields("adate"), 1)
               If Len(strTemp(6)) >= 6 Then
                   strTemp(6) = "" & convForm(ChangeTStringToTDateString(strTemp(6)), 9)
               End If
               '開發人員(複數)
               If "" & .Fields("ano") = "" Then
                    strTemp(7) = ""
               Else
                    strExc(2) = ""
                    strSql = "select st02,st03 from staff where st01 in (" & GetAddStr("" & .Fields("ano")) & ") order by st01 "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    If intI = 1 Then
                          RsTemp.MoveFirst
                          Do While Not RsTemp.EOF
                               strExc(2) = strExc(2) & RsTemp.Fields("st02") & ","
                               strTemp(8) = strTemp(8) & IIf(InStr(strTemp(8), "" & RsTemp.Fields("st03")) = 0, RsTemp.Fields("st03") & ",", "")  'Added by Lydia 2018/05/24 建立部門
                               RsTemp.MoveNext
                          Loop
                    End If
                    strTemp(7) = strExc(2)
               End If
               
               'Added by Lydia 2018/05/24 顯示"建立部門"
               strExc(2) = ""
               'Added by Lydia 2019/10/15 潛在客戶沒有建立部門,UBound(tmpArr)=-1造成For迴圈錯誤
               If strTemp(8) = "" Then
                   strTemp(8) = String(14, " ")
               'Added by Lydia 2020/02/15 非國外部：直接帶出部門
               ElseIf "" & .Fields("aType") <> "F" Then
                   strTemp(8) = convForm("" & .Fields("cdept"), 14)
               Else
               'end 2019/10/15
                    tmpArr = Split(strTemp(8), ",")
                    For intI = 0 To UBound(tmpArr)
                         If Trim(tmpArr(intI)) <> "" Then
                            Select Case Left(tmpArr(intI), 2)
                                 Case "F1": strExc(2) = strExc(2) & "外商,"
                                 Case "F2": strExc(2) = strExc(2) & "外專,"
                                 Case "F3": strExc(2) = strExc(2) & "法務,"
                                 Case "F4": strExc(2) = strExc(2) & "業拓,"
                                 Case Else: strExc(2) = strExc(2) & "其他,"
                            End Select
                         End If
                    Next
                    strTemp(8) = convForm(Mid(strExc(2), 1, Len(strExc(2)) - 1), 14)
               End If

               wksA92.Range("A" & xRows).Value = strTemp(0)
               wksA92.Range("B" & xRows).Value = strTemp(1)
               wksA92.Range("C" & xRows).Value = strTemp(8)
               wksA92.Range("D" & xRows).Value = strTemp(2)
               wksA92.Range("E" & xRows).Value = strTemp(3)
               wksA92.Range("F" & xRows).Value = strTemp(4)
               wksA92.Range("G" & xRows).Value = strTemp(10) '狀態
               wksA92.Range("H" & xRows).Value = strTemp(5)
               wksA92.Range("I" & xRows).Value = strTemp(6)
               wksA92.Range("J" & xRows).Value = strTemp(7)
               wksA92.Range("K" & xRows).Value = strTemp(9)
               xRows = xRows + 1
               
               iCount = iCount + 1
               
JumpNextRec:
               .MoveNext
           Loop
       End If
   End With

   If iCount > 0 Then
        '判斷若版本2007以上改變存檔格式
        If Val(xlsPoint.Version) < 12 Then
          xlsPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=-4143
        Else
          xlsPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=56
        End If
        xlsPoint.Workbooks.Close
        xlsPoint.Quit
        Set xlsPoint = Nothing
        Set wksA92 = Nothing
        MsgBox "OK !"
   End If
   Set rsAD = Nothing
   Exit Sub
   
DebugErr:
   If Err.Number <> 0 Then
      Set rsAD = Nothing
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2020/12/4
Private Sub Command43_Click()
Dim strCP09 As String

   'Add By Sindy 2022/4/25
   If Left(Trim(txtCaseNo), 2) = "&H" And Len(txtCaseNo) = 6 Then
      txtCaseNo = ChrW(Val(txtCaseNo))
      Exit Sub
   End If
   '2022/4/25 END
   
   strCP09 = InputBox("請輸入總收文號？")
   If Trim(strCP09) <> "" Then
      strSql = "select cp01,cp02,cp03,cp04 from caseprogress where CP09='" & strCP09 & "'"
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         txtCaseNo = adoRecordset.Fields("CP01") & adoRecordset.Fields("CP02") & _
                   adoRecordset.Fields("CP03") & adoRecordset.Fields("CP04")
      Else
         txtCaseNo = "查無資料"
      End If
   End If
   Exit Sub

'Add By Sindy 2022/3/2 [補掛期限]代理人通知核准,程序人員輸入核准通知時也要一併掛下一程序[註冊證]期限,管制證書催審!!
'**********************************************************************************
   Dim strNP08 As String, strNP09 As String, strNP22 As String
   strSql = "select c1.*,trademark.*" & _
            " from caseprogress c1,caseprogress c3,trademark where c1.cp01='T' and c1.cp10='1102' and c1.cp43 is not null and c1.cp43=c3.cp09 and c3.cp10 in('101','308') and c1.cp05<>19221111" & _
            " and tm01=c1.cp01 and tm02=c1.cp02 and tm03=c1.cp03 and tm04=c1.cp04 and tm10='020'" & _
            " and not exists (select * from nextprogress where np01=c1.cp09 and NP02=c1.cp01 and NP03=c1.cp02 and NP04=c1.cp03 and NP05=c1.cp04 AND NP07='1701')" & _
            " and not exists (select * from caseprogress c2 where c2.cp01=c1.cp01 and c2.cp02=c1.cp02 and c2.cp03=c1.cp03 and c2.cp04=c1.cp04 AND c2.cp10='1701')" & _
            " and not exists (select * from caseprogress c2 where c2.cp01=c1.cp01 and c2.cp02=c1.cp02 and c2.cp03=c1.cp03 and c2.cp04=c1.cp04 AND c2.cp10='728')" & _
            " and tm29 is null and tm57 is null"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         '申請(101)及分割(308)核准函掛下一程序1701期限NP08=NP09=系統日+8個月,NP10=申請或分割之CP14
         strNP08 = CompDate(1, 8, adoRecordset.Fields("cp05"))
         strNP09 = strNP08

         strNP22 = GetNextProgressNo()
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & adoRecordset.Fields("cp09") & "','" & adoRecordset.Fields("tm01") & "','" & adoRecordset.Fields("tm02") & "','" & adoRecordset.Fields("tm03") & "','" & adoRecordset.Fields("tm04") & "', 1701 ," & _
                          strNP08 & "," & strNP09 & ",'" & adoRecordset.Fields("cp14") & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
         
         adoRecordset.MoveNext
      Loop
   Else
      txtCaseNo = "查無資料"
   End If
   MsgBox "更新完畢!!", vbInformation
   Exit Sub

'Add By Sindy 2021/12/22 MCTF03人員留職停薪,寄發文件確認人員調整
'**********************************************************************************
Dim rsAD As New ADODB.Recordset
Dim strLP01 As String, Str01 As String, Str02 As String, Str03 As String, Str04 As String
Dim strCP13 As String

   strSql = "select * from letterprogress where lp06='MCTF03' and lp07=0"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         strLP01 = adoRecordset.Fields("lp01")
         
         strExc(1) = "select cp01,cp02,cp03,cp04" & _
                     " from caseprogress where cp09='" & strLP01 & "'"
         If rsAD.State = adStateOpen Then
            rsAD.Close
         End If
         rsAD.CursorLocation = adUseClient
         rsAD.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
         If rsAD.RecordCount > 0 Then
            Str01 = rsAD.Fields("cp01")
            Str02 = rsAD.Fields("cp02")
            Str03 = rsAD.Fields("cp03")
            Str04 = rsAD.Fields("cp04")
            '目前智權人員
            If Str01 = "FCP" Or Str01 = "FG" Then
               strCP13 = PUB_GetFCPSalesNo(Str01, Str02, Str03, Str04)
            ElseIf Str01 = "FCL" Or Str01 = "LIN" Then
               strCP13 = PUB_GetFCLSalesNo(Str01, Str02, Str03, Str04)
            ElseIf Str01 = "FCT" Then
               strCP13 = PUB_GetFCTSalesNo(Str01, Str02, Str03, Str04)
            ElseIf Str01 = "S" Then
              If adoRecordset.Fields("SP09") = "000" Then
                 strCP13 = PUB_GetFCTSalesNo(Str01, Str02, Str03, Str04)
              Else
                 strCP13 = PUB_GetAKindSalesNo(Str01, Str02, Str03, Str04)
              End If
            Else
               strCP13 = PUB_GetAKindSalesNo(Str01, Str02, Str03, Str04)
            End If
            If strCP13 <> "MCTF07" And strCP13 <> "MCTF06" And strCP13 <> "MCTF04" And strCP13 <> "MCTF01" Then
               MsgBox strCP13
            End If
            strSql = "update letterprogress set LP06='" & strCP13 & "',LP12=LP12||'110/12/22整批更新確認人員MCTF03改為" & strCP13 & "'" & _
                     " where LP01='" & strLP01 & "'"
            cnnConnection.Execute strSql
         End If
         
         adoRecordset.MoveNext
      Loop
   Else
      txtCaseNo = "查無資料"
   End If
   MsgBox "更新完畢!!", vbInformation
   Exit Sub
   
   
'**********************************************************************************
   '2. 由於需增加寄發附件Excel所示的Email【請協助以網域篩選，排除系統中的Y編號】
   '比對Y編號的 E-mail 網域
'Dim rsAD As New ADODB.Recordset
Dim strRR02 As String, strRR04 As String, strChkDa As String

   strSql = "select R02,R04 from rtemp where substr(r02,1,5)='97038'"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         strRR02 = adoRecordset.Fields("R02")
         strRR04 = adoRecordset.Fields("R04")
         strChkDa = Mid(strRR04, InStr(strRR04, "@"))
         
         strExc(1) = "select fa01,fa02,fa16,fa80,fa81,fa82" & _
                     " from fagent where (instr(fa16,'" & strChkDa & "')>0" & _
                     " or instr(fa80,'" & strChkDa & "')>0" & _
                     " or instr(fa81,'" & strChkDa & "')>0" & _
                     " or instr(fa82,'" & strChkDa & "')>0) and fa02='0'"
         If rsAD.State = adStateOpen Then
            rsAD.Close
         End If
         rsAD.CursorLocation = adUseClient
         rsAD.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
         If rsAD.RecordCount > 0 Then
'            If rsAD.RecordCount > 1 Then
'               MsgBox rsAD.RecordCount
'            End If
            strSql = "update rtemp set R01='" & rsAD.Fields("fa01") & rsAD.Fields("fa02") & "'" & _
                     " where R02='" & strRR02 & "'" & _
                     " and R04='" & strRR04 & "'"
            cnnConnection.Execute strSql
         End If
         
         adoRecordset.MoveNext
      Loop
   Else
      txtCaseNo = "查無資料"
   End If
   Exit Sub
   
'   Combo2.Clear
'   Combo2.AddItem "90 天", 0
'   Combo2.ITEMDATA(0) = 90
'   Combo2.AddItem "75 天", 0
'   Combo2.ITEMDATA(0) = 75
'   Combo2.AddItem "60 天", 0
'   Combo2.ITEMDATA(0) = 60
'   Combo2.AddItem "45 天", 0
'   Combo2.ITEMDATA(0) = 45
'   Combo2.AddItem "30 天", 0
'   Combo2.ITEMDATA(0) = 30
'   Combo2.AddItem "3 週", 0
'   Combo2.ITEMDATA(0) = 21
'   Combo2.AddItem "2 週", 0
'   Combo2.ITEMDATA(0) = 14
'   Combo2.AddItem "1 週", 0
'   Combo2.ITEMDATA(0) = 7
'
'   Dim pa() As String, intWhere As Integer
'   Dim m_CP27 As String, m_CP122 As String, iCP10m As String, m_203CP48 As String
'   Dim strCP06 As String, strCP07 As String, strCP48 As String
'
'   'FCP主動修正未發文未取消收文的
'   strSql = "SELECT * FROM caseprogress,engineerprogress WHERE cp01='FCP' AND cp10='203'" & _
'            " and cp27||cp57 is null and cp09=ep02(+) AND (instr(cp64,'命名-提申')=0 or cp64 is null)" & _
'            " order by cp09 desc"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'
'         Erase pa()
'         ReDim pa(1 To TF_PA) As String
'
'         pa(1) = Trim(adoRecordset.Fields("cp01"))
'         pa(2) = Trim(adoRecordset.Fields("cp02"))
'         pa(3) = Trim(adoRecordset.Fields("cp03"))
'         pa(4) = Trim(adoRecordset.Fields("cp04"))
'         If ClsPDReadPatentDatabase(pa(), intWhere) = False Then
'            MsgBox "QQQ"
'         End If
'
'         If pa(58) <> "" Or pa(108) <> "" Then
'            MsgBox "已閉卷或銷卷"
'            GoTo ReadNext
'         ElseIf InStr("" & adoRecordset.Fields("cp64"), "命名-提申") > 0 Then
'            MsgBox "" & adoRecordset.Fields("cp01") & "" & adoRecordset.Fields("cp02")
'            GoTo ReadNext
'         End If
'
'         strCP09 = "" & adoRecordset.Fields("cp09")
'         m_CP27 = "" & adoRecordset.Fields("cp27")
'         Text10(1).Text = "" & adoRecordset.Fields("cp10")
'         Text10(0).Text = "" & adoRecordset.Fields("cp14")
'         m_CP122 = "" & adoRecordset.Fields("cp122")
'         Text10(4).Text = "" & adoRecordset.Fields("cp06"): strCP06 = "" & adoRecordset.Fields("cp06")
'         Text10(5).Text = "" & adoRecordset.Fields("cp07"): strCP07 = "" & adoRecordset.Fields("cp07")
'         Text10(23).Text = "" & adoRecordset.Fields("cp48"): strCP48 = "" & adoRecordset.Fields("cp48")
'         iCP10m = "主動修正"
'         Text10(28).Text = "" & adoRecordset.Fields("ep06")
'         Text10(12).Text = "" & adoRecordset.Fields("cp05")
'
'         If PUB_CheckFCPshowMsg(Me.Visible, pa, m_CP27, Text10(1), Text10(0), m_CP122, _
'            Text10(4), Text10(5), Text10(23), iCP10m, Text10(28), Combo2, Text10(12), m_203CP48) Then
'
'            If strCP06 <> DBDATE(Text10(4)) Or _
'               strCP07 <> DBDATE(Text10(5)) Or _
'               strCP48 <> DBDATE(Text10(23)) Then
'
'               strSql = "update caseprogress set cp06=" & CNULL(DBDATE(Text10(4)), True) & _
'                        ",cp07=" & CNULL(DBDATE(Text10(5)), True) & _
'                        ",cp48=" & CNULL(DBDATE(Text10(23)), True) & " where cp09='" & strCP09 & "'"
'               cnnConnection.Execute strSql
'            End If
'         End If
'
'ReadNext:
'         adoRecordset.MoveNext
'      Loop
'   Else
'      MsgBox "查無資料", vbInformation
'   End If
'
'   Exit Sub
   
      
'   '補發文歸檔的承辦單
'   Dim rsAD As New ADODB.Recordset
'   Dim rsAD2 As New ADODB.Recordset
'   Dim intA As Integer
'   strExc(1) = "select eep01,cp01,cp02,cp10,eep03,eep06,eep07 From empelectronprocess,caseprogress,casepaperpdf" & _
'               " Where eep01 = CP09" & _
'               " AND substr(cp01,1,1)='T' AND substr(cp09,1,1)='C'" & _
'               " AND eep04='34'" & _
'               " AND eep01=cpp01 AND instr(upper(cpp02),'WORKSHEET')=0" & _
'               " group by eep01,cp01,cp02,cp10,eep03,eep06,eep07"
'   intA = 1
'   Set rsAD = ClsLawReadRstMsg(intA, strExc(1))
'   If intA = 1 Then
'      rsAD.MoveFirst
'      Do While Not rsAD.EOF
'         strExc(1) = "select cpp01 from casepaperpdf where cpp01='" & rsAD.Fields("eep01") & "' AND instr(upper(cpp02),'WORKSHEET')=0"
'         intA = 1
'         Set rsAD2 = ClsLawReadRstMsg(intA, strExc(1))
'         If intA = 0 Then
'            strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10,cpp12)" & _
'                     " values('" & rsAD.Fields("eep01") & "'," & _
'                          "'" & rsAD.Fields("cp01") & rsAD.Fields("cp02") & "." & rsAD.Fields("cp10") & ".WorkSheet.menu',0,'" & rsAD.Fields("eep03") & "'," & _
'                          rsAD.Fields("eep06") & "," & rsAD.Fields("eep07") & "," & _
'                          rsAD.Fields("eep06") & "," & rsAD.Fields("eep07") & ",'Y','S')"
'            cnnConnection.Execute strSql, intI
'         End If
'         rsAD.MoveNext
'      Loop
'   End If
End Sub

Private Sub Command46_Click()

   'Added by Lydia 2021/03/15 先判斷符合條件
   '--create table lydia_1100315a (sno varchar(3) ,cno varchar(12 char),cname varchar(180))
   
On Error GoTo ErrHandle

   cnnConnection.Execute "delete from lydia_1100315a"
   
   strExc(0) = "select * from lydia_1100315 where email is not null order by sno "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       RsTemp.MoveNext
       Do While Not RsTemp.EOF
           strExc(1) = UCase(Mid(Trim("" & RsTemp.Fields("email")), InStr(Trim("" & RsTemp.Fields("email")), "@")))
           '3/16 抓@後面比對,再查看重覆筆數高的,應該是免費信箱
           If InStr(UCase("@gmail.com;@aol.com;@hanmail.net;@naver.com;@akzonobel.com"), strExc(1)) = 0 Then
                '客戶
                strSql = "insert into lydia_1100315a (sno,cno,cname) select '" & Format(RsTemp.Fields("sno"), "000") & "' as sno, cu01||cu02 as cno,decode(cu05,null,cu04,cu05||' '||cu88||' '||cu89||' '||cu90) as cname " & _
                             "from customer where instr(upper(cu20||';'||cu115||';'||cu116||';'||cu117||';'||cu118),'" & strExc(1) & "') > 0 "
                cnnConnection.Execute strSql
                strSql = "insert into lydia_1100315a (sno,cno,cname) select '" & Format(RsTemp.Fields("sno"), "000") & "' as sno, cu01||cu02||'-'||pcc02 as cno,decode(cu05,null,cu04,cu05||' '||cu88||' '||cu89||' '||cu90) as cname " & _
                             "from customer,potcustcont where cu01=pcc01(+) and cu02='0' and instr(upper(pcc08),'" & strExc(1) & "') > 0 "
                cnnConnection.Execute strSql
                '代理人
                strSql = "insert into lydia_1100315a (sno,cno,cname) select '" & Format(RsTemp.Fields("sno"), "000") & "' as sno, fa01||fa02 as cno,decode(fa05,null,fa04,fa05||' '||fa63||' '||fa64||' '||fa65) as cname " & _
                             "from fagent where instr(upper(fa16||';'||fa79||';'||fa80||';'||fa81||';'||fa82||';'||fa105),'" & strExc(1) & "') > 0 "
                cnnConnection.Execute strSql
                strSql = "insert into lydia_1100315a (sno,cno,cname) select '" & Format(RsTemp.Fields("sno"), "000") & "' as sno, fa01||fa02||'-'||pcc02 as cno,decode(fa05,null,fa04,fa05||' '||fa63||' '||fa64||' '||fa65) as cname " & _
                             "from fagent,potcustcont where fa01=pcc01(+) and fa02='0' and instr(upper(pcc08),'" & strExc(1) & "') > 0 "
                cnnConnection.Execute strSql
                '潛在客戶
                strSql = "insert into lydia_1100315a (sno,cno,cname) select '" & Format(RsTemp.Fields("sno"), "000") & "' as sno, pcu01||pcu02 as cno,decode(pcu03,null,pcu08,pcu03||' '||pcu04||' '||pcu05||' '||pcu06) as cname " & _
                             "from potcustomer where instr(upper(pcu18),'" & strExc(1) & "') > 0 "
                cnnConnection.Execute strSql
                strSql = "insert into lydia_1100315a (sno,cno,cname) select '" & Format(RsTemp.Fields("sno"), "000") & "' as sno, pcu01||pcu02||'-'||pcc02 as cno,decode(pcu03,null,pcu08,pcu03||' '||pcu04||' '||pcu05||' '||pcu06) as cname " & _
                             "from potcustomer,potcustcont where pcu01=pcc01(+) and pcu02='0' and instr(upper(pcc08),'" & strExc(1) & "') > 0 "
                cnnConnection.Execute strSql
                '國內潛在客戶
                strSql = "insert into lydia_1100315a (sno,cno,cname) select '" & Format(RsTemp.Fields("sno"), "000") & "' as sno, poc01||poc02 as cno,decode(poc23,null,poc03,Poc23||' '||Poc24||' '||Poc25||' '||Poc26||' '||Poc27) as cname " & _
                             "from potcustomer1 where instr(upper(poc09),'" & strExc(1) & "') > 0 "
                cnnConnection.Execute strSql
                strSql = "insert into lydia_1100315a (sno,cno,cname) select '" & Format(RsTemp.Fields("sno"), "000") & "' as sno, poc01||poc02 as cno,decode(poc23,null,poc03,Poc23||' '||Poc24||' '||Poc25||' '||Poc26||' '||Poc27) as cname " & _
                             "from potcustomer1,potcustcont where poc01=pcc01(+) and poc02='0' and instr(upper(pcc08),'" & strExc(1) & "') > 0 "
                cnnConnection.Execute strSql
           End If
           RsTemp.MoveNext
       Loop
       strExc(1) = "select count(*) cnt from lydia_1100315a "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
       If intI = 1 Then
            If Val("" & RsTemp.Fields("cnt")) = 0 Then
                 MsgBox "查無資料!"
            Else
                 MsgBox "OK "
            End If
       End If
   End If
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
       MsgBox Err.Description
   End If
End Sub

''Mark by Lydia 2021/03/18 修正CFT英國脫歐案自動上傳檔案之進度備註
'Private Sub Command47_Click()
'
'On Error GoTo ErrHandle
'
'   strExc(0) = "Select Cp01,Cp02,Cp03,Cp04,cp09,Cp64,Nvl(Tm15,Tm12) Tm1512  From Caseprogress,Trademark " & _
'                    "Where Cp01='CFT' And Cp05>=20210101 And Cp10='1730' " & _
'                    "and cp01=tm01(+) and cp02=tm02 and cp03=tm03(+) and cp04=tm04(+) order by cp01,cp02 "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      cnnConnection.BeginTrans
'      Do While Not RsTemp.EOF
'           If InStr("" & RsTemp.Fields("tm1512"), "UK0") = 0 Then
'               'UK009+歐盟號數後8碼(拿掉第1碼0)
'               strExc(1) = "UK009" & Mid(RsTemp.Fields("tm1512"), 2)
'           Else
'               strExc(1) = RsTemp.Fields("tm1512")
'           End If
'           If InStr("," & RsTemp.Fields("cp64"), "英國脫歐案專利號數：") > 0 Then
'               strSql = "update caseprogress set cp64='" & Replace(RsTemp.Fields("cp64"), "英國脫歐案專利號數：", "英國脫歐案審定號數：") & "' where cp09='" & RsTemp.Fields("cp09") & "' "
'           Else
'               strSql = "update caseprogress set cp64=cp64||" & CNULL("英國脫歐案審定號數：" & strExc(1) & ";") & " where cp09='" & RsTemp.Fields("cp09") & "' "
'           End If
'           cnnConnection.Execute strSql
'           RsTemp.MoveNext
'      Loop
'      cnnConnection.CommitTrans
'   End If
'   MsgBox "OK !"
'   Exit Sub
'
'ErrHandle:
'   If Err.Number <> 0 Then
'        cnnConnection.RollbackTrans
'   End If
'End Sub

'Add By Sindy 2021/7/9 下載卷宗區電子檔
Private Sub Command47_Click()
Dim iRound As Integer
Dim m_TempDir(1 To 2) As String
Dim strTempFile As String
Dim strP1 As String, strP2 As String, strP3 As String

   strP3 = "": strP1 = "": strP2 = ""
   m_TempDir(1) = "C:"
   strSql = "SELECT cpp02,cpp01 FROM caseprogress,casepaperpdf WHERE cp01='FCT' AND cp10='301' AND cp27=20210708" & _
            " and cpp01(+)=cp09 and instr(cpp02,'.301.RECEIPT.pdf')>0"
   strSql = strSql & " order by cp67"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strTempFile = m_TempDir(1) & "\" & RsTemp.Fields("cpp02") '& ".pdf"
         strP1 = "" & RsTemp.Fields("cpp01")
         strP2 = "" & RsTemp.Fields("cpp02")
         If strP1 <> "" And strP2 <> "" And strTempFile <> "" Then
            If PUB_GetAttachFile_CPP(strP1, strP2, strTempFile, True) = False Then
             strP3 = strP3 & "," & strP1 & ":" & strP2 & "，無法下載"
            End If
         End If
         RsTemp.MoveNext
      Loop
   End If
   
   If strP3 <> "" Then Debug.Print Replace(Mid(strP3, 2), ",", vbCrLf)
   MsgBox "OK!"
End Sub

'Added by Lydia 2021/11/19 抓C類備註設定
Private Sub Command48_Click()
Dim strR1 As String, intR As Integer, tmpArr As Variant
Dim rsRd As New ADODB.Recordset
Dim dblAmt As Double, dblPFee As Double, dblTFee As Double


Exit Sub

'2025/08/15 將TrackingNo檔案搬到原始檔區; 流水號24467~24490(FCP-074213~FCP-074236)只做"變更",後續案件不進本所,在櫃台收文frm010005無法輸入流水號
'   Dim mSaveDir As String, intB As Integer
'   Dim bolMoveOK As Boolean
'   Dim rsBD As New ADODB.Recordset
'   Call Pub_ChkExcelPath(App.path & "\" & strUserNum)
'   mSaveDir = App.path & "\" & strUserNum & "\暫存區"
'   If Dir(mSaveDir, vbDirectory) = "" Then
'      MkDir mSaveDir
'   End If
'   strR1 = "select * from trackingcasename where tcn01>=24467 and tcn01<=24490 and tcn05 is null order by tcn01 "
'   intR = 1
'   Set rsRd = ClsLawReadRstMsg(intR, strR1)
'   strExc(0) = "74213"
'   If intR = 1 Then
'      rsRd.MoveFirst
'      Do While Not rsRd.EOF
'         bolMoveOK = False
'         strExc(1) = "select cp01,cp02,cp03,cp04,cp09,cp05 from caseprogress where cp01='FCP' and cp02='0" & strExc(0) & "' and cp03='0' and cp04='00' and cp31='Y' "
'         intB = 1
'         Set rsBD = ClsLawReadRstMsg(intB, strExc(1))
'         If intB = 1 Then
'            mSaveDir = App.path & "\" & strUserNum & "\暫存區"
'            Call PUB_UpdTCNfile(rsRd.Fields("tcn01"), rsBD.Fields("cp01") & rsBD.Fields("cp02") & rsBD.Fields("cp03") & rsBD.Fields("cp04"), rsBD.Fields("cp09"), rsBD.Fields("cp05"), mSaveDir, bolMoveOK)
'            If bolMoveOK = True Then
'                strSql = "Update trackingcasename set tcn05='" & rsBD.Fields("cp09") & "' where tcn01='" & rsRd.Fields("tcn01") & "' "
'                cnnConnection.Execute strSql
'            End If
'         End If
'         If bolMoveOK = True Then  'TrackingNO是否已搬檔完成(True無問題)，若有問題則TrackingNO和本機端的資料夾不刪除
'             Call PUB_KillAnyFile(mSaveDir)
'             RmDir mSaveDir  '移除資料夾
'         End If
'         strExc(0) = Val(strExc(0)) + 1
'         rsRd.MoveNext
'      Loop
'   End If
'   Set rsBD = Nothing
'   Set rsRd = Nothing
'   MsgBox "OK!"
'
'strExc(1) = PUB_GetReceiver("FCP", "041628", "0", "00", "605", "1")
'Debug.Print strExc(1)
'
'strExc(2) = PUB_GetNpMemo2("1", "FCP041628000", "605", strExc(1), "X65045000")
'Debug.Print strExc(2)
'
'Exit Sub
'
'  strSql = "select * from lydia_tmp order by np09 asc "
'  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'  intI = 1
'  If intI = 1 Then
'     RsTemp.MoveFirst
'     Do While Not RsTemp.EOF
'        strExc(1) = PUB_GetNpMemo2("1", RsTemp.Fields("np02") & RsTemp.Fields("np03") & RsTemp.Fields("np04") & RsTemp.Fields("np05"), RsTemp.Fields("np07"), RsTemp.Fields("pa75"), RsTemp.Fields("pa26") & "," & RsTemp.Fields("pa27") & "," & RsTemp.Fields("pa28") & "," & RsTemp.Fields("pa29") & "," & RsTemp.Fields("pa30"))
'        If strExc(1) <> "" Then
'          strSql = "Update lydia_tmp Set nm02='" & ChgSQL(strExc(1)) & "' where np01='" & RsTemp.Fields("np01") & "' and np07='" & RsTemp.Fields("np07") & "' and np22='" & RsTemp.Fields("np22") & "' "
'          cnnConnection.Execute strSql
'        End If
'        RsTemp.MoveNext
'     Loop
'  End If
'
'Exit Sub

'先產生暫存檔
'strSql = "drop table lydia_001"
'cnnConnection.Execute strSql
'strSql = "drop table lydia_002"
'cnnConnection.Execute strSql
strSql = "create table lydia_001 as Select Pa01||Pa02||Pa03||Pa04 As Pa0104,Pa75 as yno,Pa26 as xno From Patent Where Pa01 In ('FCP','P') And Pa09  In ('020','013','044','000') And Pa75 Is Not Null And Pa16||Pa108||Pa57 Is Null "
cnnConnection.Execute strSql
strSql = "create table lydia_002 (na01 VARCHAR2(4 CHAR), 代理人編號 VARCHAR2(9 CHAR), 代理人名稱 VARCHAR2(80 CHAR), 流水號 NUMBER(3), 本所案號 VARCHAR2(12 CHAR), Y編號 VARCHAR2(9 CHAR), Y編號名稱 VARCHAR2(180 CHAR), X編號 VARCHAR2(9 CHAR), X編號名稱 VARCHAR2(180 CHAR), 備註  VARCHAR2(2000 CHAR), 未准駁 NUMBER(3), 已閉  NUMBER(3) ) "
cnnConnection.Execute strSql
'strSql = "delete from lydia_002"
'cnnConnection.Execute strSql

' 分別對承辦和程序產生excel, 增加目前未准駁案件數、X/Y設定的名稱、去掉B類設定

'FCP案
strSql = "Select fa10 As na01,Pa75 As 代理人編號,Nvl(Fa04,Nvl(Fa05,Fa06)) As 代理人名稱,Im01 as 流水號,im03 as 本所案號,im04 as Y編號,im05 as X編號,Im02 as 備註,decode(Pa16||Pa108||Pa57,null,1,0) 未准駁, decode(pa57||pa108,null,0,1) 已閉卷 " & _
            "From Incommemo, Casepropertymap, Patent, Fagent Where Im06=Cpm02(+) And 'FCP'=Cpm01(+) And Substr(Im03,1,3)='FCP' " & _
            "And im03=pa01||pa02||pa03||pa04 And Substr(Pa75,1,8)=Fa01(+) And Substr(Pa75,9,1)=Fa02(+) "
'P案
strSql = strSql & "Union All Select fa10 As na01,Pa75 As 代理人編號,Nvl(Fa04,Nvl(Fa05,Fa06)) As 代理人名稱,Im01 As 流水號,Im03 As 本所案號,Im04 As Y編號,Im05 As X編號,Im02 As 備註,0 as 未准駁, 0 as 已閉卷 " & _
           "From Incommemo, Casepropertymap, Patent, Fagent Where Im06=Cpm02(+) And 'P'=Cpm01(+) And Substr(Im03,1,1)='P' " & _
           "And im03=pa01||pa02||pa03||pa04 And Substr(Pa75,1,8)=Fa01(+) And Substr(Pa75,9,1)=Fa02(+) "
'有Y編號
strSql = strSql & "Union All Select fa10 As na01,fa01||fa02 As 代理人編號,Nvl(Fa04,Nvl(Fa05,Fa06)) As 代理人名稱,Im01 as 流水號,im03 as 本所案號,im04 as Y編號,im05 as X編號,Im02 as 備註,0 as 未准駁, 0 as 已閉卷 " & _
           "From Incommemo, Casepropertymap, Fagent Where Im06=Cpm02(+) And 'FCP'=Cpm01(+) And Im03 Is Null " & _
           "and im04 is not null and substr(im04||'00',1,8)=fa01(+) and '0'=fa02(+) "
'有X編號
strSql = strSql & "Union All Select cu10 As Na01,cu01||cu02 As 代理人編號,Nvl(cu04,Nvl(cu05,cu06)) As 代理人名稱,Im01 As 流水號,Im03 As 本所案號,Im04 As Y編號,Im05 As X編號,Im02 As 備註,0 as 未准駁, 0 as 已閉卷 " & _
           "From Incommemo, Casepropertymap, customer Where Im06=Cpm02(+) And 'FCP'=Cpm01(+) And Im03||Im04 Is Null " & _
           "And im05 is not null and substr(Im05||'00',1,8)=Cu01(+) And '0'=Cu02(+) "
strSql = strSql & "order by 流水號"

intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strSql)
If intI = 1 Then
    RsTemp.MoveFirst
    Do While Not RsTemp.EOF
        strExc(1) = "": strExc(2) = ""   'X/Y編號的名稱
        strExc(3) = "0": strExc(4) = "0" '未准駁,已閉卷的件數
        strR1 = ""
        If "" & RsTemp.Fields("本所案號") <> "" Then
            strExc(3) = Val("" & RsTemp.Fields("未准駁"))
            strExc(4) = Val("" & RsTemp.Fields("已閉卷"))
        Else  '非個案
            If "" & RsTemp.Fields("Y編號") <> "" Then
               strR1 = strR1 & " and instr(yno,'" & RsTemp.Fields("Y編號") & "') > 0 "
            End If
            If "" & RsTemp.Fields("X編號") <> "" Then
               strR1 = strR1 & " and instr(xno,'" & RsTemp.Fields("X編號") & "') > 0 "
            End If
            strR1 = "select count(*) cnt from lydia_001 where " & Mid(strR1, 5)
            intR = 1
            Set rsRd = ClsLawReadRstMsg(intR, strR1)
            If intR = 1 Then
               strExc(3) = Val("" & rsRd.Fields("cnt"))
            End If
        End If
        If "" & RsTemp.Fields("Y編號") <> "" Then 'Y編號名稱
           If ClsPDGetAgent("" & RsTemp.Fields("Y編號"), strExc(1)) = True Then
               strExc(1) = Mid(strExc(1), 1, 80)
           End If
        End If
        If "" & RsTemp.Fields("X編號") <> "" Then 'X編號名稱
           If ClsPDGetCustomer("" & RsTemp.Fields("X編號"), strExc(2)) = True Then
               strExc(2) = Mid(strExc(2), 1, 80)
           Else
               strExc(2) = ""
           End If
        End If
        strExc(9) = "insert into lydia_002 (na01,代理人編號,代理人名稱,流水號,本所案號,Y編號,Y編號名稱,X編號,X編號名稱,備註,未准駁,已閉) " & _
                          "values('" & "" & RsTemp.Fields("na01") & "','" & "" & RsTemp.Fields("代理人編號") & "','" & ChgSQL("" & RsTemp.Fields("代理人名稱")) & "','" & RsTemp.Fields("流水號") & "','" & RsTemp.Fields("本所案號") & "' " & _
                          ",'" & RsTemp.Fields("Y編號") & "','" & ChgSQL(strExc(1)) & "','" & RsTemp.Fields("X編號") & "','" & ChgSQL(strExc(2)) & "','" & ChgSQL(RsTemp.Fields("備註")) & "','" & strExc(3) & "', '" & strExc(4) & "')"
        cnnConnection.Execute strExc(9)
        RsTemp.MoveNext
    Loop
End If

MsgBox "OK!"
End Sub

Private Sub Command49_Click()
Dim intPG As Integer
Dim xlsAccAddress As New Excel.Application
Dim wksAddress As New Worksheet
Dim tmpArr
Dim tmpDetail
Dim intA As Integer, intP As Integer, intCounter As Integer, intL As Integer
Dim intLine As Integer
Dim strFileName As String
Dim strTempLine As String

   strFileName = "$$地址條" & MsgText(43)
   If Dir(strExcelPath & strFileName) <> "" Then
       Kill strExcelPath & strFileName
   End If
   
   intPG = PUB_GetPaperSize(2, , False)
   If intPG = 0 Then
       MsgBox "地址條紙張格式未設定!!!", vbCritical
   Else
       strExc(0) = "１０４臺北市中山區長安東路一段１８號５樓$華南商業銀行股份有限公司|" & _
                        "７１７臺南市仁德區保安里民生路１５-１號$聯豐生物科技股份有限公司|" & _
                        "高雄市鳳山區大仁街２８號$益鑫記帳士事務所~黃小姐|"
       tmpArr = Split(strExc(0), "|")
       For intA = 0 To UBound(tmpArr)
           If Trim(tmpArr(intA)) <> "" Then
               If intCounter = 0 Then
                    xlsAccAddress.SheetsInNewWorkbook = 1
                    xlsAccAddress.Workbooks.add
                    Set wksAddress = xlsAccAddress.Worksheets(1)
                    wksAddress.Activate
                    'xlsAccAddress.Visible = True
                    If Val(xlsAccAddress.Version) < 12 Then
                        xlsAccAddress.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
                    Else
                        xlsAccAddress.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
                    End If
                    wksAddress.Range("A:A").ColumnWidth = 50
                    wksAddress.Range("A:A").Font.Size = 12
                    wksAddress.PageSetup.PaperSize = intPG
                    wksAddress.PageSetup.Orientation = xlPortrait '直印
                    wksAddress.PageSetup.Zoom = 100 '縮放比例為100%,列印頁面水平置中
                    wksAddress.PageSetup.HeaderMargin = Application.InchesToPoints(0) '頁首
                    wksAddress.PageSetup.FooterMargin = Application.InchesToPoints(0) '頁尾
                    wksAddress.PageSetup.TopMargin = xlsAccAddress.InchesToPoints(0.1) '上
                    wksAddress.PageSetup.BottomMargin = xlsAccAddress.InchesToPoints(0.1) '下
                    wksAddress.PageSetup.LeftMargin = xlsAccAddress.InchesToPoints(0.1) '左邊界
                    wksAddress.PageSetup.RightMargin = xlsAccAddress.InchesToPoints(0.1) '右邊界
               End If
               intCounter = intCounter + 1
               tmpDetail = Empty
               tmpDetail = Split(tmpArr(intA), "$")
               intLine = 1
               For intP = 0 To UBound(tmpDetail)
                   Select Case intP
                        Case 0  '地址
                            strTempLine = GetEnterLineAcc("" & tmpDetail(intP), 38, "", intL)
                        Case 1  '收件人
                            strTempLine = GetEnterLineAcc("" & tmpDetail(intP) & MsgText(104), 38, "~", intL)
                   End Select
                   '地址條空白列範圍為A1~A7： 其中列印內容超過一行寬度用折行方式，在同一儲存格列印，所以地址固定在A1而收件者在A3
                   wksAddress.Range("A" & IIf(intP = 0, "1", "3")).Value = strTempLine
                   intLine = intLine + intL '記錄列高
                   If intP = 0 Then intLine = intLine + 1
               Next intP
               xlsAccAddress.Workbooks(1).Save
               wksAddress.PrintOut Copies:=1, Collate:=True
               '清除資料
               wksAddress.Range("A1:A7").ClearContents
           End If
       Next intA
       
       xlsAccAddress.Workbooks(1).Save
       xlsAccAddress.Quit
       Set wksAddress = Nothing
       Set xlsAccAddress = Nothing
       MsgBox "OK"
   End If
End Sub
'Added by Lydia 2022/03/01 取得折行後的文字
Private Function GetEnterLineAcc(ByVal pTmp As String, ByVal intM As Single, Optional ByVal pSingle As String, Optional ByRef intRe As Integer) As String
'intM: 每行最大字數,中文算2個
'intRe: 行數
'pSingle : 區隔符號
Dim intB As Integer, strTempB As String, intY As Integer, strTempA As String
Dim tmpArr1
    
    strTempA = IIf(pSingle <> "", "　　", "") & pTmp
    If pTmp = "" Or (pSingle = "" And GetTextLength(strTempA) <= intM) Or _
             (pSingle <> "" And InStr(strTempA, pSingle) = 0 And GetTextLength(strTempA) <= intM) Then
       GetEnterLineAcc = strTempA
       intRe = 1
    Else
       tmpArr1 = Split(pTmp, IIf(pSingle <> "", pSingle, "|"))
       For intY = 0 To UBound(tmpArr1)
            '收件人需縮排
            strTempA = IIf(pSingle <> "", "　　", "") & tmpArr1(intY)
            strTempB = PUB_StrToStr(strTempA, intM)
            Do While strTempB <> ""
                GetEnterLineAcc = GetEnterLineAcc & IIf(GetEnterLineAcc <> "", vbCrLf, "") & strTempB
                intB = intB + 1
                strTempA = Replace(strTempA, strTempB, "")
                If strTempA <> "" Then
                    strTempB = IIf(pSingle <> "", "　　", "") & PUB_StrToStr(strTempA, intM)
                Else
                    strTempB = ""
                End If
            Loop
       Next intY
       intRe = intB
    End If
    
End Function

'Add By Sindy 2012/8/27
'過濾研討會的Excel檔案是否有本所客戶
Private Sub Command5_Click()
Dim strFileName As String
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim bolReadExit As Boolean
Dim iRow As Integer
Dim strCustName As String, strCompanyName As String, strCUNo As String
   
   On Error GoTo flgErr
   
   If txtFileName = "" Then
      MsgBox "研討會檔案不可空白！"
      txtFileName.SetFocus
      Exit Sub
   End If
   
   strFileName = txtFileName 'PUB_Getdesktop & "\" & strYear & "年IPC分類案件市佔分析.xls"
   
'   If Dir(strFileName) <> MsgText(601) Then
'      Kill strFileName
'   End If
   
   '開檔
   Screen.MousePointer = vbHourglass
   xlsSalesPoint.Workbooks.Open strFileName
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
   
   '新增A欄
   wksaccrpt114.Columns("A:A").Select
   xlsSalesPoint.Selection.Insert Shift:=xlToRight
   wksaccrpt114.Range("A1").Select
   wksaccrpt114.Range("A1").Value = "本所客戶"
   wksaccrpt114.Columns("A:A").ColumnWidth = 14
   
   '過濾客戶資料
   bolReadExit = False: iRow = 2
   Do While bolReadExit = False
      If wksaccrpt114.Range("B" & iRow).Value = "" Then
         bolReadExit = True
      Else
'         strCustName = Trim(wksaccrpt114.Range("E" & iRow).Value)
'         strCustName = Replace(strCustName, " ", "")
'         strCustName = Replace(strCustName, "　", "")
         strCompanyName = Trim(wksaccrpt114.Range("E" & iRow).Value)
         If InStr(strCompanyName, "(股)") > 0 Then
            strCompanyName = Replace(strCompanyName, "(股)", "股份有限")
         End If
         strCompanyName = Replace(strCompanyName, " ", "")
         strCompanyName = Replace(strCompanyName, "　", "")
         strCompanyName = Left(strCompanyName, 4)
         
         '先過濾公司名稱
         strCUNo = ""
         strSql = "select cu01||cu02 from customer where substr(cu04,1,4)='" & strCompanyName & "' or substr(upper(replace(rtrim(rtrim(cu05)||rtrim(cu88)||rtrim(cu89)||rtrim(cu90)),' ',null)),1,4)='" & UCase(strCompanyName) & "' or substr(cu06,1,4)='" & strCompanyName & "'"
         If adoRecordset.State = adStateOpen Then
            adoRecordset.Close
         End If
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount > 0 Then
            adoRecordset.MoveFirst
            Do While Not adoRecordset.EOF
               strCUNo = strCUNo & "，" & adoRecordset.Fields(0)
               adoRecordset.MoveNext
            Loop
         End If
'         If strCUNo = "" Then
'            '再過濾個人姓名
'            strSql = "select cu01||cu02 from customer where cu04='" & strCustName & "' or upper(replace(rtrim(rtrim(cu05)||rtrim(cu88)||rtrim(cu89)||rtrim(cu90)),' ',null))='" & UCase(strCustName) & "' or cu06='" & strCustName & "'"
'            If adoRecordset.State = adStateOpen Then
'               adoRecordset.Close
'            End If
'            adoRecordset.CursorLocation = adUseClient
'            adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'            If adoRecordset.RecordCount > 0 Then
'               adoRecordset.MoveFirst
'               Do While Not adoRecordset.EOF
'                  strCUNo = strCUNo & "，" & adoRecordset.Fields(0)
'                  adoRecordset.MoveNext
'               Loop
'            End If
'         End If
         '寫回Excel檔
         If strCUNo <> "" Then
            strCUNo = Right(strCUNo, Len(strCUNo) - 1)
            wksaccrpt114.Range("A" & iRow).Value = strCUNo
         End If
         
         iRow = iRow + 1
      End If
   Loop
   adoRecordset.Close
   '存檔
   xlsSalesPoint.Workbooks(1).SaveAs FileName:=Left(strFileName, Len(strFileName) - 4) & "_new.xls"
   
   '關閉
   xlsSalesPoint.Workbooks.Close
   '離開
   xlsSalesPoint.Quit
   Set wksaccrpt114 = Nothing
   Set xlsSalesPoint = Nothing
   Screen.MousePointer = vbDefault
   
   MsgBox "資料過濾完畢！"
   
   Exit Sub
   
flgErr:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2022/5/17
Private Sub Command51_Click()
  
   '更新卷宗區電子檔名案號流水號足6碼
   Dim rsAD As New ADODB.Recordset
   Dim rsAD2 As New ADODB.Recordset
   Dim strOldCPP01 As String, strOldCPP02 As String, strOldCPP13 As String
   Dim strCP01 As String, strCP02 As String, strNewCPP02 As String
   Dim intA As Integer
   
On Error GoTo flgErr
   
   '要使用的話要改一下,保留原修改人員?
   Exit Sub
   
   If MsgBox("確定要更新卷宗區電子檔名案號流水號足6碼？", vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   strExc(1) = "select *" & _
               " from casepaperpdf,caseprogress where cpp01=cp09(+)" & _
               " and cp09 is not null" & _
               " and ((instr(cpp02,cp01||cp02||'.')=0 and instr(cpp02,cp01||TO_NUMBER(cp02)||'.')>0)" & _
                 " or (instr(cpp02,cp01||cp02||'-')=0 and instr(cpp02,cp01||TO_NUMBER(cp02)||'-')>0))" & _
               " and cpp10 in('Y','P','F','T','X') and cpp06>=20210101" '& _
'               " and cp01='FCT' and cp02='042955'"
'--and (cp03<>'0' or cp04<>'00')
'--and cp04<>'00'
'--and rownum<=50
   intA = 1
   Dim ii_d As Double, ii_err As Double
   Set rsAD = ClsLawReadRstMsg(intA, strExc(1))
   If intA = 1 Then
      rsAD.MoveFirst
      Do While Not rsAD.EOF
         ii_d = ii_d + 1
         'If Val("" & rsAD.Fields("cpp13")) > 0 Then '已有FTP更新日期
            strOldCPP01 = rsAD.Fields("cpp01")
            strOldCPP02 = rsAD.Fields("cpp02")
            strOldCPP13 = "" & rsAD.Fields("cpp13")
            strCP01 = rsAD.Fields("cp01")
            strCP02 = rsAD.Fields("cp02")
            
            Me.Caption = "維護作業1  ( " & ii_d & " / Err:" & ii_err & " / " & rsAD.RecordCount & " )"
            DoEvents
            
            If InStr(strOldCPP02, strCP01 & Val(strCP02) & ".") > 0 Then
               strNewCPP02 = Replace(strOldCPP02, strCP01 & Val(strCP02) & ".", strCP01 & strCP02 & ".", , 1)
            Else
               strNewCPP02 = Replace(strOldCPP02, strCP01 & Val(strCP02) & "-", strCP01 & strCP02 & "-", , 1)
            End If
            strExc(1) = "select * from casepaperpdf where cpp01='" & strOldCPP01 & "' AND cpp02='" & strNewCPP02 & "'"
            intA = 1
            Set rsAD2 = ClsLawReadRstMsg(intA, strExc(1))
            If intA = 0 Then
               strSql = "update casepaperpdf set cpp02='" & strNewCPP02 & "' where cpp01='" & strOldCPP01 & "' AND cpp02='" & strOldCPP02 & "'"
               cnnConnection.Execute strSql, intI
               If Val(strOldCPP13) > 0 Then '已有FTP更新日期
                  strSql = "update casepaperpdf set cpp13=" & strSrvDate(1) & " where cpp01='" & strOldCPP01 & "' AND cpp02='" & strNewCPP02 & "'"
                  cnnConnection.Execute strSql, intI
               End If
            Else
               ii_err = ii_err + 1
            End If
         'End If
         
ReadNext:
         rsAD.MoveNext
      Loop
   End If
   MsgBox "完成!!!"
   
   Set rsAD = Nothing
   Set rsAD2 = Nothing
   
flgErr:
   If Err.Number <> 0 Then
      ii_err = ii_err + 1
      GoTo ReadNext
   End If
End Sub

'Add By Sindy 2022/7/28 比對電子檔並刪除(信件)
Private Sub Command52_Click()
Dim oFileSys As New FileSystemObject
Dim oFolder As Folder
Dim oFile As Object
Dim objOutLook As Object
Dim objMail As Object
Dim strMeCap As String
Dim lngRonCnt As Long
Dim strII17 As String, strII11 As String, strII12 As String, strII13 As String
Dim fs
Dim strFolder As String

On Error GoTo ErrHand
   
   strExc(6) = UCase(InputBox("請輸入要檢查的信箱資料檔，是那一個？", "信箱資料檔為(F/P/T)"))
   If Trim(strExc(6)) = "" Then
      Exit Sub
   Else
      If Trim(strExc(6)) <> "F" And Trim(strExc(6)) <> "P" And Trim(strExc(6)) <> "T" Then
         Exit Sub
      End If
   End If
                  
   strMeCap = Me.Caption
   
   If Trim(strExc(6)) = "F" Then
      strFolder = PUB_Getdesktop & "\IPDept"
   ElseIf Trim(strExc(6)) = "P" Then
      strFolder = PUB_Getdesktop & "\Patent"
   Else
      strFolder = PUB_Getdesktop & "\TM"
   End If
   
   Set oFolder = oFileSys.GetFolder(strFolder)
   Set objOutLook = CreateObject("Outlook.Application")
   Set fs = CreateObject("Scripting.FileSystemObject")
   lngRonCnt = 0
   For Each oFile In oFolder.files
      lngRonCnt = lngRonCnt + 1
      Me.Caption = strMeCap & " 已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count
      DoEvents
'      oForm.TxtIPDept = oFile.Name
      
      If UCase(Right(Trim(oFile.Name), 4)) = UCase(".msg") Then
'         Call PUB_ExLetterTransTxt(oFile, CStr(strFolder))
         
         Set objMail = objOutLook.CreateItemFromTemplate(strFolder & "\" & oFile.Name)
         DoEvents
         Screen.MousePointer = vbHourglass
         
         strII17 = ChgSQL(objMail.Subject)
'         oForm.TextII17 = objMail.Subject 'Add By Sindy 2021/4/12 Find簡體字
         
         DoEvents
         If objMail.Class = 46 Then '46.olReport
            strII11 = "未傳遞的主旨"
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
         
         '檢查是否有此筆郵件已匯入系統
         If Trim(strExc(6)) = "F" Then
            strSql = "select ii01,ii03 from IPDeptinput" & _
                     " where instr(ii17,'" & strII17 & "')>0" & _
                     " and ii11 = '" & ChgSQL(strII11) & "' and ii12 = " & strII12 & " and ii13 = " & strII13 & _
                     " order by ii01 desc,ii03 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               Call fs.DeleteFile(strFolder & "\" & oFile.Name)
            End If
         ElseIf Trim(strExc(6)) = "P" Then
            strSql = "select pi01,pi03 from patentinput" & _
                     " where pi17 = '" & ChgSQL(strII17) & "'" & _
                     " and pi11 = '" & ChgSQL(strII11) & "' and pi12 = " & strII12 & " and pi13 = " & strII13 & _
                     " order by pi01 desc,pi03 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               Call fs.DeleteFile(strFolder & "\" & oFile.Name)
            End If
         Else
            strSql = "select ti01,ti03 from TMinput" & _
                     " where ti17 = '" & ChgSQL(strII17) & "'" & _
                     " and ti11 = '" & ChgSQL(strII11) & "' and ti12 = " & strII12 & " and ti13 = " & strII13 & _
                     " order by ti01 desc,ti03 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               Call fs.DeleteFile(strFolder & "\" & oFile.Name)
            End If
         End If
      End If
   Next
   
ErrHand:
   Me.Caption = strMeCap
   Screen.MousePointer = vbDefault
   Set fs = Nothing
   Set oFolder = Nothing
   Set oFile = Nothing
   Set oFileSys = Nothing
   Set objMail = Nothing
   Set objOutLook = Nothing
End Sub

'Added by Morgan 2022/9/28
'批次下載卷宗區檔案
Private Sub Command53_Click()
Dim strToPath As String
Dim strTempFile As String
Dim strP1 As String, strP2 As String, strP3 As String

   strP3 = "": strP1 = "": strP2 = ""
   strToPath = "C:\Users\92012\Desktop\CPP"
   strSql = "select cpp02,cpp01,m00" & _
      " From morgan, patent, caseprogress, casepaperpdf" & _
      " where pa01(+)=substr(m01,1,3) and pa02(+)=substr(m01,5) and pa03(+)='0' and pa04(+)='00'" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='1912'" & _
      " and cpp01(+)=cp09 and instr(upper(cpp02),'.1912.PDF')>0" & _
      " order by m00"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strTempFile = strToPath & "\" & RsTemp.Fields("cpp02")
         strP1 = "" & RsTemp.Fields("cpp01")
         strP2 = "" & RsTemp.Fields("cpp02")
         If strP1 <> "" And strP2 <> "" And strTempFile <> "" Then
            If PUB_GetAttachFile_CPP(strP1, strP2, strTempFile, True) = False Then
             strP3 = strP3 & "," & strP1 & ":" & strP2 & "，無法下載"
            End If
         End If
         RsTemp.MoveNext
      Loop
   End If
   
   If strP3 <> "" Then Debug.Print Replace(Mid(strP3, 2), ",", vbCrLf)
   MsgBox "OK!"
End Sub

'更新FCT核准定稿
Private Sub Command55_Click()
Dim intQ As Integer

   strSql = "select ld01,ld02,ld03,ld04,ld05,ld06,ld07,ld08,ld09,ld10,ld11,ld16 from letterdemand where ld05='FCT' and ld02='99999999' order by 1,2,3 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      cnnConnection.BeginTrans
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            strExc(1) = PUB_GetUniqeLD03("" & RsTemp.Fields("ld01"), strSrvDate(1), "" & RsTemp.Fields("ld03"))
            strSql = "update letterdemand set ld02=" & strSrvDate(1) & ",ld03='" & strExc(1) & "' where ld04 ='" & RsTemp.Fields("ld04") & "' and ld02='" & RsTemp.Fields("ld02") & "' and ld09='" & RsTemp.Fields("ld09") & "' and ld10='" & RsTemp.Fields("ld10") & "'  and LD11='" & RsTemp.Fields("ld11") & "' and ld01='" & RsTemp.Fields("ld01") & "' and LD03='" & RsTemp.Fields("ld03") & "' "
            cnnConnection.Execute strSql, intQ
            strSql = "update exceptcondition set et07=" & strSrvDate(1) & " where et02 ='" & RsTemp.Fields("ld04") & "' and et07='" & RsTemp.Fields("ld02") & "'  and et01='" & RsTemp.Fields("ld10") & "' and et03='" & RsTemp.Fields("ld11") & "' and et04='" & RsTemp.Fields("ld01") & "' "
            cnnConnection.Execute strSql, intQ
            RsTemp.MoveNext
         Loop
      cnnConnection.CommitTrans
   End If
   MsgBox "OK !"
   
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Sub


Private Sub Command58_Click()
   PUB_SendMail strUserNum, txtRcvr.Text, "", txtSubj & "-" & Now & "(SMTP:" & Text8.Text & ")", "testing...", , , , , , txtCC, , , , , , , , , , , , , , , , , , , Text8.Text
   If bolMailSendOk Then
      MsgBox "已寄出！", vbOKOnly + vbInformation
   Else
      MsgBox "寄信失敗！", vbOKOnly + vbCritical
   End If
End Sub

'Add By Sindy 2012/8/27
Private Sub Command6_Click()
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.xls"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "Excel檔案 (*.xls)|*.xls"
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtFileName.Text = .FileName
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2012/9/11 補掛期限
Private Sub Command7_Click()
Dim iRow As Integer
Dim xlsSalesPoint As New Excel.Application
Dim wksrpt As New Worksheet
Dim StrPA11 As String, strPA160 As String, strTPB10 As String
Dim strTPB11 As String, strTPB12 As String, strData As String, strTPB13 As String
   
'On Error GoTo CheckingErr
   
   '開檔
    Screen.MousePointer = vbHourglass
    xlsSalesPoint.Workbooks.Open "C:\Users\97038\Desktop\申復再審案公報無分類A.xlsx"
    Set wksrpt = xlsSalesPoint.Worksheets(1)
    iRow = 2
    cnnConnection.BeginTrans
    Do While wksrpt.Range("B" & iRow).Value <> MsgText(601)
        StrPA11 = "": strPA160 = "": strTPB10 = ""
        strTPB11 = "": strTPB12 = "": strData = "": strTPB13 = ""
        
        StrPA11 = Trim(wksrpt.Range("B" & iRow).Value)
        If Trim(wksrpt.Range("C" & iRow).Value) = "" Or Trim(wksrpt.Range("C" & iRow).Value) = "-" Then
            If Trim(wksrpt.Range("E" & iRow).Value) <> "" Then
               strPA160 = Trim(wksrpt.Range("E" & iRow).Value)
               strTPB10 = strPA160 '國際分類號
               strTPB11 = "11"
            End If
        Else
            strData = Trim(wksrpt.Range("C" & iRow).Value)
            strPA160 = Left(strData, 4) '國際分類前4碼
            strTPB10 = strData '國際分類號
            strTPB11 = GetPatentIPC("1", strTPB10, "")
        End If
        
        If StrPA11 <> "" And strTPB10 <> "" Then
            '產業別分類
            strTPB12 = GetPatentIPC("2", strTPB10, "")
            '案件屬性
            strTPB13 = GetPatentIPC("3", strTPB10, "")
             
        If strTPB11 = "" And strTPB12 = "" And strTPB13 = "" Then
            MsgBox "strPA11=" & StrPA11 & ", strTPB10=" & strTPB10
        End If
        
            '更新:
            '專利公報
            strSql = "update tpbulletin set tpb10=" & CNULL(strTPB10) & ",tpb11=" & CNULL(strTPB11) & _
                     ",tpb12=" & CNULL(strTPB12) & ",tpb13=" & CNULL(strTPB13) & " where tpb01='" & StrPA11 & "' and tpb10 is null"
            cnnConnection.Execute strSql
            '專利公開公報
            strSql = "update tpgazette set tpg15=" & CNULL(strTPB10) & ",tpg16=" & CNULL(strTPB11) & _
                     ",tpg17=" & CNULL(strTPB12) & ",tpg18=" & CNULL(strTPB13) & " where tpg01='" & StrPA11 & "' and tpg15 is null"
            cnnConnection.Execute strSql
            '專利檔
            strSql = "update patent set PA160=" & CNULL(strPA160) & " where PA11='" & StrPA11 & "' and PA160 is null"
            cnnConnection.Execute strSql
        Else
            MsgBox "strPA11=" & StrPA11 & ", strTPB10=" & strTPB10
        End If
        iRow = iRow + 1
    Loop
    cnnConnection.CommitTrans
    
    '關閉
    xlsSalesPoint.Workbooks.Close
    '離開
    xlsSalesPoint.Quit
    Set wksrpt = Nothing
    Set xlsSalesPoint = Nothing

'   strSql = "select PZD02,PZD03,mz06,mz07 from mailzip," & _
'            "(select PZD02,PZD03 from postzipdata group by PZD02,PZD03) X" & _
'            " where mz04(+)=pzd02 and mz05(+)=pzd03" & _
'            " and mz01 is not null"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      cnnConnection.BeginTrans
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         strSql = "update postzipdata" & _
'                  " set PZD10='" & adoRecordset.Fields("mz06") & "',PZD11='" & adoRecordset.Fields("mz07") & "'" & _
'                  " where PZD02='" & adoRecordset.Fields("PZD02") & "'" & _
'                  " and PZD03='" & adoRecordset.Fields("PZD03") & "'"
'         cnnConnection.Execute strSql
'         adoRecordset.MoveNext
'      Loop
'      cnnConnection.CommitTrans
'   End If
   
   
''   For i = 1 To 2
''      If i = 1 Then '發文未超過2個月則掛2個月期限
''         strSql = "select c1.cp09,c1.cp01,c1.cp02,c1.cp03,c1.cp04,'305'," & _
''                  "to_char(add_months(to_date(c1.cp27,'YYYYMMDD'),2),'YYYYMMDD'),to_char(add_months(to_date(c1.cp27,'YYYYMMDD'),2),'YYYYMMDD'),c1.cp13" & _
''                  " from caseprogress c1,nextprogress" & _
''                  " where c1.cp01 in('FCT') and c1.cp10='717' and c1.cp27 is not null" & _
''                  " and c1.cp09=np01(+)" & _
''                  " and c1.cp01=np02(+)" & _
''                  " and c1.cp02=np03(+)" & _
''                  " and c1.cp03=np04(+)" & _
''                  " and c1.cp04=np05(+)" & _
''                  " and (np07 is null or np07<>'305')" & _
''                  " and not exists (select c2.cp01 from caseprogress c2 where c1.cp01=c2.cp01 and c1.cp02=c2.cp02 and c1.cp03=c2.cp03 and c1.cp04=c2.cp04 and c2.cp10 ='1701')" & _
''                  " and to_char(add_months(to_date(c1.cp27,'YYYYMMDD'),2),'YYYYMMDD')>" & strSrvDate(1)
''      Else '若發文已超過2個月則掛系統日
''         strSql = "select c1.cp09,c1.cp01,c1.cp02,c1.cp03,c1.cp04,'305'," & strSrvDate(1) & "," & strSrvDate(1) & ",c1.cp13" & _
''                  " from caseprogress c1,nextprogress" & _
''                  " where c1.cp01 in('FCT') and c1.cp10='717' and c1.cp27 is not null" & _
''                  " and c1.cp09=np01(+)" & _
''                  " and c1.cp01=np02(+)" & _
''                  " and c1.cp02=np03(+)" & _
''                  " and c1.cp03=np04(+)" & _
''                  " and c1.cp04=np05(+)" & _
''                  " and (np07 is null or np07<>'305')" & _
''                  " and not exists (select c2.cp01 from caseprogress c2 where c1.cp01=c2.cp01 and c1.cp02=c2.cp02 and c1.cp03=c2.cp03 and c1.cp04=c2.cp04 and c2.cp10 ='1701')" & _
''                  " and to_char(add_months(to_date(c1.cp27,'YYYYMMDD'),2),'YYYYMMDD')<=" & strSrvDate(1)
''      End If
'      '檢查是否為一申請書多件
'      '"and not exists (select * from caseprogress c2 where c2.cp43=c1.cp09) "
'      strSql = "select cp01,cp02,cp03,cp04,cp09,cp10,cp27,cp43,cp123,cp28 " & _
'               "from caseprogress c1,trademark " & _
'               "where c1.cp01='FCT' and c1.cp10 in('301','501','502','504') " & _
'               "and c1.cp27 is not null and c1.cp24 is null " & _
'               "and c1.cp01=tm01(+) and c1.cp02=tm02(+) and c1.cp03=tm03(+) and c1.cp04=tm04(+) " & _
'               "and tm29 is null and tm57 is null " & _
'               "and c1.cp27>=20120701 "
'      If adoRecordset.State = adStateOpen Then
'         adoRecordset.Close
'      End If
'      adoRecordset.CursorLocation = adUseClient
'      adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If adoRecordset.RecordCount > 0 Then
'         cnnConnection.BeginTrans
'         adoRecordset.MoveFirst
'         Do While Not adoRecordset.EOF
'            Call PUB_UpdateCP148(adoRecordset.Fields("cp01"), adoRecordset.Fields("cp02"), adoRecordset.Fields("cp03"), adoRecordset.Fields("cp04"), adoRecordset.Fields("cp10"), "" & adoRecordset.Fields("cp27"))
'            adoRecordset.MoveNext
'         Loop
'         cnnConnection.CommitTrans
'      End If
''   Next i
   
   Screen.MousePointer = vbDefault
   MsgBox "更新完畢"
   Exit Sub
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Sub

'Add By Sindy 2012/12/24 產生文字檔
Private Sub Command8_Click()
Dim ff As Integer
Dim A01 As String
Dim A02 As String
Dim A03 As String
Dim A04 As String
Dim A05 As String
Dim A06 As String
Dim A07 As String
Dim TempFileName As String
Dim Rs As New ADODB.Recordset
   
'   strSql = "select fa01,fa02,decode(DL10,'FAGENT','代理人','PotCustCont','聯絡人','PotCustomer','潛在客戶',DL10),dl06,dl07,dl08,dl09 from dml_log,FAGENT where dl06='89037' and dl07=20120215 and DL10='FAGENT' " & _
'            "and substr(dl09,22,8)=fa01(+) and substr(dl09,37,1)=fa02(+) Union " & _
'            "select pcc01,pcc02,decode(DL10,'FAGENT','代理人','PotCustCont','聯絡人','PotCustomer','潛在客戶',DL10),dl06,dl07,dl08,dl09 from dml_log,PotCustCont where dl06='89037' and dl07=20120215 and DL10='PotCustCont' " & _
'            "and substr(dl09,28,8)=pcc01(+) and substr(dl09,44,2)=pcc02(+) Union " & _
'            "select pcu01,pcu02,decode(DL10,'FAGENT','代理人','PotCustCont','聯絡人','PotCustomer','潛在客戶',DL10),dl06,dl07,dl08,dl09 from dml_log,PotCustomer where dl06='89037' and dl07=20120215 and DL10='PotCustomer' " & _
'            "and substr(dl09,28,8)=pcu01(+) and substr(dl09,44,1)=pcu02(+) " & _
'            "order by 1,2 "
   strSql = "select fa01,fa02,decode(DL10,'FAGENT','代理人','PotCustCont','聯絡人','PotCustomer','潛在客戶',DL10),dl06,dl07,dl08,dl09 from dml_log,FAGENT where dl06='89037' and dl07=20121222 and DL10='FAGENT' " & _
            "and substr(dl09,22,8)=fa01(+) and substr(dl09,37,1)=fa02(+) Union " & _
            "select pcc01,pcc02,decode(DL10,'FAGENT','代理人','PotCustCont','聯絡人','PotCustomer','潛在客戶',DL10),dl06,dl07,dl08,dl09 from dml_log,PotCustCont where dl06='89037' and dl07=20121222 and DL10='PotCustCont' " & _
            "and substr(dl09,28,8)=pcc01(+) and substr(dl09,44,2)=pcc02(+) " & _
            "order by 1,2 "
   CheckOC
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   With Rs
      If .RecordCount > 0 Then
         .MoveFirst
         TempFileName = ""
         Do While Not .EOF
            If TempFileName = "" Then
               TempFileName = "產生文字檔"
               ff = FreeFile
               If ff > 0 Then Close #ff
               ff = FreeFile
               Open App.path & "\" & TempFileName & ".txt" For Output As ff
'               Print #ff, "代理人編號  備註                           修改人員             修改內容    "
'               Print #ff, "======== == ============================== ==================== ============"
               Print #ff, "代理人編號  檔案          "
               Print #ff, "======== == =============="
            End If
            A01 = convForm(Trim(CheckStr(.Fields(0).Value)), 8)
            A02 = convForm(Trim(CheckStr(.Fields(1).Value)), 2)
            A03 = convForm(Trim(CheckStr(.Fields(2).Value)), 10)
            A04 = convForm(Trim(CheckStr(.Fields(3).Value)), 6)
            A05 = convForm(Trim(CheckStr(.Fields(4).Value)), 8)
            A06 = convForm(Trim(CheckStr(.Fields(5).Value)), 6)
            A07 = CheckStr(.Fields(6).Value)
            'Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05 & " " & A06 & " " & A07
            Print #ff, A01 & " " & A02 & " " & A03
            .MoveNext
         Loop
         Close ff
      End If
   End With
   MsgBox "電子檔產生完畢！"
   
   Set Rs = Nothing
End Sub

'限定字串長度
'Remove by Lydia 2018/08/24 與basQuery重複
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function

'Add By Sindy 2012/12/27 暫時-轉檔程式
Private Sub Command9_Click()
Dim m_A1K06 As String, m_A1K10 As String, m_A1k08 As String
Dim m_A1K18 As String, m_A1K02 As String, m_A1K01 As String
Dim m_A1K11 As String, strA1K06 As String
Dim strAppl(1 To 10) As String, strTPB13 As String 'Add By Sindy 2017/2/21
   
On Error GoTo CheckingErr
   
   Exit Sub
   'Add By Sindy 2015/5/15 更新CPP11=null
   strSql = "select * from tpbulletin_sonia"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         cnnConnection.BeginTrans
         strAppl(1) = "" & adoRecordset.Fields("TPB12")
         strAppl(2) = "" & adoRecordset.Fields("TPB13")
         strAppl(3) = "" & adoRecordset.Fields("TPB14")
         strAppl(4) = "" & adoRecordset.Fields("TPB15")
         strAppl(5) = "" & adoRecordset.Fields("TPB16")
         strAppl(6) = "" & adoRecordset.Fields("TPB17")
         strAppl(7) = "" & adoRecordset.Fields("TPB18")
         strAppl(8) = "" & adoRecordset.Fields("TPB19")
         strAppl(9) = "" & adoRecordset.Fields("TPB20")
         strAppl(10) = "" & adoRecordset.Fields("TPB21")
         strTPB13 = GetPatentIPC("3", "" & adoRecordset.Fields("TPB10"), "" & adoRecordset.Fields("TPB02"))
         
         strSql = "update tpbulletin set " & _
                  "TPB14=" & CNULL(strAppl(1)) & _
                  ",TPB15=" & CNULL(strAppl(2)) & _
                  ",TPB16=" & CNULL(strAppl(3)) & _
                  ",TPB17=" & CNULL(strAppl(4)) & _
                  ",TPB18=" & CNULL(strAppl(5)) & _
                  ",TPB19=" & CNULL(strAppl(6)) & _
                  ",TPB20=" & CNULL(strAppl(7)) & _
                  ",TPB21=" & CNULL(strAppl(8)) & _
                  ",TPB22=" & CNULL(strAppl(9)) & _
                  ",TPB23=" & CNULL(strAppl(10)) & _
                  " where TPB01='" & adoRecordset.Fields("TPB01") & "'"
         cnnConnection.Execute strSql
         
         strSql = "update tpbulletin set " & _
                  "TPB13=" & CNULL(strTPB13) & _
                  " where TPB01='" & adoRecordset.Fields("TPB01") & "' and TPB13 is null"
         cnnConnection.Execute strSql
         
         cnnConnection.CommitTrans
         adoRecordset.MoveNext
      Loop
   End If
   MsgBox "轉檔完畢!!!"
   Exit Sub
   
'   'Add By Sindy 2015/5/5 更新Acc1V0的substr(a1v02,1,1)='X',a1v12.案件性質和a1v13.申請國家
'   strSql = "select a1v01,a1v02,DECODE(PA09,'000',CPM03,CPM04) as Property, pa09 as nation,na03 from caseprogress, casepropertyMap, patent,acc1v0,nation where substr(a1v02,1,1)='X' and a1v12 is null and a1v01=cp09(+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and pa09=na01(+)" & _
'            " Union select a1v01,a1v02,DECODE(TM10,'000',CPM03,CPM04) as Property, tm10 as nation,na03 from caseprogress, casepropertyMap, trademark,acc1v0,nation where substr(a1v02,1,1)='X' and a1v12 is null and a1v01=cp09(+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and tm10=na01(+)" & _
'            " Union select a1v01,a1v02,DECODE(LC15,'000',CPM03,CPM04) as Property, lc15 as nation,na03 from caseprogress, casepropertyMap, lawcase,acc1v0,nation where substr(a1v02,1,1)='X' and a1v12 is null and a1v01=cp09(+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and lc15=na01(+)" & _
'            " Union select a1v01,a1v02,nvl(cpm03, cpm04) as Property, null as nation,null from caseprogress, casepropertyMap, hirecase,acc1v0 where substr(a1v02,1,1)='X' and a1v12 is null and a1v01=cp09(+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04" & _
'            " Union select a1v01,a1v02,DECODE(SP09,'000',CPM03,CPM04) as Property, sp09 as nation,na03 from caseprogress, casepropertyMap, servicepractice,acc1v0,nation where substr(a1v02,1,1)='X' and a1v12 is null and a1v01=cp09(+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and sp09=na01(+)" & _
'            " order by a1v01 asc"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      cnnConnection.BeginTrans
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         strSql = "update Acc1V0" & _
'                  " set a1v12='" & adoRecordset.Fields(2) & "',a1v13='" & adoRecordset.Fields(4) & "'" & _
'                  " where a1v01='" & adoRecordset.Fields(0) & "'" & _
'                  " and a1v02='" & adoRecordset.Fields(1) & "'"
'         cnnConnection.Execute strSql
'         adoRecordset.MoveNext
'      Loop
'      cnnConnection.CommitTrans
'   End If
'   'Exit Sub
   
   'ACC0Z0之A0Z12>0 : 無ACC1V0的資料補新增至ACC1V0
'   insert into acc1v0(a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v18)
'   select d.cp09
'   ,d.A0Z02
'   ,GetA0k11(d.cp09)
'   ,a0z12
'   ,decode(a1k29,'Y','N','Y')
'   ,a0z12
'   ,0
'   ,substr(a0y02,1,length(a0y02)-4),
'1
'   from (
'   select min(cp09) cp09,A0Z02 from acc0z0,caseprogress,acc1v0
'   where a0z12>0 and a0z02=a1v02(+) and a1v02 is null
'   and cp60(+)=a0z02
'   group by A0Z02) d
'   ,acc1k0,acc0y0,acc0z0 z2
'   Where a1k01 = D.A0Z02
'   and z2.A0Z02=d.A0Z02
'   and z2.A0Z01=a0y01
'   ;
   
   'Add By Sindy 2015/5/5 更新資料庫：ACC0Z0之A0Z12>0
   '1. 無ACC1V0的資料補新增至ACC1V0
   '2. 依程式規則更新ACC1K0之A1K35
   strSql = "select A0Z02,cuname from(" & _
            " select A0Z02,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) as cuname from (select A0Z01,A0Z02 from acc0z0 where A0Z12>0),acc0y0,customer" & _
            " where a0z01=a0y01(+) and a0y18='1'" & _
            " and substr(a0y07,1,8)=cu01 and substr(a0y07,9)=cu02" & _
            " Union select A0Z02,nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)) as cuname from (select A0Z01,A0Z02 from acc0z0 where A0Z12>0),acc0y0,fagent" & _
            " where a0z01=a0y01(+) and a0y18='1'" & _
            " and substr(a0y07,1,8)=fa01 and substr(a0y07,9)=fa02" & _
            " Union select A0Z02,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) as cuname from (select A0Z01,A0Z02 from acc0z0 where A0Z12>0),acc0y0,customer" & _
            " where a0z01=a0y01(+) and a0y18='2'" & _
            " and substr(a0y08,1,8)=cu01 and substr(a0y08,9)=cu02" & _
            " Union select A0Z02,nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)) as cuname from (select A0Z01,A0Z02 from acc0z0 where A0Z12>0),acc0y0,fagent" & _
            " where a0z01=a0y01(+) and a0y18='2'" & _
            " and substr(a0y08,1,8)=fa01 and substr(a0y08,9)=fa02" & _
            " Union select A0Z02,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) as cuname from (select A0Z01,A0Z02 from acc0z0 where A0Z12>0),acc0y0,customer" & _
            " where a0z01=a0y01(+) and a0y18='3'" & _
            " and substr(a0y09,1,8)=cu01 and substr(a0y09,9)=cu02" & _
            " Union select A0Z02,nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)) as cuname from (select A0Z01,A0Z02 from acc0z0 where A0Z12>0),acc0y0,fagent" & _
            " where a0z01=a0y01(+) and a0y18='3'" & _
            " and substr(a0y09,1,8)=fa01 and substr(a0y09,9)=fa02" & _
            ") group by A0Z02,cuname"
   If adoRecordset.State = adStateOpen Then
      adoRecordset.Close
   End If
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      cnnConnection.BeginTrans
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         strSql = "update ACC1K0" & _
                  " set A1K35='" & ChgSQL(adoRecordset.Fields(1)) & "'" & _
                  " where A1K01='" & adoRecordset.Fields(0) & "'" & _
                  " and A1K35 is null"
         cnnConnection.Execute strSql
         adoRecordset.MoveNext
      Loop
      cnnConnection.CommitTrans
   End If
   MsgBox "轉檔完畢!!!"
   Exit Sub
   
   
'   'Add By Sindy 2013/9/25
'   strSql = "select eef01,eef02,eef03,cp01||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','',cp04)||eef03" & _
'            " From empelectronfile,caseprogress" & _
'            " where substr(eef03,1,1)='.'" & _
'            " and eef01=cp09(+)"
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      cnnConnection.BeginTrans
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         strSql = "update empelectronfile " & _
'                  " set eef03='" & adoRecordset.Fields(3) & "'" & _
'                  " where eef01='" & adoRecordset.Fields("eef01") & "'" & _
'                  " and eef02=" & adoRecordset.Fields("eef02") & "" & _
'                  " and eef03='" & adoRecordset.Fields("eef03") & "'"
'         cnnConnection.Execute strSql
'         adoRecordset.MoveNext
'      Loop
'      cnnConnection.CommitTrans
'   End If
'
'   Exit Sub
'
''   【轉檔需知】
''   920201
''   幣別非USD才需要轉檔, NTD???
''   /*
''   Select a1k18,max(a1k02) a_a1k02 FROM acc1k0 group by a1k18 order by a_a1k02;
''   A1K1 A_A1K02
''   ---- ----------
''   DM 861004
''   NTD 900508
''   EUR 1011015
''   RMB 1011226
''   USD 1011227
''   選取了 5 筆資料列.
''   */
''   select count(*) from acc1k0 where a1k02>=920201 and a1k18<>'USD'; --1011228 4059筆 有折讓的有7筆
''
''   select a1k18,count(*) from acc1k0 where a1k02>=920201 and a1k18<>'USD'
''   group by a1k18; --1011228 4059筆 有折讓的有7筆
''
''   A1K1 COUNT(*)
''   ---- ----------
''   EUR 3
''   RMB 4056
''   選取了 2 筆資料列.
''   A1K06 NUMBER(13,2)   台幣折讓金額          A1k06=fix(A1k06 * A1k10)
''   A1K10 NUMBER(11,6)   請款幣別對台幣匯率    A1k10=PUB_GetUSXRate_1(A1k02, A1k18)
''   A1K08 NUMBER(13, 2)  請款幣別請款金額      A1K08=A1K11 / PUB_GetUSXRate_1(A1k02, A1k18)
''and a1k18='RMB' and a1k01='X10115563'
'   strSql = "select * " & _
'            "from acc1k0 " & _
'            "where a1k02>=920201 and a1k18='RMB' order by a1k02,a1k01 asc "
'   If adoRecordset.State = adStateOpen Then
'      adoRecordset.Close
'   End If
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount > 0 Then
'      cnnConnection.BeginTrans
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         m_A1K06 = "" & adoRecordset.Fields("A1K06")
'         m_A1K10 = "" & adoRecordset.Fields("A1K10")
'         m_A1k08 = "" & adoRecordset.Fields("A1K08")
'         m_A1K18 = "" & adoRecordset.Fields("A1K18")
'         m_A1K02 = "" & adoRecordset.Fields("A1K02")
'         m_A1K01 = "" & adoRecordset.Fields("A1K01")
'         m_A1K11 = "" & adoRecordset.Fields("A1K11")
'         If m_A1K06 = "" Then
'            strA1K06 = "Null"
'         Else
'            strA1K06 = Fix(Val(m_A1K06) * m_A1K10)
'         End If
'         'fix(m_A1K11 / PUB_GetUSXRate_1(m_A1K02, m_A1K18) * 100)/100 '取小數2位,無條件捨去
'         strSql = "update acc1k0 " & _
'                  " set A1K06=" & strA1K06 & ", " & _
'                  "A1K10=" & PUB_GetUSXRate_1(m_A1K02, m_A1K18) & ", " & _
'                  "A1K08=" & Fix(m_A1K11 / PUB_GetUSXRate_1(m_A1K02, m_A1K18)) & " " & _
'                  " where A1K01='" & m_A1K01 & "'"
'         'cnnConnection.Execute strSql
'         adoRecordset.MoveNext
'      Loop
'      cnnConnection.CommitTrans
'   End If
   
   MsgBox "轉檔完畢!!!"
   Exit Sub
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
InitialGrid
If GetStaffDepartment(strUserNum) = "M51" Then
   Command2.Visible = True
End If

txtCC = strUserNum 'Added by Morgan 2024/4/17
SSTab1.Tab = 0 'Added by Lydia 2018/01/19
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_oFileSys = Nothing
   Set frm000001 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
    If Me.MSHFlexGrid1.Rows > 1 Then
        If Me.MSHFlexGrid1.TextMatrix(1, 1) <> "" Then
            If Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 0) = "" Then
                Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 0) = "V"
            Else
                Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 0) = ""
            End If
        End If
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Me.Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub InitialGrid()
    With Me.MSHFlexGrid1
       .Clear
       .ColHeader(0) = flexColHeaderOn
       .Rows = 2
       .Cols = 2
       .row = 0
       .col = 0
       .ColWidth(0) = 500
       .ColAlignment(0) = flexAlignCenterCenter
       .Text = "V"
       .row = 0
       .col = 1
       .ColWidth(1) = 2500
       .ColAlignment(1) = flexAlignLeftCenter
       .Text = "本所案號"
    End With
End Sub

Private Sub UpdateData()
Dim StrSQLa As String
Dim ii As Integer

With Me.MSHFlexGrid1
    For ii = 1 To .Rows - 1
        '若有勾選資料
        If .TextMatrix(ii, 0) <> "" Then
            'Patent
            StrSQLa = "Update Patent Set PA26='" & Me.Text1(1).Text & "' Where " & ChgPatent(Replace(.TextMatrix(ii, 1), "-", "")) & " And PA26='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Patent Set PA27='" & Me.Text1(1).Text & "' Where " & ChgPatent(Replace(.TextMatrix(ii, 1), "-", "")) & " And PA27='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Patent Set PA28='" & Me.Text1(1).Text & "' Where " & ChgPatent(Replace(.TextMatrix(ii, 1), "-", "")) & " And PA28='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Patent Set PA29='" & Me.Text1(1).Text & "' Where " & ChgPatent(Replace(.TextMatrix(ii, 1), "-", "")) & " And PA29='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Patent Set PA30='" & Me.Text1(1).Text & "' Where " & ChgPatent(Replace(.TextMatrix(ii, 1), "-", "")) & " And PA30='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Patent Set PA75='" & Me.Text1(1).Text & "' Where " & ChgPatent(Replace(.TextMatrix(ii, 1), "-", "")) & " And PA75='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            
            'Trademark
            StrSQLa = "Update Trademark Set TM23='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And TM23='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            'Add By Sindy 2011/2/21
            StrSQLa = "Update Trademark Set TM78='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And TM78='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Trademark Set TM79='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And TM79='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Trademark Set TM80='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And TM80='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Trademark Set TM81='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And TM81='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            '2011/2/21 End
            StrSQLa = "Update Trademark Set TM44='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And TM44='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            
            'Add By Sindy 2011/2/21
            'ServicePractice
            StrSQLa = "Update ServicePractice Set SP08='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And SP08='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update ServicePractice Set SP58='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And SP58='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update ServicePractice Set SP59='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And SP59='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update ServicePractice Set SP65='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And SP65='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update ServicePractice Set SP66='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And SP66='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update ServicePractice Set SP26='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And SP26='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            'Lawcase
            StrSQLa = "Update Lawcase Set LC11='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And LC11='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Lawcase Set LC43='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And LC43='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Lawcase Set LC44='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And LC44='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Lawcase Set LC45='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And LC45='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Lawcase Set LC46='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And LC46='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Lawcase Set LC22='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And LC22='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            'Hirecase
            StrSQLa = "Update Hirecase Set HC05='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And HC05='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Hirecase Set HC24='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And HC24='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Hirecase Set HC25='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And HC25='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Hirecase Set HC26='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And HC26='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update Hirecase Set HC27='" & Me.Text1(1).Text & "' Where " & ChgTradeMark(Replace(.TextMatrix(ii, 1), "-", "")) & " And HC27='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            '2011/2/21 End
            
            'CaseProgress
            StrSQLa = "Update CaseProgress Set CP55='" & Me.Text1(1).Text & "' Where " & ChgCaseprogress(Replace(.TextMatrix(ii, 1), "-", "")) & " And CP55='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update CaseProgress Set CP56='" & Me.Text1(1).Text & "' Where " & ChgCaseprogress(Replace(.TextMatrix(ii, 1), "-", "")) & " And CP56='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update CaseProgress Set CP72='" & Me.Text1(1).Text & "' Where " & ChgCaseprogress(Replace(.TextMatrix(ii, 1), "-", "")) & " And CP72='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
            StrSQLa = "Update CaseProgress Set CP44='" & Me.Text1(1).Text & "' Where " & ChgCaseprogress(Replace(.TextMatrix(ii, 1), "-", "")) & " And CP44='" & Me.Text1(0).Text & "' "
            cnnConnection.Execute StrSQLa
                    
        End If
    Next ii
End With
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Dim StrSQLa As String
    Dim rsA As New ADODB.Recordset
    
    If Me.Text1(Index).Text = "" Then Exit Sub
    Select Case Index
    Case 0, 1 '編號
        Me.Text1(Index).Text = Left(Me.Text1(Index).Text & "000000000", 9)
        StrSQLa = "Select FA01 From Fagent Where FA01='" & Mid(Me.Text1(Index).Text, 1, 8) & "' And FA02='" & Mid(Me.Text1(Index).Text, 9, 1) & "' "
        StrSQLa = StrSQLa & "Union Select CU01 From Customer Where CU01='" & Mid(Me.Text1(Index).Text, 1, 8) & "' And CU02='" & Mid(Me.Text1(Index).Text, 9, 1) & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            Cancel = True
            MsgBox "編號輸入錯誤!!!"
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    Case Else
    End Select
End Sub

'Add By Cheng 2003/05/15
'取得公司別
Private Function GetCompany(strA1902 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCompany = ""
StrSQLa = "select Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp61 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
              "select Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp62 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
              "select Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp63 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
              "select Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp87 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
              "select Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp88 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
              "select Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp61 (+) and cp60 = a0k01 (+) and a1701 = '1' and length(a1702) = 10 and (cp61 is null and cp62 is null and cp63 is null and cp87 is null and cp88 is null) and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
              "select Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind, acc161, acc160 where a1702 = axg01 and a1702 = a1601 and AXG02=CP09 AND CP60 = a0k01 (+) and a1701 = '2' and a1702='" & strA1902 & "' and substr(axg03, 1, length(axg03) - 9)=sk01 union " & _
              "select '2' as a0k11 from acc170 where a1701 = '3' and a1702='" & strA1902 & "' union " & _
              "select '2' as a0k11 from acc170 where a1701 = '4' and a1702='" & strA1902 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCompany = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

Private Function GetNP01(strCaseNo As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetNP01 = ""
StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(strCaseNo) & " And CP09<'C' Order By CP05 Desc "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    Do While Not rsA.EOF
        GetNP01 = "" & rsA("CP09").Value
        If "" & rsA("CP10").Value = "101" Then Exit Do
        rsA.MoveNext
    Loop
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Function GetCKindCP05(strCaseNo As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCKindCP05 = ""
StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(strCaseNo) & " And CP10 In ('1601','1602') "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCKindCP05 = "" & rsA("CP05").Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Sindy 2010/7/9
Private Sub Command2_Click()
Dim strCaseNo As String, strTM09 As String, dblGoods As Double
Dim strTgText As String
Dim j  As Integer, ff As Integer
Dim varTemp As Variant, bolStarWrite As Boolean, bolCompError As Boolean
Dim strSqlText As String, strCP14 As String
   
   strSqlText = ""
   '逐筆比對商品檔案號
   'Modify By Sindy 2014/2/19 +(TG18 is null or TG18='')
   strSql = "select distinct TG01,TG02,TG03,TG04 from tmgoods where (TG18 is null or TG18='') order by 1,2,3,4"
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 Then
      AdoRecordSet3.MoveFirst
      While Not AdoRecordSet3.EOF
         '讀取主檔
         strExc(0) = "select tm09 from trademark where tm01='" & AdoRecordSet3.Fields(0) & "' and tm02='" & AdoRecordSet3.Fields(1) & "' and tm03='" & AdoRecordSet3.Fields(2) & "' and tm04='" & AdoRecordSet3.Fields(3) & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         strTM09 = ""
         If intI = 1 Then
            strTM09 = "" & RsTemp.Fields("TM09")
         End If
         '讀取商品檔數量
         'Modify By Sindy 2014/2/19 +(TG18 is null or TG18='')
         strExc(0) = "select count(*) from tmgoods where (TG18 is null or TG18='') and tg01='" & AdoRecordSet3.Fields(0) & "' and tg02='" & AdoRecordSet3.Fields(1) & "' and tg03='" & AdoRecordSet3.Fields(2) & "' and tg04='" & AdoRecordSet3.Fields(3) & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         dblGoods = 0
         If intI = 1 Then
            dblGoods = RsTemp.Fields(0)
         End If
         '開始比對商品類別
         bolCompError = False
         If strTM09 = "" Then bolCompError = True: GoTo ReadNext
         If strTM09 <> "" Then
            varTemp = Split(strTM09, ",")
            '比對TG
            For j = 0 To UBound(varTemp)
               'Modify By Sindy 2014/2/19 +(TG18 is null or TG18='')
               strExc(0) = "select count(*) from tmgoods where (TG18 is null or TG18='') and tg01='" & AdoRecordSet3.Fields(0) & "' and tg02='" & AdoRecordSet3.Fields(1) & "' and tg03='" & AdoRecordSet3.Fields(2) & "' and tg04='" & AdoRecordSet3.Fields(3) & "' and tg05='" & Trim(varTemp(j)) & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp.Fields(0) = 0 Then
                     bolCompError = True
                     GoTo ReadNext
                  End If
               End If
            Next j
         End If
         '比對TM
         'Modify By Sindy 2014/2/19 +(TG18 is null or TG18='')
         strExc(0) = "select tg05 from tmgoods where (TG18 is null or TG18='') and tg01='" & AdoRecordSet3.Fields(0) & "' and tg02='" & AdoRecordSet3.Fields(1) & "' and tg03='" & AdoRecordSet3.Fields(2) & "' and tg04='" & AdoRecordSet3.Fields(3) & "' order by tg05 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               If UBound(Split(strTM09, RsTemp.Fields(0))) = 0 Then
                  bolCompError = True
                  GoTo ReadNext
               End If
               RsTemp.MoveNext
            Loop
         End If
ReadNext:
         If bolCompError = True Then
            If strSqlText <> "" Then
               strSqlText = strSqlText & " union "
            End If
            strSqlText = strSqlText & "select tm01,tm02,tm03,tm04,tm09,cp14,st03 from trademark,caseprogress,staff where tm01='" & AdoRecordSet3.Fields(0) & "' and tm02='" & AdoRecordSet3.Fields(1) & "' and tm03='" & AdoRecordSet3.Fields(2) & "' and tm04='" & AdoRecordSet3.Fields(3) & "' and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp31(+)='Y' and cp14=st01(+) "
         End If
         AdoRecordSet3.MoveNext
      Wend
   End If
   
   bolStarWrite = False
   If strSqlText <> "" Then
      strSql = strSqlText & " order by st03,cp14,tm01,tm02,tm03,tm04 "
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
         AdoRecordSet3.MoveFirst
         While Not AdoRecordSet3.EOF
            strCaseNo = Left(AdoRecordSet3.Fields(0) & "-" & AdoRecordSet3.Fields(1) & "-" & AdoRecordSet3.Fields(2) & "-" & AdoRecordSet3.Fields(3) & "               ", 17)
            strTM09 = IIf(Len("" & AdoRecordSet3.Fields("TM09")) <= 20, Left("" & AdoRecordSet3.Fields("TM09") & "                    ", 20), "" & AdoRecordSet3.Fields("TM09") & "  ")
            strCP14 = Left("" & AdoRecordSet3.Fields("CP14") & " " & GetStaffName("" & AdoRecordSet3.Fields("CP14"), True) & "            ", 12)
            '組合TG05資料
            strTgText = ""
            'Modify By Sindy 2014/2/19 +(TG18 is null or TG18='')
            strExc(0) = "select tg05 from tmgoods where (TG18 is null or TG18='') and tg01='" & AdoRecordSet3.Fields(0) & "' and tg02='" & AdoRecordSet3.Fields(1) & "' and tg03='" & AdoRecordSet3.Fields(2) & "' and tg04='" & AdoRecordSet3.Fields(3) & "' order by tg05 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If strTgText = "" Then
                     strTgText = "" & Trim(RsTemp.Fields(0))
                  Else
                     strTgText = strTgText & ", " & "" & Trim(RsTemp.Fields(0))
                  End If
                  RsTemp.MoveNext
               Loop
            End If
            '寫文字檔
            If bolStarWrite = False Then
               If ff > 0 Then Close #ff
               ff = FreeFile
               Open App.path & "\TGTM商品類別比對結果.txt" For Output As ff
               Print #ff, "承辦人           本所案號          TM商品類別           TG商品類別"
               Print #ff, "=============   =============     ===========          ==========="
               bolStarWrite = True
            End If
            Print #ff, strCP14 & " " & strCaseNo & " " & strTM09 & " " & strTgText
            AdoRecordSet3.MoveNext
         Wend
         Close ff
         MsgBox "商品類別比對完成, 詳情請查閱文字檔!"
      Else
         MsgBox "商品類別比對完成, 無不符資料!"
      End If
   Else
      MsgBox "商品類別比對完成, 無不符資料!"
   End If
End Sub

'Add By Sindy 2010/8/30
'CFT商品類別比對
Private Sub Command3_Click()
Dim strTemp As String, bolStarWrite As Boolean, ff As Integer
   
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans
   '讀取主檔
   strSql = "select tm01,tm02,tm03,tm04,tm09 from trademark where tm01='CFT' "
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   bolStarWrite = False
   If AdoRecordSet3.RecordCount <> 0 Then
      AdoRecordSet3.MoveFirst
      While Not AdoRecordSet3.EOF
         If IsNull(AdoRecordSet3.Fields("tm09")) Or Trim(AdoRecordSet3.Fields("tm09")) = "" Then
            strSql = "delete from tmgoods where tg01='" & Trim(AdoRecordSet3.Fields("tm01")) & "' " & _
                                                           "and tg02='" & Trim(AdoRecordSet3.Fields("tm02")) & "' " & _
                                                           "and tg03='" & Trim(AdoRecordSet3.Fields("tm03")) & "' " & _
                                                           "and tg04='" & Trim(AdoRecordSet3.Fields("tm04")) & "' "
            cnnConnection.Execute strSql
         Else
            strTemp = Trim(AdoRecordSet3.Fields("tm09"))
            If Right(Trim(AdoRecordSet3.Fields("tm09")), 1) = "," Then
               strTemp = Left(Trim(AdoRecordSet3.Fields("tm09")), Len(Trim(AdoRecordSet3.Fields("tm09"))) - 1)
            End If
            strExc(0) = "select count(*) from tmgoods where tg01='" & Trim(AdoRecordSet3.Fields("tm01")) & "' " & _
                                                           "and tg02='" & Trim(AdoRecordSet3.Fields("tm02")) & "' " & _
                                                           "and tg03='" & Trim(AdoRecordSet3.Fields("tm03")) & "' " & _
                                                           "and tg04='" & Trim(AdoRecordSet3.Fields("tm04")) & "' " & _
                                                           "and tg05 not in (" & strTemp & ") "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) > 0 Then
                  '寫文字檔
                  If bolStarWrite = False Then
                     If ff > 0 Then Close #ff
                     ff = FreeFile
                     Open App.path & "\CFT商品類別比對結果.txt" For Output As ff
                     Print #ff, "TM01           TM02          TM03           TM04           TM09"
                     Print #ff, "=============   =============     ===========          ===========          ==========="
                     bolStarWrite = True
                  End If
                  Print #ff, Trim(AdoRecordSet3.Fields("tm01")) & " " & Trim(AdoRecordSet3.Fields("tm02")) & " " & Trim(AdoRecordSet3.Fields("tm03")) & " " & Trim(AdoRecordSet3.Fields("tm04")) & " " & Trim(AdoRecordSet3.Fields("tm09"))
               End If
            End If
'            strSql = "delete from tmgoods where tg01='" & Trim(AdoRecordSet3.Fields("tm01")) & "' " & _
'                                                           "and tg02='" & Trim(AdoRecordSet3.Fields("tm02")) & "' " & _
'                                                           "and tg03='" & Trim(AdoRecordSet3.Fields("tm03")) & "' " & _
'                                                           "and tg04='" & Trim(AdoRecordSet3.Fields("tm04")) & "' " & _
'                                                           "and tg05 not in (" & strTemp & ") "
'            cnnConnection.Execute strSql
         End If
         AdoRecordSet3.MoveNext
      Wend
   End If
   Close ff
   cnnConnection.CommitTrans
   MsgBox "檢查完成!!!"
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2011/1/27 五都修改地址通知函
Private Sub Command4_Click()
Dim nPageNo As Long
   
   On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   
   'CU112 : 207~253 新北市
   '              411~439 台中縣
   '              710~745 台南縣
   '              814~852 高雄縣
   'Modify By Sindy 2012/5/24 DECODE(CU15,'0','台端','貴公司')==>DECODE(CU15,'0','台端','1','貴公司','貴單位')
   strSql = "select distinct(t.id) as custid,c1.cu112 as cu112,NVL(NVL(CU104,CU04),CU05||CU88||CU89||CU90) C00,NVL(NVL(NVL(NVL(PCC05,CU08),CU104),CU04),CU05||CU88||CU89||CU90) C01,DECODE(CU15,'0','台端','1','貴公司','貴單位') C02,ST02 C03,NVL(PCC21,CU30) C04,NVL(NVL(PCC22,CU31),CU23) C05,decode(ST22,'F','小姐','先生') C06,substr(CU01,1,6) C07,c1.cu04 as cu04,c1.cu23 as cu23,c1.cu30 as cu30,c1.cu31 as cu31 from customer c1,POTCUSTCONT,STAFF,( " & _
               "SELECT pa01,pa02,pa03,pa04,pa26 as id,pa09,pa57,pa108,pa136,pa25,pa17,pa16 FROM Patent,Customer WHERE pa01='CFP' and pa09 in('011','012') and cu01=substr(pa26,1,8) and cu02=substr(pa26,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and pa57 is null and pa108 is null and pa136 is null and (pa25 is null or (pa25 is not null and pa25>" & strSrvDate(1) & " and pa17='Y')) and (pa16 is null or pa16<>'2') " & _
               "Union SELECT pa01,pa02,pa03,pa04,pa27 as id,pa09,pa57,pa108,pa136,pa25,pa17,pa16 FROM Patent,Customer WHERE pa01='CFP' and pa09 in('011','012') and cu01=substr(pa27,1,8) and cu02=substr(pa27,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and pa57 is null and pa108 is null and pa136 is null and (pa25 is null or (pa25 is not null and pa25>" & strSrvDate(1) & " and pa17='Y')) and (pa16 is null or pa16<>'2') " & _
               "Union SELECT pa01,pa02,pa03,pa04,pa28 as id,pa09,pa57,pa108,pa136,pa25,pa17,pa16 FROM Patent,Customer WHERE pa01='CFP' and pa09 in('011','012') and cu01=substr(pa28,1,8) and cu02=substr(pa28,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and pa57 is null and pa108 is null and pa136 is null and (pa25 is null or (pa25 is not null and pa25>" & strSrvDate(1) & " and pa17='Y')) and (pa16 is null or pa16<>'2') " & _
               "Union SELECT pa01,pa02,pa03,pa04,pa29 as id,pa09,pa57,pa108,pa136,pa25,pa17,pa16 FROM Patent,Customer WHERE pa01='CFP' and pa09 in('011','012') and cu01=substr(pa29,1,8) and cu02=substr(pa29,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and pa57 is null and pa108 is null and pa136 is null and (pa25 is null or (pa25 is not null and pa25>" & strSrvDate(1) & " and pa17='Y')) and (pa16 is null or pa16<>'2') " & _
               "Union SELECT pa01,pa02,pa03,pa04,pa30 as id,pa09,pa57,pa108,pa136,pa25,pa17,pa16 FROM Patent,Customer WHERE pa01='CFP' and pa09 in('011','012') and cu01=substr(pa30,1,8) and cu02=substr(pa30,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and pa57 is null and pa108 is null and pa136 is null and (pa25 is null or (pa25 is not null and pa25>" & strSrvDate(1) & " and pa17='Y')) and (pa16 is null or pa16<>'2') " & _
               "Union SELECT tm01,tm02,tm03,tm04,tm23 as id,tm10,tm29,tm57,tm73,tm22,tm17,tm16 FROM Trademark,Customer WHERE ((tm01='T' and tm10='020') or (tm01='CFT' and tm10 in('011','012','014','239'))) and cu01=substr(tm23,1,8) and cu02=substr(tm23,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and tm29 is null and tm57 is null and tm73 is null and (tm22 is null or (tm22 is not null and tm22>" & strSrvDate(1) & " and tm17='Y')) and (tm16 is null or tm16<>'2') " & _
               "Union SELECT tm01,tm02,tm03,tm04,tm78 as id,tm10,tm29,tm57,tm73,tm22,tm17,tm16 FROM Trademark,Customer WHERE ((tm01='T' and tm10='020') or (tm01='CFT' and tm10 in('011','012','014','239'))) and cu01=substr(tm78,1,8) and cu02=substr(tm78,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and tm29 is null and tm57 is null and tm73 is null and (tm22 is null or (tm22 is not null and tm22>" & strSrvDate(1) & " and tm17='Y')) and (tm16 is null or tm16<>'2') " & _
               "Union SELECT tm01,tm02,tm03,tm04,tm79 as id,tm10,tm29,tm57,tm73,tm22,tm17,tm16 FROM Trademark,Customer WHERE ((tm01='T' and tm10='020') or (tm01='CFT' and tm10 in('011','012','014','239'))) and cu01=substr(tm79,1,8) and cu02=substr(tm79,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and tm29 is null and tm57 is null and tm73 is null and (tm22 is null or (tm22 is not null and tm22>" & strSrvDate(1) & " and tm17='Y')) and (tm16 is null or tm16<>'2') " & _
               "Union SELECT tm01,tm02,tm03,tm04,tm80 as id,tm10,tm29,tm57,tm73,tm22,tm17,tm16 FROM Trademark,Customer WHERE ((tm01='T' and tm10='020') or (tm01='CFT' and tm10 in('011','012','014','239'))) and cu01=substr(tm80,1,8) and cu02=substr(tm80,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and tm29 is null and tm57 is null and tm73 is null and (tm22 is null or (tm22 is not null and tm22>" & strSrvDate(1) & " and tm17='Y')) and (tm16 is null or tm16<>'2') " & _
               "Union SELECT tm01,tm02,tm03,tm04,tm81 as id,tm10,tm29,tm57,tm73,tm22,tm17,tm16 FROM Trademark,Customer WHERE ((tm01='T' and tm10='020') or (tm01='CFT' and tm10 in('011','012','014','239'))) and cu01=substr(tm81,1,8) and cu02=substr(tm81,9,1) and ((cu112 between '２０７' and '２５３') or (cu112 between '４１１' and '４３９') or (cu112 between '７１０' and '７４５') or (cu112 between '８１４' and '８５２')) " & _
               "and tm29 is null and tm57 is null and tm73 is null and (tm22 is null or (tm22 is not null and tm22>" & strSrvDate(1) & " and tm17='Y')) and (tm16 is null or tm16<>'2') " & _
               ") t " & _
               "Where C1.cu01 = substr(t.Id, 1, 8) And C1.cu02 = substr(t.Id, 9, 1) " & _
               "AND ST01(+)=C1.CU13 AND PCC01(+)=C1.CU01 AND PCC02(+)=C1.CU127 " & _
               "order by 2,1 "
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With adoRecordset
         nPageNo = 0
         Printer.Orientation = vbPRORPortrait '橫印
         Printer.Font = "標楷體"
         Printer.FontSize = 12
         .MoveFirst
         Do While .EOF = False
            nPageNo = nPageNo + 1
            If nPageNo > 1 Then Printer.NewPage
            PrintLetter2 "" & .Fields("C00"), "" & .Fields("C01"), "" & .Fields("C02"), "" & .Fields("C03"), "" & .Fields("C04"), "" & .Fields("C05"), "" & .Fields("C06"), "Y", nPageNo
            .MoveNext
         Loop
         Printer.EndDoc
      End With
   End If
   
   Screen.MousePointer = vbDefault
   MsgBox "五都地址通知函產生完畢!!!"
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      'cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   'Resume
End Sub

'Add By Sindy 2011/1/27
Private Sub PrintLetter2(ByVal stCustName As String, ByVal stContact As String, ByVal stCustType As String, ByVal stSalesName As String, ByVal stZipNo As String, ByVal stAddr As String, ByVal stST22 As String, ByVal strAddrType As String, ByVal lngPage As Long)
Const cLMargin = 950 '左留白
Const cUMargin = 2700 '上留白
Const cVPad = 400 '列距

Dim stDoc As String '文章內容
Dim stSentence As String '一列文字
Dim lngWMax As Long '可印行寬
Dim lngWRest As Long '剩餘可印行寬
Dim iAWord As Integer '一個字寬
Dim sChar As String
Dim lPos As Long
Dim lngX As Long, lngY As Long '列印位置
Dim sBoldi As String '粗體開
Dim sBoldo As String '粗體關
Dim sUnderlinei As String '底線開
Dim sUnderlineo As String '底線關
Dim stIdData As String
Dim intLine As Integer
Dim stContact2 As String
Dim ExceptFieldData2 As String
Dim m_line As Variant
Dim ii As Integer
   
   Printer.FontSize = 14
   
   sUnderlinei = Chr(28) '底線開
   sUnderlineo = Chr(29) '底線關
   sBoldi = Chr(30) '粗體開
   sBoldo = Chr(31) '粗體關
   
   iAWord = Printer.TextWidth("　")
   lngWMax = iAWord * ((Printer.ScaleWidth - 2 * cLMargin) \ iAWord) + 50
   lngY = cUMargin
   
   '印頁次
   stSentence = Format(lngPage, "000000")
   lngX = cLMargin + lngWMax - Printer.TextWidth(stSentence)
   Printer.CurrentX = lngX: Printer.CurrentY = lngY
   Printer.Print stSentence
      
   lngX = cLMargin
   lngY = lngY + cVPad
   lngWRest = lngWMax
   stSentence = ""
   
   intLine = 0
'   If Text1(9).Text = "83008" Then
      stIdData = sUnderlinei & "王文德同仁" & sUnderlineo & "已於九十八年三月三十一月離職"
'   ElseIf Text1(9).Text = "73001" Then
'      stIdData = sUnderlinei & "林錦山同仁" & sUnderlineo & "已於九十八年三月三十一月自本所退休"
'   End If
   
   stContact2 = stContact
   If strAddrType = "N" Then
      stZipNo = ""
      stAddr = ""
      stContact = ""
   End If
   
   '處理寄信的地址及收件人
   ExceptFieldData2 = Trim(stAddr) & vbLf
   If stCustType = "台端" Then
      ExceptFieldData2 = ExceptFieldData2 & Trim(stCustName) & "　　　鈞啟" & vbLf
   Else
      ExceptFieldData2 = ExceptFieldData2 & Trim(stCustName) & vbLf & _
                                                                      Trim(stContact) & "　　　鈞啟" & vbLf
   End If
   If ExceptFieldData2 <> "" Then
      m_line = Split(ExceptFieldData2, vbLf)
      For ii = 0 To UBound(m_line)
           ExceptFieldData2 = m_line(ii)
           Do While ExceptFieldData2 <> StrToStr(ExceptFieldData2, 17)
               If InStr(1, m_line(ii), StrToStr(ExceptFieldData2, 17)) = 1 Then
                   m_line(ii) = Mid(m_line(ii), 1, InStr(1, m_line(ii), StrToStr(ExceptFieldData2, 17)) - 1) & StrToStr(ExceptFieldData2, 17) & vbLf & Replace(m_line(ii), StrToStr(ExceptFieldData2, 17), "")
               Else
                   m_line(ii) = Mid(m_line(ii), 1, InStr(1, m_line(ii), StrToStr(ExceptFieldData2, 17)) - 1) & StrToStr(ExceptFieldData2, 17) & vbLf & Replace(Mid(m_line(ii), InStr(1, m_line(ii), StrToStr(ExceptFieldData2, 17))), StrToStr(ExceptFieldData2, 17), "")
               End If
               ExceptFieldData2 = Replace(ExceptFieldData2, StrToStr(ExceptFieldData2, 17), "")
           Loop
      Next ii
      ExceptFieldData2 = Join(m_line, vbLf)
      m_line = Split(ExceptFieldData2, vbLf)
      For ii = 0 To UBound(m_line)
           m_line(ii) = m_line(ii)
      Next ii
      ExceptFieldData2 = Join(m_line, vbLf)
      m_line = Split(ExceptFieldData2, vbLf)
      If UBound(m_line) < 3 Then
           ExceptFieldData2 = ExceptFieldData2 & vbLf
      End If
   End If
   
   stDoc = stZipNo & "　　　　　　　　　　　" & sBoldi & "掛號" & sBoldo & vbLf
   stDoc = stDoc & ExceptFieldData2 & _
            "　　　　　　　　　　　　　　　　　　　　" & vbLf & _
            "　　　　　　　　　　　　　　　　　　　　" & vbLf & _
            "　　　　　　　　　　　　　　　　　　　　" & vbLf
   If stCustType = "台端" Then
      stDoc = stDoc & "致：" & stCustName & "　君台鑒" & vbLf
   Else
      stDoc = stDoc & "致：" & stCustName & vbLf & _
                                 "　　" & stContact2 & "　君台鑒" & vbLf
   End If
      stDoc = stDoc & _
            "　　　　　　　　　　　　　　　　　　　　" & vbLf & _
            "　　感謝　" & stCustType & "多年來將智慧財產權之保護與申請案件委由台一國際專利法律事務所處理。" & vbLf & _
            "　　前任職於本所之" & stIdData & "，為順利銜接　" & stCustType & "的案件，本所特指派" & sBoldi & stSalesName & stST22 & sBoldo & "接手　" & stCustType & "各項業務。日後若　" & stCustType & "有任何需要服務之案件或工作，請直接與他聯繫。" & vbLf & _
            "　　聯繫電話：04-2327-0288；Email：taie@seed.net.tw。" & vbLf & _
            "　　請特別留意，若有已委託本所尚未完成交之案件，或付款尚未接獲收據之情形，皆請主動知會" & sUnderlinei & stSalesName & stST22 & sUnderlineo & "，本所將立即處理，以維　" & stCustType & "的權益。" & vbLf & _
            "　　本所對所有的案件都有齊全的管理與整合措施，任何客戶委辦案件皆由主管做雙重審核後再由本所專業人員處理，因此　" & stCustType & "所有案件皆可由接手人員順利銜接，本所一定會做到令您放心。" & vbLf & _
            "　　台一在提供專利、商標、著作權及法律的各項服務時，一直秉持著專業能度與迅速服務的精神，為業界在智慧財產權保護上提供最完善的服務。在本所過去三十餘年的成長過程中，承蒙　" & stCustType & "持續對本所支持與肯定，期待日後能有更多的服務機會。" & vbLf & _
            "　　耑此　　順頌" & vbLf & vbLf & _
            "商祺"
   
   For lPos = 1 To Len(stDoc)
      sChar = Mid(stDoc, lPos, 1)
      
      If sChar = vbLf Then '跳行
         intLine = intLine + 1
         If intLine > 7 Then
            Printer.FontSize = 12
         End If
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         lngX = cLMargin
         lngY = lngY + cVPad
         lngWRest = lngWMax
         stSentence = ""
         
      ElseIf sChar = sBoldi Then  '粗體開
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         lngX = lngX + Printer.TextWidth(stSentence)
         lngWRest = lngWMax - lngX + cLMargin
         Printer.FontBold = True
         stSentence = ""
         
      ElseIf sChar = sBoldo Then  '粗體關
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         lngX = lngX + Printer.TextWidth(stSentence)
         lngWRest = lngWMax - lngX + cLMargin
         Printer.FontBold = False
         stSentence = ""
         
      ElseIf sChar = sUnderlinei Then  '底線開
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         lngX = lngX + Printer.TextWidth(stSentence)
         lngWRest = lngWMax - lngX + cLMargin
         Printer.FontUnderline = True
         stSentence = ""
         
      ElseIf sChar = sUnderlineo Then  '底線關
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         lngX = lngX + Printer.TextWidth(stSentence)
         lngWRest = lngWMax - lngX + cLMargin
         Printer.FontUnderline = False
         stSentence = ""
         
      ElseIf Printer.TextWidth(stSentence & sChar) > lngWRest Then '字數超過一列可印寬
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         
         If sChar = "　" Or sChar = " " Then
            stSentence = ""
         Else
            stSentence = sChar
         End If
         lngX = cLMargin
         lngY = lngY + cVPad
         lngWRest = lngWMax
         
      Else
         stSentence = stSentence & sChar
      End If
   Next
   If stSentence <> "" Then
      Printer.CurrentX = lngX: Printer.CurrentY = lngY
      Printer.Print stSentence
   End If
   
   lngY = lngY + cVPad
   stSentence = "台一國際專利商標事務所　敬上　　" & "98.05.06" 'Year(Now) - 1911 & "." & Format(Now, "MM.DD")
   lngX = cLMargin + lngWMax - Printer.TextWidth(stSentence)
   Printer.CurrentX = lngX: Printer.CurrentY = lngY
   Printer.Print stSentence
End Sub

'Added by Lydia 2018/01/22 大宗Email退件處理
Private Sub Command18_Click()
'Memo by Lydia 2020/04/17 處理Tai E Quarterly 2020Q1 (Undelivered) 備註=>  原本抓"To: "到"Subject: "之間分析,但是中間增加了base64碼或是To和Subject位置顛倒,所以人工刪除冗餘行數
'Memo by Lydia 2021/01/29 【大批退信處理】201023 Tai E's Seasons Greetings (Jerry's video) => 分析出32個email，變更資料記錄36筆(退件郵件整理日期20210127)。
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim strMid As String
Dim iCnt As Integer  '流水號
Dim tLine As Long '檔案目前列-位置
Dim strLine As String '檔案目前列-資料
Dim strToList As String '收件人信箱
Dim bStart As Boolean
Dim intP As Integer, inX As Integer, intU As Integer
Dim srKey As String '判斷收件人信箱
Dim DTime As String  '轉入時間
Dim tmpArr As Variant
'Modified by Lydia 2018/07/20
'Dim strTemp(1 To 5) As String
Dim strTemp(1 To 6) As String
Dim tmpArr2 As Variant, intJ As Integer  'Added by Lydia 2018/01/31
Dim bolAuto As Boolean 'Added by Lydia 2020/04/17

    If Trim(Text2) = "" Then
         MsgBox "請輸入來源檔案!"
         Command17_Click
         Exit Sub
    ElseIf Dir(Text2) = "" Then
         MsgBox "請輸入正確的來源檔案!"
         Command17_Click
         Exit Sub
    End If
    If Trim(Text3) = "" Or lblName.Caption = "" Then
         MsgBox "請輸入提供人員的編號!"
         Text3.SetFocus
         Exit Sub
    End If
    If Trim(Text4) = "" Then
         MsgBox "請輸入退件郵件整理日期!"
         Call Text4_Validate(False)
         Exit Sub
    End If
    
    bolAuto = False 'Added by Lydia 2020/04/17 人工處理txt檔後,重新再讀取txt
    
JumpToReStart: 'Added by Lydia 2020/04/17
    If fso.FileExists(Text2.Text) Then
         Set ts = fso.OpenTextFile(Text2.Text)
         Do While Not ts.AtEndOfStream
            strLine = ts.ReadLine
            tLine = tLine + 1
            If Trim(strLine) = "" Then GoTo SkipLine1
            
            If bStart = False Then
                 '解析的email內容
                'If InStr(strLine, "To: ") > 0 Then
                If Left(strLine, Len("To: ")) = "To: " Then
                    srKey = "Subject: "
                '對方DNS或防火牆退信內容
                'Mark by Lydia 2018/01/29 (保留) Widen回覆:只抓To: 到Subject: 之間的Email信箱
'                'ElseIf InStr(strLine, "Delivery to the following recipients failed permanently:") > 0 Then
'                ElseIf Left(strLine, Len("Delivery to the following recipients failed permanently:")) = "Delivery to the following recipients failed permanently:" Then
'                    srKey = "Reason:"
'                'ElseIf InStr(strLine, "傳遞至下列收件者或群組失敗:") > 0 Then
'                ElseIf Left(strLine, Len("傳遞至下列收件者或群組失敗:")) = "傳遞至下列收件者或群組失敗:" Then
'                    srKey = "郵件"
'                'end 2018/01/29
                End If
                
                If srKey <> "" Then
                    intP = 1
                    bStart = True
                    strMid = strLine
                    strExc(2) = strMid
                End If
            Else
                strMid = strMid & strLine
                strExc(2) = strMid & vbCrLf & strLine
                intP = intP + 1
                '在email中的關鍵字之間抓取收件人email信箱<..@>

                If InStr(strLine, srKey) > 0 Then
                    If srKey = "Subject: " Then
                        strExc(1) = Mid(strMid, InStr(strMid, "<"))
                        strExc(1) = Mid(strExc(1), 1, InStr(strExc(1), ">"))
                        strExc(1) = Mid(strExc(1), 2, Len(strExc(1)) - 2)
                    'Mark by Lydia 2018/01/29 (保留) Widen回覆:只抓To: 到Subject: 之間的Email信箱
'                    ElseIf srKey = "Reason:" Then '範例: Delivery Failure by May (5).txt 由Dns或對方防火牆退回的訊息,不能確定是真的收件人信箱
'                        strExc(1) = Mid(strMid, Len("Delivery to the following recipients failed permanently:") + 1)
'                        strExc(1) = Mid(strExc(1), 1, InStr(strExc(1), srKey) - 1)
'                        strExc(1) = Trim(Replace(strExc(1), "* ", ""))
'                    ElseIf srKey = "郵件" Then
'                        strMid = Replace(strMid, strLine, "") '因為敘述錯誤原因的內容不一定,先將原因拿掉
'                        strExc(1) = Mid(strMid, Len("傳遞至下列收件者或群組失敗:") + 1)
'                        strExc(1) = Trim(strExc(1))
                    'end 2018/01/29
                    End If
                    If InStr(strExc(1), "@") > 0 Then
                         inX = 0
                         tmpArr2 = Split(strToList, ";")
                         For intJ = 0 To UBound(tmpArr2)
                             If strExc(1) = Trim(tmpArr2(intJ)) Then
                                 inX = intJ + 1
                             End If
                         Next intJ
                         If inX = 0 Then
                             strToList = strToList & strExc(1) & ";"
                             iCnt = iCnt + 1
                         End If
                         bStart = False
                         srKey = ""
                    Else
                         GoTo JumpErrMsg
                    End If
                Else
                    If intP > 15 Then
JumpErrMsg:
                        MsgBox "請檢查檔案中的" & vbCrLf & strExc(2) & "是否有不完整的內容!" & vbCrLf, vbCritical, "來源檔案錯誤"
                        If bolAuto = True Then GoTo JumpToReStart 'Added by Lydia 2020/04/17 人工處理txt檔後,重新再讀取txt
                        Debug.Print strMid
                        Exit Sub
                    End If
                End If
                
            End If
SkipLine1:
         Loop
         ts.Close
    End If
    
    If iCnt = 0 Then
        MsgBox "來源檔案無法分析出收件人信箱，請與提供人員確認資料的正確性! ", vbCritical
        Exit Sub
    End If
    If MsgBox("來源檔案分析出 " & iCnt & " 筆收件人信箱，是否繼續處理？", vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
    End If
    
    DTime = Format(ServerTime, "000000")
    tmpArr = Split(strToList, ";")
    iCnt = 0
    
    For intP = 0 To UBound(tmpArr)
         srKey = Trim(tmpArr(intP))
         If srKey <> "" Then
JumpReInput:
              '客戶檔
              'Modified by Lydia 2018/07/20 +F06
              strExc(0) = " SELECT 1 AS ORD1, 'CUSTOMER' AS TNAME, CU01||CU02 AS TPK, CU20 AS F01, CU115 AS F02, CU116 AS F03, CU117 AS F04, CU118 AS F05, '' AS F06" & _
                                " FROM CUSTOMER WHERE (INSTR(NLS_UPPER(CU20),'" & ChgSQL(UCase(srKey)) & "') > 0 OR  INSTR(NLS_UPPER(CU115),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(CU116),'" & ChgSQL(UCase(srKey)) & "') > 0" & _
                                " OR INSTR(NLS_UPPER(CU117),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(CU118),'" & ChgSQL(UCase(srKey)) & "') > 0 )"
              '國外潛在客戶檔
              'Modified by Lydia 2018/07/20 +F06
              strExc(0) = strExc(0) & " UNION ALL SELECT 2 AS ORD1, 'POTCUSTOMER' AS TNAME, PCU01||PCU02 AS TPK, PCU18 AS F01, '' AS F02, '' AS F03, '' AS F04, '' AS F05, '' AS F06" & _
                                " FROM POTCUSTOMER WHERE (INSTR(NLS_UPPER(PCU18),'" & ChgSQL(UCase(srKey)) & "') > 0 )"
              '國內潛在客戶檔
              'Modified by Lydia 2018/07/20 +F06
              strExc(0) = strExc(0) & " UNION ALL SELECT 3 AS ORD1, 'POTCUSTOMER1' AS TNAME, POC01||POC02 AS TPK, POC09 AS F01, '' AS F02, '' AS F03, '' AS F04, '' AS F05, '' AS F06" & _
                                " FROM POTCUSTOMER1 WHERE (INSTR(NLS_UPPER(POC09),'" & ChgSQL(UCase(srKey)) & "') > 0 )"
              '代理人檔
              'Modified by Lydia 2018/07/20 +FA105 AS F06
              'strExc(0) = strExc(0) & " UNION ALL SELECT 4 AS ORD1, 'FAGENT' AS TNAME, FA01||FA02 AS TPK, FA16 AS F01, FA79 AS F02, FA80 AS F03, FA81 AS F04, FA82 AS F05" & _
                                " FROM FAGENT WHERE (INSTR(NLS_UPPER(FA16),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(FA79),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(FA80),'" & ChgSQL(UCase(srKey)) & "') > 0" & _
                                " OR INSTR(NLS_UPPER(FA81),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(FA82),'" & ChgSQL(UCase(srKey)) & "') > 0 )"
              strExc(0) = strExc(0) & " UNION ALL SELECT 4 AS ORD1, 'FAGENT' AS TNAME, FA01||FA02 AS TPK, FA16 AS F01, FA79 AS F02, FA80 AS F03, FA81 AS F04, FA82 AS F05, FA105 AS F06" & _
                                " FROM FAGENT WHERE (INSTR(NLS_UPPER(FA16),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(FA79),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(FA80),'" & ChgSQL(UCase(srKey)) & "') > 0" & _
                                " OR INSTR(NLS_UPPER(FA81),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(FA82),'" & ChgSQL(UCase(srKey)) & "') > 0 OR INSTR(NLS_UPPER(FA105),'" & ChgSQL(UCase(srKey)) & "') > 0 )"
              
              '接洽人檔
              'Modified by Lydia 2018/07/20 +F06
              strExc(0) = strExc(0) & " UNION ALL SELECT 5 AS ORD1, 'POTCUSTCONT' AS TNAME, PCC01||'0-'||PCC02 AS TPK, PCC08 AS F01, '' AS F02, '' AS F03, '' AS F04, '' AS F05, '' AS F06" & _
                                " FROM POTCUSTCONT WHERE (INSTR(NLS_UPPER(PCC08),'" & ChgSQL(UCase(srKey)) & "') > 0 )"
              '開拓客戶資料檔
              'Modified by Lydia 2018/07/20 +F06
              strExc(0) = strExc(0) & " UNION ALL SELECT 6 AS ORD1, 'EXPANDCUSDETAIL' AS TNAME, ECD02||'-'||LPAD(ECD01,6,'0') AS TPK, ECD13 AS F01, '' AS F02, '' AS F03, '' AS F04, '' AS F05, '' AS F06" & _
                                " FROM EXPANDCUSDETAIL WHERE (INSTR(NLS_UPPER(ECD13),'" & ChgSQL(UCase(srKey)) & "') > 0 )"
              strExc(0) = strExc(0) & " ORDER BY ORD1,TPK"
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
              '假如遇到O8不能接受的字元會出現缺少右括號
              If intI = 2 Then
                  strExc(6) = InputBox("假如遇到O8不能接受的字元會出現遺漏右括號的訊息," & vbCrLf & "是否用?取代不可讀取的字元", "人工修改Email", srKey)
                  If strExc(6) <> "" And strExc(6) <> srKey Then
                      srKey = strExc(6)
                      GoTo JumpReInput
                  End If
              End If
              If intI = 0 Then '記錄未符合的Email信箱
JumpToAddRec:
                    '因為不同檔案有可會重覆收件人,判斷有同一天轉入的記錄EAE05=已更新
                    iCnt = iCnt + 1
                    strExc(5) = ""
                    strExc(0) = "select eae05 from EmailAddrErr where eae01='" & strSrvDate(1) & "' and NLS_UPPER(eae08)='" & ChgSQL(UCase(srKey)) & "' and nvl(eae05,'N') <> 'N' order by eae02,eae03 "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then strExc(5) = "已更新"
                    strSql = "insert into EmailAddrErr(EAE01,EAE02,EAE03,EAE04,EAE05,EAE06,EAE07,EAE08,EAE09,EAE10) " & _
                               "values ('" & strSrvDate(1) & "','" & DTime & "','" & Format(iCnt, "0000") & "','" & Text3.Text & "','" & strExc(5) & "' ,NULL,NULL,'" & ChgSQL(srKey) & "','" & strUserNum & "'," & Trim(Text4.Text) & ") "
                    cnnConnection.Execute strSql
              Else '記錄符合的Email信箱
                  With RsTemp
                        .MoveFirst
                        Do While Not .EOF
                             strMid = ""
                             strSql = ""
                             Erase strTemp
                             'Modified by Lydia 2018/07/20
                             'For intI = 1 To 5
                             For intI = 1 To 6
                                  '清除跳行符號和多餘空白
                                  strExc(1) = Trim(PUB_StringFilter("" & .Fields("F" & Format(intI, "00"))))
                                  ' 如果沒有";"號,則需完全相同
                                  tmpArr2 = Split(strExc(1), ";")
                                  strExc(2) = ""
                                  inX = 0
                                  For intJ = 0 To UBound(tmpArr2)
                                       strExc(3) = Trim(tmpArr2(intJ))
                                       If strExc(3) <> srKey Then
                                           strExc(2) = strExc(2) & IIf(strExc(2) <> "", ";" & tmpArr2(intJ), Trim(tmpArr2(intJ)))
                                       Else
                                           inX = intJ + 1
                                       End If
                                  Next intJ
                                  
                                  If strExc(1) <> "" And inX > 0 Then
                                       '更新欄位值
                                       Select Case "" & .Fields("ORD1")
                                             Case "1" '客戶檔
                                                    If intI = 1 Then strMid = strMid & ", cu20=" & CNULL(strExc(2)): strTemp(1) = "CU20"
                                                    If intI = 2 Then strMid = strMid & ", cu115=" & CNULL(strExc(2)): strTemp(2) = "CU115"
                                                    If intI = 3 Then strMid = strMid & ", cu116=" & CNULL(strExc(2)): strTemp(3) = "CU116"
                                                    If intI = 4 Then strMid = strMid & ", cu117=" & CNULL(strExc(2)): strTemp(4) = "CU117"
                                                    If intI = 5 Then strMid = strMid & ", cu118=" & CNULL(strExc(2)): strTemp(5) = "CU118"
                                             Case "2" '國外潛在客戶檔
                                                    If intI = 1 Then strMid = strMid & ", pcu18=" & CNULL(strExc(2)): strTemp(1) = "PCU18"
                                             Case "3" '國內潛在客戶檔
                                                    If intI = 1 Then strMid = strMid & ", poc09=" & CNULL(strExc(2)): strTemp(1) = "POC09"
                                             Case "4" '代理人檔
                                                    If intI = 1 Then strMid = strMid & ", fa16=" & CNULL(strExc(2)): strTemp(1) = "FA16"
                                                    If intI = 2 Then strMid = strMid & ", fa79=" & CNULL(strExc(2)): strTemp(2) = "FA79"
                                                    If intI = 3 Then strMid = strMid & ", fa80=" & CNULL(strExc(2)): strTemp(3) = "FA80"
                                                    If intI = 4 Then strMid = strMid & ", fa81=" & CNULL(strExc(2)): strTemp(4) = "FA81"
                                                    If intI = 5 Then strMid = strMid & ", fa82=" & CNULL(strExc(2)): strTemp(5) = "FA82"
                                                    If intI = 6 Then strMid = strMid & ", fa105=" & CNULL(strExc(2)): strTemp(6) = "FA105"
                                             Case "5" '接洽人檔
                                                    If intI = 1 Then strMid = strMid & ", pcc08=" & CNULL(strExc(2)): strTemp(1) = "PCC08"
                                             Case "6" '開拓客戶資料檔
                                                    If intI = 1 Then strMid = strMid & ", ecd13=" & CNULL(strExc(2)): strTemp(1) = "ECD13"
                                       End Select
                                  End If
                             Next intI
                             
                             '更新&新增處理記錄
                             If strMid = "" Then GoTo JumpToAddRec  'Added by Lydia 2018/01/18 如果沒有";"號,則需完全相同
                             
On Error GoTo ErrHandle18
                             If strMid <> "" Then
                                  cnnConnection.BeginTrans
                                       '更新欄位
                                       Select Case "" & .Fields("ORD1")
                                             Case "1" '客戶檔
                                                    'Modified by Lydia 2018/02/05 加Update(cu84,cu85,cu86)
                                                    strSql = "update customer set cu84='QPGMR', cu85=" & strSrvDate(1) & ", cu86=" & Mid(DTime, 1, 4) & " " & strMid & _
                                                                 " where cu01='" & Mid(.Fields("TPK"), 1, 8) & "' and cu02='" & Mid(.Fields("TPK"), 9, 1) & "' "
                                             Case "2" '國外潛在客戶檔
                                                     'Modified by Lydia 2018/02/05 加Update(pcu44,pcu45,pcu46)
                                                    strSql = "update potcustomer set pcu44='QPGMR', pcu45=" & strSrvDate(1) & ", pcu46=" & Mid(DTime, 1, 4) & " " & strMid & _
                                                                 " where pcu01='" & Mid(.Fields("TPK"), 1, 8) & "' and pcu02='" & Mid(.Fields("TPK"), 9, 1) & "' "
                                             Case "3" '國內潛在客戶檔
                                                     'Modified by Lydia 2018/02/05 加Update(poc20,poc21,poc22)
                                                    strSql = "update potcustomer1 set poc20='QPGMR', poc21=" & strSrvDate(1) & ", poc22=" & Mid(DTime, 1, 4) & " " & strMid & _
                                                                " where poc01='" & Mid(.Fields("TPK"), 1, 8) & "' and poc02='" & Mid(.Fields("TPK"), 9, 1) & "' "
                                             Case "4" '代理人檔
                                                     'Modified by Lydia 2018/02/05 加Update(fa49,fa50,fa51)
                                                    strSql = "update fagent set fa49='QPGMR', fa50=" & strSrvDate(1) & ", fa51=" & Mid(DTime, 1, 4) & " " & strMid & _
                                                                 " where fa01='" & Mid(.Fields("TPK"), 1, 8) & "' and fa02='" & Mid(.Fields("TPK"), 9, 1) & "' "
                                             Case "5" '接洽人檔
                                                     'Modified by Lydia 2018/02/05 加Update(pcc17,pcc18,pcc19)
                                                    strSql = "update potcustcont set pcc17='QPGMR', pcc18=" & strSrvDate(1) & ", pcc19=" & Mid(DTime, 1, 4) & " " & strMid & _
                                                                 " where pcc01='" & Mid(.Fields("TPK"), 1, 8) & "' and pcc02='" & Mid(.Fields("TPK"), 11, 2) & "' "
                                             Case "6" '開拓客戶資料檔
                                                    strSql = "update expandcusdetail set " & Mid(strMid, 2) & _
                                                                " where ecd02='" & Mid(.Fields("TPK"), 1, InStr(.Fields("TPK"), "-") - 1) & "' and ecd01=" & Val(Mid(.Fields("TPK"), InStr(.Fields("TPK"), "-") + 1))
                                       End Select
                                       Pub_SeekTbLog strSql, "QPGMR" 'Added by Lydia 2018/02/05 加修改記錄
                                       cnnConnection.Execute strSql, intI
                                       '(因為update 語法若有誤,若先新增log則無法中斷)
                                       If intI > 0 Then
                                            intU = intU + 1
                                            '新增處理記錄
                                             For intI = 1 To 5
                                                  If "" & strTemp(intI) <> "" Then
                                                        iCnt = iCnt + 1
                                                        strSql = "insert into EmailAddrErr(EAE01,EAE02,EAE03,EAE04,EAE05,EAE06,EAE07,EAE08,EAE09,EAE10) " & _
                                                                   "values ('" & strSrvDate(1) & "','" & DTime & "','" & Format(iCnt, "0000") & "','" & Text3.Text & "','" & .Fields("TNAME") & "','" & .Fields("TPK") & "','" & strTemp(intI) & "','" & ChgSQL(srKey) & "','" & strUserNum & "'," & Trim(Text4.Text) & ") "
                                                         cnnConnection.Execute strSql
                                                  End If
                                             Next
                                        End If
                                  cnnConnection.CommitTrans
                             End If
                             .MoveNext
                        Loop
                  End With
              End If
         End If
    Next intP
    MsgBox "大宗Email退件處理完成，共 " & iCnt & " 筆處理記錄!", vbInformation
    Exit Sub
ErrHandle18:
    If DTime <> "" And strSql <> "" Then
        cnnConnection.RollbackTrans
    End If
End Sub

Private Sub Command17_Click()
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.txt"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.TXT)|*.TXT"
      .InitDir = PUB_Getdesktop
      '單選
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         Text2.Text = .FileName
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Text2_GotFocus()
     TextInverse Text2
End Sub

Private Sub Text29_Change(Index As Integer)
   If Index = 0 Then
      If Len(Text29(Index)) = 5 Then
         Label29(3) = GetStaffName(Text29(0))
      Else
         Label29(3) = ""
      End If
   End If
End Sub

Private Sub Text29_GotFocus(Index As Integer)
   CloseIme
   TextInverse Text29(Index)
End Sub

Private Sub Text29_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = UpperCase(KeyAscii)
   ElseIf KeyAscii <> 8 And (KeyAscii > 57 Or KeyAscii < 48) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text3_GotFocus()
     CloseIme
     TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
     KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    lblName.Caption = ""
    If Trim(Text3.Text) = "" Then Exit Sub
    If ClsPDGetStaff(Text3.Text, strExc(1)) = True Then
        lblName.Caption = strExc(1)
    End If
End Sub

Private Sub Text4_GotFocus()
    TextInverse Text4
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
    If Text4.Text = "" Or CheckIsDate(Text4.Text) = False Then
         Text4.SetFocus
         Text4_GotFocus
    End If
End Sub
'end 2018/01/22

'Added by Lydia 2018/02/02 補2018/2/2的退信處理的dml_log和UpdateID
Private Sub Command19_Click()
Dim intA As Integer
Dim rsAD As New ADODB.Recordset
Dim strMid As String

On Error GoTo ErrHandler
    strExc(0) = "select * from emailaddrerr where EAE01=20180202  and eae06 is not null ORDER BY EAE01,EAE02,EAE03 "
    intI = 1
    Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
         With rsAD
             cnnConnection.BeginTrans
             
             Do While Not .EOF
                    strExc(0) = "":   strExc(1) = "": strExc(2) = "": strSql = ""
                    Select Case "" & .Fields("EAE05")
                          Case "CUSTOMER" '客戶檔
                                 strExc(0) = "SELECT " & .Fields("EAE07") & " ,CU84,CU85,CU86 FROM " & .Fields("EAE05") & " WHERE CU01||CU02='" & .Fields("EAE06") & "' "
                                 strSql = "UPDATE CUSTOMER SET CU84='QPGMR', CU85=" & .Fields("EAE01") & ", CU86=" & Mid(.Fields("EAE02"), 1, 4) & " WHERE CU01||CU02='" & .Fields("EAE06") & "' "
                          Case "POTCUSTOMER" '國外潛在客戶檔
                                 strExc(0) = "SELECT " & .Fields("EAE07") & ",PCU44,PCU45,PCU46 FROM " & .Fields("EAE05") & " WHERE PCU01||PCU02='" & .Fields("EAE06") & "' "
                                 strSql = "UPDATE POTCUSTOMER SET PCU44='QPGMR', PCU45=" & .Fields("EAE01") & ", PCU46=" & Mid(.Fields("EAE02"), 1, 4) & " WHERE PCU01||PCU02='" & .Fields("EAE06") & "' "
                          Case "POTCUSTOMER1" '國內潛在客戶檔
                                 strExc(0) = "SELECT " & .Fields("EAE07") & ",POC20,POC21,POC22 FROM " & .Fields("EAE05") & " WHERE POC01||POC02='" & .Fields("EAE06") & "' "
                                 strSql = "UPDATE POTCUSTOMER1 SET POC20='QPGMR', POC21=" & .Fields("EAE01") & ", POC22=" & Mid(.Fields("EAE02"), 1, 4) & " WHERE POC01||POC02='" & .Fields("EAE06") & "' "
                          Case "FAGENT" '代理人檔
                                 strExc(0) = "SELECT " & .Fields("EAE07") & ",FA49,FA50,FA51 FROM " & .Fields("EAE05") & " WHERE FA01||FA02='" & .Fields("EAE06") & "' "
                                 strSql = "UPDATE FAGENT SET FA49='QPGMR', FA50=" & .Fields("EAE01") & ", FA51=" & Mid(.Fields("EAE02"), 1, 4) & " WHERE FA01||FA02='" & .Fields("EAE06") & "' "
                          Case "POTCUSTCONT" '接洽人檔
                                 strExc(0) = "SELECT " & .Fields("EAE07") & ",PCC17,PCC18,PCC19 FROM " & .Fields("EAE05") & " WHERE PCC01||'0-'||PCC02='" & .Fields("EAE06") & "' "
                                 strSql = "UPDATE POTCUSTCONT SET PCC17='QPGMR', PCC18=" & .Fields("EAE01") & ", PCC19=" & Mid(.Fields("EAE02"), 1, 4) & " WHERE PCC01||'0-'||PCC02='" & .Fields("EAE06") & "' "
                          Case "EXPANDCUSDETAIL" '開拓客戶資料檔
                                 strExc(0) = "SELECT " & .Fields("EAE07") & ",'' n01,'' n02,'' n03 FROM " & .Fields("EAE05") & " WHERE ECD02||'-'||LPAD(ECD01,6,'0')='" & .Fields("EAE06") & "' "
                                 strSql = ""
                    End Select
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    '抓現有資料+去掉的Email
                    strMid = ""
                    If intI = 1 Then
                         If Trim("" & RsTemp(0)) <> "" Then
                             strMid = "" & RsTemp(0) & ";" & .Fields("EAE08")
                             strMid = strMid & "=>" & RsTemp(0)
                         Else
                             strMid = .Fields("EAE08") & "=>NULL"
                         End If
                            Select Case "" & .Fields("EAE05")
                                  Case "CUSTOMER" '客戶檔
                                         strExc(1) = "修改 customer ，條件 cu01=>" & Mid(.Fields("EAE06"), 1, 8) & ", cu02=>" & Mid(.Fields("EAE06"), 9, 1) & _
                                                          " 資料；cu84[" & RsTemp(1) & "=>QPGMR]; cu85[" & RsTemp(2) & "=>" & .Fields("eae01") & "];cu86[" & RsTemp(3) & "=>" & Mid(.Fields("eae02"), 1, 4) & " ];" & .Fields("eae07") & "[abcdefgh]"
                                  Case "POTCUSTOMER" '國外潛在客戶檔
                                         strExc(1) = "修改 potcustomer ，條件 pcu01=>" & Mid(.Fields("EAE06"), 1, 8) & ", pcu02=>" & Mid(.Fields("EAE06"), 9, 1) & _
                                                          " 資料；pcu44[" & RsTemp(1) & "=>QPGMR]; pcu45[" & RsTemp(2) & "=>" & .Fields("eae01") & "];pcu46[" & RsTemp(3) & "=>" & Mid(.Fields("eae02"), 1, 4) & " ];" & .Fields("eae07") & "[abcdefgh]"
                                  Case "POTCUSTOMER1" '國內潛在客戶檔
                                         strExc(1) = "修改 potcustomer1 ，條件 poc01=>" & Mid(.Fields("EAE06"), 1, 8) & ", poc02=>" & Mid(.Fields("EAE06"), 9, 1) & _
                                                          " 資料；poc20[" & RsTemp(1) & "=>QPGMR]; poc21[" & RsTemp(2) & "=>" & .Fields("eae01") & "];poc22[" & RsTemp(3) & "=>" & Mid(.Fields("eae02"), 1, 4) & " ];" & .Fields("eae07") & "[abcdefgh]"
                                  Case "FAGENT" '代理人檔
                                         strExc(1) = "修改 fagent ，條件 fa01=>" & Mid(.Fields("EAE06"), 1, 8) & ", fa02=>" & Mid(.Fields("EAE06"), 9, 1) & _
                                                           " 資料；fa49[" & RsTemp(1) & "=>QPGMR]; fa50[" & RsTemp(2) & "=>" & .Fields("eae01") & "];fa51[" & RsTemp(3) & "=>" & Mid(.Fields("eae02"), 1, 4) & " ];" & .Fields("eae07") & "[abcdefgh]"
                                  Case "POTCUSTCONT" '接洽人檔
                                         strExc(1) = "修改 potcustcont ，條件 pcc01=>" & Mid(.Fields("EAE06"), 1, 8) & ", pcc02=>" & Mid(.Fields("EAE06"), 11, 2) & _
                                                          " 資料；pcc17[" & RsTemp(1) & "=>QPGMR]; pcc18[" & RsTemp(2) & "=>" & .Fields("eae01") & "];pcc19[" & RsTemp(3) & "=>" & Mid(.Fields("eae02"), 1, 4) & " ];" & .Fields("eae07") & "[abcdefgh]"
                                  Case "EXPANDCUSDETAIL" '開拓客戶資料檔
                                         strExc(1) = "修改 expadcusdetail ，條件 ecd02=>" & Mid(.Fields("EAE06"), 1, InStr(.Fields("EAE06"), "-") - 1) & ", ecd01=>" & Val(Mid(.Fields("EAE06"), InStr(.Fields("EAE06"), "-") + 1)) & _
                                                          " 資料；" & .Fields("eae07") & "[abcdefgh]"
                            End Select
                         strExc(2) = Replace(strExc(1), "abcdefgh", strMid)
                         strExc(1) = "INSERT INTO DML_LOG (DL06,DL07,DL08,DL09,DL10,DL11,DL12) " & _
                                           "VALUES ('QPGMR', " & .Fields("eae01") & ", " & .Fields("eae02") & ", '" & ChgSQL(strExc(2)) & "', '" & .Fields("eae05") & "', '1', '維護作業1(FRM000001)') "
                    End If
                    cnnConnection.Execute strExc(1)
                    If strSql <> "" Then cnnConnection.Execute strSql
                    
                  .MoveNext
             Loop
             cnnConnection.CommitTrans
         End With
    End If
    
    MsgBox "END!!!", vbInformation
Exit Sub

ErrHandler:
    If Err.Number <> 0 Then
        cnnConnection.RollbackTrans
        MsgBox Err.Description
    End If
End Sub

'Added by Lydia 2018/03/08 107/3/5 後 FCP中說進度發文 批次刪除english_vers案號資料夾的*.msg
'Remove by Lydia 2021/12/09 已不使用
'Private Sub Command20_Click()
'Dim intA As Integer
'Dim rsAD As New ADODB.Recordset
'Dim strList As String
'Dim strToPath As String
'Dim stFtpIP  As String
'Dim strToDir As String
'Dim f91 As Integer
'Dim TFName As String
'Dim intC As Integer
'Dim strDate As String
'Dim tmpArr As Variant
'
'On Error GoTo ErrHandler
'
'    If MsgBox("確定執行" & Command40.Caption & " ？", vbInformation + vbYesNo + vbDefaultButton2, "確定執行") = vbNo Then
'        Exit Sub
'    End If
'
'   TFName = "FCP中說發文批次處理案號資料夾"
'
'   '轉FTP
'   stFtpIP = Pub_GetSpecMan("FTP_TYPING2")
'   If stFtpIP = "" Then Exit Sub
'
'  strDate = strSrvDate(1) '要處理的發文日
'
'  '處理案件命名追蹤TrackingCaseName的TCN12改成記錄msg檔數量
'  strSql = "update trackingcasename set tcn12='1' where tcn12='Y' "
'  cnnConnection.Execute strSql, intI
'  strExc(0) = "select tcn01,tcn05,tcn12 from trackingcasename where tcn07>=20180305 "
'   intA = 1
'   Set rsAD = ClsLawReadRstMsg(intA, strExc(0))
'   If intA = 1 Then
'        rsAD.MoveFirst
'        Do While Not rsAD.EOF
'            strExc(1) = Replace(FCP命名追蹤暫存, "\", "/") & "/" & Val(rsAD.Fields("tcn01"))
'            strExc(2) = "//" & Mid(strExc(1), InStr(3, strExc(1), "/") + 1)
'            If "" & rsAD.Fields("tcn05") <> "" Then
'                  '刪除不該存在的資料夾
'                  If PUB_ChkFtpDirectory(stFtpIP, strExc(2)) = True Then
'                         If PUB_FtpDelFile(strExc(2), , , , stFtpIP) = False Then
'                         End If
'                  End If
'            ElseIf Val("" & rsAD.Fields("tcn12")) = 1 Then '記錄未收文的命名追蹤msg檔數量
'                  strList = ""
'                  intC = 0
'                  If PUB_ChkFtpDirectory(stFtpIP, strExc(2), "R", ".msg", strList) = True Then
'                      If strList <> "" Then
'                          tmpArr = Empty
'                          tmpArr = Split(strList, "&")
'                          For intI = 0 To UBound(tmpArr)
'                               If Trim(tmpArr(intI)) <> "" Then
'                                   intC = intC + 1
'                               End If
'                          Next intI
'                          If intC > 1 Then
'                               strSql = "update trackingcasename set tcn12='" & intC & "' where tcn01=" & CNULL(rsAD.Fields("tcn01"))
'                               cnnConnection.Execute strSql, intI
'                          End If
'                      End If
'                  End If
'            End If
'             rsAD.MoveNext
'        Loop
'   End If
'   intC = 0
'   'end 2018/03/09
'
'   strExc(0) = "select cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp158 from caseprogress " & _
'                     "where cp01='FCP' and cp05>=20180305 and cp158=" & strDate & " and cp159=0 and cp10 in (201,209,210,235) " & _
'                     "and exists (select * from caseprogress where cp01='FCP' and cp10 in (101,102,103) and cp05>=20180305) " & _
'                     "order by cp05,cp09 "
'    intA = 1
'    Set rsAD = ClsLawReadRstMsg(intA, strExc(0))
'    If intA = 1 Then
'         With rsAD
'              .MoveFirst
'              f91 = FreeFile
'              Open App.path & "\" & TFName & ".txt" For Output As f91
'              Print #f91, Space(40) & strSrvDate(2) & "_" & TFName
'              Print #f91, ""
'              Print #f91, convForm("本所案號", 15) & " " & "處理記錄"
'              Print #f91, String(15, "=") & " " & String(124, "=")
'              Do While Not .EOF
'                   strList = ""
'                    '測試路徑
'                    'strToDir = "//English_Vers/Test/" & Val(.Fields("cp02"))
'                    strExc(1) = Replace(Pub_GetFCPcaseFilePath("" & .Fields("cp02")), "\", "/")
'                    strToDir = "//" & Mid(strExc(1), InStr(3, strExc(1), "/") + 1)
'                   If PUB_ChkFtpDirectory(stFtpIP, strToDir, "R", "*.*", strList) Then
'                        '測試路徑
'                        'strList = strToDir & "資料夾存在，讀取：" & strList
'                        strList = Pub_GetFCPcaseFilePath("" & .Fields("cp02")) & "資料夾存在，刪除：" & strList
'                   Else
'                        '測試路徑
'                        'strList = strToDir & "資料夾不存在"
'                        strList = Pub_GetFCPcaseFilePath("" & .Fields("cp02")) & "資料夾不存在"
'                   End If
'                   Print #f91, convForm("" & .Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04"), 15) & " " & Replace(strList, "&", " ；")
'                   intC = intC + 1
'                   .MoveNext
'              Loop
'         End With
'         Print #f91, String(140, "=")
'         Print #f91, "共" & intC & "筆"
'    End If
'
'    If f91 > 0 Then
'        Close f91
'        PUB_SendMail strUserNum, "A3034", "", strSrvDate(2) & "_" & TFName, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, , App.path & "\" & TFName & ".txt"
'    End If
'
'    Set rsAD = Nothing
'    MsgBox "End !!"
'    Exit Sub
'ErrHandler:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description
'    End If
'End Sub

'Added by Lydia 2018/04/02 刪除FC撰寫信函上傳到卷宗區的檔案
Private Sub Command21_Click()
Dim intQ As Integer
Dim rsA1 As New ADODB.Recordset
Dim f21 As Integer
Dim TempName As String

    If MsgBox("確定執行" & Command40.Caption & " ？", vbInformation + vbYesNo + vbDefaultButton2, "確定執行") = vbNo Then
        Exit Sub
    End If
    
   strSql = "SELECT cp01,cp02,cp03,cp04,cpp01,cpp02,cpp14 " & _
               "FROM CASEPAPERPDF,CASEPROGRESS " & _
               "WHERE cpp01=cp09(+) and cp01='FCP' and cp10='201' and cpp05 in ('86013','73023') order by cpp06,cpp07 " 'Test: and cp05<=20151105
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   TempName = "刪除FC撰寫信函上傳到卷宗區的檔案記錄"
   If intI = 1 Then
       If MsgBox("共有" & RsTemp.RecordCount & "筆資料，是否要刪除？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
           Exit Sub
       End If
       intQ = 0
       f21 = FreeFile
       Open App.path & "\" & TempName & ".txt" For Output As f21
       Print #f21, Space(40) & TempName
       Print #f21, ""
       Print #f21, convForm("本所案號", 16) & " " & convForm("收文號 ", 9) & " " & convForm("檔名", 100) & " " & convForm("FTP路徑", 200)
       Print #f21, String(16, "=") & " " & String(9, "=") & " " & String(100, "=") & " " & String(200, "=")
       RsTemp.MoveFirst
       Do While Not RsTemp.EOF
            If "" & RsTemp.Fields("cpp01") <> "" And "" & RsTemp.Fields("cpp02") <> "" Then
                If PUB_DelFtpFile2("" & RsTemp.Fields("cpp01"), " and CPP02=" & CNULL("" & RsTemp.Fields("cpp02"))) = True Then
                     strSql = "delete from casepaperpdf where cpp01=" & CNULL(RsTemp.Fields("cpp01")) & " and cpp02=" & CNULL(RsTemp.Fields("cpp02"))
                     cnnConnection.Execute strSql, intI
                     intQ = intQ + 1
                     Print #f21, RsTemp.Fields("cp01") & "-" & RsTemp.Fields("cp02") & "-" & RsTemp.Fields("cp03") & "-" & RsTemp.Fields("cp04") & "  " & _
                                     RsTemp.Fields("cpp01") & " " & convForm(RsTemp.Fields("cpp02"), 100) & " " & convForm(RsTemp.Fields("cpp14"), 100)
                Else
                     Print #f21, "*" & RsTemp.Fields("cp01") & "-" & RsTemp.Fields("cp02") & "-" & RsTemp.Fields("cp03") & "-" & RsTemp.Fields("cp04") & " " & _
                                     RsTemp.Fields("cpp01") & " " & convForm(RsTemp.Fields("cpp02"), 100) & " " & convForm(RsTemp.Fields("cpp14"), 100)
                End If
            End If
            RsTemp.MoveNext
       Loop
       Print #f21, String(200, "=")
       Print #f21, "共" & intQ & "筆"
   End If
    
   If f21 > 0 Then
        Close f21
        '取消發信
        'PUB_SendMail strUserNum, "A3034", "", TempName, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, , App.path & "\" & TempName & ".txt"
   End If
   MsgBox "End !!", vbInformation
End Sub

'Added by Lydia 2018/04/23 更新命名-告代,主動修正901,203承辦期限
Private Sub Command22_Click()
Dim intQ As Integer
Dim rsA1 As New ADODB.Recordset
    
On Error GoTo ErrHand

    If MsgBox("確定執行" & Command40.Caption & " ？", vbInformation + vbYesNo + vbDefaultButton2, "確定執行") = vbNo Then
        Exit Sub
    End If
    
    strSql = "select cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp27,cp48,cp64 from caseprogress,staff " & _
                 "where cp05>=20180305 and cp10 in (203,901) and substr(cp09,1,1) = 'B' and cp158=0 and cp159=0 and cp65=st01(+) " & _
                 "and st03='F21' and cp48 is null order by cp05 "
    intI = 1
    Set rsA1 = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
         rsA1.MoveFirst
cnnConnection.BeginTrans
         Do While Not rsA1.EOF
              strSql = "select cp09,cp10,cp158 from caseprogress where cp01='" & rsA1.Fields("cp01") & "' and cp02='" & rsA1.Fields("cp02") & "' and cp03='" & rsA1.Fields("cp03") & "' and cp04='" & rsA1.Fields("cp04") & "' and cp31='Y' "
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
              If intI = 1 Then
                  strExc(1) = "" & RsTemp.Fields("cp09") '新申請案收文號
                  If Val("" & RsTemp.Fields("cp158")) > 0 Then
                        strExc(2) = ""
                        strExc(3) = "" & RsTemp.Fields("cp158")
                        '告代：提申後新案發文日起算6個工作天
                        If "" & rsA1.Fields("cp10") = "901" Then
                            strExc(2) = Pub_GetHandleDay("FCP", "000", "901", "" & RsTemp.Fields("cp158"))
                        Else
                        '主動修正：提申後+新案翻譯未發文=新案翻譯的本所期限；
                            strSql = "select cp09,cp06,cp158 from caseprogress where cp01='" & rsA1.Fields("cp01") & "' and cp02='" & rsA1.Fields("cp02") & "' and cp03='" & rsA1.Fields("cp03") & "' and cp04='" & rsA1.Fields("cp04") & "' and cp10 in (201,209,235,210)  and cp159=0 "
                            intI = 1
                            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                            If intI = 1 Then
                                strExc(2) = "" & RsTemp.Fields("cp06")
                                If Val(strExc(3)) = Val("" & RsTemp.Fields("cp158")) Then
                                    strExc(2) = Pub_GetHandleDay("FCP", "000", "203", "" & RsTemp.Fields("cp158"))
                                End If
                            End If
                        End If
                        strSql = "update caseprogress set cp48=" & strExc(2) & " , cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 整批更新承辦期限為提申後;'||cp64 where cp09='" & rsA1.Fields("cp09") & "' "
                        cnnConnection.Execute strSql, intI
                        strSql = "update transcasetitle set " & IIf("" & rsA1.Fields("cp10") = "901", " tct20='1' ", " tct117='1' ") & " where tct01='" & strExc(1) & "' "
                        cnnConnection.Execute strSql, intI
                  '新申請案未發文
                  Else
                        strExc(2) = Pub_GetHandleDay("FCP", "000", "" & rsA1.Fields("cp10"), "" & rsA1.Fields("cp05"))
                        strSql = "update caseprogress set cp48=" & strExc(2) & " , cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 整批更新承辦期限為提申前;'||cp64 where cp09='" & rsA1.Fields("cp09") & "' "
                        cnnConnection.Execute strSql, intI
                        strSql = "update transcasetitle set " & IIf("" & rsA1.Fields("cp10") = "901", " tct20='2' ", " tct117='2' ") & " where tct01='" & strExc(1) & "' "
                        cnnConnection.Execute strSql, intI
                  End If
              End If
              rsA1.MoveNext
         Loop
cnnConnection.CommitTrans
    End If
    MsgBox "End !!"
    Exit Sub
    
ErrHand:
    cnnConnection.RollbackTrans
    
End Sub

'Add by Amy 2018/05/04 取代造字
Private Function GetZIPSpecWord(strZSW01 As String, strZSW As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    strQ = "Select zsw05 From ZIPSpecWord " & _
                "Where zsw01='" & strZSW01 & "' And zsw02||zsw03||zsw04='" & strZSW & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetZIPSpecWord = "" & RsQ.Fields("zsw05")
    End If
    RsQ.Close
End Function

Private Sub txtCaseNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2018/09/14
Private Sub Command27_Click()
Dim oFileSys As New FileSystemObject
Dim oFile
Dim strTempPath As String
Dim strTempName As String
Dim strNo As String
Dim ii As Integer, jj As Integer, iCnt As Integer
Dim bolTest As Boolean
Dim strProcList As String
'1.讀取指定路徑的檔案名稱，分析出回覆的是代理人編號、客戶編號或潛在客戶編號以及聯絡人編號，更新基本檔的GDPR欄位。
'2.將處理完的檔案人工移到下一層資料夾。
'GDPR回函回寫
'1.編號: 抓(())或()內編號
'2.若編號最後一字為+則跑下列語法抓其他編號
'select msd06 From mailscheduledetail a
' where msd01=511 and msd02=(select b.msd02
' from mailscheduledetail b where b.msd01=a.msd01
' and b.msd06='編號') and msd03=19221111
'3.回寫: Y >> FA123, R >> PCU50, R- >> PCC26
    
'P.S. R14812000-01+ 特別處理,沒找到不用更新
ii = MsgBox("是否測試SQL語法正確？" & vbCrLf & "是：顯示語法在即時運算視窗" & vbCrLf & "否：更新DB" & vbCrLf & "取消：放棄作業", vbInformation + vbYesNoCancel + vbDefaultButton3)
If ii = 2 Then 'Cancel
     Exit Sub
ElseIf ii = 7 Then 'No
      bolTest = False
      '檢查語法
        'select count(*) from fagent where fa123 in ('Y','N');
        'select count(*) from potcustomer where pcu50 in ('Y','N');
        'select count(*) from potcustcont where pcc26 in ('Y','N');
        'select seqno,rowseq,r001,r002 from rdatafactory where formname='frm000001' and id='A3034' and seqno>=270 order by r001,rowseq desc;
      cnnConnection.Execute "delete from rdatafactory where formname=" & CNULL(Me.Name) & " and id = " & CNULL(strUserNum)
Else
      bolTest = True
End If

    For ii = 0 To 1
         strTempPath = txtGDPR(ii).Text
         iCnt = 0
         strProcList = ""
         If bolTest = True Then Debug.Print String(20, "=")
         strTempName = Dir(strTempPath & "\*.msg")
         If strTempName = "" Then
              MsgBox Label6(ii).Caption & "資料夾無msg檔案!"
              GoTo JumpNext
         End If
         Do While strTempName <> ""
              If iCnt = 0 Then
                    If bolTest = False Then
                        cnnConnection.BeginTrans
                    End If
              End If
              strNo = ""
              If InStr(strTempName, "(") > 0 Then
                  strExc(1) = UCase(Mid(strTempName, InStr(strTempName, "(") + 1))
                  For jj = 1 To Len(strExc(1))
                      strExc(2) = Mid(strExc(1), jj, 1)
                      strExc(3) = Mid(strExc(1), jj + 1, 1)
                      If (strExc(2) = "X" Or strExc(2) = "Y" Or strExc(2) = "R") And InStr("0123456789", strExc(3)) > 0 And strExc(3) <> "" Then
                           strExc(1) = Mid(strExc(1), jj, InStr(jj, strExc(1), ")"))
                           If InStr(strExc(1), ")") > 0 Then
                              strExc(1) = Mid(strExc(1), 1, InStr(jj, strExc(1), ")") - 1)
                           End If
                           If InStr(strProcList, strExc(1)) > 0 Then
                               'Debug.Print strExc(1) & "->" & strTempName '檢查
                               strNo = ""
                           Else
                               strNo = strExc(1)
                               strProcList = strProcList & "," & strNo
                               iCnt = iCnt + 1
                           End If
                           Exit For
                      End If
                  Next jj
              End If
              If strNo <> "" Then
                    Select Case Left(strNo, 1)
                         Case "Y"
                               strSql = "update fagent set FA123='" & IIf(ii = 0, "Y", "N") & "' where FA01||FA02='" & Left(strNo, 9) & "' "
                               strExc(5) = Left(strNo, 9)
                         Case "R"
                               If Mid(strNo, 10, 1) = "-" Then
                                   strSql = "update PotCustCont set PCC26='" & IIf(ii = 0, "Y", "N") & "' where PCC01||'0-'||PCC02='" & Left(strNo, 12) & "' "
                                   strExc(5) = Left(strNo, 12)
                               Else
                                   strSql = "update PotCustomer set PCU50='" & IIf(ii = 0, "Y", "N") & "' where PCU01||PCU02='" & Left(strNo, 9) & "' "
                                   strExc(5) = Left(strNo, 9)
                               End If
                    End Select
                    If bolTest = True Then
                         Debug.Print strSql
                    Else
                         cnnConnection.Execute strSql, intI
                         cnnConnection.Execute "insert into rdatafactory (formname,id,seqno,rowseq,r001,r002) values (" & CNULL(Me.Name) & ", " & CNULL(strUserNum) & ", " & "27" & ii & ", " & iCnt & ", " & CNULL(strNo) & ", " & CNULL(ChgSQL(strSql)) & ") "
                    End If
                    '最後一字為+則跑下列語法抓其他編號
                    If Right(strNo, 1) = "+" Then
                         strExc(0) = "select msd06 From mailscheduledetail a " & _
                                           "where msd01=511 and msd02=(select b.msd02 " & _
                                           "from mailscheduledetail b where b.msd01=a.msd01 " & _
                                           "and b.msd06='" & strNo & "') and msd03=19221111 "
                         intI = 1
                         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                         If intI = 1 Then
                             With RsTemp
                                   .MoveFirst
                                   Do While Not .EOF
                                        Select Case Left("" & .Fields(0), 1)
                                             Case "Y"
                                                   strSql = "update fagent set FA123='" & IIf(ii = 0, "Y", "N") & "' where FA01||FA02='" & Left("" & .Fields(0), 9) & "' "
                                                   strExc(5) = Left("" & .Fields(0), 9)
                                             Case "R"
                                                   If Mid("" & .Fields(0), 10, 1) = "-" Then
                                                       strSql = "update PotCustCont set PCC26='" & IIf(ii = 0, "Y", "N") & "' where PCC01||'0-'||PCC02='" & Left("" & .Fields(0), 12) & "' "
                                                       strExc(5) = Left("" & .Fields(0), 12)
                                                   Else
                                                       strSql = "update PotCustomer set PCU50='" & IIf(ii = 0, "Y", "N") & "' where PCU01||PCU02='" & Left("" & .Fields(0), 9) & "' "
                                                       strExc(5) = Left("" & .Fields(0), 9)
                                                   End If
                                        End Select
                                        If bolTest = True Then
                                             Debug.Print "+SQL: " & strSql
                                        Else
                                             cnnConnection.Execute strSql, intI
                                             cnnConnection.Execute "insert into rdatafactory (formname,id,seqno,rowseq,r001,r002) values (" & CNULL(Me.Name) & ", " & CNULL(strUserNum) & ", " & "27" & ii & ", " & iCnt & ", " & CNULL(.Fields(0)) & ", " & CNULL(ChgSQL(strSql)) & ") "
                                        End If
                                        strProcList = strProcList & "," & .Fields(0)
                                        iCnt = iCnt + 1
                                        .MoveNext
                                   Loop
                             End With
                         End If
                    End If '最後一字為+
              End If
              strTempName = Dir()
         Loop
         If bolTest = True Then
            Debug.Print String(20, "=")
            Debug.Print "Record: " & iCnt
            Debug.Print "No List: " & strProcList
         Else
             cnnConnection.CommitTrans
         End If
JumpNext:
    Next ii
    
    MsgBox "執行完畢！", vbInformation
    Exit Sub
    
ErrHand27:
If Err.Number <> 0 Then
    MsgBox Err.Description
    If bolTest = False Then
         cnnConnection.RollbackTrans
    End If
End If
End Sub

'Added by Lydia 2018/09/27 CFT補催審期限
Private Sub Command28_Click()
   Dim stSQL As String
   Dim rsA1 As New ADODB.Recordset
   Dim stDate1 As String
   Dim intC As Integer
   
    If MsgBox("確定執行" & Command40.Caption & " ？", vbInformation + vbYesNo + vbDefaultButton2, "確定執行") = vbNo Then
        Exit Sub
    End If
    
   'Modified by Lydia 2018/12/05
   'stSQL = "select * from a113 order by cp27 "
   stSQL = "select a.*,tm10 from a113 a, trademark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) order by cp27 "
   intI = 1
   Set rsA1 = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
        If MsgBox("是否新增催審期限？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
             Exit Sub
        End If
        cnnConnection.BeginTrans
        With rsA1
             .MoveFirst
             Do While Not .EOF
                  stDate1 = CompDate(2, Val("" & .Fields("cf05")), "" & .Fields("cp27"))
                  'Modified by Lydia 2018/12/05 改成系統日+1天
                  'If stDate1 < strSrvDate(1) Then  '期限若<系統日改設為系統日
                  '    stDate1 = strSrvDate(1)
                  strExc(1) = CompWorkDay(2, strSrvDate(1))
                  If stDate1 < strExc(1) Then
                      stDate1 = strExc(1)
                  'end 2018/12/05
                  End If
                  'Modified by Lydia 2018/12/05 改備註(整批更新催審天數=>補掛催審期限)
                  'Modified by Lydia 2018/12/05 NP10=NA69
                  'stSQL = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np08,np09,np10,np15,np22) " & _
                                "values ('" & .Fields("cp09") & "', '" & .Fields("cp01") & "', '" & .Fields("cp02") & "', '" & .Fields("cp03") & "', '" & .Fields("cp04") & "' " & _
                                ", '305', " & stDate1 & ", " & stDate1 & ", " & CNULL(PUB_GetAKindSalesNo(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))) & _
                                ", '" & ChangeWStringToWDateString(strSrvDate(1)) & "補掛催審期限;" & "', GETNP22  ) "
                  'Modified by Lydia 2023/12/22 +CP01~CP04
                  Call GetNA69("", "" & .Fields("TM10"), PUB_GetAKindSalesNo(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04")), strExc(3), .Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))
                  stSQL = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np08,np09,np10,np15,np22) " & _
                                "values ('" & .Fields("cp09") & "', '" & .Fields("cp01") & "', '" & .Fields("cp02") & "', '" & .Fields("cp03") & "', '" & .Fields("cp04") & "' " & _
                                ", '305', " & stDate1 & ", " & stDate1 & ", '" & strExc(3) & "' " & _
                                ", '" & ChangeWStringToWDateString(strSrvDate(1)) & "補掛催審期限;" & "', GETNP22  ) "
                  cnnConnection.Execute stSQL
                  intC = intC + 1
                  .MoveNext
             Loop
        End With
        cnnConnection.CommitTrans
   End If
   
   MsgBox "共新增" & intC & "筆催審期限! "
   Exit Sub
   
ErrHand28:
   If Err.Number <> 0 Then
        MsgBox Err.Description
        cnnConnection.RollbackTrans
   End If
'CP01 CP02   CP03 CP04 CP09      CP10       CP27       CF05       CP66
'---- ------ ---- ---- --------- ---- ---------- ---------- ----------
'CFT  015454 0    00   BA6044293 301    19221111         70   20171124
'CFT  001445 0    00   B00013662 101    19221111        180   19121588
'CFT  013964 0    00   BA0014453 301    19221111        180   20110519
End Sub

'Added by Lydia 2019/05/16 國外帳單檢查原文字數(參考Frmaccc2150.ChkMailTransFee)
Private Sub Command25_Click()
Dim inA As Integer
Dim rsAD As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset  '檢查範圍
Dim strA1 As String
Dim bolChk1 As Boolean 'Added by Lydia 2018/01/05
Dim strWordVal As String, inB As Integer 'Added by Lydia 2019/05/16 因為備註自動加註訊息(ex.此收文號有虧損;) ,所以先判斷有無";"
Dim bUpdate As Boolean

On Error GoTo ErrorHandle

   'cnt判斷是否為台灣案
   strSql = "select a1501,a1502,a1503,a1509,sum(decode(pa09,'000',1,0)) cnt " & _
               "From acc150, acc151, patent where a1502>=1080201 and a1503 in ('Y53541000','Y52268000','Y54868000') " & _
               "and a1501=axf01 and axf03=pa01||pa02||pa03||pa04 group by a1501,a1502,a1503,a1509 " & _
               "order by a1502,a1501 "
   intI = 1
   Set rs1 = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      rs1.MoveFirst
      cnnConnection.BeginTrans
      Do While Not rs1.EOF
            'Added by Lydia 2019/05/16  抓原文字數
            If Trim("" & rs1.Fields("a1509")) <> "" Then
                 inA = InStr("" & rs1.Fields("a1509"), ";")
                 inB = InStr("" & rs1.Fields("a1509"), "/")
                 If inB > inA Then '訊息在前
                     If inA > 0 Then
                        strWordVal = Val(Mid("" & rs1.Fields("a1509"), inA + 1))
                     Else
                        strWordVal = Val("" & rs1.Fields("a1509"))
                     End If
                 Else '無訊息或訊息在後
                     strWordVal = Val("" & rs1.Fields("a1509"))
                 End If
            End If
            '判斷代理人為舜禹(Y53541)或捷恩凱(Y52268)
            If rsAD.State <> adStateClosed Then rsAD.Close
            'Modified by Lydia 2025/03/13 改用模組取得
            'If Val(strWordVal) > 0 And InStr(外翻Y編號, "" & rs1.Fields("a1503")) > 0 Then
            If Val(strWordVal) > 0 And InStr(Pub_SetF51Order("Y", ""), "" & rs1.Fields("a1503")) > 0 Then
               '抓翻譯費有原文字數和相似度
               strA1 = "SELECT AXF01,AXF02,AXF03,B.*,CP10,NVL(CPM03,CPM04) CPM03 FROM ACC151,TRANSFEE B,CASEPROGRESS,CASEPROPERTYMAP " & _
                          "WHERE AXF01='" & rs1.Fields("a1501") & "' AND AXF02=TF01(+) AND TF01=CP09(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP10 IN ('201','927') AND NVL(TF23,0) > 0 AND NVL(TF19,0) > 0 "
               rsAD.CursorLocation = adUseClient
               rsAD.Open strA1, cnnConnection, adOpenStatic, adLockReadOnly
               If rsAD.RecordCount > 0 Then
                     bolChk1 = True 'Added by Lydia 2018/01/05
                     'Modified by Lydia 2019/05/16 "" & rs1.fields("a1509")=> strWordVal
                     strA1 = Format(Val(strWordVal) / (Val("" & rsAD.Fields("TF23")) * (1 - Val("" & rsAD.Fields("TF19")) / 100)), "0.00")
                     '比對完稿字數(備註前固定輸入數字)和原文字數(預估值),超出5%,發email通知
                     'Modified by Lydia 2018/04/02 原文字數比對請排除P案設定,因為無法要求內專輸入原文字數(ex.P119594)
                     'If Val(strA1) > 1.05 Then
                     If Val(strA1) > 1.05 And Mid("" & rsAD.Fields("AXF03"), 1, 3) = "FCP" Then
                         Call ChgCaseNo("" & rsAD.Fields("AXF03"), strExc)
                         strExc(5) = IIf(strExc(3) & strExc(4) = "000", strExc(1) & strExc(2), strExc(1) & strExc(2) & strExc(3) & strExc(4))
                         '內文
                         'Modified by Lydia 2019/05/16 "" & rs1.fields("a1509")=> strWordVal
                         strExc(6) = vbCrLf & "本所案號：" & strExc(5) & vbCrLf & _
                                    "收  文  號：" & rsAD.Fields("AXF02") & "　　" & rsAD.Fields("CPM03") & vbCrLf & _
                                    "原文字數：" & PUB_StrToStr("" & rsAD.Fields("TF23"), 6, True, True) & " 字" & "　　" & _
                                    "相  似  度：" & PUB_StrToStr("" & rsAD.Fields("TF19"), 3, True, True) & " %" & _
                                    IIf("" & rsAD.Fields("TF20") <> "", "　　相似案號：" & rsAD.Fields("TF20"), "") & vbCrLf & _
                                    "完稿字數：" & PUB_StrToStr(Val(strWordVal), 6, True, True) & " 字" & vbCrLf & _
                                    "完稿字數比對結果：" & (Val(strA1) - 1) * 100 & " %"
                         PUB_SendMail strUserNum, "86013", "", strExc(5) & "翻譯字數異常超過5%", strExc(6)
                     End If
               End If
               If rsAD.State <> adStateClosed Then rsAD.Close
               'Added by Lydia 2018/01/05 新案翻譯-舜禹,捷恩凱,迅達需比對完稿字數,備註欄位/前數字與原字數計算,
               If bolChk1 = False Then
                    'Modified by Lydi 2018/01/08 +CP66
                     strA1 = "SELECT AXF01,AXF02,AXF03,B.*,CP10,NVL(CPM03,CPM04) CPM03,CP66 FROM ACC151,TRANSFEE B,CASEPROGRESS,CASEPROPERTYMAP " & _
                             "WHERE AXF01='" & rs1.Fields("a1501") & "' AND AXF02=TF01(+) AND TF01=CP09(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP10 IN ('201') "
                     rsAD.CursorLocation = adUseClient
                     rsAD.Open strA1, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsAD.RecordCount > 0 Then
                         'Modified by Lydia 2018/04/02 原文字數比對請排除P案設定,因為無法要求內專輸入原文字數(ex.P119594)
                         'If "" & rsAD.Fields("CP66") >= "20180101" Then  'Added by Lydia 2018/01/08 從107/1/1開始控管 by Sharon
                         If "" & rsAD.Fields("CP66") >= "20180101" And Mid("" & rsAD.Fields("AXF03"), 1, 3) = "FCP" Then
                             Call ChgCaseNo("" & rsAD.Fields("AXF03"), strExc)
                             strExc(5) = IIf(strExc(3) & strExc(4) = "000", strExc(1) & strExc(2), strExc(1) & strExc(2) & strExc(3) & strExc(4))
                              '若無原文字數可比對, 自動發一Email給程序管制人員,cc:Sharon
                              If Val("" & rsAD.Fields("TF23")) = 0 Or Val("" & rsAD.Fields("TF23")) < 0 Then
                                    strExc(6) = PUB_GetFCPHandler(strExc(1), strExc(2), strExc(3), strExc(4))
                                    If strExc(6) <> "" Then
                                         PUB_SendMail strUserNum, strExc(6), "", strExc(5) & "無原文字數可比對,請後續追蹤交稿字數是否為正確", "同主旨", , , , , , "86013"
                                    End If
                              Else
                              '完稿字數大於原文字數250字, 自動發一Email至Sharon
                                    'Modified by Lydia 2018/04/24 若完稿字數大於原文字數5%(原文10000字以下)或3%(原文超過10000字)
                                    'If Val("" & rs1.fields("a1509")) > 0 And Val("" & rs1.fields("a1509")) > Val("" & rsAD.Fields("TF23")) + 250 Then
                                    'Modified by Lydia 2019/05/16 "" & rs1.fields("a1509")=> strWordVal
                                    If Val(strWordVal) > 0 And Val(strWordVal) > Val("" & rsAD.Fields("TF23")) + Format(IIf(Val("" & rsAD.Fields("TF23")) <= 10000, Val("" & rsAD.Fields("TF23")) * 0.05, Val("" & rsAD.Fields("TF23")) * 0.03), "0") Then
                                         strExc(6) = vbCrLf & "本所案號：" & strExc(5) & vbCrLf & _
                                                    "收  文  號：" & rsAD.Fields("AXF02") & "　　" & rsAD.Fields("CPM03") & vbCrLf & _
                                                    "原文字數：" & PUB_StrToStr("" & rsAD.Fields("TF23"), 6, True, True) & " 字" & "　　" & _
                                                    "完稿字數：" & PUB_StrToStr(Val(strWordVal), 6, True, True) & " 字" & vbCrLf & _
                                                    "完稿字數比對結果：大於" & Val(strWordVal) - Val("" & rsAD.Fields("TF23")) & " 字"
                                         'Modified by Lydia 2018/04/24 改主旨
                                         'PUB_SendMail strUserNum, "86013", "", strExc(5) & "翻譯字數異常大於250字", strExc(6)
                                         PUB_SendMail strUserNum, "86013", "", strExc(5) & "翻譯字數異常大於" & IIf(Val("" & rsAD.Fields("TF23")) <= 10000, "５％", "３％"), strExc(6)
                                    End If
                              End If
                         End If 'end 2018/01/08
                     End If
               End If
               'end 2018/01/05
            End If
            '台灣案清除特定備註
            If Val("" & rs1.Fields("cnt")) > 0 And "" & rs1.Fields("a1509") <> "" And InStr("" & rs1.Fields("a1509"), "最小收文號尚未輸過帳單") > 0 Then
                bUpdate = True
                strSql = "update acc150 set a1509= replace(a1509,'最小收文號尚未輸過帳單;','') where a1501='" & rs1.Fields("a1501") & "' "
                cnnConnection.Execute strSql, intI
            End If
            rs1.MoveNext
      Loop
   End If
   
   If bUpdate = True Then cnnConnection.CommitTrans
   MsgBox "完成 !", vbInformation
   
   Set rsAD = Nothing
   Set rs1 = Nothing
   
   Exit Sub
   
ErrorHandle:
   If bUpdate = True Then
       cnnConnection.RollbackTrans
   End If
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

'Added by Lydia 2019/09/05 活化客戶檔整批處理
'Memo by Lydia 2019/10/16 保留Code
Private Sub Command23_Click_old()
Dim intA As Integer, intQ As Integer
Dim rsAD As New ADODB.Recordset
Dim strA1 As String

'Memo by Lydia 2019/09/05 待整批處理記錄有11040筆，分成LYDIA_A001和LYDIA_A002(超過１萬筆無法匯入DB)；
                                        '最後匯整在LYDIA_A001，在CHKTYPE分別註記1,2
'名稱 類型
'------- -- -------------
'業務區 VARCHAR2(20)
'智權人員 VARCHAR2(20)
'聯絡日期 VARCHAR2(100)
'聯絡內容 VARCHAR2(200)
'客戶編號 VARCHAR2(9)  =>PKey
'客戶現況 VARCHAR2(50)
'統一編號 VARCHAR2(20)
'新統一編號 VARCHAR2(20)
'負責人 VARCHAR2(50)
'負責人變更 VARCHAR2(50)
'客戶名稱 VARCHAR2(100)
'郵遞區號 VARCHAR2(20)
'地址 VARCHAR2(400)
'新地址 VARCHAR2(400)
'新郵遞區號 VARCHAR2(20)
'CHKTYPE VARCHAR2(1)
'PROCTYPE VARCHAR2(10 CHAR) '處理動作
'-------TABLE END

'---挑選符合的DML_LOG丟暫存檔;客戶異動只抓DL12為個人客戶資料修改(frm210101_1)、國內案件接洽記錄單(frm090801)、案件接洽單(frm090801)、客戶基本資料維護(frm140401)、客戶變更名稱作業(frm140101)、客戶資料修改(frm210101_1)、客戶/代理人改號作業(frm12040125)、接洽紀錄單－新增－商標(frm010004)，並剔除操作人員為M51電腦中心及M31財務處人員的資料
'--CREATE TABLE LYDIA_TMP1 AS
'SELECT PKNO,COUNT(*) CNT FROM (
'SELECT DL06,DL07,DL08,SUBSTR(UPPER(DL09),INSTR(UPPER(DL09),'條件 CU01=>')+9,8) PKNO
'FROM DML_LOG,STAFF WHERE DL06=ST01(+) AND ST03<>'M51' AND ST03<>'M31' AND DL07>=20180401 AND UPPER(DL10)='CUSTOMER'
'AND DL12 IN ('個人客戶資料修改(frm210101_1)','國內案件接洽記錄單(frm090801)','案件接洽單(frm090801)','客戶基本資料維護(frm140401)','客戶變更名稱作業(frm140101)','客戶資料修改(frm210101_1)','客戶/代理人改號作業(frm12040125)','接洽紀錄單－新增－商標(frm010004)')
'AND DL09 LIKE '%修改%' ) GROUP BY PKNO ;
'ALTER TABLE LYDIA_TMP1 ADD PRIMARY KEY (PKNO) ;

   strSql = "select a.*,b.cu01||b.cu02 as custno,nvl(b.cu80,'N') cu80,nvl(c.cnt,0) cnt " & _
               "from lydia_a001 a,customer b,lydia_tmp1 c " & _
               "Where proctype is null " & _
               "and substr(客戶編號,1,8)=cu01(+) and substr(客戶編號,9,1)=cu02(+) " & _
               "and substr(客戶編號,1,8)=pkno(+) "
   strSql = strSql & "order by 客戶編號 "
   
   Debug.Print "Start: " & Format(ServerTime, "000000")
   intI = 1
   Set rsAD = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
        '處理A.整理檔有聯絡內容者，新增國內往來記錄：聯絡日期->往來日期COR02，沒有聯絡日期者填108/1/1以示區別；聯絡內容->內容COR05；客戶編號->往來對象COR03；主旨COR04->「活化客戶整理檔聯絡內容」；Create ID->QPGMR。
        '處理B.依整理檔更新客戶檔：新統一編號欄更新CU11統一編號；負責人變更欄更新CU07公司負責人；新地址欄更新CU31聯絡地址 ；新郵遞區號欄更新CU30聯絡地址郵遞區號 (轉全形)；同時寫維護記錄檔DML-LOG；客戶檔的Update ID->QPGMR。整理檔的客戶現況欄不管。
        '處理C.寫入待活化客戶檔OldCustomer，供後續收文參考。欄位：OCU01客戶編號V2(8) Primary_Key、OCU02寫入日期N(8) Primary_Key、OCU03活化日期。
        '----------------
        '更新資料條件:
        '2018/4/1以後有收文者：
        '  1.客戶檔2018/4/1以後有異動者，只做A。
        '  2.沒有異動者則做A+B。
        '2018/4/1以後沒有收文者：
        '  1.客戶檔2018/4/1以後有異動者，只做A。
        '  2.沒有異動者則做A+B；若沒有A或B者才做C，也就是Excel整理檔沒有聯絡內容和修改資料者，且客戶狀態CU80空白或非解散、廢止、撤銷、停業、死亡、遷移不明者才寫入活化客戶。
       
       rsAD.MoveFirst
       Do While Not rsAD.EOF
           cnnConnection.BeginTrans
           strA1 = "" '處理動作記錄
           '已改號刪除
           If "" & rsAD.Fields("custno") = "" Then
                '2018/4/12,4/13,4/17 的大量改號原因: 楊挺的客戶拆成專利和商標，後來承接的智權人員可以完全接手後，再次合併
                strExc(0) = "update lydia_a001 set proctype='已改號刪除' where 客戶編號=" & CNULL(rsAD.Fields("客戶編號"))
                cnnConnection.Execute strExc(0)
           Else
                '處理A.整理檔有聯絡內容者，新增國內往來記錄：
                If Trim("" & rsAD.Fields("聯絡日期") & rsAD.Fields("聯絡內容")) <> "" Then
                     strExc(1) = AutoNo("K", 6) '往來記錄編號
                     If "" & rsAD.Fields("聯絡日期") <> "" And Left("" & rsAD.Fields("聯絡日期"), 1) <> "-" Then
                         strExc(2) = DBDATE("" & rsAD.Fields("聯絡日期"))
                     Else
                         '聯絡日期欄<1070101或沒有聯絡日期者填108/1/1
                         strExc(2) = "20190101"
                     End If
                     If Val(strExc(2)) < 20180101 Then
                         strExc(2) = "20190101"
                     End If
                     
                     '聯絡日期欄<1070101者寫往來記錄時仍以108/1/1寫入，原聯絡日期欄與聯絡內容欄合併
                     'If "" & rsAD.Fields("聯絡日期") <> "" And Left("" & rsAD.Fields("聯絡日期"), 1) <> "-" Then
                     If strExc(2) = "20190101" Then
                          strExc(3) = "" & rsAD.Fields("聯絡日期") & rsAD.Fields("聯絡內容")
                     Else
                          strExc(3) = "" & rsAD.Fields("聯絡內容")
                     End If
                     
                     strExc(0) = "Insert into ContactRecord1 (cor01,cor02,cor03,cor04,cor05) " & _
                                      "values ('" & strExc(1) & "', " & strExc(2) & " , '" & rsAD.Fields("客戶編號") & "', '活化客戶整理檔聯絡內容', " & _
                                      CNULL(strExc(3)) & ") "
                     cnnConnection.Execute strExc(0)
                     '不改變Trigger, 直接改CreateID
                     strExc(0) = "Update ContactRecord1 Set COR06='QPGMR' Where COR01='" & strExc(1) & "' "
                     cnnConnection.Execute strExc(0)
                     strA1 = strA1 & ",A"
                End If
                '處理B.依整理檔更新客戶檔 (客戶檔2018/4/1以後沒有異動者 cnt=0)
                '處理B2：Excel整理檔有修改資料，但是客戶檔2018/4/1以後有異動者，例X00854010。
                'If Val("" & rsAD.Fields("cnt")) = 0 And Trim("" & rsAD.Fields("新統一編號") & rsAD.Fields("負責人變更") & rsAD.Fields("新地址") & rsAD.Fields("新郵遞區號")) <> "" Then
                If Trim("" & rsAD.Fields("新統一編號") & rsAD.Fields("負責人變更") & rsAD.Fields("新地址") & rsAD.Fields("新郵遞區號")) <> "" Then
                     If Val("" & rsAD.Fields("cnt")) > 0 Then
                            strA1 = strA1 & ",B2"
                     Else
                            strExc(1) = ""
                            If Trim("" & rsAD.Fields("新統一編號")) <> "" Then strExc(1) = strExc(1) & ", cu11=" & CNULL(ChgSQL("" & rsAD.Fields("新統一編號")))
                            If Trim("" & rsAD.Fields("負責人變更")) <> "" Then strExc(1) = strExc(1) & ", cu07=" & CNULL(ChgSQL("" & rsAD.Fields("負責人變更")))
                            If Trim("" & rsAD.Fields("新地址")) <> "" Then strExc(1) = strExc(1) & ", cu31=" & CNULL(toDblFont(ChgSQL("" & rsAD.Fields("新地址"))))   '地址=>全形
                            If Trim("" & rsAD.Fields("地址國籍")) <> "" Then strExc(1) = strExc(1) & ", cu87=" & CNULL(ChgSQL("" & rsAD.Fields("地址國籍")))
                            If Trim("" & rsAD.Fields("新郵遞區號")) <> "" And Trim("" & rsAD.Fields("新郵遞區號")) <> "找不到" Then strExc(1) = strExc(1) & ", cu30=" & CNULL(toDblFont(ChgSQL("" & rsAD.Fields("新郵遞區號"))))
                            If strExc(1) <> "" Then
                                 strExc(0) = "Update Customer Set " & Mid(strExc(1), 2) & ",cu84='QPGMR', cu85=" & strSrvDate(1) & ", cu86=" & Left(Format(ServerTime, "000000"), 4) & _
                                                  " Where cu01='" & Mid(rsAD.Fields("客戶編號"), 1, 8) & "' and cu02='" & Mid(rsAD.Fields("客戶編號"), 9, 1) & "' "
                                 
                                 Pub_SeekTbLog strExc(0), "QPGMR"
                                 cnnConnection.Execute strExc(0)
                                 strA1 = strA1 & ",B"
                            End If
                     End If
                End If
                
                '處理C. 2018/4/1以後沒有A類收文者,並且 Excel整理檔沒有聯絡內容和修改資料者，且客戶狀態CU80空白或非解散、廢止、撤銷、停業、死亡、遷移不明者才寫入活化客戶。
                
                '沒有A,B => 判斷客戶狀態
                If strA1 = "" And InStr("解散、廢止、撤銷、停業、死亡、遷移不明", "" & rsAD.Fields("cu80")) = 0 Then
                     strExc(1) = "Select Count(*) Cnt1 From Caseprogress Where Cp05>=20180401 And Cp57 Is Null and substr(cp09,1,1)='A' And (Cp01,Cp02,Cp03,Cp04) In (" & _
                                     "Select Pa01,Pa02,Pa03,Pa04 From Patent Where Instr(Pa26||','||Pa27||','||Pa28||','||Pa29||','||Pa30,'" & rsAD.Fields("客戶編號") & "') > 0 " & _
                                     "union all select tm01,tm02,tm03,tm04 from trademark where instr(tm23||','||tm78||','||tm79||','||tm80||','||tm81,'" & rsAD.Fields("客戶編號") & "') > 0 " & _
                                     "union all select sp01,sp02,sp03,sp04 from servicepractice where instr(sp08||','||sp58||','||sp59||','||sp65||','||sp66,'" & rsAD.Fields("客戶編號") & "') > 0 " & _
                                     "union all select lc01,lc02,lc03,lc04 from lawcase where instr(lc11||','||lc43||','||lc44||','||lc45||','||lc46,'" & rsAD.Fields("客戶編號") & "') > 0 ) "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
                     strExc(1) = ""
                     If intI = 1 Then strExc(1) = "" & RsTemp.Fields(0)
                     '沒有收文
                     If Val(strExc(1)) = 0 Then
                         strExc(0) = "Insert into OldCustomer (OCU01,OCU02) Values ('" & Mid(rsAD.Fields("客戶編號"), 1, 8) & "' , " & strSrvDate(1) & ") "
                         cnnConnection.Execute strExc(0)
                         strA1 = strA1 & ",C"
                     End If
                End If
                
                '暫存B2->空白
                strA1 = Replace(strA1, ",B2", "")
                
                If strA1 = "" Then
                     strExc(0) = "update lydia_a001 set proctype='無' where 客戶編號=" & CNULL(rsAD.Fields("客戶編號"))
                     cnnConnection.Execute strExc(0)
                Else
                     strExc(0) = "update lydia_a001 set proctype='" & Mid(strA1, 2) & "' where 客戶編號=" & CNULL(rsAD.Fields("客戶編號"))
                     cnnConnection.Execute strExc(0)
                End If
           End If

           cnnConnection.CommitTrans
           rsAD.MoveNext
       Loop
   End If
   Debug.Print "End:   " & Format(ServerTime, "000000")
   MsgBox "完成 !", vbInformation
   
   Set rsAD = Nothing

   Exit Sub
   
ErrorHandle:

   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Sub txtFTM_GotFocus(Index As Integer)
   TextInverse txtFTM(Index)
End Sub

Private Sub txtFTM_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtWD01_GotFocus()
   TextInverse txtWD01
End Sub

Private Sub txtWD01_Validate(Cancel As Boolean)
   If txtWD01 = "" Then Exit Sub
   strExc(0) = "select * from workday where wd01=" & DBDATE(txtWD01)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      txtWD07 = "" & RsTemp("wd07")
   Else
      MsgBox "非工作日！", vbExclamation
      Cancel = True
   End If
   txtWD01.Tag = txtWD01
End Sub

Private Sub txtWD07_GotFocus()
   TextInverse txtWD07
End Sub

Private Sub txtWD07_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2019/10/16 北所南所刪址客戶處理
Private Sub Command23_Click()
Dim intA As Integer, intQ As Integer
Dim rsAD As New ADODB.Recordset
Dim strA1 As String
Dim stDept As String, stDeptSales As String
Dim stSalesNo As String, stSalesName As String
'Added by Lydia 2019/10/23
Dim stCust01 As String, stCust02 As String '要處理的客戶代號
Dim rsBD As New ADODB.Recordset
Dim strUpdNp As String 'Added by Lydia 2019/11/05 要更新下一程序的智權人員

'Memo by Lydia 2019/10/16 匯入Lydia_Tmp1(共5671筆):  先在excel檔人工處理出”紅字判斷(整筆記錄為紅字=Y)”或”狀態判斷(客戶狀態欄為空白=Y)”
'                                         建立Lydia_A001(共5670筆):  因為X44204010有重覆記錄；加上PROCTYPE V2(1) 處理動作
'Memo by Lydia 2019/10/18 修改地址; 重新匯入Lydia_A002
'Memo by Lydia 2019/10/30 客服組刪址客戶: 匯入Lydia_A001共244筆,剔除與Lydia_A002重覆的45筆, 直接再匯入Lydia_A002
'Memo by Lydia 2019/11/05 忘記一併更新下一程序, 實際上使用要先測試
'參考文件：
'1.  各區紅色資料或客戶狀態欄為空者：依EXCEL檔更新智權人員(屬於下方第3~6之智權人員資料智權人員已修改，此處不再改)
                                                                '、負責人、客戶狀態、中文地址、聯絡地址；
                                                                '若EXCEL檔之客戶狀態欄為遷移不明、解散、廢止、撤銷、停業、死亡者，
                                                                '不管EXCEL檔之智權人員為何一律都改為各區虛建智權人員；同時寫維護記錄檔DML-LOG；客戶檔的Update ID->QPGMR。
'2.  各區非紅色且有客戶狀態欄資料：更新客戶狀態欄，並統一將智權人員改為各區虛建智權人員；同時寫維護記錄檔DML-LOG；客戶檔的Update ID->QPGMR。
'3.  杜主秘客戶僅限於EXCEL檔案中之客戶才轉給客服組W1001(已逐客戶修改)。
'4.  北二區所有客戶轉客服組W1001(已整批改以智權人員客戶轉移作業處理)。
'5.  北一區彭德明.蕭文津.葉明色的資料更改為北一區10011(已整批改以智權人員客戶轉移作業處理)。
'6.  北四區邱南豪.吳家清的資料更改為北四區10041(已整批改以智權人員客戶轉移作業處理)；
     '楊挺的資料更改為李承翰A2033(已整批改以智權人員客戶轉移作業處理)。
'--------------------------------------------------------------------------------------------------------------

    If MsgBox("確定執行" & Command40.Caption & " ？", vbInformation + vbYesNo + vbDefaultButton2, "確定執行") = vbNo Then
        Exit Sub
    End If
    
    'Modified by Lydia 2019/10/18 改成以客戶檔的CU12判斷部門
    'Modified by Lydia 2019/11/05 +CU13
    strSql = "SELECT 紅字判斷,狀態判斷,CU12 AS A0901,業務區,智權人員,國籍,客戶編號,統一編號,負責人,客戶名稱,客戶狀態," & _
                " 中文地址,中文地址郵遞區號,客戶國籍,聯絡地址,聯絡地址郵遞區號,地址國籍,電話一,電話二,CU13" & _
                " From LYDIA_A002, Customer Where PROCTYPE Is Null AND Substr(客戶編號,1,8)=Cu01(+) And Substr(客戶編號,9,1)=Cu02(+)"
    strSql = strSql & " ORDER BY 業務區,智權人員"
   Set rsAD = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
       cnnConnection.BeginTrans
       rsAD.MoveFirst
       Do While Not rsAD.EOF
           strUpdNp = "" 'Added by Lydia 2019/11/05
           If "" & rsAD.Fields("A0901") = "" Then
                GoTo JumpNextRec
           End If
           If stDept <> "" & rsAD.Fields("A0901") Then
                '抓各區虛建智權人員
                strSql = "select MIN(ST01) minno from staff where st15='" & rsAD.Fields("A0901") & "' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then
                    stDeptSales = "" & RsTemp.Fields("minno")
                End If
                stDept = "" & rsAD.Fields("A0901")
           End If
           
           'Added by Lydia 2019/10/23 舊客戶也要改: 變更智權人員+狀態+地址, 不變更負責人
           If Val(Mid("" & rsAD.Fields("客戶編號"), 9, 1)) > 0 Then
                strSql = " cu01||cu02=" & CNULL("" & rsAD.Fields("客戶編號"))
           Else
                strSql = " cu01=" & CNULL(Left("" & rsAD.Fields("客戶編號"), 8))
           End If
           strSql = "select cu01,cu02 from customer where " & strSql & " order by cu01"
           intI = 1
           Set rsBD = ClsLawReadRstMsg(intI, strSql)
           If intI = 1 Then
               rsBD.MoveFirst
               Do While Not rsBD.EOF
                    stCust01 = "" & rsBD.Fields("cu01")
                    stCust02 = "" & rsBD.Fields("cu02")
                    'Mark by Lydia 2019/11/05 補上更新
                    'strUpdNp = "" & rsAD.Fields("CU13")
                    'GoTo JumpToPlus
           
                    strExc(1) = ""
                    'Added by Lydia 2019/10/23 判斷舊客戶否存在lydia_a002
                    If Val(Mid("" & rsAD.Fields("客戶編號"), 9, 1)) = 0 And stCust02 <> "0" Then
                        strSql = "select 客戶編號 from lydia_a002 where 客戶編號=" & CNULL(stCust01 & stCust02)
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                             GoTo JumpToNext
                        End If
                    End If
                    strA1 = "" '處理動作記錄
                    '1.  各區紅色資料或客戶狀態欄為空者
                    If "" & rsAD.Fields("紅字判斷") = "Y" Or "" & rsAD.Fields("狀態判斷") = "Y" Then
                         strA1 = "1"
                         If "" & rsAD.Fields("智權人員") <> "" Then
                             'Added by Lydia 2019/10/30 排除客服組
                             If Trim("" & rsAD.Fields("智權人員")) = "客服組" Then
                             Else
                             'end 2019/10/30
                                    '更新-智權人員
                                    If InStr("杜清麟,彭德明,蕭文津,葉明色,邱南豪,吳家清,楊挺", "" & rsAD.Fields("智權人員")) = 0 And "" & rsAD.Fields("業務區") <> "北二區" Then
                                        If stSalesName <> Trim("" & rsAD.Fields("智權人員")) Then
                                            stSalesNo = ""
                                            strSql = "select st01 from staff where st15 like 'S%' and st02='" & Trim(rsAD.Fields("智權人員")) & "' order by 1"
                                            intI = 1
                                            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                            If intI = 1 Then
                                                stSalesNo = "" & RsTemp.Fields("st01")
                                            End If
                                            stSalesName = Trim("" & rsAD.Fields("智權人員"))
                                        End If
                                        If Trim("" & rsAD.Fields("客戶狀態")) <> "" And InStr("遷移不明、解散、廢止、撤銷、停業、死亡", Trim("" & rsAD.Fields("客戶狀態"))) > 0 Then
                                             strExc(1) = strExc(1) & ", CU13='" & stDeptSales & "' "
                                             strUpdNp = stDeptSales  'Added by Lydia 2019/11/05
                                        ElseIf stSalesNo <> "" Then
                                             strExc(1) = strExc(1) & ", CU13='" & stSalesNo & "' "
                                             strUpdNp = stSalesNo  'Added by Lydia 2019/11/05
                                        End If
                                    End If
                             End If
                             '負責人,客戶狀態,中文地址,中文地址郵遞區號,客戶國籍,聯絡地址,聯絡地址郵遞區號,地址國籍
                             'Modified by Lydia 2019/10/23  舊客戶也要改: 變更智權人員+狀態+地址, 不變更負責人
                             'strExc(1) = strExc(1) & ", CU07=" & CNULL(ChgSQL(Trim("" & rsAD.Fields("負責人"))))
                             If stCust02 = "0" Then strExc(1) = strExc(1) & ", CU07=" & CNULL(ChgSQL(Trim("" & rsAD.Fields("負責人"))))
                             
                             strExc(1) = strExc(1) & ", CU80=" & CNULL(ChgSQL(Trim("" & rsAD.Fields("客戶狀態"))))
                             strExc(1) = strExc(1) & ", CU23=" & CNULL(ChgSQL(toDblFont(Trim("" & rsAD.Fields("中文地址")))))
                             If Trim("" & rsAD.Fields("中文地址郵遞區號")) <> "" And Trim("" & rsAD.Fields("中文地址郵遞區號")) <> "找不到" Then strExc(1) = strExc(1) & ", CU112=" & CNULL(toDblFont(ChgSQL("" & rsAD.Fields("中文地址郵遞區號"))))
                             If Trim("" & rsAD.Fields("客戶國籍")) <> "" And Trim("" & rsAD.Fields("客戶國籍")) <> "找不到" Then strExc(1) = strExc(1) & ", CU10=" & CNULL(ChgSQL(Trim("" & rsAD.Fields("客戶國籍"))))
                             strExc(1) = strExc(1) & ", CU31=" & CNULL(ChgSQL(toDblFont(Trim("" & rsAD.Fields("聯絡地址")))))
                             If Trim("" & rsAD.Fields("聯絡地址郵遞區號")) <> "" And Trim("" & rsAD.Fields("聯絡地址郵遞區號")) <> "找不到" Then strExc(1) = strExc(1) & ", CU30=" & CNULL(toDblFont(ChgSQL("" & rsAD.Fields("聯絡地址郵遞區號"))))
                             If Trim("" & rsAD.Fields("地址國籍")) <> "" And Trim("" & rsAD.Fields("地址國籍")) <> "找不到" Then strExc(1) = strExc(1) & ", CU87=" & CNULL(ChgSQL(Trim("" & rsAD.Fields("地址國籍"))))
                         End If
                    Else '2.  各區非紅色且有客戶狀態欄資料
                         strA1 = "2"
                         '更新客戶狀態欄，並統一將智權人員改為各區虛建智權人員
                         'Modified by Lydia 2019/10/21 排除處理1~2
                         'If InStr("杜清麟,彭德明,蕭文津,葉明色,邱南豪,吳家清,楊挺", "" & rsAD.Fields("智權人員")) = 0 And "" & rsAD.Fields("業務區") <> "北二區" Then
                         'Added by Lydia 2019/10/30 排除客服組
                         If Trim("" & rsAD.Fields("智權人員")) = "客服組" Then
                         Else
                         'end 2019/10/30
                                If "" & rsAD.Fields("智權人員") <> "杜清麟" And "" & rsAD.Fields("業務區") <> "北二區" Then
                                    strExc(1) = strExc(1) & ", CU13='" & stDeptSales & "' "
                                    strUpdNp = stDeptSales 'Added by Lydia 2019/11/05
                                End If
                         End If
                         strExc(1) = strExc(1) & ", CU80=" & CNULL(ChgSQL(Trim("" & rsAD.Fields("客戶狀態"))))
                    End If
                    If strExc(1) <> "" Then
                          strExc(0) = "Update Customer Set " & Mid(strExc(1), 2) & ",cu84='QPGMR', cu85=" & strSrvDate(1) & ", cu86=" & Left(Format(ServerTime, "000000"), 4) & _
                                          " Where cu01='" & stCust01 & "' and cu02='" & stCust02 & "' "
                         
                         Pub_SeekTbLog strExc(0), "QPGMR"
                         cnnConnection.Execute strExc(0)
                         'Added by Lydia 2019/11/05 更新下一程序
JumpToPlus: '補上更新資料
                         If strUpdNp <> "" Then
                                'strExc(1) = "20191031" '刪址更新的時間
                                'strA1 = "3" '暫時區別
                                strExc(1) = strSrvDate(1)
                                strSql = "Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Patent Where NP02<>'FCP' AND NP02=PA01 And NP03=PA02 And NP04=PA03 And NP05=PA04 And NP06 Is Null And NP09>=" & strExc(1) & " And  PA26='" & stCust01 & stCust02 & "' "
                                strSql = strSql & " and np10 <> '" & strUpdNp & "' " & strNpSqlOfNoSalesDuty
                                strSql = strSql & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Trademark Where NP02<>'FCT' AND NP02=TM01 And NP03=TM02 And NP04=TM03 And NP05=TM04 And NP06 Is Null And NP09>=" & strExc(1) & " And  TM23='" & stCust01 & stCust02 & "' "
                                strSql = strSql & " and np10 <> '" & strUpdNp & "' " & strNpSqlOfNoSalesDuty
                                strSql = strSql & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Lawcase Where (NP02<>'FCL' and NP02<>'LIN') AND NP02=LC01 And NP03=LC02 And NP04=LC03 And NP05=LC04 And NP06 Is Null And NP09>=" & strExc(1) & " And  LC11='" & stCust01 & stCust02 & "' "
                                strSql = strSql & " and np10 <> '" & strUpdNp & "' " & strNpSqlOfNoSalesDuty
                                strSql = strSql & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, Hirecase Where NP02=HC01 And NP03=HC02 And NP04=HC03 And NP05=HC04 And NP06 Is Null And NP09>=" & strExc(1) & " And  HC05='" & stCust01 & stCust02 & "' "
                                strSql = strSql & " and np10 <> '" & strUpdNp & "' " & strNpSqlOfNoSalesDuty
                                strSql = strSql & " Union Select NP01, NP07, NP22,NP02,NP03,NP04,NP05 From NextProgress, ServicePractice Where NP02<>'FG' AND NP02=SP01 And NP03=SP02 And NP04=SP03 And NP05=SP04 And NP06 Is Null And NP09>=" & strExc(1) & " And  SP08='" & stCust01 & stCust02 & "' "
                                strSql = strSql & " and np10 <> '" & strUpdNp & "' " & strNpSqlOfNoSalesDuty
                                intI = 1
                                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                If intI = 1 Then
                                    RsTemp.MoveFirst
                                    Do While Not RsTemp.EOF
                                         strExc(0) = "Update Nextprogress Set NP10='" & strUpdNp & "',NP15='108/11/05       108/10/31整批更新刪址客戶補做下一程序未續辦未過期期限更新智權人員;'||NP15  Where NP01='" & RsTemp.Fields(0).Value & "' And NP07='" & RsTemp.Fields(1).Value & "' And NP22=" & RsTemp.Fields(2).Value & _
                                                           " And NP02='" & RsTemp.Fields("np02") & "' And NP03='" & RsTemp.Fields("np03") & "' And NP04='" & RsTemp.Fields("np04") & "' And NP05='" & RsTemp.Fields("np05") & "' "
                                         Pub_SeekTbLog strExc(0), "QPGMR"
                                         cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"
                                         cnnConnection.Execute strExc(0)
                                         cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"
                                         RsTemp.MoveNext
                                    Loop
                                    'strA1 = "4" '暫時區別 'Mark
                                End If
                         End If
                         'end 2019/11/05
                    End If
JumpToNext: '舊客戶已存在lydia_a002
                    rsBD.MoveNext
               Loop
           End If 'end 2019/10/23
           strExc(0) = "update lydia_a002 set proctype='" & strA1 & "' where 客戶編號=" & CNULL(rsAD.Fields("客戶編號"))
           cnnConnection.Execute strExc(0)
           'Added by Lydia 2019/10/30 一併更新
           strExc(0) = "update lydia_a001 set proctype='" & strA1 & "' where 客戶編號=" & CNULL(rsAD.Fields("客戶編號"))
           cnnConnection.Execute strExc(0)
JumpNextRec:
           rsAD.MoveNext
       Loop
       cnnConnection.CommitTrans '移到最後
   End If
   MsgBox "OK !"
   
End Sub

'Added by Lydia 2020/03/30 利益衝突: FCP案和FMP案若未銷閉卷,檢查是否有D類收文English_vers和專利案件
Private Sub Command35_Click()
Dim bolUpdate As Boolean
Dim stCP12 As String, stCP13 As String, stCP14 As String
Dim strNo As String
Dim intCnt As Long
              
On Error GoTo ErrHandle

     strSql = "Select Pa01,Pa02,Pa03,Pa04,Pa05,Pa26,Pa75,Pa58,Pa108,v09,v65,x09,x65 " & _
                 "From Patent,Caseprogress C1 " & _
                 ",(Select Cp01 V01,Cp02 V02,Cp03 V03,Cp04 V04, Cp09 As V09,Cp65 As V65 From Caseprogress  Where Substr(Cp09,1,1)='D' And Cp10='991') Vt1 " & _
                 ",(Select cp01 x01,cp02 x02,cp03 x03,cp04 x04, Cp09 As x09,cp65 as x65 From Caseprogress Where Substr(Cp09,1,1)='D' And Cp10='992') xt1 " & _
                "Where Pa01 In ('FCP','P') And Pa01=Cp01(+) And Pa02=Cp02(+) And Pa03=Cp03(+) And Pa04=Cp04(+) And Cp31='Y' And Cp12 Like 'F%' " & _
                " and cp01=v01(+) and cp02=v02(+) and cp03=v03(+) and cp04=v04(+) and cp01=x01(+) and cp02=x02(+) and cp03=x03(+) and cp04=x04(+) "
     strSql = strSql & " and pa01='FCP' and pa02 between '060000' and '064999' "
     strSql = strSql & " group by Pa01,Pa02,Pa03,Pa04,Pa05,Pa26,Pa75,Pa58,Pa108,v09,v65,x09,x65 order by 1,2 "
     intI = 1
     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
     If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
              If "" & RsTemp.Fields("pa26") & RsTemp.Fields("pa75") <> "" Then
                    stCP13 = PUB_GetFCPSalesNo(RsTemp.Fields("pa01"), RsTemp.Fields("pa02"), RsTemp.Fields("pa03"), RsTemp.Fields("pa04"))   'FCP承辦
                    stCP12 = GetSalesArea(stCP13)
              Else '舊資料無pa26,pa75 => ex.FCP-012558
                    GoTo JumpToNext
              End If
              stCP14 = "QPGMR"
              If "" & RsTemp.Fields("pa58") & RsTemp.Fields("pa108") = "" Then  '未銷閉卷
                  If bolUpdate = False Then
                      bolUpdate = True
                      cnnConnection.BeginTrans
                  End If

                  '專利案件991
                  If "" & RsTemp.Fields("v09") = "" Then
                        strNo = AutoNo("D", 6)
                        strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
                           ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp65,cp66,cp67 ) values ('" & RsTemp.Fields("pa01") & "','" & RsTemp.Fields("pa02") & "','" & RsTemp.Fields("pa03") & "','" & RsTemp.Fields("pa04") & "'," & _
                            " 19221111,'" & strNo & "','" & cnt專利案件 & "' " & _
                            ",'" & stCP12 & "','" & stCP13 & "','" & stCP14 & "','N','N',19221111,'N','QPGMR'," & strSrvDate(1) & ", " & Left(Format(ServerTime, "000000"), 4) & ")"
                        cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"
                        cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"
                  ElseIf "" & RsTemp.Fields("v65") = "QPGMR" Then '最初搬檔的CP13和CP14放反了
                        strSql = "update caseprogress set cp12='" & stCP12 & "' ,cp13='" & stCP13 & "', cp14='" & stCP14 & "'  where cp09='" & RsTemp.Fields("v09") & "' "
                        cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"
                        cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"
                  End If
                  
                  'English_vers992
                  If "" & RsTemp.Fields("x09") = "" Then
                        strNo = AutoNo("D", 6)
                        strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
                           ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp65,cp66,cp67 ) values ('" & RsTemp.Fields("pa01") & "','" & RsTemp.Fields("pa02") & "','" & RsTemp.Fields("pa03") & "','" & RsTemp.Fields("pa04") & "'," & _
                            " 19221111,'" & strNo & "','" & cntEnglish_Vers & "' " & _
                            ",'" & stCP12 & "','" & stCP13 & "','" & stCP14 & "','N','N',19221111,'N','QPGMR'," & strSrvDate(1) & ", " & Left(Format(ServerTime, "000000"), 4) & ")"
                        cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"
                        cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"
                  ElseIf "" & RsTemp.Fields("x65") = "QPGMR" Then
                        strSql = "update caseprogress set cp12='" & stCP12 & "' ,cp13='" & stCP13 & "', cp14='" & stCP14 & "'  where cp09='" & RsTemp.Fields("x09") & "' "
                        'cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"
                        cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"
                  End If
                  
                  intCnt = intCnt + 1
              '已銷閉卷／舊檔搬移
              ElseIf "" & RsTemp.Fields("v09") & RsTemp.Fields("x09") <> "" Then
                  If bolUpdate = False Then
                      bolUpdate = True
                      cnnConnection.BeginTrans
                  End If
                  '專利案件991
                  If "" & RsTemp.Fields("v09") <> "" And "" & RsTemp.Fields("v65") = "QPGMR" Then
                        strSql = "update caseprogress set cp12='" & stCP12 & "' ,cp13='" & stCP13 & "', cp14='" & stCP14 & "'  where cp09='" & RsTemp.Fields("v09") & "' "
                        'cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"
                        cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"
                  End If
                  'English_vers992
                  If "" & RsTemp.Fields("x09") <> "" And "" & RsTemp.Fields("x65") = "QPGMR" Then
                        strSql = "update caseprogress set cp12='" & stCP12 & "' ,cp13='" & stCP13 & "', cp14='" & stCP14 & "'  where cp09='" & RsTemp.Fields("x09") & "' "
                        'cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"
                        cnnConnection.Execute strSql
                        cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"
                  End If
                  
                  intCnt = intCnt + 1
              End If
JumpToNext:
              RsTemp.MoveNext
         Loop
     End If
     
     If bolUpdate = True Then
          cnnConnection.CommitTrans
          MsgBox "完成" & intCnt & "筆!", vbInformation
     End If
     
     Exit Sub
     
ErrHandle:
     If bolUpdate = True Then
         cnnConnection.RollbackTrans
         If Err.Number <> 0 Then
             MsgBox "檢查D類English_Vers和專利案件，發生錯誤：" & vbCrLf & Err.Description
         End If
     End If
End Sub

Private Sub Write2File(pText As String, pFileName As String)
   Dim stFile As String, ffa As Integer
   
On Error GoTo ErrHnd
  
   If InStr(pFileName, "\") > 0 Then
      stFile = pFileName
   Else
      stFile = App.path & "\" & pFileName
   End If
   
   ffa = FreeFile
On Error GoTo ErrHnd2

   Open stFile For Append As ffa
   Print #ffa, pText
   
ErrHnd2:
   Close ffa
   
ErrHnd:
   If Err.Number = 53 Then Resume Next
      
End Sub

'Added by Lydia 2020/09/22 (109/09/21) 智權部開拓客戶合併地址條，判斷是否為對造客戶
Private Sub Command40_Click()
Dim rsRd As New ADODB.Recordset
Dim strR1 As String, intR As Integer
Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
   
    If MsgBox("確定執行" & Command40.Caption & " ？", vbInformation + vbYesNo + vbDefaultButton2, "確定執行") = vbNo Then
        Exit Sub
    End If
    
    strSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
    strSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
    StrSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
    StrSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
    strSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
    
    strR1 = "select * from lydia_a4 order by 1,2"
    intR = 1
    Set rsRd = ClsLawReadRstMsg(intR, strR1)
    If intR = 1 Then
        rsRd.MoveFirst
        Do While Not rsRd.EOF
             Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, "" & rsRd.Fields("公司名稱"), ">0")
             strR1 = "select * from R100102_1 where id = " & CNULL(strUserNum & "@" & Me.Name)
             intR = 1
             Set RsTemp = ClsLawReadRstMsg(intR, strR1)
             If intR = 1 Then
                  strSql = " update lydia_a4 set 對造='Y' where 部門代號='" & rsRd.Fields("部門代號") & "' and 順序='" & rsRd.Fields("順序") & "' "
                  cnnConnection.Execute strSql, intI
             End If
             rsRd.MoveNext
        Loop
    End If
    MsgBox "OK !"
End Sub

'Added by Lydia 2020/09/23 檢查-更換造字欄位表Excel
Private Sub Command41_Click()
Dim xlsPoint As New Excel.Application
Dim wksPoint As New Worksheet
Dim strNowTitle As String
Dim tmpStr As String, tmpArr As Variant
Dim intC As Integer, intWks As Integer, nRows As Integer
Dim strConList As String
Dim strMid As String
Dim tmpStr2 As String, tmpArr2 As Variant
Dim rsAD As New ADODB.Recordset
Dim strFindWord As String
Dim strTempFile As String

    strFindWord = "?" '判斷輸入Unicode在儲存後產生的亂碼
    strTempFile = PUB_Getdesktop & "\檢查-更換造字欄位表.xls"
    
    If Dir(strTempFile) <> "" Then
         Kill strTempFile
    End If
    
    strExc(0) = "select * from EditOraInEudc where upper(tbname) in ('CUSTOMER','PATENT','TRADEMARK','IPDEPTKEYWORD') order by sno asc "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
         With RsTemp
             .MoveFirst
             intC = 1
             Do While Not .EOF
                  If "" & .Fields("tbname") <> "" And "" & .Fields("pkno") & .Fields("updlist") <> "" Then
                      tmpStr = "" & .Fields("updlist")
                      tmpArr = Empty
                      tmpArr = Split(tmpStr, ",")
                      '------------------------------------------
                       '將所有結果併成一個Grid
                      tmpStr2 = "" & .Fields("showlist")
                      tmpArr2 = Empty
                      tmpArr2 = Split(tmpStr2, ",")
                      strExc(2) = ""
                      strMid = " select '" & strFindWord & "' as N01, '1' as N02,'" & .Fields("sno") & "'  as N03,'" & .Fields("tbname") & "' as N04, 'XXX' as N05,"
                      For intI = 0 To UBound(tmpArr2)
                           If Trim(tmpArr2(intI)) <> "" Then
                               '用原欄位名稱; 因為ACC420、Periodical、SalesNo、ACC490的主鍵同為更新欄位或不在前面的欄位,所以改成比對所有列出欄位
                               strMid = strMid & " " & Trim(tmpArr2(intI)) & ","
                           End If
                      Next
                      strMid = Replace(strMid, "XXX", "6")
                      
                      '組合WHERE語法 (N01=KeyWord , N02=輸入排序,N03=SNo, N04=tablename,N05=開始Update欄位)
                      For intI = 0 To UBound(tmpArr)
                           If Trim(tmpArr(intI)) <> "" Then
                               strExc(2) = strExc(2) & " instr(" & tmpArr(intI) & ",'" & strFindWord & "')>0 or"
                           End If
                      Next
                  End If
                  strMid = Mid(strMid, 1, Len(strMid) - 1) & " From " & .Fields("tbname") & " Where " & Mid(strExc(2), 1, Len(strExc(2)) - 2)
                  strConList = strConList & strMid
                  If UBound(tmpArr2) < 4 Then
                       strExc(0) = ""
                       For intI = 0 To UBound(tmpArr2)
                            strExc(0) = strExc(0) & "," & intI + 6 & " asc"
                       Next intI
                       strConList = strConList & "order by " & Mid(strExc(0), 2) & " ;"
                  Else
                       strConList = strConList & " order by 6 asc, 7 asc ,8 asc, 9 asc ;"
                  End If
                  intC = intC + 1
                  .MoveNext
             Loop
         End With
    End If
    
    tmpArr = Empty
    tmpArr = Split(strConList, ";")
    For intC = 0 To UBound(tmpArr) - 1
        If Trim(tmpArr(intC)) <> "" Then
            intI = 1
            Set rsAD = ClsLawReadRstMsg(intI, Trim(tmpArr(intC)))
            If intI = 1 Then
                rsAD.MoveFirst
                If strNowTitle <> "" & rsAD.Fields("N04") Then
                     intWks = intWks + 1
                     ReDim tmpArr2(1 To rsAD.Fields.Count - 5)
                     If intWks = 1 Then
                         xlsPoint.SheetsInNewWorkbook = intWks
                         xlsPoint.Workbooks.add
                         xlsPoint.Visible = False '預設不顯示
                     Else
                         xlsPoint.Worksheets("工作表" & intWks - 1).Name = strNowTitle  '工作表：更名為Tbname
                         xlsPoint.Worksheets.add
                     End If
                     Set wksPoint = xlsPoint.Worksheets("工作表" & intWks)
                     xlsPoint.Sheets(intWks).Select '選擇工作表
                     nRows = 1
                     For intI = 1 To rsAD.Fields.Count - 5
                        strExc(0) = Pub_NumberToSystem26(intI)
                        wksPoint.Range(strExc(0) & ":" & strExc(0)).ColumnWidth = 13
                        wksPoint.Range(strExc(0) & ":" & strExc(0)).HorizontalAlignment = xlLeft
                        wksPoint.Range(strExc(0) & ":" & strExc(0)).NumberFormat = "@"
                        tmpArr2(intI) = "" & rsAD.Fields(intI + 4).Name
                     Next intI
                     wksPoint.Range("A" & nRows & ":" & Pub_NumberToSystem26(rsAD.Fields.Count - 5) & nRows).Value = tmpArr2
                     nRows = nRows + 1
                     strNowTitle = "" & rsAD.Fields("N04")
                End If
                Do While Not rsAD.EOF
                     For intI = 1 To rsAD.Fields.Count - 5
                        tmpArr2(intI) = "" & rsAD.Fields(intI + 4).Value
                     Next intI
                     wksPoint.Range("A" & nRows & ":" & Pub_NumberToSystem26(rsAD.Fields.Count - 5) & nRows).Value = tmpArr2
                     nRows = nRows + 1
                     rsAD.MoveNext
                Loop
            End If
        End If
    Next intC
    
    If intWks > 0 Then
        xlsPoint.Worksheets("工作表" & intWks).Name = strNowTitle   '工作表：更名為Tbname (最後)
        
        xlsPoint.Sheets(1).Select '選擇工作表
        '判斷版本
        If Val(xlsPoint.Version) < 12 Then
             xlsPoint.Workbooks(1).SaveAs FileName:=strTempFile, FileFormat:=-4143
        Else
             xlsPoint.Workbooks(1).SaveAs FileName:=strTempFile, FileFormat:=56
        End If
        xlsPoint.Workbooks.Close
        xlsPoint.Quit
        Set wksPoint = Nothing
        Set xlsPoint = Nothing
    End If

    MsgBox "OK !"
        
    Exit Sub
    
ErrHandle:
    
End Sub

'Added by Lydia 2021/01/21 重新下載CFT,CFP之UK檔：在測試時，把共用資料夾檔案放上M51
Private Sub Command44_Click()
Dim iRound As Integer
Dim m_TempDir(1 To 2) As String
Dim strTempFile As String
Dim strP1 As String, strP2 As String, strP3 As String

   strP3 = "":    strP1 = "":    strP2 = ""
   m_TempDir(1) = "C:\Users\A3034\Desktop\卷宗匯入區\UKIPO" 'CFT
   m_TempDir(2) = "C:\Users\A3034\Desktop\卷宗匯入區\CFP通知英國再註冊設計" 'CFP
   For iRound = 1 To 2
       If iRound = 1 Then
           strSql = "Select Cpp01,Cpp02,Nvl(Tm15,Tm12) As Pa22,'UK009'||substr(Nvl(Tm15,Tm12),2,8) Pfilename From Caseprogress,Casepaperpdf,Trademark " & _
                       "Where Cp01='CFT' And Cp10='1730' And Cp05>=20210115  And Cp09=Cpp01(+) " & _
                        "and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) "
       Else
           strSql = "Select Cpp01,Cpp02,pa22,'9'||replace(pa22,'-','') pfilename From Caseprogress,Casepaperpdf,Patent " & _
                        "Where Cp01='CFP' And Cp10='1608' And Cp05>=20210115  And Cp09=Cpp01(+) " & _
                        "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
       End If
       strSql = strSql & " order by cp67"
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
       If intI = 1 Then
           RsTemp.MoveFirst
           Do While Not RsTemp.EOF
               strTempFile = m_TempDir(iRound) & "\" & RsTemp.Fields("pfilename") & ".pdf"
               strP1 = "" & RsTemp.Fields("cpp01")
               strP2 = "" & RsTemp.Fields("cpp02")
               If strP1 <> "" And strP2 <> "" And strTempFile <> "" Then
                  If PUB_GetAttachFile_CPP(strP1, strP2, strTempFile, True) = False Then
                       strP3 = strP3 & "," & strP1 & ":" & strP2 & "，無法下載"
                  End If
               End If
               RsTemp.MoveNext
           Loop
       End If
   Next iRound
   If strP3 <> "" Then Debug.Print Replace(Mid(strP3, 2), ",", vbCrLf)
   MsgBox "OK!"
End Sub


'Added by Lydia 2022/03/11
Private Sub txtEES_GotFocus(Index As Integer)
   TextInverse txtEES(Index)
End Sub

Private Sub txtEES_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
       If txtEES(Index) <> "" Then
           txtEES(1) = "": txtEES(2) = ""
           strExc(0) = "select min(ees02) minno,max(ees02) maxno from editeudcsearch where ees01='" & txtEES(Index) & "' and ees05 is null "
           intI = 1
           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
               If Val("" & RsTemp.Fields("minno")) > 0 Then
                   txtEES(1) = "" & RsTemp.Fields("minno")
               End If
               If Val("" & RsTemp.Fields("maxno")) > 0 Then
                   txtEES(2) = "" & RsTemp.Fields("maxno")
               End If
           End If
           If Trim(txtEES(1) & txtEES(2)) = "" Then
                MsgBox "無資料可供查詢!"
                Command50.Enabled = False
           Else
                Command50.Enabled = True
           End If
       End If
   End If
End Sub

Private Sub txtEES_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii <> 8 And (KeyAscii > 57 Or KeyAscii < 48) Then
      KeyAscii = 0
      Beep
  End If
End Sub

'Added by Lydia 2022/03/11 手動執行->整批查詢造字(參考frmAutoBatchDay.strMenu113) ; 配合111/3/7萬國碼Provider上線，用來普查所有造字區是否有用到資料
Private Sub Command50_Click()
Dim strA1 As String, intA As Integer
Dim strUpd As String
Dim tmpUpd As Variant
Dim rsAD As New ADODB.Recordset
Dim rsQuery As New ADODB.Recordset
Dim intQ As Integer
Dim strBeginTime As String
Dim lngTot As Long
Dim strConList As String, strWd1 As String
Dim strTBname As String
Dim strWhere As String
   
   If Val(txtEES(1)) > Val(txtEES(2)) Then
        MsgBox "起始號碼不可大於終止號碼!!!"
        Exit Sub
   End If
   
   strBeginTime = Format(ServerTime, "000000")
   strA1 = "select ees02, ees03 from editeudcsearch where ees01='" & txtEES(0) & "' and ees05 is null and ees02>=" & Val(txtEES(1)) & " and ees02<=" & Val(txtEES(2))
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strA1)
   If intQ = 1 Then
       rsQuery.MoveFirst
       Do While Not rsQuery.EOF
            lngTot = 0
            strWd1 = Trim(PUB_StringFilter("" & rsQuery.Fields("ees03")))
            strA1 = "select * from EditOraInEudc order by sno asc "
            intA = 1
            Set rsAD = ClsLawReadRstMsg(intA, strA1)
            If intA = 1 Then
                 rsAD.MoveFirst
                 Do While Not rsAD.EOF
                    If "" & rsAD.Fields("tbname") <> "" And "" & rsAD.Fields("updlist") <> "" Then
                        strWhere = ""
                        strTBname = rsAD.Fields("tbname")
                        strA1 = "" & rsAD.Fields("updlist")
                        tmpUpd = Empty
                        tmpUpd = Split(strA1, ",")
                        For intI = 0 To UBound(tmpUpd)
                            '合併 Where
                            strWhere = strWhere & "or instr(" & tmpUpd(intI) & ",'" & strWd1 & "')>0 "
                        Next intI
                        If strWhere <> "" Then
                           strConList = "select count(*) as cnt from " & rsAD.Fields("tbname") & " where (" & Mid(strWhere, 4) & ") "
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strConList)
                           If intI = 1 Then lngTot = lngTot + Val("" & RsTemp.Fields("cnt"))
                        End If
                    End If
                    rsAD.MoveNext
                 Loop
                 strSql = "Update editeudcsearch set ees05=" & lngTot & " where ees01='" & txtEES(0) & "' and ees02=" & rsQuery.Fields("ees02")
                 cnnConnection.Execute strSql
            End If
           rsQuery.MoveNext
       Loop
       strExc(1) = "匯入日期：" & txtEES(0) & String(10, " ") & "起始號碼：" & PUB_StrToStr(txtEES(1), 10, True) & "終止號碼：" & PUB_StrToStr(txtEES(2), 10, True) & vbCrLf & _
                        "執行時間起：" & strBeginTime & vbCrLf & "執行時間止：" & Format(ServerTime, "000000")
       'PUB_SendMail strUserNum, strUserNum, "", "整批查詢造字->完成通知", strExc(1)
       MsgBox strExc(1), vbInformation, "整批查詢造字->完成通知"
   End If
   Set rsQuery = Nothing
   Set rsAD = Nothing
   
End Sub

'Mark by Lydia 2022/03/17  先保留
'Private Sub Command51_Click()
'Dim intA As Integer
'Dim tmpArr As Variant
'Dim rsAD As New ADODB.Recordset
'
'   '----------尋找有?的Table
'   strExc(0) = "select tbname, tbcname ,replace(updlist,',','||') updlist from editoraineudc order by sno "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'       RsTemp.MoveFirst
'       Do While Not RsTemp.EOF
'           strSql = "select count(*) cnt from " & RsTemp.Fields("tbname") & " where instr(" & RsTemp.Fields("updlist") & " , '?') > 0 "
'           intI = 1
'           Set rsAD = ClsLawReadRstMsg(intI, strSql)
'           If intI = 1 Then
'              If Val("" & rsAD.Fields("cnt")) > 0 Then
'                   Debug.Print RsTemp.Fields("tbname") & " (" & RsTemp.Fields("tbcname") & " ) => " & rsAD.Fields("cnt")
'                   Debug.Print "     " & strSql
'              End If
'           End If
'           RsTemp.MoveNext
'       Loop
'   End If
'   '-------比對更換欄位=Pkey的Table
'   strExc(0) = "select * from editoraineudc order by sno "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'       RsTemp.MoveFirst
'       Do While Not RsTemp.EOF
'           strExc(0) = ""
'           tmpArr = Empty
'           tmpArr = Split("" & RsTemp.Fields("pkno"), ",")
'           For intA = 0 To UBound(tmpArr)
'               If Trim("" & tmpArr(intA)) <> "" Then
'                   If InStr(RsTemp.Fields("updlist") & ",", Trim("" & tmpArr(intA)) & ",") > 0 Then
'                       strExc(0) = "Y"
'                   End If
'               End If
'           Next intA
'           tmpArr = Empty
'           tmpArr = Split("" & RsTemp.Fields("updlist"), ",")
'           For intA = 0 To UBound(tmpArr)
'               If Trim("" & tmpArr(intA)) <> "" Then
'                   If InStr(RsTemp.Fields("pkno") & ",", Trim("" & tmpArr(intA)) & ",") > 0 Then
'                       strExc(0) = "Y"
'                   End If
'               End If
'           Next intA
'           If strExc(0) = "Y" Then
'               Debug.Print RsTemp.Fields("tbname") & "    " & RsTemp.Fields("tbcname")
'           End If
'           RsTemp.MoveNext
'       Loop
'   End If
'
'   MsgBox "OK"
'End Sub


'Added by Lydia 2022/07/20
Private Sub Command45_Click()
Dim intQ As Integer
   
   If Trim(txtFM2(0)) = "" Or Trim(txtFM2(1)) = "" Then
        MsgBox "請輸入更換字詞!!"
        txtFM2(0).SetFocus
        Exit Sub
   End If
   If Len(Trim(txtFM2(0))) <> Len(Trim(txtFM2(1))) Then
       If MsgBox("更換前後長度不一致，是否繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
           txtFM2(0).SetFocus
           Exit Sub
       End If
   End If
   
   strExc(0) = "select count(*) cnt from finaltextmap where instr(ftm05||ftm08,'" & Trim(txtFM2(0)) & "')> 0  and ftm07='2' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       intQ = intQ + Val("" & RsTemp.Fields("cnt"))
   End If
   strExc(0) = "select count(*) cnt from UniTextList where instr(utl03,'" & Trim(txtFM2(0)) & "')> 0 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       intQ = intQ + Val("" & RsTemp.Fields("cnt"))
   End If
   
   If intQ = 0 Then
       MsgBox "無定稿資料可以更換!!", vbExclamation
   Else
       If MsgBox("共" & intQ & " 筆定稿資料，是否繼續更換？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
          intQ = 0
          strSql = " update finaltextmap set ftm05=replace(ftm05,'" & Trim(txtFM2(0)) & "','" & Trim(txtFM2(1)) & "'), ftm08=replace(ftm08,'" & Trim(txtFM2(0)) & "','" & Trim(txtFM2(1)) & "')  where instr(ftm05||ftm08,'" & Trim(txtFM2(0)) & "')> 0  and ftm07='2' "
          cnnConnection.Execute strSql, intI
          intQ = intQ + intI
          strSql = " update UniTextList set UTL03=replace(UTL03,'" & Trim(txtFM2(0)) & "','" & Trim(txtFM2(1)) & "')  where instr(UTL03,'" & Trim(txtFM2(0)) & "')> 0 "
          cnnConnection.Execute strSql, intI
          intQ = intQ + intI
          MsgBox "已更換" & intQ & "筆"
       End If
   End If
End Sub

'Added by Lydia 2022/07/20
Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
'/* 已被刪除的記錄
'--6
'select * from lydia_incommemo where im01 not in (select im01 from incommemo) ;
'--0
'select * from lydia_approvalps where aps01 not in (select aps01 from approvalps) ;
'--0
'select * from lydia_debitnoteps where dnps01 not in (select dnps01 from debitnoteps) ;
'--2
'select * from lydia_approvalmemo2 where am01 not in (select am01 from approvalmemo2) ;
'--3
'select * from lydia_fcpempbill where feb01 not in (select feb01 from fcpempbill) ;
'--0
'select * from lydia_npmemo where nm01 not in (select nm01 from npmemo) ;
'*/
End Sub

'Added by Lydia 2023/03/17 補ACS案B類收文
Private Sub Command54_Click()
Dim intP As Integer, intQ As Integer, tmpArr As Variant
Dim strSNo As String, strCP10 As String, strCP64 As String

   '外專備註設定還原6碼
   strExc(0) = "select substr(fa10,1,3) fa10,substr(cu10,1,3) cu10,a.im01 as a0001,a.im03 as a0003,a.im04 as a0004,a.im05 as a0005,b.im03 as b0003,b.im04 as b0004,b.im05 as b0005 " & _
                    "from Lydia_incommemo a , incommemo b,fagent ,customer where a.im01=b.im01(+) and b.im01 is not null " & _
                    "and substr(a.im04||'00',1,8)=fa01(+) and '0'=fa02(+) and substr(a.im05||'00',1,8)=cu01(+) and '0'=cu02(+) order by a.im01 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       With RsTemp
          cnnConnection.BeginTrans
          .MoveFirst
          Do While Not .EOF
               '日本區不還原
               If "" & .Fields("fa10") = "011" Or ("" & .Fields("fa10") & .Fields("cu10") = "011") Then

               Else
                   strExc(2) = ""
                   If "" & .Fields("A0004") <> "" And Mid("" & .Fields("A0004"), 1, 6) = Mid("" & .Fields("B0004"), 1, 6) Then
                       strExc(2) = strExc(2) & ", im04='" & .Fields("A0004") & "'"
                   End If
                   If "" & .Fields("A0005") <> "" And Mid("" & .Fields("A0005"), 1, 6) = Mid("" & .Fields("B0005"), 1, 6) Then
                       strExc(2) = strExc(2) & ", im05='" & .Fields("A0005") & "'"
                   End If
                   If strExc(2) <> "" Then
                      strSql = "Update IncomMemo Set " & Mid(strExc(2), 2) & " Where Im01='" & .Fields("A0001") & "' "
                      cnnConnection.Execute strSql
                   End If
               End If
              .MoveNext
          Loop
          cnnConnection.CommitTrans
       End With
   End If
   MsgBox "OK!"
   
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
       cnnConnection.RollbackTrans
       MsgBox Err.Description
   End If
''Added by Lydia 2023/03/17 補ACS案B類收文
''附件資料是要補B類或C類進度的:
''1. 紅色的實地審查通知1211、驗證通過1006是要補C類進度(C18,C20)，其他都是補B類進度；
''2. 依據資料日期抓當年度之B類或C類流水號；
''3. CP05=CP27=附件資料日期;CP10=附件表頭之案件性質代號欄;CP14=附件系統承辦人欄;
''CP11='90';CP12、CP13抓客戶目前之CU12及CU13；CP20=CP32='N'；CP64=' 西元年/月/日(補資料之系統日)整批補進度資料'。
'
'   strExc(0) = "select c01,c02,c03,c04 ,c21,st02,st04,lc11,cu12,cu13,decode(c05,null,null,to_char(to_date(c05,'mm/dd/yyyy'),'yyyymmdd')) c05 " & _
'                    ",decode(c06,null,null,to_char(to_date(c06,'mm/dd/yyyy'),'yyyymmdd')) c06,decode(c07,null,null,to_char(to_date(c07,'mm/dd/yyyy'),'yyyymmdd')) c07 " & _
'                    ",decode(c08,null,null,to_char(to_date(c08,'mm/dd/yyyy'),'yyyymmdd')) c08,decode(c09,null,null,to_char(to_date(c09,'mm/dd/yyyy'),'yyyymmdd')) c09 " & _
'                    ",decode(c10,null,null,to_char(to_date(c10,'mm/dd/yyyy'),'yyyymmdd')) c10,decode(c11,null,null,to_char(to_date(c11,'mm/dd/yyyy'),'yyyymmdd')) c11 " & _
'                    ",decode(c12,null,null,to_char(to_date(c12,'mm/dd/yyyy'),'yyyymmdd')) c12,decode(c13,null,null,to_char(to_date(c13,'mm/dd/yyyy'),'yyyymmdd')) c13 " & _
'                    ",decode(c14,null,null,to_char(to_date(c14,'mm/dd/yyyy'),'yyyymmdd')) c14,decode(c15,null,null,to_char(to_date(c15,'mm/dd/yyyy'),'yyyymmdd')) c15 " & _
'                    ",decode(c16,null,null,to_char(to_date(c16,'mm/dd/yyyy'),'yyyymmdd')) c16,decode(c17,null,null,to_char(to_date(c17,'mm/dd/yyyy'),'yyyymmdd')) c17 " & _
'                    ",decode(c18,null,null,to_char(to_date(c18,'mm/dd/yyyy'),'yyyymmdd')) c18,decode(c19,null,null,to_char(to_date(c19,'mm/dd/yyyy'),'yyyymmdd')) c19 " & _
'                    ",decode(c20,null,null,to_char(to_date(c20,'mm/dd/yyyy'),'yyyymmdd')) c20 " & _
'                    "from lydia_acs01,staff,lawcase,customer where c21=st01(+) and c01=lc01(+) and c02=lc02(+) and c03=lc03(+) and c04=lc04(+) " & _
'                    "and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
'intI = 1
'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'If intI = 1 Then
'    With RsTemp
'        tmpArr = Split("208,1013,2091,2092,2093,2093,2093,211,213,215,216,218,221,1211,220,1006", ",")
'        .MoveFirst
'        Do While Not .EOF
'            cnnConnection.BeginTrans
'               For intP = 5 To 20
'                  If "" & .Fields("C" & Format(intP, "00")) <> "" Then
'                      If InStr("18,20,", Format(intP, "00")) > 0 Then
'                         strSNo = AutoNo("C", 6)
'                      Else
'                         strSNo = AutoNo("B", 6)
'                      End If
'                      strCP10 = "": strCP64 = ""
'                      '208,1013,2091,2092,2093,2093,2093,211,213,215,216,218,221,1211,220,1006
'                      'C05 , C06, C07, C08, C09, C10, C11, C12, C15, C16, C18, C19, C22, C24, C25, C27
'                      '培訓課程-內部稽核教育訓練(2093)  培訓課程-營業秘密(2093) 培訓課程-新制度文件培訓(2093)
'                      strCP10 = tmpArr(intP - 5)
'                      Select Case intP
'                         Case 9: strCP64 = ":培訓課程-內部稽核教育訓練"
'                         Case 10: strCP64 = ":培訓課程-營業秘密"
'                         Case 11: strCP64 = ":培訓課程-新制度文件培訓"
'                      End Select
'                      If strCP10 <> "" Then
'                          strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp27,cp20,cp32,cp64) " & _
'                                "values ('" & .Fields("c01") & "', '" & .Fields("c02") & "','" & .Fields("c03") & "','" & .Fields("c04") & "','" & .Fields("C" & Format(intP, "00")) & "','" & strSNo & "' " & _
'                                ", '" & strCP10 & "' , '90' , '" & .Fields("cu12") & "', '" & .Fields("cu13") & "', '" & .Fields("c21") & "','" & .Fields("C" & Format(intP, "00")) & "','N','N','" & ChangeTStringToTDateString(strSrvDate(1)) & "整批補進度資料" & strCP64 & "' )"
'                          'Debug.Print strSql
'                          cnnConnection.Execute strSql
'                      End If
'                  End If
'               Next intP
'            cnnConnection.CommitTrans
'            .MoveNext
'        Loop
'    End With
'End If
'MsgBox "OK!"
'Exit Sub
'
'ErrHandle:
'   cnnConnection.RollbackTrans
'   If Err.Number <> 0 Then
'      MsgBox Err.Description
'   End If

End Sub

'Added by Lydia 2023/08/15
Private Sub Text6_GotFocus(Index As Integer)
   TextInverse Text6(Index)
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String
   Select Case Index
      Case 1 '代理人
         Label11(Index) = ""
         If Trim(Text6(Index)) <> "" Then
            Text6(Index) = ChangeCustomerL(Text6(Index))
            If ClsPDGetAgent(Text6(Index), strTemp) Then
               Label11(Index).Caption = strTemp
            Else
               Cancel = True
            End If
         End If
      Case 2 '申請人
         Label11(Index) = ""
         If Trim(Text6(Index)) <> "" Then
            Text6(Index) = ChangeCustomerL(Text6(Index))
            If ClsPDGetCustomer(Text6(Index), strTemp) Then
               Label11(Index).Caption = strTemp
            Else
               Cancel = True
            End If
         End If
      Case 3 '本所案號
         If Trim(Text6(Index)) <> "" Then
            strExc(0) = "select PA01||PA02||PA03||PA04 from patent where " & ChgPatent(Text6(Index))
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               MsgBox "本所案號輸入錯誤!", vbExclamation
               Cancel = True
            Else
               Text6(Index) = RsTemp(0)
            End If
         End If
      Case 4 '員工編號
         Label11(Index) = ""
         If Trim(Text6(Index)) <> "" Then
            strTemp = GetStaffName(Trim(Text6(Index)))
            If strTemp <> "" Then
               Label11(Index).Caption = strTemp
            Else
               Cancel = True
            End If
         End If
   End Select
End Sub

'Added by Lydia 2023/08/15 利益衝突權限檢查
Private Sub cmdChkCufa_Click()
Dim intErr As Integer, tmpBol As Boolean

   If Trim(Text6(4)) = "" Or Label11(4) = "" Then
      MsgBox "請輸入正確的員工編號！", vbExclamation
      intErr = 4
      GoTo EXITSUB
   End If
   If Trim(Text6(1)) <> "" Then
      Call Text6_Validate(1, tmpBol)
      If tmpBol = True Then
         MsgBox "請輸入正確的代理人編號！", vbExclamation
         intErr = 1
         GoTo EXITSUB
      End If
   End If
   If Trim(Text6(2)) <> "" Then
      Call Text6_Validate(2, tmpBol)
      If tmpBol = True Then
         MsgBox "請輸入正確的申請人編號！", vbExclamation
         intErr = 2
         GoTo EXITSUB
      End If
   End If
   If Trim(Text6(3)) <> "" Then
      Call Text6_Validate(3, tmpBol)
      If tmpBol = True Then
         intErr = 3
         GoTo EXITSUB
      End If
   End If
   
   If Trim(Text6(1) & Text6(2) & Text6(3)) = "" Then
      MsgBox "請輸入代理人、申請人或本所案號為條件！", vbExclamation
      intErr = 1
      GoTo EXITSUB
   End If
   If Trim(Text6(3)) <> "" Then
       strExc(0) = Text6(3)
       Call ChgCaseNo(strExc(0), strExc)
       If strExc(1) <> "" And Len(strExc(2)) = 6 Then
          strSql = "select pa75, nvl(fa04,nvl(fa05,fa06)) as fname, pa26, nvl(cu04,nvl(cu05,cu06)) cname " & _
                   "from patent,fagent,customer where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' " & _
                   "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strSql)
          If intI = 1 Then
             Text6(1) = "" & RsTemp.Fields("pa75")
             Label11(1) = "" & RsTemp.Fields("fname")
             Text6(2) = "" & RsTemp.Fields("pa26")
             Label11(2) = "" & RsTemp.Fields("cname")
          Else
             intErr = 3
             GoTo EXITSUB
          End If
       Else
          intErr = 3
          GoTo EXITSUB
       End If
   End If
   
   If intErr = 0 Then
      If (InStr(XY特殊權限範圍, Left(Text6(1), 8)) > 0 And Trim(Text6(1)) <> "") Or (InStr(XY特殊權限範圍, Left(Text6(2), 8)) > 0 And Trim(Text6(2)) <> "") Then
         If Text6(3) <> "" And strExc(1) <> "" And Len(strExc(2)) = 6 Then
            strExc(5) = strExc(1)
            strExc(6) = strExc(1) & strExc(2) & strExc(3) & strExc(4)
         Else
            strExc(5) = "ALL"
            strExc(6) = "FCP000001000"
         End If
         If strExc(5) = "ALL" Then
            strExc(5) = GetAllSysKind(, Trim(Text6(0)))
         End If
         If Trim(Text6(1)) <> "" Then
            strExc(7) = ChangeCustomerL(Text6(1))
         Else
            strExc(7) = "Y00000000"
         End If
         If Trim(Text6(2)) <> "" Then
            strExc(8) = ChangeCustomerL(Text6(2)) & ",,,,"
         Else
            strExc(8) = "X00000000,,,,"
         End If
         If PUB_ChkCufaByCase(Me.Name, strExc(5), strExc(6), strExc(8), strExc(7), Trim(Text6(4))) = False Then
            MsgBox Text6(4) & " " & Label11(4) & vbCrLf & "沒有權限！", vbInformation
         Else
            MsgBox Text6(4) & " " & Label11(4) & vbCrLf & "有利益衝突案件(限閱案件)的權限！", vbInformation
         End If
      Else
         MsgBox "輸入的代理人、申請人並非利益衝突案件(限閱案件)！", vbInformation
      End If
   End If
   
EXITSUB:
   If intErr > 0 Then
      Text6(intErr).SetFocus
      Text6_GotFocus intErr
   End If
End Sub

'Added by Lydia 2023/09/20
Private Sub Text7_GotFocus(Index As Integer)
   TextInverse Text7(Index)
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index < 2 Then
      If KeyAscii <> 8 And (KeyAscii > 57 Or KeyAscii < 48) Then
          KeyAscii = 0
          Beep
      End If
   Else
      If KeyAscii <> 89 And KeyAscii <> 8 Then 'Y/ null
          KeyAscii = 0
          Beep
      End If
   End If
End Sub

Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   Cancel = False
   If Index < 2 Then
      If IsEmptyText(Text7(Index)) = False Then
         If CheckIsTaiwanDate(Text7(Index), False) = False Then
            Cancel = True
            MsgBox "請輸入正確日期(民國年月日)!", vbCritical
            GoTo EXITSUB
         End If
         
         '不能大於系統日
         If DBDATE(Text7(Index)) > strSrvDate(1) Then
            Cancel = True
            MsgBox "輸入日期不能大於系統日!", vbCritical
            GoTo EXITSUB
         End If
      End If
   End If
   Exit Sub
   
EXITSUB:
   Text7(Index).SetFocus
   Call Text7_GotFocus(Index)
End Sub

'Added by Lydia 2023/09/20 網中-查名單
Private Sub Command56_Click()
Dim strGrp As String, strTQF00 As String, strTQFList As String, strTMQ01List As String
Dim intP As Integer, strP1 As String, strErrList As String
Dim rsPD As New ADODB.Recordset
Dim bolReset As Boolean
Dim strTemp(1 To 4) As String
Dim tmpArrB As Variant, tmpArrA As Variant, intA As Integer, intB As Integer 'Added by Lydia 2025/05/28
   
   If Val(Text7(0)) = 0 Or Val(Text7(1)) = 0 Then
      MsgBox "請輸入起始日期和終止日期!", vbExclamation
      Exit Sub
   ElseIf Val(Text7(0)) > Val(Text7(1)) Then
      MsgBox "請輸入起始日期不可大於終止日期!", vbExclamation
      Exit Sub
   End If
      
   Screen.MousePointer = vbHourglass
   Debug.Print "Start :" & Format(ServerTime, "000000")
   Me.Enabled = False
   
   'Added by Lydia 2025/05/28 重新匯出訓練資料：委查日1120501~1140430，圖形查名結果為”相同、近似”包含”相同本所案、近似本所案”，尚有查名結果附件(TS.PDF)；
                        '下載附件：原始查名圖(JPG檔)、查名結果附件(TS.PDF)
   '匯出Excel: SELECT tmq01 AS 查名單號,tqd05 AS 組群,tmq24 AS 圖形路徑, tqd06n AS 結果, tqf00 AS 圖形附件, tqf12list AS 查名附件 FROM lydia_tmq2tqf114 order by tmq01
   'DROP TABLE LYDIA_TMQ2TQF114;
   'CREATE TABLE LYDIA_TMQ2TQF114 (
   'TMQ18 VARCHAR2(10 CHAR),
   'TMQ01 VARCHAR2(9 CHAR),
   'TQD05 VARCHAR2(6 CHAR),
   'TMQ24 VARCHAR2(100 CHAR),
   'TQD06 VARCHAR2(1 CHAR),
   'TQD06N VARCHAR2(6 CHAR),
   'TQF00 VARCHAR2(20 CHAR),
   'TQF12LIST VARCHAR2(200 CHAR));
   strP1 = "select tmq18,tmq01,tqd05,tmq24,tqd06,decode(tqd06," & PUB_GetTMQans("3", True) & ") tqd06n,b.tqf12 tqf00,b.tqf01 b01,b.tqf02 b02,b.tqf03 b03,b.tqf04 b04 , a.tqf12,a.tqf01 a01,a.tqf02 a02,a.tqf03 a03,a.tqf04 a04 " & _
           " from trademarkquery,tmqdetail,tmqfile a, tmqfile b where tmq05>=" & DBDATE(Text7(0)) & " and tmq05<=" & DBDATE(Text7(1)) & " and tmq09=1 and tmq01=tqd02(+) and tqd06 in ('2','3','4','5')" & _
           " and tmq01=a.tqf02(+) and upper(a.tqf12) like '%.TS.PDF' and tmq18=b.tqf01(+) and b.tqf03='" & TMQ_附件F02 & "' and b.tqf04='" & TMQ_附件F04 & "' and upper(b.tqf12) like '%.JPG'"
   strP1 = strP1 & " order by tmq05,tmq01,tqf12"
   intP = 1
   Set rsPD = ClsLawReadRstMsg(intP, strP1)
   If intP = 1 Then
      cnnConnection.Execute "TRUNCATE TABLE LYDIA_TMQ2TQF114"
      rsPD.MoveFirst
      Do While Not rsPD.EOF
         If strGrp <> "" & rsPD.Fields("tmq01") Then
            '(修正)查名路徑
            strExc(1) = Replace(Replace("" & rsPD.Fields("tmq24"), " 、 ", ","), "、", ",")
            strExc(1) = Replace(Replace(strExc(1), "(", ""), ")", "") '拿掉()
            strExc(1) = Replace(Replace(strExc(1), ", ", ","), " ,", ",") '拿掉,加空白
            strExc(1) = Replace(strExc(1), " ", ",") '空白改成,
            If InStr(strExc(1), "/") > 0 Then
               strExc(2) = "": strExc(3) = ""
               tmpArrA = Empty
               tmpArrA = Split(strExc(1), ",")
               For intA = 0 To UBound(tmpArrA)
                  If Trim(tmpArrA(intA)) <> "" Then
                     strExc(3) = ""
                     tmpArrB = Empty
                     tmpArrB = Split(tmpArrA(intA), "/")
                     If UBound(tmpArrB) = 0 Then
                        If Len(Trim(tmpArrB(0))) = 5 Then
                           strExc(2) = strExc(2) & "," & Mid(Trim(tmpArrB(0)), 1, 2) & "-" & Mid(Trim(tmpArrB(0)), 3, 1) & "-" & Mid(Trim(tmpArrB(0)), 4, 2)
                        Else
                           strExc(2) = strExc(2) & "," & Trim(tmpArrB(0))
                        End If
                     Else
                        For intB = 0 To UBound(tmpArrB)
                           If Len(Trim(tmpArrB(intB))) = 5 Then
                              strExc(3) = Mid(Trim(tmpArrB(intB)), 1, 3)
                              strExc(2) = strExc(2) & "," & Trim(tmpArrB(intB))
                           ElseIf Len(Trim(tmpArrB(intB))) = 7 Then
                              strExc(3) = Mid(Trim(tmpArrB(intB)), 1, 5)
                              strExc(2) = strExc(2) & "," & Trim(tmpArrB(intB))
                           ElseIf Len(Trim(tmpArrB(intB))) <= 2 And strExc(3) <> "" Then
                              strExc(2) = strExc(2) & "," & strExc(3) & Trim(tmpArrB(intB))
                           End If
                        Next intB
                     End If
                     strExc(1) = Mid(strExc(2), 2)
                  End If
               Next intA
            Else
               strExc(2) = "": strExc(3) = ""
               tmpArrA = Empty
               tmpArrA = Split(strExc(1), ",")
               For intA = 0 To UBound(tmpArrA)
                  If Trim(tmpArrA(intA)) <> "" Then
                     If Len(Trim(tmpArrA(intA))) = 5 Then
                        strExc(3) = Mid(Trim(tmpArrA(intA)), 1, 2) & "-" & Mid(Trim(tmpArrA(intA)), 3, 1) & "-"
                        strExc(2) = strExc(2) & "," & Mid(Trim(tmpArrA(intA)), 1, 2) & "-" & Mid(Trim(tmpArrA(intA)), 3, 1) & "-" & Mid(Trim(tmpArrA(intA)), 4, 2)
                     ElseIf Len(Trim(tmpArrA(intA))) = 7 Then
                        strExc(3) = Mid(Trim(tmpArrA(intA)), 1, 5)
                        strExc(2) = strExc(2) & "," & Trim(tmpArrA(intA))
                     ElseIf Len(Trim(tmpArrA(intA))) <= 2 And strExc(3) <> "" Then
                        strExc(2) = strExc(2) & "," & strExc(3) & Trim(tmpArrA(intA))
                     End If
                  End If
               Next intA
               If strExc(2) <> "" Then
                  strExc(1) = Mid(strExc(2), 2)
               End If
            End If
            strSql = "INSERT INTO LYDIA_TMQ2TQF114 (TMQ18,TMQ01,TQD05,TMQ24,TQD06,TQD06N) VALUES ('" & rsPD.Fields("tmq18") & "','" & rsPD.Fields("tmq01") & "','" & rsPD.Fields("tqd05") & "','" & strExc(1) & "','" & rsPD.Fields("tqd06") & "','" & rsPD.Fields("tqd06n") & "') "
            cnnConnection.Execute strSql
            strTQFList = ""
            '圖形查名的附件
            strTQF00 = Mid("" & rsPD.Fields("tqf00"), InStrRev("" & rsPD.Fields("tqf00"), "/") + 1)
            If Text7(2) = "Y" Then
               If Command56_Sub2("" & rsPD.Fields("tmq01"), "" & rsPD.Fields("b01"), "" & rsPD.Fields("b02"), "" & rsPD.Fields("b03"), "" & rsPD.Fields("b04"), "" & rsPD.Fields("tqf00"), strErrList) = True Then
                  strSql = "Update Lydia_tmq2tqf114 set tqf00='" & strTQF00 & "' where tmq01='" & rsPD.Fields("tmq01") & "' "
                  cnnConnection.Execute strSql
               End If
            End If
         End If
         '查名結果附件
         If Text7(2) = "Y" And InStr(strTQFList & ",", Mid("" & rsPD.Fields("tqf12"), InStrRev("" & rsPD.Fields("tqf12"), "/") + 1)) = 0 Then
            If Command56_Sub2("" & rsPD.Fields("tmq01"), "" & rsPD.Fields("a01"), "" & rsPD.Fields("a02"), "" & rsPD.Fields("a03"), "" & rsPD.Fields("a04"), "" & rsPD.Fields("tqf12"), strErrList) = True Then
               strTQFList = strTQFList & "," & Mid("" & rsPD.Fields("tqf12"), InStrRev("" & rsPD.Fields("tqf12"), "/") + 1)
               strSql = "Update Lydia_tmq2tqf114 set tqf12list='" & Mid(strTQFList, 2) & "' where tmq01='" & rsPD.Fields("tmq01") & "' "
               cnnConnection.Execute strSql
            End If
         End If
         strGrp = "" & rsPD.Fields("tmq01")
         rsPD.MoveNext
      Loop
   End If
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   Debug.Print "End   :" & Format(ServerTime, "000000")
   If strErrList <> "" Then
      PUB_SendMail strUserNum, strUserNum, "", "網中-查名單無法下載檔案", vbCrLf & Mid(strErrList, 2)
   End If
   Set rsPD = Nothing
   MsgBox "OK!"
   Exit Sub
   'end 2025/05/28
   

'處理條件:
'1. 期間：111.01.01至112.08.31
'2. 針對案件性質「申請101」之已發文商標案件,限查名單有圖形查名
'3. 案件名稱含有「及圖」或「標章」或「圖形」(併入第2點)
'4. 提供申請案號
'5. 提供商標圖檔
'6. 提取查名結果中之圖形查名單於附件區的所有附件

'----先抓一些例子給網中
   'CREATE TABLE LYDIA_TMQ2TQF (
   'CP01 VARCHAR2(3 CHAR),
   'CP02 VARCHAR2(6 CHAR),
   'CP03 VARCHAR2(1 CHAR),
   'CP04 VARCHAR2(2 CHAR),
   'CP09 VARCHAR2(9 CHAR),
   'TMQ01L VARCHAR2(400 CHAR),
   'TQF00 VARCHAR2(200 CHAR),  ---圖形附件的檔名
   'TQFLIST VARCHAR2(600 CHAR));   ---所有查名附件(TS.PDF)的檔名,用,區隔
'--------------------
   '匯出SQL: select tm12 as 申請案號,tm05 as 名稱,tqf00 as 圖形附件,tqflist as 查名附件 from lydia_tmq2tqf,trademark
   'where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)
   'P.S.因為網中要求有查名附件TS.PDF, 所以抓還未刪除TS.PDF的案件
   strP1 = "select tm12,cp01,cp02,cp03,cp04,cp09,tmq01,tmq18 " & _
           "From caseprogress, trademark, tmqcasemap, trademarkquery " & _
           "where cp01='T' and cp05>=" & DBDATE(Text7(0)) & " and cp05<=" & DBDATE(Text7(1)) & " and cp10='101' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
           "and tm10='000' and cp09=tqc02(+) and tqc03=tmq01(+) and nvl(tmq09,0) > 0 and tm12 is not null " & _
           "and tmq01 in (select tqf02 from tmqfile where tqf02<>'" & TMQ_附件F02 & "' and tqf04<>'" & TMQ_附件F04 & "') " & _
           "order by tm12,tmq01 "
   intP = 1
   Set rsPD = ClsLawReadRstMsg(intP, strP1)
   If intP = 1 Then
      rsPD.MoveFirst
      If bolReset = False Then
         cnnConnection.Execute "TRUNCATE TABLE LYDIA_TMQ2TQF"
         bolReset = True
      End If
      Do While Not rsPD.EOF
         If strGrp <> "" & rsPD.Fields("cp09") Then
            If strGrp <> "" Then
               strSql = "INSERT INTO LYDIA_TMQ2TQF (CP01,CP02,CP03,CP04,CP09,TMQ01L,TQF00,TQFLIST) VALUES ('" & strTemp(1) & "', '" & strTemp(2) & "', '" & strTemp(3) & "', '" & strTemp(4) & "', '" & strGrp & "' " & _
                        ", '" & Mid(strTMQ01List, 2) & "', '" & Mid(strTQF00, 2) & "','" & Mid(strTQFList, 2) & "')"
               cnnConnection.Execute strSql
            End If
            strGrp = ""
            strTMQ01List = ""
            strTQF00 = ""
            strTQFList = ""
         End If
         If strGrp = "" Then
            '圖形附件的檔名
            strExc(0) = "select * from tmqfile where tqf01='" & rsPD.Fields("tmq18") & "' and tqf02='" & TMQ_附件F02 & "' and tqf04='" & TMQ_附件F04 & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTQF00 = strTQF00 & "," & Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1)
               If Text7(2) = "Y" Then
                  Call Command56_Sub("" & rsPD.Fields("tm12"), "" & RsTemp.Fields("tqf01"), "" & RsTemp.Fields("tqf02"), "" & RsTemp.Fields("tqf03"), "" & RsTemp.Fields("tqf04"), Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1), "" & rsPD.Fields("cp01"), "" & rsPD.Fields("cp02"), "" & rsPD.Fields("cp09"), strErrList)
               End If
            End If
         End If
         '圖形附件的檔名
         strExc(0) = "select * from tmqfile where tqf02='" & rsPD.Fields("tmq01") & "' and tqf02<>'" & TMQ_附件F02 & "' and tqf04<>'" & TMQ_附件F04 & "' " & _
                     "order by tqf03, tqf04 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strTQFList = strTQFList & "," & Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1)
               If Text7(2) = "Y" Then
                  Call Command56_Sub("" & rsPD.Fields("tm12"), "" & RsTemp.Fields("tqf01"), "" & RsTemp.Fields("tqf02"), "" & RsTemp.Fields("tqf03"), "" & RsTemp.Fields("tqf04"), Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1), "" & rsPD.Fields("cp01"), "" & rsPD.Fields("cp02"), "" & rsPD.Fields("cp09"), strErrList)
               End If
               RsTemp.MoveNext
            Loop
         End If
         strGrp = "" & rsPD.Fields("cp09")
         strTemp(1) = "" & rsPD.Fields("cp01")
         strTemp(2) = "" & rsPD.Fields("cp02")
         strTemp(3) = "" & rsPD.Fields("cp03")
         strTemp(4) = "" & rsPD.Fields("cp04")
         strTMQ01List = strTMQ01List & "," & rsPD.Fields("tmq01")
         rsPD.MoveNext
         Sleep 100
      Loop
   End If
   
   '抓已刪除對照檔
   If strGrp <> "" Then
      strSql = "INSERT INTO LYDIA_TMQ2TQF (CP01,CP02,CP03,CP04,CP09,TMQ01L,TQF00,TQFLIST) VALUES ('" & strTemp(1) & "', '" & strTemp(2) & "', '" & strTemp(3) & "', '" & strTemp(4) & "', '" & strGrp & "' " & _
               ", '" & Mid(strTMQ01List, 2) & "', '" & Mid(strTQF00, 2) & "','" & Mid(strTQFList, 2) & "')"
      cnnConnection.Execute strSql
   End If
   strGrp = ""
   strTMQ01List = ""
   strTQF00 = ""
   strTQFList = ""
   strP1 = "select tm12,cp01,cp02,cp03,cp04,cp09,tmq01,tmq18 " & _
           "From caseprogress, trademark, trademarkquery " & _
           "where cp01='T' and cp05>=" & DBDATE(Text7(0)) & " and cp05<=" & DBDATE(Text7(1)) & " and cp10='101' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
           "and tm10='000' and cp09=tmq01(+) and nvl(tmq09,0) > 0 and tm12 is not null " & _
           "and tmq01 in (select tqf02 from tmqfile where tqf02<>'" & TMQ_附件F02 & "' and tqf04<>'" & TMQ_附件F04 & "') " & _
           "and cp09 not in (select tqc02 from tmqcasemap where tqc02 is not null) order by tm12,tmq01 "
   intP = 1
   Set rsPD = ClsLawReadRstMsg(intP, strP1)
   If intP = 1 Then
      rsPD.MoveFirst
      If bolReset = False Then
         cnnConnection.Execute "TRUNCATE TABLE LYDIA_TMQ2TQF"
         bolReset = True
      End If
      Do While Not rsPD.EOF
         If strGrp <> "" & rsPD.Fields("cp09") Then
            If strGrp <> "" Then
               strSql = "INSERT INTO LYDIA_TMQ2TQF (CP01,CP02,CP03,CP04,CP09,TMQ01L,TQF00,TQFLIST) VALUES ('" & strTemp(1) & "', '" & strTemp(2) & "', '" & strTemp(3) & "', '" & strTemp(4) & "', '" & strGrp & "' " & _
                        ", '" & Mid(strTMQ01List, 2) & "', '" & Mid(strTQF00, 2) & "','" & Mid(strTQFList, 2) & "')"
               cnnConnection.Execute strSql
            End If
            strGrp = ""
            strTQF00 = ""
            strTQFList = ""
         End If
         If InStr(strTQF00 & ",", "" & rsPD.Fields("tmq18")) = 0 Then
            '圖形附件的檔名
            strExc(0) = "select * from tmqfile where tqf01='" & rsPD.Fields("tmq18") & "' and tqf02='" & TMQ_附件F02 & "' and tqf04='" & TMQ_附件F04 & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTQF00 = strTQF00 & "," & Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1)
               If Text7(2) = "Y" Then
                  Call Command56_Sub("" & rsPD.Fields("tm12"), "" & RsTemp.Fields("tqf01"), "" & RsTemp.Fields("tqf02"), "" & RsTemp.Fields("tqf03"), "" & RsTemp.Fields("tqf04"), Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1), "" & rsPD.Fields("cp01"), "" & rsPD.Fields("cp02"), "" & rsPD.Fields("cp09"), strErrList)
               End If
            End If
         End If
         '圖形附件的檔名
         strExc(0) = "select * from tmqfile where tqf02='" & rsPD.Fields("tmq01") & "' and tqf02<>'" & TMQ_附件F02 & "' and tqf04<>'" & TMQ_附件F04 & "' " & _
                     "order by tqf03, tqf04 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strTQFList = strTQFList & "," & Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1)
               If Text7(2) = "Y" Then
                  Call Command56_Sub("" & rsPD.Fields("tm12"), "" & RsTemp.Fields("tqf01"), "" & RsTemp.Fields("tqf02"), "" & RsTemp.Fields("tqf03"), "" & RsTemp.Fields("tqf04"), Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1), "" & rsPD.Fields("cp01"), "" & rsPD.Fields("cp02"), "" & rsPD.Fields("cp09"), strErrList)
               End If
               RsTemp.MoveNext
            Loop
         End If
         strGrp = "" & rsPD.Fields("cp09")
         strTMQ01List = strTMQ01List & "," & rsPD.Fields("tmq01")
         rsPD.MoveNext
         Sleep 100
      Loop
   End If
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   Debug.Print "End   :" & Format(ServerTime, "000000")
   If strGrp <> "" Then
      strSql = "INSERT INTO LYDIA_TMQ2TQF (CP01,CP02,CP03,CP04,CP09,TMQ01L,TQF00,TQFLIST) VALUES ('" & strTemp(1) & "', '" & strTemp(2) & "', '" & strTemp(3) & "', '" & strTemp(4) & "', '" & strGrp & "' " & _
               ", '" & Mid(strTMQ01List, 2) & "', '" & Mid(strTQF00, 2) & "','" & Mid(strTQFList, 2) & "')"
      cnnConnection.Execute strSql
   End If
   strGrp = ""
   strTMQ01List = ""
   strTQF00 = ""
   strTQFList = ""
   If strErrList <> "" Then
      PUB_SendMail strUserNum, strUserNum, "", "網中-查名單無法下載檔案", vbCrLf & Mid(strErrList, 2)
   End If
   Set rsPD = Nothing
   MsgBox "OK!"
   
   
'-----112/8/1 網中要的資料先丟LYDIA_TMQ;核駁前先行通知的圖形近似條款為AH30的案件
'  strSql = "select a.*,b.tmq01,b.tmq24 from lydia_tmq a, trademarkquery b where a.cp09=b.tmq21(+) order by a.cp09 asc "
'  strExc(0) = "": strExc(1) = "": strExc(2) = ""
'  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'  intI = 1
'  If intI = 1 Then
'     RsTemp.MoveFirst
'     Do While Not RsTemp.EOF
'        If strExc(0) <> "" & RsTemp.Fields("cp09") Then
'           If strExc(0) <> "" And (strExc(1) <> "" Or strExc(2) <> "") Then
'              strSql = "Update lydia_tmq set tmq24='" & ChgSQL(strExc(1)) & "', nolist='" & ChgSQL(strExc(2)) & "' where cp09='" & strExc(0) & "' "
'              cnnConnection.Execute strSql
'           End If
'           strExc(0) = "" & RsTemp.Fields("cp09")
'           strExc(1) = "" & RsTemp.Fields("tmq24")
'           strExc(2) = "" & RsTemp.Fields("tmq01")
'        End If
'        '記錄查名路徑
'        If InStr(strExc(1), "" & RsTemp.Fields("tmq24")) = 0 And "" & RsTemp.Fields("tmq24") <> "" Then
'           strExc(1) = strExc(1) & IIf(strExc(1) <> "", ",", "") & RsTemp.Fields("tmq24")
'        End If
'        '記錄查名單號
'        If InStr(strExc(2), "" & RsTemp.Fields("tmq01")) = 0 And "" & RsTemp.Fields("tmq01") <> "" Then
'           strExc(2) = strExc(2) & IIf(strExc(2) <> "", ",", "") & RsTemp.Fields("tmq01")
'        End If
'        RsTemp.MoveNext
'     Loop
'     If strExc(0) <> "" And (strExc(1) <> "" Or strExc(2) <> "") Then
'        strSql = "Update lydia_tmq set tmq24='" & ChgSQL(strExc(1)) & "', nolist='" & ChgSQL(strExc(2)) & "' where cp09='" & strExc(0) & "' "
'        cnnConnection.Execute strSql
'     End If
'  End If
'  MsgBox "OK"
'Exit Sub

End Sub

'Added by Lydia 2023/09/20 網中-查名單:處理下載檔案
Private Sub Command56_Sub(ByVal pTM12 As String, ByVal pTQF01 As String, ByVal pTQF02 As String, ByVal pTQF03 As String, ByVal pTQF04 As String, _
    ByVal pTQF12 As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP09 As String, ByRef pErrMsg As String)
Dim strDefPath As String, strFuName As String
   
   strDefPath = App.path & "\" & strUserNum & "\" & Left(pTM12, 3) & "\" & pTM12
   Pub_ChkExcelPath strDefPath
   
   If Dir(strDefPath, vbDirectory) <> "" Then
      strFuName = strDefPath & "\" & pTQF12
      If Dir(strFuName) = "" Then
         If PUB_TMQGetAFile("", strFuName, pTQF01, pTQF02, pTQF03, pTQF04, "") = False Then
            pErrMsg = pErrMsg & ",無法下載檔案: " & pTQF12 & "(" & pCP01 & "-" & pCP02 & "  收文號:" & pCP09 & "  申請案號:" & pTM12 & ")"
         End If
      End If
   End If
End Sub

'Added by Lydia 2025/05/28 網中-查名單:處理下載檔案(114年匯出訓練資料)
Private Function Command56_Sub2(ByVal pTMQ01 As String, ByVal pTQF01 As String, ByVal pTQF02 As String, ByVal pTQF03 As String, ByVal pTQF04 As String, ByVal pTQF12 As String, ByRef pErrMsg As String) As String
Dim strDefPath As String, strFuName As String
   
   strDefPath = App.path & "\" & strUserNum & "\" & Left(pTMQ01, 5) & "\" & pTMQ01
   Pub_ChkExcelPath strDefPath
   
   Command56_Sub2 = False
   If Dir(strDefPath, vbDirectory) <> "" Then
      pTQF12 = Mid(pTQF12, InStrRev(pTQF12, "/") + 1)
      strFuName = strDefPath & "\" & pTQF12
      If Dir(strFuName) = "" Then
         If PUB_TMQGetAFile("", strFuName, pTQF01, pTQF02, pTQF03, pTQF04, "") = False Then
            pErrMsg = pErrMsg & ",無法下載檔案: " & pTQF12
         Else
            Command56_Sub2 = True
         End If
      End If
   End If
End Function

'Added by Lydia 2023/09/24 查名附件大於20M
Private Sub Command57_Click()
Dim strDefPath As String, strNowFile As String
Dim strCode(1 To 4) As String, strTQF11 As String, strTQF13 As String
Dim sf1 As Integer

   'Memo by Lydia 2023/10/02 下載查名附件大於20M
'   strSql = "select * from tmqfile where tqf06 > 21000000 " ' and tqf02 like 'HB201%' and tqf08='B0002' "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      Screen.MousePointer = vbHourglass
'      Me.Enabled = False
'      RsTemp.MoveFirst
'      Do While Not RsTemp.EOF
'         Call Command56_Sub("ALL20", "" & RsTemp.Fields("TQF01"), "" & RsTemp.Fields("TQF02"), "" & RsTemp.Fields("TQF03"), "" & RsTemp.Fields("TQF04"), Mid("" & RsTemp.Fields("tqf12"), InStrRev("" & RsTemp.Fields("tqf12"), "/") + 1), "T", "Test", "A000", strExc(1))
'         RsTemp.MoveNext
'      Loop
'      Screen.MousePointer = vbDefault
'      Me.Enabled = True
'   End If
   
   'cnnConnection.Execute "delete from lydia_1121002" '1234
   
   strDefPath = App.path & "\" & strUserNum & "\A001"
   strNowFile = Dir(strDefPath & "\*.PDF")
   Do While strNowFile <> ""
      strCode(1) = ""
      strTQF11 = "": strTQF13 = ""
      strCode(2) = Mid(strNowFile, 1, 9)
      strCode(3) = Mid(strNowFile, 10, 1)
      strCode(4) = Mid(strNowFile, 11, 2)
      strSql = "select tmq18,b.* from trademarkquery,tmqfile b where tmq01='" & strCode(2) & "' and tmq01=tqf02(+) and instr(tqf12,'" & UCase(strNowFile) & "') > 0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strCode(1) = "" & RsTemp.Fields("tmq18")
         strTQF11 = "" & RsTemp.Fields("tqf11")
         strTQF13 = "" & RsTemp.Fields("tqf13")
      End If
      If strCode(1) <> "" Then
         If PUB_TMQAFileSave(strCode(1), strCode(2), strCode(3), strCode(4), UCase(TMQ_查名作業 & ".pdf"), strDefPath & "\" & strNowFile, "N") = True Then
            strSql = "update tmqfile set tqf11=" & CNULL(strTQF11) & ", tqf13=" & CNULL(strTQF13) & " where TQF01='" & strCode(1) & "' AND TQF02='" & strCode(2) & "' AND TQF03='" & strCode(3) & "' AND TQF04='" & strCode(4) & "' "
            cnnConnection.Execute strSql
            strSql = "insert into lydia_1121002 values('" & strCode(1) & "', '" & strCode(2) & "', '" & strCode(3) & "', '" & strCode(4) & "', '" & strNowFile & "') "
            cnnConnection.Execute strSql
         End If
      End If
      strNowFile = Dir()
   Loop
                
   MsgBox "OK !"
End Sub

'Added by Lydia 2025/03/06
Private Sub Command59_Click()
Dim tmpArr As Variant
Dim intA As Integer, intB As Integer
   
'1. 僅限CFT案件，注意多類別TM09案件，例CFT-005865、CFT-017593。
'2. 商品檔TMGOODS的商品類別TG05也要一併改，缺TMGOODS的案件不必補TMGOODS。
'3. 為免將來後悔，先將要修改的資料的TM01,TM02,TM03,TM04,TM09寫在工作檔CFTTM09BACKUP-(系統日)，留做備用。
'4. 並在電腦中心的資料刪改記錄中留下修改筆數的記錄。
   
   strSql = "Create table cfttm09backup_20250306 (TM01 VARCHAR2(3 CHAR), TM02 VARCHAR2(6 CHAR), TM03 VARCHAR2(1 CHAR), TM04 VARCHAR2(2 CHAR), TM09 VARCHAR2(395 CHAR)) "
   cnnConnection.Execute strSql
   
   strExc(0) = "select tm01,tm02,tm03,tm04,tm09 from trademark where tm01='CFT' and length(tm09) > 0 order by tm01,tm02,tm03,tm04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Me.Enabled = False
      Screen.MousePointer = vbHourglass
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strExc(1) = "" & RsTemp.Fields("tm09")
         strExc(2) = ""
         tmpArr = Empty
         tmpArr = Split(strExc(1), ",")
         For intI = 0 To UBound(tmpArr)
            strExc(3) = ""
            If Trim(tmpArr(intI)) <> "" Then
               For intA = 1 To Len(Trim(tmpArr(intI)))
                  intB = Asc(Mid(Trim(tmpArr(intI)), intA, 1))
                  If intB >= 48 And intB <= 57 Then
                     strExc(3) = strExc(3) & Chr(intB)
                  End If
               Next
               strExc(2) = strExc(2) & "," & Format(Val(strExc(3)), "00")
            End If
            If Format(Val(strExc(3)), "00") <> Trim(tmpArr(intI)) Then
               strSql = "Update tmgoods set tg05='" & Format(Val(strExc(3)), "00") & "' where tg01='" & RsTemp.Fields("tm01") & "' and tg02='" & RsTemp.Fields("tm02") & "' and tg03='" & RsTemp.Fields("tm03") & "' and tg04='" & RsTemp.Fields("tm04") & "' and tg05='" & Trim(tmpArr(intI)) & "' "
               cnnConnection.Execute strSql
            End If
         Next intI
         strExc(2) = Mid(strExc(2), 2)
         If strExc(1) <> strExc(2) Then
            strSql = "Update trademark set tm09='" & strExc(2) & "' where tm01='" & RsTemp.Fields("tm01") & "' and tm02='" & RsTemp.Fields("tm02") & "' and tm03='" & RsTemp.Fields("tm03") & "' and tm04='" & RsTemp.Fields("tm04") & "' "
            cnnConnection.Execute strSql
            strSql = "Insert into CFTTM09BACKUP_20250306(TM01,TM02,TM03,TM04,TM09) VALUES ('" & RsTemp.Fields("tm01") & "', '" & RsTemp.Fields("tm02") & "', '" & RsTemp.Fields("tm03") & "', '" & RsTemp.Fields("tm04") & "', '" & RsTemp.Fields("tm09") & "')"
            cnnConnection.Execute strSql
         End If
         RsTemp.MoveNext
      Loop
      Screen.MousePointer = vbDefault
      Me.Enabled = True
   End If
   MsgBox "OK!"
End Sub


'Added by Morgan 2025/3/10
'P程序人員工作1140211起改智權區域分配，已收文未發文資料更新承辦人
Private Sub UpdCP14()
   Dim strQ As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCaseNo As String, stNewCP14 As String
   
   ProgressBar1.Value = 0
   Label13.Caption = 0
      
   strQ = "select * from caseprogress where cp05=20250308 and cp01='P' and cp10='1605' and cp14 is null"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      With rsQuery
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Do While Not .EOF
         stCaseNo = .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
         ProgressBar1.Value = .AbsolutePosition
         Label13.Caption = .AbsolutePosition & "/" & .RecordCount
         StatusBar1.Panels.Item(1).Text = stCaseNo
         DoEvents
         
         stNewCP14 = PUB_GetPHandler(stCaseNo)
         If stNewCP14 <> "" & .Fields("cp14") Then
            strQ = "update caseprogress set cp14='" & stNewCP14 & "' where cp09='" & .Fields("cp09") & "' and cp158=0 and cp159=0"
            cnnConnection.Execute strQ, intQ
         End If
         StatusBar1.Panels.Item(1).Text = stCaseNo & "...OK"
         .MoveNext
      Loop
      End With
   End If
   Set rsQuery = Nothing
End Sub

'Added by Morgan 2025/3/10
Private Sub Command60_Click()
   If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then
      If MsgBox("目前為連線為正式資料庫，是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   PUB_BatchArchive
   
End Sub

'Added by Morgan 2025/3/25
Public Sub PUB_BatchArchive(Optional pBack As Boolean = False)
   Dim stSQL As String, intQ As Integer, ii As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stTable As String, stCon As String
   Dim stFullFtpPath As String, stFtpIP As String
   Dim bolByCopy As Boolean
   Dim bolInTran As Boolean
   
   Dim stFtpFromDir As String, stFtpToDir As String, stFilePath As String
   Dim stMoveFrom As String, stMoveTo As String, stMoveToDir As String
   Dim stFtpToDir2 As String, stMoveTo2 As String, stMoveToDir2 As String
   Dim arrDir() As String
   Dim oTime As Variant
   Dim lngSize As Long
   
On Error GoTo ErrHnd
   
   If txtCPP13 > (strSrvDate(1) - 10000) Then
      MsgBox "CPP13 不可大於系統日-2年!!", vbCritical
      Exit Sub
   End If
   
   oTime = time
   lblElapse.Caption = 0
   
   stTable = "CASEPAPERPDF"
   stCon = " and cpp13<" & txtCPP13
   If pBack Then
      stCon = stCon & " and cpp19 is not null"
      
   Else
      stCon = stCon & " and cpp19 is null"
      
      If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then
         stFtpFromDir = "\\" & Pub_GetSpecMan("FTP_VOL_IP")
         stFtpFromDir = stFtpFromDir & "\" & Replace(Pub_GetSpecMan("FTP_VOL_DIR"), "/", "\")
                  
         stFtpToDir = "\\" & Pub_GetSpecMan("FTP_VOL_IP_ARC")
         stFtpToDir = stFtpToDir & "\" & Replace(Pub_GetSpecMan("FTP_VOL_DIR_ARC"), "/", "\")
         'Added by Morgan 2025/5/19
         stFtpToDir2 = "\\" & Pub_GetSpecMan("FTP_VOL_IP_2_ARC")
         stFtpToDir2 = stFtpToDir2 & "\" & Replace(Pub_GetSpecMan("FTP_VOL_DIR_2_ARC"), "/", "\")
         'end 2025/5/19
      Else
         stFtpFromDir = "\\" & Pub_GetSpecMan("FTP_VOL_IP_2")
         stFtpFromDir = stFtpFromDir & "\" & Replace(Pub_GetSpecMan("FTP_VOL_DIR_2"), "/", "\")
         
         stFtpToDir = "\\" & Pub_GetSpecMan("FTP_VOL_IP_2_ARC")
         stFtpToDir = stFtpToDir & "\" & Replace(Pub_GetSpecMan("FTP_VOL_DIR_2_ARC"), "/", "\")
      End If
      stFtpFromDir = stFtpFromDir & "\" & stTable
      stFtpToDir = stFtpToDir & "\" & stTable
   End If
   
   PUB_WriteLog "FTP封存讀取資料,條件:" & stCon
   stSQL = "select cpp14,cpp03 from casepaperpdf where cpp13>0  and cpp14 is not null" & stCon
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      m_bolStop = False
      With rsQuery
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      PUB_WriteLog "FTP封存(" & .RecordCount & "筆)開始"
      Do While Not .EOF
         If m_bolStop = True Then
            If MsgBox("是否要繼續？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
               Exit Do
            Else
               m_bolStop = False
            End If
         End If
         
         Sleep 100

         stFullFtpPath = PUB_GetFtpTableDir(stTable, stFtpIP) & "/" & .Fields("cpp14")
         txtNow = Format(Round(.Fields("cpp03") / 1024), "#,###")
         ProgressBar1.Value = .AbsolutePosition
         Label13.Caption = .AbsolutePosition & "/" & .RecordCount
         StatusBar1.Panels.Item(1).Text = stTable & "/" & .Fields("cpp14")
         
         lngSize = DateDiff("s", oTime, time)
         lblElapse = Format(lngSize \ 3600, "00") & ":" & Format((lngSize Mod 3600) \ 60, "00") & ":" & Format(lngSize Mod 60, "00")
         DoEvents
         
         cnnConnection.BeginTrans
         bolInTran = True
         If pBack Then
            stSQL = "update casepaperpdf set cpp19='' where cpp14='" & .Fields("cpp14") & "'"
         Else
            stSQL = "update casepaperpdf set cpp19='Y' where cpp14='" & .Fields("cpp14") & "'"
         End If
         cnnConnection.Execute stSQL, intQ
         
         If pBack Then
         
            If PUB_MoveToArchive(.Fields("cpp14"), stTable, bolByCopy, pBack) Then
               cnnConnection.CommitTrans
               bolInTran = False
               If bolByCopy Then
                  PUB_FtpDelFile2 stFullFtpPath, , , stFtpIP
               End If
   
               StatusBar1.Panels.Item(1).Text = stFullFtpPath & "...OK"
            Else
               cnnConnection.RollbackTrans
               bolInTran = False
               StatusBar1.Panels.Item(1).Text = stFullFtpPath & "...Fail"
               MsgBox "封存失敗!!", vbCritical
               Exit Do
            End If
            
         Else
            stFilePath = Replace(.Fields("cpp14"), "/", "\")
            stMoveFrom = stFtpFromDir & "\" & stFilePath
            stMoveTo = stFtpToDir & "\" & stFilePath
            stMoveToDir = Left(stMoveTo, InStrRev(stMoveTo, "\") - 1)
            
            If Dir(stMoveToDir, vbDirectory) = "" Then
               stMoveToDir = stFtpToDir
               arrDir = Split(stFilePath, "\")
               For ii = LBound(arrDir) To UBound(arrDir) - 1
                  stMoveToDir = stMoveToDir & "\" & arrDir(ii)
                  If Dir(stMoveToDir, vbDirectory) = "" Then
                     MkDir stMoveToDir
                  End If
               Next
            End If
            
            'Added by Morgan 2025/5/19
            If stFtpToDir2 <> "" Then
               stMoveTo2 = stFtpToDir2 & "\" & stFilePath
               stMoveToDir2 = Left(stMoveTo2, InStrRev(stMoveTo2, "\") - 1)
               
               If Dir(stMoveToDir2, vbDirectory) = "" Then
                  stMoveToDir2 = stFtpToDir2
                  arrDir = Split(stFilePath, "\")
                  For ii = LBound(arrDir) To UBound(arrDir) - 1
                     stMoveToDir2 = stMoveToDir2 & "\" & arrDir(ii)
                     If Dir(stMoveToDir2, vbDirectory) = "" Then
                        MkDir stMoveToDir2
                     End If
                  Next
               End If
            End If
            'end 2025/5/19
            
            If fnMoveFile(stMoveFrom, stMoveTo, stMoveTo2) Then
                cnnConnection.CommitTrans
            Else
                cnnConnection.RollbackTrans
            End If
            bolInTran = False
         End If
         
         txtTot = Format(Val(Format(txtTot)) + Val(Format(txtNow)), "#,###")
         
         If .AbsolutePosition Mod 1000 = 0 Then
            PUB_WriteLog "FTP封存已完成" & .AbsolutePosition & "筆(" & txtTot & "K)"
         End If
         
         .MoveNext
      Loop
      PUB_WriteLog "FTP封存結束共" & IIf(.EOF, .RecordCount, .AbsolutePosition) & "筆(" & txtTot & "K)"
      End With
   End If
   
ErrHnd:
   If bolInTran Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      PUB_WriteLog StatusBar1.Panels.Item(1).Text & "-->" & Err.Description
      'MsgBox Err.Description, vbCritical
   End If

   Set rsQuery = Nothing
   
End Sub

Private Function fnMoveFile(pFrom As String, pTo As String, Optional pTo2 As String) As Boolean
    
On Error GoTo ErrHnd
   If pTo2 <> "" Then
      m_oFileSys.CopyFile pFrom, pTo2, True
   End If
   m_oFileSys.MoveFile pFrom, pTo
   fnMoveFile = True
   
ErrHnd:
   If Err.Number <> 0 Then
      PUB_WriteLog StatusBar1.Panels.Item(1).Text & "-->" & Err.Description
   End If
End Function

Private Sub Command61_Click()
   If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then
      If MsgBox("目前為連線為正式資料庫，是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   PUB_BatchArchive True
End Sub

Private Sub Command62_Click()
   m_bolStop = True
End Sub
