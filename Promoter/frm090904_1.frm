VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090904_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "工程師各式申請書-電子送件"
   ClientHeight    =   7670
   ClientLeft      =   410
   ClientTop       =   1500
   ClientWidth     =   10050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7670
   ScaleWidth      =   10050
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   285
      Left            =   8220
      TabIndex        =   243
      Top             =   600
      Visible         =   0   'False
      Width           =   1755
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   900
         Style           =   1  '圖片外觀
         TabIndex        =   244
         Top             =   -30
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   245
         Top             =   0
         Width           =   765
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   159
      Top             =   2200
      Width           =   10035
      _ExtentX        =   17709
      _ExtentY        =   9543
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "分割/實審"
      TabPicture(0)   =   "frm090904_1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FramePage"
      Tab(0).Control(1)=   "Frame416"
      Tab(0).Control(2)=   "Frame307"
      Tab(0).Control(3)=   "Frame12"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "修正/誤譯訂正"
      TabPicture(1)   =   "frm090904_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame204"
      Tab(1).Control(1)=   "Frame433"
      Tab(1).Control(2)=   "Check9(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "PPH審查"
      TabPicture(2)   =   "frm090904_1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Check431_2(0)"
      Tab(2).Control(2)=   "Check431_1(0)"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "再審/加速審查"
      TabPicture(3)   =   "frm090904_1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame422"
      Tab(3).Control(1)=   "Frame107"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "更正"
      TabPicture(4)   =   "frm090904_1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Line8"
      Tab(4).Control(1)=   "Line7"
      Tab(4).Control(2)=   "chk4Tab1(8)"
      Tab(4).Control(3)=   "chk4Tab1(10)"
      Tab(4).Control(4)=   "chk4Tab1(7)"
      Tab(4).Control(5)=   "chk4Tab1(2)"
      Tab(4).Control(6)=   "chk4Tab1(5)"
      Tab(4).Control(7)=   "chk4Tab1(6)"
      Tab(4).Control(8)=   "chk4Tab1(0)"
      Tab(4).Control(9)=   "chk4Tab1(1)"
      Tab(4).Control(10)=   "chk4Tab1(3)"
      Tab(4).Control(11)=   "chk4Tab1(4)"
      Tab(4).Control(12)=   "chk4Tab1(9)"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "變更頁數"
      TabPicture(5)   =   "frm090904_1.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label28"
      Tab(5).Control(1)=   "Label36"
      Tab(5).Control(2)=   "FramePA6468"
      Tab(5).Control(3)=   "FrameAddPage"
      Tab(5).Control(4)=   "FrameCP167"
      Tab(5).Control(5)=   "FrameCP168"
      Tab(5).Control(6)=   "Frame6"
      Tab(5).Control(7)=   "txtAddPageFee"
      Tab(5).Control(8)=   "txtDecreasePageFee"
      Tab(5).ControlCount=   9
      TabCaption(6)   =   "同時辦理事項"
      TabPicture(6)   =   "frm090904_1.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Frame1"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame421"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame421 
         Appearance      =   0  '平面
         ForeColor       =   &H80000008&
         Height          =   4870
         Left            =   1590
         TabIndex        =   312
         Top             =   330
         Visible         =   0   'False
         Width           =   4660
         Begin VB.CheckBox chkAtt421_3 
            Caption         =   "委任書"
            Height          =   210
            Index           =   2
            Left            =   300
            TabIndex        =   330
            Top             =   4170
            Width           =   2850
         End
         Begin VB.CheckBox chkAtt421_3 
            Caption         =   "涉及專利侵權爭議證明文件"
            Height          =   210
            Index           =   1
            Left            =   300
            TabIndex        =   329
            Top             =   3900
            Width           =   2850
         End
         Begin VB.CheckBox chkAtt421_3 
            Caption         =   "商業實施證明文件"
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   328
            Top             =   3630
            Width           =   2850
         End
         Begin VB.CheckBox chkAtt421_1 
            Caption         =   "委任代理人辦理事項聲明"
            Height          =   210
            Index           =   3
            Left            =   300
            TabIndex        =   327
            Top             =   1380
            Width           =   3300
         End
         Begin VB.CheckBox chkAtt421_1 
            Caption         =   "非專利權人申請技術報告事由"
            Height          =   210
            Index           =   2
            Left            =   300
            TabIndex        =   326
            Top             =   1110
            Width           =   3300
         End
         Begin VB.CheckBox chkAtt421_1 
            Caption         =   "專利權人申請技術報告事由"
            Height          =   210
            Index           =   1
            Left            =   300
            TabIndex        =   325
            Top             =   840
            Width           =   3300
         End
         Begin VB.CheckBox chkAtt421_1 
            Caption         =   "專利權已當然消滅"
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   324
            Top             =   570
            Width           =   3300
         End
         Begin VB.CheckBox chkAtt421_2 
            Caption         =   "變更專利權人之姓名或名稱　是"
            Height          =   210
            Index           =   3
            Left            =   300
            TabIndex        =   323
            Top             =   2880
            Width           =   2850
         End
         Begin VB.CheckBox chkAtt421_2 
            Caption         =   "變更專利權人之代表人　　　是"
            Height          =   210
            Index           =   2
            Left            =   300
            TabIndex        =   322
            Top             =   2610
            Width           =   2850
         End
         Begin VB.CheckBox chkAtt421_2 
            Caption         =   "變更專利權人之代理人　　　是"
            Height          =   210
            Index           =   1
            Left            =   300
            TabIndex        =   321
            Top             =   2340
            Width           =   3300
         End
         Begin VB.CheckBox chkAtt421_2 
            Caption         =   "變更專利權人之地址　　　　是"
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   320
            Top             =   2070
            Width           =   3300
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "【附送書件】"
            Height          =   180
            Left            =   120
            TabIndex        =   319
            Top             =   3300
            Width           =   1080
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "【同時辦理事項】"
            Height          =   180
            Left            =   120
            TabIndex        =   318
            Top             =   1800
            Width           =   1440
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "【相關事項】"
            Height          =   180
            Left            =   120
            TabIndex        =   317
            Top             =   270
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '平面
         Caption         =   "同時辦理事項"
         ForeColor       =   &H80000008&
         Height          =   1540
         Left            =   500
         TabIndex        =   311
         Top             =   510
         Width           =   5170
         Begin VB.CheckBox chkAtt3 
            Caption         =   "變更申請人之代表人"
            Height          =   210
            Index           =   3
            Left            =   270
            TabIndex        =   315
            Top             =   600
            Width           =   2490
         End
         Begin VB.CheckBox chkAtt3 
            Caption         =   "變更申請人之地址"
            Height          =   210
            Index           =   1
            Left            =   270
            TabIndex        =   314
            Top             =   330
            Width           =   2490
         End
         Begin VB.CheckBox chkAtt3 
            Caption         =   "變更申請人之國籍"
            Height          =   210
            Index           =   5
            Left            =   270
            TabIndex        =   313
            Top             =   870
            Width           =   2490
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "※其他同時辦理事項因有規費，故請程序另外產生申請書送件"
            ForeColor       =   &H00FF00FF&
            Height          =   300
            Left            =   150
            TabIndex        =   316
            Top             =   1170
            Width           =   5010
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox txtDecreasePageFee 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         Left            =   -67050
         TabIndex        =   291
         Top             =   1620
         Width           =   840
      End
      Begin VB.TextBox txtAddPageFee 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         Left            =   -67050
         Locked          =   -1  'True
         TabIndex        =   280
         Top             =   1320
         Width           =   840
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         Caption         =   "修正後中文本頁數"
         ForeColor       =   &H00FF0000&
         Height          =   2115
         Left            =   -71820
         TabIndex        =   271
         Top             =   540
         Width           =   2535
         Begin VB.TextBox txtDocCh4 
            Height          =   270
            Index           =   7
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   307
            Top             =   1770
            Width           =   420
         End
         Begin VB.TextBox txtPageCount 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Left            =   1920
            TabIndex        =   301
            Top             =   1380
            Width           =   420
         End
         Begin VB.TextBox txtDocCh4 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Index           =   4
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   275
            Top             =   1095
            Width           =   420
         End
         Begin VB.TextBox txtDocCh4 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Index           =   3
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   274
            Top             =   840
            Width           =   420
         End
         Begin VB.TextBox txtDocCh4 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   273
            Top             =   579
            Width           =   420
         End
         Begin VB.TextBox txtDocCh4 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   272
            Top             =   315
            Width           =   420
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "圖式圖數:"
            Height          =   180
            Left            =   240
            TabIndex        =   308
            Top             =   1830
            Width           =   765
         End
         Begin VB.Label lblPage 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "修正後總頁數"
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   240
            TabIndex        =   302
            Top             =   1425
            Width           =   1080
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "圖式頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   279
            Top             =   1155
            Width           =   765
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   278
            Top             =   900
            Width           =   1485
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "說明書頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   277
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "摘要頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   276
            Top             =   360
            Width           =   765
         End
      End
      Begin VB.Frame FrameCP168 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         Caption         =   "刪除已審頁數"
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   -69000
         TabIndex        =   265
         Top             =   2910
         Width           =   2535
         Begin VB.TextBox txtDocCp168 
            Height          =   270
            Index           =   0
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   293
            Top             =   315
            Width           =   420
         End
         Begin VB.TextBox txtDocCp168 
            Height          =   270
            Index           =   1
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   294
            Top             =   579
            Width           =   420
         End
         Begin VB.TextBox txtDocCp168 
            Height          =   270
            Index           =   3
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   295
            Top             =   855
            Width           =   420
         End
         Begin VB.TextBox txtDocCp168 
            Height          =   270
            Index           =   4
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   296
            Top             =   1125
            Width           =   420
         End
         Begin VB.TextBox txtCP168 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   298
            Top             =   1395
            Width           =   420
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "摘要頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   270
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "說明書頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   269
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   268
            Top             =   900
            Width           =   1485
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "圖式頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   267
            Top             =   1170
            Width           =   765
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "頁數總計:"
            ForeColor       =   &H00000040&
            Height          =   180
            Left            =   240
            TabIndex        =   266
            Top             =   1440
            Visible         =   0   'False
            Width           =   765
         End
      End
      Begin VB.Frame FrameCP167 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         Caption         =   "刪除未審頁數"
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   -71820
         TabIndex        =   259
         Top             =   2910
         Width           =   2535
         Begin VB.TextBox txtCP167 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   292
            Top             =   1395
            Width           =   420
         End
         Begin VB.TextBox txtDocCp167 
            Height          =   270
            Index           =   4
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   290
            Top             =   1125
            Width           =   420
         End
         Begin VB.TextBox txtDocCp167 
            Height          =   270
            Index           =   3
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   289
            Top             =   855
            Width           =   420
         End
         Begin VB.TextBox txtDocCp167 
            Height          =   270
            Index           =   1
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   288
            Top             =   579
            Width           =   420
         End
         Begin VB.TextBox txtDocCp167 
            Height          =   270
            Index           =   0
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   287
            Top             =   315
            Width           =   420
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "頁數總計:"
            ForeColor       =   &H00000040&
            Height          =   180
            Left            =   240
            TabIndex        =   264
            Top             =   1440
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "圖式頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   263
            Top             =   1170
            Width           =   765
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   262
            Top             =   900
            Width           =   1485
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "說明書頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   261
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "摘要頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   260
            Top             =   360
            Width           =   765
         End
      End
      Begin VB.Frame FrameAddPage 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         Caption         =   "增加頁數"
         ForeColor       =   &H00FF0000&
         Height          =   1845
         Left            =   -74670
         TabIndex        =   253
         Top             =   2910
         Width           =   2535
         Begin VB.TextBox txtDocAdd 
            Height          =   270
            Index           =   0
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   281
            Top             =   315
            Width           =   420
         End
         Begin VB.TextBox txtDocAdd 
            Height          =   270
            Index           =   1
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   282
            Top             =   579
            Width           =   420
         End
         Begin VB.TextBox txtDocAdd 
            Height          =   270
            Index           =   3
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   283
            Top             =   855
            Width           =   420
         End
         Begin VB.TextBox txtDocAdd 
            Height          =   270
            Index           =   4
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   284
            Top             =   1125
            Width           =   420
         End
         Begin VB.TextBox txtAddPage 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   286
            Top             =   1395
            Width           =   420
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "摘要頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   258
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "說明書頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   257
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   256
            Top             =   900
            Width           =   1485
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "圖式頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   255
            Top             =   1170
            Width           =   765
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "頁數總計:"
            ForeColor       =   &H00000040&
            Height          =   180
            Left            =   240
            TabIndex        =   254
            Top             =   1440
            Visible         =   0   'False
            Width           =   765
         End
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之設計說明書"
         Height          =   195
         Index           =   9
         Left            =   -74370
         TabIndex        =   242
         Tag             =   "3.FIX_DESCRIPTION.pdf"
         Top             =   2745
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之發明圖式"
         Height          =   195
         Index           =   4
         Left            =   -74370
         TabIndex        =   241
         Tag             =   "1.FIX_DRAWINGS.pdf"
         Top             =   1440
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之發明申請專利範圍"
         Height          =   195
         Index           =   3
         Left            =   -74370
         TabIndex        =   240
         Tag             =   "1.FIX_CLAIMS.pdf"
         Top             =   1215
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之發明說明書"
         Height          =   195
         Index           =   1
         Left            =   -74370
         TabIndex        =   239
         Tag             =   "1.FIX_DESCRIPTION.pdf"
         Top             =   780
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之發明摘要"
         Height          =   195
         Index           =   0
         Left            =   -74370
         TabIndex        =   238
         Tag             =   "1.FIX_ABSTRACT.pdf"
         Top             =   570
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之新型說明書"
         Height          =   195
         Index           =   6
         Left            =   -74370
         TabIndex        =   237
         Tag             =   "2.FIX_DESCRIPTION.pdf"
         Top             =   1980
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之新型摘要"
         Height          =   195
         Index           =   5
         Left            =   -74370
         TabIndex        =   236
         Tag             =   "2.FIX_ABSTRACT.pdf"
         Top             =   1755
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之序列表"
         Height          =   195
         Index           =   2
         Left            =   -74370
         TabIndex        =   235
         Tag             =   "1.FIX.SEQ.pdf"
         Top             =   1005
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之新型申請專利範圍"
         Height          =   195
         Index           =   7
         Left            =   -74370
         TabIndex        =   234
         Tag             =   "2.FIX_CLAIMS.pdf"
         Top             =   2190
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之設計圖式"
         Height          =   195
         Index           =   10
         Left            =   -74370
         TabIndex        =   233
         Tag             =   "3.FIX_DRAWINGS.pdf"
         Top             =   2970
         Width           =   3045
      End
      Begin VB.CheckBox chk4Tab1 
         Caption         =   "更正後之新型圖式"
         Height          =   195
         Index           =   8
         Left            =   -74370
         TabIndex        =   232
         Tag             =   "2.FIX_DRAWINGS.pdf"
         Top             =   2415
         Width           =   3045
      End
      Begin VB.Frame FramePA6468 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         Caption         =   "修正前中文本頁數"
         ForeColor       =   &H00FF0000&
         Height          =   2115
         Left            =   -74670
         TabIndex        =   223
         Top             =   540
         Width           =   2535
         Begin VB.TextBox txtPage 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Left            =   1920
            TabIndex        =   299
            Top             =   1380
            Width           =   420
         End
         Begin VB.TextBox txtDocCh3 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Index           =   4
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   227
            Top             =   1095
            Width           =   420
         End
         Begin VB.TextBox txtDocCh3 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Index           =   3
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   226
            Top             =   840
            Width           =   420
         End
         Begin VB.TextBox txtDocCh3 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   225
            Top             =   585
            Width           =   420
         End
         Begin VB.TextBox txtDocCh3 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   224
            Top             =   315
            Width           =   420
         End
         Begin VB.Label Label35 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "原總頁數"
            ForeColor       =   &H00000040&
            Height          =   180
            Left            =   240
            TabIndex        =   300
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "圖式頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   231
            Top             =   1155
            Width           =   765
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   230
            Top             =   900
            Width           =   1485
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "說明書頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   229
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "摘要頁數:"
            Height          =   180
            Left            =   240
            TabIndex        =   228
            Top             =   360
            Width           =   765
         End
      End
      Begin VB.CheckBox Check9 
         Caption         =   "本案續行再審查"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   1
         Left            =   -74220
         TabIndex        =   222
         Top             =   330
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Caption         =   "取消暫放著"
         Height          =   3075
         Left            =   -74160
         TabIndex        =   206
         Top             =   4170
         Visible         =   0   'False
         Width           =   5385
         Begin VB.OptionButton Option1 
            Caption         =   "PPO-Publication-Server申請專利範圍文件名稱及日期清單"
            Height          =   240
            Index           =   9
            Left            =   60
            TabIndex        =   216
            Top             =   2790
            Width           =   4665
         End
         Begin VB.OptionButton Option1 
            Caption         =   "KIPO-K-PION申請專利範圍文件名稱及日期清單"
            Height          =   240
            Index           =   8
            Left            =   60
            TabIndex        =   215
            Top             =   2494
            Width           =   4335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "SPTO-Expedientes-Digitalizados申請專利範圍文件名稱及日期清單"
            Height          =   240
            Index           =   7
            Left            =   60
            TabIndex        =   214
            Top             =   2201
            Width           =   5295
         End
         Begin VB.OptionButton Option1 
            Caption         =   "JPO-AIPN申請專利範圍文件名稱及日期清單"
            Height          =   240
            Index           =   6
            Left            =   60
            TabIndex        =   213
            Top             =   1908
            Width           =   4335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "USPTO-Public-PAIR申請專利範圍文件名稱及日期清單"
            Height          =   240
            Index           =   5
            Left            =   60
            TabIndex        =   212
            Top             =   1615
            Width           =   4545
         End
         Begin VB.OptionButton Option1 
            Caption         =   "PPO-Publication-Server審查意見書文件名稱及日期清單"
            Height          =   240
            Index           =   4
            Left            =   60
            TabIndex        =   211
            Top             =   1322
            Width           =   4515
         End
         Begin VB.OptionButton Option1 
            Caption         =   "KIPO-K-PION審查意見書文件名稱及日期清單"
            Height          =   240
            Index           =   3
            Left            =   60
            TabIndex        =   210
            Top             =   1029
            Width           =   4335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "SPTO-Expedientes-Digitalizados審查意見書文件名稱及日期清單"
            Height          =   240
            Index           =   2
            Left            =   60
            TabIndex        =   209
            Top             =   736
            Width           =   5145
         End
         Begin VB.OptionButton Option1 
            Caption         =   "JPO-AIPN審查意見書文件名稱及日期清單"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   208
            Top             =   443
            Width           =   4335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "USPTO-Public-PAIR審查意見書文件名稱及日期清單"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   207
            Top             =   150
            Width           =   4335
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "附送書件"
         Height          =   4005
         Left            =   -71190
         TabIndex        =   163
         Top             =   1350
         Width           =   6075
         Begin VB.CheckBox chk0Tab 
            Caption         =   "本分割案與原申請案之說明書差異部分之劃線本"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   21
            Left            =   750
            TabIndex        =   310
            Tag             =   " .ATT.pdf"
            Top             =   3330
            Width           =   4430
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "本分割案與原申請案之說明書差異部分說明文件"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   22
            Left            =   750
            TabIndex        =   309
            Tag             =   " .ATT.pdf"
            Top             =   3600
            Width           =   4430
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "優惠期證明文件"
            Height          =   195
            Index           =   11
            Left            =   750
            TabIndex        =   46
            Tag             =   " .EXHIBITION.pdf"
            Top             =   2250
            Width           =   2100
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "發明摘要"
            Height          =   195
            Index           =   2
            Left            =   750
            TabIndex        =   38
            Tag             =   "1.INV_ABSTRACT.pdf"
            Top             =   710
            Width           =   1230
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "發明說明書"
            Height          =   195
            Index           =   3
            Left            =   750
            TabIndex        =   39
            Tag             =   "1.INV_DESCRIPTION.pdf"
            Top             =   960
            Width           =   1230
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "發明申請專利範圍"
            Height          =   195
            Index           =   4
            Left            =   750
            TabIndex        =   40
            Tag             =   "1.INV_CLAIMS.pdf"
            Top             =   1200
            Width           =   1860
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "發明圖式"
            Height          =   195
            Index           =   5
            Left            =   750
            TabIndex        =   41
            Tag             =   "1.INV_DRAWINGS.pdf"
            Top             =   1460
            Width           =   1230
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "說明書"
            Height          =   195
            Index           =   6
            Left            =   750
            TabIndex        =   42
            Tag             =   " .ORI.pdf"
            Top             =   1710
            Width           =   980
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "簡體字本"
            Height          =   195
            Index           =   9
            Left            =   750
            TabIndex        =   45
            Tag             =   " .SEP.pdf"
            Top             =   1980
            Width           =   1230
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "國內生物材料寄存證明文件"
            Height          =   195
            Index           =   12
            Left            =   750
            TabIndex        =   47
            Tag             =   "1.DOMESTICPROOF.pdf"
            Top             =   2520
            Width           =   2535
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "國外生物材料寄存證明文件"
            Height          =   195
            Index           =   13
            Left            =   750
            TabIndex        =   48
            Tag             =   "1.FOREIGNPROOF.pdf"
            Top             =   2790
            Width           =   2535
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "生物材料為通常知識者易於獲得證明文件"
            Height          =   195
            Index           =   14
            Left            =   750
            TabIndex        =   49
            Tag             =   "1.EASILYOBTAINED.pdf"
            Top             =   3060
            Width           =   4160
         End
         Begin VB.CheckBox Check416_1 
            Caption         =   "專利修正申請書"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   750
            TabIndex        =   36
            Top             =   210
            Width           =   2175
         End
         Begin VB.CheckBox Check416_2 
            Caption         =   "專利誤譯訂正申請書"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   750
            TabIndex        =   37
            Top             =   470
            Width           =   2175
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "新型圖式"
            Height          =   195
            Index           =   18
            Left            =   3750
            TabIndex        =   53
            Tag             =   "2.UTL_DRAWINGS.pdf"
            Top             =   1130
            Width           =   1230
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "設計說明書"
            Height          =   195
            Index           =   19
            Left            =   3750
            TabIndex        =   54
            Tag             =   "3.DES_DESCRIPTION.pdf"
            Top             =   1710
            Width           =   1230
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "設計圖式"
            Height          =   195
            Index           =   20
            Left            =   3750
            TabIndex        =   55
            Tag             =   "3.DES_DRAWINGS.pdf"
            Top             =   2010
            Width           =   1860
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "新型摘要"
            Height          =   195
            Index           =   15
            Left            =   3750
            TabIndex        =   50
            Tag             =   "2.UTL_ABSTRACT.pdf"
            Top             =   240
            Width           =   1230
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "新型說明書"
            Height          =   195
            Index           =   16
            Left            =   3750
            TabIndex        =   51
            Tag             =   "2.UTL_DESCRIPTION.pdf"
            Top             =   540
            Width           =   1230
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "新型申請專利範圍"
            Height          =   195
            Index           =   17
            Left            =   3750
            TabIndex        =   52
            Tag             =   "2.UTL_CLAIMS.pdf"
            Top             =   830
            Width           =   1980
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "圖式"
            Height          =   195
            Index           =   7
            Left            =   1740
            TabIndex        =   43
            Top             =   1710
            Width           =   780
         End
         Begin VB.CheckBox chk0Tab 
            Caption         =   "序列表"
            Height          =   195
            Index           =   8
            Left            =   2595
            TabIndex        =   44
            Tag             =   " .SEQ.pdf"
            Top             =   1710
            Width           =   915
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00404000&
            BorderWidth     =   2
            X1              =   3630
            X2              =   5970
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00404000&
            BorderWidth     =   2
            X1              =   3630
            X2              =   5970
            Y1              =   1650
            Y2              =   1650
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "外文本"
            Height          =   180
            Left            =   120
            TabIndex        =   164
            Top             =   1710
            Width           =   540
         End
      End
      Begin VB.Frame Frame107 
         Appearance      =   0  '平面
         Caption         =   "再審查"
         ForeColor       =   &H00FF0000&
         Height          =   4125
         Left            =   -74790
         TabIndex        =   183
         Top             =   390
         Width           =   3165
         Begin VB.TextBox txtDocCh2 
            Height          =   270
            Index           =   7
            Left            =   1890
            MaxLength       =   4
            TabIndex        =   122
            Top             =   1328
            Width           =   420
         End
         Begin VB.CheckBox Check107_2 
            Caption         =   "專利誤譯訂正申請書"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   1110
            TabIndex        =   129
            Top             =   3450
            Width           =   1965
         End
         Begin VB.CheckBox chk3Tab2 
            Caption         =   "再審查理由書"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   1110
            TabIndex        =   128
            Tag             =   ".RE.pdf"
            Top             =   3180
            Value           =   1  '核取
            Width           =   1965
         End
         Begin VB.CheckBox Check107_1 
            Caption         =   "專利修正申請書"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   1110
            TabIndex        =   130
            Top             =   3720
            Width           =   1965
         End
         Begin VB.CheckBox Check107_1 
            Caption         =   "一併申請修正"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   118
            Top             =   210
            Width           =   1605
         End
         Begin VB.CheckBox Check107_2 
            Caption         =   "一併申請誤譯訂正"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   119
            Top             =   480
            Width           =   1785
         End
         Begin VB.TextBox txtDocCh2 
            Height          =   270
            Index           =   0
            Left            =   1890
            MaxLength       =   4
            TabIndex        =   120
            Top             =   780
            Width           =   420
         End
         Begin VB.TextBox txtDocCh2 
            Height          =   270
            Index           =   1
            Left            =   1890
            MaxLength       =   4
            TabIndex        =   121
            Top             =   1054
            Width           =   420
         End
         Begin VB.TextBox txtDocCh2 
            Height          =   270
            Index           =   2
            Left            =   1890
            MaxLength       =   4
            TabIndex        =   123
            Top             =   1602
            Width           =   420
         End
         Begin VB.TextBox txtDocCh2 
            Height          =   270
            Index           =   3
            Left            =   1890
            MaxLength       =   4
            TabIndex        =   124
            Top             =   1876
            Width           =   420
         End
         Begin VB.TextBox txtDocCh2 
            Height          =   270
            Index           =   4
            Left            =   1890
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   125
            Top             =   2150
            Width           =   420
         End
         Begin VB.TextBox txtDocCh2 
            Height          =   270
            Index           =   5
            Left            =   1890
            MaxLength       =   4
            TabIndex        =   126
            Top             =   2424
            Width           =   420
         End
         Begin VB.TextBox txtDocCh2 
            Height          =   270
            Index           =   6
            Left            =   1890
            TabIndex        =   127
            Top             =   2700
            Width           =   420
         End
         Begin VB.Label Label1 
            Caption         =   "不算超頁費"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.5
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   220
            Top             =   1373
            Width           =   900
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "序列表:"
            Height          =   180
            Left            =   360
            TabIndex        =   219
            Top             =   1373
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "附送書件"
            Height          =   180
            Left            =   360
            TabIndex        =   191
            Top             =   3180
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "摘要頁數:"
            Height          =   180
            Left            =   360
            TabIndex        =   190
            Top             =   825
            Width           =   765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "說明書頁數:"
            Height          =   180
            Left            =   360
            TabIndex        =   189
            Top             =   1099
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍頁數:"
            Height          =   180
            Left            =   360
            TabIndex        =   188
            Top             =   1647
            Width           =   1485
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "圖式頁數:"
            Height          =   180
            Left            =   360
            TabIndex        =   187
            Top             =   1921
            Width           =   765
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "頁數總計:"
            ForeColor       =   &H00000040&
            Height          =   180
            Left            =   360
            TabIndex        =   186
            Top             =   2195
            Width           =   765
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍項數:"
            ForeColor       =   &H00000040&
            Height          =   180
            Left            =   360
            TabIndex        =   185
            Top             =   2469
            Width           =   1485
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "圖式圖數:"
            Height          =   180
            Left            =   360
            TabIndex        =   184
            Top             =   2745
            Width           =   765
         End
      End
      Begin VB.Frame Frame422 
         Appearance      =   0  '平面
         Caption         =   "加速審查"
         ForeColor       =   &H00FF0000&
         Height          =   4905
         Left            =   -71610
         TabIndex        =   182
         Top             =   390
         Width           =   6555
         Begin VB.OptionButton Opt2Tab3 
            Caption         =   "為綠色技術相關案件"
            Height          =   225
            Index           =   3
            Left            =   150
            TabIndex        =   147
            Tag             =   "為綠色技術相關案件"
            Top             =   4590
            Width           =   6285
         End
         Begin VB.OptionButton Opt2Tab3 
            Caption         =   "為商業上之實施所必要"
            Height          =   225
            Index           =   2
            Left            =   150
            TabIndex        =   146
            Tag             =   "為商業上之實施所必要"
            Top             =   4320
            Width           =   6285
         End
         Begin VB.OptionButton Opt2Tab3 
            Caption         =   "外國對應申請案經美日歐專利局核發審查意見通知書及檢索報告但尚未審定"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   145
            Tag             =   "對應申請案經美日歐專利局"
            Top             =   4050
            Width           =   6285
         End
         Begin VB.OptionButton Opt2Tab3 
            Caption         =   "外國對應申請案經外國專利局實體審查而核准"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   144
            Tag             =   "外國對應申請案經外國專利局實體審查而核准"
            Top             =   3780
            Width           =   6285
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國即將公告之申請專利範圍"
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   135
            Top             =   1230
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國核准通知中譯本"
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   134
            Top             =   990
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國核准公告之申請專利範圍中譯本"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   132
            Top             =   510
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國核准公告之申請專利範圍"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   131
            Top             =   270
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "申請專利範圍中譯本與本申請案之差異說明"
            Height          =   195
            Index           =   8
            Left            =   165
            TabIndex        =   139
            Top             =   2190
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國核發之審查意見通知之申請專利範圍中譯本"
            Height          =   195
            Index           =   7
            Left            =   165
            TabIndex        =   138
            Top             =   1950
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國核發之審查意見通知之申請專利範圍"
            Height          =   195
            Index           =   6
            Left            =   165
            TabIndex        =   137
            Top             =   1710
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國即將公告之申請專利範圍中譯本"
            Height          =   195
            Index           =   5
            Left            =   165
            TabIndex        =   136
            Top             =   1470
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國核准通知"
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   133
            Top             =   750
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "外國核發之審查意見通知書及檢索報告"
            Height          =   195
            Index           =   9
            Left            =   165
            TabIndex        =   140
            Top             =   2460
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "對應案違反新穎性或進步性之非專利文獻"
            Height          =   195
            Index           =   10
            Left            =   165
            TabIndex        =   141
            Top             =   2700
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "具可專利性之理由"
            Height          =   195
            Index           =   11
            Left            =   165
            TabIndex        =   142
            Top             =   2940
            Width           =   4185
         End
         Begin VB.CheckBox chk3Tab3 
            Caption         =   "商業實施證明文件"
            Height          =   195
            Index           =   12
            Left            =   165
            TabIndex        =   143
            Top             =   3180
            Width           =   4185
         End
         Begin VB.Label Label18 
            Caption         =   "加速審查申請事由："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            TabIndex        =   201
            Top             =   3480
            Width           =   2115
         End
      End
      Begin VB.CheckBox Check431_1 
         Caption         =   "一併申請PPH修正"
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   -74910
         TabIndex        =   109
         Top             =   450
         Width           =   1725
      End
      Begin VB.CheckBox Check431_2 
         Caption         =   "一併申請誤譯訂正"
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   -74910
         TabIndex        =   110
         Top             =   720
         Width           =   1785
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '平面
         Caption         =   "附送書件"
         ForeColor       =   &H80000008&
         Height          =   2925
         Left            =   -73080
         TabIndex        =   181
         Top             =   330
         Width           =   6945
         Begin VB.CheckBox chk2Tab2 
            Caption         =   "對應之外國專利公開資訊查詢系統可取得之申請專利範圍文件名稱及日期清單"
            Height          =   195
            Index           =   5
            Left            =   225
            TabIndex        =   205
            Top             =   1570
            Width           =   6555
         End
         Begin VB.CheckBox chk2Tab2 
            Caption         =   "對應之外國專利公開資訊查詢系統可取得之審查意見書文件名稱及日期清單"
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   204
            Top             =   1304
            Width           =   6555
         End
         Begin VB.CheckBox chk2Tab2 
            Caption         =   "引用作為專利准駁判斷依據之非專利文獻"
            Height          =   195
            Index           =   6
            Left            =   225
            TabIndex        =   115
            Top             =   1836
            Width           =   3585
         End
         Begin VB.CheckBox chk2Tab2 
            Caption         =   "外國審查達到可核准之申請專利範圍中譯本或英譯本"
            Height          =   195
            Index           =   3
            Left            =   225
            TabIndex        =   114
            Top             =   1038
            Width           =   4605
         End
         Begin VB.CheckBox chk2Tab2 
            Caption         =   "外國核發之所有審查意見書中譯本或英譯本"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   112
            Top             =   506
            Width           =   4605
         End
         Begin VB.CheckBox chk2Tab2 
            Caption         =   "外國核發之所有審查意見書"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   111
            Top             =   240
            Width           =   4605
         End
         Begin VB.CheckBox Check431_2 
            Caption         =   "專利誤譯訂正申請書"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   117
            Top             =   2370
            Width           =   3045
         End
         Begin VB.CheckBox Check431_1 
            Caption         =   "發明專利PPH修正申請書"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   116
            Top             =   2102
            Width           =   3045
         End
         Begin VB.CheckBox chk2Tab2 
            Caption         =   "外國審查達到可核准之申請專利範圍"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   113
            Top             =   772
            Width           =   4605
         End
      End
      Begin VB.Frame Frame433 
         Appearance      =   0  '平面
         Caption         =   "誤譯訂正"
         ForeColor       =   &H00C00000&
         Height          =   4965
         Left            =   -71790
         TabIndex        =   180
         Top             =   360
         Width           =   6795
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之設計圖式"
            Height          =   195
            Index           =   30
            Left            =   3390
            TabIndex        =   106
            Tag             =   "3.FIX_DRAWINGS.pdf"
            Top             =   4710
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之設計說明書"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   28
            Left            =   3390
            TabIndex        =   104
            Tag             =   "3.COR_U_DESCRIPTION.pdf"
            Top             =   4176
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之設計說明書"
            Height          =   195
            Index           =   26
            Left            =   3390
            TabIndex        =   102
            Tag             =   "3.COR_DESCRIPTION.pdf"
            Top             =   3642
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之設計說明書"
            Height          =   195
            Index           =   29
            Left            =   3390
            TabIndex        =   105
            Tag             =   "3.FIX_DESCRIPTION.pdf"
            Top             =   4443
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之設計圖式"
            Height          =   195
            Index           =   27
            Left            =   3390
            TabIndex        =   103
            Tag             =   "3.COR_DRAWINGS.pdf"
            Top             =   3909
            Width           =   3375
         End
         Begin VB.CheckBox Check433 
            Caption         =   "一併申請修正"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   60
            TabIndex        =   108
            Top             =   4425
            Width           =   2385
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "佐證資料"
            Height          =   195
            Index           =   31
            Left            =   60
            TabIndex        =   107
            Top             =   4710
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之新型摘要"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   18
            Left            =   3390
            TabIndex        =   94
            Tag             =   "2.COR_U_ABSTRACT.pdf"
            Top             =   1228
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之新型說明書"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   19
            Left            =   3390
            TabIndex        =   95
            Tag             =   "2.COR_U_DESCRIPTION.pdf"
            Top             =   1490
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之新型申請專利範圍"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   20
            Left            =   3390
            TabIndex        =   96
            Tag             =   "2.COR_U_CLAIMS.pdf"
            Top             =   1752
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之新型圖式"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   21
            Left            =   3390
            TabIndex        =   97
            Tag             =   "2.COR_U_DRAWINGS.pdf"
            Top             =   2014
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之新型摘要"
            Height          =   195
            Index           =   22
            Left            =   3390
            TabIndex        =   98
            Tag             =   "2.FIX_ABSTRACT.pdf"
            Top             =   2276
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之新型說明書"
            Height          =   195
            Index           =   23
            Left            =   3390
            TabIndex        =   99
            Tag             =   "2.FIX_DESCRIPTION.pdf"
            Top             =   2538
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之新型申請專利範圍"
            Height          =   195
            Index           =   24
            Left            =   3390
            TabIndex        =   100
            Tag             =   "2.FIX_CLAIMS.pdf"
            Top             =   2800
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之新型圖式"
            Height          =   195
            Index           =   25
            Left            =   3390
            TabIndex        =   101
            Tag             =   "2.FIX_DRAWINGS.pdf"
            Top             =   3070
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之新型說明書"
            Height          =   195
            Index           =   15
            Left            =   3390
            TabIndex        =   91
            Tag             =   "2.COR_DESCRIPTION.pdf"
            Top             =   442
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之新型摘要"
            Height          =   195
            Index           =   14
            Left            =   3390
            TabIndex        =   90
            Tag             =   "2.COR_ABSTRACT.pdf"
            Top             =   180
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之新型申請專利範圍"
            Height          =   195
            Index           =   16
            Left            =   3390
            TabIndex        =   92
            Tag             =   "2.COR_CLAIMS.pdf"
            Top             =   704
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之新型圖式"
            Height          =   195
            Index           =   17
            Left            =   3390
            TabIndex        =   93
            Tag             =   "2.COR_DRAWINGS.pdf"
            Top             =   966
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之發明摘要"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   81
            Tag             =   "1.COR_U_ABSTRACT.pdf"
            Top             =   1640
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之發明說明書"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   82
            Tag             =   "1.COR_U_DESCRIPTION.pdf"
            Top             =   1926
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之序列表"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   83
            Tag             =   "1.COR_U.SEQ.pdf"
            Top             =   2212
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正部分劃線之發明申請專利範圍"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   84
            Tag             =   "1.COR_U_CLAIMS.pdf"
            Top             =   2498
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之發明圖式"
            Height          =   195
            Index           =   13
            Left            =   60
            TabIndex        =   89
            Tag             =   "1.FIX_DRAWINGS.pdf"
            Top             =   3930
            Width           =   3285
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之發明摘要"
            Height          =   195
            Index           =   9
            Left            =   60
            TabIndex        =   85
            Tag             =   "1.FIX_ABSTRACT.pdf"
            Top             =   2784
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之發明說明書"
            Height          =   195
            Index           =   10
            Left            =   60
            TabIndex        =   86
            Tag             =   "1.FIX_DESCRIPTION.pdf"
            Top             =   3070
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之序列表"
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   87
            Tag             =   "1.FIX.SEQ.pdf"
            Top             =   3356
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後修正無劃線之發明申請專利範圍"
            Height          =   195
            Index           =   12
            Left            =   60
            TabIndex        =   88
            Tag             =   "1.FIX_CLAIMS.pdf"
            Top             =   3642
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之序列表"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   78
            Tag             =   "1.COR.SEQ.pdf"
            Top             =   782
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之發明摘要"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   76
            Tag             =   "1.COR_ABSTRACT.pdf"
            Top             =   210
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之發明說明書"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   77
            Tag             =   "1.COR_DESCRIPTION.pdf"
            Top             =   496
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之發明申請專利範圍"
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   79
            Tag             =   "1.COR_CLAIMS.pdf"
            Top             =   1068
            Width           =   3375
         End
         Begin VB.CheckBox chk1Tab2 
            Caption         =   "訂正後無劃線之發明圖式"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   80
            Tag             =   "1.COR_DRAWINGS.pdf"
            Top             =   1354
            Width           =   3375
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00404000&
            BorderWidth     =   2
            X1              =   3300
            X2              =   6700
            Y1              =   3570
            Y2              =   3570
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00404000&
            BorderWidth     =   2
            X1              =   3300
            X2              =   6700
            Y1              =   150
            Y2              =   150
         End
      End
      Begin VB.Frame Frame307 
         Appearance      =   0  '平面
         Caption         =   "分割"
         ForeColor       =   &H00FF0000&
         Height          =   1005
         Left            =   -74820
         TabIndex        =   177
         Top             =   330
         Width           =   9735
         Begin VB.Frame FramePA158 
            Appearance      =   0  '平面
            BorderStyle     =   0  '沒有框線
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5640
            TabIndex        =   246
            Top             =   690
            Width           =   2745
            Begin VB.ComboBox Combo3 
               Height          =   300
               ItemData        =   "frm090904_1.frx":00C4
               Left            =   825
               List            =   "frm090904_1.frx":00C6
               TabIndex        =   247
               Top             =   0
               Width           =   1785
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "案件屬性:"
               Height          =   180
               Index           =   168
               Left            =   0
               TabIndex        =   248
               Top             =   40
               Width           =   765
            End
         End
         Begin VB.CheckBox Check4 
            Caption         =   "電子提申"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   8550
            TabIndex        =   18
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "未附英文說明書，申請書及摘要附有英文翻譯，可減免800元"
            Height          =   255
            Left            =   4020
            TabIndex        =   16
            Top             =   450
            Width           =   5115
         End
         Begin VB.CheckBox Check2 
            Caption         =   "一併提實審"
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   180
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "援用原申請案就相同創作於申請日同日-另申請新型專利之聲明"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   720
            Width           =   5295
         End
         Begin VB.ComboBox cboFavReason 
            Height          =   300
            ItemData        =   "frm090904_1.frx":00C8
            Left            =   5835
            List            =   "frm090904_1.frx":00D5
            Style           =   2  '單純下拉式
            TabIndex        =   14
            Top             =   135
            Width           =   3765
         End
         Begin VB.TextBox txtFavDate 
            Height          =   270
            Left            =   3765
            MaxLength       =   7
            TabIndex        =   13
            Top             =   150
            Width           =   1095
         End
         Begin VB.CheckBox Check9 
            Caption         =   "本案續行再審查"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   15
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label Label800 
            Caption         =   "- 800"
            ForeColor       =   &H000000C0&
            Height          =   165
            Left            =   9180
            TabIndex        =   202
            Top             =   480
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblFavReason 
            AutoSize        =   -1  'True
            Caption         =   "公開事由:"
            Height          =   180
            Left            =   5040
            TabIndex        =   179
            Top             =   210
            Width           =   765
         End
         Begin VB.Label lblFavDate 
            AutoSize        =   -1  'True
            Caption         =   "優惠期發生日期:"
            Height          =   180
            Left            =   2430
            TabIndex        =   178
            Top             =   210
            Width           =   1305
         End
      End
      Begin VB.Frame Frame416 
         Appearance      =   0  '平面
         Caption         =   "實體審查"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   -74820
         TabIndex        =   176
         Top             =   1360
         Width           =   2595
         Begin VB.CheckBox Check416_1 
            Caption         =   "一併申請修正"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   0
            Left            =   510
            TabIndex        =   20
            Top             =   180
            Width           =   1605
         End
         Begin VB.CheckBox Check416_2 
            Caption         =   "一併申請誤譯訂正"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   0
            Left            =   510
            TabIndex        =   21
            Top             =   420
            Width           =   1785
         End
      End
      Begin VB.Frame FramePage 
         Height          =   3255
         Left            =   -74820
         TabIndex        =   165
         Top             =   2100
         Width           =   3555
         Begin VB.TextBox txtDocCh 
            Enabled         =   0   'False
            Height          =   270
            Index           =   7
            Left            =   3030
            MaxLength       =   4
            TabIndex        =   28
            Top             =   1318
            Width           =   420
         End
         Begin VB.TextBox txtSimplified 
            Enabled         =   0   'False
            Height          =   270
            Left            =   3030
            TabIndex        =   35
            Top             =   2925
            Width           =   420
         End
         Begin VB.CheckBox chkDoc 
            Caption         =   "簡體字本資訊"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   34
            Top             =   2955
            Width           =   1455
         End
         Begin VB.ComboBox cboLagnuage 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "frm090904_1.frx":011B
            Left            =   1680
            List            =   "frm090904_1.frx":013A
            Style           =   2  '單純下拉式
            TabIndex        =   24
            Top             =   440
            Width           =   1770
         End
         Begin VB.CheckBox chkDoc 
            Caption         =   "中文本資訊"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   25
            Top             =   818
            Width           =   1230
         End
         Begin VB.TextBox txtDocCh 
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   3030
            MaxLength       =   4
            TabIndex        =   26
            Top             =   780
            Width           =   420
         End
         Begin VB.TextBox txtDocCh 
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   3030
            MaxLength       =   4
            TabIndex        =   27
            Top             =   1049
            Width           =   420
         End
         Begin VB.TextBox txtDocCh 
            Enabled         =   0   'False
            Height          =   270
            Index           =   2
            Left            =   3030
            MaxLength       =   4
            TabIndex        =   29
            Top             =   1587
            Width           =   420
         End
         Begin VB.TextBox txtDocCh 
            Enabled         =   0   'False
            Height          =   270
            Index           =   3
            Left            =   3030
            MaxLength       =   4
            TabIndex        =   30
            Top             =   1856
            Width           =   420
         End
         Begin VB.TextBox txtDocCh 
            Height          =   270
            Index           =   4
            Left            =   3030
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   31
            Top             =   2125
            Width           =   420
         End
         Begin VB.TextBox txtDocCh 
            Enabled         =   0   'False
            Height          =   270
            Index           =   5
            Left            =   3030
            MaxLength       =   4
            TabIndex        =   32
            Top             =   2394
            Width           =   420
         End
         Begin VB.TextBox txtDocCh 
            Enabled         =   0   'False
            Height          =   270
            Index           =   6
            Left            =   3030
            TabIndex        =   33
            Top             =   2663
            Width           =   420
         End
         Begin VB.CheckBox chkDoc 
            Caption         =   "外文本資訊"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   22
            Top             =   195
            Width           =   1230
         End
         Begin VB.TextBox txtForeign 
            Enabled         =   0   'False
            Height          =   270
            Left            =   3030
            TabIndex        =   23
            Top             =   150
            Width           =   420
         End
         Begin VB.Label Label1 
            Caption         =   "序列表："
            Height          =   180
            Index           =   71
            Left            =   1500
            TabIndex        =   218
            Top             =   1361
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "不算超頁費"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.5
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   75
            Left            =   2220
            TabIndex        =   217
            Top             =   1361
            Width           =   900
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "簡體字頁數總計:"
            Height          =   180
            Left            =   1500
            TabIndex        =   175
            Top             =   2970
            Width           =   1305
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "外文本種類:"
            Height          =   180
            Left            =   645
            TabIndex        =   174
            Top             =   500
            Width           =   945
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "摘要頁數:"
            Height          =   180
            Left            =   1500
            TabIndex        =   173
            Top             =   825
            Width           =   765
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "說明書頁數:"
            Height          =   180
            Left            =   1500
            TabIndex        =   172
            Top             =   1093
            Width           =   945
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍頁數:"
            Height          =   180
            Left            =   1500
            TabIndex        =   171
            Top             =   1629
            Width           =   1485
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "圖式頁數:"
            Height          =   180
            Left            =   1500
            TabIndex        =   170
            Top             =   1897
            Width           =   765
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "頁數總計:"
            ForeColor       =   &H00000040&
            Height          =   180
            Left            =   1500
            TabIndex        =   169
            Top             =   2165
            Width           =   765
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "申請專利範圍項數:"
            ForeColor       =   &H00000040&
            Height          =   180
            Left            =   1500
            TabIndex        =   168
            Top             =   2433
            Width           =   1485
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "圖式圖數:"
            Height          =   180
            Left            =   1500
            TabIndex        =   167
            Top             =   2701
            Width           =   765
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "外文頁數總計:"
            Height          =   180
            Left            =   1500
            TabIndex        =   166
            Top             =   195
            Width           =   1125
         End
      End
      Begin VB.Frame Frame204 
         Appearance      =   0  '平面
         Caption         =   "修正"
         ForeColor       =   &H00C00000&
         Height          =   4965
         Left            =   -74970
         TabIndex        =   161
         Top             =   360
         Width           =   3135
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之設計說明書"
            Height          =   195
            Index           =   16
            Left            =   60
            TabIndex        =   72
            Tag             =   "3.FIX_DESCRIPTION.pdf"
            Top             =   3915
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "申請專利範圍對應表"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   19
            Left            =   60
            TabIndex        =   75
            Tag             =   " .TBL.pdf"
            Top             =   4710
            Visible         =   0   'False
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之發明圖式"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   60
            Tag             =   "1.FIX_DRAWINGS.pdf"
            Top             =   1110
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之發明申請專利範圍"
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   59
            Tag             =   "1.FIX_CLAIMS.pdf"
            Top             =   885
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之發明說明書"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   57
            Tag             =   "1.FIX_DESCRIPTION.pdf"
            Top             =   450
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之發明摘要"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   56
            Tag             =   "1.FIX_ABSTRACT.pdf"
            Top             =   240
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正部分劃線之發明申請專利範圍"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   64
            Tag             =   "1.FIX_U_CLAIMS.pdf"
            Top             =   1980
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正部分劃線之序列表"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   63
            Tag             =   "1.FIX_U.SEQ.pdf"
            Top             =   1770
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正部分劃線之發明說明書"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   62
            Tag             =   "1.FIX_U_DESCRIPTION.pdf"
            Top             =   1545
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正部分劃線之發明摘要"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   61
            Tag             =   "1.FIX_U_ABSTRACT.pdf"
            Top             =   1320
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之新型說明書"
            Height          =   195
            Index           =   10
            Left            =   60
            TabIndex        =   66
            Tag             =   "2.FIX_DESCRIPTION.pdf"
            Top             =   2490
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之新型摘要"
            Height          =   195
            Index           =   9
            Left            =   60
            TabIndex        =   65
            Tag             =   "2.FIX_ABSTRACT.pdf"
            Top             =   2265
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之序列表"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   58
            Tag             =   "1.FIX.SEQ.pdf"
            Top             =   675
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之新型申請專利範圍"
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   67
            Tag             =   "2.FIX_CLAIMS.pdf"
            Top             =   2700
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正部分劃線之設計說明書"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   18
            Left            =   60
            TabIndex        =   74
            Tag             =   "3.FIX_U_DESCRIPTION.pdf"
            Top             =   4350
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正部分劃線之新型說明書"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   14
            Left            =   60
            TabIndex        =   70
            Tag             =   "2.FIX_U_DESCRIPTION.pdf"
            Top             =   3360
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正部分劃線之新型申請專利範圍"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   15
            Left            =   60
            TabIndex        =   71
            Tag             =   "2.FIX_U_CLAIMS.pdf"
            Top             =   3585
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之設計圖式"
            Height          =   195
            Index           =   17
            Left            =   60
            TabIndex        =   73
            Tag             =   "3.FIX_DRAWINGS.pdf"
            Top             =   4140
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正後之新型圖式"
            Height          =   195
            Index           =   12
            Left            =   60
            TabIndex        =   68
            Tag             =   "2.FIX_DRAWINGS.pdf"
            Top             =   2925
            Width           =   3045
         End
         Begin VB.CheckBox chk1Tab1 
            Caption         =   "修正部分劃線之新型摘要"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   13
            Left            =   60
            TabIndex        =   69
            Tag             =   "2.FIX_U_ABSTRACT.pdf"
            Top             =   3135
            Width           =   3045
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404000&
            BorderWidth     =   2
            X1              =   -30
            X2              =   3150
            Y1              =   3900
            Y2              =   3900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00404000&
            BorderWidth     =   2
            X1              =   0
            X2              =   3180
            Y1              =   2250
            Y2              =   2250
         End
      End
      Begin VB.Label Label36 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "本次頁數應退還規費"
         Height          =   180
         Left            =   -68700
         TabIndex        =   297
         Top             =   1665
         Width           =   1620
      End
      Begin VB.Label Label28 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "本次頁數應加收規費"
         Height          =   180
         Left            =   -68700
         TabIndex        =   285
         Top             =   1365
         Width           =   1620
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404000&
         BorderWidth     =   2
         X1              =   -74640
         X2              =   -71460
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00404000&
         BorderWidth     =   2
         X1              =   -74640
         X2              =   -71460
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.TextBox txtCP84 
      Height          =   270
      Left            =   7440
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame FrameFee 
      Height          =   1245
      Left            =   60
      TabIndex        =   192
      Top             =   930
      Width           =   6495
      Begin VB.TextBox txtCP136 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         Left            =   5805
         TabIndex        =   304
         Top             =   360
         Width           =   420
      End
      Begin VB.TextBox txtCP135 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         Left            =   5805
         TabIndex        =   303
         Top             =   90
         Width           =   420
      End
      Begin VB.TextBox txtCP135_tmp 
         Height          =   270
         Left            =   5970
         TabIndex        =   203
         Top             =   -75
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtAddItem 
         Height          =   270
         Left            =   3870
         TabIndex        =   1
         Top             =   135
         Width           =   420
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1950
         TabIndex        =   0
         Top             =   135
         Width           =   420
      End
      Begin VB.TextBox txtCP138 
         Height          =   270
         Left            =   3870
         TabIndex        =   3
         Top             =   405
         Width           =   420
      End
      Begin VB.TextBox txtCount 
         Height          =   270
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   675
         Width           =   420
      End
      Begin VB.TextBox txtAddItemFee 
         Height          =   270
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   945
         Width           =   840
      End
      Begin VB.TextBox txtDecreaseItemFee 
         Height          =   270
         Left            =   3870
         TabIndex        =   6
         Top             =   945
         Width           =   840
      End
      Begin VB.TextBox txtCP137 
         Height          =   270
         Left            =   1950
         TabIndex        =   2
         Top             =   405
         Width           =   420
      End
      Begin VB.Label LabelCP136 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "CP136:"
         Height          =   180
         Left            =   5250
         TabIndex        =   306
         Top             =   405
         Width           =   525
      End
      Begin VB.Label lblPageCount 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "CP135:"
         Height          =   180
         Left            =   5250
         TabIndex        =   305
         Top             =   135
         Width           =   525
      End
      Begin VB.Label lblPS1 
         AutoSize        =   -1  'True
         Caption         =   "P.S 請到「變更頁數」頁籤，輸入完整頁數。"
         ForeColor       =   &H00FF00FF&
         Height          =   480
         Left            =   4770
         TabIndex        =   221
         Top             =   720
         Width           =   1710
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "原請求項數"
         Height          =   180
         Left            =   1035
         TabIndex        =   200
         Top             =   180
         Width           =   900
      End
      Begin VB.Label lblAddItem 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "增加項數"
         Height          =   180
         Left            =   3120
         TabIndex        =   198
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label14 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪除未審項數"
         ForeColor       =   &H00000040&
         Height          =   180
         Left            =   840
         TabIndex        =   197
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label13 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "修正後請求項總項數"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   300
         TabIndex        =   196
         Top             =   720
         Width           =   1620
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "本次項數應收規費"
         Height          =   180
         Left            =   60
         TabIndex        =   195
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label10 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "本次項數應退規費"
         Height          =   180
         Left            =   2400
         TabIndex        =   194
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label4 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪除已審項數"
         ForeColor       =   &H00000040&
         Height          =   180
         Left            =   2760
         TabIndex        =   193
         Top             =   450
         Width           =   1080
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "附件文書"
      Height          =   705
      Left            =   6630
      TabIndex        =   162
      Top             =   1020
      Width           =   3315
      Begin VB.CheckBox chkAtt2 
         Caption         =   "文件檔名"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2145
         TabIndex        =   19
         Top             =   450
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "其他"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   10
         Top             =   210
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CheckBox chkAtt2 
         Caption         =   "文件描述"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2145
         TabIndex        =   11
         Top             =   210
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "基本資料表"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Tag             =   ".CONTACT.pdf"
         Top             =   210
         Value           =   1  '核取
         Width           =   1305
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "申復書"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Tag             =   ".EX.pdf"
         Top             =   450
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   7650
      TabIndex        =   148
      Top             =   60
      Width           =   690
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   8400
      TabIndex        =   149
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   153
      Top             =   12
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   152
      Top             =   12
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   151
      Top             =   12
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2610
      MaxLength       =   2
      TabIndex        =   150
      Top             =   12
      Width           =   375
   End
   Begin MSForms.Label lblCP13T 
      Height          =   180
      Left            =   3960
      TabIndex        =   252
      Top             =   360
      Width           =   690
      BackColor       =   -2147483637
      VariousPropertyBits=   268435483
      Caption         =   "lblCP13T"
      Size            =   "1217;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCP14T 
      Height          =   180
      Left            =   960
      TabIndex        =   251
      Top             =   360
      Width           =   690
      BackColor       =   -2147483637
      VariousPropertyBits=   268435483
      Caption         =   "lblCP14T"
      Size            =   "1217;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   960
      TabIndex        =   250
      Top             =   585
      Width           =   7020
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12382;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   8580
      TabIndex        =   249
      Top             =   1800
      Visible         =   0   'False
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "繳費金額:"
      Height          =   180
      Left            =   6630
      TabIndex        =   199
      Top             =   1860
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Index           =   0
      Left            =   144
      TabIndex        =   160
      Top             =   630
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   180
      Index           =   10
      Left            =   3960
      TabIndex        =   158
      Top             =   12
      Width           =   480
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3144
      TabIndex        =   157
      Top             =   12
      Width           =   768
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3150
      TabIndex        =   156
      Top             =   360
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   150
      TabIndex        =   155
      Top             =   360
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   144
      TabIndex        =   154
      Top             =   12
      Width           =   765
   End
End
Attribute VB_Name = "frm090904_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; Combo1、Label7(12)=>lblCP13T、Label7(11)=>lblCP14T
'Create By Sindy 2017/11/22
Option Explicit

Dim strReceiveNo As String, intWhere As Integer
Dim pa() As String, cp() As String
Dim pageD() As String 'Add By Sindy 2023/3/9
Dim m_CaseNo As String
Dim m_IPOSendDt As String, m_IPOSendData1 As String, m_IPOSendData2 As String
Dim strDivPA11 As String
Dim m_CP43CP10 As String
'************************************************************
Dim m_bol99Case As Boolean '是否99年後申請案件
Dim m_bolChkFee As Boolean '是否需檢查規費
Dim m_bolChkPageItem As Boolean '是否要輸頁數與項數
Dim m_lngOverPageFee As Long, m_lngOverItemFee As Long '超頁費,超項費
Dim m_lngOverPageFeeDiff As Long, m_lngOverItemFeeDiff As Long '超頁費,超項費差額
Dim m_lngRecOverPageFee As Long, m_lngRecOverItemFee As Long '已收文超頁費,超項費 Add by Morgan 2011/6/29
Dim m_FeeMemo As String '規費備註
'Dim m_lngOfficialFee As Long '原始規費
Dim m_bol107NewFee As Boolean '台灣再審是否用102年新規費計算
Dim m_bolFixNewFee As Boolean '台灣修正是否用102年新規費計算
Dim m_strReExamCP27 As String '台灣再審發文日(若再審延期發文日)
Dim bolDelay As Boolean 'Add by Morgan 2004/9/8 是否延期過
Dim m_strDelayCP09 As String 'Added by Morgan 2011/11/11 延期收文號
Dim m_Div416OfficialFee As Long  '分割案實審規費  2010/12/8 add by sonia
Dim m_bolChkItem As Boolean '是否要檢查增刪項數 Add by Morgan 2010/9/27
'************************************************************
Dim m_allPage As String, m_allItem As String '總頁數,總項數
Dim m_WriteNote As String 'Add By Sindy 2018/5/8
Dim bolHad404 As Boolean, strHad404CP09 As String 'Add By Sindy 2018/7/30 檢查是否有申請過延期
Dim oText  As Control  'Added by Lydia 2018/12/27
Dim m_PA162 As String 'Added by Morgan 2019/10/7
Dim m_DivAppPA158 As String '母案的案件屬性 Add By Sindy 2020/3/10
Dim m_AgentName As String 'Add By Sindy 2021/5/10
Dim m_433Fee As String 'Add By Sindy 2022/3/4
Dim m_IsRun As Boolean
Dim bolIsSecond As Boolean 'Add By Sindy 2023/3/17 第2個申請書


Private Sub Check3_Click()
   If Check3.Value = 1 Then
      Label800.Visible = True
   Else
      Label800.Visible = False
   End If
End Sub

'Add By Sindy 2020/8/28
Private Sub Check416_2_Click(Index As Integer)
   If Index = 0 Then
      If Check416_2(Index).Value = 1 Then
         Call Call_PUB_SetOfficialFee_P
      End If
   End If
End Sub

'Add By Sindy 2019/2/1 ex:FCP-60443
Private Sub Check9_Validate(Index As Integer, Cancel As Boolean)
   Call_PUB_SetOfficialFee_P
End Sub

'Add By Sindy 2018/1/17
Private Sub chk0Tab_Click(Index As Integer)
Dim iChecked As Single
   
   If Val(chk0Tab(Index)) > 0 Then
      iChecked = vbChecked
   Else
      iChecked = vbUnchecked
   End If
   Select Case Index
   '發明摘要
   Case 2:
      'If chk0Tab(2).Value = 1 Then
         If chk0Tab(3).Enabled = True Then chk0Tab(3).Value = iChecked
         If chk0Tab(4).Enabled = True Then chk0Tab(4).Value = iChecked
      'End If
   '發明說明書
   Case 3:
      'If chk0Tab(3).Value = 1 Then
         If chk0Tab(2).Enabled = True Then chk0Tab(2).Value = iChecked
         If chk0Tab(4).Enabled = True Then chk0Tab(4).Value = iChecked
      'End If
   '發明專利範圍
   Case 4:
      'If chk0Tab(4).Value = 1 Then
         If chk0Tab(2).Enabled = True Then chk0Tab(2).Value = iChecked
         If chk0Tab(3).Enabled = True Then chk0Tab(3).Value = iChecked
      'End If
   '新型摘要
   Case 15:
      'If chk0Tab(15).Value = 1 Then
         If chk0Tab(16).Enabled = True Then chk0Tab(16).Value = iChecked
         If chk0Tab(17).Enabled = True Then chk0Tab(17).Value = iChecked
      'End If
   '新型說明書
   Case 16:
      'If chk0Tab(16).Value = 1 Then
         If chk0Tab(15).Enabled = True Then chk0Tab(15).Value = iChecked
         If chk0Tab(17).Enabled = True Then chk0Tab(17).Value = iChecked
      'End If
   '新型專利範圍
   Case 17:
      'If chk0Tab(17).Value = 1 Then
         If chk0Tab(15).Enabled = True Then chk0Tab(15).Value = iChecked
         If chk0Tab(16).Enabled = True Then chk0Tab(16).Value = iChecked
      'End If
   End Select
End Sub

Private Sub chk1Tab1_Click(Index As Integer)
Dim chk As CheckBox
Dim bolChkAllFalse As Boolean
Dim iChecked As Single
   
   bolChkAllFalse = False
   For Each chk In chk1Tab1
      If chk.Value = 1 Then
         bolChkAllFalse = True
         '一併申請修正自動打勾
         If SSTab1.TabVisible(0) = True Then
            If Frame416.Enabled = True Then
               Check416_1(0).Value = 1
               Check416_1(1).Value = 1
               'Add By Sindy 2018/4/24
               If Check416_2(0).Value = 1 Then
                  'Check416_1(0).Value = 0
                  Check416_1(1).Value = 0
               End If
               '2018/4/24 END
            End If
         ElseIf SSTab1.TabVisible(2) = True Then
            Check431_1(0).Value = 1
            Check431_1(1).Value = 1
            'Add By Sindy 2018/4/24
            If Check431_2(0).Value = 1 Then
               'Check431_1(0).Value = 0
               Check431_1(1).Value = 0
            End If
            '2018/4/24 END
         ElseIf SSTab1.TabVisible(3) = True Then
            If Frame107.Enabled = True Then
               Check107_1(0).Value = 1
               Check107_1(1).Value = 1
               'Add By Sindy 2018/4/24
               If Check107_2(0).Value = 1 Then
                  'Check107_1(0).Value = 0
                  Check107_1(1).Value = 0
               End If
               '2018/4/24 END
            End If
         '要放IF最後面
         ElseIf SSTab1.TabVisible(1) = True Then
            If Frame433.Enabled = True Then
               Check433.Value = 1
            End If
         End If
         If cp(10) = "205" Then chkAtt(1).Value = 1: Frame204.Tag = "Y"
         Exit For
      End If
   Next
   If bolChkAllFalse = False Then
      '一併申請修正自動取消
      If SSTab1.TabVisible(0) = True Then
         Check416_1(0).Value = 0
         Check416_1(1).Value = 0
      ElseIf SSTab1.TabVisible(2) = True Then
         Check431_1(0).Value = 0
         Check431_1(1).Value = 0
      ElseIf SSTab1.TabVisible(3) = True Then
         Check107_1(0).Value = 0
         Check107_1(1).Value = 0
      '要放IF最後面
      ElseIf SSTab1.TabVisible(1) = True Then
         Check433.Value = 0
      End If
      If cp(10) = "205" Then chkAtt(1).Value = 1: Frame204.Tag = ""
   End If
   
   If Val(chk1Tab1(Index)) > 0 Then
      iChecked = vbChecked
   Else
      iChecked = vbUnchecked
   End If
   
   'If Frame433.Enabled = True Then Exit Sub
   Select Case Index
   '發明摘要
   Case 0:
      'If chk1Tab1(0).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab1(1).Enabled = True Then chk1Tab1(1).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab1(3).Enabled = True Then chk1Tab1(3).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab1(5).Enabled = True Then chk1Tab1(5).Value = iChecked
      'End If
   '發明說明書
   Case 1:
      'If chk1Tab1(1).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab1(0).Enabled = True Then chk1Tab1(0).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab1(3).Enabled = True Then chk1Tab1(3).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab1(6).Enabled = True Then chk1Tab1(6).Value = iChecked
      'End If
   '發明序列表
   Case 2:
      'If chk1Tab1(2).Value = 1 Then 'Add By Sindy 2018/1/15
         If chk1Tab1(7).Enabled = True Then chk1Tab1(7).Value = iChecked
      'End If
   '發明專利範圍
   Case 3:
      'If chk1Tab1(3).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab1(0).Enabled = True Then chk1Tab1(0).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab1(1).Enabled = True Then chk1Tab1(1).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab1(8).Enabled = True Then chk1Tab1(8).Value = iChecked
      'End If
   '新型摘要
   Case 9:
      'If chk1Tab1(9).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab1(10).Enabled = True Then chk1Tab1(10).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab1(11).Enabled = True Then chk1Tab1(11).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab1(13).Enabled = True Then chk1Tab1(13).Value = iChecked
      'End If
   '新型說明書
   Case 10:
      'If chk1Tab1(10).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab1(9).Enabled = True Then chk1Tab1(9).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab1(11).Enabled = True Then chk1Tab1(11).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab1(14).Enabled = True Then chk1Tab1(14).Value = iChecked
      'End If
   '新型專利範圍
   Case 11:
      'If chk1Tab1(11).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab1(9).Enabled = True Then chk1Tab1(9).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab1(10).Enabled = True Then chk1Tab1(10).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab1(15).Enabled = True Then chk1Tab1(15).Value = iChecked
      'End If
   '設計說明書
   Case 16:
      'If chk1Tab1(16).Value = 1 Then 'Add By Sindy 2018/1/15
         If chk1Tab1(18).Enabled = True Then chk1Tab1(18).Value = iChecked
      'End If
   End Select
End Sub

Private Sub chk1Tab2_Click(Index As Integer)
Dim chk As CheckBox
Dim bolChkAllFalse As Boolean
Dim iChecked As Single
   
   bolChkAllFalse = False
   For Each chk In chk1Tab2
      If chk.Value = 1 Then
         bolChkAllFalse = True
         '一併申請誤譯訂正自動打勾
         If SSTab1.TabVisible(0) = True Then
            If Frame416.Enabled = True Then
               Check416_2(0).Value = 1
               Check416_2(1).Value = 1
               'Add By Sindy 2018/4/24
               'Check416_1(0).Value = 0
               Check416_1(1).Value = 0
               '2018/4/24 END
            End If
         ElseIf SSTab1.TabVisible(2) = True Then
            Check431_2(0).Value = 1
            Check431_2(1).Value = 1
            'Add By Sindy 2018/4/24
            'Check431_1(0).Value = 0
            Check431_1(1).Value = 0
            '2018/4/24 END
         ElseIf SSTab1.TabVisible(3) = True Then
            If Frame107.Enabled = True Then
               Check107_2(0).Value = 1
               Check107_2(1).Value = 1
               'Add By Sindy 2018/4/24
               'Check107_1(0).Value = 0
               Check107_1(1).Value = 0
               '2018/4/24 END
            End If
         '要放IF最後面
         ElseIf SSTab1.TabVisible(1) = True Then
            'If Frame433.Enabled = True Then
               If (Index >= 9 And Index <= 13) Or _
                  (Index >= 22 And Index <= 25) Or _
                  (Index >= 29 And Index <= 30) Then
                  Check433.Value = 1
               End If
            'End If
         End If
         Exit For
      End If
   Next
   If bolChkAllFalse = False Then
      '一併申請誤譯訂正自動取消
      If SSTab1.TabVisible(0) = True Then
         Check416_2(0).Value = 0
         Check416_2(1).Value = 0
      ElseIf SSTab1.TabVisible(2) = True Then
         Check431_2(0).Value = 0
         Check431_2(1).Value = 0
      ElseIf SSTab1.TabVisible(3) = True Then
         Check107_2(0).Value = 0
         Check107_2(1).Value = 0
      '要放IF最後面
      ElseIf SSTab1.TabVisible(1) = True Then
         If (Index >= 9 And Index <= 13) Or _
            (Index >= 22 And Index <= 25) Or _
            (Index >= 29 And Index <= 30) Then
            Check433.Value = 0
         End If
      End If
   End If
   
   If Val(chk1Tab2(Index)) > 0 Then
      iChecked = vbChecked
   Else
      iChecked = vbUnchecked
   End If
   
   Select Case Index
   '發明摘要
   Case 0:
      'If chk1Tab2(0).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab2(1).Enabled = True Then chk1Tab2(1).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab2(3).Enabled = True Then chk1Tab2(3).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab2(5).Enabled = True Then chk1Tab2(5).Value = iChecked
      'End If
   '發明說明書
   Case 1:
      'If chk1Tab2(1).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab2(0).Enabled = True Then chk1Tab2(0).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab2(3).Enabled = True Then chk1Tab2(3).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab2(6).Enabled = True Then chk1Tab2(6).Value = iChecked
      'End If
   '發明序列表
   Case 2:
      'If chk1Tab2(2).Value = 1 Then 'Add By Sindy 2018/1/15
         If chk1Tab2(7).Enabled = True Then chk1Tab2(7).Value = iChecked
      'End If
   '發明專利範圍
   Case 3:
      'If chk1Tab2(3).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab2(0).Enabled = True Then chk1Tab2(0).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab2(1).Enabled = True Then chk1Tab2(1).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab2(8).Enabled = True Then chk1Tab2(8).Value = iChecked
      'End If
      
'   'Add By Sindy 2018/1/17
'   '發明摘要/發明說明書/發明專利範圍
'   Case 9, 10, 12:
'      'If chk1Tab2(9).Value = 1 Then
'         If chk1Tab2(9).Enabled = True Then chk1Tab2(9).Value = iChecked
'         If chk1Tab2(10).Enabled = True Then chk1Tab2(10).Value = iChecked
'         If chk1Tab2(12).Enabled = True Then chk1Tab2(12).Value = iChecked
'         'Modify By Sindy 2018/4/9 修正後連動訂正後修正無劃線
'         If chk1Tab1(5).Enabled = True Then chk1Tab1(5).Value = iChecked
'         If chk1Tab1(6).Enabled = True Then chk1Tab1(6).Value = iChecked
'         If chk1Tab1(8).Enabled = True Then chk1Tab1(8).Value = iChecked
''         If chk1Tab1(0).Enabled = True Then chk1Tab1(0).Value = iChecked
''         If chk1Tab1(1).Enabled = True Then chk1Tab1(1).Value = iChecked
''         If chk1Tab1(3).Enabled = True Then chk1Tab1(3).Value = iChecked
'      'End If
   'Add By Sindy 2018/1/17
   '發明摘要/發明說明書/發明專利範圍
   Case 9:
      'If chk1Tab2(9).Value = 1 Then
         If chk1Tab2(9).Enabled = True Then chk1Tab2(9).Value = iChecked
         If chk1Tab1(5).Enabled = True Then chk1Tab1(5).Value = iChecked
      'End If
   Case 10:
      'If chk1Tab2(9).Value = 1 Then
         If chk1Tab2(10).Enabled = True Then chk1Tab2(10).Value = iChecked
         If chk1Tab1(6).Enabled = True Then chk1Tab1(6).Value = iChecked
      'End If
   '發明摘要/發明說明書/發明專利範圍
   Case 12:
      'If chk1Tab2(9).Value = 1 Then
         If chk1Tab2(12).Enabled = True Then chk1Tab2(12).Value = iChecked
         If chk1Tab1(8).Enabled = True Then chk1Tab1(8).Value = iChecked
      'End If
   '發明序列表
   Case 11:
      'If chk1Tab2(11).Value = 1 Then
         If chk1Tab1(7).Enabled = True Then chk1Tab1(7).Value = iChecked
'         If chk1Tab1(2).Enabled = True Then chk1Tab1(2).Value = iChecked
      'End If
   '2018/1/17 END

'   'Add By Sindy 2018/4/9
'   '發明圖式
'   Case 13:
'      'If chk1Tab2(13).Value = 1 Then
'         If chk1Tab1(4).Enabled = True Then chk1Tab1(4).Value = iChecked
'      'End If
'   '2018/4/9 END
   '新型摘要
   Case 14:
      'If chk1Tab2(14).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab2(15).Enabled = True Then chk1Tab2(15).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab2(16).Enabled = True Then chk1Tab2(16).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab2(18).Enabled = True Then chk1Tab2(18).Value = iChecked
      'End If
   '新型說明書
   Case 15:
      'If chk1Tab2(15).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab2(14).Enabled = True Then chk1Tab2(14).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab2(16).Enabled = True Then chk1Tab2(16).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab2(19).Enabled = True Then chk1Tab2(19).Value = iChecked
      'End If
   '新型專利範圍
   Case 16:
      'If chk1Tab2(16).Value = 1 Then 'Add By Sindy 2018/1/15
         'If chk1Tab2(14).Enabled = True Then chk1Tab2(14).Value = iChecked 'Add By Sindy 2018/1/15
         'If chk1Tab2(15).Enabled = True Then chk1Tab2(15).Value = iChecked 'Add By Sindy 2018/1/15
         If chk1Tab2(20).Enabled = True Then chk1Tab2(20).Value = iChecked
      'End If
   '新型圖式
   Case 17:
      'If chk1Tab2(17).Value = 1 Then 'Add By Sindy 2018/1/15
         If chk1Tab2(21).Enabled = True Then chk1Tab2(21).Value = iChecked
      'End If
      
   'Add By Sindy 2018/1/17
'   '新型摘要/新型說明書/新型專利範圍
'   Case 22, 23, 24:
'      'If chk1Tab2(22).Value = 1 Then
'         If chk1Tab2(22).Enabled = True Then chk1Tab2(22).Value = iChecked
'         If chk1Tab2(23).Enabled = True Then chk1Tab2(23).Value = iChecked
'         If chk1Tab2(24).Enabled = True Then chk1Tab2(24).Value = iChecked
'         If chk1Tab1(13).Enabled = True Then chk1Tab1(13).Value = iChecked
'         If chk1Tab1(14).Enabled = True Then chk1Tab1(14).Value = iChecked
'         If chk1Tab1(15).Enabled = True Then chk1Tab1(15).Value = iChecked
''         If chk1Tab1(9).Enabled = True Then chk1Tab1(9).Value = iChecked
''         If chk1Tab1(10).Enabled = True Then chk1Tab1(10).Value = iChecked
''         If chk1Tab1(11).Enabled = True Then chk1Tab1(11).Value = iChecked
'      'End If
   Case 22:
      'If chk1Tab2(22).Value = 1 Then
         If chk1Tab2(22).Enabled = True Then chk1Tab2(22).Value = iChecked
         If chk1Tab1(13).Enabled = True Then chk1Tab1(13).Value = iChecked
      'End If
   '新型摘要/新型說明書/新型專利範圍
   Case 23:
      'If chk1Tab2(22).Value = 1 Then
         If chk1Tab2(23).Enabled = True Then chk1Tab2(23).Value = iChecked
         If chk1Tab1(14).Enabled = True Then chk1Tab1(14).Value = iChecked
      'End If
   '新型摘要/新型說明書/新型專利範圍
   Case 24:
      'If chk1Tab2(22).Value = 1 Then
         If chk1Tab2(24).Enabled = True Then chk1Tab2(24).Value = iChecked
         If chk1Tab1(13).Enabled = True Then chk1Tab1(13).Value = iChecked
      'End If
   '2018/1/17 END
   
'   'Add By Sindy 2018/4/9
'   '新型圖式
'   Case 25:
'      'If chk1Tab2(25).Value = 1 Then
'         If chk1Tab1(12).Enabled = True Then chk1Tab1(12).Value = iChecked
'      'End If
'   '2018/4/9 END
   '設計說明書
   Case 26:
      'If chk1Tab2(26).Value = 1 Then 'Add By Sindy 2018/1/15
         If chk1Tab2(28).Enabled = True Then chk1Tab2(28).Value = iChecked
      'End If
   'Add By Sindy 2018/1/17
   '設計說明書
   Case 29:
      'If chk1Tab2(29).Value = 1 Then
         If chk1Tab1(18).Enabled = True Then chk1Tab1(18).Value = iChecked
'         If chk1Tab1(16).Enabled = True Then chk1Tab1(16).Value = iChecked
      'End If
   '2018/1/17 END
'   'Add By Sindy 2018/4/9
'   '設計圖式
'   Case 30:
'      'If chk1Tab2(30).Value = 1 Then
'         If chk1Tab1(17).Enabled = True Then chk1Tab1(17).Value = iChecked
'      'End If
'   '2018/4/9 END
   End Select
End Sub

'Add By Sindy 2019/10/31 更正
Private Sub chk4Tab1_Click(Index As Integer)
Dim chk As CheckBox
Dim iChecked As Single
   
   'Modify By Sindy 2023/8/23 + 敏莉說工程師出"更正"申請書時，介面勾選修正哪一欄位
   '                            (例：申請專利範圍)則申請書僅出那一道，其他欄位（例：摘要、說明書）不要一併自動勾選
'   If Val(chk4Tab1(Index)) > 0 Then
'      iChecked = vbChecked
'   Else
'      iChecked = vbUnchecked
'   End If
'
'   Select Case Index
'   '發明摘要
'   Case 0:
'      If chk4Tab1(1).Enabled = True Then chk4Tab1(1).Value = iChecked
'      If chk4Tab1(3).Enabled = True Then chk4Tab1(3).Value = iChecked
'   '發明說明書
'   Case 1:
'      If chk4Tab1(0).Enabled = True Then chk4Tab1(0).Value = iChecked
'      If chk4Tab1(3).Enabled = True Then chk4Tab1(3).Value = iChecked
'   '發明專利範圍
'   Case 3:
'      If chk4Tab1(0).Enabled = True Then chk4Tab1(0).Value = iChecked
'      If chk4Tab1(1).Enabled = True Then chk4Tab1(1).Value = iChecked
'   '新型摘要
'   Case 5:
'      If chk4Tab1(6).Enabled = True Then chk4Tab1(6).Value = iChecked
'      If chk4Tab1(7).Enabled = True Then chk4Tab1(7).Value = iChecked
'   '新型說明書
'   Case 6:
'      If chk4Tab1(5).Enabled = True Then chk4Tab1(5).Value = iChecked
'      If chk4Tab1(7).Enabled = True Then chk4Tab1(7).Value = iChecked
'   '新型專利範圍
'   Case 7:
'      If chk4Tab1(5).Enabled = True Then chk4Tab1(5).Value = iChecked
'      If chk4Tab1(6).Enabled = True Then chk4Tab1(6).Value = iChecked
'   End Select
End Sub

Private Sub chkAtt_Click(Index As Integer)
   If chkAtt(2).Value = 1 Then
      chkAtt2(0).Value = 1
      chkAtt2(1).Value = 1
   Else
      chkAtt2(0).Value = 0
      chkAtt2(1).Value = 0
   End If
End Sub

'Add By Sindy 2024/8/21
Private Sub chkAtt421_1_Click(Index As Integer)
   If Index = 1 Then
      chkAtt421_3(0).Value = chkAtt421_1(Index).Value
   ElseIf Index = 2 Then
      chkAtt421_3(1).Value = chkAtt421_1(Index).Value
   ElseIf Index = 3 Then
      chkAtt421_3(2).Value = chkAtt421_1(Index).Value
   End If
End Sub
Private Sub chkAtt421_3_Click(Index As Integer)
   If Index = 0 Then
      chkAtt421_1(1).Value = chkAtt421_3(Index).Value
   ElseIf Index = 1 Then
      chkAtt421_1(2).Value = chkAtt421_3(Index).Value
   ElseIf Index = 2 Then
      chkAtt421_1(3).Value = chkAtt421_3(Index).Value
   End If
End Sub
'2024/8/21 END

Private Sub chkDoc_Click(Index As Integer)
   'Modified by Lydia 2018/12/27
   'Dim oText As TextBox, bEnabled As Boolean
   Dim bEnabled As Boolean
   
   If chkDoc(Index).Value = 1 Then
      bEnabled = True
   Else
      bEnabled = False
   End If

   Select Case Index
   Case 0 '中文
      For Each oText In txtDocCh
         oText.Enabled = bEnabled
      Next
      If pa(8) = "3" Then
         txtDocCh(0).Enabled = False
         txtDocCh(0).BackColor = Me.BackColor
         txtDocCh(2).Enabled = False
         txtDocCh(2).BackColor = Me.BackColor
         txtDocCh(5).Enabled = False
         txtDocCh(5).BackColor = Me.BackColor
         'Add By Sindy 2019/8/7 Ex:FCP-061566
         txtDocCh(7).Enabled = False
         txtDocCh(7).BackColor = Me.BackColor
         '2019/8/7 END
      End If
   Case 1 '外文
      txtForeign.Enabled = bEnabled
      cboLagnuage.Enabled = bEnabled
      chk0Tab(6).Value = chkDoc(Index).Value
   Case 2 '簡體
      txtSimplified.Enabled = bEnabled
      chk0Tab(9).Value = chkDoc(Index).Value
   End Select
End Sub

'Add By Sindy 2020/3/10
Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 <> "" Then
      Combo3 = Left(Combo3, 1) + "." + PUB_GetCaseAttributeName(Left(Combo3, 1), pa(8))
      If Combo3 = Left(Combo3, 1) + "." Then
         Combo3 = Left(Combo3, 1)
         Cancel = True
         Combo3.SetFocus
      End If
   End If
End Sub
'2020/3/10 End

Private Sub Form_Activate()
   If m_IsRun = False Then
      m_IsRun = True
      
      'Add By Sindy 2020/3/10
      If cp(10) = "307" Then 'And FramePA158.Visible = True
         Call PUB_AddCaseAttributeCombo(Combo3, pa(8)) '專利案件屬性選單
         If pa(158) = "" Then
            '預設母案的案件屬性
            Call PUB_DivAppDate(pa(1), pa(2), pa(3), pa(4), , , m_DivAppPA158) '取得母案案件屬性
            Combo3 = m_DivAppPA158 + "." + PUB_GetCaseAttributeName(m_DivAppPA158, pa(8))
            Combo3.Tag = ""
         Else
            Combo3 = pa(158) + "." + PUB_GetCaseAttributeName(pa(158), pa(8))
            Combo3.Tag = Combo3.Text
         End If
      End If
      '2020/3/10 END
      
      'Modify By Sindy 2019/3/19
      '計算規費:
      Call_PUB_SetOfficialFee_P 'Add By Sindy 2018/6/13
      '2019/3/19 END
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm090904
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   If Pub_StrUserSt03 <> "M51" Then
      LabelCP136.Visible = False
      txtCP136.Visible = False
      'Add By Sindy 2023/3/10
      lblPageCount.Visible = False
      txtCP135.Visible = False
      '2023/3/10 END
   End If
   
   m_IsRun = False 'Add By Sindy 2023/3/15
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090904_1 = Nothing
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Public Sub ReadData(Optional ByRef bolShowForm As Boolean)
Dim chk As CheckBox
Dim opt As OptionButton
Dim oText As TextBox
Dim intQ As Integer, ii As Integer
Dim strCE01 As String 'Add By Sindy 2023/2/24
   
   bolShowForm = True
   
   ReDim pageD(1 To 21) As String 'Add By Sindy 2023/3/9
   ReDim pa(1 To TF_PA) As String
   ReDim cp(TF_CP)
   
   '專利基本檔
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   'Add By Sindy 2023/8/11
   '讀取ServicePractice服務業務檔
   If pa(1) = "FG" Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
      End If
   Else
   '2023/8/11 END
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
      End If
   End If
   m_PA162 = pa(162) ' Added by Morgan 2019/10/17
   
   '進度檔
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      lblCP13T = GetPrjSalesNM(cp(13))
      lblCP14T = GetPrjSalesNM(cp(14))
      If ClsPDGetCaseProperty("FCP", cp(10), strExc(0)) Then Label7(10) = strExc(0)
   End If
   
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10), True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300
   
   '來函文號
   'Modify By Sindy 2018/5/11 先抓相關總收文號 ex:FCP-057221(申復)
   'Modify By Sindy 2019/7/26 調整SQL
   'strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09='" & cp(9) & "' AND ed11(+)=cp43 and cp43 is not null ORDER BY CP05 DESC"
   strExc(0) = "SELECT cp08,ed08,cp09 FROM caseprogress,edocument,(SELECT cp43 FROM caseprogress" & _
               " where CP09='" & cp(9) & "' AND cp43 IS NOT NULL) A" & _
               " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
               " AND CP09=A.cp43 AND ed11(+)=A.cp43" & _
               " ORDER BY CP05 DESC"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp("ED08")) Then
         m_IPOSendDt = RsTemp("ED08") - 19110000
         If Not IsNull(RsTemp("cp08")) Then
            strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
            m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
            strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
            'Modify By Sindy 2018/4/16 不要號字
            m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
         End If
      ElseIf Not IsNull(RsTemp("cp08")) Then
            strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
            m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
            strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
            'Modify By Sindy 2018/4/16 不要號字
            m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
      End If
   End If
   If m_IPOSendDt = "" Then
   '2018/5/11 END
      'Add By Sindy 2022/4/13 擇一申復:請優先抓1232通知擇一申復函號，若無則抓1202審查意見通知函函號
      'Modify By Sindy 2024/5/28 補文件:請抓通知修正（1201）函號
      If cp(10) = "239" Or cp(10) = 補文件 Then
         If cp(10) = "239" Then
            strExc(10) = "1202"
         Else
            strExc(10) = "1201"
         End If
         '2024/5/28 END
         strExc(0) = "select cp08,ed08,cp09 from caseprogress,edocument" & _
                     " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                     " AND CP10='" & strExc(10) & "' AND ed11(+)=cp09 ORDER BY CP05 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(RsTemp("ED08")) Then
               m_IPOSendDt = RsTemp("ED08") - 19110000
               If Not IsNull(RsTemp("cp08")) Then
                  strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                  m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
                  'Modify By Sindy 2018/4/16 不要號字
                  m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
               End If
            ElseIf Not IsNull(RsTemp("cp08")) Then
                  strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                  m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
                  'Modify By Sindy 2018/4/16 不要號字
                  m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
            End If
         End If
      End If
      '2022/4/13 END
      If m_IPOSendDt = "" Then
         'Added by Morgan 2022/5/12 435 續行母案再審抓分割的機關文號 --陳亭妙
         If cp(10) = "435" Then
            strExc(0) = "select cp08,ed08,cp09 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                        " AND CP10='307' AND ed11(+)=cp09"
         Else
         'end 2022/5/12
            strExc(0) = "select cp08,ed08,cp09 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                        " AND CP09='" & cp(9) & "' AND ed11(+)=cp09 ORDER BY CP05 DESC"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(RsTemp("ED08")) Then
               m_IPOSendDt = RsTemp("ED08") - 19110000
               If Not IsNull(RsTemp("cp08")) Then
                  strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                  m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
                  'Modify By Sindy 2018/4/16 不要號字
                  m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
               End If
            ElseIf Not IsNull(RsTemp("cp08")) Then
                  strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                  m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
                  'Modify By Sindy 2018/4/16 不要號字
                  m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
            End If
         End If
      End If
   End If
   
   cmdOK(0).Tag = "" 'Add By Sindy 2023/8/11
   '**********************************************
   '直接產生申請書: 408面詢,407請求面詢,403更改
   Select Case cp(10)
      'Modify By Sindy 2022/4/13 + 239擇一申復
      'Modify By Sindy 2023/8/11 + 230提供情報
      'Modify By Sindy 2024/5/28 + 202補文件
      'Modify By Sindy 2025/2/18 修正:補充說明 是工程師操作, 產生專利補正文件申請書
      Case "408", "407", "403", "239", "230", 補文件, 補充說明
         bolShowForm = False
         cmdOK(0).Tag = "僅產生申請書" 'Add By Sindy 2023/8/11
'         intQ = MsgBox("產生面詢申請書，附件文書是否有「其他(文件描述、文件檔名)」？" & vbCrLf & _
'                   "是：有「其他」" & vbCrLf & _
'                   "否：無「其他」" & vbCrLf & _
'                   "取消：放棄產生申請書", vbInformation + vbYesNoCancel + vbDefaultButton1)
         intQ = MsgBox("是否要產生申請書？", vbInformation + vbYesNo + vbDefaultButton1)
'         If intQ = vbYes Or intQ = vbNo Then
'            If intQ = vbYes Then
'               chkAtt(2).Value = 1
'            Else
'               chkAtt(2).Value = 0
'            End If
'            Call cmdOK_Click(0)
'            Exit Sub
'         End If
         If intQ = vbYes Then
            Call cmdok_Click(0)
            Exit Sub
         End If
         Call cmdok_Click(1)
         Exit Sub
   End Select
   '**********************************************
   
   'Added by Lydia 2018/12/27 中文本資訊-各項頁數
   '---分割/實審
   txtDocCh(0).Text = pa(64) '摘要頁數
   txtDocCh(1).Text = pa(65) '說明書頁數
   txtDocCh(7).Text = pa(66) '序列表頁數
   txtDocCh(2).Text = pa(67) '申請專利範圍頁數
   txtDocCh(3).Text = pa(68) '圖式頁數
   'Added by Lydia 2019/01/10
   txtDocCh(5).Text = pa(172) '申請專利範圍項數(最初項數)
   txtDocCh(6).Text = pa(173) '圖式圖數
   'end 2019/01/10
   If Val(pa(64)) + Val(pa(65)) + Val(pa(66)) + Val(pa(67)) + Val(pa(68)) > 0 Then
       chkDoc(0).Value = 1
       Call txtDocCh_Validate(0, False)
       Call txtDocCh_Validate(1, False)
       Call txtDocCh_Validate(2, False)
       Call txtDocCh_Validate(3, False)
   End If
   For Each oText In txtDocCh
       oText.Tag = oText.Text
   Next
   '---其他->再審查
   txtDocCh2(0).Text = pa(64) '摘要頁數
   txtDocCh2(1).Text = pa(65) '說明書頁數
   txtDocCh2(7).Text = pa(66) '序列表頁數
   txtDocCh2(2).Text = pa(67) '申請專利範圍頁數
   txtDocCh2(3).Text = pa(68) '圖式頁數
   'Added by Lydia 2019/01/10
   txtDocCh2(5).Text = pa(172) '申請專利範圍項數(最初項數)
   txtDocCh2(6).Text = pa(173) '圖式圖數
   'end 2019/01/10
   If Val(pa(64)) + Val(pa(65)) + Val(pa(66)) + Val(pa(67)) + Val(pa(68)) > 0 Then
       chkDoc(0).Value = 1
       Call txtDocCh2_Validate(0, False)
       Call txtDocCh2_Validate(1, False)
       Call txtDocCh2_Validate(2, False)
       Call txtDocCh2_Validate(3, False)
   End If
   For Each oText In txtDocCh2
       oText.Tag = oText.Text
   Next
   'end 2018/12/27
   'Added by Lydia 2019/01/03 修正->變更頁數
   txtDocCh3(0).Text = pa(64) '摘要頁數
   txtDocCh3(1).Text = pa(65) '說明書頁數
   txtDocCh3(3).Text = pa(67) '申請專利範圍頁數
   txtDocCh3(4).Text = pa(68) '圖式頁數
   'Add By Sindy 2023/3/29
   txtDocCh4(7).Text = pa(173) '圖式圖數
   '2023/3/29 END
   If Val(pa(64)) + Val(pa(65)) + Val(pa(67)) + Val(pa(68)) > 0 Then
      Call txtDocCh3_Validate(0, False)
   End If
   For Each oText In txtDocCh3
       oText.Tag = oText.Text
   Next
   
   '讀取原總頁數和原總項數(統計已發文)
   'Modified by Lydia 2018/12/27 預設基本檔的頁數總計
   'm_allPage = 0: m_allItem = 0
   m_allPage = Val(txtDocCh(4))
   m_allItem = Val(txtDocCh(5))
   'end 2018/12/27
   'Modify By Sindy 2023/3/9 改成共用函數
   '取得總頁數/總項數
   Call PUB_GetAllPageItem(strReceiveNo, cp, pa, m_allPage, m_allItem)
   'Add By Sindy 2023/5/2 有關工程師出實審和再審申請書，介面帶的目前項數，
   '                      請用進度檔算出來而不要直接抓基本檔的項數 ex:FCP-60771
   If m_allItem > 0 Then
      txtDocCh(5).Text = m_allItem
      txtDocCh2(5).Text = m_allItem
   End If
   '2023/5/2 END
   '讀取專利說明書頁數明細
   Call PUB_ReadPageDetail(strReceiveNo, pageD)
   '增加頁數
   txtDocAdd(0) = pageD(2)
   txtDocAdd(1) = pageD(3)
   txtDocAdd(3) = pageD(4)
   txtDocAdd(4) = pageD(5)
   '刪除未審頁數
   txtDocCp167(0) = pageD(6)
   txtDocCp167(1) = pageD(7)
   txtDocCp167(3) = pageD(8)
   txtDocCp167(4) = pageD(9)
   '刪除已審頁數
   txtDocCp168(0) = pageD(10)
   txtDocCp168(1) = pageD(11)
   txtDocCp168(3) = pageD(12)
   txtDocCp168(4) = pageD(13)
   '計算頁數合計
   Call CountPage
   '2023/3/9 END
   
   '*************************************************************************
   'Added by Morgan 2013/1/8
   m_bol107NewFee = True
   bolDelay = False
   'end 2013/1/8
   'Modify by Morgan 2006/8/18 加判斷107(再審),803(舉發),301,302,303,305(改請)才要
   'Modified by Morgan 2013/8/26 +507 -- FCP032929
   If InStr("107,803,301,302,303,305,507", cp(10)) > 0 Then
      'Add by Morgan 2004/9/8 檢查是否有延期，若有則規費預設0
      bolDelay = PUB_ChkDelay(strReceiveNo, m_strDelayCP09, strExc(1))
      If bolDelay = True Then
         If strExc(1) < "20130101" Then m_bol107NewFee = False 'Added by Morgan 2013/1/9
         cp(17) = "0"
      End If
   End If

   txtCP84.Tag = cp(17)
   txtCP84.Text = txtCP84.Tag
   
   '頁數,項數欄位
   m_bolChkFee = False
   txtAddItem.Enabled = False
   txtAddItem.BackColor = Me.BackColor
   txtCP137.Enabled = False
   txtCP137.BackColor = Me.BackColor
   txtCP138.Enabled = False
   txtCP138.BackColor = Me.BackColor
   txtCount.Enabled = False
   txtCount.BackColor = Me.BackColor
   txtAddItemFee.Enabled = False
   txtAddItemFee.BackColor = Me.BackColor
   txtDecreaseItemFee.Enabled = False
   txtDecreaseItemFee.BackColor = Me.BackColor
   'txtCP84.Enabled = True: txtCP84.Locked = False 'Modify By Sindy 2018/7/27 Mark
   
   '依不同的案件性質,顯示不同的頁籤
   SSTab1.TabVisible(0) = False
   SSTab1.TabVisible(1) = False
   SSTab1.TabVisible(2) = False
   SSTab1.TabVisible(3) = False
   SSTab1.TabVisible(4) = False 'Add By Sindy 2019/10/31
   'Added by Lydia 2019/01/03 修正另外開變更頁數的頁籤
   SSTab1.TabVisible(5) = False
   
'敏莉:
'有關工程師產生誤譯訂正申請書時，若有"205申復","204修正","203主動修正"尚未發文則:
'1. 增加彈訊息如下:
'若此道誤譯訂正不請款，則請和修正分2個申請書填送，分開送件。
'2. 請開放增、減項次欄位的填寫
'若有因增加項次需繳納規費產生，則申請書需增加如附檔黃色處的說明(和申復一樣)，繳費金額= 2000+本次應加收規費，且將增、減項次回寫到尚未發文的"申復""修正""主動修正"再審"進度檔中。
   'Add By Sindy 2022/3/4
   m_433Fee = "N"
   If cp(10) = "433" And _
      (PUB_ChkCPExist(cp, "203", 1) Or PUB_ChkCPExist(cp, "204", 1) Or PUB_ChkCPExist(cp, "205", 1)) Then
      If MsgBox("若此道誤譯訂正不請款，則請和" & _
         IIf(PUB_ChkCPExist(cp, "203", 1), "主動修正", IIf(PUB_ChkCPExist(cp, "204", 1), "修正", "申復")) & _
         "分2個申請書填送，分開送件。" & vbCrLf & vbCrLf & _
         "「不請款」請按「否」。", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
         m_433Fee = "Y"
      End If
   End If
   '2022/3/4 END
   
   m_strReExamCP27 = "" 'Added by Morgan 2013/1/10
   m_bolFixNewFee = False 'Added by Morgan 2013/1/10
   'Modified by Morgan 2013/11/6 +235核對中說格式
'   416.實體審查
'   201.新案翻譯
'   209.檢視中說
'   235.核對中說格式
'   210.製作中說
'   307.分割
'   203.主動修正
'   204.修正
'   205.申復
'   206.補充說明
   'Modify By Sindy 2022/3/4 + m_433Fee = "Y"
   If cp(10) = "416" Or cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Or _
      cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206" Or m_433Fee = "Y" Then
      m_Div416OfficialFee = 0 '2010/12/8 add by sonia
      
      'Added by Morgan 2013/1/10 判斷新舊法的費用計算
      If (cp(10) = "210" Or cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206") Then
         m_strReExamCP27 = PUB_GetReExamDate(cp)
         If m_strReExamCP27 > "20130000" Then
            m_bolFixNewFee = True '新法
         End If
      End If
      'end 2013/1/10
      
      If m_strReExamCP27 = "" Then 'Added by Morgan 2013/1/10
         m_bol99Case = Chk99NewCase(cp(1), cp(2), cp(3), cp(4))
         If m_bol99Case Then
            '實審
            If cp(10) = "416" Then
               '新案翻譯已發文
               'Modify by Morgan 2010/4/28 +307
               'Modified by Morgan 2013/11/6 +235核對中說格式
               'Modify By Sindy 2018/4/24 實審+主動修正時,才會是工程師出申請書
               If PUB_ChkCPExist(cp, "201", 2) Or PUB_ChkCPExist(cp, "209", 2) Or PUB_ChkCPExist(cp, "235", 2) Or PUB_ChkCPExist(cp, "210", 2) Or PUB_ChkCPExist(cp, "307", 2) Then
                  m_bolChkFee = True
                  txtCP84.Enabled = False
                  'Add By Sindy 2023/3/7
                  SSTab1.TabVisible(5) = True '顯示變更頁數的頁籤
                  '2023/3/7 END
                  '2018/4/24 END
               End If
'            '新案翻譯,檢視中說,製作中說
'            'Modified by Morgan 2013/11/6 +235核對中說格式
'            '中說
'            ElseIf cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
'               '實審已發文
'               If PUB_ChkCPExist(cp, "416", 2) Then
'                  m_bolChkFee = True
'                  m_bolChkPageItem = True
'                  txtCP84.Enabled = False
'                  lblAddItem.Caption = "總項數:"
'               End If
            '修正
            'Modify By Sindy 2022/3/4 + m_433Fee = "Y"
            ElseIf cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206" Or m_433Fee = "Y" Then
               'Modify By Sindy 2018/4/24
               m_bolChkFee = True
               txtCP84.Enabled = False
               'Add By Sindy 2023/3/7
               SSTab1.TabVisible(5) = True '顯示變更頁數的頁籤
               '2023/3/7 END
'               '主動修正+中說未收文未發文
'               If (PUB_ChkCPExist(cp, "201") = False And _
'                   PUB_ChkCPExist(cp, "209") = False And _
'                   PUB_ChkCPExist(cp, "235") = False And _
'                   PUB_ChkCPExist(cp, "210") = False) Or _
'                  PUB_ChkCPExist(cp, "201", 1) Or PUB_ChkCPExist(cp, "209", 1) Or _
'                  PUB_ChkCPExist(cp, "235", 1) Or PUB_ChkCPExist(cp, "210", 1) Then
                  '實審已發文
                  'MODIFY BY SONIA 2014/6/20 加入435續行母案再審 FCP-048155
                  If PUB_ChkCPExist(cp, "416", 2) Or PUB_ChkCPExist(cp, "435", 2) Then
                     'Modify By Sindy 2018/5/8
                     m_WriteNote = "N"
                     If cp(10) = "203" And _
                        (PUB_ChkCPExist(cp, "201", 1) Or PUB_ChkCPExist(cp, "209", 1) Or _
                         PUB_ChkCPExist(cp, "235", 1) Or PUB_ChkCPExist(cp, "210", 1)) Then
                        If MsgBox("是否一併送中說？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                           'm_bolChkPageItem = True 'Add By Sindy 2018/6/20
                           m_WriteNote = "Y"
                        End If
                     End If
                     '2018/5/8 END
                  '實審未收文或未發文
                  Else
                     'Modify By Sindy 2022/3/4
                     If m_433Fee = "Y" Then
                        '誤譯訂正有規費
                     Else
                        m_bolChkFee = False
                        txtCP84 = 0 '繳費金額固定為0
                     End If
                  End If
'               End If
               '2018/4/24 END
            End If
         
         '2010/12/8 ADD BY SONIA 分割案之實審發文,若母案有已收未取消的再審程序且申請日在2010/1/1以前者,分割案實審規費應為8000元
         '2011/3/24 MODIFY BY SONIA FCP-034512申復發文誤帶規費8000
         'ElseIf PUB_ChkCPExist(cp, "307") Then
         ElseIf cp(10) = "416" And PUB_ChkCPExist(cp, "307") Then
            txtCP84 = "8000"
            'Modify by Morgan 2011/7/26 會有超頁費 Ex.FCP-044051
            'txtCP84.Enabled = False
            MsgBox "本案請依舊法規則計算規費！"
            'end 2011/7/26
            cp(17) = txtCP84      '同時改收文規費
            m_Div416OfficialFee = txtCP84
         '2010/12/8 END
         End If
         
      'Added by Morgan 2013/1/10
      '有再審102年後發文,修正要收超項費
      ElseIf m_bolFixNewFee = True Then
         m_bolChkFee = True
         txtCP84.Enabled = False
         'Add By Sindy 2023/3/7
         SSTab1.TabVisible(5) = True '顯示變更頁數的頁籤
         '2023/3/7 END
      End If
      'end 2013/1/10
      
   'Added by Morgan 2013/1/9 102/1/1 起再審也要算超頁超項費(若有延期則以延期發文日判斷)
   'Modified by Morgan 2013/6/10 只有發明的再審
   'ElseIf cp(10) = "107" And m_bol107NewFee = True Then
   'Modified by Morgan 2013/10/18 +435
   'Modify By Sindy 2019/4/1 + 431高速審查
   ElseIf pa(8) = "1" And (cp(10) = "107" Or cp(10) = "435" Or cp(10) = "431") And m_bol107NewFee = True Then
      'Modify By Sindy 2018/6/11 ex:FCP-050420 再審申請+修正:AA7021170
      m_bolChkFee = True
      txtCP84.Enabled = False
      'Add By Sindy 2023/3/7
      SSTab1.TabVisible(5) = True '顯示變更頁數的頁籤
      '2023/3/7 END
      '2018/6/11 END
   'Added by Morgan 2013/1/3
   '申請技術報告要輸項數以便計算規費
   ElseIf cp(10) = "421" Or cp(10) = "807" Then
      lblAddItem.Caption = "項數:"
      txtAddItem.Enabled = True
      txtAddItem.BackColor = vbWhite
      txtCP84.Enabled = False
      'Add By Sindy 2024/8/21
      If cp(10) = "421" Then
         m_bolChkFee = True
         SSTab1.TabCaption(6) = "申請技術報告"
         Frame1.Visible = False
         Frame421.Visible = True
         Frame421.Left = 330
      End If
      '2024/8/21 END
   'end 2013/1/3
   End If
   '*************************************************************************
   
   Select Case cp(10)
      '307分割,416實體審查
      Case "307", "416"
         'Add By Sindy 2019/6/3 進入產生分割申請書介面詢問是否電子送件
         If cp(10) = "307" Then
            'Add By Sindy 2020/3/10
            If pa(8) = "3" Then '設計案
               FramePA158.Visible = True
            Else
               FramePA158.Visible = False
            End If
            If FramePA158.Visible = True Then
'               Call PUB_AddCaseAttributeCombo(Combo3, pa(8)) '專利案件屬性選單
'               If pa(158) = "" Then
'                  '預設母案的案件屬性
'                  Call PUB_DivAppDate(pa(1), pa(2), pa(3), pa(4), , , m_DivAppPA158) '取得母案案件屬性
'                  Combo3 = m_DivAppPA158 + "." + PUB_GetCaseAttributeName(m_DivAppPA158, pa(8))
'                  Combo3.Tag = ""
'               Else
'                  Combo3 = pa(158) + "." + PUB_GetCaseAttributeName(pa(158), pa(8))
'                  Combo3.Tag = Combo3.Text
'               End If
            End If
            '2020/3/10 End
            
            If MsgBox("分割案是否以電子送件？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
               Check4.Value = 1
            End If
         End If
         '2019/6/3 END
         SSTab1.TabVisible(0) = True
         Frame307.Enabled = False
         Frame416.Enabled = False
         
         If cp(10) = "416" Then
            SSTab1.TabVisible(1) = True
            Frame416.Enabled = True
            chkAtt(1).Visible = True '申復書開放可勾選
         
         ElseIf cp(10) = "307" Then
            'Modify By Sindy 2018/5/8
            m_bolChkFee = True
            'm_bolChkPageItem = True 'Modify By Sindy 2018/8/7 Mark:FCP-59380
            'lblAddItem.Caption = "總項數:"
            '2018/5/8 END
            Frame307.Enabled = True
            Check2.Enabled = False
            Check9(0).Enabled = False
            Check1.Caption = IIf(pa(8) = "2", Replace(Check1.Caption, "新型", "發明"), Check1.Caption)
            If pa(8) = "1" Then
               Check2.Enabled = True
               Check9(0).Enabled = True
               '是否一併提實審
               If PUB_ChkCPExist(pa, "416") = True Then
                  Check2.Value = 1
                  'm_bolChkPageItem = True 'Modify By Sindy 2018/8/7 Mark:FCP-59380
               End If
            End If
            If Check2.Enabled = False Then Check2.BackColor = Me.BackColor
            
            If pa(140) <> "" Then
               txtFavDate = TransDate(pa(140), 1)
            End If
            Call PUB_DivAppDate(pa(1), pa(2), pa(3), pa(4), , strDivPA11) '取得母案申請案號
         End If
         
      'Added by Morgan 2022/5/11
      Case "435"
         SSTab1.TabCaption(0) = "續行母案再審"
         SSTab1.TabVisible(0) = True
         Frame307.Enabled = False
         Frame416.Enabled = False
         SSTab1.TabVisible(1) = True
         chkAtt(1).Visible = True
         Check9(1).Visible = True
         Check9(1).Value = vbChecked
         Check9(1).Enabled = False
      'end 2022/5/11
         
      '204修正,203主動修正,433誤譯訂正,205申復(PPH修正)
      Case "204", "203", "433", "205"
         SSTab1.TabVisible(1) = True
         'Modify By Sindy 2019/3/4 有續行母案再審未發文時才顯示
         If PUB_ChkCPExist(cp, "435", 1) Then
            Check9(1).Visible = True
         End If
         '2019/3/4 END
         'Added by Lydia 2019/01/14 主動修正在中說發文後,工程師才可以修改頁數(ex.FCP-59962 因為要先請款,所以檢視中說209先設頁數64與實際頁數67不符)
         'Remove by Lydia 2019/01/17 不限制在中說發文後，工程師才可以修改頁數，始終保持最新版頁數；若程序在出申請書時，視情況做人工變更by敏莉
         'strExc(0) = "select cp09,cp10, cp158 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('201','209','210','235') and cp159=0 order by cp158 "
         'intI = 1
         'strExc(1) = "0"
         'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         'If intI = 1 Then
         '   strExc(1) = "" & RsTemp.Fields("cp158")
         'End If
         'If Val(strExc(1)) > 0 And strExc(1) < strSrvDate(1) Then
         'end 2019/01/14
            'Modify By Sindy 2023/3/14 mark
'            'Added by Lydia 2019/01/03 修正另外開變更頁數的頁籤
'            SSTab1.TabVisible(5) = True
'            FramePA6468.Visible = True
'            lblPS1.Visible = True
'            'end 2019/01/03
         'End If 'end 2019/01/14
         'end 2019/01/17
         
         '檢查是否為PPH修正
         If cp(43) <> "" Then
            strExc(0) = "select cp09,cp10 from caseprogress where cp09='" & cp(43) & "' and cp10 in ('431') and cp159=0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_CP43CP10 = RsTemp.Fields("cp10")
               chk1Tab1(19).Visible = True '申請專利範圍對應表開放可勾選
            End If
         End If
         If cp(10) = "204" Or cp(10) = "203" Or cp(10) = "205" Then
            chkAtt(1).Visible = True '申復書開放可勾選
            If cp(10) = "205" Then '申復
               chkAtt(1).Value = 1
            End If
            Frame433.Enabled = False
         ElseIf cp(10) = "433" Then
            Check433.Visible = True
            chk1Tab2(31).Visible = True
         End If
         
      '431高速審查
      Case "431"
         SSTab1.TabVisible(1) = True
         'chk1Tab1(19).Visible = True '申請專利範圍對應表開放可勾選
         SSTab1.TabVisible(2) = True
         SSTab1.TabVisible(5) = True 'Add By Sindy 2019/8/29 = 修正另外開變更頁數的頁籤
         
      '107再審申請,422加速審查
      Case "107", "422"
         SSTab1.TabVisible(3) = True
         Frame107.Enabled = False
         Frame422.Enabled = False
         
         If cp(10) = "107" Then
            SSTab1.TabVisible(1) = True
            Frame107.Enabled = True
         
         ElseIf cp(10) = "422" Then
            Frame422.Enabled = True
         End If
      
      'Add By Sindy 2019/10/31 402更正
      Case "402"
         SSTab1.TabVisible(4) = True
   End Select
   
   '控管項目是否可勾選
   If SSTab1.TabVisible(0) = True Then
      For Each chk In chk0Tab
         If pa(8) = "1" Then
            If Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
               chk.Enabled = False
            End If
         ElseIf pa(8) = "2" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "3" Then
               chk.Enabled = False
            End If
         ElseIf pa(8) = "3" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Then
               chk.Enabled = False
            End If
         End If
      Next
   End If
   If SSTab1.TabVisible(1) = True Then
      For Each chk In chk1Tab1
         If pa(8) = "1" Then
            If Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
               chk.Enabled = False
            End If
         ElseIf pa(8) = "2" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "3" Then
               chk.Enabled = False
            End If
         ElseIf pa(8) = "3" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Then
               chk.Enabled = False
            End If
         End If
      Next
      For Each chk In chk1Tab2
         If pa(8) = "1" Then
            If Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
               chk.Enabled = False
            End If
         ElseIf pa(8) = "2" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "3" Then
               chk.Enabled = False
            End If
         ElseIf pa(8) = "3" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Then
               chk.Enabled = False
            End If
         End If
      Next
   End If
   'Add By Sindy 2019/10/31
   If SSTab1.TabVisible(4) = True Then
      For Each chk In chk4Tab1
         If pa(8) = "1" Then
            If Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
               chk.Enabled = False
            End If
         ElseIf pa(8) = "2" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "3" Then
               chk.Enabled = False
            End If
         ElseIf pa(8) = "3" Then
            If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Then
               chk.Enabled = False
            End If
         End If
      Next
   End If
   '2019/10/31 END
   If Frame307.Enabled = False And SSTab1.TabVisible(0) = True Then
      Check2.BackColor = Me.BackColor
      Check9(0).Enabled = False
      txtFavDate.BackColor = Me.BackColor
      cboFavReason.BackColor = Me.BackColor
   End If
   If Frame433.Enabled = False And SSTab1.TabVisible(1) = True Then
      For Each chk In chk1Tab2
         chk.Enabled = False
      Next
   End If
   If Frame107.Enabled = False And SSTab1.TabVisible(3) = True Then
      For Each oText In txtDocCh2
         oText.Enabled = False
         oText.BackColor = Me.BackColor
      Next
      chk3Tab2(0).BackColor = Me.BackColor
      chk3Tab2(0).Value = 0
      chk3Tab2(0).Enabled = False
   End If
   If SSTab1.TabVisible(3) = True Then
      If Frame422.Enabled = False Then
         For Each chk In chk3Tab3
            chk.Enabled = False
         Next
         For Each opt In Opt2Tab3
            opt.Enabled = False
         Next
      'Add By Sindy 2025/9/25
      Else
         intI = 0
         For Each chk In chk3Tab3
            If pa(8) = "3" Then '設計
               strExc(10) = "": strExc(9) = ""
               If intI = 0 Then strExc(10) = "委任書": strExc(9) = ".POA.pdf"
               If intI = 1 Then strExc(10) = "第三人商業實施證明文件": strExc(9) = ".ATT.pdf"
               If intI = 2 Then strExc(10) = "獲獎證明文件": strExc(9) = ".ATT.pdf"
               If intI = 3 Then strExc(10) = "外國公司設立日期證明文件": strExc(9) = ".ATT.pdf"
               If intI = 4 Then strExc(10) = "外國公司設立日期證明文件中譯本": strExc(9) = ".ATT.pdf"
               If intI = 5 Then strExc(10) = "外國公司設立日期證明文件切結書": strExc(9) = ".ATT.pdf"
               chk.Caption = strExc(10)
               chk.Tag = strExc(9)
               If strExc(10) = "" Then chk.Visible = False
               intI = intI + 1
            End If
         Next
         intI = 0
         For Each opt In Opt2Tab3
            If pa(8) = "3" Then '設計
               strExc(10) = ""
               If intI = 0 Then strExc(10) = "第三人商業實施"
               If intI = 1 Then strExc(10) = "曾獲得國內外著名設計獎項"
               If intI = 2 Then strExc(10) = "新創企業之設計專利申請案"
               opt.Caption = strExc(10)
               If strExc(10) = "" Then opt.Visible = False
               intI = intI + 1
            End If
         Next
      End If
      '2018/4/11 END
   End If
   
   If SSTab1.TabVisible(0) = True Then
      SSTab1.Tab = 0
   ElseIf SSTab1.TabVisible(2) = True Then
      SSTab1.Tab = 2
   ElseIf SSTab1.TabVisible(3) = True Then
      SSTab1.Tab = 3
   'Add By Sindy 2019/10/31
   ElseIf SSTab1.TabVisible(4) = True Then
      SSTab1.Tab = 4
   '2019/10/31 END
   ElseIf SSTab1.TabVisible(1) = True Then
      SSTab1.Tab = 1
   End If
   
   'Add By Sindy 2018/7/30 檢查再審查是否有申請過延期
   bolHad404 = False
   If cp(10) = "107" Then
      strExc(0) = "select cp09,cp118,cp10,cp158,cp05 From caseprogress" & _
                  " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                  " and cp10='404' and cp43='" & cp(9) & "'" & _
                  " Union" & _
                  " select cp09,cp118,cp10,cp158,cp05 From caseprogress" & _
                  " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                  " and cp10='404' and cp43 in(select cp43 from caseprogress where cp09='" & cp(9) & "')" & _
                  " order by cp158 desc,cp05 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            'Add By Sindy 2019/2/12
            '已延過期且為紙本送件,則走未延過期的申請書流程
            If "" & RsTemp.Fields("cp118") <> "" Then
               bolHad404 = True
               strHad404CP09 = RsTemp.Fields(0) 'Add By Sindy 2019/1/31
               SSTab1.Tab = 1
            End If
            '2019/2/12 END
         End If
      End If
   End If
   
   'Add By Sindy 2018/7/26 設計案
   If pa(8) <> "1" Then
      Check2.Enabled = False '是否一併提實審
      Check3.Enabled = False '附英文摘要
      chk0Tab(8).Enabled = False '序列表
      chk0Tab(12).Enabled = False '國內生物材料寄存證明文件
      chk0Tab(13).Enabled = False '國外生物材料寄存證明文件
      chk0Tab(14).Enabled = False '生物材料為通常知識者易於獲得證明文件
      '設計
      If pa(8) = "3" Then
         '摘要頁數
         'txtDocCh(0).Enabled = False
         txtDocCh2(0).Enabled = False
         '申請專利範圍頁數
         'txtDocCh(2).Enabled = False
         txtDocCh2(2).Enabled = False
         '申請專利範圍項數
         'txtDocCh(5).Enabled = False
         txtDocCh2(5).Enabled = False
         'Add By Sindy 2019/8/7 Ex:FCP-061566
         txtDocCh2(7).Enabled = False
         '2019/8/7 END
      End If
   End If
   '2018/7/26 END
   
   'Modify By Sindy 2018/7/27
   '直接帶入目前該筆資料
   'Add By Sindy 2018/8/28 是否要輸頁數與項數 ex:FCP-52188:再審後申復(txtCP135.Enabled = False)
   'Add By Sindy 2018/8/28 + Or m_WriteNote = "Y" ex:FCP-58875:主動修正一併送中說
   'txtCP135 = m_allPage '原總頁數
   'If txtCP135.Enabled = True Or m_WriteNote = "Y" Then txtCP135 = m_allPage
   
'   If m_WriteNote = "Y" Then txtCP135 = m_allPage 'Add By Sindy 2023/3/16
   
   '2018/8/28 END
   'Modify By Sindy 2019/1/16
   'txtItem = m_allItem
   If m_allItem > 0 Then '原請求項數
      txtItem = m_allItem
   Else
      txtItem = pa(172)
   End If
   'Add By Sindy 2023/3/10
'   If m_allPage > 0 Then '原總頁數
'      txtPage = m_allPage
'   Else
      txtPage = txtCP135
'   End If
   '2023/3/10 END
   If m_WriteNote = "Y" Then
      If Val(m_allItem) > 0 And Val(pa(172)) > 0 And Val(m_allItem) > Val(pa(172)) Then
'         txtItem = pa(172)
         txtAddItem = m_allItem - pa(172)
      End If
      'Add By Sindy 2023/3/10
      If Val(m_allPage) > 0 And Val(txtCP135) > 0 And Val(m_allPage) > Val(txtCP135) Then
'         txtPage = txtCP135
         txtAddPage = m_allPage - txtCP135
      End If
      '2023/3/10 END
   End If
   '2019/1/16 END
   'Modify By Sindy 2023/3/13 + Or Val(cp(167)) > 0 Or Val(cp(168)) > 0
   If Val(cp(135)) > 0 Or Val(cp(136)) > 0 Or Val(cp(137)) > 0 Or Val(cp(138)) > 0 Or Val(cp(167)) > 0 Or Val(cp(168)) > 0 Then
      'If txtCP135.Enabled = True Then txtCP135 = cp(135) 'Modify By Sindy 2018/8/28 ex:FCP-58875
      'txtCP136 = cp(136)
      If SSTab1.TabVisible(0) = False And SSTab1.TabVisible(3) = False Then
         txtAddItem = cp(136)
         txtAddPage = cp(135) 'Add By Sindy 2023/3/10
      End If
      txtCP137 = cp(137) '刪除未審項數
      txtCP138 = cp(138) '刪除已審項數
      txtCP167 = cp(167) '刪除未審頁數'Add By Sindy 2023/3/10
      txtCP168 = cp(168) '刪除已審頁數'Add By Sindy 2023/3/10
   End If
   
   'Modify By Sindy 2023/3/15 Mark
'   'Add By Sindy 2019/1/15 ex:FCP-51198
'   If cp(10) = "416" And SSTab1.TabVisible(0) = True And txtDocCh(5).Enabled = True Then
'      'txtDocCh(5).Text = txtCP136
'      'Add By Sindy 2019/5/2 ex:FCP-50598
'      'If Val(txtCP136) = 0 And Val(m_allItem) > 0 Then
'         txtDocCh(5).Text = Val(m_allItem)
'      'End If
'      '2019/5/2 END
'   ElseIf cp(10) = "107" And SSTab1.TabVisible(3) = True And txtDocCh2(5).Enabled = True Then
'      'txtDocCh2(5).Text = txtCP136
'      'Add By Sindy 2019/5/2 ex:FCP-50598
'      'If Val(txtCP136) = 0 And Val(m_allItem) > 0 Then
'         txtDocCh2(5).Text = Val(m_allItem)
'      'End If
'      '2019/5/2 END
'   End If
'   '2019/1/15 END
   
   'Add By Sindy 2019/8/30 鎖住不需要輸入,僅填中文本資訊即可
   If SSTab1.TabVisible(0) = True Or _
      (SSTab1.TabVisible(3) = True And cp(10) = "107") Then
      txtAddItem.BackColor = Me.BackColor
      txtCP137.BackColor = Me.BackColor
      txtCP138.BackColor = Me.BackColor
      txtCount.BackColor = Me.BackColor
      txtAddItemFee.BackColor = Me.BackColor
      txtDecreaseItemFee.BackColor = Me.BackColor
      FrameFee.Enabled = False
      'Label35.Caption = "總頁數" 'Add By Sindy 2023/3/14
      SSTab1.TabVisible(5) = False '不顯示變更頁數 Add By Sindy 2023/3/7
      '僅填中文本資訊即可
      txtItem.Text = ""
      txtPage.Text = ""
      txtAddItem.Text = ""
      txtCP137.Text = ""
      txtCP138.Text = ""
      txtCount.Text = ""
      txtAddItemFee.Text = ""
      txtDecreaseItemFee.Text = ""
   'Add By Sindy 2023/3/14
   Else
      If SSTab1.TabVisible(5) = True Then
         m_bolChkItem = True
         FrameFee.Enabled = True
         lblAddItem.Caption = "增加項數:"
         txtAddItem.Enabled = True
         txtAddItem.BackColor = vbWhite
         txtCP137.Enabled = True
         txtCP137.BackColor = vbWhite
         '已輸審查意見通知
         'Modify By Sindy 2022/3/10 + 1227最後通知(審查意見最終通知書)
         If PUB_ChkCPExist(cp, "1202") = True Or PUB_ChkCPExist(cp, "1227") = True Then
            txtCP138.Enabled = True
            txtCP138.BackColor = vbWhite
            'Add By Sindy 2023/3/7
            For Each oText In txtDocCp168
               oText.Enabled = True
               oText.BackColor = vbWhite
            Next
            '2023/3/7 END
         End If
         txtCount.Enabled = True
         txtCount.BackColor = vbWhite
         txtAddItemFee.Enabled = True
         txtAddItemFee.BackColor = vbWhite
         txtDecreaseItemFee.Enabled = True
         txtDecreaseItemFee.BackColor = vbWhite
      End If
      '2023/3/14 END
   End If
   '2019/8/30 END
   lblPS1.Visible = SSTab1.TabVisible(5) 'Add By Sindy 2023/3/14
   
   'Add By Sindy 2020/6/29 FCP-59506,FCP-61132 項數沒計算出規費
   If SSTab1.TabVisible(0) = True Then
      If Val(txtDocCh(0).Text) > 0 Or Val(txtDocCh(1).Text) > 0 Or Val(txtDocCh(2).Text) > 0 Or Val(txtDocCh(3).Text) > 0 Then
         Call txtDocCh_Validate(3, False)
      End If
      If Val(txtDocCh(5).Text) > 0 Then
         Call txtDocCh_Validate(5, False)
      End If
   End If
   If SSTab1.TabVisible(3) = True Then
      If Val(txtDocCh2(0).Text) > 0 Or Val(txtDocCh2(1).Text) > 0 Or Val(txtDocCh2(2).Text) > 0 Or Val(txtDocCh2(3).Text) > 0 Then
         Call txtDocCh2_Validate(3, False)
      End If
      If Val(txtDocCh2(5).Text) > 0 Then
         Call txtDocCh2_Validate(5, False)
      End If
   End If
   '2020/6/29 END
   
   'If Trim(pa(22)) = "" Then '申請權
      chkAtt3(1).Caption = "變更申請人之地址"
      chkAtt3(3).Caption = "變更申請人之代表人"
      chkAtt3(5).Caption = "變更申請人之國籍"
   'Else '專利權
   If cp(10) = 更正 Then
      chkAtt3(1).Caption = "變更專利權人之地址"
      chkAtt3(3).Caption = "變更專利權人之代表人"
      chkAtt3(5).Caption = "變更專利權人之國籍"
   End If
   'Add By Sindy 2023/2/20 增加檢查是否有變更資料
   strExc(0) = "select cp09 From caseprogress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10='401' and cp27||cp57 is null" & _
               " order by cp66 desc,cp67 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strCE01 = RsTemp.Fields("cp09")
      '預設,同時辦理事項
      If PUB_GetChangeEvent(strCE01, 1) = True Then
         chkAtt3(1).Value = 1: chkAtt3(1).Tag = strCE01
      End If
      If PUB_GetChangeEvent(strCE01, 3) = True Then
         chkAtt3(3).Value = 1: chkAtt3(3).Tag = strCE01
      End If
      If PUB_GetChangeEvent(strCE01, 5) = True Then
         chkAtt3(5).Value = 1: chkAtt3(5).Tag = strCE01
      End If
   End If
   '2023/2/20 END
End Sub

'Add By Sindy 2023/3/13 計算頁數合計
Private Sub CountPage()
   If FrameFee.Enabled = True Then
      '合計:
      '增加頁數:
      If Val(txtDocAdd(0)) + Val(txtDocAdd(1)) + Val(txtDocAdd(3)) + Val(txtDocAdd(4)) > 0 Then
         txtAddPage = Val(txtDocAdd(0)) + Val(txtDocAdd(1)) + Val(txtDocAdd(3)) + Val(txtDocAdd(4))
      Else
         txtAddPage = ""
      End If
      '刪除未審頁數:
      If Val(txtDocCp167(0)) + Val(txtDocCp167(1)) + Val(txtDocCp167(3)) + Val(txtDocCp167(4)) > 0 Then
         txtCP167 = Val(txtDocCp167(0)) + Val(txtDocCp167(1)) + Val(txtDocCp167(3)) + Val(txtDocCp167(4))
      Else
         txtCP167 = ""
      End If
      '刪除已審頁數:
      If Val(txtDocCp168(0)) + Val(txtDocCp168(1)) + Val(txtDocCp168(3)) + Val(txtDocCp168(4)) > 0 Then
         txtCP168 = Val(txtDocCp168(0)) + Val(txtDocCp168(1)) + Val(txtDocCp168(3)) + Val(txtDocCp168(4))
      Else
         txtCP168 = ""
      End If
      '修正後中文本頁數:
      If Val(txtAddPage) + Val(txtCP167) + Val(txtCP168) > 0 Then
         '摘要頁數
         txtDocCh4(0) = Val(txtDocCh3(0)) + Val(txtDocAdd(0)) - Val(txtDocCp167(0)) - Val(txtDocCp168(0))
         '說明書頁數
         txtDocCh4(1) = Val(txtDocCh3(1)) + Val(txtDocAdd(1)) - Val(txtDocCp167(1)) - Val(txtDocCp168(1))
         '申請專利範圍頁數
         txtDocCh4(3) = Val(txtDocCh3(3)) + Val(txtDocAdd(3)) - Val(txtDocCp167(3)) - Val(txtDocCp168(3))
         '圖式頁數
         txtDocCh4(4) = Val(txtDocCh3(4)) + Val(txtDocAdd(4)) - Val(txtDocCp167(4)) - Val(txtDocCp168(4))
         '修正後總頁數
         txtPageCount = Val(txtPage) + Val(txtAddPage) - Val(txtCP167) - Val(txtCP168)
      Else
         txtPageCount = ""
      End If
   End If
End Sub

'Add By Sindy 2018/6/13 有關規費計算或相關控制
Private Sub Call_PUB_SetOfficialFee_P()
Dim bolPageFee As Boolean 'Add By Sindy 2019/3/5

   If Me.Visible = False Then Exit Sub
   
   '分割案/再審:Check9(0).本案續行再審查
   If cp(10) = "307" And Check9(0).Visible = True Then
      'Check2.一併提實審
      If Check9(0).Value = 1 Or Check2.Value = 1 Then
         m_bolChkPageItem = True
         m_bolChkFee = True
      Else
         m_bolChkPageItem = False
      End If
   'Add By Sindy 2023/4/24 出實審時,若中說已發文,要算頁項數 ex:FCP-065189
   ElseIf cp(10) = "416" And m_bolChkFee = True Then
      m_bolChkPageItem = True
      '2023/4/24 END
   'Add By Sindy 2023/5/3 再審申請時,一律要算頁項數 ex:FCP-060131
   ElseIf cp(10) = "107" And m_bolChkFee = True Then
      m_bolChkPageItem = True
      '2023/5/3 END
   Else
      If SSTab1.TabVisible(0) = False And Check9(1).Visible = True Then
         'Check9(1).本案續行再審查
         If Check9(1).Value = 1 Then
            m_bolChkPageItem = True
            m_bolChkFee = True
         ElseIf Check9(1).Value = 0 Then
            m_bolChkPageItem = False
         End If
      End If
   End If
   
'   'Modify By Sindy 2018/8/7 Mark:FCP-59380 分割:一併提實審
'   If Check2.Visible = True And Check2.Value = 1 And cp(10) = "307" Then
'      m_bolChkPageItem = True
'   ElseIf Check2.Visible = True And Check2.Value = 0 And cp(10) = "307" Then
'      m_bolChkPageItem = False
'   End If
'   '2018/8/7 END
   
   'Add By Sindy 2018/7/27
   'Add By Sindy 2019/5/2 剔除實審和再審 ex:FCP-50598
   'If cp(10) <> "416" And cp(10) <> "107" And cp(10) <> "435" Then
      If FrameFee.Enabled = True Then
         Call CountPage 'Add By Sindy 2023/3/13
         
         'Add By Sindy 2018/8/28 ex:FCP-58875:主動修正一併送中說
         'Modify By Sindy 2019/3/4 ex:FCP-60443:主動修正,續行母案再審
         'Modify By Sindy 2019/4/16 ex:FCP-60679 +  - Val(txtCP137) - Val(txtCP138)
         'Modify By Sindy 2019/4/23 ex:FCP-60518:主動修正,原36,新增17,刪36,本次應退19項(19*800=15200元)
         'Modify By Sindy 2019/4/29 Mark
         '    ex:FCP-60518:主動修正,原36,新增17,刪36,本次應退19項(19*800=15200元)
         '    ex:FCP-60709:主動修正,原8,新增1,規費0元
   '      If Val(txtCP137) > 0 Or Val(txtCP138) > 0 Then
   '         txtCP136 = Val(txtAddItem)
   '      Else
            'Modify By Sindy 2024/9/9 + Or cp(10) = 421
            If m_WriteNote = "Y" Or Check9(1).Value = 1 Or cp(10) = 421 Then
               txtCP136 = Val(txtItem) + Val(txtAddItem) ' - Val(txtCP137) - Val(txtCP138)
               txtCP135 = Val(txtPage) + Val(txtAddPage) 'Add By Sindy 2023/3/10
            Else
            '2018/8/28 END
               txtCP136 = Val(txtAddItem) ' - Val(txtCP137) - Val(txtCP138)
               txtCP135 = Val(txtAddPage) 'Add By Sindy 2023/3/10
            End If
   '      End If
         'Add By Sindy 2019/1/15 ex:FCP-60147
         If SSTab1.TabVisible(0) = True And txtDocCh(5).Enabled = True Then
            txtDocCh(5).Text = txtCP136
         ElseIf SSTab1.TabVisible(3) = True And txtDocCh2(5).Enabled = True Then
            txtDocCh2(5).Text = txtCP136
         End If
         
         'Modify By Sindy 2018/6/13 代表有異動值
         'Modify By Sindy 2018/6/21 + Left(lblAddItem.Caption, 4) = "增加項數"
         'Modify By Sindy 2018/7/30 ex:FCP-051538
   '      If (Val(txtAddItem) + Val(txtCP137) + Val(txtCP138) > 0) Or _
   '         Left(lblAddItem.Caption, 4) = "增加項數" Then
         If (Val(txtAddItem) + Val(txtCP137) + Val(txtCP138)) > 0 Then
            txtCount = Val(txtItem) + Val(txtAddItem) - Val(txtCP137) - Val(txtCP138)
            If (Val(txtAddItem) + Val(txtCP137)) > 0 Then 'Modify By Sindy 2019/3/5 有輸入增減項,要計算項數規費
               m_bolChkPageItem = True 'Add By Sindy 2018/8/28 ex:FCP-52188:再審後申復
            End If
            'Add By Sindy 2019/1/15 ex:FCP-60147
            If SSTab1.TabVisible(0) = True And txtDocCh(5).Enabled = True Then
               txtDocCh(5).Text = txtCount
            ElseIf SSTab1.TabVisible(3) = True And txtDocCh2(5).Enabled = True Then
               txtDocCh2(5).Text = txtCount
            End If
         Else
            txtCount = 0
         End If
         'Add By Sindy 2023/3/13
         If (Val(txtAddPage) + Val(txtCP167) + Val(txtCP168)) > 0 Then
            If (Val(txtAddPage) + Val(txtCP167)) > 0 Then
               m_bolChkPageItem = True
            End If
         Else
            txtPageCount = 0
         End If
         '2023/3/13 END
      End If
   'End If
   
   If m_bolChkFee Then
      'Modify By Sindy 2018/7/31 txtCP135_tmp:為了要代表頁/項數為空白,ex:FCP-057109申復
      'Modify By Sindy 2018/8/7 + FCP:59380分割 txtCP136 ==> IIf(m_bolChkPageItem = True, txtCP136, txtCP135_tmp)
      'Modify By Sindy 2019/2/1 + 本案續行再審查,一樣要加7000元 ex:FCP-60443 IIf(Check2.Value = 1 Or Check9.Value = 1, True, False)
      'Modify By Sindy 2019/3/5 實審和再審(一併提實審,本案續行再審查)才要計算頁數規費
      'IIf(m_bolChkPageItem = True, txtCP135, txtCP135_tmp) ==> IIf(bolPageFee = True, txtCP135, txtCP135_tmp)
      If m_bolChkPageItem = True Then
         bolPageFee = True
         If Not (cp(10) = "107" Or cp(10) = "416" Or _
               Check2.Value = 1 Or Check9(0).Value = 1 Or Check9(1).Value = 1) Then
            bolPageFee = False
         End If
      End If
      '2019/3/5 END
      
      'Modify By Sindy 2021/12/24 + m_WriteNote = "Y"  修正實審已發文，中說+修正申請書，補述內容  FCP-065711
'      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
'                                IIf(bolPageFee = True Or m_WriteNote = "Y", txtCP135, txtCP135_tmp), _
'                                IIf(m_bolChkPageItem = True Or m_WriteNote = "Y", txtCP136, txtCP135_tmp), txtCP137, txtCP84, txtAddItemFee, txtDecreaseItemFee, _
'                                m_lngOverPageFee, m_lngOverItemFee, IIf(SSTab1.TabVisible(0) = True And Check4.Value = 1, True, False), IIf(Check2.Value = 1 Or Check9(0).Value = 1 Or Check9(1).Value = 1, True, False))
      'Modify By Sindy 2023/4/7 + m_WriteNote
      'Modify By Sindy 2024/9/9 + Or cp(10) = 421
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                IIf(m_bolChkPageItem = True Or m_WriteNote = "Y" Or cp(10) = 421, txtCP135, txtCP135_tmp), _
                                IIf(m_bolChkPageItem = True Or m_WriteNote = "Y" Or cp(10) = 421, txtCP136, txtCP135_tmp), txtCP137, txtCP84, txtAddItemFee, txtDecreaseItemFee, _
                                m_lngOverPageFee, m_lngOverItemFee, IIf(SSTab1.TabVisible(0) = True And Check4.Value = 1, True, False), _
                                IIf(Check2.Value = 1 Or Check9(0).Value = 1 Or Check9(1).Value = 1, True, False), , _
                                , txtAddPageFee, txtDecreasePageFee, txtCP167, m_WriteNote)
   End If
   
   'Add By Sindy 2020/8/28 亭妙(FCP-061760):當【實體審查一併誤譯訂正時】請在實審申請書中原來的規費規則下，增加2000元（為誤譯訂正之規費）
   If cp(10) = "416" And Check416_2(0).Value = 1 Then
      txtCP84 = Val(txtCP84) + Val(GetFMPOfficialFee(pa(1), "433", pa(9)))
   End If
   '2020/8/28 END
   
   'Add By Sindy 2019/1/16 一併送中說,此處繳費為0,繳費金額出在中說申請書裡
   If m_WriteNote = "Y" Then '主動修正一併送中說
      txtCP84 = 0
   End If
   'Add By Sindy 2019/4/2
   If Val(txtCP84) < 0 Then txtCP84 = 0
   If FrameFee.Enabled = True Then
      If Val(txtAddItemFee) < 0 Then txtAddItemFee = 0
      If Val(txtAddPageFee) < 0 Then txtAddPageFee = 0 'Add By Sindy 2023/3/14
   Else
      txtAddItemFee.Text = ""
      txtAddPageFee.Text = ""
   End If
   '2019/4/2 END
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strFolder As String, strFileName As String
   Dim Cancel As Boolean
   Dim intFileCnt As Integer 'Add By Sindy 2019/10/31 申請書電子檔數量
   Dim m_Representative As String 'Add By Sindy 2023/2/22 代表人
   Dim strET03 As String 'Add By Sindy 2025/9/25
   
   bolIsSecond = False 'Add By Sindy 2023/3/17 非 第2個申請書
   Select Case Index
      Case 0
         'Add By Sindy 2023/8/11
         If cmdOK(0).Tag <> "僅產生申請書" Then
         '2023/8/11 END
            'Add By Sindy 2020/10/8
            If SSTab1.TabVisible(0) = True Then
               Call txtDocCh_Validate(3, False)
               Call txtDocCh_Validate(5, False) 'Add By Sindy 2021/7/1
            End If
            If SSTab1.TabVisible(3) = True Then
               Call txtDocCh2_Validate(3, False)
               Call txtDocCh2_Validate(5, False) 'Add By Sindy 2021/7/1
            End If
            '2020/10/8 END
            
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'Added by Lydia 2020/02/17 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
            If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
                MsgBox MsgText(1111), vbInformation
                If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                    Exit Sub
                End If
            End If
            'end 2020/02/17
            
            If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         End If
                  
         'm_CaseNo = pa(1) & IIf(Left(pa(2), 1) = "0", Mid(pa(2), 2), pa(2)) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "")
         m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
         
         'If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
            strFolder = PUB_Getdesktop
         Else
            strFolder = FCP電子送件檔案存放路徑
         End If
         strFolder = strFolder & "\" & m_CaseNo
         If Dir(strFolder, vbDirectory) = "" Then
            MkDir strFolder
         End If
         
         'Add By Sindy 2023/2/22 抓變更檔中代表人資料
         m_Representative = ""
         If chkAtt3(3).Value = 1 And chkAtt3(3).Tag <> "" Then
            Call PUB_GetChangeEvent(chkAtt3(3).Tag, 3, m_Representative)
         End If
         '2023/2/22 END
         
         '1.基本資料
         'Modify By Sindy 2018/6/27 分割:基本資料表要顯示發明人資料
         'Modify By Sindy 2023/2/22 + IIf(chkAtt3(1).Value = 1 And chkAtt3(1).Tag <> "", True, False) => True : 要抓客戶檔資料
         '                          + m_Representative
         'Add By Sindy 2023/8/11
         If cp(1) = "FG" Then
            '目前只有230提供情報是給空白的基本資料表
            NowPrint strReceiveNo, "01", "97", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9), strFileName)
         Else
         '2023/8/11 END
            StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, IIf(cp(10) = "307", True, False), _
               IIf(chkAtt3(1).Value = 1 And chkAtt3(1).Tag <> "", True, False), , m_Representative
            NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9), strFileName)
         End If
         
         '2.申請書
         Select Case cp(10)
            'Add By Sindy 2023/8/11
            Case "230" '提供情報
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "發明專利申請案第三方意見書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1
               
            'Add By Sindy 2022/4/13 + 239擇一申復
            Case "239" '擇一申復
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "一案兩請擇一申復申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1
               
            'Add By Sindy 2019/12/12
            Case "403" '更改
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "一般事項申復申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1
               
            'Add By Sindy 2019/10/31
            Case "402" '更正
               If StartLetter2("01", "25") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "25", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利更正申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1
            
            'Modified by Morgan 2022/5/12 +435 續行母案再審 也是出修正申請書
            Case "203", "204", "435" '修正
               'Modify By Sindy 2019/1/15 ex:FCP-59771
               'If m_CP43CP10 = "431" Then 'PPH修正
               '有431高速(PPH)審查之後的修正,就都要出PPH修正申請書
               'Modify By Sindy 2019/2/27 高速(PPH)審查要判斷已發文
               'Modify By Sindy 2023/6/12 敏莉:有關修正申請書，只要提過PPH，之後若有修正系統就會出PPH 修正申請書，
               '   但因進入再審階段，就不能用PPH 修正申請書，需用一般修正申請書，故若進度檔已有107 再審申請且已上發文日，則日後產生之修正申請書，產生一般修正申請書即可
               If PUB_ChkCPExist(cp, "431", 2) And Not PUB_ChkCPExist(cp, "107", 2) Then
               '2019/1/15 END
                  If StartLetter2("01", "17") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "17", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "發明專利PPH修正申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
                  
               Else
                  'Modify By Sindy 2019/8/16 修正-主動修正 ex:FCP-60590;反應為何要清除,先Mark
'                  'Add By Sindy 2018/4/27 單純的修正申請書不顯示”辦理依據”
'                  m_IPOSendDt = ""
'                  m_IPOSendData1 = "" 'Add By Sindy 2018/11/28
'                  m_IPOSendData2 = "" 'Add By Sindy 2018/11/28
'                  '2018/4/27 END
                  If StartLetter2("01", "15") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "15", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "專利修正申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               End If
               
               'Added by Morgan 2024/11/18
               '出主動修正申請書時，若有"447再審查加速審查"未發文，則一併產出"再審查加速審查"申請書 --敏莉
               If cp(10) = "203" Then
                  If PUB_ChkCPExist(cp, "447", 1) Then
                     m_IPOSendData1 = ""
                     If StartLetter2("01", "26") = False Then Exit Sub
                     NowPrint strReceiveNo, "01", "26", False, strUserNum, , , True, strExc(9)
                     strFileName = strFolder & "\" & "發明專利再審查加速審查申請書"
                     Call PUB_MakeDoc(strExc(9), strFileName)
                     intFileCnt = intFileCnt + 1
                  End If
               End If
               'end 2024/11/18
            
            Case "205" '申復
               'If chkAtt(1).Value = 1 Then
               If Frame204.Tag = "Y" Then
                  'Modify By Sindy 2019/1/29 有431高速(PPH)審查之後的修正,就都要出PPH修正申請書 ex:FCP-55352
                  'Modify By Sindy 2019/2/27 高速(PPH)審查要判斷已發文
                  'Modify By Sindy 2023/6/12 敏莉:有關修正申請書，只要提過PPH，之後若有修正系統就會出PPH 修正申請書，
                  '   但因進入再審階段，就不能用PPH 修正申請書，需用一般修正申請書，故若進度檔已有107 再審申請且已上發文日，則日後產生之修正申請書，產生一般修正申請書即可
                  If PUB_ChkCPExist(cp, "431", 2) And Not PUB_ChkCPExist(cp, "107", 2) Then
                     If StartLetter2("01", "17") = False Then Exit Sub
                     NowPrint strReceiveNo, "01", "17", False, strUserNum, , , True, strExc(9)
                     strFileName = strFolder & "\" & "發明專利PPH修正申請書"
                     Call PUB_MakeDoc(strExc(9), strFileName)
                     intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
                     
                  Else
                  '2019/1/29 END
                     If StartLetter2("01", "15") = False Then Exit Sub
                     NowPrint strReceiveNo, "01", "15", False, strUserNum, , , True, strExc(9)
                     strFileName = strFolder & "\" & "專利修正申請書"
                     Call PUB_MakeDoc(strExc(9), strFileName)
                     intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
                  End If
               Else
                  If StartLetter2("01", "20") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "20", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "審查意見申復申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               End If
               
            Case "416" '實體審查
               'If StartLetter2("01", "03", False) = False Then Exit Sub
               If StartLetter2("01", "03") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "03", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "發明專利實體審查申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               
               bolIsSecond = True 'Add By Sindy 2023/3/17 第2個申請書
               If Check416_1(0).Value = 1 And Check416_2(0).Value = 0 Then
                  'Modify By Sindy 2023/2/15 工程師點選實體審查並有勾選修正內容，
                  '同時產生實體審查+修正申請書，則主動修正申請書一律不帶辦理依據內容
                  m_IPOSendData1 = ""
                  '2023/2/15 END
                  'If StartLetter2("01", "15") = False Then Exit Sub
                  If StartLetter2("01", "15", IIf(intFileCnt = 0, True, False)) = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "15", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "專利修正申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               End If
               
               If Check416_2(0).Value = 1 Then
                  'If StartLetter2("01", "21") = False Then Exit Sub
                  If StartLetter2("01", "21", IIf(intFileCnt = 0, True, False)) = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "21", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "專利誤譯訂正申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               End If
               
            Case "107" '再審查
               'Add By Sindy 2018/7/30 再審查有延期過,並且沒有修正時出補正申請書
               If bolHad404 = True And _
                  (Check107_1(0).Value = 0 And Check107_2(0).Value = 0) Then
                  'If StartLetter2("01", "02", False) = False Then Exit Sub
                  If StartLetter2("01", "02") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "02", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "專利補正文件申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
                  
               ElseIf bolHad404 = False Then
               '2018/7/30 END
                  'If StartLetter2("01", "16", False) = False Then Exit Sub
                  If StartLetter2("01", "16") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "16", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "再審查申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               End If
               
               bolIsSecond = True 'Add By Sindy 2023/3/17 第2個申請書
               If Check107_1(0).Value = 1 And Check107_2(0).Value = 0 Then
'                  'Modify By Sindy 2019/1/29 有431高速(PPH)審查之後的修正,就都要出PPH修正申請書
'                  'Modify By Sindy 2019/2/27 高速(PPH)審查要判斷已發文
'                  If PUB_ChkCPExist(cp, "431", 2) Then
'                     'If StartLetter2("01", "17") = False Then Exit Sub
'                     If StartLetter2("01", "17", False) = False Then Exit Sub
'                     NowPrint strReceiveNo, "01", "17", False, strUserNum, , , True, strExc(9)
'                     strFileName = strFolder & "\" & "發明專利PPH修正申請書"
'                     Call PUB_MakeDoc(strExc(9), strFileName)
'                     intFileCnt = intFileCnt + 1'Add By Sindy 2019/10/31
'                  Else
'                  '2019/1/29 END
                     'If StartLetter2("01", "15") = False Then Exit Sub
                     If StartLetter2("01", "15", IIf(intFileCnt = 0, True, False)) = False Then Exit Sub
                     NowPrint strReceiveNo, "01", "15", False, strUserNum, , , True, strExc(9)
                     strFileName = strFolder & "\" & "專利修正申請書"
                     Call PUB_MakeDoc(strExc(9), strFileName)
                     intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
'                  End If
               End If
               
               If Check107_2(0).Value = 1 Then
                  'If StartLetter2("01", "21") = False Then Exit Sub
                  If StartLetter2("01", "21", IIf(intFileCnt = 0, True, False)) = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "21", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "專利誤譯訂正申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               End If
            
            Case "431" 'PPH審查
               'Modify By Sindy 2019/1/15 ex:FCP-59771 PPH審查的修正資料要出在審查申請書,修正申請書不出
               'If StartLetter2("01", "18", False) = False Then Exit Sub
               If StartLetter2("01", "18") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "18", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "發明專利PPH審查申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               
               bolIsSecond = True 'Add By Sindy 2023/3/17 第2個申請書
               'If Check431_1(0).Value = 1 Then
               If Check431_1(0).Value = 1 And Check431_2(0).Value = 0 Then
                  'If StartLetter2("01", "17") = False Then Exit Sub
                  If StartLetter2("01", "17", IIf(intFileCnt = 0, True, False)) = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "17", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "發明專利PPH修正申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
                  
               End If
               If Check431_2(0).Value = 1 Then
                  'If StartLetter2("01", "21") = False Then Exit Sub
                  If StartLetter2("01", "21", IIf(intFileCnt = 0, True, False)) = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "21", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "專利誤譯訂正申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               End If
               
            Case "422" '加速審查
               'Modify By Sindy 2025/9/25
               If pa(8) = "3" Then
                  strET03 = "27"
               Else
                  strET03 = "19"
               End If
               If StartLetter2("01", strET03) = False Then Exit Sub
               NowPrint strReceiveNo, "01", strET03, False, strUserNum, , , True, strExc(9)
               '2025/9/25 END
               strFileName = strFolder & "\" & "發明專利加速審查申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               
            Case "307" '分割
               If pa(8) = "1" Then
                  If StartLetter2("01", "01") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "發明專利分割申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
                  
               ElseIf pa(8) = "2" Then
                  If StartLetter2("01", "02") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "02", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "新型專利分割申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
                  
               Else
                  If StartLetter2("01", "03") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "03", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "設計專利分割申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
                  intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               End If
               
            Case "433" '誤譯訂正
               If StartLetter2("01", "21") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "21", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "誤譯訂正申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               'Modify By Sindy 2018/1/17 一併申請修正時不帶修正申請書
'               If Check433.Value = 1 Then
'                  If StartLetter2("01", "15") = False Then Exit Sub
'                  NowPrint strReceiveNo, "01", "15", False, strUserNum, , , True, strExc(9)
'                  strFileName = strFolder & "\" & "專利修正申請書"
'                  Call PUB_MakeDoc(strExc(9), strFileName)
'                  intFileCnt = intFileCnt + 1'Add By Sindy 2019/10/31
'               End If
               
            Case "408", "407" '面詢
               If StartLetter2("01", "23") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "23", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "面詢申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1 'Add By Sindy 2019/10/31
               
            'Add By Sindy 2024/5/28
            'Modify By Sindy 2025/2/18 修正:補充說明 是工程師操作, 產生專利補正文件申請書
            Case 補文件, 補充說明
               If StartLetter2("01", "02") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "02", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利補正文件申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1
            
            'Add By Sindy 2024/8/21
            Case "421" '申請技術報告
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "新型專利技術報告申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               intFileCnt = intFileCnt + 1
         End Select
         frm090904.Show
         frm090904.ClearForm
         Unload Me
      Case 1
         frm090904.Show
         frm090904.cmdok_Click 1
         Unload Me
   End Select
End Sub

Private Function ConvertNameFormat(pName As String) As String
   Dim strTmp As String

   If InStr(pName, ",") = 0 Then
      strTmp = Left(pName, 1) & "," & Mid(pName, 2)
   Else
      strTmp = pName
   End If
   ConvertNameFormat = strTmp
End Function

'申請書
'Modify By Sindy 2018/2/8 Optional ByVal bolAttachments As Boolean = True 是否含附送書件
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String, _
   Optional ByVal bolAttachments As Boolean = True) As Boolean
Dim strTxt(110) As String, strTmp As String, strTmp1 As String, strTmp2 As String
Dim ii As Integer, jj As Integer
Dim strInventor As String
Dim chk As CheckBox
Dim opt As OptionButton
Dim strED08 As String, strCP05 As String, strSendData1 As String, strSendData2 As String
Dim dblPayFee As Double
Dim strTmpCP64 As String, strCP64No As String
Dim strTmpET03_17 As String 'Add By Sindy 2023/3/27
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   'Add By Sindy 2020/9/21 抓取收據號碼
   '若工程師的主動修正申請書【本次應退還規費】有金額，則申請書的【備註】會帶: 檢附收據號碼第 號之電子收據，
   '憑辦理退費（退費支票抬頭請開台一國際專利法律事務所）。可將收據號碼帶入上述的" 第 號"，
   '其抓取收據號碼的判斷順序為: 有無收文1. (107)再審申請 2. (911)補收款(掛實體審查的相關總收文號) 3. (416)實體審查
   'Modify By Sindy 2023/3/14 + Or Val(txtDecreasePageFee) > 0
   If Val(txtDecreaseItemFee) > 0 Or Val(txtDecreasePageFee) > 0 Then
      strExc(0) = "select cp09||cp64 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='107'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTmpCP64 = RsTemp.Fields(0)
      Else
         strExc(0) = "select cp09||cp64,cp43 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='911' AND cp43 is not null AND exists (select c.cp09 from caseprogress c where c.cp09=cp43 and c.cp10='416')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strTmpCP64 = RsTemp.Fields(0)
         Else
            strExc(0) = "select cp09||cp64 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='416'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTmpCP64 = RsTemp.Fields(0)
            End If
         End If
      End If
      If InStr(strTmpCP64, "收據號碼:") > 0 Then
         strCP64No = Mid(strTmpCP64, InStr(strTmpCP64, "收據號碼:") + 5, 11)
      End If
   End If
   '2020/9/21 END
   
   'Add By Sindy 2023/2/20 同時辦理事項
   If chkAtt3(1).Value = 1 Or chkAtt3(3).Value = 1 Or chkAtt3(5).Value = 1 Then
      strTmp = "變更申請人之"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','同時辦理事項','♀')"
      If chkAtt3(1).Value = 1 Then
         strTmp = strTmp & "地址"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & chkAtt3(1).Caption & "','是')"
      End If
      If chkAtt3(3).Value = 1 Then
         strTmp = strTmp & IIf(chkAtt3(1).Value = 1, "/", "") & "代表人"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & chkAtt3(3).Caption & "','是')"
      End If
      If chkAtt3(5).Value = 1 Then
         strTmp = strTmp & IIf(chkAtt3(1).Value = 1 Or chkAtt3(3).Value = 1, "/", "") & "國籍"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & chkAtt3(5).Caption & "','是')"
      End If
      '申請書（非變更申請書）一律帶備註：變更申請人之地址/代表人/國籍（看是勾 選哪一個，就帶哪一個）如基本資料表上所述。
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','同時辦理變更的備註','" & strTmp & "如基本資料表上所述。')"
   End If
   '2023/2/20 END
   
   'Add By Sindy 2019/10/3
   '修改工程師產生的實審(416.實體審查)+主動修正申請書及PPH(431.PPH審查)+主動修正申請書上的【事務所或申請人案件編號】為 FCP-0XXXXX.2
   'Modify By Sindy 2019/10/15 10/15 分割+實審或分割+續行母案;也要加.2
   'Modify By Sindy 2023/3/17 排除 bolIsSecond = True : 第2個申請書
   If bolIsSecond = False Then
   '2023/3/17 END
      If (cp(10) = "431" And (Check431_1(0).Value = 1 Or Check431_2(0).Value = 1)) Or _
         cp(10) = "416" Or _
         (cp(10) = "307" And (Check2.Value = 1 Or Check9(0).Value = 1)) Or _
         (chkAtt3(1).Value = 1 Or chkAtt3(3).Value = 1 Or chkAtt3(5).Value = 1) Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','電子稽核防止漏發文','♀')"
      End If
   End If
   '2019/10/3 END
   
   '本所案號
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   'Add By Sindy 2018/2/9
   If cp(10) = "107" Then '再審查
      'Modify By Sindy 2019/1/31 補正文件申請書:有1004.延期受理,抓延期受理函號,若無再抓核駁的函號
      'Modify By Sindy 2019/2/12 取消 And ET03 = "02";修正申請書亦同
      If bolHad404 = True And strHad404CP09 <> "" Then
         strExc(0) = "select cp05,cp08,ed08 from caseprogress,edocument" & _
                     " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                     " AND ed11(+)=cp09 AND cp43='" & strHad404CP09 & "' and cp10='1004'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(RsTemp("ED08")) Then
               m_IPOSendDt = RsTemp("ED08") - 19110000
               If Not IsNull(RsTemp("cp08")) Then
                  strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                  m_IPOSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), m_IPOSendData1 & "字第", "")
                  m_IPOSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
               End If
            End If
         End If
      End If
      '2019/1/31 END
      '1002.核駁-新申請案
      strExc(0) = "select cp05,cp08,ed08 from caseprogress,edocument" & _
                  " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                  " AND ed11(+)=cp09 AND cp10='1002'" & _
                  " AND cp43 in(select cp09 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10 in(" & NewCasePtyList & ")) ORDER BY CP05 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp("ED08")) Then
            strED08 = RsTemp("ED08") - 19110000
            If Not IsNull(RsTemp("cp08")) Then
               strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
               strSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
               strExc(0) = Replace(strExc(0), strSendData1 & "字第", "")
               strSendData2 = Replace(Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1), "號", "")
            End If
         End If
         If Not IsNull(RsTemp("cp05")) Then
            strCP05 = RsTemp("cp05") - 19110000
         End If
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','初審官方發文日','" & ChangeTStringToTDateString(strED08) & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','初審智專字','" & strSendData1 & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','初審發文號','" & strSendData2 & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','初審收文日期','" & ChangeTStringToTDateString(strCP05) & "')"
   End If
   
   '辦理依據
   'Modify By Sindy 2018/11/28 + Or m_IPOSendData1 <> ""
   If m_IPOSendDt <> "" Or m_IPOSendData1 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(m_IPOSendDt) & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & m_IPOSendData1 & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & m_IPOSendData2 & "')"
   'Modify By Sindy 2024/5/28
   ElseIf cp(10) = 補文件 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','♀')"
      '2024/5/28 END
   End If
   
   '申請人
   'Add By Sindy 2023/8/11
   If cp(1) = "FG" Then
      '目前只有230提供情報是給工程師自行填寫意見提交人
   Else
   '2023/8/11 END
      Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())
   End If
   
   '出名代理人
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03" '取消order by OA03:依存入的順序
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   
   'Modify By Sindy 2018/5/11 有無繳費金額 ex:FCP-049790(申復)
   'Modify By Sindy 2019/1/16 一併送中說 ex:FCP-059648(主動修正)
   'If Val(txtCP84) = 0 And Val(txtAddItemFee) = 0 Then
   'Modify By Sindy 2019/2/15 再審查+修正(無增刪項數) ex:FCP-056056(再審查)
'   If Val(txtAddItem) = 0 And Val(txtCP137) = 0 And Val(txtCP138) = 0 Then
   'Modify By Sindy 2019/2/15 有新增項數和刪除未審項數,才算有變動規費 或 實審未發文規費不變
   'Modify By Sindy 2019/4/1 + PPH審查+修正,修正時若有增刪項數要顯示規費資訊 And Not (cp(10) = "431" And (Val(txtAddItem) = 0 And Val(txtCP137) = 0))
   'Modify By Sindy 2019/7/1 + (PUB_ChkCPExist(cp, "435", 2) = False And PUB_ChkCPExist(cp, "435", 0) = True And Check2.Value = 0) ex:FCP-60369
   'Modify By Sindy 2019/8/30 + 307分割,416實審,107再審查會鎖住增減項,帶規費不變動 => FrameFee.Enabled = False
   'Modify By Sindy 2019/10/4 設計案沒有專利範圍
   'Modify By Sindy 2019/12/13 有修正專利範圍,才須要顯示規費有無變動
   'Modified by Morgan 2023/1/9 改發明案才要帶--敏莉
   If pa(8) = "1" And SSTab1.TabVisible(1) = True And _
      (chk1Tab1(0).Value = 1 Or chk1Tab1(1).Value = 1 Or chk1Tab1(3).Value = 1 Or chk1Tab1(4).Value = 1 _
       Or chk1Tab1(5).Value = 1 Or chk1Tab1(6).Value = 1 Or chk1Tab1(8).Value = 1 _
       Or chk1Tab2(0).Value = 1 Or chk1Tab2(1).Value = 1 Or chk1Tab2(3).Value = 1 Or chk1Tab2(4).Value = 1 _
       Or chk1Tab2(5).Value = 1 Or chk1Tab2(6).Value = 1 Or chk1Tab2(8).Value = 1 _
       Or chk1Tab2(9).Value = 1 Or chk1Tab2(10).Value = 1 Or chk1Tab2(12).Value = 1 Or chk1Tab2(13).Value = 1 _
      ) Then
      'Modify By Sindy 2023/5/8 + FCP-68744,實審只繳7000沒繳超項4000,但這次的應加收4000元在中文本資訊裡原始是15項,主動修正的規費項算不變
      '                         : + (m_WriteNote = "Y" And (Val(txtAddItem) = 0 And Val(txtCP137) = 0 And Val(txtAddPage) = 0 And Val(txtCP167) = 0))
      'Modify By Sindy 2023/11/16 + FCP-69836,有輸增加1頁但未超過級距沒有應加收規費,無變動
      If FrameFee.Enabled = False Or _
         (PUB_ChkCPExist(cp, "416", 2) = False And PUB_ChkCPExist(cp, "416", 0) = True And Check2.Value = 0) Or _
         (PUB_ChkCPExist(cp, "435", 2) = False And PUB_ChkCPExist(cp, "435", 0) = True And Check2.Value = 0) Or _
         (bolAttachments = False And cp(10) <> "431") Or _
         (Val(txtAddItemFee) = 0 And Val(txtDecreaseItemFee) = 0 And Val(txtAddPageFee) = 0 And Val(txtDecreasePageFee) = 0) Then 'Or _
         (m_WriteNote = "Y" And ( _
                  (Val(txtAddItem) = 0 Or Val(txtAddItemFee) = 0) And _
                  (Val(txtCP137) = 0 Or Val(txtDecreaseItemFee) = 0) And _
                  (Val(txtAddPage) = 0 Or Val(txtAddPageFee) = 0) And _
                  (Val(txtCP167) = 0 Or Val(txtDecreasePageFee) = 0) _
                                )) Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','規費不變','♀')"
         'Add By Sindy 2023/3/10
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數規費不變','♀')"
         '2023/3/10 END
         strTmpET03_17 = "發明專利案修正後總頁數與修正前總頁數相較-應繳規費不變" 'Add By Sindy 2023/3/27
         
      Else
         'If (Val(txtAddItem) = 0 And Val(txtCP137) = 0) Then
         'If Val(txtAddItemFee) = 0 And Val(txtDecreaseItemFee) = 0 Then
         'Modify By Sindy 2023/11/16 + FCP-69836
         If (Val(txtAddItemFee) = 0 Or Val(txtAddItem) = 0) And _
            (Val(txtDecreaseItemFee) = 0 Or Val(txtCP137) = 0) Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','規費不變','♀')"
         Else
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','規費有變動','♀')"
            '規費變動
            If Val(txtAddItem) > 0 Or Val(txtCP137) > 0 Or Val(txtCP138) > 0 Then
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','原請求項數','" & Val(txtItem) & "')"
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','新增項數','" & Val(txtAddItem) & "')"
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','刪除項數','" & IIf(Val(txtCP137) + Val(txtCP138) = 0, 0, Val(txtCP137) + Val(txtCP138)) & "')"
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','修正後總項數','" & Val(txtCount) & "')"
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本次應加規費','" & Val(txtAddItemFee) & "')"
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本次應退規費','" & Val(txtDecreaseItemFee) & "')"
            End If
         End If
         'Add By Sindy 2023/3/10
         'If (Val(txtAddPage) = 0 And Val(txtCP167) = 0) Then
         'If Val(txtAddPageFee) = 0 And Val(txtDecreasePageFee) = 0 Then
         'Modify By Sindy 2023/11/16 + FCP-69836
         If (Val(txtAddPageFee) = 0 Or Val(txtAddPage) = 0) And _
            (Val(txtDecreasePageFee) = 0 Or Val(txtCP167) = 0) Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數規費不變','♀')"
            strTmpET03_17 = "發明專利案修正後總頁數與修正前總頁數相較-應繳規費不變" 'Add By Sindy 2023/3/27
         Else
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數規費有變動','♀')"
            strTmpET03_17 = "發明專利案修正後總頁數與修正前總頁數相較-應繳規費有變動" & vbCrLf 'Add By Sindy 2023/3/27
            '頁數規費變動
            If Val(txtAddPage) > 0 Or Val(txtCP167) > 0 Or Val(txtCP168) > 0 Then
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','修正前總頁數','" & Val(txtPage) & "')"
               strTmpET03_17 = strTmpET03_17 & "修正前總頁數：" & Val(txtPage) & vbCrLf 'Add By Sindy 2023/3/27
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','修正後總頁數','" & Val(txtPageCount) & "')"
               strTmpET03_17 = strTmpET03_17 & "修正後總頁數：" & Val(txtPageCount) & vbCrLf 'Add By Sindy 2023/3/27
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','新增頁數','" & Val(txtAddPage) & "')"
               strTmpET03_17 = strTmpET03_17 & "新增頁數：" & Val(txtAddPage) & vbCrLf 'Add By Sindy 2023/3/27
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','刪除頁數','" & IIf(Val(txtCP167) + Val(txtCP168) = 0, 0, Val(txtCP167) + Val(txtCP168)) & "')"
               strTmpET03_17 = strTmpET03_17 & "刪除頁數：" & IIf(Val(txtCP167) + Val(txtCP168) = 0, 0, Val(txtCP167) + Val(txtCP168)) & vbCrLf 'Add By Sindy 2023/3/27
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本次頁數應加規費','" & Val(txtAddPageFee) & "')"
               strTmpET03_17 = strTmpET03_17 & "本次應加收規費：" & Val(txtAddPageFee) & vbCrLf 'Add By Sindy 2023/3/27
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本次頁數應退規費','" & Val(txtDecreasePageFee) & "')"
               strTmpET03_17 = strTmpET03_17 & "本次應退還規費：" & Val(txtDecreasePageFee) & vbCrLf 'Add By Sindy 2023/3/27
            End If
         End If
         '2023/3/10 END
      End If
   End If
   
   '繳費金額
   ii = ii + 1
   dblPayFee = 0 'Add By Sindy 2019/8/27 繳費金額
   'Modif By Sindy 2018/7/30 若實審,再審,PPH審查同時修正時,修正誤譯規費為0;規費出在審查書
   'Add By Sindy 2018/6/11 107.再審查+修正時繳費金額為0,繳費金額顯示於另外的申請書中
   'Add By Sindy 2018/7/30 再審查有延期過,則不再出再審查申請書;因此不管補正或修正若須繳費均出在其申請書上
   If (cp(10) = "107" And ET03 <> "16") And (Check107_1(0).Value = 1 Or Check107_2(0).Value = 1) And bolHad404 = False Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','0')"
   'Add By Sindy 2018/6/11 416.實體審查+修正時繳費金額為0,繳費金額顯示於另外的申請書中
   ElseIf (cp(10) = "416" And ET03 <> "03") And (Check416_1(0).Value = 1 Or Check416_2(0).Value = 1) Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','0')"
   'Add By Sindy 2018/6/11 431.PPH審查+修正時繳費金額為0,繳費金額顯示於另外的申請書中
   ElseIf (cp(10) = "431" And ET03 <> "18") And (Check431_1(0).Value = 1 Or Check431_2(0).Value = 1) Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','0')"
   Else
   '2018/6/11 END
      dblPayFee = Val(txtCP84) 'Add By Sindy 2019/8/27 繳費金額
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   End If
   
   '*********************************************************************************
   '依案件性質讀取的資料區
   '*********************************************************************************
   If ET03 = "15" Then '修正
      'Modify By Sindy 2018/5/8 Mark:繳費金額改放到中說申請書
'      If m_WriteNote = "Y" And _
'         (m_lngOverPageFee > 0 Or m_lngOverItemFee > 0) Then
'         If m_lngOverPageFee > 0 And m_lngOverItemFee > 0 Then
'            strTmp = "繳費金額為超頁規費" & Format(m_lngOverPageFee, "#,##0") & "元整，超項規費" & Format(m_lngOverItemFee, "#,##0") & "元整，共計" & Format(m_lngOverPageFee + m_lngOverItemFee, "#,##0") & "元整。"
'         ElseIf m_lngOverPageFee > 0 Then
'            strTmp = "繳費金額為超頁規費" & Format(m_lngOverPageFee, "#,##0") & "元整。"
'         Else
'            strTmp = "繳費金額為超項規費" & Format(m_lngOverItemFee, "#,##0") & "元整。"
'         End If
      'Modify By Sindy 2019/3/4 本案續行再審查 ex:FCP-060443(主動修正)
      If Check9(1).Visible = True And Check9(1).Value = 1 Then
         strTmp = "一併補繳續行母案再審查規費。"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
      'Modify By Sindy 2019/1/16 一併送中說 ex:FCP-059648(主動修正)
      ElseIf m_WriteNote = "Y" Then
         strTmp = ""
         If (Val(m_lngOverPageFee) > 0 Or Val(m_lngOverItemFee) > 0) Then 'And Val(txtAddItemFee) > 0
            'Modify By Sindy 2021/12/24 修正實審已發文，中說+修正申請書，補述內容  FCP-065711 (暫緩)
            strTmp = "應繳納之金額於同日補正中文本時繳納。"
            'strTmp = "本案同日另函提出補正中文本，修正後" & IIf(m_lngOverPageFee > 0, "超頁規費" & IIf(m_lngOverItemFee > 0, "、超項規費", ""), "超項規費") & "請見補正申請書。"
         'Add By Sindy 2023/4/7
         ElseIf Val(txtDecreaseItemFee) > 0 Or Val(txtDecreaseItemFee) > 0 Then
            strTmp = "應退費之金額於同日補正中文本時申請退費。"
         End If
         '2023/4/7 END
         If strTmp <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
         End If
      'Add By Sindy 2019/8/27 ex:FCP-054503 再審查有延期,有修正,產出修正申請書
      'Modify By Sindy 2020/2/25
'      ElseIf dblPayFee > 0 And cp(10) = "107" And (m_lngOverPageFee > 0 Or m_lngOverItemFee > 0) Then
      ElseIf dblPayFee > 0 And cp(10) = "107" Then
      '2020/2/25 END
         If m_lngOverPageFee > 0 Or m_lngOverItemFee > 0 Then
            strTmp = "繳費金額為再審查規費及" & IIf(m_lngOverPageFee > 0, "超頁規費" & IIf(m_lngOverItemFee > 0, "、超項規費", ""), "超項規費") & "。"
         'Add By Sindy 2020/2/25
         Else
            strTmp = "繳費金額為再審查規費。"
         End If
         '2020/2/25 END
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
      '2019/8/27 END
      'Modify By Sindy 2019/8/16 Mark,敏莉反應不須加註,因實審/再審查會和修正申請書一起送出去
'      '實審
'      ElseIf cp(10) = "416" And Val(txtAddItemFee) > 0 Then
'         strTmp = "應繳交之金額，於申請實體審查時一併繳納。"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
'      '再審查
'      ElseIf cp(10) = "107" And Val(txtAddItemFee) > 0 Then
'         strTmp = "應繳交之金額，於申請再審查時一併繳納。"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
      '2019/8/16 END
      'Add By Sindy 2019/4/1
      'Modify By Sindy 2023/3/14 + Or Val(txtDecreasePageFee) > 0
      ElseIf Val(txtDecreaseItemFee) > 0 Or Val(txtDecreasePageFee) > 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','檢附收據號碼第" & IIf(strCP64No <> "", strCP64No, "           ") & "號之電子收據，憑辦理退費（退費支票抬頭請開" & CompNameQuery("2") & "）。')"
      '2019/4/1 END
'      Else
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','否')"
      'Add By Sindy 2023/3/23
      'Modify By Sindy 2023/4/11 敏莉說不要帶此備註=>本案一併繳納超項規費。
'      ElseIf m_lngOverPageFee > 0 Or m_lngOverItemFee > 0 Then
'         strTmp = ""
'         If m_lngOverPageFee > 0 Then
'            strTmp = "超頁"
'         End If
'         If m_lngOverItemFee > 0 Then
'            If strTmp <> "" Then strTmp = strTmp & "、"
'            strTmp = "超項"
'         End If
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','本案一併繳納" & strTmp & "規費。')"
         '2023/3/23 END
      'Add By Sindy 2023/3/23
      ElseIf bolIsSecond = True Then 'And (cp(10) = "416" Or cp(10) = "107")
         strTmp = GetPrjState6(cp(1), cp(10))
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','此修正與" & strTmp & "同時提出申請，相關繳費資訊請參見" & strTmp & "申請書。')"
         '2023/3/23 END
      End If
      '2018/5/8 END
      'Modify By Sindy 2019/4/29 Mark
'      'Add By Sindy 2018/12/17
'      If cp(10) = "107" And chk3Tab2(0).Value = 1 Then '再審查
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-再審查理由書','" & m_CaseNo & chk3Tab2(0).Tag & "')"
'      End If
'      '2018/12/17 END
   'Add By Sindy 2019/4/1
   ElseIf ET03 = "17" Then 'PPH修正
      strTmp = ""
      'Add By Sindy 2021/2/17 431.高速審查若已發文,不顯示下列備註
      If PUB_ChkCPExist(cp, "431", 2) = False Then
      '2021/2/17 END
         'Modify By Sindy 2023/3/14 + Or Val(txtAddPageFee) > 0
         If Val(txtAddItemFee) > 0 Or Val(txtAddPageFee) > 0 Then
            strTmp = "本次應加收之規費，於發明專利PPH審查申請書繳納。"
         'Modify By Sindy 2023/3/14 + Or Val(txtDecreasePageFee) > 0
         ElseIf Val(txtDecreaseItemFee) > 0 Or Val(txtDecreasePageFee) > 0 Then
            strTmp = "本次應退還規費，於發明專利PPH審查申請書退費。"
         End If
      End If
      If strTmpET03_17 <> "" Then
         strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmpET03_17
      End If
      If strTmp <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
      End If
   '2019/4/1 END
   ElseIf ET03 = "21" Then '誤譯訂正
      'If Check433.Value = 1 Then
      If Check433.Value = 1 Or Check107_1(0).Value = 1 Or Check416_1(0).Value = 1 Or Check431_1(0).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','是')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','修正事項','♀')"
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','否')"
      End If
      'Modify By Sindy 2019/4/29 Mark
'      'Add By Sindy 2018/12/17
'      If cp(10) = "107" And chk3Tab2(0).Value = 1 Then '再審查
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-再審查理由書','" & m_CaseNo & chk3Tab2(0).Tag & "')"
'      End If
'      '2018/12/17 END
      'Add By Sindy 2019/4/1 431.PPH審查
      If cp(10) = "431" Then
         'Modify By Sindy 2023/3/14 + Or Val(txtAddPageFee) > 0
         If Val(txtAddItemFee) > 0 Or Val(txtAddPageFee) > 0 Then
            strTmp = "本次應加收之規費，於發明專利PPH審查申請書繳納。"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
         'Modify By Sindy 2023/3/14 + Or Val(txtDecreasePageFee) > 0
         ElseIf Val(txtDecreaseItemFee) > 0 Or Val(txtDecreasePageFee) > 0 Then
            strTmp = "本次應退還規費，於發明專利PPH審查申請書退費。"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
         End If
'      Else
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','否')"
      End If
      '2019/4/1 END
   ElseIf cp(10) = "107" Then '再審查
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-再審查理由書','" & m_CaseNo & chk3Tab2(0).Tag & "')"
      If Check107_1(0).Value = 1 And Check107_2(0).Value = 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','是')"
         If Check107_2(0).Value = 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-專利修正申請書','" & m_CaseNo & ".AMD.DATA.ATT.pdf')" '專利修正申請書.pdf
         End If
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','否')"
      End If
      If Check107_2(0).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','是')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-專利誤譯訂正申請書','" & m_CaseNo & ".COR.DATA.ATT.pdf')"
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','否')"
      End If
      'Add By Sindy 2019/1/31 補正文件申請書不須顯示此資料
      If ET03 <> "02" Then
      '2019/1/31 END
         If txtDocCh2(0).Enabled = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','摘要頁數','" & txtDocCh2(0) & "')"
         End If
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明書頁數','" & txtDocCh2(1) & "')"
         If txtDocCh2(2).Enabled = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請專利範圍頁數','" & txtDocCh2(2) & "')"
         End If
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式頁數','" & txtDocCh2(3) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數','" & txtDocCh2(4) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','項數','" & txtDocCh2(5) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式圖數','" & txtDocCh2(6) & "')"
      End If
      'Modify By Sindy 2019/5/2 序列表 加註備註
      'Modify By Sindy 2019/8/27 ex:FCP-054503 再審查有延期沒修正時出補正申請書
      'Modify By Sindy 2020/2/25
'      If Val(txtDocCh2(7)) > 0 Or _
'         (dblPayFee > 0 And ET03 = "02" And (m_lngOverPageFee > 0 Or m_lngOverItemFee > 0)) Then
      If Val(txtDocCh2(7)) > 0 Or _
         (dblPayFee > 0 And ET03 = "02") Then
      '2020/2/25 END
         strTmp = ""
         If Val(txtDocCh2(7)) > 0 Then
            strTmp = "序列表" & Val(txtDocCh2(7)) & "頁不納入超頁費之計算。"
         End If
         If dblPayFee > 0 And ET03 = "02" Then
            If m_lngOverPageFee > 0 Or m_lngOverItemFee > 0 Then
               strTmp = strTmp & "繳費金額為再審查規費及" & IIf(m_lngOverPageFee > 0, "超頁規費" & IIf(m_lngOverItemFee > 0, "、超項規費", ""), "超項規費") & "。"
            'Add By Sindy 2020/2/25
            Else
               strTmp = strTmp & "繳費金額為再審查規費。"
            End If
            '2020/2/25 END
         End If
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
         '2019/8/27 END
'      Else
'      '2019/5/2 END
'         'Add By Sindy 2018/7/30
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','否')"
      End If
   ElseIf cp(10) = "431" Then '高速審查
      If Check431_1(0).Value = 1 And Check431_2(0).Value = 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請PPH修正','是')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-專利PPH修正申請書','" & m_CaseNo & ".PPH.AMD.DATA.ATT.pdf')"
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請PPH修正','否')"
      End If
      If Check431_2(0).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','是')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-專利誤譯訂正申請書','" & m_CaseNo & ".COR.DATA.ATT.pdf')"
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','否')"
      End If
      'Add By Sindy 2019/4/1
      'Modify By Sindy 2023/3/14 + Or Val(txtDecreasePageFee) > 0
      If Val(txtDecreaseItemFee) > 0 Or Val(txtDecreasePageFee) > 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','檢附收據號碼第" & IIf(strCP64No <> "", strCP64No, "           ") & "號之電子收據，憑辦理退費（退費支票抬頭請開" & CompNameQuery("2") & "）。')"
      End If
      '2019/4/1 END
   ElseIf cp(10) = "416" Then '實體審查
      If Check416_1(0).Value = 1 And Check416_2(0).Value = 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','是')"
         If Check416_2(0).Value = 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-專利修正申請書','" & m_CaseNo & ".AMD.DATA.ATT.pdf')" '專利修正申請書.pdf
         End If
      'Add By Sindy 2018/1/17
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','否')"
      '2018/1/17 END
      End If
      If Check416_2(0).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','是')"
         'Modify By Sindy 2020/8/28 敏莉:
         '另增加判斷，當工程師出實審申請書若申請書上的"【一併申請誤譯訂正】"為" 是"，
         '則申請書的" 【附送書件】"的"【專利誤譯訂正申請書】 "的檔名請帶"FCP0XXXXX.COR.DATA.ATT.pdf"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-專利誤譯訂正申請書','" & m_CaseNo & ".COR.DATA.ATT.pdf')" '專利誤譯訂正申請書.pdf
      'Add By Sindy 2018/1/17
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','否')"
      '2018/1/17 END
      End If
      If txtDocCh(0).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','摘要頁數','" & txtDocCh(0) & "')"
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明書頁數','" & txtDocCh(1) & "')"
      If txtDocCh(2).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍頁數','" & txtDocCh(2) & "')"
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式頁數','" & txtDocCh(3) & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數總計','" & txtDocCh(4) & "')"
      If txtDocCh(5).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍項數','" & txtDocCh(5) & "')"
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式圖數','" & txtDocCh(6) & "')"
      
      'Modify By Sindy 2019/5/2 序列表 加註備註
      'Modify By Sindy 2020/8/28 亭妙(FCP-061760):並於申請書【備註】增加段落:繳納金額已包含誤譯訂正之規費
      If Val(txtDocCh(7)) > 0 Or Check416_2(0).Value = 1 Then
         strTmp = ""
         If Val(txtDocCh(7)) > 0 Then strTmp = IIf(strTmp <> "", strTmp & "；", "") & "序列表" & Val(txtDocCh(7)) & "頁不納入超頁費之計算。"
         If Check416_2(0).Value = 1 Then strTmp = IIf(strTmp <> "", strTmp & "；", "") & "繳納金額已包含誤譯訂正之規費。"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
      End If
      '2019/5/2 END
   ElseIf cp(10) = "307" Then '分割
      If Frame307.Enabled = True Then
         If Check2.Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請實體審查','是')"
         End If
         If Check9(0).Value = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本案續行再審查','是')"
         End If
         'Add By Sindy 2018/4/23
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','另申請專利之聲明','" & IIf(Check1.Value = 1, "是", "否") & "')"
         '2018/4/23 END
      End If
      '專利案件屬性名稱
      If pa(158) <> "" Then
         ii = ii + 1
         'Modify By Sindy 2019/8/7 PUB_GetCaseAttributeName(pa(158)) & IIf(pa(8) = "3", "設計", "") ==> PUB_GetCaseAttributeName(pa(158), pa(8)) & IIf(pa(8) = "3", "設計", "")
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利案件屬性名稱','" & PUB_GetCaseAttributeName(pa(158), pa(8)) & IIf(pa(8) = "3", "設計", "") & "')"
      End If
      '原申請案號
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','原申請案號','" & strDivPA11 & "')"
      
      'Add By Sindy 2018/7/26
      '附英文摘要
      If Check3.Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附英文摘要','♀')"
         'Modify By Sindy 2019/5/8
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','本案未附英文說明書，但所檢附之申請書中發明名稱、申請人姓名或名稱、發明人姓名及摘要同時附有英文翻譯，故可減收申請規費800元整。')"
         '2019/5/8 END
      End If
      '2018/7/26 END
      
      '讀取發明人資料
      If pa(8) = "1" Then
         strExc(1) = "發明人"
      ElseIf pa(8) = "2" Then
         strExc(1) = "新型創作人"
      Else
         strExc(1) = "設計人"
      End If
      strInventor = ""
      strExc(0) = " SELECT IN03,IN04,IN05,IN11,NA72" & _
                  " FROM PatentInventor,INVENTOR,NATION" & _
                  " WHERE pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
                  " AND IN01=substr(pi06,1,8) AND IN02=substr(pi06,9,2)" & _
                  " AND NA01(+)=IN11" & _
                  " order by pi05 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         jj = 1
         Do While Not RsTemp.EOF
            If strInventor <> "" Then strInventor = strInventor & vbCrLf
            'Modify By Sindy 2018/10/25 增加英文名稱格式化 PUB_FCPIN05Format_EName
            strInventor = strInventor & "【" & strExc(1) & jj & "】" & vbCrLf & _
                                        "　　【國籍】　　　　　　" & RsTemp("NA72") & vbCrLf & _
                                        "　　【中文姓名】　　　　" & ChgSQL("" & RsTemp("IN04")) & _
                                        IIf("" & RsTemp("IN05") = "", "", vbCrLf & "　　【英文姓名】　　　　" & ChgSQL(PUB_FCPIN05Format_EName("" & RsTemp("IN05"), "" & RsTemp("NA72"))))
            jj = jj + 1
            RsTemp.MoveNext
         Loop
      Else
         strInventor = "【" & strExc(1) & "1】" & vbCrLf & _
                       "　　【國籍】　　　　　　" & vbCrLf & _
                       "　　【中文姓名】　　　　"
      End If
      If Not (pa(1) = "FCP" And InStr("101,102,103,307", cp(10)) > 0) Then 'Added by Lydia 2022/03/03 排除FCP分割申請書，改用<申請書發明人資料>；ex.FCP-066642發明申請書的發明人資料超過4000字
          ii = ii + 1
          strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strExc(1) & "資料','" & strInventor & "')"
      End If
      '優惠期發生日期
      If txtFavDate <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優惠期發生日期','" & ChangeTStringToWDateString(txtFavDate) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優惠期原因','" & cboFavReason.Text & "')"
      End If
      '中文本資訊
      If chkDoc(0).Value = 1 Then
         If txtDocCh(0).Enabled = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','摘要頁數','" & Val(txtDocCh(0)) & "')"
         End If
         If txtDocCh(1).Enabled = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明書頁數','" & Val(txtDocCh(1)) & "')"
         End If
         If txtDocCh(2).Enabled = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍頁數','" & Val(txtDocCh(2)) & "')"
         End If
         'Modify By Sindy 2019/8/7 Ex:FCP-061566
         'If Val(txtDocCh(3)) > 0 And txtDocCh(2).Enabled = True Then
         If txtDocCh(3).Enabled = True Then
         '2019/8/7 END
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式頁數','" & Val(txtDocCh(3)) & "')"
         End If
         If txtDocCh(4).Enabled = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數總計','" & Val(txtDocCh(4)) & "')"
         End If
         If txtDocCh(5).Enabled = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍項數','" & Val(txtDocCh(5)) & "')"
         End If
         If Val(txtDocCh(6)) > 0 And txtDocCh(6).Enabled = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式圖數','" & Val(txtDocCh(6)) & "')"
         End If
      End If
      '外文本資訊
      If chkDoc(1).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','外文頁數總計','" & Val(txtForeign) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','外文本種類','" & Trim(cboLagnuage) & "')"
      End If
      '簡體字本資訊
      If chkDoc(2).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','簡體字頁數總計','" & Val(txtSimplified) & "')"
      End If
   End If
   
   '*******************************************************************************
   '附送書件
   '*******************************************************************************
   strTmp = ""
   '修正/誤譯訂正Tab
   'If SSTab1.TabVisible(1) = True And cp(10) <> "416" Then '實體審查申請書不顯示修正/誤譯訂正附送書件
   If SSTab1.TabVisible(1) = True And bolAttachments = True Then
      'If Check416_1(0).Value = 1 Then
         For Each chk In chk1Tab1
            If chk.Value = 1 Then
               strTmp1 = "": strTmp2 = ""
               strTmp1 = chk.Caption
               If strTmp = "" Then
                  strTmp1 = "　【" & strTmp1 & "】"
                  If Len(strTmp1) < 14 Then
                     strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
                  End If
               Else
                  strTmp1 = "　　【" & strTmp1 & "】"
                  If Len(strTmp1) < 15 Then
                     strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
                  End If
               End If
               If chk.Tag <> "" Then
                  If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
                     strTmp2 = Mid(chk.Tag, 2)
                  Else
                     strTmp2 = Trim(chk.Tag)
                  End If
                  If strTmp2 <> "" Then
                     strTmp2 = m_CaseNo & strTmp2
                  End If
               End If
               strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
            End If
         Next
         'Add By Sindy 2020/2/15 + FCP-053135 '再審查有延期過,只出修正申請書,附送書件要出現"再審查理由書"
         If SSTab1.TabVisible(3) = True And bolHad404 = True And chk3Tab2(0).Value = 1 Then
            strTmp1 = chk3Tab2(0).Caption
            If strTmp = "" Then
               strTmp1 = "　【" & strTmp1 & "】"
               If Len(strTmp1) < 14 Then
                  strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
               End If
            Else
               strTmp1 = "　　【" & strTmp1 & "】"
               If Len(strTmp1) < 15 Then
                  strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
               End If
            End If
            strTmp2 = m_CaseNo & chk3Tab2(0).Tag
            strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
         End If
         '2020/2/15 END
      'End If
      'If Check416_2(0).Value = 1 Then
         For Each chk In chk1Tab2
            If chk.Value = 1 Then
               strTmp1 = "": strTmp2 = ""
               strTmp1 = chk.Caption
               If strTmp = "" Then
                  strTmp1 = "　【" & strTmp1 & "】"
                  If Len(strTmp1) < 14 Then
                     strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
                  End If
               Else
                  strTmp1 = "　　【" & strTmp1 & "】"
                  If Len(strTmp1) < 15 Then
                     strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
                  End If
               End If
               If chk.Tag <> "" Then
                  If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
                     strTmp2 = Mid(chk.Tag, 2)
                  Else
                     strTmp2 = Trim(chk.Tag)
                  End If
                  If strTmp2 <> "" Then
                     strTmp2 = m_CaseNo & strTmp2
                  End If
               End If
               strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
            End If
         Next
      'End If
   End If
   '分割/實審Tab
   If SSTab1.TabVisible(0) = True And bolAttachments = True Then
      If strTmp <> "" And cp(10) = "416" Then GoTo ReadEnd_416 'Add By Sindy 2020/1/7
      For Each chk In chk0Tab
         If chk.Value = 1 Then
            strTmp1 = "": strTmp2 = ""
            strTmp1 = chk.Caption
            If strTmp1 = "說明書" Or strTmp1 = "圖式" Or strTmp1 = "序列表" Then
               strTmp1 = "外文本"
               'Add By Sindy 2019/8/7
               If InStr(strTmp, strTmp1) > 0 Then
                  GoTo ReadNext
               End If
               '2019/8/7 END
            End If
            If strTmp = "" Then
               strTmp1 = "　【" & strTmp1 & "】"
               If Len(strTmp1) < 14 Then
                  strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
               End If
            Else
               strTmp1 = "　　【" & strTmp1 & "】"
               If Len(strTmp1) < 15 Then
                  strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
               End If
            End If
            If chk.Tag <> "" Then
               If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
                  strTmp2 = Mid(chk.Tag, 2)
               Else
                  strTmp2 = Trim(chk.Tag)
               End If
               If strTmp2 <> "" Then
                  strTmp2 = m_CaseNo & strTmp2
               End If
            End If
            strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
ReadNext:
         End If
      Next
ReadEnd_416:
   End If
   'Add By Sindy 2019/10/31
   '更正Tab
   If SSTab1.TabVisible(4) = True And bolAttachments = True Then
      For Each chk In chk4Tab1
         If chk.Value = 1 Then
            strTmp1 = "": strTmp2 = ""
            strTmp1 = chk.Caption
            If strTmp = "" Then
               strTmp1 = "　【" & strTmp1 & "】"
               If Len(strTmp1) < 14 Then
                  strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
               End If
            Else
               strTmp1 = "　　【" & strTmp1 & "】"
               If Len(strTmp1) < 15 Then
                  strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
               End If
            End If
            If chk.Tag <> "" Then
               If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
                  strTmp2 = Mid(chk.Tag, 2)
               Else
                  strTmp2 = Trim(chk.Tag)
               End If
               If strTmp2 <> "" Then
                  strTmp2 = m_CaseNo & strTmp2
               End If
            End If
            strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
         End If
      Next
   End If
   '2019/10/31 END
   'PPH審查Tab
   If SSTab1.TabVisible(2) = True And ET03 = "18" Then 'PPH審查
      For Each chk In chk2Tab2
         If chk.Value = 1 Then
            strTmp1 = "": strTmp2 = ""
            strTmp1 = chk.Caption
            If strTmp = "" Then
               strTmp1 = "　【" & strTmp1 & "】"
               If Len(strTmp1) < 14 Then
                  strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
               End If
            Else
               strTmp1 = "　　【" & strTmp1 & "】"
               If Len(strTmp1) < 15 Then
                  strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
               End If
            End If
            If chk.Tag <> "" Then
               If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
                  strTmp2 = Mid(chk.Tag, 2)
               Else
                  strTmp2 = Trim(chk.Tag)
               End If
               If strTmp2 <> "" Then
                  strTmp2 = m_CaseNo & strTmp2
               End If
            End If
            strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
         End If
      Next
'      For Each opt In Option1
'         If opt.Value = True Then
'            strTmp1 = "": strTmp2 = ""
'            strTmp1 = opt.Caption
'            If strTmp = "" Then
'               strTmp1 = "　【" & strTmp1 & "】"
'               If Len(strTmp1) < 14 Then
'                  strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
'               End If
'            Else
'               strTmp1 = "　　【" & strTmp1 & "】"
'               If Len(strTmp1) < 15 Then
'                  strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
'               End If
'            End If
'            If opt.Tag <> "" Then
'               If Mid(opt.Tag, 1, 1) = "1" Or Mid(opt.Tag, 1, 1) = "2" Or Mid(opt.Tag, 1, 1) = "3" Then
'                  strTmp2 = Mid(opt.Tag, 2)
'               Else
'                  strTmp2 = Trim(opt.Tag)
'               End If
'               If strTmp2 <> "" Then
'                  strTmp2 = m_CaseNo & strTmp2
'               End If
'            End If
'            strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
'         End If
'      Next
   End If
   '其他Tab
   'Modify By Sindy 2025/9/25 +Or ET03 = "27"
   If SSTab1.TabVisible(3) = True And (ET03 = "19" Or ET03 = "27") Then '加速審查
      '加速審查
      If Frame422.Enabled = True Then
         For Each chk In chk3Tab3
            If chk.Value = 1 Then
               strTmp1 = "": strTmp2 = ""
               strTmp1 = chk.Caption
               If strTmp = "" Then
                  strTmp1 = "　【" & strTmp1 & "】"
                  If Len(strTmp1) < 14 Then
                     strTmp1 = strTmp1 & String(14 - Len(strTmp1), "　")
                  End If
               Else
                  strTmp1 = "　　【" & strTmp1 & "】"
                  If Len(strTmp1) < 15 Then
                     strTmp1 = strTmp1 & String(15 - Len(strTmp1), "　")
                  End If
               End If
               If chk.Tag <> "" Then
                  If Mid(chk.Tag, 1, 1) = "1" Or Mid(chk.Tag, 1, 1) = "2" Or Mid(chk.Tag, 1, 1) = "3" Then
                     strTmp2 = Mid(chk.Tag, 2)
                  Else
                     strTmp2 = Trim(chk.Tag)
                  End If
                  If strTmp2 <> "" Then
                     strTmp2 = m_CaseNo & strTmp2
                  End If
               End If
               strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & strTmp2
            End If
         Next
         'Add By Sindy 2018/4/12
         For Each opt In Opt2Tab3
            If opt.Value = True Then
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & opt.Caption & "','♀')"
            End If
         Next
         '2018/4/12 END
      End If
   End If
   'Add By Sindy 2019/10/30
   '1603.專利證書(ET03 = "21".誤譯訂正)
   '402.更正
   If (PUB_ChkCPExist(cp, "1603") = True And ET03 = "21") Or cp(10) = "402" Then
      If (PUB_ChkCPExist(cp, "1603") = True And ET03 = "21") Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利申請案已領專利證書','♀')"
      End If
      If pa(8) = "2" Then   '新型時
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','新型適用更正時機','♀')"
         strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "　　【新型專利權有訴訟案件繫屬中之證明文件】" & m_CaseNo & ".ATT.PDF"
      End If
   End If
   '2019/10/30 END
   If strTmp <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附送書件','　" & strTmp & "')"
   End If
   '*******************************************************************************
   
   'Add By Sindy 2019/12/12
   If cp(10) = "403" Then '更改
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','因「　　　」原因，申請/聲明「　　　    」。')"
   'Add By Sindy 2022/4/13
   ElseIf cp(10) = "239" Then '擇一申復
      Dim tmpCM1(1 To 4) As String '發明案號
      Dim tmpCM2(1 To 4) As String '新型案號
      tmpCM1(1) = pa(1): tmpCM1(2) = pa(2): tmpCM1(3) = pa(3): tmpCM1(4) = pa(4)
      '抓新型案號
      Call PUB_IsDualApply(tmpCM1, tmpCM2, , , strTmp)
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','新型申請案號','" & strTmp & "')"
   'Add By Sindy 2023/3/22 審查意見申復申請書
   ElseIf cp(10) = "205" And ET03 = "20" And (chkAtt3(1).Value = 1 Or chkAtt3(3).Value = 1 Or chkAtt3(5).Value = 1) Then
      strTmp = "1.復貴局來函，為本案提出申復，並檢附申復書。" & vbCrLf & _
               "2.一併變更地址如基本資料上所述。"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','" & strTmp & "')"
   'Add By Sindy 2024/8/21
   ElseIf cp(10) = "421" Then '申請技術報告
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請求項','" & txtCP136 & "')" 'txtAddItem ex:FCP-71335
      '相關事項
      If chkAtt421_1(0).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利權已當然消滅','♀')"
      End If
      If chkAtt421_1(1).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利權人報告事由','♀')"
      End If
      If chkAtt421_1(2).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','非專利權人報告事由','♀')"
      End If
      If chkAtt421_1(3).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','辦理事項聲明','♀')"
      End If
      '同時辦理事項
      strExc(10) = ""
      If chkAtt421_2(0).Value = 1 Then strExc(10) = IIf(strExc(10) <> "", strExc(10) & Chr(13) & Chr(10), "") & "　　" & chkAtt421_2(0).Caption
      If chkAtt421_2(1).Value = 1 Then strExc(10) = IIf(strExc(10) <> "", strExc(10) & Chr(13) & Chr(10), "") & "　　" & chkAtt421_2(1).Caption
      If chkAtt421_2(2).Value = 1 Then strExc(10) = IIf(strExc(10) <> "", strExc(10) & Chr(13) & Chr(10), "") & "　　" & chkAtt421_2(2).Caption
      If chkAtt421_2(3).Value = 1 Then strExc(10) = IIf(strExc(10) <> "", strExc(10) & Chr(13) & Chr(10), "") & "　　" & chkAtt421_2(3).Caption
      If strExc(10) <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','同時辦理事項','" & strExc(10) & "')"
      End If
      '附送書件
      If chkAtt421_3(0).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-商業實施證明文件','" & m_CaseNo & ".')"
      End If
      If chkAtt421_3(1).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-涉及專利侵權爭議證明文件','" & m_CaseNo & ".')"
      End If
      If chkAtt421_3(2).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-委任書','" & m_CaseNo & ".')"
      End If
      If pa(22) = "" Then
         '無證書號者再帶出
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','本案已另函提出領證申請（尚未公告的案件）')"
      End If
   End If
   '2019/12/12 END
   
   '附件-基本資料表
   'Add By Sindy 2019/1/15 高速審查的修正不顯示附送書件
   'If cp(10) = "431" And bolAttachments = False Then
   If bolAttachments = False Then
      '不顯示附送書件
   Else
   '2019/1/15 END
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & chkAtt(0).Tag & "')"
   End If
   'If cp(10) <> "416" Then
   '實體審查申請書不顯示其他
   '高速審查的修正不顯示其他
   'Modify By Sindy 2019/1/31 補正文件申請書不顯示其他 ex:FCP-057881 => + Not (cp(10) = "107" And ET03 = "02")
'   If Not (cp(10) = "416" And (ET03 = "03" Or bolAttachments = False)) And _
'      Not (cp(10) = "431" And bolAttachments = False) And _
'      Not (cp(10) = "107" And (ET03 = "02" Or bolAttachments = False)) Then
   'Modify By Sindy 2023/3/14 + And Val(txtDecreasePageFee) = 0
   If bolAttachments = True And _
      Not (cp(10) = "431" And Val(txtDecreaseItemFee) = 0 And Val(txtDecreasePageFee) = 0) Then
      
      'Modify By Sindy 2019/11/26 再審查若有含修正時,才需要出現其他
      If cp(10) = "107" Then
         If Check107_1(0).Value = 0 And Check107_2(0).Value = 0 Then
            GoTo ExitShowOther
         End If
      End If
      '2019/11/26 END
      
      '附件-其他
      'If chkAtt(2).Value = 1 Then
      'Modify By Sindy 2019/4/22 Mark
      'Add By Sindy 2019/4/2 高速審查有退費時才顯示其他
'      If cp(10) = "431" And Val(txtDecreaseItemFee) > 0 Then
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他','♀')"
'      ElseIf cp(10) <> "431" Then
      '2019/4/2 END
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他','♀')"
'      End If
      'End If
      
      'Add By Sindy 2019/12/12 403=更改
      'Modify By Sindy 2024/5/28 + 補文件
      If cp(10) = "403" Or cp(10) = 補文件 Then '更改
         'Add By Sindy 2024/5/28
         If cp(10) = 補文件 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-優惠期證明文件','♀')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-申復書','♀')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-再審查理由書','♀')"
         End If
         '2024/5/28 END
         
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件描述','♀')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件檔名','" & m_CaseNo & ".ATT.PDF')"
      Else
         'If chkAtt2(0).Value = 1 Then
         'Modify By Sindy 2019/4/22 FCP-60518
         'Modify By Sindy 2023/3/14 + Or Val(txtDecreasePageFee) > 0
         If Val(txtDecreaseItemFee) > 0 Or Val(txtDecreasePageFee) > 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件描述','辦理退費之電子收據')"
         Else
         '2019/4/22 END
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件描述','英文說明書、申請專利範圍、及摘要修正替換本')"
         End If
         'If chkAtt2(1).Value = 1 Then
         'Modify By Sindy 2019/4/22
         'Modify By Sindy 2023/3/14 + Or Val(txtDecreasePageFee) > 0
         If Val(txtDecreaseItemFee) > 0 Or Val(txtDecreasePageFee) > 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件檔名','" & m_CaseNo & ".ATT.PDF')"
         Else
         '2019/4/22 END
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件檔名','" & m_CaseNo & ".FIX.ORI.pdf')"
         End If
      End If
ExitShowOther:
   End If
   
   'Add By Sindy 2018/1/25
   If chkAtt(1).Value = 1 Or cp(10) = "205" Then '205申復
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-申復書','" & m_CaseNo & chkAtt(1).Tag & "')"
   End If
   '2018/1/25 END
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

'************************************************
' 儲存專利案件資料
'
'************************************************
Private Function FormSave() As Boolean
Dim ii As Integer, strUpdDA As String
Dim strUpdDA_Spec As String, strCP09_Spec As String 'Add By Sindy 2020/2/20
Dim strCP10Whe As String 'Add By Sindy 2022/3/4
'Add By Sindy 2025/10/9
Dim bolChk As Boolean
Dim chk As CheckBox
'2025/10/9 END
   
On Error GoTo CheckingErr
   
   cnnConnection.BeginTrans
   
   '頁數總計
   'Modify By Sindy 2018/5/8 + Or cp(10) = "307"
   'Memo by Morgan 2022/5/12 +435 續行母案再審
   If SSTab1.TabVisible(0) = True Then '分割/實審
      If Val(txtDocCh(4)) > 0 Then
         cp(135) = txtDocCh(4).Text
         strUpdDA = strUpdDA & ",cp135=" & CNULL(cp(135), True)
      End If
      If Val(txtDocCh(5)) > 0 Then
         cp(136) = txtDocCh(5).Text
         strUpdDA = strUpdDA & ",cp136=" & CNULL(cp(136), True)
      End If
      'Add By Sindy 2019/5/8
      '附英文摘要
      If Check3.Value = 1 And cp(10) = "307" Then '分割
         strUpdDA = strUpdDA & ",cp64='未附英文說明書，所檢附之申請書及摘要附有英文翻譯，可減收申請規費800元;" & cp(64) & "'"
      Else
      '2019/5/8 END
         'Add By Sindy 2019/5/2 序列表 加註進度備註
         If Val(txtDocCh(7)) > 0 And (InStr(cp(64), "序列表" & Val(txtDocCh(7)) & "頁不納入超頁費之計算") = 0) Then
            strUpdDA = strUpdDA & ",cp64='" & "序列表" & Val(txtDocCh(7)) & "頁不納入超頁費之計算;" & cp(64) & "'"
         End If
         '2019/5/2 END
      End If
      
      'Add By Sindy 2020/3/10
      If cp(10) = "307" And FramePA158.Visible = True Then '分割
         If Combo3.Tag <> Combo3.Text Then
            strSql = "UPDATE patent SET pa158='" & Left(Trim(Combo3.Text), 1) & "'" & _
                     " WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'"
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strSql
            Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strSql
         End If
      End If
      '2020/3/10 END
   'Modify By Sindy 2021/3/23 +107.再審申請才需要回存
   ElseIf SSTab1.TabVisible(3) = True And cp(10) = "107" Then '其他頁籤
      If Val(txtDocCh2(4)) > 0 Then
         cp(135) = txtDocCh2(4).Text
         strUpdDA = strUpdDA & ",cp135=" & CNULL(cp(135), True)
      End If
      If Val(txtDocCh2(5)) > 0 Then
         cp(136) = txtDocCh2(5).Text
         strUpdDA = strUpdDA & ",cp136=" & CNULL(cp(136), True)
      End If
      'Add By Sindy 2019/5/2 序列表 加註進度備註
      If Val(txtDocCh2(7)) > 0 And (InStr(cp(64), "序列表" & Val(txtDocCh2(7)) & "頁不納入超頁費之計算") = 0) Then
         strUpdDA = strUpdDA & ",cp64='" & "序列表" & Val(txtDocCh2(7)) & "頁不納入超頁費之計算;" & cp(64) & "'"
      End If
      '2019/5/2 END
   End If
   'Add By Sindy 2019/2/18 剔除實審和再審,同時修正時只要記錄上列程式段最後的頁/項數即可
   'Modified by Morgan 2022/5/9 +剔除 435 續行母案再審
   If cp(10) <> "416" And cp(10) <> "107" And cp(10) <> "435" Then
   '2019/2/18 END
      'Add By Sindy 2020/2/20 431高速審查有勾選修正事項且增、減項次有數值，
      '                       則請將增、減項次的數值回存到未發文之203主動修正，不存在高速審查
      'Modify By Sindy 2020/12/10 因發現有工程師未勾選修正事項，故請改程式為若高速審查增、減項次有數值，
      '                           則請一律將增、減項次的數值回存到未發文之主動修正，不存在高速審查
      'If cp(10) = "431" And (Check431_1(0).Value = 1 Or Check431_2(0).Value = 1) Then
      'Modify By Sindy 2022/3/4 + m_433Fee = "Y"
      If cp(10) = "431" Or m_433Fee = "Y" Then
      '2020/12/10 END
         If m_433Fee = "Y" Then
            strCP10Whe = " AND CP10 in('203','204','205','107')"
         Else
            strCP10Whe = " AND CP10='203'"
         End If
         strExc(0) = "SELECT cp09 FROM caseprogress" & _
                     " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                     strCP10Whe & " AND CP158=0 AND CP159=0" & _
                     " ORDER BY CP05 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strCP09_Spec = RsTemp.Fields("cp09")
         End If
      End If
      '2020/2/20 END
      'Update修正頁/項數:
      'Modify By Sindy 2023/4/27 +if
      If Me.Visible = True Then
      '2023/4/27 END
         If txtAddItem.Enabled = True And FrameFee.Enabled = True Then
            '增加項數
            'Modify By Sindy 2019/1/16
            cp(136) = IIf(Val(txtAddItem) > 0, Val(txtAddItem), "")
            'Add By Sindy 2020/2/20
            If strCP09_Spec <> "" Then
               strUpdDA_Spec = strUpdDA_Spec & ",cp136=" & CNULL(cp(136), True)
            Else
            '2020/2/20 END
               strUpdDA = strUpdDA & ",cp136=" & CNULL(cp(136), True)
            End If
            '2019/1/16 END
            
            '刪除未審項數
            cp(137) = IIf(Val(txtCP137) > 0, Val(txtCP137), "")
            'Add By Sindy 2020/2/20
            If strCP09_Spec <> "" Then
               strUpdDA_Spec = strUpdDA_Spec & ",cp137=" & CNULL(cp(137), True)
            Else
            '2020/2/20 END
               strUpdDA = strUpdDA & ",cp137=" & CNULL(cp(137), True)
            End If
            
            '刪除已審項數
            cp(138) = IIf(Val(txtCP138) > 0, Val(txtCP138), "")
            'Add By Sindy 2020/2/20
            If strCP09_Spec <> "" Then
               strUpdDA_Spec = strUpdDA_Spec & ",cp138=" & CNULL(cp(138), True)
            Else
            '2020/2/20 END
               strUpdDA = strUpdDA & ",cp138=" & CNULL(cp(138), True)
            End If
            
            'Add By Sindy 2023/3/13
            '增加頁數
            cp(135) = IIf(Val(txtAddPage) > 0, Val(txtAddPage), "")
            If strCP09_Spec <> "" Then
               strUpdDA_Spec = strUpdDA_Spec & ",cp135=" & CNULL(cp(135), True)
            Else
               strUpdDA = strUpdDA & ",cp135=" & CNULL(cp(135), True)
            End If
            '刪除未審頁數
            cp(167) = IIf(Val(txtCP167) > 0, Val(txtCP167), "")
            If strCP09_Spec <> "" Then
               strUpdDA_Spec = strUpdDA_Spec & ",cp167=" & CNULL(cp(167), True)
            Else
               strUpdDA = strUpdDA & ",cp167=" & CNULL(cp(167), True)
            End If
            '刪除已審頁數
            cp(168) = IIf(Val(txtCP168) > 0, Val(txtCP168), "")
            If strCP09_Spec <> "" Then
               strUpdDA_Spec = strUpdDA_Spec & ",cp168=" & CNULL(cp(168), True)
            Else
               strUpdDA = strUpdDA & ",cp168=" & CNULL(cp(168), True)
            End If
            '更新專利說明書頁數明細
            'If Val(cp(135)) + Val(cp(167)) + Val(cp(168)) > 0 Then ==> 都新增,此記錄為發文時辨識是否要回寫基本檔
            'Modify By Sindy 2023/12/4 檢查資料是否已存在
            strExc(0) = "SELECT * FROM pagedetail WHERE pd01='" & IIf(strCP09_Spec <> "", strCP09_Spec, strReceiveNo) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            'If pageD(1) = "" Then
            If intI = 0 Then
            '2023/12/4 END
               strSql = "INSERT INTO pagedetail(pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10,pd11,pd12,pd13,pd21)" & _
                        " VALUES('" & IIf(strCP09_Spec <> "", strCP09_Spec, strReceiveNo) & "'" & _
                        "," & CNULL(txtDocAdd(0), True) & "," & CNULL(txtDocAdd(1), True) & "," & CNULL(txtDocAdd(3), True) & "," & CNULL(txtDocAdd(4), True) & _
                        "," & CNULL(txtDocCp167(0), True) & "," & CNULL(txtDocCp167(1), True) & "," & CNULL(txtDocCp167(3), True) & "," & CNULL(txtDocCp167(4), True) & _
                        "," & CNULL(txtDocCp168(0), True) & "," & CNULL(txtDocCp168(1), True) & "," & CNULL(txtDocCp168(3), True) & "," & CNULL(txtDocCp168(4), True) & _
                        "," & CNULL(txtDocCh4(7), True) & ")"
               Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
               cnnConnection.Execute strSql
            Else
               strExc(10) = ""
               If txtDocAdd(0) <> pageD(2) Then strExc(10) = strExc(10) & ",pd02=" & CNULL(txtDocAdd(0), True)
               If txtDocAdd(1) <> pageD(3) Then strExc(10) = strExc(10) & ",pd03=" & CNULL(txtDocAdd(1), True)
               If txtDocAdd(3) <> pageD(4) Then strExc(10) = strExc(10) & ",pd04=" & CNULL(txtDocAdd(3), True)
               If txtDocAdd(4) <> pageD(5) Then strExc(10) = strExc(10) & ",pd05=" & CNULL(txtDocAdd(4), True)
               If txtDocCp167(0) <> pageD(6) Then strExc(10) = strExc(10) & ",pd06=" & CNULL(txtDocCp167(0), True)
               If txtDocCp167(1) <> pageD(7) Then strExc(10) = strExc(10) & ",pd07=" & CNULL(txtDocCp167(1), True)
               If txtDocCp167(3) <> pageD(8) Then strExc(10) = strExc(10) & ",pd08=" & CNULL(txtDocCp167(3), True)
               If txtDocCp167(4) <> pageD(9) Then strExc(10) = strExc(10) & ",pd09=" & CNULL(txtDocCp167(4), True)
               If txtDocCp168(0) <> pageD(10) Then strExc(10) = strExc(10) & ",pd10=" & CNULL(txtDocCp168(0), True)
               If txtDocCp168(1) <> pageD(11) Then strExc(10) = strExc(10) & ",pd11=" & CNULL(txtDocCp168(1), True)
               If txtDocCp168(3) <> pageD(12) Then strExc(10) = strExc(10) & ",pd12=" & CNULL(txtDocCp168(3), True)
               If txtDocCp168(4) <> pageD(13) Then strExc(10) = strExc(10) & ",pd13=" & CNULL(txtDocCp168(4), True)
               If txtDocCh4(7) <> pageD(21) Then strExc(10) = strExc(10) & ",pd21=" & CNULL(txtDocCh4(7), True)
               If strExc(10) <> "" Then
                  strExc(10) = Mid(strExc(10), 2)
                  strSql = "UPDATE pagedetail SET " & strExc(10) & _
                           " WHERE pd01='" & IIf(strCP09_Spec <> "", strCP09_Spec, strReceiveNo) & "'"
                  Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
                  cnnConnection.Execute strSql
               End If
            End If
         Else
            strExc(0) = "SELECT * FROM pagedetail WHERE pd01='" & strReceiveNo & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strSql = "DELETE FROM pagedetail WHERE pd01='" & strReceiveNo & "'"
               Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
               cnnConnection.Execute strSql
            End If
         '2023/3/13 END
         End If
      End If
      
      'Add By Sindy 2020/2/20 有項數要儲存主動修正
      If strCP09_Spec <> "" And strUpdDA_Spec <> "" Then
         strUpdDA_Spec = Mid(strUpdDA_Spec, 2)
         strSql = "UPDATE CASEPROGRESS SET " & strUpdDA_Spec & _
                  " WHERE CP09='" & strCP09_Spec & "'"
         'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
         'Pub_SeekTbLog strSql
         Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
         cnnConnection.Execute strSql
      End If
      '2020/2/20 END
   End If
   
'   cp(110) = ""
'   For ii = 0 To lstNameAgent.ListCount - 1
'      If lstNameAgent.Selected(ii) = True Then
'         '員工編號已可非數字需做轉換
'         cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
'      End If
'   Next
'   If cp(110) <> "" Then
'      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
'   End If
   'Add By Sindy 2025/10/9 記錄是否有一併修正、主動修正有併入中說送件
   If cp(10) = "205" Or cp(10) = "107" Then
      bolChk = False
      For Each chk In chk1Tab1
         If chk.Enabled = True And chk.Value = 1 Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = True Then
         strUpdDA = strUpdDA & ",cp148='Y'"
      Else
         strUpdDA = strUpdDA & ",cp148=null"
      End If
   Else
      If m_WriteNote = "Y" Then
         strUpdDA = strUpdDA & ",cp148='Y'"
      ElseIf m_WriteNote = "N" Then
         strUpdDA = strUpdDA & ",cp148=null"
      End If
   End If
   '2025/10/9 END
   strUpdDA = strUpdDA & ",cp84=" & Val(txtCP84) '發文規費
   strUpdDA = strUpdDA & ",cp110='" & cp(110) & "'" '出名代理人
'   'If frm090904.Text6 = "1" Then '1.電子送件
'   If SSTab1.TabVisible(0) = True And Check4.Value = 1 Then
'      strUpdDA = strUpdDA & ",cp118='Y'" '電子送件
'   End If
   
   'Add By Sindy 2018/5/21 先清除此案號總頁/項數,後面SQL會將總頁/項數儲存在此筆文號中
   If cp(10) = "307" Then '分割
      strSql = "UPDATE CASEPROGRESS SET cp135=null,cp136=null" & _
               " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'"
      'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
      'Pub_SeekTbLog strSql 'Add By Sindy 2019/1/15 新增log
      Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
      cnnConnection.Execute strSql
   End If
   '2018/5/21 END
   'If lstNameAgent.Visible = True Then
   If strUpdDA <> "" Then
      strUpdDA = Mid(strUpdDA, 2)
      strSql = " UPDATE CASEPROGRESS SET " & strUpdDA & " WHERE CP09='" & strReceiveNo & "' and nvl(cp27,0)=0"
      'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
      'Pub_SeekTbLog strSql 'Add By Sindy 2019/1/15 新增log
      Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
      cnnConnection.Execute strSql
   End If
   
   'Added by Lydia 2018/12/27 存中文本資訊
   'If SSTab1.TabVisible(0) = True Or (SSTab1.TabVisible(3) = True Or Frame107.Visible = True) Then
   'Modify By Sindy 2023/3/10 Mark,修正的發文才能回寫基本檔 : 拿掉 or SSTab1.TabVisible(5) = True
   If SSTab1.TabVisible(0) = True Or (SSTab1.TabVisible(3) = True And Frame107.Visible = True) Then
        strSql = ""
        'Modify By Sindy 2023/3/10 Mark,修正的發文才能回寫基本檔
'        '修正-變更頁數
'        If SSTab1.TabVisible(5) = True Then
'            For Each oText In txtDocCh3
'               If oText.Tag <> oText.Text Then
'                   Select Case oText.Index
'                        Case 0 '摘要頁數
'                             strSql = strSql & ", PA64=" & CNULL(oText.Text, True)
'                        Case 1 '說明書頁數
'                             strSql = strSql & ", PA65=" & CNULL(oText.Text, True)
'                        Case 2 '序列表頁數
'                             strSql = strSql & ", PA66=" & CNULL(oText.Text, True)
'                        Case 3 '申請專利範圍頁數
'                             strSql = strSql & ", PA67=" & CNULL(oText.Text, True)
'                        Case 4 '圖式頁數
'                             strSql = strSql & ", PA68=" & CNULL(oText.Text, True)
'                        'Added by Lydia 2019/01/10
'                        Case 7 '6 '圖式圖數 Modify By Sindy 2019/8/29 6=>7
'                             strSql = strSql & ", PA173=" & CNULL(oText.Text, True)
'                   End Select
'               End If
'            Next
'        Else
            For Each oText In IIf(SSTab1.TabVisible(0) = True, txtDocCh, txtDocCh2)
               'Modified by Lydia 2019/01/10 +6
               'If (oText.Index <= 3 Or oText.Index = 7 Or oText.Index = 6) And oText.Tag <> oText.Text Then
               If (oText.Index <> 4) And oText.Tag <> oText.Text Then
                   Select Case oText.Index
                        Case 0 '摘要頁數
                             strSql = strSql & ", PA64=" & CNULL(oText.Text, True)
                        Case 1 '說明書頁數
                             strSql = strSql & ", PA65=" & CNULL(oText.Text, True)
                        Case 7 '序列表頁數
                             strSql = strSql & ", PA66=" & CNULL(oText.Text, True)
                        Case 2 '申請專利範圍頁數
                             strSql = strSql & ", PA67=" & CNULL(oText.Text, True)
                        Case 3 '圖式頁數
                             strSql = strSql & ", PA68=" & CNULL(oText.Text, True)
                        'Add By Sindy 2023/3/15 申請專利範圍項數(最初項數)
                        Case 5
                             strSql = strSql & ", PA172=" & CNULL(oText.Text, True)
                        'Added by Lydia 2019/01/10
                        Case 6 '圖式圖數
                             strSql = strSql & ", PA173=" & CNULL(oText.Text, True)
                   End Select
               End If
            Next
'        End If
        If strSql <> "" Then
            strSql = "UPDATE PATENT SET " & Mid(strSql, 2) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            strSql = "begin user_data.user_enabled:=0; " & strSql & "; end;"
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strSql '新增log
            Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strSql
        End If
   End If
   'end 2018/12/27
   
   'Added by Morgan 2019/10/7
   If m_PA162 <> pa(162) Then
      strSql = "UPDATE patent SET pa162='" & m_PA162 & "'" & _
               " WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'"
      'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
      'Pub_SeekTbLog strSql
      Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
      cnnConnection.Execute strSql
   End If
   'end 2019/10/7
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function

CheckingErr:
   cnnConnection.RollbackTrans

End Function

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim i As Integer
   
   Cancel = False
   lstNameAgent_Validate Cancel
   If Cancel = True Then
      If lstNameAgent.Visible = True And lstNameAgent.Enabled = True Then lstNameAgent.SetFocus
      Exit Function
   End If
   
   Select Case cp(10)
      Case "307"
         If txtFavDate <> "" Then
            txtFavDate_Validate Cancel
            If Cancel = True Then
               txtFavDate.SetFocus
               Exit Function
            Else
               If cboFavReason.ListIndex < 0 Then
                  MsgBox "請點選優惠期發生原因！", vbInformation
                  cboFavReason.SetFocus
                  Exit Function
               End If
            End If
         End If
      
         If chkDoc(0).Value + chkDoc(1).Value + chkDoc(2).Value = 0 Then
            MsgBox "外文本、中文本或簡體字本資訊至少要選擇一種！", vbCritical
            Exit Function
         End If
      
         If chkDoc(1).Value = vbChecked Then
            If Val(txtForeign) = 0 Then
               MsgBox "請輸入外文頁數總計！", vbInformation
               txtForeign.SetFocus
               Exit Function
            End If
            If cboLagnuage.ListIndex < 0 Then
               MsgBox "請點選外文本種類！", vbInformation
               cboLagnuage.SetFocus
               Exit Function
            End If
         End If
      
         If chkDoc(0).Value = vbChecked Then
            If Val(txtDocCh(0)) = 0 And txtDocCh(0).Enabled = True Then
               MsgBox "請輸入摘要頁數！", vbInformation
               txtDocCh(0).SetFocus
               Exit Function
            End If
            If Val(txtDocCh(1)) = 0 And txtDocCh(1).Enabled = True Then
               MsgBox "請輸入說明書頁數！", vbInformation
               txtDocCh(1).SetFocus
               Exit Function
            End If
            If Val(txtDocCh(2)) = 0 And txtDocCh(2).Enabled = True Then
               MsgBox "請輸入申請專利範圍頁數！", vbInformation
               txtDocCh(2).SetFocus
               Exit Function
            End If
            If Val(txtDocCh(5)) = 0 And txtDocCh(5).Enabled = True Then
               MsgBox "請輸入申請專利範圍項數！", vbInformation
               txtDocCh(5).SetFocus
               Exit Function
            End If
         End If
      
         If chkDoc(2).Value = vbChecked Then
            If Val(txtSimplified) = 0 Then
               MsgBox "請輸入簡體字頁數總計！", vbInformation
               txtSimplified.SetFocus
               Exit Function
            End If
         End If
         
         'Add By Sindy 2020/3/10
         If FramePA158.Visible = True Then
            If Trim(Combo3.Text) = "" Then Combo3 = m_DivAppPA158 + "." + PUB_GetCaseAttributeName(m_DivAppPA158, pa(8))
            If Combo3.Tag = "" And Left(Trim(Combo3.Text), 1) = m_DivAppPA158 Then
               If MsgBox("請確認案件屬性是否為「預設屬性」？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                  Combo3.SetFocus
                  Exit Function
               End If
            End If
         End If
         '2020/3/10 END
   End Select
   
   'Add By Sindy 2019/8/29 有”變更頁數”頁籤時,檢查資料有無輸入或異動
   '                       分割,實審,再審查亦須檢查
'   If SSTab1.TabVisible(5) = True And _
'      Not ((cp(10) = "205" Or cp(10) = "431") And Frame204.Tag = "") Then
'      If txtDocCh3(0).Tag = txtDocCh3(0).Text And txtDocCh3(1).Tag = txtDocCh3(1).Text And _
'         txtDocCh3(3).Tag = txtDocCh3(3).Text And _
'         txtDocCh3(4).Tag = txtDocCh3(4).Text Then
'         If MsgBox("中文本資訊頁數沒有異動，是否要修改？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
'            SSTab1.Tab = 5
'            If txtDocCh3(0).Enabled = True Then txtDocCh3(0).SetFocus
'            Exit Function
'         End If
'      End If
'   Else
   'Add By Sindy 2023/4/27 +if ex:FCP-56167(更改)
   If Me.Visible = True Then
   '2023/4/27 END
      If SSTab1.TabVisible(0) = True Then
         If Val(txtDocCh(0).Text) = 0 And Val(txtDocCh(1).Text) = 0 And _
            Val(txtDocCh(7).Text) = 0 And Val(txtDocCh(2).Text) = 0 And _
            Val(txtDocCh(3).Text) = 0 And Val(txtDocCh(6).Text) = 0 And Val(txtDocCh(5).Text) = 0 Then
            MsgBox "最終頁數、項數未填請補填！", vbExclamation
            SSTab1.Tab = 0
            chkDoc(0).Value = 1
            If txtDocCh(0).Enabled = True Then
               If txtDocCh(0).Enabled = True Then txtDocCh(0).SetFocus
            End If
            Exit Function
         ElseIf txtDocCh(0).Tag = txtDocCh(0).Text And txtDocCh(1).Tag = txtDocCh(1).Text And _
               txtDocCh(7).Tag = txtDocCh(7).Text And txtDocCh(2).Tag = txtDocCh(2).Text And _
               txtDocCh(3).Tag = txtDocCh(3).Text And txtDocCh(6).Tag = txtDocCh(6).Text Then
            If MsgBox("中文本資訊頁數沒有異動，是否要修改？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
               SSTab1.Tab = 0
               chkDoc(0).Value = 1
               If txtDocCh(0).Enabled = True Then txtDocCh(0).SetFocus
               Exit Function
            End If
         End If
      ElseIf SSTab1.TabVisible(3) = True And cp(10) = "107" Then
         If Val(txtDocCh2(0).Text) = 0 And Val(txtDocCh2(1).Text) = 0 And _
            Val(txtDocCh2(7).Text) = 0 And Val(txtDocCh2(2).Text) = 0 And _
            Val(txtDocCh2(3).Text) = 0 And Val(txtDocCh2(6).Text) = 0 And Val(txtDocCh2(5).Text) = 0 Then
            MsgBox "最終頁數、項數未填請補填！", vbExclamation
            SSTab1.Tab = 3
            If txtDocCh2(0).Enabled = True Then txtDocCh2(0).SetFocus
            Exit Function
         ElseIf txtDocCh2(0).Tag = txtDocCh2(0).Text And txtDocCh2(1).Tag = txtDocCh2(1).Text And _
               txtDocCh2(7).Tag = txtDocCh2(7).Text And txtDocCh2(2).Tag = txtDocCh2(2).Text And _
               txtDocCh2(3).Tag = txtDocCh2(3).Text And txtDocCh2(6).Tag = txtDocCh2(6).Text Then
            If MsgBox("中文本資訊頁數沒有異動，是否要修改？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
               SSTab1.Tab = 3
               If txtDocCh2(0).Enabled = True Then txtDocCh2(0).SetFocus
               Exit Function
            End If
         End If
      End If
      '2019/8/29 END
   End If
   
   'Add by Morgan 2010/1/6
   m_lngOverPageFee = 0
   m_lngOverItemFee = 0
   m_FeeMemo = ""
   If m_bolChkFee Then
      '檢查條件
      'Modify By Sindy 2023/3/23 +, txtCP167, txtCP168
      If Not PUB_CheckOfficialFee_P(cp(), False, m_bolChkItem, _
                                    txtCP135, txtCP136, txtCP137, txtCP138, txtCP84, _
                                    m_lngRecOverPageFee, m_lngRecOverItemFee, m_FeeMemo, _
                                    m_lngOverPageFee, m_lngOverItemFee, _
                                    m_lngOverPageFeeDiff, m_lngOverItemFeeDiff, txtCP167, txtCP168) Then
         Exit Function
      End If
   End If
   'end 2010/1/6
   Call_PUB_SetOfficialFee_P 'Add By Sindy 2018/6/13
   
   'Modify By Sindy 2018/7/27 附英文摘要可以減免800元
   If Label800.Visible = True Then
      txtCP84 = Val(txtCP84) - 800
      txtCP84.Tag = txtCP84.Text
   End If
   'Add By Sindy 2023/5/26 若工程師在產生加速審查申請書時，若是勾選：
   '1.以商業上之實施所必要
   '2.為綠色技術相關案件:申請書的繳費金額帶4000，回存進度檔的發文規費
   If Opt2Tab3(2).Value = True Or Opt2Tab3(3).Value = True Then
      txtCP84 = 4000
      txtCP84.Tag = txtCP84.Text
   End If
   '2023/5/26 END
   
   'Added by Morgan 2019/10/7
   '發明/新型的 1. 主動修正 2. 修正 3. 申復 4. 再審查  要確認是否加註核准分割建議(提申前主動修正不會有申請書故不需如發文另加判斷)
   If (pa(8) = "1" Or pa(8) = "2") And (cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "107") Then
      If m_PA162 = "" Then
         intI = MsgBox("是否加註核准分割建議？", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
         If intI = vbYes Then
            m_PA162 = "Y"
         ElseIf intI = vbNo Then
            m_PA162 = "N"
         Else
            Exit Function
         End If
      ElseIf m_PA162 = "Y" Then
         intI = MsgBox("目前是否加註核准分割建議為""是""，是否更改？", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
         If intI = vbYes Then
            m_PA162 = "N"
         ElseIf intI = vbCancel Then
            Exit Function
         End If
      ElseIf m_PA162 = "N" Then
         intI = MsgBox("目前是否加註核准分割建議為""否""，是否更改？", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
         If intI = vbYes Then
            m_PA162 = "Y"
         ElseIf intI = vbCancel Then
            Exit Function
         End If
      End If
   End If
   'end 2019/10/7
   
   TxtValidate = True
End Function

'是否一併提實審
Private Sub Check2_Validate(Cancel As Boolean)
   Call_PUB_SetOfficialFee_P 'Add By Sindy 2018/6/13
End Sub

'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   cp(110) = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      'Modify By Sindy 2025/10/14 ex:FCP-074558實體審查
      'MsgBox "出名代理人不可空白！", vbExclamation
      MsgBox "此程序尚無出名代理人，請於程序人員確認！", vbExclamation
   Else
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

Private Sub txtAddItem_GotFocus()
   TextInverse txtAddItem
   CloseIme
End Sub

Private Sub txtAddItem_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtAddItem_Validate(Cancel As Boolean)
   'Add By Sindy 2019/2/18 剔除實審和再審
   'If cp(10) <> "416" And cp(10) <> "107" And cp(10) <> "435" Then
   If FrameFee.Enabled = True Then
   '2019/2/18 END
      Call_PUB_SetOfficialFee_P
   End If
End Sub

Private Sub txtCP135_GotFocus()
   TextInverse txtCP135
   CloseIme
End Sub

Private Sub txtCP135_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP135_Validate(Cancel As Boolean)
   Call_PUB_SetOfficialFee_P 'Add By Sindy 2018/6/13
End Sub

Private Sub txtCP136_GotFocus()
   TextInverse txtCP136
   CloseIme
End Sub

Private Sub txtCP136_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP136_Validate(Cancel As Boolean)
   Call_PUB_SetOfficialFee_P 'Add By Sindy 2018/6/13
End Sub

Private Sub txtCP137_GotFocus()
   TextInverse txtCP137
   CloseIme
End Sub

Private Sub txtCP137_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP137_Validate(Cancel As Boolean)
   'Add By Sindy 2019/2/18 剔除實審和再審
   'If cp(10) <> "416" And cp(10) <> "107" And cp(10) <> "435" Then
   If FrameFee.Enabled = True Then
   '2019/2/18 END
      Call_PUB_SetOfficialFee_P 'Add By Sindy 2018/6/13
   End If
End Sub

Private Sub txtCP138_GotFocus()
   TextInverse txtCP138
   CloseIme
End Sub

Private Sub txtCP138_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP138_Validate(Cancel As Boolean)
   'Add By Sindy 2019/2/18 剔除實審和再審
   'If cp(10) <> "416" And cp(10) <> "107" And cp(10) <> "435" Then
   If FrameFee.Enabled = True Then
   '2019/2/18 END
      Call_PUB_SetOfficialFee_P 'Add By Sindy 2018/6/13
   End If
End Sub

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Sindy 2023/3/14
Private Sub txtDocAdd_GotFocus(Index As Integer)
   TextInverse txtDocAdd(Index)
   CloseIme
End Sub
Private Sub txtDocAdd_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub txtDocAdd_Validate(Index As Integer, Cancel As Boolean)
   If FrameFee.Enabled = True Then
      Call_PUB_SetOfficialFee_P
   End If
End Sub

Private Sub txtDocCh4_GotFocus(Index As Integer)
   TextInverse txtDocCh4(Index)
   CloseIme
End Sub

Private Sub txtDocCh4_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCp167_GotFocus(Index As Integer)
   TextInverse txtDocCp167(Index)
   CloseIme
End Sub
Private Sub txtDocCp167_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub txtDocCp167_Validate(Index As Integer, Cancel As Boolean)
   If FrameFee.Enabled = True Then
      Call_PUB_SetOfficialFee_P
   End If
End Sub
Private Sub txtDocCp168_GotFocus(Index As Integer)
   TextInverse txtDocCp168(Index)
   CloseIme
End Sub
Private Sub txtDocCp168_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub txtDocCp168_Validate(Index As Integer, Cancel As Boolean)
   If FrameFee.Enabled = True Then
      Call_PUB_SetOfficialFee_P
   End If
End Sub
'2023/3/14 END

''Add by Morgan 2004/8/11
'Private Sub txtCP84_Validate(Cancel As Boolean)
'   '台灣
'   If pa(9) = "000" Then
'      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
'         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
'            txtCP84.Tag = txtCP84.Text
'         Else
'            txtCP84_GotFocus
'            Cancel = True
'         End If
'      End If
'   End If
'End Sub

Private Sub txtDocCh_GotFocus(Index As Integer)
   TextInverse txtDocCh(Index)
   CloseIme
End Sub

Private Sub txtDocCh_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCh_Validate(Index As Integer, Cancel As Boolean)
   Dim iChecked As Single

   If Val(txtDocCh(Index)) > 0 Then
      iChecked = vbChecked
   Else
      iChecked = vbUnchecked
   End If

   Select Case Index
   '摘要
   Case 0:
      If chk0Tab(2).Enabled = True And pa(8) = Mid(chk0Tab(2).Tag, 1, 1) Then chk0Tab(2).Value = iChecked
      If chk0Tab(15).Enabled = True And pa(8) = Mid(chk0Tab(15).Tag, 1, 1) Then chk0Tab(15).Value = iChecked
   '說明書
   Case 1:
      If chk0Tab(3).Enabled = True And pa(8) = Mid(chk0Tab(3).Tag, 1, 1) Then chk0Tab(3).Value = iChecked
      If chk0Tab(16).Enabled = True And pa(8) = Mid(chk0Tab(16).Tag, 1, 1) Then chk0Tab(16).Value = iChecked
      If chk0Tab(19).Enabled = True And pa(8) = Mid(chk0Tab(19).Tag, 1, 1) Then chk0Tab(19).Value = iChecked
   '專利範圍
   Case 2:
      If chk0Tab(4).Enabled = True And pa(8) = Mid(chk0Tab(4).Tag, 1, 1) Then chk0Tab(4).Value = iChecked
      If chk0Tab(17).Enabled = True And pa(8) = Mid(chk0Tab(17).Tag, 1, 1) Then chk0Tab(17).Value = iChecked
   '圖式
   Case 3:
      If chk0Tab(5).Enabled = True And pa(8) = Mid(chk0Tab(5).Tag, 1, 1) Then chk0Tab(5).Value = iChecked
      If chk0Tab(18).Enabled = True And pa(8) = Mid(chk0Tab(18).Tag, 1, 1) Then chk0Tab(18).Value = iChecked
      If chk0Tab(20).Enabled = True And pa(8) = Mid(chk0Tab(20).Tag, 1, 1) Then chk0Tab(20).Value = iChecked
   End Select
   
   'Memo by Lydia 2018/12/27 序列表(不算超頁費=不算錢),所以不加入進度檔的總頁數
   If Index <= 3 Then
      txtDocCh(4) = Val(txtDocCh(0)) + Val(txtDocCh(1)) + Val(txtDocCh(2)) + Val(txtDocCh(3))
      txtCP135 = txtDocCh(4) 'Add By Sindy 2018/5/8
      Call txtCP135_Validate(False) 'Add By Sindy 2018/5/8
   ElseIf Index = 5 Then
      'Modify By Sindy 2019/2/12 是否一併主動修正,沒有修正
      'If Check416_1(0).Value = 0 And Check416_2(0).Value = 0 Then
'      'Modify By Sindy 2019/2/13 沒有增刪項數,就以此項數計算規費
'      If Val(txtAddItem) = 0 And Val(txtCP137) = 0 And Val(txtCP138) = 0 Then
'      '2019/2/12 END
         txtCP136 = txtDocCh(5) 'Add By Sindy 2018/5/8
         Call txtCP136_Validate(False) 'Add By Sindy 2018/5/8
'      End If
   End If
End Sub

Private Sub txtFavDate_GotFocus()
   TextInverse txtFavDate
   CloseIme
End Sub

Private Sub txtFavDate_Validate(Cancel As Boolean)
   If txtFavDate <> "" Then
      If ChkDate(txtFavDate) = False Then
         txtFavDate_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub txtDocCh2_GotFocus(Index As Integer)
   TextInverse txtDocCh2(Index)
   CloseIme
End Sub

Private Sub txtDocCh2_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCh2_Validate(Index As Integer, Cancel As Boolean)
   Dim iChecked As Single

   If Val(txtDocCh2(Index)) > 0 Then
      iChecked = vbChecked
   Else
      iChecked = vbUnchecked
   End If
   
   'Memo by Lydia 2018/12/27 序列表(不算超頁費=不算錢),所以不加入進度檔的總頁數
   If Index <= 3 Then
      txtDocCh2(4) = Val(txtDocCh2(0)) + Val(txtDocCh2(1)) + Val(txtDocCh2(2)) + Val(txtDocCh2(3))
      'Modify By Sindy 2018/6/11 Mark:再審申請應該只要輸增修項數 EX:FCP-50420 / FCP-54083
      'txtCP135 = txtDocCh2(4) 'Add By Sindy 2017/12/14
      txtCP135 = txtDocCh2(4) 'Add By Sindy 2018/7/27
      Call txtCP135_Validate(False) 'Add By Sindy 2018/7/27
   'Add By Sindy 2018/8/9 Mark : Jack反應ex:FCP-57374再審申請-有可能之前已有申請項數,這次只是修改
   ElseIf Index = 5 Then
      'Modify By Sindy 2018/6/11 Mark:再審申請應該只要輸增修項數
      'txtCP136 = txtDocCh2(5) 'Add By Sindy 2017/12/14
      'Modify By Sindy 2019/2/12 是否一併主動修正 ex:FCP-049314:只出再審申請書,沒有修正
      'If Check107_1(0).Value = 0 And Check107_2(0).Value = 0 Then
'      'Modify By Sindy 2019/2/13 沒有增刪項數,就以此項數計算規費
'      If Val(txtAddItem) = 0 And Val(txtCP137) = 0 And Val(txtCP138) = 0 Then
'      '2019/2/12 END
         txtCP136 = txtDocCh2(5) 'Add By Sindy 2018/7/27
         Call txtCP136_Validate(False) 'Add By Sindy 2018/7/27
'      End If
   End If
End Sub
'Added by Lydia 2019/01/03
Private Sub txtDocCh3_GotFocus(Index As Integer)
   TextInverse txtDocCh3(Index)
   CloseIme
End Sub

Private Sub txtDocCh3_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCh3_Validate(Index As Integer, Cancel As Boolean)
   '序列表(不算超頁費=不算錢),所以不加入進度檔的總頁數
   'If Index <> 2 And Index <> 5 Then
      'Modified by Morgan 2023/1/9
      'txtDocCh3(5) = Val(txtDocCh3(0)) + Val(txtDocCh3(1)) + Val(txtDocCh3(3)) + Val(txtDocCh3(4))
      'txtCP135 = txtDocCh3(5)
      txtCP135 = Val(txtDocCh3(0)) + Val(txtDocCh3(1)) + Val(txtDocCh3(3)) + Val(txtDocCh3(4))
      'end 2023/1/9
      Call txtCP135_Validate(False)
      txtPage = txtCP135 'Add By Sindy 2023/3/13 讓使用者可以看到頁數合計
   'End If
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
End Sub
