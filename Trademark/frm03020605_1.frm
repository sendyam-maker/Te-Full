VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020605_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-申請, 延展, 補換發證書"
   ClientHeight    =   6036
   ClientLeft      =   72
   ClientTop       =   996
   ClientWidth     =   8664
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6036
   ScaleWidth      =   8664
   Begin VB.TextBox txtTM136 
      Height          =   270
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   68
      Top             =   3720
      Width           =   300
   End
   Begin VB.Frame Frame2 
      Height          =   1845
      Left            =   2940
      TabIndex        =   49
      Top             =   4128
      Width           =   4215
      Begin VB.Label lblC 
         Caption         =   "申請人1:"
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   255
         Width           =   885
      End
      Begin VB.Label lblC 
         Caption         =   "申請人2:"
         Height          =   165
         Index           =   1
         Left            =   120
         TabIndex        =   63
         Top             =   570
         Width           =   885
      End
      Begin VB.Label lblC 
         Caption         =   "申請人3:"
         Height          =   165
         Index           =   2
         Left            =   120
         TabIndex        =   62
         Top             =   892
         Width           =   885
      End
      Begin VB.Label lblC 
         Caption         =   "申請人4:"
         Height          =   165
         Index           =   3
         Left            =   120
         TabIndex        =   61
         Top             =   1214
         Width           =   885
      End
      Begin VB.Label lblC 
         Caption         =   "申請人5:"
         Height          =   165
         Index           =   4
         Left            =   120
         TabIndex        =   60
         Top             =   1538
         Width           =   885
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   0
         Left            =   990
         TabIndex        =   59
         Top             =   180
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1940;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   1
         Left            =   990
         TabIndex        =   58
         Top             =   495
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   2
         Left            =   990
         TabIndex        =   57
         Top             =   825
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   3
         Left            =   990
         TabIndex        =   56
         Top             =   1140
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   4
         Left            =   990
         TabIndex        =   55
         Top             =   1470
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   0
         Left            =   2100
         TabIndex        =   54
         Top             =   203
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   1
         Left            =   2100
         TabIndex        =   53
         Top             =   525
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   2
         Left            =   2100
         TabIndex        =   52
         Top             =   847
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   3
         Left            =   2100
         TabIndex        =   51
         Top             =   1169
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   4
         Left            =   2100
         TabIndex        =   50
         Top             =   1493
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CheckBox ChkPD 
      Caption         =   "主張優先權"
      Height          =   225
      Left            =   1530
      TabIndex        =   47
      Top             =   3450
      Width           =   1275
   End
   Begin VB.CheckBox ChkPart 
      Caption         =   "部分延展"
      Height          =   225
      Left            =   1620
      TabIndex        =   45
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CheckBox ChkColor 
      Caption         =   "彩色"
      Height          =   210
      Left            =   1530
      TabIndex        =   44
      Top             =   2805
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   28
      Top             =   2460
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "附送書件"
      Height          =   1428
      Left            =   2940
      TabIndex        =   26
      Top             =   2376
      Width           =   2925
      Begin VB.CheckBox chkAtt1 
         Caption         =   "團體商標使用規範書"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   73
         Tag             =   ".ATT.pdf"
         Top             =   2064
         Width           =   2175
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "法人資格證明文件"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   72
         Tag             =   ".ATT.pdf"
         Top             =   1752
         Width           =   2175
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "已取得識別性之具體事證"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   71
         Tag             =   ".ATT.pdf"
         Top             =   1464
         Width           =   2364
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "具結書"
         Height          =   255
         Index           =   5
         Left            =   2328
         TabIndex        =   42
         Tag             =   ".declaration.pdf"
         Top             =   852
         Width           =   1485
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "變更證明文件"
         Height          =   255
         Index           =   4
         Left            =   1056
         TabIndex        =   41
         Tag             =   ".change.pdf"
         Top             =   852
         Width           =   1485
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "展覽會優先權證明文件"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Tag             =   ".PRI.pdf"
         Top             =   1152
         Width           =   2175
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "優先權證明文件"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Tag             =   ".PRI.pdf"
         Top             =   852
         Width           =   1695
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "委任書"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   ".poa.pdf"
         Top             =   546
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "基本資料表"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Tag             =   ".contact.pdf"
         Top             =   240
         Value           =   1  '核取
         Width           =   1215
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1890
      MaxLength       =   7
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7770
      TabIndex        =   2
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5820
      TabIndex        =   0
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6660
      TabIndex        =   1
      Top             =   45
      Width           =   1080
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm03020605_1.frx":0000
      Left            =   1260
      List            =   "frm03020605_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   870
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   7
      Top             =   210
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1575
      MaxLength       =   6
      TabIndex        =   8
      Top             =   210
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2415
      MaxLength       =   1
      TabIndex        =   9
      Top             =   210
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2655
      MaxLength       =   2
      TabIndex        =   10
      Top             =   210
      Width           =   375
   End
   Begin MSForms.CheckBox CheckBox2 
      Height          =   300
      Left            =   210
      TabIndex        =   70
      Top             =   3412
      Width           =   1836
      BackColor       =   -2147483633
      ForeColor       =   16711680
      DisplayStyle    =   4
      Size            =   "3238;529"
      Value           =   "0"
      Caption         =   "變更地址"
      FontName        =   "新細明體"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1:電子 2:紙本"
      Height          =   180
      Index           =   1
      Left            =   1530
      TabIndex        =   69
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "證書形式："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   67
      Top             =   3780
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否要變更申請人:"
      Height          =   180
      Left            =   2970
      TabIndex        =   66
      Top             =   3870
      Width           =   1545
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   312
      Left            =   4536
      TabIndex        =   65
      Top             =   3816
      Width           =   3924
      BackColor       =   -2147483633
      ForeColor       =   255
      DisplayStyle    =   4
      Size            =   "6921;550"
      Value           =   "0"
      Caption         =   "若申請書不需顯示變更申請人，請勿勾選。"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   340
      Left            =   7050
      TabIndex        =   48
      Top             =   2400
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;600"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblPart 
      Caption         =   "延展範圍及內容："
      Height          =   225
      Left            =   210
      TabIndex        =   46
      Top             =   3120
      Width           =   1605
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      Caption         =   "商標圖樣顏色："
      Height          =   210
      Left            =   210
      TabIndex        =   43
      Top             =   2820
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSForms.Label lblData 
      Height          =   252
      Index           =   10
      Left            =   7392
      TabIndex        =   40
      Top             =   532
      Width           =   1068
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "1884;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "商標種類:"
      Height          =   180
      Left            =   6588
      TabIndex        =   39
      Top             =   532
      Width           =   768
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   8600
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   8600
      Y1              =   2325
      Y2              =   2325
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   9
      Left            =   4710
      TabIndex        =   38
      Top             =   1950
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   8
      Left            =   1260
      TabIndex        =   37
      Top             =   1950
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   7
      Left            =   4710
      TabIndex        =   36
      Top             =   1605
      Width           =   3615
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "6376;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   6
      Left            =   1260
      TabIndex        =   35
      Top             =   1590
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   5
      Left            =   4710
      TabIndex        =   34
      Top             =   1262
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   33
      Top             =   1230
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   32
      Top             =   870
      Width           =   6615
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "11668;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   2
      Left            =   4710
      TabIndex        =   31
      Top             =   532
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   1
      Left            =   1260
      TabIndex        =   30
      Top             =   532
      Width           =   1665
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "2937;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   0
      Left            =   4710
      TabIndex        =   29
      Top             =   233
      Width           =   1035
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "1826;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "規費 :                                      元"
      Height          =   180
      Left            =   216
      TabIndex        =   27
      Top             =   2508
      Width           =   2484
   End
   Begin VB.Label lblNameAgent 
      Caption         =   "出名代理人 :"
      Height          =   180
      Left            =   6000
      TabIndex        =   25
      Top             =   2460
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期 :"
      Height          =   180
      Left            =   870
      TabIndex        =   24
      Top             =   4020
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3810
      TabIndex        =   22
      Top             =   255
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3810
      TabIndex        =   21
      Top             =   1605
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   210
      TabIndex        =   20
      Top             =   1605
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3810
      TabIndex        =   19
      Top             =   1262
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   210
      TabIndex        =   18
      Top             =   1262
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   17
      Top             =   210
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   16
      Top             =   532
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "審定號數:"
      Height          =   180
      Left            =   3810
      TabIndex        =   15
      Top             =   532
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "商標名稱:"
      Height          =   180
      Left            =   210
      TabIndex        =   14
      Top             =   870
      Width           =   765
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   0
      Left            =   3810
      TabIndex        =   13
      Top             =   1950
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   210
      TabIndex        =   12
      Top             =   1950
      Width           =   765
   End
End
Attribute VB_Name = "frm03020605_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2019/03/29 改成Form2.0 (lblData和lstNameAgent)
'Create by Lydia 2019/03/29 各式申請書:申請, 延展, 補換發證書
Option Explicit
Dim tm() As String '商標基本檔
Dim intWhere As Integer, intLastRow As Integer
Dim strReceiveNo As String '收文號
Dim m_CP110 As String, m_AgentName As String  '出名代理人
Dim m_CP10 As String '案件性質
Dim m_CP17 As String '收文規費
Dim m_CP118  As String '是否電子送件'
Dim m_CaseNo As String '電子送件-本所案號
Dim m_F21st07 As String 'FCT程序分機
Dim oObj As Control 'Added by Lydia 2023/11/30
Dim mChkType As String 'Added by Lydia 2023/11/14 確定商標種類：特殊商標TM72>商標種類TM08;
                        '商標註冊種類:要申請的權利主體,商標型態:以立體形狀、聲音等形式呈現，而這些表彰商品或服務來源之標識，為商標法規範之商標「型態」
                        
Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean
Dim i As Integer
Dim strFolder As String, strFileName As String
Dim ET01 As String, ET03 As String
Dim ET03type As String
Dim strChkVal As String
Dim strContent As String 'Added by Lydia 2019/08/08
   
   Select Case Index
      Case 0 '確定
         '檢查101商申註冊的種類
         If m_CP10 = "101" Then
              If m_CP118 = "Y" Then '電子送件申請書
                   'Modified by Lydia 2023/11/14
                   'If InStr("1,7,9,A,C", tm(8)) = 0 And tm(8) <> "" Then
                   'Modified by Lydia 2023/11/30 增加各商標種類的電子送件申請書
                   'If InStr("1,7,9,A,B", mChkType) = 0 Then
                   If InStr("1,7,8,9,A,B,C,D,E,F,G,H,I,J,K", mChkType) = 0 Then
                      MsgBox "該商標種類並無電子送件申請書！", vbCritical
                      Exit Sub
                   End If
              Else   '紙本申請書
                   'Modified by Lydia 2023/11/14
                   'If InStr("1,7,A,C", tm(8)) = 0 And tm(8) <> "" Then
                   If InStr("1,7,A,B", mChkType) = 0 Then
                      MsgBox "該商標種類並無申請書！", vbCritical
                      Exit Sub
                   End If
              End If
         End If
         If m_CP118 = "Y" Then
            bolChk = False
            'Modified by Lydia 2023/11/30 改用For Each
            For Each oObj In chkAtt1
               If oObj.Value = 1 Then
                   bolChk = True
                   Exit For
               End If
            Next
            If bolChk = False Then
               MsgBox "請選擇附送書件 !", vbCritical
               Exit Sub
            End If
         End If
         
         If TxtValidate = False Then Exit Sub
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         If m_CP118 = "Y" Then
            m_CaseNo = PUB_FCPCaseNo2FileName(tm(1), tm(2), tm(3), tm(4))
            '桌面上建立案號資料夾
            strFolder = PUB_Getdesktop
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
                MkDir strFolder
            End If
            
            strLetterDate = Text5.Text
            
            ET01 = "90"
            '1.基本資料 'Memo by Lydia 2019/08/08 移到下方

            ET03 = ""
            If m_CP10 = "101" Then '申請
                'Modified by Lydia 2023/11/14 tm(8)=>mChkType
                Select Case mChkType
                    Case "7" '證明標章
                        ET03 = "21"  'FCT,90,101,01
                    Case "9" '團體商標
                        ET03 = "22"
                    Case "A" '立體商標
                        ET03 = "23"
                    'Modified by Lydia 2023/11/14 "C"=>"B"
                    Case "B" '顏色商標
                        ET03 = "24"
                    'Added by Lydia 2023/11/30 增加各商標種類的電子送件申請書
                    Case "8"  '團體標章
                        ET03 = "28"
                    Case "C"  '聲音商標
                        ET03 = "29"
                    Case "D", "E", "F" '其他商標
                        ET03 = "30"
                    Case "G"  '動態商標
                        ET03 = "31"
                    Case "H"  '全像圖商標
                        ET03 = "32"
                    Case "I"  '立體團體商標
                        ET03 = "33"
                    Case "J"  '顏色團體商標
                        ET03 = "34"
                    Case "K"  '聲音團體商標
                        ET03 = "35"
                    'end 2023/11/30
                    Case Else '商標，含未設定種類
                        ET03 = "20"  'FCT,90,101,20
                End Select
                'Added by Lydia 2023/11/30
                If InStr("D,E,F", mChkType) > 0 Then
                   ET03type = "其他商標註冊"
                Else
                'end 2023/11/30
                   ET03type = lblData(10).Caption & "註冊"
                End If 'Added by Lydia 2023/11/30
            Else
                If m_CP10 = "102" Then '延展
                     ET03 = "25"
                ElseIf m_CP10 = "103" Then '補換發證書
                     ET03 = "26"
                'Added by Lydia 2022/12/19 申請註冊證副本
                ElseIf m_CP10 = "314" Then
                     ET03 = "27"
                'end 2022/12/19
                End If
                ET03type = lblData(0).Caption
            End If
            If ET03 <> "" Then
                   '2.申請書
                   If StartLetter2(ET01, ET03, strReceiveNo, "2") = False Then Exit Sub
                   'Added by Lydia 2019/08/08 判斷要基本資料表,先不存檔
                   If chkAtt1(0).Value = 1 Then
                        NowPrint strReceiveNo, ET01, ET03, False, strUserNum, , , True, strContent
                        strFileName = strFolder & "\" & m_CaseNo & "." & ET03type & "申請書"
                   Else
                   'end 2019/08/08
                        NowPrint strReceiveNo, ET01, ET03, False, strUserNum, , , True, strContent
                        strFileName = strFolder & "\" & m_CaseNo & "." & ET03type & "申請書"
                        Call PUB_MakeDoc(strContent, strFileName)
                   End If
            End If
            'Move by Lydia 2019/08/08 從申請書上面移下來(經外商實測申請案的基本資料表要和申請書在同一份文件,才能轉檔)
            '基本資料
            If chkAtt1(0).Value = 1 Then
                   'Modified by Lydia 2020/12/31 電子送件-基本資料表03=>11
                   ET03 = "11"
                   If StartLetter2(ET01, ET03, strReceiveNo, "1") = False Then Exit Sub
                   'Modified by Lydia 2019/08/08 統一將基本資料表要和申請書放在同一份文件
                   'NowPrint strReceiveNo, ET01, ET03, False, strUserNum, , , True, strContent
                   'strFileName = strFolder & "\" & m_CaseNo & ".contact"
                   'Call PUB_MakeDoc(strContent, strFileName)
                   NowPrint strReceiveNo, ET01, ET03, False, strUserNum, , strContent, True, strContent
                   If strFileName = "" Then strFileName = strFolder & "\" & m_CaseNo & ".contact"
                   'Modified by Lydia 2020/09/25 增加分節處理頁碼
                   'Call PUB_MakeDoc(strContent, strFileName)
                   strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(分節)#|")    '換頁符號Chr(12)替換為分節符號 "|#(分節)#|"
                   Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '分節處理頁碼
                   'end 2019/08/08
                   'end 2020/09/25
            End If
         Else '紙本申請書
            '勾選附件
            If m_CP10 = "101" Then '101申請
                strChkVal = strChkVal & IIf(chkAtt1(1).Value = 1, "A1", "B1") & "," '委任書
                strChkVal = strChkVal & IIf(chkAtt1(2).Value = 1, "A2", "B2") & "," '優先權
                strChkVal = strChkVal & IIf(chkAtt1(3).Value = 1, "A3", "B3") & "," '展覽會優先權
            ElseIf m_CP10 = "102" Then '延展
                strChkVal = strChkVal & IIf(chkAtt1(1).Value = 1, "A1", "B1") & ","  '委任書
                strChkVal = strChkVal & IIf(chkAtt1(4).Value = 1, "A2", "B2") & ","  '變更證明
            ElseIf m_CP10 = "103" Then '補換發證書
                strChkVal = strChkVal & IIf(chkAtt1(1).Value = 1, "A1", "B1") & ","  '委任書
                'strChkVal = strChkVal & IIf(chkAtt1(5).Value = 1, "A2", "B2") & "," '紙本沒有具結書
            End If
            '商標圖樣顏色
            If ChkColor.Visible = True Then
                strChkVal = strChkVal & IIf(ChkColor.Value = 1, "彩色", "墨色") & ","
            End If
            
            '紙本申請書(樣本doc與T台灣案相同)
            Call PUB_GetApplBook(tm(1) & "-" & tm(2) & "-" & tm(3) & "-" & tm(4), m_CP10, , , , , , strReceiveNo, strChkVal)
         End If
         frm030206_1.Show
         '回到原畫面要清除畫面
         frm030206_1.ClearForm
         
      Case 1 '回前畫面
         frm030206_1.Show
         
      Case 2 '結束
         Unload frm030206_1
   End Select
   Unload Me
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         lblData(3) = tm(5)
      Case "英"
         lblData(3) = tm(6)
      Case "日"
         lblData(3) = tm(7)
   End Select
End Sub

Private Sub Form_Load()
Dim tKind As String '特殊申請書

   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm030206_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      tKind = .Text6
      strReceiveNo = .Tag
      If tKind = "2" Then m_CP118 = "Y"
   End With
   ReDim tm(TF_TM)
   ReadTradeMark
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   'Modified by Lydia 2019/08/02 阿蓮:預設出名代理人為林+閻
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110, , True
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True
   'Added by Lydia 2021/04/20 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)

   '電子送件
   If tKind = "2" Then
       m_CP118 = "Y"
   Else
      chkAtt1(0).Enabled = False '紙本不勾選基本資料表
      chkAtt1(0).Value = 0
   End If
   '原本設計電子送件才有勾附件,阿蓮要求紙本也可以勾
    Frame1.Visible = True
    chkAtt1(4).Left = chkAtt1(2).Left
    chkAtt1(5).Left = chkAtt1(3).Left
    '調整顯示附件選項
    'Modified by Lydia 2023/11/30 改用For Each
    For Each oObj In chkAtt1
       If oObj.Index > 1 Then
         oObj.Visible = False
       End If
    Next
    ChkColor.Visible = False: lblColor.Visible = False
    ChkPart.Visible = False:   lblPart.Visible = False 'Added by Lydia 2019/11/05 部分延展
    ChkPD.Visible = False 'Added by Lydia 2020/02/05 主張優先權
    If m_CP10 = "101" Then '申請
        chkAtt1(2).Visible = True
        chkAtt1(3).Visible = True
        'Added by Lydia 2023/11/30
        chkAtt1(6).Visible = True
        If InStr(lblData(10), "證明") > 0 Or InStr(lblData(10), "團體") > 0 Then
           chkAtt1(7).Visible = True
           chkAtt1(8).Visible = True
        End If
        'end 2023/11/30
        ChkColor.Visible = True: lblColor.Visible = True
        'Modified by Lydia 2023/11/14 顏色商標
        'If tm(8) = "C" Then
        'Modified by Lydia 2023/11/30 +J顏色團體商標
        If mChkType = "B" Or mChkType = "J" Then
            ChkColor.Value = 1
        'Added by Lydia 2022/03/08 當案件基本檔之代表圖已放商標圖並點選「彩色」時，請系統自動於商標圖樣顏色欄勾選彩色
        Else
            'Memo by Lydia 2022/03/09 GetPicColor模組是概括彩色圖(IBF05=1 or 2)，如果後面要用IBF05=2只能另外寫判斷；
                  'Amy: 因為ibf05='2'才是真正彩圖Ex: FCT042565, 當初請作單1070827-01 ; FCT案過去有區分整批要不要彩圖，所以有IBF05=1並且IBF06=2的舊制
            strExc(0) = GetPicColor(tm(1), tm(2), tm(3), tm(4))
            If InStr("," & strExc(0), "彩色") > 0 Then
               ChkColor.Value = 1
            End If
        'end 2022/03/08
        End If
        'Added by Lydia 2020/02/05
        ChkPD.Visible = True
        ChkPD.Top = lblPart.Top
    ElseIf m_CP10 = "102" Then '延展
        chkAtt1(4).Visible = True
        Frame2.Left = 2940  'Added by Lydia 2023/11/30
        'Added by Lydia 2019/11/05 電子送件-部分延展
        If m_CP118 = "Y" Then
            ChkPart.Visible = True
            lblPart.Visible = True
        End If
    ElseIf m_CP10 = "103" Then '補換發證書
        chkAtt1(5).Visible = True
    'Added by Lydia 2022/12/19
    ElseIf m_CP10 = "314" Then '申請註冊證副本
        chkAtt1(2).Visible = False
        chkAtt1(3).Visible = False
    End If
    'Added by Lydia 2022/12/28
    If InStr("103,314", m_CP10) > 0 Then
       Label4(0).Visible = True: Label4(1).Visible = True
       txtTM136.Visible = True
    Else
       Label4(0).Visible = False: Label4(1).Visible = False
       txtTM136.Visible = False
    End If
    'end 2022/12/28
    
    'Added by Lydia 2022/04/19 請於「延展」案，增加「是否要變更申請人選項」 (因延展可同時變更申請人)，俾申請書帶出正確之申請人資料。
    Frame2.Enabled = False
    If m_CP10 <> "102" Then
        'Modified by Lydia 2022/12/26
        'Me.Height = 4515
        'Modified by Lydia 2023/11/30
        'Me.Height = 4800
        Me.Height = 5260
        Frame1.Height = 2390
        'end 2023/11/30
        Label10.Visible = False: CheckBox1.Visible = False
        Frame2.Visible = False
        CheckBox2.Visible = False 'Added by Lydia 2023/11/08
    ElseIf m_CP10 = "102" And m_CP118 = "Y" Then
        Me.Height = 6460
        Frame1.Height = 1450 'Added by Lydia 2023/11/30
        Label10.Visible = True: CheckBox1.Visible = True
        Frame2.Visible = True
        Frame2.Enabled = True 'Added by Lydia 2022/07/12 申請人不限制要勾選「變更申請人」才能輸入
        CheckBox2.Visible = True 'Added by Lydia 2023/11/08
    End If
    For intI = 0 To 4
        textFM2(intI).Text = ""
        textFM2(intI).Tag = ""
        lblCName(intI).Caption = ""
    Next intI
    'end 2022/04/19
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm03020605_1 = Nothing
End Sub

Private Sub ReadTradeMark()
Dim rsRd As New ADODB.Recordset

   'Modified by Lydia 2023/11/30 改用oObj
   For Each oObj In lblData
      oObj.Caption = ""
   Next
   
   tm(1) = Text1
   tm(2) = Text2
   tm(3) = Text3
   tm(4) = Text4
   If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
      Text5 = tm(11)
      lblData(1) = tm(12)
      lblData(2) = tm(15)
      lblData(3) = tm(5)
      txtTM136 = tm(136) 'Added by Lydia 2022/12/26
   End If
   '商標種類
   'Added by Lydia 2023/11/30
   If tm(72) <> "" Then
      lblData(10) = PUB_GetSpecialPTName("2", tm(72))
   Else
   'end 2023/11/30
      If ClsPDGetPatentTrademarkKind(商標, tm(8), strExc(0), False) = 1 Then
         lblData(10) = strExc(0)
      End If
   End If 'Added by Lydia 2023/11/30
   
   'Added by Lydia 2023/11/14 確定商標種類：特殊商標TM72>商標種類TM08;
   mChkType = IIf(tm(72) <> "", tm(72), tm(8))
   If mChkType = "" Then mChkType = "1"
   'end 2023/11/14
   'Added by Lydia 2023/11/30 區別選項：法人資格證明文件、團體商標使用規範書
   Select Case mChkType
      Case "7" '證明標章
         chkAtt1(7).Caption = "法人團體機關證明文件"
         chkAtt1(8).Caption = "證明標章使用規範書"
      Case "8"  '團體標章
         chkAtt1(8).Caption = "團體標章使用規範書"
   End Select
   'end 2023/11/30
   
   strExc(0) = "select cpm03,s1.st02 as st1,s2.st02 as st2,cp43,cp10,cp05,cp06,cp07,cp84,cp110,cp64,cp27,cp17,s3.st07  " & _
                    "from caseprogress,casepropertymap,staff s1 ,staff s2,staff s3 " & _
                    "where cp09='" & strReceiveNo & "' " & _
                    "and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) " & _
                    "and cp13=s2.st01(+) and s2.st57=s3.st01(+) "
   intI = 1
   Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       With rsRd
          m_CP110 = "" & .Fields("CP110")
          m_CP10 = "" & .Fields("CP10")
          m_CP17 = Val("" & .Fields("cp17"))
          '收文規費
          If m_CP118 <> "" Then   '電子送件的規費有千分位,會造成轉檔錯誤
               Text9.Text = Val(m_CP17)
          Else
               Text9.Text = Format(Val("" & .Fields("cp17")), "#,##0")
          End If
          
          lblData(0) = "" & .Fields("cpm03") '案件性質
          lblData(4) = "" & .Fields("st1") '承辦人
          lblData(5) = "" & .Fields("st2") '智權人員
          m_F21st07 = "" & .Fields("st07")
          If "" & .Fields("cp43") <> "" Then
             strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP09='" & .Fields("cp43") & "'"
             intI = 1
             Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                lblData(6) = TransDate("" & rsRd.Fields("CP05"), 1) '來函收文日
                lblData(7) = "" & rsRd.Fields("CP08") '機關文號
             End If
          End If
          lblData(8) = TransDate("" & .Fields("cp06"), 1) '本所期限
          lblData(9) = TransDate("" & .Fields("cp07"), 1) '法定期限
       End With
   End If
   
   '101申請: 判斷是否有主張優先權
   If m_CP10 = "101" Then
        If PUB_ChkCPExist(tm, "108") = True Then
             chkAtt1(2).Value = vbChecked
             ChkPD.Value = vbChecked  'Added by Lydia 2020/02/05
        End If
   End If
   
   'FCT向智慧局提出之各式申請書上之分機號碼，請將日本區設定為011國家檔管制人分機
   strExc(0) = "select fa10,st07 from fagent, nation, staff where fa01||fa02='" & ChangeCustomerL(tm(44)) & "' and substr(fa10,1,3)=na01(+) and na55=st01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Left("" & RsTemp.Fields("fa10"), 3) = "011" Then
         m_F21st07 = "" & RsTemp.Fields("st07")
      End If
   End If
   
   Set rsRd = Nothing
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   'Added by Lydia 2022/04/19
   For intI = 0 To 4
       If textFM2(intI).Text <> "" And lblCName(intI).Caption = "" Then
            MsgBox "申請人代碼<" & textFM2(intI).Text & ">不正確", vbExclamation
            textFM2(intI).SetFocus
            textFM2_GotFocus intI
            Exit Function
       End If
   Next intI
   '申請人必須依序輸入
   If textFM2(0) <> "" Or textFM2(1) <> "" Or textFM2(2) <> "" Or textFM2(3) <> "" Or textFM2(4) <> "" Then
      If (Trim(textFM2(1)) <> "" And Trim(textFM2(0)) = "") Or _
         (Trim(textFM2(2)) <> "" And Trim(textFM2(1)) = "") Or _
         (Trim(textFM2(3)) <> "" And Trim(textFM2(2)) = "") Or _
         (Trim(textFM2(4)) <> "" And Trim(textFM2(3)) = "") Then
         MsgBox "請依序輸入！", vbExclamation
         Exit Function
      End If
      If (textFM2(1) <> "" And Trim(textFM2(1)) = Trim(textFM2(0))) Or _
         (textFM2(2) <> "" And Trim(textFM2(2)) = Trim(textFM2(1))) Or _
         (textFM2(3) <> "" And Trim(textFM2(3)) = Trim(textFM2(2))) Or _
         (textFM2(4) <> "" And Trim(textFM2(4)) = Trim(textFM2(3))) Then
         MsgBox "資料重覆！", vbExclamation
         Exit Function
      End If
   End If
   'end 2022/04/19
   TxtValidate = True
End Function

'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'Modified by Lydia 2021/04/23 改模組
         'm_CP110 = m_CP110 & "," & PUB_GetListBoxTagVal(lstNameAgent, ii)
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2)
   End If
End Sub

'電子送件-申請書
Private Function StartLetter2(ByVal iET01 As String, ByVal iET03 As String, ByVal iCp09 As String, ByVal iKind As String) As Boolean
   Dim strTxt(1 To 30) As String
   Dim ii As Integer, jj As Integer
   Dim tmpArr1 As Variant, tmpArr2 As Variant
   Dim intA As Integer
   Dim bolReadTM As Boolean 'Added by Lydia 2023/11/08
   
   EndLetter iET01, iCp09, iET03, strUserNum
   
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   '申請人資料
   'Modified by Lydia 2020/09/29 +案件性質
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, tm(), False)
   'Modified by Lydia 2022/04/19 延展: 可同時變更申請人
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), False)
   strExc(0) = ""
   'Modified by Lydia 2022/07/28 不限制CheckBox
   'If Frame2.Visible = True And CheckBox1.Value = True Then
   If Frame2.Visible = True Then
     If textFM2(0).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(0).Text)
     If textFM2(1).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(1).Text)
     If textFM2(2).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(2).Text)
     If textFM2(3).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(3).Text)
     If textFM2(4).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(4).Text)
     If strExc(0) <> "" Then strExc(0) = Mid(strExc(0), 2)
   End If
   'Modified by Lydia 2023/11/08
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), False, strExc(0), , tm(1))
   ''end 2022/04/19
   '原本預設抓申請人基本檔之地址;現在改成預設抓案件申請人資料之地址，當勾選「變更地址」，基本資料表之申請人地址改抓申請人基本檔之地址。
   'Modified by Lydia 2023/12/29 單純勾選「變更地址」，只有地址改抓申請人基本檔之地址。ex. FCT-033079
   'If CheckBox1.Value = True Or CheckBox2.Value = True Then
   strExc(1) = ""
   If CheckBox2.Value = True Then
      strExc(1) = strExc(1) & ",申請人地址"
   End If
   If CheckBox1.Value = True Then
      strExc(1) = ""
   'end 2023/12/29
      bolReadTM = False
   Else
      bolReadTM = True
   End If
   'Modified by Lydia 2023/12/29 +指定讀取申請人的資料
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), bolReadTM, strExc(0), , tm(1))
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), bolReadTM, strExc(0), , tm(1), , strExc(1))
   'end 2023/11/08
   
   '出名代理人: 改成共用模組取得資料
   strExc(0) = PUB_GetAgentCP110(iCp09, m_CP110, "FCT", "4")
   If strExc(0) <> "" Then
       tmpArr1 = Empty
       tmpArr1 = Split(strExc(0), "|")
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
               tmpArr2 = Empty
               tmpArr2 = Split(tmpArr1(jj), ",")
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
           End If
       Next jj
   End If
   
   If iKind = "1" Then '基本資料表
        ii = ii + 1
        'FCT程序分機
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','FCT程序分機','" & m_F21st07 & "')"
   End If
   
   If iKind = "2" Then '電子送件申請書
        ii = ii + 1
        '繳費金額
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','繳費金額','" & Text9.Text & "')"
       '收據抬頭 (內商才用)
'        strExc(1) = ""
'        strExc(1) = GetPrjPeople1(ChangeCustomerL(tm(23)))
'        For intI = 78 To 81 '申請人2~4
'            If tm(intI) <> "" Then
'               strExc(1) = strExc(1) & "、" & GetPrjPeople1(ChangeCustomerL(tm(intI)))
'            End If
'        Next intI
'        ii = ii + 1
'        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'           " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','收據抬頭', " & CNULL(ChgSQL(strExc(1))) & ")"
         
        'Added by Lydia 2022/12/19 註冊證形式
        If strSrvDate(1) >= "20230101" Then
           If m_CP10 = "314" Then '申請註冊證副本314: 註冊證形式
              '申請內容1
              ii = ii + 1
              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1','申請本件商標之紙本商標註冊證副本。')"
           End If
           If tm(136) = "1" Then
              strExc(1) = "電子"
           ElseIf tm(136) = "2" Then
              strExc(1) = "紙本"
           Else
              strExc(1) = "電子/紙本"
           End If
           ii = ii + 1
           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','註冊證形式','" & strExc(1) & "')"
        End If
        'end 2022/12/19
        'Added by Lydia 2024/09/12 補證103
        If m_CP10 = "103" Then
            '申請內容1
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1','1.申請補發本件註冊證。" & vbCrLf & "2.具結：本件註冊商標/標章註冊證確實遺失。')"
        End If
        'end 2024/09/12
        
        If m_CP10 = "101" Then '申請
            '主張優先權
            'Modified by Lydia 2019/07/31 + NA72
            'Modified by Lydia 2020/02/19 商標申請書之主張優先權, 國家為239者改寫死帶 EU歐盟----阿蓮 (與專利案不同)
            'strExc(0) = "select sqldatet(PD05) as PD05 ,PD06,NA03,NA72,NVL(A1.TM01||A1.TM02||A1.TM03||A1.TM04,A2.TM01||A2.TM02||A2.TM03||A2.TM04) AS caseno,PD10 " & _
                     "from PRIDATE,NATION,TRADEMARK A1,TRADEMARK A2 " & _
                     "WHERE PD01='" & tm(1) & "' AND PD02='" & tm(2) & "' AND PD03='" & tm(3) & "' AND PD04 ='" & tm(4) & "' " & _
                     "AND PD06=A1.TM12(+) AND PD05=A1.TM11(+) AND PD07=A1.TM10(+) " & _
                     "AND PD06=A2.TM15(+) AND PD05=A2.TM11(+) AND PD07=A2.TM10(+) " & _
                     "AND PD07=NA01(+) " & _
                     "ORDER BY PD01,PD02,PD03,PD04"
            strExc(0) = "select sqldatet(PD05) as PD05 ,PD06,NA03,DECODE(PD07,'239','EU歐盟',NA72) NA72,NVL(A1.TM01||A1.TM02||A1.TM03||A1.TM04,A2.TM01||A2.TM02||A2.TM03||A2.TM04) AS caseno,PD10 " & _
                     "from PRIDATE,NATION,TRADEMARK A1,TRADEMARK A2 " & _
                     "WHERE PD01='" & tm(1) & "' AND PD02='" & tm(2) & "' AND PD03='" & tm(3) & "' AND PD04 ='" & tm(4) & "' " & _
                     "AND PD06=A1.TM12(+) AND PD05=A1.TM11(+) AND PD07=A1.TM10(+) " & _
                     "AND PD06=A2.TM15(+) AND PD05=A2.TM11(+) AND PD07=A2.TM10(+) " & _
                     "AND PD07=NA01(+) " & _
                     "ORDER BY PD01,PD02,PD03,PD04"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                With RsTemp
                    .MoveFirst
                    jj = 1
                    strExc(1) = ""
                    Do While Not .EOF
                         'Modified by Lydia 2019/07/31 na03=>na72 IPO國籍代碼+中文國名
                         strExc(1) = strExc(1) & _
                                          "【主張優先權" & jj & "】  " & vbCrLf & _
                                          "　　【優先權日】　　　　　　　" & .Fields("pd05") & vbCrLf & _
                                          "　　【受理國家或地區】　　　　" & .Fields("na72") & vbCrLf & _
                                          "　　【申請案號】　　　　　　　" & .Fields("pd06") & vbCrLf
                         jj = jj + 1
                         .MoveNext
                    Loop
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','主張優先權','" & ChgSQL(strExc(1)) & "')"
                End With
            'Added by Lydia 2019/09/18 無優先權資料,先帶出空項目(因為資料後補 by阿蓮)
            'Modified by Lydia 2020/02/05 與附送書件分開（因為有時候沒有文件）
            'ElseIf chkAtt1(2).Value = 1 Then
            ElseIf ChkPD.Value = 1 Then
                         jj = 1: strExc(1) = ""
                         strExc(1) = strExc(1) & _
                                          "【主張優先權" & jj & "】  " & vbCrLf & _
                                          "　　【優先權日】　　　　　　　" & vbCrLf & _
                                          "　　【受理國家或地區】　　　　" & vbCrLf & _
                                          "　　【申請案號】　　　　　　　" & vbCrLf
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','主張優先權','" & ChgSQL(strExc(1)) & "')"
            'end 2019/09/18
            End If
        End If
        '商標顏色
        If ChkColor.Visible = True Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商標顏色','" & IIf(ChkColor.Value = 1, "彩色", "墨色") & "')"
        End If
        
        'Mark by Lydia 2025/02/18 重覆; ex.FCT-050544
        'If m_CP10 = "103" Then '補換發證書
        '     strExc(1) = "    1.  申請補發本件註冊證。" & vbCrLf & _
        '                      "    2.  具結：本件註冊商標/標章註冊證確實遺失。"
        '     ii = ii + 1
        '     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '                     " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1','" & ChgSQL(strExc(1)) & "')"
        'End If
         'end 2025/02/18
         
        'Added by Lydia 2019/11/05 判斷畫面選擇"部分延展":自動帶出全部類別和商品服務名稱的內容
        If m_CP10 = "102" And ChkPart.Value = 0 Then
            ii = ii + 1
            strExc(2) = "　　【全部延展】　　　　　　　是"
            'Modified by Lydia 2022/09/15 改變名稱
            'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','部分延展','" & ChgSQL(strExc(2)) & "')"
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別及名稱-部分延展','" & ChgSQL(strExc(2)) & "')"
        Else
        'end 2019/11/05
        
        '商品服務類別及名稱
            If m_CP10 = "101" Or m_CP10 = "102" Then '101申請,102延展
                'Mark by Lydia 2022/09/15 改成即時抓基本檔資料; ex.FCT-49623有38類描述長度有5857字
'                strExc(1) = "": strExc(2) = ""
'                strExc(0) = BeforePrintGetDBData("TMGoods:" & tm(1) & "-" & tm(2) & "-" & tm(3) & "-" & tm(4) & "-||區隔", True)
'                If Trim(strExc(0)) <> "" Then
'                     tmpArr1 = Empty
'                     tmpArr1 = Split(strExc(0), "||")
'                     jj = 1
'                     For intA = 0 To UBound(tmpArr1)
'                         strExc(1) = Trim(tmpArr1(intA))
'                         If strExc(1) <> "" Then
'                              'Modified by Lydia 2019/07/31 阿蓮:FCT不要組群代碼 ; 嘉雯: 內商的商品服務類別內容是用智慧局插件產生,代空白也可以,而外商多半無法用智慧局所以才用人工撰寫
'                              'strExc(2) = strExc(2) & _
'                                               "【指定使用商品服務類別及名稱" & jj & "】  " & vbCrLf & _
'                                               "　　【類別】　　　　　　　　　" & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1) & vbCrLf & _
'                                               IIf(m_CP10 = "101", "　　【組群代碼】　　　　　　　" & vbCrLf, "") & _
'                                               "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
'                              strExc(2) = strExc(2) & _
'                                               "【指定使用商品服務類別及名稱" & jj & "】  " & vbCrLf & _
'                                               "　　【類別】　　　　　　　　　" & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1) & vbCrLf & _
'                                               "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
'                              jj = jj + 1
'                         End If
'                     Next intA
'               'Added by Lydia 2019/07/31 阿蓮: 因為在產生申請書時才撰寫商品服務,所以依收文的類別產生
'               'Memo by Lydia 2019/07/31 嘉雯: 內商的商品服務類別內容是用智慧局插件產生,代空白也可以,而外商多半無法用智慧局所以才用人工撰寫
'               ElseIf tm(9) <> "" Then
'                     tmpArr1 = Empty
'                     tmpArr1 = Split(tm(9), ",")
'                     jj = 1
'                     For intA = 0 To UBound(tmpArr1)
'                         strExc(1) = Trim(tmpArr1(intA))
'                         If strExc(1) <> "" Then
'                              strExc(2) = strExc(2) & _
'                                               "【指定使用商品服務類別及名稱" & jj & "】  " & vbCrLf & _
'                                               "　　【類別】　　　　　　　　　" & strExc(1) & vbCrLf & _
'                                               "　　【商品服務名稱】　　　　　" & vbCrLf
'                              jj = jj + 1
'                         End If
'                     Next intA
'                'end 2019/07/31
'                Else
'                     'Modified by Lydia 2019/07/31 阿蓮:FCT不要組群代碼
'                     'strExc(2) = "【指定使用商品服務類別及名稱1】  " & vbCrLf & _
'                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
'                                     IIf(m_CP10 = "101", "　　【組群代碼】　　　　　　　" & vbCrLf, "") & _
'                                      "　　【商品服務名稱】　　　　　" & vbCrLf
'                     strExc(2) = "【指定使用商品服務類別及名稱1】  " & vbCrLf & _
'                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
'                                      "　　【商品服務名稱】　　　　　" & vbCrLf
'                End If
'                ii = ii + 1
'                '申請
'                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','指定使用商品服務類別及名稱','" & ChgSQL(strExc(2)) & "')"
'                If m_CP10 = "102" Then '延展
'                     strTxt(ii) = Replace(strTxt(ii), "指定使用商品服務類別及名稱", "部分延展")
'                End If
                'end 2022/09/15
            End If
        End If
        '附送書件
        'Modified by Lydia 2023/11/30 改用For Each
        For Each oObj In chkAtt1
            If oObj.Value = 1 Then
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & oObj.Caption & "', '" & m_CaseNo & oObj.Tag & "')"
               If oObj.Index = 3 Then '展覽會優先權證明
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','有展覽會','♀')"
               End If
            End If
        Next
        '若不勾選基本資料表，則附件名稱「未變更本案基本資料」並且不用產生.contact檔案
        If chkAtt1(0).Value = 0 Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & chkAtt1(0).Caption & "', '未變更本案基本資料')"
        End If
        
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

Private Function FormSave() As Boolean
Dim strSqlText As String

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   '出名代理人
   If lstNameAgent.Visible = True Then
      strSql = " UPDATE CASEPROGRESS SET "
      If strSqlText = "" Then
         strSqlText = " cp110 = " & CNULL(m_CP110)
      Else
         strSqlText = strSqlText & " ,cp110 = " & CNULL(m_CP110)
      End If
      strSql = strSql & strSqlText & " WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
   End If
   '預設為電子送件
   If m_CP118 = "Y" Then
      'Modified by Morgan 2019/7/17 目前FCT尚未自動扣款
      'strSql = " UPDATE CASEPROGRESS SET CP118='A' WHERE CP09='" & strReceiveNo & "' AND CP158=0 AND CP118 IS NULL"
      strSql = " UPDATE CASEPROGRESS SET CP118='Y' WHERE CP09='" & strReceiveNo & "' AND CP158=0 AND CP118 IS NULL"
      cnnConnection.Execute strSql, intI
   End If
   
   'Added by Lydia 2022/07/15 (延展案)在"未勾註變更申請人"下再產生申請書, 將一併刪除變更申請人事項記錄
   If Frame2.Visible = True And CheckBox1.Value = False Then
       strSql = "delete from changeevent where ce01='" & strReceiveNo & "'"
       cnnConnection.Execute strSql, intI
   End If
   'end 2022/07/15
   
   'Added by Lydia 2022/04/19 變更申請人
   If Frame2.Visible = True And CheckBox1.Value = True Then
      '檢查是否有此筆文號變更資料
      strExc(0) = "select ce01 from changeevent where ce01='" & strReceiveNo & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "update changeevent" & _
                  " set ce04='" & textFM2(0) & "'" & _
                  " ,ce05='" & textFM2(1) & "'" & _
                  " ,ce06='" & textFM2(2) & "'" & _
                  " ,ce07='" & textFM2(3) & "'" & _
                  " ,ce08='" & textFM2(4) & "'" & _
                  " WHERE CE01='" & strReceiveNo & "'"
      Else
         strSql = "insert into changeevent(ce01,ce04,ce05,ce06,ce07,ce08) values(" & _
                  CNULL(strReceiveNo) & "," & CNULL(textFM2(0)) & "," & CNULL(textFM2(1)) & "," & _
                  CNULL(textFM2(2)) & "," & CNULL(textFM2(3)) & "," & CNULL(textFM2(4)) & ")"
      End If
      cnnConnection.Execute strSql
   End If
   'end 2022/04/19
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

'Added by Lydia 2022/04/19
Private Sub CheckBox1_Click()
    'Modified by Lydia 2022/07/12 申請人不限制要勾選「變更申請人」才能輸入
    'If CheckBox1.Value = True Then
    '    Frame2.Enabled = True
    'Else
    '    Frame2.Enabled = False
    'End If
    'end 2022/07/12
    For intI = 0 To 4
       textFM2(intI).Text = ""
       textFM2(intI).Tag = ""
       lblCName(intI).Caption = ""
    Next intI
End Sub

Private Sub textFM2_GotFocus(Index As Integer)
    TextInverse textFM2(Index)
End Sub

Private Sub textFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFM2_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String
   
   If textFM2(Index).Tag <> textFM2(Index).Text Then
       lblCName(Index).Caption = ""
       If textFM2(Index).Text <> "" Then
           textFM2(Index).Text = ChangeCustomerL(textFM2(Index).Text)
           strTemp = GetCustomerName(textFM2(Index).Text)
           If strTemp = "" Then
               MsgBox "申請人代碼<" & textFM2(Index).Text & ">不正確", vbExclamation
               Cancel = True
           Else
               lblCName(Index).Caption = strTemp
           End If
           textFM2(Index).Tag = textFM2(Index).Text
       End If
   End If
   
   If Cancel = True Then
      textFM2(Index).SetFocus
      textFM2_GotFocus Index
   End If
End Sub

'end 2022/04/19
