VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000018&
   Caption         =   "電腦中心作業"
   ClientHeight    =   4510
   ClientLeft      =   4080
   ClientTop       =   2860
   ClientWidth     =   9120
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  '最大化
   Begin VB.Timer tmrConnect 
      Left            =   30
      Top             =   2190
   End
   Begin VB.Timer tmrSalary 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1395
      Top             =   3600
   End
   Begin VB.Timer Timer2 
      Left            =   30
      Top             =   1710
   End
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   1260
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   280
      Left            =   0
      TabIndex        =   1
      Top             =   4230
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   459
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   917
      ButtonWidth     =   617
      ButtonHeight    =   811
      Appearance      =   1
      _Version        =   393216
      Begin VB.TextBox TxtTestDB 
         Alignment       =   2  '置中對齊
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   380
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "測試資料庫"
         Top             =   60
         Visible         =   0   'False
         Width           =   1990
      End
      Begin VB.ListBox List1 
         Height          =   220
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   3000
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin VB.Menu mnuPop2 
      Caption         =   "公用彈跳選單2"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopItem2 
         Caption         =   "剪下(&T)"
         Index           =   0
      End
      Begin VB.Menu mnuPopItem2 
         Caption         =   "複製(&C)"
         Index           =   1
      End
      Begin VB.Menu mnuPopItem2 
         Caption         =   "貼上(&P)"
         Index           =   2
      End
      Begin VB.Menu mnuPopItem2 
         Caption         =   "刪除(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "系統"
      Index           =   0
      Begin VB.Menu mnu00 
         Caption         =   "切換連線"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00 
         Caption         =   "結束"
         Index           =   1
      End
      Begin VB.Menu mnu00 
         Caption         =   "維護作業1"
         Index           =   2
      End
      Begin VB.Menu mnu00 
         Caption         =   "信頭維護"
         Index           =   4
      End
      Begin VB.Menu mnu00 
         Caption         =   "FCP/FCT特殊信函"
         Index           =   6
      End
      Begin VB.Menu mnu00 
         Caption         =   "其他特殊信函"
         Index           =   7
      End
      Begin VB.Menu mnu00 
         Caption         =   "專利資料匯入(代繳年費用)"
         Index           =   8
      End
      Begin VB.Menu mnu00 
         Caption         =   "整批更新造字資料"
         Index           =   9
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "共同查詢"
      Index           =   10
      Begin VB.Menu mnu10 
         Caption         =   "案件查詢"
         Index           =   1
         Begin VB.Menu mnu101 
            Caption         =   "案件資料及進度查詢"
            Index           =   1
         End
         Begin VB.Menu mnu101 
            Caption         =   "申請人查詢(查本所客戶)"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnu101 
            Caption         =   "申請人查詢（含客戶及對造）"
            Index           =   3
            Shortcut        =   ^S
         End
         Begin VB.Menu mnu101 
            Caption         =   "以發明人查詢"
            Index           =   4
         End
         Begin VB.Menu mnu101 
            Caption         =   "以案件名稱查詢"
            Index           =   5
         End
         Begin VB.Menu mnu101 
            Caption         =   "以國籍查詢代理人/申請人"
            Index           =   6
         End
         Begin VB.Menu mnu101 
            Caption         =   "以國別查詢"
            Index           =   7
         End
         Begin VB.Menu mnu101 
            Caption         =   "申請人查詢案件變更記錄"
            Index           =   8
         End
         Begin VB.Menu mnu101 
            Caption         =   "爭議案件查詢"
            Index           =   9
         End
         Begin VB.Menu mnu101 
            Caption         =   "代理人案件查詢"
            Index           =   10
            Shortcut        =   ^A
         End
         Begin VB.Menu mnu101 
            Caption         =   "監視系統案件查詢"
            Index           =   11
         End
         Begin VB.Menu mnu101 
            Caption         =   "條碼廠商號碼查詢"
            Index           =   12
         End
         Begin VB.Menu mnu101 
            Caption         =   "優先權資料查詢"
            Index           =   13
         End
         Begin VB.Menu mnu101 
            Caption         =   "庭期資料查詢"
            Index           =   14
         End
         Begin VB.Menu mnu101 
            Caption         =   "後金案件及結果查詢"
            Index           =   15
         End
         Begin VB.Menu mnu101 
            Caption         =   "後金收回查詢"
            Index           =   16
         End
         Begin VB.Menu mnu101 
            Caption         =   "關聯案件資料及正聯商標查詢"
            Index           =   17
         End
         Begin VB.Menu mnu101 
            Caption         =   "專利法律案件關聯查詢"
            Index           =   18
         End
         Begin VB.Menu mnu101 
            Caption         =   "介紹法律所案源查詢"
            Index           =   19
         End
      End
      Begin VB.Menu mnu10 
         Caption         =   "收發文查詢"
         Index           =   2
         Begin VB.Menu mnu102 
            Caption         =   "以收/發文日查詢"
            Index           =   1
         End
         Begin VB.Menu mnu102 
            Caption         =   "以收/發文量查詢"
            Index           =   2
         End
         Begin VB.Menu mnu102 
            Caption         =   "收文未發文查詢"
            Index           =   3
         End
         Begin VB.Menu mnu102 
            Caption         =   "以收文日查詢來函"
            Index           =   4
         End
         Begin VB.Menu mnu102 
            Caption         =   "以期限管制日查詢"
            Index           =   5
         End
         Begin VB.Menu mnu102 
            Caption         =   "承辦人收/發文量查詢"
            Index           =   6
         End
         Begin VB.Menu mnu102 
            Caption         =   "發文日查詢代理人作業進度"
            Index           =   7
         End
      End
      Begin VB.Menu mnu10 
         Caption         =   "國外部查詢"
         Index           =   3
         Begin VB.Menu mnu103 
            Caption         =   "客戶重新委任案件查詢列印"
            Index           =   1
         End
      End
      Begin VB.Menu mnu10 
         Caption         =   "員工姓名查詢員工資料"
         Index           =   4
      End
      Begin VB.Menu mnu10 
         Caption         =   "程式公告查詢"
         Index           =   5
      End
      Begin VB.Menu mnu10 
         Caption         =   "臺灣地址郵遞區號查詢"
         Index           =   6
      End
      Begin VB.Menu mnu10 
         Caption         =   "各項指示分類查詢"
         Index           =   7
      End
      Begin VB.Menu mnu10 
         Caption         =   "不得宣傳客戶名稱資料查詢"
         Index           =   8
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "共同程序"
      Index           =   11
      Begin VB.Menu mnu11 
         Caption         =   "期限管制解除"
         Index           =   1
         Begin VB.Menu mnu1101 
            Caption         =   "解除期限"
            Index           =   1
         End
         Begin VB.Menu mnu1101 
            Caption         =   "取消收文"
            Index           =   2
         End
         Begin VB.Menu mnu1101 
            Caption         =   "閉卷"
            Index           =   3
         End
      End
      Begin VB.Menu mnu11 
         Caption         =   "內部收文"
         Index           =   2
         Begin VB.Menu mnu1102 
            Caption         =   "新增"
            Index           =   1
         End
         Begin VB.Menu mnu1102 
            Caption         =   "修改"
            Index           =   2
         End
         Begin VB.Menu mnu1102 
            Caption         =   "查詢"
            Index           =   3
         End
      End
      Begin VB.Menu mnu11 
         Caption         =   "聯絡單列印及E-Mail"
         Index           =   3
      End
      Begin VB.Menu mnu11 
         Caption         =   "相關卷號資料維護"
         Index           =   4
      End
      Begin VB.Menu mnu11 
         Caption         =   "多案相關卷號關係建立"
         Index           =   5
      End
      Begin VB.Menu mnu11 
         Caption         =   "分割案件關係維護"
         Index           =   6
      End
      Begin VB.Menu mnu11 
         Caption         =   "撰寫信函作業"
         Index           =   7
      End
      Begin VB.Menu mnu11 
         Caption         =   "作業失誤"
         Index           =   8
         Begin VB.Menu mnu1107 
            Caption         =   "作業失誤資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu1107 
            Caption         =   "作業失誤明細表"
            Index           =   2
         End
      End
      Begin VB.Menu mnu11 
         Caption         =   "代理人帳目"
         Index           =   9
         Begin VB.Menu mnu1104 
            Caption         =   "資料維護"
            Index           =   1
            Begin VB.Menu mnu110401 
               Caption         =   "帳單輸入"
               Index           =   1
            End
            Begin VB.Menu mnu110401 
               Caption         =   "抵帳單輸入"
               Index           =   2
            End
            Begin VB.Menu mnu110401 
               Caption         =   "帳單作廢作業"
               Index           =   3
            End
            Begin VB.Menu mnu110401 
               Caption         =   "請款單輸入"
               Index           =   4
            End
            Begin VB.Menu mnu110401 
               Caption         =   "折讓輸入"
               Index           =   5
            End
            Begin VB.Menu mnu110401 
               Caption         =   "請款單作廢作業"
               Index           =   6
            End
            Begin VB.Menu mnu110401 
               Caption         =   "請款項目資料"
               Index           =   7
            End
            Begin VB.Menu mnu110401 
               Caption         =   "美金請款匯率資料維護"
               Index           =   8
            End
            Begin VB.Menu mnu110401 
               Caption         =   "預估結匯匯率資料維護"
               Index           =   9
            End
            Begin VB.Menu mnu110401 
               Caption         =   "相同案件性質請款作業"
               Index           =   10
            End
            Begin VB.Menu mnu110401 
               Caption         =   "主管審核作業"
               Index           =   11
            End
            Begin VB.Menu mnu110401 
               Caption         =   "其他幣別請款匯率資料維護"
               Index           =   12
            End
            Begin VB.Menu mnu110401 
               Caption         =   "客製化請款項目資料維護"
               Index           =   13
            End
         End
         Begin VB.Menu mnu1104 
            Caption         =   "資料查詢"
            Index           =   2
            Begin VB.Menu mnu110402 
               Caption         =   "國外代理人帳目查詢"
               Index           =   1
            End
            Begin VB.Menu mnu110402 
               Caption         =   "國外案件帳目查詢"
               Index           =   2
            End
            Begin VB.Menu mnu110402 
               Caption         =   "國外請款金額查詢"
               Index           =   3
            End
            Begin VB.Menu mnu110402 
               Caption         =   "案件損益查詢"
               Index           =   4
            End
            Begin VB.Menu mnu110402 
               Caption         =   "各幣別最新請款匯率查詢"
               Index           =   5
            End
         End
         Begin VB.Menu mnu1104 
            Caption         =   "報表列印"
            Index           =   3
            Begin VB.Menu mnu110403 
               Caption         =   "催款單列印"
               Index           =   1
            End
            Begin VB.Menu mnu110403 
               Caption         =   "請款單列印"
               Index           =   2
            End
            Begin VB.Menu mnu110403 
               Caption         =   "請款單整批列印"
               Index           =   3
            End
            Begin VB.Menu mnu110403 
               Caption         =   "國外FC帳款明細表"
               Index           =   4
            End
            Begin VB.Menu mnu110403 
               Caption         =   "代理人帳目排名"
               Index           =   5
            End
            Begin VB.Menu mnu110403 
               Caption         =   "FC業務請款／收款明細表"
               Index           =   6
            End
            Begin VB.Menu mnu110403 
               Caption         =   "FC專業收款明細表"
               Index           =   7
            End
            Begin VB.Menu mnu110403 
               Caption         =   "代理人逾期帳款分析"
               Index           =   8
            End
            Begin VB.Menu mnu110403 
               Caption         =   "折讓單列印"
               Index           =   9
            End
            Begin VB.Menu mnu110403 
               Caption         =   "請款單折扣案件明細"
               Index           =   10
            End
         End
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘單查詢"
         Index           =   10
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘資料維護"
         Index           =   11
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘單案件明細查詢"
         Index           =   12
      End
      Begin VB.Menu mnu11 
         Caption         =   "部門別送件清單列印"
         Index           =   13
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "電腦中心"
      Index           =   12
      Begin VB.Menu mnu1204 
         Caption         =   "資料維護"
         Index           =   1
         Begin VB.Menu mnu120401 
            Caption         =   "自動編號資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu120401 
            Caption         =   "系統種類對照資料維護"
            Index           =   2
         End
         Begin VB.Menu mnu120401 
            Caption         =   "員工資料維護"
            Index           =   3
         End
         Begin VB.Menu mnu120401 
            Caption         =   "員工密碼資料維護"
            Index           =   4
         End
         Begin VB.Menu mnu120401 
            Caption         =   "員工等級資料維護"
            Index           =   5
         End
         Begin VB.Menu mnu120401 
            Caption         =   "員工權限資料維護"
            Index           =   6
         End
         Begin VB.Menu mnu120401 
            Caption         =   "員工跨部門權限資料維護"
            Index           =   7
         End
         Begin VB.Menu mnu120401 
            Caption         =   "案件來源對照資料維護"
            Index           =   8
         End
         Begin VB.Menu mnu120401 
            Caption         =   "案件性質對照資料維護"
            Index           =   9
         End
         Begin VB.Menu mnu120401 
            Caption         =   "國家基本資料維護"
            Index           =   10
         End
         Begin VB.Menu mnu120401 
            Caption         =   "專利商標種類對照資料維護"
            Index           =   11
         End
         Begin VB.Menu mnu120401 
            Caption         =   "解除期限原因資料維護"
            Index           =   12
         End
         Begin VB.Menu mnu120401 
            Caption         =   "資料刪除記錄維護"
            Index           =   13
         End
         Begin VB.Menu mnu120401 
            Caption         =   "專利年費對照資料維護"
            Index           =   14
         End
         Begin VB.Menu mnu120401 
            Caption         =   "專利基本資料維護"
            Index           =   15
         End
         Begin VB.Menu mnu120401 
            Caption         =   "商標基本資料維護"
            Index           =   16
         End
         Begin VB.Menu mnu120401 
            Caption         =   "法務基本資料維護"
            Index           =   17
         End
         Begin VB.Menu mnu120401 
            Caption         =   "顧問基本資料維護"
            Index           =   18
         End
         Begin VB.Menu mnu120401 
            Caption         =   "服務業務基本資料維護"
            Index           =   19
            Begin VB.Menu mnu12040117 
               Caption         =   "專利服務業務"
               Index           =   1
            End
            Begin VB.Menu mnu12040117 
               Caption         =   "商標服務業務－條碼"
               Index           =   2
            End
            Begin VB.Menu mnu12040117 
               Caption         =   "商標服務業務－監視系統"
               Index           =   3
            End
            Begin VB.Menu mnu12040117 
               Caption         =   "商標服務業務－網域"
               Index           =   4
            End
            Begin VB.Menu mnu12040117 
               Caption         =   "商標服務業務－著作權"
               Index           =   5
            End
            Begin VB.Menu mnu12040117 
               Caption         =   "商標服務業務－其他業務"
               Index           =   6
            End
         End
         Begin VB.Menu mnu120401 
            Caption         =   "案件進度資料維護"
            Index           =   20
         End
         Begin VB.Menu mnu120401 
            Caption         =   "下一程序資料維護"
            Index           =   21
         End
         Begin VB.Menu mnu120401 
            Caption         =   "個人目標資料維護"
            Index           =   22
         End
         Begin VB.Menu mnu120401 
            Caption         =   "案件國家收費表維護"
            Index           =   23
         End
         Begin VB.Menu mnu120401 
            Caption         =   "員工群組資料維護"
            Index           =   24
         End
         Begin VB.Menu mnu120401 
            Caption         =   "工作天維護"
            Index           =   25
         End
         Begin VB.Menu mnu120401 
            Caption         =   "客戶資料維護"
            Index           =   26
         End
         Begin VB.Menu mnu120401 
            Caption         =   "客戶發明人資料維護"
            Index           =   27
         End
         Begin VB.Menu mnu120401 
            Caption         =   "客戶變更名稱作業"
            Index           =   28
         End
         Begin VB.Menu mnu120401 
            Caption         =   "代理人資料維護"
            Index           =   29
         End
         Begin VB.Menu mnu120401 
            Caption         =   "代理人變更名稱作業"
            Index           =   30
         End
         Begin VB.Menu mnu120401 
            Caption         =   "特殊專利商標資料維護"
            Index           =   31
         End
         Begin VB.Menu mnu120401 
            Caption         =   "系統特殊設定"
            Index           =   32
         End
         Begin VB.Menu mnu120401 
            Caption         =   "郵件排程維護"
            Index           =   33
         End
         Begin VB.Menu mnu120401 
            Caption         =   "電子收文接洽單查詢"
            Index           =   34
            Visible         =   0   'False
         End
         Begin VB.Menu mnu120401 
            Caption         =   "不得代理案件之客戶或代理人資料維護"
            Index           =   35
         End
         Begin VB.Menu mnu120401 
            Caption         =   "非本所實質客戶資料維護"
            Index           =   36
         End
         Begin VB.Menu mnu120401 
            Caption         =   "LEDES基本資料維護"
            Index           =   37
         End
         Begin VB.Menu mnu120401 
            Caption         =   "特殊客戶/代理人收文費用維護"
            Index           =   38
         End
         Begin VB.Menu mnu120401 
            Caption         =   "程式公告維護"
            Index           =   39
         End
         Begin VB.Menu mnu120401 
            Caption         =   "造字與UnidCode字對照表"
            Index           =   40
         End
         Begin VB.Menu mnu120401 
            Caption         =   "案件表單簽核人員設定"
            Index           =   41
         End
         Begin VB.Menu mnu120401 
            Caption         =   "各項指示分類維護"
            Index           =   42
         End
         Begin VB.Menu mnu120401 
            Caption         =   "國外部關聯企業分類維護"
            Index           =   43
         End
         Begin VB.Menu mnu120401 
            Caption         =   "國內收據點數分配"
            Index           =   44
            Visible         =   0   'False
         End
         Begin VB.Menu mnu120401 
            Caption         =   "委任契約書用印記錄查詢"
            Index           =   45
            Visible         =   0   'False
         End
         Begin VB.Menu mnu120401 
            Caption         =   "核判表設定作業"
            Index           =   46
         End
         Begin VB.Menu mnu120401 
            Caption         =   "查詢特殊置換字對照表"
            Index           =   47
            Visible         =   0   'False
         End
         Begin VB.Menu mnu120401 
            Caption         =   "電子報特殊名單維護"
            Index           =   48
         End
      End
      Begin VB.Menu mnu1205 
         Caption         =   "定期作業"
         Index           =   1
         Begin VB.Menu mnu120501 
            Caption         =   "來文資料稽核表"
            Index           =   1
         End
         Begin VB.Menu mnu120501 
            Caption         =   "規費資料稽核表"
            Index           =   2
         End
         Begin VB.Menu mnu120501 
            Caption         =   "自動核准"
            Index           =   3
         End
         Begin VB.Menu mnu120501 
            Caption         =   "新客戶清單"
            Index           =   5
         End
         Begin VB.Menu mnu120501 
            Caption         =   "收文未發文明細表"
            Index           =   6
         End
         Begin VB.Menu mnu120501 
            Caption         =   "逾期未處理案件明細表"
            Index           =   7
         End
         Begin VB.Menu mnu120501 
            Caption         =   "本所期限工作天推算"
            Index           =   8
         End
      End
      Begin VB.Menu mnu1206 
         Caption         =   "不定期作業"
         Begin VB.Menu mnu120601 
            Caption         =   "客戶／代理人改號作業"
            Index           =   1
         End
         Begin VB.Menu mnu120601 
            Caption         =   "案件改號作業"
            Index           =   2
         End
         Begin VB.Menu mnu120601 
            Caption         =   "刪除記錄統計表"
            Index           =   4
         End
         Begin VB.Menu mnu120601 
            Caption         =   "智權人員客戶轉移作業"
            Index           =   5
         End
         Begin VB.Menu mnu120601 
            Caption         =   "智權人員調區作業"
            Index           =   6
         End
         Begin VB.Menu mnu120601 
            Caption         =   "收文未發文明細表"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnu120601 
            Caption         =   "備份作業--全部"
            Index           =   8
         End
         Begin VB.Menu mnu120601 
            Caption         =   "備份作業--總帳年度"
            Index           =   9
         End
         Begin VB.Menu mnu120601 
            Caption         =   "客戶案件總簿輸出"
            Index           =   10
         End
         Begin VB.Menu mnu120601 
            Caption         =   "代理人案件總簿列印"
            Index           =   11
         End
         Begin VB.Menu mnu120601 
            Caption         =   "客戶/代理人名冊、地址條列印"
            Index           =   12
         End
         Begin VB.Menu mnu120601 
            Caption         =   "國內客戶名條"
            Index           =   13
         End
         Begin VB.Menu mnu120601 
            Caption         =   "閉卷清單"
            Index           =   14
         End
         Begin VB.Menu mnu120601 
            Caption         =   "分所案號檢核表"
            Index           =   15
         End
         Begin VB.Menu mnu120601 
            Caption         =   "國外客戶/代理人地址條列印"
            Index           =   16
         End
         Begin VB.Menu mnu120601 
            Caption         =   "智權人員客戶名冊 (依最後收文日期區間)"
            Index           =   17
         End
         Begin VB.Menu mnu120601 
            Caption         =   "員工查詢印表記錄資料查詢"
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu mnu120601 
            Caption         =   "客戶案件預算預估表"
            Index           =   19
         End
         Begin VB.Menu mnu120601 
            Caption         =   "開拓名單轉入國內潛在客戶作業"
            Index           =   20
         End
         Begin VB.Menu mnu120601 
            Caption         =   "新建指紋整批匯入"
            Index           =   21
         End
         Begin VB.Menu mnu120601 
            Caption         =   "員工指紋卡片資料"
            Index           =   22
         End
         Begin VB.Menu mnu120601 
            Caption         =   "考勤機設定"
            Index           =   23
         End
         Begin VB.Menu mnu120601 
            Caption         =   "客戶案件總簿-紙本作業列印"
            Index           =   24
         End
         Begin VB.Menu mnu120601 
            Caption         =   "更換FC代理人作業"
            Index           =   25
         End
         Begin VB.Menu mnu120601 
            Caption         =   "國外部人員離職修改資料"
            Index           =   26
         End
         Begin VB.Menu mnu120601 
            Caption         =   "設定颱風假作業"
            Index           =   27
         End
         Begin VB.Menu mnu120601 
            Caption         =   "外專案件清單Excel"
            Index           =   28
         End
      End
      Begin VB.Menu mnu1207 
         Caption         =   "查詢作業"
         Index           =   1
         Begin VB.Menu mnu120701 
            Caption         =   "員工查詢印表記錄資料查詢"
            Index           =   1
         End
         Begin VB.Menu mnu120701 
            Caption         =   "委任契約書用印記錄查詢"
            Index           =   2
         End
         Begin VB.Menu mnu120701 
            Caption         =   "電子收文接洽單查詢"
            Index           =   3
         End
         Begin VB.Menu mnu120701 
            Caption         =   "查詢特殊置換字對照表"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "設定"
      Index           =   16
      Begin VB.Menu mnu16 
         Caption         =   "系統印表機設定"
         Index           =   0
      End
      Begin VB.Menu mnu16 
         Caption         =   "報表紙張格式設定"
         Index           =   1
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "國外開拓"
      Index           =   22
      Begin VB.Menu mnu22 
         Caption         =   "潛在客戶資料維護"
         Index           =   0
      End
      Begin VB.Menu mnu22 
         Caption         =   "客戶/代理人聯絡人資料維護"
         Index           =   1
      End
      Begin VB.Menu mnu22 
         Caption         =   "往來記錄資料維護"
         Index           =   2
      End
      Begin VB.Menu mnu22 
         Caption         =   "交換名片記錄維護"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnu22 
         Caption         =   "互惠代理人資料維護"
         Index           =   4
      End
      Begin VB.Menu mnu22 
         Caption         =   "互惠代理人案件統計表"
         Index           =   5
      End
      Begin VB.Menu mnu22 
         Caption         =   "潛在客戶名條列印"
         Index           =   6
      End
      Begin VB.Menu mnu22 
         Caption         =   "潛在客戶資料查詢"
         Index           =   7
      End
      Begin VB.Menu mnu22 
         Caption         =   "往來記錄資料查詢"
         Index           =   8
      End
      Begin VB.Menu mnu22 
         Caption         =   "往來記錄統計"
         Index           =   9
      End
      Begin VB.Menu mnu22 
         Caption         =   "國外部新客戶/代理人查詢"
         Index           =   10
      End
      Begin VB.Menu mnu22 
         Caption         =   "不得宣傳客戶名稱資料查詢"
         Index           =   11
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "一般作業"
      Index           =   23
      Begin VB.Menu mnu23 
         Caption         =   "預約作業"
         Index           =   1
      End
      Begin VB.Menu mnu23 
         Caption         =   "出缺勤作業"
         Index           =   2
         Begin VB.Menu mnu232 
            Caption         =   "表單"
            Index           =   1
            Begin VB.Menu mnu2321 
               Caption         =   "目前表單"
               Index           =   1
            End
            Begin VB.Menu mnu2321 
               Caption         =   "職代/簽核主管代填表單"
               Index           =   2
            End
            Begin VB.Menu mnu2321 
               Caption         =   "打卡異常個人處理"
               Index           =   3
            End
         End
         Begin VB.Menu mnu232 
            Caption         =   "簽核"
            Index           =   2
            Begin VB.Menu mnu2322 
               Caption         =   "簽核作業"
               Index           =   1
            End
            Begin VB.Menu mnu2322 
               Caption         =   "每月出缺勤統計確認"
               Index           =   2
            End
            Begin VB.Menu mnu2322 
               Caption         =   "員工個人資料明細確認"
               Index           =   3
            End
            Begin VB.Menu mnu2322 
               Caption         =   "簽核人員異動作業"
               Index           =   4
            End
            Begin VB.Menu mnu2322 
               Caption         =   "打卡異常主管處理"
               Index           =   5
            End
         End
         Begin VB.Menu mnu232 
            Caption         =   "查詢"
            Index           =   3
            Begin VB.Menu mnu2323 
               Caption         =   "近日請假公佈欄/工作所在地"
               Index           =   1
            End
            Begin VB.Menu mnu2323 
               Caption         =   "出缺勤查詢"
               Index           =   2
            End
            Begin VB.Menu mnu2323 
               Caption         =   "職代/簽核主管關聯查詢"
               Index           =   3
            End
            Begin VB.Menu mnu2323 
               Caption         =   "打卡資料查詢"
               Index           =   4
            End
         End
      End
      Begin VB.Menu mnu23 
         Caption         =   "教育訓練登錄作業"
         Index           =   3
      End
      Begin VB.Menu mnu23 
         Caption         =   "客戶端平台帳號管理作業"
         Index           =   4
      End
      Begin VB.Menu mnu23 
         Caption         =   "薪資查詢系統"
         Index           =   6
         Begin VB.Menu mnu2306 
            Caption         =   "員工薪資明細"
            Index           =   1
         End
         Begin VB.Menu mnu2306 
            Caption         =   "勞保/健保/勞退金明細"
            Index           =   2
         End
         Begin VB.Menu mnu2306 
            Caption         =   "年度各項所得明細"
            Index           =   3
         End
         Begin VB.Menu mnu2306 
            Caption         =   "年終獎金明細"
            Index           =   4
         End
         Begin VB.Menu mnu2306 
            Caption         =   "薪資查詢密碼修改"
            Index           =   5
         End
      End
      Begin VB.Menu mnu23 
         Caption         =   "信箱分信紀錄查詢"
         Index           =   7
      End
      Begin VB.Menu mnu23 
         Caption         =   "圖書借閱資料查詢"
         Index           =   8
      End
      Begin VB.Menu mnu23 
         Caption         =   "行事曆提醒通知"
         Index           =   9
      End
      Begin VB.Menu mnu23 
         Caption         =   "風險檢查對象資料維護"
         Index           =   10
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "說明"
      Index           =   97
      Begin VB.Menu mnu15 
         Caption         =   "說明主題"
         Index           =   1
      End
      Begin VB.Menu mnu15 
         Caption         =   "索引"
         Index           =   2
      End
      Begin VB.Menu mnu15 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnu15 
         Caption         =   "關於"
         Index           =   4
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "視窗"
      Index           =   99
      Begin VB.Menu mnu99 
         Caption         =   "最近開啟畫面"
         Index           =   0
      End
   End
   Begin VB.Menu mnuChUser 
      Caption         =   "更改使用者"
   End
   Begin VB.Menu mnuDML 
      Caption         =   "查維護紀錄"
      Index           =   0
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPop 
      Caption         =   "公用彈跳選單"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopItem 
         Caption         =   "新增"
         Index           =   0
      End
      Begin VB.Menu mnuPopItem 
         Caption         =   "修改"
         Index           =   1
      End
      Begin VB.Menu mnuPopItem 
         Caption         =   "刪除"
         Index           =   2
      End
      Begin VB.Menu mnuPopItem 
         Caption         =   "檢視"
         Index           =   3
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'sonia 2010/8/19 日期欄已修改
Option Explicit

'intPCaseKind分案之系統分類，intPWhere 0國內  1國外CF  2國外FC
Public intPCaseKind As Integer, intPWhere As Integer
Public m_ChkIsOpenFrm180203 As Boolean 'Add By Sindy 2013/7/8


Private Sub MDIForm_Activate()
   'Modify By Sindy 2025/11/3 改為共用函數
   Call MDIFormStarProc
End Sub

'Add By Sindy 2025/11/3
Public Sub SetTmpForm()
   Set Tmpfrm210147 = frm210147
   Set Tmpfrm210148 = frm210148
   'Add By Sindy 2025/11/3
   Set Tmpfrm180201 = frm180201
   Set Tmpfrm180101 = frm180101
   Set Tmpfrm180203_1 = frm180203_1
   Set Tmpfrm160102 = frm160102
   Set Tmpfrm160018 = frm160018
   Set Tmpfrm010035_2 = frm010035_2
   '2025/11/3 END
End Sub

'Add By Sindy 2011/10/7
Public Sub SysStartCallForm()
   '此函數在各系統一啟動時,因出缺勤待辦提示納入之故,共用會使用到,所以不可刪除
End Sub

Private Sub MDIForm_Load()
Dim strSysKind As String
Dim lngValue, lngBufferSize As Long, intCounter As Integer
Dim strUserId As String * 10, strLocalId As String

    '若登入成功
    If pub_str_LoginSucceeded = "1" Then
        Me.Timer1.Interval = 100
        'Modify By Cheng 2003/07/10
        'Begin
    '   Set cnnConnection = objPublicData.Connection
        'End
       strSysKind = GetSystemKindByNick
       'add by nickc 2006/06/09 可以查詢維護紀錄
       If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
            mnuDML(0).Visible = True
            mnuChUser.Visible = True 'Added by Sindy 2013/7/23
       Else
            mnuDML(0).Visible = False
            mnuChUser.Visible = False 'Added by Sindy 2013/7/23
       End If

'       'Added by Morgan 2016/1/22 薪資查詢測試
'       If Pub_StrUserSt03 = "M51" Then
'         mnu23(6).Visible = True
'       Else
'         mnu23(6).Visible = False
'       End If
'       'end
      
'      'Add By Sindy 2018/4/12
'      If strUserNum = "97038" Then
'         mnu120401(46).Visible = True
'      Else
'         mnu120401(46).Visible = False
'      End If
'      '2018/4/12 END
      
       If bolFNation = False Then
    'Ken 90/07/06
    '      mnu10(14).Visible = False
    '      mnu10(15).Caption = "以國籍查詢申請人"
          mnu101(10).Visible = False 'Modify by Amy 2014/05/05 '原:mnu101(9)
          mnu102(7).Visible = False
    'Ken 90/07/06
          mnu101(6).Visible = False 'Modify by Amy 2014/05/05 '原:mnu101(5)
          '92.3.17 add by sonia
          mnu102(2).Visible = False
          '92.3.17 end
            'Add By Cheng 2003/08/13
            '業務收/發文量比較查詢
'          mnu10(23).Visible = False
       End If
        If Me.ImageList1.ListImages.Count <= 0 Then
           ImageList1.ListImages.add , "graphic1", LoadPicture(strPicPath & "misc41.ico")
           ImageList1.ListImages.add , "graphic2", LoadPicture(strPicPath & "note16.ico")
           ImageList1.ListImages.add , "graphic3", LoadPicture(strPicPath & "erase02.ico")
           ImageList1.ListImages.add , "graphic4", LoadPicture(strPicPath & "drive03.ico")
           ImageList1.ListImages.add , "graphic5", LoadPicture(strPicPath & "trash02.ico")
           ImageList1.ListImages.add , "graphic6", LoadPicture(strPicPath & "explorer.ico")
           ImageList1.ListImages.add , "graphic7", LoadPicture(strPicPath & "printfld.ico")
           ImageList1.ListImages.add , "graphic8", LoadPicture(strPicPath & "first.ico")
           ImageList1.ListImages.add , "graphic9", LoadPicture(strPicPath & "prior.ico")
           ImageList1.ListImages.add , "graphic10", LoadPicture(strPicPath & "next.ico")
           ImageList1.ListImages.add , "graphic11", LoadPicture(strPicPath & "last.ico")
           ImageList1.ListImages.add , "graphic12", LoadPicture(strPicPath & "net14.ico")
           ImageList1.ListImages.add , "graphic13", LoadPicture(strPicPath & "w95mbx01.ico")
           Toolbar1.ImageList = ImageList1
           Toolbar1.Buttons.add , "function1", , tbrDefault, "graphic1"
           Toolbar1.Buttons.add , "none1", , tbrSeparator
           Toolbar1.Buttons.add , "none2", , tbrSeparator
           Toolbar1.Buttons.add , "function2", , tbrDefault, "graphic2"
           Toolbar1.Buttons.add , "function3", , tbrDefault, "graphic3"
           Toolbar1.Buttons.add , "function4", , tbrDefault, "graphic4"
           Toolbar1.Buttons.add , "function12", , tbrDefault, "graphic13"
           Toolbar1.Buttons.add , "function5", , tbrDefault, "graphic5"
           Toolbar1.Buttons.add , "function6", , tbrDefault, "graphic6"
        '   Toolbar1.Buttons.Add , "function7", , tbrDefault, "graphic7"
           Toolbar1.Buttons.add , "none5", , tbrSeparator
           Toolbar1.Buttons.add , "none3", , tbrSeparator
           Toolbar1.Buttons.add , "none4", , tbrSeparator
           Toolbar1.Buttons.add , "function8", , tbrDefault, "graphic8"
           Toolbar1.Buttons.add , "function9", , tbrDefault, "graphic9"
           Toolbar1.Buttons.add , "function10", , tbrDefault, "graphic10"
           Toolbar1.Buttons.add , "function11", , tbrDefault, "graphic11"
           Toolbar1.Buttons.Item(1).ToolTipText = "關閉(Esc)"
           Toolbar1.Buttons.Item(4).ToolTipText = "新增(F2)"
           Toolbar1.Buttons.Item(5).ToolTipText = "修改(F3)"
           Toolbar1.Buttons.Item(6).ToolTipText = "存檔(F9)"
           Toolbar1.Buttons.Item(7).ToolTipText = "放棄(F10)"
           Toolbar1.Buttons.Item(8).ToolTipText = "刪除(F5)"
           Toolbar1.Buttons.Item(9).ToolTipText = "查詢(F4)"
        '   Toolbar1.Buttons.Item(10).ToolTipText = "列印(F7)"
           Toolbar1.Buttons.Item(13).ToolTipText = "第一筆(Home)"
           Toolbar1.Buttons.Item(14).ToolTipText = "上一筆(PageUp)"
           Toolbar1.Buttons.Item(15).ToolTipText = "下一筆(PageDown)"
           Toolbar1.Buttons.Item(16).ToolTipText = "最後(End)"
        End If
       Me.TxtTestDB.Left = Toolbar1.Buttons.Item(16).Left + 500 'Add By Sindy 2023/11/8
       tool4_enabled
       strFormName = MsgText(601)
       strExitControl = MsgText(602)
        If Me.StatusBar1.Panels.Count < 5 Then
            For intCounter = 1 To 4
               StatusBar1.Panels.add
            Next intCounter
        End If
       StatusBar1.Height = 300
       StatusBar1.Panels.Item(1).Width = 5500
       StatusBar1.Panels.Item(2).Width = 1000
       StatusBar1.Panels.Item(3).Text = CFDate(ACDate(ServerDate))
       StatusBar1.Panels.Item(4).Text = time
'       Me.Icon = LoadPicture(strIcoPath)
       ToolHide
       Systemkind_g = GetSystemKindByNick
       Systemkind_g_P = GetSystemKindByNickP
       Systemkind_g_T = GetSystemKindByNickT
       Systemkind_g_TnoS = GetSystemKindByNickTnoS
    End If
End Sub

'edit by nick 2004/12/14
'Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Dim frm As Form
'
'    'Add By Cheng 2003/01/30
'    '使用者從表單上的控制功能表中選取「關閉」指令, 則取消動入
'    If UnloadMode = 0 Then
'        MsgBox "請按 [系統] --> [結束]，以結束本系統!!!", vbExclamation + vbOKOnly
'        Cancel = True
'    Else
'        '關閉尚未關閉的子視窗
'        For Each frm In Forms
'            If frm.Name <> mdiMain.Name Then
'                Unload frm
'            End If
'        Next
'    End If
'End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'edit by nickc 2007/02/09 不用 dll 了
'Set obj001 = Nothing
'Set objPublicData = Nothing
   ' 90.08.16 modify by louis (釋放Word物件)
   EndOfficeAp
   PUB_AddAuditLog AL_登出 'Added by Morgan 2025/7/31
End Sub
Private Sub mnu00_Click(Index As Integer)
'Modify By Cheng 2003/02/17
Select Case Index
    Case 0   '切換連線
         If PUB_Connect2DB(True) = False Then
            Unload Me
         End If
         'Add By Sindy 2023/11/8
         If InStr(pub_HostName, "97038") > 0 Then '測試中
         If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then '測試資料庫
            ToolHide
            Me.TxtTestDB.Visible = False
         Else
            ToolShow
            Me.TxtTestDB.Visible = True
         End If
         End If
         '2023/11/8 END
    Case 1   '結束
        Unload Me
    Case 2   '維護作業1
        If CheckUse("frm000001", strExec) = True Then
            frm000001.Show
        End If
    'Add By Cheng 2003/05/13
    'Removed by Morgan 2022/1/11 刪除(不再使用)
    'Case 3   '代理人電子檔
    '    If CheckUse("frm000002", strExec) = True Then
    '        frm000002.Show
    '    End If
    'end 2022/1/11
        
    'Added by Morgan 2015/6/26
    Case 4   '信頭維護
        If CheckUse("frm000003", strExec) = True Then
            frm000003.Show
        End If
        
    'Add by Morgan 2009/2/26
    Case 6  'FCP/FCT特殊信函
        frm12040149.Show
        
    'Add by Morgan 2010/1/18
    Case 7  '其他特殊信函
        frm12040151.Show
        
    'Added by Morgan 2012/2/4
    Case 8 '專利資料匯入(代繳年費用)
         frm12040153.Show
         
    'Added by Lydia 2022/03/21
    Case 9  '整批更新造字資料
        frm001_1.Show
End Select
End Sub

Private Sub mnu10_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 4
         'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm100121_1") = False Then
            Set frm100121_1 = Nothing
         End If
         'end 2021/12/16
         'Modify by Amy 2014/04/30 Mark CheckUse
         'If CheckUse("frm100121_1", strExec) Then
            frm100121_1.Show
         'End If
      Case 5 'Add By Amy 2013/05/08 程式公告查詢
            'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm100131") = False Then
               Set frm100131 = Nothing
            End If
            'end 2021/12/16
            frm100131.Show
      Case 6 'Add By Sindy 2015/3/20 臺灣地址郵遞區號查詢
         'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm100134") = False Then
               Set frm100134 = Nothing
            End If
            'end 2021/12/16
            frm100134.Show
      'Added by Lydia 2020/08/31
      Case 7   '各項指示分類查詢
            'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm140415_1") = False Then
               Set frm140415_1 = Nothing
            End If
            'end 2021/12/16
            frm140415_1.Show
      'Added by Lydia 2023/06/17
      Case 8  '不得宣傳客戶名稱資料查詢
           If PUB_CheckFormExist("frm100136") = False Then
              Set frm100136 = Nothing
           End If
           If frm100136.ChkUseRight = True Then
              frm100136.Show
           End If
   End Select
End Sub

Private Sub mnu101_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '案件資料及進度查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100101_1") = False Then
           Set frm100101_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100101_1", strExec) Then
            frm100101_1.Show
         End If
      'Modify  by Amy 2014/04/30 申請人查詢分兩個項目
      Case 2 '以申請人查詢(查本所客戶)
        'Mark by Amy 2025/02/03 之前因分所速度慢分兩支,因資料庫升級後不再分兩支-薛經理
'        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
'        If PUB_CheckFormExist("frm100102_1") = False Then
'           Set frm100102_1 = Nothing
'        End If
'        'end 2021/12/16
'         If CheckUse("frm100102_1", strExec) Then
'            frm100102_1.IsSearchNew = False
'            frm100102_1.Caption = "申請人查詢(查本所客戶)"
'            frm100102_1.Show
'         End If
      'Add by Amy 2014/04/30
      Case 3 '以申請人查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100102_1") = False Then
           Set frm100102_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100102_1", strExec) Then
            frm100102_1.IsSearchNew = True
            'Modify by Amy 2025/02/03 改顯示名稱 原:申請人查詢(查新客戶-含對造)
            frm100102_1.Caption = "申請人查詢（含客戶及對造）"
            frm100102_1.Show
         End If
      Case 4 '以發明人查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100120_1") = False Then
           Set frm100120_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100120_1", strExec) Then
            frm100120_1.Show
         End If
      Case 5 '以案件名稱查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100103_1") = False Then
           Set frm100103_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100103_1", strExec) Then
            frm100103_1.Show
         End If
      Case 6 '以國籍查詢代理人/申請人
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100115_1") = False Then
           Set frm100115_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100115_1", strExec) Then
            frm100115_1.Show
         End If
      Case 7 '以國別查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100116_1") = False Then
           Set frm100116_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100116_1", strExec) Then
            frm100116_1.Show
         End If
      Case 8 '申請人查詢案件變更紀錄
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100113_1") = False Then
           Set frm100113_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100113_1", strExec) Then
            frm100113_1.Show
         End If
      Case 9 '爭議案件查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100110_1") = False Then
           Set frm100110_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100110_1", strExec) Then
            frm100110_1.Show
         End If
      Case 10 '代理人案件查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100114_1") = False Then
           Set frm100114_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100114_1", strExec) Then
            frm100114_1.Show
         End If
      Case 11 '監視系統案件查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100118_1") = False Then
           Set frm100118_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100118_1", strExec) Then
            frm100118_1.Show
         End If
      Case 12 '條碼廠商號碼查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100119_1") = False Then
           Set frm100119_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100119_1", strExec) Then
            frm100119_1.Show
         End If
      Case 13 '優先權資料查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100127_1") = False Then
           Set frm100127_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100127_1", strExec) Then
            frm100127_1.Show
         End If
      Case 14 '庭期資料查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm072001") = False Then
           Set frm072001 = Nothing
        End If
        'end 2021/12/16
         'If CheckUse("frm072001", strExec) = True Then
            frm072001.Show
         'End If
      Case 15 '後金案件及結果查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm10011201_1") = False Then
           Set frm10011201_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm10011201_1", strExec) Then
            frm10011201_1.Show
         End If
      Case 16 '後金收回查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm10011202_1") = False Then
           Set frm10011202_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm10011202_1", strExec) Then
            frm10011202_1.Show
         End If
      Case 17 '關聯案件資料及正聯商標查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100108_1") = False Then
           Set frm100108_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100108_1", strExec) Then
            frm100108_1.Show
         End If
      Case 18 '專利法律案件關聯查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100101_24") = False Then
           Set frm100101_24 = Nothing
        End If
        'end 2021/12/16
         frm100101_24.Show
      'Add By Sindy 2020/5/5
      Case 19 '介紹法律所案源查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm077004") = False Then
           Set frm077004 = Nothing
        End If
        'end 2021/12/16
         frm077004.Show
   End Select
End Sub

Private Sub mnu102_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '以收/發文日查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100104_1") = False Then
           Set frm100104_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100104_1", strExec) Then
            frm100104_1.Show
         End If
      Case 2 '以收/發文量查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100105_1") = False Then
           Set frm100105_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100105_1", strExec) Then
            frm100105_1.Show
         End If
      Case 3 '收文未發文查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100107_1") = False Then
           Set frm100107_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100107_1", strExec) Then
            frm100107_1.Show
         End If
      Case 4 '以收文日查詢來函
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100109_1") = False Then
           Set frm100109_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100109_1", strExec) Then
            frm100109_1.Show
         End If
      Case 5 '以期限管制日查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100106_1") = False Then
           Set frm100106_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100106_1", strExec) Then
            frm100106_1.Show
         End If
      Case 6 '承辦人收/發文量查詢
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100111_1") = False Then
           Set frm100111_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100111_1", strExec) Then
            frm100111_1.Show
         End If
      Case 7 '發文日查詢代理人作業進度
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100117_1") = False Then
           Set frm100117_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100117_1", strExec) Then
            frm100117_1.Show
         End If
      'Add By Sindy 2013/8/28
      'Remove by Lydia 2019/10/30 若要使用,請到內專->專利公報
      'Case 8 '專利公報產業分類案件市佔分析
      '   If CheckUse("frm100133", strExec) Then
      '      frm100133.Show
      '   End If
   End Select
End Sub

Private Sub mnu103_Click(Index As Integer)
   ToolHide
   Select Case Index
      'add by nickc 2007/07/04
      Case 1 '客戶重新委任案件查詢列印
        'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
        If PUB_CheckFormExist("frm100126_1") = False Then
           Set frm100126_1 = Nothing
        End If
        'end 2021/12/16
         If CheckUse("frm100126_1", strExec) Then
            frm100126_1.Show
         End If
   End Select
End Sub

Private Sub mnu11_Click(Index As Integer)
   ToolHide
   Select Case Index
      'Add By Cheng 2002/08/21
      Case 3   '聯絡單列印及E-Mail
         frm1106.Show
      Case 4   '相關卷號
         If CheckUse("frm1103_1", strExec) = True Then
            frm1103_1.Show
         End If
      Case 5   '多國案卷號關係建立
         If CheckUse("frm1104", strExec) = True Then
            frm1104.Show
         End If
      '94.1.28 add by sonia
      Case 6   '分割案件關係維護
         If CheckUse("frm02010604_1", strExec) = True Then
            frm02010604_1.Show
         End If
      'Add By Cheng 2003/06/26
      Case 7   '撰寫信函作業
         frm090401.Show
      'add by nickc 2005/07/22
      Case 10  'CF 結餘單查詢
         If CheckUse("frm040202", strExec) = True Then
            frm040202.Show
         End If
      'add by nickc 2005/07/22
      Case 11  'CF 結餘資料維護
         If CheckUse("frm040206", strExec) = True Then
            frm040206.Show
         End If
      'add by nickc 2007/11/13
      'Remove by Lydia 2022/05/09 改放在Promoter的智權部->區主管作業
      'Case 13  '客戶特殊紀錄異動
      '   If CheckUse("frm010022", strExec) = True Then
      '      frm010022.Show
      '   End If
      'end 2022/05/09
      'add by nickc 2008/03/27
      Case 12  'CF 結餘單案件明細查詢
         If CheckUse("frm040208", strExec) = True Then
            frm040208.Show
         End If
      'Add by Morgan 2008/3/14
      Case 13  '部門別送件清單列印
         If CheckUse("frm1108", strExec) = True Then
            frm1108.Show
         End If
   End Select
End Sub
Private Sub mnu1101_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '解除期限
         If CheckUse("frm110101_1", strExec) = True Then
            frm110101_1.Show
         End If
      Case 2   '取消收文
         If CheckUse("frm110102_1", strExec) = True Then
            frm110102_1.Show
         End If
      Case 3   '閉卷
         If CheckUse("frm110103_1", strExec) = True Then
            frm110103_1.Show
         End If
   End Select
End Sub
Private Sub mnu1102_Click(Index As Integer)
   If CheckUse("frm010001", strExec) = True Then
      ToolHide
      frm010001.intChoose = 1
      frm010001.intReceiveKind = 0
      frm010001.intModifyKind = Index - 1
      Select Case Index
         Case 1
                    frm010001.Caption = "內部收文－新增"
         Case 2
                    frm010001.Caption = "內部收文－修改"
         Case 3
                    frm010001.Caption = "內部收文－查詢"
      End Select
   End If
End Sub

Private Sub mnu110401_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 1   '帳單輸入
         If CheckUse("Frmacc2150", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2150.Show
      Case 2   '抵帳單輸入
         If CheckUse("Frmacc2160", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2160.Show
      Case 3   '帳單作廢作業
         If CheckUse("Frmacc21j0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21j0.Show
      Case 4   '請款單輸入
         If CheckUse("Frmacc21h0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21h0.Show
      Case 5   '折讓輸入
         If CheckUse("Frmacc21i0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21i0.Show
         ToolShow
         tool8_enabled
         Screen.MousePointer = vbDefault
         Exit Sub
      Case 6   '請款單作廢作業
         If CheckUse("Frmacc21k0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21k0.Show
      Case 7   '請款項目資料維護
         If CheckUse("Frmacc21g0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21g0.Show
      Case 8   '美金匯率資料維護
         If CheckUse("Frmacc21m0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21m0.Show
      Case 9   '預估結匯匯率資料維護
         If CheckUse("Frmacc21o0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21o0.Show
      Case 10  '相同案件性質整批請款作業
         If CheckUse("Frmacc21p0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21p0.Show
         ToolShow
         tool3_enabled
         Screen.MousePointer = vbDefault
         Exit Sub
      Case 11  '帳單審核作業
         If CheckUse("Frmacc2153", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2153.Show
         ToolShow
         tool3_enabled
         Screen.MousePointer = vbDefault
         Exit Sub
      Case 12  '其他幣別請款匯率資料維護
         If CheckUse("Frmacc21s0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21s0.Show
      'Add by Morgan 2010/12/1
      Case 13  '客製化請款項目資料維護
         If CheckUse("Frmacc21t0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21t0.Show
   End Select
   ToolShow
   tool1_enabled
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnu110402_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 1   '國外代理人帳目查詢
         If CheckUse("Frmacc2210", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2210.Show
      Case 2   '國外案件帳目查詢
         If CheckUse("Frmacc2220", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2220.Show
      Case 3   '國外請款金額查詢
         If CheckUse("Frmacc2230", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2230.Show
      Case 4   '案件損益查詢
         If CheckUse("Frmacc2240", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2240.Show
      'Add by Amy 20130914
     Case 5 '各幣別最新請款匯率查詢
        If CheckUse("frmacc2142", strExec) = False Then
            Exit Sub
        End If
        Frmacc2142.Show
   End Select
   ToolShow
   tool3_enabled
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnu110403_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 1   'FC催款單
         If CheckUse("Frmacc2470", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2470.Show
      Case 2   'FC請款單
         If CheckUse("Frmacc2480", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2480.Show
      Case 3   '請款單整批列印
         If CheckUse("Frmacc24g0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24g0.Show
      'Add By Cheng 2002/09/03
      Case 4   '國外FC帳款明細表
         If CheckUse("Frmacc24i0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24i0.Show
      
      Case 5   '代理人帳目排名
         If CheckUse("Frmacc24b0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24b0.Show
      Case 6   'FC業務請款／收款明細表
         If CheckUse("Frmacc24c0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24c0.Show
      'Modify by Amy 2021/04/23
      Case 7 'FC專業收款明細表
         If CheckUse("Frmacc24n0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24n0.Show
      Case 8   '代理人逾期帳款分析表
         If CheckUse("Frmacc24f0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24f0.Show
      Case 9   '折讓單列印
         If CheckUse("Frmacc24h0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24h0.Show
      'Added by Lydia 2018/11/30
      Case 10 '請款單折扣案件明細
      'end 2021/04/23
         If CheckUse("Frmacc24o0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24o0.Show
   End Select
   ToolShow
   tool3_enabled
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnu1107_Click(Index As Integer)
    'Add By Cheng 2003/07/30
    ToolHide
    Select Case Index
       Case 1 '作業失誤資料維護
          If CheckUse("frm050714", strExec) = True Then
             frm050714.Show
          End If
       Case 2 '失業失誤明細表
          If CheckUse("frm040327", strExec) = True Then
             frm040327.Show
          End If
    End Select
End Sub

Private Sub mnu120401_Click(Index As Integer)
   ToolHide
   strSysKind = ""
   Select Case Index
      Case 1 '自動編號檔
         If CheckUse("frm12040103", strExec) = True Then
            frm12040103.Show
         End If
      Case 2 '系統種類對照表
         If CheckUse("frm12040104", strExec) = True Then
            frm12040104.Show
         End If
      Case 3 '員工檔維護
         If CheckUse("frm12040105", strExec) = True Then
            frm12040105.Show
         End If
      Case 4 '員工密碼資料維護
         If CheckUse("frm12040132", strExec) = True Then
            frm12040132.Show
         End If
      Case 5 '員工等級檔維護
         If CheckUse("frm12040106", strExec) = True Then
            frm12040106.Show
         End If
      Case 6 '員工權限檔維護
         If CheckUse("frm12040107", strExec) = True Then
            frm12040107.Show
         End If
      'ADD BY TONI 2008/10/03
      Case 7 '員工跨部門權限檔維護
         If CheckUse("frm12040148", strExec) = True Then
            frm12040148.Show
         End If
      'END 2008/10/03
      Case 8 '案件來源對照檔
         If CheckUse("frm12040108", strExec) = True Then
            frm12040108.Show
         End If
      Case 9 '案件性質對照檔
         If CheckUse("frm12040109", strExec) = True Then
            frm12040109.Show
         End If
      Case 10 '國家基本檔維護
         If CheckUse("frm12040101", strExec) = True Then
            frm12040101.Show
         End If
      Case 11 '專利商標種類/著作權登記項目對照表
         If CheckUse("frm12040110", strExec) = True Then
            frm12040110.Show
         End If
      Case 12 '解除期限原因檔
         If CheckUse("frm12040111", strExec) = True Then
            frm12040111.Show
         End If
      Case 13 '資料刪除記錄檔
         If CheckUse("frm12040112_1", strExec) = True Then
            frm12040112_1.Show
         End If
      Case 14 '專利年費資料檔
         If CheckUse("frm12040113", strExec) = True Then
            frm12040113.Show
         End If
      Case 15 '專利案件基本資料維護
         If CheckUse("frm050701", strExec) = True Then
            frm050701.Show
         End If
      Case 16 '商標基本資料維護
         If CheckUse("frm020501", strExec) = True Then
            frm020501.SetSystem 0 'Add By Sindy 2013/1/22
            frm020501.Show
         End If
      Case 17 '法務案件基本資料維護
         If CheckUse("frm075002", strExec) = True Then
            frm075002.Show
         End If
      Case 18 '顧問案件資料維護
         If CheckUse("frm075006", strExec) = True Then
            frm075006.Show
         End If
      Case 20 '案件進度檔資料維護
         If CheckUse("frm075004_1", strExec) = True Then
            frm075004_1.Show
         End If
      Case 21 '下一程序資料
         If CheckUse("frm075007_1", strExec) = True Then
            frm075007_1.Show
         End If
      Case 22 '個人目標資料檔
         If CheckUse("frm12040120", strExec) = True Then
            frm12040120.Show
         End If
      Case 23 '案件國家收費表維護
         If CheckUse("frm12040102", strExec) = True Then
            frm12040102.Show
         End If
      Case 24 '員工群組檔維護
         If CheckUse("frm12040136", strExec) = True Then
            frm12040136.Show
         End If
      Case 25 '工作天維護
         If CheckUse("frm12040137", strExec) = True Then
            frm12040137.Show
         End If
      Case 26 '客戶基本資料維護
         If CheckUse("frm140401", strExec) = True Then
            frm140401.Show
         End If
      Case 27 '客戶發明人資料維護
         If CheckUse("frm050709", strExec) = True Then
            frm050709.Show
         End If
      Case 28 '客戶變更名稱作業
         If CheckUse("frm140101", strExec) = True Then
            frm140101.Show
         End If
      Case 29 '國外代理人資料
         If CheckUse("frm050705", strExec) = True Then
            frm050705.Show
         End If
      Case 30 '代理人變更名稱作業
         If CheckUse("frm140103", strExec) = True Then
            frm140103.Show
         End If
      Case 31 '特殊專利商標資料維護
         If CheckUse("frm12040146", strExec) = True Then
            frm12040146.Show
         End If
      'add by nickc 2007/10/17
      Case 32 '系統特殊設定
         If CheckUse("frm050716", strExec) = True Then
            frm050716.Show
         End If
      'Add by morgan 2009/2/18
      Case 33 '郵件排程維護
         'Modified by Morgan 2012/1/3 改用新程式
         'If CheckUse("frm140410", strExec) = True Then
         '   frm140410.Show
         'End If
         If CheckUse("frm140410_1", strExec) = True Then
            frm140410_1.Show
         End If
'      'Add By Sindy 2010/6/29
'      Case 34 '接洽記錄單查詢及列印
'         'Memo by Lydia 2021/05/18 更名為「自動收文接洽單查詢/列印」
'         'Modify By Sindy 2023/1/6 更名為「電子收文接洽單查詢」
'         'If CheckUse("frm12040152", strExec) = True Then
'            frm12040152.Show
'         'End If
      'Add By Sindy 2012/3/15
      Case 35 '不得代理案件之客戶或代理人資料維護
         'If CheckUse("frm12040154", strExec) = True Then
            frm12040154.Show
         'End If
      'Add By Sindy 2012/4/10
      Case 36 '非本所實質客戶資料維護
         If CheckUse("frm12040155", strExec) = True Then
            frm12040155.Show
         End If
      'Added by Morgan 2012/4/24
      Case 37 'LEDES基本資料維護
         If CheckUse("frm12040156", strExec) = True Then
            frm12040156.Show
         End If
      'Add By Sindy 2012/11/21
      Case 38 '特殊客戶/代理人收文費用維護
         If CheckUse("frm12040157", strExec) = True Then
            frm12040157.Show
         End If
      'Add By Amy 2013/03/20
      Case 39 '程式公告維護
         If CheckUse("frm140413", strExec) = True Then
            frm140413.Show
         End If
      'Added By Morgan 2013/6/28
      Case 40 '造字與UnidCode字對照表
         If CheckUse("frm12040158", strExec) = True Then
            frm12040158.Show
         End If
      'Added By Sindy 2015/1/6
      Case 41 '案件表單簽核人員設定
         If CheckUse("frm140414", strExec) = True Then
            frm140414.Show
         End If
      'Added by Lydia 2016/11/09
      Case 42  '各項指示分類維護
         If CheckUse("frm140415", strExec) = True Then
            frm140415.Show
         End If
      Case 43  '國外部關聯企業分類維護
         If CheckUse("frm140416", strExec) = True Then
            frm140416.Show
         End If
      'Added by Lydia 2016/11/09
      Case 43  '國外部關聯企業分類維護
      'Added by Lydia 2015/06/12
      'Modified by Lydia 2016/11/09 index 42 =>44
      Case 44 '國內收據點數分配輸入
         'Memo by Lydia 2020/04/20 因為法務改用工作點數frm071021, 所以這支直接隱藏
         If CheckUse("Frmacc21h5", strExec) = True Then
            Frmacc21h5.Show
         End If
'      'Added by Lydia 2017/03/24
'      Case 45 '委任契約書用印記錄查詢
'         If CheckUse("frm140417", strExec) = True Then
'            frm140417.Show
'         End If
      'Add By Sindy 2018/3/15
      Case 46 '核判表設定作業
         If CheckUse("frm12040161", strExec) = True Then
            frm12040161.Show
         End If
'      'Add by Amy 2023/08/15
'      Case 47 '查詢特殊置換字對照表
'         If CheckUse("frm12040162", strExec) = True Then
'            frm12040162.Show
'         End If
      'Add By Sindy 2023/8/24
      Case 48 '電子報特殊名單維護
         If CheckUse("frm030617", strExec) = True Then
            frm030617.m_WorkType = "M"
            frm030617.Show
         End If
   End Select
End Sub

Private Sub mnu13_Click(Index As Integer)
   Select Case Index
      Case 1
      Case 2
      Case 3
      Case 4
         'edit by nickc 2007/02/09 不用 dll 了
         'objPublicData.ShowAbout
   End Select
End Sub

Private Sub mnu12040101_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1
         strSysKind = "P"
      Case 2
         strSysKind = "CFP"
      Case 3
         strSysKind = "FCP"
   End Select
   frm050701.Show
End Sub

Private Sub mnu12040117_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '服務業務基本資料維護
         If CheckUse("frm050702", strExec) = True Then
            'strSysKind = "PS','FG','CPS"
            strSysKind = "'PS','FG','CPS'"
            frm050702.Show
         End If
      Case 2 '服務業務基本資料維護 (條碼)
         If CheckUse("frm02050201", strExec) = True Then
            frm02050201.Show
         End If
      Case 3 '服務業務基本資料維護 (監視系統)
         If CheckUse("frm02050202", strExec) = True Then
            frm02050202.Show
         End If
      Case 4 '服務業務基本資料維護 (網域)
         If CheckUse("frm02050203", strExec) = True Then
            frm02050203.Show
         End If
      Case 5 '服務業務基本資料維護 (著作權)
         If CheckUse("frm02050204", strExec) = True Then
            frm02050204.Show
         End If
      Case 6 '服務業務基本資料維護 (其它業務)
         If CheckUse("frm02050205", strExec) = True Then
            frm02050205.Show
         End If
   End Select
End Sub

Private Sub mnu120501_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '來文資料稽核表
         If CheckUse("frm12040121", strExec) = True Then
            frm12040121.Show
         End If
      Case 2   '規費資料稽核表
         If CheckUse("frm12040122", strExec) = True Then
            frm12040122.Show
         End If
      Case 3   '案件自動核准作業
         If CheckUse("frm12040123", strExec) = True Then
            frm12040123.Show
         End If
        'Add By Cheng 2003/02/19
      Case 5 '新客戶清單
         If CheckUse("frm12040141", strExec) = True Then
            frm12040141.Show
         End If
        'Add By Cheng 2003/02/24
      Case 6 '收文未發文明細表
         If CheckUse("frm12040142", strExec) = True Then
            frm12040142.Show
         End If
        'Add By Cheng 2003/08/04
      Case 7 '逾期未處理案件明細表
         If CheckUse("frm12040143", strExec) = True Then
            frm12040143.Show
         End If
        'Add By Cheng 2004/01/06
      Case 8 '本所期限工作天推算
         If CheckUse("frm12040145", strExec) = True Then
            frm12040145.Show
         End If
   End Select
End Sub

Private Sub mnu120601_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '客戶/代理人改號作業
         If CheckUse("frm12040125", strExec) = True Then
            frm12040125.Show
         End If
      Case 2   '案件改號作業
         If CheckUse("frm12040126", strExec) = True Then
            frm12040126.Show
         End If
      Case 4   '刪除記錄統計表
         If CheckUse("frm12040128", strExec) = True Then
            frm12040128.Show
         End If
      Case 5   '智權人員客戶轉移作業
         If CheckUse("frm12040129", strExec) = True Then
            frm12040129.Show
         End If
      Case 6   '智權人員調區作業
         If CheckUse("frm12040130", strExec) = True Then
            frm12040130.Show
         End If
      Case 10  '客戶案件總簿
         If CheckUse("frm050317", strExec) = True Then
            frm0503171.Show 'Modify by Amy 2014/07/21
         End If
      Case 11  '代理人案件總簿
         If CheckUse("frm050316", strExec) = True Then
            frm050316.Show
         End If
      Case 12  '客戶/代理人名冊、地址條列印
         If CheckUse("frm12040138", strExec) = True Then
            frm12040138.Show
         End If
      'Add By Cheng 2002/05/14
      Case 13  '國內客戶名條
         If CheckUse("frm12040131", strExec) = True Then
            frm12040131.Show
         End If
      'Add By Cheng 2002/05/27
      Case 14  '閉卷清單
         If CheckUse("frm12040139", strExec) = True Then
            frm12040139.Show
         End If
      'Add By Cheng 2002/05/27
      Case 15  '分所案號檢核表
         If CheckUse("frm12040140", strExec) = True Then
            frm12040140.Show
         End If
      'Add By Cheng 2002/05/27
      Case 16 '國外客戶/代理人地址條列印
         If CheckUse("frm12040144", strExec) = True Then
            frm12040144.Show
         End If
      Case 17 '智權人員客戶名冊 (依最後收文日期區間)
         If CheckUse("frm12040150", strExec) = True Then
            frm12040150.Show
         End If
'      'Add By Sindy 2010/01/07
'      Case 18  '員工查詢印表記錄資料查詢
'         If CheckUse("frm050207", strExec) = True Then
'            StrStartSystemByNick = GetSystemKindByNick
'            frm050207.Show
'         End If
      'Add By Sindy 2012/7/24
      Case 19  '客戶案件預算預估表
         If CheckUse("frm210138", strExec) = True Then
            frm210138.Show
         End If
     'Add By Amy 2013/04/02
      Case 20 '開拓名單轉入國內潛在客戶作業
         If CheckUse("frm210140", strExec) = True Then
            frm210140.Show
         End If
      
      'Added by Morgan 2013/7/18
      Case 21 '新建指紋整批匯入
          If CheckUse("frm160013", strExec) = False Then
              Exit Sub
          End If
          frm160013.Show
      
      Case 22 '員工指紋卡片資料
          If CheckUse("frm160014", strExec) = False Then
              Exit Sub
          End If
          frm160014.Show
      
      Case 23 '考勤機設定
          If CheckUse("frm160015", strExec) = False Then
              Exit Sub
          End If
          frm160015.Show
          
      Case 24  '客戶案件總簿
         'Add by Amy 2022/01/17 改Form2.0 Word 目前無英文及日文版,故由電腦中心產生
         frm0503171Old.Show
         'Mark by Amy 2017/07/25 暫存檔加欄位且程式已不再維護,故下架
'         If CheckUse("frm050317", strExec) = True Then
'            frm050317.Show 'Modify by Amy 2014/07/21
'         End If
      'Add By Sindy 2014/10/27
      Case 25 '更換FC代理人作業
         If CheckUse("frm110104_1", strExec) = True Then
            frm110104_1.Show
         End If
      'Add By Sindy 2018/4/12
      Case 26 '國外部人員離職修改資料
         If CheckUse("frm140118", strExec) = True Then
            frm140118.Show
         End If
      'Add By Sindy 2025/5/9
      Case 27  '設定颱風假作業
         If CheckUse("frm140422", strExec) = True Then
            frm140422.Show
         End If
      'Added by Lydia 2024/02/01
      Case 28 '外專案件清單Excel
         If CheckUse("frm060511", strExec) = True Then
            frm060511.Show
         End If
   End Select
End Sub

Public Sub ToolShow()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub


'*************************************************
'  工具列按鈕失效設定1
'
'*************************************************
Public Sub tool1_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = True
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = True
   Toolbar1.Buttons.Item(9).Enabled = True
   Toolbar1.Buttons.Item(10).Enabled = True
   Toolbar1.Buttons.Item(13).Enabled = True
   Toolbar1.Buttons.Item(14).Enabled = True
   Toolbar1.Buttons.Item(15).Enabled = True
   Toolbar1.Buttons.Item(16).Enabled = True
End Sub


'*************************************************
'  工具列按鈕失效設定2
'
'*************************************************
Public Sub tool2_enabled()
   Toolbar1.Buttons.Item(1).Enabled = False
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = True
   Toolbar1.Buttons.Item(7).Enabled = True
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub


'*************************************************
'  工具列按鈕失效設定3
'
'*************************************************
Public Sub tool3_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub


'*************************************************
'  工具列按鈕失效設定4
'
'*************************************************
Public Sub tool4_enabled()
   Toolbar1.Buttons.Item(1).Enabled = False
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  工具列按鈕失效設定5
'
'*************************************************
Public Sub tool5_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = True
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = True
   Toolbar1.Buttons.Item(9).Enabled = True
   Toolbar1.Buttons.Item(10).Enabled = True
   Toolbar1.Buttons.Item(13).Enabled = True
   Toolbar1.Buttons.Item(14).Enabled = True
   Toolbar1.Buttons.Item(15).Enabled = True
   Toolbar1.Buttons.Item(16).Enabled = True
End Sub

'*************************************************
'  工具列按鈕失效設定6
'
'*************************************************
Public Sub tool6_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = True
   Toolbar1.Buttons.Item(14).Enabled = True
   Toolbar1.Buttons.Item(15).Enabled = True
   Toolbar1.Buttons.Item(16).Enabled = True
End Sub

'*************************************************
'  工具列按鈕失效設定7
'
'*************************************************
Public Sub tool7_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  工具列按鈕失效設定8
'
'*************************************************
Public Sub tool8_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = True
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = True
   Toolbar1.Buttons.Item(14).Enabled = True
   Toolbar1.Buttons.Item(15).Enabled = True
   Toolbar1.Buttons.Item(16).Enabled = True
End Sub

'*************************************************
'  工具列按鈕失效設定9
'
'*************************************************
Public Sub tool9_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = True
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  工具列按鈕失效設定10
'
'*************************************************
Public Sub tool10_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = True
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = True
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  工具列按鈕失效設定11
'
'*************************************************
Public Sub tool11_enabled()
   Toolbar1.Buttons.Item(1).Enabled = False
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = True
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  工具列按鈕失效設定12
'
'*************************************************
Public Sub tool12_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = True
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = True
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  工具列按鈕失效設定13
'
'*************************************************
Public Sub tool13_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = True
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'查詢作業
Private Sub mnu120701_Click(Index As Integer)
   ToolHide
   strSysKind = ""
   Select Case Index
      'Add By Sindy 2010/01/07
      Case 1  '員工查詢印表記錄資料查詢
         If CheckUse("frm050207", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm050207.Show
         End If
      'Added by Lydia 2017/03/24
      Case 2 '委任契約書用印記錄查詢
         If CheckUse("frm140417", strExec) = True Then
            frm140417.Show
         End If
      'Add By Sindy 2010/6/29
      Case 3 '接洽記錄單查詢及列印
         'Memo by Lydia 2021/05/18 更名為「自動收文接洽單查詢/列印」
         'Modify By Sindy 2023/1/6 更名為「電子收文接洽單查詢」
         'If CheckUse("frm12040152", strExec) = True Then
            frm12040152.Show
         'End If
      'Add by Amy 2023/08/15
      Case 4 '查詢特殊置換字對照表
         If CheckUse("frm12040162", strExec) = True Then
            frm12040162.Show
         End If
   End Select
End Sub

Private Sub mnu15_Click(Index As Integer)
ToolHide
End Sub

Public Sub ToolHide()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
End Sub

'Add by Morgan 2007/12/12
Private Sub mnu22_Click(Index As Integer)
   Select Case Index
      Case 0 '潛在客戶資料維護
         If CheckUse("frm140402", strExec) = True Then
            frm140402.Show
         End If
      
      Case 1 '客戶/代理人聯絡人資料維護
         If CheckUse("frm140403", strExec) = True Then
            frm140403.Show
         End If
         
      Case 2 '往來記錄資料維護
         If CheckUse("frm140404", strExec) = True Then
            frm140404.Show
         End If
      'Add by Morgan 2009/2/6
      'Memo by Lydia 2022/01/12 經過確認往來記錄代碼不存在，並且最後資料在2008年；所以先隱藏功能選單
      Case 3 '交換名片紀錄維護
         If CheckUse("frm140411", strExec) = True Then
            frm140411.Show
         End If
      'Add by Morgan 2008/2/25
      Case 4 '代理人互惠資料維護
         If CheckUse("frm140405", strExec) = True Then
            frm140405.Show
         End If
      'Add by Morgan 2008/4/15
      Case 5 '互惠代理人案件統計表
         If CheckUse("frm050408", strExec) = True Then
            frm050408.Show
         End If
      'Add by TONI 2008/12/4
      Case 6 '潛在客戶名條列印
         If CheckUse("frm140409", strExec) = True Then
            frm140409.Show
         End If
      'Add by TONI 2008/12/4
      Case 7 '潛在客戶資料查詢
         If CheckUse("frm140407", strExec) = True Then
            frm140407.Show
         End If
      'Add by TONI 2008/12/4
      Case 8 '往來記錄資料查詢
         If CheckUse("frm140408", strExec) = True Then
            frm140408.Show
         End If
      'Add by Sindy 2019/12/27
      Case 9   '往來記錄統計
         If CheckUse("frm140420_1", strExec) = True Then
            frm140420_1.Show
         End If
      'Add by Sindy 2010/9/2
      Case 10 '國外部新客戶/代理人查詢
         If CheckUse("frm140412", strExec) = True Then
            frm140412.Show
         End If
      'Added by Lydia 2023/06/17
      Case 11  '不得宣傳客戶名稱資料查詢
         If frm100136.ChkUseRight = True Then
            frm100136.Show
         End If
   End Select
End Sub

'Added by Morgan 2015/6/3
Private Sub mnu16_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 0   '印表機設定
         frm880011.bolAppOnly = True
         frm880011.Show 1
      
      Case 1   '報表紙張格式設定
         frm880013.Show vbModal
   End Select
End Sub

Private Sub mnu23_Click(Index As Integer)
Dim nFrm As Form
   
   Select Case Index
      Case 1 '會議室預約作業
         frm140112.Show
      Case 3 '專利處研討會
         frm140113.Show
      'Add By Sindy 2012/9/5
      Case 4 '客戶端平台帳號管理作業
         frm140114.Show
      'Add By Sindy 2017/12/25 + 信箱分信紀錄查詢
      Case 7
         '檢查表單是否已開啟，若是，則關閉
         For Each nFrm In Forms
            If StrComp(nFrm.Name, "frm06010613", vbTextCompare) = 0 Then
               Unload frm06010613
            End If
         Next
         frm06010613.m_WorkType = "0" '信箱主檔
         frm06010613.Show
      Case 8 '圖書借閱資料查詢 Add by Amy 2017/01/25
        frm010035.Show
        'Add by Amy 2017/02/03 判斷是否有圖書借閱記錄需簽核
        If GetLoanRecordApply = True Then
            frm010035.bolLoanRecordApply = True
            Call frm010035.cmdLoanRecord_Click
         End If
     'Added by Lydia 2020/01/15
     Case 9 '行事曆提醒通知
         frm060209.m_Role = "F41"
         frm060209.Show
     'Add by Amy 2024/01/22
     Case 10 '風險檢查對象資料維護
         frm12040163.Show
   End Select
End Sub

Private Sub mnu2321_Click(Index As Integer)
   Select Case Index
      Case 1 '目前表單
         frm180101.Show
      Case 2 '職代/簽核主管代填表單
         frm180103.Show
      'Add By Sindy 2013/7/11
      Case 3 '打卡異常個人處理
         frm180105.Show
   End Select
End Sub

Private Sub mnu2322_Click(Index As Integer)
   Select Case Index
      Case 1 '簽核作業
         frm180201.Show
      Case 2 '每月出缺勤統計確認
'         frm160201.intChoose = 1
'         frm160201.Hide
'         Call frm160201.cmdOK_Click(0)
         frm180203_1.Show
      Case 3 '員工個人資料明細確認
         frm160102.intChoose = 1
         frm160102.Hide
         Call frm160102.cmdok_Click(0)
      Case 4 '簽核人員異動作業
         frm180104.Show
      'Add By Sindy 2013/7/11
      Case 5 '打卡異常主管處理
         frm180204.Show
   End Select
End Sub

Private Sub mnu2323_Click(Index As Integer)
   Select Case Index
      Case 1 '近日請假公佈欄
         frm180302.Show
      Case 2 '出缺勤查詢
         frm180301.Show
      Case 3 '職代/審核主管關聯查詢
         frm180403.Show
      Case 4 '打卡資料查詢
         frm180303.Show
   End Select
End Sub

Private Sub mnu99_Click(Index As Integer)
Dim frm As Form
    
    For Each frm In Forms
        If frm.Name <> "mdiMain" Then
            If frm.Name = mnu99(Index).Tag Then
                '將子視窗排在頂層
                frm.ZOrder (0)
                Exit For
            End If
        End If
    Next
End Sub

Private Sub mnuChUser_Click()
   frmChgUser.Show
End Sub

Private Sub mnuDML_Click(Index As Integer)
    frmDML.Show
End Sub

Private Sub Timer1_Timer()
   'Added by Morgan 2024/8/8 定時執行一次語法以確保跨網段連線時網路不會被切斷
   Static dtNow As Date
   
   If Now > dtNow Then
      dtNow = DateAdd("n", cntAutoQueryInterval, Now)
      ClsLawReadRstMsg 1, "select * from dual"
   End If
   'end 2024/8/8
   
'Modify By Cheng 2002/11/22
'Dim tmpformnick As Form
'Dim tmpFormI As Integer
'Dim tmpFormJ As Integer
'tmpFormI = 0
'For Each tmpformnick In Forms
'   tmpFormI = tmpFormI + 1
'Next
'   If tmpFormI > 1 Then
'      If mnuTitle(0).Enabled = True Then
'         mnuTitle(0).Enabled = False
'         mnuTitle(10).Enabled = False
'         mnuTitle(11).Enabled = False
'         mnuTitle(12).Enabled = False
'         mnuTitle(15).Enabled = False
'      End If
'   Else
'      If tmpFormI = 1 Then
'         If mnuTitle(0).Enabled = False Then
'            mnuTitle(0).Enabled = True
'            mnuTitle(10).Enabled = True
'            mnuTitle(11).Enabled = True
'            mnuTitle(12).Enabled = True
'            mnuTitle(15).Enabled = True
'         End If
'      End If
'   End If

   'Add By Sindy 2023/11/8
   If InStr(pub_HostName, "97038") > 0 Then '測試中
   If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then '測試資料庫
      ToolShow
      Me.TxtTestDB.Visible = True
   Else
      Me.TxtTestDB.Visible = False
   End If
   End If
   '2023/11/8 END

'Add By Cheng 2002/11/22
Dim frm As Form
Dim intfrm10 As Integer
Dim intFrmacc2 As Integer

'控制共同查詢
intfrm10 = 0
For Each frm In Forms
    If Left(frm.Name, 5) = "frm10" Then
        intfrm10 = 1
        Exit For
    End If
Next
If intfrm10 = 1 Then
    If mnuTitle(10).Enabled = True Then mnuTitle(10).Enabled = False
Else
    If mnuTitle(10).Enabled = False Then mnuTitle(10).Enabled = True
End If
'控制代理人帳目
intFrmacc2 = 0
For Each frm In Forms
    If Left(frm.Name, 7) = "Frmacc2" Then
        intFrmacc2 = 1
        Exit For
    End If
Next
If intFrmacc2 = 1 Then
'    If mnu11(6).Enabled = True Then mnu11(6).Enabled = False
    'Modified by Lydia 2018/11/22 控制代理人帳目
    'If mnu11(8).Enabled = True Then mnu11(8).Enabled = False
    If mnu11(9).Enabled = True Then mnu11(9).Enabled = False
Else
'    If mnu11(6).Enabled = False Then mnu11(6).Enabled = True
    'Modified by Lydia 2018/11/22 控制代理人帳目
    'If mnu11(8).Enabled = False Then mnu11(8).Enabled = True
    If mnu11(9).Enabled = False Then mnu11(9).Enabled = True
End If
'Add By Cheng 2003/12/19
'控制"視窗"Menu
MenuForFormControl
'End
StatusBar1.Panels.Item(4).Text = time

End Sub

Private Sub Timer2_Timer()
    '若登入失敗
    If pub_str_LoginSucceeded <> "1" Then
        Me.Timer1.Interval = 0
    '若登入成功
    Else
        Me.Timer1.Interval = 100
        Me.Timer2.Interval = 0
        MDIForm_Load
        Me.Show
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.KEY
      Case "function1"
         KeyEnter vbKeyEscape
      Case "function2"
         KeyEnter vbKeyF2
      Case "function3"
         KeyEnter vbKeyF3
      Case "function4"
         KeyEnter vbKeyF9
      Case "function12"
         KeyEnter vbKeyF10
      Case "function5"
         KeyEnter vbKeyF5
      Case "function6"
         KeyEnter vbKeyF4
      Case "function7"
         KeyEnter vbKeyF7
      Case "function8"
         KeyEnter vbKeyHome
      Case "function9"
         KeyEnter vbKeyPageUp
      Case "function10"
         KeyEnter vbKeyPageDown
      Case "function11"
         KeyEnter vbKeyEnd
   End Select
End Sub

Private Sub MenuForFormControl()
Dim frm As Form
Dim ii As Integer
Dim objMnu99 As Menu
Dim intMaxIndex As Integer
Dim blnFormNameMatch As Boolean

On Error Resume Next
'若無任何子視窗
If Forms.Count <= 1 Then
    If Me.mnuTitle(99).Enabled = True Then
        Me.mnuTitle(99).Enabled = False
        For Each objMnu99 In Me.mnu99
            If objMnu99.Index = 0 Then
                Me.mnu99(objMnu99.Index).Tag = ""
            Else
                Unload Me.mnu99(objMnu99.Index)
            End If
        Next
    End If
'若有子視窗
Else
    If Me.mnuTitle(99).Enabled = False Then Me.mnuTitle(99).Enabled = True
    '若子視窗數與視窗menu數不同
    If Forms.Count - 1 <> Me.mnu99.Count Then
        For Each frm In Forms
            'edit by nickc 2007/09/27
            'If frm.Name <> "mdiMain" Then
            If frm.Name <> "mdiMain" And frm.Caption <> "" Then
                '若子視窗被隱藏
                If frm.Visible = False Then
                    For Each objMnu99 In Me.mnu99
                        If frm.Name = Me.mnu99(objMnu99.Index).Tag Then
                            If Me.mnu99(objMnu99.Index).Enabled = True Then Me.mnu99(objMnu99.Index).Enabled = False
                            Exit For
                        End If
                    Next
                '若子視窗未被隱藏
                Else
                    blnFormNameMatch = False
                    For Each objMnu99 In Me.mnu99
                        If frm.Name = Me.mnu99(objMnu99.Index).Tag Then
                            If Me.mnu99(objMnu99.Index).Enabled = False Then Me.mnu99(objMnu99.Index).Enabled = True
                            blnFormNameMatch = True
                            Exit For
                        End If
                    Next
                    '若子視窗未出現在視窗Menu上
                    If blnFormNameMatch = False Then
                        For Each objMnu99 In mnu99
                            intMaxIndex = Me.mnu99(objMnu99.Index).Index
                        Next
                        Load Me.mnu99(intMaxIndex + 1)
                        Me.mnu99(intMaxIndex + 1).Caption = frm.Caption
                        Me.mnu99(intMaxIndex + 1).Tag = frm.Name
                        Me.mnu99(intMaxIndex + 1).Enabled = True
                        Exit For
                    End If
                End If
            End If
        Next
        For Each objMnu99 In Me.mnu99
            blnFormNameMatch = False
            For Each frm In Forms
                'edit by nickc 2007/09/27
                'If frm.Name <> "mdiMain" Then
                If frm.Name <> "mdiMain" And frm.Caption <> "" Then
                    blnFormNameMatch = False
                    If frm.Name = Me.mnu99(objMnu99.Index).Tag Then
                        blnFormNameMatch = True
                        Exit For
                    End If
                End If
            Next
            If blnFormNameMatch = False Then
                Exit For
            End If
        Next
        '若視窗Menu相對應的子視窗不存在
        If blnFormNameMatch = False Then
            If objMnu99.Index = 0 Then
                For Each objMnu99 In Me.mnu99
                    If objMnu99.Index <> 0 Then
                        Unload Me.mnu99(objMnu99.Index)
                    End If
                Next
                ii = 0
                For Each frm In Forms
                    'edit by nickc 2007/09/27
                    'If frm.Name <> "mdiMain" Then
                    If frm.Name <> "mdiMain" And frm.Caption <> "" Then
                        If ii = 0 Then
                            Me.mnu99(ii).Caption = frm.Caption
                            Me.mnu99(ii).Tag = frm.Name
                            If Me.mnu99(ii).Enabled = False Then Me.mnu99(ii).Enabled = True
                        Else
                            Load Me.mnu99(ii)
                            Me.mnu99(ii).Caption = frm.Caption
                            Me.mnu99(ii).Tag = frm.Name
                            Me.mnu99(ii).Enabled = True
                        End If
                        ii = ii + 1
                    End If
                Next
            Else
                Unload Me.mnu99(objMnu99.Index)
            End If
        End If
    '若子視窗數與視窗menu數皆為1
    ElseIf Forms.Count - 1 = 1 And Me.mnu99.Count = 1 Then
        For Each frm In Forms
            'edit by nickc 2007/09/27
            'If frm.Name <> "mdiMain" Then
            If frm.Name <> "mdiMain" And frm.Caption <> "" Then
                If frm.Name <> Me.mnu99(0).Tag Then
                    For Each objMnu99 In mnu99
                        intMaxIndex = Me.mnu99(0).Index
                    Next
                    If Me.mnu99(0).Enabled = False Then Me.mnu99(0).Enabled = True
                    Me.mnu99(0).Caption = frm.Caption
                    Me.mnu99(0).Tag = frm.Name
                    Exit For
                End If
            End If
        Next
    End If
End If
'若有子視窗
If Forms.Count - 1 >= 1 Then
    For Each objMnu99 In Me.mnu99
        If mdiMain.ActiveForm.Name = Me.mnu99(objMnu99.Index).Tag Then
            If Me.mnu99(objMnu99.Index).Checked = False Then Me.mnu99(objMnu99.Index).Checked = True
        Else
            If Me.mnu99(objMnu99.Index).Checked = True Then Me.mnu99(objMnu99.Index).Checked = False
        End If
    Next
End If
End Sub

Private Sub mnuPopItem_Click(Index As Integer)
   'Modified by Morgan 2021/4/26
   'frm140112.ShowNextForm Index
   frm140112.Timer2.Enabled = True
   frm140112.Timer2.Tag = Index
End Sub

'Add By Sindy 2012/3/1 依系統各作業的需求呼叫Form
Public Function SysCallSpecForm(oForm As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String)
   Select Case UCase(oForm)
      Case UCase("frm010001")
         Select Case strCP01
'            Case "CFT", "T", "TF", "FCT"
'               If strCP01 = "T" Or strCP01 = "TF" Then
'                  Call CallFormData(frm020501, "frm020501", strCP01, strCP02, strCP03, strCP04)
'               End If
'            Case "TB"
'               Call CallFormData(frm02050201, "frm02050201", strCP01, strCP02, strCP03, strCP04)
'            Case "TM"
'               Call CallFormData(frm02050202, "frm02050202", strCP01, strCP02, strCP03, strCP04)
'            Case "TD"
'               Call CallFormData(frm02050203, "frm02050203", strCP01, strCP02, strCP03, strCP04)
'            Case "TC"
'               Call CallFormData(frm02050204, "frm02050204", strCP01, strCP02, strCP03, strCP04)
'            Case Else
'               Call CallFormData(frm02050205, "frm02050205", strCP01, strCP02, strCP03, strCP04)
         End Select
   End Select
End Function
'Added by Morgan 2016/1/22
Private Sub mnu2306_Click(Index As Integer)
   Select Case Index
      Case 1 '員工薪資明細
         If PUB_SalaryEnabled Then
            frm170236.Show: Exit Sub
         Else
            frm170107.setNextForm "frm170236"
         End If
      Case 2 '勞保/健保/勞退金明細
         If PUB_SalaryEnabled Then
            frm170237.Show: Exit Sub
         Else
            frm170107.setNextForm "frm170237"
         End If
      Case 3 '年度各項所得明細
         If PUB_SalaryEnabled Then
            frm170238.Show: Exit Sub
         Else
            frm170107.setNextForm "frm170238"
         End If
      Case 4 '年終獎金明細
         If PUB_SalaryEnabled Then
            frm170239.Show: Exit Sub
         Else
            frm170107.setNextForm "frm170239"
         End If
      Case 5 '薪資查詢密碼修改
         frm170107.setNextForm "frm170108"
'      Case 6 '薪資查詢離線時間設定修改 (先不做)
   End Select
   
   frm170107.Show
End Sub

'Added by Morgan 2016/1/22
'薪資畫面計時器:60秒
Private Sub tmrSalary_Timer()
   tmrSalary.Tag = Val(tmrSalary.Tag) + 1
   If Val(tmrSalary.Tag) > 60 Then
      tmrSalary.Enabled = False
      Pub_CloseSalaryQueryForm
   End If
End Sub

'Added by Morgan 2020/1/17
'以名稱取得表單--通用不可刪
Public Function GetForm(pFormName As String) As Form
   Select Case pFormName
   '新增專案會用到的Form
   Case "frm090801"
         Set GetForm = frm090801
   Case "frm090801_11"
         Set GetForm = frm090801_11
   'Add By Sindy 2024/11/5
   Case "frm090801_Q"
         Set GetForm = frm090801_Q
         '2024/11/5 END
   'Add by Amy 2024/01/22
   Case "frm090801_14" '接洽單-對造
         Set GetForm = frm090801_14
   'Add By Sindy 2020/5/29
   Case "frm180301"
         Set GetForm = frm180301
   '2020/5/29 END
   'Added by Morgan 2021/1/18
   Case "frm090401_1"
         Set GetForm = frm090401_1
   'Add by Amy 2023/09/21 共同查詢用
   Case "frm083014" '地址條列印
         Set GetForm = frm083014
   Case "frm880022" '寄發信函-往來記錄
         Set GetForm = frm880022
   Case "frm100101_10" '代理人資料
         Set GetForm = frm100101_10
   Case "frm100101_11" '申請人資料
         Set GetForm = frm100101_11
   Case "frm100101_14" '國外潛在客戶資料
         Set GetForm = frm100101_14
   Case "frm100101_15" '往來記錄資料
         Set GetForm = frm100101_15
   Case "frm100101_17" '國外部聯絡人/接洽人資料
         Set GetForm = frm100101_17
   Case "frm100101_18" '[非]國外部聯絡人/接洽人資料
         Set GetForm = frm100101_18
   Case "frm100101_21" '國內潛在客戶資料
         Set GetForm = frm100101_21
   Case "frm100101_27" '客戶端平台帳號資料"
         Set GetForm = frm100101_27
   Case "frm100101_22" '投資法務開拓客戶資料
         Set GetForm = frm100101_22
   Case "frm100101_25" '代理案件之客戶或代理人
         Set GetForm = frm100101_25
   Case "frm100101_h" '專利相關案
         Set GetForm = frm100101_h
   Case "frm100102_2" '申請人案件資料
         Set GetForm = frm100102_2
   Case "frm100102_4" '相關多申請人
         Set GetForm = frm100102_4
   Case "frm100114_2" '代理人案件資料
         Set GetForm = frm100114_2
   Case "frm100114_6" '案件統計
         Set GetForm = frm100114_6
   Case "frm210145" '申請人一個月內寄送資料
         Set GetForm = frm210145
   'Add by Amy 2024/07/17
   Case "frm12040154" '不得代理案件之客戶或代理人資料維護
         Set GetForm = frm12040154
   'Add by Amy 2024/09/16
   Case "frm100134" '臺灣地址郵遞區號查詢
         Set GetForm = frm100134
   Case "frm100135" '臺灣地址格式說明畫面
         Set GetForm = frm100135
   'end 2024/09/16
   End Select
End Function

'Added by Morgan 2021/4/22
'複製貼上彈跳視窗
Public Sub PopupMenu2(oTextBox As Control)
   If oTextBox.Enabled = True And oTextBox.Locked = False Then
      mnuPopItem2(0).Enabled = False
      mnuPopItem2(1).Enabled = False
      mnuPopItem2(2).Enabled = False
      mnuPopItem2(3).Enabled = False
      If oTextBox.SelLength > 0 Then
         mnuPopItem2(0).Enabled = True
         mnuPopItem2(1).Enabled = True
         mnuPopItem2(3).Enabled = True
      End If
      If Clipboard.GetText <> "" Then
         mnuPopItem2(2).Enabled = True
      End If
      PopupMenu mnuPop2
   End If
End Sub

'Added by Morgan 2021/4/22
'複製貼上選單
Private Sub mnuPopItem2_Click(Index As Integer)
   Select Case Index
   Case 0 '剪下
      SendKeys "+{DELETE}"
   Case 1 '複製
      SendKeys "^C"
   Case 2 '貼上
      SendKeys "^V"
   Case 3 '刪除
      SendKeys "{DELETE}"
   End Select
End Sub
