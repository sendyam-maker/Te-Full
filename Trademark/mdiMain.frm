VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000018&
   Caption         =   "商標系統"
   ClientHeight    =   4530
   ClientLeft      =   4240
   ClientTop       =   3310
   ClientWidth     =   8880
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  '最大化
   Begin VB.Timer tmrSalary 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1395
      Top             =   3600
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '對齊表單上方
      Height          =   45
      Left            =   0
      ScaleHeight     =   10
      ScaleWidth      =   8840
      TabIndex        =   2
      Top             =   590
      Visible         =   0   'False
      Width           =   8880
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1035
      Top             =   1710
   End
   Begin VB.Timer tmrConnect 
      Left            =   1020
      Top             =   1260
   End
   Begin VB.Timer Timer2 
      Left            =   210
      Top             =   1950
   End
   Begin VB.Timer Timer1 
      Left            =   210
      Top             =   1500
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   276
      Left            =   0
      TabIndex        =   1
      Top             =   4248
      Width           =   8880
      _ExtentX        =   15663
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
      Height          =   590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1041
      ButtonWidth     =   617
      ButtonHeight    =   935
      Appearance      =   1
      _Version        =   393216
      Begin VB.TextBox TextComp 
         Height          =   285
         Left            =   990
         TabIndex        =   4
         Text            =   "暫存文字框"
         Top             =   210
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.ListBox List1 
         Height          =   220
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Visible         =   0   'False
         Width           =   825
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
      Begin VB.Menu mnuPopItem2 
         Caption         =   "輸入(&E)"
         Index           =   4
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
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "內商"
      Index           =   2
      Begin VB.Menu mnu02 
         Caption         =   "資料處理"
         Index           =   1
         Begin VB.Menu mnu0201 
            Caption         =   "分案"
            Index           =   1
         End
         Begin VB.Menu mnu0201 
            Caption         =   "發文"
            Index           =   2
         End
         Begin VB.Menu mnu0201 
            Caption         =   "申請案號輸入"
            Index           =   3
         End
         Begin VB.Menu mnu0201 
            Caption         =   "商申審查機關來函"
            Index           =   4
            Begin VB.Menu mnu020104 
               Caption         =   "核准"
               Index           =   1
            End
            Begin VB.Menu mnu020104 
               Caption         =   "核駁"
               Index           =   2
            End
            Begin VB.Menu mnu020104 
               Caption         =   "審查報告"
               Index           =   3
            End
            Begin VB.Menu mnu020104 
               Caption         =   "註冊證"
               Index           =   4
            End
            Begin VB.Menu mnu020104 
               Caption         =   "取消催審期限"
               Index           =   5
            End
            Begin VB.Menu mnu020104 
               Caption         =   "被禁止處分"
               Index           =   6
            End
            Begin VB.Menu mnu020104 
               Caption         =   "延期受理"
               Index           =   7
            End
            Begin VB.Menu mnu020104 
               Caption         =   "其他來函"
               Index           =   8
            End
            Begin VB.Menu mnu020104 
               Caption         =   "服務業務結果"
               Index           =   9
            End
            Begin VB.Menu mnu020104 
               Caption         =   "廣告刊出來函輸入"
               Index           =   10
            End
            Begin VB.Menu mnu020104 
               Caption         =   "智慧局註冊費通知函"
               Index           =   11
            End
            Begin VB.Menu mnu020104 
               Caption         =   "大陸商標審定公告及通知續展匯入作業"
               Index           =   12
            End
         End
         Begin VB.Menu mnu0201 
            Caption         =   "商爭審查機關來函"
            Index           =   5
            Begin VB.Menu mnu020105 
               Caption         =   "勝訴"
               Index           =   1
            End
            Begin VB.Menu mnu020105 
               Caption         =   "敗訴"
               Index           =   2
            End
            Begin VB.Menu mnu020105 
               Caption         =   "被異議/被評定/被撤銷/對方補充理由/對方延期"
               Index           =   3
            End
            Begin VB.Menu mnu020105 
               Caption         =   "發回補理由/發回補答辯"
               Index           =   4
            End
            Begin VB.Menu mnu020105 
               Caption         =   "撤銷原處分／和解"
               Index           =   5
            End
            Begin VB.Menu mnu020105 
               Caption         =   "受理"
               Index           =   6
            End
            Begin VB.Menu mnu020105 
               Caption         =   "延長審查時間"
               Index           =   7
            End
            Begin VB.Menu mnu020105 
               Caption         =   "對方撤回"
               Index           =   8
            End
            Begin VB.Menu mnu020105 
               Caption         =   "其他來函"
               Index           =   9
            End
            Begin VB.Menu mnu020105 
               Caption         =   "延期受理"
               Index           =   10
            End
            Begin VB.Menu mnu020105 
               Caption         =   "部分勝部分敗"
               Index           =   11
            End
         End
         Begin VB.Menu mnu0201 
            Caption         =   "代理人來函"
            Index           =   6
            Begin VB.Menu mnu020106 
               Caption         =   "已收達/已提申"
               Index           =   1
            End
            Begin VB.Menu mnu020106 
               Caption         =   "通知修正"
               Index           =   2
            End
            Begin VB.Menu mnu020106 
               Caption         =   "其他來函"
               Index           =   3
            End
         End
         Begin VB.Menu mnu0201 
            Caption         =   "電子公文來函"
            Index           =   7
         End
         Begin VB.Menu mnu0201 
            Caption         =   "待送件區"
            Index           =   8
         End
         Begin VB.Menu mnu0201 
            Caption         =   "待處理區"
            Index           =   9
         End
         Begin VB.Menu mnu0201 
            Caption         =   "掃瞄資料匯入(未來會撤掉)"
            Index           =   10
         End
         Begin VB.Menu mnu0201 
            Caption         =   "電子收據匯入"
            Index           =   11
         End
         Begin VB.Menu mnu0201 
            Caption         =   "收據/回執整批匯入"
            Index           =   12
         End
         Begin VB.Menu mnu0201 
            Caption         =   "公文來函文檔整批匯入"
            Index           =   13
         End
         Begin VB.Menu mnu0201 
            Caption         =   "代理人來函匯入"
            Index           =   14
         End
      End
      Begin VB.Menu mnu02 
         Caption         =   "查詢作業"
         Index           =   2
         Begin VB.Menu mnu0202 
            Caption         =   "代理人新案案件統計"
            Index           =   1
         End
         Begin VB.Menu mnu0202 
            Caption         =   "未請款明細查詢"
            Index           =   3
         End
         Begin VB.Menu mnu0202 
            Caption         =   "審查委員准駁統計"
            Index           =   4
         End
         Begin VB.Menu mnu0202 
            Caption         =   "FC收款請款點數查詢"
            Index           =   5
         End
         Begin VB.Menu mnu0202 
            Caption         =   "代理人案件性質統計"
            Index           =   6
         End
         Begin VB.Menu mnu0202 
            Caption         =   "延展前商標無效管制表"
            Index           =   7
         End
         Begin VB.Menu mnu0202 
            Caption         =   "員工查詢印表記錄資料查詢"
            Index           =   8
         End
         Begin VB.Menu mnu0202 
            Caption         =   "商標處期限通知"
            Index           =   9
         End
      End
      Begin VB.Menu mnu02 
         Caption         =   "報表列印"
         Index           =   3
         Begin VB.Menu mnu0203 
            Caption         =   "智權人員期限管制表"
            Index           =   1
         End
         Begin VB.Menu mnu0203 
            Caption         =   "承辦人期限管制表"
            Index           =   2
         End
         Begin VB.Menu mnu0203 
            Caption         =   "代理人案件收達/提申管制表"
            Index           =   3
         End
         Begin VB.Menu mnu0203 
            Caption         =   "對外案件延展未提申明細表"
            Index           =   4
         End
         Begin VB.Menu mnu0203 
            Caption         =   "收文未發文明細表"
            Index           =   5
         End
         Begin VB.Menu mnu0203 
            Caption         =   "催審函/催審表"
            Index           =   6
         End
         Begin VB.Menu mnu0203 
            Caption         =   "智權人員收文明細表"
            Index           =   7
         End
         Begin VB.Menu mnu0203 
            Caption         =   "承辦人案件明細表"
            Index           =   8
         End
         Begin VB.Menu mnu0203 
            Caption         =   "申請意見書案件明細表"
            Index           =   9
         End
         Begin VB.Menu mnu0203 
            Caption         =   "商品類別/群組明細表"
            Index           =   10
         End
         Begin VB.Menu mnu0203 
            Caption         =   "後金案件表"
            Index           =   11
         End
         Begin VB.Menu mnu0203 
            Caption         =   "延期明細表"
            Index           =   12
         End
         Begin VB.Menu mnu0203 
            Caption         =   "不出名案件明細表"
            Index           =   13
         End
         Begin VB.Menu mnu0203 
            Caption         =   "代理人案件總簿"
            Index           =   14
         End
         Begin VB.Menu mnu0203 
            Caption         =   "客戶案件總簿輸出"
            Index           =   15
         End
         Begin VB.Menu mnu0203 
            Caption         =   "代理人/申請人名單"
            Index           =   16
         End
         Begin VB.Menu mnu0203 
            Caption         =   "地址條列印"
            Index           =   18
         End
         Begin VB.Menu mnu0203 
            Caption         =   "國外FC帳款明細表"
            Index           =   19
         End
         Begin VB.Menu mnu0203 
            Caption         =   "智慧局註冊費通知函列印"
            Index           =   20
         End
         Begin VB.Menu mnu0203 
            Caption         =   "下載商標圖參考報表"
            Index           =   21
         End
         Begin VB.Menu mnu0203 
            Caption         =   "台灣商標延展開拓(貝爾)"
            Index           =   22
         End
         Begin VB.Menu mnu0203 
            Caption         =   "台灣商標公告近三年開拓函"
            Index           =   23
         End
         Begin VB.Menu mnu0203 
            Caption         =   "台灣商標延展開拓(智慧局)"
            Index           =   24
         End
         Begin VB.Menu mnu0203 
            Caption         =   "期限通知檢核及報表"
            Index           =   25
         End
      End
      Begin VB.Menu mnu02 
         Caption         =   "統計報表"
         Index           =   4
         Begin VB.Menu mnu0204 
            Caption         =   "商申案智權人員收/發文統計表"
            Index           =   1
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商申案智權人員准駁統計表"
            Index           =   2
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商爭案智權人員收/發文統計表"
            Index           =   3
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商爭案智權人員勝敗統計表"
            Index           =   4
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商申案承辦人收/發文統計表"
            Index           =   5
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商申案承辦人准駁統計表"
            Index           =   6
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商爭案承辦人收/發文統計表"
            Index           =   7
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商爭案承辦人勝敗統計表"
            Index           =   8
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商爭案承辦人預估勝敗統計表"
            Index           =   9
         End
         Begin VB.Menu mnu0204 
            Caption         =   "申請案收/發件數月統計表"
            Index           =   10
         End
         Begin VB.Menu mnu0204 
            Caption         =   "各區收/發文達成比較表"
            Index           =   11
         End
         Begin VB.Menu mnu0204 
            Caption         =   "各區收/發文件數明細表"
            Index           =   12
         End
         Begin VB.Menu mnu0204 
            Caption         =   "代理人新案案件統計表"
            Index           =   13
         End
         Begin VB.Menu mnu0204 
            Caption         =   "代理人新案案件年度統計表"
            Index           =   14
         End
         Begin VB.Menu mnu0204 
            Caption         =   "代理人案件准駁統計表"
            Index           =   15
         End
         Begin VB.Menu mnu0204 
            Caption         =   "代理人/申請人新申請案排行榜"
            Index           =   16
         End
         Begin VB.Menu mnu0204 
            Caption         =   "逾期未結案統計表"
            Index           =   17
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商標承辦人績效表"
            Index           =   18
         End
         Begin VB.Menu mnu0204 
            Caption         =   "商標處國內客戶收/發文件數月報"
            Index           =   19
         End
         Begin VB.Menu mnu0204 
            Caption         =   "台灣商標爭議案件補充資料次數明細及統計"
            Index           =   20
         End
         Begin VB.Menu mnu0204 
            Caption         =   "MCT收發文件數及點數統計"
            Index           =   21
         End
         Begin VB.Menu mnu0204 
            Caption         =   "大陸商申查名統計表"
            Index           =   22
         End
      End
      Begin VB.Menu mnu02 
         Caption         =   "檔案維護"
         Index           =   5
         Begin VB.Menu mnu0205 
            Caption         =   "商標案件基本資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu0205 
            Caption         =   "服務業務基本資料維護"
            Index           =   2
            Begin VB.Menu mnu020502 
               Caption         =   "條碼"
               Index           =   1
            End
            Begin VB.Menu mnu020502 
               Caption         =   "監視系統"
               Index           =   2
            End
            Begin VB.Menu mnu020502 
               Caption         =   "網域"
               Index           =   3
            End
            Begin VB.Menu mnu020502 
               Caption         =   "著作權"
               Index           =   4
            End
            Begin VB.Menu mnu020502 
               Caption         =   "其它業務"
               Index           =   5
            End
         End
         Begin VB.Menu mnu0205 
            Caption         =   "案件進度資料維護"
            Index           =   3
         End
         Begin VB.Menu mnu0205 
            Caption         =   "下一程序資料維護"
            Index           =   4
         End
         Begin VB.Menu mnu0205 
            Caption         =   "國外代理人資料維護"
            Index           =   5
         End
         Begin VB.Menu mnu0205 
            Caption         =   "變更事項資料維護"
            Index           =   6
         End
         Begin VB.Menu mnu0205 
            Caption         =   "延期記錄資料維護"
            Index           =   7
         End
         Begin VB.Menu mnu0205 
            Caption         =   "案件國家收費表維護"
            Index           =   8
         End
         Begin VB.Menu mnu0205 
            Caption         =   "商標條款資料維護"
            Index           =   9
         End
         Begin VB.Menu mnu0205 
            Caption         =   "主張內容分類資料維護"
            Index           =   10
         End
         Begin VB.Menu mnu0205 
            Caption         =   "代理人變更名稱作業"
            Index           =   12
         End
         Begin VB.Menu mnu0205 
            Caption         =   "承辦人目標點數資料維護"
            Index           =   13
         End
         Begin VB.Menu mnu0205 
            Caption         =   "客戶資料維護"
            Index           =   14
         End
         Begin VB.Menu mnu0205 
            Caption         =   "著作權案件登記項目資料維護"
            Index           =   15
         End
         Begin VB.Menu mnu0205 
            Caption         =   "非本所實質客戶資料維護"
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu mnu0205 
            Caption         =   "更換FC代理人作業"
            Index           =   17
         End
         Begin VB.Menu mnu0205 
            Caption         =   "系統特殊設定"
            Index           =   18
         End
         Begin VB.Menu mnu0205 
            Caption         =   "註冊費報價資料維護"
            Index           =   19
         End
      End
      Begin VB.Menu mnu02 
         Caption         =   "商標公報"
         Index           =   6
         Begin VB.Menu mnu0206 
            Caption         =   "商標公報資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu0206 
            Caption         =   "更新審定號作業"
            Index           =   2
         End
         Begin VB.Menu mnu0206 
            Caption         =   "國內商標公報查詢"
            Index           =   3
         End
         Begin VB.Menu mnu0206 
            Caption         =   "國內公報資料檢核表"
            Index           =   4
         End
         Begin VB.Menu mnu0206 
            Caption         =   "國內公報代理人資料"
            Index           =   5
         End
         Begin VB.Menu mnu0206 
            Caption         =   "國內公報代理人合併作業"
            Index           =   6
         End
         Begin VB.Menu mnu0206 
            Caption         =   "國內公報代理人/事務所名稱查詢"
            Index           =   7
         End
         Begin VB.Menu mnu0206 
            Caption         =   "表一＆表二、商標全國市場統計"
            Index           =   8
         End
         Begin VB.Menu mnu0206 
            Caption         =   "表三、各區市場佔有統計"
            Index           =   9
         End
         Begin VB.Menu mnu0206 
            Caption         =   "表四、各類別市場佔有統計"
            Index           =   10
         End
         Begin VB.Menu mnu0206 
            Caption         =   "表五、國外市場排名"
            Index           =   11
         End
         Begin VB.Menu mnu0206 
            Caption         =   "表一～表五統計表列印"
            Index           =   12
         End
         Begin VB.Menu mnu0206 
            Caption         =   "國外前十大申請國及其商品類別排名"
            Index           =   13
         End
         Begin VB.Menu mnu0206 
            Caption         =   "代理人國外案件排名分析表"
            Index           =   14
         End
         Begin VB.Menu mnu0206 
            Caption         =   "國內公報代理人資料列印"
            Index           =   15
         End
         Begin VB.Menu mnu0206 
            Caption         =   "商標公報轉檔作業"
            Index           =   16
         End
         Begin VB.Menu mnu0206 
            Caption         =   "公報特定公司不列印者"
            Index           =   17
         End
         Begin VB.Menu mnu0206 
            Caption         =   "公報開拓資料維護"
            Index           =   18
         End
         Begin VB.Menu mnu0206 
            Caption         =   "商標公報開拓函列印"
            Index           =   19
         End
         Begin VB.Menu mnu0206 
            Caption         =   "商標公報大陸清單列印"
            Index           =   20
         End
         Begin VB.Menu mnu0206 
            Caption         =   "本所案件申請人國籍統計表"
            Index           =   21
         End
         Begin VB.Menu mnu0206 
            Caption         =   "商標公報資料統計-Excel"
            Index           =   22
            Begin VB.Menu mnu020622 
               Caption         =   "申請人國籍及洲別統計(含同業)"
               Index           =   1
            End
            Begin VB.Menu mnu020622 
               Caption         =   "各單位公報類別數統計"
               Index           =   2
            End
            Begin VB.Menu mnu020622 
               Caption         =   "三部門案件來源比較"
               Index           =   3
            End
            Begin VB.Menu mnu020622 
               Caption         =   "同業案件來源比較"
               Index           =   4
            End
            Begin VB.Menu mnu020622 
               Caption         =   "同業台灣各區類別數比較"
               Index           =   5
            End
         End
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
            Caption         =   "介紹法律所案源查詢"
            Index           =   18
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
         Caption         =   "來文資料稽核表"
         Index           =   3
      End
      Begin VB.Menu mnu11 
         Caption         =   "聯絡單列印及E-Mail"
         Index           =   4
      End
      Begin VB.Menu mnu11 
         Caption         =   "相關卷號資料維護"
         Index           =   5
      End
      Begin VB.Menu mnu11 
         Caption         =   "多案相關卷號關係建立"
         Index           =   6
      End
      Begin VB.Menu mnu11 
         Caption         =   "分割案件關係維護"
         Index           =   7
      End
      Begin VB.Menu mnu11 
         Caption         =   "撰寫信函作業"
         Index           =   8
      End
      Begin VB.Menu mnu11 
         Caption         =   "作業失誤"
         Index           =   9
         Begin VB.Menu mnu1108 
            Caption         =   "作業失誤資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu1108 
            Caption         =   "作業失誤明細表"
            Index           =   2
         End
      End
      Begin VB.Menu mnu11 
         Caption         =   "代理人帳目"
         Index           =   10
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
               Visible         =   0   'False
            End
            Begin VB.Menu mnu110401 
               Caption         =   "帳單輸入-整批"
               Index           =   14
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
               Caption         =   "代理人逾期帳款分析"
               Index           =   7
            End
            Begin VB.Menu mnu110403 
               Caption         =   "折讓單列印"
               Index           =   8
            End
            Begin VB.Menu mnu110403 
               Caption         =   "國外請款點數分析表"
               Index           =   9
            End
            Begin VB.Menu mnu110403 
               Caption         =   "請款單月報表"
               Index           =   10
            End
            Begin VB.Menu mnu110403 
               Caption         =   "請款單折扣案件明細"
               Index           =   11
            End
         End
      End
      Begin VB.Menu mnu11 
         Caption         =   "部門別送件清單列印"
         Index           =   14
      End
      Begin VB.Menu mnu11 
         Caption         =   "部門別電子送件清單列印"
         Index           =   15
      End
      Begin VB.Menu mnu11 
         Caption         =   "銷案延遲日期輸入作業"
         Index           =   16
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘單查詢"
         Index           =   17
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘資料維護"
         Index           =   18
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘單案件明細查詢"
         Index           =   19
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "說明"
      Enabled         =   0   'False
      Index           =   15
      Visible         =   0   'False
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
      Caption         =   "定稿"
      Index           =   16
      Begin VB.Menu mnu16 
         Caption         =   "整批列印定稿"
         Index           =   1
         Begin VB.Menu mnu1601 
            Caption         =   "橫式定稿"
            Index           =   1
         End
         Begin VB.Menu mnu1601 
            Caption         =   "英文定稿"
            Index           =   2
         End
         Begin VB.Menu mnu1601 
            Caption         =   "直式定稿"
            Index           =   3
         End
         Begin VB.Menu mnu1601 
            Caption         =   "日文定稿"
            Index           =   4
         End
         Begin VB.Menu mnu1601 
            Caption         =   "申請書"
            Index           =   5
         End
         Begin VB.Menu mnu1601 
            Caption         =   "報價定稿"
            Index           =   6
         End
         Begin VB.Menu mnu1601 
            Caption         =   "橫式定稿(不印信頭)"
            Index           =   7
         End
         Begin VB.Menu mnu1601 
            Caption         =   "橫式雙面列印定稿"
            Index           =   8
         End
      End
      Begin VB.Menu mnu16 
         Caption         =   "定稿資料維護"
         Index           =   2
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "品名查詢"
      Index           =   20
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
         Caption         =   "案件表單查詢及簽核"
         Index           =   3
         Begin VB.Menu mnu2303 
            Caption         =   "目前表單"
            Index           =   1
         End
         Begin VB.Menu mnu2303 
            Caption         =   "簽核作業"
            Index           =   2
         End
         Begin VB.Menu mnu2303 
            Caption         =   "專業部 審核/補看"
            Index           =   3
         End
      End
      Begin VB.Menu mnu23 
         Caption         =   "教育訓練登錄作業"
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
         Caption         =   "系統收件區"
         Index           =   7
      End
      Begin VB.Menu mnu23 
         Caption         =   "信箱分信紀錄查詢"
         Index           =   8
      End
      Begin VB.Menu mnu23 
         Caption         =   "圖書借閱資料查詢"
         Index           =   9
      End
      Begin VB.Menu mnu23 
         Caption         =   "行事曆提醒通知"
         Index           =   10
      End
      Begin VB.Menu mnu23 
         Caption         =   "風險檢查對象資料維護"
         Index           =   11
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "設定"
      Index           =   98
      Begin VB.Menu mnu98 
         Caption         =   "系統印表機設定"
         Index           =   0
      End
      Begin VB.Menu mnu98 
         Caption         =   "報表紙張格式設定"
         Index           =   1
      End
      Begin VB.Menu mnu98 
         Caption         =   "解除畫面擷取限制"
         Index           =   2
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
   Begin VB.Menu mnuMouseR 
      Caption         =   "按右鍵用"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "複製"
      End
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
   Begin VB.Menu mnuPopEDoc 
      Caption         =   "電子機關來函彈跳選單"
      Visible         =   0   'False
      Begin VB.Menu mnuPopEDoc01 
         Caption         =   "商申審查機關來函"
         Index           =   1
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "核准"
            Index           =   1
         End
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "核駁"
            Index           =   2
         End
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "審查報告"
            Index           =   3
         End
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "註冊證"
            Index           =   4
         End
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "取消催審期限"
            Index           =   5
         End
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "被禁止處分"
            Index           =   6
         End
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "延期受理"
            Index           =   7
         End
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "其他來函"
            Index           =   8
         End
         Begin VB.Menu mnuPopEDoc0101 
            Caption         =   "智慧局註冊費通知函"
            Index           =   11
         End
      End
      Begin VB.Menu mnuPopEDoc02 
         Caption         =   "商爭審查機關來函"
         Index           =   2
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "勝訴"
            Index           =   1
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "敗訴"
            Index           =   2
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "被異議/被評定/被撤銷/對方補充理由/對方延期"
            Index           =   3
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "發回補理由/發回補答辯"
            Index           =   4
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "撤銷原處分／和解"
            Index           =   5
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "受理"
            Index           =   6
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "延長審查時間"
            Index           =   7
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "對方撤回"
            Index           =   8
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "其他來函"
            Index           =   9
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "延期受理"
            Index           =   10
         End
         Begin VB.Menu mnuPopEDoc0201 
            Caption         =   "部分勝部分敗"
            Index           =   11
         End
      End
      Begin VB.Menu mnuPopEDoc03 
         Caption         =   "FC商申審查機關來函"
         Visible         =   0   'False
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "核准"
            Index           =   1
         End
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "核駁"
            Index           =   2
         End
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "審查報告"
            Index           =   3
         End
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "註冊證"
            Index           =   4
         End
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "取消催審期限"
            Index           =   5
         End
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "被禁止處分"
            Index           =   6
         End
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "延期受理"
            Index           =   7
         End
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "其他來函"
            Index           =   8
         End
         Begin VB.Menu mnuPopEDoc0301 
            Caption         =   "通知已轉他所"
            Index           =   12
         End
      End
   End
   Begin VB.Menu mnuPopEMail1 
      Caption         =   "T案EMail來函彈跳選單"
      Visible         =   0   'False
      Begin VB.Menu mnuPopTEMailItem 
         Caption         =   "申請案號輸入"
         Index           =   1
      End
      Begin VB.Menu mnuPopTEMailItem 
         Caption         =   "商申審查機關來函"
         Index           =   2
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "核准"
            Index           =   1
         End
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "核駁"
            Index           =   2
         End
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "審查報告"
            Index           =   3
         End
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "註冊證"
            Index           =   4
         End
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "取消催審期限"
            Index           =   5
         End
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "被禁止處分"
            Index           =   6
         End
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "延期受理"
            Index           =   7
         End
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "其他來函"
            Index           =   8
         End
         Begin VB.Menu mnuPopTEMailItem2 
            Caption         =   "服務業務結果"
            Index           =   9
         End
      End
      Begin VB.Menu mnuPopTEMailItem 
         Caption         =   "商爭審查機關來函"
         Index           =   3
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "勝訴"
            Index           =   1
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "敗訴"
            Index           =   2
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "被異議/被評定/被撤銷/對方補充理由/對方延期"
            Index           =   3
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "發回補理由/發回補答辯"
            Index           =   4
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "撤銷原處分／和解"
            Index           =   5
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "受理"
            Index           =   6
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "延長審查時間"
            Index           =   7
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "對方撤回"
            Index           =   8
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "其他來函"
            Index           =   9
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "延期受理"
            Index           =   10
         End
         Begin VB.Menu mnuPopTEMailItem3 
            Caption         =   "部分勝部分敗"
            Index           =   11
         End
      End
      Begin VB.Menu mnuPopTEMailItem 
         Caption         =   "代理人來函"
         Index           =   4
         Begin VB.Menu mnuPopTEMailItem4 
            Caption         =   "已收達/已提申"
            Index           =   1
         End
         Begin VB.Menu mnuPopTEMailItem4 
            Caption         =   "通知修正"
            Index           =   2
         End
         Begin VB.Menu mnuPopTEMailItem4 
            Caption         =   "其他來函"
            Index           =   3
         End
      End
      Begin VB.Menu mnuPopTEMailItem 
         Caption         =   "其他"
         Index           =   5
         Begin VB.Menu mnuPopTEMailItem5 
            Caption         =   "內部收文"
            Index           =   1
         End
         Begin VB.Menu mnuPopTEMailItem5 
            Caption         =   "帳單輸入"
            Index           =   2
         End
         Begin VB.Menu mnuPopTEMailItem5 
            Caption         =   "抵帳單輸入"
            Index           =   3
         End
         Begin VB.Menu mnuPopTEMailItem5 
            Caption         =   "帳單作廢輸入"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

'Add by Morgan 2003/12/23
Dim WithEvents eventConn As ADODB.Connection
Attribute eventConn.VB_VarHelpID = -1
Public bolReOpen As Boolean

'move to basquery 2007/02/07
'intPCaseKind分案之系統分類，intPWhere 0國內  1國外CF  2國外FC
'Public intPCaseKind As Integer, intPWhere As Integer
'add by nick 2004/09/27 品名查詢用
Public CopyWord As String
'Add by Morgan 2008/12/2 是否已經做過
Dim m_blnActivated As Boolean
Public m_wasMaximized As Boolean 'Added by Morgan 畫面最小化後判斷原來是否為最大化用
Public m_ChkIsOpenFrm180203 As Boolean 'Add By Sindy 2013/7/8
Public FCT修改承辦人 As String 'Added by Lydia 2018/01/12 有權限執行FCT修改承辦人的人員
Dim m_UserNo As String 'Added by Morgan 2019/5/13 報價定稿人員
Dim oControl As Control  'Added by Morgan 2022/1/22


'Add by Morgan 2003/12/23
Private Sub eventConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   tmrConnect.Tag = 0
End Sub

Private Sub eventConn_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
   Debug.Print "Err=" & pError
   Debug.Print "state=" & adStatus
End Sub

'Add by Morgan 2003/12/29
Private Sub SwitchMenu(Optional bolEnable As Boolean = True)
   Dim mnuTmp As Menu
   For Each mnuTmp In mnuTitle
      If mnuTmp.Index <> 0 Then mnuTmp.Enabled = bolEnable
   Next
   If bolEnable = False Then Toolbar1.Visible = False
End Sub

'Add by Morgan 2003/12/29
Private Sub CloseAllChild()
   Dim frmTemp As Form
   For Each frmTemp In Forms
      If frmTemp.Name <> "mdiMain" Then Unload frmTemp
   Next
End Sub

'Add by Morgan 2003/12/29
Private Sub ReConnect()
'Modify by Morgan 2005/1/10 不需再控制
'      Call SwitchMenu(True)
'2005/1/10 end
      Timer1.Enabled = True
      Timer1.Interval = 100
      tmrConnect.Tag = 0
End Sub

'Add By Sindy 2009/08/24
Private Sub FcpDueCaseQurey()
Dim tmpST15 As String
   
   'Modify By Sindy 2020/8/25 Trademark:限P2X部門人員
   tmpST15 = PUB_GetStaffST15(strUserNum, 1)
   If UCase(Mid(tmpST15, 1, 2)) <> "P2" Then Exit Sub
   '2020/8/25 END
   
   '電腦中心除外
   If Pub_StrUserSt03 <> "M51" Then
      If CheckUse("frm020201", strExec, False) = True And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         '一天僅自動彈跳通知一次
         'strSQL = "select * from executelog where el01='frm020201' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
         strSql = "select * from executelog where el01='frm020201' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI <> 1 Then
            Load frm020201
            frm020201.cmdQuery(0).Value = True
         End If
      End If
   End If
End Sub

'Add By Sindy 2025/11/3
Public Sub SetTmpForm()
   Set Tmpfrm210147 = frm210147
   Set Tmpfrm210148 = frm210148
   Set Tmpfrm180201 = frm180201
   Set Tmpfrm180101 = frm180101
   Set Tmpfrm180203_1 = frm180203_1
   Set Tmpfrm160102 = frm160102
   Set Tmpfrm160018 = frm160018
   Set Tmpfrm010035_2 = frm010035_2
End Sub

Private Sub MDIForm_Activate()
   'Modify By Sindy 2025/11/3 改為共用函數
   Call MDIFormStarProc
End Sub

'Add By Sindy 2011/10/7
Public Sub SysStartCallForm()
   'Add by Morgan 2008/12/2
   If m_blnActivated = False Then
      m_blnActivated = True
      '智權人員期限資料查詢
      SalesDueCaseQuery
      '報價定稿
      PrintLetter
   End If
   
   'Add By Sindy 2009/08/24
   '本查詢需考慮當閒置太久重新登入且已經是下午時須再次執行故與單獨控制
   If pub_bolInformCheck = True Then
      '商標處期限通知
      FcpDueCaseQurey
      
      'Added by Lydia 2020/01/15 非外專人員行事曆提醒通知
      If PUB_CheckStaffCalendarDue = True Then
           frm060209.m_Role = "F41"
           frm060209.Show
      End If
      'end 2020/01/15
      
      pub_bolInformCheck = False
   End If
End Sub

'Add by Morgan 2008/12/2
Private Sub PrintLetter()
   Dim stTemp As String, arrNum() As String, ii As Integer 'Added by Morgan 2019/5/13
   
   'Added by Lydia 2016/08/31 輸入提高點數簽核主管及簽核點數
   Set Tmpfrm880004_4 = frm880004
   
   'Added by Lydia 2016/12/22 設定呼叫FC催款單
   Set TmpFrmAcc2470 = Frmacc2470
   
   If PUB_Cache2Letter(, , False, , True) = True Then
       If MsgBox("有報價定稿待列印，是否現在執行！", vbYesNo + vbDefaultButton2, "報價定稿提醒") = vbYes Then
          mnu1601_Click 6
       End If
   End If
   
   'Added by Morgan 2019/5/13 --桂英
   Pub_SetForOthersEmpCombo strUserNum, , , stTemp
   If stTemp <> "" Then
     arrNum = Split(stTemp, ";")
     For ii = LBound(arrNum) To UBound(arrNum)
        If arrNum(ii) <> "" Then
           If PUB_Cache2Letter(, , False, , True, , arrNum(ii)) = True Then
              stTemp = GetPrjSales(arrNum(ii))
              If MsgBox(stTemp & " 有報價定稿待列印，是否現在執行！", vbQuestion + vbYesNo + vbDefaultButton2, "報價定稿提醒") = vbYes Then
                 m_UserNo = arrNum(ii)
                 mnu1601_Click 6
                 m_UserNo = ""
              End If
           End If
        End If
     Next
   End If
   'end 2019/5/13
   
End Sub

'Modify by Morgan 2008/12/2 原來放在Timer內執行,現改以呼叫方式執行
Private Sub SalesDueCaseQuery()
   Dim tmpST15 As String
   Dim tmpST03 As String
   ''add by nickc 2005/09/02 非員工不跑
   '2011/3/30 MODIFY BY SONIA
   'If strUserNum >= "63001" And strUserNum < "A" Then
   If strUserNum >= "63001" And strUserNum < "F" Then
      tmpST15 = PUB_GetStaffST15(strUserNum, 1)
      'add by nickc 2005/09/20
      tmpST03 = PUB_GetST03(strUserNum)
      'edit by nickc 2005/09/20
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Then
      'edit by nickc 2006/04/04 加入非智權人員但有收文的人
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Then
      '2006/6/1 MODIFY BY SONIA 取消P2,P3中所之控制,改由PUB_ChkNotSalesButHaveCase控制
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
      If UCase(Mid(tmpST15, 1, 1)) = "S" Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
'          If ServerTime <= 100000 Then
'               pub_CallNextForm = True
'               frm100123.Show
'               frm100123.cmdSearch_Click
'         Else
'            If MsgBox("是否執行  智權人員期限資料查詢  功能", vbYesNo, "功能！") = vbYes Then
'               pub_CallNextForm = True
'               frm100123.Show
'               frm100123.cmdSearch_Click
'            End If
'         End If
         'Modify By Sindy 2015/2/10 Mark,因Menu已無掛智權部
'         '電腦中心除外
'         If Pub_StrUserSt03 <> "M51" Then
'            strSql = "select * from executelog where el01='frm100123' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI <> 1 Then
'               pub_CallNextForm = True
'               frm100123.Show
'               frm100123.cmdSearch_Click
'            Else
'               If MsgBox("是否執行  智權人員期限資料查詢  功能", vbYesNo, "功能！") = vbYes Then
'                  pub_CallNextForm = True
'                  frm100123.Show
'                  frm100123.cmdSearch_Click
'               End If
'            End If
'         End If
      End If
   End If
End Sub

Private Sub MDIForm_Resize()
   'Added by Morgan 2011/12/14 紀錄是否為最大化狀態
   If Me.WindowState = 2 Then
      m_wasMaximized = True
   ElseIf Me.WindowState = 0 Then
      m_wasMaximized = False
   End If
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
      'Add By Sindy 2020/5/5
      Case 18 '介紹法律所案源查詢
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
   End Select
End Sub

Private Sub mnu23_Click(Index As Integer)
Dim nFrm As Form
   
   Select Case Index
      Case 1 '預約作業
         frm140112.Show
      Case 4 'Add by Amy 2020/11/02 教育訓練登入作業
        frm140113.Show
      'Add By Sindy 2020/5/25
      Case 7 '系統收件區
         frm090225.Show
      '2020/5/25 END
      'Add By Sindy 2017/12/13 + 信箱分信紀錄查詢
      Case 8
         '檢查表單是否已開啟，若是，則關閉
         For Each nFrm In Forms
            If StrComp(nFrm.Name, "frm06010613", vbTextCompare) = 0 Then
               Unload frm06010613
            End If
         Next
         frm06010613.m_WorkType = "0" '信箱主檔
         frm06010613.Show
      Case 9 '圖書借閱資料查詢 Add by Amy 2017/01/25
         frm010035.Show
         'Add by Amy 2017/02/03 判斷是否有圖書借閱記錄需簽核
         If GetLoanRecordApply = True Then
            frm010035.bolLoanRecordApply = True
            Call frm010035.cmdLoanRecord_Click
         End If
      'Added by Lydia 2020/01/15
      Case 10 '行事曆提醒通知
         frm060209.m_Role = "F41"
         frm060209.Show
      'Add by Amy 2024/01/22
      Case 11 '風險檢查對象資料維護
         frm12040163.Show
   End Select
End Sub

'Add by Amy 2018/08/17 案件表單查詢及簽核
Private Sub mnu2303_Click(Index As Integer)
    Select Case Index
      Case 1 '目前表單
         frm210147.Show
      Case 2 '簽核作業
         frm210148.Show
      Case 3 '審核/補看作業
         frm040118.Show
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

'設定
Private Sub mnu98_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 0 '系統印表機設定
         frm880011.bolAppOnly = True
         frm880011.Show 1
         
      'Add by Morgan 2008/3/27
      Case 1 '報表紙張格式設定
         frm880013.Show vbModal
         
      'Added by Morgan 2015/3/20
      Case 2 '解除畫面擷取限制
         frmChgUser.Caption = "解除畫面擷取限制"
         frmChgUser.SSTab1.TabVisible(1) = True
         frmChgUser.SSTab1.TabVisible(0) = False
         frmChgUser.Show
   End Select
End Sub

Private Sub mnuChUser_Click()
   frmChgUser.Show
End Sub

'add by nick 2004/09/27
Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText CopyWord
End Sub

Private Sub mnuDML_Click(Index As Integer)
    frmDML.Show
End Sub

'Add By Sindy 2019/5/10
Private Sub mnuPopTEMailItem_Click(Index As Integer)
   OpenForm_P1 Index
End Sub
Private Sub mnuPopTEMailItem2_Click(Index As Integer)
   OpenForm_P2 Index
End Sub
Private Sub mnuPopTEMailItem3_Click(Index As Integer)
   OpenForm_P3 Index
End Sub
Private Sub mnuPopTEMailItem4_Click(Index As Integer)
   OpenForm_P4 Index
End Sub
Private Sub mnuPopTEMailItem5_Click(Index As Integer)
   OpenForm_P5 Index
End Sub
Private Sub OpenForm_P1(Index As Integer)
   Select Case Index
      Case 1 '申請案號輸入
         If CheckUse("frm02010301_1", strExec) = True Then
            Call frm02010301_1.SetParent(frm090225)
            frm02010301_1.m_strIR01 = frm090225.m_strIR01
            frm02010301_1.m_strIR02 = frm090225.m_strIR02
            frm02010301_1.m_strIR03 = frm090225.m_strIR03
            frm02010301_1.m_strIR04 = frm090225.m_strIR04
            frm02010301_1.strTM01 = frm090225.txtTi18
            frm02010301_1.strTM02 = frm090225.txtTi19
            frm02010301_1.strTM03 = frm090225.txtTi20
            frm02010301_1.strTM04 = frm090225.txtTi21
            frm02010301_1.m_RDate = frm090225.m_strTi12
            frm02010301_1.Show
         End If
   End Select
End Sub
Private Sub OpenForm_P2(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1:   '非爭議案核准輸入
         If CheckUse("frm02010401_1", strExec) = True Then
            Call frm02010401_1.SetParent(frm090225)
            frm02010401_1.m_strIR01 = frm090225.m_strIR01
            frm02010401_1.m_strIR02 = frm090225.m_strIR02
            frm02010401_1.m_strIR03 = frm090225.m_strIR03
            frm02010401_1.m_strIR04 = frm090225.m_strIR04
            frm02010401_1.strTM01 = frm090225.txtTi18
            frm02010401_1.strTM02 = frm090225.txtTi19
            frm02010401_1.strTM03 = frm090225.txtTi20
            frm02010401_1.strTM04 = frm090225.txtTi21
            frm02010401_1.m_AppNo = frm090225.m_AppNo
            frm02010401_1.m_RegNo = frm090225.m_RegNo
            frm02010401_1.m_RDate = frm090225.m_strTi12
            frm02010401_1.Show
         End If
      Case 2:   '非爭議案核駁輸入
         If CheckUse("frm02010402_1", strExec) = True Then
            Call frm02010402_1.SetParent(frm090225)
            frm02010402_1.m_strIR01 = frm090225.m_strIR01
            frm02010402_1.m_strIR02 = frm090225.m_strIR02
            frm02010402_1.m_strIR03 = frm090225.m_strIR03
            frm02010402_1.m_strIR04 = frm090225.m_strIR04
            frm02010402_1.strTM01 = frm090225.txtTi18
            frm02010402_1.strTM02 = frm090225.txtTi19
            frm02010402_1.strTM03 = frm090225.txtTi20
            frm02010402_1.strTM04 = frm090225.txtTi21
            frm02010402_1.m_AppNo = frm090225.m_AppNo
            frm02010402_1.m_RegNo = frm090225.m_RegNo
            frm02010402_1.m_RDate = frm090225.m_strTi12
            frm02010402_1.Show
         End If
      Case 3:   '審查報告輸入
         If CheckUse("frm02010403_1", strExec) = True Then
            Call frm02010403_1.SetParent(frm090225)
            frm02010403_1.m_strIR01 = frm090225.m_strIR01
            frm02010403_1.m_strIR02 = frm090225.m_strIR02
            frm02010403_1.m_strIR03 = frm090225.m_strIR03
            frm02010403_1.m_strIR04 = frm090225.m_strIR04
            frm02010403_1.strTM01 = frm090225.txtTi18
            frm02010403_1.strTM02 = frm090225.txtTi19
            frm02010403_1.strTM03 = frm090225.txtTi20
            frm02010403_1.strTM04 = frm090225.txtTi21
            frm02010403_1.m_AppNo = frm090225.m_AppNo
            frm02010403_1.m_RegNo = frm090225.m_RegNo
            frm02010403_1.m_RDate = frm090225.m_strTi12
            frm02010403_1.Show
         End If
      Case 4:   '註冊證輸入
         If CheckUse("frm02010404_1", strExec) = True Then
            Call frm02010404_1.SetParent(frm090225)
            frm02010404_1.m_strIR01 = frm090225.m_strIR01
            frm02010404_1.m_strIR02 = frm090225.m_strIR02
            frm02010404_1.m_strIR03 = frm090225.m_strIR03
            frm02010404_1.m_strIR04 = frm090225.m_strIR04
            frm02010404_1.strTM01 = frm090225.txtTi18
            frm02010404_1.strTM02 = frm090225.txtTi19
            frm02010404_1.strTM03 = frm090225.txtTi20
            frm02010404_1.strTM04 = frm090225.txtTi21
            frm02010404_1.m_AppNo = frm090225.m_AppNo
            frm02010404_1.m_RegNo = frm090225.m_RegNo
            frm02010404_1.m_RDate = frm090225.m_strTi12
            frm02010404_1.Show
         End If
      Case 5:   '非爭議案取消催審期限
         If CheckUse("frm02010405_1", strExec) = True Then
            Call frm02010405_1.SetParent(frm090225)
            frm02010405_1.m_strIR01 = frm090225.m_strIR01
            frm02010405_1.m_strIR02 = frm090225.m_strIR02
            frm02010405_1.m_strIR03 = frm090225.m_strIR03
            frm02010405_1.m_strIR04 = frm090225.m_strIR04
            frm02010405_1.strTM01 = frm090225.txtTi18
            frm02010405_1.strTM02 = frm090225.txtTi19
            frm02010405_1.strTM03 = frm090225.txtTi20
            frm02010405_1.strTM04 = frm090225.txtTi21
            frm02010405_1.m_AppNo = frm090225.m_AppNo
            frm02010405_1.m_RegNo = frm090225.m_RegNo
            frm02010405_1.m_RDate = frm090225.m_strTi12
            frm02010405_1.Show
         End If
      Case 6:   '商標案被禁止處分
         If CheckUse("frm02010406_1", strExec) = True Then
            Call frm02010406_1.SetParent(frm090225)
            frm02010406_1.m_strIR01 = frm090225.m_strIR01
            frm02010406_1.m_strIR02 = frm090225.m_strIR02
            frm02010406_1.m_strIR03 = frm090225.m_strIR03
            frm02010406_1.m_strIR04 = frm090225.m_strIR04
            frm02010406_1.strTM01 = frm090225.txtTi18
            frm02010406_1.strTM02 = frm090225.txtTi19
            frm02010406_1.strTM03 = frm090225.txtTi20
            frm02010406_1.strTM04 = frm090225.txtTi21
            frm02010406_1.m_AppNo = frm090225.m_AppNo
            frm02010406_1.m_RegNo = frm090225.m_RegNo
            frm02010406_1.m_RDate = frm090225.m_strTi12
            frm02010406_1.Show
         End If
      Case 7:   '延期受理
         If CheckUse("frm02010407_1", strExec) = True Then
            Call frm02010407_1.SetParent(frm090225)
            frm02010407_1.m_strIR01 = frm090225.m_strIR01
            frm02010407_1.m_strIR02 = frm090225.m_strIR02
            frm02010407_1.m_strIR03 = frm090225.m_strIR03
            frm02010407_1.m_strIR04 = frm090225.m_strIR04
            frm02010407_1.strTM01 = frm090225.txtTi18
            frm02010407_1.strTM02 = frm090225.txtTi19
            frm02010407_1.strTM03 = frm090225.txtTi20
            frm02010407_1.strTM04 = frm090225.txtTi21
            frm02010407_1.m_AppNo = frm090225.m_AppNo
            frm02010407_1.m_RegNo = frm090225.m_RegNo
            frm02010407_1.m_RDate = frm090225.m_strTi12
            frm02010407_1.Show
         End If
      Case 8:   '其他來函輸入
         If CheckUse("frm02010408_1", strExec) = True Then
            Call frm02010408_1.SetParent(frm090225)
            frm02010408_1.m_strIR01 = frm090225.m_strIR01
            frm02010408_1.m_strIR02 = frm090225.m_strIR02
            frm02010408_1.m_strIR03 = frm090225.m_strIR03
            frm02010408_1.m_strIR04 = frm090225.m_strIR04
            frm02010408_1.strTM01 = frm090225.txtTi18
            frm02010408_1.strTM02 = frm090225.txtTi19
            frm02010408_1.strTM03 = frm090225.txtTi20
            frm02010408_1.strTM04 = frm090225.txtTi21
            frm02010408_1.m_AppNo = frm090225.m_AppNo
            frm02010408_1.m_RegNo = frm090225.m_RegNo
            frm02010408_1.m_RDate = frm090225.m_strTi12
            frm02010408_1.Show
         End If
      Case 9:   '服務業務結果輸入
         If CheckUse("frm02010409_1", strExec) = True Then
            Call frm02010409_1.SetParent(frm090225)
            frm02010409_1.m_strIR01 = frm090225.m_strIR01
            frm02010409_1.m_strIR02 = frm090225.m_strIR02
            frm02010409_1.m_strIR03 = frm090225.m_strIR03
            frm02010409_1.m_strIR04 = frm090225.m_strIR04
            frm02010409_1.m_SP01 = frm090225.txtTi18
            frm02010409_1.m_SP02 = frm090225.txtTi19
            frm02010409_1.m_SP03 = frm090225.txtTi20
            frm02010409_1.m_SP04 = frm090225.txtTi21
            frm02010409_1.m_RDate = frm090225.m_strTi12
            frm02010409_1.Show
         End If
   End Select
End Sub
Private Sub OpenForm_P3(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1:   '爭議案勝訴輸入
         If CheckUse("frm02010501_1", strExec) = True Then
            Call frm02010501_1.SetParent(frm090225)
            frm02010501_1.m_strIR01 = frm090225.m_strIR01
            frm02010501_1.m_strIR02 = frm090225.m_strIR02
            frm02010501_1.m_strIR03 = frm090225.m_strIR03
            frm02010501_1.m_strIR04 = frm090225.m_strIR04
            frm02010501_1.strTM01 = frm090225.txtTi18
            frm02010501_1.strTM02 = frm090225.txtTi19
            frm02010501_1.strTM03 = frm090225.txtTi20
            frm02010501_1.strTM04 = frm090225.txtTi21
            frm02010501_1.m_AppNo = frm090225.m_AppNo
            frm02010501_1.m_RegNo = frm090225.m_RegNo
            frm02010501_1.m_RDate = frm090225.m_strTi12
            frm02010501_1.Show
         End If
      Case 2:   '爭議案敗訴輸入
         If CheckUse("frm02010502_1", strExec) = True Then
            Call frm02010502_1.SetParent(frm090225)
            frm02010502_1.m_strIR01 = frm090225.m_strIR01
            frm02010502_1.m_strIR02 = frm090225.m_strIR02
            frm02010502_1.m_strIR03 = frm090225.m_strIR03
            frm02010502_1.m_strIR04 = frm090225.m_strIR04
            frm02010502_1.strTM01 = frm090225.txtTi18
            frm02010502_1.strTM02 = frm090225.txtTi19
            frm02010502_1.strTM03 = frm090225.txtTi20
            frm02010502_1.strTM04 = frm090225.txtTi21
            frm02010502_1.m_AppNo = frm090225.m_AppNo
            frm02010502_1.m_RegNo = frm090225.m_RegNo
            frm02010502_1.m_RDate = frm090225.m_strTi12
            frm02010502_1.Show
         End If
      Case 3:   '被異議/被評定/被撤銷/對方補充理由/對方延期/通知復審答辯
         If CheckUse("frm02010503_1", strExec) = True Then
            Call frm02010503_1.SetParent(frm090225)
            frm02010503_1.m_strIR01 = frm090225.m_strIR01
            frm02010503_1.m_strIR02 = frm090225.m_strIR02
            frm02010503_1.m_strIR03 = frm090225.m_strIR03
            frm02010503_1.m_strIR04 = frm090225.m_strIR04
            frm02010503_1.strTM01 = frm090225.txtTi18
            frm02010503_1.strTM02 = frm090225.txtTi19
            frm02010503_1.strTM03 = frm090225.txtTi20
            frm02010503_1.strTM04 = frm090225.txtTi21
            frm02010503_1.m_AppNo = frm090225.m_AppNo
            frm02010503_1.m_RegNo = frm090225.m_RegNo
            frm02010503_1.m_RDate = frm090225.m_strTi12
            frm02010503_1.Show
         End If
      Case 4:   '發回補理由/發回補答辯
         If CheckUse("frm02010504_1", strExec) = True Then
            Call frm02010504_1.SetParent(frm090225)
            frm02010504_1.m_strIR01 = frm090225.m_strIR01
            frm02010504_1.m_strIR02 = frm090225.m_strIR02
            frm02010504_1.m_strIR03 = frm090225.m_strIR03
            frm02010504_1.m_strIR04 = frm090225.m_strIR04
            frm02010504_1.strTM01 = frm090225.txtTi18
            frm02010504_1.strTM02 = frm090225.txtTi19
            frm02010504_1.strTM03 = frm090225.txtTi20
            frm02010504_1.strTM04 = frm090225.txtTi21
            frm02010504_1.m_AppNo = frm090225.m_AppNo
            frm02010504_1.m_RegNo = frm090225.m_RegNo
            frm02010504_1.m_RDate = frm090225.m_strTi12
            frm02010504_1.Show
         End If
      Case 5:   '撤銷原處分／和解輸入
         If CheckUse("frm02010505_1", strExec) = True Then
            Call frm02010505_1.SetParent(frm090225)
            frm02010505_1.m_strIR01 = frm090225.m_strIR01
            frm02010505_1.m_strIR02 = frm090225.m_strIR02
            frm02010505_1.m_strIR03 = frm090225.m_strIR03
            frm02010505_1.m_strIR04 = frm090225.m_strIR04
            frm02010505_1.strTM01 = frm090225.txtTi18
            frm02010505_1.strTM02 = frm090225.txtTi19
            frm02010505_1.strTM03 = frm090225.txtTi20
            frm02010505_1.strTM04 = frm090225.txtTi21
            frm02010505_1.m_AppNo = frm090225.m_AppNo
            frm02010505_1.m_RegNo = frm090225.m_RegNo
            frm02010505_1.m_RDate = frm090225.m_strTi12
            frm02010505_1.Show
         End If
      Case 6:   '受理
         If CheckUse("frm02010506_1", strExec) = True Then
            Call frm02010506_1.SetParent(frm090225)
            frm02010506_1.m_strIR01 = frm090225.m_strIR01
            frm02010506_1.m_strIR02 = frm090225.m_strIR02
            frm02010506_1.m_strIR03 = frm090225.m_strIR03
            frm02010506_1.m_strIR04 = frm090225.m_strIR04
            frm02010506_1.strTM01 = frm090225.txtTi18
            frm02010506_1.strTM02 = frm090225.txtTi19
            frm02010506_1.strTM03 = frm090225.txtTi20
            frm02010506_1.strTM04 = frm090225.txtTi21
            frm02010506_1.m_AppNo = frm090225.m_AppNo
            frm02010506_1.m_RegNo = frm090225.m_RegNo
            frm02010506_1.m_RDate = frm090225.m_strTi12
            frm02010506_1.Show
         End If
      Case 7:   '延長審查時間
         If CheckUse("frm02010507_1", strExec) = True Then
            Call frm02010507_1.SetParent(frm090225)
            frm02010507_1.m_strIR01 = frm090225.m_strIR01
            frm02010507_1.m_strIR02 = frm090225.m_strIR02
            frm02010507_1.m_strIR03 = frm090225.m_strIR03
            frm02010507_1.m_strIR04 = frm090225.m_strIR04
            frm02010507_1.strTM01 = frm090225.txtTi18
            frm02010507_1.strTM02 = frm090225.txtTi19
            frm02010507_1.strTM03 = frm090225.txtTi20
            frm02010507_1.strTM04 = frm090225.txtTi21
            frm02010507_1.m_AppNo = frm090225.m_AppNo
            frm02010507_1.m_RegNo = frm090225.m_RegNo
            frm02010507_1.m_RDate = frm090225.m_strTi12
            frm02010507_1.Show
         End If
      Case 8:   '對方撤回
         If CheckUse("frm02010508_1", strExec) = True Then
            Call frm02010508_1.SetParent(frm090225)
            frm02010508_1.m_strIR01 = frm090225.m_strIR01
            frm02010508_1.m_strIR02 = frm090225.m_strIR02
            frm02010508_1.m_strIR03 = frm090225.m_strIR03
            frm02010508_1.m_strIR04 = frm090225.m_strIR04
            frm02010508_1.strTM01 = frm090225.txtTi18
            frm02010508_1.strTM02 = frm090225.txtTi19
            frm02010508_1.strTM03 = frm090225.txtTi20
            frm02010508_1.strTM04 = frm090225.txtTi21
            frm02010508_1.m_AppNo = frm090225.m_AppNo
            frm02010508_1.m_RegNo = frm090225.m_RegNo
            frm02010508_1.m_RDate = frm090225.m_strTi12
            frm02010508_1.Show
         End If
      Case 9:   '其他來函輸入
         If CheckUse("frm02010408_1", strExec) = True Then
            Call frm02010408_1.SetParent(frm090225)
            frm02010408_1.m_strIR01 = frm090225.m_strIR01
            frm02010408_1.m_strIR02 = frm090225.m_strIR02
            frm02010408_1.m_strIR03 = frm090225.m_strIR03
            frm02010408_1.m_strIR04 = frm090225.m_strIR04
            frm02010408_1.strTM01 = frm090225.txtTi18
            frm02010408_1.strTM02 = frm090225.txtTi19
            frm02010408_1.strTM03 = frm090225.txtTi20
            frm02010408_1.strTM04 = frm090225.txtTi21
            frm02010408_1.m_AppNo = frm090225.m_AppNo
            frm02010408_1.m_RegNo = frm090225.m_RegNo
            frm02010408_1.m_RDate = frm090225.m_strTi12
            frm02010408_1.Show
         End If
      Case 10:   '延期受理
         If CheckUse("frm02010407_1", strExec) = True Then
            Call frm02010407_1.SetParent(frm090225)
            frm02010407_1.m_strIR01 = frm090225.m_strIR01
            frm02010407_1.m_strIR02 = frm090225.m_strIR02
            frm02010407_1.m_strIR03 = frm090225.m_strIR03
            frm02010407_1.m_strIR04 = frm090225.m_strIR04
            frm02010407_1.strTM01 = frm090225.txtTi18
            frm02010407_1.strTM02 = frm090225.txtTi19
            frm02010407_1.strTM03 = frm090225.txtTi20
            frm02010407_1.strTM04 = frm090225.txtTi21
            frm02010407_1.m_AppNo = frm090225.m_AppNo
            frm02010407_1.m_RegNo = frm090225.m_RegNo
            frm02010407_1.m_RDate = frm090225.m_strTi12
            frm02010407_1.Show
         End If
      Case 11:   '部分勝部分敗
         If CheckUse("frm02010509_1", strExec) = True Then
            Call frm02010509_1.SetParent(frm090225)
            frm02010509_1.m_strIR01 = frm090225.m_strIR01
            frm02010509_1.m_strIR02 = frm090225.m_strIR02
            frm02010509_1.m_strIR03 = frm090225.m_strIR03
            frm02010509_1.m_strIR04 = frm090225.m_strIR04
            frm02010509_1.strTM01 = frm090225.txtTi18
            frm02010509_1.strTM02 = frm090225.txtTi19
            frm02010509_1.strTM03 = frm090225.txtTi20
            frm02010509_1.strTM04 = frm090225.txtTi21
            frm02010509_1.m_AppNo = frm090225.m_AppNo
            frm02010509_1.m_RegNo = frm090225.m_RegNo
            frm02010509_1.m_RDate = frm090225.m_strTi12
            frm02010509_1.Show
         End If
   End Select
End Sub
Private Sub OpenForm_P4(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1:   '代理人已收達/已提申
         If CheckUse("frm02010601_01", strExec) = True Then
            Call frm02010601_01.SetParent(frm090225)
            frm02010601_01.m_strIR01 = frm090225.m_strIR01
            frm02010601_01.m_strIR02 = frm090225.m_strIR02
            frm02010601_01.m_strIR03 = frm090225.m_strIR03
            frm02010601_01.m_strIR04 = frm090225.m_strIR04
            frm02010601_01.m_TM01 = frm090225.txtTi18
            frm02010601_01.m_TM02 = frm090225.txtTi19
            frm02010601_01.m_TM03 = frm090225.txtTi20
            frm02010601_01.m_TM04 = frm090225.txtTi21
            frm02010601_01.m_RDate = frm090225.m_strTi12
            frm02010601_01.Show
         End If
      Case 2:   '代理人通知修正
         If CheckUse("frm02010602_01", strExec) = True Then
            Call frm02010602_01.SetParent(frm090225)
            frm02010602_01.m_strIR01 = frm090225.m_strIR01
            frm02010602_01.m_strIR02 = frm090225.m_strIR02
            frm02010602_01.m_strIR03 = frm090225.m_strIR03
            frm02010602_01.m_strIR04 = frm090225.m_strIR04
            frm02010602_01.m_TM01 = frm090225.txtTi18
            frm02010602_01.m_TM02 = frm090225.txtTi19
            frm02010602_01.m_TM03 = frm090225.txtTi20
            frm02010602_01.m_TM04 = frm090225.txtTi21
            frm02010602_01.m_RDate = frm090225.m_strTi12
            frm02010602_01.Show
         End If
      Case 3:   '其他來函輸入
         If CheckUse("frm02010408_1", strExec) = True Then
            Call frm02010408_1.SetParent(frm090225)
            frm02010408_1.m_strIR01 = frm090225.m_strIR01
            frm02010408_1.m_strIR02 = frm090225.m_strIR02
            frm02010408_1.m_strIR03 = frm090225.m_strIR03
            frm02010408_1.m_strIR04 = frm090225.m_strIR04
            frm02010408_1.strTM01 = frm090225.txtTi18
            frm02010408_1.strTM02 = frm090225.txtTi19
            frm02010408_1.strTM03 = frm090225.txtTi20
            frm02010408_1.strTM04 = frm090225.txtTi21
            frm02010408_1.m_AppNo = frm090225.m_AppNo
            frm02010408_1.m_RegNo = frm090225.m_RegNo
            frm02010408_1.m_RDate = frm090225.m_strTi12
            frm02010408_1.Show
         End If
   End Select
End Sub
Private Sub OpenForm_P5(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1 '內部收文
         If CheckUse("frm010001", strExec) = True Then
            Call frm010001.SetParent(frm090225)
            frm010001.m_strIR01 = frm090225.m_strIR01
            frm010001.m_strIR02 = frm090225.m_strIR02
            frm010001.m_strIR03 = frm090225.m_strIR03
            frm010001.m_strIR04 = frm090225.m_strIR04
            frm010001.m_strCP01 = frm090225.txtTi18
            frm010001.m_strCP02 = frm090225.txtTi19
            frm010001.m_strCP03 = frm090225.txtTi20
            frm010001.m_strCP04 = frm090225.txtTi21
            frm010001.m_RDate = frm090225.m_strTi12
            frm010001.intChoose = 1
            frm010001.intReceiveKind = 0
            frm010001.intModifyKind = 0
            frm010001.Caption = "內部收文－新增"
         End If
      Case 2 '帳單輸入
         'Modify By Sindy 2021/1/18
         If CheckUse("Frmacc21u0", strExec) = True Then
            ToolShow
            'tool1_enabled
            tool3_enabled
            Frmacc21u0.m_strIR01 = frm090225.m_strIR01
            Frmacc21u0.m_strIR02 = frm090225.m_strIR02
            Frmacc21u0.m_strIR03 = frm090225.m_strIR03
            Frmacc21u0.m_strIR04 = frm090225.m_strIR04
            Frmacc21u0.m_strCP01 = frm090225.txtTi18
            Frmacc21u0.m_strCP02 = frm090225.txtTi19
            Frmacc21u0.m_strCP03 = frm090225.txtTi20
            Frmacc21u0.m_strCP04 = frm090225.txtTi21
            Frmacc21u0.m_RDate = frm090225.m_strTi12
            Set Frmacc21u0.m_PrevForm = frm090225
            Frmacc21u0.Show
         End If
         '2021/1/18 END
'         If CheckUse("Frmacc2150", strExec) = True Then
'            ToolShow
'            tool1_enabled
'            Frmacc2150.m_strIR01 = frm090225.m_strIR01
'            Frmacc2150.m_strIR02 = frm090225.m_strIR02
'            Frmacc2150.m_strIR03 = frm090225.m_strIR03
'            Frmacc2150.m_strIR04 = frm090225.m_strIR04
'            Frmacc2150.m_CP01 = frm090225.txtTi18
'            Frmacc2150.m_CP02 = frm090225.txtTi19
'            Frmacc2150.m_CP03 = frm090225.txtTi20
'            Frmacc2150.m_CP04 = frm090225.txtTi21
'            Frmacc2150.m_RDate = frm090225.m_strTi12
'            Set Frmacc2150.m_ParentForm = Me
'            Frmacc2150.Show
'         End If
      Case 3 '抵帳單輸入
         If CheckUse("Frmacc2160", strExec) = True Then
            ToolShow
            tool1_enabled
            Frmacc2160.m_strIR01 = frm090225.m_strIR01
            Frmacc2160.m_strIR02 = frm090225.m_strIR02
            Frmacc2160.m_strIR03 = frm090225.m_strIR03
            Frmacc2160.m_strIR04 = frm090225.m_strIR04
            Frmacc2160.m_strCP01 = frm090225.txtTi18
            Frmacc2160.m_strCP02 = frm090225.txtTi19
            Frmacc2160.m_strCP03 = frm090225.txtTi20
            Frmacc2160.m_strCP04 = frm090225.txtTi21
            Frmacc2160.m_RDate = frm090225.m_strTi12
            Set Frmacc2160.m_PrevForm = frm090225
            Frmacc2160.Show
         End If
      Case 4 '帳單作廢輸入
         If CheckUse("Frmacc21j0", strExec) = True Then
            ToolShow
            tool1_enabled
            Frmacc21j0.m_strIR01 = frm090225.m_strIR01
            Frmacc21j0.m_strIR02 = frm090225.m_strIR02
            Frmacc21j0.m_strIR03 = frm090225.m_strIR03
            Frmacc21j0.m_strIR04 = frm090225.m_strIR04
            Frmacc21j0.m_strCP01 = frm090225.txtTi18
            Frmacc21j0.m_strCP02 = frm090225.txtTi19
            Frmacc21j0.m_strCP03 = frm090225.txtTi20
            Frmacc21j0.m_strCP04 = frm090225.txtTi21
            Frmacc21j0.m_RDate = frm090225.m_strTi12
            Set Frmacc21j0.m_PrevForm = frm090225
            Frmacc21j0.Show
         End If
   End Select
End Sub
'2019/5/10 END

'Add By Sindy 2018/11/14
Public Sub EDocSubOpenForm(pType As Integer, Index As Integer, _
   ByRef oForm As Form, ByRef iStiu As Integer)
   
   'T商申
   If pType = 1 Then
      Select Case Index
      Case 1   '非爭議案核准輸入
         Set oForm = frm02010401_1

      Case 2   '非爭議案核駁輸入
         Set oForm = frm02010402_1

      Case 3:   '審查報告輸入
         Set oForm = frm02010403_1

      'Added by Morgan 2023/1/13 證書只會有紙本
      Case 4:   '註冊證輸入
         Set oForm = frm02010404_1

      Case 5:   '非爭議案取消催審期限
         Set oForm = frm02010405_1

      Case 6:   '商標案被禁止處分
         Set oForm = frm02010406_1

      Case 7:   '延期受理
         Set oForm = frm02010407_1

      Case 8:   '其他來函輸入
         Set oForm = frm02010408_1

      'Removed by Morgan 2017/4/24 主管機關不會是智慧局--秀玲
      'Case 9:   '服務業務結果輸入
      '   Set oForm = frm02010409_1
      'Removed by Morgan 2017/4/24 大陸才有已經沒有用了--秀玲
      'Case 10:   '廣告刊出來函輸入
      '   Set oForm = frm02010410_1

      Case 11:   '智慧局註冊費通知函輸入
         Set oForm = frm02010411_1

      End Select

   'T商爭
   ElseIf pType = 2 Then
      Select Case Index
      Case 1:   '爭議案勝訴輸入
         Set oForm = frm02010501_1

      Case 2:   '爭議案敗訴輸入
         Set oForm = frm02010502_1

      Case 3:   '被異議/被評定/被撤銷/對方補充理由/對方延期/通知復審答辯
         Set oForm = frm02010503_1

      Case 4:   '發回補理由/發回補答辯
         Set oForm = frm02010504_1

      Case 5:   '撤銷原處分／和解輸入
         Set oForm = frm02010505_1

      Case 6:   '受理
         Set oForm = frm02010506_1

      Case 7:   '延長審查時間
         Set oForm = frm02010507_1

      Case 8:   '對方撤回
         Set oForm = frm02010508_1

      Case 9:   '其他來函輸入
         Set oForm = frm02010408_1

      Case 10:   '延期受理
         Set oForm = frm02010407_1

      Case 11:   '部分勝部分敗
         Set oForm = frm02010509_1

      End Select
      
   'FCT商申
   Else
'      Select Case Index
'      Case 1   '非爭議案核准輸入
'         Set oForm = frm03020401_01
'
'      Case 2   '非爭議案核駁輸入
'         Set oForm = frm03020402_01
'
'      Case 3:   '審查報告輸入
'         Set oForm = frm03020403_01
'
'      Case 5:   '非爭議案取消催審期限
'         Set oForm = frm03020405_01
'         iStiu = 1
'
'      Case 6:   '商標案被禁止處分
'         Set oForm = frm03020406_01
'
'      Case 7:   '延期受理
'         Set oForm = frm03020407_01
'
'      Case 8:   '其他來函輸入
'         Set oForm = frm03020408_01
'         iStiu = 1
'
'      Case 12:   '通知已轉他所
'         Set oForm = frm03020405_01
'         iStiu = 2
'      End Select
   End If
End Sub

Private Sub mnuPopEDoc0101_Click(Index As Integer)
   frm02010412.OpenForm 1, Index
End Sub

Private Sub mnuPopEDoc0201_Click(Index As Integer)
   frm02010412.OpenForm 2, Index
End Sub

Private Sub mnuPopEDoc0301_Click(Index As Integer)
   frm02010412.OpenForm 3, Index
End Sub

'Add by Morgan 2005/3/2 控制不可拷貝畫面
Private Sub Timer3_Timer()
On Error Resume Next 'Added by Morgan 2017/8/29 若有其他軟體也在使用剪貼簿時會發生521(無法開啟剪貼簿)的錯誤(Ex.Word開啟剪貼簿並擷取畫面)
   
   'Added by Morgan 2024/8/8 定時執行一次語法以確保跨網段連線時網路不會被切斷
   Static dtNow As Date
   
   If tmrConnect.Interval = 0 Then
      If Now > dtNow Then
         dtNow = DateAdd("n", cntAutoQueryInterval, Now)
         ClsLawReadRstMsg 1, "select * from dual"
      End If
   End If
   'end 2024/8/8

   'add by nickc 2005/05/02 電腦中心的不管
   If Pub_StrUserSt03 = "M51" Or Pub_Can_Copy_Pic = True Then Exit Sub
   '圖檔才清
   If Clipboard.GetFormat(1) = False And Clipboard.GetFormat(2) = True And Clipboard.GetFormat(3) = False Then
       Clipboard.Clear
   End If
End Sub

'Add by Morgan 2003/12/23
'控制連線閒置超過30分鐘自動離線
Private Sub tmrConnect_Timer()
   tmrConnect.Tag = Val(tmrConnect.Tag) + 1
   'Modify by Morgan 2005/2/3 改成10分鐘--副理
   'If tmrConnect.Tag = 30 Then
   If tmrConnect.Tag = 10 Then
      Timer1.Enabled = False

'Modify by Morgan 2005/1/10 改保留原畫面不結束
'      Call CloseAllChild
'      Call SwitchMenu(False)
'2005/1/10 end

      'Remove by Morgan 2005/2/23 移到重連線視窗
      'cnnConnection.Close

      bolReOpen = False
      PUB_SendMailCache 'Add by Morgan 2010/6/11
      frmReopen.Show vbModal, Me
      If bolReOpen = True Then
         Call ReConnect
      Else
         Call mnu00_Click(1)
      End If
   End If
End Sub

Private Sub MDIForm_Load()
'Add by Morgan 2003/12/23
'控制連線閒置超過30分鐘自動關閉程式
If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
   Set eventConn = cnnConnection
   tmrConnect.Interval = 60000
End If
'Add end 2003/12/23

Dim strSysKind As String
Dim lngValue, lngBufferSize As Long, intCounter As Integer
Dim strUserId As String * 10, strLocalId As String
   
    '若成功登入
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
            mnuChUser.Visible = True 'Add by Sindy 2015/2/12
            tmrConnect.Interval = 0 'Added by Morgan 2023/6/1
       Else
            mnuDML(0).Visible = False
            mnuChUser.Visible = False 'Add by Sindy 2015/2/12
       End If
       
'       'Added by Morgan 2016/1/22 薪資查詢測試
'       If Pub_StrUserSt03 = "M51" Then
'         mnu23(6).Visible = True
'       Else
'         mnu23(6).Visible = False
'       End If
'       'end
         
       'Add By Sindy 2010/01/07 M51及王副總才可以看到
       'If Pub_StrUserSt03 = "M51" Or strUserNum = "71011" Then
       'If Pub_StrUserSt03 = "M51" Or CheckUse("frm050207", strExec) = True Then
       If Pub_StrUserSt03 = "M51" Then
          mnu0202(8).Visible = True
'          mnu0201(8).Visible = True '待送件區
       Else
          mnu0202(8).Visible = False
'          mnu0201(8).Visible = False '待送件區
       End If
       'Add by Amy 2018/08/09 待處理區
       If Val(strSrvDate(1)) < 非P結案電子化啟用日 Then
          mnu0201(9).Visible = False
       End If
       'end 2018/08/09
       
       'Add By Sindy 2023/8/4
       '將原訂的公報資訊，設計至背面，即可於一張Ａ４紙完整揭露
       '可上線時間訂於2023年9月1日，屆時將取消公報單獨寄送
       If strSrvDate(1) >= 20230901 Then
         mnu0206(18).Visible = False '公報開拓資料維護
         mnu0206(19).Visible = False '商標公報開拓函列印
       End If
       '2023/8/4 END
      
       If bolFNation = False Then
    'Ken 90/07/06
    '      mnu10(14).Visible = False
    '      mnu10(15).Caption = "以國籍查詢申請人"
          mnu101(10).Visible = False 'Modify by Amy 2014/05/05
          mnu102(7).Visible = False
    'Ken 90/07/06
          mnu101(6).Visible = False 'Modify by Amy 2014/05/05
          '92.3.17 add by sonia
          mnu102(2).Visible = False
          '92.3.17 end
          'Add By Cheng 2003/08/13
          '智權人員收/發文量比較查詢
          '2005/8/3 CANCEL BY SONIA
          'mnu10(23).Visible = False
       End If
       'add by nickc 2008/05/01
'        If Pub_StrUserSt03 = "M51" Then
'            mnu2103(3).Visible = True
'        Else
'            mnu2103(3).Visible = False
'        End If
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
   PUB_SendMailCache 'Add by Morgan 2010/6/11
   PUB_AddAuditLog AL_登出 'Added by Morgan 2025/7/31
   
'edit by nickc 2007/02/06 不用 dll 了
'Set obj001 = Nothing
'Set objPublicData = Nothing
   ' 90.08.16 modify by louis
   EndOfficeAp
End Sub

'Modify by Morgan 2005/12/16 加切換連線選擇
Private Sub mnu00_Click(Index As Integer)
   Select Case Index
      Case 0 '切換連線
         If PUB_Connect2DB(True) = False Then
            Unload Me
         End If
      Case 1 '結束
         bolUnloading = True 'Add by Morgan 2011/3/11
         Unload Me
   End Select
End Sub

Private Sub mnu0201_Click(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1: '分案
         If CheckUse("frm020101_01", strExec) = True Then
            frm020101_01.Show
         End If
      Case 2: '發文
         If CheckUse("frm020102_01", strExec) = True Then
            frm020102_01.Show
         End If
      Case 3: '商標申請案號輸入
         If CheckUse("frm02010301_1", strExec) = True Then
            frm02010301_1.Show
         End If
      'Added by Morgan 2017/3/15
      Case 7: '電子公文來函
         If CheckUse("frm02010412", strExec) = True Then
            frm02010412.Show
         End If
      'Add By Sindy 2018/4/26
      Case 8 '待送件區
         If CheckUse("frm090202_4", strExec) = True Then
            frm090202_4.m_ProState = "T"
            frm090202_4.Show
         End If
      'Add by Amy 2018/06/21 '待處理區
      Case 9
        If CheckUse("frm210149", strExec) = True Then
           frm210149.m_ProState = "T"
           frm210149.Show
        End If
      'Add By Sindy 2018/10/5
      Case 10 '掃瞄資料匯入
         If CheckUse("frm010033", strExec) = True Then
            frm010033.Show
         End If
      'Add by Sindy 2020/7/30
      Case 11 '電子收據匯入
        If CheckUse("frm040121", strExec) = True Then
            frm040121.m_ProState = "T"
            frm040121.Show
         End If
      'Add By Sindy 2019/12/30
      Case 12 '收據/回執整批匯入
         If CheckUse("frm040115", strExec) = True Then
            frm040115.m_ProState = "T"
            frm040115.Show
         End If
      'Modify by Amy 2020/01/08 加公文來函文檔整批匯入,調整順序
      Case 13 '公文來函文檔整批匯入
        If CheckUse("frm040112", strExec) = True Then
            frm040112.m_ProState = "T"
            frm040112.Show
        End If
      'Add by Amy 2020/01/08
      Case 14 '代理人來函匯入
        If CheckUse("frm040120", strExec) = True Then
            frm040120.m_ProState = "T" 'Add by Amy 2020/02/19
            frm040120.Show
        End If
   End Select
End Sub

'Add By Sindy 2018/5/2
Public Sub frm090202_4CallFrm(strSendFrmType As String, m_PA11 As String, _
   m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, _
   m_EEP01 As String)
Dim m_SendRecvForm As Form '發文作業
   Select Case strSendFrmType
      Case "T"
         Set m_SendRecvForm = frm020102_01
         m_SendRecvForm.Show
         m_SendRecvForm.bolIsEMPFlow = True
         m_SendRecvForm.radio(1).Value = True
         If m_CP01 = "TF" Then
            m_SendRecvForm.textTM01 = m_CP01
            m_SendRecvForm.textTM02 = Left(m_CP02, 5)
            m_SendRecvForm.textTM02_2 = Right(m_CP02, 1)
            m_SendRecvForm.textTM03 = m_CP03
            m_SendRecvForm.textTM04 = m_CP04
         Else
            m_SendRecvForm.textTM01 = m_CP01
            m_SendRecvForm.textTM02 = m_CP02
            m_SendRecvForm.textTM03 = m_CP03
            m_SendRecvForm.textTM04 = m_CP04
         End If
         m_SendRecvForm.cmdQuery.Value = True
         If m_SendRecvForm.grdList.Rows = 2 Then
            m_SendRecvForm.cmdOK.Value = True
         End If
         Set m_SendRecvForm = Nothing
   End Select
End Sub

Private Sub mnu020104_Click(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1:   '非爭議案核准輸入
         If CheckUse("frm02010401_1", strExec) = True Then
            frm02010401_1.Show
         End If
      Case 2:   '非爭議案核駁輸入
         If CheckUse("frm02010402_1", strExec) = True Then
            frm02010402_1.Show
         End If
      Case 3:   '審查報告輸入
         If CheckUse("frm02010403_1", strExec) = True Then
            frm02010403_1.Show
         End If
      Case 4:   '註冊證輸入
         If CheckUse("frm02010404_1", strExec) = True Then
            frm02010404_1.Show
         End If
      Case 5:   '非爭議案取消催審期限
         If CheckUse("frm02010405_1", strExec) = True Then
            frm02010405_1.Show
         End If
      Case 6:   '商標案被禁止處分
         If CheckUse("frm02010406_1", strExec) = True Then
            frm02010406_1.Show
         End If
      Case 7:   '延期受理
         If CheckUse("frm02010407_1", strExec) = True Then
            frm02010407_1.Show
         End If
      Case 8:   '其他來函輸入
         If CheckUse("frm02010408_1", strExec) = True Then
            frm02010408_1.Show
         End If
      Case 9:   '服務業務結果輸入
         If CheckUse("frm02010409_1", strExec) = True Then
            frm02010409_1.Show
         End If
      Case 10:   '廣告刊出來函輸入
         If CheckUse("frm02010410_1", strExec) = True Then
            frm02010410_1.Show
         End If
      Case 11:   '智慧局註冊費通知函輸入
         If CheckUse("frm02010411_1", strExec) = True Then
            frm02010411_1.Show
         End If
      'Add By Sindy 2010/4/7
      Case 12:   '大陸商標審定公告及通知續展匯入作業
         If CheckUse("frm020320", strExec) = True Then
            frm020320.Show
         End If
      '2010/4/7 End
   End Select
End Sub

Private Sub mnu020105_Click(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1:   '爭議案勝訴輸入
         If CheckUse("frm02010501_1", strExec) = True Then
            frm02010501_1.Show
         End If
      Case 2:   '爭議案敗訴輸入
         If CheckUse("frm02010502_1", strExec) = True Then
            frm02010502_1.Show
         End If
      Case 3:   '被異議/被評定/被撤銷/對方補充理由/對方延期/通知復審答辯
         If CheckUse("frm02010503_1", strExec) = True Then
            frm02010503_1.Show
         End If
      Case 4:   '發回補理由/發回補答辯
         If CheckUse("frm02010504_1", strExec) = True Then
            frm02010504_1.Show
         End If
      Case 5:   '撤銷原處分／和解輸入
         If CheckUse("frm02010505_1", strExec) = True Then
            frm02010505_1.Show
         End If
      Case 6:   '受理
         If CheckUse("frm02010506_1", strExec) = True Then
            frm02010506_1.Show
         End If
      Case 7:   '延長審查時間
         If CheckUse("frm02010507_1", strExec) = True Then
            frm02010507_1.Show
         End If
      Case 8:   '對方撤回
         If CheckUse("frm02010508_1", strExec) = True Then
            frm02010508_1.Show
         End If
      Case 9:   '其他來函輸入
         If CheckUse("frm02010408_1", strExec) = True Then
            frm02010408_1.Show
         End If
      'Add By Cheng 2002/02/01
      Case 10:   '延期受理
         If CheckUse("frm02010407_1", strExec) = True Then
            frm02010407_1.Show
         End If
      'Add By Sindy 98/04/13
      Case 11:   '部分勝部分敗
         If CheckUse("frm02010509_1", strExec) = True Then
            frm02010509_1.Show
         End If
   End Select
End Sub

Private Sub mnu020106_Click(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1:   '代理人已收達/已提申
         If CheckUse("frm02010601_01", strExec) = True Then
            frm02010601_01.Show
         End If
      Case 2:   '代理人通知修正
         If CheckUse("frm02010602_01", strExec) = True Then
            frm02010602_01.Show
         End If
      Case 3:   '其他來函輸入
         If CheckUse("frm02010408_1", strExec) = True Then
            frm02010408_1.Show
         End If
   End Select
End Sub

Private Sub mnu0202_Click(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1 '代理人新案案件統計
         If CheckUse("frm050201", strExec) = True Then
            StrStartSystemByNick = "T,TF,TS,TB,TC,TD,TM,TR,TT"
            frm050201.Show
         End If
'edit by nickc 2005/07/22
'      Case 2 '案件結餘查詢
'         If CheckUse("frm040202", strExec) = True Then
'            frm040202.Show
'         End If
      Case 3 '未請款明細查詢
         If CheckUse("frm050203", strExec) = True Then
            StrStartSystemByNick = "T,TF,TS,TB,TC,TD,TM,TR,TT"
            frm050203.Show
         End If
      Case 4 '審查委員准駁統計
         If CheckUse("frm040204", strExec) = True Then
            StrStartSystemByNick = "T,FCT,CFT,TF"
            frm040204.Show
         End If
      Case 5 'FC收款請款點數查詢
         If CheckUse("frm040205", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm040205.Show
         End If
      'Add By Cheng 2002/09/24
      Case 6 '代理人案件性質統計
         If CheckUse("frm050204_1", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm050204_1.Show
         End If
      'add by nickc 2007/07/31 延展前商標無效管制表
      Case 7
         If CheckUse("frm040207", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm040207.Show
         End If
      'Add By Sindy 2010/01/07 員工查詢印表記錄檔查詢
      Case 8
         If CheckUse("frm050207", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm050207.Show
         End If
      'Add By Sindy 2020/1/13 商標處期限通知
      Case 9
         If CheckUse("frm020201", strExec) = True Then
            frm020201.Show
         End If
      Case Else
   End Select
End Sub

Private Sub mnu0203_Click(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1 '智權人員期限管制表
         If CheckUse("frm020301", strExec) = True Then
           frm020301.Show
         End If
      Case 2 '承辦人期限管制表
         If CheckUse("frm020302", strExec) = True Then
           frm020302.Show
         End If
      Case 3 '代理人案件收達/提申管制表
         If CheckUse("frm050303", strExec) = True Then
           frm050303.Show
         End If
      Case 4 '對外案件延展未提申明細表
         If CheckUse("frm020310", strExec) = True Then
           frm020310.Show
         End If
      Case 5 '收文未發文明細表
         If CheckUse("frm050304", strExec) = True Then
           frm050304.Show
         End If
      Case 6 '催審函/催審表
         If CheckUse("frm020305", strExec) = True Then
           frm020305.Show
         End If
      Case 7 '智權人員收文明細表
         If CheckUse("frm020306", strExec) = True Then
           frm020306.Show
         End If
      Case 8 '承辦人案件明細表
         If CheckUse("frm020307", strExec) = True Then
           frm020307.Show
         End If
      Case 9 '申請意見書案件明細表
         If CheckUse("frm020308", strExec) = True Then
           frm020308.Show
         End If
      Case 10 '商品類別/組群案件明細表
         If CheckUse("frm020309", strExec) = True Then
           frm020309.Show
         End If
      Case 11 '後金案件表
         If CheckUse("frm050314", strExec) = True Then
           frm050314.Show
         End If
      Case 12 '延期明細表
         If CheckUse("frm050315", strExec) = True Then
           frm050315.Show
         End If
      Case 13 '不出名案件明細表
         If CheckUse("frm020312", strExec) = True Then
           frm020312.Show
         End If
      Case 14 '代理人案件總簿
         If CheckUse("frm050316", strExec) = True Then
           frm050316.Show
         End If
      Case 15 '客戶案件總簿
         If CheckUse("frm050317", strExec) = True Then
           'Modify by Amy 2017/08/04 原:frm050317
           frm0503171.Show
         End If
      Case 16 '代理人/申請人名單
         If CheckUse("frm050318", strExec) = True Then
           frm050318.Show
         End If
'edit by nickc 2005/11/10 取消舊結餘單
'      Case 17
'         If CheckUse("frm040320", strExec) = True Then
'           frm040320.Show
'         End If
      Case 18 '地址條列印
         If CheckUse("frm083014", strExec) = True Then
           frm083014.Show
         End If
      'add by nickc 2006/06/14
      Case 19 '內商-國外FC帳款明細表
         If CheckUse("frm050324", strExec) = True Then
           frm050324.Show
         End If
      '2007/9/10 ADD BY SONIA 智慧局註冊費通知函定稿地址條列印
      Case 20 '智慧局註冊費通知函列印
         If CheckUse("frm020311", strExec) = True Then
           frm020311.Show
         End If
      'Add By Sindy 98/03/20
      Case 21 '下載商標圖參考報表
         If CheckUse("frm020319", strExec) = True Then
            frm020319.Show
         End If
      'Add By Sindy 101/7/13
      Case 22 '台灣商標延展開拓(貝爾)
         If CheckUse("frm020321", strExec) = True Then
            frm020321.Show
         End If
      'Add By Sindy 2018/11/9
      Case 23 '台灣商標公告近三年開拓函
         If CheckUse("frm020322", strExec) = True Then
            frm020322.Show
         End If
      'Add By Sindy 2019/1/4
      Case 24 '台灣商標延展開拓(智慧局)
         If CheckUse("frm020323", strExec) = True Then
            frm020323.Show
         End If
      'Add By Sindy 2025/4/16
      Case 25 '期限通知檢核及報表列印
         If CheckUse("frm040335T", strExec) = True Then
            frm040335.m_ProState = "T"
            frm040335.Show
         End If
   End Select
End Sub

Private Sub mnu0204_Click(Index As Integer)
   ToolHide
   intPCaseKind = 商標
   intPWhere = 國內
   Select Case Index
      Case 1 '商申案智權人員收/發文統計表
         If CheckUse("frm020401", strExec) = True Then
           frm020401.Show
         End If
      Case 2 '商申案智權人員准駁統計表
         If CheckUse("frm020402", strExec) = True Then
           frm020402.Show
         End If
      Case 3 '商爭案智權人員收/發文統計表
         If CheckUse("frm020403", strExec) = True Then
           frm020403.Show
         End If
      Case 4 '商爭案智權人員勝敗統計表
         If CheckUse("frm020404", strExec) = True Then
           frm020404.Show
         End If
      Case 5 '商申案承辦人收/發文統計表
         If CheckUse("frm020405", strExec) = True Then
           frm020405.Show
         End If
      Case 6 '商申案承辦人准駁統計表
         If CheckUse("frm020406", strExec) = True Then
           frm020406.Show
         End If
      Case 7 '商爭案承辦人收/發文統計表
         If CheckUse("frm020407", strExec) = True Then
           frm020407.Show
         End If
      Case 8 '商爭案承辦人勝敗統計表
         If CheckUse("frm020408", strExec) = True Then
           frm020408.Show
         End If
      Case 9 '商爭案承辦人預估勝敗統計表
         If CheckUse("frm020409", strExec) = True Then
           frm020409.Show
         End If
      Case 10 '申請案收/發件數月統計表
         If CheckUse("frm020410", strExec) = True Then
           frm020410.Show
         End If
      Case 11 '各區收/發文達成比較表
         If CheckUse("frm020411", strExec) = True Then
           frm020411.Show
         End If
      Case 12 '各區收/發文件數明細表
         If CheckUse("frm020412", strExec) = True Then
           frm020412.Show
         End If
      Case 13 '代理人新案案件統計表
         If CheckUse("frm050404", strExec) = True Then
           frm050404.Show
         End If
        'Add By Cheng 2003/12/02
      Case 14 '代理人新案案件年度統計表
         If CheckUse("frm050407", strExec) = True Then
           frm050407.Show
         End If
        'End
      Case 15 '代理人案件准駁統計表
         If CheckUse("frm020414", strExec) = True Then
           frm020414.Show
         End If
      Case 16 '代理人/申請人新申請案排行榜
         If CheckUse("frm050405", strExec) = True Then
           StrStartSystemByNick = "T"
           frm050405.Show
         End If
      Case 17 '逾期未結案統計表
         If CheckUse("frm084004", strExec) = True Then
           frm084004.Tag = 5
           frm084004.Show
         End If
      Case 18 '商標承辦人績效表
         If CheckUse("frm020417", strExec) = True Then
           frm020417.Show
         End If
      'add by nickc 2006/07/04
      Case 19 '收/發文件數月報
         If CheckUse("frm020418", strExec) = True Then
           frm020418.Show
         End If
      'Add By Sindy 2012/5/7
      Case 20 '台灣商標爭議案件補充資料次數明細及統計
         If CheckUse("frm020419", strExec) = True Then
           frm020419.Show
         End If
       'Add By Amy 2019/05/28
      Case 21 'MCT收發文件數及點數統計
         If CheckUse("frm020420", strExec) = True Then
           frm020420.Show
         End If
      'Added by Lydia 2023/02/01
      Case 22 '大陸商申查名統計表
         If CheckUse("frm020421", strExec) = True Then
           frm020421.Show
         End If
   End Select
End Sub

Private Sub mnu0205_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1: '商標基本資料維護
         If CheckUse("frm020501", strExec) = True Then
            frm020501.SetSystem 0
            frm020501.Show
         End If
      Case 3 '案件進度檔資料維護
         If CheckUse("frm075004_1", strExec) = True Then
            '91.12.8 ADD BY SONIA
            strSysKind = "T"
            '91.12.8 END
            frm075004_1.Show
         End If
      Case 4 '下一程序資料
         If CheckUse("frm075007_1", strExec) = True Then
            '91.12.8 ADD BY SONIA
            strSysKind = "T"
            '91.12.8 END
            frm075007_1.Show
         End If
      Case 5 '國外代理人資料維護
         If CheckUse("frm050705", strExec) = True Then
            frm050705.Show
         End If
      Case 6 '變更事項
         If CheckUse("frm050706", strExec) = True Then
            frm050706.Show
         End If
      Case 7 '延期記錄資料維護
         If CheckUse("frm050707", strExec) = True Then
            frm050707.Show
         End If
      Case 8 '案件國家收費表維護
         If CheckUse("frm12040102", strExec) = True Then
            frm12040102.Show
         End If
      Case 9 '商標條款資料維護
         If CheckUse("frm020507", strExec) = True Then
            frm020507.Show
         End If
      Case 10 '主張內容分類資料維護
         If CheckUse("frm020508", strExec) = True Then
            frm020508.Show
         End If
      Case 12 '代理人變更名稱作業
         If CheckUse("frm140103", strExec) = True Then
            frm140103.Show
         End If
      Case 13 '承辦人目標點數資料
         If CheckUse("frm020505", strExec) = True Then
            frm020505.Show
         End If
      Case 14 '客戶基本資料維護
         If CheckUse("frm140401", strExec) = True Then
            frm140401.Show
         End If
      Case 15 '著作權登記項目資料維護
         If CheckUse("frm050708", strExec) = True Then
            frm050708.Show
         End If
      'Add By Sindy 2012/4/10
      Case 16 '非本所實質客戶資料維護
         If CheckUse("frm12040155", strExec) = True Then
            frm12040155.Show
         End If
      'Add By Sindy 2014/10/27
      Case 17 '更換FC代理人作業
         If CheckUse("frm110104_1", strExec) = True Then
            frm110104_1.Show
         End If
      'Added by Lydia 2024/01/10
      Case 18  '系統特殊設定
         If CheckUse("frm050716", strExec) = True Then
            frm050716.Show
         End If
      'Add By Sindy 2024/5/22
      Case 19 '註冊費報價資料維護
         If CheckUse("frm050720", strExec) = True Then
            frm050720.Show
         End If
   End Select
End Sub

Private Sub mnu020502_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1: '服務業務基本資料維護 (條碼)
         If CheckUse("frm02050201", strExec) = True Then
            frm02050201.Show
         End If
      Case 2: '服務業務基本資料維護 (監視系統)
         If CheckUse("frm02050202", strExec) = True Then
            frm02050202.Show
         End If
      Case 3: '服務業基本資料維護 (網域)
         If CheckUse("frm02050203", strExec) = True Then
            frm02050203.Show
         End If
      Case 4: '服務業基本資料維護 (著作權)
         If CheckUse("frm02050204", strExec) = True Then
            frm02050204.SetSystem 0
            frm02050204.Show
         End If
      Case 5: '服務業基本資料維護 (其它業務)
         If CheckUse("frm02050205", strExec) = True Then
            frm02050205.SetSystem 0
            frm02050205.Show
         End If
   End Select
End Sub

Private Sub mnu0206_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '商標公報資料維護
         If CheckUse("frm030602", strExec) = True Then
            frm030602.Show
         End If
      Case 2 '更新審定號作業
         If CheckUse("frm030603", strExec) = True Then
            frm030603.Show
         End If
      'Add By Cheng 2002/01/16
      Case 3 '國內商標公報查詢
         If CheckUse("frm030614", strExec) = True Then
            frm030614.Show
         End If
        'Add By Cheng 2003/11/18
      Case 4 '國內公報資料檢核表
         If CheckUse("frm030615", strExec) = True Then
            frm030615.Show
         End If
      Case 5 '國內公報代理人資料維護
         If CheckUse("frm030601", strExec) = True Then
            frm030601.Show
         End If
      Case 6 '國內公報代理人合併作業
         If CheckUse("frm030604", strExec) = True Then
            frm030604.Show
         End If
      Case 7 '國內公報代理人/事務所名稱查詢
         If CheckUse("frm030605", strExec) = True Then
            frm030605.Show
         End If
      Case 8 '表一＆表二.商標全國市場統計表
         If CheckUse("frm030606", strExec) = True Then
            frm030606.Show
         End If
      Case 9 '表三.各區市場佔有率統計表
         If CheckUse("frm030608", strExec) = True Then
            frm030608.Show
         End If
      Case 10 '表四.各類別市場佔有統計表
         If CheckUse("frm030609", strExec) = True Then
            frm030609.Show
         End If
      Case 11 '表五.國外市場排名
         If CheckUse("frm030610", strExec) = True Then
            frm030610.Show
         End If
      Case 12 '表一～表五
         If CheckUse("frm030611", strExec) = True Then
            frm030611.Show
         End If
      Case 13 '國外前十大申請國及其商品類別排名
         If CheckUse("frm030612", strExec) = True Then
            frm030612.Show
         End If
      Case 14 '代理人國外案件排名分析表
         If CheckUse("frm030613", strExec) = True Then
            frm030613.Show
         End If
      'Add By Sindy 2011/5/30
      Case 15 '國內公報代理人資料列印
         If CheckUse("frm04060109", strExec) = True Then
            frm04060109.strTA01 = "T"
            frm04060109.Show
         End If
      'Add By Sindy 2011/11/21
      Case 16 '商標公報轉檔作業
         If CheckUse("frm030616", strExec) = True Then
            'Modify By Sindy 2018/12/12
            frm030616.Option1(0).Value = True '公報
            frm030616.Caption = "商標公報轉檔作業"
            frm030616.Show vbModal
         End If
      Case 17 '公報特定公司不列印者／公報特殊字對照檔
         If CheckUse("frm030617", strExec) = True Then
            frm030617.Show
         End If
      Case 18 '公報開拓資料維護
         If CheckUse("frm030618", strExec) = True Then
            frm030618.Show
         End If
      Case 19 '商標公報開拓函列印
         If CheckUse("frm030619", strExec) = True Then
            frm030619.Show
         End If
      '2011/11/21 End
      'Add By Sindy 2011/12/29
      Case 20 '商標公報大陸清單列印
         If CheckUse("frm030620", strExec) = True Then
            frm030620.Show
         End If
      'Add By Sindy 2014/2/21
      'Modified by Lydia 2015/12/08 改成"商標公報資料統計-Excel"申請人國籍及洲別統計(含同業)
'      Case 21 '本所案件申請人國籍統計表
'         If CheckUse("frm030621", strExec) = True Then
'            frm030621.Show
'         End If
      'Modified by Lydia 2015/12/18 保留原程式,更名為frm030621_1
      Case 21 '本所案件申請人國籍統計表
         If CheckUse("frm030621_1", strExec) = True Then
            frm030621_1.Show
         End If
   End Select
End Sub

Private Sub mnu11_Click(Index As Integer)
   ToolHide
   Select Case Index
        'Add By Cheng 2002/12/09
      Case 3 '來文資料稽核表
         If CheckUse("frm12040121", strExec) = True Then
            frm12040121.Show
         End If
      'Add By Cheng 2002/08/21
      Case 4 '聯絡單列印及E-Mail
         frm1106.Show
      Case 5 '相關卷號
         If CheckUse("frm1103_1", strExec) = True Then
            frm1103_1.Show
         End If
      Case 6 '多國案卷號關係建立
         If CheckUse("frm1104", strExec) = True Then
            frm1104.Show
         End If
      'Add By Cheng 2004/03/16
      Case 7 '分割案件關係維護
         If CheckUse("frm02010604_1", strExec) = True Then
            frm02010604_1.Show
         End If
      'Add By Cheng 2003/06/26
      Case 8 '撰寫信函作業
'         If CheckUse("frm090401", strExec) = True Then
            frm090401.Show
'         End If
      'Add by Morgan 2004//13
      '部門別送件清單列印
      Case 14
         If CheckUse("frm1108", strExec) = True Then
            frm1108.Show
         End If
      
      'Add by Morgan 2011/6/2
      '部門別電子送件清單列印
      Case 15
         If CheckUse("frm1109", strExec) = True Then
            frm1109.Show
         End If
         
      'add by nickc 2005/05/03 銷案延遲日期輸入作業
      Case 16
         If CheckUse("frm140501", strExec) = True Then
            frm140501.Show
         End If
      'add by nickc 2005/07/22 CF 結餘單查詢
      Case 17
         If CheckUse("frm040202", strExec) = True Then
            frm040202.Show
         End If
      'add by nickc 2005/07/22 CF 結餘資料維護
      Case 18
         If CheckUse("frm040206", strExec) = True Then
            frm040206.Show
         End If
      'add by nickc 2008/03/27 CF 結餘單案件明細查詢
      Case 19
         If CheckUse("frm040208", strExec) = True Then
            frm040208.Show
         End If
   End Select
End Sub

Private Sub mnu1101_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '解除期限
         If CheckUse("frm110101_1", strExec) = True Then
            frm110101_1.Show
         End If
      Case 2 '取消收文
         If CheckUse("frm110102_1", strExec) = True Then
            frm110102_1.Show
         End If
      Case 3 '閉卷
         If CheckUse("frm110103_1", strExec) = True Then
            frm110103_1.Show
         End If
   End Select
End Sub

Private Sub mnu1102_Click(Index As Integer)
   ToolHide
   If CheckUse("frm010001", strExec) = True Then
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

Public Sub mnu110401_Click(Index As Integer)
'Add By Cheng 2003/04/02
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 1 '帳單輸入
         If CheckUse("Frmacc2150", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2150.Show
      Case 2 '抵帳單輸入
         If CheckUse("Frmacc2160", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2160.Show
      Case 3 '帳單作廢作業
         If CheckUse("Frmacc21j0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21j0.Show
      Case 4 '請款單輸入
            If CheckUse("Frmacc21h0", strExec) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            'Add By Cheng 2003/04/02
            '預設輸入請款資料不印地址條
            pub_blnARPrintAddress = False
            StrSQLa = "Select * From Staff Where ST01='" & strUserNum & "' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                '若為外商人員或電腦中心人員
                If "" & rsA("ST03").Value = "F10" Or "" & rsA("ST03").Value = "F11" Or "" & rsA("ST03").Value = "F12" Or "" & rsA("ST03").Value = "M51" Then
                    '輸入請款資料時印地址條
                    pub_blnARPrintAddress = True
                End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Frmacc21h0.Show
      Case 5 '折讓輸入
         If CheckUse("Frmacc21i0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21i0.Show
         ToolShow
         tool8_enabled
         Screen.MousePointer = vbDefault
         Exit Sub
      Case 6 '請款單作廢作業
         If CheckUse("Frmacc21k0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21k0.Show
      Case 7 '請款項目資料維護
         If CheckUse("Frmacc21g0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21g0.Show
      Case 8 '美金匯率資料維護
         If CheckUse("Frmacc21m0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21m0.Show
      Case 9 '預估結匯匯率資料維護
         If CheckUse("Frmacc21o0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21o0.Show
      Case 10 '相同案件性質整批請款作業
         If CheckUse("Frmacc21p0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21p0.Show
         ToolShow
         tool3_enabled
         Screen.MousePointer = vbDefault
         Exit Sub
      Case 11 '帳單審核作業
         If CheckUse("Frmacc2153", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2153.Show
         ToolShow
         tool3_enabled
         Screen.MousePointer = vbDefault
         Exit Sub
      'Add By Sindy 2009/06/06
      Case 12 '其他幣別請款匯率資料維護
         If CheckUse("Frmacc21s0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21s0.Show
      'Add by Morgan 2010/11/19
      'Memo by Morgan 2025/6/11 目前沒開放給User操作(選單已設不顯示)
      Case 13 '客製化請款項目資料維護
         If CheckUse("Frmacc21t0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21t0.Show
      'Add by Sindy 2021/1/15
      Case 14 '帳單輸入-整批
         If CheckUse("Frmacc21u0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Me.MousePointer = vbHourglass
         tool3_enabled
         Frmacc21u0.Show
         Me.MousePointer = vbDefault
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
      Case 1 '國外代理人帳目查詢
         If CheckUse("Frmacc2210", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2210.Show
      Case 2 '國外案件帳目查詢
         If CheckUse("Frmacc2220", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2220.Show
      Case 3 '國外請款金額查詢
         If CheckUse("Frmacc2230", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2230.Show
      Case 4 '案件損益查詢
         If CheckUse("Frmacc2240", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2240.Show
      'Add by Amy 20130914
     Case 5 '各幣別最新請款匯率查詢
        If CheckUse("Frmacc2142", strExec) = False Then
            Exit Sub
        End If
        Frmacc2142.Show
   End Select
   'Add by Amy 20130914 +各幣別最新請款匯率查詢不顯示toolbar
   If Index = 5 Then
      ToolHide
   Else
     ToolShow
      tool3_enabled
   End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu110403_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 1 'FC催款單
         If CheckUse("Frmacc2470", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2470.Show
      Case 2 'FC請款單
         If CheckUse("Frmacc2480", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2480.Show
      Case 3 '請款單整批列印
         If CheckUse("Frmacc24g0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24g0.Show
      'Add By Cheng 2002/09/04
      Case 4 '國外FC帳款明細表
         If CheckUse("Frmacc24i0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         '2007/11/29 ADD BY SONIA 外商使用者預設條件
         If Mid(GetStaffDepartment(strUserNum), 1, 2) = "F1" Then
            Frmacc24i0.Text7 = "F10"
            Frmacc24i0.Text8 = "F19"
         End If
         '2007/11/29 END
         Frmacc24i0.Show
         
      Case 5 '代理人帳目排名
         If CheckUse("Frmacc24b0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24b0.Show
      Case 6 'FC業務請款／收款明細表
         If CheckUse("Frmacc24c0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         '2007/11/6 ADD BY SONIA 外商使用者預設條件
         If Mid(GetStaffDepartment(strUserNum), 1, 2) = "F1" Then
            Frmacc24c0.Text9 = "F10"
            Frmacc24c0.Text10 = "F19"
         End If
         '2007/11/6 END
         Frmacc24c0.Show
      Case 7 '代理人逾期帳款分析表
         If CheckUse("Frmacc24f0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24f0.Show
      Case 8 '折讓單列印
         If CheckUse("Frmacc24h0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24h0.Show
      Case 9 'Add by Morgan 2010/11/29 國外請款點數分析表
         If CheckUse("Frmacc24k0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24k0.Show
        
      'Added by Morgan 2014/2/26
      Case 10 '請款單月報表
         If CheckUse("Frmacc24l0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24l0.Show
         
      'Added by Lydia 2018/11/30
      Case 11 '請款單折扣案件明細
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

Private Sub mnu1108_Click(Index As Integer)
    'Add By Cheng 2003/07/30
    ToolHide
    Select Case Index
       Case 1 '作業失誤資料維護
          If CheckUse("frm050714", strExec) = True Then
             frm050714.Show
          End If
       Case 2 '作業失誤明細表
          If CheckUse("frm040327", strExec) = True Then
             frm040327.Show
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

Private Sub mnu15_Click(Index As Integer)
ToolHide
End Sub

Public Sub ToolHide()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
End Sub

Private Sub mnu16_Click(Index As Integer)
   ToolHide
   Select Case Index
        'Modify By Cheng 2003/01/21
'      'Add By Cheng 2002/10/11
'      Case 1: PrinterLetterDemand
      Case 2 '定稿資料維護
         If CheckUse("frm1105", strExec) = True Then
            frm1105.Show
         End If
   End Select
End Sub

Private Sub mnu1601_Click(Index As Integer)
Dim strTemp As String 'Added by Lydia 2019/06/11

   'Dim iDefaultPrinter As Integer 'Remove by Morgan 2010/2/3
   ToolHide
   
   'Added by Lydia 2019/05/23
   If Index = 8 Then
       MsgBox "請記得選擇預設紙張為雙面的印表機！", vbInformation, "橫式雙面列印定稿"
   End If
   
ReSetPrinter: 'Added by Lydia 2019/06/11

   'Add by Morgan 2008/8/29
   '設定控制台&Word預設印表機
   Load frm880011
   'Modify by Morgan 2010/2/3
   'iDefaultPrinter = frm880011.GetPrinterIndex
   pub_OsPrinter = PUB_GetOsDefaultPrinter
   frm880011.Show 1
   'end 2008/8/29
   
   'Added by Lydia 2019/06/11 檢查是否為雙面列印的印表機
   If Index = 8 Then
      strTemp = PUB_GetOsDefaultPrinter
      If InStr(strTemp, "雙面") = 0 And InStr(UCase(strTemp), UCase("PDFCreator")) = 0 And InStr(UCase(strTemp), UCase("PDF reDirect")) = 0 Then
          'Modified by Lydia 2019/06/12 桂英反應選錯了要跳出
          'MsgBox "選擇印表機:" & strTemp & vbCrLf & "請記得選擇預設紙張為雙面的印表機！", vbInformation, "橫式雙面列印定稿"
          If MsgBox("選擇印表機:" & strTemp & vbCrLf & "請問是否要繼續列印橫式雙面列印定稿？" & vbCrLf & "選""是""會重新選擇印表機！", vbYesNo + vbDefaultButton1 + vbInformation, "橫式雙面列印定稿") = vbYes Then
              PUB_SetOsDefaultPrinter pub_OsPrinter
              GoTo ReSetPrinter
          'Added by Lydia 2019/06/12
          Else
              'Remove by Lydia 2020/04/10 debug-繼續列印
              'PUB_SetOsDefaultPrinter pub_OsPrinter
              'Exit Sub
              'end 2020/04/10
          End If
      End If
   'Added by Lydia 2019/06/13 使用者先自行在設定切換印表機,怕回到單面忘記換回來
   Else
      strTemp = PUB_GetOsDefaultPrinter
      If InStr(strTemp, "雙面") > 0 Then
          If MsgBox("選擇印表機:" & strTemp & vbCrLf & "請問是否要繼續列印單面定稿？" & vbCrLf & "選""是""會重新選擇印表機！", vbYesNo + vbDefaultButton2 + vbInformation, "檢查印表機") = vbYes Then
              PUB_SetOsDefaultPrinter pub_OsPrinter
              GoTo ReSetPrinter
          Else
              'Remove by Lydia 2020/04/10 debug-繼續列印
              'PUB_SetOsDefaultPrinter pub_OsPrinter
              'Exit Sub
              'end 2020/04/10
          End If
      End If
   'end 2019/06/13
   End If
      
   'Add By Cheng 2003/01/21
   Select Case Index
      Case 1 '橫式
         PrinterLetterDemand "1"
      Case 2 '英文
         PrinterLetterDemand "4"
      Case 3 '直式
         PrinterLetterDemand "2"
      Case 4 '日文
         PrinterLetterDemand "3"
      'add by nick 2004/12/16
      Case 5 '申請書
         PrinterLetterDemand "5"
      Case 6 '報價通知定稿
         'Modified by Morgan 2019/5/13 +m_UserNo
          PUB_Cache2Letter , , , , , , m_UserNo
      Case 7 '橫式(不印信頭)
         PrinterLetterDemand "7"
      'Added by Lydia 2019/05/23
      Case 8 '橫式雙面
         PrinterLetterDemand "8"
   End Select
   
   'Add by Morgan 2006/10/19
   '還原控制台&Word預設印表機
   'Modify by Morgan 2010/2/3
   'If iDefaultPrinter <> -1 Then
   '   If Printers(iDefaultPrinter).DeviceName <> Printer.DeviceName Then
   '      Printer.TrackDefault = True
   '      CreateObject("WScript.Network").SetDefaultPrinter Printers(iDefaultPrinter).DeviceName
   '      PUB_SetWordActivePrinter
   '   End If
   'End If
   PUB_SetOsDefaultPrinter pub_OsPrinter
   'end 2010/2/3
   'end 2006/10/19
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

Private Sub mnuTitle_Click(Index As Integer)
   ' 90.08.22 modify by louis
   Select Case Index
      'Modify By Cheng 2002/10/15
'      Case 16: PrinterLetterDemand
      Case 20 '商品查名
         If CheckUse("frm20", strExec) = True Then
            frm20.Show
         End If
      Case Else
   End Select
End Sub

Private Sub Timer1_Timer()
'Modify By Cheng 2002/11/21
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
'         mnuTitle(2).Enabled = False
'         mnuTitle(10).Enabled = False
'         mnuTitle(11).Enabled = False
'         mnuTitle(15).Enabled = False
'         mnuTitle(16).Enabled = False
'         'Add By Cheng 2002/08/15
'         mnuTitle(20).Enabled = False
'      End If
'   Else
'      If tmpFormI = 1 Then
'         If mnuTitle(0).Enabled = False Then
'            mnuTitle(0).Enabled = True
'            mnuTitle(2).Enabled = True
'            mnuTitle(10).Enabled = True
'            mnuTitle(11).Enabled = True
'            mnuTitle(15).Enabled = True
'            mnuTitle(16).Enabled = True
'            'Add By Cheng 2002/08/15
'            mnuTitle(20).Enabled = True
'         End If
'      End If
'   End If
'Add By Cheng 2002/11/22
Dim frm As Form
Dim intfrm10 As Integer
Dim intFrmacc2 As Integer
'Added by Morgan 2014/7/17
Dim bXForm As Boolean
Dim frmX As Form
'end 2014/7/17

'控制共同查詢
intfrm10 = 0
For Each frm In Forms
    'Modified by Morgan 2014/7/17 +frm100123 除外
    If Left(frm.Name, 5) = "frm10" And frm.Name <> "frm100123" Then
        intfrm10 = 1
        Exit For
    End If
Next

'Added by Morgan 2014/7/17
For Each frm In Forms
   If frm.Name = "frm100123" Then
      bXForm = True
      Set frmX = frm
      Exit For
   End If
Next
'end 2014/7/17

If intfrm10 = 1 Then
    If mnuTitle(10).Enabled = True Then mnuTitle(10).Enabled = False
    'add by nickc 2008/05/01
    'Modified by Morgan 2014/7/17
    'If mnu2102(2).Enabled = True Then mnu2102(2).Enabled = False
    If bXForm Then
      frmX.cmdOK(0).Enabled = False
      frmX.cmdOK(1).Enabled = False
    End If
    'end 2014/7/17
Else
    If mnuTitle(10).Enabled = False Then mnuTitle(10).Enabled = True
    'add by nickc 2008/05/01
    'Modified by Morgan 2014/7/17
    'If mnu2102(2).Enabled = False Then mnu2102(2).Enabled = True
    If bXForm Then
      frmX.cmdOK(0).Enabled = True
      frmX.cmdOK(1).Enabled = True
    End If
    'end 2014/7/17
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
    'Modify By Cheng 2002/12/13
'    If mnu11(6).Enabled = True Then mnu11(6).Enabled = False
'    If mnu11(7).Enabled = True Then mnu11(7).Enabled = False
    If mnu11(10).Enabled = True Then mnu11(10).Enabled = False
Else
    'Modify By Cheng 2002/12/13
'    If mnu11(6).Enabled = False Then mnu11(6).Enabled = True
'    If mnu11(7).Enabled = False Then mnu11(7).Enabled = True
    If mnu11(10).Enabled = False Then mnu11(10).Enabled = True
End If
'Add By Cheng 2003/12/19
'控制"視窗"Menu
MenuForFormControl
'End
StatusBar1.Panels.Item(4).Text = time
'Add by Morgan 2003/12/29
mnuTitle(99).Visible = mnuTitle(99).Enabled
End Sub

Private Sub Timer2_Timer()
    '若登入失敗
    If pub_str_LoginSucceeded <> "1" Then
        Me.Timer1.Interval = 0
    '若登入成功
    Else
        Me.Timer1.Interval = 100
        Me.Timer2.Interval = 0
        'MDIForm_Load
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
    'Modify By Sindy 2018/10/18 + Or Forms(Forms.Count - 1).Name <> Me.mnu99(Me.mnu99.Count - 1).Tag
    If Forms.Count - 1 <> Me.mnu99.Count Or Forms(Forms.Count - 1).Name <> Me.mnu99(Me.mnu99.Count - 1).Tag Then
        For Each frm In Forms
            'edit by nickc 2007/09/28
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
                'edit by nickc 2007/09/28
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
                    'edit by nickc 2007/09/28
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
            'edit by nickc 2007/09/28
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
   If Not mdiMain.ActiveForm Is Nothing Then 'Added by Morgan 2015/10/30
    For Each objMnu99 In Me.mnu99
        If mdiMain.ActiveForm.Name = Me.mnu99(objMnu99.Index).Tag Then
            If Me.mnu99(objMnu99.Index).Checked = False Then Me.mnu99(objMnu99.Index).Checked = True
        Else
            If Me.mnu99(objMnu99.Index).Checked = True Then Me.mnu99(objMnu99.Index).Checked = False
        End If
    Next
   End If 'Added by Morgan 2015/10/30
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
            Case "CFT", "T", "TF", "FCT"
               If strCP01 = "T" Or strCP01 = "TF" Then
                  Call CallFormData(frm020501, "frm020501", strCP01, strCP02, strCP03, strCP04)
               End If
            Case "TB"
               Call CallFormData(frm02050201, "frm02050201", strCP01, strCP02, strCP03, strCP04)
            Case "TM"
               Call CallFormData(frm02050202, "frm02050202", strCP01, strCP02, strCP03, strCP04)
            Case "TD"
               Call CallFormData(frm02050203, "frm02050203", strCP01, strCP02, strCP03, strCP04)
            Case "TC"
               Call CallFormData(frm02050204, "frm02050204", strCP01, strCP02, strCP03, strCP04)
            Case Else
               Call CallFormData(frm02050205, "frm02050205", strCP01, strCP02, strCP03, strCP04)
         End Select
   End Select
End Function

'Add By Sindy 2015/10/20
Public Sub SetTmpfrm1103_2()
   Set Tmpfrm1103_2 = frm1103_2
End Sub
'2015/10/20 END

'Added by Lydia 2015/12/08 商標公報資料統計-Excel
Private Sub mnu020622_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '申請人國籍及洲別統計(含同業)
         If CheckUse("frm030621", strExec) Then
            frm030621.Show
         End If
      Case 2 '各單位公報類別數統計
         If CheckUse("frm030623", strExec) Then
            frm030623.Show
         End If
      Case 3 '三部門案件來源比較
         If CheckUse("frm030622", strExec) Then
            frm030622.Show
         End If
      Case 4 '同業案件來源比較
         If CheckUse("frm030624", strExec) Then
            frm030624.Show
         End If
      Case 5 '同業台灣各區類別數比較
         If CheckUse("frm030625", strExec) Then
            frm030625.Show
         End If
   End Select
End Sub


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

'Added by Morgan 2020/1/16
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
   'Add by Amy 2024/01/22 接洽單-對造
   Case "frm090801_14"
         Set GetForm = frm090801_14
   'Add By Sindy 2020/5/29
   Case "frm180301"
         Set GetForm = frm180301
   '2020/5/29 END
   'Add by Amy 2023/09/22 共同查詢用
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
   'Add by Amy 2025/08/28
   Case "frm210133" '結案單
      Set GetForm = frm210133
   Case "frm210133_F" '國外部結案單
      Set GetForm = frm210133_F
   Case "Frmacc21h0" '請款單
      Set GetForm = Frmacc21h0
   Case "frm210133_INV"
      Set GetForm = frm210133_INV
   'Add by Amy 2025/10/28
   Case "mdiMain"
      Set GetForm = mdiMain
   End Select
End Function

'Added by Morgan 2021/4/22
'複製貼上彈跳視窗
Public Sub PopupMenu2(oTextBox As Control)
   Set oControl = oTextBox 'Added by Morgan 2022/5/24
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
   'Added by Morgan 2022/5/25
   Case 4 '輸入
      frm880023.SetTextBox oControl
      frm880023.Show vbModal
   End Select
End Sub
