VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000018&
   Caption         =   "承辦人作業"
   ClientHeight    =   5390
   ClientLeft      =   60
   ClientTop       =   960
   ClientWidth     =   10080
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '最大化
   Begin VB.Timer tmrSalary 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1395
      Top             =   3600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3645
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4815
      Top             =   2490
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1125
      Top             =   2175
   End
   Begin VB.Timer tmrConnect 
      Left            =   1125
      Top             =   1725
   End
   Begin VB.Timer Timer2 
      Left            =   450
      Top             =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   450
      Top             =   1725
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   550
      Top             =   610
      _ExtentX        =   953
      _ExtentY        =   953
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1130
      Top             =   610
      _ExtentX        =   953
      _ExtentY        =   953
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   917
      ButtonWidth     =   494
      ButtonHeight    =   811
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   280
      Left            =   0
      TabIndex        =   0
      Top             =   5110
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   494
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   15
      Top             =   615
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
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
      Caption         =   "國外部"
      Index           =   3
      Begin VB.Menu mnu03 
         Caption         =   "撰寫信函作業"
         Index           =   1
      End
      Begin VB.Menu mnu03 
         Caption         =   "工作進度資料維護"
         Index           =   2
      End
      Begin VB.Menu mnu03 
         Caption         =   "待核判區"
         Index           =   3
      End
      Begin VB.Menu mnu03 
         Caption         =   "工作進度資料查詢"
         Index           =   4
      End
      Begin VB.Menu mnu03 
         Caption         =   "承辦人工作進度資料查詢"
         Index           =   5
      End
      Begin VB.Menu mnu03 
         Caption         =   "承辦人請款/發文明細表"
         Index           =   6
      End
      Begin VB.Menu mnu03 
         Caption         =   "國外部專利處期限通知"
         Index           =   7
      End
      Begin VB.Menu mnu03 
         Caption         =   "法務處期限通知"
         Index           =   8
      End
      Begin VB.Menu mnu03 
         Caption         =   "行事曆提醒通知"
         Index           =   9
      End
      Begin VB.Menu mnu03 
         Caption         =   "國外FC帳款明細表"
         Index           =   10
      End
      Begin VB.Menu mnu03 
         Caption         =   "各幣別最新請款匯率查詢"
         Index           =   11
      End
      Begin VB.Menu mnu03 
         Caption         =   "電話聯絡單發文"
         Index           =   12
      End
      Begin VB.Menu mnu03 
         Caption         =   "專利日文資料維護作業"
         Index           =   13
      End
      Begin VB.Menu mnu03 
         Caption         =   "客戶日文資料維護作業"
         Index           =   14
      End
      Begin VB.Menu mnu03 
         Caption         =   "代理人日文資料維護作業"
         Index           =   15
      End
      Begin VB.Menu mnu03 
         Caption         =   "工程師各式申請書"
         Index           =   16
      End
      Begin VB.Menu mnu03 
         Caption         =   "外專新案未命名區"
         Index           =   17
         Begin VB.Menu mnu0316 
            Caption         =   "待分案/待確認"
            Index           =   1
         End
         Begin VB.Menu mnu0316 
            Caption         =   "待命名"
            Index           =   2
         End
      End
      Begin VB.Menu mnu03 
         Caption         =   "工程師上傳作業"
         Index           =   18
      End
      Begin VB.Menu mnu03 
         Caption         =   "外專翻譯分案-認翻譯"
         Index           =   19
      End
      Begin VB.Menu mnu03 
         Caption         =   "外專-專利連結通知維護作業"
         Index           =   20
      End
      Begin VB.Menu mnu03 
         Caption         =   "外專新案認領區"
         Index           =   21
      End
      Begin VB.Menu mnu03 
         Caption         =   "外專-藥證號維護作業"
         Index           =   22
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "專利處"
      Index           =   5
      Begin VB.Menu mnu05 
         Caption         =   "承辦人作業"
         Index           =   2
         Begin VB.Menu mnu0502 
            Caption         =   "工作進度資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu0502 
            Caption         =   "待核判區"
            Index           =   2
         End
         Begin VB.Menu mnu0502 
            Caption         =   "承辦人支援記錄維護"
            Index           =   3
         End
         Begin VB.Menu mnu0502 
            Caption         =   "承辦人外出記錄維護"
            Index           =   4
         End
         Begin VB.Menu mnu0502 
            Caption         =   "承辦人特殊案件記錄維護"
            Index           =   5
         End
         Begin VB.Menu mnu0502 
            Caption         =   "承辦人修改記錄維護"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnu0502 
            Caption         =   "承辦人衍生記錄維護"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnu0502 
            Caption         =   "未齊備、未完稿、未發文查詢"
            Index           =   8
         End
         Begin VB.Menu mnu0502 
            Caption         =   "工作進度資料查詢"
            Index           =   9
         End
         Begin VB.Menu mnu0502 
            Caption         =   "承辦人達成情形查詢"
            Index           =   10
         End
         Begin VB.Menu mnu0502 
            Caption         =   "每週速度查詢"
            Index           =   11
         End
         Begin VB.Menu mnu0502 
            Caption         =   "工作進度資料列印"
            Index           =   13
         End
         Begin VB.Menu mnu0502 
            Caption         =   "專利案例管理"
            Index           =   14
            Begin VB.Menu mnu050205 
               Caption         =   "專利案例個人輸入作業"
               Index           =   1
            End
            Begin VB.Menu mnu050205 
               Caption         =   "專利案例資料查詢"
               Index           =   2
            End
            Begin VB.Menu mnu050205 
               Caption         =   "專利案例資料彙整作業"
               Index           =   3
            End
            Begin VB.Menu mnu050205 
               Caption         =   "專利案例資料維護"
               Index           =   4
            End
         End
         Begin VB.Menu mnu0502 
            Caption         =   "公報簡訊管理"
            Index           =   15
            Begin VB.Menu mnu050207 
               Caption         =   "公報簡訊個人輸入作業"
               Index           =   1
            End
            Begin VB.Menu mnu050207 
               Caption         =   "公報簡訊資料查詢/列印"
               Index           =   2
            End
            Begin VB.Menu mnu050207 
               Caption         =   "公報簡訊資料彙整作業"
               Index           =   3
            End
            Begin VB.Menu mnu050207 
               Caption         =   "公報簡訊資料維護"
               Index           =   4
            End
            Begin VB.Menu mnu050207 
               Caption         =   "公報簡訊索引資料維護"
               Index           =   5
            End
         End
         Begin VB.Menu mnu0502 
            Caption         =   "期刊資料管理"
            Index           =   16
            Begin VB.Menu mnu050208 
               Caption         =   "期刊資料查詢/列印"
               Index           =   1
            End
            Begin VB.Menu mnu050208 
               Caption         =   "期刊資料維護"
               Index           =   2
            End
            Begin VB.Menu mnu050208 
               Caption         =   "期刊索引資料維護"
               Index           =   3
            End
         End
         Begin VB.Menu mnu0502 
            Caption         =   "月考核"
            Index           =   17
         End
         Begin VB.Menu mnu0502 
            Caption         =   "季考核"
            Index           =   18
         End
         Begin VB.Menu mnu0502 
            Caption         =   "英文核稿查詢"
            Index           =   19
         End
      End
      Begin VB.Menu mnu05 
         Caption         =   "繪圖人員作業"
         Index           =   3
         Begin VB.Menu mnu0503 
            Caption         =   "工作進度資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu0503 
            Caption         =   "待核判區"
            Index           =   2
         End
         Begin VB.Menu mnu0503 
            Caption         =   "繪圖人員支援記錄維護"
            Index           =   3
         End
         Begin VB.Menu mnu0503 
            Caption         =   "繪圖人員外出記錄維護"
            Index           =   4
         End
         Begin VB.Menu mnu0503 
            Caption         =   "未齊備、未完稿查詢"
            Index           =   5
         End
         Begin VB.Menu mnu0503 
            Caption         =   "工作進度資料查詢"
            Index           =   6
         End
         Begin VB.Menu mnu0503 
            Caption         =   "工作進度資料列印"
            Index           =   8
         End
         Begin VB.Menu mnu0503 
            Caption         =   "繪圖人員達成情形查詢"
            Index           =   9
         End
         Begin VB.Menu mnu0503 
            Caption         =   "專利案例資料查詢"
            Index           =   10
         End
         Begin VB.Menu mnu0503 
            Caption         =   "公報簡訊管理"
            Index           =   11
            Begin VB.Menu mnu050310 
               Caption         =   "公報簡訊個人輸入作業"
               Index           =   1
            End
            Begin VB.Menu mnu050310 
               Caption         =   "公報簡訊資料查詢/列印"
               Index           =   2
            End
         End
         Begin VB.Menu mnu0503 
            Caption         =   "期刊資料查詢/列印"
            Index           =   12
         End
         Begin VB.Menu mnu0503 
            Caption         =   "每週速度查詢"
            Index           =   13
         End
         Begin VB.Menu mnu0503 
            Caption         =   "月考核"
            Index           =   14
         End
         Begin VB.Menu mnu0503 
            Caption         =   "季考核"
            Index           =   15
         End
      End
      Begin VB.Menu mnu05 
         Caption         =   "撰寫信函作業"
         Index           =   4
      End
      Begin VB.Menu mnu05 
         Caption         =   "P案國外新案指示信"
         Index           =   6
      End
      Begin VB.Menu mnu05 
         Caption         =   "P案各式申請書"
         Index           =   7
      End
      Begin VB.Menu mnu05 
         Caption         =   "聯絡單列印及E-Mail"
         Index           =   8
      End
      Begin VB.Menu mnu05 
         Caption         =   "承辦人工作管理"
         Index           =   9
         Begin VB.Menu mnu0507 
            Caption         =   "查詢及報表"
            Index           =   3
            Begin VB.Menu mnu050703 
               Caption         =   "承辦人工作進度資料查詢"
               Index           =   1
            End
            Begin VB.Menu mnu050703 
               Caption         =   "承辦人達成情形查詢"
               Index           =   2
            End
            Begin VB.Menu mnu050703 
               Caption         =   "承辦人工作量查詢"
               Index           =   3
            End
            Begin VB.Menu mnu050703 
               Caption         =   "承辦人每日分案情形查詢"
               Index           =   4
            End
            Begin VB.Menu mnu050703 
               Caption         =   "承辦天數統計查詢"
               Index           =   5
            End
            Begin VB.Menu mnu050703 
               Caption         =   "未齊備未完稿未發文查詢"
               Index           =   6
            End
            Begin VB.Menu mnu050703 
               Caption         =   "案件處理時間統計查詢"
               Index           =   7
            End
            Begin VB.Menu mnu050703 
               Caption         =   "工程師每週完稿明細"
               Index           =   8
            End
            Begin VB.Menu mnu050703 
               Caption         =   "案件逾期及異常查詢"
               Index           =   9
            End
            Begin VB.Menu mnu050703 
               Caption         =   "加乘註記修改歷史查詢列印"
               Index           =   10
            End
            Begin VB.Menu mnu050703 
               Caption         =   "英文核稿查詢"
               Index           =   11
            End
            Begin VB.Menu mnu050703 
               Caption         =   "智權人員收文高低標查詢"
               Index           =   12
            End
            Begin VB.Menu mnu050703 
               Caption         =   "預定會稿日異常案件查詢"
               Index           =   13
            End
            Begin VB.Menu mnu050703 
               Caption         =   "支援記錄獎金統計"
               Index           =   14
            End
            Begin VB.Menu mnu050703 
               Caption         =   "待辦案件量統計查詢"
               Index           =   15
            End
            Begin VB.Menu mnu050703 
               Caption         =   "每周承辦會議統計查詢"
               Index           =   16
            End
            Begin VB.Menu mnu050703 
               Caption         =   "支援次數統計"
               Index           =   17
            End
         End
         Begin VB.Menu mnu0507 
            Caption         =   "人員考核管理"
            Index           =   4
            Begin VB.Menu mnu050704 
               Caption         =   "專利處每週速度考核"
               Index           =   1
            End
            Begin VB.Menu mnu050704 
               Caption         =   "月考核"
               Index           =   2
            End
            Begin VB.Menu mnu050704 
               Caption         =   "季考核"
               Index           =   3
            End
            Begin VB.Menu mnu050704 
               Caption         =   "工程師每月目標基數設定"
               Index           =   4
            End
            Begin VB.Menu mnu050704 
               Caption         =   "個人目標資料維護"
               Index           =   5
            End
            Begin VB.Menu mnu050704 
               Caption         =   "獎金輸入作業"
               Index           =   6
            End
            Begin VB.Menu mnu050704 
               Caption         =   "獎金明細表"
               Index           =   7
            End
         End
         Begin VB.Menu mnu0507 
            Caption         =   "基本資料維護"
            Index           =   5
            Begin VB.Menu mnu050705 
               Caption         =   "承辦人支援記錄維護"
               Index           =   1
            End
            Begin VB.Menu mnu050705 
               Caption         =   "承辦人特殊案件記錄維護"
               Index           =   2
            End
            Begin VB.Menu mnu050705 
               Caption         =   "承辦人修改記錄維護"
               Index           =   3
            End
            Begin VB.Menu mnu050705 
               Caption         =   "承辦人衍生記錄維護"
               Index           =   4
            End
            Begin VB.Menu mnu050705 
               Caption         =   "國內外案件資料維護"
               Index           =   5
            End
            Begin VB.Menu mnu050705 
               Caption         =   "每月目次重編作業"
               Index           =   6
            End
            Begin VB.Menu mnu050705 
               Caption         =   "特殊加乘註記維護"
               Index           =   8
            End
            Begin VB.Menu mnu050705 
               Caption         =   "英文核稿人欄修改權限設定"
               Index           =   9
            End
            Begin VB.Menu mnu050705 
               Caption         =   "免費修正事由維護"
               Index           =   10
            End
         End
      End
      Begin VB.Menu mnu05 
         Caption         =   "繪圖人員工作管理"
         Index           =   10
         Begin VB.Menu mnu0508 
            Caption         =   "查詢及報表"
            Index           =   1
            Begin VB.Menu mnu050801 
               Caption         =   "繪圖人員工作進度資料查詢"
               Index           =   1
            End
            Begin VB.Menu mnu050801 
               Caption         =   "繪圖超時案件查詢"
               Index           =   2
            End
            Begin VB.Menu mnu050801 
               Caption         =   "繪圖人員達成情形查詢"
               Index           =   3
            End
            Begin VB.Menu mnu050801 
               Caption         =   "繪圖人員工作量查詢"
               Index           =   4
            End
            Begin VB.Menu mnu050801 
               Caption         =   "繪圖人員每日分案情形查詢"
               Index           =   5
            End
            Begin VB.Menu mnu050801 
               Caption         =   "繪圖人員作業天數統計查詢"
               Index           =   6
            End
            Begin VB.Menu mnu050801 
               Caption         =   "未齊備、未完稿、未發文查詢"
               Index           =   7
            End
         End
         Begin VB.Menu mnu0508 
            Caption         =   "人員考核管理"
            Index           =   2
            Begin VB.Menu mnu050802 
               Caption         =   "個人目標資料維護"
               Index           =   1
            End
            Begin VB.Menu mnu050802 
               Caption         =   "月考核"
               Index           =   2
            End
            Begin VB.Menu mnu050802 
               Caption         =   "季考核"
               Index           =   3
            End
         End
         Begin VB.Menu mnu0508 
            Caption         =   "繪圖人員支援記錄維護"
            Index           =   3
         End
         Begin VB.Menu mnu0508 
            Caption         =   "繪圖分案作業"
            Index           =   4
         End
      End
      Begin VB.Menu mnu05 
         Caption         =   "主管機關處理記錄"
         Index           =   11
         Begin VB.Menu mnu0511 
            Caption         =   "來電記錄"
            Index           =   1
         End
         Begin VB.Menu mnu0511 
            Caption         =   "去電記錄"
            Index           =   2
         End
      End
      Begin VB.Menu mnu05 
         Caption         =   "公文來函判發作業"
         Index           =   12
      End
      Begin VB.Menu mnu05 
         Caption         =   "發後補看作業"
         Index           =   13
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "商標處"
      Index           =   7
      Begin VB.Menu mnu07 
         Caption         =   "商標委查作業"
         Index           =   0
         Begin VB.Menu mnu0701 
            Caption         =   "查名/待查區(網中)"
            Index           =   1
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名/查覆區(網中)"
            Index           =   2
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名/覆核區(網中)"
            Index           =   3
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名單維護(網中)"
            Index           =   4
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名期限資料查詢"
            Index           =   5
         End
         Begin VB.Menu mnu0701 
            Caption         =   "委查組群統計"
            Index           =   6
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名人查覆明細表"
            Index           =   7
         End
         Begin VB.Menu mnu0701 
            Caption         =   "期限過期明細表"
            Index           =   8
         End
         Begin VB.Menu mnu0701 
            Caption         =   "委查人委查明細表"
            Index           =   9
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名人查覆統計表"
            Index           =   10
         End
         Begin VB.Menu mnu0701 
            Caption         =   "委查人委查統計表"
            Index           =   11
         End
         Begin VB.Menu mnu0701 
            Caption         =   "商品組群委查統計表"
            Index           =   12
         End
         Begin VB.Menu mnu0701 
            Caption         =   "委查資料刪除作業"
            Index           =   13
         End
         Begin VB.Menu mnu0701 
            Caption         =   "刪除組群維護"
            Index           =   14
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名人員維護"
            Index           =   15
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名人狀態"
            Index           =   16
         End
         Begin VB.Menu mnu0701 
            Caption         =   "組群和圖形路徑維護"
            Index           =   17
            Begin VB.Menu mnu070117 
               Caption         =   "查名組群及本數維護"
               Index           =   1
            End
            Begin VB.Menu mnu070117 
               Caption         =   "圖形查名路徑-大分類維護"
               Index           =   2
            End
            Begin VB.Menu mnu070117 
               Caption         =   "圖形查名路徑-中分類維護"
               Index           =   3
            End
            Begin VB.Menu mnu070117 
               Caption         =   "圖形查名路徑-小分類維護"
               Index           =   4
            End
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名/待查區"
            Index           =   18
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名/覆核區"
            Index           =   19
         End
         Begin VB.Menu mnu0701 
            Caption         =   "查名單維護(限電腦中心)"
            Index           =   20
         End
      End
      Begin VB.Menu mnu07 
         Caption         =   "商品名稱維護"
         Index           =   1
      End
      Begin VB.Menu mnu07 
         Caption         =   "承辦人作業"
         Index           =   2
         Begin VB.Menu mnu0702 
            Caption         =   "工作進度資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu0702 
            Caption         =   "待核判區"
            Index           =   2
         End
         Begin VB.Menu mnu0702 
            Caption         =   "未齊備、未完稿、未發文查詢"
            Index           =   8
         End
         Begin VB.Menu mnu0702 
            Caption         =   "工作進度資料查詢"
            Index           =   9
         End
         Begin VB.Menu mnu0702 
            Caption         =   "承辦人達成情形查詢"
            Index           =   10
         End
         Begin VB.Menu mnu0702 
            Caption         =   "工作進度資料列印"
            Index           =   13
         End
         Begin VB.Menu mnu0702 
            Caption         =   "電話回覆主管機關"
            Index           =   18
         End
         Begin VB.Menu mnu0702 
            Caption         =   "主管機關來電處理記錄"
            Index           =   19
         End
         Begin VB.Menu mnu0702 
            Caption         =   "商標委任書正本案號維護"
            Index           =   22
         End
         Begin VB.Menu mnu0702 
            Caption         =   "台灣商標延展開拓"
            Index           =   23
         End
         Begin VB.Menu mnu0702 
            Caption         =   "商申承辦人內部及機關收發文統計表"
            Index           =   24
         End
         Begin VB.Menu mnu0702 
            Caption         =   "台灣商標委任狀中譯文"
            Index           =   25
         End
         Begin VB.Menu mnu0702 
            Caption         =   "未發文案件原因註記"
            Index           =   26
         End
         Begin VB.Menu mnu0702 
            Caption         =   "案件催審延緩維護"
            Index           =   27
         End
         Begin VB.Menu mnu0702 
            Caption         =   "案件催審作業"
            Index           =   28
         End
         Begin VB.Menu mnu0702 
            Caption         =   "台灣案催審申請書(定稿資料維護)"
            Index           =   29
         End
         Begin VB.Menu mnu0702 
            Caption         =   "國外代理人帳目查詢"
            Index           =   30
         End
         Begin VB.Menu mnu0702 
            Caption         =   "國外案件帳目查詢"
            Index           =   31
         End
         Begin VB.Menu mnu0702 
            Caption         =   "各幣別最新請款匯率查詢"
            Index           =   32
         End
      End
      Begin VB.Menu mnu07 
         Caption         =   "撰寫信函作業"
         Index           =   4
      End
      Begin VB.Menu mnu07 
         Caption         =   "T案大陸指示信"
         Index           =   5
      End
      Begin VB.Menu mnu07 
         Caption         =   "TC陸代申請書輸入"
         Index           =   6
      End
      Begin VB.Menu mnu07 
         Caption         =   "聯絡單列印及E-Mail"
         Index           =   8
      End
      Begin VB.Menu mnu07 
         Caption         =   "承辦人工作管理"
         Index           =   9
         Begin VB.Menu mnu0707 
            Caption         =   "查詢及報表"
            Index           =   3
            Begin VB.Menu mnu070703 
               Caption         =   "承辦人工作進度資料查詢"
               Index           =   1
            End
            Begin VB.Menu mnu070703 
               Caption         =   "承辦人達成情形查詢"
               Index           =   2
            End
            Begin VB.Menu mnu070703 
               Caption         =   "承辦人工作量查詢"
               Index           =   3
            End
            Begin VB.Menu mnu070703 
               Caption         =   "承辦人每日分案情形查詢"
               Index           =   4
            End
            Begin VB.Menu mnu070703 
               Caption         =   "承辦天數統計查詢"
               Index           =   5
            End
            Begin VB.Menu mnu070703 
               Caption         =   "未齊備未完稿未發文查詢"
               Index           =   6
            End
            Begin VB.Menu mnu070703 
               Caption         =   "MCT收發文件數及點數統計"
               Index           =   7
            End
         End
         Begin VB.Menu mnu0708 
            Caption         =   "基本資料維護"
            Index           =   4
            Begin VB.Menu mnu070801 
               Caption         =   "商申承辦人責任業務區分配維護"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnu07 
         Caption         =   "商標處信件"
         Index           =   10
         Begin VB.Menu mnu0710 
            Caption         =   "商標處收件夾信件處理"
            Index           =   0
         End
         Begin VB.Menu mnu0710 
            Caption         =   "郵件分信關鍵字對照表維護"
            Index           =   1
         End
         Begin VB.Menu mnu0710 
            Caption         =   "未處理信件查詢"
            Index           =   2
         End
      End
      Begin VB.Menu mnu07 
         Caption         =   "商標處期限通知"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnu07 
         Caption         =   "公文來函判發作業"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnu07 
         Caption         =   "發後補看作業"
         Index           =   13
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "共同查詢(&Q)"
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
            Caption         =   "外專未完成核稿明細查詢/列印"
            Index           =   1
         End
         Begin VB.Menu mnu103 
            Caption         =   "外專收文未發文明細查詢/列印"
            Index           =   2
         End
         Begin VB.Menu mnu103 
            Caption         =   "客戶重新委任案件查詢/列印"
            Index           =   3
         End
         Begin VB.Menu mnu103 
            Caption         =   "外專承辦人請款/發文明細表"
            Index           =   4
         End
         Begin VB.Menu mnu103 
            Caption         =   "外專工程師請款點數和OA發文統計表"
            Index           =   5
         End
         Begin VB.Menu mnu103 
            Caption         =   "外商每月請款點數統計表"
            Index           =   6
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
      Caption         =   "品名查詢"
      Index           =   20
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "智權部"
      Index           =   21
      Begin VB.Menu mnu20 
         Caption         =   "個人常用區"
         Index           =   1
         Begin VB.Menu mnu2001 
            Caption         =   "常用1"
            Index           =   1
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用2"
            Index           =   2
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用3"
            Index           =   3
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用4"
            Index           =   4
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用5"
            Index           =   5
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用6"
            Index           =   6
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用7"
            Index           =   7
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用8"
            Index           =   8
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用9"
            Index           =   9
         End
         Begin VB.Menu mnu2001 
            Caption         =   "常用10"
            Index           =   10
         End
      End
      Begin VB.Menu mnu21 
         Caption         =   "程序作業"
         Index           =   1
         Begin VB.Menu mnu2101 
            Caption         =   "委任契約書"
            Index           =   1
         End
         Begin VB.Menu mnu2101 
            Caption         =   "案件接洽單"
            Index           =   2
         End
         Begin VB.Menu mnu2101 
            Caption         =   "案件接洽單(電子收文)"
            Index           =   3
            Shortcut        =   ^O
         End
         Begin VB.Menu mnu2101 
            Caption         =   "案件結案單"
            Index           =   4
            Shortcut        =   ^E
         End
         Begin VB.Menu mnu2101 
            Caption         =   "銷案／銷帳單"
            Index           =   5
         End
         Begin VB.Menu mnu2101 
            Caption         =   "期限資料查詢"
            Index           =   6
         End
         Begin VB.Menu mnu2101 
            Caption         =   "回覆單"
            Index           =   7
         End
         Begin VB.Menu mnu2101 
            Caption         =   "未列印收據查詢"
            Index           =   8
         End
         Begin VB.Menu mnu2101 
            Caption         =   "電子收文接洽單查詢"
            Index           =   9
         End
         Begin VB.Menu mnu2101 
            Caption         =   "寄發文件"
            Index           =   10
         End
         Begin VB.Menu mnu2101 
            Caption         =   "寄件查詢"
            Index           =   11
         End
      End
      Begin VB.Menu mnu21 
         Caption         =   "專利商標作業"
         Index           =   2
         Begin VB.Menu mnu2102 
            Caption         =   "專利／商標會稿"
            Index           =   1
            Shortcut        =   ^D
         End
         Begin VB.Menu mnu2102 
            Caption         =   "專利案件彙整表"
            Index           =   2
         End
         Begin VB.Menu mnu2102 
            Caption         =   "商標查名／查覆區"
            Index           =   3
            Shortcut        =   ^T
         End
         Begin VB.Menu mnu2102 
            Caption         =   "商標查名報告"
            Index           =   4
         End
         Begin VB.Menu mnu2102 
            Caption         =   "商標著作權案件齊備管制"
            Index           =   5
         End
         Begin VB.Menu mnu2102 
            Caption         =   "商標未發文原因註記"
            Index           =   6
         End
         Begin VB.Menu mnu2102 
            Caption         =   "合併查名／查覆區"
            Index           =   7
         End
      End
      Begin VB.Menu mnu21 
         Caption         =   "財務作業"
         Index           =   3
         Begin VB.Menu mnu2103 
            Caption         =   "請款作業及應收查詢"
            Index           =   1
            Shortcut        =   ^R
         End
         Begin VB.Menu mnu2103 
            Caption         =   "繳款作業及收據PDF"
            Index           =   2
            Shortcut        =   ^M
         End
         Begin VB.Menu mnu2103 
            Caption         =   "繳款查詢及簽收查詢"
            Index           =   3
         End
         Begin VB.Menu mnu2103 
            Caption         =   "點數輸入作業及查詢"
            Index           =   4
         End
         Begin VB.Menu mnu2103 
            Caption         =   "每月點數結算及查詢"
            Index           =   5
         End
         Begin VB.Menu mnu2103 
            Caption         =   "暫收款查詢"
            Index           =   6
         End
         Begin VB.Menu mnu2103 
            Caption         =   "未列印收據查詢"
            Index           =   7
         End
      End
      Begin VB.Menu mnu21 
         Caption         =   "查詢資料"
         Index           =   4
         Begin VB.Menu mnu2104 
            Caption         =   "未發文案件管制"
            Index           =   1
         End
         Begin VB.Menu mnu2104 
            Caption         =   "來函期限查詢"
            Index           =   2
         End
         Begin VB.Menu mnu2104 
            Caption         =   "定稿報價查詢"
            Index           =   3
         End
         Begin VB.Menu mnu2104 
            Caption         =   "新客戶來源分析"
            Index           =   4
         End
         Begin VB.Menu mnu2104 
            Caption         =   "新舊客戶收款貢獻度分析"
            Index           =   5
         End
         Begin VB.Menu mnu2104 
            Caption         =   "智權人員收/發文量分析"
            Index           =   6
         End
         Begin VB.Menu mnu2104 
            Caption         =   "價目表"
            Index           =   7
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnu2104 
            Caption         =   "各國年費預估報價"
            Index           =   8
         End
         Begin VB.Menu mnu2104 
            Caption         =   "CFP領證預估報價"
            Index           =   9
         End
         Begin VB.Menu mnu2104 
            Caption         =   "下一程序接洽單列印"
            Index           =   10
         End
      End
      Begin VB.Menu mnu21 
         Caption         =   "其他"
         Index           =   5
         Begin VB.Menu mnu2105 
            Caption         =   "客戶資料修改"
            Index           =   1
         End
         Begin VB.Menu mnu2105 
            Caption         =   "案件資料修改"
            Index           =   2
         End
         Begin VB.Menu mnu2105 
            Caption         =   "行事曆"
            Index           =   3
         End
         Begin VB.Menu mnu2105 
            Caption         =   "撰寫信函作業"
            Index           =   4
         End
         Begin VB.Menu mnu2105 
            Caption         =   "聯絡單"
            Index           =   5
         End
      End
      Begin VB.Menu mnu21 
         Caption         =   "區主管作業"
         Index           =   6
         Begin VB.Menu mnu2106 
            Caption         =   "每日點數輸入"
            Index           =   1
         End
         Begin VB.Menu mnu2106 
            Caption         =   "各區業績點數統計"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnu2106 
            Caption         =   "專業達成點數表-秘書"
            Index           =   3
         End
         Begin VB.Menu mnu2106 
            Caption         =   "各區業務工作報告統計"
            Index           =   4
         End
         Begin VB.Menu mnu2106 
            Caption         =   "各所業務工作報告統計"
            Index           =   5
         End
         Begin VB.Menu mnu2106 
            Caption         =   "智權部工作報告-總所"
            Index           =   6
         End
         Begin VB.Menu mnu2106 
            Caption         =   "新申請案收文至發文件數日數比較表"
            Index           =   7
         End
         Begin VB.Menu mnu2106 
            Caption         =   "客戶案件整理表記錄查詢"
            Index           =   8
         End
         Begin VB.Menu mnu2106 
            Caption         =   "業績達成日報表"
            Index           =   9
         End
         Begin VB.Menu mnu2106 
            Caption         =   "業績達成月報表"
            Index           =   10
         End
         Begin VB.Menu mnu2106 
            Caption         =   "業務收/發文量比較查詢"
            Index           =   11
         End
         Begin VB.Menu mnu2106 
            Caption         =   "智權部點數分析表"
            Index           =   12
         End
         Begin VB.Menu mnu2106 
            Caption         =   "未收款、未收齊清單列印"
            Index           =   13
         End
         Begin VB.Menu mnu2106 
            Caption         =   "業績年度統計表"
            Index           =   14
         End
         Begin VB.Menu mnu2106 
            Caption         =   "客戶特殊紀錄異動"
            Index           =   15
         End
      End
      Begin VB.Menu mnu21 
         Caption         =   "國內業務開拓"
         Index           =   7
         Begin VB.Menu mnu2107 
            Caption         =   "潛在客戶資料維護"
            Index           =   1
         End
         Begin VB.Menu mnu2107 
            Caption         =   "往來記錄資料維護"
            Index           =   2
         End
         Begin VB.Menu mnu2107 
            Caption         =   "潛在客戶資料查詢"
            Index           =   3
         End
         Begin VB.Menu mnu2107 
            Caption         =   "往來記錄資料查詢"
            Index           =   4
         End
         Begin VB.Menu mnu2107 
            Caption         =   "台灣商標公告近三年開拓函"
            Index           =   5
         End
         Begin VB.Menu mnu2107 
            Caption         =   "台灣商標延展開拓(智慧局)"
            Index           =   6
         End
         Begin VB.Menu mnu2107 
            Caption         =   "網頁提供國內專利公報資訊"
            Index           =   7
         End
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
         Caption         =   "互惠代理人案件統計表"
         Index           =   3
      End
      Begin VB.Menu mnu22 
         Caption         =   "潛在客戶名條列印"
         Index           =   4
      End
      Begin VB.Menu mnu22 
         Caption         =   "潛在客戶資料查詢"
         Index           =   5
      End
      Begin VB.Menu mnu22 
         Caption         =   "往來記錄資料查詢"
         Index           =   6
      End
      Begin VB.Menu mnu22 
         Caption         =   "往來記錄統計"
         Index           =   7
      End
      Begin VB.Menu mnu22 
         Caption         =   "國外部新客戶/代理人查詢"
         Index           =   8
      End
      Begin VB.Menu mnu22 
         Caption         =   "整批匯入至往來記錄"
         Index           =   9
      End
      Begin VB.Menu mnu22 
         Caption         =   "潛在案量客戶名稱比對"
         Index           =   10
      End
      Begin VB.Menu mnu22 
         Caption         =   "行事曆資料維護"
         Index           =   11
      End
      Begin VB.Menu mnu22 
         Caption         =   "行事曆提醒通知"
         Index           =   12
      End
      Begin VB.Menu mnu22 
         Caption         =   "整批匯入為潛在客戶"
         Index           =   13
      End
      Begin VB.Menu mnu22 
         Caption         =   "不得宣傳客戶名稱資料查詢"
         Index           =   14
      End
      Begin VB.Menu mnu22 
         Caption         =   "代理人編號匯出案件統計及互惠狀況"
         Index           =   15
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
         Begin VB.Menu mnu2303 
            Caption         =   "專業部 主管分案"
            Index           =   4
         End
      End
      Begin VB.Menu mnu23 
         Caption         =   "教育訓練登錄作業"
         Index           =   4
      End
      Begin VB.Menu mnu23 
         Caption         =   "客戶端平台帳號管理作業"
         Index           =   5
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
         Caption         =   "商標監控系統"
         Index           =   11
      End
      Begin VB.Menu mnu23 
         Caption         =   "風險檢查對象資料維護"
         Index           =   12
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "研發開拓"
      Index           =   24
      Begin VB.Menu mnu24 
         Caption         =   "專利公報IPC分類案件市佔分析"
         Index           =   0
      End
      Begin VB.Menu mnu24 
         Caption         =   "參考名條/不列印名單/新舊縣市名稱維護作業"
         Index           =   2
      End
      Begin VB.Menu mnu24 
         Caption         =   "開拓資料本所客戶檢查"
         Index           =   3
      End
      Begin VB.Menu mnu24 
         Caption         =   "價目表查詢權限維護"
         Index           =   4
      End
      Begin VB.Menu mnu24 
         Caption         =   "價目表資料維護"
         Index           =   5
      End
      Begin VB.Menu mnu24 
         Caption         =   "價目表公告公文資料維護"
         Index           =   6
      End
      Begin VB.Menu mnu24 
         Caption         =   "電子報特殊名單維護"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnu24 
         Caption         =   "專利公報excel"
         Index           =   8
         Begin VB.Menu mnu2408 
            Caption         =   "專利公報市場排名"
            Index           =   1
         End
         Begin VB.Menu mnu2408 
            Caption         =   "專利公報市場占有率比較"
            Index           =   2
         End
         Begin VB.Menu mnu2408 
            Caption         =   "各單位專利公報件數統計"
            Index           =   3
         End
         Begin VB.Menu mnu2408 
            Caption         =   "專利公報國內各區同業排名"
            Index           =   4
         End
         Begin VB.Menu mnu2408 
            Caption         =   "專利公報國外同業排名"
            Index           =   5
         End
         Begin VB.Menu mnu2408 
            Caption         =   "國籍及洲別統計(含同業)"
            Index           =   6
         End
      End
      Begin VB.Menu mnu24 
         Caption         =   "專利公報統計"
         Index           =   9
         Begin VB.Menu mnu2409 
            Caption         =   "公開及公告市場統計表"
            Index           =   1
         End
         Begin VB.Menu mnu2409 
            Caption         =   "公報產業分類案件市佔分析"
            Index           =   2
         End
      End
      Begin VB.Menu mnu24 
         Caption         =   "商標公報統計"
         Index           =   12
         Begin VB.Menu mnu2412 
            Caption         =   "表一＆表二、商標全國市場統計"
            Index           =   1
         End
         Begin VB.Menu mnu2412 
            Caption         =   "表三、各區市場佔有統計"
            Index           =   2
         End
         Begin VB.Menu mnu2412 
            Caption         =   "表四、各類別市場佔有統計"
            Index           =   3
         End
         Begin VB.Menu mnu2412 
            Caption         =   "表五、國外市場排名"
            Index           =   4
         End
         Begin VB.Menu mnu2412 
            Caption         =   "表一～表五統計表列印"
            Index           =   5
         End
         Begin VB.Menu mnu2412 
            Caption         =   "國外前十大申請國及其商品類別排名"
            Index           =   6
         End
         Begin VB.Menu mnu2412 
            Caption         =   "代理人國外案件排名分析表"
            Index           =   7
         End
         Begin VB.Menu mnu2412 
            Caption         =   "本所案件申請人國籍統計表"
            Index           =   8
         End
      End
      Begin VB.Menu mnu24 
         Caption         =   "商標公報資料統計-Excel"
         Index           =   13
         Begin VB.Menu mnu2413 
            Caption         =   "申請人國籍及洲別統計(含同業)"
            Index           =   1
         End
         Begin VB.Menu mnu2413 
            Caption         =   "各單位公報類別數統計"
            Index           =   2
         End
         Begin VB.Menu mnu2413 
            Caption         =   "三部門案件來源比較"
            Index           =   3
         End
         Begin VB.Menu mnu2413 
            Caption         =   "同業案件來源比較"
            Index           =   4
         End
         Begin VB.Menu mnu2413 
            Caption         =   "同業台灣各區類別數比較"
            Index           =   5
         End
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
   Begin VB.Menu mnuMouseR 
      Caption         =   "按右鍵用"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "複製"
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
   Begin VB.Menu mnuPopEMail1 
      Caption         =   "mnuPopEMail1"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPopEMail2 
      Caption         =   "mnuPopEMail2"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPopEMail3 
      Caption         =   "mnuPopEMail3"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

'Add by Morgan 2003/12/22
Dim WithEvents eventConn As ADODB.Connection
Attribute eventConn.VB_VarHelpID = -1
Public bolReOpen As Boolean

'move to basquery 2007/02/07
'intPCaseKind分案之系統分類，intPWhere 0國內  1國外CF  2國外FC
'Public intPCaseKind As Integer, intPWhere As Integer
'add by nick 2004/09/27 品名查詢用
Public CopyWord As String
'Add by Morgan 2008/11/7 是否已經做過
Dim m_blnActivated As Boolean
Public m_wasMaximized As Boolean 'Added by Morgan 畫面最小化後判斷原來是否為最大化用
Public m_ChkIsOpenFrm180203 As Boolean 'Add By Sindy 2013/7/8
'Added by Lydia 2019/06/27
Dim PersonDate1 As String '個人使用記錄的前一天
Dim PersonList(1 To 10) As String '個人常用區的清單(前10名)
Dim oControl As Control  'Added by Morgan 2022/1/22
Dim bolUpdNew As Boolean 'Added by Lydia 2023/03/09 個人常用區是否使用接洽單(電子收文)


'Add by Morgan 2003/12/22
Private Sub eventConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   If strUserNum = "92012" Then
      Pub_WriteSysLog pCommand.CommandText
   End If
   tmrConnect.Tag = 0
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

Private Sub SalesDueCaseQuery()
   Dim tmpST15 As String
   'add by nickc 2005/09/20
   Dim tmpST03 As String
   Dim strTo As String
   
   'add by nickc 2005/09/02 非員工不跑
   '2011/3/30 MODIFY BY SONIA
   'If strUserNum >= "63001" And strUserNum < "A" Then
   If strUserNum >= "63001" And strUserNum < "F" Then
      tmpST15 = PUB_GetStaffST15(strUserNum, 1)
      'add by nickc 2005/09/20
      tmpST03 = PUB_GetST03(strUserNum)
      'add by nickc 2005/10/14
      '2005/12/28 CANCEL BY SONIA 杜副總說開放分所使用
      'If pub_strUserOffice <> "1" Then
      '   mnu21(14).Visible = False
      '   mnu21(15).Visible = False
      'End If
      '2005/12/28 END
      'edit by nickc 2005/09/20
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Then
      'edit by nickc 2006/04/04 加入非智權人員但有收文的人
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Then
      '2006/6/1 MODIFY BY SONIA 取消P2,P3中所之控制,改由PUB_ChkNotSalesButHaveCase控制
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
      '2012/7/18 MODIFY BY SONIA 商標處所有人都要
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
      'modify by sonia 2019/7/18 W10部門也要  2019/9/5 改W部門都要
      If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST15, 1, 2)) = "P2" Or UCase(Mid(tmpST15, 1, 1)) = "W" Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
'         If ServerTime <= 100000 Then
'            pub_CallNextForm = True
'            frm100123.Show
'            frm100123.cmdSearch_Click
'         Else
'            If MsgBox("是否執行  智權人員期限資料查詢  功能", vbYesNo, "功能！") = vbYes Then
'               pub_CallNextForm = True
'               frm100123.Show
'               frm100123.cmdSearch_Click
'            End If
'         End If
         '電腦中心除外
         'If Pub_StrUserSt03 <> "M51" Then   'cancel by sonia 2017/9/28 否則薛經理74001不會跑
            strSql = "select * from executelog where el01='frm100123' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI <> 1 Then
               pub_CallNextForm = True
               frm100123.Show
               frm100123.cmdSearch_Click
            Else
               If MsgBox("是否執行  智權人員期限資料查詢  功能", vbYesNo, "功能！") = vbYes Then
                  pub_CallNextForm = True
                  frm100123.Show
                  frm100123.cmdSearch_Click
               End If
            End If
         'End If   'cancel by sonia 2017/9/28 否則薛經理74001不會跑
      End If
      
      'Add By Sindy 2020/9/10
      strTo = GetCaseDutyAgent(Pub_GetSpecMan("商標處信件檢核表收受者"), "", False) '抓請假職代
      If InStr(Pub_GetSpecMan("商標處信件檢核表收受者") & IIf(strTo <> "", ";" & strTo, ""), strUserNum) > 0 Then
         If CheckUse("frm100106_9", strExec, False) = True Then
            '一天僅自動彈跳通知一次
            strSql = "select * from executelog where el01='frm100106_9' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI <> 1 Then
               frm100106_9.m_WorkType = 1
               Load frm100106_9 '未處理信件查詢
            End If
         End If
      End If
      
'      'Add By Sindy 2020/2/3
'      If CheckUse("frm020201", strExec, False) = True And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'         '一天僅自動彈跳通知一次
'         'strSQL = "select * from executelog where el01='frm020201' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
'         strSql = "select * from executelog where el01='frm020201' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI <> 1 Then
'            Load frm020201 '商標處期限通知
'            frm020201.cmdQuery(0).Value = True
'         End If
'      End If
   End If
End Sub

Private Sub FcpDueCaseQurey()
Dim tmpST15 As String
   
   'Modify By Sindy 2020/8/25 Promoter:限FXX及P2X部門人員
   tmpST15 = PUB_GetStaffST15(strUserNum, 1)
   If UCase(Mid(tmpST15, 1, 1)) <> "F" And UCase(Mid(tmpST15, 1, 2)) <> "P2" Then Exit Sub
   '2020/8/25 END
   
   '電腦中心除外
   If Pub_StrUserSt03 <> "M51" Then
'      'Add by Morgan 2009/12/15 改先執行FMP案已達約定期限通知
'      If CheckUse("frm060206", strExec, False) = True Then
'         strSql = "select * from executelog where el01='frm060206' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI <> 1 Then
'            Load frm060206 '外專非台灣案已達約定期限通知
'            frm060206.cmdQuery.Value = True
'         End If
'      End If
'      'end 2009/12/15
      
      If CheckUse("frm060204", strExec, False) = True Then
         strSql = "select * from executelog where el01='frm060204' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI <> 1 Then
            Load frm060204 '國外部專利處期限通知
            frm060204.cmdQuery(0).Value = True
         End If
      End If
   End If
End Sub

'Add By Sindy 2015/7/2
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

'Modify by Morgan 2008/11/7 登入系統後要執行的程式集中到這裡做
Private Sub MDIForm_Activate()
   'Modify By Sindy 2025/11/3 改為共用函數
   Call MDIFormStarProc
End Sub

'Add By Sindy 2011/10/7
Public Sub SysStartCallForm()
   '只要執行一次
   If m_blnActivated = False Then
      m_blnActivated = True
      '智權人員期限資料查詢
      SalesDueCaseQuery
      'Added by Lydia 2016/05/06 提示查名人員有到期的查名單
      If Left(Pub_StrUserSt03, 2) = "P2" Then
         Call CheckTMQ10Alert
      End If
   End If
   
   '本查詢需考慮當閒置太久重新登入且已經是下午時須再次執行故與單獨控制
   If pub_bolInformCheck = True Then
      '國外部期限查詢
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

Private Sub MDIForm_Resize()
   'Added by Morgan 2011/12/14 紀錄是否為最大化狀態
   If Me.WindowState = 2 Then
      m_wasMaximized = True
   ElseIf Me.WindowState = 0 Then
      m_wasMaximized = False
   End If
End Sub

Private Sub mnu03_Click(Index As Integer)
   ToolHide
   Select Case Index
   'Modified by Lydia 2016/01/11 調整功能表項目
'      Case 0 '撰寫信函
'        frm090401.Show
'      Case 1 '工作進度資料維護(國外部)
'         If CheckUse("frm090901", strExec) Then
'            frm090901.Show
'         End If
'      Case 2 '承辦人請款/發文明細表(國外部)
'         If CheckUse("frm060312", strExec) = True Then
'            frm060312.Show
'         End If
'      'Add By Morgan 2008/9/22
'      Case 3 '國外部專利處期限通知
'         If CheckUse("frm060204", strExec) = True Then
'            frm060204.Show
'         End If
'      '2009/5/19 add by sonia
'      Case 5 '國外FC帳款明細表
'         If CheckUse("Frmacc24i0", strExec) = False Then
'            Screen.MousePointer = vbDefault
'            Exit Sub
'         End If
'         If Mid(GetStaffDepartment(strUserNum), 1, 2) = "F2" Then
'            Frmacc24i0.Text7 = "F20"
'            Frmacc24i0.Text8 = "F29"
'         End If
'         Frmacc24i0.Show
'      '2009/6/25 ADD BY SONIA
'      Case 6 '預估結匯匯率資料維護
'         If CheckUse("Frmacc21o0", strExec) = False Then
'            Screen.MousePointer = vbDefault
'            Exit Sub
'         End If
'         Frmacc21o0.Show
'      'Add By Sindy 2009/08/28
'      Case 7 '國外部法務處期限通知
'         If CheckUse("frm082007", strExec) = True Then
'            frm082007.Show
'         End If
'      '2009/08/28 End
'      'Add By Morgan 2009/12/15
'      Case 8 '外專非台灣案已達約定期限通知
'         If CheckUse("frm060206", strExec) = True Then
'            frm060206.Show
'         End If
'      'Add By Morgan 2009/11/19
'      Case 9 '定稿資料維護
'         If CheckUse("frm1105", strExec) = True Then
'            frm1105.Show
'         End If
'      'Add By Sindy 2010/7/14
'      Case 10 '專利日文資料維護作業
'         If CheckUse("frm140110", strExec) = True Then
'            frm140110.Show
'         End If
'      Case 11 '客戶日文資料維護作業
'         If CheckUse("frm140108", strExec) = True Then
'            frm140108.Show
'         End If
'      Case 12 '代理人日文資料維護作業
'         If CheckUse("frm140109", strExec) = True Then
'            frm140109.Show
'         End If
'      '2010/7/14 End
'      '2010/12/30 ADD BY SONIA
'      Case 13 '專利基本檔維護
'         If CheckUse("frm050701", strExec) = True Then
'            strSysKind = "FCP"
'            frm050701.Show
'         End If
'      Case 14 '客戶基本資料維護
'         If CheckUse("frm140401", strExec) = True Then
'            frm140401.Show
'         End If
'      Case 15 '國外代理人資料
'         If CheckUse("frm050705", strExec) = True Then
'            strSysKind = "FCP"
'            frm050705.Show
'         End If
'      '2010/12/30 END
'      '2011/10/11 自智權部作業移過來
'      Case 16 '國外代理人帳目查詢
'         If CheckUse("Frmacc2210", strExec) = False Then
'            Screen.MousePointer = vbDefault
'            Exit Sub
'         End If
'         Frmacc2210.Show
'         ToolShow
'         tool3_enabled
'      Case 17 '國外案件帳目查詢
'         If CheckUse("Frmacc2220", strExec) = False Then
'            Screen.MousePointer = vbDefault
'            Exit Sub
'         End If
'         Frmacc2220.Show
'         ToolShow
'         tool3_enabled
'      '2011/10/11 end
'      'Added by Morgan 2012/6/21
'      Case 18 '期限資料結案單
'         frm210133.Show
'      'Add by Amy 2013/07/17
'      Case 19 '案件命名追蹤
'        If CheckUse("frm060504", strExec) Then
'            frm060504.Show
'        End If
'     'Add by Amy 20130914
'     Case 20 '各幣別最新請款匯率查詢
'        If CheckUse("Frmacc2142", strExec) = False Then
'            Exit Sub
'        End If
'        Frmacc2142.Show
'      'Add By Sindy 2013/9/3
'      Case 21 '待核判區
'         'If CheckUse("frm090202_1", strExec) Then
'            frm090202_1.m_ProSysState = "1" '承辦人
'            frm090202_1.Show
'         'End If
'      'add by sonia 2013/12/5 因洪丹怡留職停薪,故開放79034王俊傑,94006宗家澔可使用
'      Case 22 '電話聯絡單發文
'         If CheckUse("frm060104_h", strExec) Then
'            frm060104_h.Show
'         End If
'      'Added by Lydia 2015/12/10
'      Case 23   '案件基本資料-外專承辦組
'         If CheckUse("frm060116", strExec) = True Then
'            strSysKind = "FCP"
'            frm060116.Caption = "案件基本資料-外專承辦組"
'            frm060116.Show
'         End If
      Case 1 '撰寫信函
        frm090401.Show
      Case 2 '工作進度資料維護(國外部)
         If strSrvDate(1) >= 外專承辦歷程啟用日 Then
            'Add By Sindy 2025/3/3
            If PUB_ChkFormIsClose("frm090909", "工作進度資料維護") = False Then
               Exit Sub
            Else
            '2025/3/3 END
               If CheckUse("frm090201_4", strExec) Then
                  ProState = "1" '個人
                  ProSysState = "1" '承辦人
                  frm090201_4.StrMenu1 '當天本所期限案件資料,無資料時由frm090201_4的nextstep執行下一畫面
                  If frm090201_4.TextOk = True Then frm090201_4.Show
               End If
            End If
         Else
            If CheckUse("frm090901", strExec) Then
               frm090901.Show
            End If
         End If
      'Add By Sindy 2023/9/28
      Case 3 '待核判區
         frm090202_1.m_ProSysState = "1" '承辦人
         frm090202_1.Show
      'Add By Sindy 2023/12/18
      Case 4 '工作進度資料查詢
         If CheckUse("frm090203_1", strExec) Then
            ProState = "1"
            ProSysState = "1"
            frm090203_1.Show
         End If
      Case 5 '承辦人工作進度資料查詢
         If CheckUse("frm090614", strExec) Then
            ProState = "2"
            ProSysState = "1"
            frm090614.m_ProState = "FCP" 'Add By Sindy 2024/2/23
            frm090614.Show
         End If
      Case 6 '承辦人請款/發文明細表(國外部)
         If CheckUse("frm060312", strExec) = True Then
            frm060312.Show
         End If
      'Add By Morgan 2008/9/22
      Case 7 '國外部專利處期限通知
         If CheckUse("frm060204", strExec) = True Then
            frm060204.Show
         End If
'      'Add By Morgan 2009/12/15
'      Case 5 '外專非台灣案已達約定期限通知
'         If CheckUse("frm060206", strExec) = True Then
'            frm060206.Show
'         End If
      'Add By Sindy 2009/08/28
      Case 8 '法務處期限通知
'         If CheckUse("frm082007", strExec) = True Then
'            frm082007.Show
'         End If
         If CheckUse("frm072005", strExec) = True Then
            frm072005.Show
         End If
      'Added By Lydia 2016/01/12
      Case 9 '國外部行事曆提醒通知
         'Memo by Lydia 2020/01/15 更名為「行事曆提醒通知」
         If CheckUse("frm060209", strExec) = True Then
            frm060209.Show
         End If
      '2009/5/19 add by sonia
      Case 10 '國外FC帳款明細表
         If CheckUse("Frmacc24i0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         If Mid(GetStaffDepartment(strUserNum), 1, 2) = "F2" Then
            Frmacc24i0.Text7 = "F20"
            Frmacc24i0.Text8 = "F29"
         End If
         Frmacc24i0.Show
     'Add by Amy 20130914
      Case 11 '各幣別最新請款匯率查詢
         If CheckUse("Frmacc2142", strExec) = False Then
            Exit Sub
         End If
         Frmacc2142.Show
      'add by sonia 2013/12/5 因洪丹怡留職停薪,故開放79034王俊傑,94006宗家澔可使用
      'modify by sonia 2016/5/26 因宗家澔離職故再開放A0022任政宏處理王俊傑案件
      Case 12 '電話聯絡單發文
         If CheckUse("frm060104_h", strExec) Then
            frm060104_h.Show
         End If
      'Add By Sindy 2010/7/14
      Case 13 '專利日文資料維護作業
         If CheckUse("frm140110", strExec) = True Then
            frm140110.Show
         End If
      'Add By Sindy 2010/7/14
      Case 14 '客戶日文資料維護作業
         If CheckUse("frm140108", strExec) = True Then
            frm140108.Show
         End If
      'Add By Sindy 2010/7/14
      Case 15 '代理人日文資料維護作業
         If CheckUse("frm140109", strExec) = True Then
            frm140109.Show
         End If
      'Add By Sindy 2017/11/10
      Case 16 '工程師各式申請書
         If CheckUse("frm090904", strExec) = True Then
            frm090904.Show
         End If
      'Added by Lydia 2018/03/12
      Case 18 '工程師上傳作業
         If CheckUse("frm090905", strExec) = True Then
            frm090905.Show
         End If
      'Added by Lydia 2018/08/15
      Case 19 '外專翻譯分案作業-認翻譯
         If CheckUse("frm060122", strExec) = True Then
             frm060122.PubRole = "2"
             frm060122.Show
         End If
      'Added by Lydia 2021/08/16
      Case 20 '外專-專利連結通知維護作業
         If CheckUse("frm090907", strExec) = True Then
             frm090907.Show
         End If
      'Added by Lydia 2023/05/05
      Case 21 '外專新案認領區
         If CheckUse("frm090908", strExec) = True Then
             frm090908.Show
         End If
      'Added by Lydia 2023/06/28
      Case 22 '外專-藥證號維護作業
         If CheckUse("frm090910", strExec) = True Then
            frm090910.Show
         End If
   End Select
End Sub

'Add by Morgan 2004/1/6
'專利處-繪圖人員作業-公報簡訊管理
Private Sub mnu050310_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '公報簡訊個人輸入作業
      If CheckUse("frm090208_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090208_1.Show
      End If
   Case 2   '公報簡訊查詢列印
      If CheckUse("frm090212_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090212_1.Show
      End If
   End Select
End Sub


'Add by Amy 2015/09/04 將承辦人拆成專利處及商標處
'商標處
Private Sub mnu07_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1  'add by nickc 2006/06/13 商標商品名稱維護
      If CheckUse("frm03010303_05", strExec) Then
         frm03010303_05.Show
      End If
   Case 4 '撰寫信函
      frm090401.Show
   'Add by Sindy 2013/8/30
   Case 5 'T案大陸指示信
      If CheckUse("frm020107_1", strExec) = True Then
         frm020107_1.Show
      End If
   'Added by Lydia 2016/09/07 TC案陸代申請書輸入
   Case 6
      If CheckUse("frm020109", strExec) = True Then
         frm020109.Show
      End If
   Case 8 '聯絡單列印及E-Mail    '2011/9/22 加入
      frm1106.Show
   'Add By Sindy 2020/2/3
   Case 11 '商標處期限通知
      If CheckUse("frm020201", strExec) = True Then
         frm020201.Show
      End If
   'Add by Amy 2019/12/09
   Case 12 '公文來函判發作業
      frm040113.Show
   'Add By Sindy 2020/12/7
   Case 13 '發後補看作業
      '考慮職代問題,不必鎖權限
      frm040117.m_ProState = "T"
      frm040117.Show
   Case Else
   End Select
End Sub

'商標處-商品委查作業
'Mark by Lydia 2024/11/07 重新整理功能表
'Private Sub mnu0701_Click_old(Index As Integer)
'   ToolHide
'   Select Case Index
'      'Memo by Lydia 2016/04/28 查名單電子化已上線,隱藏
'      'Remark by Lydia 2022/02/22 已刪除表單
''      Case 1 '商標委查資料維護
''         If CheckUse("frm090101", strExec) Then
''            frm090101.Show
''         End If
''      Case 2 '查覆日期輸入
''         If CheckUse("frm090102", strExec) Then
''            frm090102.Show
''         End If
'      'end 2016/04/28
'      'end 2022/02/22
'      Case 3 '查名期限資料查詢
'         If CheckUse("frm090103_1", strExec) Then
'            frm090103_1.Show
'         End If
'      Case 4 '委查組群統計
'         If CheckUse("frm090112", strExec) Then
'            frm090112.Show
'         End If
'      Case 5 '查名人查覆明細表
'         If CheckUse("frm090104_1", strExec) Then
'            frm090104_1.Show
'         End If
'      Case 6 '期限過期明細表
'         If CheckUse("frm090105_1", strExec) Then
'            frm090105_1.Show
'         End If
'      Case 7 '委查人委查明細表
'         If CheckUse("frm090106_1", strExec) Then
'            frm090106_1.Show
'         End If
'      Case 8 '委名人委覆統計表
'         If CheckUse("frm090107_1", strExec) Then
'            frm090107_1.Show
'         End If
'      Case 9 '委查人委查統計表
'         If CheckUse("frm090108_1", strExec) Then
'            frm090108_1.Show
'         End If
'      Case 10 '商品組群委查統計表
'         If CheckUse("frm090109_1", strExec) Then
'            frm090109_1.Show
'         End If
'      Case 11 '委查資料刪除作業
'         If CheckUse("frm090110", strExec) Then
'            frm090110.Show
'         End If
'      'Remark by Lydia 2022/02/22 已刪除表單
'      'Case 12 '商品組群查名工作天資料維護
'      '   If CheckUse("frm090111", strExec) Then
'      '      frm090111.Show
'      '   End If
'      'end 2022/02/22
'      Case 13 '圖形查名單未查覆統計表
'         If CheckUse("frm090113", strExec) Then
'            frm090113.Show
'         End If
'      'Remark by Lydia 2022/02/22 已刪除表單
''      Case 14 '已查未覆狀態登陸
''         If CheckUse("frm090114", strExec) Then
''            frm090114.Show
''         End If
''      'add by nick 2005/02/04
''      Case 15 '組群總本數維護
''         If CheckUse("frm090115", strExec) Then
''            frm090115.Show
''         End If
''      'add by nickc 2007/11/15   加入類似組群維護
''      Case 16
''         If CheckUse("frm090116", strExec) Then
''            frm090116.Show
''         End If
''      'add by nickc 2007/11/15   加入查名單列印
''      Case 17
''            frm090117.Show
''      'add by nickc 2008/03/03   加入查名順序
''      Case 18
''          If CheckUse("frm090119", strExec) Then
''            frm090119.Show
''          End If
'      'end 2022/02/22
'      'Add By Sindy 2010/01/15
'      Case 19 '刪除組群維護
'         If CheckUse("frm090120", strExec) = True Then
'            frm090120.Show
'         End If
'      '2010/01/15 End
'      'Added by Lydia 2015/05/26
'      Case 20 '查名人員維護
'         If CheckUse("frm090122", strExec) = True Then
'            frm090122.Show
'         End If
'      Case 21 '查名人狀態
'         If CheckUse("frm090124", strExec) = True Then
'            Set frm090124.mPreForm = Me
'            frm090124.Show
'         End If
'      'Modified by Lydia 2024/07/09 和圖形查名路徑合併為子選單
'      'Case 22 '查名組群及本數維護
'      '   If CheckUse("frm090123", strExec) = True Then
'      '      frm090123.Show
'      '   End If
'      'end 2015/05/26
'      'end 2024/07/09
'      'Added by Lydia 2015/07/21
'      'Remark by Lydia 2022/02/22 已刪除表單
'      'Case 23 '委查單資料維護
'      '   If CheckUse("frm090125", strExec) = True Then
'      '      frm090125.Show
'      '   End If
'      'end 2022/02/22
'      'Added by Lydia 2015/11/05 查名單電子化-主功能表
'      Case 24 '待查區
'            If frm090127.IsRolePlay("待查") = True Then
'               SetTmpTMQ
'               frm090127.Show
'            End If
'      Case 25 '查覆區
'            If frm090127.IsRolePlay("查覆") = True Then
'               SetTmpTMQ
'               frm090127.Show
'            End If
'      Case 26 '覆核區
'            If frm090127.IsRolePlay("覆核") = True Then
'               SetTmpTMQ
'               frm090127.Show
'            End If
'      'end 2015/11/05
'      'Added by Lydia 2015/11/17 查名單維護(限電腦中心)
'      Case 27
'            If frm090127.IsRolePlay("維護") = True Then
'               SetTmpTMQ
'               frm090127.Show
'            End If
'   End Select
'End Sub
'end 2024/11/07


'Added by Lydia 2024/11/07
'商標處-商品委查作業
Private Sub mnu0701_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '待查區(網中)
         If frm090127_New.IsRolePlay("待查") = True Then
            frm090127_New.Show
         End If
      Case 2 '查覆區(網中)   ----先暫時放在這,上線後移到智權部
         If frm090127_New.IsRolePlay("查覆") = True Then
            frm090127_New.Show
         End If
      Case 3 '覆核區(網中)
         If frm090127_New.IsRolePlay("覆核") = True Then
            frm090127_New.Show
         End If
      Case 4 '查名單維護(限電腦中心)
         If frm090127_New.IsRolePlay("維護") = True Then
            frm090127_New.Show
         End If
      Case 5 '查名期限資料查詢
         If CheckUse("frm090103_1", strExec) Then
            frm090103_1.Show
         End If
      Case 6 '委查組群統計
         If CheckUse("frm090112", strExec) Then
            frm090112.Show
         End If
      Case 7 '查名人查覆明細表
         If CheckUse("frm090104_1", strExec) Then
            frm090104_1.Show
         End If
      Case 8 '期限過期明細表
         If CheckUse("frm090105_1", strExec) Then
            frm090105_1.Show
         End If
      Case 9 '委查人委查明細表
         If CheckUse("frm090106_1", strExec) Then
            frm090106_1.Show
         End If
      Case 10 '委名人委覆統計表
         If CheckUse("frm090107_1", strExec) Then
            frm090107_1.Show
         End If
      Case 11 '委查人委查統計表
         If CheckUse("frm090108_1", strExec) Then
            frm090108_1.Show
         End If
      Case 12 '商品組群委查統計表
         If CheckUse("frm090109_1", strExec) Then
            frm090109_1.Show
         End If
      Case 13 '委查資料刪除作業
         If CheckUse("frm090110", strExec) Then
            frm090110.Show
         End If
      Case 14 '刪除組群維護
         If CheckUse("frm090120", strExec) = True Then
            frm090120.Show
         End If
      Case 15 '查名人員維護
         If CheckUse("frm090122", strExec) = True Then
            frm090122.Show
         End If
      Case 16 '查名人狀態
         If CheckUse("frm090124", strExec) = True Then
            Set frm090124.mPreForm = Me
            frm090124.Show
         End If
      'Memo  by Lydia 2024/07/09 和圖形查名路徑合併為子選單
      'Case 17 '查名組群及本數維護
      'Added by Lydia 2015/11/05 查名單電子化-主功能表
      Case 18 '待查區
            If frm090127.IsRolePlay("待查") = True Then
               SetTmpTMQ
               frm090127.Show
            End If
      Case 19 '覆核區
            If frm090127.IsRolePlay("覆核") = True Then
               SetTmpTMQ
               frm090127.Show
            End If
      'end 2015/11/05
      'Added by Lydia 2015/11/17 查名單維護(限電腦中心)
      Case 20
            If frm090127.IsRolePlay("維護") = True Then
               SetTmpTMQ
               frm090127.Show
            End If
   End Select
End Sub

'Added by Lydia 2015/11/05 查名單電子化-主功能表
Private Sub SetTmpTMQ()
   Set frm090127.Tmpfrm090126 = frm090126
   Set frm090127.Tmpfrm090128 = frm090128
   Set frm090128.Tmpfrm090129 = frm090129
   Set frm090801.Tmpfrm090126 = frm090126 'Added by Lydia 2016/05/10
End Sub
'end 2015/10/14

'Added by Lydia 2024/07/17
'Modified by Lydia 2024/11/07 重新整理功能表
'Private Sub mnu070122_Click(Index As Integer)
Private Sub mnu070117_Click(Index As Integer)
   Select Case Index
      Case 1 '查名組群及本數維護
         If CheckUse("frm090123", strExec) = True Then
            frm090123.Show
         End If
      Case 2 '圖形查名路徑-大分類維護
         If CheckUse("frm090133", strExec) = True Then
            frm090133.Show
         End If
      Case 3 '圖形查名路徑-中分類維護
         If CheckUse("frm090133", strExec) = True Then
            frm090133_1.Show
         End If
      Case 4 '圖形查名路徑-小分類維護
         If CheckUse("frm090133", strExec) = True Then
            frm090133_2.Show
         End If
   End Select
End Sub


'商標處-承辦人作業
Private Sub mnu0702_Click(Index As Integer)
'Dim bolNoCheck As Boolean 'Added by Morgan 2013/10/8
Dim strSysID As String 'Add By Sindy 2014/7/4
Dim nFrm As Form

'ProState = "1"
'ProSysState = "1"
ToolHide
Select Case Index
   Case 1 '工作進度資料維護
      If CheckUse("frm090201_4", strExec) Then
         ProState = "1"
         ProSysState = "1"
         'Added by Lydia 2015/1/13 新增查名單對應
         Set frm090201_b.Tmpfrm090130 = frm090130
'         'Modify by Sindy 2018/7/24
'         If Left(Pub_StrUserSt03, 2) = "P2" Then
'            '第2次以上可選擇
'            strSql = "select * from executelog where el01='frm090201_a' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               If MsgBox("是否執行 當天本所期限案件... 等功能", vbYesNo + vbQuestion + vbDefaultButton1, "功能！") = vbNo Then
'                  bolNoCheck = True
'               End If
'            End If
'         End If
'
'         If bolNoCheck = True Then
'            frm090201_2.Show
'         Else
'         'end 2013/10/8
         
         'Add By Sindy 2025/3/3
         If PUB_ChkFormIsClose("frm090201_b", "工作進度資料維護") = False Then
            Exit Sub
         Else
         '2025/3/3 END
            '2009/11/12 modify by sonia 改寫法無資料不顯示畫面
            'frm090201_4.Show
            frm090201_4.StrMenu1   '當天本所期限案件資料,無資料時由frm090201_4的nextstep執行下一畫面
            If frm090201_4.TextOk = True Then frm090201_4.Show
            '2009/11/12 end
            
            'Add By Sindy 2019/10/18
            '檢查表單是否已開啟，若是，則關閉
            For Each nFrm In Forms
               If StrComp(nFrm.Name, "frm090202_1", vbTextCompare) = 0 Then
                  Unload frm090202_1
               End If
            Next
            frm090202_1.m_ProSysState = "1" '承辦人
            frm090202_1.Hide
            If frm090202_1.Tag = "有黃色期限" Then
               If MsgBox("您有待核判案件逾時，是否直接進入待核判區處理？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                  frm090202_1.Show
               Else
                  Unload frm090202_1
               End If
            Else
               Unload frm090202_1
            End If
            '2019/10/18 END
         End If
'         End If 'Added by Morgan 2013/10/8
      End If
   
   'Add By Sindy 2018/4/26
   Case 2 '待核判區
      frm090202_1.m_ProSysState = "1" '承辦人
      frm090202_1.Show
      
   Case 8 '未齊備、未完稿、未發文查詢
      If CheckUse("frm0906121", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090612.Show
      End If
      
   Case 9 '工作進度資料查詢
      If CheckUse("frm090203_1", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090203_1.Show
      End If
   Case 10 '承辦人達成情形查詢
      If CheckUse("frm0906081", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090608.Show
      End If
   Case 13
      If CheckUse("frm090205_1", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090205_1.Show
      End If
   'Add By Sindy 2010/01/21 電話回覆主管機關
   'Modified by Lydia 2015/11/23 改index=18
   'Case 20
   Case 18
      If CheckUse("frm090219", strExec) Then
         frm090219.Show
      End If
   'Added by Lydia 2015/11/23 主管機關來電處理記錄
   Case 19
      If CheckUse("frm020108", strExec) Then
         frm020108.Show
      End If
   'Add by Morgan 2011/6/14
   'Modify By Sindy 2023/12/8 mark,此作業功能已併入frm090202_2電子承辦單簽辦作業
'   Case 21 '承辦人電子送件作業
'      If CheckUse("frm090220", strExec) Then
'         frm090220.Show
'      End If
   'Add by Sindy 2012/2/9
   Case 22 '商標委任書正本案號維護
      If CheckUse("frm090221", strExec) Then
         'Modify By Sindy 2014/7/4
         If GetStaffDepartment(strUserNum) = "M51" Then
AgainRunFrm090221:
            strSysID = UCase(InputBox("欲操作專利還是商標？請輸入(P.專利(含空白) T.商標)：", "委任書正本案號維護", "P"))
            If Trim(strSysID) = "P" Or Trim(strSysID) = "T" Then
               frm090221.strSysID = strSysID
               frm090221.Show
            ElseIf Trim(strSysID) <> "" Then
               MsgBox "只可輸入P或T"
               GoTo AgainRunFrm090221
            End If
         Else
            frm090221.Show
         End If
      End If
   'Add By Sindy 2012/8/30
   Case 23 '台灣商標延展開拓
      If CheckUse("frm020321", strExec) = True Then
         frm020321.Show
      End If
   'Add By Sindy 2013/5/29
   Case 24 '商申承辦人內部及機關收發文統計表
      If CheckUse("frm090622", strExec) = True Then
         frm090622.Show
      End If
  'Add By Amy 2013/07/11
   Case 25 '委任狀中譯文
         If CheckUse("frm030207", strExec) = True Then
            frm030207.Show
         End If
   'Add by Amy 2015/09/04
   Case 26 '未發文案件原因註記
            frm090638.intPeople = 1
            frm090638.Show
   'Added by Lydia 2016/01/11
   Case 27 '案件催審延緩維護
         If CheckUse("frm090222", strExec) = True Then
            frm090222.Show
         End If
   'Added by Lydia 2016/01/11
   Case 28 '案件催審作業
         If CheckUse("frm090223", strExec) = True Then
            frm090223.Show
         End If
   'Added by Lydia 2016/01/15
   Case 29 '台灣案催審申請書(定稿資料維護)
         If CheckUse("frm1105", strExec) = True Then
            frm1105.Show
         End If
         
   'Added by Morgan 2025/8/20
   Case 30  '國外代理人帳目查詢
      If CheckUse("Frmacc2210", strExec) = False Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      
      If PUB_CheckFormExist("Frmacc2220") Then
         MsgBox "一次只可開啟一個國外帳目查詢的功能！", vbExclamation
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      
      Frmacc2210.Show
      ToolShow
      tool3_enabled
      
   Case 31  '國外案件帳目查詢
      If CheckUse("Frmacc2220", strExec) = False Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      
      If PUB_CheckFormExist("Frmacc2210") Then
         MsgBox "一次只可開啟一個國外帳目查詢的功能！", vbExclamation
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      
      Frmacc2220.Show
      ToolShow
      tool3_enabled
         
   Case 32 '各幣別最新請款匯率查詢
      If CheckUse("Frmacc2142", strExec) = False Then
         Exit Sub
      End If
      Frmacc2142.Show
   'end 2025/8/20
   Case Else
End Select
End Sub

'商標處-承辦人工作管理-查詢及報表
Private Sub mnu070703_Click(Index As Integer)
ToolHide
Select Case Index
Case 1 '承辦人工作進度資料查詢
   If CheckUse("frm090614", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090614.m_ProState = "T" 'Add By Sindy 2024/2/23
      frm090614.Show
   End If
Case 2 '承辦人達成情形查詢
   If CheckUse("frm0906082", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090608.Show
   End If
Case 3 '承辦人工作量查詢
   If CheckUse("frm090609", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090609.Show
   End If
Case 4 '承辦人每日分案情形查詢
   If CheckUse("frm090610", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090610.Show
   End If
Case 5 '承辦天數統計查詢
   If CheckUse("frm090611", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090611.Show
   End If
Case 6 '未齊備未完稿未發文查詢
   'CheckUse時於FormName後面加 1,2 區分個人及管理
   If CheckUse("frm0906122", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090612.Show
   End If
'Add by Amy 2019/07/22
Case 7 'MCT收發文件數及點數統計
    If CheckUse("frm020420", strExec) = True Then
        frm020420.Show
    End If
Case Else
End Select
End Sub

Private Sub mnu070801_Click(Index As Integer)
    ToolHide
    Select Case Index
        Case 0 '商申承辦人責任業務區分配維護
            If CheckUse("frm090226", strExec) = True Then
                frm090226.Show
            End If
    End Select
End Sub

'Add By Sindy 2019/5/29
Private Sub mnu0710_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 0 '商標處收件夾信件處理
      If CheckUse("frm090224", strExec) = True Then
         frm090224.Show
      End If
   Case 1 '郵件分信關鍵字對照表維護
      If CheckUse("frm06010614", strExec) = True Then
         frm06010614.m_strLK12 = "T"
         frm06010614.Show
      End If
   Case 2 '未處理信件查詢
      If CheckUse("frm100106_9", strExec) = True Then
         frm100106_9.m_WorkType = 1
         frm100106_9.Show
      End If
   Case Else
   End Select
End Sub

'商標處-商標處-承辦人工作管理-人員考核管理
'Remove by Lydia 2018/02/12 限專利處使用
'Private Sub mnu070704_Click(Index As Integer)
'ToolHide
'Select Case Index
'Case 1
'   If CheckUse("frm090624", strExec) Then '專利處每週速度考核表
'    ProState = "2"
'    ProSysState = "1"
'      frm090624.Show
'   End If
'Case 2 '月考核
'   If CheckUse("frm090616M", strExec) Then
'    ProState = "2"
'    ProSysState = "1"
'      frm090616_0.Show
'   End If
'Case 3 '季考核
'   If CheckUse("frm090618M", strExec) Then
'    ProState = "2"
'    ProSysState = "1"
'      frm090618.Show
'   End If
'Case 4 '工程師每月目標基數設定
'   If CheckUse("frm090615", strExec) Then
'    ProState = "2"
'    ProSysState = "1"
'      frm090631.Show
'   End If
'Case 5 '個人目標資料維護
'   If CheckUse("frm090615", strExec) Then
'    ProState = "2"
'    ProSysState = "1"
'      frm090615.Show
'   End If
'Case 6 '獎金輸入作業
'   If CheckUse("frm090617", strExec) Then
'    ProState = "2"
'    ProSysState = "1"
'      frm090617.Show
'   End If
'Case 7 '獎金明細表
'   If CheckUse("frm090619", strExec) Then
'    ProState = "2"
'    ProSysState = "1"
'      frm090619.Show
'   End If
'Case Else
'End Select
'End Sub
'end 2018/02/12

'商標處-承辦人工作管理-基本資料維護
'Remove by Lydia 2018/02/12 限專利處使用
'Private Sub mnu070705_Click(Index As Integer)
'ToolHide
'Select Case Index
'Case 1
''承辦人支援記錄維護
'   If CheckUse("frm090623M", strExec) Then
'      ProState = "2"
'      ProSysState = "1"
'      frm090623.Show
'   End If
'Case 2
''承辦人特殊案件記錄維護
'   If CheckUse("frm090627M", strExec) Then
'      ProState = "2"
'      ProSysState = "1"
'      frm090627.Show
'   End If
'
'   'Add by Morgan 2011/7/27
'   Case 3   '承辦人修改記錄維護
'      If CheckUse("frm090633M", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090633.Show
'      End If
'
'   'Add by Morgan 2011/8/1
'   Case 4   '承辦人衍生記錄維護
'      If CheckUse("frm090634M", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090634.Show
'      End If
'
'Case 5 '國內外案件資料維護
'   If CheckUse("frm050106_1", strExec) = True Then
'      ProState = "2"
'      ProSysState = "1"
'      frm050106_1.intWhereToGo = 0
'      frm050106_1.Show
'   End If
'Case 6 '每月目次重編作業
'    If CheckUse("frm090606", strExec) Then
'      ProState = "2"
'      ProSysState = "1"
'      frm090606.Show
'   End If
'Case 7 '承辦人、核稿人對照資料維護
'   If CheckUse("frm090621", strExec) Then
'      ProState = "2"
'      ProSysState = "1"
'      frm090621.Show
'   End If
'Case 8 '特別加乘註記維護
'   If CheckUse("frm090629", strExec) Then
'      ProState = "2"
'      ProSysState = "1"
'      frm090629.Show
'   End If
''英文核稿人欄修改權限設定
'Case 9
'   If CheckUse("frm090202_6", strExec) Then
'      ProState = "2"
'      ProSysState = "1"
'      frm090202_6.Show
'   End If
'Case Else
'End Select
'End Sub
''end 2015/09/04
'end 2018/02/12

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
            'Modify by Amy 2025/02/03 改顯示名稱
            'frm100102_1.Caption = "申請人查詢(查新客戶-含對造)"
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
   End Select
End Sub

Private Sub mnu103_Click(Index As Integer)
   ToolHide
   Select Case Index
      'Add by Morgan 2007/5/31
      Case 1 '外專未完成核稿明細查詢/列印
         'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm060320") = False Then
            Set frm060320 = Nothing
         End If
         'end 2021/12/16
         If CheckUse("frm060320", strExec) Then
            frm060320.Show
         End If
      'Add by Morgan 2007/5/31
      Case 2 '外專收文未發文明細查詢/列印
         'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm060309") = False Then
            Set frm060309 = Nothing
         End If
         'end 2021/12/16
         If CheckUse("frm060309", strExec) Then
            frm060309.Show
         End If
      'add by nickc 2007/07/04
      Case 3 '客戶重新委任案件查詢列印
         'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm100126_1") = False Then
            Set frm100126_1 = Nothing
         End If
         'end 2021/12/16
         If CheckUse("frm100126_1", strExec) Then
            frm100126_1.Show
         End If
      'Add by Morgan 2007/8/9
      Case 4 '外專承辦人請款/發文明細表
         'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm060312") = False Then
            Set frm060312 = Nothing
         End If
         'end 2021/12/16
         If CheckUse("frm060312", strExec) = True Then
            frm060312.Show
         End If
      'Added by Lydia 2019/02/14
      Case 5 '外專工程師請款點數和OA發文統計表
         'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm060332") = False Then
            Set frm060332 = Nothing
         End If
         'end 2021/12/16
         If CheckUse("frm060332", strExec) = True Then
            frm060332.Show
         End If
      'Added by Lydia 2019/02/27
      Case 6 '外商每月請款點數統計表
         'Added by Lydia 2021/12/16 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm060333") = False Then
            Set frm060333 = Nothing
         End If
         'end 2021/12/16
         If CheckUse("frm060333", strExec) = True Then
            frm060333.Show
         End If
   End Select
End Sub

'Mark by Lydia 2019/06/27 調整表單
'Private Sub mnu2101_Click(Index As Integer)
'   Select Case Index
'      Case 1   '國內案件接洽記錄單
'         Set frm090801.Tmpfrm090126 = frm090126 'Added by Lydia 2016/03/21 增加查名單輸入
'         frm090801.Show
'      'Add By Sindy 2010/4/23
'      Case 2 '期限資料結案單
'         frm210133.Show
'      'Add By Sindy 2013/3/15
'      Case 3 '銷案銷帳單
'         frm210139.Show
'      'add by nickc 2006/04/27
'      Case 4 '案件委任契約書
'         frm210114.Show
'      Case 5   '個人客戶資料修改
'         'Modify by Morgan 2003/12/16
'         'Dim strDeptNo As String
'         'If frm210101.IsAuthorized(strUserNum, "", False) = True And frm210101.getSalesDept(strUserNum, strDeptNo) = True Then
'         '   frm210101_1.setSalesNo strUserNum
'         '   frm210101_1.setDeptNo strDeptNo
'         '   frm210101_1.setCaller Screen.ActiveForm
'         '   frm210101_1.Show
'         'Else
'         '   frm210101.setCaller Screen.ActiveForm
'            frm210101.setNextForm "frm210101_1"
'            frm210101.Show
'         'End If
'      'Add by Morgan 2004/4/21
'      Case 6   '客戶案件資料維護
'         frm210101.setNextForm "frm210102"
'         frm210101.Show
'      Case 7   '個人行事曆維護
'          frm210110.Show
'      '2005/4/21 modify by sonia 由mnu10移至此
'      Case 8 '撰寫信函
'          frm090401.Show
''cancel by sonia 2016/7/21 查名單電子化後隱藏
''      'add by nickc 2007/11/15 加入查名單列印
''      Case 9
''            frm090117.Show
''end 2016/7/21
'      'Modified by Lydia 2015/11/05 移到查覆區後面
'      'Add by Morgan 2006/4/6
''      Case 10 '聯絡單列印及E-Mail
''         frm1106.Show
'      'Add By Sindy 2010/3/19
'      'Case 11 '查名報告
'      Case 10 '查名報告
'         frm090121.Show
''cancel by sonia 2016/7/21 已改由查覆區進入,不掛menu
''      'Added by Lydia 2015/11/05 查名單電子化
''      Case 11 '查名單輸入
''            frm090126.Show
''end 2016/7/21
'      Case 12 '查覆區 (日常或查詢)
'            If frm090127.IsRolePlay("查覆") = True Then
'               SetTmpTMQ
'               frm090127.Show
'            End If
'      'Modified by Lydia 2015/11/05 後續編號+2
'      Case 13 '聯絡單列印及E-Mail
'         frm1106.Show
'      'Add By Sindy 2010/8/19
'      Case 14 '智權部自行管制未發文案件作業
'         frm210134.Show
'      'Add By Sindy 2012/5/7
'      Case 15 '台灣商標爭議案齊備日輸入
'         frm210136.Show
'      'Add By Sindy 2013/5/16
'      Case 16 '待會稿區
'         frm090202_3.Show
'      'Added by Morgan 2013/12/2
'      Case 17 '智權人員繳款資料輸入
'         frm210141.Show
'      'Add By Amy 2015/09/04
'      Case 18 '商標未發文案件原因註記
'         frm090638.intPeople = 2
'         frm090638.Show
'      'Added by Morgan 2014/5/9
'      Case 19 '文件寄送確認
'         frm210144.Show
'      'Modify By Sindy 2014/7/10
'      Case 20 '寄件查詢
'         frm210145.intWorkItem = 1
'         frm210145.Show
'      '2014/7/10 END
'      'Modify by Amy 2016/01/14
'      Case 21 '智權點數實績與結餘輸入
'        Dim bolIsAreaAg As Boolean, bolIsRest As Boolean, bolRest1Day As Boolean
'        Dim stA0908 As String
'        'Add by Amy 2016/03/30
'        Dim st04 As String
'
'        st04 = "Y"
'        bolIsAreaAg = IsAreaAgent(strUserNum, True, stA0908, st04)
'        If stA0908 <> MsgText(601) Then bolIsRest = CheckIsPersonRest(Left(stA0908, 5), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day)
'        If bolIsAreaAg = True And (st04 = "2" Or (bolIsRest = True And bolRest1Day = True)) Then
'            If MsgBox("要以「 " & Mid(stA0908, 7) & "」職代身份進入嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
'                frm210152.strAreaManNo = Left(stA0908, 5)
'            End If
'        End If
'        'end 2016/03/30
'        frm210152.Show
'    End Select
'End Sub

'Add By Sindy 2014/12/31
'Mark by Lydia 2024/11/07
'Private Sub mnu210104_Click(Index As Integer)
'   Select Case Index
'      Case 1 '目前表單
'         frm210147.Show
'      Case 2 '簽核作業
'         frm210148.Show
'   End Select
'End Sub
'end 2024/11/07

'Mark by Lydia 2019/06/27 調整表單
'Private Sub mnu2102_Click(Index As Integer)
'Select Case Index
'      Case 1 '客戶專利案件整理表
'         'If CheckUse("frm210115", strExec) = True Then
'            frm210115.Show
'         'End If
'      'add by nickc 2005/08/22
'      Case 2   '智權人員期限資料查詢
'          frm100123.Show
'      'Add by Morgan 2005/4/7
'      Case 3 '簽收資料查詢
'         'Modified by Lydia 2017/01/26 是否需要輸入密碼
'         'frm210106.Show
'         If frm210106_1.setNextForm = "" Then
'            frm210106.Show
'         Else
'            frm210106_1.setCaller frm210106
'            frm210106_1.Show
'         End If
'
'      'Add by Morgan 2005/3/18
'      Case 4 '業績點數查詢
'         frm210104.Show
'      'Add by Morgan 2005/3/22
'      Case 5 '暫收款查詢
'         frm210105.Show
'      Case 6   '新客戶來源分析
'          frm210111.Show
'       'add by nickc 2008/03/07 收款貢獻度分析
'      Case 7
''edit by nickc 2008/04/25 改所有智權人員都可以用
''         If CheckUse("frm210120", strExec) Then
'            frm210120.Show
''         End If
'      Case 8   '智權人員收/發文量分析
'          frm210112.Show
'       'add by nickc 2008/04/15
'      Case 9  '客戶應收帳款查詢
'         frm210122.Show
'       'Add by Lydia 2014/09/26
'      Case 10 '客戶請款明細表
'         frm210146.Show
'       'add by nickc 2008/05/05 案件回覆單/接洽結案單列印
'      Case 11 ' 案件回覆單
'         frm210126.Show
'      'Add by Morgan 2008/5/15
'      Case 12 '審查機關來函期限查詢
'         frm210125.Show
'
'      'Add by Morgan 2008/5/15
'      Case 13 '專業部定稿報價查詢
'         frm210124.Show
'
'      'Add By Sindy 2010/4/19
'      Case 14   '未列印收據查詢
'          frm210132.Show
'      'Add By Sindy 2011/3/1
'      Case 15 '接洽記錄單查詢及列印
'         'If CheckUse("frm12040152", strExec) = True Then
'            frm12040152.Show
'         'End If
'      'Add by Amy 2013/05/24
'      Case 16 '各區業績點數統計
'         frm210137.Show
'      'Added by Morgan 2013/12/24
'      Case 17 '智權人員繳款資料查詢
'         frm210142.Show
'      'Added by Sindy 2014/3/6
'      Case 18 '價目表查詢
'         frm210143.Show
'      'Added by Lydia 2015/11/30
'      Case 19 'CFP常辦國家年費(延展費)預估報價
'         frm210151.Show
'      'Add by Amy 2019/05/31
'      Case 20 '下一程序接洽單列印
'         frm210154.Show
'    End Select
'End Sub

'Mark by Lydia 2019/06/27 調整表單
'Private Sub mnu2103_Click(Index As Integer)
'Select Case Index
'      'Add by Morgan 2005/3/17
'      Case 1 '每日業績點數輸入
'         'Modify by Morgan 2005/4/26 為使同區智權人員皆可輸入不控制使用權限
'         'If CheckUse("frm210103", strExec) = True Then
'            frm210103.Show
'         'End If
'      'Add by Morgan 2005/10/24
'      Case 2 '各區業務工作報告統計
'         If CheckUse("frm210113", strExec) Then
'            frm210113.Show
'         End If
'      'Add by Morgan 2007/9/26
'      Case 3 '各所業務工作報告統計
'         If CheckUse("frm210117", strExec) Then
'            frm210117.Show
'         End If
'      'add by nickc 2007/10/02
'      Case 4 '客戶案件整理表記錄查詢
'         If CheckUse("frm210118", strExec) Then
'            frm210118.Show
'        End If
'      'Add by Morgan 2005/5/12
'      Case 5 '業務目標及達成通知日報表
'         If CheckUse("frm210107", strExec) Then
'            frm210107.Show
'         End If
'      'Add by Morgan 2005/5/12
'      Case 6 '業務目標及達成通知月報表
'         If CheckUse("frm210108", strExec) Then
'            frm210108.Show
'         End If
'      Case 7 '業務收/發文量比較查詢
'         If CheckUse("frm100122_1", strExec) Then
'            frm100122_1.Show
'         End If
'       'add by nickc 2008/03/28 智權部點數分析表
'      Case 8
'         If CheckUse("frm210121", strExec) Then
'            frm210121.Show
'         End If
'       'add by nickc 2008/04/15 未收款、未收齊清單列印
'      Case 9
'         If CheckUse("frm210123", strExec) Then
'            frm210123.Show
'         End If
'      'Add by Morgan 2008/9/17
'      Case 10 '新申請案收文至發文件數日數比較表
'         If CheckUse("frm210127", strExec) Then
'            frm210127.Show
'         End If
'      'Add by Sindy 2010/11/4
'      Case 11 '業績年度統計表
'         If CheckUse("frm210135", strExec) Then
'            frm210135.Show
'         End If
'      'Modify by Amy 2013/05/23 開放所有人使用
'      'Add by Morgan 2012/6/11
'      Case 12 '各區業績點數統計
'         'If CheckUse("frm210137", strExec) Then
'            frm210137.Show
'         'End If
'      'Add by Amy 2015/04/28
'      Case 13 '智權部工作報告-邱素蓮用
'         If CheckUse("frm210150", strExec) Then
'            frm210150.Show
'         End If
'      'Add by Amy 2016/03/24 從Patpro搬過來
'      Case 14 '專業達成點數表-秘書用
'         If CheckUse("Frmacc44r0", strExec) Then
'            Frmacc44r0.Show
'         End If
'   End Select
'End Sub

' Add By Sindy 98/02/25
'Modified by Lydia 2019/06/27 改menu
'Private Sub mnu2104_Click(Index As Integer)
Private Sub mnu2107_Click(Index As Integer)
   Dim bolOK As Boolean 'Added by Morgan 2019/6/26

   Select Case Index
      Case 1 '國內潛在客戶資料維護
         'If CheckUse("frm210128", strExec) = True Then
            Call Pub_AddPersonRec("frm210128") 'Added by Lydia 2019/06/27
            frm210128.Show
         'End If

      Case 2 '國內往來記錄資料維護
         'If CheckUse("frm210129", strExec) = True Then
            Call Pub_AddPersonRec("frm210129") 'Added by Lydia 2019/06/27
            frm210129.Show
         'End If

      Case 3 '國內潛在客戶資料查詢
         'If CheckUse("frm210130", strExec) = True Then
            Call Pub_AddPersonRec("frm210130") 'Added by Lydia 2019/06/27
            frm210130.Show
         'End If

      Case 4 '國內往來記錄資料查詢
         'If CheckUse("frm210131", strExec) = True Then
            Call Pub_AddPersonRec("frm210131") 'Added by Lydia 2019/06/27
            frm210131.Show
         'End If

      Case 5 '台灣商標公告近三年開拓函
         If CheckUse("frm020322", strExec) = True Then
            Call Pub_AddPersonRec("frm020322") 'Added by Lydia 2019/06/27
            frm020322.Show
         End If

      'Add By Sindy 2019/1/17
      Case 6 '台灣商標延展開拓(智慧局)
         If CheckUse("frm020323", strExec) = True Then
            Call Pub_AddPersonRec("frm020323") 'Added by Lydia 2019/06/27
            frm020323.Show
         End If
      'Added by Morgan 2019/5/13
      Case 7 '網頁提供國內專利公報資訊
         bolOK = False
         '智權部
         If Left(Pub_StrUserSt15, 1) = "S" Then
            bolOK = True
         '北所業務助理人員
         ElseIf InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) > 0 Then
            bolOK = True
         '業務助理 S1,CS,N1,K1,SA
         ElseIf CheckUse("frm210153", strExec) = True Then
            bolOK = True
         End If
         'Modified by Lydia 2019/06/27
         'If BolOk Then frm210153.Show
         If bolOK Then
             Call Pub_AddPersonRec("frm210153")
             frm210153.Show
         End If
         'end 2019/06/27
   End Select
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
      'Add by Morgan 2008/6/6
      Case 3 '互惠代理人案件統計表
         If CheckUse("frm050408", strExec) = True Then
            frm050408.Show
         End If
      'Add by TONI 2008/12/4
      Case 4 '潛在客戶名條列印
         If CheckUse("frm140409", strExec) = True Then
            frm140409.Show
         End If
      'Add by TONI 2008/12/4
      Case 5 '潛在客戶資料查詢
         If CheckUse("frm140407", strExec) = True Then
            frm140407.Show
         End If
      'Add by TONI 2008/12/4
      Case 6 '往來記錄資料查詢
         If CheckUse("frm140408", strExec) = True Then
            frm140408.Show
         End If
      'Add by Sindy 2019/12/27
      Case 7 '往來記錄統計
         If CheckUse("frm140420_1", strExec) = True Then
            frm140420_1.Show
         End If
      'Add by Sindy 2010/9/2
      Case 8 '國外部新客戶/代理人查詢
         If CheckUse("frm140412", strExec) = True Then
            frm140412.Show
         End If
      'Add By Sindy 2018/6/6
      Case 9 '整批匯入至往來記錄
         If CheckUse("frm140418", strExec) = True Then
            frm140418.Show
         End If
      'Add by Amy 2018/10/03
      Case 10 '潛在案量客戶名稱比對
         'Add by Amy 2018/10/12 +權限
         'Modify by Amy 2021/04/15 +CheckUse權限
         If CheckUse("frm140419", strExec, False) = False Then
            If GetMenuLimit("frm140419") = True Then
                frm140419.Show
            End If
         Else
            frm140419.Show
         End If
         'end 2021/04/15
      'add by sonia 2019/12/24
      Case 11 '國外部行事曆資料維護
         'Memo by Lydia 2020/01/15 更名為「行事曆資料維護」
         If CheckUse("frm06010610", strExec) = True Then
            frm06010610.Show
         End If
      Case 12 '國外部行事曆提醒通知
         'Memo by Lydia 2020/01/15 更名為「行事曆提醒通知」
         If CheckUse("frm060209", strExec) = True Then
            frm060209.m_Role = "F41" 'Added by Lydia 2020/01/15
            frm060209.Show
         End If
      'end 2019/12/24
      'Add By Sindy 2021/6/25
      Case 13 '整批匯入為潛在客戶
         If CheckUse("frm140421", strExec) = True Then
            frm140421.Show
         End If
      'Added by Lydia 2023/06/17
      Case 14  '不得宣傳客戶名稱資料查詢
         If frm100136.ChkUseRight = True Then
            frm100136.Show
         End If
      'Added by Lydia 2025/09/11
      Case 15  '代理人編號匯出案件統計及互惠狀況
         If CheckUse("frm140423", strExec) = True Then
            frm140423.Show
         End If
   End Select
End Sub

Private Sub mnu23_Click(Index As Integer)
Dim nFrm As Form
   
   Select Case Index
      Case 1 '預約作業
         frm140112.Show
      'Added by Morgan 2012/2/14
      Case 4 '專利處研討會
         frm140113.Show
      'Added by Sindy 2012/10/2
      Case 5 '客戶端平台帳號管理作業
         If CheckUse("frm140114", strExec) = True Then
            frm140114.Show
         End If
      'Add By Sindy 2016/3/21
      Case 7 '系統收件區
         'Modify By Sindy 2020/5/20
         'If Left(Pub_StrUserSt03, 2) = "P2" Then '商標處
            frm090225.Show 'MDIForm_Load有控制
         'End If
         '2020/5/20 END
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
      Case 9 '圖書借閱資料查詢 Add by Amy 2016/12/09
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
      
      'Added by Morgan 2021/1/8
      Case 11 '商標監控系統
         PUB_OpenTMMonitor
      'Add by Amy 2024/01/22
      Case 12 '風險檢查對象資料維護
         frm12040163.Show
   End Select
End Sub

'Add By Sindy 2014/12/31 案件表單查詢及簽核
Private Sub mnu2303_Click(Index As Integer)
   Select Case Index
      Case 1 '目前表單
         frm210147.Show
      Case 2 '簽核作業
         frm210148.Show
      'Add by Amy 2018/08/17
      Case 3 '結案單審核作業
         frm040118.Show
      'Add by Sindy 2022/7/22
      Case 4 '專業部主管分案作業
         If CheckUse("frm210156", strExec) = True Then 'Added by Lydia 2022/10/31
            frm210156.Show
         End If
   End Select
End Sub

'Added by Morgan 2016/1/19
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

Private Sub mnu24_Click(Index As Integer)
   Select Case Index
      'Move by Lydia 2019/10/30 從共同查詢->收發文查詢搬來
      Case 0 'Add By Sindy 2012/8/20 專利公報IPC分類案件市佔分析
         If CheckUse("frm100130", strExec) Then
            frm100130.Show
         End If
      'end 2019/10/30
'      Case 1 '大陸專利公告轉檔作業
'         If CheckUse("frm140115", strExec) = True Then
'            frm140115.Show
'         End If
      Case 2 '參考名條/不列印名單/新舊縣市名稱維護作業
         If CheckUse("frm140116", strExec) = True Then
            frm140116.Show
         End If
      Case 3 'Add by Amy 2013/05/31 開拓資料本所客戶檢查
         If CheckUse("frm210140_1", strExec) = True Then
            frm210140_1.Show
         End If
      'Add by Sindy 2014/2/25
      Case 4 '價目表查詢權限維護
         If CheckUse("frm090635", strExec) = True Then
            frm090635.Show
         End If
      'Add by Sindy 2014/2/25
      Case 5 '價目表資料維護
         If CheckUse("frm090636", strExec) = True Then
            frm090636.Show
         End If
      'Add by Sindy 2014/2/26
      Case 6 '價目表公告公文資料維護
         If CheckUse("frm090637", strExec) = True Then
            frm090637.Show
         End If
'      'Add By Sindy 2016/3/9
'      Case 7 '上網站截取資料
'         If CheckUse("frm140117", strExec) = True Then
'            frm140117.Show
'         End If
'      'Add By Sindy 2023/9/1
'      Case 7 '電子報特殊名單維護
'         If CheckUse("frm030617", strExec) = True Then
'            frm030617.m_WorkType = "M"
'            frm030617.Show
'         End If
   End Select
End Sub

'add by sonia 2017/12/7 林總同意開放楊特助可列印商標公報統計
'Modified by Lydia 2020/01/13 mnu2409=>mnu2412
Private Sub mnu2412_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '表一＆表二.商標全國市場統計表
         If CheckUse("frm030606", strExec) = True Then
            frm030606.Show
         End If
      Case 2 '表三.各區市場佔有率統計表
         If CheckUse("frm030608", strExec) = True Then
            frm030608.Show
         End If
      Case 3 '表四.各類別市場佔有統計表
         If CheckUse("frm030609", strExec) = True Then
            frm030609.Show
         End If
      Case 4 '表五.國外市場排名
         If CheckUse("frm030610", strExec) = True Then
            frm030610.Show
         End If
      Case 5 '表一～表五
         If CheckUse("frm030611", strExec) = True Then
            frm030611.Show
         End If
      Case 6 '國外前十大申請國及其商品類別排名
         If CheckUse("frm030612", strExec) = True Then
            frm030612.Show
         End If
      Case 7 '代理人國外案件排名分析表
         If CheckUse("frm030613", strExec) = True Then
            frm030613.Show
         End If
      Case 8 '本所案件申請人國籍統計表
         If CheckUse("frm030621_1", strExec) = True Then
            frm030621_1.Show
         End If
   End Select
End Sub
'end 2017/12/7

'Added by Lydia 2020/01/09 專利公報統計
'Modified by Lydia 2020/01/13 mnu2410=>mnu2409
Private Sub mnu2409_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '公開及公告市場統計表
         If CheckUse("frm04060108", strExec) = True Then
            frm04060108.Show
         End If
      Case 2   '公報產業分類案件市佔分析
         If CheckUse("frm100133", strExec) = True Then
            frm100133.Show
         End If
     Case Else
   End Select
End Sub

'Added by Lydia 2021/01/06 商標公報資料統計-Excel
Private Sub mnu2413_Click(Index As Integer)
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

'Add By Sindy 2015/3/19
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
            'Added by Morgan 2015/3/19
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

'add by nick 2004/10/06
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
   'Modify by Morgan 2008/8/1 代表圖畫面開啟時不限制
   'If Pub_StrUserSt03 = "M51" Then Exit Sub
   'Modified by Morgan 2014/4/23 +繪圖也不鎖
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "P13" Or Pub_Can_Copy_Pic = True Then Exit Sub
   'S 開頭的部門不能 copy
   'edit by nickc 2005/03/02
   'If Mid(UCase(Pub_StrUserSt03), 1, 1) = "S" Then
       '圖檔才清 edit by nickc 2005/03/02
   '    If Clipboard.GetFormat(2) = True Or Clipboard.GetFormat(3) = True Or Clipboard.GetFormat(9) = True Then
        If Clipboard.GetFormat(2) = True And Clipboard.GetFormat(3) = False And Clipboard.GetFormat(1) = False Then
           Clipboard.Clear
       End If
   'End If
End Sub

'Add by Morgan 2003/12/23
'控制連線閒置超過30分鐘自動離線
Private Sub tmrConnect_Timer()
   tmrConnect.Tag = Val(tmrConnect.Tag) + 1
   'Modify by Morgan 2005/2/3 改成10分鐘--薛副理
   'If tmrConnect.Tag = 30 Then
   'Modified by Morgan 2013/9/25 改回30分鐘--薛經理
   'If tmrConnect.Tag = 10 Then
   If tmrConnect.Tag = 30 Then
      Timer1.Enabled = False
      
'Modify by Morgan 2005/1/10 改保留原畫面不結束
'      Call CloseAllChild
'      Call SwitchMenu(False)
'2005/1/10 end
      
      'Remove by Morgan 2005/2/23 移到重連線視窗
      'cnnConnection.Close
      bolReOpen = False
      frmReopen.Show vbModal, Me
      If bolReOpen = True Then
         Call ReConnect
      Else
         Call mnu00_Click(1)
      End If
   End If
End Sub

Private Sub MDIForm_Load()
   bolReOpen = True 'Add By Sindy 2020/4/10
   
   'Add by Morgan 2003/12/22
   '控制連線閒置超過30分鐘自動關閉程式
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
      Set eventConn = cnnConnection
      tmrConnect.Interval = 60000
   End If
   
   Dim strSysKind As String
   Dim lngValue, lngBufferSize As Long, intCounter As Integer
   Dim strUserId As String * 10, strLocalId As String
    
   '若登入成功
   If pub_str_LoginSucceeded = "1" Then
      'add by nickc 2005/03/11 測試用
      If Pub_StrUserSt03 = "M51" Then
         mnuChUser.Visible = True
         tmrConnect.Interval = 0
      Else
         mnuChUser.Visible = False
         'Modify by Amy 2015/09/04 原:mnu0901(23) 將承辦人拆成專利處及商標處
         'Modified b Lydia 2024/11/07 Index 23=>20
         mnu0701(20).Visible = False 'Added by Lydia 2015/07/21 限電腦中心使用
         'Add by Amy 2019/06/05 只有電腦中心及外商人員可用
         'Modified by Lydia 2019/06/27 調整表單
'         mnu2102(20).Visible = False
'         If Left(Pub_StrUserSt03, 2) = "F1" Then
'             mnu2102(20).Visible = True
'         End If
         'Modified by Lydia 2022/05/18 下一程序接洽單列印: 因為增加CFP領證預估報價 index=9 改為10
         mnu2104(10).Visible = False
         'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
         If Left(Pub_StrUserSt03, 2) = "F1" Or InStr("01,08", Pub_strUserST05) > 0 Then
             mnu2104(10).Visible = True
         End If
         'end 2019/06/27
      End If
      
      'Add By Sindy 2023/4/20
      mnu23(7).Visible = False '系統收件區
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      If Left(Pub_StrUserSt03, 2) = "P2" Or Pub_StrUserSt03 = "M51" Or InStr("01,08", Pub_strUserST05) > 0 Then '商標處
         mnu23(7).Visible = True
      End If
      '2023/4/20 END
      
'      'Added by Lydia 2023/05/05 外專新案認領：控制顯示
'      If Pub_StrUserSt03 = "M51" Or strSrvDate(1) >= 外專新案認領啟用日 Then
'          mnu03(20).Visible = True
'      Else
'          mnu03(20).Visible = False
'      End If
'      'end 2023/05/05
      'Add By Sindy 2023/12/18
      If strSrvDate(1) >= 外專承辦歷程啟用日 Then
         mnu03(4).Visible = True
         mnu03(5).Visible = True
      Else
         mnu03(4).Visible = False
         mnu03(5).Visible = False
      End If
      '2023/12/18 END
      
      'Add By Sindy 2018/5/4
'      If Pub_StrUserSt03 = "M51" Then
'         mnu0702(2).Visible = True
'      Else
'         mnu0702(2).Visible = False
'      End If
      '2018/5/4 END
      
'       'Added by Morgan 2016/1/19 薪資查詢測試
'       If strUserNum = "94007" Or strUserNum = "68009" Then
'         mnu23(6).Visible = True
'       Else
'         mnu23(6).Visible = False
'       End If
'       'end
'
      'Modify 2017/01/25 Add by Amy 2016/12/09 圖書查詢測試
      'If Pub_strUserST05 = "F1" Or Pub_strUserST05 = "F2" Or Pub_StrUserSt03 = "M51" Or strUserNum = "A4023" Then
      If strSrvDate(1) >= 20170202 Then
        mnu23(8).Visible = True
      Else
        mnu23(8).Visible = False
      End If
      'end 2016/12/09
      
      'Add By Sindy 2013/3/7 研發開拓
      'Modified by Lydia 2017/02/08 +總經理 Or Pub_strUserST05 = "01" Or Pub_strUserST05 = "08"
      'Modified by Lydia 2019/04/03 +何主秘同意開放權限給FCP四位工程師主管(42,38)
      'If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "D01" Or Pub_strUserST05 = "01" Or Pub_strUserST05 = "08" Then
      'modify by sonia 2019/10/18 69009楊特助改監察人但仍有研發開拓權限
      'Modified by Lydia 2020/09/08 總經理同意開放權限給55(國外業務拓展)使用
      'If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "D01" Or Pub_strUserST05 = "01" Or Pub_strUserST05 = "08" Or Pub_strUserST05 = "42" Or Pub_strUserST05 = "38" Or strUserNum = "69009" Then
      
      'Add by Sindy 2020/10/29 + Or strUserNum = "A4023" Or strUserNum = "77027" : 開放價目表3支作業
      If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "D01" Or InStr("01,08,42,38,55", Pub_strUserST05) > 0 Or _
         strUserNum = "69009" Or strUserNum = "A4023" Or strUserNum = "77027" Then
         mnuTitle(24).Visible = True
      Else
         mnuTitle(24).Visible = False
      End If
      '2013/3/7 End
      
      'Add By Sindy 2022/9/30
      mnu2303(4).Visible = False '主管待分案
      mnu2101(3).Visible = False '案件接洽單 (電子收文)
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      If Pub_StrUserSt03 = "M51" Or InStr("01,08", Pub_strUserST05) > 0 Then
         mnu2101(2).Caption = "案件接洽單(法務案件)" 'Add By Sindy 2022/12/25
         mnu2101(3).Visible = True '案件接洽單 (電子收文)
         mnu2303(4).Visible = True '主管待分案
      'Added by Lydia 2022/12/16 配合接洽單電子收文
      Else
         If strSrvDate(1) >= 接洽單電子收文啟用日 Then
            'Modify By Sindy 2023/2/17 + Or Left(Pub_StrUserSt03, 1) = "F" '外專外商也是用舊的接洽單輸案源,因不走簽核
            'Add By Sindy 2023/5/12 F1.外商開放掛2支接洽單,因為要收CFT案件
            If Left(Pub_StrUserSt03, 2) = "F1" Then
               mnu2101(2).Caption = "案件接洽單(法務案件)"
               mnu2101(2).Visible = True
               mnu2101(3).Visible = True '案件接洽單 (電子收文)
               bolUpdNew = False
            '2023/5/12 END
            'F單位沒有電子接洽單所以不能用電子收文那一支
            ElseIf Left(Pub_StrUserSt03, 1) = "L" Or Left(Pub_StrUserSt03, 1) = "F" Then '法律所用紙本接洽單
               mnu2101(2).Caption = "案件接洽單(法務案件)" 'Add By Sindy 2022/12/25
               mnu2101(2).Visible = True
               mnu2101(3).Visible = False
               bolUpdNew = False  'Added by Lydia 2023/03/09
            Else
               mnu2101(2).Visible = False
               mnu2101(3).Visible = True '案件接洽單 (電子收文)
               bolUpdNew = True 'Added by Lydia 2023/03/09 強迫使用電子收文顯示
            End If
            mnu2303(4).Visible = True
         End If
      'end 2022/12/16
      End If
      
'      'Add By Sindy 2013/5/16 承辦歷程
'      If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'         mnu2101(14).Visible = True
'         mnu0507(6).Visible = True
'         mnu0508(5).Visible = True
'      Else
'         mnu2101(14).Visible = False
'         mnu0507(6).Visible = False
'         mnu0508(5).Visible = False
'      End If
'      '2013/5/16 End
      
      Me.Timer1.Interval = 100
      
      'Modify By Cheng 2003/07/10
      'Set cnnConnection = objPublicData.Connection
      Systemkind_g = GetSystemKindByNick
      Systemkind_g_P = GetSystemKindByNickP
      strSysKind = Systemkind_g
      'add by nickc 2006/06/09 可以查詢維護紀錄
      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
           mnuDML(0).Visible = True
'           mnu24(7).Visible = True 'Add By Sindy 2016/3/9
      Else
           mnuDML(0).Visible = False
'           mnu24(7).Visible = False 'Add By Sindy 2016/3/9
      End If
      
      If bolFNation = False Then
         mnu101(10).Visible = False 'Modify by Amy 2014/05/05 '原:mnu101(9)
         mnu102(7).Visible = False
         mnu101(6).Visible = False 'Modify by Amy 2014/05/05 '原:mnu101(5)
      End If
       
      'add by nickc 2007/10/02 林淑真、杜副總、電腦中心 才秀的
      'modify by sonia 2014/6/9 +美珍77027
      '2015/7/24 modify by sonia +林總94007改01,何主秘68009,-江總68001-小真65001
      'Modified by Lydia 2020/01/09 + 69005 簡協理
      'Modified by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      If strUserNum = "77027" Or strUserNum = "68006" Or strUserNum = "68009" Or InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0 _
                    Or Pub_StrUserSt03 = "M51" Or InStr("01,08", Pub_strUserST05) > 0 Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         'Modified by Lydia 2019/06/27 調整表單
'         mnu2103(5).Visible = True
'         mnu2103(6).Visible = True
'         mnu2103(7).Visible = True
'         'add by nickc 2008/04/18
'         mnu2103(8).Visible = True
'         'add by nickc 2008/04/28
'         mnu2103(9).Visible = True
         mnu2106(9).Visible = True
         mnu2106(10).Visible = True
         mnu2106(11).Visible = True
         mnu2106(12).Visible = True
         mnu2106(13).Visible = True
         'end 2019/06/27
      'add by nickc 2008/04/28 加入邱素蓮可以的權限
      'modify by sonia 2019/4/12邱素蓮調職改成莊敏惠73017
      'modify by sonia 2019/5/15再改為判斷'北所業務助理人員'
      ElseIf InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) > 0 Then
         'Modified by Lydia 2019/06/27 調整表單
         'mnu2103(9).Visible = True
         mnu2106(13).Visible = True
         'Modified by Lydia 2022/05/10 北所業務助理人員只開放mnu2106(9),mnu2106(13)
         mnu2106(9).Visible = True
         mnu2106(10).Visible = False
         mnu2106(11).Visible = False
         mnu2106(12).Visible = False
         'end 2022/05/10
      Else
         'Modified by Lydia 2019/06/27 調整表單
'         mnu2103(5).Visible = False
'         mnu2103(6).Visible = False
'         mnu2103(7).Visible = False
'         'add by nickc 2008/04/18
'         mnu2103(8).Visible = False
'         'add by nickc 2008/04/28
'         mnu2103(9).Visible = False
         'modify by sonia 2019/7/18 文雄A4023及敏惠73017開放業務目標及達成通知日報表權限
         'mnu2106(9).Visible = False
         'modify by sonia 2020/10/6 + 69009.楊毓純,75033.夏慧珠
         'Modified by Lydia 2021/08/31 +74018杜經理
         'Modified by Lydia 2022/05/10 北所業務助理人員只開放mnu2106(9),mnu2106(13)
         'If strUserNum = "A4023" Or strUserNum = "69009" Or strUserNum = "74018" Or _
            InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) > 0 Then
         'modify by sonia 2022/12/8 69009楊監察人權限改用等級08主任秘書
         'If strUserNum = "A4023" Or strUserNum = "69009" Or strUserNum = "74018" Then
         If strUserNum = "A4023" Or Pub_strUserST05 = "08" Or strUserNum = "74018" Then
            mnu2106(9).Visible = True
         Else
            mnu2106(9).Visible = False
         End If
         'end 2019/7/18
         mnu2106(10).Visible = False
         mnu2106(11).Visible = False
         mnu2106(12).Visible = False
         mnu2106(13).Visible = False
         'end 2019/06/27
      End If
      
      'Add By Sindy 2010/11/4
      '74028.邱素蓮、杜副總、電腦中心 才秀的
      '2015/7/24 modify by sonia +林總94007改01,何主秘68009
      'modify by sonia 2019/4/12邱素蓮調職改成莊敏惠73017
      'modify by sonia 2019/5/15再改為判斷'北所業務助理人員'
      'Modified by Lydia 2020/01/09 + 69005 簡協理
      'Modified by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      If InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) > 0 Or strUserNum = "68006" Or strUserNum = "68009" Or InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0 _
          Or Pub_StrUserSt03 = "M51" Or InStr("01,08", Pub_strUserST05) > 0 Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         'Modified by Lydia 2019/06/27 調整表單
         'mnu2103(11).Visible = True
         mnu2106(14).Visible = True
      Else
         'Modified by Lydia 2019/06/27 調整表單
         'mnu2103(11).Visible = False
         mnu2106(14).Visible = False
      End If
      
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or InStr("01,08", Pub_strUserST05) > 0 Then
         'Modified by Lydia 2019/06/27 調整表單
         'mnu2103(4).Visible = True
         mnu2106(8).Visible = True
      Else
         'Modified by Lydia 2019/06/27 調整表單
         'mnu2103(4).Visible = False
         mnu2106(8).Visible = False
      End If
      
      'Added by Lydia 2022/05/09 客戶特殊紀錄異動: 有執行權限的人才出現此功能選項
      strExc(0) = ""
      strSql = "select distinct decode(st01,null,SR01,st01) from staff_right,staff" & _
                 " where upper(sr02)='FRM010022' and sr08='Y' and sr01=st05(+)" & _
            " and decode(st01,null,SR01,st01)='" & strUserNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then strExc(0) = "Y"
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      If Pub_StrUserSt03 = "M51" Or strExc(0) = "Y" Or InStr("01,08", Pub_strUserST05) > 0 Then
          mnu2106(15).Visible = True
      Else
          mnu2106(15).Visible = False
      End If
      'end 2022/05/09
      
'      'Add By Sindy 2014/5/1 T大陸指示信
'      If strUserNum = "86048" Or Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'         mnu05(5).Visible = True
'      Else
'         mnu05(5).Visible = False
'      End If
      
      'Added by Morgan 2013/9/12
      If Left(Pub_StrUserSt03, 2) = "P1" Then
         mnu23(1).Caption = "會議室/檢索系統預約作業"
      End If
      'end 2013/9/12
        
      'Modfiy by Amy 2015/09/04 將承辦人拆成專利處及商標處,改F2及P1顯示國外部及專利處
      'Add by Morgan 2008/9/3 加控制國外部專用功能
      'Modified by Lydia 2017/02/08 +總經理 Or InStr("01,08", Pub_strUserST05) > 0
      'Modified by Lydia 2024/03/05 +系統特殊設定「協助機械組內專主管」的人員也開啟國外部的功能。
      If Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt15, 1) = "F" Or InStr("01,08", Pub_strUserST05) > 0 Or InStr(Pub_GetSpecMan("協助機械組內專主管") & ";", strUserNum) > 0 Then
        mnuTitle(3).Visible = True '2015/09/04 原:mnuTitle(8)-國外部
      Else
        mnuTitle(3).Visible = False
      End If
      
      'Mark by Amy 2015/10/07 文雄無權限核稿所以開放不限制(只留mnuTitle(5).Visible = True其餘Mark)
      'Modify by Amy 2020/12/08 避免商標處進入,誤按專利處,故修改
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      'Modified by Lydia 2024/03/07 避免外專工程師(尤其是協辦機械組工程師)，國外部不顯示
      'If Pub_StrUserSt03 = "M51" Or Mid(Pub_StrUserSt03, 1, 2) = "F2" Or Mid(Pub_StrUserSt03, 1, 2) = "P1" Or InStr("01,08", Pub_strUserST05) > 0 Or strUserNum = "A4023" Then
      If Pub_StrUserSt03 = "M51" Or Mid(Pub_StrUserSt03, 1, 2) = "P1" Or InStr("01,08", Pub_strUserST05) > 0 Or strUserNum = "A4023" Then
          mnuTitle(5).Visible = True '專利處
      Else
          mnuTitle(5).Visible = False
      End If
      'end 2015/09/04
      
'      'Add By Sindy 2012/10/2 客戶端平台帳號管理作業
'      If InStr(Pub_GetSpecMan("客戶平台系統設定之管理人員"), strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
'         mnu23(5).Visible = True
'      Else
'         mnu23(5).Visible = False
'      End If
      
      'Add by Morgan 2009/4/29
      'Modify by Amy 2015/09/04 原:mnuTitle(9) 將承辦人拆成商標處
      'If Left(Pub_StrUserSt15, 1) = "F" Then
      'Modified by Lydia 2017/02/08 +總經理 Or InStr("01,08", Pub_strUserST05) > 0
      'Modified by Lydia 2021/07/16 江協理98020進入Promoter系統時可以看到「商標處」的大項功能
      If Pub_StrUserSt03 = "M51" Or Mid(Pub_StrUserSt03, 1, 2) = "F1" Or Mid(Pub_StrUserSt03, 1, 2) = "P2" Or _
                  InStr("01,08", Pub_strUserST05) > 0 Or strUserNum = "98020" Then
          mnuTitle(7).Visible = True
      Else
          mnuTitle(7).Visible = False
      End If
      
      'Add By Sindy 2016/2/19 專利處不要看到品名查詢 及 國外開拓
      'Modified by Lydia 2019/07/01 包含收文部門
      'If Mid(Pub_StrUserSt03, 1, 2) = "P1" Then
      If Mid(Pub_StrUserSt03, 1, 2) = "P1" And Mid(Pub_StrUserSt15, 1, 2) = "P1" Then
         mnuTitle(20).Visible = False '品名查詢
         mnuTitle(22).Visible = False '國外開拓
      Else
         mnuTitle(20).Visible = True
         mnuTitle(22).Visible = True
      End If
      '2016/2/19 END
      
      'Added by Lydia 2024/11/07 查名單(網中) ---未上線隱藏
      If Pub_StrUserSt03 <> "M51" Then
         mnu0701(1).Visible = False   '查名區
         mnu0701(2).Visible = False   '查覆區---上線後,移到智權部
         mnu0701(3).Visible = False   '覆核區
         mnu0701(4).Visible = False   '查名單維護
         mnu2102(7).Visible = False 'Added by Lydia 2025/04/30 合併查名／查覆區：(原)查名單、查名單(網中)
      End If
      'end 2024/11/07
      'Added by Lydia 2025/04/14 暫時開放(網中)查名單給人員檢查內容
      strExc(0) = Pub_GetSpecMan("協助檢查網中查名單")
      If InStr(strExc(0) & ";", strUserNum) > 0 Then
         mnu0701(1).Visible = True   '查名區
         mnu0701(2).Visible = True   '查覆區 'Added by Lydia 2025/04/21 開放測試模式
         mnu0701(3).Visible = True   '覆核區
      End If
      'end 2025/04/14
      
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
      'Me.Icon = LoadPicture(strIcoPath)
      ToolHide
      'Systemkind_g_T = GetSystemKindByNickT
      'Systemkind_g_TnoS = GetSystemKindByNickTnoS
   End If
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim frm As Form
    
   'Add By Cheng 2003/01/30
   '使用者從表單上的控制功能表中選取「關閉」指令, 則取消動入
   If UnloadMode = 0 Then
      MsgBox "請按 [系統] --> [結束]，以結束本系統!!!", vbExclamation + vbOKOnly
      Cancel = True
   Else
      '關閉尚未關閉的子視窗
      For Each frm In Forms
          If frm.Name <> mdiMain.Name Then
              Unload frm
          End If
      Next
      
      'Add By Sindy 2020/4/10 人員反應若出現斷線按結束時;人員沒輸入密碼狀況下,還是會因為下列詢問再登入系統,但覺得不安全
      If bolReOpen = True Then
      '2020/4/10 END
         'Add By Sindy 2020/3/20 專利處主管,專利處工程師,專利處繪圖要提醒詢問,但排除王副總因為沒承辦案件
         If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then 'Add By Sindy 2020/3/26 + if
            'Modify By Sindy 2025/7/17 柏翰經理提:這個提醒主要是疫情期間隨時可能居家辦公所以提醒要上傳檔案
            '                                     現在已無這需求,故請移除這個提醒
'            If (Pub_StrUserSt03 = "P10" Or Pub_StrUserSt03 = "P11" Or Pub_StrUserSt03 = "P13") And _
'               strUserNum <> "71011" Then
'               If MsgBox("未完成稿件是否已上傳暫存區？", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption & " 重要訊息！") = vbNo Then
'                  If Pub_StrUserSt03 <> "P13" Then
'                     ProState = "1"
'                     ProSysState = "1"
'                     frm090201_2.Show
'                  End If
'                  Cancel = True
'               End If
'            End If
'            '2020/3/20 END
            'Add By Sindy 2023/1/10 檢查接洽單待收文區是否還有資料,不可結束,防止人員忘記送出
            strSql = "select CRL01 From ConsultRecordList,ConsultRecCMP,flow003" & _
                     " where crl02>=" & 接洽單電子收文啟用日 & " and CRL01=f0301(+) and f0309 is null" & _
                     " and CRL01=crc01(+) and crc02 is not null" & _
                     " and crl78='" & strUserNum & "'" & _
                     " group by CRL01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               MsgBox "接洽單待收文區尚有資料，不可結束系統，請送出！", vbExclamation
               Cancel = True
               If PUB_CheckFormExist("frm090801_New") = False Then
                  frm090801_New.Show
               Else
                  frm090801_New.ZOrder 1
               End If
            End If
            '2023/1/10 END
         End If
      End If
   End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2011/3/9
   PUB_AddAuditLog AL_登出 'Added by Morgan 2025/7/31

'edit by nickc 2007/02/06 不用 dll 了
'Set obj001 = Nothing
'Set objPublicData = Nothing
   ' 90.08.22 modify by louis
   EndOfficeAp
   'Add By Cheng 2002/07/18
   Set mdiMain = Nothing
End Sub
'Modify by Morgan 2005/12/14 加切換連線選擇
Private Sub mnu00_Click(Index As Integer)
   Select Case Index
      Case 0
         If PUB_Connect2DB(True) = False Then
            Unload Me
         End If
      Case 1
         bolUnloading = True 'Add by Morgan 2011/3/11
         Unload Me
   End Select
End Sub

'Modify by Amy 2015/09/04 原:mnu09 將承辦人拆成專利處及商標處
Private Sub mnu05_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 4 '撰寫信函
      frm090401.Show
   'Add by Morgan 2007/12/25
   Case 6 'P案國外指示信
      If CheckUse("frm040106_1P", strExec) = True Then
         'Modified by Morgan 2016/5/19 +程序(跑歷程)
         If Pub_StrUserSt03 = "P12" Then
            frm040106_1.iFrom = 2
         Else
            frm040106_1.iFrom = 1
         End If
         'end 2016/5/19
         frm040106_1.Caption = "P案國外指示信"
         frm040106_1.Show
      End If
   'Add by Morgan 2011/9/22
   Case 7 'P案各式申請書
      If CheckUse("frm04010301_1", strExec) = True Then
         frm04010301_1.Show
      End If
   Case 8 '聯絡單列印及E-Mail    '2011/9/22 加入
      frm1106.Show
   'Modified by Lydia 2015/11/05 新增"主管機關處理記錄"功能表,後續index + 1
   'Added by Morgan 2014/4/28
   Case 12 '公文來函判發作業
      '考慮職代問題,不必鎖權限
      'Modified by Morgan 2015/4/21 改回用權限控制,目前為游經理及王副總可執行
      'Modified by Morgan 2017/4/24 配合非臺灣案有其他判發人再改成可看自己及當時請假之被代理人(patpro於2016/6/17修改)
      'If CheckUse("frm040113", strExec) = True Then
         frm040113.Show
      'End If
      
   'Added by Morgan 2014/12/17
   Case 13 '發後補看作業
      '考慮職代問題,不必鎖權限
      frm040117.m_ProState = "P" 'Add By Sindy 2020/12/7
      frm040117.Show
      
   'Mark by Amy 2018/08/17 結案單審核作業搬至一般作業
'   'Added by Sindy 2015/1/23
'   Case 14 '
'      frm040118.Show
   'end 2015/11/05
   Case Else
   End Select
End Sub

'專利處-承辦人作業
Private Sub mnu0502_Click(Index As Integer)
Dim bolNoCheck As Boolean 'Added by Morgan 2013/10/8
Dim strSysID As String 'Add By Sindy 2014/7/4
Dim nFrm As Form 'Add By Sindy 2018/1/24

'ProState = "1"
'ProSysState = "1"
ToolHide
Select Case Index
   Case 1 '工作進度資料維護
'      'Add By Sindy 2018/1/24
'      '檢查表單是否已開啟，若是，則關閉
'      For Each nFrm In Forms
'         If StrComp(nFrm.Name, "frm090201_2", vbTextCompare) = 0 Then
'            Unload frm090201_2
'         End If
'      Next
'      For Each nFrm In Forms
'         If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'            Unload frm090202_2
'            If strSaveConfirm = True Then frm090202_2.ZOrder: Exit Sub 'Add By Sindy 2020/1/17 有資料要儲存,尚需處理...
'         End If
'      Next
'      '2018/1/24 END
'      If PUB_ChkFormIsClose("frm090201_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
'      If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
      
      If CheckUse("frm090201_4", strExec) Then
         'Add By Sindy 2025/3/3
         If PUB_ChkFormIsClose("frm090201_2", "工作進度資料維護") = False Then
            Exit Sub
         Else
         '2025/3/3 END
            'Added by Morgan 2013/10/8
            If Left(Pub_StrUserSt03, 2) = "P1" Then
               '第2次以上可選擇
               strSql = "select * from executelog where el01='frm090201_a' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  'Modified by Morgan 2016/4/18
                  'If MsgBox("是否執行 當天本所期限案件... 等功能", vbYesNo + vbQuestion + vbDefaultButton1, "功能！") = vbNo Then
                  If MsgBox("是否執行 3個工作天內達本所期限案件... 等功能", vbYesNo + vbQuestion + vbDefaultButton1, "功能！") = vbNo Then
                  'end 2016/4/18
                     bolNoCheck = True
                  End If
               End If
            End If
         End If
         
         'Modify By Sindy 2021/5/3
         If bolNoCheck = True Then
            If PUB_ChkFormIsClose("frm090201_2", "工作進度資料維護") = False Then
               Exit Sub
            Else
               ProState = "1"
               ProSysState = "1"
         '2021/5/3 END
               frm090201_2.Show
            End If
         Else
         'end 2013/10/8
            'Modify By Sindy 2021/5/3
            ProState = "1"
            ProSysState = "1"
            '2021/5/3 END
            '2009/11/12 modify by sonia 改寫法無資料不顯示畫面
            'frm090201_4.Show
            frm090201_4.StrMenu1   '當天本所期限案件資料,無資料時由frm090201_4的nextstep執行下一畫面
            If frm090201_4.TextOk = True Then frm090201_4.Show
            '2009/11/12 end
            
         End If 'Added by Morgan 2013/10/8
      End If
   'Add By Sindy 2013/9/3
   Case 2 '待核判區
      'If CheckUse("frm090202_1", strExec) Then
         frm090202_1.m_ProSysState = "1" '承辦人
         frm090202_1.Show
      'End If
   'Add By Cheng 2003/06/17
   Case 3 '承辦人支援記錄維護
      If CheckUse("frm090623P", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090623.Show
      End If
      
    'Add By Morgan 2003/12/25
   Case 4 '承辦人外出記錄維護
      If CheckUse("frm090626", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090626.Show
      End If
      
    'Add By Cheng 2003/12/16
   Case 5 '承辦人特殊案件記錄維護
      If CheckUse("frm090627P", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090627.Show
      End If
      
   'Add by Morgan 2011/7/27
   Case 6   '承辦人修改記錄維護
      '目前控制只有中所能用
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      If pub_strUserOffice = "2" Or Pub_StrUserSt03 = "M51" Or InStr("01,08", Pub_strUserST05) > 0 Then
         If CheckUse("frm090633P", strExec) Then
            ProState = "1"
            ProSysState = "1"
            frm090633.Show
         End If
      End If
   
   'Add by Morgan 2011/8/1
   Case 7   '承辦人衍生記錄維護
      '目前控制只有中所能用
      'Modified by Lydia 2024/02/23 開放總經理的權限+01,08
      If pub_strUserOffice = "2" Or Pub_StrUserSt03 = "M51" Or InStr("01,08", Pub_strUserST05) > 0 Then
         If CheckUse("frm090634P", strExec) Then
            ProState = "1"
            ProSysState = "1"
            frm090634.Show
         End If
      End If
            
   Case 8 '未齊備、未完稿、未發文查詢
      If CheckUse("frm0906121", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090612.Show
      End If
      
   Case 9 '工作進度資料查詢
      If CheckUse("frm090203_1", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090203_1.Show
      End If
   'Add By Cheng 2002/08/27 CheckUse時於FormName後面加 1,2 區分個人及管理
    'Modify By Cheng 2003/07/30
'   Case 5 '承辦人目標資料查詢
'      If CheckUse("frm0906221", strExec) Then
'         frm090622.Show
'      End If
   Case 10 '承辦人達成情形查詢
      If CheckUse("frm0906081", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090608.Show
      End If
   'add by nickc 2005/03/01  加入個人可查的速度評分
   Case 11
      If CheckUse("frm090624P", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090624.Show
      End If
   
   'Case 8 前面加一功能
   
'Removed by Morgan 2022/1/17 沒在用,刪除 (原來就沒顯示)
'   Case 12 '同仁評分作業
'      If CheckUse("frm090204_1", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'         frm090204_1.Show
'      End If

   'Case 9  前面加一功能
   Case 13 '工作進度資料查詢
      If CheckUse("frm090205_1", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090205_1.Show
      End If
   'add by nickc 2005/03/07 月考核
   Case 17
      If CheckUse("frm090616P", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090616_0.Show
      End If
   'add by nickc 2005/03/07 季考核
   Case 18
      If CheckUse("frm090618P", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090618.Show
      End If
   'add by nickc 2007/08/22 英文核稿查詢
   Case 19
      If CheckUse("frm090218", strExec) Then
         ProState = "1"
         ProSysState = "1"
         frm090218.Show
      End If
   Case Else
End Select
End Sub

'專利處-承辦人作業-專利案件管理
Private Sub mnu050205_Click(Index As Integer)
'ProState = "1"
'ProSysState = "2"
ToolHide
Select Case Index
   '專利案例個人輸入作業
   Case 1
      If CheckUse("frm090206_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090206_1.Show
      End If
   
   '專利案例資料查詢
   Case 2
      If CheckUse("frm090207_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090207_1.Show
      End If

   '專利案例資料彙整
   Case 3
      If CheckUse("frm090217_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090217_1.Show
      End If
   '專利案例資料維護
   Case 4
      If CheckUse("frm090206_2", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090206_2.Show
      End If

   Case Else
   
End Select
End Sub

'專利處-承辦人作業-公報簡訊管理
Private Sub mnu050207_Click(Index As Integer)
'ProState = "1"
'ProSysState = "2"
ToolHide
Select Case Index
   Case 1 '公報簡訊個人輸入作業
      If CheckUse("frm090208_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090208_1.Show
      End If
   Case 2 '公報簡訊資料查詢/列印
      If CheckUse("frm090212_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090212_1.Show
      End If
   Case 3 '公報簡訊資料彙整作業
      If CheckUse("frm090209_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090209_1.Show
      End If
   Case 4 '公報簡訊資料維護
      If CheckUse("frm090210_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090210_1.Show
      End If
   Case 5 '公報簡訊索引資料維護
      If CheckUse("frm090211_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090211_1.Show
      End If
End Select

End Sub

'專利處-承辦人作業-期刊資料管理
Private Sub mnu050208_Click(Index As Integer)
'ProState = "1"
'ProSysState = "2"
ToolHide
Select Case Index
   Case 2 '期刊資料維護
      If CheckUse("frm090213", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090213.Show
      End If
   Case 3 '期刊索引資料維護
      If CheckUse("frm090214", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090214.Show
      End If
   Case 1 '期刊資料查詢/列印
      If CheckUse("frm090215_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090215_1.Show
      End If
   Case Else
End Select

End Sub

'專利處-繪圖人員作業
Private Sub mnu0503_Click(Index As Integer)
Dim bolNoCheck As Boolean 'Added by Morgan 2016/4/18
'ProState = "1"
'ProSysState = "2"
ToolHide
Select Case Index
   Case 1 '工作進度資料維護
        'Modify By Cheng 2003/06/27
'      If CheckUse("frm090711", strExec) Then
'         frm090711.Show
'      End If
      If CheckUse("frm090711_2", strExec) Then
         'Add By Sindy 2025/3/3
         If PUB_ChkFormIsClose("frm090711_2", "工作進度資料維護") = False Then
            Exit Sub
         Else
         '2025/3/3 END
            ProState = "1"
            ProSysState = "2"
            'Modified by Morgan 2016/4/18 比照工程師彈3個工作天內達所限案件
            'frm090711_2.Show
            '第2次以上可選擇
            strSql = "select * from executelog where el01='frm090201_4' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If MsgBox("是否執行 3個工作天內達本所期限案件功能", vbYesNo + vbQuestion + vbDefaultButton1, "功能！") = vbNo Then
                  bolNoCheck = True
               End If
            End If
            If bolNoCheck = True Then
               frm090711_2.Show
            Else
               frm090201_4.m_bolIsDrawer = True
               frm090201_4.StrMenu1
               If frm090201_4.TextOk = True Then frm090201_4.Show
            End If
            'end 2016/4/18
         End If
      End If
   'Add By Sindy 2013/9/3
   Case 2 '待核判區
      'If CheckUse("frm090202_1", strExec) Then
         frm090202_1.m_ProSysState = "2" '繪圖人員
         frm090202_1.Show
      'End If
   'Add By Cheng 2003/06/17
   Case 3 '繪圖人員支援記錄維護
      If CheckUse("frm090623P", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090623.Show
      End If
      
    'Add By Morgan 2003/12/16
   Case 4 '繪圖人員外出記錄維護
      If CheckUse("frm090626", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090626.Show
      End If
      
   Case 5 '未齊備、未完稿查詢
      If CheckUse("frm090705", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090705.Show
      End If
   Case 6 '工作進度資料查詢
      If CheckUse("frm090303_1", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090303_1.Show
      End If
   
'Removed by Morgan 2022/1/17 沒在用,刪除 --翔龍、游經理確認
'   Case 7 '同仁評分作業
'      If CheckUse("frm090204_1", strExec) Then
'        ProState = "1"
'        ProSysState = "2"
'         frm090204_1.Show
'      End If

   Case 8 '工作進度資料列印
      If CheckUse("frm090706", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090706.Show
      End If
    'Add By Cheng 2003/07/30
   Case 9 '繪圖人員達成情形查詢
      If CheckUse("frm0907011", strExec) Then
        ProState = "1"
        ProSysState = "2"
         frm090701.Show
      End If
   'Add by Morgan 2004/1/6
   Case 10   '專利案例資料查詢
      If CheckUse("frm090207_1", strExec) Then
         ProState = "1"
         ProSysState = "2"
         frm090207_1.Show
      End If
   'Add by Morgan 2004/1/6
   Case 12  '期刊資料查詢列印
      If CheckUse("frm090215_1", strExec) Then
         ProState = "1"
         ProSysState = "2"
         frm090215_1.Show
      End If
    'add by nickc 2005/03/01 加入 每週速度查詢
   Case 13
      If CheckUse("frm090624P", strExec) Then
         ProState = "1"
         ProSysState = "2"
         frm090624.Show
      End If
   'add by nickc 2005/03/07 月考核
   Case 14
      If CheckUse("frm090616D", strExec) Then
         ProState = "1"
         ProSysState = "2"
         frm090616_0.Show
      End If
   'add by nickc 2005/03/07 季考核
   Case 15
      If CheckUse("frm090618D", strExec) Then
         ProState = "1"
         ProSysState = "2"
         frm090618.Show
      End If
End Select
End Sub

'專利處-承辦人工作管理-查詢及報表
Private Sub mnu050703_Click(Index As Integer)
ToolHide
Select Case Index
Case 1 '承辦人工作進度資料查詢
   If CheckUse("frm090614", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090614.m_ProState = "P" 'Add By Sindy 2024/2/23
      frm090614.Show
   End If
Case 2 '承辦人達成情形查詢
   If CheckUse("frm0906082", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090608.Show
   End If
Case 3 '承辦人工作量查詢
   If CheckUse("frm090609", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090609.Show
   End If
Case 4 '承辦人每日分案情形查詢
   If CheckUse("frm090610", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090610.Show
   End If
Case 5 '承辦天數統計查詢
   If CheckUse("frm090611", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090611.Show
   End If
Case 6 '未齊備未完稿未發文查詢
   'CheckUse時於FormName後面加 1,2 區分個人及管理
   If CheckUse("frm0906122", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090612.Show
   End If
Case 7 '案件處理時間統計查詢
   If CheckUse("frm090613", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090613.Show
   End If
Case 8 '工程師每週完稿明細
   If CheckUse("frm090625", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090625.Show
   End If
Case 9 '案件逾期及異常查詢 (已發文未輸會稿完成日查詢)
   If CheckUse("frm090628", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090628.Show
   End If
Case 10 '加乘註記修改歷史查詢列印
   If CheckUse("frm090630", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090630.Show
   End If
Case 11 '英文核稿查詢
      If CheckUse("frm090218", strExec) Then
        ProState = "2"
        ProSysState = "1"
         frm090218.Show
      End If
Case 12 '智權人員收文高低標查詢
   If CheckUse("frm090607", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090607.Show
   End If
'Add by Morgan 2010/10/12
Case 13 '預定會稿日異常案件查詢
   If CheckUse("frm090632", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090632.Show
   End If
'Added by Lydia 2023/09/18
Case 14   '支援記錄獎金統計
   If CheckUse("frm090639", strExec) Then
      frm090639.Show
   End If
'Added by Lydia 2023/09/18
Case 15 '待辦案件量統計查詢
   If CheckUse("frm090641", strExec) Then
      frm090641.Show
   End If
'Added by Lydia 2024/12/02
Case 16 '待辦案件量統計查詢
   If CheckUse("frm090642", strExec) Then
      frm090642.Show
   End If
'Added by Morgan 2025/4/22
Case 17 '支援次數統計
   If CheckUse("frm090643", strExec) Then
      frm090643.Show
   End If
   
Case Else
End Select
End Sub

'專利處-承辦人工作管理-人員考核管理
Private Sub mnu050704_Click(Index As Integer)
'''''edit by nickc 2007/12/12 專利處修改
'''''''ProState = "2"
'''''''ProSysState = "1"
''''''ToolHide
''''''Select Case Index
''''''Case 1
''''''   If CheckUse("frm090615", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090615.Show
''''''   End If
''''''Case 2
''''''   'edit by nickc 2006/04/19
''''''   'If CheckUse("frm090616_0", strExec) Then
''''''   If CheckUse("frm090616M", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090616_0.Show
''''''   End If
''''''Case 3
''''''   If CheckUse("frm090617", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090617.Show
''''''   End If
''''''Case 4
''''''   'edit by nickc 2006/0419
''''''   'If CheckUse("frm090618", strExec) Then
''''''   If CheckUse("frm090618M", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090618.Show
''''''   End If
''''''Case 5
''''''   If CheckUse("frm090619", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090619.Show
''''''   End If
''''''Case 6
''''''   If CheckUse("frm090615", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090631.Show
''''''   End If
''''''Case Else
''''''End Select

ToolHide
Select Case Index
Case 1
   If CheckUse("frm090624", strExec) Then '專利處每週速度考核表
    ProState = "2"
    ProSysState = "1"
      frm090624.Show
   End If
Case 2 '月考核
   If CheckUse("frm090616M", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090616_0.Show
   End If
Case 3 '季考核
   If CheckUse("frm090618M", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090618.Show
   End If
Case 4 '工程師每月目標基數設定
   'Modified by Lydia 2019/10/25 權限分開來
   'If CheckUse("frm090615", strExec) Then
   If CheckUse("frm090631", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090631.Show
   End If
Case 5 '個人目標資料維護
   If CheckUse("frm090615", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090615.Show
   End If
Case 6 '獎金輸入作業
   If CheckUse("frm090617", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090617.Show
   End If
Case 7 '獎金明細表
   If CheckUse("frm090619", strExec) Then
    ProState = "2"
    ProSysState = "1"
      frm090619.Show
   End If
Case Else
End Select

End Sub

'Private Sub mnu05070402_Click(Index As Integer)
'ProState = "2"
'ProSysState = "1"
'ToolHide
'Select Case Index
'Case 1
'   If CheckUse("frm090616", strExec) Then
'      frm090616.Show
'   End If
'Case 2
'   If CheckUse("frm090616_3", strExec) Then
'      frm090616_3.Show
'   End If
'Case Else
'End Select
'End Sub

'專利處-承辦人工作管理-基本資料維護
Private Sub mnu050705_Click(Index As Integer)
'''''edit by nickc 2007/12/12 專利處修改
'''''''ProState = "2"
'''''''ProSysState = "1"
''''''ToolHide
''''''Select Case Index
'''''''edit by nickc 2005/03/07 已經不需要
'''''''Case 1
'''''''   If CheckUse("frm090620", strExec) Then
'''''''    ProState = "2"
'''''''    ProSysState = "1"
'''''''      frm090620.Show
'''''''   End If
''''''Case 2
''''''   If CheckUse("frm090621", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090621.Show
''''''   End If
'''''''add by nickc 2005/03/16
''''''Case 3
''''''   If CheckUse("frm090629", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090629.Show
''''''   End If
''''''Case Else
''''''End Select

ToolHide
Select Case Index
Case 1
'承辦人支援記錄維護
   If CheckUse("frm090623M", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090623.Show
   End If
Case 2
'承辦人特殊案件記錄維護
   If CheckUse("frm090627M", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090627.Show
   End If
   
   'Add by Morgan 2011/7/27
   Case 3   '承辦人修改記錄維護
      If CheckUse("frm090633M", strExec) Then
         ProState = "2"
         ProSysState = "1"
         frm090633.Show
      End If
   
   'Add by Morgan 2011/8/1
   Case 4   '承辦人衍生記錄維護
      If CheckUse("frm090634M", strExec) Then
         ProState = "2"
         ProSysState = "1"
         frm090634.Show
      End If
      
Case 5 '國內外案件資料維護
   If CheckUse("frm050106_1", strExec) = True Then
      ProState = "2"
      ProSysState = "1"
      frm050106_1.intWhereToGo = 0
      frm050106_1.Show
   End If
Case 6 '每月目次重編作業
    If CheckUse("frm090606", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090606.Show
   End If
   
'Memo by Morgan 2018/5/23 操作超慢,目前沒用,選單已設不顯示
'Removed by Morgan 2022/1/17 刪除
'Case 7 '承辦人、核稿人對照資料維護
'   If CheckUse("frm090621", strExec) Then
'      ProState = "2"
'      ProSysState = "1"
'      frm090621.Show
'   End If
'end 2022/1/17
   
'Memo by Morgan 2018/5/23 操作超慢,目前沒用選單已設不顯示
Case 8 '特殊加乘註記維護
   If CheckUse("frm090629", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090629.Show
   End If
'英文核稿人欄修改權限設定
Case 9
   If CheckUse("frm090202_6", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090202_6.Show
   End If
'Added by Lydia 2023/09/18
Case 10  '免費修正事由維護
   If CheckUse("frm090640", strExec) Then
      ProState = "2"
      ProSysState = "1"
      frm090640.Show
   End If
Case Else
End Select
End Sub

'專利處-繪圖人員工作管理
Private Sub mnu0508_Click(Index As Integer)
'ProState = "2"
'ProSysState = "2"
ToolHide
Select Case Index
'Add By Cheng 2003/06/17
Case 3 '繪圖人員支援記錄維護
'    ProState = "3"
'    ProSysState = "2"
    If CheckUse("frm090623M", strExec) Then
        ProState = "3"
        ProSysState = "2"
        frm090623.Show
    End If
Case 4 '繪圖分案作業
'    ProState = "3"
'    ProSysState = "2"
    If CheckUse("frm090712", strExec) Then
        ProState = "3"
        ProSysState = "2"
        frm090712.Show
    End If
''Add By Sindy 2013/5/16
'Case 5 '待核判區
'   If CheckUse("frm090202_1", strExec) Then
'      frm090202_1.m_ProSysState = "2" '繪圖人員
'      frm090202_1.Show
'   End If
Case Else
End Select
End Sub

'Modified by Morgan 2012/9/25 調整順序
'專利處-繪圖人員工作管理-查詢及報表
Private Sub mnu050801_Click(Index As Integer)
'ProState = "2"
'ProSysState = "2"
ToolHide
Select Case Index
Case 1
   'modify by sonia 2018/5/3 原未區分會因個人之工作進度資料列印權限而開放主管權限
   'If CheckUse("frm090706", strExec) Then
   If CheckUse("frm0907062", strExec) Then
    ProState = "2"
    ProSysState = "2"
      frm090706.Show
   End If
'Added by Morgan 2012/9/25
Case 2   '繪圖超時案件查詢
   If CheckUse("frm090707", strExec) Then
      ProState = "2"
      ProSysState = "2"
      frm090707.Show
   End If
   
Case 3
    'Modify By Cheng 2003/07/30
    '表單名稱後加一碼以區別個人或管理權限
'   If CheckUse("frm090701", strExec) Then
   If CheckUse("frm0907012", strExec) Then
    ProState = "2"
    ProSysState = "2"
      frm090701.Show
   End If
Case 4
   If CheckUse("frm090702", strExec) Then
    ProState = "2"
    ProSysState = "2"
      frm090702.Show
   End If
Case 5
   If CheckUse("frm090703", strExec) Then
    ProState = "2"
    ProSysState = "2"
      frm090703.Show
   End If
Case 6
   If CheckUse("frm090704", strExec) Then
    ProState = "2"
    ProSysState = "2"
      frm090704.Show
   End If
Case 7
   If CheckUse("frm090705", strExec) Then
    ProState = "2"
    ProSysState = "2"
      frm090705.Show
   End If
Case Else
End Select
End Sub

'專利處-繪圖人員工作管理-人員考核管理
Private Sub mnu050802_Click(Index As Integer)
'ProState = "2"
'ProSysState = "2"
ToolHide
Select Case Index
Case 1 '個人目標資料維護
   If CheckUse("frm090615", strExec) Then
    ProState = "2"
    ProSysState = "2"
      frm090615.Show
   End If
Case 2 '月考核
   'edit by nickc 2005/03/07
   'If CheckUse("frm090709", strExec) Then
   If CheckUse("frm090616DM", strExec) Then
    ProState = "2"
    ProSysState = "2"
      'edit by nickc 2005/03/07
      'frm090709.Show
      frm090616_0.Show
   End If
Case 3 '季考核
   'edit by nickc 2005/03/07
   'If CheckUse("frm090710", strExec) Then
   If CheckUse("frm090618DM", strExec) Then
    ProState = "2"
    ProSysState = "2"
      'edit by nickc 2005/03/07
      'frm090710.Show
      frm090618.Show
   End If
Case Else
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

Public Sub ToolHide()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
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
   Select Case Index
      'Modify By Cheng 2002/10/15
'      Case 16: PrinterLetterDemand
      'Add By Cheng 2002/10/22
      Case 20 '商品查名
         If CheckUse("frm20", strExec) = True Then
            frm20.Show
         End If
      Case Else:
   End Select
End Sub

Private Sub Timer1_Timer()
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
'         mnuTitle(9).Enabled = False
'         mnuTitle(10).Enabled = False
'         mnuTitle(15).Enabled = False
'         mnuTitle(16).Enabled = False
'         'Add By Cheng 2002/10/22
'         mnuTitle(20).Enabled = False
'      End If
'   Else
'      If tmpFormI = 1 Then
'         If mnuTitle(0).Enabled = False Then
'            mnuTitle(0).Enabled = True
'            mnuTitle(9).Enabled = True
'            mnuTitle(10).Enabled = True
'            mnuTitle(15).Enabled = True
'            mnuTitle(16).Enabled = True
'            'Add By Cheng 2002/10/22
'            mnuTitle(20).Enabled = True
'         End If
'      End If
'   End If

'Add By Cheng 2002/11/13
Dim frm As Form
Dim intfrm10 As Integer
'Added by Morgan 2014/7/17
Dim bXForm As Boolean
Dim frmX As Form
'end 2014/7/17

'控制共同查詢
intfrm10 = 0
For Each frm In Forms
    'Modified by Morgan 2014/7/17 +frm100123 除外
    'Modified by Lydia 2019/10/30 +frm100130除外
    If Left(frm.Name, 5) = "frm10" And frm.Name <> "frm100123" And frm.Name <> "frm100130" Then
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
    'add by nickc 2005/08/22
    'If mnu21(9).Enabled = True Then mnu21(9).Enabled = False
    
    'Modified by Morgan 2014/7/17
    'If mnu2102(2).Enabled = True Then mnu2102(2).Enabled = False
    If bXForm Then
      frmX.cmdOK(0).Enabled = False
      frmX.cmdOK(1).Enabled = False
    End If
    'end 2014/7/17
Else
    If mnuTitle(10).Enabled = False Then mnuTitle(10).Enabled = True
    'add by nickc 2005/08/22
    'If mnu21(9).Enabled = False Then mnu21(9).Enabled = True
    'Modified by Morgan 2014/7/17
    'If mnu2102(2).Enabled = False Then mnu2102(2).Enabled = True
    If bXForm Then
      frmX.cmdOK(0).Enabled = True
      frmX.cmdOK(1).Enabled = True
    End If
    'end 2014/7/17
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
        'add by nickc 2007/09/06 加入判斷，以免獨佔視窗會出錯
        If Not mdiMain.ActiveForm Is Nothing Then
            If mdiMain.ActiveForm.Name = Me.mnu99(objMnu99.Index).Tag Then
                If Me.mnu99(objMnu99.Index).Checked = False Then Me.mnu99(objMnu99.Index).Checked = True
            Else
                If Me.mnu99(objMnu99.Index).Checked = True Then Me.mnu99(objMnu99.Index).Checked = False
            End If
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
'Add By Sindy 2015/7/15
Public Sub SetTmpfrm090401()
   Set Tmpfrm090401 = frm090401
End Sub

'Added by Lydia 2015/11/05 主管機關處理記錄
Private Sub mnu0511_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '來電記錄
         If CheckUse("frm04010515", strExec) Then
            frm04010515.Show
         End If
      Case 2 '去電記錄
         If CheckUse("frm04010517", strExec) Then
            frm04010517.Show
         End If
    End Select
End Sub

'Added by Morgan 2016/1/19
'薪資畫面計時器:60秒
Private Sub tmrSalary_Timer()
   tmrSalary.Tag = Val(tmrSalary.Tag) + 1
   If Val(tmrSalary.Tag) > 60 Then
      tmrSalary.Enabled = False
      Pub_CloseSalaryQueryForm
   End If
End Sub
'Added by Lydia 2016/05/06 提示查名人員有到期的查名單
Private Sub CheckTMQ10Alert()
  
   strSql = "select count(*) CNT1 from trademarkquery where tmq11 is null and tmq06<=" & strSrvDate(1) & " and tmq10=" & CNULL(strUserNum)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Val("" & RsTemp(0)) > 0 Then
         If MsgBox("你今天有 " & RsTemp(0) & "張查名單到期／過期，是否進入待查區？", vbInformation + vbYesNo, "到期查名單") = vbYes Then
            'Modified by Lydia 2024/11/13
            'Call mnu0701_Click(25)
            Call mnu0701_Click(18)
         End If
      End If
   End If
   'Added by Lydia 2024/11/13
   If strSrvDate(1) >= 查名單網中系統平行測試 Then
      strSql = "select count(*) cnt from tmqappform where nvl(tma14,'N')='N' and tma11<=" & strSrvDate(1) & " and tma10=" & CNULL(strUserNum)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Val("" & RsTemp(0)) > 0 Then
            If MsgBox("你今天有 " & RsTemp(0) & "張查名單(網中)到期／過期，是否進入待查區？", vbInformation + vbYesNo, "到期查名單(網中)") = vbYes Then
               Call mnu0701_Click(1)
            End If
         End If
      End If
   End If
   'end 2024/11/13
End Sub

'Added by Lydia 2017/01/05 專利公報excel
Private Sub mnu2408_Click(Index As Integer)
   ToolHide
   Select Case Index
       '專利公報市場排名
       Case 1
          If CheckUse("frm04060307", strExec) = True Then
             frm04060307.Show
          End If
       '專利公報市場占有率比較
       Case 2
          If CheckUse("frm04060308", strExec) = True Then
             frm04060308.Show
          End If
       '各單位專利公報件數統計
       Case 3
          If CheckUse("frm04060309", strExec) = True Then
             frm04060309.Show
          End If
       '專利公報國內各區同業排名
       Case 4
          If CheckUse("frm04060310", strExec) = True Then
             frm04060310.Show
          End If
       '專利公報國外同業排名
       Case 5
          If CheckUse("frm04060311", strExec) = True Then
             frm04060311.Show
          End If
       '國籍及洲別統計(含同業)
       Case 6
          If CheckUse("frm04060312", strExec) = True Then
             frm04060312.Show
          End If
   End Select
End Sub

'Added by Lydia 2017/11/15 外專新案未命名區
Private Sub mnu0316_Click(Index As Integer)
   ToolHide
   Select Case Index
       '待分案/待確認
       Case 1
          'Added by Lydia 2024/03/15
          If strUserNum = "89020" Then
              MsgBox "無此使用權限...", , "警告!!"
              Exit Sub
          End If
          'end 2024/03/15
          If CheckUse("frm090902", strExec) = True Then
             frm090902.Show
          End If
       '待命名
       Case 2
          If CheckUse("frm090903", strExec) = True Then
             frm090903.Show
          End If
   End Select
End Sub

'Added by Lydia 2019/06/27 設定個人常用區
'Modify By Sindy 2025/11/7 Private => Public
Public Sub SetPersonMenu()
'2019/06/27 END
Dim intP As Integer, intN As Integer
Dim rsRD As New ADODB.Recordset
Dim strA As String
Dim strTmp As String
   
   'Added by Lydia 2022/12/16 配合接洽單電子收文，修改智權部人員使用表單
   'Modified by Lydia 2023/03/09 外專外商也是用舊的接洽單輸案源,因不走簽核
   'If strSrvDate(1) >= 接洽單電子收文啟用日 And Left(Pub_StrUserSt03, 1) <> "L" Then
   If strSrvDate(1) >= 接洽單電子收文啟用日 And bolUpdNew = True Then
      strA = "select decode(uf01,'frm090801','frm090801_New',uf01) uf01,uf02,uf03,uf04,nvl(fo05,fo02) frmname from useform u1, form f1 " & _
                "where uf02='" & strUserNum & "' and uf03 in ( select nvl(max(uf03),0) as maxdate from useform where uf02='" & strUserNum & "' and uf03<" & strSrvDate(1) & ") " & _
                "and decode(uf01,'frm090801','frm090801_New',uf01)=fo01(+) order by uf04 desc,1 asc "
   Else
   'end 2022/12/16
      strA = "select u1.*,nvl(fo05,fo02) frmname from useform u1, form f1 where uf02='" & strUserNum & "' and uf03 in (" & _
                " select nvl(max(uf03),0) as maxdate from useform where uf02='" & strUserNum & "' and uf03<" & strSrvDate(1) & ")" & _
                " and uf01=fo01(+) order by uf04 desc, uf01 asc "
   End If 'Added by Lydia 2022/12/16
   intP = 1
   PersonDate1 = "0"
   Set rsRD = ClsLawReadRstMsg(intP, strA)
   If intP = 1 Then
        With rsRD
             .MoveFirst
             intP = 1
             PersonDate1 = "" & .Fields("uf03")
             Do While Not .EOF
                 If intP < UBound(PersonList) + 1 Then
                     mnu2001(intP).Caption = "" & .Fields("frmname")
                     PersonList(intP) = "" & .Fields("uf01")
                     intP = intP + 1
                 Else
                     strTmp = strTmp & .Fields("uf01") & ","
                 End If
                 .MoveNext
             Loop
        End With
        strA = "delete from useform where uf02='" & strUserNum & "' and uf03<" & PersonDate1 '刪除前一日之前的記錄
        cnnConnection.Execute strA
        If strTmp <> "" Then '刪除前一日之第X名之後
             strA = "delete from useform where uf02='" & strUserNum & "' and uf03=" & PersonDate1 & " and uf01 in ( " & GetAddStr(strTmp) & ") "
             cnnConnection.Execute strA
        End If
   End If
   '隱藏沒有使用的項目
   If PersonDate1 = "0" Or PersonList(1) = "" Then
       mnu20(1).Visible = False
   ElseIf intP <= UBound(PersonList) Then
       For intN = intP To UBound(PersonList)
          mnu2001(intN).Visible = False
       Next
   End If
   
   Set rsRD = Nothing
   Exit Sub
   
ErrHandlePerson:
   If Err.Number <> 0 Then
        Resume Next
   End If
End Sub

'Added by Lydia 2019/06/27 個人常用區
Private Sub mnu2001_Click(Index As Integer)
'Remove by Lydia 2019/08/01
'Dim myForm As Form
'Dim nName

'Added by Lydia 2019/08/01
Dim bolOK As Boolean
Dim nName As String

   ToolHide
   nName = PersonList(Index)
   
   If nName = "" Then Exit Sub
   'Modify By Sindy 2019/7/19 + frm090127
   'Modified by Lydia 2019/07/29 +先排除frm090801(因為Forms.add會另外開啟同名稱表單)
   'Remove by Lydia 2019/08/01 目前無法解決Forms.add會再產生一個相同表單的問題,只能全部用固定表單呼叫
   'If InStr("frm210145,frm210152,frm090127,frm090801", nName) = 0 Then
   '    If myForm Is Nothing Then Set myForm = Forms.add(nName)
   'End If
   
   Select Case nName '傳入各表單的變數
         'Modified by Lydia 2019/08/01 固定表單呼叫
        '智權部-程序作業
            Case "frm210114" '案件委任契約書=>委任契約書
               Call Pub_AddPersonRec("frm210114")
               frm210114.Show
            Case "frm090801"   '國內案件接洽記錄單=>案件接洽單
               Set frm090801.Tmpfrm090126 = frm090126
               Call Pub_AddPersonRec("frm090801")
               frm090801.Show
            'Added by Lydia 2022/12/16
            Case "frm090801_New"   '案件接洽單(電子收文)
               Set frm090801_New.Tmpfrm090126 = frm090126
               Call Pub_AddPersonRec("frm090801_New")
               frm090801_New.Show
            'end 2022/12/16
            Case "frm210133" '期限資料結案單=>案件結案單
               'Add by Amy 2025/07/28 FCP結案單上線後,不可使用舊結案單
               If Left(Pub_StrUserSt03, 2) = "F2" Then
                  MsgBox "結案單電子化已上線" & vbCrLf & _
                                    "請使用新結案單操作"
                  Exit Sub
               End If
               Call Pub_AddPersonRec("frm210133")
               'Modify by Amy 2025/08/11 +if 若為外商結案單,需先輸案號,再判斷進 國內or國外部 結案單
               If strSrvDate(1) >= FCT結案單電子化啟用日 And Left(Pub_StrUserSt03, 2) = "F1" Then
                  frm210133_2.Show
               Else
                  frm210133.Show
               End If
            Case "frm210139" '銷案／銷帳單
               Call Pub_AddPersonRec("frm210139")
               frm210139.Show
            Case "frm100123"   '智權人員期限資料查詢=>期限資料查詢
               Call Pub_AddPersonRec("frm100123")
                frm100123.Show
            Case "frm210126" ' 案件回覆單列印=>回覆單
               Call Pub_AddPersonRec("frm210126")
               frm210126.Show
            Case "frm210132"   '未列印收據/請款單查詢
               Call Pub_AddPersonRec("frm210132")
                frm210132.Show
            Case "frm12040152"  '接洽記錄單查詢及列印=>接洽記錄單查詢／列印
               'Memo by Lydia 2021/05/18 更名為「自動收文接洽單查詢/列印」
               Call Pub_AddPersonRec("frm12040152")
               frm12040152.Show
            Case "frm210144" '文件寄送確認=>寄發文件
               Call Pub_AddPersonRec("frm210144")
               frm210144.Show
            Case "frm210145" '寄件查詢
               frm210145.intWorkItem = 1
               Call Pub_AddPersonRec("frm210145")
               frm210145.Show
        '智權部-專利商標作業
            Case "frm090202_3"   '待會稿區=>專利／商標會稿
                Call Pub_AddPersonRec("frm090202_3")
                frm090202_3.Show
            Case "frm210115"  '客戶專利案件整理表=> 專利案件彙整表
                  Call Pub_AddPersonRec("frm210115")
                  frm210115.Show
            Case "frm090127" '查覆區 (日常或查詢) => 商標查名／查覆區
                  If frm090127.IsRolePlay("查覆") = True Then
                     SetTmpTMQ
                     Call Pub_AddPersonRec("frm090127")
                     frm090127.Show
                  End If
            'Added by Lydia 2025/04/30
            Case "frm090127_1 '查覆區 (日常或查詢) => 合併查名／查覆區"
                  If frm090127_1.IsRolePlay("查覆") = True Then
                     Call Pub_AddPersonRec("frm090127_1")
                     frm090127_1.Show
                  End If
            'end 2025/04/30
            'Added by Lydia 2024/11/07
            Case "frm090127_New" '查覆區 (日常或查詢) => 商標查名／查覆區
                  If frm090127_New.IsRolePlay("查覆") = True Then
                     Call Pub_AddPersonRec("frm090127_New")
                     frm090127.Show
                  End If
            Case "frm090121" '商標查名報告
               Call Pub_AddPersonRec("frm090121")
               frm090121.Show
            Case "frm210136" '台灣商標案齊備日輸入=>商標案件齊備管制
               Call Pub_AddPersonRec("frm210136")
               frm210136.Show
            Case "frm090638" '商標未發文案件原因註記=>商標未發文原因註記
               frm090638.intPeople = 2
               Call Pub_AddPersonRec("frm090638")
               frm090638.Show
        '智權部-財務作業
            Case "frm210141" '智權人員繳款資料輸入=>繳款輸入=>繳款作業及收據查詢
               If ChkSpecRule = True Then 'Added by Lydia 2024/01/31
                  Call Pub_AddPersonRec("frm210141")
                  frm210141.Show
               End If
            Case "frm210106"  '簽收資料查詢
               If frm210106_1.setNextForm = "" Then
                  frm210106.Show
               Else
                  frm210106_1.setCaller frm210106
                  frm210106_1.Show
               End If
            Case "frm210104" '業績點數查詢
               Call Pub_AddPersonRec("frm210104")
               frm210104.Show
            Case "frm210105" '暫收款查詢
               Call Pub_AddPersonRec("frm210105")
               frm210105.Show
            Case "frm210142" '智權人員繳款資料查詢=>繳款資料查詢
               If ChkSpecRule = True Then 'Added by Lydia 2024/01/31
                  Call Pub_AddPersonRec("frm210142")
                  frm210142.Show
               End If
            Case "frm210137"  '各區業績點數統計=>業績點數統計
               Call Pub_AddPersonRec("frm210137")
               frm210137.Show
            Case "frm210122"  '客戶應收帳款查詢=>應收帳款查詢
               Call Pub_AddPersonRec("frm210122")
               frm210122.Show
            Case "frm210146"  '客戶請款明細表=>請款明細表
               Call Pub_AddPersonRec("frm210146")
               frm210146.Show
            Case "frm210152" '智權點數實績與結餘輸入=>每月點數查詢／輸入
              'Modified by Lydia 2021/07/27 更名為「每月點數結算及查詢」
              'Modify by Amy 2019/10/16 原程式寫至frm210152Limit
              Call Frm210152Limit
        '智權部-查詢資料
            Case "frm210134" '智權部自行管制未發文案件作業=>未發文案件管制
               Call Pub_AddPersonRec("frm210134")
               frm210134.Show
            Case "frm210125" '審查機關來函期限查詢=>來函期限查詢
               Call Pub_AddPersonRec("frm210125")
               frm210125.Show
            Case "frm210124" '專業部定稿報價查詢=>定稿報價查詢
               Call Pub_AddPersonRec("frm210124")
               frm210124.Show
            Case "frm210111"   '新客戶來源分析
               Call Pub_AddPersonRec("frm210111")
               frm210111.Show
            Case "frm210120"   '新舊客戶收款貢獻度分析
               Call Pub_AddPersonRec("frm210120")
               frm210120.Show
            Case "frm210112" '智權人員收/發文量分析
               Call Pub_AddPersonRec("frm210112")
               frm210112.Show
            Case "frm210143" '價目表查詢=>價目表
               Call Pub_AddPersonRec("frm210143")
               frm210143.Show
            Case "frm210151" 'CFP常辦國家年費(延展費)預估報價=>CFP年費預估報價
               Call Pub_AddPersonRec("frm210151")
               frm210151.Show
            'Added by Lydia 2022/05/11
            Case "frm210155" 'CFP領證預估報價
               Call Pub_AddPersonRec("frm210155")
               frm210155.Show
            'end 2022/05/11
            Case "frm210154" '下一程序接洽單列印=>外商和M51才會顯示
               Call Pub_AddPersonRec("frm210154")
               frm210154.Show
        '智權部-其他
            Case "frm210101_1"   '個人客戶資料修改=>客戶資料修改
                  frm210101.setNextForm "frm210101_1"
                  frm210101.Caption = "客戶資料修改-登入"
                  frm210101.Show
            Case "frm210102"   '客戶案件資料維護=>案件資料維護
               frm210101.setNextForm "frm210102"
               frm210101.Caption = "案件資料修改-登入"
               frm210101.Show
            Case "frm210110"   '個人行事曆維護=>行事曆
               Call Pub_AddPersonRec("frm210110")
               frm210110.Show
            Case "frm090401" '撰寫信函作業
               Call Pub_AddPersonRec("frm090401")
               frm090401.Show
            Case "frm1106"  '聯絡單列印及E-Mail=>聯絡單 (Menu改名,但是表單不用改名)
               Call Pub_AddPersonRec("frm1106")
               frm1106.Show
        '智權部-區主管作業
            Case "frm210103" '每日業績點數輸入
               If ChkSpecRule = True Then 'Added by Lydia 2024/01/31
                 Call Pub_AddPersonRec("frm210103")
                 frm210103.Show
               End If
            Case "frm210137" '各區業績點數統計
               Call Pub_AddPersonRec("frm210137")
               frm210137.Show
            Case "Frmacc44r0" '專業達成點數表-秘書(從Patpro搬過來)
               If CheckUse("Frmacc44r0", strExec) Then
                  Call Pub_AddPersonRec("Frmacc44r0")
                  Frmacc44r0.Show
               End If
            Case "frm210113" '各區業務工作報告統計
               If ChkSpecRule = True Then 'Added by Lydia 2024/01/31
                  If CheckUse("frm210113", strExec) Then
                     Call Pub_AddPersonRec("frm210113")
                     frm210113.Show
                  End If
               End If
            Case "frm210117" '各所業務工作報告統計
               If CheckUse("frm210117", strExec) Then
                  Call Pub_AddPersonRec("frm210117")
                  frm210117.Show
               End If
            Case "frm210150" '智權部工作報告-總所
               If CheckUse("frm210150", strExec) Then
                  Call Pub_AddPersonRec("frm210150")
                  frm210150.Show
               End If
            Case "frm210127" '新申請案收文至發文件數日數比較表
               If CheckUse("frm210127", strExec) Then
                  Call Pub_AddPersonRec("frm210127")
                  frm210127.Show
               End If
            Case "frm210118" '客戶案件整理表記錄查詢
               If CheckUse("frm210118", strExec) Then
                  Call Pub_AddPersonRec("frm210118")
                  frm210118.Show
              End If
            Case "frm210107" '業務目標及達成通知日報表
               If CheckUse("frm210107", strExec) Then
                  Call Pub_AddPersonRec("frm210107")
                  frm210107.Show
               End If
            Case "frm210108" '業務目標及達成通知月報表
               If CheckUse("frm210108", strExec) Then
                  Call Pub_AddPersonRec("frm210108")
                  frm210108.Show
               End If
            Case "frm100122_1" '業務收/發文量比較查詢
               If CheckUse("frm100122_1", strExec) Then
                  Call Pub_AddPersonRec("frm100122_1")
                  frm100122_1.Show
               End If
            Case "frm210121" '智權部點數分析表
               If CheckUse("frm210121", strExec) Then
                  Call Pub_AddPersonRec("frm210121")
                  frm210121.Show
               End If
            Case "frm210123" '未收款、未收齊清單列印
               If CheckUse("frm210123", strExec) Then
                  Call Pub_AddPersonRec("frm210123")
                  frm210123.Show
               End If
            Case "frm210135" '業績年度統計表
               If CheckUse("frm210135", strExec) Then
                  Call Pub_AddPersonRec("frm210135")
                  frm210135.Show
               End If
            'Added by Lydia 2022/05/09
            Case "frm010022" '客戶特殊紀錄異動
               If CheckUse("frm010022", strExec) Then
                  Call Pub_AddPersonRec("frm010022")
                  frm010022.Show
               End If
            'end 2022/05/09
        '智權部-國內業務開拓
            Case "frm210128" '國內潛在客戶資料維護
                  Call Pub_AddPersonRec("frm210128")
                  frm210128.Show
            Case "frm210129" '國內往來記錄資料維護
                  Call Pub_AddPersonRec("frm210129")
                  frm210129.Show
            Case "frm210130" '國內潛在客戶資料查詢
                  Call Pub_AddPersonRec("frm210130")
                  frm210130.Show
            Case "frm210131" '國內往來記錄資料查詢
                  Call Pub_AddPersonRec("frm210131")
                  frm210131.Show
            Case "frm020322" '台灣商標公告近三年開拓函
               If CheckUse("frm020322", strExec) = True Then
                  Call Pub_AddPersonRec("frm020322")
                  frm020322.Show
               End If
            Case "frm020323" '台灣商標延展開拓(智慧局)
               If CheckUse("frm020323", strExec) = True Then
                  Call Pub_AddPersonRec("frm020323")
                  frm020323.Show
               End If
            Case "frm210153" '網頁提供國內專利公報資訊
               bolOK = False
               '智權部
               If Left(Pub_StrUserSt15, 1) = "S" Then
                  bolOK = True
               '北所業務助理人員
               ElseIf InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) > 0 Then
                  bolOK = True
               '業務助理 S1,CS,N1,K1,SA
               ElseIf CheckUse("frm210153", strExec) = True Then
                  bolOK = True
               End If
               If bolOK Then
                   Call Pub_AddPersonRec("frm210153")
                   frm210153.Show
               End If
          '例外處理
          Case Else
                 MsgBox "無法呼叫常用設定，請通知電腦中心！", vbCritical, "個人常用區表單未設定"
   End Select
   
   'Remove by Lydia 2019/08/01
   'If InStr("frm210101_1,frm210102,frm210106", nName) = 0 Then '排除-在FormLoad記錄今天使用次數
   '   Call Pub_AddPersonRec(nName)   '記錄今天使用次數
   'End If
   
   ''Add By Sindy 2019/7/19
   'If Not myForm Is Nothing Then
   ''2019/7/19 END
    '  myForm.Show
   'End If
   
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
        MsgBox "個人常用區呼叫失敗(" & nName & "): " & Err.Description
        Resume Next
   End If
End Sub

'Added by Lydia 2019/06/27 智權部-程序作業
Private Sub mnu2101_Click(Index As Integer)
   Select Case Index
      Case 1 '案件委任契約書=>委任契約書
         Call Pub_AddPersonRec("frm210114")
         frm210114.Show
      Case 2 '國內案件接洽記錄單=>案件接洽單
         Set frm090801.Tmpfrm090126 = frm090126
         Call Pub_AddPersonRec("frm090801")
         frm090801.Show
      'Add By Sindy 2022/8/10
      Case 3 '國內案件接洽記錄單=>案件接洽單(自動收文)
         Set frm090801_New.Tmpfrm090126 = frm090126
         Call Pub_AddPersonRec("frm090801_New")
         frm090801_New.Show
      Case 4 '期限資料結案單=>案件結案單
         'Add by Amy 2025/07/28 FCP結案單上線後,不可使用舊結案單
         If Left(Pub_StrUserSt03, 2) = "F2" Then
            MsgBox "結案單電子化已上線" & vbCrLf & _
                              "請使用新結案單操作"
            Exit Sub
         End If
         Call Pub_AddPersonRec("frm210133")
         'Modify by Amy 2025/08/11 +if 若為外商結案單,需先輸案號,再判斷進 國內or國外部 結案單
         If strSrvDate(1) >= FCT結案單電子化啟用日 And Left(Pub_StrUserSt03, 2) = "F1" Then
            frm210133_2.Show
         Else
            frm210133.Show
         End If
      Case 5 '銷案／銷帳單
         Call Pub_AddPersonRec("frm210139")
         frm210139.Show
      Case 6 '智權人員期限資料查詢=>期限資料查詢
         Call Pub_AddPersonRec("frm100123")
          frm100123.Show
      Case 7 '案件回覆單列印=>回覆單
         Call Pub_AddPersonRec("frm210126")
         frm210126.Show
      Case 8 '未列印收據/請款單查詢 'Memo by Lydia 2021/07/27 更名為「未列印收據查詢」
         Call Pub_AddPersonRec("frm210132")
          frm210132.Show
      Case 9 '接洽記錄單查詢及列印=>接洽記錄單查詢／列印
         'Memo by Lydia 2021/05/18 更名為「自動收文接洽單查詢/列印」
         'Modify By Sindy 2023/1/6 更名為「電子收文接洽單查詢」
         Call Pub_AddPersonRec("frm12040152")
         frm12040152.Show
      Case 10 '文件寄送確認=>寄發文件
         Call Pub_AddPersonRec("frm210144")
         frm210144.Show
      Case 11 '寄件查詢
         frm210145.intWorkItem = 1
         Call Pub_AddPersonRec("frm210145")
         frm210145.Show
   End Select
End Sub

'Added by Lydia 2019/06/27 智權部-專利商標作業
Private Sub mnu2102_Click(Index As Integer)
    Select Case Index
        Case 1    '待會稿區=>專利／商標會稿
            Call Pub_AddPersonRec("frm090202_3")
            frm090202_3.Show
        Case 2  '客戶專利案件整理表=> 專利案件彙整表
              Call Pub_AddPersonRec("frm210115")
              frm210115.Show
        Case 3 '查覆區 (日常或查詢) => 商標查名／查覆區
              If frm090127.IsRolePlay("查覆") = True Then
                 SetTmpTMQ
                 Call Pub_AddPersonRec("frm090127")
                 frm090127.Show
              End If
        Case 4 '商標查名報告
           Call Pub_AddPersonRec("frm090121")
           frm090121.Show
        Case 5 '台灣商標案齊備日輸入=>商標案件齊備管制
           'Memo by Lydia 2019/11/06 更名「商標案件齊備管制」=>「台灣商標案件齊備管制」(P.S. 加上避免人員誤解)
           'Memo by Lydia 2022/07/15 更名「台灣商標案件齊備管制」=>「商標著作權案件齊備管制」
           Call Pub_AddPersonRec("frm210136")
           frm210136.Show
        Case 6 '商標未發文案件原因註記=>商標未發文原因註記
           frm090638.intPeople = 2
           Call Pub_AddPersonRec("frm090638")
           frm090638.Show
        'Added by Lydia 2025/04/30 暫時放在這裡
        Case 7  '合併查名／查覆區：(原)查名單、查名單(網中)
         If frm090127_1.IsRolePlay("查覆") = True Then
            Call Pub_AddPersonRec("frm090127_1")
            frm090127_1.Show
         End If
    End Select
End Sub

'Added by Lydia 2019/06/27 智權部-財務作業
'Modified by Lydia 2021/07/27 調整財務系統 'Memo by Lydia 2021/08/27 上線
'Private Sub mnu2103_Click(Index As Integer)
'    Select Case Index
'        Case 1 '智權人員繳款資料輸入=>繳款輸入
'           Call Pub_AddPersonRec("frm210141")
'           frm210141.Show
'        Case 2  '簽收資料查詢
'           If frm210106_1.setNextForm = "" Then
'              frm210106.Show
'           Else
'              frm210106_1.setCaller frm210106
'              frm210106_1.Show
'           End If
'        Case 3 '業績點數查詢
'           Call Pub_AddPersonRec("frm210104")
'           frm210104.Show
'        Case 4 '暫收款查詢
'           Call Pub_AddPersonRec("frm210105")
'           frm210105.Show
'        Case 5 '智權人員繳款資料查詢=>繳款資料查詢
'           Call Pub_AddPersonRec("frm210142")
'           frm210142.Show
'        Case 6 '各區業績點數統計=>業績點數統計
'           Call Pub_AddPersonRec("frm210137")
'           frm210137.Show
'        Case 7  '客戶應收帳款查詢=>應收帳款查詢
'           Call Pub_AddPersonRec("frm210122")
'           frm210122.Show
'        Case 8  '客戶請款明細表=>請款明細表
'           Call Pub_AddPersonRec("frm210146")
'           frm210146.Show
'        Case 9 '智權點數實績與結餘輸入=>每月點數查詢／輸入
'           'Modify by Amy 2019/10/16 原程式寫至frm210152Limit
'           Call Frm210152Limit
'    End Select
'End Sub

Private Sub mnu2103_Click(Index As Integer)
    Select Case Index
        Case 1  '請款作業及應收查詢（包含了原請款明細表frm210146及應收帳款查詢frm210122）
            If PUB_CheckFormExist("frm210122") Then
                MsgBox "請先關閉〔應收帳款查詢〕畫面！"
                Exit Sub
            End If
           Call Pub_AddPersonRec("frm210146")
           frm210146.Show
        Case 2  '繳款作業及收據PDF（包含了原繳款輸入frm210141）
           If ChkSpecRule = True Then 'Added by Lydia 2024/01/31
              Call Pub_AddPersonRec("frm210141")
              frm210141.Show
           End If
        Case 3  '繳款查詢及簽收查詢（包含了原繳款資料查詢frm210142及簽收資料查詢frm210106)
           If ChkSpecRule = True Then 'Added by Lydia 2024/01/31
              Call Pub_AddPersonRec("frm210142")
              frm210142.Show
           End If
        Case 4 ''點數輸入作業及查詢（包含了原業績點數查詢frm210104、業績點數統計frm210137及區主管作業中之”每日業績點輸入frm210103”及”各區業績點數統計frm210137(menu名字不同)”）
            If PUB_CheckFormExist("frm210103") Then
                MsgBox "請先關閉〔每日點數輸入〕畫面！"
                Exit Sub
            End If
            If PUB_CheckFormExist("frm210137") Then
                MsgBox "請先關閉〔業績點數統計〕畫面！"
                Exit Sub
            End If
           Call Pub_AddPersonRec("frm210104")
           frm210104.Show
        Case 5 '每月點數結算及查詢（包含了原每月點數查詢/輸入）
           Call Frm210152Limit
        Case 6 '暫收款查詢（包含了原暫收款查詢）
           Call Pub_AddPersonRec("frm210105")
           frm210105.Show
        Case 7 '未列印收據查詢（包含了原未列印收據/請款單查詢）
           Call Pub_AddPersonRec("frm210132")
           frm210132.Show
    End Select
End Sub
'end 2021/07/27

'Added by Lydia 2019/06/27 智權部-查詢資料
Private Sub mnu2104_Click(Index As Integer)
    Select Case Index
        Case 1 '智權部自行管制未發文案件作業=>未發文案件管制
           Call Pub_AddPersonRec("frm210134")
           frm210134.Show
        Case 2 '審查機關來函期限查詢=>來函期限查詢
           Call Pub_AddPersonRec("frm210125")
           frm210125.Show
        Case 3 '專業部定稿報價查詢=>定稿報價查詢
           Call Pub_AddPersonRec("frm210124")
           frm210124.Show
        Case 4   '新客戶來源分析
           Call Pub_AddPersonRec("frm210111")
           frm210111.Show
        Case 5    '新舊客戶收款貢獻度分析
           Call Pub_AddPersonRec("frm210120")
           frm210120.Show
        Case 6   '智權人員收/發文量分析
           Call Pub_AddPersonRec("frm210112")
           frm210112.Show
        Case 7 '價目表查詢=>價目表
           Call Pub_AddPersonRec("frm210143")
           frm210143.Show
        Case 8 'CFP常辦國家年費(延展費)預估報價=>CFP年費預估報價
                   'Memo by Lydia 2019/11/13 更名為:各國年費預估報價
           Call Pub_AddPersonRec("frm210151")
           frm210151.Show
        'Added by Lydia 2022/05/11
        Case 9 'CFP領證預估報價
           Call Pub_AddPersonRec("frm210155")
           frm210155.Show
        'end 2022/05/11
        'Modified by Lydia 2022/05/11 Case 9 => Case 10 'Memo by Lydia 2022/05/18 變更index需要修改顯示的控制
        Case 10 '下一程序接洽單列印=>外商和M51才會顯示
           Call Pub_AddPersonRec("frm210154")
           frm210154.Show
    End Select
End Sub

'Added by Lydia 2019/06/27 智權部-其他
Private Sub mnu2105_Click(Index As Integer)
    Select Case Index
        Case 1   '個人客戶資料修改=>客戶資料修改
              frm210101.setNextForm "frm210101_1"
              frm210101.Caption = "客戶資料修改-登入"
              frm210101.Show
        Case 2   '客戶案件資料維護=>案件資料修改
           frm210101.setNextForm "frm210102"
           frm210101.Caption = "案件資料修改-登入"
           frm210101.Show
        Case 3   '個人行事曆維護=>行事曆
           Call Pub_AddPersonRec("frm210110")
           frm210110.Show
        Case 4 '撰寫信函作業
           Call Pub_AddPersonRec("frm090401")
           frm090401.Show
        Case 5  '聯絡單列印及E-Mail=>聯絡單 (Menu改名,但是表單不用改名)
           Call Pub_AddPersonRec("frm1106")
           frm1106.Show
    End Select
End Sub

'Added by Lydia 2019/06/27 智權部-區主管作業
Private Sub mnu2106_Click(Index As Integer)
    Select Case Index
          Case 1 '每日業績點數輸入 'Memo by Lydia 2021/08/04 「每日業績點數輸入」更名為「每日點數輸入」
             If ChkSpecRule = True Then 'Added by Lydia 2024/01/31
               'Added by Lydia 2021/08/04
               If PUB_CheckFormExist("frm210103") Then
                   MsgBox "請先關閉〔每日點數輸入〕畫面！"
                   Exit Sub
               End If
               'end 2021/08/04
               Call Pub_AddPersonRec("frm210103")
               frm210103.Show
             End If
          Case 2 '各區業績點數統計  'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909)： 隱藏選單改從frm210104呼叫;
             Call Pub_AddPersonRec("frm210137")
             frm210137.Show
          Case 3 '專業達成點數表-秘書(從Patpro搬過來)
             If CheckUse("Frmacc44r0", strExec) Then
                Call Pub_AddPersonRec("Frmacc44r0")
                Frmacc44r0.Show
             End If
          Case 4 '各區業務工作報告統計
             If ChkSpecRule = True Then 'Added by Lydia 2024/01/31
               If CheckUse("frm210113", strExec) Then
                  Call Pub_AddPersonRec("frm210113")
                  frm210113.Show
               End If
             End If
          Case 5 '各所業務工作報告統計
             If CheckUse("frm210117", strExec) Then
                Call Pub_AddPersonRec("frm210117")
                frm210117.Show
             End If
          Case 6 '智權部工作報告-總所
             If CheckUse("frm210150", strExec) Then
                Call Pub_AddPersonRec("frm210150")
                frm210150.Show
             End If
          Case 7 '新申請案收文至發文件數日數比較表
             If CheckUse("frm210127", strExec) Then
                Call Pub_AddPersonRec("frm210127")
                frm210127.Show
             End If
    '---------------以下程式有特定權限才會顯示
          Case 8 '客戶案件整理表記錄查詢
             If CheckUse("frm210118", strExec) Then
                Call Pub_AddPersonRec("frm210118")
                frm210118.Show
            End If
          Case 9 '業務目標及達成通知日報表 'Memo by Lydia 2021/07/27 更名為「業績達成日報表」
             If CheckUse("frm210107", strExec) Then
                Call Pub_AddPersonRec("frm210107")
                frm210107.Show
             End If
          Case 10 '業務目標及達成通知月報表  'Memo by Lydia 2021/07/27 更名為「業績達成月報表」
             If CheckUse("frm210108", strExec) Then
                Call Pub_AddPersonRec("frm210108")
                frm210108.Show
             End If
          Case 11 '業務收/發文量比較查詢
             If CheckUse("frm100122_1", strExec) Then
                Call Pub_AddPersonRec("frm100122_1")
                frm100122_1.Show
             End If
          Case 12 '智權部點數分析表
             If CheckUse("frm210121", strExec) Then
                Call Pub_AddPersonRec("frm210121")
                frm210121.Show
             End If
          Case 13 '未收款、未收齊清單列印
             If CheckUse("frm210123", strExec) Then
                Call Pub_AddPersonRec("frm210123")
                frm210123.Show
             End If
          Case 14 '業績年度統計表
             If CheckUse("frm210135", strExec) Then
                Call Pub_AddPersonRec("frm210135")
                frm210135.Show
             End If
          'Added by Lydia 2022/05/09
          Case 15 '客戶特殊紀錄異動
             If CheckUse("frm010022", strExec) Then
                Call Pub_AddPersonRec("frm010022")
                frm010022.Show
             End If
    End Select
End Sub

'Mark by Amy 2021/11/11
'Add by Amy 2019/10/16 原程式搬過來寫成模組
Private Sub Frm210152Limit_Old()
'    Dim bolIsAreaAg As Boolean, bolIsRest As Boolean, bolRest1Day As Boolean
'    Dim stA0908 As String, st04 As String
'    Dim stA0908List As String 'Add by Amy 2019/10/16
'    Dim IsAreaMan As Boolean, m_ST05 As String, m_ST15 As String 'Add by Amy 2021/07/14是區主管/等級/部門
'    Dim IsAgentLimit As Boolean 'Add by Amy 2021/07/16 是職代權限
'
'    m_ST05 = PUB_GetST05(strUserNum) 'Add by Amy 2021/07/14 等級
'    'Modify by Amy 2021/08/02 1100802 中三區人員調部門,要輸11007月資料部門會有問題
'    'Modify by Amy 2021/08/04 bug-日期帶錯
'    m_ST15 = GetST15(strUserNum, , Left(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 6))
'    If m_ST15 = MsgText(601) Then m_ST15 = Pub_StrUserSt15 'Add by Amy 2021/07/14 部門
'    'end 2021/08/02
'    st04 = "Y"
'    bolIsAreaAg = IsAreaAgent(strUserNum, True, stA0908, st04)
'
'    'Modify by Amy 2021/07/14 +if 葉易雲(78011) 登入,判斷江郁仁(98020) 請假一整,可以區主管職傳身份登入操作
'    If strUserNum = "78011" Then
'        bolIsRest = CheckIsPersonRest("98020", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day)
'        If bolIsRest = True And bolRest1Day = True Then
'            bolIsAreaAg = True
'            stA0908 = "98020 江郁仁"
'        End If
'    Else
'        If stA0908 <> MsgText(601) Then bolIsRest = CheckIsPersonRest(Left(stA0908, 5), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day)
'    End If
'    'end 2021/07/14
'
'    '業績工作職務代理人且其區主管請假一整天或離職
'    'Memo by Amy 2019/08/02 非S部門區主管職代需區主管請假才可進入 ex:陳怡如/林軒吉/洪琬姿
'    'Modify by Amy 2021/07/16 原與m_ST05 = "00" Or m_ST05 = "01" Or m_ST05 = "08" ...同一個if 判斷,因職代原無權限又不是以區主管登入,權限仍要判斷
'                                                  'ex:杜經理請假 79053 為其職代,但彈詢問訊息按「否」,不可以進入
'    If bolIsAreaAg = True And (st04 = "2" Or (bolIsRest = True And bolRest1Day = True)) Then
'        If MsgBox("要以「 " & Mid(stA0908, 7) & "」職代身份進入嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
'            frm210152.strAreaManNo = Left(stA0908, 5)
'            'Add by Amy 2021/07/14
'            m_ST15 = PUB_GetStaffST15(Left(stA0908, 5), 1)
'            If Left(stA0908, 5) = "98020" Then m_ST15 = "F11"
'            IsAreaMan = True
'            'end 2021/07/14
'            IsAgentLimit = True 'Add by Amy 2021/07/16
'        End If
'    End If
'    'Modify by Amy 2021/07/16 職代以區主管登入
'    If IsAgentLimit = True Then
'        '職代以區主管登入,不需再做下面權限判斷
'    'Modify by Amy 2021/07/14 因外商由江協理98020確認,但由洪琬姿80030輸F4106,May 78011輸F4107,故調整權限
'    ElseIf m_ST05 = "00" Or m_ST05 = "01" Or m_ST05 = "08" Or strUserNum = "71011" Or strUserNum = "98020" Then
'        '電腦中心,財務,總經理,主任秘書(等級08),王副總(78011),江協理(98020)可操作此程式
'        'm_ST15 = Pub_StrUserSt15 'Mark by Amy 2021/08/04 上面已抓
'        If strUserNum = "71011" Then
'            '王副總可看自己和P1001
'            IsAreaMan = True
'        ElseIf strUserNum = "98020" Then
'            '江協理操作外商,以區主管權限輸區主管欄,可看「全區資料」頁籤(可做「主管確認」)
'            m_ST15 = "F11"
'            IsAreaMan = True
'        End If
'    '非S部門輸入人員確認(江協理 98020確認外商資料)
'    ElseIf InStr(Replace(智權點數實績與結餘輸入部門, "S", ""), Left(Pub_StrUserSt15, 1)) > 0 Then
'        'F部門抓st14
'        If Left(Pub_StrUserSt15, 2) = "F1" Then
'            stA0908List = GetDeptList(1, strUserNum)
'            If stA0908List = MsgText(601) Then
'                MsgBox "無權限可操作…", , MsgText(5)
'                Exit Sub
'            ElseIf InStr(stA0908List, ",") > 0 Then
'                MsgBox "多重身份請洽電腦中心", , MsgText(5)
'                Exit Sub
'            End If
'        '82026 為P11 會由此確認
'        Else
'            stA0908List = GetDeptList(2, strUserNum)
'            'Modify by Amy 2021/08/03 張宜萱以W1001登入操作W1001
'            If strUserNum = "W1001" Then
'                stA0908List = m_ST15
'            ElseIf stA0908List = MsgText(601) Then
'                MsgBox "無權限可操作…", , MsgText(5)
'                Exit Sub
'            ElseIf InStr(stA0908List, ",") > 0 Then
'                MsgBox "多重身份請洽電腦中心", , MsgText(5)
'                Exit Sub
'            End If
'        End If
'        m_ST15 = Replace(stA0908List, "'", "")
'        'Modify by Amy 2021/08/03 張宜萱以W1001登入輸入W1001「個人」欄位
'        If Left(Pub_StrUserSt15, 2) = "F1" Or strUserNum = "W1001" Then
'            '外商由洪琬姿 80030 輸F4106,May 78011 輸F4107 輸「個人」欄位,但不可看「全區資料」頁籤(非區主管)
'        Else
'            '外專由王文安輸F4102/04/05,W及P2部門 以區主管輸「個人」欄位
'            IsAreaMan = True
'        End If
'    'Modify  by Amy 2019/10/16 增加文雄(st15=S14)輸客服組資料,原文雄需輸北四區,故彈詢問
'    ElseIf Left(Pub_StrUserSt15, 1) = "S" Then
'        'm_ST15 = Pub_StrUserSt15 'Mark by Amy 2021/08/04 上面已抓
'        stA0908List = GetDeptList(3, strUserNum, IIf(strUserNum = "A4023", " And a0901<>'" & m_ST15 & "'", ""))
'        '使用者為區主管但收文所屬部門與區主管部門不同 ex:文雄收文所屬部門北四區,但為客服組區主管
'        'Modify by Amy 2019/11/21 +And strUserNum <> "82026" 排除 柄佑
'        'Modify by Amy 2021/07/14 拿掉 2019/11/21柄佑 82026 判斷,改由非S部門判斷
'        If stA0908List <> MsgText(601) Then
'            If InStr(stA0908List, ",") > 0 Then
'                'Modify by Amy 2021/10/07 客服組區主管改為魏經理(75007),為S11/W10 區主管,需選擇部門進入
'                'MsgBox "多重身份請洽電腦中心", , MsgText(5)
'                'Exit Sub
'                strPublicTemp = stA0908List
'                frm210152_1.Show vbModal
'                If strPublicTemp = MsgText(601) Then
'                    Exit Sub
'                Else
'                    m_ST15 = strPublicTemp
'                    IsAreaMan = True
'                    strPublicTemp = ""
'                End If
'                'end 2021/10/07
'            ElseIf InStr(stA0908List, m_ST15) = 0 Then
'                'ST15非區主管的區 ex:文雄
'                If MsgBox("要以「 " & GetDepartmentName(Replace(stA0908List, "'", "")) & "」部門主管身份進入嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
'                    m_ST15 = Replace(stA0908List, "'", "")
'                    IsAreaMan = True
'                End If
'            ElseIf InStr(stA0908List, m_ST15) > 0 Then
'                '區主管
'                IsAreaMan = True
'            End If
'        End If
'    Else
'        MsgBox "無權限可操作…", , MsgText(5)
'        Exit Sub
'    End If
'    'end 2019/10/16
'    'end 2021/07/14
'    Call Pub_AddPersonRec("frm210152")
'    'Add by Amy 2021/07/14
'    frm210152.stST15 = m_ST15
'    frm210152.bolAreaMan = IsAreaMan
'    'end 2021/07/14
'    'Modified by Lydia 2021/07/27 「每月點數查詢／輸入」更名為「每月點數結算及查詢」
'    frm210152.Caption = "每月點數結算及查詢" '與Account不同
'    frm210152.Show
End Sub

'Add by Amy 2021/11/11 判斷流程改變,故重修改 Fucntion(舊判斷看Frm210152Limit_OLD)
Private Sub Frm210152Limit()
    Dim bolIsAreaAg As Boolean, bolIsRest As Boolean, bolRest1Day As Boolean
    Dim stA0908 As String, st04 As String, stA0908Dept As String, stA0908List As String '登入者之區主管/登入者之區主管是離職/登入者職代之部門/區主管管理區list
    Dim m_ST05 As String, m_ST15 As String, m_UseEmp As String '登入者之等級/登入者之部門/操作員編
    Dim IsAreaMan As Boolean, IsAgentLimit As Boolean '是區主管/是職代權限
    Dim intCnt As Integer, stTP(2) As String
    
    m_ST05 = PUB_GetST05(strUserNum) '登入者 等級
    '以點數數入部門為主, 沒才抓st15,因 1100802 中三區人員調部門,要輸11007月資料部門會有問題
    m_ST15 = GetST15(strUserNum, , Left(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 6))
    If m_ST15 = MsgText(601) Then m_ST15 = Pub_StrUserSt15 '登入者 部門
    
    st04 = "Y"
    '判斷登入者是否為區主管職代
    bolIsAreaAg = IsAreaAgent(strUserNum, True, stA0908, st04, stA0908Dept)
    '為區主管職代->判斷區主管是否請假
    If stA0908 <> MsgText(601) Then
        bolIsRest = CheckIsPersonRest(Left(stA0908, 5), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day)
    End If
    
    '為每月點數輸入工作職務代理人且其區主管請假一整天或離職 (ex:杜經理請假 79053 為其職代,但彈詢問訊息按「否」,不可以進入)
    If bolIsAreaAg = True And (st04 = "2" Or (bolIsRest = True And bolRest1Day = True)) Then
        If MsgBox("要以「 " & Mid(stA0908, 7) & "」職代身份進入嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            frm210152.strAreaManNo = Left(stA0908, 5)
            m_ST15 = stA0908Dept
            IsAreaMan = True '是區主管
            IsAgentLimit = True '是區主管職代
        End If
    End If
    
    '職代以區主管登入
    If IsAgentLimit = True Then
        '職代以區主管登入,不需再做下面權限判斷
    '電腦中心,財務,總經理(等級01),主任秘書(等級08),王副總(71011)可操作此程式
    ElseIf m_ST05 = "00" Or m_ST05 = "01" Or m_ST05 = "08" Or strUserNum = "71011" Then
        If strUserNum = "71011" Then
            '王副總可看自己和P1001
            IsAreaMan = True
            m_UseEmp = strUserNum & ";P1001"
        End If
    '登入者 非S 部門(都以 ST14 判斷)
    ElseIf Left(Pub_StrUserSt15, 1) <> "S" Then
        '抓取 登入者=st14且為智權點數實績與結餘特殊員編(Memo 張宜萱 以W1001登入操作W1001)
        If GetInputPointST14Data(1, strUserNum, intCnt, stTP(1), stTP(2), True) = True Then
            If intCnt > 1 Then
               'Modify by Amy 2023/08/07 +W3001讓A5024輸
'                MsgBox "多重身份請洽電腦中心", , MsgText(5)
'                Exit Sub
                strPublicTemp = "'" & Replace(stTP(1), ";", "','") & "'"
                frm210152_1.Show vbModal
                If strPublicTemp = MsgText(601) Then
                     Exit Sub
                Else
                     m_ST15 = strPublicTemp
                     IsAreaMan = True
                     strPublicTemp = ""
                End If
                'end 2023/08/07
            Else
                stTP(0) = GetDeptMan(stTP(1), 1)
                '登入人員若為區主管
                If stTP(0) <> MsgText(601) And strUserNum = stTP(0) Then
                    IsAreaMan = True
                End If
                m_ST15 = stTP(1)
                m_UseEmp = stTP(2) '操作的編號
            End If
        '部門區主管 ex:江協理
        'Modify by Amy 2022/12/02 +顏裕洋部門F23 確認 F4104/F4105部門F21(王文安協理退休)
        ElseIf GetInputPointData(1, strUserNum, stA0908List, intCnt) = True And intCnt >= 1 Then
            If intCnt = 1 Then
                m_ST15 = stA0908List
                IsAreaMan = True
            Else
        'end 2022/12/02
                strPublicTemp = "'" & Replace(stA0908List, ";", "','") & "'"
                frm210152_1.Show vbModal
                If strPublicTemp = MsgText(601) Then
                    Exit Sub
                Else
                    m_ST15 = strPublicTemp
                    Call GetInputPointST14Data(2, m_ST15, , , stTP(2), True)
                    m_UseEmp = stTP(2) '操作的編號
                    IsAreaMan = True
                    strPublicTemp = ""
                End If
            End If
        Else
            MsgBox "無權限可操作…", , MsgText(5)
            Exit Sub
        End If
    '登入者為 S 部門
    ElseIf Left(Pub_StrUserSt15, 1) = "S" Then
        stA0908List = GetDeptList(3, strUserNum)
        '多重區主管身份
        If InStr(stA0908List, ",") > 0 Then
            '客服組區主管改為魏經理(75007),為S11/W10 區主管,需選擇部門進入
            strPublicTemp = stA0908List
            frm210152_1.Show vbModal
            If strPublicTemp = MsgText(601) Then
                Exit Sub
            Else
                m_ST15 = strPublicTemp
                IsAreaMan = True
                strPublicTemp = ""
            End If
        '一般區主管
        ElseIf Replace(stA0908List, "'", "") = m_ST15 Then
            IsAreaMan = True
        End If
    Else
        MsgBox "無權限可操作…", , MsgText(5)
        Exit Sub
    End If

    Call Pub_AddPersonRec("frm210152")
    frm210152.stST15 = m_ST15
    frm210152.bolAreaMan = IsAreaMan
    frm210152.IsAgentLimit = IsAgentLimit 'Add by Amy 2023/02/02 職代
    If m_UseEmp <> MsgText(601) Then
        frm210152.strInputEmp = m_UseEmp
    End If
    frm210152.Caption = "每月點數結算及查詢" '與Account不同
    frm210152.Show
End Sub

'Added by Morgan 2020/2/19
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
   'Add By Sindy 2025/8/1
   Case "frm090801_New"
         Set GetForm = frm090801_New
         '202025/8/1 END
   'Add by Amy 2024/01/22
   Case "frm090801_14" '接洽單-對造
         Set GetForm = frm090801_14
   'Add By Sindy 2020/5/29
   Case "frm180301"
         Set GetForm = frm180301
   'Added by Morgan 2020/12/29
   Case "frm090401_1"
         Set GetForm = frm090401_1
   'Added by Lydia 2021/04/13
   Case "frm090126"
         Set GetForm = frm090126
   'Added by Lydia 2021/05/10
   Case "frm100123_2"
         Set GetForm = frm100123_2
   'end 2021/05/10
   'Added by Sindy 2022/12/14
   Case "frm210146"
         Set GetForm = frm210146
   'Added by Morgan 2023/5/25
   Case "frm04010304_1"
      Set GetForm = frm04010304_1
   'Modify By Sindy 2023/6/20
   Case "frm090201_5"
      Set GetForm = frm090201_5
   Case "frm090711_2"
      Set GetForm = frm090711_2
   Case "frm090201_b"
      Set GetForm = frm090201_b
   Case "frm090201_2"
      Set GetForm = frm090201_2
   Case "frm090201_d"
      Set GetForm = frm090201_d
   Case "frm090909"
      Set GetForm = frm090909
   '2023/6/20 END
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
   'Added by Lydia 2024/07/17 內商【查名單(TradeMarkQurey)】、【查名單-網中(TMQAPPForm)】
   Case "frm090131"  '圖形路徑輸入
         Set GetForm = frm090131
   Case "frm090132"  '3519組群輸入
         Set GetForm = frm090132
   'end 2024/07/09
   'Added by Lydia 2024/09/24 查名單-網中
   Case "frm090126_New"  '查名單輸入
         Set GetForm = frm090126_New
   Case "frm090127_New"  '查名/查覆區 or 覆核區
         Set GetForm = frm090127_New
   Case "frm090128_New"  '查名單明細作業
         Set GetForm = frm090128_New
   Case "frm090129"  '查名單-->顯示圖檔
         Set GetForm = frm090129
   'end 2024/09/24
   'Added by Lydia 2025/04/30
   Case "frm090127_1"  '合併查名/查覆區 or 覆核區
         Set GetForm = frm090127_1
   Case "frm090128"  '(原)查名單明細作業
         Set GetForm = frm090128
   'end 2025/04/30
   'Add by Amy 2025/08/11
   Case "frm210133" '結案單
      Set GetForm = frm210133
   Case "frm210133_F" '國外部結案單
      Set GetForm = frm210133_F
   Case "frm210133_2" '案號輸入
      Set GetForm = frm210133_2
   Case "frm210133_INV"
      Set GetForm = frm210133_INV
   'end 2025/08/11
   End Select
End Function

'Added by Morgan 2021/4/22
'複製貼上彈跳視窗
Public Sub PopupMenu2(oTextBox As Control)
   Set oControl = oTextBox 'Added by Morgan 2022/1/22
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
      
   'Added by Morgan 2022/1/22
   Case 4 '輸入
      frm880023.SetTextBox oControl
      frm880023.Show vbModal
   End Select
End Sub

'Added by Lydia 2024/01/31 登入特殊控制
Private Function ChkSpecRule() As Boolean
   
   ChkSpecRule = True
   
   '即日起不可使用
   If strUserNum = "75007" And strSrvDate(1) >= "20240202" Then
       MsgBox "無此使用權限...", , "警告!!"
      ChkSpecRule = False
   End If
   
End Function

