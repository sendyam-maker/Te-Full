VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000018&
   Caption         =   "法務系統"
   ClientHeight    =   5540
   ClientLeft      =   4330
   ClientTop       =   2450
   ClientWidth     =   10430
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '最大化
   Begin VB.Timer tmrSalary 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1395
      Top             =   3600
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   150
      Top             =   2250
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   900
      Top             =   1425
   End
   Begin VB.Timer tmrConnect 
      Left            =   900
      Top             =   1815
   End
   Begin VB.Timer Timer2 
      Left            =   150
      Top             =   1830
   End
   Begin VB.Timer Timer1 
      Left            =   150
      Top             =   1425
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   280
      Left            =   0
      TabIndex        =   1
      Top             =   5260
      Width           =   10430
      _ExtentX        =   18397
      _ExtentY        =   494
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
      Height          =   560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10430
      _ExtentX        =   18397
      _ExtentY        =   988
      ButtonWidth     =   494
      ButtonHeight    =   882
      Appearance      =   1
      _Version        =   393216
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
      Caption         =   "法務"
      Index           =   4
      Begin VB.Menu mnuTitle3 
         Caption         =   "內法"
         Index           =   1
         Begin VB.Menu mnu07 
            Caption         =   "資料處理"
            Index           =   1
            Begin VB.Menu mnu0701 
               Caption         =   "分案"
               Index           =   1
            End
            Begin VB.Menu mnu0701 
               Caption         =   "會稿日輸入"
               Index           =   2
            End
            Begin VB.Menu mnu0701 
               Caption         =   "發文"
               Index           =   3
            End
            Begin VB.Menu mnu0701 
               Caption         =   "機關來函"
               Index           =   4
               Begin VB.Menu mnu070104 
                  Caption         =   "法院文書"
                  Index           =   1
               End
               Begin VB.Menu mnu070104 
                  Caption         =   "回執"
                  Index           =   2
               End
               Begin VB.Menu mnu070104 
                  Caption         =   "開庭通知"
                  Index           =   3
               End
               Begin VB.Menu mnu070104 
                  Caption         =   "其他來函"
                  Index           =   4
               End
               Begin VB.Menu mnu070104 
                  Caption         =   "信件退回"
                  Index           =   5
               End
            End
            Begin VB.Menu mnu0701 
               Caption         =   "顧問案件電話諮詢"
               Index           =   5
            End
            Begin VB.Menu mnu0701 
               Caption         =   "合約基本資料維護"
               Index           =   6
            End
            Begin VB.Menu mnu0701 
               Caption         =   "律師執業登錄地區維護"
               Index           =   7
            End
            Begin VB.Menu mnu0701 
               Caption         =   "出庭費確認維護"
               Index           =   8
            End
         End
         Begin VB.Menu mnu07 
            Caption         =   "查詢作業"
            Index           =   2
            Begin VB.Menu mnu0702 
               Caption         =   "庭期資料查詢"
               Index           =   1
            End
            Begin VB.Menu mnu0702 
               Caption         =   "承辦/協辦人員案件查詢"
               Index           =   2
            End
            Begin VB.Menu mnu0702 
               Caption         =   "代理人案件性質統計"
               Index           =   3
            End
            Begin VB.Menu mnu0702 
               Caption         =   "法務處期限通知"
               Index           =   4
            End
            Begin VB.Menu mnu0702 
               Caption         =   "出庭費查詢"
               Index           =   5
            End
         End
         Begin VB.Menu mnu07 
            Caption         =   "報表列印"
            Index           =   3
            Begin VB.Menu mnu0703 
               Caption         =   "期限管制表"
               Index           =   1
            End
            Begin VB.Menu mnu0703 
               Caption         =   "收文未發文明細表"
               Index           =   2
            End
            Begin VB.Menu mnu0703 
               Caption         =   "催審表"
               Index           =   3
            End
            Begin VB.Menu mnu0703 
               Caption         =   "收/發文明細表"
               Index           =   4
            End
            Begin VB.Menu mnu0703 
               Caption         =   "人員工作分析表"
               Index           =   5
            End
            Begin VB.Menu mnu0703 
               Caption         =   "收款業績比較表"
               Index           =   6
            End
            Begin VB.Menu mnu0703 
               Caption         =   "業績分析表"
               Index           =   7
            End
            Begin VB.Menu mnu0703 
               Caption         =   "客戶案件總簿"
               Index           =   8
            End
            Begin VB.Menu mnu0703 
               Caption         =   "顧問到期明細表"
               Index           =   9
            End
            Begin VB.Menu mnu0703 
               Caption         =   "顧問客戶資料表"
               Index           =   10
            End
            Begin VB.Menu mnu0703 
               Caption         =   "顧問地址條"
               Index           =   11
            End
            Begin VB.Menu mnu0703 
               Caption         =   "地址條列印"
               Index           =   12
            End
         End
         Begin VB.Menu mnu07 
            Caption         =   "統計報表"
            Index           =   4
            Begin VB.Menu mnu0704 
               Caption         =   "收/發文統計表"
               Index           =   1
            End
            Begin VB.Menu mnu0704 
               Caption         =   "法務委任案件統計表"
               Index           =   2
            End
            Begin VB.Menu mnu0704 
               Caption         =   "法務案件年度統計表"
               Index           =   3
            End
            Begin VB.Menu mnu0704 
               Caption         =   "逾期未結案統計表"
               Index           =   4
            End
         End
         Begin VB.Menu mnu07 
            Caption         =   "檔案維護"
            Index           =   5
            Begin VB.Menu mnu0705 
               Caption         =   "法務案件基本資料維護"
               Index           =   1
            End
            Begin VB.Menu mnu0705 
               Caption         =   "顧問案件基本資料維護"
               Index           =   2
            End
            Begin VB.Menu mnu0705 
               Caption         =   "案件進度資料維護"
               Index           =   3
            End
            Begin VB.Menu mnu0705 
               Caption         =   "下一程序資料維護"
               Index           =   4
            End
            Begin VB.Menu mnu0705 
               Caption         =   "機關單位資料維護"
               Index           =   5
            End
            Begin VB.Menu mnu0705 
               Caption         =   "庭期資料維護"
               Index           =   6
            End
            Begin VB.Menu mnu0705 
               Caption         =   "案件國家收費表維護"
               Index           =   7
            End
            Begin VB.Menu mnu0705 
               Caption         =   "更換FC代理人作業"
               Index           =   8
            End
            Begin VB.Menu mnu0705 
               Caption         =   "法務工作點數分配"
               Index           =   9
               Visible         =   0   'False
            End
         End
      End
      Begin VB.Menu mnuTitle3 
         Caption         =   "外法"
         Index           =   2
         Begin VB.Menu mnu08 
            Caption         =   "資料處理"
            Index           =   1
            Begin VB.Menu mnu0801 
               Caption         =   "分案"
               Index           =   1
            End
            Begin VB.Menu mnu0801 
               Caption         =   "會稿日輸入"
               Index           =   2
            End
            Begin VB.Menu mnu0801 
               Caption         =   "發文"
               Index           =   3
            End
            Begin VB.Menu mnu0801 
               Caption         =   "機關來函"
               Index           =   4
               Begin VB.Menu mnu080103 
                  Caption         =   "法院文書"
                  Index           =   1
               End
               Begin VB.Menu mnu080103 
                  Caption         =   "回執"
                  Index           =   2
               End
               Begin VB.Menu mnu080103 
                  Caption         =   "開庭通知"
                  Index           =   3
               End
               Begin VB.Menu mnu080103 
                  Caption         =   "其他來函"
                  Index           =   4
               End
               Begin VB.Menu mnu080103 
                  Caption         =   "信件退回"
                  Index           =   5
               End
            End
            Begin VB.Menu mnu0801 
               Caption         =   "代理人來函"
               Index           =   5
               Begin VB.Menu mnu080105 
                  Caption         =   "已收達/已提申"
                  Index           =   1
               End
               Begin VB.Menu mnu080105 
                  Caption         =   "其他來函"
                  Index           =   2
               End
            End
            Begin VB.Menu mnu0801 
               Caption         =   "開拓客戶資料維護"
               Index           =   6
            End
         End
         Begin VB.Menu mnu08 
            Caption         =   "查詢作業"
            Index           =   2
            Begin VB.Menu mnu0802 
               Caption         =   "庭期資料查詢"
               Index           =   1
            End
            Begin VB.Menu mnu0802 
               Caption         =   "承辦/協辦人員案件查詢"
               Index           =   2
            End
            Begin VB.Menu mnu0802 
               Caption         =   "代理人新案案件查詢"
               Index           =   3
            End
            Begin VB.Menu mnu0802 
               Caption         =   "未請款明細查詢"
               Index           =   5
            End
            Begin VB.Menu mnu0802 
               Caption         =   "FC收款請款點數查詢"
               Index           =   6
            End
            Begin VB.Menu mnu0802 
               Caption         =   "開拓客戶資料查詢"
               Index           =   7
            End
            Begin VB.Menu mnu0802 
               Caption         =   "代理人案件性質統計"
               Index           =   8
            End
            Begin VB.Menu mnu0802 
               Caption         =   "法務處期限通知"
               Index           =   9
            End
            Begin VB.Menu mnu0802 
               Caption         =   "員工查詢印表記錄資料查詢"
               Index           =   10
            End
            Begin VB.Menu mnu0802 
               Caption         =   "國外業務帳款查詢"
               Index           =   11
            End
         End
         Begin VB.Menu mnu08 
            Caption         =   "報表列印"
            Index           =   3
            Begin VB.Menu mnu0803 
               Caption         =   "期限管制表"
               Index           =   1
            End
            Begin VB.Menu mnu0803 
               Caption         =   "代理人案件收達/提申管制表"
               Index           =   2
            End
            Begin VB.Menu mnu0803 
               Caption         =   "收文未發文明細表"
               Index           =   3
            End
            Begin VB.Menu mnu0803 
               Caption         =   "催審表"
               Index           =   4
            End
            Begin VB.Menu mnu0803 
               Caption         =   "智權人員收文明細表"
               Index           =   5
            End
            Begin VB.Menu mnu0803 
               Caption         =   "收/發文明細表"
               Index           =   6
            End
            Begin VB.Menu mnu0803 
               Caption         =   "請款點數明細表"
               Index           =   7
            End
            Begin VB.Menu mnu0803 
               Caption         =   "法務人員工作分析表"
               Index           =   8
            End
            Begin VB.Menu mnu0803 
               Caption         =   "收款業績比較表"
               Index           =   9
            End
            Begin VB.Menu mnu0803 
               Caption         =   "業績分析表"
               Index           =   10
            End
            Begin VB.Menu mnu0803 
               Caption         =   "代理人案件總簿"
               Index           =   11
            End
            Begin VB.Menu mnu0803 
               Caption         =   "客戶案件總簿"
               Index           =   12
            End
            Begin VB.Menu mnu0803 
               Caption         =   "代理人/申請人名單"
               Index           =   13
            End
            Begin VB.Menu mnu0803 
               Caption         =   "地址條列印"
               Index           =   14
            End
            Begin VB.Menu mnu0803 
               Caption         =   "開拓客戶地址條"
               Index           =   15
            End
            Begin VB.Menu mnu0803 
               Caption         =   "DHL列印"
               Index           =   16
            End
         End
         Begin VB.Menu mnu08 
            Caption         =   "統計報表"
            Index           =   4
            Begin VB.Menu mnu0804 
               Caption         =   "收/發文統計表"
               Index           =   1
            End
            Begin VB.Menu mnu0804 
               Caption         =   "法務委任案統計表"
               Index           =   2
            End
            Begin VB.Menu mnu0804 
               Caption         =   "法務案件年度統計表"
               Index           =   3
            End
            Begin VB.Menu mnu0804 
               Caption         =   "逾期未結案統計表"
               Index           =   4
            End
         End
         Begin VB.Menu mnu08 
            Caption         =   "檔案維護"
            Index           =   5
            Begin VB.Menu mnu0805 
               Caption         =   "法務案件基本資料維護"
               Index           =   1
            End
            Begin VB.Menu mnu0805 
               Caption         =   "案件進度資料維護"
               Index           =   2
            End
            Begin VB.Menu mnu0805 
               Caption         =   "下一程序資料維護"
               Index           =   3
            End
            Begin VB.Menu mnu0805 
               Caption         =   "機關單位維護"
               Index           =   4
            End
            Begin VB.Menu mnu0805 
               Caption         =   "庭期資料維護"
               Index           =   5
            End
            Begin VB.Menu mnu0805 
               Caption         =   "案件國家收費表維護"
               Index           =   6
            End
            Begin VB.Menu mnu0805 
               Caption         =   "國外代理人資料維護"
               Index           =   7
            End
            Begin VB.Menu mnu0805 
               Caption         =   "代理人變更名稱作業"
               Index           =   8
            End
            Begin VB.Menu mnu0805 
               Caption         =   "客戶資料維護"
               Index           =   9
            End
            Begin VB.Menu mnu0805 
               Caption         =   "客戶變更名稱作業"
               Index           =   10
            End
            Begin VB.Menu mnu0805 
               Caption         =   "案件聯絡人修改作業"
               Index           =   11
            End
            Begin VB.Menu mnu0805 
               Caption         =   "非本所實質客戶資料維護"
               Index           =   12
               Visible         =   0   'False
            End
            Begin VB.Menu mnu0805 
               Caption         =   "更換FC代理人作業"
               Index           =   13
            End
            Begin VB.Menu mnu0805 
               Caption         =   "法務工作點數分配"
               Index           =   14
               Visible         =   0   'False
            End
         End
      End
      Begin VB.Menu mnuTitle3 
         Caption         =   "ＡＣＳ"
         Index           =   3
         Begin VB.Menu mnu09 
            Caption         =   "資料處理"
            Index           =   1
            Begin VB.Menu mnu0901 
               Caption         =   "分案"
               Index           =   1
            End
            Begin VB.Menu mnu0901 
               Caption         =   "待送件區"
               Index           =   2
            End
            Begin VB.Menu mnu0901 
               Caption         =   "發文"
               Index           =   3
            End
            Begin VB.Menu mnu0901 
               Caption         =   "機關來函"
               Index           =   4
               Begin VB.Menu mnu090101 
                  Caption         =   "一般來函"
                  Index           =   1
               End
            End
            Begin VB.Menu mnu0901 
               Caption         =   "智財顧問案重新計算各部門實際比例"
               Index           =   5
            End
            Begin VB.Menu mnu0901 
               Caption         =   "TIPS案請款階段設定"
               Index           =   6
            End
            Begin VB.Menu mnu0901 
               Caption         =   "TIPS案請款階段分配比例維護作業"
               Index           =   7
            End
            Begin VB.Menu mnu0901 
               Caption         =   "TIPS案請款階段分配比例-年度結算作業"
               Index           =   8
            End
         End
         Begin VB.Menu mnu09 
            Caption         =   "承辦人工作"
            Index           =   2
            Begin VB.Menu mnu0902 
               Caption         =   "工作進度資料維護"
               Index           =   1
            End
            Begin VB.Menu mnu0902 
               Caption         =   "待核判區"
               Index           =   2
            End
            Begin VB.Menu mnu0902 
               Caption         =   "承辦人工作進度資料查詢"
               Index           =   3
            End
         End
         Begin VB.Menu mnu09 
            Caption         =   "查詢"
            Index           =   3
            Begin VB.Menu mnu0903 
               Caption         =   "ACS案件期限通知"
               Index           =   1
            End
         End
         Begin VB.Menu mnu09 
            Caption         =   "報表列印"
            Index           =   4
            Begin VB.Menu mnu0904 
               Caption         =   "期限管制表"
               Index           =   1
            End
            Begin VB.Menu mnu0904 
               Caption         =   "收文未發文明細表"
               Index           =   2
            End
            Begin VB.Menu mnu0904 
               Caption         =   "收/發文明細表"
               Index           =   3
            End
         End
         Begin VB.Menu mnu09 
            Caption         =   "檔案維護"
            Index           =   5
            Begin VB.Menu mnu0905 
               Caption         =   "創新業務案件基本資料維護"
               Index           =   1
            End
            Begin VB.Menu mnu0905 
               Caption         =   "案件進度資料維護"
               Index           =   2
            End
            Begin VB.Menu mnu0905 
               Caption         =   "下一程序資料維護"
               Index           =   3
            End
            Begin VB.Menu mnu0905 
               Caption         =   "案件國家收費表維護"
               Index           =   4
            End
         End
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "會稿判發"
      Index           =   5
      Begin VB.Menu mnu05 
         Caption         =   "專利／商標會稿"
         Index           =   1
      End
      Begin VB.Menu mnu05 
         Caption         =   "待核判區"
         Index           =   2
      End
      Begin VB.Menu mnu05 
         Caption         =   "發後補看作業"
         Index           =   3
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "智慧所案件"
      Index           =   6
      Begin VB.Menu mnu06 
         Caption         =   "顧問記錄"
         Index           =   1
      End
      Begin VB.Menu mnu06 
         Caption         =   "顧問明細及統計"
         Index           =   2
      End
      Begin VB.Menu mnu06 
         Caption         =   "介紹案源管理"
         Index           =   3
      End
      Begin VB.Menu mnu06 
         Caption         =   "介紹案源查詢"
         Index           =   4
      End
      Begin VB.Menu mnu06 
         Caption         =   "智財訴訟案需專業部配合通知補收文作業"
         Index           =   5
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
         Caption         =   "撰寫信函作業"
         Index           =   7
      End
      Begin VB.Menu mnu11 
         Caption         =   "作業失誤"
         Index           =   8
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
               Caption         =   "請款單折扣案件明細"
               Index           =   10
            End
         End
      End
      Begin VB.Menu mnu11 
         Caption         =   "銷案延遲日期輸入作業"
         Index           =   10
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘單查詢"
         Index           =   11
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘資料維護"
         Index           =   12
      End
      Begin VB.Menu mnu11 
         Caption         =   "CF 結餘單案件明細查詢"
         Index           =   14
      End
      Begin VB.Menu mnu11 
         Caption         =   "客戶應收帳款收文檢查上限"
         Enabled         =   0   'False
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnu11 
         Caption         =   "客戶預定收款日放寬月數上限"
         Enabled         =   0   'False
         Index           =   16
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "定稿"
      Index           =   15
      Begin VB.Menu mnu15 
         Caption         =   "整批列印定稿"
         Index           =   1
         Begin VB.Menu mnu1501 
            Caption         =   "橫式定稿"
            Index           =   1
         End
         Begin VB.Menu mnu1501 
            Caption         =   "直式定稿"
            Index           =   2
         End
         Begin VB.Menu mnu1501 
            Caption         =   "日文定稿"
            Index           =   3
         End
      End
      Begin VB.Menu mnu15 
         Caption         =   "定稿資料維護"
         Index           =   2
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "國外開拓"
      Index           =   17
      Begin VB.Menu mnu17 
         Caption         =   "潛在客戶資料維護"
         Index           =   1
      End
      Begin VB.Menu mnu17 
         Caption         =   "客戶/代理人聯絡人資料維護"
         Index           =   2
      End
      Begin VB.Menu mnu17 
         Caption         =   "往來記錄維護"
         Index           =   3
      End
      Begin VB.Menu mnu17 
         Caption         =   "交換名片紀錄維護"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnu17 
         Caption         =   "互惠代理人案件統計表"
         Index           =   5
      End
      Begin VB.Menu mnu17 
         Caption         =   "潛在客戶名條列印"
         Index           =   6
      End
      Begin VB.Menu mnu17 
         Caption         =   "潛在客戶資料查詢"
         Index           =   7
      End
      Begin VB.Menu mnu17 
         Caption         =   "往來記錄資料查詢"
         Index           =   8
      End
      Begin VB.Menu mnu17 
         Caption         =   "往來記錄統計"
         Index           =   9
      End
      Begin VB.Menu mnu17 
         Caption         =   "國外部新客戶/代理人查詢"
         Index           =   10
      End
      Begin VB.Menu mnu17 
         Caption         =   "不得宣傳客戶名稱資料查詢"
         Index           =   11
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "一般作業"
      Index           =   23
      Begin VB.Menu mnu23 
         Caption         =   "會議室/檢索系統預約作業"
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnu23 
         Caption         =   "圖書借閱資料查詢 "
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
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/11 智權部mnuTitle(16)功能表已拆出不使用,故拿掉選單及程式-秀玲
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Modify By Sindy 2010/8/31 日期欄已修改
'Memo by Lydia 2019/07/01 智權部mnuTitle(16)功能表全部隱藏
Option Explicit

'Add by Morgan 2003/12/23
Dim WithEvents eventConn As ADODB.Connection
Attribute eventConn.VB_VarHelpID = -1
Public bolReOpen As Boolean
'intPCaseKind分案之系統分類，intPWhere 0國內  1國外CF  2國外FC
Public intPCaseKind As Integer, intPWhere As Integer
'Add by Morgan 2008/11/7 是否已經做過
Dim m_blnActivated As Boolean
Public m_wasMaximized As Boolean 'Added by Morgan 畫面最小化後判斷原來是否為最大化用
Public m_ChkIsOpenFrm180203 As Boolean 'Add By Sindy 2013/7/8
Dim oControl As Control  'Added by Morgan 2022/1/22


'Add by Morgan 2003/12/23
Private Sub eventConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
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

Private Sub MDIForm_Resize()
   'Added by Morgan 2011/12/14 紀錄是否為最大化狀態
   If Me.WindowState = 2 Then
      m_wasMaximized = True
   ElseIf Me.WindowState = 0 Then
      m_wasMaximized = False
   End If
End Sub

Private Sub mnu05_Click(Index As Integer)
   Select Case Index
      'Add By Sindy 2015/8/6
      Case 1 '待會稿區 'Memo by Lydia 2019/07/01 更名: 專利／商標會稿
         frm090202_3.Show
      'Add By Sindy 2015/5/21
      Case 2 '待核判區
         frm090202_1.m_ProSysState = "1" '承辦人
         frm090202_1.Show
      'Add By Sindy 2021/11/11
      Case 3 '發後補看作業
         '考慮職代問題,不必鎖權限
         frm040117.m_ProState = "T"
         frm040117.Show
      Case Else
   End Select
End Sub

Private Sub mnu080103_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '法院文書
         If CheckUse("frm071008", strExec) = True Then
            frm071008.Show
         End If
      Case 2  '回執
         If CheckUse("frm071010", strExec) = True Then
            frm071010.Show
         End If
      Case 3  '開庭通知
         If CheckUse("frm071012", strExec) = True Then
            frm071012.Show
         End If
      Case 4  '其他來函
         If CheckUse("frm071014", strExec) = True Then
            frm071014.Show
         End If
      'Add By Sindy 2011/6/14
      Case 5  '信件退回
         If CheckUse("frm071019", strExec) = True Then
            frm071019.Show
         End If
   End Select
End Sub

'add by sonia 2019/7/24
Private Sub mnu0901_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1  '分案
      If CheckUse("frm081031", strExec) = True Then
         frm081031.Show
      End If
   Case 2  '待送件區
      If CheckUse("frm090202_4", strExec) = True Then
         frm090202_4.m_ProState = "A" '一般
         frm090202_4.Show
      End If
   Case 3  '發文
      'Modified by Lydia 2024/03/25 改成獨立程式
      'If CheckUse("frm071004", strExec) = True Then
      '   frm071004.Show
      If CheckUse("frm081035", strExec) = True Then
         frm081035.Show
      'end 2024/03/25
      End If
   'Added by Lydia 2021/06/03
   Case 5 '智財顧問案重新計算各部門實際比例
      If CheckUse("frm081033", strExec) = True Then
         frm081033.Show
      End If
   'end 2021/06/03
   'Added by Lydia 2024/03/25
   Case 6   'TIPS案請款階段設定
      If CheckUse("frm081034", strExec) = True Then
         frm081034.Show
      End If
   'end 2024/03/25
   'Added by Lydia 2025/04/18
   Case 7   'TIPS案請款階段分配比例維護作業
      If CheckUse("frm081036", strExec) = True Then
         frm081036.Show
      End If
   Case 8   'TIPS案請款階段分配比例-年度結算作業
      If CheckUse("frm081036_1", strExec) = True Then
         frm081036_1.Show
      End If
   'end 2025/04/18
   End Select

End Sub

Private Sub mnu090101_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1  '一般來函
      If CheckUse("frm071014", strExec) = True Then
         frm071014.Show
         frm071014.Caption = "一般來函" 'Added by Lydia 2023/03/17
      End If
   End Select
End Sub

Private Sub mnu0902_Click(Index As Integer)
Dim nFrm As Form
   
   ToolHide
   Select Case Index
      Case 1 '工作進度資料維護
      If CheckUse("frm090201_4", strExec) Then
         ProState = "1"
         ProSysState = "1"
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
         
            '2009/11/12 modify by sonia 改寫法無資料不顯示畫面
            'frm090201_4.Show
            frm090201_4.StrMenu1   '當天本所期限案件資料,無資料時由frm090201_4的nextstep執行下一畫面
            If frm090201_4.TextOk = True Then frm090201_4.Show
            '2009/11/12 end
            
            'Add By Sindy 2021/10/14
            '沒有核完的天數控管, 2個工作天(不含當天) 發mail提醒，進承辦人工作進度時也要提醒
            '沒有判發的天數控管, 2個工作天(不含當天) 發mail提醒，進承辦人工作進度時也要提醒
            '判發後隔天沒有發文的要發mail提醒 , 進承辦人工作進度時也要提醒
            Dim rsTmp As New ADODB.Recordset
            If PUB_ChkEmpElePro(strUserNum, "A", rsTmp) = True Then
               If MsgBox("您有已逾時的待核判案件，是否直接進入待核判區處理？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                  '檢查表單是否已開啟，若是，則關閉
                  For Each nFrm In Forms
                     If StrComp(nFrm.Name, "frm090202_1", vbTextCompare) = 0 Then
                        Unload frm090202_1
                     End If
                  Next
                  frm090202_1.m_ProSysState = "1" '承辦人
                  frm090202_1.Show
               End If
               rsTmp.Close
            End If
            If PUB_ChkEmpElePro(strUserNum, "B", rsTmp) = True Then
               If MsgBox("您有之前尚未歸檔欲發文的案件，是否直接進入待送件區處理？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                  '檢查表單是否已開啟，若是，則關閉
                  For Each nFrm In Forms
                     If StrComp(nFrm.Name, "frm090202_4", vbTextCompare) = 0 Then
                        Unload frm090202_4
                     End If
                  Next
                  frm090202_4.m_ProState = "A" '一般
                  frm090202_4.Show
               End If
               rsTmp.Close
            End If
            Set rsTmp = Nothing
            '2021/10/14 END
'         End If 'Added by Morgan 2013/10/8
      End If
      
      'Add By Sindy 2021/7/16
      Case 2 '待核判區
         frm090202_1.m_ProSysState = "1" '承辦人
         frm090202_1.Show
      
      'Add By Sindy 2021/9/24
      Case 3 '承辦人工作進度資料查詢
         If CheckUse("frm090614", strExec) Then
            ProState = "2"
            ProSysState = "1"
            frm090614.Show
         End If
   End Select
End Sub

'Added by Lydia 2020/12/09
Private Sub mnu0903_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1  'ACS案件期限通知
      If CheckUse("frm081032", strExec) = True Then
         frm081032.Show
      End If
   End Select
End Sub

Private Sub mnu0904_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '期限管制表
         If CheckUse("frm083001", strExec) = True Then
             frm083001.Tag = 2
             frm083001.Show
         End If
      Case 2  '收文未發文明細表
         If CheckUse("frm083003", strExec) = True Then
             frm083003.Tag = 2
             frm083003.Show
         End If
      Case 3  '收 / 發文明細表
         If CheckUse("frm083006", strExec) = True Then
             frm083006.Tag = 2
             frm083006.Show
         End If
   End Select
End Sub

Private Sub mnu0905_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1 '創新業務案件基本資料維護
         If CheckUse("frm075002", strExec) = True Then
            strSysKind = "ACS"
            frm075002.Show
         End If
     
   Case 2 '案件進度檔資料維護
         If CheckUse("frm075004_1", strExec) = True Then
            strSysKind = "ACS"
            frm075004_1.Show
         End If
     
   Case 3 '下一程序資料維護
         If CheckUse("frm075007_1", strExec) = True Then
            strSysKind = "ACS"
            frm075007_1.Show
         End If
     
   Case 4 '案件國家收費表維護
         If CheckUse("frm12040102", strExec) = True Then
            frm12040102.Show
         End If
   End Select
End Sub
'end 2019/7/24

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
      Case 3 '以申請人查詢(查新客戶)
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

'Add by Morgan 2007/12/12
Private Sub mnu17_Click(Index As Integer)
   Select Case Index
      Case 1 '潛在客戶資料維護
         If CheckUse("frm140402", strExec) = True Then
            frm140402.Show
         End If
      
      Case 2 '客戶/代理人聯絡人資料維護
         If CheckUse("frm140403", strExec) = True Then
            frm140403.Show
         End If
         
      Case 3 '往來記錄資料維護
         If CheckUse("frm140404", strExec) = True Then
            frm140404.Show
         End If
      'Add by Morgan 2009/2/6
      'Memo by Lydia 2022/01/12 經過確認往來記錄代碼不存在，並且最後資料在2008年；所以先隱藏功能選單
      Case 4 '交換名片紀錄維護
         If CheckUse("frm140411", strExec) = True Then
            frm140411.Show
         End If
      'Add by Morgan 2008/6/3
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

Private Sub mnu23_Click(Index As Integer)
   ToolHide
   Select Case Index
      'Add by Morgan 2011/5/18
      Case 1 '會議室預約作業
         frm140112.Show
      'Added by Morgan 2018/1/18
      Case 4 '專利處研討會
         frm140113.Show
'      'Add By Sindy 2016/3/21
'      Case 7 '系統收件區
'         frm06010612.Show
'      '2016/3/21 END
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

'Add By Sindy 2009/08/26
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

'Add By Sindy 2015/5/21
Private Sub mnuChUser_Click()
   frmChgUser.Show
End Sub

Private Sub mnuDML_Click(Index As Integer)
    frmDML.Show
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

''2005/12/20 ADD BY SONIA
'Private Sub Timer4_Timer()
'If IsLoginFormIsUnload = True Then
'    IsLoginFormIsUnload = False
'    Dim tmpST15 As String
'    Dim tmpST03 As String
'    Dim tmpST06 As String   '2005/12/20 ADD BY SONIA
'    ''add by nickc 2005/09/02 非員工不跑
'    If strUserNum >= "63001" And strUserNum < "A" Then
'      tmpST15 = PUB_GetStaffST15(strUserNum, 1)
'      'add by nickc 2005/09/20
'      tmpST03 = PUB_GetST03(strUserNum)
'      tmpST06 = PUB_GetST06(strUserNum)   '2005/12/20 ADD BY SONIA
'      'edit by nickc 2005/09/20
'      'If UCase(Mid(tmpST15, 1, 1)) = "S" Then
'      '2005/12/20 MODIFY BY SONIA 加入中所法務
'      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Then
'      'edit by nickc 2006/04/04
'      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Or (UCase(Mid(tmpST03, 1, 2)) = "P3" And tmpST06 = "2") Then
'      '2006/6/1 MODIFY BY SONIA 取消P2,P3中所之控制,改由PUB_ChkNotSalesButHaveCase控制
'      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Or (UCase(Mid(tmpST03, 1, 2)) = "P3" And tmpST06 = "2") Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
'      If UCase(Mid(tmpST15, 1, 1)) = "S" Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
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
'      End If
'   End If
'End If
'End Sub
''2005/12/20 END

'Add by Morgan 2003/12/23
'控制連線閒置超過30分鐘自動離線
Private Sub tmrConnect_Timer()
   tmrConnect.Tag = tmrConnect.Tag + 1
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

Private Sub MDIForm_Activate()
   'Modify By Sindy 2025/11/3 改為共用函數
   Call MDIFormStarProc
End Sub

'Add By Sindy 2011/10/7
Public Sub SysStartCallForm()
   'Add By Sindy 2011/3/17
   '只要執行一次
   If m_blnActivated = False Then
      m_blnActivated = True
      '智權人員期限資料查詢
      SalesDueCaseQuery
   End If
   
   'Add By Sindy 2009/08/26
   '本查詢需考慮當閒置太久重新登入且已經是下午時須再次執行故與單獨控制
   If pub_bolInformCheck = True Then
      '國外部法務處期限通知
      FcpDueCaseQurey
      
      'Added by Lydia 2020/01/15 非外專人員行事曆提醒通知
      If PUB_CheckStaffCalendarDue = True Then
           frm060209.m_Role = "F41"
           frm060209.Show
      End If
      'end 2020/01/15
      'Added by Lydia 2020/12/10 ACS案：顧服組W20及其部門主管A0908
      If Pub_StrUserSt03 = "W20" Or InStr(GetDeptMan("W20") & ",", strUserNum) > 0 Then
         strSql = "select * from executelog where el01='frm081032' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI <> 1 Then
            If CheckUse("frm081032", strExec) = True Then
               pub_bolInformCheck = True
               Load frm081032
               frm081032.CmdQuery(0).Value = True
               pub_bolInformCheck = False
            End If
         End If
      End If
      'end 2020/12/10
      
      LawOfficeCaseQuery 'Added by Morgan 2020/4/28 法律所案源查詢
      
      pub_bolInformCheck = False
   End If
End Sub

'Added by Morgan 2020/4/28
'法律所案源查詢
Private Sub LawOfficeCaseQuery()
   '等級53、L1人員每天早上、中午第一次進入系統都要自動執行，若不是第一次則詢問是否要執行
   'Modified by Morgan 2020/7/20 +52,L5
   If mnu06(3).Visible = True And (Pub_strUserST05 = "52" Or Pub_strUserST05 = "53" Or Pub_strUserST05 = "L1" Or Pub_strUserST05 = "L5") Then
      strSql = "select * from executelog where el01='frm077003' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If MsgBox("是否執行  介紹案源管理  功能", vbYesNo, "功能！") = vbYes Then
            intI = 0
         End If
      End If
      If intI = 0 Then
         pub_CallNextForm = True
         frm077003.Show
         frm077003.CmdSearch.Value = True
         frm077003.ZOrder 1 '若有期限彈跳視窗時顯示在後面
      End If
   End If
End Sub

Private Sub MDIForm_Load()
'Add by Morgan 2003/12/23
If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
   Set eventConn = cnnConnection
   tmrConnect.Interval = 60000
End If

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
       Else
            mnuDML(0).Visible = False
       End If

'       'Added by Morgan 2016/1/22 薪資查詢測試
'       If Pub_StrUserSt03 = "M51" Then
'         mnu23(6).Visible = True
'       Else
'         mnu23(6).Visible = False
'       End If
'       'end
'
      'Add by Amy 2017/01/25
      If strSrvDate(1) >= 20170202 Then
        mnu23(8).Visible = True
      Else
        mnu23(8).Visible = False
      End If
      
     
       'Add By Sindy 2010/01/07 M51及王副總才可以看到
       'If Pub_StrUserSt03 = "M51" Or strUserNum = "71011" Then
       'If Pub_StrUserSt03 = "M51" Or CheckUse("frm050207", strExec) = True Then
       If Pub_StrUserSt03 = "M51" Then
          mnu0802(10).Visible = True
          mnuChUser.Visible = True 'Add By Sindy 2015/5/21
       Else
          mnu0802(10).Visible = False
          mnuChUser.Visible = False 'Add By Sindy 2015/5/2
       End If
      
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
'            mnu10(23).Visible = False
       End If
       'add by nickc 2008/05/01
       'Mark by Amy 2021/11/11 智權部已拆出不使用,故拿掉選單及程式-秀玲
'        If Pub_StrUserSt03 = "M51" Then
'            mnu1603(3).Visible = True
'        Else
'            mnu1603(3).Visible = False
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
'edit by nickc 2007/02/07 不用 dll 了
'Set obj001 = Nothing
'Set objPublicData = Nothing
   ' 90.08.16 modify by louis 釋放Word物件
   EndOfficeAp
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

Private Sub mnu0701_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1  '分案
      If CheckUse("frm071001", strExec) = True Then
         frm071001.Show
      End If
   Case 2  '會稿日輸入
      If CheckUse("frm071016", strExec) = True Then
         frm071016.Show
      End If
   Case 3  '發文
      If CheckUse("frm071004", strExec) = True Then
         frm071004.Show
      End If
   '2005/12/28 ADD BY SONIA 顧問案件電話諮詢  '2011/5/26自共同程序作業移過來
   Case 5
      If CheckUse("frm010014", strExec) = True Then
         frm010014.Show
      End If
   'Add by Amy 2017/12/27 合約基本資料維護
   Case 6
      If CheckUse("frm075008", strExec, False) = True Or ChkContractLimit(True, , False) = True Then
         frm075008.Show
      Else
         MsgBox "無此使用權限...", , "警告!!"
      End If
   'Add by Amy 2019/01/08 律師執業登錄地區維護
   Case 7
      If CheckUse("frm071022", strExec, False) = True Then
         frm071022.Show
      End If
   'Added by Lydia 2024/07/29 (113/11/01上線)出庭費確認維護
   Case 8
       If CheckUse("frm075013", strExec, False) = True Then
         frm075013.Show
       End If
   End Select
End Sub

Private Sub mnu070104_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1  '法院文書
      If CheckUse("frm071008", strExec) = True Then
         frm071008.Show
      End If
   Case 2  '回執
      If CheckUse("frm071010", strExec) = True Then
         frm071010.Show
      End If
   Case 3  '開庭通知
      If CheckUse("frm071012", strExec) = True Then
         frm071012.Show
      End If
   Case 4  '其他來函
      If CheckUse("frm071014", strExec) = True Then
         frm071014.Show
      End If
   'Add By Sindy 2011/6/14
   Case 5  '信件退回
      If CheckUse("frm071019", strExec) = True Then
         frm071019.Show
      End If
   End Select
End Sub

Private Sub mnu0702_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1  '庭期資料查詢
      If CheckUse("frm072001", strExec) = True Then
         frm072001.Show
      End If
   Case 2  '承辦/協辦人員案件查詢
      If CheckUse("frm072003", strExec) = True Then
          frm072003.Tag = 0
          frm072003.Show
      End If
   'Add By Cheng 2002/09/24
   Case 3  '代理人案件性質統計
      If CheckUse("frm050204_1", strExec) = True Then
          frm050204_1.Show
      End If
   'Add By Sindy 2011/6/20
   Case 4  '法務處期限通知
      If CheckUse("frm072005", strExec) = True Then
          frm072005.Show
      End If
   'Added by Lydia 2024/07/29 (113/11/01上線)出庭費查詢
   Case 5
      If CheckUse("frm075013_2", strExec, False) = True Then
        frm075013_2.Show
      End If
   End Select
End Sub

Private Sub mnu0703_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '期限管制表
         If CheckUse("frm083001", strExec) = True Then
             frm083001.Tag = 0
             frm083001.Show
         End If
      Case 2  '收文未發文明細表
         If CheckUse("frm083003", strExec) = True Then
             frm083003.Tag = 0
             frm083003.Show
         End If
      Case 3  '催審表
         If CheckUse("frm083004", strExec) = True Then
             frm083004.Tag = 0
             frm083004.Show
         End If
      Case 4  '收 / 發文明細表
         If CheckUse("frm083006", strExec) = True Then
             frm083006.Tag = 0
             frm083006.Show
         End If
      Case 5  '人員工作分析表
         If CheckUse("frm083007", strExec) = True Then
             frm083007.Tag = 0
             frm083007.Show
         End If
      Case 6  '收款業績比較表
         If CheckUse("frm083008", strExec) = True Then
             frm083008.Tag = 0
             frm083008.Show
         End If
      Case 7  '業績分析表
         If CheckUse("frm083009", strExec) = True Then
            frm083009.Tag = 0
            frm083009.Show
          End If
      Case 8  '客戶案件總簿
         If CheckUse("frm050317", strExec) = True Then
            'Modify by Amy 2017/08/04 原:frm050317
             frm0503171.Tag = 0
             frm0503171.Show
         End If
          'frm083011.Tag = 0
          'frm083011.Show
      Case 9  '顧問到期明細表
         If CheckUse("frm073008", strExec) = True Then
            frm073008.Tag = 0
            frm073008.Show
         End If
        'Add By Cheng 2003/04/18
      Case 10 '顧問客戶資料表
         If CheckUse("frm073011", strExec) = True Then
            frm073011.Tag = 0
            frm073011.Show
         End If
      Case 11  '顧問地址條
         If CheckUse("frm073009", strExec) = True Then
            frm073009.Tag = 0
            frm073009.Show
         End If
      Case 12  '地址條列印
         If CheckUse("frm083014", strExec) = True Then
            frm083014.Tag = 0
            frm083014.Show
         End If
   End Select
End Sub

Private Sub mnu0704_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '收／發文統計表
       If CheckUse("frm084001", strExec) = True Then
             frm084001.Tag = 0
             frm084001.Show
             'Modified by Lydia 2023/01/16 增加類別
             'frm084001.Label5(2).Caption = "系統別：                (1.LA, 2.L, 3.全部)"
             frm084001.Label5(2).Caption = "系統別：                (1.LA, 2.L, 3.L民刑事, 4.全部)"
       End If
      
      Case 2  '法務委任案件統計表
         If CheckUse("frm084002", strExec) = True Then
             frm084002.Tag = 0
             frm084002.Show
         End If
         'frm084002.Label1(1).Caption = "系統別：                         (1.L, 2.LA, 3.全部)"
      Case 3  '法務案件年度統計表
        If CheckUse("frm084003", strExec) = True Then
             frm084003.Tag = 0
             frm084003.Show
         End If
       
         'frm084003.Label1(2).Caption = "系統別：                           (1.L, 2.LA, 3.全部)"
      Case 4  '逾期未結案統計表
         If CheckUse("frm084004", strExec) = True Then
             frm084004.Tag = 0
             frm084004.Show
         End If
        
   End Select
End Sub

Private Sub mnu0705_Click(Index As Integer)
   ToolHide
   Select Case Index
   Case 1 '法務案件基本資料維護
        If CheckUse("frm075002", strExec) = True Then
             strSysKind = "L"
             frm075002.Tag = 0
             frm075002.Show
         End If
     
   Case 2 '顧問案件資料維護
        If CheckUse("frm075006", strExec) = True Then
             frm075006.Tag = 0
             frm075006.Show
         End If
   
   Case 3 '案件進度檔資料維護
         If CheckUse("frm075004_1", strExec) = True Then
             frm075004_1.Tag = 0
             frm075004_1.Show
         End If
     
   Case 4 '下一程序資料維護
         If CheckUse("frm075007_1", strExec) = True Then
             frm075007_1.Tag = 0
             frm075007_1.Show
         End If
     
   Case 5 '機關單位資料維護
          If CheckUse("frm075010", strExec) = True Then
             frm075010.Tag = 0
             frm075010.Show
         End If
    
   Case 6 '庭期資料維護
         If CheckUse("frm075011", strExec) = True Then
             frm075011.Tag = 0
             frm075011.Show
         End If
   Case 7 '案件國家收費表維護
         If CheckUse("frm12040102", strExec) = True Then
             frm12040102.Tag = 0
             frm12040102.Show
         End If
   'Add By Sindy 2014/10/27
   Case 8 '更換FC代理人作業
      If CheckUse("frm110104_1", strExec) = True Then
         frm110104_1.Show
      End If
   'Added by Lydia 2015/08/06 預先上傳
   Case 9 '法務工作點數分配
      If CheckUse("frm071021", strExec) = True Then
         frm071021.Show
      End If
   End Select
End Sub

Private Sub mnu0801_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '分案
         If CheckUse("frm081001", strExec) = True Then
            frm081001.Show
         End If
         
      Case 2  '會稿日輸入
         If CheckUse("frm071016", strExec) = True Then
            frm071016.Show
         End If

      Case 3  '發文
         If CheckUse("frm071004", strExec) = True Then
            frm071004.Show
         End If

      Case 6  '開拓客戶資料維護
        If CheckUse("frm081020", strExec) = True Then
            frm081020.Show
         End If

       
   End Select
End Sub

Private Sub mnu080104_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '法院文書
         If CheckUse("frm071008", strExec) = True Then
            frm071008.Show
         End If
      
      Case 2  '回執
         If CheckUse("frm071010", strExec) = True Then
            frm071010.Show
         End If

      Case 3  '開庭通知
       If CheckUse("frm071012", strExec) = True Then
            frm071012.Show
         End If

      Case 4  '其他來函
        If CheckUse("frm071014", strExec) = True Then
            frm071014.Show
         End If
        
   End Select
End Sub

Private Sub mnu080105_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '代理人已收達/已提申
        If CheckUse("frm081016", strExec) = True Then
            frm081016.Show
         End If
         
      Case 2  '其他來函
        If CheckUse("frm071014", strExec) = True Then
            frm071014.Show
         End If
       
   End Select
End Sub

Private Sub mnu0802_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '庭期資料查詢
        If CheckUse("frm072001", strExec) = True Then
            frm072001.Show
        End If
          
      Case 2  '承辦/協辦人員案件查詢
        If CheckUse("frm072003", strExec) = True Then
           frm072003.Tag = 1
           frm072003.Show
        End If
      Case 3  '代理人新案案件統計
       If CheckUse("frm050201", strExec) = True Then
            frm050201.Show
        End If
'edit by nickc 2005/07/22
'      Case 4
'        If CheckUse("frm040202", strExec) = True Then
'            frm040202.Show
'        End If
   
      Case 5  '未請款明細查詢
        If CheckUse("frm050203", strExec) = True Then
           'Modify By Sindy 2009/07/24 增加LIN系統類別
           'modify by sonia 2019/7/29 +ACS系統類別
           StrStartSystemByNick = "FCL,CFL,LIN,ACS"
           frm050203.Show
        End If
         
      Case 6  'FC收款請款點數查詢
        If CheckUse("frm040205", strExec) = True Then
           '2007/11/16 modify by sonia 改預設所有系統類別,因外商會收法務系統
           'StrStartSystemByNick = GetSystemKindByNick
           StrStartSystemByNick = "ALL"
           '2007/11/16 ADD BY SONIA 外法使用者預設條件
           If (Mid(GetStaffDepartment(strUserNum), 1, 2) = "F3" Or Mid(GetStaffDepartment(strUserNum), 1, 2) = "F4") Then
              frm040205.txt1(10) = "F30"
              frm040205.txt1(11) = "F49"
           End If
           '2007/11/16 END
           frm040205.Show
        End If
         'frm082011.Show
       
      Case 7  '開拓客戶查詢
        If CheckUse("frm082005", strExec) = True Then
            frm082005.Show
        End If
      'Add By Cheng 2002/09/24
      Case 8 '代理人案件性質統計
        If CheckUse("frm050204_1", strExec) = True Then
            frm050204_1.Show
        End If
        
      'Add By Sindy 2009/08/26
      'Modify By Sindy 2017/9/12 內外法期限通知改同一支
      Case 9 '法務處期限通知
         If CheckUse("frm072005", strExec) = True Then
            frm072005.Show
         End If
      
      'Add By Sindy 2010/01/07 員工查詢印表記錄檔查詢
      Case 10
         If CheckUse("frm050207", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm050207.Show
         End If
         
      'Add By Sindy 2011/1/14 國外業務帳款查詢
      Case 11
         If CheckUse("Frmacc2260", strExec) = True Then
            Frmacc2260.Show
            ToolShow
            tool3_enabled
         End If
      End Select
End Sub

Private Sub mnu0803_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1    '期限管制表
        If CheckUse("frm083001", strExec) = True Then
           frm083001.Tag = 1
           frm083001.Show
        End If
        
      Case 2    '代理人案件收達 / 提申管制表
         If CheckUse("frm083002", strExec) = True Then
            frm083002.Show
         End If

      Case 3    '收文未發文明細表
        If CheckUse("frm083003", strExec) = True Then
             frm083003.Tag = 1
             frm083003.Show
        End If
        
      Case 4     '催審表
          If CheckUse("frm083004", strExec) = True Then
            frm083004.Tag = 1
            frm083004.Show
          End If
         
      Case 5     '智權人員員收文明細表
        If CheckUse("frm083005", strExec) = True Then
            frm083005.Show
        End If

      Case 6     '收 / 發文明細表
        If CheckUse("frm083006", strExec) = True Then
             frm083006.Tag = 1
             frm083006.Show
        End If
      'Added by Lydia 2016/12/13
      Case 7   '請款點數明細表
        If CheckUse("frm083016", strExec) = True Then
            frm083016.Show
        End If
      'Modified by Lydia 2016/12/12 Case 7=>Case 8
      Case 8     '人員工作分析表
           If CheckUse("frm083007", strExec) = True Then
             frm083007.Tag = 1
             frm083007.Show
        End If
      'Modified by Lydia 2016/12/12 Case 8=>Case 9
      Case 9     '收款業績比較表
        If CheckUse("frm083008", strExec) = True Then
            frm083008.Tag = 1
            frm083008.Show
        End If
        
      'Modified by Lydia 2016/12/12 Case 9=>Case 10
      Case 10     '業績分析表
         If CheckUse("frm083009", strExec) = True Then
             frm083009.Tag = 1
             frm083009.Show
         End If
         
      'Modified by Lydia 2016/12/12 Case 10=>Case 11
      Case 11     '代理人案件總簿
          If CheckUse("frm050316", strExec) = True Then
            frm050316.Show
          End If
          
      'Modified by Lydia 2016/12/12 Case 11=>Case 12
      Case 12     '客戶案件總簿
        If CheckUse("frm050317", strExec) = True Then
            'Modify by Amy 2017/08/04 原:frm050317
            frm0503171.Show
        End If

      'Modified by Lydia 2016/12/12 Case 12=>Case 13
      Case 13     '代理人/申請人名單
       If CheckUse("frm050318", strExec) = True Then
            frm050318.Show
        End If

         'frm083012.Show
'edit by nickc 2005/11/10 取消舊結餘
'      Case 13
'         If CheckUse("frm040320", strExec) = True Then
'            frm040320.Show
'        End If
      Case 14     '地址條列印
        If CheckUse("frm083014", strExec) = True Then
            frm083014.Show
        End If

      Case 15     '開拓客戶地址條
        If CheckUse("frm083015", strExec) = True Then
            frm083015.Show
        End If

      'Added by Lydia 2023/01/16
      Case 16  'DHL列印
        If CheckUse("frm060330", strExec) = True Then
           frm060330.Show
        End If
   End Select
End Sub

Private Sub mnu0804_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '收／發文統計表
         If CheckUse("frm084001", strExec) = True Then
            frm084001.Tag = 1
            frm084001.Show
            frm084001.Label5(2).Caption = "系統別：                (1.FCL, 2.CFL, 3.全部)"
         End If
      Case 2  '法務委任案件統計表
         If CheckUse("frm084002", strExec) = True Then
           frm084002.Tag = 1
           frm084002.Show
         End If

      Case 3  '法務案件年度統計表
         If CheckUse("frm084003", strExec) = True Then
            frm084003.Tag = 1
            frm084003.Show
         End If
         
        ' frm084003.Label1(2).Caption = "系統別：                           (1.FCL, 2.CFL, 3.全部)"
      Case 4  '逾期未結案統計表
         If CheckUse("frm084004", strExec) = True Then
            frm084004.Tag = 1
            frm084004.Show
         End If
         
   End Select
End Sub

Private Sub mnu0805_Click(Index As Integer)
    ToolHide
    Select Case Index
       Case 1  '法務案件基本資料維護
       If CheckUse("frm075002", strExec) = True Then
           strSysKind = ""
           frm075002.Tag = 1
           frm075002.Show
        End If
           intWSysKindThis = "2"
  
       Case 2  '案件進度檔資料維護
       If CheckUse("frm075004_1", strExec) = True Then
           frm075004_1.Tag = 1
           frm075004_1.Show
        End If
   
       Case 3  '下一程序資料維護
       If CheckUse("frm075007_1", strExec) = True Then
           frm075007_1.Tag = 1
           frm075007_1.Show
        End If
     
       Case 4  '機關單位資料維護
       If CheckUse("frm075010", strExec) = True Then
           frm075010.Tag = 1
           frm075010.Show
        End If
          
       Case 5  '庭期資料維護
       If CheckUse("frm075011", strExec) = True Then
           frm075011.Tag = 1
           frm075011.Show
        End If
           
       Case 6  '案件國家收費表維護
       If CheckUse("frm12040102", strExec) = True Then
           frm12040102.Tag = 1
           frm12040102.Show
        End If
           
       Case 7  '國外代理人資料維護
       If CheckUse("frm050705", strExec) = True Then
           frm050705.Tag = 1
           frm050705.Show
        End If
           
       Case 8  '代理人變更名稱作業
       If CheckUse("frm140103", strExec) = True Then
           frm140103.Tag = 1
           frm140103.Show
        End If
        
       Case 9  '客戶基本資料維護
       If CheckUse("frm140401", strExec) = True Then
           frm140401.Tag = 1
           frm140401.Show
        End If
      
       Case 10  '客戶變更名稱作業
       If CheckUse("frm140101", strExec) = True Then
           frm140101.Tag = 1
           frm140101.Show
        End If
           
       Case 11  '案件聯絡人修改作業
         If CheckUse("frm050713", strExec) = True Then
             frm050713.Tag = 1
             frm050713.Show
          End If
      
      'Add By Sindy 2012/4/10
      Case 12 '非本所實質客戶資料維護
         If CheckUse("frm12040155", strExec) = True Then
            frm12040155.Show
         End If
      'Add By Sindy 2014/10/27
      Case 13 '更換FC代理人作業
         If CheckUse("frm110104_1", strExec) = True Then
            frm110104_1.Show
         End If
      'Added by Lydia 2015/08/06 預先上傳
      Case 14 '法務工作點數分配
         If CheckUse("frm071021", strExec) = True Then
            frm071021.Show
         End If
    End Select
End Sub

Private Sub mnu10_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 4 '員工姓名查詢員工資料
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
      'Add By Cheng 2003/06/26
      Case 7 '撰寫信函作業
'         If CheckUse("frm090401", strExec) = True Then
            frm090401.Show
'         End If
      'add by nickc 2005/04/08 銷案延遲日期輸入作業
      Case 10
         If CheckUse("frm140501", strExec) = True Then
            frm140501.Show
         End If
      'add by nickc 2005/07/22 CF 結餘單查詢
      Case 11
         If CheckUse("frm040202", strExec) = True Then
            frm040202.Show
         End If
      'add by nickc 2005/07/22 CF 結餘資料維護
      Case 12
         If CheckUse("frm040206", strExec) = True Then
            frm040206.Show
         End If
      'add by nickc 2007/11/13 加入客戶特殊紀錄
      'Remove by Lydia 2022/05/09 改放在Promoter的智權部->區主管作業
      'Case 13
      '   If CheckUse("frm010022", strExec) = True Then
      '      frm010022.Show
      '   End If
      'end 2022/05/09
      'add by nickc 2008/03/27 CF 結餘單案件明細查詢
      Case 14
         If CheckUse("frm040208", strExec) = True Then
            frm040208.Show
         End If
'Remove by Lydia 2015/10/14 移到Patpro之共同程序
'      'Add By Sindy 2012/12/10 客戶應收帳款收文檢查上限
'      Case 15
'         If CheckUse("frm140502", strExec) = True Then
'            frm140502.Show
'         End If
'      'Add By Sindy 2015/9/17 客戶預定收款日放寬月數上限
'      Case 16
'         If CheckUse("frm140503", strExec) = True Then
'            frm140503.Show
'         End If
'end 2015/10/14
End Select
End Sub
Private Sub mnu1101_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1  '解除期限
           If CheckUse("frm110101_1", strExec) = True Then
              frm110101_1.Show
           End If
          
      Case 2  '取消收文
           If CheckUse("frm110102_1", strExec) = True Then
              frm110102_1.Show
           End If
          
      Case 3  '閉卷
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

Private Sub mnu110401_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   
   Select Case Index
      Case 1  '帳單輸入
         If CheckUse("Frmacc2150", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2150.Show
      Case 2  '抵帳單輸入
         If CheckUse("Frmacc2160", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc2160.Show
      Case 3  '帳單作廢作業
         If CheckUse("Frmacc21j0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21j0.Show
      Case 4  '請款單輸入
         If CheckUse("Frmacc21h0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21h0.Show
      Case 5  '折讓輸入
         If CheckUse("Frmacc21i0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21i0.Show
         ToolShow
         tool8_enabled
         Screen.MousePointer = vbDefault
         Exit Sub
      Case 6  '請款單作廢作業
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
         '2007/11/29 ADD BY SONIA 外法使用者預設條件
         If (Mid(GetStaffDepartment(strUserNum), 1, 2) = "F3" Or Mid(GetStaffDepartment(strUserNum), 1, 2) = "F4") Then
            Frmacc24i0.Text7 = "F30"
            Frmacc24i0.Text8 = "F49"
         End If
         '2007/11/29 END
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
         '2007/11/6 ADD BY SONIA 外法使用者預設條件
         If (Mid(GetStaffDepartment(strUserNum), 1, 2) = "F3" Or Mid(GetStaffDepartment(strUserNum), 1, 2) = "F4") Then
            Frmacc24c0.Text9 = "F30"
            Frmacc24c0.Text10 = "F49"
         End If
         '2007/11/6 END
         Frmacc24c0.Show
      Case 7   '代理人逾期帳款分析表
         If CheckUse("Frmacc24f0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc24f0.Show
      Case 8   '折讓單列印
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
      'Added by Lydia 2018/11/30
      Case 10 '請款單折扣案件明細
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

Private Sub mnu15_Click(Index As Integer)
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

Private Sub mnu1501_Click(Index As Integer)
   ToolHide
   
   'Modify by Morgan 2011/9/16
   '設定控制台&Word預設印表機
   Load frm880011
   pub_OsPrinter = PUB_GetOsDefaultPrinter
   frm880011.Show 1
   
   Select Case Index
      Case 1 '橫式
          PrinterLetterDemand "1"
      Case 2 '直式
          PrinterLetterDemand "2"
      Case 3 '日文
          PrinterLetterDemand "3"
   End Select
    
   PUB_SetOsDefaultPrinter pub_OsPrinter

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

Private Sub mnu98_Click(Index As Integer)
ToolHide
   'Added by Lydia 2016/09/12
   Select Case Index
      Case 0 '系統印表機設定
         frm880011.bolAppOnly = True
         frm880011.Show 1

      Case 1 '報表紙張格式設定
         frm880013.Show vbModal

      Case 2 '解除畫面擷取限制
         frmChgUser.Caption = "解除畫面擷取限制"
         frmChgUser.SSTab1.TabVisible(1) = True
         frmChgUser.SSTab1.TabVisible(0) = False
         frmChgUser.Show
   End Select
   'end 2016/09/12
End Sub

Public Sub ToolHide()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
End Sub

Private Sub Timer1_Timer()
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
'         mnuTitle(4).Enabled = False
'         mnuTitle(10).Enabled = False
'         mnuTitle(11).Enabled = False
'         mnuTitle(15).Enabled = False
'      End If
'   Else
'      If tmpFormI = 1 Then
'         If mnuTitle(0).Enabled = False Then
'            mnuTitle(0).Enabled = True
'            mnuTitle(4).Enabled = True
'            mnuTitle(10).Enabled = True
'            mnuTitle(11).Enabled = True
'            mnuTitle(15).Enabled = True
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
      
      'Added by Morgan 2014/7/17
      If bXForm Then
        frmX.cmdOK(0).Enabled = False
        frmX.cmdOK(1).Enabled = False
      End If
      'end 2014/7/17
   Else
       If mnuTitle(10).Enabled = False Then mnuTitle(10).Enabled = True
      
      'Added by Morgan 2014/7/17
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
       If mnu11(9).Enabled = True Then mnu11(9).Enabled = False
       'add by nickc 2008/05/01
       'Removed by Morgan 2014/7/17
       'If mnu1602(2).Enabled = True Then mnu1602(2).Enabled = False
   Else
       'Modify By Cheng 2002/12/13
   '    If mnu11(6).Enabled = False Then mnu11(6).Enabled = True
   '    If mnu11(7).Enabled = False Then mnu11(7).Enabled = True
       If mnu11(9).Enabled = False Then mnu11(9).Enabled = True
       'add by nickc 2008/05/01
       'Removed by Morgan 2014/7/17
       'If mnu1602(2).Enabled = False Then mnu1602(2).Enabled = True
   End If
   
   'Add By Sindy 2009/08/26
   '控制"視窗"Menu
   MenuForFormControl
   mnuTitle(99).Visible = mnuTitle(99).Enabled
   '2009/08/26 End
   
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

'Add By Sindy 2009/08/26
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

Private Sub SalesDueCaseQuery()
Dim tmpST15 As String
'add by nickc 2005/09/20
Dim tmpST03 As String
   
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
      'edit by nickc 2006/04/04 加入非業務但有收文的人
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Then
      '2006/6/1 MODIFY BY SONIA 取消P2,P3中所之控制,改由PUB_ChkNotSalesButHaveCase控制
      'If UCase(Mid(tmpST15, 1, 1)) = "S" Or UCase(Mid(tmpST03, 1, 2)) = "P2" Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
      If UCase(Mid(tmpST15, 1, 1)) = "S" Or PUB_ChkNotSalesButHaveCase(strUserNum) Then
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
         If Pub_StrUserSt03 <> "M51" Then
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
         End If
      End If
   End If
End Sub

'Add By Sindy 2009/08/26
Private Sub FcpDueCaseQurey()
'Dim strSql As String, bolRun As Boolean
Dim tmpST15 As String
   
   'Modify By Sindy 2020/8/25 Law:限LXX部門人員
   tmpST15 = PUB_GetStaffST15(strUserNum, 1)
   If UCase(Mid(tmpST15, 1, 1)) <> "L" Then Exit Sub
   '2020/8/25 END
   
   '電腦中心,89037.蘇月星除外
   If Pub_StrUserSt03 <> "M51" And strUserNum <> "89037" Then
'      bolRun = False
'      strSql = "select count(*) from caseprogress " & _
'                     "where cp01 in ('CFL','FCL','LIN') " & _
'                     "and cp27||cp57 is null " & _
'                     "and not cp06 is null " & _
'                     "and cp14='" & strUserNum & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If RsTemp.Fields(0) > 0 Then
'            bolRun = True
'         End If
'      End If
'      If strGroup = "F1" Or strGroup = "F2" Or strGroup = "D4" Or strGroup = "G1" Or bolRun = True Then
'         If CheckUse("frm082007", strExec, False) = True Or bolRun = True Then
'            strSql = "select * from executelog where el01='frm082007' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI <> 1 Then
'               Load frm082007
'               frm082007.cmdQuery(0).Value = True
'            End If
'         End If
'      End If
      
      'Add By Sindy 2011/6/22 ST11為G1者,其電腦自動執行一次即可。
'      If strGroup = "F1" Or strGroup = "F2" Or strGroup = "D4" Or strGroup = "G1" Then
        'modify by sonia 2018/4/24 桂所長76012先不看期限
        If CheckUse("frm072005", strExec, False) = True And strUserNum <> "76012" Then
           strSql = "select * from executelog where el01='frm072005' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) '& " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
           intI = 1
           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
           If intI <> 1 Then
              Load frm072005
              frm072005.CmdQuery(0).Value = True
           End If
        End If
'      End If
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

'Add By Sindy 2015/10/20
Public Sub SetTmpfrm071004()
   Set Tmpfrm071004 = frm071004
End Sub
Public Sub SetTmpfrm071005()
   Set Tmpfrm071005 = frm071005
End Sub
Public Sub SetTmpfrm1103_2()
   Set Tmpfrm1103_2 = frm1103_2
End Sub
'2015/10/20 END

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
'Added by Lydia 2016/03/31 查名單電子化-主功能表
Private Sub SetTmpTMQ()
   Set frm090127.Tmpfrm090126 = frm090126
   Set frm090127.Tmpfrm090128 = frm090128
   Set frm090128.Tmpfrm090129 = frm090129
   Set frm090801.Tmpfrm090126 = frm090126 'Added by Lydia 2016/05/10
End Sub
'end 2015/10/14

'Added by Morgan 2020/1/17
'以名稱取得表單--通用不可刪
Public Function GetForm(pFormName As String) As Form
   Select Case pFormName
   '新增專案會用到的Form
   'Added by Morgan 2020/4/22
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
   'Add By Sindy 2023/6/20
   Case "frm090201_d"
      Set GetForm = frm090201_d
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
   'Added by Lydia 2024/07/29
   Case "frm100101_6" '顧問基本檔
         Set GetForm = frm100101_6
   Case "frm100101_5" '法務基本檔
         Set GetForm = frm100101_5
   Case "frm100101_2" '進度檔
         Set GetForm = frm100101_2
   End Select
End Function

'Added by Lydia 2020/04/15 智慧所案件
Private Sub mnu06_Click(Index As Integer)
    Select Case Index
        Case 1 '顧問記錄 'Memo by Lydia 2020/04/27  4/28上線
            If CheckUse("frm077001", strExec) = True Then
               frm077001.SetData "LA", 0, True
               frm077001.SetData "999999", 1, False
               frm077001.SetData "0", 2, False
               frm077001.SetData "00", 3, False
               frm077001.Show
               frm077001.QueryData
            End If
        Case 2 '顧問明細及統計 'Memo by Lydia 2020/04/27  4/28上線
            If CheckUse("frm077002", strExec) = True Then
               frm077002.Show
            End If
        Case 3 '介紹案源管理
            If CheckUse("frm077003", strExec) = True Then
               frm077003.Show
            End If
        Case 4 '介紹案源查詢
            frm077004.Caption = "介紹案源查詢"
            frm077004.Show 'Add By Sindy 2020/5/5
        'Added by Lydia 2020/06/16
        Case 5  '智財訴訟案需專業部配合通知補收文作業
            If CheckUse("frm077005", strExec) = True Then
               frm077005.Show
            End If
    End Select
End Sub

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

'Add By Sindy 2021/7/16
Public Sub frm090202_4CallFrm(strSendFrmType As String, m_PA11 As String, _
   m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, _
   m_EEP01 As String)
Dim m_SendRecvForm As Form '發文作業
   Select Case strSendFrmType
      Case "A"
         Set m_SendRecvForm = frm071004
         m_SendRecvForm.Show
         'm_SendRecvForm.bolIsEMPFlow = True
         m_SendRecvForm.Option1.Value = True
         m_SendRecvForm.txtDNum.Text = m_EEP01 '收文號
         m_SendRecvForm.cmdGoInput(0).Value = True
         Set m_SendRecvForm = Nothing
   End Select
End Sub
